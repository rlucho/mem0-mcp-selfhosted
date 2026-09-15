"""Integration: indexed entity lookups vs mem0ai's 10,000-row listings, on throwaway Qdrant collections.

The ``stock`` variants run mem0ai's own methods and are expected to fail with an AssertionError
(strict xfail, ``raises=AssertionError``): they are the negative control that proves each test can
see the bug. Any other exception, including one raised by a fixture, fails the test instead of
passing itself off as the expected failure. If mem0ai fixes the listing upstream, the stock variants
start passing, the strict xfail turns red, and entity_store.py can be retired.
"""

from __future__ import annotations

import os
import tempfile
import uuid

import pytest

pytestmark = pytest.mark.integration

CAP = 10_000
REPLACED = ("_existing_entities_by_text", "_remove_memory_from_entity_store", "_link_entities_for_memory", "entity_store")
VARIANTS = [
    pytest.param(
        "stock",
        marks=pytest.mark.xfail(
            strict=True, raises=AssertionError, reason="mem0ai 2.0.18 lists at most 10,000 entity rows"
        ),
    ),
    pytest.param("patched"),
]


@pytest.fixture(scope="module")
def scaling_memory(qdrant_url, ollama_url):
    if os.environ.get("MEM0_LLM_PROVIDER", "anthropic") == "anthropic":
        from mem0_mcp_selfhosted.auth import resolve_token

        if not resolve_token():
            pytest.skip("No Anthropic token available (required for MEM0_LLM_PROVIDER=anthropic)")

    from qdrant_client import QdrantClient, models

    collection = f"mem0_entity_scaling_{uuid.uuid4().hex[:8]}"
    saved = {k: os.environ.get(k) for k in ("MEM0_COLLECTION", "MEM0_HISTORY_DB_PATH", "MEM0_ENABLE_GRAPH")}
    tmp = tempfile.TemporaryDirectory()
    os.environ["MEM0_COLLECTION"] = collection
    os.environ["MEM0_HISTORY_DB_PATH"] = os.path.join(tmp.name, "history.db")
    os.environ["MEM0_ENABLE_GRAPH"] = "false"
    try:
        from mem0_mcp_selfhosted.config import build_config
        from mem0_mcp_selfhosted.server import register_providers

        config_dict, providers_info, _ = build_config()
        register_providers(providers_info)

        from mem0 import Memory

        memory = Memory.from_config(config_dict)
        store = memory.entity_store  # create the entity collection up front
        # Exact search is fast enough for 10k points and spares the live Qdrant an HNSW build.
        store.client.update_collection(
            collection_name=store.collection_name,
            optimizers_config=models.OptimizersConfigDiff(indexing_threshold=0),
        )
        yield memory
    finally:
        client = QdrantClient(url=qdrant_url, api_key=os.environ.get("MEM0_QDRANT_API_KEY") or None)
        for name in (collection, f"{collection}_entities"):
            try:
                client.delete_collection(name)
            except Exception:
                pass
        for key, value in saved.items():
            if value is None:
                os.environ.pop(key, None)
            else:
                os.environ[key] = value
        tmp.cleanup()


@pytest.fixture
def stock():
    """Run mem0ai's own entity methods; refuse if something in this process already patched the class."""
    from mem0.memory.main import Memory

    ours = [
        name
        for name in REPLACED
        if getattr(getattr(Memory.__dict__[name], "fget", Memory.__dict__[name]), "__module__", "").startswith(
            "mem0_mcp_selfhosted"
        )
    ]
    if ours:
        raise RuntimeError(f"Memory is already patched in this process ({ours}); the stock control would be meaningless")


@pytest.fixture
def patched(monkeypatch):
    """Apply entity_store's patch to the real Memory class for one test, then restore mem0ai's methods."""
    from mem0.memory.main import Memory

    from mem0_mcp_selfhosted import entity_store as es

    for name in REPLACED:
        monkeypatch.setattr(Memory, name, Memory.__dict__[name])
    monkeypatch.setattr(Memory, es.PATCH_MARKER, False, raising=False)
    monkeypatch.delenv("MEM0_ENTITY_INDEXED_LOOKUPS", raising=False)
    assert es.patch_entity_store_scaling(Memory) is True


@pytest.fixture(scope="module")
def seeded(scaling_memory):
    """10,050 entity rows for one user; returns an entity that is NOT in the first 10,000 scrolled rows."""
    import numpy as np
    from qdrant_client import models

    store = scaling_memory.entity_store
    client, coll = store.client, store.collection_name
    user = "entity-scaling-user"
    rng = np.random.default_rng(7)
    ids = [str(uuid.uuid4()) for _ in range(CAP + 50)]
    texts = {pid: f"Scaling Entity {i:05d}" for i, pid in enumerate(ids)}
    for start in range(0, len(ids), 1000):
        chunk = ids[start:start + 1000]
        vecs = rng.standard_normal((len(chunk), 1024)).astype("float32")
        vecs = np.round(vecs / np.linalg.norm(vecs, axis=1, keepdims=True), 3)
        client.upsert(
            collection_name=coll,
            points=[
                models.PointStruct(
                    id=pid,
                    vector={"": vec.tolist()},
                    payload={
                        "data": texts[pid],
                        "data_norm": texts[pid].lower(),
                        "entity_type": "TOPIC",
                        "linked_memory_ids": [f"seed-{pid}"],
                        "user_id": user,
                    },
                )
                for pid, vec in zip(chunk, vecs)
            ],
        )
    head, _ = client.scroll(coll, limit=CAP, with_payload=False, with_vectors=False)
    first = {str(p.id) for p in head}
    if len(first) != CAP:  # not an assert: an AssertionError here would count as the stock variants' expected failure
        raise RuntimeError(f"setup: expected {CAP} rows in the first page, got {len(first)}")
    target = sorted(pid for pid in ids if pid not in first)[0]
    return scaling_memory, user, target, texts[target]


@pytest.mark.parametrize("variant", VARIANTS)
def test_exact_lookup_finds_entity_beyond_first_10000_rows(seeded, variant, request):
    memory, user, target_id, target_text = seeded
    request.getfixturevalue(variant)
    row = memory._existing_entities_by_text({"user_id": user}).get(target_text.lower())
    assert row is not None and str(row.id) == target_id


@pytest.mark.parametrize("variant", VARIANTS)
def test_delete_cleanup_unlinks_entity_beyond_first_10000_rows(seeded, variant, request):
    memory, user, target_id, _ = seeded
    request.getfixturevalue(variant)
    store = memory.entity_store
    doomed = f"cleanup-{variant}-{uuid.uuid4().hex[:6]}"
    store.client.set_payload(store.collection_name, payload={"linked_memory_ids": ["keep-me", doomed]}, points=[target_id])
    memory._remove_memory_from_entity_store(doomed, {"user_id": user})
    links = store.client.retrieve(store.collection_name, ids=[target_id], with_payload=True)[0].payload["linked_memory_ids"]
    assert doomed not in links and "keep-me" in links


def test_update_and_delete_link_then_unlink_entities(scaling_memory, patched):
    from qdrant_client import models

    user = "entity-roundtrip-user"
    added = scaling_memory.add([{"role": "user", "content": "roundtrip seed"}], user_id=user, infer=False)
    memory_id = added["results"][0]["id"]
    scaling_memory.update(memory_id, text="Lucho moved Proxmox Backup Server onto the Samsung SSD in Madrid")

    store = scaling_memory.entity_store
    linked_to_memory = models.Filter(must=[models.FieldCondition(key="linked_memory_ids", match=models.MatchValue(value=memory_id))])
    linked, _ = store.client.scroll(store.collection_name, scroll_filter=linked_to_memory, limit=100, with_payload=True)
    assert linked, "update() should link at least one entity"
    assert all(p.payload["data_norm"] == " ".join(p.payload["data"].lower().split()) for p in linked), [p.payload for p in linked]
    schema = store.client.get_collection(store.collection_name).payload_schema
    assert {"data_norm", "linked_memory_ids"} <= set(schema)

    scaling_memory.delete(memory_id)
    remaining, _ = store.client.scroll(store.collection_name, scroll_filter=linked_to_memory, limit=100)
    assert remaining == []


def test_maintenance_backfills_merges_and_prunes_real_rows(scaling_memory):
    from qdrant_client import models

    from mem0_mcp_selfhosted import entity_maintenance as em

    store = scaling_memory.entity_store
    client, coll = store.client, store.collection_name
    user = "entity-maintenance-user"
    rows = {
        str(uuid.uuid4()): {"data": "Duplicate Name", "linked_memory_ids": ["m1"]},
        str(uuid.uuid4()): {"data": "duplicate  NAME", "linked_memory_ids": ["m2", "gone-1"]},
        str(uuid.uuid4()): {"data": "Solo", "linked_memory_ids": ["gone-2"]},
    }
    client.upsert(
        collection_name=coll,
        points=[
            models.PointStruct(id=pid, vector={"": [1.0] + [0.0] * 1023}, payload={**p, "entity_type": "TOPIC", "user_id": user})
            for pid, p in rows.items()
        ],
    )
    scope = models.Filter(must=[models.FieldCondition(key="user_id", match=models.MatchValue(value=user))])
    current, _ = client.scroll(coll, scroll_filter=scope, limit=100, with_payload=True)
    plan = em.plan_entity_maintenance(current, {"m1", "m2"})
    em.apply_entity_maintenance(client, coll, plan)

    after, _ = client.scroll(coll, scroll_filter=scope, limit=100, with_payload=True)
    assert len(after) == 1
    survivor = after[0].payload
    assert survivor["data_norm"] == "duplicate name"
    assert sorted(survivor["linked_memory_ids"]) == ["m1", "m2"]
    assert em.plan_entity_maintenance(after, {"m1", "m2"}).is_empty()
