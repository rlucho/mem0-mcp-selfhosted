"""Tests for entity_store.py — indexed entity lookups that replace mem0ai's 10,000-row listings."""

from __future__ import annotations

import pytest
from qdrant_client import models

from mem0_mcp_selfhosted import entity_store as es
from tests.unit.qdrant_fakes import FakeEmbedder, FakeEntityStore, FakeQdrantClient, seed_entities

CAP = 10_000
USER = {"user_id": "lucho"}


class FakeMemory:
    """Minimal host for the replacement methods: a lazily created entity store plus an embedder, as on mem0's Memory."""

    _existing_entities_by_text = es.existing_entities_by_text
    _remove_memory_from_entity_store = es.remove_memory_from_entity_store
    _link_entities_for_memory = es.link_entities_for_memory

    def __init__(self, store, embedder=None):
        self._pending_store = store
        self._entity_store = None
        self.embedding_model = embedder or FakeEmbedder()

    @property
    def entity_store(self):
        if self._entity_store is None:
            self._entity_store = self._pending_store
        return self._entity_store


@pytest.fixture
def store():
    return FakeEntityStore()


@pytest.fixture
def extracted(monkeypatch):
    """Entities the (stubbed) spaCy extractor reports for any text."""
    found: list[tuple[str, str]] = []
    monkeypatch.setattr(es, "_extract_entities", lambda text: list(found))
    return found


class TestNormalizeEntityText:
    def test_lowercases_trims_and_collapses_whitespace(self):
        assert es.normalize_entity_text("  Proxmox   Backup\tServer ") == "proxmox backup server"


class TestExistingEntitiesByText:
    def test_finds_entity_beyond_first_10000_rows(self, store):
        seed_entities(store, CAP + 50)
        row = FakeMemory(store)._existing_entities_by_text(USER).get("entity 10049")
        assert row is not None and row.id == "e010049"

    def test_never_lists_the_whole_collection(self, store):
        seed_entities(store, CAP + 50)
        FakeMemory(store)._existing_entities_by_text(USER).get("entity 10049")
        limits = [c["limit"] for c in store.client.calls_named("scroll")]
        assert limits and max(limits) <= 10, limits

    def test_is_scoped_to_the_callers_user(self, store):
        seed_entities(store, 1, user="alice")
        seed_entities(store, 1, user="bob", start=1, text=lambda i: "Entity 00000")
        row = FakeMemory(store)._existing_entities_by_text({"user_id": "bob"}).get("entity 00000")
        assert row is not None and row.payload["user_id"] == "bob"

    def test_missing_or_empty_key_returns_default(self, store):
        seed_entities(store, 3)
        lookup = FakeMemory(store)._existing_entities_by_text(USER)
        assert lookup.get("no such entity") is None
        assert lookup.get("") is None
        assert lookup.get("nope", "sentinel") == "sentinel"

    def test_query_errors_return_default_so_semantic_dedup_still_runs(self, store, monkeypatch):
        def boom(*a, **k):
            raise RuntimeError("qdrant down")

        monkeypatch.setattr(store.client, "scroll", boom)
        assert FakeMemory(store)._existing_entities_by_text(USER).get("entity 00001") is None

    def test_unavailable_entity_store_returns_empty_lookup(self):
        class BrokenMemory(FakeMemory):
            @property
            def entity_store(self):
                raise RuntimeError("qdrant down")

        assert BrokenMemory(None)._existing_entities_by_text(USER).get("entity 00001") is None


class TestRemoveMemoryFromEntityStore:
    def test_unlinks_memory_from_entities_beyond_first_10000_rows(self, store):
        seed_entities(store, CAP + 50)
        store.rows["e010040"].payload["linked_memory_ids"] = ["target", "keep-me"]
        store.rows["e000003"].payload["linked_memory_ids"] = ["target"]
        FakeMemory(store)._remove_memory_from_entity_store("target", USER)
        assert store.rows["e010040"].payload["linked_memory_ids"] == ["keep-me"]
        assert "e000003" not in store.rows

    def test_does_not_re_embed_entities(self, store):
        seed_entities(store, 5)
        store.rows["e000001"].payload["linked_memory_ids"] = ["target", "other"]
        mem = FakeMemory(store)
        mem._remove_memory_from_entity_store("target", USER)
        assert mem.embedding_model.embed_calls == [] and mem.embedding_model.batch_calls == []

    def test_works_before_the_entity_store_was_first_used(self, store):
        seed_entities(store, 2)
        store.rows["e000001"].payload["linked_memory_ids"] = ["target"]
        mem = FakeMemory(store)
        assert mem._entity_store is None
        mem._remove_memory_from_entity_store("target", USER)
        assert "e000001" not in store.rows

    def test_leaves_other_users_entities_alone(self, store):
        seed_entities(store, 1, user="alice", links=lambda i: ["shared-id"])
        seed_entities(store, 1, user="bob", start=1, links=lambda i: ["shared-id"])
        FakeMemory(store)._remove_memory_from_entity_store("shared-id", {"user_id": "bob"})
        assert "e000000" in store.rows and "e000001" not in store.rows

    def test_client_errors_are_logged_not_raised(self, store, monkeypatch):
        def boom(*a, **k):
            raise RuntimeError("qdrant down")

        monkeypatch.setattr(store.client, "scroll", boom)
        FakeMemory(store)._remove_memory_from_entity_store("target", USER)  # must not raise


class TestLinkEntitiesForMemory:
    def test_existing_entity_gets_the_link_without_being_embedded(self, store, extracted):
        seed_entities(store, 3)
        extracted.append(("TOPIC", "ENTITY 00001"))
        mem = FakeMemory(store)
        mem._link_entities_for_memory("new-mem", "text", USER)
        assert store.rows["e000001"].payload["linked_memory_ids"] == ["mem-00001", "new-mem"]
        assert mem.embedding_model.embed_calls == [] and mem.embedding_model.batch_calls == []

    def test_new_entities_are_embedded_in_one_batch_and_stored_with_normalized_text(self, store, extracted):
        extracted.extend([("PROPER", "Proxmox VE"), ("TOPIC", "backup job"), ("QUOTED", "zq7431")])
        mem = FakeMemory(store)
        mem._link_entities_for_memory("new-mem", "text", USER)
        assert mem.embedding_model.batch_calls == [["Proxmox VE", "backup job", "zq7431"]]
        assert mem.embedding_model.embed_calls == []
        stored = {r.payload["data_norm"]: r.payload for r in store.rows.values()}
        assert set(stored) == {"proxmox ve", "backup job", "zq7431"}
        assert all(p["linked_memory_ids"] == ["new-mem"] and p["user_id"] == "lucho" for p in stored.values())

    def test_close_semantic_match_gets_the_link_instead_of_a_new_entity(self, store, extracted):
        seed_entities(store, 1, text=lambda i: "Proxmox Virtual Environment")
        store.search_results["Proxmox VE"] = [
            models.ScoredPoint(id="e000000", version=0, score=0.97, payload=dict(store.rows["e000000"].payload))
        ]
        extracted.append(("PROPER", "Proxmox VE"))
        FakeMemory(store)._link_entities_for_memory("new-mem", "text", USER)
        assert len(store.rows) == 1
        assert store.rows["e000000"].payload["linked_memory_ids"] == ["mem-00000", "new-mem"]

    def test_weak_semantic_match_creates_a_new_entity(self, store, extracted):
        seed_entities(store, 1, text=lambda i: "Proxmox Virtual Environment")
        store.search_results["Proxmox VE"] = [models.ScoredPoint(id="e000000", version=0, score=0.90, payload={})]
        extracted.append(("PROPER", "Proxmox VE"))
        FakeMemory(store)._link_entities_for_memory("new-mem", "text", USER)
        assert len(store.rows) == 2

    def test_repeated_mentions_are_linked_once(self, store, extracted):
        extracted.extend([("TOPIC", "Backup Job"), ("TOPIC", "backup  job")])
        mem = FakeMemory(store)
        mem._link_entities_for_memory("new-mem", "text", USER)
        assert len(store.rows) == 1
        assert mem.embedding_model.batch_calls == [["Backup Job"]]

    def test_existing_link_is_not_duplicated(self, store, extracted):
        seed_entities(store, 1, links=lambda i: ["new-mem"])
        extracted.append(("TOPIC", "Entity 00000"))
        FakeMemory(store)._link_entities_for_memory("new-mem", "text", USER)
        assert store.rows["e000000"].payload["linked_memory_ids"] == ["new-mem"]

    @pytest.mark.parametrize("failure", ["raises", "short"])
    def test_failed_batch_embed_falls_back_to_one_embed_per_entity(self, store, extracted, failure):
        extracted.extend([("PROPER", "Proxmox VE"), ("TOPIC", "backup job")])
        embedder = FakeEmbedder()

        def broken_batch(texts, memory_action=None):
            if failure == "raises":
                raise RuntimeError("batch embed failed")
            return [[1.0]]

        embedder.embed_batch = broken_batch
        FakeMemory(store, embedder)._link_entities_for_memory("new-mem", "text", USER)
        assert embedder.embed_calls == ["Proxmox VE", "backup job"]
        assert {r.payload["data_norm"] for r in store.rows.values()} == {"proxmox ve", "backup job"}

    def test_extractor_failure_is_logged_not_raised(self, store, monkeypatch):
        def boom(text):
            raise RuntimeError("spaCy missing")

        monkeypatch.setattr(es, "_extract_entities", boom)
        FakeMemory(store)._link_entities_for_memory("new-mem", "text", USER)  # must not raise


class TestWrapEntityStore:
    def test_insert_adds_normalized_text(self, store):
        es.wrap_entity_store(store)
        store.insert(vectors=[[0.0]], payloads=[{"data": "  Mixed CASE  text", "linked_memory_ids": ["m"]}], ids=["x1"])
        assert store.rows["x1"].payload["data_norm"] == "mixed case text"

    def test_payload_update_adds_normalized_text_when_data_is_present(self, store):
        seed_entities(store, 1, norm=False)
        es.wrap_entity_store(store)
        store.update(vector_id="e000000", vector=None, payload={"data": "Entity 00000", "linked_memory_ids": ["a"]})
        assert store.rows["e000000"].payload["data_norm"] == "entity 00000"

    def test_wrapping_twice_does_not_stack(self, store):
        es.wrap_entity_store(store)
        es.wrap_entity_store(store)
        store.insert(vectors=[[0.0]], payloads=[{"data": "A"}], ids=["x"])
        assert len(store.client.calls_named("insert")) == 1
        assert len(store.client.calls_named("create_payload_index")) == 2

    def test_creates_keyword_indexes_for_the_lookup_fields(self, store):
        es.wrap_entity_store(store)
        created = {c["field_name"]: str(c["field_schema"]).lower() for c in store.client.calls_named("create_payload_index")}
        assert set(created) == {"data_norm", "linked_memory_ids"}
        assert all(schema.endswith("keyword") for schema in created.values())

    def test_skips_indexes_for_local_qdrant(self):
        local = FakeEntityStore(FakeQdrantClient(), is_local=True)
        es.wrap_entity_store(local)
        assert local.client.calls_named("create_payload_index") == []


class DummyMemory:
    def __init__(self):
        self._entity_store = None

    @property
    def entity_store(self):
        return self._entity_store

    def _existing_entities_by_text(self, filters):
        return "stock"

    def _remove_memory_from_entity_store(self, memory_id, filters):
        return "stock"

    def _link_entities_for_memory(self, memory_id, text, filters):
        return "stock"


class TestPatchEntityStoreScaling:
    @pytest.fixture(autouse=True)
    def _default_env(self, monkeypatch):
        monkeypatch.delenv("MEM0_ENTITY_INDEXED_LOOKUPS", raising=False)

    def test_replaces_the_entity_methods(self):
        cls = type("PatchedMemory", (DummyMemory,), {})
        assert es.patch_entity_store_scaling(cls) is True
        assert cls._existing_entities_by_text is es.existing_entities_by_text
        assert cls._remove_memory_from_entity_store is es.remove_memory_from_entity_store
        assert cls._link_entities_for_memory is es.link_entities_for_memory
        assert isinstance(cls.__dict__["entity_store"], property)

    def test_second_call_changes_nothing(self):
        cls = type("PatchedMemory", (DummyMemory,), {})
        es.patch_entity_store_scaling(cls)
        prop = cls.__dict__["entity_store"]
        assert es.patch_entity_store_scaling(cls) is False
        assert cls.__dict__["entity_store"] is prop

    def test_skips_when_mem0_internals_are_missing(self):
        def stock(self, filters):
            return None

        cls = type("OtherMemory", (), {"_existing_entities_by_text": stock})
        assert es.patch_entity_store_scaling(cls) is False
        assert cls.__dict__["_existing_entities_by_text"] is stock

    @pytest.mark.parametrize("value", ["false", "0", "off"])
    def test_disabled_by_env_flag(self, monkeypatch, value):
        monkeypatch.setenv("MEM0_ENTITY_INDEXED_LOOKUPS", value)
        cls = type("PatchedMemory", (DummyMemory,), {})
        assert es.patch_entity_store_scaling(cls) is False
        assert "_existing_entities_by_text" not in cls.__dict__

    def test_patched_entity_store_property_wraps_the_store(self):
        cls = type("PatchedMemory", (DummyMemory,), {})
        es.patch_entity_store_scaling(cls)
        inst = cls()
        inst._entity_store = FakeEntityStore()
        inst.entity_store.insert(vectors=[[0.0]], payloads=[{"data": "X  Y"}], ids=["p"])
        assert inst._entity_store.rows["p"].payload["data_norm"] == "x y"
