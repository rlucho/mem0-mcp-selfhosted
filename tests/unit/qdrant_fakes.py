"""In-memory stand-ins for the Qdrant surface used by entity_store.py and entity_maintenance.py.

They implement only what those modules call, with Qdrant's semantics where it matters:
``MatchValue`` on a list field matches when the list contains the value, ``scroll`` pages by
point id with ``next_page_offset`` naming the first id of the next page, and ``set_payload``
merges keys into the existing payload.
"""

from __future__ import annotations

from types import SimpleNamespace

from qdrant_client import models


def _norm(value: str) -> str:
    # Deliberately independent of the code under test.
    return " ".join(value.strip().lower().split())


class FakeQdrantClient:
    def __init__(self):
        self.collections: dict[str, dict[str, SimpleNamespace]] = {}
        self.calls: list[tuple[str, dict]] = []

    def calls_named(self, name: str) -> list[dict]:
        return [kwargs for method, kwargs in self.calls if method == name]

    @staticmethod
    def _matches(payload: dict, flt) -> bool:
        if flt is None:
            return True
        for cond in flt.must or []:
            want = cond.match.value
            have = payload.get(cond.key)
            if isinstance(have, list):
                if want not in have:
                    return False
            elif have != want:
                return False
        return True

    def scroll(self, collection_name, scroll_filter=None, limit=10, offset=None,
               with_payload=True, with_vectors=False, **_):
        self.calls.append(("scroll", {"limit": limit, "filter": scroll_filter, "offset": offset}))
        rows = self.collections.get(collection_name, {})
        ids = [pid for pid in sorted(rows) if self._matches(rows[pid].payload, scroll_filter)]
        start = 0 if offset is None else next((i for i, pid in enumerate(ids) if pid >= offset), len(ids))
        page = ids[start:start + limit]
        next_offset = ids[start + limit] if start + limit < len(ids) else None
        return [models.Record(id=pid, payload=dict(rows[pid].payload)) for pid in page], next_offset

    def set_payload(self, collection_name, payload, points, **_):
        self.calls.append(("set_payload", {"payload": dict(payload), "points": list(points)}))
        for pid in points:
            self.collections[collection_name][pid].payload.update(payload)

    def delete(self, collection_name, points_selector, **_):
        ids = list(points_selector.points)
        self.calls.append(("delete", {"points": ids}))
        for pid in ids:
            self.collections[collection_name].pop(pid, None)

    def create_payload_index(self, collection_name, field_name, field_schema, **_):
        self.calls.append(("create_payload_index", {"field_name": field_name, "field_schema": field_schema}))


class FakeEntityStore:
    """The part of mem0's Qdrant wrapper that entity code touches.

    ``update(vector=None, payload=...)`` is a payload merge, exactly as in mem0's wrapper.
    """

    def __init__(self, client=None, collection_name="ents", is_local=False):
        self.client = client or FakeQdrantClient()
        self.collection_name = collection_name
        self.is_local = is_local
        self.client.collections.setdefault(collection_name, {})
        self.search_results: dict[str, list] = {}
        self.search_calls: list[dict] = []

    @property
    def rows(self) -> dict[str, SimpleNamespace]:
        return self.client.collections[self.collection_name]

    def insert(self, vectors, payloads=None, ids=None):
        self.client.calls.append(("insert", {"ids": list(ids), "payloads": [dict(p) for p in payloads]}))
        for pid, vector, payload in zip(ids, vectors, payloads):
            self.rows[pid] = SimpleNamespace(payload=dict(payload), vector=vector)

    def update(self, vector_id, vector=None, payload=None):
        if vector is None:
            if payload is not None:
                self.client.set_payload(self.collection_name, payload, [vector_id])
            return
        self.client.calls.append(("upsert", {"id": vector_id}))
        self.rows[vector_id] = SimpleNamespace(payload=dict(payload or {}), vector=vector)

    def search(self, query, vectors, top_k=5, filters=None):
        self.search_calls.append({"query": query, "top_k": top_k, "filters": filters})
        return self.search_results.get(query, [])


class FakeEmbedder:
    def __init__(self):
        self.embed_calls: list[str] = []
        self.batch_calls: list[list[str]] = []

    def embed(self, text, memory_action=None):
        self.embed_calls.append(text)
        return [float(len(text))]

    def embed_batch(self, texts, memory_action=None):
        self.batch_calls.append(list(texts))
        return [[float(len(t))] for t in texts]


def seed_entities(store, count, *, user="lucho", start=0, text=None, links=None, norm=True):
    """Add entity rows with sortable ids e000000, e000001, ... so 'beyond the first N rows' is deterministic."""
    for i in range(start, start + count):
        data = text(i) if text else f"Entity {i:05d}"
        payload = {
            "data": data,
            "entity_type": "TOPIC",
            "linked_memory_ids": list(links(i)) if links else [f"mem-{i:05d}"],
            "user_id": user,
        }
        if norm:
            payload["data_norm"] = _norm(data)
        store.rows[f"e{i:06d}"] = SimpleNamespace(payload=payload, vector=[0.0])
