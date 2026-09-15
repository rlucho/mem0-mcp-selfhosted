"""Indexed entity-store lookups for mem0ai's ``Memory`` (Qdrant only).

mem0ai 2.0.18 finds entities by listing them. ``_existing_entities_by_text`` and
``_remove_memory_from_entity_store`` read at most ``top_k=10000`` rows per scope and search that page
in Python, so once a user has more than 10,000 entities, exact-text dedup stops seeing the rest
(duplicates pile up) and deleting or updating a memory leaves stale links behind. On the update path
every extracted entity also pays that listing plus an embedding call, and every unlinked entity is
re-embedded.

``patch_entity_store_scaling()`` swaps those three methods for filtered queries on two keyword-indexed
payload fields: ``data_norm`` (the normalized entity text, now written with every entity insert and
update) and ``linked_memory_ids``. Unlinking is payload-only; linking embeds only entities that are not
stored yet, in one batch. Only the synchronous ``Memory`` this server uses is patched.

Call it before ``Memory.from_config()``. Entities written before the patch need ``data_norm`` once:
``python -m mem0_mcp_selfhosted.entity_maintenance --apply``. ``MEM0_ENTITY_INDEXED_LOOKUPS=false``
keeps mem0ai's own methods.
"""

from __future__ import annotations

import inspect
import logging
import uuid

from qdrant_client import models

from mem0_mcp_selfhosted.env import bool_env

logger = logging.getLogger(__name__)

NORM_FIELD = "data_norm"
LINK_FIELD = "linked_memory_ids"
SCOPE_KEYS = ("user_id", "agent_id", "run_id")
SEMANTIC_MATCH_SCORE = 0.95  # same threshold as mem0ai's _upsert_entity
PATCH_MARKER = "_mem0_mcp_entity_store_patched"
_WRAPPED_MARKER = "_mem0_mcp_entity_store_wrapped"
_PAGE_SIZE = 256


def normalize_entity_text(value: str) -> str:
    """Same normalization as mem0ai's ``Memory._normalize_entity_text``."""
    return " ".join(value.strip().lower().split())


def _extract_entities(text: str) -> list[tuple[str, str]]:
    from mem0.utils.entity_extraction import extract_entities

    return extract_entities(text)


def _scope(filters) -> dict:
    return {k: v for k, v in (filters or {}).items() if k in SCOPE_KEYS and v}


def _match(key: str, value) -> models.FieldCondition:
    return models.FieldCondition(key=key, match=models.MatchValue(value=value))


def _with_norm(payload):
    """Copy of ``payload`` with ``data_norm`` set from ``data``; anything else is returned unchanged."""
    if isinstance(payload, dict) and isinstance(payload.get("data"), str):
        return {**payload, NORM_FIELD: normalize_entity_text(payload["data"])}
    return payload


class IndexedEntityLookup:
    """Stands in for mem0ai's ``{normalized text: row}`` dict. Each ``.get()`` is one indexed query."""

    def __init__(self, store, filters):
        self._store = store
        self._scope = [_match(k, v) for k, v in _scope(filters).items()]

    def get(self, key, default=None):
        if not key:
            return default
        try:
            points, _ = self._store.client.scroll(
                collection_name=self._store.collection_name,
                scroll_filter=models.Filter(must=[*self._scope, _match(NORM_FIELD, key)]),
                limit=1,
                with_payload=True,
                with_vectors=False,
            )
        except Exception as e:
            # Same fallback as mem0ai: without an exact match, callers dedup semantically.
            logger.debug("Exact entity lookup failed, falling back to semantic dedup: %s", e)
            return default
        return points[0] if points else default


def existing_entities_by_text(self, filters):
    """Replaces ``Memory._existing_entities_by_text``. Callers only ever call ``.get()`` on the result."""
    try:
        store = self.entity_store
    except Exception as e:
        logger.debug("Exact entity lookup failed, falling back to semantic dedup: %s", e)
        return {}
    return IndexedEntityLookup(store, filters)


def remove_memory_from_entity_store(self, memory_id, filters):
    """Replaces ``Memory._remove_memory_from_entity_store``: strip ``memory_id`` from every linked entity.

    Finds the entities by the indexed link field, rewrites their links payload-only (no re-embedding)
    and deletes any entity left without links. Unlike mem0ai's version it also runs when this process
    has not used the entity store yet, e.g. a delete right after a restart. Never raises.
    """
    try:
        store = self.entity_store
        linked = models.Filter(
            must=[*(_match(k, v) for k, v in _scope(filters).items()), _match(LINK_FIELD, memory_id)]
        )
        offset = None
        while True:
            points, offset = store.client.scroll(
                collection_name=store.collection_name,
                scroll_filter=linked,
                limit=_PAGE_SIZE,
                offset=offset,
                with_payload=True,
                with_vectors=False,
            )
            for point in points:
                try:
                    _unlink(store, point, memory_id)
                except Exception as e:
                    logger.debug("Entity cleanup failed for id=%s: %s", point.id, e)
            if offset is None:
                break
    except Exception as e:
        logger.warning("Entity store cleanup failed for memory_id=%s: %s", memory_id, e)


def _unlink(store, point, memory_id) -> None:
    remaining = [mid for mid in (point.payload or {}).get(LINK_FIELD, []) if mid != memory_id]
    if remaining:
        store.client.set_payload(
            collection_name=store.collection_name, payload={LINK_FIELD: remaining}, points=[point.id]
        )
    else:
        store.client.delete(
            collection_name=store.collection_name, points_selector=models.PointIdsList(points=[point.id])
        )


def link_entities_for_memory(self, memory_id, text, filters):
    """Replaces ``Memory._link_entities_for_memory``: extract entities from ``text`` and link them.

    Entities already stored under the same normalized text are linked without being embedded. The
    rest are embedded in one batch, then linked to a close semantic match (score >= 0.95, as in
    mem0ai) or inserted. Never raises.
    """
    try:
        entities = _extract_entities(text)
        if not entities:
            return
        store = self.entity_store
        scope = _scope(filters)
        lookup = IndexedEntityLookup(store, scope)
        pending = []
        seen = set()
        for entity_type, entity_text in entities:
            key = normalize_entity_text(entity_text)
            if not key or key in seen:
                continue
            seen.add(key)
            row = lookup.get(key)
            if row is None:
                pending.append((entity_type, entity_text, key))
                continue
            try:
                _add_link(store, row, memory_id)
            except Exception as e:
                logger.debug("Entity link failed for '%s': %s", entity_text, e)
        if not pending:
            return

        vectors = _embed_batch(self.embedding_model, [entity_text for _, entity_text, _ in pending])
        for i, (entity_type, entity_text, key) in enumerate(pending):
            try:
                vector = vectors[i] if vectors is not None else self.embedding_model.embed(entity_text, "add")
                hits = store.search(query=entity_text, vectors=vector, top_k=1, filters=scope)
                if hits and hits[0].score >= SEMANTIC_MATCH_SCORE:
                    _add_link(store, hits[0], memory_id)
                else:
                    payload = {"data": entity_text, "entity_type": entity_type, LINK_FIELD: [memory_id], NORM_FIELD: key}
                    store.insert(vectors=[vector], ids=[str(uuid.uuid4())], payloads=[{**payload, **scope}])
            except Exception as e:
                logger.debug("Entity link failed for '%s': %s", entity_text, e)
    except Exception as e:
        logger.warning("Entity linking failed for memory_id=%s: %s", memory_id, e)


def _add_link(store, row, memory_id) -> None:
    payload = row.payload or {}
    links = list(payload.get(LINK_FIELD) or [])
    if memory_id in links:
        return
    store.update(vector_id=row.id, vector=None, payload=_with_norm({**payload, LINK_FIELD: [*links, memory_id]}))


def _embed_batch(embedder, texts):
    """One embedding call for all ``texts``, or None when that call fails or returns the wrong count."""
    try:
        vectors = embedder.embed_batch(texts, "add")
    except Exception as e:
        logger.debug("Batch entity embedding failed, embedding one by one: %s", e)
        return None
    if len(vectors) != len(texts):
        logger.warning(
            "embed_batch returned %d vectors for %d entity texts, embedding one by one", len(vectors), len(texts)
        )
        return None
    return vectors


def create_entity_indexes(client, collection_name: str) -> None:
    """Keyword indexes on the two payload fields the lookups filter by. Qdrant treats a repeat as a no-op."""
    for field_name in (NORM_FIELD, LINK_FIELD):
        client.create_payload_index(
            collection_name=collection_name, field_name=field_name, field_schema=models.PayloadSchemaType.KEYWORD
        )


def wrap_entity_store(store):
    """Make ``store`` write ``data_norm`` on every insert and payload update and ensure the lookup indexes exist.

    This also covers the entity writes mem0ai still makes itself (batch linking in ``add()``).
    Wrapping an already wrapped store does nothing.
    """
    if getattr(store, _WRAPPED_MARKER, False):
        return store
    insert, update = store.insert, store.update

    def insert_with_norm(vectors, payloads=None, ids=None):
        return insert(vectors=vectors, payloads=[_with_norm(p) for p in payloads] if payloads else payloads, ids=ids)

    def update_with_norm(vector_id, vector=None, payload=None):
        return update(vector_id=vector_id, vector=vector, payload=_with_norm(payload))

    store.insert = insert_with_norm
    store.update = update_with_norm
    if not getattr(store, "is_local", False):  # embedded Qdrant has no payload indexes
        try:
            create_entity_indexes(store.client, store.collection_name)
        except Exception as e:
            logger.warning("Could not create entity lookup indexes on %s: %s", store.collection_name, e)
    setattr(store, _WRAPPED_MARKER, True)
    return store


_REPLACEMENTS = {
    "_existing_entities_by_text": existing_entities_by_text,
    "_remove_memory_from_entity_store": remove_memory_from_entity_store,
    "_link_entities_for_memory": link_entities_for_memory,
}


def patch_entity_store_scaling(memory_cls=None) -> bool:
    """Swap mem0ai's listing-based entity methods for indexed lookups. Returns True if this call patched.

    Must run before ``Memory.from_config()``. Does nothing when ``MEM0_ENTITY_INDEXED_LOOKUPS`` is false,
    when the class is already patched, or when mem0ai's entity internals no longer look as expected.
    """
    if not bool_env("MEM0_ENTITY_INDEXED_LOOKUPS", "true"):
        logger.info("MEM0_ENTITY_INDEXED_LOOKUPS is off, keeping mem0ai's entity store methods")
        return False
    if memory_cls is None:
        from mem0.memory.main import Memory as memory_cls
    if getattr(memory_cls, PATCH_MARKER, False):
        return False
    prop = inspect.getattr_static(memory_cls, "entity_store", None)
    missing = [name for name in _REPLACEMENTS if not callable(inspect.getattr_static(memory_cls, name, None))]
    if missing or not isinstance(prop, property):
        logger.warning(
            "mem0ai entity store internals changed (missing: %s), not patching",
            ", ".join(missing) or "entity_store property",
        )
        return False

    def entity_store(self):
        store = prop.fget(self)
        return wrap_entity_store(store) if store is not None else store

    for name, replacement in _REPLACEMENTS.items():
        setattr(memory_cls, name, replacement)
    memory_cls.entity_store = property(entity_store, doc=prop.__doc__)
    setattr(memory_cls, PATCH_MARKER, True)
    logger.info("Patched mem0ai entity store: indexed lookups instead of 10,000-row listings")
    return True
