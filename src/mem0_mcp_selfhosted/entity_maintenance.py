"""One-off maintenance for the entity store that entity_store.py queries by index.

    python -m mem0_mcp_selfhosted.entity_maintenance           # dry run: report only
    python -m mem0_mcp_selfhosted.entity_maintenance --apply   # write

Connects with the server's MEM0_QDRANT_URL, MEM0_QDRANT_API_KEY and MEM0_COLLECTION environment
variables (no .env loading). With --apply it creates the data_norm and linked_memory_ids keyword
indexes, then:

- backfills data_norm on entities written before the patch, or with a stale value;
- merges entities that share normalized text within one user/agent/run scope into the one with the
  most links, keeping every link;
- removes links to memories that no longer exist and deletes entities left with none.

Stop the MCP server first: an entity write that lands between the scan and the apply can be lost.
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass, field

from qdrant_client import models

from mem0_mcp_selfhosted.entity_store import (
    LINK_FIELD,
    NORM_FIELD,
    SCOPE_KEYS,
    create_entity_indexes,
    normalize_entity_text,
)
from mem0_mcp_selfhosted.env import env, opt_env

_PAGE_SIZE = 1000


@dataclass
class MaintenancePlan:
    set_norm: dict = field(default_factory=dict)  # entity id -> data_norm
    set_links: dict = field(default_factory=dict)  # entity id -> linked_memory_ids
    delete: list = field(default_factory=list)  # entity ids
    merged_groups: int = 0
    orphan_links_removed: int = 0

    def is_empty(self) -> bool:
        return not (self.set_norm or self.set_links or self.delete)


def plan_entity_maintenance(rows, memory_ids) -> MaintenancePlan:
    """Changes needed for entity ``rows`` (Qdrant records), given the ids of memories that still exist."""
    groups: dict[tuple, list] = {}
    for row in rows:
        payload = row.payload or {}
        if isinstance(payload.get("data"), str):
            scope = tuple(payload.get(k) for k in SCOPE_KEYS)
            groups.setdefault((scope, normalize_entity_text(payload["data"])), []).append(row)

    plan = MaintenancePlan()
    for (_, norm), members in groups.items():
        kept = {}
        for row in members:
            links = row.payload.get(LINK_FIELD)
            links = links if isinstance(links, list) else []
            kept[row.id] = [mid for mid in links if mid in memory_ids]
            plan.orphan_links_removed += len(links) - len(kept[row.id])
        keeper = max(members, key=lambda r: len(kept[r.id]))  # the first row wins a tie
        others = [row for row in members if row is not keeper]
        if others:
            plan.merged_groups += 1
            plan.delete.extend(row.id for row in others)
        links = list(dict.fromkeys(mid for row in (keeper, *others) for mid in kept[row.id]))
        if not links:
            plan.delete.append(keeper.id)
            continue
        if links != keeper.payload.get(LINK_FIELD):
            plan.set_links[keeper.id] = links
        if keeper.payload.get(NORM_FIELD) != norm:
            plan.set_norm[keeper.id] = norm
    return plan


def apply_entity_maintenance(client, collection_name, plan) -> None:
    """Write ``plan``: payload changes first, then deletes, so an interrupted merge never loses links."""
    for point_id in dict.fromkeys([*plan.set_norm, *plan.set_links]):
        payload = {}
        if point_id in plan.set_norm:
            payload[NORM_FIELD] = plan.set_norm[point_id]
        if point_id in plan.set_links:
            payload[LINK_FIELD] = plan.set_links[point_id]
        client.set_payload(collection_name=collection_name, payload=payload, points=[point_id])
    if plan.delete:
        client.delete(collection_name=collection_name, points_selector=models.PointIdsList(points=list(plan.delete)))


def _connect():
    from mem0.memory.main import _entity_collection_name
    from qdrant_client import QdrantClient

    memories = env("MEM0_COLLECTION", "mem0_mcp_selfhosted")
    client = QdrantClient(
        url=env("MEM0_QDRANT_URL", "http://localhost:6333"), api_key=opt_env("MEM0_QDRANT_API_KEY") or None
    )
    return client, _entity_collection_name("qdrant", memories), memories


def _scroll_all(client, collection_name, with_payload):
    offset = None
    while True:
        points, offset = client.scroll(
            collection_name=collection_name,
            limit=_PAGE_SIZE,
            offset=offset,
            with_payload=with_payload,
            with_vectors=False,
        )
        yield from points
        if offset is None:
            return


def _load(client, entity_collection, memory_collection):
    rows = list(_scroll_all(client, entity_collection, with_payload=True))
    memory_ids = {str(point.id) for point in _scroll_all(client, memory_collection, with_payload=False)}
    return rows, memory_ids


def _describe(plan, collection_name, entity_count, memory_count) -> str:
    return (
        f"{collection_name}: {entity_count} entities, {memory_count} memories | "
        f"data_norm to backfill: {len(plan.set_norm)} | duplicate groups to merge: {plan.merged_groups} | "
        f"stale links to remove: {plan.orphan_links_removed} | entities to relink: {len(plan.set_links)} | "
        f"entities to delete: {len(plan.delete)}"
    )


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        prog="python -m mem0_mcp_selfhosted.entity_maintenance",
        description="Backfill data_norm, merge duplicate entities and prune stale links in mem0's entity store.",
    )
    parser.add_argument("--apply", action="store_true", help="write the changes (default: dry run, report only)")
    args = parser.parse_args(argv)

    client, entities, memories = _connect()
    rows, memory_ids = _load(client, entities, memories)
    plan = plan_entity_maintenance(rows, memory_ids)
    print(_describe(plan, entities, len(rows), len(memory_ids)))
    if not args.apply:
        print("dry run: nothing written (pass --apply to write)")
        return 0

    create_entity_indexes(client, entities)
    apply_entity_maintenance(client, entities, plan)
    rows, memory_ids = _load(client, entities, memories)
    remaining = plan_entity_maintenance(rows, memory_ids)
    if remaining.is_empty():
        print(f"applied: {entities} is clean ({len(rows)} entities)")
        return 0
    print("applied, but changes remain: " + _describe(remaining, entities, len(rows), len(memory_ids)))
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
