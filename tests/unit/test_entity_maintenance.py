"""Tests for entity_maintenance.py — one-off backfill, duplicate merge and stale-link pruning."""

from __future__ import annotations

from types import SimpleNamespace

from qdrant_client import models

from mem0_mcp_selfhosted import entity_maintenance as em
from tests.unit.qdrant_fakes import FakeQdrantClient


def row(pid, data, links, *, user="u", norm=None):
    payload = {"data": data, "entity_type": "TOPIC", "linked_memory_ids": list(links), "user_id": user}
    if norm is not None:
        payload["data_norm"] = norm
    return models.Record(id=pid, payload=payload)


class TestPlanEntityMaintenance:
    def test_backfills_missing_or_stale_normalized_text(self):
        plan = em.plan_entity_maintenance(
            [row("a", "Proxmox VE", ["m1"]), row("b", "Samsung SSD", ["m1"], norm="old value")], {"m1"}
        )
        assert plan.set_norm == {"a": "proxmox ve", "b": "samsung ssd"}
        assert plan.set_links == {} and plan.delete == []

    def test_merges_same_text_in_same_scope_into_the_entity_with_most_links(self):
        rows = [
            row("a", "Backup Job", ["m1"], norm="backup job"),
            row("b", "backup  job", ["m2", "m3"], norm="backup job"),
        ]
        plan = em.plan_entity_maintenance(rows, {"m1", "m2", "m3"})
        assert plan.delete == ["a"]
        assert sorted(plan.set_links["b"]) == ["m1", "m2", "m3"]
        assert plan.merged_groups == 1

    def test_never_merges_across_users(self):
        rows = [
            row("a", "Backup Job", ["m1"], user="alice", norm="backup job"),
            row("b", "Backup Job", ["m2"], user="bob", norm="backup job"),
        ]
        assert em.plan_entity_maintenance(rows, {"m1", "m2"}).is_empty()

    def test_prunes_links_to_deleted_memories_and_deletes_entities_left_empty(self):
        rows = [row("a", "Keep", ["m1", "gone"], norm="keep"), row("b", "Drop", ["gone"], norm="drop")]
        plan = em.plan_entity_maintenance(rows, {"m1"})
        assert plan.set_links == {"a": ["m1"]}
        assert plan.delete == ["b"]
        assert plan.orphan_links_removed == 2

    def test_clean_collection_needs_no_changes(self):
        assert em.plan_entity_maintenance([row("a", "Keep", ["m1"], norm="keep")], {"m1"}).is_empty()


class TestApplyEntityMaintenance:
    def test_writes_payload_changes_and_deletes(self):
        client = FakeQdrantClient()
        client.collections["ents"] = {
            "a": SimpleNamespace(payload={"data": "Keep", "linked_memory_ids": ["m1", "gone"], "user_id": "u"}),
            "b": SimpleNamespace(payload={"data": "Drop", "linked_memory_ids": ["gone"], "user_id": "u"}),
        }
        plan = em.MaintenancePlan(set_norm={"a": "keep"}, set_links={"a": ["m1"]}, delete=["b"])
        em.apply_entity_maintenance(client, "ents", plan)
        assert client.collections["ents"]["a"].payload["data_norm"] == "keep"
        assert client.collections["ents"]["a"].payload["linked_memory_ids"] == ["m1"]
        assert "b" not in client.collections["ents"]


class TestMain:
    def test_dry_run_reports_without_writing(self, monkeypatch, capsys):
        rows = [row("a", "Keep", ["m1", "gone"], norm="keep")]
        monkeypatch.setattr(em, "_connect", lambda: (FakeQdrantClient(), "ents", "mems"))
        monkeypatch.setattr(em, "_load", lambda client, entities, memories: (rows, {"m1"}))

        def must_not_write(*a, **k):
            raise AssertionError("dry run must not write")

        monkeypatch.setattr(em, "apply_entity_maintenance", must_not_write)
        assert em.main([]) == 0
        out = capsys.readouterr().out.lower()
        assert "dry run" in out and "stale links" in out

    def test_apply_writes_indexes_and_changes_then_reports_clean(self, monkeypatch, capsys):
        client = FakeQdrantClient()
        client.collections["ents"] = {
            "a": SimpleNamespace(payload={"data": "Keep", "linked_memory_ids": ["m1", "gone"], "user_id": "u"}),
        }

        def load(c, entities, memories):
            return [models.Record(id=pid, payload=dict(r.payload)) for pid, r in c.collections[entities].items()], {"m1"}

        monkeypatch.setattr(em, "_connect", lambda: (client, "ents", "mems"))
        monkeypatch.setattr(em, "_load", load)
        assert em.main(["--apply"]) == 0
        assert client.collections["ents"]["a"].payload["data_norm"] == "keep"
        assert client.collections["ents"]["a"].payload["linked_memory_ids"] == ["m1"]
        assert {c["field_name"] for c in client.calls_named("create_payload_index")} == {"data_norm", "linked_memory_ids"}
        assert "clean" in capsys.readouterr().out.lower()

    def test_apply_refuses_a_plan_that_deletes_more_than_max_delete(self, monkeypatch, capsys):
        client = FakeQdrantClient()
        rows = [row(f"e{i}", f"Entity {i}", ["gone"], norm=f"entity {i}") for i in range(3)]
        monkeypatch.setattr(em, "_connect", lambda: (client, "ents", "mems"))
        monkeypatch.setattr(em, "_load", lambda c, entities, memories: (rows, {"m1"}))
        assert em.main(["--apply", "--max-delete", "2"]) == 3
        assert "refusing" in capsys.readouterr().out.lower()
        assert client.calls == []

    def test_apply_refuses_when_the_memories_collection_is_empty(self, monkeypatch, capsys):
        """An empty memories collection makes every link look stale, so the plan would delete every entity."""
        client = FakeQdrantClient()
        rows = [row("a", "Keep", ["m1"], norm="keep")]
        monkeypatch.setattr(em, "_connect", lambda: (client, "ents", "mems"))
        monkeypatch.setattr(em, "_load", lambda c, entities, memories: (rows, set()))
        assert em.main(["--apply"]) == 3
        assert "refusing" in capsys.readouterr().out.lower()
        assert client.calls == []
