#!/usr/bin/env python3
"""Assign the Banking team to the activities the Payment restructure created.

Same contract as the Collections side, proven on this instance:

    idProject;refType;refId;idResource;rate     with refType = Activity

One difference that matters. Collections was a single project, so every row
carried `idProject = 14`. Payment's units are separate sub-projects, so idProject
is per row -- PS is 31, IB is 35, and so on -- taken from the activity's own
project rather than assumed.

WHICH ACTIVITIES. Only the ones this restructure created, and only the leaves:

  yes  25 new flat tasks (the 5 added to PS, BE, IT, DE, PMS-TMS)
  yes  66 new tasks under the 6 systems
  no   the 6 system activities themselves -- group rows. The Collections work
       established that a group row still renders editable cells on the
       timesheet, so assigning one invites time landing on it instead of
       rolling up from its children.
  no   the 48 tasks that predate this work, including the 18 under IB/UK/FR --
       they already carry their own assignments, and touching them here would
       duplicate.

The roster comes from the Collections kit rather than a second copy: it keys by
numeric resource id because nine Banking names contain double spaces that HTML
collapses on screen, so name matching looks right and silently misses people.

Usage
-----
    python3 build_assignments.py --resources EXPORT_Resource.csv \\
                                 --activities EXPORT_Activity.csv \\
                                 [--split-projects]
"""

from __future__ import annotations

import argparse
import csv
import io
import os
import sys
import types

import build

# The roster logic lives in the sibling kit. Importing it keeps one definition of
# "who is on the Banking team" instead of two that can drift apart.
_COLLECTIONS = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                            "..", "collections-restructure")
sys.path.insert(0, _COLLECTIONS)
import build_assignments as collections_assignments  # noqa: E402

DELIMITER = ";"
OUTPUT_ENCODING = "cp1252"
COLUMNS = ["idProject", "refType", "refId", "idResource", "rate"]
REF_TYPE = "Activity"
DEFAULT_RATE = 100


def read_export(path: str) -> list[dict]:
    for enc in ("cp1252", "latin-1", "utf-8-sig"):
        try:
            with open(path, encoding=enc, newline="") as fh:
                reader = csv.DictReader(fh, delimiter=DELIMITER)
                rows = list(reader)
                if reader.fieldnames and len(reader.fieldnames) > 1:
                    return rows
        except UnicodeDecodeError:
            continue
    sys.exit(f"ERROR: could not parse {path} as a ProjeQtor CSV export.")


def resolve_targets(rows: list[dict]) -> list[tuple[str, int, str]]:
    """(activity id, idProject, project name) for every leaf needing assignment."""
    baseline = {str(v) for p in build.EXISTING for v in build.EXISTING[p].values()}
    # Under the re-parent plan these 18 are deleted: each duplicates an older task
    # moving under P02, which already carries its own assignments. Assigning a row
    # that is about to be deleted wastes 25 rows apiece and leaves orphans behind
    # if the deletion happens after the import.
    doomed = {new for new, _c, _n, _old in build.p02_duplicates(rows)}
    out = []
    for r in rows:
        wbs = (r.get("wbs") or "").strip()
        project = build.canon((r.get("project") or "").strip())
        if not wbs.startswith("2.1.") or project not in build.PAYMENT_PROJECTS:
            continue
        depth = len(wbs.split("."))
        name = (r.get("name") or "").strip()
        if depth == 4 and name in build.SYSTEMS.get(project, []):
            continue                      # group row
        if depth == 4 and r["id"] in baseline:
            continue                      # predates this work, already assigned
        if r["id"] in doomed:
            continue                      # duplicate, being deleted
        if depth not in (4, 5):
            continue
        out.append((r["id"], build.PAYMENT_PROJECTS[project], project))
    return sorted(out, key=lambda t: int(t[0]))


def write_csv(path: str, rows: list[dict]) -> None:
    buf = io.StringIO(newline="")
    writer = csv.DictWriter(buf, fieldnames=COLUMNS, delimiter=DELIMITER,
                            quoting=csv.QUOTE_MINIMAL, lineterminator="\r\n")
    writer.writeheader()
    writer.writerows(rows)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "wb") as fh:
        fh.write(buf.getvalue().encode(OUTPUT_ENCODING, errors="replace"))


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--resources", metavar="CSV", required=True)
    ap.add_argument("--activities", metavar="CSV", required=True)
    ap.add_argument("--team", default="Banking")
    ap.add_argument("--exclude", default="211",
                    help="resource ids to leave out (default 211 = TEST DUMMY)")
    ap.add_argument("--rate", type=int, default=DEFAULT_RATE)
    ap.add_argument("--split-projects", action="store_true",
                    help="one file per sub-project rather than one big one. The "
                         "Collections import survived 2024 rows but only by "
                         "committing before the response died.")
    ap.add_argument("--out", default=os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "out",
        "06_assign_team_to_new_activities.csv"))
    args = ap.parse_args()

    targets = resolve_targets(read_export(args.activities))
    if not targets:
        sys.exit("ERROR: no activities to assign -- is the export current?")

    roster, note = collections_assignments.resolve_roster(types.SimpleNamespace(
        resources=args.resources, allocations=None, assignments=None,
        team=args.team, exclude=args.exclude))

    rows = [{"idProject": pid, "refType": REF_TYPE, "refId": aid,
             "idResource": person, "rate": args.rate}
            for aid, pid, _project in targets for person in roster]

    print(f"columns:  {COLUMNS}   (refType={REF_TYPE!r})")
    print(f"roster:   {len(roster)} from {note}")
    print(f"targets:  {len(targets)} activities x {len(roster)} = {len(rows)} rows")

    if args.split_projects:
        base, ext = os.path.splitext(args.out)
        for n, project in enumerate(build.PAYMENT_PROJECTS, 1):
            part = [r for r in rows
                    if r["idProject"] == build.PAYMENT_PROJECTS[project]]
            if not part:
                continue
            slug = project.lower().replace(" ", "-")
            path = f"{base}_{n:02d}_{slug}{ext}"
            write_csv(path, part)
            n_act = len({r['refId'] for r in part})
            print(f"  wrote {os.path.basename(path):46s} {len(part):4d} rows  "
                  f"{n_act} activities")
        return

    # Smoke + REMAINDER, never smoke + full. These rows carry no `id`, so every
    # one INSERTs: importing a 1-row smoke and then a file that still contains
    # that row assigns the same person twice. That is exactly how PS ended up
    # with two `Payment Run Issues`, and how the Collections assignment file had
    # to be regenerated. The two files are disjoint by construction.
    smoke = args.out.replace(".csv", "_SMOKE_1row.csv")
    remainder = args.out.replace(".csv", "_REMAINDER.csv")
    write_csv(smoke, rows[:1])
    write_csv(remainder, rows[1:])
    print(f"\nwrote {smoke}      [1 row]           <- import FIRST")
    print(f"wrote {remainder}  [{len(rows) - 1} rows]  <- then this")
    print(f"together: {len(rows)} rows, no row in both.")


if __name__ == "__main__":
    main()
