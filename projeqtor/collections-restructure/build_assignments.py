#!/usr/bin/env python3
"""Generate an Assignment import file: every Banking resource x every leaf task.

The `automaticAssignment` column on Activity is a dead end -- it sets the toggle
without generating the assignment rows (see README.md). Assignment is its own
element type in the import dropdown, which is a different code path: it creates
real assignment records rather than relying on a UI side effect.

    "To assign more than one resource, you have to add as many lines for each
     resource."  -- projeqtor.org forum

So the file is one row per (resource, activity): 25 resources x 90 leaves = 2250.

COLUMNS. Taken from the "?" button on the Import Data screen with Assignment
selected, which prints the object's real schema:

    id             int(12)        id
    idProject      int(12)        project
    refType        varchar(100)   element        <- class name, e.g. Activity
    refId          int(12)        element id     <- the activity id
    idResource     int(12)        resource
    uniqueResource int(1)         unique resource
    idRole         int(12)        function
    rate           int(3)         rate (%)
    capacity       decimal(5,2)   capacity (FTE)
    comment        varchar(4000)  comments
    idle           int(1)         closed
    assignedWork   decimal(12,5)  assigned work
    leftWork       decimal(12,5)  left work

Note there is NO idActivity. An assignment can hang off an Activity, a Ticket, a
Meeting and so on, so the target is the generic pair refType + refId. A first
attempt using `idResource;idActivity;rate` -- guessed from a forum snippet --
failed with a bare "An error occurred", which is what the one-row smoke file is
for.

Usage
-----
    python3 build_assignments.py --resources EXPORT_Resource.csv --from-id 554

Roster, one of:
  --resources    Resource export, filtered to --team. Keys by NUMERIC id, which
                 sidesteps the double spaces in several real names.
  --allocations  Allocations export, scoped to --roster-project.
  --assignments  Assignment export, scoped to the rows for --roster-activity.

Leaves, one of:
  --activities   Activity export; may be unfiltered, leaf resolution already
                 filters on project Collections plus a 2.2 WBS prefix.
  --leaf-ids     explicit range, e.g. 545-634.
  --from-id      skip leaves below this id, to resume after manual toggling.
"""

from __future__ import annotations

import argparse
import csv
import io
import os
import sys

import build

DELIMITER = ";"
OUTPUT_ENCODING = "cp1252"

# The subset of the schema above that an assignment actually needs. Columns with
# no data are simply not updated, so the optional ones are left out rather than
# written blank.
COLUMNS = ["idProject", "refType", "refId", "idResource", "rate"]

# refType carries the class name of the thing being assigned to. The activity
# export is "export_Activity_<date>.csv" and Activity is the element type these
# rows were imported as, so Activity is the class -- Task is its *type*, held in
# the activity's own `activity type` column, not a class of its own.
REF_TYPE = "Activity"
DEFAULT_RATE = 100

RESOURCE_COLUMNS = ("idResource", "resource", "resource name", "name")
ACTIVITY_COLUMNS = ("refId", "idActivity", "activity", "element id")


def read_csv(path: str) -> tuple[list[str], list[dict]]:
    for enc in ("cp1252", "latin-1", "utf-8-sig"):
        try:
            with open(path, encoding=enc, newline="") as fh:
                reader = csv.DictReader(fh, delimiter=DELIMITER)
                rows = list(reader)
                if reader.fieldnames and len(reader.fieldnames) > 1:
                    return [f for f in reader.fieldnames if f], rows
        except UnicodeDecodeError:
            continue
    sys.exit(f"ERROR: could not parse {path} as a ProjeQtor CSV export "
             f"(expected ';' delimited with a header line).")


def pick(header: list[str], candidates: tuple[str, ...], what: str) -> str:
    lowered = {h.strip().lower(): h for h in header}
    for candidate in candidates:
        if candidate.lower() in lowered:
            return lowered[candidate.lower()]
    sys.exit(f"ERROR: no {what} column found.\n"
             f"  looked for: {', '.join(candidates)}\n"
             f"  header was: {', '.join(header)}")


def resolve_leaves(args) -> list[int]:
    target = sum(len(t) for _, ss in build.TREE for _, t in ss)
    if args.activities:
        ids = build.resolve_ids(read_csv(args.activities)[1])
        leaves = sorted(v for k, v in ids.items() if len(k) == 3)
        if len(leaves) != target:
            missing = [" > ".join(k) for c, ss in build.TREE for s, ts in ss
                       for t in ts if (k := (c, s, t)) not in ids]
            sys.exit(f"ERROR: found {len(leaves)} of {target} leaf tasks in the "
                     f"activity export. Import 03_create_tasks.csv first.\n"
                     f"  missing: " + "; ".join(missing[:5])
                     + (" ..." if len(missing) > 5 else ""))
    elif args.leaf_ids:
        lo, _, hi = args.leaf_ids.partition("-")
        leaves = list(range(int(lo), int(hi) + 1))
        if len(leaves) != target:
            sys.exit(f"ERROR: --leaf-ids {args.leaf_ids} spans {len(leaves)} ids "
                     f"but the tree has {target} leaves.")
    else:
        sys.exit("ERROR: pass --activities or --leaf-ids.")
    return leaves


def resolve_roster(args) -> tuple[list[str], str]:
    """Return (resource ids or names, a description of where they came from)."""
    if args.resources:
        # Keyed by NUMERIC id. Nine Banking names carry double spaces ("Alba
        # Fernandez Lopez"), which HTML collapses on screen, so matching by name
        # would look right and silently miss those people.
        header, source = read_csv(args.resources)
        id_col = pick(header, ("id",), "id")
        team_col = pick(header, ("team", "idTeam"), "team")
        name_col = pick(header, ("real name", "name"), "name")
        closed_col = next((h for h in header if h.strip().lower() == "closed"), None)
        excluded = {e.strip() for e in args.exclude.split(",") if e.strip()}

        roster, dropped = [], []
        for row in source:
            if (row.get(team_col) or "").strip() != args.team:
                continue
            rid = (row.get(id_col) or "").strip()
            name = " ".join((row.get(name_col) or "").split())
            if rid in excluded:
                dropped.append(f"{rid} {name} (--exclude)")
            elif closed_col and (row.get(closed_col) or "").strip() == "1":
                dropped.append(f"{rid} {name} (closed)")
            else:
                roster.append(rid)
        if not roster:
            sys.exit(f"ERROR: no resources on team '{args.team}'. Values seen: "
                     + ", ".join(sorted({(r.get(team_col) or '').strip()
                                         for r in source})[:10]))
        note = f"team '{args.team}'" + (
            f", minus {len(dropped)} [{'; '.join(dropped)}]" if dropped else "")
        return roster, note

    if args.allocations:
        header, source = read_csv(args.allocations)
        res_col = pick(header, RESOURCE_COLUMNS, "resource")
        proj_col = pick(header, ("idProject", "project"), "project")
        scoped = [r for r in source
                  if (r.get(proj_col) or "").strip() == str(args.roster_project)]
        if not scoped:
            sys.exit(f"ERROR: no allocations for project {args.roster_project}.")
        return _distinct(scoped, res_col), f"allocations on project {args.roster_project}"

    if args.assignments:
        header, source = read_csv(args.assignments)
        res_col = pick(header, RESOURCE_COLUMNS, "resource")
        act_col = pick(header, ACTIVITY_COLUMNS, "activity")
        scoped = [r for r in source
                  if (r.get(act_col) or "").strip() == str(args.roster_activity)]
        note = f"assignments on activity #{args.roster_activity}"
        if not scoped:
            scoped, note = source, "ALL rows in the export (unscoped)"
        return _distinct(scoped, res_col), note

    sys.exit("ERROR: pass --resources, --allocations or --assignments.")


def _distinct(rows: list[dict], column: str) -> list[str]:
    out, seen = [], set()
    for row in rows:
        value = (row.get(column) or "").strip()
        if value and value not in seen:
            seen.add(value)
            out.append(value)
    return out


def write_csv(path: str, columns: list[str], rows: list[dict]) -> None:
    buf = io.StringIO(newline="")
    writer = csv.DictWriter(buf, fieldnames=columns, delimiter=DELIMITER,
                            quoting=csv.QUOTE_MINIMAL, lineterminator="\r\n")
    writer.writeheader()
    writer.writerows(rows)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "wb") as fh:
        fh.write(buf.getvalue().encode(OUTPUT_ENCODING, errors="replace"))


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--resources", metavar="CSV")
    ap.add_argument("--allocations", metavar="CSV")
    ap.add_argument("--assignments", metavar="CSV")
    ap.add_argument("--team", default="Banking")
    ap.add_argument("--exclude", default="211",
                    help="resource ids to leave out (default 211 = TEST DUMMY)")
    ap.add_argument("--roster-project", type=int, default=build.ID_PROJECT_COLLECTIONS)
    ap.add_argument("--roster-activity", type=int, default=80)
    ap.add_argument("--activities", metavar="CSV")
    ap.add_argument("--leaf-ids", metavar="LO-HI")
    ap.add_argument("--from-id", type=int, default=0)
    ap.add_argument("--id-role", type=int,
                    help="idRole (function) to write on every row. Omit unless the "
                         "import complains that it is required -- the '?' screen "
                         "lists it as int(12), so it needs the numeric role id.")
    ap.add_argument("--rate", type=int, default=DEFAULT_RATE)
    ap.add_argument("--out", default=os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "out",
        "05_assign_team_to_leaves.csv"))
    args = ap.parse_args()

    leaves = resolve_leaves(args)
    done = [i for i in leaves if i < args.from_id]
    leaves = [i for i in leaves if i >= args.from_id]
    if not leaves:
        sys.exit(f"ERROR: --from-id {args.from_id} leaves nothing to assign.")

    roster, note = resolve_roster(args)

    columns = list(COLUMNS)
    if args.id_role is not None:
        columns.insert(columns.index("rate"), "idRole")
    rows = [{"idProject": build.ID_PROJECT_COLLECTIONS,
             "refType": REF_TYPE,
             "refId": aid,
             "idResource": person,
             **({"idRole": args.id_role} if args.id_role is not None else {}),
             "rate": args.rate}
            for aid in leaves for person in roster]

    write_csv(args.out, columns, rows)
    smoke = args.out.replace(".csv", "_SMOKE_1row.csv")
    write_csv(smoke, columns, rows[:1])

    print(f"columns:  {columns}   (refType={REF_TYPE!r})")
    print(f"roster:   {len(roster)} from {note}")
    if done:
        print(f"skipped:  {len(done)} leaves already done by hand "
              f"(ids {min(done)}-{max(done)})")
    print(f"leaves:   {len(leaves)}   ids {min(leaves)}-{max(leaves)}")
    print(f"\nwrote {smoke}  [1 row]      <- import this FIRST")
    print(f"wrote {args.out}  [{len(rows)} rows]")


if __name__ == "__main__":
    main()
