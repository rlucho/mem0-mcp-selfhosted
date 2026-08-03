#!/usr/bin/env python3
"""Generate an Assignment import file: every Banking resource x every leaf task.

The `automaticAssignment` column on Activity is a dead end -- it sets the toggle
without generating the assignment rows (see README.md). Assignment is its own
element type in the import dropdown, though, which is a different code path: it
creates real assignment records rather than relying on a UI side effect.

    "To assign more than one resource, you have to add as many lines for each
     resource."  -- projeqtor.org forum

So the file is one row per (resource, activity): 25 resources x 90 leaves = 2250.

Rather than guess the column names, this script uses YOUR OWN assignment export as
the template -- it reads the header, works out which columns carry the resource
and the activity, and emits a file with those exact names. ProjeQtor's exports are
re-importable, so a file built from one is the safest possible shape.

Usage
-----
    python3 build_assignments.py \
        --assignments EXPORT_Assignment.csv \
        --activities  EXPORT_Activity.csv \
        [--out out/05_assign_team_to_leaves.csv]

  --assignments  export from the Assignments screen. Supplies the column layout
                 and the roster. It does NOT need to be filtered -- the roster is
                 taken from the rows pointing at one known Collections activity
                 (--roster-activity, default #80 Sap P02, which carries all 25),
                 so an export of the whole instance works fine.
  --activities   OPTIONAL export from the Activities list; also fine unfiltered,
                 since leaf resolution already filters on project Collections plus
                 a 2.2 WBS prefix. Omit it to use --leaf-ids instead.
  --leaf-ids     explicit id range for the 90 leaves, e.g. 545-634. Used when no
                 activity export is given.
  --from-id      skip leaves below this id, for resuming after manual toggling.
                 e.g. --from-id 554 after having done up to 553 by hand.

Import the result with element type Assignment. TEST WITH ONE ROW FIRST -- keep
the header plus a single line, confirm it comes back "inserted" and shows up on
the activity's Progress tab, then run the whole file.
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

# Candidate header names, in order of preference. ProjeQtor accepts either the
# object-class field name or the on-screen label, so an export may carry either.
RESOURCE_COLUMNS = ("idResource", "resource", "resource name", "name")
ACTIVITY_COLUMNS = ("idActivity", "activity", "activity name", "task")
RATE_COLUMNS = ("rate", "assignment rate", "%")


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
    sys.exit(f"ERROR: no {what} column in the assignment export.\n"
             f"  looked for: {', '.join(candidates)}\n"
             f"  header was: {', '.join(header)}\n"
             f"Add the right name to the *_COLUMNS tuple at the top of this file. "
             f"The '?' button on the import screen, with Assignment selected, "
             f"lists the valid column names.")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--assignments", metavar="CSV")
    ap.add_argument("--allocations", metavar="CSV",
                    help="fallback roster source when no Assignment list exists to "
                         "export: an Allocations export, scoped to --roster-project")
    ap.add_argument("--roster-project", type=int, default=build.ID_PROJECT_COLLECTIONS,
                    help="project id to scope the allocations roster to")
    ap.add_argument("--activities", metavar="CSV")
    ap.add_argument("--leaf-ids", metavar="LO-HI",
                    help="id range of the 90 leaves when --activities is omitted")
    ap.add_argument("--from-id", type=int, default=0,
                    help="skip leaves below this id (resume after manual toggling)")
    ap.add_argument("--roster-activity", type=int, default=80,
                    help="take the roster from assignments on this activity id")
    ap.add_argument("--out", default=os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "out",
        "05_assign_team_to_leaves.csv"))
    args = ap.parse_args()

    target = sum(len(t) for _, ss in build.TREE for _, t in ss)

    # --- the 90 leaves: from an activity export, or from an explicit id range
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

    done = [i for i in leaves if i < args.from_id]
    leaves = [i for i in leaves if i >= args.from_id]
    if not leaves:
        sys.exit(f"ERROR: --from-id {args.from_id} leaves nothing to assign.")

    # --- the roster and the column layout
    if args.assignments:
        # Preferred: the real Assignment export dictates the exact column names,
        # and ProjeQtor's own exports are re-importable.
        header, source = read_csv(args.assignments)
        res_col = pick(header, RESOURCE_COLUMNS, "resource")
        act_col = pick(header, ACTIVITY_COLUMNS, "activity")
        rate_col = next((h for h in header
                         if h.strip().lower() in {c.lower() for c in RATE_COLUMNS}),
                        None)
        # An unfiltered export covers the whole instance, so narrow to the rows for
        # one activity known to carry the full Banking team. Numeric activity
        # columns are matched by id; a name column cannot be trusted (other
        # projects reuse names like "Sap P02"), so fall back to every row and say so.
        scoped = [r for r in source
                  if (r.get(act_col) or "").strip() == str(args.roster_activity)]
        scope_note = f"assignments on activity #{args.roster_activity}"
        if not scoped:
            scoped, scope_note = source, "ALL rows in the export (unscoped)"
    elif args.allocations:
        # Fallback when there is no Assignment list to export from. Allocations are
        # resource-to-project, so scoping to project 14 gives exactly the Banking
        # team. Column names then come from the forum-documented trio rather than
        # from an export -- which is why the one-row smoke file matters here.
        header, source = read_csv(args.allocations)
        res_col = pick(header, RESOURCE_COLUMNS, "resource")
        proj_col = pick(header, ("idProject", "project"), "project")
        act_col, rate_col = "idActivity", "rate"
        scoped = [r for r in source
                  if (r.get(proj_col) or "").strip() == str(args.roster_project)]
        scope_note = f"allocations on project {args.roster_project}"
        if not scoped:
            sys.exit(f"ERROR: no allocations for project {args.roster_project} in "
                     f"column '{proj_col}'. Values seen: "
                     + ", ".join(sorted({(r.get(proj_col) or '').strip()
                                         for r in source})[:10]))
        res_col = "idResource" if res_col.lower() == "idresource" else res_col
    else:
        sys.exit("ERROR: pass --assignments (preferred) or --allocations.")

    resources, seen = [], set()
    for row in scoped:
        value = (row.get(res_col) or "").strip()
        if value and value not in seen:
            seen.add(value)
            resources.append(value)
    if not resources:
        sys.exit(f"ERROR: no resources found in column '{res_col}' of "
                 f"{args.assignments}.")

    rates = {v for r in scoped if (v := (r.get(rate_col) or "").strip())} \
        if rate_col else set()
    rate = rates.pop() if len(rates) == 1 else "100"

    columns = [res_col, act_col] + ([rate_col] if rate_col else [])
    rows = [{res_col: person, act_col: aid, **({rate_col: rate} if rate_col else {})}
            for aid in leaves
            for person in resources]

    buf = io.StringIO(newline="")
    writer = csv.DictWriter(buf, fieldnames=columns, delimiter=DELIMITER,
                            quoting=csv.QUOTE_MINIMAL, lineterminator="\r\n")
    writer.writeheader()
    writer.writerows(rows)
    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    with open(args.out, "wb") as fh:
        fh.write(buf.getvalue().encode(OUTPUT_ENCODING, errors="replace"))

    # a one-row twin, so the format can be proven before the full run
    smoke = args.out.replace(".csv", "_SMOKE_1row.csv")
    with open(smoke, "wb") as fh:
        fh.write("\r\n".join([DELIMITER.join(columns),
                              DELIMITER.join(str(rows[0][c]) for c in columns), ""])
                 .encode(OUTPUT_ENCODING, errors="replace"))

    print(f"columns taken from your export: {columns}")
    print(f"roster:    {len(resources)} from {scope_note}"
          f"  ({', '.join(resources[:3])}, ...)")
    if done:
        print(f"skipped:   {len(done)} leaves already done by hand "
              f"(ids {min(done)}-{max(done)})")
    print(f"leaves:    {len(leaves)}   ids {min(leaves)}-{max(leaves)}")
    print(f"\nwrote {smoke}  [1 row]      <- import this FIRST")
    print(f"wrote {args.out}  [{len(rows)} rows]")
    if not rate_col:
        print("\nNOTE: no rate column in the export, so none is written. If the "
              "import complains, add one named as the '?' screen shows.")


if __name__ == "__main__":
    main()
