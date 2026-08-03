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
                 and the roster: every resource currently assigned to a
                 Collections activity. Filter it to Collections before exporting
                 so the roster is the Banking team and not the whole company.
  --activities   export from the Activities list, filtered to project Collections.
                 Supplies the 90 leaf ids.

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
    ap.add_argument("--assignments", required=True, metavar="CSV")
    ap.add_argument("--activities", required=True, metavar="CSV")
    ap.add_argument("--out", default=os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "out",
        "05_assign_team_to_leaves.csv"))
    args = ap.parse_args()

    # --- the 90 leaves, straight out of the activity export
    ids = build.resolve_ids(read_csv(args.activities)[1])
    leaves = {k: v for k, v in ids.items() if len(k) == 3}
    target = sum(len(t) for _, ss in build.TREE for _, t in ss)
    if len(leaves) != target:
        missing = [" > ".join(k) for c, ss in build.TREE for s, ts in ss
                   for t in ts if (k := (c, s, t)) not in leaves]
        sys.exit(f"ERROR: found {len(leaves)} of {target} leaf tasks in the "
                 f"activity export. Import 03_create_tasks.csv first.\n  missing: "
                 + "; ".join(missing[:5]) + (" ..." if len(missing) > 5 else ""))

    # --- the roster, straight out of the assignment export
    header, assignments = read_csv(args.assignments)
    res_col = pick(header, RESOURCE_COLUMNS, "resource")
    act_col = pick(header, ACTIVITY_COLUMNS, "activity")
    rate_col = next((h for h in header
                     if h.strip().lower() in {c.lower() for c in RATE_COLUMNS}), None)

    resources, seen = [], set()
    for row in assignments:
        value = (row.get(res_col) or "").strip()
        if value and value not in seen:
            seen.add(value)
            resources.append(value)
    if not resources:
        sys.exit(f"ERROR: no resources found in column '{res_col}' of "
                 f"{args.assignments}.")

    rates = {(r.get(rate_col) or "").strip() for r in assignments} if rate_col else set()
    rate = rates.pop() if len(rates) == 1 else "100"

    columns = [res_col, act_col] + ([rate_col] if rate_col else [])
    rows = [{res_col: person, act_col: aid, **({rate_col: rate} if rate_col else {})}
            for aid in sorted(leaves.values())
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
    print(f"resources: {len(resources)}  ({', '.join(resources[:3])}, ...)")
    print(f"leaves:    {len(leaves)}   ids {min(leaves.values())}-{max(leaves.values())}")
    print(f"\nwrote {smoke}  [1 row]      <- import this FIRST")
    print(f"wrote {args.out}  [{len(rows)} rows]")
    if not rate_col:
        print("\nNOTE: no rate column in the export, so none is written. If the "
              "import complains, add one named as the '?' screen shows.")


if __name__ == "__main__":
    main()
