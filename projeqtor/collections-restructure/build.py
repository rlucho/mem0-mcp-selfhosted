#!/usr/bin/env python3
"""Generate the ProjeQtor CSV import files for the Collections (id 14) restructure.

Target shape (Banking > Collections is the PROJECT; everything below is ACTIVITIES):

    Banking
      Collections            <- project id 14, wbs 2.2
        Italia               <- activity  (level 1)
          SAP E01            <- activity  (level 2, parent = Italia)
            Autobank         <- activity  (level 3, parent = SAP E01)
            Cheque
            Riba
            ...

ProjeQtor recalculates WBS on import, so hierarchy CANNOT be expressed by a wbs
column. It is expressed by the parent column (`idActivity`), which needs the
parent's NUMERIC id. Those ids only exist after the parent rows are imported, so
creation is a 3-pass process:

    pass 1 -> 01_create_countries.csv   (5 rows, no parent)
    export Activities from ProjeQtor
    pass 2 -> 02_create_systems.csv     (14 rows, parent = country id)
    export Activities from ProjeQtor
    pass 3 -> 03_create_tasks.csv       (90 rows, parent = system id)

Re-run this script with --export <fresh ProjeQtor activity export.csv> after each
pass and it fills the real parent ids in for the next one.

Usage
-----
    python3 build.py                          # emit everything (parents as placeholders)
    python3 build.py --export ACT.csv         # resolve parent ids from a fresh export
    python3 build.py --mode ids               # technical field names + numeric ref ids
"""

from __future__ import annotations

import argparse
import csv
import io
import os
import sys

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

# Banking > Collections. NOT Q2C > Collections (that one is project id 17).
ID_PROJECT_COLLECTIONS = 14

# Values as ProjeQtor renders them in its own CSV export (used by --mode labels).
LABEL_ACTIVITY_TYPE = "Task"
LABEL_PLANNING_MODE = "as soon as possible"
LABEL_STATUS_NEW = "recorded"
LABEL_STATUS_CLOSED = "closed"  # VERIFY: exact name of the closing status

# Numeric reference ids (used by --mode ids). These are instance-specific and are
# NOT present in the activity export -- read them off the Administration screens
# before using this mode. See README.md, "Reference ids you must confirm".
ID_ACTIVITY_TYPE_TASK = None
ID_ACTIVITY_PLANNING_MODE_ASAP = None
ID_STATUS_NEW = None
ID_STATUS_CLOSED = None

# "automatic assignment of the project team" -- the per-activity toggle on the
# Progress tab. When on, every resource allocated to the project is assigned to
# THAT activity, and the assignment list re-syncs whenever an allocation changes.
# It applies to one activity only; it does not cascade to sub-activities. So it
# goes on the 90 leaf tasks (where time is booked) and stays off the countries and
# systems (which are grouping rows nobody should book against).
#
# VERIFY the column name against the "?" button next to `element type to import`
# with Activity selected -- it lists the valid column names for the object. Set to
# None to leave the column out entirely and toggle by hand.
AUTO_ASSIGN_COLUMN = "automaticAssignment"
AUTO_ASSIGN_LEVELS = {"task"}          # of: country, system, task

# Activity #524 (Italia), created by the smoke test. Used to emit a one-row file
# that flips the toggle on an EXISTING record, so the column name can be proven
# without creating a throwaway activity to clean up afterwards.
SMOKE_TEST_ACTIVITY_ID = 524

# Setting the flag on an EXISTING activity by import (00b) turned the toggle on
# but left the assignment table empty -- the import writes the column without
# running whatever populates assignments. 00c tests the other path: a row CREATED
# with the flag already set. Throwaway activity, delete it once checked.
TEST_ACTIVITY_NAME = "ZZ TEST auto-assign (delete me)"

# Written to match the encoding of the ProjeQtor export (Espana carries an enye).
OUTPUT_ENCODING = "cp1252"
DELIMITER = ";"

# ---------------------------------------------------------------------------
# The target structure  (decisions 1A / 2A / 3A)
#   1A flat: Country > System > Task, three levels, no nesting under Emails
#   2A normalised spellings (see NORMALISATIONS below)
#   3A plain system names, no country prefix
# ---------------------------------------------------------------------------

TREE: list[tuple[str, list[tuple[str, list[str]]]]] = [
    ("Italia", [
        ("SAP E01", ["Autobank", "Cheque", "Riba", "Emails",
                     "Conciliaciones", "Reportes", "Reuniones"]),
        ("SAP PP2", ["FEBAN", "Riba", "Emails", "Reuniones"]),
    ]),
    ("Francia", [
        ("SAP E01", ["Autobank", "Cheque/BOE", "Emails",
                     "Conciliaciones", "Reportes", "Reuniones"]),
        ("SAP PER", ["FEBAN", "Emails", "Conciliaciones", "Reuniones"]),
        ("SAP P02", ["FEBAN", "Cheque", "DD", "Emails",
                     "Conciliaciones", "Reuniones"]),
        ("Navision", ["FEBAN", "Conciliaciones", "Reuniones"]),
    ]),
    ("España", [
        ("SAP E01", ["FEBAN", "Cheque", "DD", "Emails",
                     "Conciliaciones", "Reportes", "Reuniones"]),
        ("SAP PER", ["FEBAN", "Cheque", "DD", "Emails",
                     "Conciliaciones", "Reportes", "Reuniones"]),
        ("SAP P02", ["FEBAN", "Cheque", "DD", "Emails",
                     "Conciliaciones", "Reportes", "Reuniones"]),
    ]),
    ("Portugal", [
        ("SAP P02", ["FEBAN", "DD", "Emails",
                     "Conciliaciones", "Reportes", "Reuniones"]),
        ("SAP PER", ["FEBAN", "Emails", "Conciliaciones",
                     "Reportes", "Reuniones"]),
        ("SAP GP1", ["FEBAN", "Emails", "Conciliaciones", "Reuniones"]),
    ]),
    ("Marruecos", [
        ("SAP E01", ["Autobank", "Cheque/BOE/Encaissement", "Emails",
                     "Conciliaciones", "Reportes", "Reuniones"]),
        ("SAP PER", ["FEBAN", "Emails", "Conciliaciones", "Reuniones"]),
    ]),
]

# Tasks that exist under EVERY system, on top of what Time_Pro.md listed.
# Appended after each system's own tasks; a system that already names one keeps
# its original position.
COMMON_TASKS = ["Training"]

TREE = [(country, [(system, tasks + [t for t in COMMON_TASKS if t not in tasks])
                   for system, tasks in systems])
        for country, systems in TREE]

# Decision 2A. Cells where the source Time_Pro.md differed from the spelling used
# above -- kept here so the change is auditable and reversible.
NORMALISATIONS = [
    ("Francia", "SAP PER", "Email", "Emails"),
    ("Francia", "SAP P02", "Cheques", "Cheque"),
    ("Francia", "SAP P02", "Email", "Emails"),
    ("España", "SAP E01", "Reporte", "Reportes"),
    ("España", "SAP P02", "Cheques", "Cheque"),
    ("España", "SAP P02", "Email", "Emails"),
]

# ---------------------------------------------------------------------------
# The 11 activities currently sitting directly under Collections (id 14).
# Taken from export_Activity_20260803_143423.csv, wbs 2.2.1 - 2.2.11.
# Ordered by DESCENDING id: the community import guide notes that closing a
# parent before its children errors out, and descending id closes leaves first.
# ---------------------------------------------------------------------------

CURRENT_COLLECTIONS_ACTIVITIES = [
    (521, "-TAS-528", "Mailbox", "2.2.11"),
    (520, "-TAS-527", "Riba", "2.2.10"),
    (519, "-TAS-526", "SAP E01", "2.2.9"),
    (356, "-TAS-363", "Project - Robotic", "2.2.8"),
    (177, "-TAS-184", "Direct Debit", "2.2.7"),
    (85, "-TAS-92", "Check process", "2.2.6"),
    (84, "-TAS-91", "Navision FR", "2.2.5"),
    (83, "-TAS-90", "Sap GP1", "2.2.4"),
    (82, "-TAS-89", "Sap PER", "2.2.3"),
    (81, "-TAS-88", "Sap PP2 IT", "2.2.2"),
    (80, "-TAS-87", "Sap P02", "2.2.1"),
]

# ---------------------------------------------------------------------------
# Column sets
# ---------------------------------------------------------------------------

# Technical DB field names. Confirmed against the community MSProject->ProjeQtor
# import guide, which ships a working activity import file using exactly these.
COLS_IDS_CREATE = ["name", "idProject", "idActivity",
                   "idActivityType", "idStatus", "idActivityPlanningMode"]
COLS_IDS_CLOSE = ["id", "name", "idStatus"]

# ProjeQtor's own export labels. The developer's guidance on the forum is that an
# exported file can be re-imported, so these are the round-trip-safe names.
COLS_LABELS_CREATE = ["name", "idProject", "idActivity",
                      "activity type", "status", "planning mode"]
COLS_LABELS_CLOSE = ["id", "name", "status"]


def placeholder(*parts: str) -> str:
    """Marker left in the parent column when no export has been supplied yet."""
    return "<<" + " > ".join(parts) + ">>"


def rows_countries(mode: str, ids: dict) -> list[dict]:
    """Countries not already present in the export."""
    return [_create_row(mode, name, parent="", level="country")
            for name, _ in TREE if (name,) not in ids]


def rows_systems(mode: str, ids: dict) -> list[dict]:
    out = []
    for country, systems in TREE:
        parent = ids.get((country,), placeholder(country))
        for system, _ in systems:
            if (country, system) not in ids:
                out.append(_create_row(mode, system, parent=parent, level="system"))
    return out


def rows_tasks(mode: str, ids: dict) -> list[dict]:
    out = []
    for country, systems in TREE:
        for system, tasks in systems:
            parent = ids.get((country, system), placeholder(country, system))
            for task in tasks:
                if (country, system, task) not in ids:
                    out.append(_create_row(mode, task, parent=parent, level="task"))
    return out


def _create_row(mode: str, name: str, parent, level: str) -> dict:
    if mode == "ids":
        row = {
            "name": name,
            "idProject": ID_PROJECT_COLLECTIONS,
            "idActivity": parent,
            "idActivityType": _require(ID_ACTIVITY_TYPE_TASK, "ID_ACTIVITY_TYPE_TASK"),
            "idStatus": _require(ID_STATUS_NEW, "ID_STATUS_NEW"),
            "idActivityPlanningMode": _require(
                ID_ACTIVITY_PLANNING_MODE_ASAP, "ID_ACTIVITY_PLANNING_MODE_ASAP"),
        }
    else:
        row = {
            "name": name,
            "idProject": ID_PROJECT_COLLECTIONS,
            "idActivity": parent,
            "activity type": LABEL_ACTIVITY_TYPE,
            "status": LABEL_STATUS_NEW,
            "planning mode": LABEL_PLANNING_MODE,
        }
    if AUTO_ASSIGN_COLUMN:
        row[AUTO_ASSIGN_COLUMN] = 1 if level in AUTO_ASSIGN_LEVELS else 0
    return row


def rows_close(mode: str) -> list[dict]:
    key, value = (("idStatus", _require(ID_STATUS_CLOSED, "ID_STATUS_CLOSED"))
                  if mode == "ids" else ("status", LABEL_STATUS_CLOSED))
    return [{"id": aid, "name": name, key: value}
            for aid, _ref, name, _wbs in CURRENT_COLLECTIONS_ACTIVITIES]


def _require(value, label):
    if value is None:
        sys.exit(f"ERROR: --mode ids needs {label} set at the top of build.py. "
                 f"See README.md, 'Reference ids you must confirm'.")
    return value


# ---------------------------------------------------------------------------
# Resolving parent ids from a fresh ProjeQtor activity export
# ---------------------------------------------------------------------------

def read_export(path: str) -> list[dict]:
    for enc in ("cp1252", "latin-1", "utf-8-sig"):
        try:
            with open(path, encoding=enc, newline="") as fh:
                rows = list(csv.DictReader(fh, delimiter=DELIMITER))
            if rows and "name" in rows[0]:
                return rows
        except UnicodeDecodeError:
            continue
    sys.exit(f"ERROR: could not parse {path} as a ProjeQtor CSV export.")


def resolve_ids(export_rows: list[dict]) -> dict:
    """Map every already-imported node of TREE to its ProjeQtor id.

        ('Italia',)                        -> 524
        ('Italia', 'SAP E01')              -> 529
        ('Italia', 'SAP E01', 'Autobank')  -> 543

    Ancestry comes from the WBS, not from the `parent activity` name column:
    `SAP E01` appears under four countries, so a name alone cannot say which
    branch a row belongs to, whereas `2.2.12.1` unambiguously sits under
    `2.2.12`. Depth under Collections is the level: 3 = country, 4 = system,
    5 = task.

    Only rows inside Collections (id 14) count -- project name 'Collections'
    AND a wbs under 2.2 -- which is what keeps Q2C's Collections (id 17,
    wbs 3.3) out.

    Everything found here is treated as already done and is left out of the
    generated files, so each pass can be re-run safely without creating
    duplicates.
    """
    scoped = {}
    for row in export_rows:
        wbs = (row.get("wbs") or "").strip()
        if (row.get("project") or "").strip() != "Collections":
            continue
        if not wbs.startswith("2.2."):
            continue
        scoped[wbs] = row

    if not scoped:
        return {}

    countries = {c for c, _ in TREE}
    systems_of = {c: {s for s, _ in ss} for c, ss in TREE}
    tasks_of = {(c, s): set(t) for c, ss in TREE for s, t in ss}

    ids: dict = {}
    key_by_wbs: dict = {}
    dupes: list[str] = []

    def claim(wbs: str, key: tuple, row: dict) -> None:
        try:
            rid = int((row.get("id") or "").strip())
        except ValueError:
            return
        if key in ids and ids[key] != rid:
            dupes.append(f"{' > '.join(key)}: ids {ids[key]} and {rid}")
            return
        ids[key] = rid
        key_by_wbs[wbs] = key

    # Depth 3 -> country, 4 -> system, 5 -> task. Shallowest first, so each pass
    # can look its parent up in key_by_wbs.
    for depth in (3, 4, 5):
        for wbs, row in scoped.items():
            if wbs.count(".") + 1 != depth:
                continue
            name = (row.get("name") or "").strip()
            if depth == 3:
                if name in countries:
                    claim(wbs, (name,), row)
                continue
            parent_key = key_by_wbs.get(wbs.rsplit(".", 1)[0])
            if parent_key is None or len(parent_key) != depth - 3:
                continue
            if depth == 4 and name in systems_of[parent_key[0]]:
                claim(wbs, parent_key + (name,), row)
            elif depth == 5 and name in tasks_of.get(parent_key, ()):
                claim(wbs, parent_key + (name,), row)

    if dupes:
        sys.exit("ERROR: the same node exists twice in Collections -- delete the "
                 "duplicate in ProjeQtor before continuing:\n  " + "\n  ".join(dupes))
    return ids


# ---------------------------------------------------------------------------
# Writing
# ---------------------------------------------------------------------------

def write_csv(path: str, columns: list[str], rows: list[dict]) -> None:
    buf = io.StringIO(newline="")
    writer = csv.DictWriter(buf, fieldnames=columns, delimiter=DELIMITER,
                            quoting=csv.QUOTE_MINIMAL, lineterminator="\r\n",
                            extrasaction="ignore")
    writer.writeheader()
    writer.writerows(rows)
    with open(path, "wb") as fh:
        fh.write(buf.getvalue().encode(OUTPUT_ENCODING, errors="replace"))


def write_structure_md(path: str) -> None:
    """Regenerate structure.md so the docs can never drift from TREE."""
    n_c = len(TREE)
    n_s = sum(len(s) for _, s in TREE)
    n_t = sum(len(t) for _, s in TREE for _, t in s)

    out = [
        "# Target structure — Banking > Collections (project id 14, WBS 2.2)\n",
        "Generated by `build.py` from `TREE` — edit there, not here.\n",
        "Decisions applied: **1A** flat 3 levels · **2A** normalised spellings · "
        "**3A** plain system names.\n",
        f"`{n_c}` countries + `{n_s}` systems + `{n_t}` tasks = "
        f"**{n_c + n_s + n_t} activities**, all with `idProject = {ID_PROJECT_COLLECTIONS}`.\n",
        "```",
        "Banking",
        f"└── Collections                    (project id {ID_PROJECT_COLLECTIONS}, wbs 2.2)",
    ]
    for ci, (country, systems) in enumerate(TREE):
        c_last = ci == len(TREE) - 1
        out.append(f"    {'└──' if c_last else '├──'} {country}")
        c_pad = "    " + ("    " if c_last else "│   ")
        for si, (system, tasks) in enumerate(systems):
            s_last = si == len(systems) - 1
            out.append(f"{c_pad}{'└──' if s_last else '├──'} {system}")
            s_pad = c_pad + ("    " if s_last else "│   ")
            for ti, task in enumerate(tasks):
                out.append(f"{s_pad}{'└──' if ti == len(tasks) - 1 else '├──'} {task}")
    out.append("```\n")

    if COMMON_TASKS:
        out.append("## Tasks added to every system\n")
        out.append("Not in `Time_Pro.md` — added on request, once under each of the "
                   f"{n_s} systems.\n")
        out += [f"- `{t}`" for t in COMMON_TASKS]
        out.append("")

    out.append("## Spelling normalisations applied (decision 2A)\n")
    out.append("| Country | System | Time_Pro.md | Imported as |")
    out.append("|---|---|---|---|")
    out += [f"| {c} | {s} | `{a}` | `{b}` |" for c, s, a, b in NORMALISATIONS]
    out.append("\nTo go back to the verbatim `.md` spellings, edit `TREE` in "
               "`build.py` and re-run it.\n")

    out.append("## Activities being closed\n")
    out.append("The 11 activities currently sitting flat under Collections, ordered "
               "by descending id —")
    out.append("the order `04_close_old_collections.csv` uses.\n")
    out.append("| id | reference | name | current WBS |")
    out.append("|---|---|---|---|")
    out += [f"| {aid} | `{ref}` | {name} | {wbs} |"
            for aid, ref, name, wbs in CURRENT_COLLECTIONS_ACTIVITIES]
    out.append("")

    with open(path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(out))


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--mode", choices=["labels", "ids"], default="labels",
                    help="labels: ProjeQtor export column labels + names (default). "
                         "ids: technical DB field names + numeric reference ids.")
    ap.add_argument("--export", metavar="CSV",
                    help="fresh ProjeQtor activity export, used to fill parent ids")
    ap.add_argument("--out", default=os.path.join(os.path.dirname(__file__), "out"))
    args = ap.parse_args()

    ids = resolve_ids(read_export(args.export)) if args.export else {}
    os.makedirs(args.out, exist_ok=True)

    cols_create = list(COLS_IDS_CREATE if args.mode == "ids" else COLS_LABELS_CREATE)
    cols_close = COLS_IDS_CLOSE if args.mode == "ids" else COLS_LABELS_CLOSE
    if AUTO_ASSIGN_COLUMN:
        cols_create.append(AUTO_ASSIGN_COLUMN)

    countries = rows_countries(args.mode, ids)
    systems = rows_systems(args.mode, ids)
    tasks = rows_tasks(args.mode, ids)
    close = rows_close(args.mode)

    if ids:
        n_c = sum(1 for k in ids if len(k) == 1)
        n_s = sum(1 for k in ids if len(k) == 2)
        n_t = sum(1 for k in ids if len(k) == 3)
        print(f"found in the export, so left out of the files below: "
              f"{n_c} countries, {n_s} systems, {n_t} tasks\n")

    files = [
        ("00_smoke_test_one_activity.csv", cols_create,
         rows_countries(args.mode, {})[:1]),
        ("01_create_countries.csv", cols_create, countries),
        ("02_create_systems.csv", cols_create, systems),
        ("03_create_tasks.csv", cols_create, tasks),
        ("04_close_old_collections.csv", cols_close, close),
    ]
    if AUTO_ASSIGN_COLUMN and SMOKE_TEST_ACTIVITY_ID:
        files.insert(1, ("00b_test_auto_assign_column.csv",
                         ["id", AUTO_ASSIGN_COLUMN],
                         [{"id": SMOKE_TEST_ACTIVITY_ID, AUTO_ASSIGN_COLUMN: 1}]))
    if AUTO_ASSIGN_COLUMN:
        files.insert(2, ("00c_test_auto_assign_on_create.csv", cols_create,
                         [_create_row(args.mode, TEST_ACTIVITY_NAME,
                                      parent="", level="task")]))
    for filename, columns, rows in files:
        path = os.path.join(args.out, filename)
        write_csv(path, columns, rows)
        pending = sum(1 for r in rows if str(r.get("idActivity", "")).startswith("<<"))
        note = f"  ({pending} parent ids still unresolved)" if pending else ""
        print(f"wrote {path}  [{len(rows)} rows]{note}")

    structure_md = os.path.join(os.path.dirname(os.path.abspath(__file__)), "structure.md")
    write_structure_md(structure_md)
    print(f"wrote {structure_md}")

    total = len(countries) + len(systems) + len(tasks)
    target = len(TREE) + sum(len(ss) for _, ss in TREE) + \
        sum(len(t) for _, ss in TREE for _, t in ss)
    print(f"\nmode={args.mode}  still to create: {len(countries)} countries + "
          f"{len(systems)} systems + {len(tasks)} tasks = {total} activities "
          f"(target tree is {target})")
    print(f"closing {len(close)} existing activities in project "
          f"{ID_PROJECT_COLLECTIONS} (Banking > Collections)")
    if not ids:
        print("\nNo --export given: files 02/03 carry <<placeholders>> in idActivity "
              "and are NOT importable yet. Import 01, export Activities, then re-run "
              "with --export to fill them in.")


if __name__ == "__main__":
    main()
