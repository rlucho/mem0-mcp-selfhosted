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


def rows_countries(mode: str) -> list[dict]:
    return [_create_row(mode, name, parent="") for name, _ in TREE]


def rows_systems(mode: str, ids: dict) -> list[dict]:
    out = []
    for country, systems in TREE:
        parent = ids.get((country,), placeholder(country))
        for system, _ in systems:
            out.append(_create_row(mode, system, parent=parent))
    return out


def rows_tasks(mode: str, ids: dict) -> list[dict]:
    out = []
    for country, systems in TREE:
        for system, tasks in systems:
            parent = ids.get((country, system), placeholder(country, system))
            for task in tasks:
                out.append(_create_row(mode, task, parent=parent))
    return out


def _create_row(mode: str, name: str, parent) -> dict:
    if mode == "ids":
        return {
            "name": name,
            "idProject": ID_PROJECT_COLLECTIONS,
            "idActivity": parent,
            "idActivityType": _require(ID_ACTIVITY_TYPE_TASK, "ID_ACTIVITY_TYPE_TASK"),
            "idStatus": _require(ID_STATUS_NEW, "ID_STATUS_NEW"),
            "idActivityPlanningMode": _require(
                ID_ACTIVITY_PLANNING_MODE_ASAP, "ID_ACTIVITY_PLANNING_MODE_ASAP"),
        }
    return {
        "name": name,
        "idProject": ID_PROJECT_COLLECTIONS,
        "idActivity": parent,
        "activity type": LABEL_ACTIVITY_TYPE,
        "status": LABEL_STATUS_NEW,
        "planning mode": LABEL_PLANNING_MODE,
    }


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
    """Map ('Italia',) -> id and ('Italia', 'SAP E01') -> id.

    Only rows inside Collections (id 14) count: project name 'Collections' AND a
    wbs under 2.2. That is what keeps Q2C's Collections (id 17, wbs 3.3) out.
    """
    scoped = [r for r in export_rows
              if (r.get("project") or "").strip() == "Collections"
              and (r.get("wbs") or "").strip().startswith("2.2")]

    countries = {c for c, _ in TREE}
    ids: dict = {}
    dupes: list[str] = []

    for row in scoped:
        name = (row.get("name") or "").strip()
        parent = (row.get("parent activity") or "").strip()
        try:
            rid = int((row.get("id") or "").strip())
        except ValueError:
            continue

        key = None
        if not parent and name in countries:
            key = (name,)
        elif parent in countries:
            systems = dict(TREE)[parent]
            if name in {s for s, _ in systems}:
                key = (parent, name)

        if key is None:
            continue
        if key in ids and ids[key] != rid:
            dupes.append(f"{' > '.join(key)}: ids {ids[key]} and {rid}")
        ids.setdefault(key, rid)

    if dupes:
        sys.exit("ERROR: ambiguous parents in the export -- fix in ProjeQtor "
                 "before continuing:\n  " + "\n  ".join(dupes))
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

    cols_create = COLS_IDS_CREATE if args.mode == "ids" else COLS_LABELS_CREATE
    cols_close = COLS_IDS_CLOSE if args.mode == "ids" else COLS_LABELS_CLOSE

    countries = rows_countries(args.mode)
    systems = rows_systems(args.mode, ids)
    tasks = rows_tasks(args.mode, ids)
    close = rows_close(args.mode)

    files = [
        ("00_smoke_test_one_activity.csv", cols_create, countries[:1]),
        ("01_create_countries.csv", cols_create, countries),
        ("02_create_systems.csv", cols_create, systems),
        ("03_create_tasks.csv", cols_create, tasks),
        ("04_close_old_collections.csv", cols_close, close),
    ]
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
    print(f"\nmode={args.mode}  structure={len(countries)} countries + "
          f"{len(systems)} systems + {len(tasks)} tasks = {total} activities")
    print(f"closing {len(close)} existing activities in project "
          f"{ID_PROJECT_COLLECTIONS} (Banking > Collections)")
    if not ids:
        print("\nNo --export given: files 02/03 carry <<placeholders>> in idActivity "
              "and are NOT importable yet. Import 01, export Activities, then re-run "
              "with --export to fill them in.")


if __name__ == "__main__":
    main()
