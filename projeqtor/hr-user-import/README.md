# ProjeQtor HR user/resource import — macro fixes

`import_hr_users_Automated_v3.xlsm` generates the ProjeQtor import CSVs.
[`NewModuleMacro.bas`](NewModuleMacro.bas) is the corrected module;
[`NewModuleMacro.ORIGINAL.bas`](NewModuleMacro.ORIGINAL.bas) is what it replaces.

## The bug that mattered: "Update" did not update

The UPDATE file was `name;idTeam;capacity;maxWeeklyWork;idRole;idCalendarDefinition`
— **no `id` column**. ProjeQtor's rule is *row has an `id` → UPDATE; no `id` →
INSERT*, so every "update" run was creating a **second resource with the same
name**, not editing the first.

The same rule already had form here: it duplicated `Payment Run Issues` under PS,
and it turned one assignment import into 1825 duplicate rows.

## Five more, in order of how much they cost

**Names were blindly proper-cased.** `StrConv(vbProperCase)` mangled **11 of 214**
real names — `Silvia De la Fuente Peña` → `De La`, `Cristina Garcia-Rojo` →
`Garcia-rojo`. Since the old UPDATE matched *by name*, a mangled name matched
nothing. `SmartName()` now normalises only input arriving ALL CAPS or all
lowercase and leaves mixed-case spellings alone.

**Sending `name` on an update renames the resource.** With the id present the name
is not needed to find the record — and **55 of 214 names carry a double space**
(`Francisco  Manzanilla`) that HTML collapses on screen, so a retyped name looks
identical and is not. The update file no longer sends `name` at all.

**Allocations were keyed by name.** Now by `idResource` wherever the id is known.
Create rows have no id yet, so those still go by name — hence two files.

**A team change never removed the old allocations.** Nothing in the import can
delete. Column Q (previous team) now produces a checklist naming exactly which
project allocations to remove by hand.

**`idProfile` was hardcoded to `4`** in allocations, so a Project Leader (3) could
not be allocated. It now comes from column J.

## Project maps

Banking was **stale**: missing `74 PL & Others` and `75 Marruecos`, both created in
the 2026-08 Payment restructure. Added.

`Management` (39 AP / 51 Banking / 54 Q2C / 55 R2R) is **deliberately excluded** —
regular users are not allocated to it. Confirmed with the user, not an oversight.

Verified against `export_Project_20260806_140106.csv` by **WBS prefix**, not by
name: this instance has five projects called `Management`, two called
`Collections` and three called `IB`, so a name-based check reports phantom gaps.

## New sheet columns

| Col | Meaning |
|---|---|
| **P** | ProjeQtor resource id — **required for Update**; the row is skipped and reported without it |
| **Q** | previous team — optional, drives the removal checklist on a team change |

## Files the macro writes

| File | Element type | Note |
|---|---|---|
| `import_hr_new_resources.csv` | Resource | unchanged |
| `import_hr_update_resources.csv` | Resource | now `id`-keyed |
| `import_allocations_by_id.csv` | Allocation | Update rows |
| `import_allocations_by_name.csv` | Allocation | Create rows |
| `CHECKLIST_allocations_to_remove.txt` | — | UI job; the import cannot delete |

**Import each allocation file once.** Allocation rows always INSERT — there is no
upsert — so re-running duplicates them.
