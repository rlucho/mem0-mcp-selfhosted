# Closing Manager — Analysis & V4-CIO hardened macros

Analysis of the legacy SAP/Excel VBA automation
`Closing_Manager_IP_Legacy_changed_v3.xlsm`, plus a hardened **V4-CIO** set of
VBA modules that fix the defects found during the review.

> **Nothing here executes the macro or contains any credentials or live data.**
> The VBA was extracted statically from `vbaProject.bin` with `olevba`.

---

## Contents

| Path | What it is |
|------|------------|
| **`Closing_Manager_IP_V4-CIO.xlsm`** | **The ready-to-use workbook** — the original with the V4 fixes baked in. |
| `report/Closing-Manager-Analysis.html` | The shareable report (open in a browser; also published as an artifact). |
| `report/Closing-Manager-Analysis.md` | Same report in Markdown. |
| `report/v4-changes.diff` | Unified diff of every V4 change vs the original modules. |
| `vba-v4/` | The V4-CIO modules as source (for import, or review). |
| `vba-original/` | The original modules, extracted verbatim, for reference/diffing. |
| `build/` | The offline rebuild pipeline (Python) that produced the `.xlsm`. |

## Two ways to get V4

1. **Use the built workbook (recommended):** open `Closing_Manager_IP_V4-CIO.xlsm`.
   It is the original file with only the VBA changed — see *How the workbook was
   built* below.
2. **Import the modules yourself:** import `vba-v4/mCloseEnv_V4.bas` and replace
   `GlobalModule` / `Closing` / `Printing` (steps further down). Use this if you
   prefer to apply the fix to your own current copy of the workbook.

---

## What the macro does (one paragraph)

A Capgemini "Closing Manager" for International Paper's month-end **Record-to-Report**
close. It attaches to a running **SAP GUI** session via SAP GUI Scripting, drives
transactions to export data as flat files into the workbook's own folder, parses
them back into worksheets, posts closing entries, prints statutory reports to PDF
via **PDFCreator**, merges them with a bundled **`GiosPSMC.exe`**, and files a single
dated PDF under `C:\_Files to Transfer\MONTH END CLOSE\<year>\<mm>\`. It is a
Windows-desktop RPA screen-scraper — see the full report for the pipeline, the
dependency list, and the ranked failure points.

---

## V4-CIO — the fixes

All path/environment logic is centralised in one new module, `mCloseEnv_V4.bas`.
The three edited modules keep the same procedure names and call signatures, so the
rest of the workbook (buttons, other macros) is unaffected.

| # | Fix | Where | Failure it removes |
|---|-----|-------|--------------------|
| 1 | **Single working drive** — one `CM_BASE_DRIVE` used everywhere (was `D:\`-if-present in `CreatePaths` but hard-coded `C:\` in `RunClosing`). | `GlobalModule`, `Closing`, `mCloseEnv_V4` | Silent print/merge failure on any PC with a `D:` drive. |
| 2 | **OneDrive / SharePoint guard** — `AssertLocalWorkbook` stops with a clear message if the workbook path is a URL. | `GlobalModule`, `mCloseEnv_V4` | Cryptic mid-run failure when run from the cloud. |
| 3 | **Parent-aware folder creation** — `EnsureFolders` builds any missing ancestor folders on the correct drive. | `mCloseEnv_V4` | `CreateFolder` failing when a parent doesn't exist. |
| 4 | **PDFCreator presence check** — clear error instead of an unguarded `Shell` of the (x86) path. | `Printing` | Run-time crash when PDFCreator isn't installed. |
| 5 | **Bounded merge waits** — the two PDF wait loops get a safety cap. | `Printing` | Excel hanging forever if SAP/PDFCreator stalls. |
| 6 | **Merger fast-fail** — `CombinePDF` verifies `GiosPSMC.exe` before shelling out. | `Printing` | Silent no-output when the tool is missing. |
| 7 | **Name-clash loop fix** — `fN` now increments. | `Printing` | Infinite loop when a same-named report already exists. |
| 8 | **Preflight Check** *(new)* — one-click `PreflightCheck` reports every dependency as pass/fail before a run. | `mCloseEnv_V4` | Discovering a missing dependency halfway through a close. |

Functional close logic and the SAP command sequences are **unchanged** — the diff
is intentionally small (`report/v4-changes.diff`).

### Configuration — change locations in one place

Open `mCloseEnv_V4.bas` and edit the `CONFIGURATION` block if your environment differs:

```vba
Public Const CM_BASE_DRIVE  As String = "C:\"                                  ' working drive for \pdf\ folders
Public Const CM_REPORT_ROOT As String = "C:\_Files to Transfer\MONTH END CLOSE\"
Public Const CM_MERGER_SRC  As String = "\\pl-krabpo-fsc01\...\GiosPSMC.exe"   ' network master copy
```

> If your standard build genuinely stores data on `D:\`, set `CM_BASE_DRIVE = "D:\"`
> — that one line now governs *every* folder consistently.

---

## How the workbook was built (and how it was validated)

The `.xlsm` was produced **without Excel**, by surgically rewriting only the
workbook's `xl/vbaProject.bin` (the VBA container) and copying every other part
of the file byte-for-byte. Pipeline in `build/`:

1. Decompress the original `GlobalModule` / `Closing` / `Printing` source
   (MS-OVBA), apply the exact V4 edits, and re-compress. The `mCloseEnv_V4`
   helpers are **folded into `GlobalModule`** (see
   `vba-v4/GlobalModule_folded_for_xlsm.bas`) so no new module had to be added.
2. Rebuild each module stream keeping its original compiled-code prefix and
   appending the new compressed source.
3. Force Excel to recompile from the new source on first open: bump the
   `_VBA_PROJECT` version stamp and empty the stale `__SRP_*` compiled-cache
   streams.
4. Rebuild the compound file and drop it back into the xlsm zip.

**Validation performed here (no Excel required):**

- The compound-file writer was proven with a **byte-perfect identity
  round-trip** (rebuild the original with no edits → 116/116 streams identical).
- `olevba` re-extracts all 26 modules; exactly 3 carry the V4 code, the other 23
  are identical to the original; every V4 marker is present.
- `olefile` reads every stream; all UserForms and other modules are untouched.
- `openpyxl` opens the workbook; **only `xl/vbaProject.bin` differs** from the
  original — all sheets, forms, images, external links and the sensitivity
  label are byte-identical.

> **One caveat:** the actual *recompile-and-run in Excel* step can't be exercised
> in this Linux build (no Excel/Windows here). The technique used — invalidating
> the compiled cache so Excel rebuilds from the embedded source — is the standard
> mechanism, and everything checkable offline passes. On first open, Excel may
> take a moment to recompile, and if macro security prompts appear, choose
> **Enable Content**. If your Excel ever rejects the baked file, the module
> import route below is the guaranteed fallback and produces the identical result.

## How to install V4 (≈ 2 minutes, in Excel on Windows)

VBA source can't be re-injected into an `.xlsm` outside Excel, so V4 ships as
importable modules. In the workbook:

1. **Back up** the `.xlsm` first (copy the file).
2. Press **Alt + F11** to open the VBA editor. If the Project pane is hidden, press **Ctrl + R**.
3. **Import the new module:** right-click the project → **Import File…** → choose
   `vba-v4/mCloseEnv_V4.bas`.
4. **Replace the three edited modules.** For each of `GlobalModule`, `Closing`,
   `Printing`:
   - right-click the existing module → **Remove …** → **No** when asked to export;
   - right-click the project → **Import File…** → choose the matching file in `vba-v4/`.

   *(Alternatively, open each module and paste the V4 contents over the old code —
   the procedure names are identical.)*
5. **Compile:** menu **Debug → Compile VBAProject**. It should compile with no errors.
6. **(Recommended)** add a button on the `START` sheet assigned to the macro
   **`PreflightCheck`**, and run it once to confirm the environment.
7. Save as macro-enabled (`.xlsm`).

> `mCloseEnv_V4.bas` uses `Option Explicit`; the three edited modules deliberately
> do **not**, to stay byte-for-byte faithful to the original except for the fixes.

---

## Recommended follow-ups (not included — they need live-Excel testing)

- Wrap the SAP transaction blocks in structured error handling (log + optional
  screenshot on failure) instead of the current bare `.findById` chains.
- Bind the SAP session by **connection description** rather than `Children(0)`, so
  the macro can't drive the wrong open SAP window.
- Add `Option Explicit` to every module and resolve the resulting undeclared-variable
  warnings (e.g. the dead, procedure-local `FShared`).
