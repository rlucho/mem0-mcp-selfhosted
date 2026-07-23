# Closing Manager — What It Does, and Where It Breaks

**Artifact:** `Closing_Manager_IP_Legacy_changed_v3.xlsm`
**System:** SAP GUI Scripting + Excel VBA (desktop RPA) · 26 modules, ~5,800 lines
**Domain:** R2R month-end close, International Paper (Capgemini automation)
**Method:** static source review (`olevba` extraction of `vbaProject.bin`; no macro executed)

---

## Bottom line

A tightly environment-coupled desktop automation — an SAP-GUI screen-scraper wired
to PDFCreator and a bundled PDF merger. It runs cleanly **only** on a Windows PC, on
the corporate network, with SAP open, **from a local folder**. It is **not** safe to
run from OneDrive/SharePoint, and it carries a silent drive-selection defect that
breaks printing on any PC with a `D:` drive.

Severity tally: **0 Critical · 5 High · 3 Medium · 4 Low/code-smell.**

---

## 1. What the file is

A Capgemini "Closing Manager" for International Paper's Record-to-Report month-end
close. It drives SAP to pull data, post entries, print statutory reports to PDF, and
file a single merged closing report per cost centre. It is an **RPA screen-scraper**,
not an API integration. The `_changed_v3` build has most old SharePoint tracker calls
commented out and leans on the workbook's own sheets plus a prompt for the SAP user.

## 2. How it works — the run pipeline

Entry points: `UpdateData` (refresh master data), `RunClosing` (the close),
`CombinePDF` (merge). A close proceeds:

1. **Attach to SAP** — `GetObject("SAPGUI")` grabs an already-running, logged-in
   session (`Common.SAPsess()`). No session → aborts with "You are not logged in SAP."
2. **Extract from SAP** — drives transactions (SE16 dumps `T001`/`ZCCOD`/`SKB1`,
   reports `ZGLRME`/`ZGE132`/`AA02`…), saving each as a fixed-name `.txt`/`.csv` into
   the workbook's own folder.
3. **Parse back in** — `ADODB.Stream` reads each export; a custom pipe-delimited parser
   loads sheets; `Kill` deletes the temp files.
4. **Post & report** — posts closing entries and runs reporting transactions
   (`Postings.bas`), again by scripted screen IDs.
5. **Print to PDF** — sets SAP's front-end printer to the OS printer named
   `PDFCreator`, launches `PDFCreator.exe`, waits for each PDF in `\pdf\final\`.
6. **Merge** — shells out to the bundled `GiosPSMC.exe` to concatenate the PDFs.
7. **Deliver** — moves the merged report to
   `C:\_Files to Transfer\MONTH END CLOSE\<year>\<mm>\` under a dated name.

## 3. Points of failure (ranked)

| # | Failure point | Severity | What happens & when |
|---|---------------|----------|---------------------|
| 1 | **Working-drive split-brain** (Defect) | HIGH | `CreatePaths` picks `D:\` if a D: drive exists, but `RunClosing` hard-codes folder creation on `C:\`. On a D:-drive PC the folders + merger are created on C: but read from D: → **print & merge fail silently.** |
| 2 | **Run from OneDrive/SharePoint** (Environment) | HIGH | `FPath = ThisWorkbook.Path` becomes an `https://…` URL when cloud-hosted. SAP's download dialog, `ADODB.Stream` I/O and `Kill` all reject URLs → dies at the first export. |
| 3 | **SAP not open / not scriptable** (Dependency) | HIGH | `GetObject("SAPGUI")` needs SAP running, logged in, scripting enabled client- & server-side. Binds only the *first* connection/session — wrong window if several are open. |
| 4 | **Brittle hard-coded SAP screen IDs** (Fragility) | HIGH | Hundreds of `.findById("wnd[1]/usr/…")` with almost no error handling. Any SAP version/theme change, altered variant, or pop-up halts the macro and leaves orphaned temp files. |
| 5 | **PDFCreator dependency** (Dependency + Defect) | HIGH | Needs PDFCreator with the OS printer named exactly `PDFCreator`. If absent, `SetPDFCreator` `Shell`-launches the (x86) path *without checking it exists* → run-time error. |
| 6 | **Fixed network share / VPN** (Dependency) | MEDIUM | Merger copied from — and reports archived to — `\\pl-krabpo-fsc01\ipa$\R2R\…`. Off-network, the tool can't be fetched (if not already local) and archiving fails. |
| 7 | **Legacy SharePoint SOAP endpoint** (Dependency) | MEDIUM | Active calls hit `troom-x.capgemini.com/…/_vti_bin/Lists.asmx` (SP2010 SOAP) via Windows auth. If retired, calls error or silently no-op. |
| 8 | **Unbounded wait loops** (Robustness) | MEDIUM | PDF steps poll with `Application.Wait` and **no timeout**. If SAP or PDFCreator stalls, Excel hangs indefinitely. |
| 9 | **Infinite loop on report name clash** (Defect) | LOW | In `CombinePDF` the de-dup loop sets `FName = …&fN` but never increments `fN`. If base + `-1` names both exist, it spins forever. |
| 10 | **Stale temp files after a crash** (Robustness) | LOW | Fixed-name exports are `Kill`ed at the end; a prior crash leaves them, so the next run may read stale data. |

Also noted (code-smell, no runtime impact): `FShared` is computed in `CreatePaths` but,
being undeclared, is procedure-local and never used; the project runs without
`Option Explicit`, so variable-name typos fail silently.

## 4. Your questions, answered

**Does it need fixed folder paths? — Yes.**
- `C:\_Files to Transfer\MONTH END CLOSE\<year>\<mm>\` — final report (always C:).
- `C:\pdf\{temp,merger,mergedFiles,printed,final}\` (or `D:\pdf\…`) — PDF working area.
- `\\pl-krabpo-fsc01\ipa$\R2R\…\GiosPSMC.exe` — source of the merger tool.
- `\\pl-krabpo-fsc01\ipa$\R2R\…\CLOSING REPORTS\…` — archive share.
- *Relative:* the workbook's own folder (`ThisWorkbook.Path`) — the SAP export scratch
  area, which **must be a real filesystem path**.

**Does it need a file in a given folder, with a fixed name? — Yes.**
- `GiosPSMC.exe` at `<drive>\pdf\merger\` — auto-copied from the share if missing; if the
  share is unreachable *and* it isn't already there, the merge dies.
- `PDFCreator.exe` under `C:\Program Files\PDFCreator\`, printer named `PDFCreator`.
- Fixed-name SAP exports created at runtime: `t001.txt · zgxmit.txt · skb1.txt · CC.csv
  · PC.csv · zglrme.txt · aa02.txt · zge132.txt · DocL.csv · DocG.csv` …
- The workbook's own `config` (cost centre in `B2`) and `SAP config` sheets must be
  populated.

**What are the requirements?**
- Windows + desktop Excel (Windows), macros trusted.
- SAP GUI for Windows — logged in, scripting enabled; SAP authorisations for the
  driven transactions.
- PDFCreator installed; printer named `PDFCreator`.
- Corporate network / VPN (UNC shares + SharePoint).
- Local write access to `C:\`.
- COM: ADODB, MSXML2, Scripting, WScript.Shell.
- Runtime input: SAP user + cost centre in `config!B2`.
- Windows-desktop only — will not run on Mac Excel or Excel for the web.

**Can it run locally, or is OneDrive/SharePoint OK? — Local only.**
Run it from a plain local folder. When opened from a OneDrive/SharePoint (cloud)
location, `ThisWorkbook.Path` returns an `https://…` address instead of a drive path —
breaking the SAP download dialog, the `ADODB.Stream` reads/writes, and the `Kill`
cleanups at once. Even a locally-synced OneDrive folder is risky (AutoSave can surface
a URL path; constant temp-file churn causes sync locking). The other fixed paths are
absolute and unaffected — the only constraint on where the workbook lives is that
**its own path must resolve to a genuine local filesystem folder.**

> **Operator trap:** the **D: split-brain defect (#1)** makes a close appear to work
> right up until printing produces nothing, with no error, on any PC that has a D:
> drive. Fixed in the V4-CIO build.

## 5. External dependencies (Capgemini / IP-Polska)

Two Capgemini-hosted systems are reached at run time; the rest are reference-only.
Status is for this `_changed_v3` build.

**Capgemini SharePoint — SOAP (ACTIVELY called)**
- `https://troom-x.capgemini.com/sites/InternationalPaper/r2a/_vti_bin/Lists.asmx`
  — live dependency; 11 calls (Admin ×7, GlobalModule ×4) via the `sp*` helpers,
  used by `UpdateData` and list logging during the close. Windows-auth; needs VPN.
- `https://troom-x.capgemini.com/sites/InternationalPaper/CC/CG/_vti_bin/Lists.asmx`
  — Employees/SAP-ID lookup; **commented out** in v3.

**IP-Polska file server — UNC (`\\pl-krabpo-fsc01`, share `ipa$`)**
- `\\pl-krabpo-fsc01\ipa$\R2R\R2R - IP EU GL West\USEFUL\pdf\merger\GiosPSMC.exe`
  — **ACTIVE, load-bearing** (merger copied locally on first run).
- `\\pl-krabpo-fsc01\ipa$\R2R\R2R - IP EU\MONTH-END\CLOSING REPORTS\<year>\<mm>`
  — `FShared`; **dead** (never read in v3).

**Reference only (not called)**
- `https://capgemini.sharepoint.com/sites/InternationalPaper/capgemini/Lists/Automations_list/Allitemsg.aspx`
  — governance register in module header comments.
- `https://capgemini-my.sharepoint.com/IP All/…/NEW AR AP Intercompany Matching Master.xlsm`
  (+ twin `U:\IP All\…\NEW AR AP Intercompany Matching Master.xlsm`) — external
  workbook formula link.

*Not* dependencies: `schemas.microsoft.com/sharepoint/soap/…`,
`schemas.xmlsoap.org/soap/envelope/`, `www.w3.org/2001/XMLSchema…` are XML
namespaces, never contacted. Runtime reachability needed: **troom-x.capgemini.com**
and **\\pl-krabpo-fsc01\ipa$** (corporate network / VPN). In V4 these are the named
constants `CM_SP_BASE`, `CM_MERGER_SRC`, `CM_ARCHIVE_ROOT`.

## 6. V4-CIO — what was fixed

See `../README.md` for the full change table and import steps. Summary: one consistent
working drive, a OneDrive/URL guard, parent-aware folder creation, a PDFCreator presence
check, bounded merge waits, a merger fast-fail, the name-clash loop fix, and a new
one-click `PreflightCheck`. The fragile SAP-scripting core is left untouched; the diff
is intentionally minimal (`v4-changes.diff`).

**Recommended follow-ups** (need live-Excel testing): structured error handling around
SAP blocks; bind SAP by connection description not `Children(0)`; add `Option Explicit`
project-wide.
