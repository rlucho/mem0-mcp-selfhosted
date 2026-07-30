# PP2 audit extract — FEBAN statement evidence

Pulls the SAP-side evidence for the 56 payment samples in the auditor's
`Samples_Paper_SURL260716_152455.962.xlsx`: the FEBAN bank-statement line, the FI document it
posted to, the vendor behind it, and the attached invoice image where the configuration allows
it to be exported.

The macro is driven from a workbook, not from hardcoded screen IDs. Adapting it to your SAP
release is a paste-the-recording exercise — see [`RECORDING_GUIDE.md`](RECORDING_GUIDE.md),
which also answers the `Alt`+`F12` question.

---

## Why FEBAN is not read-only

You asked whether there is any risk, since everything is reading and exporting. Mostly yes —
with one real exception worth knowing before you run this.

**FEBAN is a post-processing transaction, not a display transaction.** Its whole purpose is to
let a user finish off bank-statement items that did not clear automatically: assign them,
change them, and post them. A user who can run FEBAN can post with it. The read-only-looking
part (browse the items, export the list) sits in the same transaction as the write part,
sometimes one keystroke apart — `F11` saves, and on some releases the *Post* button lands in
the toolbar slot that a recording from a different release will have captured as *Execute*.

Three consequences for this project:

1. **A guard module ships with it.** `modSafety` allowlists transactions, blocks committing
   OK-codes (`BU`, `SICH`, `POST`, `STOR`, `FBRA`, …), blocks `F11`, refuses to press a button
   whose own label reads *Post* / *Save* / *Reverse* / *Change*, and stops the run on any
   popup it does not recognise rather than pressing Enter through it.
2. **The guard is defence in depth, not a guarantee.** It cannot stop a recorded ID that
   happens to be a Post button, and it cannot undo anything. **The control that actually makes
   this safe is the authorisation profile of the user ID you run as** — ask Basis for a
   display-only role for the extract. That is a one-line request and it removes the risk
   properly.
3. **If you only need to *see* the statements, use `FF.6`** (Display Electronic Bank
   Statement) instead — it is genuinely display-only. Put it in *Transaction for statement
   search* on the Control sheet. Use FEBAN when you specifically need its post-processing view,
   which shows each item's posting *status* — often exactly what an auditor is asking about.

Everything else in the chain is display-only already: `FB03`, `FBL1N`, `FAGLL03`, `OAOR`, and
the `%pc` / ALV export path. The only writes are files on your own filesystem.

Two smaller things, stated once:

- **Scripting has to be enabled, and the attach is logged.** `sapgui/user_scripting = TRUE`
  plus the client-side option. `SM20` will show it. For audit-support work that is helpful, but
  tell Basis rather than letting them find it in the log.
- **The extract will contain vendor and payment data for a live company code.** Point the
  download root at a location already approved for it.

---

## What the auditor's file actually asks for

Worth reading before you run anything, because it changes what "done" looks like.

- The 10 month tabs are **not empty**. They already hold **55 pasted screenshots** of Citibank
  `Account Statement Report` pages, one per sample, with the sample line highlighted. The
  bank-side evidence is already collected.
- Column G of `Paper Samples` — **"Payment to Supplier?"** — is **blank on all 56 rows**. That
  is the open question, and it is the one thing the bank statement cannot answer.
- **18 of the 56 samples show `-` as the party.** These are the `ACH PYMTS - LCL BULK FNDG`
  bulk fundings, where the statement names no beneficiary at all. For those lines the SAP
  drill-down is the *only* way to answer column G. They are the samples that matter.
- **`Oct 25` has 8 sample rows but only 7 screenshots**, so one October line has no bank
  evidence attached.
- **One October screenshot is for a different account** — 12343649 / `GB49CITI18500812343649`
  in EUR, where every other screenshot is 12343657 / `GB27CITI18500812343657` in GBP. Either a
  second account is in scope or a screenshot was mis-pasted. Worth resolving before you send
  anything back, because it changes which house bank the FEBAN search should cover.

The full list of data-quality findings is on the `Data Issues` sheet of the generated workbook.

---

## Files

```
sap-audit-macro/
├── FEBAN_Audit_Control.xlsx    the workbook you run from
├── samples.csv                 the 56 samples, extracted and normalised
├── RECORDING_GUIDE.md          Alt+F12, what to record, how to map it
├── scripts/
│   ├── extract_samples.py      auditor's xlsx  -> samples.csv
│   └── build_control_workbook.py   samples.csv -> FEBAN_Audit_Control.xlsx
└── vba/
    ├── modConfig.bas           reads every setting and element ID from the workbook
    ├── modSafety.bas           the read-only guard
    ├── modSapConnect.bas       attaches to your running session; verifies the SID
    ├── modFeban.bas            per-month statement search and per-sample matching
    ├── modDrilldown.bas        statement item -> FI document -> vendor -> attachment
    ├── modExport.bas           ALV and classic list export to local files
    ├── modLog.bas              the audit trail
    ├── modUtil.bas             date, amount, filename helpers
    └── modMain.bas             entry points
```

## Setting it up

1. **Rebuild the workbook** if you want to regenerate it from the auditor's file:

   ```bash
   python3 scripts/extract_samples.py /path/to/Samples_Paper_SURL260716_152455.962.xlsx
   python3 scripts/build_control_workbook.py
   ```

2. **Save `FEBAN_Audit_Control.xlsx` as `.xlsm`** — Excel will not keep macros in `.xlsx`.

3. **Import the modules:** `Alt`+`F11` → File → Import File → each `.bas` in `vba/`. Add a
   reference to *Microsoft Scripting Runtime* if `Scripting.Dictionary` will not resolve
   (Tools → References).

4. **Fill in the yellow cells on `Control`.** The `SAP date format` cell matters more than it
   looks — `SU3` → Defaults → Date format is the authority. Wrong value means FEBAN searches a
   period that does not exist and reports nothing found, with no error.

5. **Record a session and fill in `Screen Map`** — see
   [`RECORDING_GUIDE.md`](RECORDING_GUIDE.md). Every required row must be filled; the run
   aborts naming any that are blank rather than guessing at an ID.

## Running it

Log on to SAP yourself first. The macro attaches to your session; it never logs on, so the
extract stays attributable to a named person in SAP's own log.

| Order | Macro | What it does |
|---|---|---|
| 1 | `modMain.CheckSetup` | verifies the config and the connection. Touches no data. Prints how it will format a date — check that. |
| 2 | `modFeban.DumpGridColumns` | with a FEBAN result list on screen, prints the grid's real column names for `FEBAN.Col.*` |
| 3 | `modMain.RunExtract` with **Run mode = DRY RUN** | walks every screen, logs what it would export, writes nothing |
| 4 | `modMain.RunSingleMonth` | one month with **EXTRACT**, to confirm the files look right |
| 5 | `modMain.RunExtract` with **EXTRACT** | the whole job |

FEBAN selects a period, not a payment, so the run opens each month once and walks the samples
inside it — 10 FEBAN executions for 56 samples, not 56.

## Output

```
<download root>/
  Sep 25/
    01_2025-09-03_8072447_FEBAN_statement.txt
    ...
```

and, in the workbook itself:

- **`Samples`** — status, statement item, FI document, vendor invoices and file count per line.
  Grey columns; the macro overwrites them each run.
- **`Log`** — one row per action, stamped with SAP system, client and user. Keep this with the
  extract; it is the record of who pulled what, from where, and when.

## Known limits

- **Invoice images often cannot be scripted out of SAP.** Listing the attachments works.
  Extracting the PDF or TIFF itself usually does not: SAP hands the document to the registered
  external viewer, and a script cannot capture that. The macro logs those as `MANUAL` rather
  than reporting a success that left no file behind. For all 56 in one go, a read-only extract
  from the content server by Basis is the better route — `modDrilldown`'s header explains why.
- **Ties are not resolved automatically.** Matching is on value date plus amount. Several
  samples are for exactly 2,450,000.00 on month-end dates, so ties are a real prospect; those
  are reported `AMBIGUOUS` for a human to settle rather than resolved by picking the first row.
- **`FBL1N` selection-screen IDs are not in the default Screen Map.** The vendor line-item
  export is stubbed and logs `SKIPPED` until you record that screen too.
- **Statement lines that were never posted** come back `UNPOSTED` with no document. That is a
  finding to report, not a failure to fix.
