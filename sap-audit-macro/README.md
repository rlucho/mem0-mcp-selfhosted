# PP2 audit extract — FEBAN statement evidence

Pulls the SAP-side evidence for an auditor's payment samples: for each payment line, the FEBAN
statement item, the FI document, the clearing document, the batch of ZP payments behind it, the
largest of those payments, and the largest invoice inside that — ending in that invoice's PDF.

**Import request** reads any of the auditor's sample files. Seven of them have turned up in
three layouts with the header on a different row in almost every one, so the importer finds it
by content rather than by position and shows you what it matched before writing anything. The
company code is asked once per file, because the requests are not all the same company.

The macro is driven from a workbook, not from hardcoded screen IDs: every `findById` string is
read from the `Screen Map` sheet, so correcting one is an edit to a cell rather than to code.
See [`RECORDING_GUIDE.md`](RECORDING_GUIDE.md), which also answers the `Alt`+`F12` question.

**Choosing which rows to run:** you pick samples, never months — the run works out which FEBAN
periods it needs and opens them itself. **RUN SAMPLES** opens with the scope per request and
one question: run the rows marked `Include?`, run everything regardless, or cancel. Set a row
to `No` to leave it out; blank counts as included, and **Include/Exclude every sample** bulk-set
the column so excluding a handful is not 141 edits.

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
├── RECORDING_GUIDE.md          Alt+F12, what each recording gave us, what is still missing
├── SETUP.md                    two steps to a first test run
├── build_xlsm.vbs              optional one-click .xlsm build
├── recordings/                 all six, decoded from UTF-16 to plain text
│   ├── Audit.vbs               steps 1-7, first pass
│   ├── Audit2.vbs              steps 1-7 incl. the ALV export and the ZP list export
│   ├── Audit3.vbs              steps 1-6, stops at the Payment Usage list
│   ├── Audit5.vbs              steps 1-7, ZP list exported
│   ├── N1.vbs                  steps 6-10, Santander SCF, ending in the PDF
│   └── B2.vbs                  steps 6-10, regular vendor, ending in the PDF
├── scripts/
│   ├── extract_samples.py      auditor's xlsx  -> samples.csv
│   ├── build_control_workbook.py   samples.csv -> FEBAN_Audit_Control.xlsx
│   ├── test_list_parser.py     tests for the parsing logic in modExportRead.bas
│   ├── test_import_detection.py  runs modImport's header detection over the real requests
│   ├── vbacheck.py             cross-module refs, UDTs, blocks, GoTo labels
│   ├── vbarules.py             VBA compile rules a structure check misses
│   ├── undeclared2.py          every identifier resolves to something
│   └── keycheck.py             Screen Map and Control keys the code asks for
└── vba/
    ├── modConfig.bas           reads every setting and element ID from the workbook
    ├── modSafety.bas           the read-only guard
    ├── modSapConnect.bas       attaches to your running session; verifies the SID
    ├── modProbe.bas            checks predicted IDs against the live screen
    ├── modFeban.bas            steps 1-2: per-month search, per-sample matching
    ├── modChain.bas            steps 2-7 and 10: the walk for one sample
    ├── modFbl1n.bas            steps 8-9: the ZP payments of one batch
    ├── modExportRead.bas       reads an export back (workbook or text)
    ├── modExport.bas           ALV and classic list export to local files
    ├── modImport.bas           reads an auditor's request into the Samples sheet
    ├── modReport.bas           the per-sample report workbook
    ├── modLog.bas              the audit trail
    ├── modUtil.bas             date, amount, filename helpers
    ├── modMain.bas             entry points
```

## The chain, per sample

```
 1  FEBAN            company code + statement dates from the audit row
 2  result list      export it; find the row by date+amount; open it
 3  item detail      Posting Area 1 Doc. number -> F2 -> FI document
 4  FI document      first posting-key-40 line with a clearing doc
 5  line detail      clearing doc field -> F2 -> clearing document
 6  Environment > Payment Usage -> the batch's ZP documents
 7  export it, read the ZP document numbers back off disk
 8  FBL1N            company code, all items, month, those ZP numbers
 9  export, take the LARGEST ZP payment of the batch
10  Payment usage on it -> the most NEGATIVE row: a payment is a debit and
    the invoice it settles is a credit, so the biggest invoice is the most
    negative. Santander SCF pays a finance provider, so that route takes one
    more hop -- open the settlement, its clearing line, payment usage again --
    before the supplier invoices appear. Then open the invoice and save the
    attachment titled 'Invoice'.
```

Every step is covered by a recording in [`recordings/`](recordings/), so the IDs are captured
values rather than guesses. A stage whose Screen Map rows are blank still finishes at the last
step that worked and records the document numbers a person needs to carry on.

**One route, not two.** `N1.vbs` walked a Santander SCF payment and `B2.vbs` a regular vendor,
and they are the same shape. The regular one only looked shorter because there were fewer rows
to sort through. `IsConfirmingPayment` now only colours the report.

**The invoice is the KR document.** The list step 10 picks from holds several document types at
once — ZP payments, KR vendor invoices, SB statement documents — and the KR rows carry
*negative* amounts. Filtering to KR is what stops the largest row coming back as a payment;
comparing on magnitude is what stops the sign mattering.

Two "largest" decisions, both read from exported data rather than assumed from a sort order.
Worth knowing why: the recordings sort the list on screen and click a row by position
(`lbl[164,8]`), which is a fixed screen coordinate — across 56 batches of different sizes that
opens whichever document happens to land on row 8, silently.

## Verifying an ID without recording

Steps 8–10 are predictions, and there is a faster way to check them than another Alt+F12
session:

1. Enter `FBL1N` by hand and stop on the selection screen.
2. Run **`modProbe.ProbeCurrentScreen`**. It reports which of the mapped IDs are actually on
   that screen, and what type of control each one is. It only reads — it presses nothing.
3. For any it cannot find, run **`modProbe.DumpScreen`** to list every named element there,
   and paste the right one into column F.
4. Execute FBL1N by hand, then run `ProbeCurrentScreen` again for the result-grid rows and
   **`modFbl1n.DumpFbl1nColumns`** for its column names.

One thing to expect: on standard FBL1N the **document number is not on the main selection
screen** — it lives in dynamic selections. If the probe cannot find `Fbl1n.DocNumberFrom`,
leave the whole `Fbl1n.*` block blank. The run then stops at step 7 having exported the
Payment Usage list and logged the ZP numbers, which is still the bulk of the work, and steps
8–10 can be done from that list by hand.

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

5. **Check `Screen Map`.** It is already filled in from `recordings/Audit.vbs`. Deal with the
   rows marked `VERIFY` and `BLOCKED` before an EXTRACT run — see
   [`RECORDING_GUIDE.md`](RECORDING_GUIDE.md). Every required row must hold a value; the run
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

One folder per sample, named after the sample, with fixed names inside it. The folder says
which sample it is and the file name says which step of the chain produced it, so a folder can
be zipped and sent to the auditor as it stands.

```
<download root>/
  Sep 25/
    1 - FEBAN statement list.xlsx           the period's list, exported once
    01 - 8072447.42/
      01 - 8072447.42 - GBKM.xlsx           the report — read this first
      1 - FEBAN statement list.xlsx         a copy, so the folder stands alone
      2 - Payment usage - batch of payments.xlsx
      3 - FBL1N - payments in the batch.xlsx
      4 - Documents behind the largest payment.xlsx
      5 - Supplier invoices behind the largest SCF settlement.xlsx   (confirming payments only)
      6 - Largest invoice.pdf
    05 - 8617262.93/
      05 - 8617262.93 - GBKM.xlsx
      1 - FEBAN statement list.xlsx
      2 - FI document line items (not cleared).xlsx
```

The report workbook carries the trail top to bottom — what the auditor asked for, the statement
line it matched, the FI and clearing documents, the batch, the largest payment, the largest
invoice — with each evidence file listed against the step it proves and the PDF embedded as an
object. Its `Log` sheet holds that sample's rows from the run's audit trail.

The second folder above is the shape a sample takes when the FI document clears nothing: a
transfer between the company's own bank accounts settles no supplier, so there is no batch, no
payment and no invoice. The line items are the evidence, and the status is `NO CLEARING` rather
than an error.

A re-run clears the numbered files out of a sample's folder before re-exporting, because the
SAP save dialog never overwrites — without it a second run would leave `..._2.xlsx` beside
every file. Only names this macro writes are removed.

The recordings name every export `.XLSX`, so these come back as real workbooks rather than
delimited text. `modExportRead` sniffs the first two bytes — a ZIP signature means a workbook
and it is opened with `Workbooks.Open`; anything else is parsed as text. Trusting the
extension would not work, because SAP writes plain text into a `.XLSX` name often enough.

and, in the workbook itself:

- **`Samples`** — status, statement item, FI document, vendor invoices and file count per line.
  Grey columns; the macro overwrites them each run.
- **`Log`** — one row per action, stamped with SAP system, client and user. Keep this with the
  extract; it is the record of who pulled what, from where, and when.

## Known limits

- **Level 2, the Santander SCF route, is blocked.** `Audit2.vbs` captured no steps — just the
  scripting boilerplate and a `resizeWorkingPane` call. Those samples run all of level 1, then
  log `BLOCKED_SCF` carrying the clearing document, the SCF item's document number and the path
  to the exported cleared-items list, so they can be finished by hand. Filling in
  `Scf.OpenPayment` and `Scf.InvoiceListMenu` from a fresh recording turns level 2 on with **no
  code change**. Eight of the 56 samples name `SANTANDER SCF` directly in the auditor's request,
  and more may turn out to be SCF once the cleared items are read — the branch is decided by
  what the cleared-items list says, not by the request.
- **The regular-supplier PDF download is unverified.** `Invoice.*` on the Screen Map is
  standard SAP rather than recorded, because that walk wasn't captured either. Where those
  rows are blank the run logs `MANUAL` with the document numbers, rather than reporting a
  success that left no file behind.
- **Invoice images sometimes cannot be scripted out of SAP at all.** Listing attachments works;
  extracting the PDF often does not, because SAP hands the document to the registered external
  viewer and a script cannot capture that. If that turns out to be the case here, a read-only
  extract from the content server by Basis beats fighting the GUI for 56 documents.
- **Ties are not resolved automatically.** Matching is on value date plus amount. Several
  samples are for exactly 2,450,000.00 on month-end dates, so ties are a real prospect; those
  are reported `AMBIGUOUS` for a human to settle rather than resolved by picking the first row.
- **Statement lines that were never posted** come back `PARTIAL` with no document. That is a
  finding to report, not a failure to fix.
- **None of this has been run against SAP.** It was written from the recordings, and the
  structural checks that can be done without SAP have been done — every cross-module
  reference, block and `GoTo` label resolves, and the list-parsing logic is tested in
  `scripts/test_list_parser.py`. The first `CheckSetup` and `DRY RUN` on your machine are the
  real tests.
