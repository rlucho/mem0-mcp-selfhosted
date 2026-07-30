# Recording the SAP script the macro is built from

## The shortcut

**`Alt` + `F12`** — the "Customize Local Layout" menu on the SAP GUI toolbar.
Then pick **Script Recording and Playback**.

In SAP GUI 7.60 and later the entry sits one level down: `Alt`+`F12` → **More…** →
**Script Recording and Playback**. Either way the dialog is the same, and there is no
direct key combination for it.

In the dialog:

1. **Record** (the red dot) → choose where to save the `.vbs`
2. Do the clicks you want captured
3. **Stop** (the square)

Ctrl-clicking the toolbar's "Customize Local Layout" button gets you to the same menu if
`Alt`+`F12` is intercepted by your terminal or remote-desktop client.

## If Alt+F12 shows no scripting entry

Scripting is switched off. It needs both halves:

- **Server:** profile parameter `sapgui/user_scripting = TRUE` (Basis sets it; `RZ11` shows
  its current value). Read-only check: `RZ11` → parameter name → Display.
- **Client:** SAP GUI Options → *Accessibility & Scripting* → *Scripting* → **Enable
  scripting**. Untick *Notify when a script attaches to SAP GUI* and *Notify when a script
  opens a connection* as well, or every step of the macro waits for a click.

Worth knowing before you ask: enabling scripting is logged, and on most landscapes the
security audit log (`SM19`/`SM20`) records the attach. For audit-support work that is a
feature — it evidences who pulled the data — but tell Basis what you are doing rather than
having them find out from the log.

## What the two recordings gave us

Both are in [`recordings/`](recordings/), decoded from UTF-16 to plain text.

**`Audit.vbs` — 38 captured steps, and it covers the main chain.** Everything it touched is
already filled into the Screen Map, marked `recorded` in the Source column. The flow it
proves out:

```
FEBAN
  selection popup (wnd[1]!)   GBKM, 01092025 .. 30092025, Execute
  result grid                 wnd[0]/shellcont/shell, sorted on KWBTR
  double-click the row
statement item detail         txtD2201_BELNR          -> the FI document
  F2
FI document overview          double-click the DMBTR line
line item detail              txtBSEG-AUGBL           -> the clearing document
  F2
clearing document             mbar/menu[5]/menu[3]    -> cleared items + supplier names
  mbar/menu[0]/menu[3]/menu[1]                        -> save-list dialog
save dialog                   ctxtDY_PATH, btn[0]
```

Three things in it were worth catching:

- **The FEBAN selection screen is a modal popup (`wnd[1]`), not a full screen.** Anything
  written against `wnd[0]/usr/...` would fail outright.
- **Dates were typed as 8 bare digits** — `01092025`, no separators. The Control sheet's
  `SAP date format` is set to `DDMMYYYY` to match. This one is a silent failure if wrong:
  FEBAN just returns nothing.
- **The save dialog confirms with `btn[0]`, not `btn[11]`.** And the recording never set
  `DY_FILENAME`, so it kept SAP's default name — that field's ID is the one guess left in the
  save path, and it's marked `VERIFY`.

**`Audit2.vbs` captured nothing.** It holds the scripting boilerplate and a single
`resizeWorkingPane` call — no steps at all. The recorder was almost certainly stopped before
the actions, or started after them. So the Santander SCF confirming-payment route is
**BLOCKED**: those samples run as far as identifying the supplier, then log `BLOCKED_SCF` with
the clearing document and invoice numbers so they can be finished by hand.

## What still needs recording

Three gaps. Everything marked `VERIFY` or `BLOCKED` on the Screen Map:

| Gap | What to record | Screen Map keys |
|---|---|---|
| **Level 2, the SCF route** | From the cleared-items list: get **into** the Santander SCF payment, then list the invoices it settled | `Scf.OpenPayment`, `Scf.InvoiceListMenu` |
| **The regular-supplier PDF** | The same walk for a non-Santander supplier: open the payment, *Services for Object* → *Attachment list*, save the PDF | `Invoice.GosToolbox`, `Invoice.AttachListGrid`, `Invoice.SaveButton` |
| **The save dialog's file-name field** | Type a file name into the save dialog this time, rather than accepting the default | `Save.FileName` |

### The two steps level 2 actually needs

Only **two** IDs unblock it. Everything after them reuses machinery that already works — the
export goes through the same `Export.ListMenu` / `Save.*` path as the cleared-items list, and
the largest-invoice pick reuses the same parser.

1. **`Scf.OpenPayment`** — how you get from the cleared-items list *into* the Santander SCF
   payment. Whatever you actually do: click a menu entry, press a button, or put the cursor on
   a document-number field and press F2. The macro reads the control's type and does the right
   thing, so just record the click.
2. **`Scf.InvoiceListMenu`** — the menu path that lists the supplier invoices that SCF payment
   settled. Probably another `mbar/menu[x]/menu[y]`, like the `menu[5]/menu[3]` that produced
   the cleared-items list.

Optionally **`Scf.InvoiceListAnchor`** — any element on that invoice list. It is used purely to
confirm step 2 landed where expected; leave it blank to skip the check.

Then, on the first run, open the exported `*_scf_invoices.txt` and check the macro picked the
right amount and supplier column. If it guessed wrong, copy the real captions into the
**`SCF invoice list ...`** settings on the Control sheet. The log says which columns it chose,
so you do not have to guess whether it got it right.

Plus one that isn't a recording: **run `modFeban.DumpGridColumns`** with a FEBAN result list on
screen. `Audit.vbs` only ever touched the `KWBTR` column, so the *value date* column name is
still a guess (`VALUT`), and matching every sample depends on it. This prints the real names.

When you re-record, keep it slow and single-purpose, and **close any attachment list with
Cancel rather than Enter** — Enter hands the document to the external viewer and blocks the
script behind it.

## Which sample to re-record Audit2 against

**Not Sep 25.** The first recording was made on Sep 25, and Sep 25 contains no Santander SCF
line at all — so the confirming route never comes up and you would capture the regular
supplier path by mistake. Dec 25 has none either.

Record against **sample #8: Oct 25, 01/10/2025, 4,483,676.08, `SANTANDER SCF`**. It is the
earliest sample the auditor's request names as SCF, so the confirming route is guaranteed to
appear. The other seven are #16 Nov, #24 Jan, #30 Feb, #34 Mar, #39 Apr, #46 May, #53 Jun.

Note that in no month is the SCF line the *largest* line of the month — so if the "biggest
payment" you were describing is the biggest item **within** a payment's cleared items rather
than the biggest statement line, say so, because the macro currently reads it the second way
(largest cleared item behind whichever sample line it is processing).

### Confirm the recorder is actually capturing

This is what went wrong last time — `Audit2.vbs` came back holding only the boilerplate and a
`resizeWorkingPane` call, which is what SAP writes the instant recording starts. So before the
real walk:

1. Press **Record**.
2. Make one throwaway click — open any menu, click a grid cell.
3. **Stop**, open the `.vbs`, and check there is at least one `findById(...)` line *after* the
   `resizeWorkingPane` line.
4. If there is, record again for real. If there is not, the recorder is not capturing and
   nothing you do next will be saved.

### Grab the other two gaps in the same session

Since you are recording anyway, these close out the remaining `VERIFY` rows:

- **In the save dialog, type a file name** rather than accepting SAP's default. That confirms
  `Save.FileName`, the last guess left in the export path.
- **Then do the same walk for a non-Santander supplier** — any `DS SMITH` line will do — to
  capture the regular-supplier PDF route: *Services for Object* → *Attachment list* → save.
  That fills `Invoice.GosToolbox`, `Invoice.AttachListGrid` and `Invoice.SaveButton`.

## Turning a recording into the Screen Map

Open the `.vbs` in Notepad — SAP writes it as UTF-16, so it looks fine in Notepad but spaced
out in some editors. It is a list of lines like:

```vbscript
session.findById("wnd[1]/usr/ctxtSL_BUKRS-LOW").text = "gbkm"
session.findById("wnd[1]/tbar[0]/btn[8]").press
```

Copy the string inside each `findById("…")` into **column F** of the **Screen Map** sheet, next
to the matching key. Column F is the only column the macro reads. Column D says where the
current value came from:

| Source | Meaning |
|---|---|
| `recorded` | Came out of `Audit.vbs`. Should be right for this system. |
| `VERIFY` | A standard SAP ID, but the recording never touched it. Check before an EXTRACT run. |
| `BLOCKED` | Unknown, because `Audit2.vbs` captured nothing. |

Ignore the recording's noise lines: `resizeWorkingPane`, `setFocus`, `caretPosition`,
`sendVKey 4` followed by `wnd[2].close` (that's an F4 browse someone opened and cancelled).
None of them are steps.

## Things a recording will not tell you

- **Whether the date format is what SAP expects.** `Audit.vbs` typed `01092025`, so
  `DDMMYYYY` is known to work here. If you switch it, check `SU3` → *Defaults* →
  *Date format* first. A wrong value means FEBAN searches a period that does not exist and
  reports nothing found, with no error at all.
- **Which house bank and account.** The samples appear to span at least two Citibank
  accounts (see `Data Issues`), so leaving both filters blank is the safer default.
- **What an unlabelled button does.** `Audit.vbs` presses `wnd[0]/tbar[1]/btn[14]` right after
  Execute. That is in the Screen Map as `FEBAN.PostExecuteButton` and it is optional — check
  what it actually does before relying on it, or clear the cell to skip it.
- **Theme differences.** A recording made under a different SAP GUI theme (Blue Crystal vs
  Quartz vs the Fiori visual theme) can produce different container IDs for the same grid.
  Record on the theme you will run on.
