# Getting to a first test

## Fastest path: `build_xlsm.vbs`

Unzip everything to one folder and double-click **`build_xlsm.vbs`**. It makes
`FEBAN_Audit_Control.xlsm` from the `.xlsx`, imports all thirteen modules, and puts the
buttons on the Control sheet. Then open it, log on to SAP, and press **1. Check setup**.

It needs Excel's **Trust access to the VBA project object model**
(File ▸ Options ▸ Trust Center ▸ Trust Center Settings ▸ Macro Settings). That is off by
default on many managed laptops. If the script says so, do it by hand — it takes two minutes:

1. Open `FEBAN_Audit_Control.xlsx` → **Save As** → **Excel Macro-Enabled Workbook (.xlsm)**
2. `Alt`+`F11` → **File ▸ Import File…** → every `.bas` in `vba` (thirteen of them)
3. Run `modSetup.AddButtons` once, then save

If `Scripting.Dictionary` will not resolve, add **Tools ▸ References ▸ Microsoft Scripting
Runtime**.

## Updating a workbook you have already been using

`build_xlsm.vbs` **replaces** the .xlsm, so anything filled in on its sheets is lost. Once you
have a workbook you are actually working in, use **`update_xlsm.vbs`** instead. It refreshes
the modules and applies any Screen Map corrections in place, and leaves your Control settings,
buttons, Probe sheet and Log alone. The Screen Map half of it works even when trust access is
switched off.

## 3. Run CheckSetup

Log on to SAP first — the macro attaches to the session you are already in, it never logs
on itself.

In the VBA editor, open `modMain`, click anywhere inside `Public Sub CheckSetup()`, press
`F5`.

It reads only. It enters no transaction, presses no button and writes no file. What it
does:

- loads the `Screen Map` sheet and stops if a required row is blank, naming which
- attaches to your SAP session
- **refuses to continue if the SID is not the one on the Control sheet** — so an audit
  extract cannot be taken from the wrong system by accident
- reports the SID, client, user, SAP release and current transaction
- prints **how it would type today's date**

That last line is the one to check. The Control sheet says `DDMMYYYY`, so today should
come out as eight digits with no separators, matching what your recordings typed. If it
does not match what SAP expects from your user, FEBAN searches a period that does not
exist and reports nothing found — with no error at all. That is the failure mode most
likely to waste an afternoon.

## 4. Then, in order

| Macro | What it does |
|---|---|
| `modProbe.ProbeCurrentScreen` | put SAP on any screen; reports which mapped IDs are really there. Read-only. Use it on the FEBAN selection popup and the FBL1N selection screen to clear the `VERIFY` rows. |
| `modFeban.DumpGridColumns` | with a FEBAN result list on screen, prints the grid's real column names |
| `modMain.RunSingleMonth` → `Sep 25`, Run mode still **DRY RUN** | walks every screen for that month's 7 samples and logs what it *would* export, writing nothing |
| `modMain.RunSingleMonth` → `Sep 25`, Run mode **EXTRACT** | the real thing, one month |
| `modMain.RunExtract` | all 56 |

Run mode is on the `Control` sheet and starts at **DRY RUN**. Leave it there until a dry
run comes back clean.

## What to send me after the dry run

The `Log` sheet. It records every step, what it found, and which column it picked out of
each export — that is what tells us whether the remaining `VERIFY` rows are right without
you having to check them one by one.
