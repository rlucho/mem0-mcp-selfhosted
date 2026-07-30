# Getting to a first test

Excel will not keep macros in a `.xlsx`, and a `.bas` file is just text until it is
imported — so there are two short steps before anything runs.

## 1. Make the workbook macro-enabled

Open `FEBAN_Audit_Control.xlsx` → **File ▸ Save As** → set the type to
**Excel Macro-Enabled Workbook (*.xlsm)** → save it as `FEBAN_Audit_Control.xlsm`.

## 2. Import the modules

`Alt`+`F11` to open the VBA editor, then **File ▸ Import File…** and pick each `.bas`
in the `vba` folder. There are twelve. Import all of them.

If `Scripting.Dictionary` fails to resolve when you run something, add the reference:
**Tools ▸ References ▸ Microsoft Scripting Runtime**.

*Shortcut:* `build_xlsm.vbs` in this folder does both steps for you — double-click it.
It only works if **File ▸ Options ▸ Trust Center ▸ Trust Center Settings ▸ Macro Settings
▸ Trust access to the VBA project object model** is ticked, which it often is not on a
managed laptop. If it errors, just do steps 1 and 2 by hand; it takes about two minutes.

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
