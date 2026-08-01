' Updates an existing FEBAN_Audit_Control.xlsm in place:
'   1. applies the Screen Map corrections listed below
'   2. removes the sap-audit modules and re-imports them from .\vba
'
' Everything else is left alone -- your Control settings, the buttons, the
' Probe sheet and the Log all survive. That is the point of updating rather
' than rebuilding: a fresh workbook would discard them.
'
' Step 1 needs nothing special. Step 2 needs Excel's "Trust access to the VBA
' project object model" (File > Options > Trust Center > Trust Center Settings
' > Macro Settings), which is off by default on many managed laptops. If it is
' off, step 1 still runs and the script tells you to import the modules by
' hand -- which is the only part you would have to do manually.
'
' Close the workbook in Excel before running this.
Option Explicit

Dim fso, here, target, excel, book, folder, file
Dim comps, i, removed, imported, name
Dim fixedCells, addedRows, msg, vbaOk, vbaErr

Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)

' --- find the workbook ------------------------------------------------------
target = fso.BuildPath(here, "FEBAN_Audit_Control.xlsm")
If Not fso.FileExists(target) Then
    target = InputBox("Full path to your FEBAN_Audit_Control.xlsm:", "Update", target)
    If Trim(target) = "" Then WScript.Quit 0
    If Not fso.FileExists(target) Then
        MsgBox "Not found:" & vbCrLf & target, 16, "Update"
        WScript.Quit 1
    End If
End If

If Not fso.FolderExists(fso.BuildPath(here, "vba")) Then
    MsgBox "No 'vba' folder next to this script.", 16, "Update"
    WScript.Quit 1
End If

On Error Resume Next
Set excel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    MsgBox "Could not start Excel: " & Err.Description, 16, "Update"
    WScript.Quit 1
End If
Err.Clear
On Error GoTo 0

excel.Visible = False
excel.DisplayAlerts = False

On Error Resume Next
Set book = excel.Workbooks.Open(target)
If Err.Number <> 0 Then
    MsgBox "Could not open the workbook: " & Err.Description & vbCrLf & vbCrLf & _
           "Is it still open in Excel? Close it and run this again.", 16, "Update"
    excel.Quit
    WScript.Quit 1
End If
On Error GoTo 0

' --- 1. Screen Map corrections ---------------------------------------------
' SAP's own column dump on PP2 reported neither ESTAT nor SGTXT. VB1OK is
' 'Update 1 OK' and VWEZW is 'Payment Notes'. The FB03 rows are new: the
' largest payment is now opened by document number rather than by clicking a
' row, because FBL1N renders a classic list whose row positions shift.
fixedCells = 0
addedRows = 0

fixedCells = fixedCells + SetId(book, "FEBAN.Col.Status", "VB1OK")
fixedCells = fixedCells + SetId(book, "FEBAN.Col.Reference", "VWEZW")

addedRows = addedRows + AddId(book, "FB03.DocNumber", "wnd[0]/usr/txtRF05L-BELNR", _
                              "Document number on the FB03 entry screen. VERIFY")
addedRows = addedRows + AddId(book, "FB03.CompanyCode", "wnd[0]/usr/ctxtRF05L-BUKRS", _
                              "Company code on that screen. VERIFY")
addedRows = addedRows + AddId(book, "FB03.FiscalYear", "wnd[0]/usr/txtRF05L-GJAHR", _
                              "Fiscal year on that screen. VERIFY")

' F2 on the FBL1N list lands on the line item; Environment > Payment usage
' lives on the document. N1.vbs bridged that with tbar[1]/btn[9].
addedRows = addedRows + AddId(book, "Doc.OverviewButton", "wnd[0]/tbar[1]/btn[9]", _
                              "Document-overview button. Only used when the payment-usage " & _
                              "menu is missing, so clearing it just skips the attempt")

' 'Payment document type' now takes a list, which is how a batch paid outside
' the normal payment run gets picked up.
SetNote book, "Payment document type", _
        "Only documents of these types are taken from the Payment Usage list and fed to " & _
        "FBL1N. ZP is the SAP standard for a payment document. Takes a comma-separated " & _
        "list -- 'ZP, ZV, KZ' -- for batches paid outside the normal payment run. When " & _
        "nothing matches, the Log names the types the file actually held."

' An invoice document carries more than one attachment; picking row 0 got the
' workflow notes instead of the invoice.
AddSetting book, "Invoice attachment title contains", "Invoice", _
        "An invoice document carries more than one attachment -- on this system the " & _
        "workflow notes sit above the document itself -- so the row whose text contains " & _
        "this word is the one downloaded. The titles come from the archiving system " & _
        "rather than from SAP, so they do not follow the SAP logon language. Nothing " & _
        "matching takes the last row and says so in the Log, naming every attachment."

' The first invoice this system exported came back as document type RN, not
' the KR the setting shipped with, so the KR filter matched nothing. Widen it
' -- but only if the operator has not already put their own value there.
fixedCells = fixedCells + SetValueIfDefault(book, "Invoice document type", "KR", "KR, RN")
SetNote book, "Invoice document type", _
        "CROSS-CHECK ONLY -- this no longer decides anything. The invoice is picked by " & _
        "SIGN: in a vendor line-item list the payment is a debit and the invoice it " & _
        "settles is a credit, so the invoice is the negative row and the biggest invoice " & _
        "is the most negative one. That holds whatever the type is called in this company " & _
        "code, which matters because it is RN here and KR on the SAP standard. If the row " & _
        "picked is not one of the types listed, the Log says so and takes it anyway. " & _
        "Blank to switch the cross-check off."

AddSetting book, "Max rows for a settlement", 8, _
    "A clearing document with no vendor payments and no more than this many " & _
    "bookkeeping rows is a treasury, tax or FX settlement -- there is no invoice " & _
    "behind it. Above it, the run assumes the document-type column was misread. " & _
    "A real payment run here has held between 34 and 656 payments, never a handful."

' --- 1b. Samples sheet: the columns the importer writes ---------------------
' Appended at Q, R and S so the Control sheet's formulas over A:P are
' untouched. Existing rows are stamped with the request they came from --
' the Paper samples, company code from the Control sheet -- so a workbook
' that has already run keeps writing to the same folders.
AddSampleColumn book, 17, "Request", 30
AddSampleColumn book, 18, "Company code", 13
AddSampleColumn book, 19, "Auditor's comment", 46
AddSampleColumn book, 20, "Start document", 16
AddSampleColumn book, 21, "Start at", 12
' A follow-up request shows its working: an SAP extract under each sample
' naming the ZP the auditor picked. Captured so the run can be checked
' against their answer instead of merely trusted. Blank on every row whose
' file showed no working, which is most of them.
AddSampleColumn book, 22, "Auditor's ZP", 14
StampExistingSamples book, "Paper Samples"

' --- 2. modules -------------------------------------------------------------
vbaOk = False
removed = 0
imported = 0

On Error Resume Next
Set comps = book.VBProject.VBComponents
vbaErr = Err.Description
On Error GoTo 0

If Not comps Is Nothing Then
    On Error Resume Next
    For i = comps.Count To 1 Step -1
        name = comps(i).Name
        If IsOurModule(name) Then
            comps.Remove comps(i)
            If Err.Number = 0 Then removed = removed + 1
            Err.Clear
        End If
    Next

    Set folder = fso.GetFolder(fso.BuildPath(here, "vba"))
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "bas" Then
            comps.Import file.Path
            If Err.Number = 0 Then imported = imported + 1
            Err.Clear
        End If
    Next
    On Error GoTo 0
    vbaOk = (imported > 0)
End If

book.Save
book.Close False
excel.Quit

' --- report -----------------------------------------------------------------
msg = "Screen Map:  " & fixedCells & " corrected, " & addedRows & " added" & vbCrLf

If vbaOk Then
    msg = msg & "Modules:     " & removed & " removed, " & imported & " imported" & _
          vbCrLf & vbCrLf & _
          "Open the workbook and run Check setup." & vbCrLf & vbCrLf & _
          "Your Control settings, buttons, Probe sheet and Log were left alone."
    MsgBox msg, 64, "Update"
Else
    msg = msg & "Modules:     NOT updated" & vbCrLf & vbCrLf & _
          "Excel would not let the script touch the VBA project" & _
          IIfStr(Len(vbaErr) > 0, " (" & vbaErr & ")", "") & ". That is almost always " & _
          """Trust access to the VBA project object model"" being switched off." & _
          vbCrLf & vbCrLf & _
          "The Screen Map corrections above DID apply. For the modules: open the " & _
          "workbook, Alt+F11, remove the old mod* modules and import the ones in the " & _
          "vba folder."
    MsgBox msg, 48, "Update"
End If

' ---------------------------------------------------------------------------
Function IsOurModule(n)
    Dim known
    known = "|modChain|modConfig|modExport|modExportRead|modFbl1n|modFeban|" & _
            "modListFile|modDrilldown|modImport|modLog|modMain|modProbe|modReport|" & _
            "modSafety|" & _
            "modSapConnect|modSetup|modUtil|"
    IsOurModule = (InStr(1, known, "|" & n & "|", 1) > 0)
End Function

' Set column F for an existing key. Returns 1 if it changed something.
Function SetId(wb, key, value)
    Dim sheet, r, cell
    SetId = 0
    Set sheet = wb.Worksheets("Screen Map")
    For r = 6 To 200
        If Trim(CStr(sheet.Cells(r, 2).Value)) = key Then
            If Trim(CStr(sheet.Cells(r, 6).Value)) <> value Then
                sheet.Cells(r, 6).Value = value
                SetId = 1
            End If
            Exit Function
        End If
    Next
End Function

' Add a key that is not there yet, just below the last existing key so it
' stays inside the block rather than under the explanatory notes.
Function AddId(wb, key, value, note)
    Dim sheet, r, lastKey
    AddId = 0
    Set sheet = wb.Worksheets("Screen Map")

    lastKey = 0
    For r = 6 To 200
        If Trim(CStr(sheet.Cells(r, 2).Value)) = key Then Exit Function   ' already there
        If InStr(CStr(sheet.Cells(r, 2).Value), ".") > 0 And _
           Len(Trim(CStr(sheet.Cells(r, 6).Value))) > 0 Then lastKey = r
    Next
    If lastKey = 0 Then Exit Function

    sheet.Rows(lastKey + 1).Insert
    sheet.Cells(lastKey + 1, 2).Value = key
    sheet.Cells(lastKey + 1, 3).Value = note
    sheet.Cells(lastKey + 1, 4).Value = "VERIFY"
    sheet.Cells(lastKey + 1, 5).Value = "No"
    sheet.Cells(lastKey + 1, 6).Value = value
    AddId = 1
End Function

' Change a Control VALUE, but only when it is still the value this project
' shipped. Anything the operator has typed themselves is left alone.
Function SetValueIfDefault(wb, settingName, oldDefault, newDefault)
    Dim sheet, r
    SetValueIfDefault = 0
    Set sheet = wb.Worksheets("Control")
    For r = 1 To 200
        If Trim(CStr(sheet.Cells(r, 2).Value)) = settingName Then
            If Trim(CStr(sheet.Cells(r, 3).Value)) = oldDefault Then
                sheet.Cells(r, 3).Value = newDefault
                SetValueIfDefault = 1
            End If
            Exit Function
        End If
    Next
End Function

' Replace the Notes cell of a Control setting, leaving its VALUE alone -- the
' operator's own settings must survive an update.
Sub SetNote(wb, settingName, note)
    Dim sheet, r
    Set sheet = wb.Worksheets("Control")
    For r = 1 To 200
        If Trim(CStr(sheet.Cells(r, 2).Value)) = settingName Then
            sheet.Cells(r, 4).Value = note
            Exit Sub
        End If
    Next
End Sub

' Add a Control setting below the last existing one, if it is not there yet.
Sub AddSetting(wb, settingName, value, note)
    Dim sheet, r, lastRow
    Set sheet = wb.Worksheets("Control")
    lastRow = 0
    For r = 1 To 200
        If Trim(CStr(sheet.Cells(r, 2).Value)) = settingName Then Exit Sub   ' already there
        If Len(Trim(CStr(sheet.Cells(r, 2).Value))) > 0 And _
           Len(Trim(CStr(sheet.Cells(r, 4).Value))) > 0 Then lastRow = r
    Next
    If lastRow = 0 Then Exit Sub

    sheet.Rows(lastRow + 1).Insert
    sheet.Cells(lastRow + 1, 2).Value = settingName
    sheet.Cells(lastRow + 1, 3).Value = value
    sheet.Cells(lastRow + 1, 4).Value = note
End Sub

' Add a heading to the Samples sheet if the column is not already there.
Sub AddSampleColumn(wb, col, title, width)
    Dim sheet
    Set sheet = wb.Worksheets("Samples")
    If Trim(CStr(sheet.Cells(4, col).Value)) = title Then Exit Sub
    sheet.Cells(4, col).Value = title
    sheet.Cells(4, col).Font.Bold = True
    sheet.Columns(col).ColumnWidth = width
End Sub

' Give rows that predate the importer a request name and the Control
' sheet's company code, so they keep behaving exactly as before.
Sub StampExistingSamples(wb, requestName)
    Dim sheet, r, lastRow, companyCode
    Set sheet = wb.Worksheets("Samples")

    companyCode = ""
    For r = 1 To 200
        If Trim(CStr(wb.Worksheets("Control").Cells(r, 2).Value)) = "Company code" Then
            companyCode = Trim(CStr(wb.Worksheets("Control").Cells(r, 3).Value))
            Exit For
        End If
    Next

    lastRow = sheet.Cells(sheet.Rows.Count, 5).End(-4162).Row   ' xlUp
    For r = 5 To lastRow
        If IsDate(sheet.Cells(r, 5).Value) Then
            If Trim(CStr(sheet.Cells(r, 17).Value)) = "" Then
                sheet.Cells(r, 17).Value = requestName
            End If
            If Trim(CStr(sheet.Cells(r, 18).Value)) = "" Then
                sheet.Cells(r, 18).Value = companyCode
            End If
        End If
    Next
End Sub

Function IIfStr(cond, a, b)
    If cond Then IIfStr = a Else IIfStr = b
End Function
