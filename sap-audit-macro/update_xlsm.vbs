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
            "modListFile|modDrilldown|modLog|modMain|modProbe|modSafety|" & _
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

Function IIfStr(cond, a, b)
    If cond Then IIfStr = a Else IIfStr = b
End Function
