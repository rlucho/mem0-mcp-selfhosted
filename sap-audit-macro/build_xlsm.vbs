' Builds FEBAN_Audit_Control.xlsm from the .xlsx and the .bas modules beside it,
' then puts the buttons on the Control sheet.
'
' This REPLACES any existing .xlsm, so anything already filled in on its sheets
' is lost. To keep a workbook you have been using, run update_xlsm.vbs instead:
' it refreshes the modules and the Screen Map in place and leaves your Control
' settings, buttons, Probe sheet and Log alone.
'
' Needs Excel's "Trust access to the VBA project object model" to be enabled
' (File > Options > Trust Center > Trust Center Settings > Macro Settings).
' That is off by default on many managed laptops -- if this errors out, do it
' by hand instead, which takes about two minutes. See SETUP.md.
Option Explicit

Dim fso, shell, here, xlsx, xlsm, excel, book, folder, file, imported
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

here = fso.GetParentFolderName(WScript.ScriptFullName)
xlsx = fso.BuildPath(here, "FEBAN_Audit_Control.xlsx")
xlsm = fso.BuildPath(here, "FEBAN_Audit_Control.xlsm")

If Not fso.FileExists(xlsx) Then
    MsgBox "FEBAN_Audit_Control.xlsx is not next to this script.", 16, "Build"
    WScript.Quit 1
End If

If fso.FileExists(xlsm) Then
    If MsgBox(xlsm & vbCrLf & vbCrLf & "already exists. Overwrite it?" & vbCrLf & _
              "Anything you have filled in on its sheets will be lost.", _
              vbExclamation + vbYesNo + vbDefaultButton2, "Build") <> vbYes Then
        WScript.Quit 0
    End If
    fso.DeleteFile xlsm, True
End If

On Error Resume Next
Set excel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    MsgBox "Could not start Excel: " & Err.Description, 16, "Build"
    WScript.Quit 1
End If
On Error GoTo 0

excel.Visible = False
excel.DisplayAlerts = False

Set book = excel.Workbooks.Open(xlsx)
book.SaveAs xlsm, 52          ' 52 = xlOpenXMLWorkbookMacroEnabled

Set folder = fso.GetFolder(fso.BuildPath(here, "vba"))
imported = 0

On Error Resume Next
For Each file In folder.Files
    If LCase(fso.GetExtensionName(file.Name)) = "bas" Then
        book.VBProject.VBComponents.Import file.Path
        If Err.Number = 0 Then
            imported = imported + 1
        Else
            excel.DisplayAlerts = True
            book.Close False
            excel.Quit
            MsgBox "Could not import " & file.Name & ":" & vbCrLf & vbCrLf & _
                   Err.Description & vbCrLf & vbCrLf & _
                   "This is almost always 'Trust access to the VBA project object " & _
                   "model' being switched off. Import the modules by hand instead " & _
                   "-- see SETUP.md.", 48, "Build"
            WScript.Quit 1
        End If
    End If
Next
On Error GoTo 0

' Put the buttons on the Control sheet too. Trust access is already proven
' by the imports above, so this cannot fail for that reason -- and it saves
' rebuilding them by hand every time the modules are refreshed.
Dim buttonsAdded
buttonsAdded = False
On Error Resume Next
excel.Run "modSetup.AddButtonsQuiet"
buttonsAdded = (Err.Number = 0)
Err.Clear
On Error GoTo 0

book.Save
book.Close False
excel.Quit

MsgBox "Built:" & vbCrLf & xlsm & vbCrLf & vbCrLf & _
       imported & " modules imported." & vbCrLf & _
       IIfStr(buttonsAdded, "Buttons added to the Control sheet.", _
              "Buttons NOT added -- run modSetup.AddButtons yourself.") & vbCrLf & vbCrLf & _
       "Open it, log on to SAP, then press '1. Check setup'.", 64, "Build"

Function IIfStr(cond, a, b)
    If cond Then IIfStr = a Else IIfStr = b
End Function
