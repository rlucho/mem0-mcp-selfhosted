Attribute VB_Name = "modSetup"
'=======================================================================
' modSetup -- puts buttons on the Control sheet.
'
' The workbook cannot ship with buttons already on it, because a .xlsx
' carries no macros for a button to call and the .xlsm has to be made on a
' machine that has Excel. So this builds them instead: run AddButtons once
' after importing the modules, save, and the sheet has its own toolbar
' from then on.
'
' Nothing here touches SAP.
'=======================================================================
Option Explicit

Private Const BUTTON_TOP As Double = 12
Private Const BUTTON_LEFT As Double = 640
Private Const BUTTON_WIDTH As Double = 210
Private Const BUTTON_HEIGHT As Double = 30
Private Const BUTTON_GAP As Double = 6
Private Const TAG_PREFIX As String = "sapaudit_"

'-----------------------------------------------------------------------
' Parameterless, so it still appears in Excel's Macro dialog -- a Sub that
' takes arguments does not.
Public Sub AddButtons()
    Build False
End Sub

' Same, without the confirmation dialog. build_xlsm.vbs calls this, because an
' unattended build must not stop on a message box.
Public Sub AddButtonsQuiet()
    Build True
End Sub

Private Sub Build(ByVal quiet As Boolean)
    Dim sheet As Worksheet
    Dim top As Double
    Dim built As Long

    On Error GoTo Failed

    Set sheet = ThisWorkbook.Worksheets(modConfig.SHEET_CONTROL)
    RemoveButtons

    top = BUTTON_TOP

    AddOne sheet, top, "0.  Import request", "modImport.ImportRequest", built
    AddOne sheet, top, "1.  Check setup", "modMain.CheckSetup", built
    AddOne sheet, top, "2.  Probe this SAP screen", "modProbe.ProbeCurrentScreen", built
    AddOne sheet, top, "3.  Dump FEBAN grid columns", "modFeban.DumpGridColumns", built
    AddOne sheet, top, "3b. FBL1N columns (see note)", "modFbl1n.DumpFbl1nColumns", built
    AddOne sheet, top, "4.  Run one month", "modMain.RunSingleMonth", built
    AddOne sheet, top, "5.  Run all samples", "modMain.RunExtract", built
    AddOne sheet, top, "     Dump this SAP screen", "modProbe.DumpScreen", built
    AddOne sheet, top, "     Clear the log", "modLog.ClearLog", built

    sheet.Activate
    If Not quiet Then
        MsgBox built & " buttons added to the '" & modConfig.SHEET_CONTROL & "' sheet." & _
               vbCrLf & vbCrLf & "Save the workbook so they persist.", _
               vbInformation, "Add buttons"
    End If
    Exit Sub

Failed:
    If Not quiet Then
        MsgBox "Could not add the buttons: " & Err.Description, vbExclamation, "Add buttons"
    End If
End Sub

Private Sub AddOne(ByVal sheet As Worksheet, ByRef top As Double, ByVal caption As String, _
                   ByVal macroName As String, ByRef built As Long)
    Dim shape As Object

    Set shape = sheet.Buttons.Add(BUTTON_LEFT, top, BUTTON_WIDTH, BUTTON_HEIGHT)
    shape.Name = TAG_PREFIX & built
    shape.Caption = caption
    shape.OnAction = macroName

    ' Buttons that move or resize with the cells under them end up in odd
    ' places the first time someone widens a column.
    On Error Resume Next
    shape.Placement = 3               ' xlFreeFloating
    On Error GoTo 0

    top = top + BUTTON_HEIGHT + BUTTON_GAP
    built = built + 1
End Sub

'-----------------------------------------------------------------------
' Only removes buttons this module made -- anything else on the sheet is
' left alone, so running AddButtons twice does not destroy someone's work.
'-----------------------------------------------------------------------
Public Sub RemoveButtons()
    Dim sheet As Worksheet
    Dim i As Long

    Set sheet = ThisWorkbook.Worksheets(modConfig.SHEET_CONTROL)

    On Error Resume Next
    For i = sheet.Buttons.Count To 1 Step -1
        If Left$(sheet.Buttons(i).Name, Len(TAG_PREFIX)) = TAG_PREFIX Then
            sheet.Buttons(i).Delete
        End If
    Next i
    On Error GoTo 0
End Sub
