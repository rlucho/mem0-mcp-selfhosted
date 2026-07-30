Attribute VB_Name = "modProbe"
'=======================================================================
' modProbe -- verify predicted element IDs without recording anything.
'
' Steps 8 to 10 are written from standard SAP rather than from a
' recording, so their Screen Map rows are predictions. Rather than asking
' for another Alt+F12 session, this checks them against the live screen:
' put SAP on the screen in question, run ProbeCurrentScreen, and it
' reports which of the predicted IDs actually exist, what type of control
' each one is, and -- for the ones that are missing -- what else is on the
' screen that looks like a plausible substitute.
'
' It reads. It presses nothing and types nothing, so it is safe to run on
' any screen, including in a production client.
'
' Typical use:
'   1. enter FBL1N by hand, stop on the selection screen
'   2. run ProbeCurrentScreen -> tells you which Fbl1n.* predictions hold
'   3. paste the corrections into column F of the Screen Map
'   4. execute FBL1N by hand, then run ProbeCurrentScreen again for the
'      result-grid rows, and modFbl1n.DumpGridColumns for its columns
'=======================================================================
Option Explicit

'-----------------------------------------------------------------------
' Check every mapped ID against whatever is on screen right now.
'-----------------------------------------------------------------------
Public Sub ProbeCurrentScreen()
    Dim sheet As Worksheet
    Dim row As Long
    Dim key As String, elementId As String
    Dim present As String, absent As String, blank As String
    Dim found As Long, missing As Long
    Dim report As String

    On Error GoTo Failed

    modConfig.LoadScreenMap
    modSapConnect.SapAttach

    Set sheet = ThisWorkbook.Worksheets(modConfig.SHEET_SCREENMAP)

    For row = modConfig.SCREENMAP_FIRST_ROW To _
              sheet.Cells(sheet.Rows.Count, modConfig.SCREENMAP_COL_KEY).End(xlUp).Row

        key = Trim$(CStr(sheet.Cells(row, modConfig.SCREENMAP_COL_KEY).Value))
        If Len(key) > 0 And InStr(key, ".") > 0 Then
            elementId = Trim$(CStr(sheet.Cells(row, modConfig.SCREENMAP_COL_VALUE).Value))

            If Len(elementId) = 0 Then
                blank = blank & "  " & key & vbCrLf
            ElseIf Not LooksLikeElementId(elementId) Then
                ' A grid column name or a function code, not an element path --
                ' nothing to probe, those are checked by DumpGridColumns.
            ElseIf modSapConnect.Exists(elementId) Then
                present = present & "  " & key & "  [" & ControlType(elementId) & "]" & vbCrLf
                found = found + 1
            Else
                absent = absent & "  " & key & vbCrLf & "      " & elementId & vbCrLf
                missing = missing + 1
            End If
        End If
    Next row

    report = "Probe of " & modSapConnect.CurrentTransaction() & _
             " on " & modSapConnect.gSystemId & vbCrLf & String$(58, "-") & vbCrLf & vbCrLf

    If Len(present) > 0 Then
        report = report & "ON THIS SCREEN (" & found & ")" & vbCrLf & present & vbCrLf
    End If
    If Len(absent) > 0 Then
        report = report & "NOT ON THIS SCREEN (" & missing & ")" & vbCrLf & _
                 "Expected if they belong to another screen. If one of these SHOULD be " & _
                 "here, the prediction is wrong -- run DumpScreen to see what is." & _
                 vbCrLf & absent & vbCrLf
    End If
    If Len(blank) > 0 Then
        report = report & "STILL BLANK ON THE SCREEN MAP" & vbCrLf & blank
    End If

    Debug.Print report
    modLog.LogAction 0, "Probe", _
                 "Probed " & modSapConnect.CurrentTransaction() & ": " & found & _
                 " of the mapped IDs are on this screen, " & missing & " are not.", _
                 "OK", vbNullString

    ShowLongMessage report, "Probe screen"
    Exit Sub

Failed:
    MsgBox "Probe failed: " & Err.Description, vbExclamation, "Probe screen"
End Sub

'-----------------------------------------------------------------------
' Every named element on the current screen, so a wrong prediction can be
' replaced with the real thing. This is the recording substitute.
'-----------------------------------------------------------------------
Public Sub DumpScreen()
    Dim report As String
    Dim windowIndex As Long

    On Error GoTo Failed

    modSapConnect.SapAttach

    report = "Elements on " & modSapConnect.CurrentTransaction() & vbCrLf & _
             String$(58, "-") & vbCrLf

    For windowIndex = 0 To 2
        If modSapConnect.Exists("wnd[" & windowIndex & "]") Then
            report = report & vbCrLf & "=== wnd[" & windowIndex & "] " & _
                     modSapConnect.Element("wnd[" & windowIndex & "]").Text & " ===" & vbCrLf
            Describe modSapConnect.Element("wnd[" & windowIndex & "]"), 0, report
        End If
    Next windowIndex

    Debug.Print report
    ShowLongMessage report, "Dump screen"
    Exit Sub

Failed:
    MsgBox "Dump failed: " & Err.Description, vbExclamation, "Dump screen"
End Sub

' Walk the control tree. Only input-ish and pressable controls are listed --
' a full dump of every label is unreadable and not what anyone needs.
Private Sub Describe(ByVal control As Object, ByVal depth As Long, ByRef report As String)
    Dim child As Object
    Dim kind As String
    Dim caption As String

    If depth > 12 Then Exit Sub

    On Error Resume Next
    kind = control.Type
    On Error GoTo 0

    Select Case kind
        Case "GuiTextField", "GuiCTextField", "GuiPasswordField", "GuiComboBox", _
             "GuiCheckBox", "GuiRadioButton", "GuiButton", "GuiShell", _
             "GuiGridView", "GuiTableControl", "GuiTree"
            caption = vbNullString
            On Error Resume Next
            caption = control.Text
            If Len(caption) = 0 Then caption = control.Tooltip
            On Error GoTo 0

            report = report & "  " & kind & vbTab & control.Id & _
                     IIf(Len(caption) > 0, vbTab & """" & Left$(caption, 40) & """", "") & vbCrLf
    End Select

    On Error Resume Next
    For Each child In control.Children
        Describe child, depth + 1, report
    Next child
    On Error GoTo 0
End Sub

'-----------------------------------------------------------------------
Private Function LooksLikeElementId(ByVal value As String) As Boolean
    ' Element paths start with a window; grid column names and OK-codes do not.
    LooksLikeElementId = (Left$(value, 4) = "wnd[")
End Function

Private Function ControlType(ByVal elementId As String) As String
    On Error Resume Next
    ControlType = modSapConnect.Element(elementId).Type
    On Error GoTo 0
    If Len(ControlType) = 0 Then ControlType = "?"
End Function

' MsgBox truncates around 1024 characters, and these reports run longer.
' Write the whole thing to a sheet and show the head of it.
Private Sub ShowLongMessage(ByVal text As String, ByVal caption As String)
    Dim sheet As Worksheet
    Dim lines() As String
    Dim i As Long

    On Error Resume Next
    Set sheet = ThisWorkbook.Worksheets("Probe")
    On Error GoTo 0

    If sheet Is Nothing Then
        Set sheet = ThisWorkbook.Worksheets.Add
        sheet.Name = "Probe"
    End If

    sheet.Cells.Clear
    lines = Split(text, vbCrLf)
    For i = LBound(lines) To UBound(lines)
        sheet.Cells(i + 1, 1).Value = "'" & lines(i)
    Next i
    sheet.Columns(1).AutoFit
    sheet.Activate

    MsgBox Left$(text, 900) & vbCrLf & vbCrLf & _
           "(full report on the 'Probe' sheet and in the Immediate window, Ctrl+G)", _
           vbInformation, caption
End Sub
