Attribute VB_Name = "modConfig"
'=======================================================================
' modConfig -- reads every setting and every SAP element ID from the
'              workbook, so that adapting the macro to your SAP release
'              is a paste-the-recording exercise, not a code change.
'
' Nothing in this project hardcodes a findById string. If a required ID
' is missing from the 'Screen Map' sheet the run aborts naming the key,
' rather than guessing at an ID and clicking something unintended.
'=======================================================================
Option Explicit

Public Const SHEET_CONTROL As String = "Control"
Public Const SHEET_SCREENMAP As String = "Screen Map"
Public Const SHEET_SAMPLES As String = "Samples"
Public Const SHEET_LOG As String = "Log"

' Row/column geometry of the generated workbook. Kept in one place so a
' layout change is a single edit here.
Public Const SAMPLES_FIRST_ROW As Long = 5
Public Const SCREENMAP_FIRST_ROW As Long = 6
Public Const SCREENMAP_COL_KEY As Long = 2
Public Const SCREENMAP_COL_REQUIRED As Long = 5
Public Const SCREENMAP_COL_VALUE As Long = 6

Private mScreenMap As Object          ' Scripting.Dictionary: key -> element ID
Private mRequiredKeys As Object       ' Scripting.Dictionary: key -> True

'-----------------------------------------------------------------------
' Settings -- looked up by the label in column B of the Control sheet.
' The labels are the contract between that sheet and this module.
'-----------------------------------------------------------------------
Public Function Setting(ByVal label As String) As String
    Dim sheet As Worksheet
    Dim found As Range

    Set sheet = ThisWorkbook.Worksheets(SHEET_CONTROL)
    Set found = sheet.Columns(2).Find(What:=label, LookAt:=xlWhole, MatchCase:=False)

    If found Is Nothing Then
        Err.Raise vbObjectError + 513, "modConfig.Setting", _
                  "Setting '" & label & "' is not on the '" & SHEET_CONTROL & "' sheet. " & _
                  "It may have been renamed -- the macro looks it up by exact label."
    End If

    Setting = Trim$(CStr(found.Offset(0, 1).Value))
End Function

Public Function SettingNumber(ByVal label As String, ByVal fallback As Double) As Double
    Dim raw As String
    raw = Setting(label)
    If Len(raw) = 0 Or Not IsNumeric(raw) Then
        SettingNumber = fallback
    Else
        SettingNumber = CDbl(raw)
    End If
End Function

Public Function SettingIsYes(ByVal label As String) As Boolean
    SettingIsYes = (UCase$(Setting(label)) = "YES")
End Function

' The single switch that decides whether anything leaves SAP.
Public Function IsDryRun() As Boolean
    IsDryRun = (UCase$(Replace(Setting("Run mode"), " ", "")) <> "EXTRACT")
End Function

'-----------------------------------------------------------------------
' Screen map
'-----------------------------------------------------------------------
Public Sub LoadScreenMap()
    Dim sheet As Worksheet
    Dim row As Long
    Dim key As String, value As String, required As String

    Set mScreenMap = CreateObject("Scripting.Dictionary")
    Set mRequiredKeys = CreateObject("Scripting.Dictionary")
    mScreenMap.CompareMode = vbTextCompare
    mRequiredKeys.CompareMode = vbTextCompare

    Set sheet = ThisWorkbook.Worksheets(SHEET_SCREENMAP)

    For row = SCREENMAP_FIRST_ROW To sheet.Cells(sheet.Rows.Count, SCREENMAP_COL_KEY).End(xlUp).Row
        key = Trim$(CStr(sheet.Cells(row, SCREENMAP_COL_KEY).Value))
        If Len(key) > 0 And InStr(key, ".") > 0 Then      ' skip the section headings
            value = Trim$(CStr(sheet.Cells(row, SCREENMAP_COL_VALUE).Value))
            required = UCase$(Trim$(CStr(sheet.Cells(row, SCREENMAP_COL_REQUIRED).Value)))

            If Len(value) > 0 Then mScreenMap(key) = value
            If required = "YES" Then mRequiredKeys(key) = True
        End If
    Next row
End Sub

' Raises with a list of everything still blank, so the operator fixes the
' whole sheet in one pass instead of one key per failed run.
Public Sub AssertScreenMapComplete()
    Dim key As Variant
    Dim missing As String

    If mScreenMap Is Nothing Then LoadScreenMap

    For Each key In mRequiredKeys.Keys
        If Not mScreenMap.Exists(key) Then missing = missing & vbCrLf & "  - " & key
    Next key

    If Len(missing) > 0 Then
        Err.Raise vbObjectError + 514, "modConfig.AssertScreenMapComplete", _
                  "These required element IDs are blank on the '" & SHEET_SCREENMAP & _
                  "' sheet:" & missing & vbCrLf & vbCrLf & _
                  "Record the transaction with Alt+F12 > Script Recording and Playback " & _
                  "and paste the findById strings from the generated .vbs into column F."
    End If
End Sub

' Element ID for a key. Raises if absent -- never returns a guess.
Public Function ElementId(ByVal key As String) As String
    If mScreenMap Is Nothing Then LoadScreenMap

    If Not mScreenMap.Exists(key) Then
        Err.Raise vbObjectError + 515, "modConfig.ElementId", _
                  "No element ID recorded for '" & key & "'. Add it to the '" & _
                  SHEET_SCREENMAP & "' sheet."
    End If

    ElementId = mScreenMap(key)
End Function

Public Function HasElementId(ByVal key As String) As Boolean
    If mScreenMap Is Nothing Then LoadScreenMap
    HasElementId = mScreenMap.Exists(key)
End Function

' Optional IDs: returns "" when the key was left blank, so callers can
' treat a step as 'not configured on this release' and skip it.
Public Function ElementIdOrBlank(ByVal key As String) As String
    If HasElementId(key) Then
        ElementIdOrBlank = ElementId(key)
    Else
        ElementIdOrBlank = vbNullString
    End If
End Function
