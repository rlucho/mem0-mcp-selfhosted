Attribute VB_Name = "modSapConnect"
'=======================================================================
' modSapConnect -- attaches to an already-running SAP GUI session.
'
' This module never logs on and never starts SAP. You log on yourself,
' with your own credentials, and the macro borrows the session you are
' already sitting in. That keeps the extract attributable to a named
' person in SAP's own security audit log, which is what an auditor will
' want to see anyway.
'=======================================================================
Option Explicit

Public gSession As Object       ' GuiSession
Public gConnection As Object    ' GuiConnection

' Session identity, captured once at attach time and written into every
' log row so the extract is self-describing.
Public gSystemId As String
Public gClient As String
Public gSapUser As String
Public gSapRelease As String

'-----------------------------------------------------------------------
' Attach to the running session and verify it is the system we expect.
'-----------------------------------------------------------------------
Public Sub SapAttach()
    Dim guiAuto As Object
    Dim engine As Object

    ' Two ROT names are in use. recordings/Audit.vbs on this system emits
    ' SAPGUISERVER, so try that first and fall back to the more common SAPGUI.
    On Error Resume Next
    Set guiAuto = GetObject("SAPGUISERVER")
    If guiAuto Is Nothing Then Set guiAuto = GetObject("SAPGUI")
    On Error GoTo 0

    If guiAuto Is Nothing Then
        Err.Raise vbObjectError + 520, "modSapConnect.SapAttach", _
                  "No running SAP GUI found. Log on to SAP first, then run the macro." & _
                  vbCrLf & vbCrLf & _
                  "If you are logged on and still see this, scripting is switched off. " & _
                  "It needs sapgui/user_scripting = TRUE on the application server and " & _
                  "'Enable scripting' ticked in SAP GUI Options > Accessibility & Scripting."
    End If

    On Error Resume Next
    Set engine = guiAuto.GetScriptingEngine
    On Error GoTo 0

    If engine Is Nothing Then
        Err.Raise vbObjectError + 521, "modSapConnect.SapAttach", _
                  "SAP GUI is running but would not hand out its scripting engine. " & _
                  "Ask Basis to confirm sapgui/user_scripting is TRUE on this server."
    End If

    If engine.Children.Count = 0 Then
        Err.Raise vbObjectError + 522, "modSapConnect.SapAttach", _
                  "SAP GUI is running but has no open connection."
    End If

    Set gConnection = engine.Children(0)
    If gConnection.Children.Count = 0 Then
        Err.Raise vbObjectError + 523, "modSapConnect.SapAttach", _
                  "The SAP connection has no open session."
    End If

    Set gSession = gConnection.Children(0)

    ' Some landscapes pop 'a script is attaching' or 'a script is opening a
    ' connection'. If either is on, scripting still works but every step
    ' waits for a human click, which defeats an unattended run.
    On Error Resume Next
    gSession.TestToolMode = 1        ' faster, and suppresses per-step redraw
    On Error GoTo 0

    CaptureSessionIdentity
    AssertExpectedSystem
End Sub

Private Sub CaptureSessionIdentity()
    On Error Resume Next
    gSystemId = gSession.Info.SystemName
    gClient = gSession.Info.Client
    gSapUser = gSession.Info.User
    gSapRelease = gSession.Info.SapRelease
    On Error GoTo 0
End Sub

'-----------------------------------------------------------------------
' Refuse to run against a system other than the one on the Control sheet.
' An audit extract taken from the wrong client is worse than no extract,
' because it looks legitimate.
'-----------------------------------------------------------------------
Private Sub AssertExpectedSystem()
    Dim expectedSid As String
    Dim expectedClient As String

    expectedSid = modConfig.Setting("Expected SAP system ID (SID)")
    expectedClient = modConfig.Setting("Expected SAP client")

    If Len(expectedSid) > 0 Then
        If UCase$(Trim$(gSystemId)) <> UCase$(Trim$(expectedSid)) Then
            Err.Raise vbObjectError + 524, "modSapConnect.AssertExpectedSystem", _
                      "Attached session is system '" & gSystemId & "' but the Control " & _
                      "sheet expects '" & expectedSid & "'." & vbCrLf & vbCrLf & _
                      "Switch to the right system, or correct the Control sheet if you " & _
                      "genuinely mean to extract from this one."
        End If
    End If

    If Len(expectedClient) > 0 Then
        If Trim$(gClient) <> Trim$(expectedClient) Then
            Err.Raise vbObjectError + 525, "modSapConnect.AssertExpectedSystem", _
                      "Attached session is client '" & gClient & "' but the Control " & _
                      "sheet expects '" & expectedClient & "'."
        End If
    End If
End Sub

'-----------------------------------------------------------------------
' Screen helpers
'-----------------------------------------------------------------------
Public Function CurrentTransaction() As String
    On Error Resume Next
    CurrentTransaction = gSession.Info.Transaction
    On Error GoTo 0
End Function

Public Function StatusBarType() As String
    ' "S" success, "W" warning, "E" error, "A" abort, "" nothing shown
    On Error Resume Next
    StatusBarType = gSession.findById("wnd[0]/sbar").MessageType
    On Error GoTo 0
End Function

Public Function StatusBarText() As String
    On Error Resume Next
    StatusBarText = gSession.findById("wnd[0]/sbar").Text
    On Error GoTo 0
End Function

' True when a modal window is up. Anything unexpected here stops the run
' rather than being dismissed blind -- a blind Enter on an unknown popup
' is exactly how a 'read-only' script ends up posting something.
Public Function ModalWindowOpen() As Boolean
    Dim probe As Object
    On Error Resume Next
    Set probe = gSession.findById("wnd[1]")
    On Error GoTo 0
    ModalWindowOpen = Not (probe Is Nothing)
End Function

Public Function ModalWindowTitle() As String
    On Error Resume Next
    ModalWindowTitle = gSession.findById("wnd[1]").Text
    On Error GoTo 0
End Function

'-----------------------------------------------------------------------
' A cheap fingerprint of the screen, so a navigation step can prove it
' actually navigated.
'
' F2 on a classic list only drills in when the cursor sits on a hotspot.
' When it does not, SAP does nothing at all -- no popup, no status message,
' no error to catch -- and the run carries on believing it has moved. That
' is how one sample exported the FBL1N list twice and then reported it had
' found no invoices in it.
'
' Title plus user-area shape plus the first line of text is enough to tell
' 'a different screen' from 'the same screen'. It says nothing about WHICH
' screen, which is the point: no wording, no language.
'-----------------------------------------------------------------------
Public Function ScreenSignature() As String
    Dim area As Object, child As Object
    Dim children As Long
    Dim firstText As String

    On Error Resume Next
    ScreenSignature = gSession.findById("wnd[0]").Text

    Set area = gSession.findById("wnd[0]/usr")
    If Not area Is Nothing Then
        children = area.Children.Count
        For Each child In area.Children
            firstText = child.Text
            If Len(Trim$(firstText)) > 0 Then Exit For
        Next child
    End If
    On Error GoTo 0

    ScreenSignature = ScreenSignature & "|" & children & "|" & Left$(firstText, 60)
End Function

' Does an element exist on the current screen?
Public Function Exists(ByVal elementId As String) As Boolean
    Dim probe As Object
    On Error Resume Next
    Set probe = gSession.findById(elementId)
    On Error GoTo 0
    Exists = Not (probe Is Nothing)
End Function

Public Function Element(ByVal elementId As String) As Object
    Dim result As Object
    On Error Resume Next
    Set result = gSession.findById(elementId)
    On Error GoTo 0

    If result Is Nothing Then
        Err.Raise vbObjectError + 526, "modSapConnect.Element", _
                  "Element not found on the current screen:" & vbCrLf & "  " & elementId & _
                  vbCrLf & vbCrLf & "Current transaction: " & CurrentTransaction() & _
                  ". The recorded ID may belong to a different screen or release."
    End If

    Set Element = result
End Function

' Wait for SAP to finish the previous round trip. GuiSession exposes Busy
' on most releases; where it does not, fall back to a bounded sleep loop.
Public Sub WaitForSap()
    Dim waited As Double
    Dim limit As Double
    Dim isBusy As Boolean

    limit = modConfig.SettingNumber("Max seconds to wait per screen", 60)

    Do
        isBusy = False
        On Error Resume Next
        isBusy = gSession.Busy
        On Error GoTo 0

        If Not isBusy Then Exit Do

        modUtil.SleepSeconds 0.2
        waited = waited + 0.2

        If waited >= limit Then
            Err.Raise vbObjectError + 527, "modSapConnect.WaitForSap", _
                      "SAP was still busy after " & limit & " seconds. Raise " & _
                      "'Max seconds to wait per screen' on the Control sheet, or " & _
                      "check the system."
        End If
    Loop
End Sub

Public Sub SapDetach()
    Set gSession = Nothing
    Set gConnection = Nothing
End Sub
