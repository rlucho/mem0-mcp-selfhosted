Attribute VB_Name = "modSafety"
'=======================================================================
' modSafety -- keeps the run to reading and exporting.
'
' READ THIS BEFORE TRUSTING IT.
'
' This module is defence in depth, not a guarantee. It can only block
' what it is asked to block: an allowlisted transaction, an allowlisted
' OK-code, a button whose text does not look like a write action. It
' cannot stop a recorded ID that happens to be a Post button, and it
' cannot undo anything.
'
' The control that actually makes this safe is the SAP authorisation
' profile of the user ID you are logged on as. Ask Basis for a
' display-only role for the extract -- FEBAN in particular is a
' post-processing transaction (see README), so a user who can run it
' normally can also post with it.
'=======================================================================
Option Explicit

' Transactions the run is allowed to enter. Everything else aborts.
' FEBAN is here because the audit request names it, not because it is
' display-only.
Private Const ALLOWED_TCODES As String = _
    "FEBAN|FEBA_BANK_STATEMENT|FF.6|FF_6|FB03|FBL1N|FBL3N|FAGLL03|FBL5N|OAOR|BD87"

' OK-codes that commit something. Matched whole, case-insensitively.
Private Const BLOCKED_OKCODES As String = _
    "BU|SICH|POST|SAVE|DELE|LOES|LOSC|STOR|PARK|HALT|FBRA|FBR2|BUCH|AB|ABBR_SAVE"

' Substrings that mark a write action in a button tooltip or menu entry.
Private Const BLOCKED_TEXT_FRAGMENTS As String = _
    "POST|SAVE|SICHERN|BUCHEN|DELETE|REVERSE|CHANGE|EDIT|PARK|HOLD|RESET|CLEAR|RELEASE"

' Function keys that commit on standard SAP screens. 11 is Save.
Private Const BLOCKED_VKEYS As String = "11"

' Modal windows the run knows how to handle. Anything else stops the run
' instead of being dismissed blind -- a blind Enter on an unrecognised
' popup is exactly how a 'read-only' script ends up committing something.
Private Const KNOWN_POPUP_TITLES As String = _
    "Save list in file|Select Spreadsheet|Export list object to XXL|" & _
    "Information|Attachment list|Service: Attachment list|" & _
    "Business Document Navigator|Save As|Directory Browse"

Public gWriteAttemptBlocked As Long     ' counted, and surfaced in the summary

'-----------------------------------------------------------------------
' Transaction guard
'-----------------------------------------------------------------------
Public Sub AssertTcodeAllowed(ByVal tcode As String)
    Dim normalised As String
    normalised = UCase$(Trim$(tcode))

    If Not InList(normalised, ALLOWED_TCODES) Then
        gWriteAttemptBlocked = gWriteAttemptBlocked + 1
        Err.Raise vbObjectError + 530, "modSafety.AssertTcodeAllowed", _
                  "Transaction '" & tcode & "' is not on the read-only allowlist." & _
                  vbCrLf & vbCrLf & "Allowed: " & Replace(ALLOWED_TCODES, "|", ", ") & _
                  vbCrLf & vbCrLf & "If you genuinely need another transaction, add it to " & _
                  "ALLOWED_TCODES in modSafety and note why in the commit message."
    End If
End Sub

' Enter a transaction by OK-code, via the guard.
Public Sub StartTransaction(ByVal tcode As String)
    AssertTcodeAllowed tcode
    GuardedOkCode "/n" & tcode
    modSapConnect.WaitForSap
End Sub

'-----------------------------------------------------------------------
' OK-code guard
'-----------------------------------------------------------------------
Public Sub GuardedOkCode(ByVal code As String)
    Dim bare As String

    bare = UCase$(Trim$(code))
    ' Strip the navigation prefixes so '/nBU' is checked as 'BU'.
    Do While Left$(bare, 1) = "/"
        If Left$(bare, 2) = "/N" Or Left$(bare, 2) = "/O" Then
            bare = Mid$(bare, 3)
        Else
            bare = Mid$(bare, 2)
        End If
    Loop

    If InList(bare, BLOCKED_OKCODES) Then
        gWriteAttemptBlocked = gWriteAttemptBlocked + 1
        Err.Raise vbObjectError + 531, "modSafety.GuardedOkCode", _
                  "Blocked OK-code '" & code & "' -- it commits data. " & _
                  "This run is read-only."
    End If

    ' Deliberately no 'settling' Enter before this. On a selection screen a
    ' blind Enter executes the selection, and on a document screen it can
    ' confirm whatever is pending -- exactly the class of accident this
    ' module exists to prevent.
    modSapConnect.gSession.findById("wnd[0]/tbar[0]/okcd").Text = code
    modSapConnect.Element("wnd[0]").sendVKey 0
End Sub

'-----------------------------------------------------------------------
' Keystroke guard
'-----------------------------------------------------------------------
Public Sub GuardedSendVKey(ByVal windowId As String, ByVal vkey As Long)
    If InList(CStr(vkey), BLOCKED_VKEYS) Then
        gWriteAttemptBlocked = gWriteAttemptBlocked + 1
        Err.Raise vbObjectError + 532, "modSafety.GuardedSendVKey", _
                  "Blocked VKey " & vkey & " (Save) on " & windowId & _
                  ". This run is read-only."
    End If

    modSapConnect.Element(windowId).sendVKey vkey
    modSapConnect.WaitForSap
End Sub

'-----------------------------------------------------------------------
' Button guard -- inspects what the control actually says before pressing.
' A recording gives you an ID like 'wnd[0]/tbar[1]/btn[8]'; on a different
' release that same slot can be a different function.
'-----------------------------------------------------------------------
Public Sub GuardedPress(ByVal elementId As String)
    Dim control As Object
    Dim label As String

    Set control = modSapConnect.Element(elementId)

    label = vbNullString
    On Error Resume Next
    label = UCase$(control.Text & " " & control.Tooltip)
    On Error GoTo 0

    If LooksLikeWriteAction(label) Then
        gWriteAttemptBlocked = gWriteAttemptBlocked + 1
        Err.Raise vbObjectError + 533, "modSafety.GuardedPress", _
                  "Refused to press '" & elementId & "'. Its label reads """ & _
                  Trim$(label) & """, which looks like a write action." & vbCrLf & vbCrLf & _
                  "Check the recorded ID -- button positions move between releases."
    End If

    control.Press
    modSapConnect.WaitForSap
End Sub

Public Function LooksLikeWriteAction(ByVal label As String) As Boolean
    Dim fragments() As String
    Dim i As Long
    Dim upper As String

    upper = UCase$(label)
    fragments = Split(BLOCKED_TEXT_FRAGMENTS, "|")

    For i = LBound(fragments) To UBound(fragments)
        If InStr(upper, fragments(i)) > 0 Then
            ' 'Change layout' and 'Change display' are display-side only, and
            ' 'Reset layout' likewise -- they touch the ALV, not the document.
            If Not IsLayoutOnly(upper) Then
                LooksLikeWriteAction = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Function IsLayoutOnly(ByVal upper As String) As Boolean
    IsLayoutOnly = (InStr(upper, "LAYOUT") > 0) Or _
                   (InStr(upper, "DISPLAY") > 0) Or _
                   (InStr(upper, "VARIANT") > 0)
End Function

'-----------------------------------------------------------------------
' Popup guard
'-----------------------------------------------------------------------
Public Sub AssertPopupKnown()
    Dim title As String

    If Not modSapConnect.ModalWindowOpen() Then Exit Sub

    title = Trim$(modSapConnect.ModalWindowTitle())
    If Not TitleIsKnown(title) Then
        Err.Raise vbObjectError + 534, "modSafety.AssertPopupKnown", _
                  "Unrecognised popup: """ & title & """" & vbCrLf & vbCrLf & _
                  "The run stopped rather than clicking through it. Deal with it by " & _
                  "hand, note what it was, and either add it to KNOWN_POPUP_TITLES in " & _
                  "modSafety or handle it explicitly."
    End If
End Sub

Private Function TitleIsKnown(ByVal title As String) As Boolean
    Dim known() As String
    Dim i As Long

    known = Split(KNOWN_POPUP_TITLES, "|")
    For i = LBound(known) To UBound(known)
        If InStr(1, title, known(i), vbTextCompare) > 0 Then
            TitleIsKnown = True
            Exit Function
        End If
    Next i
End Function

'-----------------------------------------------------------------------
' Dry-run gate. Every routine that would write a file or leave a trace
' calls this first, so DRY RUN is enforced in one place.
'-----------------------------------------------------------------------
Public Function BlockedByDryRun(ByVal what As String) As Boolean
    If modConfig.IsDryRun() Then
        modLog.LogAction 0, "DRY RUN", what, "SKIPPED", vbNullString
        BlockedByDryRun = True
    End If
End Function

'-----------------------------------------------------------------------
Private Function InList(ByVal needle As String, ByVal pipeList As String) As Boolean
    InList = (InStr(1, "|" & pipeList & "|", "|" & needle & "|", vbTextCompare) > 0)
End Function
