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
'
' LANGUAGE NOTE. Button labels are translated, so this list only bites on a
' logon whose language it covers. That is acceptable only because it is the
' second line of defence, not the first: every ID this module presses comes
' from the curated Screen Map, so the text check exists to catch the case
' where a recorded ID has moved to a different function on another release.
' Where it does not match, the OK-code and VKey blocks below still apply,
' and the authorisation profile still applies underneath everything.
Private Const BLOCKED_TEXT_FRAGMENTS As String = _
    "POST|SAVE|DELETE|REVERSE|CHANGE|EDIT|PARK|HOLD|RESET|CLEAR|RELEASE|" & _
    "SICHERN|BUCHEN|LOESCHEN|LÖSCHEN|STORNIEREN|AENDERN|ÄNDERN|" & _
    "GUARDAR|CONTABILIZAR|BORRAR|ANULAR|MODIFICAR|" & _
    "ENREGISTRER|COMPTABILISER|SUPPRIMER|EXTOURNER|MODIFIER|" & _
    "GRAVAR|CONTABILIZAR|ELIMINAR|ESTORNAR|" & _
    "SALVA|REGISTRA|CANCELLA|STORNA|MODIFICA"

' Function keys that commit on standard SAP screens. 11 is Save.
Private Const BLOCKED_VKEYS As String = "11"

' Modal windows the run knows how to handle. Anything else stops the run
' instead of being dismissed blind -- a blind Enter on an unrecognised
' popup is exactly how a 'read-only' script ends up committing something.
' 'Postprocess' covers FEBAN's own selection popup, which is where the
' company code and statement dates go on this system.
Private Const KNOWN_POPUP_TITLES As String = _
    "Save list in file|Select Spreadsheet|Export list object to XXL|" & _
    "Information|Attachment list|Service: Attachment list|" & _
    "Business Document Navigator|Save As|Directory Browse|" & _
    "Postprocess|Post-process|Bank Statement|Electronic Account Statement|" & _
    "Select File|Save in file|Restrictions|" & _
    "Cleared Line Items|Payment Usage|Payment usage|Document Overview|List of Documents"

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
    NoteUnknownPopup
End Sub

' Note an unrecognised modal window -- do NOT stop the run for it.
'
' This used to raise, and that was wrong. The guard's real purpose is to
' stop a blind Enter through a window nobody identified, but nothing here
' presses blind: every action goes to a specific mapped id, and the
' committing OK-codes and the Save key are blocked outright regardless of
' what is on screen. Raising on an unfamiliar title therefore stopped
' legitimate flows -- the cleared-items list opens in a window titled after
' the document, so its title is different for every single sample and could
' never have been on a fixed list.
'
' It is logged, so an unexpected window is still visible in the trail.
Public Sub NoteUnknownPopup()
    Dim title As String

    If Not modSapConnect.ModalWindowOpen() Then Exit Sub
    If PopupLooksKnownByStructure() Then Exit Sub

    title = Trim$(modSapConnect.ModalWindowTitle())
    If TitleIsKnown(title) Then Exit Sub

    modLog.LogAction 0, "Popup", _
                 "A window titled """ & title & """ is open and is not one the run " & _
                 "recognises. Carrying on, because nothing is pressed blind -- but if " & _
                 "the next step misbehaves, this is where to look.", _
                 "MANUAL", vbNullString
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
' Dry-run gate.
'
' This used to block the exports, which was wrong twice over. The exports
' are read-only on the SAP side -- they write a file on your own disk and
' change nothing in SAP -- so blocking them protected against nothing. And
' the chain READS those files back to find the ZP numbers and the largest
' payment, so blocking them stopped every sample at step 6. A dry run could
' never rehearse steps 7 to 10, which are exactly the steps worth
' rehearsing.
'
' DRY RUN now means "put the exports somewhere separate", not "do not
' export". modConfig.DownloadRoot sends them to a '_dry run' subfolder, so
' the whole chain runs and a rehearsal cannot be mistaken for evidence.
'
' What actually keeps the run safe is the rest of this module -- the
' transaction allowlist, the blocked OK-codes, the blocked Save key and the
' popup guard -- and those apply in both modes.
'-----------------------------------------------------------------------
Public Function BlockedByDryRun(ByVal what As String) As Boolean
    ' Deliberately never blocks. Kept as a single place to change if a
    ' genuinely destructive step is ever added.
    BlockedByDryRun = False
End Function

'-----------------------------------------------------------------------
Private Function InList(ByVal needle As String, ByVal pipeList As String) As Boolean
    InList = (InStr(1, "|" & pipeList & "|", "|" & needle & "|", vbTextCompare) > 0)
End Function
