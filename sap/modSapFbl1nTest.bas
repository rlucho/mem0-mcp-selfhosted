Attribute VB_Name = "modSapFbl1nTest"
Option Explicit

'===============================================================================
' modSapFbl1nTest - connectivity test.
'
' Attaches to the SAP session for system PER that is ALREADY OPEN on this PC
' and navigates it to FBL1N (Vendor Line Item Display).
'
' Entry point (Alt+F8):   Test_SAP_PER_FBL1N
'
' WHAT IT DOES / DOES NOT DO
'   - It does NOT log on to SAP. SAP GUI must already be running and logged on
'     to PER; GetObject() only attaches to that running window via local COM.
'   - It does NOT execute the report. It stops on the FBL1N selection screen.
'   - Apart from the /nFBL1N navigation itself, nothing is changed or posted.
'
' DESIGN NOTES - all four were paid for in debugging time, keep them:
'
'   1. TWO MONIKERS EXIST. Classic SAP GUI (saplogon.exe) publishes "SAPGUI".
'      SAP Business Client / NWBC publishes "SAPGUISERVER". A machine answers
'      to one and fails the other with -2147221020 (MK_E_SYNTAX), which VBA
'      reports as "Automation error / Invalid syntax" and which looks exactly
'      like an Office macro security block but is only an unresolvable name.
'      Both are tried; the winner is cached for subsequent calls.
'
'   2. NEVER INDEX A SAP GUI COLLECTION WITH A Long. app.Children(0) works only
'      because VBA types the literal as Integer (VT_I2); a Long loop counter
'      raises error 618 "Bad index type for collection access". Hence CInt().
'
'   3. THE SESSION IS PICKED BY SystemName, not by Children(0).Children(0).
'      With two systems open, index 0 is whichever was opened first, so the
'      classic pattern can silently drive the wrong system.
'
'   4. UNDER On Error Resume Next, a property read that throws while an
'      argument is being evaluated abandons the ENTIRE statement - the log line
'      simply vanishes. And "For j = 0 To coll.Count - 1" with a throwing limit
'      expression enters the body ONCE with j uninitialised. So: risky
'      properties are read into their own variable first, and collection counts
'      are captured into a variable BEFORE the For.
'
' REQUIREMENTS
'   - SAP GUI open and logged on to PER.
'   - Scripting enabled client side (Options > Accessibility & Scripting >
'     Scripting > Enable scripting) and server side (sapgui/user_scripting).
'   - Excel and SAP GUI running at the same elevation (both normal, or both
'     "as administrator") - COM will not cross that boundary.
'===============================================================================

Private Const MODULE_NAME     As String = "modSapFbl1nTest"
Private Const TARGET_SYSTEM   As String = "PER"
Private Const TARGET_TCODE    As String = "FBL1N"

' Leave False to keep the user's window layout untouched.
Private Const MAXIMIZE_WINDOW As Boolean = False

Private mLog     As String   ' collected report
Private mMoniker As String   ' cached winning moniker name


'-------------------------------------------------------------------------------
' Entry point.
'-------------------------------------------------------------------------------
Public Sub Test_SAP_PER_FBL1N()
    Dim sess     As Object
    Dim txBefore As String
    Dim txAfter  As String
    Dim winTitle As String
    Dim sbarText As String
    Dim sbarType As String
    Dim popups   As Integer
    Dim errNum   As Long
    Dim eDesc    As String

    mLog = ""
    Say "SAP " & TARGET_SYSTEM & " -> " & TARGET_TCODE & " connectivity test"
    Say String$(58, "-")

    On Error GoTo Fail

    '--- [1] attach to the running SAP GUI --------------------------------------
    Say "[1] attaching to the running SAP GUI"
    Set sess = PickSession(TARGET_SYSTEM)

    txBefore = ""
    On Error Resume Next
    txBefore = CStr(sess.Info.Transaction)
    On Error GoTo Fail
    Say "    attached - transaction before = " & Dsp(txBefore)

    '--- [2] navigate -----------------------------------------------------------
    Say "[2] sending /n" & TARGET_TCODE & " to the command field"
    If MAXIMIZE_WINDOW Then sess.findById("wnd[0]").Maximize
    sess.findById("wnd[0]/tbar[0]/okcd").Text = "/n" & TARGET_TCODE
    sess.findById("wnd[0]").sendVKey 0

    '--- [3] verify where we landed ---------------------------------------------
    txAfter = "": winTitle = "": sbarText = "": sbarType = "": popups = 0
    On Error Resume Next
    txAfter = CStr(sess.Info.Transaction)
    winTitle = CStr(sess.findById("wnd[0]").Text)
    popups = CInt(sess.Children.Count) - 1
    sbarText = CStr(sess.findById("wnd[0]/sbar").Text)
    sbarType = CStr(sess.findById("wnd[0]/sbar").MessageType)
    On Error GoTo Fail

    Say "[3] result"
    Say "    transaction = " & Dsp(txAfter)
    Say "    window      = " & Dsp(winTitle)
    If Len(sbarText) > 0 Then Say "    status bar  = [" & sbarType & "] " & sbarText
    If popups > 0 Then Say "    NOTE: " & popups & " popup window(s) open on top of wnd[0]"

    Say ""
    If UCase$(Trim$(txAfter)) = TARGET_TCODE Then
        Say "PASS - attached to " & TARGET_SYSTEM & " and " & TARGET_TCODE & " is on screen."
    Else
        Say "FAIL - navigation did not land on " & TARGET_TCODE & "."
        Say "       Check the status bar text above (missing authorisation, or the"
        Say "       previous screen refused to leave with a data-loss prompt)."
    End If

    Finish
    Exit Sub

Fail:
    errNum = Err.Number
    eDesc = Err.Description
    Say ""
    Say "FAIL - " & eDesc
    Say "       (error " & errNum & ")"
    Finish
End Sub


'-------------------------------------------------------------------------------
' Returns the GuiApplication (scripting engine), trying both monikers.
' Raises a descriptive error if neither answers.
'-------------------------------------------------------------------------------
Private Function GetSapApp() As Object
    Dim order   As Variant
    Dim i       As Integer
    Dim guiAuto As Object
    Dim app     As Object
    Dim errNum  As Long
    Dim seen    As String

    ' Try the moniker that worked last time first - a workbook with many
    ' transaction modules connects repeatedly.
    If mMoniker = "SAPGUI" Then
        order = Array("SAPGUI", "SAPGUISERVER")
    Else
        order = Array("SAPGUISERVER", "SAPGUI")
    End If

    For i = LBound(order) To UBound(order)
        Set guiAuto = Nothing
        errNum = 0
        On Error Resume Next
        Set guiAuto = GetObject(CStr(order(i)))
        errNum = Err.Number
        On Error GoTo 0

        If guiAuto Is Nothing Then
            Say "    moniker " & order(i) & " -> not registered (error " & errNum & ")"
            seen = seen & CStr(order(i)) & "=" & errNum & " "
        Else
            Set app = Nothing
            errNum = 0
            On Error Resume Next
            Set app = guiAuto.GetScriptingEngine
            errNum = Err.Number
            On Error GoTo 0

            If app Is Nothing Then
                Say "    moniker " & order(i) & " -> found, but GetScriptingEngine failed (error " & errNum & ")"
                seen = seen & CStr(order(i)) & "=engine" & errNum & " "
            Else
                mMoniker = CStr(order(i))
                Say "    moniker " & mMoniker & " -> OK"
                Set GetSapApp = app
                Exit Function
            End If
        End If
    Next i

    Err.Raise vbObjectError + 513, MODULE_NAME, _
        "Could not reach the SAP scripting engine. Tried both monikers (" & Trim$(seen) & ")." & vbCrLf & _
        "Check: (1) SAP GUI is open and logged on, (2) scripting is enabled in " & _
        "SAP GUI Options and on the server, (3) Excel and SAP GUI run at the same elevation."
End Function


'-------------------------------------------------------------------------------
' Enumerates every connection/session and returns the one whose SystemName
' matches wantSystem. Logs the full inventory so a failure is self-diagnosing.
'-------------------------------------------------------------------------------
Private Function PickSession(ByVal wantSystem As String) As Object
    Dim app          As Object
    Dim conn         As Object
    Dim sess         As Object
    Dim nConn        As Integer
    Dim nSess        As Integer
    Dim i            As Integer
    Dim j            As Integer
    Dim total        As Integer
    Dim descr        As String
    Dim sysName      As String
    Dim client       As String
    Dim usr          As String
    Dim tx           As String
    Dim busy         As Boolean
    Dim wins         As Integer
    Dim flags        As String
    Dim idleMatch    As Object
    Dim blockedMatch As Object
    Dim descrMatch   As Object

    Set app = GetSapApp()

    ' Count captured BEFORE the loop - see design note 4.
    nConn = 0
    On Error Resume Next
    nConn = CInt(app.Children.Count)
    On Error GoTo 0
    Say "    connections open: " & nConn

    If nConn = 0 Then
        Err.Raise vbObjectError + 514, MODULE_NAME, _
            "The scripting engine answered but no SAP connection is open. " & _
            "Log on to " & wantSystem & " in SAP GUI first, then run this again."
    End If

    For i = 0 To nConn - 1
        Set conn = Nothing
        On Error Resume Next
        Set conn = app.Children(CInt(i))       ' CInt - see design note 2
        On Error GoTo 0

        If conn Is Nothing Then
            Say "    conn(" & i & ") unreadable"
        Else
            descr = ""
            On Error Resume Next
            descr = CStr(conn.Description)     ' own statement - see design note 4
            On Error GoTo 0

            nSess = 0
            On Error Resume Next
            nSess = CInt(conn.Children.Count)
            On Error GoTo 0

            Say "    conn(" & i & ") """ & descr & """  sessions=" & nSess

            For j = 0 To nSess - 1
                Set sess = Nothing
                On Error Resume Next
                Set sess = conn.Children(CInt(j))
                On Error GoTo 0

                If sess Is Nothing Then
                    Say "      sess(" & j & ") unreadable"
                Else
                    sysName = "": client = "": usr = "": tx = "": busy = False: wins = 1
                    On Error Resume Next
                    sysName = CStr(sess.Info.SystemName)
                    client = CStr(sess.Info.Client)
                    usr = CStr(sess.Info.User)
                    tx = CStr(sess.Info.Transaction)
                    busy = CBool(sess.Busy)
                    wins = CInt(sess.Children.Count)
                    On Error GoTo 0

                    total = total + 1
                    flags = ""
                    If busy Then flags = flags & "  [BUSY]"
                    If wins > 1 Then flags = flags & "  [POPUP OPEN]"

                    Say "      sess(" & j & ") System=" & Dsp(sysName) & _
                        "  Client=" & Dsp(client) & _
                        "  User=" & Dsp(usr) & _
                        "  Tx=" & Dsp(tx) & flags

                    If UCase$(Trim$(sysName)) = UCase$(Trim$(wantSystem)) Then
                        If busy Or wins > 1 Then
                            If blockedMatch Is Nothing Then Set blockedMatch = sess
                        Else
                            If idleMatch Is Nothing Then Set idleMatch = sess
                        End If
                    ElseIf Len(Trim$(sysName)) = 0 Then
                        ' SystemName unreadable - fall back to the connection
                        ' description, but only then. A readable SystemName is
                        ' authoritative and must never be overridden by a name.
                        If InStr(1, descr, wantSystem, vbTextCompare) > 0 Then
                            If descrMatch Is Nothing Then Set descrMatch = sess
                        End If
                    End If
                End If
            Next j
        End If
    Next i

    If Not idleMatch Is Nothing Then
        Say "    selected: the " & wantSystem & " session (idle, no popup)"
        Set PickSession = idleMatch
        Exit Function
    End If

    If Not blockedMatch Is Nothing Then
        Err.Raise vbObjectError + 515, MODULE_NAME, _
            "Found the " & wantSystem & " session, but it is busy or has a popup open, " & _
            "so it was left alone. Close the dialog in SAP / wait for it to finish, " & _
            "then run this again."
    End If

    If Not descrMatch Is Nothing Then
        Say "    selected: matched the connection description (SystemName unreadable)"
        Set PickSession = descrMatch
        Exit Function
    End If

    Err.Raise vbObjectError + 516, MODULE_NAME, _
        "No open session reports SystemName = " & wantSystem & ". " & total & _
        " session(s) were found and are listed above. Log on to " & wantSystem & _
        ", or change TARGET_SYSTEM to the SID shown in the list. " & _
        "Nothing was sent - refusing to drive a system that was not asked for."
End Function


'-------------------------------------------------------------------------------
' Small helpers.
'-------------------------------------------------------------------------------
Private Sub Say(ByVal s As String)
    mLog = mLog & s & vbCrLf
End Sub

Private Function Dsp(ByVal s As String) As String
    If Len(Trim$(s)) = 0 Then Dsp = "<empty>" Else Dsp = s
End Function

Private Sub Finish()
    Dim p As String
    Debug.Print mLog
    p = SaveLog()
    If Len(p) > 0 Then
        MsgBox mLog & vbCrLf & "Log also saved to:" & vbCrLf & p, _
               vbInformation, "SAP " & TARGET_SYSTEM & " / " & TARGET_TCODE & " test"
    Else
        MsgBox mLog, vbInformation, "SAP " & TARGET_SYSTEM & " / " & TARGET_TCODE & " test"
    End If
End Sub

Private Function SaveLog() As String
    Dim p As String
    Dim f As Integer

    On Error GoTo Nope
    p = Environ$("TEMP")
    If Len(p) = 0 Then Exit Function
    p = p & "\SAP_" & TARGET_SYSTEM & "_" & TARGET_TCODE & "_Test.txt"

    f = FreeFile
    Open p For Output As #f
    Print #f, mLog
    Close #f

    SaveLog = p
    Exit Function

Nope:
    On Error Resume Next
    If f <> 0 Then Close #f
    SaveLog = ""
End Function
