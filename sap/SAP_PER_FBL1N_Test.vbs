Option Explicit

'===============================================================================
' SAP_PER_FBL1N_Test.vbs - connectivity test, VBScript twin of modSapFbl1nTest.
'
' Attaches to the SAP session for system PER that is ALREADY OPEN on this PC
' and navigates it to FBL1N (Vendor Line Item Display).
'
' Run:   double-click, or   cscript.exe SAP_PER_FBL1N_Test.vbs /nologo
' Exit:  0 = pass, 1 = fail.  A copy of the log is written next to this script.
'
' This calls exactly the same SAP COM API as the VBA module - a .vbs is NOT
' "a SAP script", it is plain VBScript run by wscript/cscript. The only thing
' that differs between the two files is which process makes the call
' (wscript.exe vs EXCEL.EXE), which matters only if something blocks one host.
'
' DESIGN NOTES - see modSapFbl1nTest.bas for the full write-up:
'   1. Two monikers: classic SAP GUI publishes "SAPGUI", SAP Business Client /
'      NWBC publishes "SAPGUISERVER". The wrong one fails -2147221020
'      (MK_E_SYNTAX), which looks like a security block but is not. Try both.
'   2. Never index a SAP GUI collection with a Long -> error 618. Use CInt().
'   3. Pick the session BY SystemName, never Children(0).Children(0), or with
'      two systems open the script silently drives the wrong one.
'   4. Under On Error Resume Next a throwing property read abandons the WHOLE
'      statement, and a throwing For limit runs the body once with an
'      uninitialised counter. Read risky properties one per statement and
'      capture collection counts BEFORE the For.
'
' Keep this file saved WITHOUT a BOM - cscript rejects a UTF-8 BOM with
' "Invalid character" at (1,1).
'===============================================================================

Const TARGET_SYSTEM = "PER"
Const TARGET_TCODE  = "FBL1N"
Const MAXIMIZE_WINDOW = False

Dim gLog, gMoniker, gErrMsg
gLog = ""
gMoniker = ""
gErrMsg = ""

Dim gExitCode
gExitCode = Main()
WScript.Quit gExitCode


'-------------------------------------------------------------------------------
Function Main()
    Dim sess, txBefore, txAfter, winTitle, sbarText, sbarType, popups, okcd

    Say "SAP " & TARGET_SYSTEM & " -> " & TARGET_TCODE & " connectivity test"
    Say String(58, "-")

    '--- [1] attach -------------------------------------------------------------
    Say "[1] attaching to the running SAP GUI"
    Set sess = PickSession(TARGET_SYSTEM)
    If sess Is Nothing Then
        Say ""
        Say "FAIL - " & gErrMsg
        Main = Finish(1)
        Exit Function
    End If

    txBefore = ""
    On Error Resume Next
    txBefore = CStr(sess.Info.Transaction)
    On Error GoTo 0
    Say "    attached - transaction before = " & Dsp(txBefore)

    '--- [2] navigate -----------------------------------------------------------
    Say "[2] sending /n" & TARGET_TCODE & " to the command field"

    On Error Resume Next
    If MAXIMIZE_WINDOW Then sess.findById("wnd[0]").Maximize

    Set okcd = Nothing
    Set okcd = sess.findById("wnd[0]/tbar[0]/okcd")
    If okcd Is Nothing Then
        On Error GoTo 0
        Say ""
        Say "FAIL - the command field wnd[0]/tbar[0]/okcd was not found."
        Say "       The session is probably sitting on a modal dialog."
        Main = Finish(1)
        Exit Function
    End If

    okcd.Text = "/n" & TARGET_TCODE
    If Err.Number <> 0 Then
        Say "    could not write the command field (error " & Err.Number & ": " & Err.Description & ")"
        Err.Clear
    End If

    sess.findById("wnd[0]").sendVKey 0
    If Err.Number <> 0 Then
        Say "    sendVKey failed (error " & Err.Number & ": " & Err.Description & ")"
        Err.Clear
    End If
    On Error GoTo 0

    '--- [3] verify -------------------------------------------------------------
    txAfter = "" : winTitle = "" : sbarText = "" : sbarType = "" : popups = 0
    On Error Resume Next
    txAfter = CStr(sess.Info.Transaction)
    winTitle = CStr(sess.findById("wnd[0]").Text)
    popups = CInt(sess.Children.Count) - 1
    sbarText = CStr(sess.findById("wnd[0]/sbar").Text)
    sbarType = CStr(sess.findById("wnd[0]/sbar").MessageType)
    On Error GoTo 0

    Say "[3] result"
    Say "    transaction = " & Dsp(txAfter)
    Say "    window      = " & Dsp(winTitle)
    If Len(sbarText) > 0 Then Say "    status bar  = [" & sbarType & "] " & sbarText
    If popups > 0 Then Say "    NOTE: " & popups & " popup window(s) open on top of wnd[0]"

    Say ""
    If UCase(Trim(txAfter)) = TARGET_TCODE Then
        Say "PASS - attached to " & TARGET_SYSTEM & " and " & TARGET_TCODE & " is on screen."
        Main = Finish(0)
    Else
        Say "FAIL - navigation did not land on " & TARGET_TCODE & "."
        Say "       Check the status bar text above (missing authorisation, or the"
        Say "       previous screen refused to leave with a data-loss prompt)."
        Main = Finish(1)
    End If
End Function


'-------------------------------------------------------------------------------
' Returns the GuiApplication, trying both monikers. Nothing + gErrMsg on failure.
'-------------------------------------------------------------------------------
Function GetSapApp()
    Dim order, i, guiAuto, app, eNum, seen

    Set GetSapApp = Nothing
    seen = ""

    If gMoniker = "SAPGUI" Then
        order = Array("SAPGUI", "SAPGUISERVER")
    Else
        order = Array("SAPGUISERVER", "SAPGUI")
    End If

    For i = LBound(order) To UBound(order)
        Set guiAuto = Nothing
        eNum = 0
        On Error Resume Next
        Set guiAuto = GetObject(CStr(order(i)))
        eNum = Err.Number
        On Error GoTo 0

        If guiAuto Is Nothing Then
            Say "    moniker " & order(i) & " -> not registered (error " & eNum & ")"
            seen = seen & CStr(order(i)) & "=" & eNum & " "
        Else
            Set app = Nothing
            eNum = 0
            On Error Resume Next
            Set app = guiAuto.GetScriptingEngine
            eNum = Err.Number
            On Error GoTo 0

            If app Is Nothing Then
                Say "    moniker " & order(i) & " -> found, but GetScriptingEngine failed (error " & eNum & ")"
                seen = seen & CStr(order(i)) & "=engine" & eNum & " "
            Else
                gMoniker = CStr(order(i))
                Say "    moniker " & gMoniker & " -> OK"
                Set GetSapApp = app
                Exit Function
            End If
        End If
    Next

    gErrMsg = "Could not reach the SAP scripting engine. Tried both monikers (" & Trim(seen) & ")." & vbCrLf & _
              "       Check: (1) SAP GUI is open and logged on, (2) scripting is enabled in" & vbCrLf & _
              "       SAP GUI Options and on the server, (3) this script and SAP GUI run at" & vbCrLf & _
              "       the same elevation."
End Function


'-------------------------------------------------------------------------------
' Enumerates every connection/session and returns the one whose SystemName
' matches wantSystem. Nothing + gErrMsg if there is no safe match.
'-------------------------------------------------------------------------------
Function PickSession(wantSystem)
    Dim app, conn, sess
    Dim nConn, nSess, i, j, total
    Dim descr, sysName, client, usr, tx, busy, wins, flags
    Dim idleMatch, blockedMatch, descrMatch

    Set PickSession = Nothing
    Set idleMatch = Nothing
    Set blockedMatch = Nothing
    Set descrMatch = Nothing
    total = 0

    Set app = GetSapApp()
    If app Is Nothing Then Exit Function

    ' Count captured BEFORE the loop - see design note 4.
    nConn = 0
    On Error Resume Next
    nConn = CInt(app.Children.Count)
    On Error GoTo 0
    Say "    connections open: " & nConn

    If nConn = 0 Then
        gErrMsg = "The scripting engine answered but no SAP connection is open. " & _
                  "Log on to " & wantSystem & " in SAP GUI first, then run this again."
        Exit Function
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
                    sysName = "" : client = "" : usr = "" : tx = ""
                    busy = False : wins = 1
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

                    If UCase(Trim(sysName)) = UCase(Trim(wantSystem)) Then
                        If busy Or wins > 1 Then
                            If blockedMatch Is Nothing Then Set blockedMatch = sess
                        Else
                            If idleMatch Is Nothing Then Set idleMatch = sess
                        End If
                    ElseIf Len(Trim(sysName)) = 0 Then
                        ' SystemName unreadable - fall back to the connection
                        ' description, but only then. A readable SystemName is
                        ' authoritative and must never be overridden by a name.
                        If InStr(1, descr, wantSystem, vbTextCompare) > 0 Then
                            If descrMatch Is Nothing Then Set descrMatch = sess
                        End If
                    End If
                End If
            Next
        End If
    Next

    If Not idleMatch Is Nothing Then
        Say "    selected: the " & wantSystem & " session (idle, no popup)"
        Set PickSession = idleMatch
        Exit Function
    End If

    If Not blockedMatch Is Nothing Then
        gErrMsg = "Found the " & wantSystem & " session, but it is busy or has a popup open, " & _
                  "so it was left alone. Close the dialog in SAP / wait for it to finish, " & _
                  "then run this again."
        Exit Function
    End If

    If Not descrMatch Is Nothing Then
        Say "    selected: matched the connection description (SystemName unreadable)"
        Set PickSession = descrMatch
        Exit Function
    End If

    gErrMsg = "No open session reports SystemName = " & wantSystem & ". " & total & _
              " session(s) were found and are listed above. Log on to " & wantSystem & _
              ", or change TARGET_SYSTEM to the SID shown in the list. " & _
              "Nothing was sent - refusing to drive a system that was not asked for."
End Function


'-------------------------------------------------------------------------------
' Small helpers.
'-------------------------------------------------------------------------------
Sub Say(s)
    gLog = gLog & s & vbCrLf
End Sub

Function Dsp(s)
    If Len(Trim(s & "")) = 0 Then Dsp = "<empty>" Else Dsp = s
End Function

Function IsCScript()
    IsCScript = (LCase(Right(WScript.FullName, 11)) = "cscript.exe")
End Function

Function Finish(code)
    Dim p
    p = SaveLog()
    If Len(p) > 0 Then gLog = gLog & vbCrLf & "Log also saved to:" & vbCrLf & p

    If IsCScript() Then
        WScript.Echo gLog
    Else
        MsgBox gLog, vbInformation, "SAP " & TARGET_SYSTEM & " / " & TARGET_TCODE & " test"
    End If
    Finish = code
End Function

Function SaveLog()
    Dim fso, folder, p, ts

    SaveLog = ""
    On Error Resume Next

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Err.Number <> 0 Then Exit Function

    folder = fso.GetParentFolderName(WScript.ScriptFullName)
    If Len(folder) = 0 Then folder = "."
    p = fso.BuildPath(folder, "SAP_" & TARGET_SYSTEM & "_" & TARGET_TCODE & "_Test.txt")

    Set ts = fso.CreateTextFile(p, True)
    If Err.Number <> 0 Then
        ' Script folder is read-only (network share, etc.) - fall back to TEMP.
        Err.Clear
        p = fso.BuildPath(fso.GetSpecialFolder(2), "SAP_" & TARGET_SYSTEM & "_" & TARGET_TCODE & "_Test.txt")
        Set ts = fso.CreateTextFile(p, True)
        If Err.Number <> 0 Then Exit Function
    End If

    ts.Write gLog
    ts.Close
    If Err.Number = 0 Then SaveLog = p
End Function
