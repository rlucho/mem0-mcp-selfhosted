'============================================================================
' mCloseEnv_V4  -  Closing Manager (V4-CIO)
'----------------------------------------------------------------------------
' NEW module. Import it once; it adds the environment / path helpers used by
' the V4 versions of CreatePaths, RunClosing and CombinePDF.
'
' It centralises everything that used to be scattered and inconsistent in V3:
'   * ONE working drive for all \pdf\ folders (fixes the D:\ vs C:\ split-brain)
'   * a guard that blocks the workbook running from OneDrive / SharePoint (URL path)
'   * folder creation that also builds any missing parent folders
'   * a one-click PreflightCheck that reports every dependency before a run
'
' All tunable locations live in the CONFIGURATION block below - change them
' here in one place, not deep inside the SAP routines.
'============================================================================
Option Explicit

'---------------------------- CONFIGURATION ---------------------------------
' Working drive for the \pdf\ scratch folders. V3 auto-picked "D:\ if present
' else C:\", but only ever created the folders on C:\ - so on any PC that had
' a D:\ drive the print/merge step looked in the wrong place and failed.
' V4 uses ONE value everywhere. Default C:\ (matches the final-report drive).
' If your standard build genuinely uses D:\, change this single line.
Public Const CM_BASE_DRIVE As String = "C:\"

' Final delivered report tree (year\month is appended automatically).
Public Const CM_REPORT_ROOT As String = "C:\_Files to Transfer\MONTH END CLOSE\"

' Master copy of the PDF merge tool, pulled from the network share if missing.
Public Const CM_MERGER_SRC As String = _
    "\\pl-krabpo-fsc01\ipa$\R2R\R2R - IP EU GL West\USEFUL\pdf\merger\GiosPSMC.exe"

' Capgemini SharePoint SOAP endpoint (Lists.asmx) used by UpdateData and the
' sp* helper subs. Change here if the site is moved/decommissioned.
Public Const CM_SP_BASE As String = "https://troom-x.capgemini.com/sites/InternationalPaper/r2a/_vti_bin/Lists.asmx"

' UNC archive root for the closing reports (year\month is appended). NB: this
' is assigned to FShared, which the original v3 flow never reads downstream.
Public Const CM_ARCHIVE_ROOT As String = "\\pl-krabpo-fsc01\ipa$\R2R\R2R - IP EU\MONTH-END\CLOSING REPORTS\"

' --- SAP objects the closing macro drives, tested by the optional preflight
'     authorisation sweep. Keep these in step with the code if a transaction,
'     table or report group is ever added.
Public Const CM_SAP_TCODES  As String = "SE16,GR55,SM35,ZGE132,ZGLRME,ZGR215,ZGLGWUL,ZGE1174"
Public Const CM_SAP_TABLES  As String = "T001,T001B,T001Z,SKB1,ZCCOD,ZGXMIT"
Public Const CM_SAP_RGROUPS As String = "AA02,EIS4,GIS4,GTB1"

' The "/..." display layouts the close types into selection-screen fields.
' Format:  tcode | selection-screen field | layout name   (entries separated by ;)
' These are ALV display layouts (P_VARID / P_VARIE / P_ALV), not selection
' variants - the macro types them in and expects them to already exist.
Public Const CM_SAP_LAYOUTS As String = _
    "ZGLRME|P_VARID|/default;ZGLRME|P_VARIE|/closing;" & _
    "ZGLRME|P_VARIE|/default;ZGR215|P_ALV|/arek2"
'----------------------------------------------------------------------------


'--- TRUE when a path is an http(s) address, i.e. an Excel-in-the-cloud file -
Public Function IsUrlPath(ByVal p As String) As Boolean
    Dim s As String
    s = LCase$(Trim$(p))
    IsUrlPath = (Left$(s, 7) = "http://") Or (Left$(s, 8) = "https://")
End Function


'--- Stop early, with a clear message, if run from OneDrive / SharePoint -----
' Returns True when the workbook lives on a real local/UNC path.
Public Function AssertLocalWorkbook() As Boolean
    If IsUrlPath(ThisWorkbook.Path) Then
        MsgBox "This workbook is open from a OneDrive / SharePoint (web) location:" & vbCrLf & vbCrLf & _
               ThisWorkbook.Path & vbCrLf & vbCrLf & _
               "The Closing Manager saves SAP export files next to the workbook and " & _
               "cannot use a web address." & vbCrLf & vbCrLf & _
               "Please SAVE A COPY to a local folder (for example C:\Closing\) and run it from there.", _
               vbCritical, "Closing Manager - unsupported location"
        AssertLocalWorkbook = False
    Else
        AssertLocalWorkbook = True
    End If
End Function


'--- Create a folder and any missing parents (V3 could only create 1 level) --
Private Sub EnsureFolderChain(ByVal fso As Object, ByVal path As String)
    Dim p As String, parent As String
    p = path
    If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    If Len(p) = 0 Then Exit Sub
    If fso.FolderExists(p) Then Exit Sub
    parent = fso.GetParentFolderName(p)
    If Len(parent) > 0 Then
        If Not fso.FolderExists(parent) Then EnsureFolderChain fso, parent
    End If
    On Error Resume Next
    fso.CreateFolder p
    On Error GoTo 0
End Sub


'--- Copy the merge tool locally if it is not already there -----------------
Private Sub EnsureMerger(ByVal fso As Object)
    Dim dst As String
    dst = CM_BASE_DRIVE & "pdf\merger\GiosPSMC.exe"
    If fso.FileExists(dst) Then Exit Sub
    On Error Resume Next
    If fso.FileExists(CM_MERGER_SRC) Then fso.CopyFile CM_MERGER_SRC, dst
    On Error GoTo 0        ' silence: PreflightCheck / CombinePDF report a real absence
End Sub


'--- Build every folder the close needs, on ONE consistent drive ------------
Public Sub EnsureFolders()
    Dim fso As Object, base As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    base = CM_BASE_DRIVE

    EnsureFolderChain fso, base & "pdf\"
    EnsureFolderChain fso, base & "pdf\mergedFiles\"
    EnsureFolderChain fso, base & "pdf\merger\"
    EnsureFolderChain fso, base & "pdf\temp\"
    EnsureFolderChain fso, base & "pdf\printed\"
    EnsureFolderChain fso, base & "pdf\final\"

    ' Final report tree - only when the accounting date is known (LastDay set by
    ' CreateVariants). Guards the UpdateData path where LastDay may still be 0.
    If LastDay > 0 Then
        EnsureFolderChain fso, CM_REPORT_ROOT & Year(LastDay) & "\"
        EnsureFolderChain fso, CM_REPORT_ROOT & Year(LastDay) & "\" & Right$("0" & Month(LastDay), 2) & "\"
    End If

    EnsureMerger fso
End Sub


'============================================================================
' PreflightCheck  -  one-click "will the close run?" report.
' Assign it to a button on the START sheet. It turns every known point of
' failure into a plain-language yes/no BEFORE the operator kicks off a run.
'============================================================================
Public Sub PreflightCheck()
    Dim fso As Object, msg As String, okAll As Boolean
    Dim g As Object, eng As Object, sapOK As Boolean
    Dim sapSess As Object, sapBlocked As Boolean, runData As Boolean
    Dim cc As String, dst As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    okAll = True

    Call CreateVariants        ' make sure LastDay is set for the path checks

    ' 1) Workbook location -----------------------------------------------------
    If IsUrlPath(ThisWorkbook.Path) Then
        msg = msg & "[X] Workbook is running from the web (OneDrive/SharePoint)." & vbCrLf & _
                    "       " & ThisWorkbook.Path & vbCrLf & _
                    "       Save a local copy and run from there." & vbCrLf
        okAll = False
    Else
        msg = msg & "[OK] Workbook location is local." & vbCrLf
    End If

    ' 2) Working drive ---------------------------------------------------------
    If fso.DriveExists(Left$(CM_BASE_DRIVE, 2)) Then
        msg = msg & "[OK] Working drive " & CM_BASE_DRIVE & " is present." & vbCrLf
    Else
        msg = msg & "[X] Working drive " & CM_BASE_DRIVE & " was not found." & vbCrLf
        okAll = False
    End If

    ' 3) SAP GUI session -------------------------------------------------------
    sapOK = False
    On Error Resume Next
    Set g = GetObject("SAPGUI")
    If Not g Is Nothing Then
        Set eng = g.GetScriptingEngine
        If Not eng Is Nothing Then
            If eng.Children.Count > 0 Then
                'keep the session itself - the optional authorisation sweep needs it
                Set sapSess = eng.Children(0).Children(0)
                If Not sapSess Is Nothing Then sapOK = True
            End If
        End If
    End If
    On Error GoTo 0
    If sapOK Then
        msg = msg & "[OK] SAP GUI session detected." & vbCrLf
    Else
        msg = msg & "[X] SAP GUI not logged in, or GUI scripting is disabled." & vbCrLf
        okAll = False
    End If

    ' 4) PDFCreator ------------------------------------------------------------
    If fso.FileExists("C:\Program Files\PDFCreator\PDFCreator.exe") Or _
       fso.FileExists("C:\Program Files (x86)\PDFCreator\PDFCreator.exe") Then
        msg = msg & "[OK] PDFCreator is installed." & vbCrLf
    Else
        msg = msg & "[X] PDFCreator.exe not found (printer must be named 'PDFCreator')." & vbCrLf
        okAll = False
    End If

    ' 5) PDF merge tool --------------------------------------------------------
    dst = CM_BASE_DRIVE & "pdf\merger\GiosPSMC.exe"
    If fso.FileExists(dst) Then
        msg = msg & "[OK] PDF merger present." & vbCrLf
    ElseIf fso.FileExists(CM_MERGER_SRC) Then
        msg = msg & "[~] PDF merger will be copied from the network share on first run." & vbCrLf
    Else
        msg = msg & "[X] PDF merger missing and the network share is unreachable (VPN?)." & vbCrLf
        okAll = False
    End If

    ' 6) Cost centre -----------------------------------------------------------
    On Error Resume Next
    cc = Trim$(CStr(Sheets("config").Range("B2").Value))
    On Error GoTo 0
    If cc = "" Then
        msg = msg & "[X] config!B2 (Company Code) is empty." & vbCrLf
        okAll = False
    Else
        msg = msg & "[OK] Company Code = " & cc & "." & vbCrLf
    End If

    ' 7) SAP authorisations (optional - it navigates SAP, so it asks first) -----
    If sapOK Then
        If MsgBox("Environment checks done." & vbCrLf & vbCrLf & _
                  "Also test the SAP authorisations this close needs?" & vbCrLf & _
                  "  - " & CM_CountList(CM_SAP_TCODES) & " transactions" & vbCrLf & _
                  "  - " & CM_CountList(CM_SAP_TABLES) & " tables read through SE16" & vbCrLf & _
                  "  - " & CM_CountList(CM_SAP_RGROUPS) & " GR55 report groups" & vbCrLf & _
                  "  - " & CM_CountLayouts() & " report layouts (/default, /closing, /arek2)" & vbCrLf & vbCrLf & _
                  "This part runs no report and posts nothing, but it does move your" & vbCrLf & _
                  "SAP session between screens and can take up to a minute.", _
                  vbYesNo + vbQuestion, "Closing Manager - Preflight") = vbYes Then
            'second, separate consent: this one actually executes a report
            runData = (MsgBox("Also test access to company code " & cc & "'s data?" & vbCrLf & vbCrLf & _
                       "This runs ZGLRME for " & Monthx & "/" & Yearx & " with the transmit flag" & vbCrLf & _
                       "cleared - it only selects and displays, it posts nothing." & vbCrLf & _
                       "It is the only way to prove company-code level access, but" & vbCrLf & _
                       "it does read live data and may take a moment.", _
                       vbYesNo + vbQuestion, "Closing Manager - Preflight") = vbYes)
            msg = msg & CM_SapAuthReport(sapSess, sapBlocked, runData)
            If sapBlocked Then okAll = False
        Else
            msg = msg & "[~] SAP authorisations not tested (skipped)." & vbCrLf
        End If
    End If

    msg = "CLOSING MANAGER  -  PREFLIGHT CHECK" & vbCrLf & _
          "--------------------------------------" & vbCrLf & _
          msg & _
          "--------------------------------------" & vbCrLf & _
          IIf(okAll, "READY TO RUN.", "NOT READY - resolve the [X] items above.")
    MsgBox msg, IIf(okAll, vbInformation, vbExclamation), "Closing Manager - Preflight"
End Sub


'============================================================================
' SAP AUTHORISATION SWEEP
'----------------------------------------------------------------------------
' Everything the closing macro drives in SAP, tested before the close starts:
'   * the transactions it calls
'   * the tables it reads through SE16
'   * the GR55 report groups it runs
'
' Method: navigate to each object and read the status bar (wnd[0]/sbar). Any
' error/abort message is reported verbatim rather than matched on keywords, so
' the check behaves correctly whatever the SAP logon language is.
'
' It navigates only - no report is executed and nothing is posted - but it does
' move the SAP session, which is why PreflightCheck asks before running it.
' The session is returned to a blank screen (/n) afterwards.
'
' Display layouts (the "/default", "/closing", "/arek2" values the macro types
' into P_VARID / P_VARIE / P_ALV) ARE checked: the sweep opens the transaction,
' confirms the selection-screen field still exists, enters the layout and reads
' SAP's reply. Note these are ALV display layouts rather than selection variants.
'
' NOT covered: company-code / cost-centre level authorisation, which only bites
' when a report actually runs against real data.
'============================================================================
' (the object lists themselves live in the CONFIGURATION block at the top of
'  this module, because VBA requires module Consts before the first procedure)


Private Function CM_CountList(ByVal csv As String) As Long
    CM_CountList = UBound(Split(csv, ",")) - LBound(Split(csv, ",")) + 1
End Function


Private Function CM_CountLayouts() As Long
    CM_CountLayouts = UBound(Split(CM_SAP_LAYOUTS, ";")) - LBound(Split(CM_SAP_LAYOUTS, ";")) + 1
End Function


'--- close any modal popup so the next navigation starts from a clean screen -
Private Sub CM_ClearPopups(ByVal sess As Object)
    Dim n As Long
    On Error Resume Next
    For n = 1 To 4
        If sess.Children.Count > 1 Then
            sess.findById("wnd[1]").sendVKey 12      'F12 = cancel
        Else
            Exit For
        End If
    Next n
    Err.Clear
    On Error GoTo 0
End Sub


'--- read the status bar and turn it into "OK|" or "X|<reason>" --------------
Private Function CM_Verdict(ByVal sess As Object) As String
    Dim mType As String, txt As String
    On Error Resume Next
    mType = UCase$(sess.findById("wnd[0]/sbar").MessageType)
    txt = Trim$(sess.findById("wnd[0]/sbar").Text)
    Err.Clear
    On Error GoTo 0
    If mType = "E" Or mType = "A" Then
        If txt = "" Then txt = "blocked by SAP (no message text)"
        CM_Verdict = "X|" & txt
    Else
        CM_Verdict = "OK|"
    End If
End Function


Private Function CM_CheckTcode(ByVal sess As Object, ByVal tcode As String) As String
    On Error Resume Next
    CM_ClearPopups sess
    sess.findById("wnd[0]/tbar[0]/okcd").Text = "/n" & tcode
    sess.findById("wnd[0]").sendVKey 0
    If Err.Number <> 0 Then
        Err.Clear: On Error GoTo 0
        CM_CheckTcode = "?|could not drive the SAP command field"
        Exit Function
    End If
    On Error GoTo 0
    CM_CheckTcode = CM_Verdict(sess)
End Function


Private Function CM_CheckTable(ByVal sess As Object, ByVal tbl As String) As String
    On Error Resume Next
    CM_ClearPopups sess
    sess.findById("wnd[0]/tbar[0]/okcd").Text = "/nse16"
    sess.findById("wnd[0]").sendVKey 0
    sess.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").Text = tbl
    sess.findById("wnd[0]").sendVKey 0
    If Err.Number <> 0 Then
        Err.Clear: On Error GoTo 0
        CM_CheckTable = "?|SE16 screen not as expected"
        Exit Function
    End If
    On Error GoTo 0
    CM_CheckTable = CM_Verdict(sess)
End Function


Private Function CM_CheckReportGroup(ByVal sess As Object, ByVal rg As String) As String
    On Error Resume Next
    CM_ClearPopups sess
    sess.findById("wnd[0]/tbar[0]/okcd").Text = "/ngr55"
    sess.findById("wnd[0]").sendVKey 0
    sess.findById("wnd[0]/usr/ctxtRGRWJ-JOB").Text = rg
    sess.findById("wnd[0]").sendVKey 8       'to the selection screen only
    If Err.Number <> 0 Then
        Err.Clear: On Error GoTo 0
        CM_CheckReportGroup = "?|GR55 screen not as expected"
        Exit Function
    End If
    On Error GoTo 0
    CM_CheckReportGroup = CM_Verdict(sess)
End Function


'--- fill a selection screen's mandatory fields before testing a layout ------
' Without this the screen answers a bare Enter with "Enter Transm. Group or
' Comp. Code or Profit Center or Cost Center", which has nothing to do with the
' layout - it was reported as a failure in earlier builds. Same fields the macro
' itself fills. P_XMIT is explicitly cleared so nothing is transmitted.
Private Sub CM_FillContext(ByVal sess As Object, ByVal tcode As String, _
                           ByVal cc As String, ByVal mth As String, ByVal yr As String)
    On Error Resume Next
    Select Case UCase$(tcode)
        Case "ZGLRME"
            sess.findById("wnd[0]/usr/chkP_XMIT").Selected = False
            sess.findById("wnd[0]/usr/ctxtS_BUKRS-LOW").Text = cc
            sess.findById("wnd[0]/usr/txtP_MONAT").Text = mth
            sess.findById("wnd[0]/usr/txtP_GJAHR").Text = yr
        Case "ZGR215"
            sess.findById("wnd[0]/usr/ctxtSBUKRS").Text = cc
            sess.findById("wnd[0]/usr/txtSYEAR").Text = yr
    End Select
    Err.Clear
    On Error GoTo 0
End Sub


'--- is a "/..." display layout present, and is the field still on the screen? -
' Two things are established here:
'   1. the selection-screen field the macro drives still exists (definitive -
'      if it has moved or been renamed the close would fail on that line)
'   2. SAP accepts the layout name when it is entered
' Any SAP message is passed through verbatim, so a "fill in required fields"
' reply is visible as such rather than being mistaken for a missing layout.
Private Function CM_CheckLayout(ByVal sess As Object, ByVal tcode As String, _
                                ByVal fld As String, ByVal lay As String, _
                                ByVal cc As String, ByVal mth As String, _
                                ByVal yr As String) As String
    Dim o As Object
    On Error Resume Next
    CM_ClearPopups sess
    sess.findById("wnd[0]/tbar[0]/okcd").Text = "/n" & tcode
    sess.findById("wnd[0]").sendVKey 0
    If Err.Number <> 0 Then
        Err.Clear: On Error GoTo 0
        CM_CheckLayout = "?|could not open " & tcode
        Exit Function
    End If

    'give the selection screen what it needs, or it complains about that instead
    CM_FillContext sess, tcode, cc, mth, yr

    Set o = Nothing
    Set o = sess.findById("wnd[0]/usr/ctxt" & fld)
    If o Is Nothing Or Err.Number <> 0 Then
        Err.Clear: On Error GoTo 0
        CM_CheckLayout = "X|field " & fld & " is not on the " & tcode & " screen"
        Exit Function
    End If

    o.Text = lay
    'Enter only - a selection-screen round trip that validates the fields.
    'The report is NOT executed (no F8), so nothing is selected, written or posted.
    sess.findById("wnd[0]").sendVKey 0
    If Err.Number <> 0 Then
        Err.Clear: On Error GoTo 0
        CM_CheckLayout = "?|could not enter the layout on " & tcode
        Exit Function
    End If
    On Error GoTo 0
    CM_CheckLayout = CM_Verdict(sess)
End Function


'--- data-level (company-code) authorisation --------------------------------
' The object checks above prove the user can REACH each transaction. They do not
' prove the user may read a given company code's data - that is a separate
' authorisation and it only surfaces once a report actually selects data.
'
' ZGLRME is used as the probe because it is a pure reporting transaction and the
' macro's own read path clears P_XMIT, so nothing is transmitted and nothing is
' posted. Deliberately NOT used: ZGLGWUL (carries a P_POST checkbox that writes)
' and SM35 (processes batch-input sessions, i.e. it posts).
Private Function CM_CheckDataAuth(ByVal sess As Object, ByVal cc As String, _
                                  ByVal mth As String, ByVal yr As String) As String
    On Error Resume Next
    CM_ClearPopups sess
    sess.findById("wnd[0]/tbar[0]/okcd").Text = "/nzglrme"
    sess.findById("wnd[0]").sendVKey 0
    If Err.Number <> 0 Then
        Err.Clear: On Error GoTo 0
        CM_CheckDataAuth = "?|could not open ZGLRME"
        Exit Function
    End If

    sess.findById("wnd[0]/usr/chkP_XMIT").Selected = False    'read only - no transmit
    sess.findById("wnd[0]/usr/ctxtS_BUKRS-LOW").Text = cc
    sess.findById("wnd[0]/usr/txtP_MONAT").Text = mth
    sess.findById("wnd[0]/usr/txtP_GJAHR").Text = yr
    If Err.Number <> 0 Then
        Err.Clear: On Error GoTo 0
        CM_CheckDataAuth = "?|ZGLRME selection screen not as expected"
        Exit Function
    End If

    sess.findById("wnd[0]").sendVKey 8                        'execute (select only)
    If Err.Number <> 0 Then
        Err.Clear: On Error GoTo 0
        CM_CheckDataAuth = "?|could not execute ZGLRME"
        Exit Function
    End If
    On Error GoTo 0
    CM_CheckDataAuth = CM_Verdict(sess)
End Function


'--- format one result line and keep the tally ------------------------------
Private Function CM_Line(ByVal nm As String, ByVal res As String, _
                         ByRef nBad As Long) As String
    Dim p As Long, v As String, why As String
    p = InStr(res, "|")
    v = Left$(res, p - 1)
    why = Mid$(res, p + 1)
    If v = "OK" Then
        CM_Line = "     [OK] " & nm & vbCrLf
    ElseIf v = "X" Then
        nBad = nBad + 1
        CM_Line = "     [X]  " & nm & " - " & Left$(why, 55) & vbCrLf
    Else
        CM_Line = "     [~]  " & nm & " - " & Left$(why, 55) & vbCrLf
    End If
End Function


Public Function CM_SapAuthReport(ByVal sess As Object, _
                                 ByRef anyBlocked As Boolean, _
                                 ByVal runData As Boolean) As String
    Dim out As String, arr As Variant, prt As Variant, i As Long, nBad As Long
    Dim cc As String
    anyBlocked = False
    nBad = 0

    out = "--------------------------------------" & vbCrLf & _
          "SAP AUTHORISATIONS" & vbCrLf

    out = out & "   Transactions:" & vbCrLf
    arr = Split(CM_SAP_TCODES, ",")
    For i = LBound(arr) To UBound(arr)
        out = out & CM_Line(arr(i), CM_CheckTcode(sess, arr(i)), nBad)
    Next i

    out = out & "   Tables (via SE16):" & vbCrLf
    arr = Split(CM_SAP_TABLES, ",")
    For i = LBound(arr) To UBound(arr)
        out = out & CM_Line(arr(i), CM_CheckTable(sess, arr(i)), nBad)
    Next i

    out = out & "   GR55 report groups:" & vbCrLf
    arr = Split(CM_SAP_RGROUPS, ",")
    For i = LBound(arr) To UBound(arr)
        out = out & CM_Line(arr(i), CM_CheckReportGroup(sess, arr(i)), nBad)
    Next i

    On Error Resume Next
    cc = Trim$(CStr(Sheets("config").Range("B2").Value))
    On Error GoTo 0
    out = out & "   Report layouts (the ""/..."" variants):" & vbCrLf
    arr = Split(CM_SAP_LAYOUTS, ";")
    For i = LBound(arr) To UBound(arr)
        prt = Split(arr(i), "|")
        out = out & CM_Line(prt(0) & " " & prt(2) & " (" & prt(1) & ")", _
                            CM_CheckLayout(sess, CStr(prt(0)), CStr(prt(1)), CStr(prt(2)), _
                                           cc, CStr(Monthx), CStr(Yearx)), nBad)
    Next i

    If runData Then
        out = out & "   Company-code data access (ZGLRME, read-only):" & vbCrLf
        If cc = "" Then
            out = out & "     [~]  skipped - config!B2 (Company Code) is empty" & vbCrLf
        Else
            out = out & CM_Line("company code " & cc & ", period " & Monthx & "/" & Yearx, _
                                CM_CheckDataAuth(sess, cc, CStr(Monthx), CStr(Yearx)), nBad)
        End If
    End If

    'leave SAP on a neutral screen
    On Error Resume Next
    CM_ClearPopups sess
    sess.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    sess.findById("wnd[0]").sendVKey 0
    Err.Clear
    On Error GoTo 0

    If nBad > 0 Then
        anyBlocked = True
        out = out & "   => " & nBad & " SAP object(s) blocked - ask Security for the" & vbCrLf & _
                    "      missing authorisation before running the close." & vbCrLf
    Else
        out = out & "   => all SAP objects reachable." & vbCrLf
    End If
    If Not runData Then
        out = out & "   Note: company-code data access was not tested (skipped)." & vbCrLf
    End If

    CM_SapAuthReport = out
End Function
