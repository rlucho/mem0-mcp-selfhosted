Attribute VB_Name = "GlobalModule"
Global FPath As String, FTemp As String, Fmerger As String, Fmerged As String, Fprinted As String, FFinal As String
Global FShared As String, DiscN As String   'V4-CIO: promoted to module scope so every routine shares one working drive
Global LastDay As Date
Global Monthx, Yearx
Dim araj()
Dim datax
Dim Data
'============================================================================
' mCloseEnv_V4  -  Closing Manager (V4-CIO)
'----------------------------------------------------------------------------
' V4-CIO environment / path helpers (folded into GlobalModule for the baked build),
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
' Format:  [~] tcode | selection-screen field | layout name   (entries by ;)
' These are ALV display layouts (P_VARID / P_VARIE / P_ALV), not selection
' variants - the macro types them in and expects them to already exist.
'
' A leading "~" marks an entry as INFORMATIONAL: it is still attempted, but a
' negative answer is reported as [~] and never fails the preflight. ZGR215's
' /arek2 is one of these - the close reaches that layout field only after the
' document-number popup, so it cannot honestly be tested from a cold selection
' screen, and reporting [X] there sent people looking for a problem that was
' not real.
Public Const CM_SAP_LAYOUTS As String = _
    "ZGLRME|P_VARID|/default;ZGLRME|P_VARIE|/closing;" & _
    "ZGLRME|P_VARIE|/default;~ZGR215|P_ALV|/arek2"
'----------------------------------------------------------------------------

' --- Decimal notation SAP writes amounts in, when it cannot be worked out ------
' Leave as "auto" and the macro learns it from the first unambiguous amount it
' sees in a run (anything with both separators, or with 1, 2 or 4+ decimals) and
' applies it to any ambiguous one afterwards. Set it to "." or "," only if a site
' hits the ambiguity message - it then never has to guess or learn.
'   "auto"   work it out from the data          (default)
'   "."      SAP writes 1,234.56
'   ","      SAP writes 1.234,56
Public Const CM_SAP_DECIMAL As String = "auto"
'----------------------------------------------------------------------------


' Breadcrumb: the plain-language description of what the macro is doing right
' now. Set by CM_Note at each stage, shown live on Excel's status bar, and
' read back by CM_Explain when a run stops.
Public CM_Step As String
Public CM_StepNo As Long          'which stage we are on
Public CM_StepMax As Long         'how many stages the close has
Public CM_Started As Double       'Timer value when the run began
Public CM_DecSeen As String        'decimal separator learned during this run
Public CM_SrcFile As String        'which SAP extract is being read right now
Public CM_SrcNo As Long            'line number within it
Public CM_SrcLine As String        'that line, verbatim - it carries the document
                                   'number / account / profit centre SAP printed
'----------------------------------------------------------------------------

Sub CreatePaths()
'V4-CIO: one consistent working drive (CM_BASE_DRIVE), a local-path guard for
'OneDrive/SharePoint, and centralised folder creation. Callers keep using the
'same FPath / FTemp / Fmerger / ... globals exactly as before.

If Not AssertLocalWorkbook() Then End      'stop hard if the file was opened from the web

FPath = ThisWorkbook.Path
If Right(FPath, 1) <> "\" Then FPath = FPath & "\"

DiscN = CM_BASE_DRIVE                       'V4-CIO FIX: was "D:\ if present else C:\"; now one value everywhere
FTemp = DiscN & "pdf\temp"
Fmerger = DiscN & "pdf\merger"
Fmerged = DiscN & "pdf\mergedFiles"
Fprinted = DiscN & "pdf\printed"
FFinal = DiscN & "pdf\final"
FShared = CM_ARCHIVE_ROOT & Year(LastDay) & "\" & Right("0" & Month(LastDay), 2)

Call EnsureFolders                          'build \pdf\* + report tree, provision GiosPSMC.exe

End Sub
Sub CreateVariants()

LastDay = DateAdd("d", -Day(Date), Date)
Monthx = Month(LastDay)
Yearx = Year(LastDay)

End Sub
Sub CreateArray(nazwa)
    Call ImportTitle(nazwa)
    Call ProperArray
End Sub
Sub ImportTitle(nazwa)

Set strix = New ADODB.Stream

strix.Charset = "utf-8"
strix.Open
strix.LoadFromFile (nazwa)

FirstColumn = ""

'V4-CIO FIX: the only way out of this loop was finding a line that starts
'with "|". An empty or pipe-less export spun here for ever, hanging Excel
'with no message. It now stops at end of file and says which file.
Dim cmFound As Boolean
Do Until strix.EOS
    Data = strix.ReadText(-2)
    If VBA.Left(VBA.Trim(Data), 1) = "|" Then
        cmFound = True
        For i = 2 To Len(Trim(Data))
            If Mid(Trim(Data), i, 1) = "|" Then
                FirstColumn = Trim(Mid(Trim(Data), 2, i - 2))
                Exit For
            End If
        Next i
        Exit Do
    End If
Loop
If Not cmFound Then
    strix.Close
    Set strix = Nothing
    CM_Step = "reading the SAP export " & nazwa
    Err.Raise vbObjectError + 515, "ClosingManager", "EMPTY" & Chr$(1) & nazwa
End If

strix.Close
Set strix = Nothing

End Sub
Sub ProperArray()

Dim a, b, n
Dim cmLast As Long
Dim nStart
Dim nEnd

a = 1
n = 0
nStart = 1
nEnd = 1
Do
    If nStart >= VBA.Len(Data) Then Exit Do
    cmLast = nStart
    For a = nStart To VBA.Len(Data)
        If VBA.Mid(Data, a, 1) = "|" Then
            nStart = a + 1
            nEnd = nStart
            Exit For
        End If
    Next
    'V4-CIO FIX: nStart only moves when a "|" is found. With none left it
    'stayed put and this loop ran for ever, while ReDim Preserve grew the
    'array on every pass - a hang that ended in out-of-memory, if at all.
    If nStart = cmLast Then Exit Do
        
    For b = nEnd To VBA.Len(Data)
        If VBA.Mid(Data, b, 1) = "|" Then
            nEnd = b
            Exit For
        End If
    Next
    
    n = n + 1
    
    ReDim Preserve araj(3, n)
    
    araj(0, n - 1) = VBA.Trim(VBA.Mid(Data, nStart, nEnd - nStart))
    araj(1, n - 1) = nStart
    araj(2, n - 1) = nEnd
    
    nStart = nEnd
Loop

End Sub
Function getLineData(line, match, occurence) As String

If Left(Trim(line), 1) <> "|" Then
    getLineData = ""
    Exit Function
End If

Dim knt
Dim occur: occur = 0
For knt = LBound(araj, 2) To UBound(araj, 2)
    If VBA.Trim(VBA.LCase(araj(0, knt))) = VBA.LCase(match) Then
        occur = occur + 1
    End If
    If VBA.Trim(VBA.LCase(araj(0, knt))) = VBA.LCase(match) And occur = occurence Then
        If VBA.Mid(line, araj(1, knt) - 1, 1) = "|" And VBA.Mid(line, araj(2, knt), 1) = "|" Then
            getLineData = VBA.Trim(VBA.Mid(line, araj(1, knt), araj(2, knt) - araj(1, knt)))
        Else
            getLineData = czek4marker(line, araj(1, knt), araj(2, knt))
        End If
        Exit For
    End If
Next

If occur = 0 Then
    qpa = 0
End If
    
End Function

Function czek4marker(line, startx, endx)

Dim offset, starta, enda

offset = 0

Do
    If VBA.Mid(line, startx - 1 + VBA.Abs(offset), 1) = "|" Then
        starta = startx - 1 + VBA.Abs(offset) + 1
        Exit Do
    ElseIf VBA.Mid(line, startx - VBA.Abs(offset), 1) = "|" Then
        starta = startx - VBA.Abs(offset) + 1
        Exit Do
    Else
        offset = offset + 1
    End If
    If offset > 10 Then Exit Function
Loop
    
offset = 0
Do
    If VBA.Mid(line, endx + VBA.Abs(offset), 1) = "|" Then
        enda = endx + VBA.Abs(offset)
        Exit Do
    ElseIf VBA.Mid(line, endx - VBA.Abs(offset), 1) = "|" Then
        enda = endx - VBA.Abs(offset)
        Exit Do
    Else
        offset = offset + 1
    End If
    If offset > 10 Then Exit Function
Loop

czek4marker = VBA.Mid(line, starta, enda - starta)

End Function
Sub spGetList(List)

Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
xmlDoc.async = False

Url = CM_SP_BASE
List = "CCCrossList"

request = "<?xml version='1.0' encoding='utf-8'?>" & _
            "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'" & _
            " xmlns:xsd='http://www.w3.org/2001/XMLSchema'" & _
            " xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>" & _
            " <soap:Body>" & _
                "<GetListItems xmlns='http://schemas.microsoft.com/sharepoint/soap/'>" & _
                "<listName>" & List & "</listName>" & _
                        "<QueryOptions>" & _
                            "<IncludeMandatoryColumns>FALSE</IncludeMandatoryColumns>" & _
                        "</QueryOptions>" & _
                "<rowLimit>50000</rowLimit>" & _
                " </GetListItems>" & _
            " </soap:Body>" & _
            "</soap:Envelope>"

'post it up and look at the response
With CreateObject("Microsoft.XMLHTTP")
    .Open "Get", Url, False, "", ""
    .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    .setRequestHeader "SOAPAction", "http://schemas.microsoft.com/sharepoint/soap/GetListItems"
    .send request
    
    xmlDoc.LoadXML (.responsetext)
    
    Dim X
    For Each X In xmlDoc.getElementsByTagName("z:row")
        
        EmptRow = FindLastRow(1, 3, 1, 0, "config")
        Cells(EmptRow, 3) = X.getAttribute("ows_ID")
        Cells(EmptRow, 4) = X.getAttribute("ows_CC")
        Cells(EmptRow, 5) = X.getAttribute("ows_PostingBlock")
        Cells(EmptRow, 6) = X.getAttribute("ows_LocationClosed")
        Cells(EmptRow, 7) = X.getAttribute("ows_CPCAllowedGAAP")
        Cells(EmptRow, 8) = X.getAttribute("ows_CPCNotAllowed")
        Cells(EmptRow, 9) = X.getAttribute("ows_Comment")
        
    Next
End With

End Sub
Sub spClearList()

Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
xmlDoc.async = False

Url = CM_SP_BASE
List = "ProfitCenters"

w = 2
Do Until Sheets("config").Cells(w, 3) = ""
    updates = "<Batch> <Method ID='1' Cmd='Delete'>" & _
                    "<Field Name='ID'>" & _
                        Sheets("config").Cells(w, 3) & _
                    "</Field>" & _
                "</Method></Batch>"
    
    request = "<?xml version='1.0' encoding='utf-8'?>" & _
                "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'" & _
                " xmlns:xsd='http://www.w3.org/2001/XMLSchema'" & _
                " xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>" & _
                " <soap:Body>" & _
                    "<UpdateListItems xmlns='http://schemas.microsoft.com/sharepoint/soap/'>" & _
                        "<listName>" & List & "</listName>" & _
                        "<updates>" & updates & "</updates>" & _
                    " </UpdateListItems>" & _
                " </soap:Body>" & _
                "</soap:Envelope>"
    
    'post it up and look at the response
    With CreateObject("Microsoft.XMLHTTP")
    
        .Open "POST", Url, False, "", ""
        .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
        .setRequestHeader "SOAPAction", "http://schemas.microsoft.com/sharepoint/soap/UpdateListItems"
        .send request
        
    End With
    w = w + 1
Loop
End Sub
Sub spAddToList(updates, List)

Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
xmlDoc.async = False

Url = CM_SP_BASE

request = "<?xml version='1.0' encoding='utf-8'?>" & _
                "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'" & _
                " xmlns:xsd='http://www.w3.org/2001/XMLSchema'" & _
                " xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>" & _
                " <soap:Body>" & _
                    "<UpdateListItems xmlns='http://schemas.microsoft.com/sharepoint/soap/'>" & _
                        "<listName>" & List & "</listName>" & _
                        "<updates>" & updates & "</updates>" & _
                    " </UpdateListItems>" & _
                " </soap:Body>" & _
                "</soap:Envelope>"

With CreateObject("Microsoft.XMLHTTP")
    .Open "POST", Url, False, "", ""
    .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    .setRequestHeader "SOAPAction", "http://schemas.microsoft.com/sharepoint/soap/UpdateListItems"
    .send request
End With

End Sub
Sub spUpdateList(updates, List)

Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
xmlDoc.async = False

Url = CM_SP_BASE

request = "<?xml version='1.0' encoding='utf-8'?>" & _
                "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'" & _
                " xmlns:xsd='http://www.w3.org/2001/XMLSchema'" & _
                " xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>" & _
                " <soap:Body>" & _
                    "<UpdateListItems xmlns='http://schemas.microsoft.com/sharepoint/soap/'>" & _
                        "<listName>" & List & "</listName>" & _
                        "<updates>" & updates & "</updates>" & _
                    " </UpdateListItems>" & _
                " </soap:Body>" & _
                "</soap:Envelope>"

With CreateObject("Microsoft.XMLHTTP")
    .Open "POST", Url, False, "", ""
    .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    .setRequestHeader "SOAPAction", "http://schemas.microsoft.com/sharepoint/soap/UpdateListItems"
    .send request
End With

End Sub
Sub SAPSelectFields(sess As Object, k As Long)

LastRow = FindLastRow(1, k, 0, 0, "SAP config")

j = 1
Do Until LastRow = j
    Set Area = sess.findById("wnd[1]/usr")
    Set Children = Area.Children()
    For i = 0 To Children.Count() - 1
        Set obj = Children(CInt(i))
        If obj.Type = "GuiLabel" And obj.Text <> "" Then
            w = 2
            Do Until Sheets("SAP config").Cells(w, k) = ""
                FieldName = Sheets("SAP config").Cells(w, k)
                If obj.Text = FieldName Then
                    Set obj = Children(CInt(i - 1))
                    obj.Selected = True
                    Set obj = Children(CInt(i))
                    j = j + 1
                End If
                w = w + 1
            Loop
        End If
    Next
    sess.findById("wnd[1]").sendVKey 82
Loop

End Sub

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
    Dim cc As String, dst As String, winPrn As String
    Dim pdfOK As Boolean, prtBlocked As Boolean, ans As VbMsgBoxResult
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
    pdfOK = fso.FileExists("C:\Program Files\PDFCreator\PDFCreator.exe") Or _
            fso.FileExists("C:\Program Files (x86)\PDFCreator\PDFCreator.exe")
    If pdfOK Then
        msg = msg & "[OK] PDFCreator is installed." & vbCrLf
    Else
        msg = msg & "[X] PDFCreator.exe not found (printer must be named 'PDFCreator')." & vbCrLf
        okAll = False
    End If

    ' 3b) where a failure will be written -------------------------------------
    msg = msg & "[OK] Failures are logged to " & CM_LogPath() & "." & vbCrLf

    ' 4a) how SAP writes amounts -------------------------------------------------
    If CM_DecimalSep() = "" Then
        msg = msg & "[OK] SAP decimal notation: worked out automatically at run time." & vbCrLf
    Else
        msg = msg & "[OK] SAP decimal notation is set to """ & CM_DecimalSep() & """ " & _
                    "(1" & IIf(CM_DecimalSep() = ".", ",", ".") & "234" & CM_DecimalSep() & "56)." & vbCrLf
    End If

    ' 4b) the Windows printer itself - SAP is told to print to the name below,
    '     so a renamed printer fails the close even with PDFCreator installed.
    winPrn = CM_PdfCreatorPrinter()
    If winPrn <> "" Then
        msg = msg & "[OK] Windows printer found: " & winPrn & "." & vbCrLf
    Else
        msg = msg & "[X] No Windows printer whose name contains 'PDFCreator'." & vbCrLf
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

    ' 8) Print + merge rehearsal (optional - it really prints, so it asks) -----
    If Not IsUrlPath(ThisWorkbook.Path) Then
        ans = MsgBox("Test printing and merging for real?" & vbCrLf & vbCrLf & _
              "YES     print two test pages, then merge them." & vbCrLf & _
              "        Proves the whole chain: SAP -> PDFCreator -> " & CM_BASE_DRIVE & "pdf\temp" & vbCrLf & _
              "        -> GiosPSMC.exe. Allow up to two minutes." & vbCrLf & vbCrLf & _
              "NO      test the PDF merger only." & vbCrLf & _
              "        Uses two built-in test pages, so nothing is printed and" & vbCrLf & _
              "        PDFCreator is not involved. Takes a few seconds." & vbCrLf & vbCrLf & _
              "CANCEL  skip this check." & vbCrLf & vbCrLf & _
              "Either test writes only under " & CM_BASE_DRIVE & "pdf\ and clears up after" & vbCrLf & _
              "itself. Printing uses SE16/T001 - a display, so no data is created" & vbCrLf & _
              "or changed in SAP.", _
              vbYesNoCancel + vbQuestion, "Closing Manager - Preflight")

        If ans = vbCancel Then
            msg = msg & "[~] Print/merge not tested (skipped)." & vbCrLf
        Else
            msg = msg & CM_PrintMergeTest(sapSess, prtBlocked, (ans = vbYes))
            If prtBlocked Then okAll = False
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
    Dim info As Boolean, res As String
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
        info = (Left$(arr(i), 1) = "~")
        If info Then arr(i) = Mid$(arr(i), 2)
        prt = Split(arr(i), "|")
        res = CM_CheckLayout(sess, CStr(prt(0)), CStr(prt(1)), CStr(prt(2)), _
                             cc, CStr(Monthx), CStr(Yearx))
        If info And Left$(res, 2) <> "OK" Then
            'informational: the close picks this layout up later in the dialogue
            res = "?|only reachable after the document-number popup - not a fault"
        End If
        out = out & CM_Line(prt(0) & " " & prt(2) & " (" & prt(1) & ")", res, nBad)
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


'============================================================================
' PLAIN-LANGUAGE FAILURE REPORTING
'----------------------------------------------------------------------------
' When the close stops, the operator used to get a raw VBA dialog ("Run-time
' error 13: Type mismatch") with no clue what the macro was doing or whether
' anything had been posted. CM_Note leaves a breadcrumb at each stage and
' CM_Explain turns the failure into something a finance user can act on.
'============================================================================


'--- live progress -----------------------------------------------------------
' The close spends minutes at a time inside SAP, and the old code gave no sign
' of life: users could not tell "still working" from "frozen". Every stage now
' reports itself on Excel's status bar, and every wait ticks a counter, so a
' run that is simply slow looks different from a run that is stuck.

'--- start of a run: reset the counters --------------------------------------
Public Sub CM_Begin(ByVal totalSteps As Long)
    CM_DecSeen = ""
    CM_SrcFile = "": CM_SrcNo = 0: CM_SrcLine = ""
    CM_StepNo = 0
    CM_StepMax = totalSteps
    CM_Started = Timer
    CM_Step = "starting the close"
    CM_Paint ""
End Sub


'--- leave a breadcrumb: what the macro is doing right now -------------------
Public Sub CM_Note(ByVal what As String)
    CM_Step = what
    CM_StepNo = CM_StepNo + 1
    If CM_Started = 0 Then CM_Started = Timer
    CM_Paint ""
End Sub


'--- end of a run: hand the status bar back to Excel -------------------------
Public Sub CM_Done()
    CM_Step = ""
    CM_StepNo = 0
    CM_StepMax = 0
    CM_Started = 0
    On Error Resume Next
    Application.StatusBar = False
    Err.Clear
    On Error GoTo 0
End Sub


'--- draw the status bar -----------------------------------------------------
Public Sub CM_Paint(ByVal extra As String)
    Dim s As String, el As Long
    el = CM_Elapsed()
    s = "CLOSING MANAGER"
    If CM_StepMax > 0 Then s = s & "   [" & CM_StepNo & "/" & CM_StepMax & "]"
    s = s & "   " & CM_Step
    If extra <> "" Then s = s & "   -   " & extra
    s = s & "      (running " & Format$(el \ 60, "00") & ":" & Format$(el Mod 60, "00") & _
        " - press Esc to stop)"
    On Error Resume Next
    Application.StatusBar = s
    Err.Clear
    On Error GoTo 0
    DoEvents
End Sub


'--- seconds since the run started (safe across midnight) --------------------
Private Function CM_Elapsed() As Long
    Dim t As Double
    If CM_Started = 0 Then Exit Function
    t = Timer - CM_Started
    If t < 0 Then t = t + 86400
    CM_Elapsed = CLng(t)
End Function


'--- one second of waiting that keeps Excel alive ----------------------------
Private Sub CM_Tick(ByVal extra As String)
    CM_Paint extra
    Application.Wait (Now + TimeValue("0:00:01"))
    DoEvents
End Sub


'--- file size, or -1 if the file is not there right now ---------------------
Private Function CM_FileSize(ByVal fso As Object, ByVal p As String) As Double
    On Error Resume Next
    CM_FileSize = -1
    CM_FileSize = fso.GetFile(p).Size
    Err.Clear
    On Error GoTo 0
End Function


'============================================================================
' WAIT FOR A PRINTED PDF
'----------------------------------------------------------------------------
' Replaces fifteen copies of this loop - ten in Printing, five in Postings:
'
'     Do Until File <> ""
'         Set objShell = CreateObject("Shell.Application")
'         Set objFolder = objShell.Namespace(FTemp & "\")
'         ...
'     Loop
'
' If PDFCreator never dropped a file, the folder was empty, the For Each body
' never ran, File stayed "", and the loop went straight back to CreateObject
' with NO delay: a tight spin that pegged a CPU core, left Excel showing "Not
' Responding", and hammered the shell COM server until it dropped the
' connection - which is where "Run-time error -2147417848 (80010108): Method
' 'NameSpace' of object 'IShellDispatch6' failed" came from. The disconnect was
' the symptom; the missing PDF was the cause, and the run could never recover.
'
' This version uses Dir (no COM server to lose), waits a bounded time, keeps
' Excel responsive and reporting, and if the PDF really never arrives it stops
' with an explained error instead of hanging for ever.
'============================================================================
Public Function CM_WaitForPrint(ByVal folder As String, ByVal fso As Object, _
                                ByVal what As String) As String
    Const APPEAR_MAX As Long = 240      'seconds to wait for the PDF to appear
    Const SETTLE_MAX As Long = 180      'seconds to wait for it to stop growing
    Dim base As String, nm As String, p As String
    Dim waited As Long, s1 As Double, s2 As Double

    base = folder
    If Right$(base, 1) <> "\" Then base = base & "\"

    'wait for PDFCreator to drop a file into the temp folder
    Do
        nm = Dir(base & "*.*")
        If nm <> "" Then Exit Do
        If waited >= APPEAR_MAX Then
            Err.Raise vbObjectError + 514, "ClosingManager", _
                      "PRINT" & Chr$(1) & what & Chr$(1) & waited & Chr$(1) & base
        End If
        CM_Tick "waiting for the PDF of " & what & " (" & waited & "s)"
        waited = waited + 1
    Loop

    p = base & nm

    'wait for it to stop growing (PDFCreator may still be writing it)
    waited = 0
    Do
        s1 = CM_FileSize(fso, p)
        CM_Tick "writing the PDF of " & what & " (" & waited & "s)"
        s2 = CM_FileSize(fso, p)
        If s1 >= 0 And s1 = s2 And s1 > 0 Then Exit Do
        If s1 < 0 Then
            'renamed under us (temp name -> final name): pick up what is there
            nm = Dir(base & "*.*")
            If nm <> "" Then p = base & nm
        End If
        waited = waited + 1
        If waited >= SETTLE_MAX Then Exit Do
    Loop

    If Not fso.FileExists(p) Then
        Err.Raise vbObjectError + 514, "ClosingManager", _
                  "PRINT" & Chr$(1) & what & Chr$(1) & (APPEAR_MAX + SETTLE_MAX) & Chr$(1) & base
    End If

    CM_WaitForPrint = p
End Function



'============================================================================
' READING A SAP AMOUNT
'----------------------------------------------------------------------------
' SAP writes negatives as "1.234,56-" and formats numbers using the SAP user's
' decimal notation, which need not match this PC's Windows regional settings.
' VBA's Round()/CDbl()/implicit conversion all read Windows, so a mismatch
' raised "Run-time error 13: Type mismatch" - or, worse, quietly read
' "1.234,56" as 1.23456.
'
' This reads the literal instead of the locale. The rules, in order:
'
'   already a number          use it unchanged (worksheet-sourced values)
'   both "." and ","          the LAST one is the decimal separator
'   one separator, repeated   it must be grouping     1.234.567 -> 1234567
'   one separator, 3 digits   AMBIGUOUS - see below
'   one separator, otherwise  it is the decimal point  1234,56  -> 1234.56
'
' AMBIGUITY IS NOT GUESSED. "1,234" is 1234 to an English reader and 1.234 to a
' German one, and nothing in the string says which. Guessing would put a figure
' wrong by a factor of 1000 into a signed report pack, so the close stops and
' says so. This cannot fire on a 2-decimal currency, which is every currency
' this close handles - the last separator is followed by two digits, not three.
'============================================================================
Public Function CM_ToAmount(ByVal v As Variant, ByRef ok As Boolean, _
                            ByVal blankOk As Boolean, ByRef reason As String) As Double
    Dim s As String, neg As Boolean, i As Long, c As String
    Dim dots As Long, pDot As Long, pCom As Long, nSep As Long, after As Long
    Dim sep As String, other As String
    ok = False
    reason = "AMOUNT"
    CM_ToAmount = 0

    'a value that is already numeric (a worksheet cell, say) needs no parsing
    Select Case VarType(v)
        Case vbDouble, vbSingle, vbCurrency, vbInteger, vbLong, vbByte, vbDecimal
            CM_ToAmount = CDbl(v)
            ok = True
            Exit Function
        Case vbBoolean, vbDate, vbObject
            reason = "AMOUNT"
            Exit Function
    End Select

    s = Trim$(CStr(v))
    s = Replace(s, " ", "")
    s = Replace(s, Chr$(160), "")          'SAP groups with a hard space in places

    If s = "" Then
        If blankOk Then
            ok = True                       'no amount on this row = nothing to add
        Else
            reason = "BLANK"
        End If
        Exit Function
    End If

    If Right$(s, 1) = "-" Then
        neg = True
        s = Left$(s, Len(s) - 1)
    End If
    If Left$(s, 1) = "-" Then
        neg = True
        s = Mid$(s, 2)
    End If
    If s = "" Then Exit Function

    pDot = InStrRev(s, ".")
    pCom = InStrRev(s, ",")

    If pDot > 0 And pCom > 0 Then
        'Both kinds present: whichever comes last is the decimal separator. This
        'is the most decisive evidence there is, so it is what the run learns from.
        If pDot > pCom Then
            If Not CM_Learn(".") Then reason = "CLASH": Exit Function
            s = Replace(s, ",", "")
        Else
            If Not CM_Learn(",") Then reason = "CLASH": Exit Function
            s = Replace(s, ".", "")
            s = Replace(s, ",", ".")
        End If

    ElseIf pDot > 0 Or pCom > 0 Then
        If pDot > 0 Then
            sep = ".": other = ",": after = Len(s) - pDot
        Else
            sep = ",": other = ".": after = Len(s) - pCom
        End If
        nSep = CM_CountChar(s, sep)

        If nSep > 1 Then
            'a number has one decimal point, so a repeated separator is grouping -
            'which also tells us the OTHER character is this system's decimal point
            If Not CM_Learn(other) Then reason = "CLASH": Exit Function
            s = Replace(s, sep, "")

        ElseIf after <> 3 Then
            '1, 2 or 4+ digits after it: it is the decimal point, unambiguously
            If Not CM_Learn(sep) Then reason = "CLASH": Exit Function
            If sep = "," Then s = Replace(s, ",", ".")

        Else
            'Exactly three digits after a single separator, and nothing else in
            'the string to go on: "1,234" is 1234 to an English reader and 1.234
            'to a German one. Resolve it from CM_SAP_DECIMAL, or from what this
            'run has already learned from unambiguous values. Only if neither is
            'known does the close stop - it will not put a guessed figure into a
            'signed report pack.
            If CM_DecimalSep() = "" Then
                reason = "AMBIG"
                Exit Function
            End If
            If CM_DecimalSep() = sep Then
                If sep = "," Then s = Replace(s, ",", ".")   'it is the decimal point
            Else
                s = Replace(s, sep, "")                      'it is grouping
            End If
        End If
    End If

    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If c = "." Then
            dots = dots + 1
            If dots > 1 Then Exit Function
        ElseIf c < "0" Or c > "9" Then
            Exit Function
        End If
    Next i

    CM_ToAmount = Val(s)                    'Val never reads the locale
    If neg Then CM_ToAmount = -CM_ToAmount
    ok = True
End Function



'============================================================================
' WHERE THE BAD VALUE CAME FROM
'----------------------------------------------------------------------------
' Knowing that "1.234" could not be read is only half an answer - the operator
' still has to find it. Every extract the close reads is announced with
' CM_Source, and every line with CM_Reading, so a failure can name the file, the
' line number and the line itself. SAP prints the document number, account and
' profit centre on that same line, which is what somebody actually needs to look
' the figure up.
'============================================================================
Public Sub CM_Source(ByVal fileName As String)
    CM_SrcFile = fileName
    CM_SrcNo = 0
    CM_SrcLine = ""
End Sub


Public Sub CM_Reading(ByVal rawLine As String)
    CM_SrcNo = CM_SrcNo + 1
    If Len(rawLine) > 300 Then
        CM_SrcLine = Left$(rawLine, 300) & " ..."
    Else
        CM_SrcLine = rawLine
    End If
End Sub



'--- record the convention; False if it contradicts what we already knew -----
' One SAP user writes amounts one way, so two extracts in one close cannot
' disagree. If they do, something is wrong that guessing would only hide - the
' close stops rather than read one literal two different ways in one run.
Private Function CM_Learn(ByVal sep As String) As Boolean
    If CM_SAP_DECIMAL = "." Or CM_SAP_DECIMAL = "," Then
        CM_Learn = (CM_SAP_DECIMAL = sep)      'pinned by configuration
        Exit Function
    End If
    If CM_DecSeen <> "" And CM_DecSeen <> sep Then Exit Function
    CM_DecSeen = sep
    CM_Learn = True
End Function


'--- which character SAP is using as the decimal point, "" if not yet known --
' An explicit CM_SAP_DECIMAL always wins; otherwise it is whatever this run has
' worked out from an unambiguous amount. Public so PreflightCheck can show it.
Public Function CM_DecimalSep() As String
    If CM_SAP_DECIMAL = "." Or CM_SAP_DECIMAL = "," Then
        CM_DecimalSep = CM_SAP_DECIMAL
    Else
        CM_DecimalSep = CM_DecSeen
    End If
End Function


'--- how many times a character occurs --------------------------------------
Private Function CM_CountChar(ByVal s As String, ByVal c As String) As Long
    CM_CountChar = Len(s) - Len(Replace(s, c, ""))
End Function


'--- amount that MAY legitimately be blank (a worksheet cell with no value) --
' Used for the ZGLRME extract, where the original code also treated an empty
' cell as nothing to add. A blank there is a row with no amount, not a fault.
Public Function CM_Amount(ByVal v As Variant, ByVal rowNo As Long, _
                          ByVal where As String) As Double
    CM_Amount = CM_AmountCore(v, rowNo, where, True)
End Function


'--- amount that MUST be there ----------------------------------------------
' Used everywhere a value is parsed out of a SAP text export. The original code
' raised "Type mismatch" on a blank there and stopped the close; so does this,
' with an explanation instead of a bare error number. Do NOT relax this - a
' missing amount silently read as zero would let an out-of-balance close pass.
Public Function CM_AmountReq(ByVal v As Variant, ByVal rowNo As Long, _
                             ByVal where As String) As Double
    CM_AmountReq = CM_AmountCore(v, rowNo, where, False)
End Function


Private Function CM_AmountCore(ByVal v As Variant, ByVal rowNo As Long, _
                               ByVal where As String, ByVal blankOk As Boolean) As Double
    Dim ok As Boolean, d As Double, reason As String
    d = CM_ToAmount(v, ok, blankOk, reason)
    If Not ok Then
        CM_Step = where
        'Chr(1) separates the fields - a SAP report line is full of "|"
        Err.Raise vbObjectError + 513, "ClosingManager", _
                  reason & Chr$(1) & CStr(v) & Chr$(1) & rowNo & Chr$(1) & _
                  CM_SrcFile & Chr$(1) & CM_SrcNo & Chr$(1) & CM_SrcLine
    End If
    CM_AmountCore = d
End Function


'--- turn a failure into plain language --------------------------------------
Public Sub CM_Explain(ByVal errNum As Long, ByVal errDesc As String)
    Dim kind As String, what As String, why As String, todo As String
    Dim parts As Variant, m As String, whereTo As String, logPath As String

    'our own explained data errors carry their detail in the description
    'the printed PDF never arrived (was: Excel hangs for ever)
    If errNum = vbObjectError + 514 And Left$(errDesc, 6) = "PRINT" & Chr$(1) Then
        parts = Split(errDesc, Chr$(1))
        kind = "PRINTING"
        what = "SAP was asked to print a report, but no PDF ever appeared." & vbCrLf & _
               "        The report was:  " & parts(1) & vbCrLf & _
               "        Waited:          " & parts(2) & " seconds" & vbCrLf & _
               "        Watching folder: " & parts(3)
        why = "The print left SAP but PDFCreator did not finish the file. Usually" & vbCrLf & _
              "        one of: PDFCreator is showing a dialog waiting for an answer, it" & vbCrLf & _
              "        is not set to save automatically to the folder above, it is not" & vbCrLf & _
              "        running, or SAP put up a pop-up so nothing was actually printed."
        todo = "1. Look for a PDFCreator window behind Excel and close/answer it." & vbCrLf & _
               "        2. Check SAP for a pop-up left open." & vbCrLf & _
               "        3. Run Preflight Check - it confirms PDFCreator is installed." & vbCrLf & _
               "        4. Empty the folder above of any leftover files, then run again." & vbCrLf & _
               "        Earlier reports that already printed are kept - the run can be" & vbCrLf & _
               "        repeated from the START sheet."
        GoTo Report
    End If

    'a SAP export came back with no data at all (was: an unbreakable hang)
    If errNum = vbObjectError + 515 And Left$(errDesc, 6) = "EMPTY" & Chr$(1) Then
        parts = Split(errDesc, Chr$(1))
        kind = "DATA"
        what = "A report SAP was asked for came back with nothing in it." & vbCrLf & _
               "        The file was:  " & parts(1)
        why = "SAP selected no rows at all. Usually the company code or the" & vbCrLf & _
              "        period on the START sheet does not match anything posted," & vbCrLf & _
              "        the period is not open, or the report was cancelled in SAP" & vbCrLf & _
              "        before it finished."
        todo = "1. Check the company code and period on the START sheet." & vbCrLf & _
               "        2. Run the same report by hand in SAP for that company code" & vbCrLf & _
               "           and period - if it is empty there too, the data is not" & vbCrLf & _
               "           posted yet and the close cannot run." & vbCrLf & _
               "        3. If SAP shows rows but this file is empty, send this" & vbCrLf & _
               "           message to the CI Team."
        GoTo Report
    End If

    'the shell COM object dropped out - only ever seen as a knock-on of the above
    If errNum = -2147417848 Or errNum = -2147417851 Or errNum = 462 Then
        kind = "TECHNICAL"
        what = "The link between Excel and Windows was lost part-way through the run."
        why = "In earlier versions this followed a print that never completed: the" & vbCrLf & _
              "        macro spun without pausing until Windows dropped the connection." & vbCrLf & _
              "        That loop is fixed in this version, so if you are seeing this" & vbCrLf & _
              "        message the CI Team wants to know."
        todo = "Close Excel completely (Task Manager if it will not close), reopen" & vbCrLf & _
               "        the workbook and run again. If it happens twice, send this message" & vbCrLf & _
               "        to the Continuous Improvement Team."
        GoTo Report
    End If

    If errNum = vbObjectError + 513 And InStr(errDesc, Chr$(1)) > 0 Then
        parts = Split(errDesc, Chr$(1))
        kind = "DATA"
        Select Case parts(0)
            Case "BLANK"
                what = "An amount SAP should have sent was blank."
            Case "CLASH"
                what = "Two SAP reports in this run wrote amounts in different" & vbCrLf & _
                       "        number formats, which cannot both be right." & vbCrLf & _
                       "        The value that disagreed:  [" & parts(1) & "]"
            Case "AMBIG"
                what = "SAP sent an amount that could mean two things, and" & vbCrLf & _
                       "        nothing else in this run said which." & vbCrLf & _
                       "        The value was:  [" & parts(1) & "]" & vbCrLf & _
                       "        That is either " & Replace(Replace(parts(1), ".", ""), ",", "") & _
                       " or about " & Split(Replace(parts(1), ",", "."), ".")(0) & "." & vbCrLf & _
                       "        The macro will not guess a figure for a signed report."
            Case Else
                what = "SAP sent an amount the macro could not read as a number." & vbCrLf & _
                       "        The value was:  [" & parts(1) & "]"
        End Select
        If Val(parts(2)) > 0 Then what = what & vbCrLf & "        On data row:    " & parts(2)
        If UBound(parts) >= 5 Then
            If parts(3) <> "" Then
                whereTo = "Extract file:   " & FPath & parts(3) & vbCrLf & _
                          "        Line " & parts(4) & " of that file. It is still there - the run" & vbCrLf & _
                          "        stopped before deleting it." & vbCrLf & _
                          "        This is the line SAP sent (it carries the document" & vbCrLf & _
                          "        number, account and profit centre):" & vbCrLf & _
                          "        " & parts(5)
            ElseIf Val(parts(2)) > 0 Then
                whereTo = "Sheet ""ZGLRME"", row " & parts(2) & ", the amount column." & vbCrLf & _
                          "        Open that sheet in this workbook and look at the row."
            End If
        End If
        why = "Almost always the number format. SAP writes amounts using the" & vbCrLf & _
              "        SAP user's decimal notation, and this PC reads them using its" & vbCrLf & _
              "        Windows regional settings. If one uses 1.234,56 and the other" & vbCrLf & _
              "        expects 1,234.56, the amount cannot be read."
        If parts(0) = "CLASH" Then
            todo = "Send this message to the CI Team. One of the SAP reports is" & vbCrLf & _
                   "        being produced with a different decimal notation from the" & vbCrLf & _
                   "        others, which usually means a report variant or a user" & vbCrLf & _
                   "        default was changed. Setting CM_SAP_DECIMAL does not fix" & vbCrLf & _
                   "        this - the reports themselves disagree."
        ElseIf parts(0) = "AMBIG" Then
            todo = "This one is settled permanently by a single setting." & vbCrLf & _
                   "        Look up SAP: System > User Profile > Own Data >" & vbCrLf & _
                   "        Defaults > Decimal Notation, then ask the CI Team to set" & vbCrLf & _
                   "        CM_SAP_DECIMAL in the macro to ""."" or "","" to match." & vbCrLf & _
                   "        After that the macro never has to work it out."
        ElseIf parts(0) = "BLANK" Then
            todo = "The amount column was empty where a figure was expected." & vbCrLf & _
                   "        Re-run the SAP report by hand for this company code and" & vbCrLf & _
                   "        period and check the column really has a value. If it does," & vbCrLf & _
                   "        send this message to the CI Team."
        Else
            todo = "If the value above is not a number at all (for example *****)," & vbCrLf & _
                   "        the SAP report column is too narrow to show the figure -" & vbCrLf & _
                   "        tell the CI Team. Otherwise send them this message: the" & vbCrLf & _
                   "        macro reads both 1.234,56 and 1,234.56, so a value it" & vbCrLf & _
                   "        cannot read is not a regional-settings problem."
        End If
    Else
        Select Case errNum
            Case 13, 6, 11
                kind = "DATA"
                what = "A value coming back from SAP was not the kind of value" & vbCrLf & _
                       "        the macro expected."
                why = "Every amount the close reads is now converted independently" & vbCrLf & _
                      "        of Windows regional settings, so this is NOT the usual" & vbCrLf & _
                      "        SAP-vs-Windows number-format difference. It is more" & vbCrLf & _
                      "        likely a report column showing ***** instead of a figure," & vbCrLf & _
                      "        or a SAP screen returning something unexpected."
                todo = "Run the SAP report by hand for this company code and period" & vbCrLf & _
                       "        and look for a column of asterisks or a missing value." & vbCrLf & _
                       "        Then send this whole message to the CI Team."
            Case 9
                kind = "DATA"
                what = "The macro expected a list of data and found it empty or shorter" & vbCrLf & _
                       "        than expected."
                why = "Usually SAP returned no rows for this company code and period -" & vbCrLf & _
                      "        for example the period is not open, or the data is not posted yet."
                todo = "Check the company code and period on the START sheet, and that" & vbCrLf & _
                       "        the period really is open in SAP. Then run again."
            Case 70, 75, 76, 53, 3004, 3002, 3001
                kind = "FILE"
                what = "The macro could not write or read one of its small working files."
                why = "These are written into the folder this workbook sits in. The folder" & vbCrLf & _
                      "        may be read-only, synced by OneDrive, or a leftover file from a" & vbCrLf & _
                      "        previous run may still be locked."
                todo = "1. Copy this workbook to a plain local folder such as C:\Closing\" & vbCrLf & _
                       "        2. Delete any leftover .csv / .txt files sitting beside it." & vbCrLf & _
                       "        3. Close and reopen Excel, then run again."
            Case 424, 438, 91, 462, 619
                kind = "SAP"
                what = "SAP was not showing the screen the macro expected."
                why = "Usually one of: you are not authorised for the transaction, an" & vbCrLf & _
                      "        unexpected SAP pop-up appeared, the SAP session was closed, or" & vbCrLf & _
                      "        the screen changed after an SAP update."
                todo = "1. Run Preflight Check (Preflight sheet) - it tests every" & vbCrLf & _
                       "           transaction the close needs." & vbCrLf & _
                       "        2. Make sure only one SAP window is open and you are logged in." & vbCrLf & _
                       "        3. If Preflight is all green, send this message to the CI Team."
            Case 1004
                kind = "EXCEL"
                what = "The macro could not read or write one of the sheets in this workbook."
                why = "A sheet may have been renamed, deleted, or is protected."
                todo = "Use a fresh copy of the workbook and run again."
            Case Else
                kind = "TECHNICAL"
                what = "The macro stopped with an unexpected error."
                why = "This one is not a known data problem - it needs a look by the CI Team."
                todo = "Send this whole message to the Continuous Improvement Team."
        End Select
    End If

Report:
    If CM_Step = "" Then CM_Step = "(not recorded)"

    m = "CLOSING MANAGER - COULD NOT CONTINUE" & vbCrLf & _
        "--------------------------------------------" & vbCrLf & vbCrLf & _
        "WHAT IT WAS DOING" & vbCrLf & _
        "        " & CM_StepLabel() & vbCrLf & vbCrLf & _
        "WHAT WENT WRONG   (" & kind & ")" & vbCrLf & _
        "        " & what & vbCrLf & vbCrLf & _
        "WHY THIS USUALLY HAPPENS" & vbCrLf & _
        "        " & why & vbCrLf & vbCrLf & _
        IIf(whereTo = "", "", "WHERE TO LOOK" & vbCrLf & "        " & whereTo & vbCrLf & vbCrLf) & _
        "WHAT TO TRY" & vbCrLf & _
        "        " & todo & vbCrLf & vbCrLf & _
        "--------------------------------------------" & vbCrLf & _
        "This failure itself posted nothing. Anything already posted to SAP" & vbCrLf & _
        "earlier in this run stays posted - check before running again." & vbCrLf & vbCrLf & _
        "For the CI Team:  error " & errNum & " - " & errDesc

    logPath = CM_SaveLog(m & vbCrLf & "For the CI Team:  error " & errNum & " - " & errDesc)
    If logPath <> "" Then
        m = m & vbCrLf & vbCrLf & "A copy of this message was saved to:" & vbCrLf & logPath
    Else
        m = m & vbCrLf & vbCrLf & "This message could not be saved to a file - please copy it" & vbCrLf & _
                "(Ctrl+C works on this dialog) before clicking OK."
    End If

    On Error Resume Next
    Application.StatusBar = False
    Err.Clear
    On Error GoTo 0

    MsgBox m, vbExclamation, "Closing Manager - stopped"
End Sub


'--- keep the message after the dialog is dismissed --------------------------
' Appended, not overwritten, so a sequence of failures across a close can be
' sent to the CI Team in one go. Returns the path, or "" if it could not write.
Private Function CM_SaveLog(ByVal text As String) As String
    'Beside the workbook first - that is where the SAP extracts land too, so
    'everything needed to investigate a failure sits in one folder. If that
    'folder will not take it (read-only share, or the workbook opened from the
    'web) fall back to the working folder, which the macro creates itself.
    CM_SaveLog = CM_TryLog(CM_LogFolder(), text)
    If CM_SaveLog <> "" Then Exit Function
    CM_SaveLog = CM_TryLog(CM_BASE_DRIVE & "pdf\", text)
End Function


'--- where the log goes: beside the workbook, "" if that is not usable --------
Public Function CM_LogFolder() As String
    Dim p As String
    On Error Resume Next
    p = ThisWorkbook.Path
    Err.Clear
    On Error GoTo 0
    If p = "" Then Exit Function
    If IsUrlPath(p) Then Exit Function
    If Right$(p, 1) <> "\" Then p = p & "\"
    CM_LogFolder = p
End Function


'--- the file the next failure will be written to ----------------------------
Public Function CM_LogPath() As String
    Dim p As String
    p = CM_LogFolder()
    If p = "" Then p = CM_BASE_DRIVE & "pdf\"
    CM_LogPath = p & "ClosingManager_errors.log"
End Function


Private Function CM_TryLog(ByVal folder As String, ByVal text As String) As String
    Dim f As Integer, p As String
    If folder = "" Then Exit Function
    On Error Resume Next
    p = folder & "ClosingManager_errors.log"
    f = FreeFile
    Open p For Append As #f
    Print #f, String$(70, "=")
    Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & "   user " & Environ$("username")
    Print #f, text
    Print #f, ""
    Close #f
    If Err.Number = 0 Then CM_TryLog = p
    Err.Clear
    On Error GoTo 0
End Function


'--- "step 12 of 27 - printing ZGLRME" ---------------------------------------
Public Function CM_StepLabel() As String
    If CM_Step = "" Then
        CM_StepLabel = "(not recorded)"
    ElseIf CM_StepMax > 0 Then
        CM_StepLabel = "step " & CM_StepNo & " of " & CM_StepMax & " - " & CM_Step
    Else
        CM_StepLabel = CM_Step
    End If
End Function


'============================================================================
' PRINT + MERGE REHEARSAL
'----------------------------------------------------------------------------
' The close breaks most often at the printing step, and the old code could not
' tell the difference between "PDFCreator is slow" and "PDFCreator is never
' going to produce anything" - it simply waited for ever. Every other preflight
' check is a look; this one is a rehearsal, because the only way to know the
' chain works is to run it:
'
'     SAP  ->  front-end printer LOCLX  ->  Windows printer "PDFCreator"
'          ->  auto-saved PDF in C:\pdf\temp  ->  GiosPSMC.exe merge
'
' It writes only into C:\pdf\ and removes what it wrote. In SAP it displays
' SE16/T001 (company-code names) and prints that list: a display, so no
' business data is created or changed. Printing does raise a temporary SAP
' print job, which is why PreflightCheck asks first. If SAP is not available,
' or its print dialog does not appear, it falls back to printing a page from
' Excel and says so - that still proves PDFCreator and the merger.
'============================================================================
Public Function CM_PrintMergeTest(ByVal sess As Object, ByRef blocked As Boolean, _
                                  ByVal doPrint As Boolean) As String
    Dim fso As Object, out As String, route As String, why As String, src As String
    Dim tmp As String, work As String, exePath As String, merged As String
    Dim p As String, n As Long, i As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    Call EnsureFolders

    tmp     = CM_BASE_DRIVE & "pdf\temp\"
    work    = CM_BASE_DRIVE & "pdf\_preflight\"
    exePath = CM_BASE_DRIVE & "pdf\merger\GiosPSMC.exe"
    merged  = work & "merged.pdf"

    On Error Resume Next
    fso.DeleteFolder work, True
    Err.Clear
    On Error GoTo 0
    'guarded: if the delete above could not remove it (open in Explorer, say),
    'an unguarded CreateFolder would raise error 58 and kill the preflight
    EnsureFolderChain fso, work

    If Not doPrint Then GoTo DoMerge

    'the close picks up whatever is sitting in \temp, so leftovers are a fault
    'in their own right - and they would make the print test meaningless
    n = CM_CountFiles(fso, tmp)
    If n > 0 Then
        If MsgBox(n & " leftover file(s) are sitting in" & vbCrLf & tmp & vbCrLf & vbCrLf & _
                  "The close treats whatever is in that folder as its freshly printed" & vbCrLf & _
                  "report, so leftovers can end up inside the report pack." & vbCrLf & vbCrLf & _
                  "Delete them now?", vbYesNo + vbExclamation, _
                  "Closing Manager - Preflight") = vbYes Then
            On Error Resume Next
            fso.DeleteFile tmp & "*.*", True
            Err.Clear
            On Error GoTo 0
            out = out & "[OK] Cleared " & n & " leftover file(s) from \pdf\temp." & vbCrLf
        Else
            blocked = True
            out = out & "[X] " & n & " leftover file(s) in " & tmp & _
                        " - printing not tested." & vbCrLf
            GoTo DoMerge
        End If
    End If

    Call SetPDFCreator          'same launch the close does

    'two test pages, printed one at a time so each can be identified
    For i = 1 To 2
        CM_Step = "preflight: printing test page " & i & " of 2"
        CM_Paint ""

        why = ""
        route = "SAP"
        If sess Is Nothing Then
            route = "Excel"
        Else
            why = CM_SapTestPrint(sess)
            If why <> "" Then route = "Excel"
        End If
        If route = "Excel" Then why = CM_ExcelTestPrint()

        If why <> "" Then
            blocked = True
            out = out & "[X] Could not send test page " & i & " to PDFCreator - " & why & "." & vbCrLf
            GoTo DoMerge
        End If

        p = CM_WaitForFile(fso, tmp, 90, "test page " & i)
        If p = "" Then
            blocked = True
            out = out & "[X] Test page " & i & " was printed from " & route & _
                        " but no PDF appeared in" & vbCrLf & _
                        "       " & tmp & " within 90 seconds." & vbCrLf & _
                        "       PDFCreator is not saving automatically to that folder," & vbCrLf & _
                        "       or it is waiting on a dialog. This is what makes the" & vbCrLf & _
                        "       close appear to freeze." & vbCrLf
            GoTo DoMerge
        End If

        On Error Resume Next
        fso.MoveFile p, work & "t" & i & ".pdf"
        If Err.Number <> 0 Then
            Err.Clear: On Error GoTo 0
            blocked = True
            out = out & "[X] The test PDF was created but could not be moved out of \pdf\temp." & vbCrLf
            GoTo DoMerge
        End If
        On Error GoTo 0

        out = out & "[OK] Test page " & i & ": printed from " & route & _
                    " and saved as PDF (" & CM_KB(fso, work & "t" & i & ".pdf") & " KB)." & vbCrLf
    Next i

DoMerge:
    'The merger is tested whether or not the printing worked. If we did not get
    'two printed PDFs, two built-in one-page PDFs are written straight to disk,
    'so GiosPSMC.exe is proved on its own and a broken PDFCreator cannot hide a
    'broken merger (or the other way round).
    src = "the 2 printed test pages"
    If Not (fso.FileExists(work & "t1.pdf") And fso.FileExists(work & "t2.pdf")) Then
        On Error Resume Next
        fso.DeleteFile work & "*.*", True
        Err.Clear
        On Error GoTo 0
        If CM_SeedPdf(work & "t1.pdf", "PDF merge test page 1") And _
           CM_SeedPdf(work & "t2.pdf", "PDF merge test page 2") Then
            src = "2 built-in test pages"
        Else
            blocked = True
            out = out & "[X] Could not write the built-in test pages to " & work & _
                        " - merge not tested." & vbCrLf
            GoTo CleanUp
        End If
    End If

    If Not fso.FileExists(exePath) Then
        blocked = True
        out = out & "[X] PDF merger not present at " & exePath & " - merge not tested." & vbCrLf
        GoTo CleanUp
    End If

    CM_Step = "preflight: merging " & src
    CM_Paint ""
    On Error Resume Next
    Err.Clear
    CreateObject("WScript.Shell").Run "%COMSPEC% /c " & exePath & " " & _
        Chr$(34) & work & "t1.pdf" & Chr$(34) & " " & _
        Chr$(34) & work & "t2.pdf" & Chr$(34) & " output " & merged, 0, False
    If Err.Number <> 0 Then
        Err.Clear: On Error GoTo 0
        blocked = True
        out = out & "[X] The PDF merger would not start: " & exePath & vbCrLf
        GoTo CleanUp
    End If
    On Error GoTo 0

    If Not CM_WaitForPath(fso, merged, 60, "the merged test file") Then
        blocked = True
        out = out & "[X] The merger ran but produced no merged file within 60 seconds." & vbCrLf & _
                    "       " & exePath & vbCrLf & _
                    "       Check the file is a real GiosPSMC.exe and is not blocked by" & vbCrLf & _
                    "       Windows (Properties > Unblock) or by antivirus." & vbCrLf
    ElseIf CM_KB(fso, merged) < 0 Then
        blocked = True
        out = out & "[X] The merger produced an unreadable file." & vbCrLf
    ElseIf Not CM_LooksLikePdf(merged) Then
        blocked = True
        out = out & "[X] The merger produced a file that is not a PDF." & vbCrLf
    Else
        out = out & "[OK] PDF merge works: " & src & " -> 1 file (" & _
                    CM_KB(fso, merged) & " KB)." & vbCrLf
    End If

CleanUp:
    On Error Resume Next
    fso.DeleteFolder work, True
    If doPrint Then fso.DeleteFile tmp & "*.*", True
    Err.Clear
    On Error GoTo 0
    CM_Done
    CM_PrintMergeTest = out
End Function


'--- a real PDF starts with %PDF- and ends with %%EOF ------------------------
Private Function CM_LooksLikePdf(ByVal path As String) As Boolean
    Dim f As Integer, head As String * 5, sz As Long, tail As String * 6
    On Error Resume Next
    f = FreeFile
    Open path For Binary Access Read As #f
    sz = LOF(f)
    If sz > 32 Then
        Get #f, 1, head
        Get #f, sz - 5, tail
    End If
    Close #f
    If Err.Number = 0 Then
        CM_LooksLikePdf = (head = "%PDF-") And (InStr(tail, "%%EOF") > 0)
    End If
    Err.Clear
    On Error GoTo 0
End Function


'============================================================================
' A BUILT-IN ONE-PAGE PDF
'----------------------------------------------------------------------------
' So the merger can be tested on its own, with no printer involved. This writes
' a minimal but fully valid PDF 1.4 - catalog, page tree, one A4 page, one
' Helvetica text block - with a correct cross-reference table, because a merger
' will refuse a file whose xref offsets are wrong.
'
' Everything written is 7-bit ASCII and the line ending is LF, so the byte
' length of the string equals its character length and the offsets below are
' exact on any Windows regional setting.
'============================================================================
Private Function CM_SeedPdf(ByVal path As String, ByVal caption As String) As Boolean
    Dim NL As String, body As String, out As String
    Dim o(1 To 5) As String, offs(1 To 5) As Long
    Dim i As Long, xrefAt As Long, f As Integer
    Dim b() As Byte

    NL = Chr$(10)
    body = "BT" & NL & _
           "/F1 16 Tf" & NL & _
           "60 760 Td" & NL & _
           "(Closing Manager - preflight test) Tj" & NL & _
           "0 -28 Td" & NL & _
           "(" & caption & ") Tj" & NL & _
           "ET" & NL

    o(1) = "<< /Type /Catalog /Pages 2 0 R >>"
    o(2) = "<< /Type /Pages /Kids [3 0 R] /Count 1 >>"
    o(3) = "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] " & _
           "/Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>"
    o(4) = "<< /Length " & Len(body) & " >>" & NL & "stream" & NL & body & "endstream"
    o(5) = "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"

    out = "%PDF-1.4" & NL
    For i = 1 To 5
        offs(i) = Len(out)
        out = out & i & " 0 obj" & NL & o(i) & NL & "endobj" & NL
    Next i

    xrefAt = Len(out)
    out = out & "xref" & NL & "0 6" & NL & "0000000000 65535 f " & NL
    For i = 1 To 5
        out = out & Format$(offs(i), "0000000000") & " 00000 n " & NL
    Next i
    out = out & "trailer" & NL & _
                "<< /Size 6 /Root 1 0 R >>" & NL & _
                "startxref" & NL & xrefAt & NL & "%%EOF" & NL

    On Error Resume Next
    If Dir(path) <> "" Then Kill path
    Err.Clear
    b = StrConv(out, vbFromUnicode)
    f = FreeFile
    Open path For Binary Access Write As #f
    Put #f, 1, b
    Close #f
    If Err.Number = 0 Then CM_SeedPdf = True
    Err.Clear
    On Error GoTo 0
End Function


'--- print SE16/T001 - a display, so nothing is created or changed -----------
Private Function CM_SapTestPrint(ByVal sess As Object) As String
    On Error Resume Next
    Err.Clear
    CM_ClearPopups sess

    sess.findById("wnd[0]/tbar[0]/okcd").Text = "/nse16"
    sess.findById("wnd[0]").sendVKey 0
    sess.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").Text = "T001"
    sess.findById("wnd[0]").sendVKey 0
    If Err.Number <> 0 Then
        Err.Clear: On Error GoTo 0
        CM_SapTestPrint = "SE16 did not open"
        Exit Function
    End If

    'keep the printed list to a single short page
    sess.findById("wnd[0]/usr/txtMAX_SEL").Text = "3"
    Err.Clear

    sess.findById("wnd[0]/tbar[1]/btn[8]").press          'F8 - display
    If Err.Number <> 0 Then
        Err.Clear: On Error GoTo 0
        CM_SapTestPrint = "SE16 would not display table T001"
        Exit Function
    End If

    sess.findById("wnd[0]").sendVKey 86                   'Ctrl+P - print
    sess.findById("wnd[1]/usr/ctxtPRI_PARAMS-PDEST").Text = "LOCLX"
    sess.findById("wnd[1]").sendVKey 0
    If Err.Number <> 0 Then
        Err.Clear
        CM_SapBackOut sess
        On Error GoTo 0
        CM_SapTestPrint = "the SAP print dialog did not appear"
        Exit Function
    End If

    sess.findById("wnd[1]/usr/cmbPRIPAR_EXT-OSPRINTER").Key = "PDFCreator"
    If Err.Number <> 0 Then
        Err.Clear
        'the printer list can sit one window deeper, exactly as in Print_ZGLRME
        sess.findById("wnd[2]/tbar[0]/btn[0]").press
        sess.findById("wnd[1]/usr/cmbPRIPAR_EXT-OSPRINTER").Key = "PDFCreator"
        If Err.Number <> 0 Then
            Err.Clear
            CM_SapBackOut sess
            On Error GoTo 0
            CM_SapTestPrint = "SAP does not offer a front-end printer called 'PDFCreator'"
            Exit Function
        End If
    End If

    sess.findById("wnd[1]/tbar[0]/btn[13]").press         'print
    If Err.Number <> 0 Then
        Err.Clear
        CM_SapBackOut sess
        On Error GoTo 0
        CM_SapTestPrint = "SAP would not accept the print"
        Exit Function
    End If

    CM_SapBackOut sess
    Err.Clear
    On Error GoTo 0
End Function


'--- leave SAP on a clean screen ---------------------------------------------
Private Sub CM_SapBackOut(ByVal sess As Object)
    On Error Resume Next
    sess.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    sess.findById("wnd[0]").sendVKey 0
    Err.Clear
    On Error GoTo 0
End Sub


'--- fallback: print one page from Excel to the same printer -----------------
Private Function CM_ExcelTestPrint() As String
    Dim prn As String, ws As Object
    prn = CM_PdfCreatorPrinter()
    If prn = "" Then
        CM_ExcelTestPrint = "no Windows printer called 'PDFCreator'"
        Exit Function
    End If

    On Error Resume Next
    Err.Clear
    Set ws = ThisWorkbook.Sheets("Preflight")
    If ws Is Nothing Then Set ws = ThisWorkbook.Sheets(1)
    Err.Clear
    ws.PrintOut Copies:=1, ActivePrinter:=prn, Collate:=True
    If Err.Number <> 0 Then CM_ExcelTestPrint = "Excel could not print to " & prn
    Err.Clear
    On Error GoTo 0
End Function


'--- "PDFCreator on Ne00:" - the name Excel needs, "" if there is none -------
Public Function CM_PdfCreatorPrinter() As String
    Dim wmi As Object, col As Object, p As Object
    On Error Resume Next
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    If wmi Is Nothing Then
        Err.Clear: On Error GoTo 0
        Exit Function
    End If
    Set col = wmi.ExecQuery("SELECT Name, PortName FROM Win32_Printer")
    For Each p In col
        If InStr(1, CStr(p.Name), "PDFCreator", vbTextCompare) > 0 Then
            CM_PdfCreatorPrinter = CStr(p.Name) & " on " & CStr(p.PortName)
            Exit For
        End If
    Next
    Err.Clear
    On Error GoTo 0
End Function


'--- how many files are sitting in a folder ----------------------------------
Private Function CM_CountFiles(ByVal fso As Object, ByVal folder As String) As Long
    Dim f As Object
    On Error Resume Next
    If Not fso.FolderExists(folder) Then Exit Function
    For Each f In fso.GetFolder(folder).Files
        CM_CountFiles = CM_CountFiles + 1
    Next
    Err.Clear
    On Error GoTo 0
End Function


'--- file size in whole KB, -1 if it is not there ----------------------------
Private Function CM_KB(ByVal fso As Object, ByVal p As String) As Long
    On Error Resume Next
    CM_KB = -1
    CM_KB = CLng(fso.GetFile(p).Size \ 1024)
    Err.Clear
    On Error GoTo 0
End Function


'--- bounded wait for one named file to appear and stop growing --------------
Private Function CM_WaitForPath(ByVal fso As Object, ByVal p As String, _
                                ByVal maxSecs As Long, ByVal what As String) As Boolean
    Dim waited As Long, s1 As Double, s2 As Double
    Do
        If fso.FileExists(p) Then Exit Do
        If waited >= maxSecs Then Exit Function
        CM_Tick "waiting for " & what & " (" & waited & "/" & maxSecs & "s)"
        waited = waited + 1
    Loop
    Do
        s1 = -1: s2 = -1
        On Error Resume Next
        s1 = fso.GetFile(p).Size
        Err.Clear
        On Error GoTo 0
        CM_Tick "writing " & what & " (" & waited & "s)"
        On Error Resume Next
        s2 = fso.GetFile(p).Size
        Err.Clear
        On Error GoTo 0
        If s1 >= 0 And s1 = s2 And s1 > 0 Then Exit Do
        waited = waited + 1
        If waited >= maxSecs Then Exit Do
    Loop
    CM_WaitForPath = (s1 > 0)
End Function


'--- bounded wait used by the rehearsal; "" if nothing arrived ---------------
Private Function CM_WaitForFile(ByVal fso As Object, ByVal folder As String, _
                                ByVal maxSecs As Long, ByVal what As String) As String
    Dim base As String, nm As String, waited As Long
    Dim s1 As Double, s2 As Double

    base = folder
    If Right$(base, 1) <> "\" Then base = base & "\"

    Do
        nm = Dir(base & "*.*")
        If nm <> "" Then Exit Do
        If waited >= maxSecs Then Exit Function
        CM_Tick "waiting for " & what & " (" & waited & "/" & maxSecs & "s)"
        waited = waited + 1
    Loop

    'let it finish being written
    Do
        s1 = -1: s2 = -1
        On Error Resume Next
        s1 = fso.GetFile(base & nm).Size
        On Error GoTo 0
        CM_Tick "writing " & what & " (" & waited & "s)"
        On Error Resume Next
        s2 = fso.GetFile(base & nm).Size
        Err.Clear
        On Error GoTo 0
        If s1 >= 0 And s1 = s2 And s1 > 0 Then Exit Do
        nm = Dir(base & "*.*")
        If nm = "" Then Exit Function
        waited = waited + 1
        If waited >= maxSecs Then Exit Do
    Loop

    CM_WaitForFile = base & nm
End Function
