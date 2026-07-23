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

Do
    Data = strix.ReadText(-2)
    If VBA.Left(VBA.Trim(Data), 1) = "|" Then
        For i = 2 To Len(Trim(Data))
            If Mid(Trim(Data), i, 1) = "|" Then
                FirstColumn = Trim(Mid(Trim(Data), 2, i - 2))
                Exit For
            End If
        Next i
        Exit Do
    End If
Loop

strix.Close
Set strix = Nothing

End Sub
Sub ProperArray()

Dim a, b, n
Dim nStart
Dim nEnd

a = 1
n = 0
nStart = 1
nEnd = 1
Do
    If nStart >= VBA.Len(Data) Then Exit Do
    For a = nStart To VBA.Len(Data)
        If VBA.Mid(Data, a, 1) = "|" Then
            nStart = a + 1
            nEnd = nStart
            Exit For
        End If
    Next
        
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
            If eng.Children.Count > 0 Then sapOK = True
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
        msg = msg & "[X] config!B2 (Cost Centre) is empty." & vbCrLf
        okAll = False
    Else
        msg = msg & "[OK] Cost Centre = " & cc & "." & vbCrLf
    End If

    msg = "CLOSING MANAGER  -  PREFLIGHT CHECK" & vbCrLf & _
          "--------------------------------------" & vbCrLf & _
          msg & _
          "--------------------------------------" & vbCrLf & _
          IIf(okAll, "READY TO RUN.", "NOT READY - resolve the [X] items above.")
    MsgBox msg, IIf(okAll, vbInformation, vbExclamation), "Closing Manager - Preflight"
End Sub
