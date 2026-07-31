Attribute VB_Name = "modUtil"
'=======================================================================
' modUtil -- date, amount, string and filesystem helpers.
'
' The date and amount formatting here is the part most likely to bite.
' SAP renders dates and numbers in the logged-on user's own format, and
' the sample workbook was written by someone whose Excel uses
' dd/mm/yyyy with '.' thousands and ',' decimals. Both are handled
' explicitly rather than left to CStr.
'=======================================================================
Option Explicit

'-----------------------------------------------------------------------
' How the extract is laid out on disk.
'
'   Audit GBKM\
'     Sep 25\
'       01 - 8072447.42\
'         01 - 8072447.42 - GBKM.xlsx        <- the report, read this first
'         1 - FEBAN statement list.xlsx
'         2 - Payment usage - batch.xlsx
'         3 - FBL1N - payments in the batch.xlsx
'         4 - Invoices behind the largest payment.xlsx
'         5 - Supplier invoices behind the largest SCF settlement.xlsx
'         6 - Largest invoice.pdf
'
' One folder per sample, and fixed names inside it. The sample used to be
' encoded in every file name because everything sat in one folder; now the
' folder says which sample it is and the file name says which step of the
' chain it came from, which is what an auditor needs to read a pack they
' did not watch being made.
'
' The leading number keeps them in chain order in Explorer.
'-----------------------------------------------------------------------
Public Const FILE_FEBAN As String = "1 - FEBAN statement list.xlsx"
Public Const FILE_BATCH As String = "2 - Payment usage - batch of payments.xlsx"
Public Const FILE_FIDOC As String = "2 - FI document line items (not cleared).xlsx"
Public Const FILE_ZPLIST As String = "3 - FBL1N - payments in the batch.xlsx"
Public Const FILE_INVOICES As String = "4 - Documents behind the largest payment.xlsx"
Public Const FILE_SCF As String = "5 - Supplier invoices behind the largest SCF settlement.xlsx"
Public Const FILE_PDF As String = "6 - Largest invoice.pdf"

#If VBA7 Then
    Private Declare PtrSafe Sub SleepApi Lib "kernel32" Alias "Sleep" (ByVal ms As Long)
#Else
    Private Declare Sub SleepApi Lib "kernel32" Alias "Sleep" (ByVal ms As Long)
#End If

Public Sub SleepSeconds(ByVal seconds As Double)
    SleepApi CLng(seconds * 1000)
End Sub

'-----------------------------------------------------------------------
' Dates
'-----------------------------------------------------------------------
' SAP input fields accept the format of the logged-on user, so read that
' from the session rather than assuming. Falls back to dd.mm.yyyy, the
' most common setting on European systems.
Public Function SapDate(ByVal value As Date) As String
    Dim pattern As String

    pattern = SapDatePattern()

    Select Case pattern
        ' What recordings/Audit.vbs actually typed: 8 digits, no separators.
        Case "DDMMYYYY": SapDate = Format$(value, "ddmmyyyy")   ' 01092025
        Case "MMDDYYYY": SapDate = Format$(value, "mmddyyyy")   ' 09012025
        Case "DMY":      SapDate = Format$(value, "dd.mm.yyyy") ' 31.12.2025
        Case "DMY/":     SapDate = Format$(value, "dd/mm/yyyy") ' 31/12/2025
        Case "MDY":      SapDate = Format$(value, "mm/dd/yyyy") ' 12/31/2025
        Case "YMD":      SapDate = Format$(value, "yyyy-mm-dd") ' 2025-12-31
        Case Else
            Err.Raise vbObjectError + 560, "modUtil.SapDate", _
                      "'SAP date format' on the Control sheet reads '" & pattern & _
                      "'. It must be one of DDMMYYYY, MMDDYYYY, DMY, DMY/, MDY or YMD."
    End Select
End Function

' SAP renders dates in grid cells with separators even when it accepts them
' typed without. So a grid value is compared on its digits alone, which makes
' the comparison independent of both the input format and the display format.
Public Function DateDigits(ByVal value As Date) As String
    DateDigits = Format$(value, "yyyymmdd")
End Function

' Pull a date out of whatever a grid cell holds -- '03.09.2025', '03/09/2025',
' '2025-09-03' or '03092025' -- and return it as yyyymmdd, or "" if unreadable.
Public Function GridDateDigits(ByVal text As String, ByVal pattern As String) As String
    Dim digits As String
    Dim i As Long
    Dim ch As String
    Dim dd As String, mm As String, yyyy As String

    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch >= "0" And ch <= "9" Then digits = digits & ch
    Next i

    If Len(digits) <> 8 Then Exit Function

    Select Case UCase$(pattern)
        Case "YMD"
            GridDateDigits = digits
            Exit Function
        Case "MDY", "MMDDYYYY"
            mm = Left$(digits, 2): dd = Mid$(digits, 3, 2): yyyy = Right$(digits, 4)
        Case Else                       ' every DMY variant
            dd = Left$(digits, 2): mm = Mid$(digits, 3, 2): yyyy = Right$(digits, 4)
    End Select

    GridDateDigits = yyyy & mm & dd
End Function

' Probe the session's own date format. GuiSession does not expose it
' directly on every release, so this reads it from a known date field
' where possible and otherwise returns the default.
Private Function SapDatePattern() As String
    Dim override As String

    ' An operator who knows their system's setting can pin it on the
    ' Control sheet rather than relying on the probe.
    On Error Resume Next
    override = UCase$(Trim$(modConfig.Setting("SAP date format")))
    On Error GoTo 0

    If Len(override) > 0 Then
        SapDatePattern = override
    Else
        SapDatePattern = "DDMMYYYY"      ' what the recording used
    End If
End Function

Public Function SapDatePatternPublic() As String
    SapDatePatternPublic = SapDatePattern()
End Function

Public Function MonthStart(ByVal value As Date) As Date
    MonthStart = DateSerial(Year(value), Month(value), 1)
End Function

Public Function MonthEnd(ByVal value As Date) As Date
    MonthEnd = DateSerial(Year(value), Month(value) + 1, 0)
End Function

' 'Sep 25' -- matches the auditor's per-month tab names.
Public Function MonthTabName(ByVal value As Date) As String
    MonthTabName = Format$(value, "mmm yy")
End Function

'-----------------------------------------------------------------------
' Amounts
'-----------------------------------------------------------------------
' Parse an amount as SAP rendered it. Handles '8.072.447,42' and
' '8,072,447.42', and a trailing minus, which SAP uses for credits.
Public Function ParseSapAmount(ByVal text As String) As Double
    Dim cleaned As String
    Dim isNegative As Boolean
    Dim lastDot As Long, lastComma As Long

    cleaned = Trim$(text)
    If Len(cleaned) = 0 Then Exit Function

    If Right$(cleaned, 1) = "-" Then
        isNegative = True
        cleaned = Trim$(Left$(cleaned, Len(cleaned) - 1))
    ElseIf Left$(cleaned, 1) = "-" Then
        isNegative = True
        cleaned = Trim$(Mid$(cleaned, 2))
    End If

    cleaned = Replace(cleaned, " ", vbNullString)
    cleaned = Replace(cleaned, Chr$(160), vbNullString)   ' non-breaking space

    ' Whichever separator appears last is the decimal separator.
    lastDot = InStrRev(cleaned, ".")
    lastComma = InStrRev(cleaned, ",")

    If lastComma > lastDot Then
        cleaned = Replace(cleaned, ".", vbNullString)
        cleaned = Replace(cleaned, ",", ".")
    Else
        cleaned = Replace(cleaned, ",", vbNullString)
    End If

    If Not IsNumeric(cleaned) Then Exit Function

    ParseSapAmount = CDbl(Val(cleaned))                   ' Val is locale-independent
    If isNegative Then ParseSapAmount = -ParseSapAmount
End Function

' The sample list holds unsigned amounts even where the statement shows a
' debit, so compare on magnitude.
Public Function AmountsMatch(ByVal a As Double, ByVal b As Double, _
                             ByVal tolerance As Double) As Boolean
    AmountsMatch = (Abs(Abs(a) - Abs(b)) <= tolerance)
End Function

'-----------------------------------------------------------------------
' Strings and filenames
'-----------------------------------------------------------------------
Public Function Squeeze(ByVal text As String) As String
    Dim result As String
    result = Trim$(Replace(Replace(text, vbTab, " "), vbCrLf, " "))
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    Squeeze = result
End Function

Public Function SafeFileName(ByVal text As String) As String
    Dim illegal As Variant
    Dim item As Variant
    Dim result As String

    illegal = Array("\", "/", ":", "*", "?", """", "<", ">", "|", vbTab, vbCr, vbLf)
    result = Squeeze(text)

    For Each item In illegal
        result = Replace(result, CStr(item), "_")
    Next item

    ' Windows will not accept a name ending in a dot or space.
    Do While Len(result) > 0 And (Right$(result, 1) = "." Or Right$(result, 1) = " ")
        result = Left$(result, Len(result) - 1)
    Loop

    If Len(result) > 120 Then result = Left$(result, 120)
    If Len(result) = 0 Then result = "unnamed"

    SafeFileName = result
End Function

' Create a folder and any missing parents. Returns the path.
Public Function EnsureFolder(ByVal path As String) As String
    Dim fso As Object
    Dim parent As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(path) Then
        parent = fso.GetParentFolderName(path)
        If Len(parent) > 0 And Not fso.FolderExists(parent) Then EnsureFolder parent
        fso.CreateFolder path
    End If

    EnsureFolder = path
End Function

Public Function JoinPath(ByVal folder As String, ByVal leaf As String) As String
    If Right$(folder, 1) = "\" Then
        JoinPath = folder & leaf
    Else
        JoinPath = folder & "\" & leaf
    End If
End Function

Public Function FileExists(ByVal path As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(path)
End Function

Public Function FileSizeBytes(ByVal path As String) As Double
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(path) Then FileSizeBytes = fso.GetFile(path).Size
End Function

'-----------------------------------------------------------------------
' SAP's spreadsheet export writes the file and then opens it in Excel. That
' is how the operator ended a run with the export sitting on top of the
' workbook that produced it, and a ~$ lock file left in the evidence folder.
'
' Nothing here needs the file open, so close it again. Matching on the full
' path means only the file just written is touched -- never the control
' workbook, and never anything the operator opened themselves.
'-----------------------------------------------------------------------
'-----------------------------------------------------------------------
' Close any open workbook with this file NAME, wherever it came from.
'
' Excel refuses to hold two workbooks with the same name at once, whatever
' folder each is in. That never mattered while every export carried the
' sample number in its name; it started mattering the moment each sample
' got its own folder and the names inside became fixed. Sample 1's
' '2 - Payment usage - batch of payments.xlsx' is auto-opened by SAP, and
' when sample 2's file of the same name arrives Excel puts up a modal and
' waits for a human -- and refuses the open, so the macro's own read of
' that file fails too.
'
' Matching on name rather than path is the point: it is the name Excel
' objects to. Returns how many were closed.
'-----------------------------------------------------------------------
Public Function CloseWorkbooksNamed(ByVal fileName As String) As Long
    Dim book As Workbook
    Dim doomed As Collection
    Dim item As Variant
    Dim previousAlerts As Boolean

    If Len(fileName) = 0 Then Exit Function

    Set doomed = New Collection

    On Error Resume Next
    For Each book In Application.Workbooks
        If Not book Is ThisWorkbook Then
            If StrComp(book.Name, fileName, vbTextCompare) = 0 Then doomed.Add book
        End If
    Next book
    On Error GoTo 0

    previousAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    On Error Resume Next
    For Each item In doomed
        item.Close SaveChanges:=False
    Next item
    On Error GoTo 0

    Application.DisplayAlerts = previousAlerts

    CloseWorkbooksNamed = doomed.Count
End Function

Public Sub CloseWorkbookIfOpen(ByVal path As String)
    Dim book As Workbook
    Dim previousAlerts As Boolean

    If Len(path) = 0 Then Exit Sub

    previousAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    On Error Resume Next
    For Each book In Application.Workbooks
        If StrComp(book.FullName, path, vbTextCompare) = 0 Then
            If Not book Is ThisWorkbook Then book.Close SaveChanges:=False
            Exit For
        End If
    Next book
    On Error GoTo 0

    Application.DisplayAlerts = previousAlerts
End Sub

' Same thing, for every export under the evidence folder at once. SAP opens
' the file asynchronously, so it can land after the macro has already read
' and closed it; sweeping between samples and once at the end catches those.
' Returns how many were closed, so the caller can say so.
Public Function CloseExportWorkbooksUnder(ByVal root As String) As Long
    Dim book As Workbook
    Dim doomed As Collection
    Dim item As Variant
    Dim previousAlerts As Boolean

    If Len(root) = 0 Then Exit Function

    Set doomed = New Collection

    ' Collect first, close second: closing while iterating Workbooks skips
    ' entries, because the collection is renumbered underneath the loop.
    For Each book In Application.Workbooks
        If Not book Is ThisWorkbook Then
            If InStr(1, book.FullName, root, vbTextCompare) = 1 Then
                doomed.Add book
            End If
        End If
    Next book

    previousAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    On Error Resume Next
    For Each item In doomed
        item.Close SaveChanges:=False
    Next item
    On Error GoTo 0

    Application.DisplayAlerts = previousAlerts

    CloseExportWorkbooksUnder = doomed.Count
End Function

' 'Paper Samples - GBKM'. One folder per audit request, with the company
' code in the name, because several requests are open at once and they are
' not all the same company. A row that predates the import -- no request
' name -- goes straight under the root as it always did, so nothing that
' already ran moves.
Public Function RequestFolderName(ByVal request As String, _
                                  ByVal companyCode As String) As String
    Dim name As String

    name = Trim$(request)
    If Len(name) = 0 Then Exit Function

    If Len(Trim$(companyCode)) > 0 Then name = name & " - " & Trim$(companyCode)

    RequestFolderName = SafeFileName(name)
End Function

' '01 - 8072447.42'. The index leads so the folders sort in sample order;
' the amount is what the auditor's request identifies the line by. A plain
' decimal point, not the local separator, because this is a folder name and
' has to be the same on every machine that opens the pack.
Public Function SampleFolderName(ByVal sampleIdx As Long, ByVal amount As Double) As String
    SampleFolderName = SafeFileName(Format$(sampleIdx, "00") & " - " & AmountForName(amount))
End Function

'-----------------------------------------------------------------------
' Empty a sample's folder of the previous run's evidence.
'
' The exports never overwrite: a second run used to leave
' '..._ZP_batch_list_2.xlsx' beside the first, and with one folder per
' sample that turns the pack an auditor receives into a pile of near
' duplicates. So a re-run starts the folder clean.
'
' It only ever deletes names THIS macro writes -- the numbered evidence
' files and their _2 variants -- inside a folder this macro created. Any
' other file in there, including anything the operator put there, is left
' alone. Returns how many were removed.
'-----------------------------------------------------------------------
Public Function ClearSampleFolder(ByVal folder As String) As Long
    Dim fso As Object
    Dim file As Object
    Dim doomed As Collection
    Dim item As Variant
    Dim name As String

    If Len(folder) = 0 Then Exit Function

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folder) Then Exit Function

    Set doomed = New Collection

    On Error Resume Next
    For Each file In fso.GetFolder(folder).Files
        name = file.Name
        If IsEvidenceName(name) Then doomed.Add file.path
    Next file

    For Each item In doomed
        fso.DeleteFile CStr(item), True
    Next item
    On Error GoTo 0

    ClearSampleFolder = doomed.Count
End Function

' '1 - ...' through '5 - ...' are the numbered evidence names; anything
' matching '<digit> - ' is ours. The report itself is overwritten in place
' by SaveAs, so it does not need deleting.
Private Function IsEvidenceName(ByVal name As String) As Boolean
    If Len(name) < 4 Then Exit Function
    If Mid$(name, 2, 3) <> " - " Then Exit Function

    IsEvidenceName = (Left$(name, 1) >= "1" And Left$(name, 1) <= "6")
End Function

' '01 - 8072447.42 - GBKM.xlsx'
Public Function ReportFileName(ByVal sampleIdx As Long, ByVal amount As Double, _
                               ByVal companyCode As String) As String
    ReportFileName = SafeFileName(Format$(sampleIdx, "00") & " - " & AmountForName(amount) & _
                                  " - " & companyCode) & ".xlsx"
End Function

Private Function AmountForName(ByVal amount As Double) As String
    Dim text As String

    ' Format$ follows the machine's locale, so a Spanish or German Excel
    ' would write 8072447,42 and the same sample would get two different
    ' folder names on two laptops. Normalise to a point.
    text = Format$(Abs(amount), "0.00")
    AmountForName = Replace(text, ",", ".")
End Function

' Stem for every file belonging to one sample, so the extract sorts in
' sample order and each file names the evidence it is.
Public Function EvidenceStem(ByVal sampleIdx As Long, ByVal paymentDate As Date, _
                             ByVal amount As Double) As String
    EvidenceStem = Format$(sampleIdx, "00") & "_" & _
                   Format$(paymentDate, "yyyy-mm-dd") & "_" & _
                   Format$(Abs(amount), "0")
End Function
