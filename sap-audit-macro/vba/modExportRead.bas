Attribute VB_Name = "modExportRead"
'=======================================================================
' modExportRead -- reads a SAP export back off disk.
'
' The recordings name every export .XLSX (FEBAN1.XLSX, ZPList.XLSX,
' Zppayments.XLSX), so the earlier assumption that these are delimited
' text was wrong. Two shapes turn up in practice:
'
'   ALV  "&XXL"                      a real Excel workbook
'   classic  List > Save/Send > File whatever the format popup's default
'            radio was -- often plain text, regardless of the .XLSX name
'
' So the format is sniffed from the file's first bytes rather than its
' extension: a ZIP signature ("PK") means a real workbook and is opened
' with Workbooks.Open; anything else is parsed as delimited text.
'
' The loaded table is held in module-level state, so only one export is
' open at a time. Every caller loads, queries, and moves on -- which is
' how the chain uses it. Loading a second export discards the first.
'=======================================================================
Option Explicit

Public Type ListRow
    Found As Boolean
    Supplier As String
    Amount As Double
    DocumentNumber As String
    RowsConsidered As Long
    SourceRow As Long
End Type

Private mCells() As String        ' 1-based (row, column)
Private mRowCount As Long
Private mColCount As Long
Private mHeaderRow As Long        ' 0 when no header was recognised
Private mPath As String
Private mWasWorkbook As Boolean

' SAP writes these lists in the logged-on language, so the captions to look
' for are settings rather than constants. These are the fallbacks.
Private Const DEFAULT_AMOUNT_CAPTIONS As String = _
    "Amount in local currency|Amount in LC|Amount|LC amount|Amnt in loc.cur.|" & _
    "Amount in doc. curr.|DC amount|DMBTR|WRBTR|KWBTR"
Private Const DEFAULT_SUPPLIER_CAPTIONS As String = _
    "Name|Name 1|Name of vendor|Vendor name|Supplier|Account name|NAME1|Text|Assignment"
Private Const DEFAULT_DOCUMENT_CAPTIONS As String = _
    "Document Number|DocumentNo|Document no.|Doc. Number|Doc.no.|BELNR|Invoice reference"
Private Const DEFAULT_DOCTYPE_CAPTIONS As String = _
    "Document Type|Doc. Type|DocumentType|Type|BLART"

'-----------------------------------------------------------------------
' Load
'-----------------------------------------------------------------------
Public Function LoadExport(ByVal path As String, ByVal sampleIdx As Long) As Boolean
    Reset

    If Not modUtil.FileExists(path) Then
        modLog.LogAction sampleIdx, "Read export", _
                     "File not found: " & path, "ERROR", path
        Exit Function
    End If

    mPath = path

    If LooksLikeZip(path) Then
        mWasWorkbook = True
        LoadExport = LoadFromWorkbook(path, sampleIdx)
    Else
        mWasWorkbook = False
        LoadExport = LoadFromText(path, sampleIdx)
    End If

    If LoadExport Then
        modLog.LogAction sampleIdx, "Read export", _
                     "Loaded " & mRowCount & " rows x " & mColCount & " columns from " & _
                     IIf(mWasWorkbook, "the workbook ", "the text file ") & path, _
                     "OK", path
    End If
End Function

Private Sub Reset()
    Erase mCells
    mRowCount = 0
    mColCount = 0
    mHeaderRow = 0
    mPath = vbNullString
    mWasWorkbook = False
End Sub

' A real .xlsx is a ZIP, so it starts with the bytes "PK". Sniffing beats
' trusting the extension: SAP happily writes plain text into a .XLSX name,
' and handing that to Workbooks.Open pops a modal warning mid-run.
Private Function LooksLikeZip(ByVal path As String) As Boolean
    Dim handle As Integer
    Dim signature As String * 2

    On Error GoTo Done

    handle = FreeFile
    Open path For Binary Access Read As #handle
    If LOF(handle) >= 2 Then Get #handle, 1, signature
    Close #handle

    LooksLikeZip = (signature = "PK")
    Exit Function

Done:
    On Error Resume Next
    Close #handle
    On Error GoTo 0
End Function

Private Function LoadFromWorkbook(ByVal path As String, ByVal sampleIdx As Long) As Boolean
    Dim book As Workbook
    Dim sheet As Worksheet
    Dim area As Range
    Dim r As Long, c As Long
    Dim previousAlerts As Boolean

    previousAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    On Error GoTo Failed

    ' UpdateLinks:=0 stops a linked export prompting; ReadOnly leaves the
    ' operator's file untouched.
    Set book = Application.Workbooks.Open(fileName:=path, UpdateLinks:=0, ReadOnly:=True)
    Set sheet = book.Worksheets(1)
    Set area = sheet.UsedRange

    mRowCount = area.Rows.Count
    mColCount = area.Columns.Count

    If mRowCount = 0 Or mColCount = 0 Then
        book.Close SaveChanges:=False
        Application.DisplayAlerts = previousAlerts
        modLog.LogAction sampleIdx, "Read export", _
                     "The workbook " & path & " has no used range -- the export is empty.", _
                     "ERROR", path
        Exit Function
    End If

    ReDim mCells(1 To mRowCount, 1 To mColCount)

    For r = 1 To mRowCount
        For c = 1 To mColCount
            mCells(r, c) = modUtil.Squeeze(CStr(area.Cells(r, c).Text))
        Next c
    Next r

    book.Close SaveChanges:=False
    Set book = Nothing
    Application.DisplayAlerts = previousAlerts

    LoadFromWorkbook = True
    Exit Function

Failed:
    On Error Resume Next
    If Not book Is Nothing Then book.Close SaveChanges:=False
    Application.DisplayAlerts = previousAlerts
    On Error GoTo 0

    modLog.LogAction sampleIdx, "Read export", _
                 "Could not open " & path & " as a workbook: " & Err.Description, _
                 "ERROR", path
End Function

Private Function LoadFromText(ByVal path As String, ByVal sampleIdx As Long) As Boolean
    Dim lines() As String
    Dim delimiter As String
    Dim keep() As String
    Dim kept As Long
    Dim i As Long, c As Long
    Dim fields() As String
    Dim widest As Long

    lines = ReadLines(path)
    If UBound(lines) < LBound(lines) Then Exit Function

    delimiter = DetectDelimiter(lines)

    ' Drop rules, blanks and page furniture first, so row indices below line
    ' up with what a reader would call row 1, row 2 ...
    ReDim keep(LBound(lines) To UBound(lines))
    For i = LBound(lines) To UBound(lines)
        If IsDataLine(lines(i)) Then
            keep(kept) = lines(i)
            kept = kept + 1
            fields = Split(lines(i), delimiter)
            If UBound(fields) + 1 > widest Then widest = UBound(fields) + 1
        End If
    Next i

    If kept = 0 Or widest = 0 Then
        modLog.LogAction sampleIdx, "Read export", _
                     "No data lines found in " & path & ". Delimiter guessed as " & _
                     IIf(delimiter = vbTab, "TAB", delimiter) & ".", "ERROR", path
        Exit Function
    End If

    mRowCount = kept
    mColCount = widest
    ReDim mCells(1 To mRowCount, 1 To mColCount)

    For i = 0 To kept - 1
        fields = Split(keep(i), delimiter)
        For c = 0 To UBound(fields)
            If c < mColCount Then mCells(i + 1, c + 1) = modUtil.Squeeze(fields(c))
        Next c
    Next i

    LoadFromText = True
End Function

Private Function ReadLines(ByVal path As String) As String()
    Dim stream As Object
    Dim text As String
    Dim handle As Integer

    On Error GoTo Fallback
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile path
    text = stream.ReadText
    stream.Close

    ReadLines = SplitLines(text)
    Exit Function

Fallback:
    On Error GoTo 0
    handle = FreeFile
    Open path For Input As #handle
    If LOF(handle) > 0 Then text = Input$(LOF(handle), handle)
    Close #handle
    ReadLines = SplitLines(text)
End Function

Private Function SplitLines(ByVal text As String) As String()
    SplitLines = Split(Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf), vbLf)
End Function

' SAP writes '|' for the unconverted format and tabs for text-with-tabs.
Private Function DetectDelimiter(ByRef lines() As String) As String
    Dim i As Long
    Dim pipes As Long, tabs As Long

    For i = LBound(lines) To UBound(lines)
        pipes = pipes + CountChar(lines(i), "|")
        tabs = tabs + CountChar(lines(i), vbTab)
        If i > 60 Then Exit For
    Next i

    DetectDelimiter = IIf(pipes >= tabs And pipes > 0, "|", vbTab)
End Function

Private Function CountChar(ByVal text As String, ByVal ch As String) As Long
    CountChar = Len(text) - Len(Replace(text, ch, vbNullString))
End Function

' A data line, as opposed to a rule, a page header or a blank.
Private Function IsDataLine(ByVal text As String) As Boolean
    Dim stripped As String

    stripped = Replace(Replace(Replace(text, "-", ""), "|", ""), " ", "")
    stripped = Replace(Replace(Replace(stripped, vbTab, ""), "_", ""), "=", "")

    If Len(stripped) = 0 Then Exit Function          ' blank or a rule
    If Len(Trim$(text)) < 3 Then Exit Function

    IsDataLine = True
End Function

'-----------------------------------------------------------------------
' Query
'-----------------------------------------------------------------------
Public Property Get LoadedRowCount() As Long
    LoadedRowCount = mRowCount
End Property

Public Property Get LoadedPath() As String
    LoadedPath = mPath
End Property

Public Function CellValue(ByVal row As Long, ByVal column As Long) As String
    If row < 1 Or row > mRowCount Then Exit Function
    If column < 1 Or column > mColCount Then Exit Function
    CellValue = mCells(row, column)
End Function

' Locate the header row and the columns named by the given settings. Returns
' True and fills the column indices, or False when nothing matched.
Public Function FindColumns(ByVal sampleIdx As Long, _
                            ByVal amountSetting As String, _
                            ByVal supplierSetting As String, _
                            ByVal documentSetting As String, _
                            ByRef amountCol As Long, ByRef supplierCol As Long, _
                            ByRef documentCol As Long) As Boolean
    Dim amountCaptions As String, supplierCaptions As String, documentCaptions As String
    Dim r As Long

    amountCaptions = CaptionSetting(amountSetting, DEFAULT_AMOUNT_CAPTIONS)
    supplierCaptions = CaptionSetting(supplierSetting, DEFAULT_SUPPLIER_CAPTIONS)
    documentCaptions = CaptionSetting(documentSetting, DEFAULT_DOCUMENT_CAPTIONS)

    ' Headers sit near the top of these lists.
    For r = 1 To LesserOf(mRowCount, 80)
        amountCol = MatchColumn(r, amountCaptions)
        supplierCol = MatchColumn(r, supplierCaptions)

        If amountCol > 0 And supplierCol > 0 Then
            documentCol = MatchColumn(r, documentCaptions)
            mHeaderRow = r
            FindColumns = True

            modLog.LogAction sampleIdx, "Read export", _
                         "Header on row " & r & ". Amount = column " & amountCol & _
                         " (""" & mCells(r, amountCol) & """), supplier = column " & _
                         supplierCol & " (""" & mCells(r, supplierCol) & """)" & _
                         IIf(documentCol > 0, _
                             ", document = column " & documentCol & " (""" & _
                             mCells(r, documentCol) & """)", _
                             ", no document column matched"), "OK", mPath
            Exit Function
        End If
    Next r

    modLog.LogAction sampleIdx, "Read export", _
                 "No header row in " & mPath & " matched the expected captions. Open the " & _
                 "file, then copy its real column headings into the '" & amountSetting & _
                 "' and '" & supplierSetting & "' settings on the Control sheet.", _
                 "ERROR", mPath
End Function

' Index of the first column on the given row whose caption matches one of the
' pipe-separated candidates. Compared on letters and digits only, so
' 'Amnt in loc.cur.' matches 'Amnt in loc cur'.
Public Function MatchColumn(ByVal row As Long, ByVal candidates As String) As Long
    Dim options() As String
    Dim c As Long, j As Long
    Dim wanted As String

    options = Split(candidates, "|")

    For j = LBound(options) To UBound(options)
        wanted = Normalise(options(j))
        If Len(wanted) > 0 Then
            For c = 1 To mColCount
                If Normalise(mCells(row, c)) = wanted Then
                    MatchColumn = c
                    Exit Function
                End If
            Next c
        End If
    Next j
End Function

Private Function CaptionSetting(ByVal label As String, ByVal fallback As String) As String
    Dim configured As String

    If Len(label) > 0 Then
        On Error Resume Next
        configured = Trim$(modConfig.Setting(label))
        On Error GoTo 0
    End If

    If Len(configured) > 0 Then
        CaptionSetting = configured & "|" & fallback
    Else
        CaptionSetting = fallback
    End If
End Function

Private Function Normalise(ByVal text As String) As String
    Dim i As Long
    Dim ch As String
    Dim upper As String

    upper = UCase$(text)
    For i = 1 To Len(upper)
        ch = Mid$(upper, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then
            Normalise = Normalise & ch
        End If
    Next i
End Function

Private Function LesserOf(ByVal a As Long, ByVal b As Long) As Long
    LesserOf = IIf(a < b, a, b)
End Function

'-----------------------------------------------------------------------
' The largest-magnitude row of a loaded export.
'-----------------------------------------------------------------------
Public Function LargestRow(ByVal path As String, ByVal sampleIdx As Long, _
                           ByVal amountSetting As String, _
                           ByVal supplierSetting As String, _
                           ByVal documentSetting As String) As ListRow
    Dim result As ListRow
    Dim amountCol As Long, supplierCol As Long, documentCol As Long
    Dim r As Long
    Dim value As Double, best As Double

    If Not LoadExport(path, sampleIdx) Then Exit Function
    If Not FindColumns(sampleIdx, amountSetting, supplierSetting, documentSetting, _
                       amountCol, supplierCol, documentCol) Then Exit Function

    For r = mHeaderRow + 1 To mRowCount
        ' A total or subtotal line would otherwise win outright, and it is not
        ' an invoice. Skip any row whose supplier cell is empty or whose text
        ' reads like a total.
        If Not IsTotalRow(r, supplierCol) Then
            value = Abs(modUtil.ParseSapAmount(mCells(r, amountCol)))

            If value > 0 Then result.RowsConsidered = result.RowsConsidered + 1

            ' Strictly greater, so the first of two equal rows wins and the
            ' answer does not depend on the order SAP happened to print.
            If value > best Then
                best = value
                result.Found = True
                result.Amount = value
                result.SourceRow = r
                result.Supplier = mCells(r, supplierCol)
                If documentCol > 0 Then result.DocumentNumber = mCells(r, documentCol)
            End If
        End If
    Next r

    If result.Found And result.RowsConsidered = 1 Then
        modLog.LogAction sampleIdx, "Read export", _
                     "Only one usable row in " & path & ". Check the file -- a list " & _
                     "that should hold several payments may have exported wrongly.", _
                     "MANUAL", path
    End If

    LargestRow = result
End Function

' Totals rows are the trap here: SAP prints them with the largest number on
' the list, so a naive maximum picks the total instead of an invoice.
Private Function IsTotalRow(ByVal row As Long, ByVal supplierCol As Long) As Boolean
    Dim label As String
    Dim c As Long
    Dim filled As Long

    If supplierCol > 0 Then
        label = Normalise(mCells(row, supplierCol))
        If Len(label) = 0 Then
            ' No supplier at all. A subtotal line looks exactly like this.
            IsTotalRow = True
            Exit Function
        End If
        If label = "TOTAL" Or label = "GRANDTOTAL" Or label = "SUBTOTAL" Or _
           label = "TOTALS" Or Left$(label, 5) = "TOTAL" Then
            IsTotalRow = True
            Exit Function
        End If
    End If

    ' A row with only one or two cells filled is furniture, not data.
    For c = 1 To mColCount
        If Len(mCells(row, c)) > 0 Then filled = filled + 1
    Next c
    If filled <= 2 Then IsTotalRow = True
End Function

'-----------------------------------------------------------------------
' Distinct values of one column, optionally filtered on another.
'
' Used for step 5 -> 6: pull the ZP document numbers out of the exported
' Payment Usage list so they can be fed to FBL1N. Returns them newline-
' separated, and reports how many rows were rejected by the filter so a
' wrong document-type caption is visible rather than silent.
'-----------------------------------------------------------------------
Public Function DocumentNumbersOfType(ByVal path As String, ByVal sampleIdx As Long, _
                                      ByVal wantedType As String, _
                                      ByRef matched As Long, _
                                      ByRef rejected As Long) As String
    Dim documentCol As Long, typeCol As Long
    Dim r As Long
    Dim seen As Object
    Dim number As String, rowType As String
    Dim results As String

    If Not LoadExport(path, sampleIdx) Then Exit Function

    ' Find the header row by looking for a document-number column on its own,
    ' since the Payment Usage list has no supplier column to anchor on.
    For r = 1 To LesserOf(mRowCount, 80)
        documentCol = MatchColumn(r, CaptionSetting("Payment usage document column", _
                                                   DEFAULT_DOCUMENT_CAPTIONS))
        If documentCol > 0 Then
            mHeaderRow = r
            typeCol = MatchColumn(r, CaptionSetting("Payment usage type column", _
                                                   DEFAULT_DOCTYPE_CAPTIONS))
            Exit For
        End If
    Next r

    If documentCol = 0 Then
        modLog.LogAction sampleIdx, "ZP numbers", _
                     "No document-number column found in " & path & ". Open it and name " & _
                     "the heading in 'Payment usage document column' on the Control sheet.", _
                     "ERROR", path
        Exit Function
    End If

    If typeCol = 0 Then
        modLog.LogAction sampleIdx, "ZP numbers", _
                     "No document-type column found in " & path & ", so every document " & _
                     "number was taken and none could be filtered to " & wantedType & _
                     ". Name the heading in 'Payment usage type column' on the Control " & _
                     "sheet, or FBL1N will be given non-" & wantedType & " documents too.", _
                     "MANUAL", path
    End If

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    For r = mHeaderRow + 1 To mRowCount
        number = OnlyDigits(mCells(r, documentCol))

        If Len(number) > 0 Then
            If typeCol > 0 Then
                rowType = Normalise(mCells(r, typeCol))
                If rowType <> Normalise(wantedType) Then
                    rejected = rejected + 1
                    GoTo NextRow
                End If
            End If

            If Not seen.Exists(number) Then
                seen.Add number, True
                results = results & IIf(Len(results) > 0, vbLf, "") & number
                matched = matched + 1
            End If
        End If
NextRow:
    Next r

    modLog.LogAction sampleIdx, "ZP numbers", _
                 matched & " distinct " & wantedType & " document number(s) taken from " & _
                 path & IIf(rejected > 0, ", " & rejected & " row(s) rejected as a " & _
                 "different document type", "") & ".", _
                 IIf(matched > 0, "OK", "ERROR"), path

    DocumentNumbersOfType = results
End Function

' SAP prints document numbers with leading zeros and sometimes a company-code
' suffix. Digits only makes them comparable and safe to type into a filter.
Private Function OnlyDigits(ByVal text As String) As String
    Dim i As Long
    Dim ch As String

    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch >= "0" And ch <= "9" Then OnlyDigits = OnlyDigits & ch
    Next i
End Function
