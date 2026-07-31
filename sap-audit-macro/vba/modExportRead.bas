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

' LANGUAGE INDEPENDENCE
'
' Column captions are translated -- 'Document Type' is 'Belegart' on a
' German logon, 'Tipo de documento' on a Spanish one -- so matching on them
' only works for whoever happened to write the list. Everything below tries
' three things in order, and only the first is language-dependent:
'
'   1. a caption the operator named on the Control sheet, or one of the
'      built-in captions listed here
'   2. the SAP technical field name, which is the same in every language:
'      BELNR, BLART, DMBTR, WRBTR, LIFNR, NAME1
'   3. what the column actually contains -- the amount column is the one
'      whose cells parse as signed decimals, the document-type column is
'      the one that actually holds the wanted code, the document-number
'      column is the one holding long digit strings
'
' Step 3 is the one that always works. The captions are kept because when
' they do match they are unambiguous, and because a wrong guess at step 3
' is worth avoiding on a list with several numeric columns.
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
            mCells(r, c) = CellText(area.Cells(r, c))
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

' .Text is what the cell DISPLAYS, which is what we want: it carries the
' thousands and decimal separators SAP wrote, so ParseSapAmount can read
' them. But a column too narrow for its number displays as ##### -- and
' that would parse to zero, silently, making the largest payment look like
' nothing. Fall back to the underlying value in that case.
Private Function CellText(ByVal cell As Object) As String
    Dim shown As String

    On Error Resume Next
    shown = CStr(cell.Text)
    On Error GoTo 0

    If Len(Trim$(shown)) = 0 Or InStr(shown, "##") > 0 Then
        On Error Resume Next
        If Not IsEmpty(cell.Value) Then shown = CStr(cell.Value)
        On Error GoTo 0
    End If

    CellText = modUtil.Squeeze(shown)
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

    ' No caption matched -- which is the normal case on a non-English logon.
    ' Fall back to what the columns contain, and say which it picked so a bad
    ' guess is visible on the first sample rather than after all 56.
    mHeaderRow = 1
    amountCol = ColumnThatLooksLikeAmount(2)

    If amountCol = 0 Then
        modLog.LogAction sampleIdx, "Read export", _
                     "No column in " & mPath & " matched a known caption or looks like " & _
                     "an amount. Open the file and name its headings in the '" & _
                     amountSetting & "' setting on the Control sheet.", "ERROR", mPath
        Exit Function
    End If

    supplierCol = ColumnThatLooksLikeName(2, amountCol)
    documentCol = ColumnThatLooksLikeDocumentNumber(2, amountCol)
    FindColumns = True

    modLog.LogAction sampleIdx, "Read export", _
                 "No caption matched, so columns were identified by content: amount = " & _
                 "column " & amountCol & ", name = column " & supplierCol & _
                 ", document = column " & documentCol & ". Check that against " & mPath & _
                 " -- if it is wrong, name the headings on the Control sheet.", _
                 "MANUAL", mPath
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
' Content-based column detection -- the language-independent fallback.
'-----------------------------------------------------------------------
' The column whose cells most often parse as a non-zero decimal amount.
' Ties break towards the column with the largest magnitude, since a
' quantity or a tax rate column rarely carries the biggest numbers.
Public Function ColumnThatLooksLikeAmount(ByVal firstDataRow As Long) As Long
    Dim c As Long, r As Long
    Dim hits As Long, bestHits As Long
    Dim biggest As Double, bestBiggest As Double
    Dim value As Double
    Dim text As String

    For c = 1 To mColCount
        hits = 0
        biggest = 0

        For r = firstDataRow To mRowCount
            text = mCells(r, c)
            ' A decimal separator is what distinguishes money from a document
            ' number -- both are digit strings otherwise.
            If InStr(text, ",") > 0 Or InStr(text, ".") > 0 Then
                value = Abs(modUtil.ParseSapAmount(text))
                If value <> 0 Then
                    hits = hits + 1
                    If value > biggest Then biggest = value
                End If
            End If
        Next r

        If hits > bestHits Or (hits = bestHits And hits > 0 And biggest > bestBiggest) Then
            bestHits = hits
            bestBiggest = biggest
            ColumnThatLooksLikeAmount = c
        End If
    Next c

    If bestHits = 0 Then ColumnThatLooksLikeAmount = 0
End Function

' The column that actually contains the wanted code. Finding the document
' type column by looking for 'ZP' in it needs no translation at all.
Public Function ColumnContaining(ByVal firstDataRow As Long, ByVal wanted As String) As Long
    Dim c As Long, r As Long
    Dim hits As Long, bestHits As Long
    Dim target As String

    target = Normalise(wanted)
    If Len(target) = 0 Then Exit Function

    For c = 1 To mColCount
        hits = 0
        For r = firstDataRow To mRowCount
            If Normalise(mCells(r, c)) = target Then hits = hits + 1
        Next r

        If hits > bestHits Then
            bestHits = hits
            ColumnContaining = c
        End If
    Next c
End Function

' The column of long digit strings that is not the amount column -- a
' document number. Ignores anything with a decimal separator.
Public Function ColumnThatLooksLikeDocumentNumber(ByVal firstDataRow As Long, _
                                                  ByVal skipColumn As Long) As Long
    Dim c As Long, r As Long
    Dim hits As Long, bestHits As Long
    Dim text As String

    For c = 1 To mColCount
        If c <> skipColumn Then
            hits = 0
            For r = firstDataRow To mRowCount
                text = mCells(r, c)
                If Len(text) >= 8 And IsAllDigits(text) Then hits = hits + 1
            Next r

            If hits > bestHits Then
                bestHits = hits
                ColumnThatLooksLikeDocumentNumber = c
            End If
        End If
    Next c
End Function

' The column with the most alphabetic content -- a name.
Public Function ColumnThatLooksLikeName(ByVal firstDataRow As Long, _
                                        ByVal skipColumn As Long) As Long
    Dim c As Long, r As Long
    Dim score As Long, bestScore As Long

    For c = 1 To mColCount
        If c <> skipColumn Then
            score = 0
            For r = firstDataRow To mRowCount
                score = score + LetterCount(mCells(r, c))
            Next r

            If score > bestScore Then
                bestScore = score
                ColumnThatLooksLikeName = c
            End If
        End If
    Next c
End Function

Private Function IsAllDigits(ByVal text As String) As Boolean
    Dim i As Long
    Dim ch As String

    If Len(text) = 0 Then Exit Function
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    IsAllDigits = True
End Function

Private Function LetterCount(ByVal text As String) As Long
    Dim i As Long
    Dim ch As String

    For i = 1 To Len(text)
        ch = UCase$(Mid$(text, i, 1))
        If ch >= "A" And ch <= "Z" Then LetterCount = LetterCount + 1
    Next i
End Function

'-----------------------------------------------------------------------
' The largest-magnitude row of a loaded export.
'-----------------------------------------------------------------------
Public Function LargestRow(ByVal path As String, ByVal sampleIdx As Long, _
                           ByVal amountSetting As String, _
                           ByVal supplierSetting As String, _
                           ByVal documentSetting As String) As ListRow
    LargestRow = LargestRowOfType(path, sampleIdx, amountSetting, supplierSetting, _
                                  documentSetting, vbNullString)
End Function

'-----------------------------------------------------------------------
' The invoice: the most negative row of a Payment Usage list.
'
' This used to filter on document type, and the type was wrong -- KR is the
' SAP standard but this system's first exported invoice came back as RN, so
' the filter matched nothing on a file that plainly held an invoice. Codes
' are configuration: they differ by company code, by client, by release.
'
' The SIGN does not. In a vendor line-item list the payment is a debit and
' the invoice it settles is a credit, so the invoice is the negative row and
' the biggest invoice is the most negative one. Nothing to configure and
' nothing to keep in step with a customising table.
'
' 'Invoice document type' survives as a CROSS-CHECK only: when the row this
' picks is not one of those types, the Log says so and still takes it. It
' can no longer decide the answer, so it can no longer break it.
'-----------------------------------------------------------------------
Public Function MostNegativeRow(ByVal path As String, ByVal sampleIdx As Long, _
                                ByVal amountSetting As String, _
                                ByVal supplierSetting As String, _
                                ByVal documentSetting As String, _
                                ByVal expectedTypes As String) As ListRow
    Dim result As ListRow
    Dim amountCol As Long, supplierCol As Long, documentCol As Long
    Dim typeCol As Long
    Dim r As Long
    Dim value As Double, best As Double
    Dim chosenType As String
    Dim positives As Long
    Dim biggestPositive As Double

    If Not LoadExport(path, sampleIdx) Then Exit Function
    If Not FindColumns(sampleIdx, amountSetting, supplierSetting, documentSetting, _
                       amountCol, supplierCol, documentCol) Then Exit Function

    typeCol = MatchColumn(mHeaderRow, CaptionSetting(vbNullString, DEFAULT_DOCTYPE_CAPTIONS))

    For r = mHeaderRow + 1 To mRowCount
        If Not IsTotalRow(r, supplierCol) Then
            value = modUtil.ParseSapAmount(mCells(r, amountCol))

            If value < 0 Then
                result.RowsConsidered = result.RowsConsidered + 1

                ' Strictly less, so the first of two equal rows wins and the
                ' answer does not depend on the order SAP happened to print.
                If value < best Then
                    best = value
                    result.Found = True
                    result.Amount = Abs(value)
                    result.SourceRow = r
                    result.Supplier = mCells(r, supplierCol)
                    If documentCol > 0 Then result.DocumentNumber = mCells(r, documentCol)
                    If typeCol > 0 Then chosenType = Normalise(mCells(r, typeCol))
                End If
            ElseIf value > 0 Then
                positives = positives + 1
                If value > biggestPositive Then biggestPositive = value
            End If
        End If
    Next r

    If Not result.Found Then
        modLog.LogAction sampleIdx, "Read export", _
                     "No row in " & path & " carries a negative amount, so there is no " & _
                     "invoice to take. " & positives & " positive row(s), the largest " & _
                     Format$(biggestPositive, "#,##0.00") & ". Column " & amountCol & _
                     " was read as the amount -- check that against the file.", _
                     "ERROR", path
        Exit Function
    End If

    modLog.LogAction sampleIdx, "Read export", _
                 "Most negative of " & result.RowsConsidered & " credit row(s): " & _
                 Format$(-best, "#,##0.00") & _
                 IIf(Len(result.DocumentNumber) > 0, ", document " & result.DocumentNumber, "") & _
                 IIf(Len(chosenType) > 0, ", type " & chosenType, "") & _
                 IIf(Len(result.Supplier) > 0, ", " & result.Supplier, "") & _
                 ". Picked by sign, not by document type.", "OK", path

    ' Cross-check, never a veto.
    If Len(expectedTypes) > 0 And Len(chosenType) > 0 Then
        If Not TypeWanted(chosenType, expectedTypes) Then
            modLog.LogAction sampleIdx, "Read export", _
                         "That row is document type " & chosenType & ", which is not in " & _
                         "'Invoice document type' (" & expectedTypes & "). It was taken " & _
                         "anyway, because the sign identifies the invoice. Add " & _
                         chosenType & " to that setting if it is the invoice type here.", _
                         "MANUAL", path
        End If
    End If

    MostNegativeRow = result
End Function

' Same, restricted to one document type.
'
' Kept for the ZP payment step, which needs the largest by MAGNITUDE across
' a list where every row has the same sign. The invoice step no longer uses
' it -- see MostNegativeRow above for why.
Public Function LargestRowOfType(ByVal path As String, ByVal sampleIdx As Long, _
                                 ByVal amountSetting As String, _
                                 ByVal supplierSetting As String, _
                                 ByVal documentSetting As String, _
                                 ByVal wantedType As String) As ListRow
    Dim result As ListRow
    Dim amountCol As Long, supplierCol As Long, documentCol As Long
    Dim r As Long
    Dim value As Double, best As Double
    Dim typeCol As Long, skipped As Long
    Dim skippedTypes As Object
    Dim typesSeen As String

    If Not LoadExport(path, sampleIdx) Then Exit Function
    If Not FindColumns(sampleIdx, amountSetting, supplierSetting, documentSetting, _
                       amountCol, supplierCol, documentCol) Then Exit Function

    ' Locate the document-type column the same language-independent way: by
    ' finding the column that actually holds the code.
    If Len(wantedType) > 0 Then
        typeCol = MatchColumn(mHeaderRow, CaptionSetting(vbNullString, DEFAULT_DOCTYPE_CAPTIONS))
        If typeCol = 0 Then typeCol = ColumnContaining(mHeaderRow + 1, FirstType(wantedType))

        If typeCol = 0 Then
            modLog.LogAction sampleIdx, "Read export", _
                         "No document-type column found in " & path & ", so the largest " & _
                         "row was taken across ALL document types rather than only " & _
                         wantedType & ". Check the result -- a payment may have been " & _
                         "picked where an invoice was wanted.", "MANUAL", path
        Else
            modLog.LogAction sampleIdx, "Read export", _
                         "Restricting to document type " & wantedType & " using column " & _
                         typeCol & ".", "OK", path
        End If
    End If

    Set skippedTypes = CreateObject("Scripting.Dictionary")
    skippedTypes.CompareMode = vbTextCompare

    For r = mHeaderRow + 1 To mRowCount
        If typeCol > 0 Then
            If Not TypeWanted(Normalise(mCells(r, typeCol)), wantedType) Then
                skipped = skipped + 1
                If Len(Normalise(mCells(r, typeCol))) > 0 Then
                    If Not skippedTypes.Exists(Normalise(mCells(r, typeCol))) Then
                        skippedTypes.Add Normalise(mCells(r, typeCol)), True
                    End If
                End If
                GoTo NextRow
            End If
        End If

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

NextRow:
    Next r

    If Len(wantedType) > 0 And typeCol > 0 Then
        ' Naming the types that WERE there is the difference between 'no
        ' invoice found' and 'the invoices on this system are type RN, not
        ' KR' -- which is what column 6 of the first invoice export held.
        If Not result.Found And skippedTypes.Count > 0 Then
            typesSeen = " The types in the file: " & Join(skippedTypes.Keys, ", ") & _
                        ". If the invoice is one of those, add it to 'Invoice document " & _
                        "type' on the Control sheet -- it takes a comma-separated list."
        End If

        modLog.LogAction sampleIdx, "Read export", _
                     result.RowsConsidered & " row(s) of type " & wantedType & " considered, " & _
                     skipped & " of other types skipped." & typesSeen, _
                     IIf(result.Found, "OK", "ERROR"), path
    End If

    If result.Found And result.RowsConsidered = 1 Then
        modLog.LogAction sampleIdx, "Read export", _
                     "Only one usable row in " & path & ". Check the file -- a list " & _
                     "that should hold several payments may have exported wrongly.", _
                     "MANUAL", path
    End If

    LargestRowOfType = result
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
    Dim rejectedTypes As Object
    Dim number As String, rowType As String
    Dim results As String, typesSeen As String

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

    ' Caption match failed. Find the type column by looking for the wanted
    ' code itself -- 'ZP' is 'ZP' in every language -- and the document-number
    ' column by its shape. This is what makes the step language-independent.
    If typeCol = 0 Then
        typeCol = ColumnContaining(2, FirstType(wantedType))
        If typeCol > 0 Then
            mHeaderRow = 1
            modLog.LogAction sampleIdx, "ZP numbers", _
                         "No caption matched, so column " & typeCol & " was taken as the " & _
                         "document type because it holds '" & wantedType & "'.", _
                         "OK", path
        End If
    End If

    If documentCol = 0 Then
        documentCol = ColumnThatLooksLikeDocumentNumber(2, ColumnThatLooksLikeAmount(2))
        If documentCol > 0 Then
            mHeaderRow = 1
            modLog.LogAction sampleIdx, "ZP numbers", _
                         "No caption matched, so column " & documentCol & " was taken as " & _
                         "the document number because it holds long digit strings. Check " & _
                         "that against " & path & ".", "MANUAL", path
        End If
    End If

    If documentCol = 0 Then
        modLog.LogAction sampleIdx, "ZP numbers", _
                     "No document-number column found in " & path & ", by caption or by " & _
                     "content. Open it and name the heading in 'Payment usage document " & _
                     "column' on the Control sheet.", "ERROR", path
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

    Set rejectedTypes = CreateObject("Scripting.Dictionary")
    rejectedTypes.CompareMode = vbTextCompare

    For r = mHeaderRow + 1 To mRowCount
        number = OnlyDigits(mCells(r, documentCol))

        If Len(number) > 0 Then
            If typeCol > 0 Then
                rowType = Normalise(mCells(r, typeCol))
                If Not TypeWanted(rowType, wantedType) Then
                    rejected = rejected + 1
                    If Len(rowType) > 0 Then
                        If Not rejectedTypes.Exists(rowType) Then rejectedTypes.Add rowType, True
                    End If
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

    ' When nothing matched, the useful thing to say is what the file DID hold.
    ' Two samples came back '0 ZP documents, 2 rows rejected' with no way to
    ' tell whether the column was wrong or the batch simply is not a ZP run.
    If matched = 0 And rejectedTypes.Count > 0 Then
        typesSeen = " The types actually in the file: " & _
                    Join(rejectedTypes.Keys, ", ") & ". If one of those is the " & _
                    "payment document here, add it to 'Payment document type' on the " & _
                    "Control sheet -- it takes a comma-separated list."
    End If

    modLog.LogAction sampleIdx, "ZP numbers", _
                 matched & " distinct " & wantedType & " document number(s) taken from " & _
                 path & IIf(rejected > 0, ", " & rejected & " row(s) rejected as a " & _
                 "different document type", "") & "." & typesSeen, _
                 IIf(matched > 0, "OK", "ERROR"), path

    DocumentNumbersOfType = results
End Function

' The first code of the list, for the content probe that looks for a column
' actually holding it. Probing for the literal string "ZP, ZV" would match
' nothing.
Private Function FirstType(ByVal wantedList As String) As String
    Dim parts() As String

    parts = Split(Replace(Replace(wantedList, ";", ","), "|", ","), ",")
    FirstType = Trim$(parts(LBound(parts)))
End Function

' 'Payment document type' takes one code or several -- 'ZP' or 'ZP, ZV, KZ'.
' A batch paid outside the normal payment run carries a different code, and
' the operator should be able to add it without a code change.
Private Function TypeWanted(ByVal rowType As String, ByVal wantedList As String) As Boolean
    Dim wanted() As String
    Dim i As Long

    wanted = Split(Replace(Replace(wantedList, ";", ","), "|", ","), ",")

    For i = LBound(wanted) To UBound(wanted)
        If Len(Trim$(wanted(i))) > 0 Then
            If StrComp(rowType, Normalise(wanted(i)), vbTextCompare) = 0 Then
                TypeWanted = True
                Exit Function
            End If
        End If
    Next i
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
