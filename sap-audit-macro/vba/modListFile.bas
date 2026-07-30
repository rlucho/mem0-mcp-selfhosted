Attribute VB_Name = "modListFile"
'=======================================================================
' modListFile -- reads a SAP classic-list export back off disk.
'
' Deciding which cleared item is the largest, and whether its supplier is
' the confirming party, is done by parsing the exported file rather than
' by scraping the list off the screen. Two reasons: the export happens
' anyway, and a classic SAP list on screen is a grid of lbl[x,y] elements
' that has to be reassembled by pixel position, which breaks the moment
' anyone changes a column width.
'
' The parser is deliberately loud about what it decided. It logs which
' column it took as the amount and which as the supplier, so a wrong guess
' shows up on the first sample instead of quietly skewing all 56.
'=======================================================================
Option Explicit

Public Type ListRow
    Found As Boolean
    Supplier As String
    Amount As Double
    DocumentNumber As String
    RowsConsidered As Long
End Type

' SAP writes these lists in the logged-on language, so the captions to look
' for are settings rather than constants. These are the fallbacks.
Private Const DEFAULT_AMOUNT_CAPTIONS As String = _
    "Amount in local currency|Amount in LC|Amount|LC amount|Amnt in loc.cur.|DMBTR|WRBTR"
Private Const DEFAULT_SUPPLIER_CAPTIONS As String = _
    "Name|Name 1|Name of vendor|Vendor name|Supplier|Account name|NAME1|Text"
Private Const DEFAULT_DOCUMENT_CAPTIONS As String = _
    "Document Number|DocumentNo|Doc. Number|Document no.|BELNR|Invoice reference"

'-----------------------------------------------------------------------
' The largest-magnitude row in an exported cleared-items list.
'-----------------------------------------------------------------------
Public Function LargestRow(ByVal path As String, ByVal sampleIdx As Long) As ListRow
    Dim result As ListRow
    Dim lines() As String
    Dim delimiter As String
    Dim headerIndex As Long
    Dim fields() As String
    Dim amountCol As Long, supplierCol As Long, documentCol As Long
    Dim i As Long
    Dim value As Double
    Dim best As Double

    If Not modUtil.FileExists(path) Then Exit Function

    lines = ReadLines(path)
    If UBound(lines) < LBound(lines) Then Exit Function

    delimiter = DetectDelimiter(lines)
    headerIndex = FindHeaderIndex(lines, delimiter, amountCol, supplierCol, documentCol)

    If headerIndex < 0 Then
        ' No recognisable header. Fall back to picking the column that looks
        ' most like money, and log that this is what happened.
        amountCol = GuessAmountColumn(lines, delimiter)
        supplierCol = GuessSupplierColumn(lines, delimiter, amountCol)
        documentCol = -1

        If amountCol < 0 Or supplierCol < 0 Then
            modLog.LogAction sampleIdx, "Parse list", _
                         "No header row recognised in " & path & " and the columns " & _
                         "could not be guessed either. Add this file's caption to the " & _
                         "'Cleared list ...' settings on the Control sheet.", _
                         "ERROR", path
            Exit Function
        End If

        modLog.LogAction sampleIdx, "Parse list", _
                     "No header row matched the configured captions in " & path & _
                     ". Fell back to column " & amountCol & " for the amount and " & _
                     supplierCol & " for the supplier. Check that against the file " & _
                     "before trusting the result.", "MANUAL", path
    Else
        modLog.LogAction sampleIdx, "Parse list", _
                     "Header on line " & (headerIndex + 1) & ". Amount = column " & _
                     amountCol & ", supplier = column " & supplierCol & _
                     IIf(documentCol >= 0, ", document = column " & documentCol, _
                         ", no document column found"), "OK", path
    End If

    For i = headerIndex + 1 To UBound(lines)
        If IsDataLine(lines(i)) Then
            fields = Split(lines(i), delimiter)

            If UBound(fields) >= amountCol Then
                value = Abs(modUtil.ParseSapAmount(Field(fields, amountCol)))

                ' Strictly greater, so the first of two equal rows wins and the
                ' result does not depend on the order SAP happened to print.
                If value > best Then
                    best = value
                    result.Found = True
                    result.Amount = value
                    result.Supplier = Field(fields, supplierCol)
                    result.DocumentNumber = Field(fields, documentCol)
                End If

                result.RowsConsidered = result.RowsConsidered + 1
            End If
        End If
    Next i

    LargestRow = result
End Function

'-----------------------------------------------------------------------
' File reading
'-----------------------------------------------------------------------
Private Function ReadLines(ByVal path As String) As String()
    Dim stream As Object
    Dim text As String

    ' ADODB reads whatever encoding SAP wrote without mangling the accented
    ' characters that turn up in supplier names.
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
    Dim handle As Integer
    handle = FreeFile
    Open path For Input As #handle
    text = Input$(LOF(handle), handle)
    Close #handle
    ReadLines = SplitLines(text)
End Function

Private Function SplitLines(ByVal text As String) As String()
    SplitLines = Split(Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf), vbLf)
End Function

'-----------------------------------------------------------------------
' Structure detection
'-----------------------------------------------------------------------
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

' Find the header line and, from it, the columns we need. Returns -1 when no
' line matches enough of the configured captions.
Private Function FindHeaderIndex(ByRef lines() As String, ByVal delimiter As String, _
                                 ByRef amountCol As Long, ByRef supplierCol As Long, _
                                 ByRef documentCol As Long) As Long
    Dim amountCaptions As String, supplierCaptions As String, documentCaptions As String
    Dim i As Long
    Dim fields() As String

    amountCaptions = CaptionSetting("Cleared list amount column", DEFAULT_AMOUNT_CAPTIONS)
    supplierCaptions = CaptionSetting("Cleared list supplier column", DEFAULT_SUPPLIER_CAPTIONS)
    documentCaptions = CaptionSetting("Cleared list document column", DEFAULT_DOCUMENT_CAPTIONS)

    FindHeaderIndex = -1

    For i = LBound(lines) To UBound(lines)
        If InStr(lines(i), delimiter) > 0 Then
            fields = Split(lines(i), delimiter)

            amountCol = MatchColumn(fields, amountCaptions)
            supplierCol = MatchColumn(fields, supplierCaptions)

            ' Both are needed for the row to be a usable header.
            If amountCol >= 0 And supplierCol >= 0 Then
                documentCol = MatchColumn(fields, documentCaptions)
                FindHeaderIndex = i
                Exit Function
            End If
        End If

        If i > 80 Then Exit For      ' headers are near the top of these lists
    Next i
End Function

Private Function CaptionSetting(ByVal label As String, ByVal fallback As String) As String
    Dim configured As String

    On Error Resume Next
    configured = Trim$(modConfig.Setting(label))
    On Error GoTo 0

    If Len(configured) > 0 Then
        CaptionSetting = configured & "|" & fallback
    Else
        CaptionSetting = fallback
    End If
End Function

' Index of the first field whose caption matches one of the pipe-separated
' candidates. Compared on letters and digits, so 'Amnt in loc.cur.' matches
' 'Amnt in loc cur'.
Private Function MatchColumn(ByRef fields() As String, ByVal candidates As String) As Long
    Dim options() As String
    Dim i As Long, j As Long
    Dim caption As String, wanted As String

    options = Split(candidates, "|")
    MatchColumn = -1

    For j = LBound(options) To UBound(options)
        wanted = Normalise(options(j))
        If Len(wanted) > 0 Then
            For i = LBound(fields) To UBound(fields)
                caption = Normalise(fields(i))
                If Len(caption) > 0 And caption = wanted Then
                    MatchColumn = i
                    Exit Function
                End If
            Next i
        End If
    Next j
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

'-----------------------------------------------------------------------
' Fallbacks used only when no header caption matched
'-----------------------------------------------------------------------
' The column where the most fields parse as a non-zero number.
Private Function GuessAmountColumn(ByRef lines() As String, ByVal delimiter As String) As Long
    Dim counts(0 To 63) As Long
    Dim i As Long, c As Long
    Dim fields() As String
    Dim best As Long, bestCount As Long

    For i = LBound(lines) To UBound(lines)
        If IsDataLine(lines(i)) Then
            fields = Split(lines(i), delimiter)
            For c = LBound(fields) To UBound(fields)
                If c <= 63 Then
                    If modUtil.ParseSapAmount(fields(c)) <> 0 Then counts(c) = counts(c) + 1
                End If
            Next c
        End If
    Next i

    best = -1
    For c = 0 To 63
        If counts(c) > bestCount Then
            bestCount = counts(c)
            best = c
        End If
    Next c

    GuessAmountColumn = best
End Function

' The column with the most alphabetic content, ignoring the amount column.
Private Function GuessSupplierColumn(ByRef lines() As String, ByVal delimiter As String, _
                                     ByVal skipColumn As Long) As Long
    Dim scores(0 To 63) As Long
    Dim i As Long, c As Long
    Dim fields() As String
    Dim best As Long, bestScore As Long

    For i = LBound(lines) To UBound(lines)
        If IsDataLine(lines(i)) Then
            fields = Split(lines(i), delimiter)
            For c = LBound(fields) To UBound(fields)
                If c <= 63 And c <> skipColumn Then
                    scores(c) = scores(c) + LetterCount(fields(c))
                End If
            Next c
        End If
    Next i

    best = -1
    For c = 0 To 63
        If scores(c) > bestScore Then
            bestScore = scores(c)
            best = c
        End If
    Next c

    GuessSupplierColumn = best
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
' A data line, as opposed to a rule, a page header or a blank.
'-----------------------------------------------------------------------
Private Function IsDataLine(ByVal text As String) As Boolean
    Dim stripped As String

    stripped = Replace(Replace(Replace(text, "-", ""), "|", ""), " ", "")
    stripped = Replace(Replace(stripped, vbTab, ""), "_", "")

    If Len(stripped) = 0 Then Exit Function              ' blank or a rule
    If Len(Trim$(text)) < 3 Then Exit Function

    IsDataLine = True
End Function

Private Function Field(ByRef fields() As String, ByVal index As Long) As String
    If index < 0 Then Exit Function
    If index > UBound(fields) Then Exit Function
    Field = modUtil.Squeeze(fields(index))
End Function
