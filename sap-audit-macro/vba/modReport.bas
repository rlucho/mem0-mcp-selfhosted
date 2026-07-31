Attribute VB_Name = "modReport"
'=======================================================================
' modReport -- one workbook per sample, for sending to the auditor.
'
' The evidence files are exports: raw SAP lists with thirty columns and no
' explanation of why they are in the folder. This builds the cover sheet
' that turns them into a pack -- what was asked for, what was found, how it
' was reached, and which file proves each step.
'
' Written into the sample's own folder as '01 - 8072447.42 - GBKM.xlsx', so
' the folder can be zipped and sent as it stands.
'
' One sheet: the trail, top to bottom, with the evidence files listed
' against the step each one proves and hyperlinked to their place in the
' folder.
'
' The PDF was embedded as an OLE object once. It depended on a PDF handler
' being registered, it failed on this machine, and it took the whole report
' with it -- and since the PDF is in the same folder as the report, it was
' never worth the failure mode.
'
' No log sheet either, and no absolute paths anywhere: this workbook is the
' thing that leaves the building. File names, yes -- they name the file
' sitting next to it in the same folder. Machine paths, no.
'
' Nothing here talks to SAP. It runs after the chain has finished, reads
' what the chain returned, and never raises -- a report that will not build
' must not undo evidence already on disk.
'=======================================================================
Option Explicit

Private Const COL_LABEL As Long = 2
Private Const COL_VALUE As Long = 3

Private Const HEAD_FILL As Long = 15921906      ' pale blue
Private Const SECTION_FILL As Long = 14277081   ' pale grey

'-----------------------------------------------------------------------
' Build the report. Returns the path written, or "" if it could not be.
'-----------------------------------------------------------------------
Public Function BuildSampleReport(ByVal sampleIdx As Long, ByVal monthTab As String, _
                                  ByVal payDate As Date, ByVal amount As Double, _
                                  ByVal party As String, ByVal reference As String, _
                                  ByRef match As FebanMatch, ByRef chain As ChainResult, _
                                  ByVal folder As String) As String
    Dim book As Workbook
    Dim sheet As Worksheet
    Dim path As String
    Dim row As Long
    Dim previousAlerts As Boolean
    Dim problem As String

    modUtil.EnsureFolder folder
    path = modUtil.JoinPath(folder, _
                            modUtil.ReportFileName(sampleIdx, amount, _
                                                   modConfig.Setting("Company code")))

    previousAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    On Error GoTo Failed

    Set book = Application.Workbooks.Add
    DropExtraSheets book

    Set sheet = book.Worksheets(1)
    sheet.Name = "Report"

    row = WriteHeader(sheet, sampleIdx, amount)
    row = WriteRequest(sheet, row, sampleIdx, monthTab, payDate, amount, party, reference)
    row = WriteTrail(sheet, row, match, chain)
    row = WriteOutcome(sheet, row, chain, folder)
    row = WriteFiles(sheet, row, folder, chain)
    row = WriteProvenance(sheet, row, monthTab)

    Layout sheet

    ' SAVE BEFORE EMBEDDING. Embedding the PDF is a nice-to-have that reaches
    ' outside Excel for a PDF OLE handler, and when that fails it took the
    ' whole report with it: every sample that reached an invoice -- 1, 3 and
    ' 4 -- ended up with no report at all, while sample 2, which had no PDF,
    ' got one. Writing the file first means the worst case is a report
    ' without an embedded copy of a PDF that is sitting next to it anyway.
    book.SaveAs fileName:=path, FileFormat:=51      ' xlOpenXMLWorkbook

    book.Close SaveChanges:=False
    Application.DisplayAlerts = previousAlerts

    BuildSampleReport = path
    Exit Function

Failed:
    problem = Err.Description

    On Error Resume Next
    If Not book Is Nothing Then book.Close SaveChanges:=False
    Application.DisplayAlerts = previousAlerts
    On Error GoTo 0

    ' Raise rather than return "". Returning a blank meant the caller logged
    ' neither a success nor a failure, so a report that never got written
    ' left no trace anywhere -- the only way to notice was to look in the
    ' folder and find it missing.
    Err.Raise vbObjectError + 580, "modReport.BuildSampleReport", problem
End Function

'-----------------------------------------------------------------------
' Sections
'-----------------------------------------------------------------------
Private Function WriteHeader(ByVal sheet As Worksheet, ByVal sampleIdx As Long, _
                             ByVal amount As Double) As Long
    sheet.Cells(2, COL_LABEL).Value = "Audit sample " & Format$(sampleIdx, "00") & _
                                      " -- " & Format$(amount, "#,##0.00") & _
                                      " -- company code " & modConfig.Setting("Company code")
    sheet.Cells(2, COL_LABEL).Font.Bold = True
    sheet.Cells(2, COL_LABEL).Font.Size = 14

    sheet.Cells(3, COL_LABEL).Value = _
        "Extracted from SAP " & modSapConnect.gSystemId & " client " & _
        modSapConnect.gClient & " by " & modSapConnect.gSapUser & " on " & _
        Format$(Now, "dd/mm/yyyy hh:mm") & _
        IIf(modConfig.IsDryRun(), "  (DRY RUN)", "")
    sheet.Cells(3, COL_LABEL).Font.Italic = True

    WriteHeader = 5
End Function

Private Function WriteRequest(ByVal sheet As Worksheet, ByVal startRow As Long, _
                              ByVal sampleIdx As Long, ByVal monthTab As String, _
                              ByVal payDate As Date, ByVal amount As Double, _
                              ByVal party As String, ByVal reference As String) As Long
    Dim row As Long

    row = Section(sheet, startRow, "What the auditor asked for")

    row = Pair(sheet, row, "Sample number", Format$(sampleIdx, "00"))
    row = Pair(sheet, row, "Period", monthTab)
    row = Pair(sheet, row, "Payment date", Format$(payDate, "dd/mm/yyyy"))
    row = Pair(sheet, row, "Amount", Format$(amount, "#,##0.00"))
    row = Pair(sheet, row, "Named party", _
               IIf(Len(party) > 0, party, "(none named in the request)"))
    row = Pair(sheet, row, "Payment reference", reference)

    WriteRequest = row + 1
End Function

Private Function WriteTrail(ByVal sheet As Worksheet, ByVal startRow As Long, _
                            ByRef match As FebanMatch, ByRef chain As ChainResult) As Long
    Dim row As Long

    row = Section(sheet, startRow, "How it was traced in SAP")

    row = Pair(sheet, row, "1. Bank statement line", _
               IIf(match.Found, _
                   "FEBAN " & match.StatementDate & ", row " & match.GridRow & ", " & _
                   Format$(match.StatementAmount, "#,##0.00"), _
                   "no matching statement line"))
    row = Pair(sheet, row, "2. FI document", Blank(chain.FiDocument))
    row = Pair(sheet, row, "3. Clearing document", Blank(chain.ClearingDocument))
    row = Pair(sheet, row, "4. Payments in the batch", _
               IIf(chain.ZpNumberCount > 0, _
                   chain.ZpNumberCount & " " & modConfig.Setting("Payment document type") & _
                   " document(s)", "--"))
    row = Pair(sheet, row, "5. Largest payment of the batch", _
               IIf(Len(chain.ZpPaymentDocument) > 0, _
                   Format$(chain.ZpPaymentAmount, "#,##0.00") & _
                   "  document " & chain.ZpPaymentDocument & _
                   IIf(Len(chain.ZpPaymentVendor) > 0, "  " & chain.ZpPaymentVendor, "") & _
                   IIf(chain.IsConfirmingPayment, _
                       "   (confirming party -- supply-chain finance)", ""), "--"))
    row = Pair(sheet, row, "6. Largest invoice behind it", _
               IIf(Len(chain.InvoiceNumber) > 0, _
                   Format$(chain.InvoiceAmount, "#,##0.00") & _
                   "  document " & chain.InvoiceNumber & _
                   IIf(Len(chain.InvoiceSupplier) > 0, "  " & chain.InvoiceSupplier, ""), "--"))

    WriteTrail = row + 1
End Function

Private Function WriteOutcome(ByVal sheet As Worksheet, ByVal startRow As Long, _
                              ByRef chain As ChainResult, ByVal folder As String) As Long
    Dim row As Long

    row = Section(sheet, startRow, "Outcome")

    row = Pair(sheet, row, "Status", chain.Status)
    Select Case chain.Status
        Case "DONE"
            sheet.Cells(row - 1, COL_VALUE).Font.Color = RGB(0, 112, 48)
        Case "NO CLEARING"
            sheet.Cells(row - 1, COL_VALUE).Font.Color = RGB(128, 96, 0)
        Case Else
            sheet.Cells(row - 1, COL_VALUE).Font.Color = RGB(192, 0, 0)
    End Select
    sheet.Cells(row - 1, COL_VALUE).Font.Bold = True

    row = Pair(sheet, row, "Notes", WithoutPaths(chain.Notes, folder))
    sheet.Cells(row - 1, COL_VALUE).WrapText = True

    WriteOutcome = row + 1
End Function

' Strip machine paths out of anything that reaches the auditor, leaving the
' file name. The notes are written for whoever is running the extract and
' happily quote 'C:\Users\eslucres\Documents\Audit GBKM\_dry run\Sep 25\
' 02 - 2161788.23\2 - Payment usage - batch of payments.xlsx' -- which says
' nothing to the reader and rather a lot about the operator's laptop.
'
' Folder names here contain spaces, so no amount of pattern-matching finds
' where a path ends. It does not have to: the two folders that can appear
' are both known, so this replaces those exact strings and leaves the file
' name behind. A backstop -- the messages themselves name files directly.
Private Function WithoutPaths(ByVal text As String, ByVal folder As String) As String
    Dim result As String

    result = text
    result = DropAll(result, folder & "\")
    result = DropAll(result, folder)
    result = DropAll(result, modConfig.DownloadRoot() & "\")
    result = DropAll(result, modConfig.DownloadRoot())

    WithoutPaths = Trim$(result)
End Function

Private Function DropAll(ByVal text As String, ByVal fragment As String) As String
    If Len(fragment) = 0 Then
        DropAll = text
    Else
        DropAll = Replace(text, fragment, vbNullString, 1, -1, vbTextCompare)
    End If
End Function

Private Function WriteFiles(ByVal sheet As Worksheet, ByVal startRow As Long, _
                            ByVal folder As String, ByRef chain As ChainResult) As Long
    Dim row As Long

    row = Section(sheet, startRow, "Evidence in this folder")

    row = FileLine(sheet, row, folder, modUtil.FILE_FEBAN, _
                   "The FEBAN bank statement list for the whole period, with the " & _
                   "sample's line in it")
    row = FileLine(sheet, row, folder, modUtil.FILE_FIDOC, _
                   "The FI document's line items -- the evidence where nothing clears")
    row = FileLine(sheet, row, folder, modUtil.FILE_BATCH, _
                   "Payment usage of the clearing document: every payment in the batch")
    row = FileLine(sheet, row, folder, modUtil.FILE_ZPLIST, _
                   "FBL1N over those payments, with vendor names and amounts")
    row = FileLine(sheet, row, folder, modUtil.FILE_INVOICES, _
                   "Payment usage of the largest payment: the invoices it settles")
    row = FileLine(sheet, row, folder, modUtil.FILE_PDF, _
                   "The invoice document itself, as attached in SAP")

    WriteFiles = row + 1
End Function

Private Function WriteProvenance(ByVal sheet As Worksheet, ByVal startRow As Long, _
                                 ByVal monthTab As String) As Long
    Dim row As Long

    row = Section(sheet, startRow, "How to read this")

    row = Note(sheet, row, _
        "Every file here is a SAP export or a SAP attachment. Nothing was retyped, " & _
        "and nothing was posted or changed in SAP to produce it -- the extract only " & _
        "displays and downloads.")
    row = Note(sheet, row, _
        "The largest payment of the batch and the largest invoice behind it are read " & _
        "from the exported files, not from what was on screen. The invoice is the most " & _
        "negative row: a payment is a debit and the invoice it settles is a credit.")
    row = Note(sheet, row, _
        "Every step above is timestamped in the run's audit trail, which is kept in the " & _
        "extract's control workbook against the SAP system, client and user it was " & _
        "taken from.")

    WriteProvenance = row + 1
End Function

'-----------------------------------------------------------------------
' Small builders
'-----------------------------------------------------------------------
Private Function Section(ByVal sheet As Worksheet, ByVal row As Long, _
                         ByVal title As String) As Long
    sheet.Cells(row, COL_LABEL).Value = title
    sheet.Cells(row, COL_LABEL).Font.Bold = True
    sheet.Range(sheet.Cells(row, COL_LABEL), sheet.Cells(row, COL_VALUE)).Interior.Color = SECTION_FILL
    Section = row + 1
End Function

Private Function Pair(ByVal sheet As Worksheet, ByVal row As Long, _
                      ByVal label As String, ByVal value As String) As Long
    sheet.Cells(row, COL_LABEL).Value = label
    sheet.Cells(row, COL_VALUE).Value = value
    Pair = row + 1
End Function

Private Function Note(ByVal sheet As Worksheet, ByVal row As Long, _
                      ByVal text As String) As Long
    sheet.Cells(row, COL_VALUE).Value = text
    sheet.Cells(row, COL_VALUE).WrapText = True
    Note = row + 1
End Function

' One row per evidence file, marked present or not. A file that is missing
' is stated as missing rather than left off the list -- an auditor should
' see what was not obtained as well as what was.
Private Function FileLine(ByVal sheet As Worksheet, ByVal row As Long, _
                          ByVal folder As String, ByVal fileName As String, _
                          ByVal what As String) As Long
    Dim path As String

    path = modUtil.JoinPath(folder, fileName)

    If modUtil.FileExists(path) Then
        sheet.Hyperlinks.Add Anchor:=sheet.Cells(row, COL_LABEL), _
                             Address:=fileName, TextToDisplay:=fileName
        sheet.Cells(row, COL_VALUE).Value = what
    Else
        sheet.Cells(row, COL_LABEL).Value = fileName
        sheet.Cells(row, COL_LABEL).Font.Color = RGB(128, 128, 128)
        sheet.Cells(row, COL_VALUE).Value = "not produced for this sample -- " & what
        sheet.Cells(row, COL_VALUE).Font.Color = RGB(128, 128, 128)
    End If

    FileLine = row + 1
End Function

Private Function Blank(ByVal text As String) As String
    Blank = IIf(Len(text) > 0, text, "--")
End Function

Private Sub Layout(ByVal sheet As Worksheet)
    sheet.Columns(1).ColumnWidth = 2
    sheet.Columns(COL_LABEL).ColumnWidth = 34
    sheet.Columns(COL_VALUE).ColumnWidth = 95
    sheet.Cells.VerticalAlignment = xlTop
    sheet.Rows(2).RowHeight = 20
End Sub

' A new workbook can come with three sheets depending on the operator's
' Excel options. The report builds its own, so drop the rest.
Private Sub DropExtraSheets(ByVal book As Workbook)
    Dim i As Long
    Dim previousAlerts As Boolean

    previousAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    On Error Resume Next
    For i = book.Worksheets.Count To 2 Step -1
        book.Worksheets(i).Delete
    Next i
    On Error GoTo 0

    Application.DisplayAlerts = previousAlerts
End Sub
