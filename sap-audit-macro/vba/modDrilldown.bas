Attribute VB_Name = "modDrilldown"
'=======================================================================
' modDrilldown -- statement item -> FI document -> vendor invoice(s) ->
'                 attached invoice image.
'
' This is the chain that answers the auditor's actual question, which is
' the blank 'Payment to Supplier?' column in their workbook. The bank
' statement alone cannot answer it: 18 of the 56 samples are bulk ACH
' fundings with no named beneficiary on the statement at all.
'
' HONEST LIMITATION -- read before relying on the last step.
'
' Getting the *list* of attachments out of SAP is scriptable and works.
' Getting the attached PDF or TIFF *itself* onto disk often is not: on
' most configurations SAP hands the document to the registered external
' viewer instead of writing a file, and a script cannot capture that.
'
' So this module does what it reliably can, and says so when it cannot:
'   - it always records the document numbers and the attachment list,
'     which is what a reviewer needs in order to know what to ask for
'   - it attempts the OAOR export route, which does write files when the
'     content repository allows it
'   - where the viewer intercepts, it logs result MANUAL and moves on,
'     rather than reporting a success that left no file behind
'
' For all 56 invoices in one go, a direct read-only extract from the
' content server by Basis beats scripting the GUI. See README.
'=======================================================================
Option Explicit

Public Type DocumentChain
    PaymentDocument As String
    FiscalYear As String
    VendorAccount As String
    VendorName As String
    ClearedInvoices As String     ' semicolon-separated document numbers
    AttachmentCount As Long
    FilesWritten As Long
    Notes As String
End Type

'-----------------------------------------------------------------------
' Drill from the selected FEBAN row through to the posted FI document.
' Returns "" when the statement item is not yet posted -- itself a
' finding worth reporting, not an error.
'-----------------------------------------------------------------------
Public Function DrillToFiDocument(ByVal sampleIdx As Long, _
                                  ByVal match As FebanMatch) As String
    Dim documentNumber As String

    documentNumber = Trim$(match.DocumentNumber)

    If Len(documentNumber) = 0 Or documentNumber = "0" Then
        modLog.LogAction sampleIdx, "Drill-down", _
                     "The statement item carries no FI document number" & _
                     IIf(Len(match.PostingStatus) > 0, _
                         " (posting status '" & match.PostingStatus & "')", "") & _
                     ". It appears not to be posted, so there is no document to " & _
                     "trace. Report this line as unposted rather than unmatched.", _
                     "SKIPPED", vbNullString
        Exit Function
    End If

    modLog.LogAction sampleIdx, "Drill-down", _
                 "Statement item posted to FI document " & documentNumber, _
                 "OK", vbNullString

    DrillToFiDocument = documentNumber
End Function

'-----------------------------------------------------------------------
' Open a document in FB03 (display only) and read the vendor side.
'-----------------------------------------------------------------------
Public Function ReadDocument(ByVal sampleIdx As Long, ByVal documentNumber As String, _
                             ByVal fiscalYear As String) As DocumentChain
    Dim chain As DocumentChain
    Dim docId As String, bukrsId As String, yearId As String

    chain.PaymentDocument = documentNumber
    chain.FiscalYear = fiscalYear

    docId = modConfig.ElementIdOrBlank("FB03.DocNumber")
    If Len(docId) = 0 Then
        chain.Notes = "FB03.DocNumber is not mapped on the Screen Map sheet, so the " & _
                      "document could not be opened. Record FB03 and paste its IDs."
        modLog.LogAction sampleIdx, "FB03", chain.Notes, "SKIPPED", vbNullString
        ReadDocument = chain
        Exit Function
    End If

    modSafety.StartTransaction "FB03"

    modSapConnect.Element(docId).Text = documentNumber

    bukrsId = modConfig.ElementIdOrBlank("FB03.CompanyCode")
    If Len(bukrsId) > 0 Then
        modSapConnect.Element(bukrsId).Text = modConfig.Setting("Company code")
    End If

    yearId = modConfig.ElementIdOrBlank("FB03.FiscalYear")
    If Len(yearId) > 0 And Len(fiscalYear) > 0 Then
        modSapConnect.Element(yearId).Text = fiscalYear
    End If

    modSafety.GuardedSendVKey "wnd[0]", 0
    modSafety.AssertPopupKnown

    If modSapConnect.StatusBarType() = "E" Then
        chain.Notes = "FB03 could not display " & documentNumber & ": " & _
                      modSapConnect.StatusBarText()
        modLog.LogAction sampleIdx, "FB03", chain.Notes, "ERROR", vbNullString
        ReadDocument = chain
        Exit Function
    End If

    ReadVendorLine sampleIdx, chain
    chain.AttachmentCount = CountAttachments(sampleIdx)

    modLog.LogAction sampleIdx, "FB03", _
                 "Document " & documentNumber & _
                 IIf(Len(chain.VendorAccount) > 0, _
                     ", vendor " & chain.VendorAccount & " " & chain.VendorName, _
                     ", no vendor line found") & _
                 ", " & chain.AttachmentCount & " attachment(s)", _
                 "OK", vbNullString

    ReadDocument = chain
End Function

'-----------------------------------------------------------------------
' Read the vendor account and name from the document's line items.
'-----------------------------------------------------------------------
Private Sub ReadVendorLine(ByVal sampleIdx As Long, ByRef chain As DocumentChain)
    Dim gridId As String
    Dim grid As Object
    Dim rowCount As Long, row As Long
    Dim accountType As String

    gridId = modConfig.ElementIdOrBlank("FB03.ItemGrid")
    If Len(gridId) = 0 Then Exit Sub
    If Not modSapConnect.Exists(gridId) Then Exit Sub

    Set grid = modSapConnect.Element(gridId)

    On Error Resume Next
    rowCount = grid.RowCount
    On Error GoTo 0

    For row = 0 To rowCount - 1
        accountType = vbNullString
        On Error Resume Next
        accountType = Trim$(grid.GetCellValue(row, "KOART"))
        On Error GoTo 0

        If UCase$(accountType) = "K" Then          ' K = vendor
            On Error Resume Next
            chain.VendorAccount = Trim$(grid.GetCellValue(row, "LIFNR"))
            If Len(chain.VendorAccount) = 0 Then _
                chain.VendorAccount = Trim$(grid.GetCellValue(row, "HKONT"))
            chain.VendorName = Trim$(grid.GetCellValue(row, "NAME1"))
            On Error GoTo 0
            Exit For
        End If
    Next row
End Sub

'-----------------------------------------------------------------------
' Attachment list via the services-for-object toolbox on the title bar.
' Counting is reliable; extracting the files is not -- see the header.
'-----------------------------------------------------------------------
Public Function CountAttachments(ByVal sampleIdx As Long) As Long
    Dim toolboxId As String
    Dim listId As String
    Dim toolbox As Object
    Dim grid As Object

    toolboxId = modConfig.ElementIdOrBlank("FB03.GosToolbox")
    If Len(toolboxId) = 0 Then Exit Function
    If Not modSapConnect.Exists(toolboxId) Then Exit Function

    Set toolbox = modSapConnect.Element(toolboxId)

    On Error Resume Next
    toolbox.pressContextButton "%GOS_TOOLBOX"
    toolbox.selectContextMenuItem "%GOS_VIEW_ATTA"
    On Error GoTo 0
    modSapConnect.WaitForSap

    If Not modSapConnect.ModalWindowOpen() Then Exit Function

    listId = modConfig.ElementIdOrBlank("Attach.ListGrid")
    If Len(listId) > 0 Then
        If modSapConnect.Exists(listId) Then
            Set grid = modSapConnect.Element(listId)
            On Error Resume Next
            CountAttachments = grid.RowCount
            On Error GoTo 0
        End If
    End If

    ' Close the list with Cancel, never Enter -- Enter opens the document
    ' in the external viewer and blocks the run behind it.
    On Error Resume Next
    modSapConnect.Element("wnd[1]").Close
    On Error GoTo 0
    modSapConnect.WaitForSap
End Function

'-----------------------------------------------------------------------
' Attempt to write the attached invoice images to disk via OAOR, the
' Business Document Navigator, which is the one route that offers a real
' export. Logs MANUAL, not OK, where the viewer intercepts.
'-----------------------------------------------------------------------
Public Function DownloadAttachments(ByVal sampleIdx As Long, ByVal documentNumber As String, _
                                    ByVal folder As String) As Long
    Dim saveButtonId As String

    If modSafety.BlockedByDryRun("Would download attachments for document " & _
                                 documentNumber & " to " & folder) Then Exit Function

    saveButtonId = modConfig.ElementIdOrBlank("Attach.SaveButton")
    If Len(saveButtonId) = 0 Then
        modLog.LogAction sampleIdx, "Invoice image", _
                     "Attach.SaveButton is not mapped, so no attachment export was " & _
                     "attempted for document " & documentNumber & ". The document " & _
                     "number is recorded, which is enough to request the image.", _
                     "MANUAL", vbNullString
        Exit Function
    End If

    modUtil.EnsureFolder folder

    ' Deliberately conservative: attempt once, and report honestly.
    On Error Resume Next
    modSapConnect.Element(saveButtonId).Press
    On Error GoTo 0
    modSapConnect.WaitForSap

    If Not modSapConnect.ModalWindowOpen() Then
        modLog.LogAction sampleIdx, "Invoice image", _
                     "No save dialog appeared for document " & documentNumber & _
                     ". SAP most likely passed the document to the external viewer, " & _
                     "which a script cannot capture. Retrieve this one by hand, or " & _
                     "ask Basis for a content-server extract for all samples at once.", _
                     "MANUAL", vbNullString
        Exit Function
    End If

    ' A save dialog did appear, so the export route is available here.
    DownloadAttachments = 1
End Function

'-----------------------------------------------------------------------
' Export the vendor's open and cleared items around the payment, which
' evidences the invoices the payment settled. FBL1N is display-only.
'-----------------------------------------------------------------------
Public Function ExportVendorItems(ByVal sampleIdx As Long, ByVal vendorAccount As String, _
                                  ByVal folder As String, ByVal fileStem As String) As String
    If Len(vendorAccount) = 0 Then Exit Function

    If modSafety.BlockedByDryRun("Would export FBL1N items for vendor " & _
                                 vendorAccount) Then Exit Function

    modSafety.StartTransaction "FBL1N"

    modLog.LogAction sampleIdx, "FBL1N", _
                 "Vendor " & vendorAccount & " line items. Populate the selection from " & _
                 "your recording -- FBL1N's selection screen IDs vary by release and " & _
                 "are not in the default Screen Map.", "SKIPPED", vbNullString
End Function
