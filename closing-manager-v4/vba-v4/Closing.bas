Attribute VB_Name = "Closing"
Option Base 1
Public printN As Integer
Public EIS4 As Double
Dim k As Long
Public CC As String
Public PB, GAAP, NA, Cur
Dim Var
Dim rng As Range
Public OtherF As Boolean, BPCF As Boolean, CPCF As Boolean, BPCWDF As Boolean, ZGEISF As Boolean, BPCL As Boolean, CPCL As Boolean, BPCG As Boolean, CPCG As Boolean
Public BL As Boolean, BG As Boolean
Public FirstColumn As String, SAPID As String
Sub RunClosing()

'V4-CIO: show progress on the status bar so a slow run cannot be mistaken
'for a frozen one, and explain any failure in plain language.
On Error GoTo CM_Fail
CM_Begin 27

Dim sess As Object

Call CreateVariants
Call CreatePaths

SAPID = InputBox("Provide SAP user")
If SAPID = "" Then
    CM_Done
    Exit Sub
End If

SAPID = UCase(SAPID)

RunAgain = False
RunAgainXY = False
RunAgainXYAG = False
If Sheets("config").Range("B2") = Sheets("config").Range("AA1") And Sheets("config").Range("AA2") = "Error" Then
    RunAgain = True
End If
If Sheets("config").Range("B2") = Sheets("config").Range("AA1") And Left(Sheets("config").Range("AA12"), 2) = "XY" Then
    RunAgainXY = True
    printN = Replace(Sheets("config").Range("AA12"), "XY-", "")
End If
If Sheets("config").Range("B2") = Sheets("config").Range("AA1") And Left(Sheets("config").Range("AA17"), 2) = "XY" Then
    RunAgainXYAG = True
    printN = Replace(Sheets("config").Range("AA17"), "XY-", "")
End If

'----------------------------------------------------------------
'check tracker
'----------------------------------------------------------------
'Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
'xmlDoc.async = False
'
'Url = "https://troom-x.capgemini.com/sites/InternationalPaper/r2a/_vti_bin/Lists.asmx"
'List = "ClosingTracker"
'
'request = "<?xml version='1.0' encoding='utf-8'?>" & _
'            "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'" & _
'            " xmlns:xsd='http://www.w3.org/2001/XMLSchema'" & _
'            " xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>" & _
'            " <soap:Body>" & _
'                "<GetListItems xmlns='http://schemas.microsoft.com/sharepoint/soap/'>" & _
'                "<listName>" & List & "</listName>" & _
'                        "<QueryOptions>" & _
'                            "<IncludeMandatoryColumns>FALSE</IncludeMandatoryColumns>" & _
'                        "</QueryOptions>" & _
'                "<rowLimit>50000</rowLimit>" & _
'                " </GetListItems>" & _
'            " </soap:Body>" & _
'            "</soap:Envelope>"
'
'With CreateObject("Microsoft.XMLHTTP")
'    .Open "Get", Url, False, "", ""
'    .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
'    .setRequestHeader "SOAPAction", "http://schemas.microsoft.com/sharepoint/soap/GetListItems"
'    .send request
'
'    xmlDoc.LoadXML (.responsetext)
'    found = False
'    For Each X In xmlDoc.getElementsByTagName("z:row")
'        If Round(X.getAttribute("ows_Periodx"), 0) = Month(LastDay) And Round(X.getAttribute("ows_Yearx"), 0) = Year(LastDay) Then
'            found = True
'        End If
'    Next
'    If found = False Then
'        strError = "Data not updated for " & Month(LastDay) & "/" & Year(LastDay) & "." & vbNewLine & vbNewLine & "Please run it firstly."
'        With UF_Error
'            .Lbl_Error.Caption = strError
'            .StartUpPosition = 0
'            .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
'            .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
'            .Show
'        End With
'        Exit Sub
'    End If
'End With

'----------------------------------------------------------------
'SAP
'----------------------------------------------------------------
CM_Note "connecting to your SAP session"
Set sess = SAPsess

If Sheets("config").Range("B1") = "Yes" Then
    strError = "You are not logged in SAP. Please log in and run macro again."
    With UF_Error
        .Lbl_Error.Caption = strError
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        .Show
    End With
    CM_Done
    Exit Sub
End If

'----------------------------------------------------------------
'update tracker
'----------------------------------------------------------------
CM_Note "reading the company code and period"
CC = Sheets("config").Range("B2")
'List = "ClosingTracker"
'Date1 = Format(Now(), "yyyy-MM-ddTHH:mm:ssZ")
'
'updates = "<Batch> <Method ID='1' Cmd='New'>" & _
'            "<Field Name='UName'>" & UCase(Environ("username")) & "</Field>" & _
'            "<Field Name='LogType'>" & "Closing reports - start" & "</Field>" & _
'            "<Field Name='CC'>" & CC & "</Field>" & _
'            "<Field Name='Periodx'>" & Month(LastDay) & "</Field>" & _
'            "<Field Name='Yearx'>" & Year(LastDay) & "</Field>" & _
'            "<Field Name='Success'>1</Field>" & _
'            "<Field Name='Timestamp'>" & Date1 & "</Field>" & _
'            "</Method></Batch>"
'
'Call spAddToList(updates, List)

'----------------------------------------------------------------
'get SAP ID
'----------------------------------------------------------------
'List = "Employees"
'Url = "https://troom-x.capgemini.com/sites/InternationalPaper/CC/CG/_vti_bin/Lists.asmx"
'
'request = "<?xml version='1.0' encoding='utf-8'?>" & _
'            "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'" & _
'            " xmlns:xsd='http://www.w3.org/2001/XMLSchema'" & _
'            " xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>" & _
'            " <soap:Body>" & _
'                "<GetListItems xmlns='http://schemas.microsoft.com/sharepoint/soap/'>" & _
'                "<listName>" & List & "</listName>" & _
'                        "<QueryOptions>" & _
'                            "<IncludeMandatoryColumns>FALSE</IncludeMandatoryColumns>" & _
'                        "</QueryOptions>" & _
'                "<rowLimit>50000</rowLimit>" & _
'                " </GetListItems>" & _
'            " </soap:Body>" & _
'            "</soap:Envelope>"
'
'With CreateObject("Microsoft.XMLHTTP")
'    .Open "Get", Url, False, "", ""
'    .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
'    .setRequestHeader "SOAPAction", "http://schemas.microsoft.com/sharepoint/soap/GetListItems"
'    .send request
'
'    xmlDoc.LoadXML (.responsetext)
'    For Each X In xmlDoc.getElementsByTagName("z:row")
'
'        If UCase(X.getAttribute("ows_CG_x0020_Network_x0020_ID")) = UCase(Environ("Username")) Then
'            If X.getAttribute("ows_Active") = 1 And X.getAttribute("ows_onMaternity") = 0 Then
'                SAPID = X.getAttribute("ows_IP_x0020_network_x0020_ID")
'                Exit For
'            End If
'        End If
'    Next
'End With

'----------------------------------------------------------------
'clearing
'----------------------------------------------------------------
CM_Note "clearing the previous run's data"
LastRow = FindLastRow(1, 1, 0, 0, "ZGLRME")
If LastRow > 1 Then Sheets("ZGLRME").Range("A2", Sheets("ZGLRME").Cells(LastRow, 11)).ClearContents
If RunAgainXY = False Then
    LastRow = FindLastRow(1, 26, 0, 0, "config")
    Sheets("config").Range("AA1", Sheets("config").Cells(LastRow, 27)).ClearContents
End If
LastRow = FindLastRow(1, 2, 0, 0, "Errors")
Set rng = Sheets("Errors").Range("A1", Sheets("Errors").Cells(LastRow, 30))
Call ApplyBorders(0, 0, rng)
With Sheets("Errors").Cells
    .ClearContents
    .ClearFormats
    .Font.Bold = False
    .Interior.Pattern = xlSolid
    .Interior.PatternColorIndex = xlAutomatic
    .Interior.ThemeColor = xlThemeColorDark1
    .Interior.TintAndShade = 0
    .Interior.PatternTintAndShade = 0
End With
Sheets("Errors").Visible = xlVeryHidden
'LastRow = FindLastRow(1, 17, 0, 0, "config")
'If LastRow > 1 Then Sheets("config").Range("Q2", Sheets("config").Cells(LastRow, 21)).ClearContents
'LastRow = FindLastRow(1, 11, 0, 0, "config")
'If LastRow > 1 Then Sheets("config").Range("K2", Sheets("config").Cells(LastRow, 15)).ClearContents

'----------------------------------------------------------------
'create folders
'----------------------------------------------------------------
CM_Note "creating the working folders"
Set fso = CreateObject("Scripting.FileSystemObject")

Call EnsureFolders   'V4-CIO FIX: single-drive, parent-aware folder creation (was hardcoded C:\ here while CreatePaths used D:\ if present -> broke print/merge on D:-drive PCs).

'----------------------------------------------------------------
'import settings
'----------------------------------------------------------------
'variants
'List = "ClosingVariants"
'Url = "https://troom-x.capgemini.com/sites/InternationalPaper/r2a/_vti_bin/Lists.asmx"
'
'request = "<?xml version='1.0' encoding='utf-8'?>" & _
'            "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'" & _
'            " xmlns:xsd='http://www.w3.org/2001/XMLSchema'" & _
'            " xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>" & _
'            " <soap:Body>" & _
'                "<GetListItems xmlns='http://schemas.microsoft.com/sharepoint/soap/'>" & _
'                "<listName>" & List & "</listName>" & _
'                        "<QueryOptions>" & _
'                            "<IncludeMandatoryColumns>FALSE</IncludeMandatoryColumns>" & _
'                        "</QueryOptions>" & _
'                "<rowLimit>50000</rowLimit>" & _
'                " </GetListItems>" & _
'            " </soap:Body>" & _
'            "</soap:Envelope>"
'
'With CreateObject("Microsoft.XMLHTTP")
'    .Open "Get", Url, False, "", ""
'    .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
'    .setRequestHeader "SOAPAction", "http://schemas.microsoft.com/sharepoint/soap/GetListItems"
'    .send request
'
'    xmlDoc.LoadXML (.responsetext)
'    For Each X In xmlDoc.getElementsByTagName("z:row")
'        EmptRow = FindLastRow(1, 17, 1, 0, "config")
'        Sheets("config").Cells(EmptRow, 17) = X.getAttribute("ows_ID")
'        Sheets("config").Cells(EmptRow, 18) = X.getAttribute("ows_CC")
'        Sheets("config").Cells(EmptRow, 19) = X.getAttribute("ows_VariantName")
'        Sheets("config").Cells(EmptRow, 20) = X.getAttribute("ows_Currency")
'        Sheets("config").Cells(EmptRow, 21) = X.getAttribute("ows_FCUSD")
'    Next
'End With

'Profit Centers
'List = "ProfitCenters"
'
'request = "<?xml version='1.0' encoding='utf-8'?>" & _
'            "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'" & _
'            " xmlns:xsd='http://www.w3.org/2001/XMLSchema'" & _
'            " xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>" & _
'            " <soap:Body>" & _
'                "<GetListItems xmlns='http://schemas.microsoft.com/sharepoint/soap/'>" & _
'                "<listName>" & List & "</listName>" & _
'                        "<QueryOptions>" & _
'                            "<IncludeMandatoryColumns>FALSE</IncludeMandatoryColumns>" & _
'                        "</QueryOptions>" & _
'                "<rowLimit>50000</rowLimit>" & _
'                " </GetListItems>" & _
'            " </soap:Body>" & _
'            "</soap:Envelope>"
'
'With CreateObject("Microsoft.XMLHTTP")
'    .Open "Get", Url, False, "", ""
'    .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
'    .setRequestHeader "SOAPAction", "http://schemas.microsoft.com/sharepoint/soap/GetListItems"
'    .send request
'    xmlDoc.LoadXML (.responsetext)
'
'    For Each X In xmlDoc.getElementsByTagName("z:row")
'        EmptRow = FindLastRow(1, 11, 1, 0, "config")
'        Sheets("config").Cells(EmptRow, 11) = X.getAttribute("ows_ID")
'        Sheets("config").Cells(EmptRow, 12) = X.getAttribute("ows_CC")
'        Sheets("config").Cells(EmptRow, 13) = X.getAttribute("ows_PC")
'        Sheets("config").Cells(EmptRow, 14) = X.getAttribute("ows_Transmit")
'        Sheets("config").Cells(EmptRow, 15) = X.getAttribute("ows_GAAPindicator")
'    Next
'End With

'----------------------------------------------------------------
'assign variables
'----------------------------------------------------------------
CM_Note "reading the settings from the config sheet"
LastRow = FindLastRow(1, 18, 0, 0, "config")
ArrVar = Sheets("config").Range("R2", Sheets("config").Cells(LastRow, 21))
LastRow = FindLastRow(1, 12, 0, 0, "config")
ArrPC = Sheets("config").Range("L2", Sheets("config").Cells(LastRow, 15))
LastRow = FindLastRow(1, 4, 0, 0, "config")
ArrCPC = Sheets("config").Range("D2", Sheets("config").Cells(LastRow, 9))
LastRow = FindLastRow(1, 23, 0, 0, "config")
ArrInd = Sheets("config").Range("W2", Sheets("config").Cells(LastRow, 24))

'----------------------------------------------------------------
'select closing variant
'----------------------------------------------------------------
CM_Note "selecting the closing variant for this company code"
For i = 1 To UBound(ArrVar, 1)
    If CC = ArrVar(i, 1) Then
        Var = ArrVar(i, 2)
        Cur = ArrVar(i, 3)
        FCUSD = ArrVar(i, 4)
        Exit For
    End If
Next i

Sheets("config").Range("AA1") = CC

'----------------------------------------------------------------
'check if CC is opened
'----------------------------------------------------------------
CM_Note "checking the company code and period are open in SAP"
If CCOpened(sess, ArrPC) = False Then

    'update tracker
'    List = "ClosingTracker"
'    Date1 = Format(Now(), "yyyy-MM-ddTHH:mm:ssZ")
'
'    updates = "<Batch> <Method ID='1' Cmd='New'>" & _
'                "<Field Name='UName'>" & UCase(Environ("username")) & "</Field>" & _
'                "<Field Name='LogType'>" & "Closing reports - error" & "</Field>" & _
'                "<Field Name='CC'>" & CC & "</Field>" & _
'                "<Field Name='Periodx'>" & Month(LastDay) & "</Field>" & _
'                "<Field Name='Yearx'>" & Year(LastDay) & "</Field>" & _
'                "<Field Name='Success'>0</Field>" & _
'                "<Field Name='Timestamp'>" & Date1 & "</Field>" & _
'                "<Field Name='Comment'>" & "Company Code already closed" & "</Field>" & _
'                "</Method></Batch>"
'
'    Call spAddToList(updates, List)

    strError = "Company Code " & CC & " closed for " & Month(LastDay) & "/" & Year(LastDay) & "."
    With UF_Error
        .Lbl_Error.Caption = strError
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        .Show
    End With
    CM_Done
    Exit Sub
End If

'----------------------------------------------------------------
'select reports to be printed
'----------------------------------------------------------------
CM_Note "asking which reports to print"
If RunAgainXY = False And RunAgainXYAG = False Then
    Sheets("config").Range("AA3") = "No"
    Sheets("config").Range("AA4") = "No"
    Sheets("config").Range("AA5") = "No"
    Sheets("config").Range("AA6") = "No"
    Sheets("config").Range("AA7") = "No"
    Sheets("config").Range("AA8") = "No"
    Sheets("config").Range("AA9") = "No"
    Sheets("config").Range("AA10") = "No"
    Sheets("config").Range("AA11") = "No"
    Sheets("config").Range("AA12") = "No"
    Sheets("config").Range("AA13") = "No"
    Sheets("config").Range("AA14") = "No"
    Sheets("config").Range("AA15") = "No"
    If FCUSD = 0 And Cur <> "USD" Then
        Sheets("config").Range("AA16") = "No"
        Sheets("config").Range("AA17") = "No"
    Else
        Sheets("config").Range("AA16") = "NA"
        Sheets("config").Range("AA17") = "NA"
    End If
    Sheets("config").Range("AA18") = "No"
    Sheets("config").Range("AA19") = "No"
    Sheets("config").Range("AA20") = "No"
End If

'----------------------------------------------------------------
'check ZGLRME and AA02
'----------------------------------------------------------------
CM_Note "checking ZGLRME and AA02 for differences"
If RunAgainXY = False And RunAgainXYAG = False Then
    ZGLRMEErr = False
    If CheckZGLRME(sess, ArrInd, ArrPC, ArrCPC) = False Or CheckAA02(sess) = False Then
    
        If RunAgain = True Then
            Sheets("config").Range("B3").ClearContents
    
            strOption = "You run the reports again and there were pre-close errors. Are you sure you want to run closing routine for Company Code " & CC & "?"
            With UF_Option
                .Lbl_Option.Caption = strOption
                .StartUpPosition = 0
                .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
                .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
                .Show
            End With
            If Sheets("config").Range("B3") <> "Yes" Then
                CM_Done
                Exit Sub
            Else
                LastRow = FindLastRow(1, 2, 0, 0, "Errors")
                Set rng = Sheets("Errors").Range("A1", Sheets("Errors").Cells(LastRow, 30))
                Call ApplyBorders(0, 0, rng)
                With Sheets("Errors").Cells
                    .ClearContents
                    .ClearFormats
                    .Font.Bold = False
                    .Interior.Pattern = xlSolid
                    .Interior.PatternColorIndex = xlAutomatic
                    .Interior.ThemeColor = xlThemeColorDark1
                    .Interior.TintAndShade = 0
                    .Interior.PatternTintAndShade = 0
                End With
                Sheets("config").Range("AA2") = "Yes"
            End If
        Else
    
            Sheets("config").Range("AA2") = "Error"
        
            'update tracker
            List = "ClosingTracker"
            Date1 = Format(Now(), "yyyy-MM-ddTHH:mm:ssZ")
        
            strErr = ""
            If OtherF = True Then
                If strErr = "" Then strErr = "Other" Else strErr = strErr & ";Other"
            End If
            If BPCF = True Then
                If strErr = "" Then strErr = "BPC" Else strErr = strErr & ";BPC"
            End If
            If CPCF = True Then
                If strErr = "" Then strErr = "CPC" Else strErr = strErr & ";CPC"
            End If
            If BPCWDF = True Then
                If strErr = "" Then strErr = "BPCWD" Else strErr = strErr & ";BPCWD"
            End If
        
'            updates = "<Batch> <Method ID='1' Cmd='New'>" & _
'                        "<Field Name='UName'>" & UCase(Environ("username")) & "</Field>" & _
'                        "<Field Name='LogType'>" & "Closing reports - error" & "</Field>" & _
'                        "<Field Name='CC'>" & CC & "</Field>" & _
'                        "<Field Name='Periodx'>" & Month(LastDay) & "</Field>" & _
'                        "<Field Name='Yearx'>" & Year(LastDay) & "</Field>" & _
'                        "<Field Name='Success'>0</Field>" & _
'                        "<Field Name='Timestamp'>" & Date1 & "</Field>" & _
'                        "<Field Name='Comment'>" & "Pre-close errors: " & strErr & "</Field>" & _
'                        "</Method></Batch>"
'
'            Call spAddToList(updates, List)
        
            Sheets("Errors").Visible = True
            Sheets("Errors").Select
            Range("A1").Select
            strError = "There are errors which do not allow closing the entity. Verify tab 'Errors'."
            With UF_Error
                .Lbl_Error.Caption = strError
                .StartUpPosition = 0
                .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
                .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
                .Show
            End With
            CM_Done
            Exit Sub
        End If
    Else
        LastRow = FindLastRow(1, 2, 0, 0, "Errors")
        Set rng = Sheets("Errors").Range("A1", Sheets("Errors").Cells(LastRow, 30))
        Call ApplyBorders(0, 0, rng)
        With Sheets("Errors").Cells
            .ClearContents
            .ClearFormats
            .Font.Bold = False
            .Interior.Pattern = xlSolid
            .Interior.PatternColorIndex = xlAutomatic
            .Interior.ThemeColor = xlThemeColorDark1
            .Interior.TintAndShade = 0
            .Interior.PatternTintAndShade = 0
        End With
        Sheets("config").Range("AA2") = "Yes"
    End If
    printN = 12
End If

'----------------------------------------------------------------
'print ZGLRME
'----------------------------------------------------------------
CM_Note "printing ZGLRME"
If RunAgainXY = False And RunAgainXYAG = False Then
    Call Print_ZGLRME(sess, printN)
    Sheets("config").Range("AA3") = "Yes"
End If

'----------------------------------------------------------------
'print EIS4
'----------------------------------------------------------------
CM_Note "printing report group EIS4"
If RunAgainXY = False And RunAgainXYAG = False Then Sheets("config").Range("AA4") = Print_EIS4(sess, printN)

'----------------------------------------------------------------
'print GIS4
'----------------------------------------------------------------
CM_Note "printing report group GIS4"
If RunAgainXY = False And RunAgainXYAG = False Then Sheets("config").Range("AA10") = Print_GIS4(sess, printN)

'----------------------------------------------------------------
'print and check ZGE132
'----------------------------------------------------------------
CM_Note "printing and checking ZGE132"
If RunAgainXY = False And RunAgainXYAG = False Then
    BPCL = False
    CPCL = False
    BPCG = False
    CPCG = False
    ZGEISF = False
    ZGGISF = False
    If Print_ZGE132(sess, ArrPC, ArrInd) = False Then
        'update tracker
        strErr = ""
    
        List = "ClosingTracker"
        Date1 = Format(Now(), "yyyy-MM-ddTHH:mm:ssZ")
    
        If BPCL = True Then
            If strErr = "" Then strErr = "BPCL" Else strErr = strErr & ";BPCL"
        End If
        If BPCG = True Then
            If strErr = "" Then strErr = "BPCG" Else strErr = strErr & ";BPCG"
        End If
        If CPCL = True Then
            If strErr = "" Then strErr = "CPCL" Else strErr = strErr & ";CPCL"
        End If
        If CPCG = True Then
            If strErr = "" Then strErr = "CPCG" Else strErr = strErr & ";CPCG"
        End If
        If ZGEISF = True Then
            If strErr = "" Then strErr = "ZGEIS" Else strErr = strErr & ";ZGEIS"
        End If
        If ZGGISF = True Then
            If strErr = "" Then strErr = "ZGGIS" Else strErr = strErr & ";ZGGIS"
        End If
    
'        updates = "<Batch> <Method ID='1' Cmd='New'>" & _
'                    "<Field Name='UName'>" & UCase(Environ("username")) & "</Field>" & _
'                    "<Field Name='LogType'>" & "Closing reports - error" & "</Field>" & _
'                    "<Field Name='CC'>" & CC & "</Field>" & _
'                    "<Field Name='Periodx'>" & Month(LastDay) & "</Field>" & _
'                    "<Field Name='Yearx'>" & Year(LastDay) & "</Field>" & _
'                    "<Field Name='Success'>0</Field>" & _
'                    "<Field Name='Timestamp'>" & Date1 & "</Field>" & _
'                    "<Field Name='Comment'>" & "ZGE132: " & strErr & "</Field>" & _
'                    "</Method></Batch>"
'
'        Call spAddToList(updates, List)
    
        'info
        strError = ""
        If BPCL = True Or BPCG = True Or CPCL = True Or CPCG = True Then
            strError = "There are errors which do not allow closing the entity caused by the postings done during running the closing routine. Verify tab 'Errors'."
        End If
        If ZGEISF = True Or ZGGISF = True Then
            If strError = "" Then
                strError = "Profit and Loss Clearing entry is not equal to EIS4/GIS4 due to postings done during running the closing routine."
            Else
                strError = strError & vbNewLine & vbNewLine & "Profit and Loss Clearing entry is not equal to EIS4/GIS4 due to postings done during running the closing routine."
            End If
        End If
        If strError <> "" Then
            With UF_Error
                .Lbl_Error.Caption = strError
                .StartUpPosition = 0
                .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
                .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
                .Show
            End With
            CM_Done
            Exit Sub
        End If
    End If
End If

'----------------------------------------------------------------
'post ZGE132
'----------------------------------------------------------------
CM_Note "posting the ZGE132 entries in SAP"
If RunAgainXY = False And RunAgainXYAG = False Then
    If Round(Sheets("config").Range("AA4"), 2) <> 0 Then
        Call Post_ZGE132(sess)
    ElseIf Round(Sheets("config").Range("AA10"), 2) <> 0 Then
        Call Post_ZGE132GC(sess)
    End If
End If

'----------------------------------------------------------------
'check ZGE132 after posting
'----------------------------------------------------------------
CM_Note "checking ZGE132 after posting"
If RunAgainXYAG = False Then
    BG = False
    BL = False
    CPCL = False
    CPCG = False
    If Check_ZGE132AP(sess) = False Then
        'check errors in SM35
        
        If Cur = "USD" Then
            sessN = 1
        Else
            sessN = 2
        End If
        Call GetSM35(sess, sessN)
        If sessN = 1 Then
            Sheets("config").Range("AA12") = "No posting"
        End If
        
        strErr = ""
        List = "ClosingTracker"
        Date1 = Format(Now(), "yyyy-MM-ddTHH:mm:ssZ")
        
        If Sheets("config").Range("AA12") = "XY" Then
            strErr = "XY"
        Else
            If BL = True Then
                If strErr = "" Then strErr = "BL" Else strErr = strErr & ";BL"
            End If
            If BG = True Then
                If strErr = "" Then strErr = "BG" Else strErr = strErr & ";BG"
            End If
            If CPCL = True Then
                If strErr = "" Then strErr = "CPCL" Else strErr = strErr & ";CPCL"
            End If
            If CPCG = True Then
                If strErr = "" Then strErr = "CPCG" Else strErr = strErr & ";CPCG"
            End If
        End If
        
        'update tracker
        
'        updates = "<Batch> <Method ID='1' Cmd='New'>" & _
'                    "<Field Name='UName'>" & UCase(Environ("username")) & "</Field>" & _
'                    "<Field Name='LogType'>" & "Closing reports - error" & "</Field>" & _
'                    "<Field Name='CC'>" & CC & "</Field>" & _
'                    "<Field Name='Periodx'>" & Month(LastDay) & "</Field>" & _
'                    "<Field Name='Yearx'>" & Year(LastDay) & "</Field>" & _
'                    "<Field Name='Success'>0</Field>" & _
'                    "<Field Name='Timestamp'>" & Date1 & "</Field>" & _
'                    "<Field Name='Comment'>" & "ZGE132: " & strErr & "</Field>" & _
'                    "</Method></Batch>"
'
'        Call spAddToList(updates, List)
    
        'info
        strError = ""
        If strErr = "XY" Then
            strError = "XY error. Send e-mail to GFIM team and run macro again."
        Else
            If BL = True Or BG = True Then
                strError = "Profit and Loss Clearing not completed due to postings done during running the closing routine."
            End If
            If CPCL = True Or CPCG = True Then
                If strError = "" Then
                    strError = "Balance Control Entry not completed due to postings done during running the closing routine."
                Else
                    strError = strError & vbNewLine & vbNewLine & "Balance Control Entry not completed due to postings done during running the closing routine."
                End If
            End If
        End If
    
        If strError <> "" Then
            With UF_Error
                .Lbl_Error.Caption = strError
                .StartUpPosition = 0
                .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
                .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
                .Show
            End With
            CM_Done
            Exit Sub
        End If
    Else
        Sheets("config").Range("AA8") = "Yes"
        Sheets("config").Range("AA14") = "Yes"
    End If
End If

'----------------------------------------------------------------
'get document numbers from SM35
'----------------------------------------------------------------
CM_Note "reading the document numbers from SM35"
If RunAgainXYAG = False Then
    If Round(Sheets("config").Range("AA4"), 2) <> 0 Then
        If Sheets("config").Range("AA12") = "No posting" Then
            sessN = 1
        Else
            If Cur = "USD" Then
                sessN = 1
            Else
                sessN = 2
            End If
        End If
        Call GetSM35(sess, sessN)
        If sessN = 1 Then
            Sheets("config").Range("AA12") = "No posting"
        End If
    ElseIf Round(Sheets("config").Range("AA10"), 2) <> 0 Then
        Sheets("config").Range("AA6") = "No posting"
        sessN = 1
        Call GetSM35(sess, sessN)
    Else
        Sheets("config").Range("AA6") = "No posting"
        Sheets("config").Range("AA12") = "No posting"
    End If
End If

'----------------------------------------------------------------
'print documents in ZGR215
'----------------------------------------------------------------
CM_Note "printing the posted documents in ZGR215"
If RunAgainXYAG = False Then
    If Round(Sheets("config").Range("AA4"), 2) <> 0 Then
        Call Print_ZGR215(sess)
        Sheets("config").Range("AA7") = "Yes"
        If Sheets("config").Range("AA12") <> "No posting" Then Sheets("config").Range("AA13") = "Yes"
    End If
End If

'----------------------------------------------------------------
'print EIS4
'----------------------------------------------------------------
CM_Note "printing report group EIS4 again (after posting)"
If RunAgainXYAG = False Then
    Sheets("config").Range("AA9") = Print_EIS4(sess, printN)
    If Round(Sheets("config").Range("AA9"), 2) <> 0 Then
    
        List = "ClosingTracker"
        Date1 = Format(Now(), "yyyy-MM-ddTHH:mm:ssZ")
        
'        updates = "<Batch> <Method ID='1' Cmd='New'>" & _
'                    "<Field Name='UName'>" & UCase(Environ("username")) & "</Field>" & _
'                    "<Field Name='LogType'>" & "Closing reports - error" & "</Field>" & _
'                    "<Field Name='CC'>" & CC & "</Field>" & _
'                    "<Field Name='Periodx'>" & Month(LastDay) & "</Field>" & _
'                    "<Field Name='Yearx'>" & Year(LastDay) & "</Field>" & _
'                    "<Field Name='Success'>0</Field>" & _
'                    "<Field Name='Timestamp'>" & Date1 & "</Field>" & _
'                    "<Field Name='Comment'>" & "EIS4" & "</Field>" & _
'                    "</Method></Batch>"
'
'        Call spAddToList(updates, List)
    
        strError = "EIS4 not zeroed due to postings done during running the closing routine."
        With UF_Error
            .Lbl_Error.Caption = strError
            .StartUpPosition = 0
            .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
            .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
            .Show
        End With
        CM_Done
        Exit Sub
    End If
End If

'----------------------------------------------------------------
'print GIS4
'----------------------------------------------------------------
CM_Note "printing report group GIS4 again (after posting)"
If RunAgainXYAG = False Then
    Sheets("config").Range("AA15") = Print_GIS4(sess, printN)
    If Round(Sheets("config").Range("AA15"), 2) <> 0 Then
    
        List = "ClosingTracker"
        Date1 = Format(Now(), "yyyy-MM-ddTHH:mm:ssZ")
        
'        updates = "<Batch> <Method ID='1' Cmd='New'>" & _
'                    "<Field Name='UName'>" & UCase(Environ("username")) & "</Field>" & _
'                    "<Field Name='LogType'>" & "Closing reports - error" & "</Field>" & _
'                    "<Field Name='CC'>" & CC & "</Field>" & _
'                    "<Field Name='Periodx'>" & Month(LastDay) & "</Field>" & _
'                    "<Field Name='Yearx'>" & Year(LastDay) & "</Field>" & _
'                    "<Field Name='Success'>0</Field>" & _
'                    "<Field Name='Timestamp'>" & Date1 & "</Field>" & _
'                    "<Field Name='Comment'>" & "GIS4" & "</Field>" & _
'                    "</Method></Batch>"
'
'        Call spAddToList(updates, List)
    
        strError = "GIS4 not zeroed due to postings done during running the closing routine."
        With UF_Error
            .Lbl_Error.Caption = strError
            .StartUpPosition = 0
            .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
            .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
            .Show
        End With
        CM_Done
        Exit Sub
    End If
End If

'----------------------------------------------------------------
'ZGLGWUL
'----------------------------------------------------------------
CM_Note "running ZGLGWUL"
If RunAgainXYAG = False Then
    If Sheets("config").Range("AA16") <> "NA" Then
        Call Post_ZGLGWUL(sess)
        
        If Sheets("config").Range("AA16") = "Rounding Difference Too Large. Check the report." Then
        
            List = "ClosingTracker"
            Date1 = Format(Now(), "yyyy-MM-ddTHH:mm:ssZ")
            
'            updates = "<Batch> <Method ID='1' Cmd='New'>" & _
'                        "<Field Name='UName'>" & UCase(Environ("username")) & "</Field>" & _
'                        "<Field Name='LogType'>" & "Closing reports - error" & "</Field>" & _
'                        "<Field Name='CC'>" & CC & "</Field>" & _
'                        "<Field Name='Periodx'>" & Month(LastDay) & "</Field>" & _
'                        "<Field Name='Yearx'>" & Year(LastDay) & "</Field>" & _
'                        "<Field Name='Success'>0</Field>" & _
'                        "<Field Name='Timestamp'>" & Date1 & "</Field>" & _
'                        "<Field Name='Comment'>" & "ZGLGWUL" & "</Field>" & _
'                        "</Method></Batch>"
'
'            Call spAddToList(updates, List)
            
            strError = "ZGLGWUL could not be run due to postings done during running the closing routine."
            With UF_Error
                .Lbl_Error.Caption = strError
                .StartUpPosition = 0
                .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
                .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
                .Show
            End With
            CM_Done
            Exit Sub
        End If
    End If
End If

'----------------------------------------------------------------
'after ZGLGWUL
'----------------------------------------------------------------
CM_Note "checking the result of ZGLGWUL"
If RunAgainXYAG = False Then
    If Sheets("config").Range("AA17") <> "NA" Then
        Call Post_ZGE132AG(sess)
        If Left(Sheets("config").Range("AA17"), 2) = "XY" Then
            CM_Done
            Exit Sub
        End If
    End If
Else
    Call GetSM35AG(sess)
End If

'----------------------------------------------------------------
'Print GTB1
'----------------------------------------------------------------
CM_Note "printing report group GTB1"
If Sheets("config").Range("AA18") <> "NA" Then
    Call Print_GTB1(sess, printN)
    If Sheets("config").Range("AA16") <> "NA" Then
        If Round(Sheets("config").Range("AA18"), 2) - Round(Sheets("config").Range("AA16"), 2) <> 0 Then
            List = "ClosingTracker"
            Date1 = Format(Now(), "yyyy-MM-ddTHH:mm:ssZ")
            
'            updates = "<Batch> <Method ID='1' Cmd='New'>" & _
'                        "<Field Name='UName'>" & UCase(Environ("username")) & "</Field>" & _
'                        "<Field Name='LogType'>" & "Closing reports - error" & "</Field>" & _
'                        "<Field Name='CC'>" & CC & "</Field>" & _
'                        "<Field Name='Periodx'>" & Month(LastDay) & "</Field>" & _
'                        "<Field Name='Yearx'>" & Year(LastDay) & "</Field>" & _
'                        "<Field Name='Success'>0</Field>" & _
'                        "<Field Name='Timestamp'>" & Date1 & "</Field>" & _
'                        "<Field Name='Comment'>" & "GTB1" & "</Field>" & _
'                        "</Method></Batch>"
'
'            Call spAddToList(updates, List)
            
            If Sheets("Errors").Range("A10") = "" Then
                Sheets("Errors").Range("A10") = "Error in GTB1 - amount from ZGLGWUL transaction does not equal the amount in GTB1 report for account 44400200."
                Sheets("Errors").Range("A10").Font.Bold = True
            Else
                EmptRow = FindLastRow(1, 1, 2, 0, "Errors")
                EmptRow1 = FindLastRow(1, 2, 2, 0, "Errors")
                If EmptRow > EmptRow1 Then
                    Sheets("Errors").Cells(EmptRow, 1) = "Error in GTB1 - amount from ZGLGWUL transaction does not equal the amount in GTB1 report for account 44400200."
                    Sheets("Errors").Range(EmptRow, 1).Font.Bold = True
                Else
                    Sheets("Errors").Cells(EmptRow1, 1) = "Error in GTB1 - amount from ZGLGWUL transaction does not equal the amount in GTB1 report for account 44400200."
                    Sheets("Errors").Range(EmptRow1, 1).Font.Bold = True
                End If
            End If
        End If
    End If
End If

'----------------------------------------------------------------
'Print ZGE1174
'----------------------------------------------------------------
CM_Note "printing ZGE1174"
If Sheets("config").Range("AA19") <> "NA" Then
    Call Print_ZGE1174(sess, printN)
End If

'----------------------------------------------------------------
'check ZGLRME
'----------------------------------------------------------------
CM_Note "checking ZGLRME again"
ZGLRMEErr = False
If CheckZGLRME(sess, ArrInd, ArrPC, ArrCPC) = False Or CheckAA02(sess) = False Then

    'update tracker
    List = "ClosingTracker"
    Date1 = Format(Now(), "yyyy-MM-ddTHH:mm:ssZ")

    strErr = ""
    If OtherF = True Then
        If strErr = "" Then strErr = "Other" Else strErr = strErr & ";Other"
    End If
    If BPCF = True Then
        If strErr = "" Then strErr = "BPC" Else strErr = strErr & ";BPC"
    End If
    If CPCF = True Then
        If strErr = "" Then strErr = "CPC" Else strErr = strErr & ";CPC"
    End If
    If BPCWDF = True Then
        If strErr = "" Then strErr = "BPCWD" Else strErr = strErr & ";BPCWD"
    End If

'    updates = "<Batch> <Method ID='1' Cmd='New'>" & _
'                "<Field Name='UName'>" & UCase(Environ("username")) & "</Field>" & _
'                "<Field Name='LogType'>" & "Closing reports - error" & "</Field>" & _
'                "<Field Name='CC'>" & CC & "</Field>" & _
'                "<Field Name='Periodx'>" & Month(LastDay) & "</Field>" & _
'                "<Field Name='Yearx'>" & Year(LastDay) & "</Field>" & _
'                "<Field Name='Success'>0</Field>" & _
'                "<Field Name='Timestamp'>" & Date1 & "</Field>" & _
'                "<Field Name='Comment'>" & "Post-close errors: " & strErr & "</Field>" & _
'                "</Method></Batch>"
'
'    Call spAddToList(updates, List)
Else
    LastRow = FindLastRow(1, 2, 0, 0, "Errors")
    Set rng = Sheets("Errors").Range("A1", Sheets("Errors").Cells(LastRow, 30))
    Call ApplyBorders(0, 0, rng)
    With Sheets("Errors").Cells
        .ClearContents
        .ClearFormats
        .Font.Bold = False
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.ThemeColor = xlThemeColorDark1
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
    End With
    Sheets("config").Range("AA20") = "Yes"
End If

'----------------------------------------------------------------
'print ZGLRME
'----------------------------------------------------------------
CM_Note "printing the final ZGLRME"
Call Print_ZGLRME(sess, printN)

'----------------------------------------------------------------
'combine PDFs
'----------------------------------------------------------------
CM_Note "merging all the PDFs into the final report pack"
Call CombinePDF(printN)

'----------------------------------------------------------------
'check final errors
'----------------------------------------------------------------
CM_Note "checking for any remaining errors"
LastRow = FindLastRow(1, 1, 0, 0, "Errors")
If LastRow > 1 Then
    Sheets("Errors").Visible = 1
    Sheets("Errors").Select
    strError = "There are some post-close errors, please verify the tab 'Errors'."
    With UF_Error
        .Lbl_Error.Caption = strError
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        .Show
    End With
    
'    List = "ClosingTracker"
'    Date1 = Format(Now(), "yyyy-MM-ddTHH:mm:ssZ")
'    updates = "<Batch> <Method ID='1' Cmd='New'>" & _
'                "<Field Name='UName'>" & UCase(Environ("username")) & "</Field>" & _
'                "<Field Name='LogType'>" & "Closing reports - finish" & "</Field>" & _
'                "<Field Name='CC'>" & CC & "</Field>" & _
'                "<Field Name='Periodx'>" & Month(LastDay) & "</Field>" & _
'                "<Field Name='Yearx'>" & Year(LastDay) & "</Field>" & _
'                "<Field Name='Success'>1</Field>" & _
'                "<Field Name='Timestamp'>" & Date1 & "</Field>" & _
'                "<Field Name='Comment'>" & "Finished with post-close errors." & strErr & "</Field>" & _
'                "</Method></Batch>"
'
'    Call spAddToList(updates, List)
    
Else
    Sheets("Errors").Visible = xlVeryHidden
    Sheets("START").Select
    strInfo = "Closing routine successful. Report printed to folder:" & vbNewLine & FPathReport & vbNewLine & "Name of the file:" & vbNewLine & FName
    With UF_Info
        .Lbl_Info.Caption = strInfo
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        .Show
    End With
    
'    List = "ClosingTracker"
'    Date1 = Format(Now(), "yyyy-MM-ddTHH:mm:ssZ")
'    updates = "<Batch> <Method ID='1' Cmd='New'>" & _
'                "<Field Name='UName'>" & UCase(Environ("username")) & "</Field>" & _
'                "<Field Name='LogType'>" & "Closing reports - finish" & "</Field>" & _
'                "<Field Name='CC'>" & CC & "</Field>" & _
'                "<Field Name='Periodx'>" & Month(LastDay) & "</Field>" & _
'                "<Field Name='Yearx'>" & Year(LastDay) & "</Field>" & _
'                "<Field Name='Success'>1</Field>" & _
'                "<Field Name='Timestamp'>" & Date1 & "</Field>" & _
'                "<Field Name='Comment'>" & "Finished with success." & strErr & "</Field>" & _
'                "</Method></Batch>"
'
'    Call spAddToList(updates, List)
    
End If


'V4-CIO: one handler for the whole close. RunClosing turns error handling on
'nowhere else, so anything that fails here - or in any routine it calls -
'lands below and is explained in plain language instead of showing a bare
'"Run-time error 13" dialog with no idea what the macro was doing.
CM_Done
Exit Sub

CM_Fail:
    CM_Explain Err.Number, Err.Description
    CM_Done

End Sub
Function CCOpened(sess As Object, ArrPC)

'check if CC is opened in T001B (1st period)
With sess
    .findById("wnd[0]").maximize
    .findById("wnd[0]/tbar[0]/okcd").Text = "/nse16"
    .findById("wnd[0]").sendVKey 0
    .findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").Text = "t001b"
    .findById("wnd[0]").sendVKey 0
    On Error Resume Next
    .findById("wnd[1]/tbar[0]/btn[0]").press
    On Error GoTo 0
    .findById("wnd[0]/mbar/menu[3]/menu[2]").Select
    .findById("wnd[1]/tbar[0]/btn[14]").press
    
    k = 3
    Call SAPSelectFields(sess, k)
    
    .findById("wnd[1]/tbar[0]/btn[0]").press
    .findById("wnd[0]/usr/ctxtI1-LOW").Text = Var
    .findById("wnd[0]/usr/ctxtI2-LOW").Text = "s"
    .findById("wnd[0]/usr/txtI3-LOW").Text = Yearx
    .findById("wnd[0]/usr/txtI4-LOW").Text = Monthx
    .findById("wnd[0]/tbar[1]/btn[8]").press

    TName = .ActiveWindow.Text
    If InStr(1, TName, "Hits") > 0 Or InStr(1, TName, "Select Entries") > 0 Or InStr(1, TName, "Table Content") > 0 Then
        CCOpened = True
    Else
        CCOpened = False
    End If

    .findById("wnd[0]/tbar[0]/btn[3]").press
    .findById("wnd[0]/tbar[0]/btn[3]").press
    .findById("wnd[0]/tbar[0]/btn[3]").press

    'check if CC is opened in T001b (2nd period)
    If CCOpened = False Then
        .findById("wnd[0]").maximize
        .findById("wnd[0]/tbar[0]/okcd").Text = "/nse16"
        .findById("wnd[0]").sendVKey 0
        .findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").Text = "t001b"
        .findById("wnd[0]").sendVKey 0
        On Error Resume Next
        .findById("wnd[1]/tbar[0]/btn[0]").press
        On Error GoTo 0
        .findById("wnd[0]/mbar/menu[3]/menu[2]").Select
        .findById("wnd[1]/tbar[0]/btn[14]").press
        
        k = 3
        Call SAPSelectFields(sess, k)
        
        .findById("wnd[1]/tbar[0]/btn[0]").press
        .findById("wnd[0]/usr/ctxtI1-LOW").Text = Var
        .findById("wnd[0]/usr/ctxtI2-LOW").Text = "s"
        .findById("wnd[0]/usr/txtI5-LOW").Text = Yearx
        .findById("wnd[0]/usr/txtI6-LOW").Text = Monthx
        .findById("wnd[0]/tbar[1]/btn[8]").press
        
        TName = .ActiveWindow.Text
        If InStr(1, TName, "Hit") > 0 Or InStr(1, TName, "Select Entries") > 0 Or InStr(1, TName, "Table Content") > 0 Then
            CCOpened = True
        ElseIf Left(TName, 12) = "Data Browser" Then
            CCOpened = False
        End If
        
        .findById("wnd[0]/tbar[0]/btn[3]").press
        .findById("wnd[0]/tbar[0]/btn[3]").press
        .findById("wnd[0]/tbar[0]/btn[3]").press
        
    End If
    
    'check if CC is opened in ZGXMIT if opened in T001B
    If CCOpened = True Then
    
        Set alloutput = New ADODB.Stream
        
        alloutput.Charset = "utf-8"
        alloutput.Open
        
        For i = 1 To UBound(ArrPC, 1)
            If ArrPC(i, 1) = CC Then
                alloutput.writetext (ArrPC(i, 2) & vbCrLf)
            End If
        Next i
                                            
        alloutput.SaveToFile FPath & "\PC.csv", 2
        alloutput.Close
    
        .findById("wnd[0]").maximize
        .findById("wnd[0]/tbar[0]/okcd").Text = "/nse16"
        .findById("wnd[0]").sendVKey 0
        .findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").Text = "ZGXMIT"
        .findById("wnd[0]").sendVKey 0
        On Error Resume Next
        .findById("wnd[1]/tbar[0]/btn[0]").press
        On Error GoTo 0
        .findById("wnd[0]/mbar/menu[3]/menu[2]").Select
        .findById("wnd[1]/tbar[0]/btn[14]").press
        
        k = 4
        Call SAPSelectFields(sess, k)
        
        .findById("wnd[1]/tbar[0]/btn[0]").press
        .findById("wnd[0]/usr/btn%_I1_%_APP_%-VALU_PUSH").press
        .findById("wnd[1]/tbar[0]/btn[23]").press
        .findById("wnd[2]/usr/ctxtDY_PATH").Text = FPath
        .findById("wnd[2]/usr/ctxtDY_FILENAME").Text = "PC.csv"
        .findById("wnd[2]/tbar[0]/btn[0]").press
        .findById("wnd[1]/tbar[0]/btn[8]").press
        .findById("wnd[0]/usr/txtI2-LOW").Text = Yearx
        .findById("wnd[0]/usr/txtI3-LOW").Text = Monthx
        
        .findById("wnd[0]/mbar/menu[3]/menu[1]").Select
        .findById("wnd[1]/usr/tabsG_TABSTRIP/tabp0400/ssubTOOLAREA:SAPLWB_CUSTOMIZING:0400/ctxtRSEUMOD-TBLISTBR").Text = "250"
        .findById("wnd[1]/usr/tabsG_TABSTRIP/tabp0400/ssubTOOLAREA:SAPLWB_CUSTOMIZING:0400/txtRSEUMOD-TBMAXSEL").Text = "999999999"
        .findById("wnd[1]/usr/tabsG_TABSTRIP/tabp0400/ssubTOOLAREA:SAPLWB_CUSTOMIZING:0400/radRSEUMOD-TBALV_GRID").Select
        .findById("wnd[1]/usr/tabsG_TABSTRIP/tabp0400/ssubTOOLAREA:SAPLWB_CUSTOMIZING:0400/radSEUCUSTOM-FIELDNAME").Select
        .findById("wnd[1]/tbar[0]/btn[0]").press
        .findById("wnd[0]/mbar/menu[3]/menu[0]/menu[1]").Select
        .findById("wnd[1]/tbar[0]/btn[14]").press
        
        LastRow = FindLastRow(1, 9, 0, 0, "SAP config")
        j = 1
        Do Until LastRow = j
            Set Area = .findById("wnd[1]/usr")
            Set Children = Area.Children()
            For i = 0 To Children.Count() - 1
                Set obj = Children(CInt(i))
                If obj.Type = "GuiLabel" And obj.Text <> "" Then
                    w = 2
                    Do Until Sheets("SAP config").Cells(w, 9) = ""
                        FieldName = Sheets("SAP config").Cells(w, 9)
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
            If LastRow > j Then .findById("wnd[1]").sendVKey 82
        Loop
    
        .findById("wnd[1]/tbar[0]/btn[6]").press
        .findById("wnd[0]/tbar[1]/btn[8]").press
        
        TName = .ActiveWindow.Text
        
        If InStr(1, TName, "Hits") = 0 And InStr(1, TName, "Select Entries") = 0 And InStr(1, TName, "Table Content") = 0 Then
            CCOpened = True
            .findById("wnd[0]/mbar/menu[3]/menu[1]").Select
            .findById("wnd[1]/usr/tabsG_TABSTRIP/tabp0400/ssubTOOLAREA:SAPLWB_CUSTOMIZING:0400/radSEUCUSTOM-FIELDTEXT").Select
            .findById("wnd[1]").sendVKey 0
        Else
            .findById("wnd[0]/tbar[1]/btn[32]").press
    
            VR = .findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").visiblerowcount
            Do Until VR = 0
                .findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "0-" & VR - 1
                .findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press
                VR = .findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").visiblerowcount
            Loop
            
            .findById("wnd[1]/tbar[0]/btn[0]").press
            .findById("wnd[0]/tbar[1]/btn[45]").press
            .findById("wnd[1]").sendVKey 0
            .findById("wnd[1]/usr/ctxtDY_PATH").Text = FPath
            .findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "closestatus.txt"
            .findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").Text = "4120"
            .findById("wnd[1]/tbar[0]/btn[11]").press
            
            Set strim = New ADODB.Stream
            strim.Charset = "utf-8"
            strim.Open
            
            Call CreateArray(FPath & "closestatus.txt")
            strim.LoadFromFile (FPath & "closestatus.txt")
            
            Do Until strim.EOS
                Dim line As String, tablica() As String
                line = strim.ReadText(-2)
            
                If Left(Trim(line), 1) = "|" And Trim(getLineData(line, FirstColumn, 1)) <> FirstColumn Then
                    
                    If getLineData(line, "CLOSE_STAT", 1) = "Z" Then
                        CCOpened = False
                    End If
                    
                End If
            Loop
            
        End If
        
        .findById("wnd[0]/tbar[0]/btn[3]").press
        .findById("wnd[0]/tbar[0]/btn[3]").press
        .findById("wnd[0]/tbar[0]/btn[3]").press
        
        Kill FPath & "PC.csv"
        
    End If
End With

End Function
Function CheckZGLRME(sess As Object, ArrInd, ArrPC, ArrCPC)

OtherF = False
BPCF = False
CPCF = False

With sess
    .findById("wnd[0]/tbar[0]/okcd").Text = "/nzglrme"
    .findById("wnd[0]").sendVKey 0
    .findById("wnd[0]/usr/chkP_XMIT").Selected = False
    .findById("wnd[0]/usr/ctxtS_BUKRS-LOW").Text = CC
    .findById("wnd[0]/usr/txtP_MONAT").Text = Monthx
    .findById("wnd[0]/usr/txtP_GJAHR").Text = Yearx
    .findById("wnd[0]/usr/ctxtP_VARID").Text = "/default"
    .findById("wnd[0]/usr/ctxtP_VARIE").Text = "/closing"
    .findById("wnd[0]/tbar[1]/btn[8]").press
    .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").pressToolbarContextButton "&MB_FILTER"
    .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").selectContextMenuItem "&DELETE_FILTER"
    .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").setCurrentCell -1, "SACCT"
    .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").selectColumn "SACCT"
    .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").pressToolbarButton "&MB_FILTER"
    On Error Resume Next
    .findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
    If Err.Number <> 0 Then
        On Error GoTo 0
        .findById("wnd[1]/tbar[0]/btn[0]").press
        .findById("wnd[0]/tbar[0]/btn[3]").press
        .findById("wnd[0]/tbar[0]/btn[3]").press
        CheckZGLRME = True
        Exit Function
    End If
    On Error GoTo 0
    .findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV").Select
    .findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").Text = "85500080"
    .findById("wnd[2]").sendVKey 8
    .findById("wnd[1]").sendVKey 0
    .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").setCurrentCell -1, "MESSNUM"
    .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").selectColumn "MESSNUM"
    .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").pressToolbarButton "&MB_FILTER"
    .findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN002_%_APP_%-VALU_PUSH").press
    .findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV").Select
    .findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").Text = "E10"
    .findById("wnd[2]/tbar[0]/btn[8]").press
    .findById("wnd[1]/tbar[0]/btn[0]").press
    .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").pressToolbarContextButton "&MB_VARIANT"
    .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").pressToolbarContextButton "&MB_VIEW"
    .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").selectContextMenuItem "&PRINT_BACK_PREVIEW"
    .findById("wnd[0]/mbar/menu[3]/menu[6]/menu[0]").Select
    .findById("wnd[0]/tbar[1]/btn[45]").press
    .findById("wnd[1]/tbar[0]/btn[0]").press
    .findById("wnd[1]/usr/ctxtDY_PATH").Text = FPath
    .findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "zglrme.txt"
    .findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").Text = "4120"
    .findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").SetFocus
    .findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").caretPosition = 4
    .findById("wnd[1]/tbar[0]/btn[11]").press
    .findById("wnd[0]/tbar[0]/btn[3]").press
    .findById("wnd[0]/tbar[0]/btn[3]").press
    .findById("wnd[0]/tbar[0]/btn[3]").press
End With

Set strim = New ADODB.Stream
strim.Charset = "utf-8"
strim.Open

Call CreateArray(FPath & "zglrme.txt")
strim.LoadFromFile (FPath & "zglrme.txt")

Do Until strim.EOS
    Dim line As String, tablica() As String
    line = strim.ReadText(-2)
    
    If Left(line, 6) = "| List" Then
        CheckZGLRME = True
        Exit Function
    Else
    
        If Left(Trim(line), 1) = "|" And Trim(getLineData(line, FirstColumn, 1)) <> FirstColumn And Trim(getLineData(line, "Status", 1)) <> "*" Then
            EmptRow = FindLastRow(1, 1, 1, 0, "ZGLRME")
            k = 1
            Do Until Sheets("ZGLRME").Cells(1, k) = ""
                Sheets("ZGLRME").Cells(EmptRow, k) = getLineData(line, Sheets("ZGLRME").Cells(1, k), 1)
                k = k + 1
            Loop
        End If
    End If
Loop

'assign transmit and GAAP indicator
LastRow = FindLastRow(1, 1, 0, 0, "ZGLRME")
ArrZGL = Sheets("ZGLRME").Range("A2", Sheets("ZGLRME").Cells(LastRow, 11))

For i = 1 To UBound(ArrZGL, 1)
    'V4-CIO FIX: SAP writes amounts in the SAP user's decimal notation, which
    'need not match this PC's Windows regional settings; the old implicit
    'conversion then raised "Run-time error 13: Type mismatch". CM_Amount reads
    'both conventions and the trailing minus, and explains itself if it cannot.
    ArrZGL(i, 5) = CM_Amount(ArrZGL(i, 5), i, "reading the amounts from the ZGLRME extract")
    For j = 1 To UBound(ArrPC, 1)
        If CStr(ArrZGL(i, 3)) = CStr(ArrPC(j, 2)) Then
            ArrZGL(i, 9) = ArrPC(j, 3)
            ArrZGL(i, 10) = ArrPC(j, 4)
            For k = 1 To UBound(ArrInd, 1)
                If CStr(ArrZGL(i, 10)) = CStr(ArrInd(k, 1)) Then
                    ArrZGL(i, 11) = ArrInd(k, 2)
                    Exit For
                End If
            Next k
            Exit For
        End If
    Next j
Next i

Sheets("ZGLRME").Range("A2", Sheets("ZGLRME").Cells(LastRow, 11)) = ArrZGL

'----------------------------------------------------------------
'other errors
'----------------------------------------------------------------

If Sheets("Errors").Range("A10") = "" Then
    Sheets("Errors").Range("A10") = "Other errors in ZGLRME transaction:"
    Sheets("Errors").Range("A10").Font.Bold = True
Else
    EmptRow = FindLastRow(1, 2, 2, 0, "Errors")
    EmptRow1 = FindLastRow(1, 1, 2, 0, "Errors")
    If EmptRow1 > EmptRow Then
        Sheets("Errors").Cells(EmptRow1, 1) = "Other errors in ZGLRME transaction:"
        Sheets("Errors").Cells(EmptRow1, 1).Font.Bold = True
    Else
        Sheets("Errors").Cells(EmptRow, 1) = "Other errors in ZGLRME transaction:"
        Sheets("Errors").Cells(EmptRow, 1).Font.Bold = True
    End If
End If
EmptRow = FindLastRow(1, 1, 1, 0, "Errors")
k = 2
Do Until Sheets("ZGLRME").Cells(1, k) = ""
    Sheets("Errors").Cells(EmptRow, k) = Sheets("ZGLRME").Cells(1, k)
    k = k + 1
Loop

l = 1
Dim ArrZGL2
LastRow = FindLastRow(1, 1, 0, 0, "ZGLRME")
ReDim ArrZGL2(LastRow, 10)
For i = 1 To UBound(ArrZGL, 1)
    If ArrZGL(i, 8) <> "Assets <> Liability/Equity" And ArrZGL(i, 8) <> "Income Statement Out of Balance" Then
        found = False
        For j = 1 To UBound(ArrZGL2, 1)
            If CStr(ArrZGL(i, 3)) = CStr(ArrZGL2(j, 2)) And ArrZGL(i, 4) = ArrZGL2(j, 3) And ArrZGL(i, 8) = ArrZGL2(j, 7) Then
                found = True
                ArrZGL2(j, 4) = ArrZGL2(j, 4) + ArrZGL(i, 5)
            End If
        Next j
        If found = False Then
            For k = 2 To UBound(ArrZGL, 2)
                ArrZGL2(l, k - 1) = ArrZGL(i, k)
            Next k
            l = l + 1
        End If
    End If
Next i

EmptRow = FindLastRow(1, 2, 1, 0, "Errors")
Sheets("Errors").Range(Sheets("Errors").Cells(EmptRow, 2), Sheets("Errors").Cells(EmptRow + UBound(ArrZGL2, 1) - 1, 11)) = ArrZGL2

LastRow = FindLastRow(1, 2, 0, 0, "Errors")
If Sheets("Errors").Cells(LastRow - 1, 1) = "Other errors in ZGLRME transaction:" Then 'if no errors
    Sheets("Errors").Cells(LastRow, 1).EntireRow.ClearContents
    Sheets("Errors").Cells(LastRow, 2) = "No errors."
Else 'if some errors
    OtherF = True
    FirstRow = FindLastRow(1, 1, 1, 0, "Errors")
    Set rng = Sheets("Errors").Range(Sheets("Errors").Cells(FirstRow, 2), Sheets("Errors").Cells(LastRow, 11))
    Call ApplyBorders(1, 1, rng)
    Call ApplyBorders(2, 2, rng)
    rng.HorizontalAlignment = xlCenter
    With Sheets("Errors").Range(Sheets("Errors").Cells(FirstRow, 2), Sheets("Errors").Cells(FirstRow, 11))
        .EntireColumn.AutoFit
        .Font.Bold = True
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.Color = 14071936
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
    End With
    Sheets("Errors").Range(Sheets("Errors").Cells(FirstRow + 1, 5), Sheets("Errors").Cells(LastRow, 5)).NumberFormat = "#,##0.00"
End If

'----------------------------------------------------------------
'blank PC errors
'----------------------------------------------------------------
BalL = 0
BalG = 0
For i = 1 To UBound(ArrZGL, 1)
    If ArrZGL(i, 8) = "Assets <> Liability/Equity" Or ArrZGL(i, 8) = "Income Statement Out of Balance" Then
        If ArrZGL(i, 4) = "L" Then
            BalL = BalL + Round(ArrZGL(i, 5), 2)
        ElseIf ArrZGL(i, 4) = "G" Then
            BalG = BalG + Round(ArrZGL(i, 5), 2)
        End If
    End If
Next i

EmptRow = FindLastRow(1, 2, 2, 0, "Errors")
Sheets("Errors").Cells(EmptRow, 1) = "Blank Profit Center in local currency:"
Sheets("Errors").Cells(EmptRow, 1).Font.Bold = True

If Round(BalL, 2) <> 0 Then
    Sheets("Errors").Cells(EmptRow + 1, 2) = Format(BalL, "#,##0.00") & " " & Cur
    BPCF = True
Else
    Sheets("Errors").Cells(EmptRow + 1, 2) = "No errors."
End If

EmptRow = FindLastRow(1, 2, 2, 0, "Errors")
Sheets("Errors").Cells(EmptRow, 1) = "Blank Profit Center in group currency:"
Sheets("Errors").Cells(EmptRow, 1).Font.Bold = True

If Round(BalG, 2) <> 0 Then
    Sheets("Errors").Cells(EmptRow + 1, 2) = Format(BalG, "#,##0.00") & " USD"
    BPCF = True
Else
    Sheets("Errors").Cells(EmptRow + 1, 2) = "No errors."
End If

'----------------------------------------------------------------
'cross PC errors
'----------------------------------------------------------------

EmptRow = FindLastRow(1, 2, 2, 0, "Errors")
Sheets("Errors").Cells(EmptRow, 1) = "Cross Profit Center errors:"
Sheets("Errors").Cells(EmptRow, 1).Font.Bold = True
Sheets("Errors").Cells(EmptRow + 1, 2) = "CoCode"
Sheets("Errors").Cells(EmptRow + 1, 3) = "Profit Ctr"
Sheets("Errors").Cells(EmptRow + 1, 4) = "Loc/Grp Cu"
Sheets("Errors").Cells(EmptRow + 1, 5) = "Amount"
Sheets("Errors").Cells(EmptRow + 1, 6) = "Transmit"
Sheets("Errors").Cells(EmptRow + 1, 7) = "GAAP ind"
Sheets("Errors").Cells(EmptRow + 1, 8) = "GAAP ind description"

For i = 1 To UBound(ArrCPC, 1)
    If CC = ArrCPC(i, 1) Then
        PB = ArrCPC(i, 2)
        GAAP = ArrCPC(i, 4)
        NA = ArrCPC(i, 5)
        Exit For
    End If
Next i

If PB = 0 And GAAP = 0 And NA = 0 Then
    Sheets("Errors").Cells(EmptRow + 1, 1).EntireRow.ClearContents
    Sheets("Errors").Cells(EmptRow + 1, 2) = "Cross Profit Center postings allowed in all cases."
Else
    LastRow = FindLastRow(1, 1, 0, 0, "ZGLRME")
    Dim ArrZGLS()
    ReDim ArrZGLS(LastRow, 7)
    
    'making subtotals by Profit Center
    l = 1
    For i = 1 To UBound(ArrZGL, 1)
        If ArrZGL(i, 8) = "Assets <> Liability/Equity" Or ArrZGL(i, 8) = "Income Statement Out of Balance" Then
            PC1 = ArrZGL(i, 5)
            For j = 1 To UBound(ArrZGLS, 1)
                PC2 = ArrZGLS(j, 2)
                found = False
                If CStr(ArrZGL(i, 3)) = CStr(ArrZGLS(j, 2)) And ArrZGL(i, 4) = ArrZGLS(j, 3) Then
                    found = True
                    ArrZGLS(j, 4) = ArrZGLS(j, 4) + ArrZGL(i, 5)
                    Exit For
                End If
            Next j
            If found = False Then
                ArrZGLS(l, 1) = ArrZGL(i, 2)
                ArrZGLS(l, 2) = ArrZGL(i, 3)
                ArrZGLS(l, 3) = ArrZGL(i, 4)
                ArrZGLS(l, 4) = ArrZGL(i, 5)
                ArrZGLS(l, 5) = ArrZGL(i, 9)
                ArrZGLS(l, 6) = ArrZGL(i, 10)
                ArrZGLS(l, 7) = ArrZGL(i, 11)
                l = l + 1
            End If
        End If
    Next i
    
    Dim ArrZGLG()
    ReDim ArrZGLG(UBound(ArrZGLS, 1), 6)
    
    'making subtotals by GAAP indicator
    l = 1
    For i = 1 To UBound(ArrZGLS, 1)
        found = False
        If ArrZGLS(i, 1) <> "" Then
            For j = 1 To UBound(ArrZGLG, 1)
                If ArrZGLS(i, 6) = ArrZGLG(j, 5) And ArrZGLS(i, 3) = ArrZGLG(j, 2) Then
                    found = True
                    ArrZGLG(j, 3) = ArrZGLG(j, 3) + ArrZGLS(i, 4)
                    Exit For
                End If
            Next j
            If found = False Then
                ArrZGLG(l, 1) = ArrZGLS(i, 1)
                ArrZGLG(l, 2) = ArrZGLS(i, 3)
                ArrZGLG(l, 3) = ArrZGLS(i, 4)
                ArrZGLG(l, 4) = ArrZGLS(i, 5)
                ArrZGLG(l, 5) = ArrZGLS(i, 6)
                ArrZGLG(l, 6) = ArrZGLS(i, 7)
                l = l + 1
            End If
        End If
    Next i
    
    'checking cross-PC
    If PB = 1 Then
        'no cross allowed
        For i = 1 To UBound(ArrZGLS, 1)
            If Round(ArrZGLS(i, 4), 2) <> 0 Then
                EmptRow = FindLastRow(1, 2, 1, 0, "Errors")
                For k = 1 To UBound(ArrZGLS, 2)
                    Sheets("Errors").Cells(EmptRow, k + 1) = ArrZGLS(i, k)
                Next k
            End If
        Next i
    ElseIf GAAP = 0 Then
        'no cross allowed
        For i = 1 To UBound(ArrZGLS, 1)
            If Round(ArrZGLS(i, 4), 2) <> 0 Then
                EmptRow = FindLastRow(1, 2, 1, 0, "Errors")
                For k = 1 To UBound(ArrZGLS, 2)
                    Sheets("Errors").Cells(EmptRow, k + 1) = ArrZGLS(i, k)
                Next k
            End If
        Next i
    ElseIf NA = 1 Then
        'no cross allowed
        For i = 1 To UBound(ArrZGLS, 1)
            If Round(ArrZGLS(i, 4), 2) <> 0 Then
                EmptRow = FindLastRow(1, 2, 1, 0, "Errors")
                For k = 1 To UBound(ArrZGLS, 2)
                    Sheets("Errors").Cells(EmptRow, k + 1) = ArrZGLS(i, k)
                Next k
            End If
        Next i
    ElseIf GAAP = 1 Then
        'cross within one GAAP ind allowed
        For i = 1 To UBound(ArrZGLG, 1)
            If Round(ArrZGLG(i, 3), 2) <> 0 Then
                For j = 1 To UBound(ArrZGLS, 1)
                    If Round(ArrZGLS(j, 4), 2) <> 0 Then
                        EmptRow = FindLastRow(1, 2, 1, 0, "Errors")
                        For k = 1 To UBound(ArrZGLS, 2)
                            Sheets("Errors").Cells(EmptRow, k + 1) = ArrZGLS(j, k)
                        Next k
                    End If
                Next j
                Exit For
            End If
        Next i
    End If
    
    LastRow = FindLastRow(1, 2, 0, 0, "Errors")
    FirstRow = FindLastRow(1, 1, 1, 0, "Errors")
    If LastRow = FirstRow Then
        Sheets("Errors").Cells(LastRow, 1).EntireRow.ClearContents
        Sheets("Errors").Cells(LastRow, 2) = "No cross Profit Center errors."
    Else
        CPCF = True
        Set rng = Sheets("Errors").Range(Sheets("Errors").Cells(FirstRow, 2), Sheets("Errors").Cells(LastRow, 8))
        Call ApplyBorders(1, 1, rng)
        Call ApplyBorders(2, 2, rng)
        With Sheets("Errors").Sort
            .SortFields.Clear
            .SortFields.Add Key:=Range("D" & FirstRow & ":D" & LastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add Key:=Range("G" & FirstRow & ":G" & LastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange rng
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        rng.HorizontalAlignment = xlCenter
        With Sheets("Errors").Range(Sheets("Errors").Cells(FirstRow, 2), Sheets("Errors").Cells(FirstRow, 8))
            .EntireColumn.AutoFit
            .Font.Bold = True
            .Interior.Pattern = xlSolid
            .Interior.PatternColorIndex = xlAutomatic
            .Interior.Color = 14071936
            .Interior.TintAndShade = 0
            .Interior.PatternTintAndShade = 0
        End With
        Sheets("Errors").Range(Sheets("Errors").Cells(FirstRow + 1, 5), Sheets("Errors").Cells(LastRow, 5)).NumberFormat = "#,##0.00"
        w = FirstRow + 1
        With Sheets("Errors")
            Do Until .Cells(w, 2) = ""
                If .Cells(w, 4) <> .Cells(w + 1, 4) Or .Cells(w, 7) <> .Cells(w + 1, 7) Then
                    .Range(.Cells(w, 2), .Cells(w, 8)).Borders(xlEdgeBottom).Weight = xlMedium
                End If
                w = w + 1
            Loop
        End With
    End If
End If

If OtherF = True Or BPCF = True Or CPCF = True Then
    CheckZGLRME = False
Else
    CheckZGLRME = True
End If

Kill FPath & "zglrme.txt"

End Function
Function CheckAA02(sess)

Dim Am As Double

BPCWDF = False

With sess
    .findById("wnd[0]").maximize
    .findById("wnd[0]/tbar[0]/okcd").Text = "/ngr55"
    .findById("wnd[0]").sendVKey 0
    .findById("wnd[0]/usr/ctxtRGRWJ-JOB").Text = "AA02"
    .findById("wnd[0]").sendVKey 8
    .findById("wnd[0]/usr/txt$CUR-YR").Text = Yearx
    .findById("wnd[0]/usr/txt$CUR-PER").Text = Monthx
    .findById("wnd[0]/usr/ctxt$GL-PC01").Text = ""
    .findById("wnd[0]/usr/ctxt_COCD1-LOW").Text = CC
    .findById("wnd[0]/usr/ctxt$CST-REV").Text = ""
    .findById("wnd[0]/usr/ctxt$CST-REV").SetFocus
    .findById("wnd[0]/usr/ctxt$CST-REV").caretPosition = 0
    .findById("wnd[0]/tbar[1]/btn[8]").press
    .findById("wnd[0]/mbar/menu[6]/menu[5]/menu[2]/menu[2]").Select
    .findById("wnd[1]/tbar[0]/btn[0]").press
    .findById("wnd[1]/usr/ctxtDY_PATH").Text = FPath
    .findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "aa02.txt"
    .findById("wnd[1]/tbar[0]/btn[11]").press
    .findById("wnd[0]/tbar[0]/btn[3]").press
    On Error Resume Next
    .findById("wnd[1]/usr/btnBUTTON_YES").press
    On Error GoTo 0
    .findById("wnd[0]/tbar[0]/btn[3]").press
    .findById("wnd[0]/tbar[0]/btn[3]").press
End With

Set strim = New ADODB.Stream
strim.Charset = "utf-8"
strim.Open

strim.LoadFromFile (FPath & "aa02.txt")

If Sheets("Errors").Range("A10") = "" Then
    Sheets("Errors").Range("A10") = "Blank Profit Center on all documents' lines (local currency):"
    Sheets("Errors").Range("A10").Font.Bold = True
Else
    EmptRow = FindLastRow(1, 2, 2, 0, "Errors")
    EmptRow1 = FindLastRow(1, 1, 2, 0, "Errors")
    If EmptRow1 > EmptRow Then
        Sheets("Errors").Cells(EmptRow1, 1) = "Blank Profit Center on all documents' lines (local currency):"
        Sheets("Errors").Cells(EmptRow1, 1).Font.Bold = True
    Else
        Sheets("Errors").Cells(EmptRow, 1) = "Blank Profit Center on all documents' lines (local currency):"
        Sheets("Errors").Cells(EmptRow, 1).Font.Bold = True
    End If
End If
EmptRow = FindLastRow(1, 1, 1, 0, "Errors")
Sheets("Errors").Cells(EmptRow, 2) = "Account"
Sheets("Errors").Cells(EmptRow, 3) = "Balance"

Do Until strim.EOS
    Dim line As String
    line = strim.ReadText(-2)

    If IsNumeric(Left(Trim(line), 1)) Then
        n = InStr(1, line, "  ", vbBinaryCompare)
        Do Until n = 0
            line = Replace(line, "  ", " ")
            n = InStr(1, line, "  ", vbBinaryCompare)
        Loop
        arr = Split(line, " ")
        'V4-CIO FIX: locale-proof amount. Round()/CDbl() on a SAP string use this
        'PC's Windows regional format, so a mismatch with the SAP user's decimal
        'notation raised "Run-time error 13: Type mismatch" on the whole column.
        Am = Round(CM_Amount(Replace(Replace(arr(UBound(arr, 1) - 1), "(", "-"), ")", ""), 0, _
                   "reading an amount from the AA02 report group"), 2)
        If Round(Am, 2) <> 0 Then
            EmptRow = FindLastRow(1, 2, 1, 0, "Errors")
            Sheets("Errors").Cells(EmptRow, 2) = arr(1)
            Sheets("Errors").Cells(EmptRow, 3) = Am
        End If
    End If
Loop

LastRow = FindLastRow(1, 2, 0, 0, "Errors")
FirstRow = FindLastRow(1, 1, 1, 0, "Errors")
If LastRow = FirstRow Then
    Sheets("Errors").Cells(LastRow, 1).EntireRow.ClearContents
    Sheets("Errors").Cells(LastRow, 2) = "No errors."
    CheckAA02 = True
Else
    BPCWDF = True
    CheckAA02 = False
    Set rng = Sheets("Errors").Range(Sheets("Errors").Cells(FirstRow, 2), Sheets("Errors").Cells(LastRow, 3))
    Call ApplyBorders(1, 1, rng)
    Call ApplyBorders(2, 2, rng)
    rng.HorizontalAlignment = xlCenter
    With Sheets("Errors").Range(Sheets("Errors").Cells(FirstRow, 2), Sheets("Errors").Cells(FirstRow, 3))
        .EntireColumn.AutoFit
        .Font.Bold = True
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.Color = 14071936
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
    End With
    Sheets("Errors").Range(Sheets("Errors").Cells(FirstRow + 1, 3), Sheets("Errors").Cells(LastRow, 3)).NumberFormat = "#,##0.00"
End If

Kill FPath & "aa02.txt"

End Function
