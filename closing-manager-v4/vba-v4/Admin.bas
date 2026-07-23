Attribute VB_Name = "Admin"
'Owner: Engagement Executive
'Description: https://capgemini.sharepoint.com/sites/InternationalPaper/capgemini/Lists/Automations_list/Allitemsg.aspx
'Risk category: https://capgemini.sharepoint.com/sites/InternationalPaper/capgemini/Lists/Automations_list/Allitemsg.aspx
'Data Category (PII/ Non-PII *):  https://capgemini.sharepoint.com/sites/InternationalPaper/capgemini/Lists/Automations_list/Allitemsg.aspx
'Reviewer:  https://capgemini.sharepoint.com/sites/InternationalPaper/capgemini/Lists/Automations_list/Allitemsg.aspx
'Date of validation: https://capgemini.sharepoint.com/sites/InternationalPaper/capgemini/Lists/Automations_list/Allitemsg.aspx
Sub UpdateData()

Dim sess As Object
Dim k As Long

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
    Exit Sub
End If

Call CreatePaths
Call CreateVariants

'----------------------------------------------------------------
'check tracker
'----------------------------------------------------------------
Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
xmlDoc.async = False

Url = CM_SP_BASE
List = "ClosingTracker"

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

With CreateObject("Microsoft.XMLHTTP")
    .Open "Get", Url, False, "", ""
    .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    .setRequestHeader "SOAPAction", "http://schemas.microsoft.com/sharepoint/soap/GetListItems"
    .send request
    
    xmlDoc.LoadXML (.responsetext)
    found = False
    For Each X In xmlDoc.getElementsByTagName("z:row")
        If Round(X.getAttribute("ows_Periodx"), 0) = Month(LastDay) And Round(X.getAttribute("ows_Yearx"), 0) = Year(LastDay) Then
            found = True
            If X.getAttribute("ows_Timestamp") > Date1 Then
                Date1 = X.getAttribute("ows_Timestamp")
                UName = X.getAttribute("ows_UName")
            End If
        End If
    Next
    If found = True Then
        Sheets("config").Range("B3").ClearContents
        strOption = "Data updated already for " & Month(LastDay) & "/" & Year(LastDay) & " on " & Date1 & " by " & UName & ". Do you want to update it again?"
        With UF_Option
            .Lbl_Option.Caption = strOption
            .StartUpPosition = 0
            .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
            .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
            .Show
        End With
        If Sheets("config").Range("B3") <> "Yes" Then Exit Sub
    End If
End With

With UF_Run
    .StartUpPosition = 0
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    .Show (vbModeless)
    .Repaint
End With

With Application
    .ScreenUpdating = False
    .DisplayAlerts = False
End With

'clearing
LastRow = FindLastRow(1, 13, 0, 0, "Changes")
If LastRow > 1 Then Sheets("Changes").Range("M2", Sheets("Changes").Cells(LastRow, 14)).ClearContents
LastRow = FindLastRow(1, 8, 0, 0, "Changes")
If LastRow > 1 Then Sheets("Changes").Range("H2", Sheets("Changes").Cells(LastRow, 11)).ClearContents

'----------------------------------------------------------------
'importing data from SAP
'----------------------------------------------------------------
With sess
    'list of variants for closing
    .findById("wnd[0]").maximize
    .findById("wnd[0]/tbar[0]/okcd").Text = "/nse16"
    .findById("wnd[0]").sendVKey 0
    .findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").Text = "T001"
    .findById("wnd[0]").sendVKey 0
    On Error Resume Next
    .findById("wnd[1]/tbar[0]/btn[0]").press
    On Error GoTo 0
    .findById("wnd[0]/mbar/menu[3]/menu[1]").Select
    .findById("wnd[1]/usr/tabsG_TABSTRIP/tabp0400/ssubTOOLAREA:SAPLWB_CUSTOMIZING:0400/ctxtRSEUMOD-TBLISTBR").Text = "250"
    .findById("wnd[1]/usr/tabsG_TABSTRIP/tabp0400/ssubTOOLAREA:SAPLWB_CUSTOMIZING:0400/txtRSEUMOD-TBMAXSEL").Text = "999999999"
    .findById("wnd[1]/usr/tabsG_TABSTRIP/tabp0400/ssubTOOLAREA:SAPLWB_CUSTOMIZING:0400/radRSEUMOD-TBALV_GRID").Select
    .findById("wnd[1]/usr/tabsG_TABSTRIP/tabp0400/ssubTOOLAREA:SAPLWB_CUSTOMIZING:0400/radSEUCUSTOM-FIELDNAME").Select
    .findById("wnd[1]/tbar[0]/btn[0]").press
    .findById("wnd[0]/mbar/menu[3]/menu[0]/menu[1]").Select
    .findById("wnd[1]/tbar[0]/btn[14]").press
    
    LastRow = FindLastRow(1, 1, 0, 0, "SAP config")
    j = 1
    Do Until LastRow = j
        Set Area = .findById("wnd[1]/usr")
        Set Children = Area.Children()
        For i = 0 To Children.Count() - 1
            Set obj = Children(CInt(i))
            If obj.Type = "GuiLabel" And obj.Text <> "" Then
                w = 2
                Do Until Sheets("SAP config").Cells(w, 1) = ""
                    FieldName = Sheets("SAP config").Cells(w, 1)
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
    .findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "t001.txt"
    .findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").Text = "4120"
    .findById("wnd[1]/tbar[0]/btn[11]").press
    .findById("wnd[0]/tbar[0]/btn[3]").press
    .findById("wnd[0]/mbar/menu[3]/menu[1]").Select
    .findById("wnd[1]/usr/tabsG_TABSTRIP/tabp0400/ssubTOOLAREA:SAPLWB_CUSTOMIZING:0400/radSEUCUSTOM-FIELDTEXT").Select
    .findById("wnd[1]").sendVKey 0
    .findById("wnd[0]/tbar[0]/btn[3]").press
    .findById("wnd[0]/tbar[0]/btn[3]").press
    
    'list of Profit Centers
    .findById("wnd[0]").maximize
    .findById("wnd[0]/tbar[0]/okcd").Text = "/nse16"
    .findById("wnd[0]").sendVKey 0
    .findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").Text = "ZCCOD"
    .findById("wnd[0]").sendVKey 0
    On Error Resume Next
    .findById("wnd[1]/tbar[0]/btn[0]").press
    On Error GoTo 0
    .findById("wnd[0]/mbar/menu[3]/menu[1]").Select
    .findById("wnd[1]/usr/tabsG_TABSTRIP/tabp0400/ssubTOOLAREA:SAPLWB_CUSTOMIZING:0400/ctxtRSEUMOD-TBLISTBR").Text = "250"
    .findById("wnd[1]/usr/tabsG_TABSTRIP/tabp0400/ssubTOOLAREA:SAPLWB_CUSTOMIZING:0400/txtRSEUMOD-TBMAXSEL").Text = "999999999"
    .findById("wnd[1]/usr/tabsG_TABSTRIP/tabp0400/ssubTOOLAREA:SAPLWB_CUSTOMIZING:0400/radRSEUMOD-TBALV_GRID").Select
    .findById("wnd[1]/usr/tabsG_TABSTRIP/tabp0400/ssubTOOLAREA:SAPLWB_CUSTOMIZING:0400/radSEUCUSTOM-FIELDNAME").Select
    .findById("wnd[1]/tbar[0]/btn[0]").press
    .findById("wnd[0]/mbar/menu[3]/menu[0]/menu[1]").Select
    .findById("wnd[1]/tbar[0]/btn[14]").press
    
    LastRow = FindLastRow(1, 2, 0, 0, "SAP config")
    j = 1
    Do Until LastRow = j
        Set Area = .findById("wnd[1]/usr")
        Set Children = Area.Children()
        For i = 0 To Children.Count() - 1
            Set obj = Children(CInt(i))
            If obj.Type = "GuiLabel" And obj.Text <> "" Then
                w = 2
                Do Until Sheets("SAP config").Cells(w, 2) = ""
                    FieldName = Sheets("SAP config").Cells(w, 2)
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
    .findById("wnd[0]/tbar[1]/btn[32]").press
    
    sess.findById("wnd[0]/tbar[1]/btn[32]").press
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
    .findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "zgxmit.txt"
    .findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").Text = "4120"
    .findById("wnd[1]/tbar[0]/btn[11]").press
    .findById("wnd[0]/tbar[0]/btn[3]").press
    .findById("wnd[0]/mbar/menu[3]/menu[1]").Select
    .findById("wnd[1]/usr/tabsG_TABSTRIP/tabp0400/ssubTOOLAREA:SAPLWB_CUSTOMIZING:0400/radSEUCUSTOM-FIELDTEXT").Select
    .findById("wnd[1]").sendVKey 0
    .findById("wnd[0]/tbar[0]/btn[3]").press
    .findById("wnd[0]/tbar[0]/btn[3]").press
End With

'----------------------------------------------------------------
'importing data from files
'----------------------------------------------------------------
Set strim = New ADODB.Stream
strim.Charset = "utf-8"
strim.Open

Call CreateArray(FPath & "t001.txt")
strim.LoadFromFile (FPath & "t001.txt")

Do Until strim.EOS
    Dim line As String, tablica() As String
    line = strim.ReadText(-2)

    If Left(Trim(line), 1) = "|" And Trim(getLineData(line, FirstColumn, 1)) <> FirstColumn Then
        
        EmptRow = FindLastRow(1, 13, 1, 0, "Changes")
        Sheets("Changes").Cells(EmptRow, 13) = getLineData(line, "BUKRS", 1)
        Sheets("Changes").Cells(EmptRow, 14) = getLineData(line, "OPVAR", 1)
        Sheets("Changes").Cells(EmptRow, 15) = getLineData(line, "WAERS", 1)
        
    End If
Loop

Call CreateArray(FPath & "zgxmit.txt")
strim.LoadFromFile (FPath & "zgxmit.txt")

Do Until strim.EOS
    line = strim.ReadText(-2)

    If Left(Trim(line), 1) = "|" And Trim(getLineData(line, FirstColumn, 1)) <> FirstColumn Then
        
        EmptRow = FindLastRow(1, 8, 1, 0, "Changes")
        If getLineData(line, "INDIC", 1) <> "" And getLineData(line, "SP_COCODE", 1) <> "" And getLineData(line, "LG_LOCNUM", 1) <> "" Then
            Sheets("Changes").Cells(EmptRow, 8) = getLineData(line, "SP_COCODE", 1)
            Sheets("Changes").Cells(EmptRow, 9) = getLineData(line, "LG_LOCNUM", 1)
            Sheets("Changes").Cells(EmptRow, 10) = getLineData(line, "XMIT", 1)
            Sheets("Changes").Cells(EmptRow, 11) = getLineData(line, "INDIC", 1)
        End If
        
    End If
Loop

LastRow = FindLastRow(1, 8, 0, 0, "Changes")
ArrUp = Sheets("Changes").Range("H2", Sheets("Changes").Cells(LastRow, 11))

For i = 1 To UBound(ArrUp, 1)
    If ArrUp(i, 3) = "X" Then
        ArrUp(i, 3) = 1
    Else
        ArrUp(i, 3) = 0
    End If
Next i
Sheets("Changes").Range("H2", Sheets("Changes").Cells(LastRow, 11)) = ArrUp

'----------------------------------------------------------------
'update variant table
'----------------------------------------------------------------
'import current table
LastRow = FindLastRow(1, 17, 0, 0, "config")
If LastRow > 1 Then Sheets("config").Range("Q2", Sheets("config").Cells(LastRow, 21)).ClearContents

List = "ClosingVariants"

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

With CreateObject("Microsoft.XMLHTTP")
    .Open "Get", Url, False, "", ""
    .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    .setRequestHeader "SOAPAction", "http://schemas.microsoft.com/sharepoint/soap/GetListItems"
    .send request
    
    xmlDoc.LoadXML (.responsetext)
    For Each X In xmlDoc.getElementsByTagName("z:row")
        EmptRow = FindLastRow(1, 17, 1, 0, "config")
        Sheets("config").Cells(EmptRow, 17) = X.getAttribute("ows_ID")
        Sheets("config").Cells(EmptRow, 18) = X.getAttribute("ows_CC")
        Sheets("config").Cells(EmptRow, 19) = X.getAttribute("ows_VariantName")
        Sheets("config").Cells(EmptRow, 20) = X.getAttribute("ows_Currency")
        Sheets("config").Cells(EmptRow, 21) = X.getAttribute("ows_FCUSD")
    Next
End With

'import data from SAP
With sess
    .findById("wnd[0]").maximize
    .findById("wnd[0]/tbar[0]/okcd").Text = "/nse16"
    .findById("wnd[0]").sendVKey 0
    .findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").Text = "t001z"
    .findById("wnd[0]").sendVKey 0
    On Error Resume Next
    .findById("wnd[1]/tbar[0]/btn[0]").press
    On Error GoTo 0
    .findById("wnd[0]/mbar/menu[3]/menu[1]").Select
    .findById("wnd[1]/usr/tabsG_TABSTRIP/tabp0400/ssubTOOLAREA:SAPLWB_CUSTOMIZING:0400/ctxtRSEUMOD-TBLISTBR").Text = "250"
    .findById("wnd[1]/usr/tabsG_TABSTRIP/tabp0400/ssubTOOLAREA:SAPLWB_CUSTOMIZING:0400/txtRSEUMOD-TBMAXSEL").Text = "999999999"
    .findById("wnd[1]/usr/tabsG_TABSTRIP/tabp0400/ssubTOOLAREA:SAPLWB_CUSTOMIZING:0400/radRSEUMOD-TBALV_GRID").Select
    .findById("wnd[1]/usr/tabsG_TABSTRIP/tabp0400/ssubTOOLAREA:SAPLWB_CUSTOMIZING:0400/radSEUCUSTOM-FIELDNAME").Select
    .findById("wnd[1]/tbar[0]/btn[0]").press
    .findById("wnd[0]/mbar/menu[3]/menu[2]").Select
    .findById("wnd[1]/tbar[0]/btn[14]").press
    
    k = 8
    Call SAPSelectFields(sess, k)
    
    .findById("wnd[1]/tbar[0]/btn[9]").press
    .findById("wnd[1]/tbar[0]/btn[0]").press
    .findById("wnd[0]/usr/ctxtI2-LOW").Text = "0XMTGP"
    .findById("wnd[0]/usr/ctxtI2-LOW").SetFocus
    .findById("wnd[0]/usr/ctxtI2-LOW").caretPosition = 6
    .findById("wnd[0]/usr/btn%_I3_%_APP_%-VALU_PUSH").press
    .findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,0]").Text = "x"
    .findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,1]").Text = "X"
    .findById("wnd[1]/tbar[0]/btn[8]").press
    
    .findById("wnd[0]/mbar/menu[3]/menu[0]/menu[1]").Select
    .findById("wnd[1]/tbar[0]/btn[14]").press
    
    LastRow = FindLastRow(1, 7, 0, 0, "SAP config")
    j = 1
    Do Until LastRow = j
        Set Area = .findById("wnd[1]/usr")
        Set Children = Area.Children()
        For i = 0 To Children.Count() - 1
            Set obj = Children(CInt(i))
            If obj.Type = "GuiLabel" And obj.Text <> "" Then
                w = 2
                Do Until Sheets("SAP config").Cells(w, 7) = ""
                    FieldName = Sheets("SAP config").Cells(w, 7)
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
    .findById("wnd[0]/tbar[1]/btn[32]").press
    
    sess.findById("wnd[0]/tbar[1]/btn[32]").press
    VR = .findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").visiblerowcount
    Do Until VR = 0
        .findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "0-" & VR - 1
        .findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press
        VR = .findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").visiblerowcount
    Loop
    
    .findById("wnd[1]/tbar[0]/btn[0]").press
    .findById("wnd[0]/tbar[1]/btn[45]").press
    .findById("wnd[1]/tbar[0]/btn[0]").press
    .findById("wnd[1]/usr/ctxtDY_PATH").Text = FPath
    .findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "t001z.txt"
    .findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
    .findById("wnd[1]/tbar[0]/btn[11]").press
    .findById("wnd[0]/tbar[0]/btn[3]").press
    .findById("wnd[0]/mbar/menu[3]/menu[1]").Select
    .findById("wnd[1]/usr/tabsG_TABSTRIP/tabp0400/ssubTOOLAREA:SAPLWB_CUSTOMIZING:0400/radSEUCUSTOM-FIELDTEXT").Select
    .findById("wnd[1]/tbar[0]/btn[0]").press
    .findById("wnd[0]/tbar[0]/btn[3]").press
    .findById("wnd[0]/tbar[0]/btn[3]").press
End With

LastRow = FindLastRow(1, 13, 0, 0, "Changes")
ArrUp = Sheets("Changes").Range("M2", Sheets("Changes").Cells(LastRow, 16))

Call CreateArray(FPath & "t001z.txt")

For i = 1 To UBound(ArrUp, 1)
    found = False
    strim.LoadFromFile (FPath & "t001z.txt")
    Do Until strim.EOS Or found = True
        line = strim.ReadText(-2)
        If Left(Trim(line), 1) = "|" And Trim(getLineData(line, FirstColumn, 1)) <> FirstColumn Then
            If ArrUp(i, 1) = getLineData(line, "BUKRS", 1) Then
                found = True
                ArrUp(i, 4) = 1
            End If
        End If
    Loop
    If found = False Then ArrUp(i, 4) = 0
Next i

'update table if new item
LastRow = FindLastRow(1, 17, 0, 0, "config")
ArrCur = Sheets("config").Range("Q2", Sheets("config").Cells(LastRow, 21))

For i = 1 To UBound(ArrUp, 1)
    found = False
    For j = 1 To UBound(ArrCur, 1)
        If ArrCur(j, 2) = ArrUp(i, 1) Then
            found = True
            Exit For
        End If
    Next j
    If found = False Then
        updates = "<Batch> <Method ID='1' Cmd='New'>" & _
                    "<Field Name='CC'>" & ArrUp(i, 1) & "</Field>" & _
                    "<Field Name='VariantName'>" & ArrUp(i, 2) & "</Field>" & _
                    "<Field Name='Currency'>" & ArrUp(i, 3) & "</Field>" & _
                    "<Field Name='FCUSD'>" & ArrUp(i, 4) & "</Field>" & _
                        "</Method></Batch>"
        
        Call spAddToList(updates, List)
        
    End If
Next i

'update table if changed item
For i = 1 To UBound(ArrUp, 1)
    For j = 1 To UBound(ArrCur, 1)
        If ArrUp(i, 1) = ArrCur(j, 2) Then
            If ArrUp(i, 2) <> ArrCur(j, 3) Or ArrUp(i, 3) <> ArrCur(j, 4) Or CStr(ArrUp(i, 4)) <> CStr(ArrCur(j, 5)) Then
                updates = "<Batch> <Method ID='1' Cmd='Update'>" & _
                            "<Field Name='ID'>" & ArrCur(j, 1) & "</Field>" & _
                            "<Field Name='VariantName'>" & ArrUp(i, 2) & "</Field>" & _
                            "<Field Name='Currency'>" & ArrUp(i, 3) & "</Field>" & _
                            "<Field Name='FCUSD'>" & ArrUp(i, 4) & "</Field>" & _
                            "</Method></Batch>"
                
                Call spUpdateList(updates, List)
            End If
            Exit For
        End If
    Next j
Next i

'import table after changes
LastRow = FindLastRow(1, 17, 0, 0, "config")
If LastRow > 1 Then Sheets("config").Range("Q2", Sheets("config").Cells(LastRow, 21)).ClearContents

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

With CreateObject("Microsoft.XMLHTTP")
    .Open "Get", Url, False, "", ""
    .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    .setRequestHeader "SOAPAction", "http://schemas.microsoft.com/sharepoint/soap/GetListItems"
    .send request
    
    xmlDoc.LoadXML (.responsetext)
    For Each X In xmlDoc.getElementsByTagName("z:row")
        EmptRow = FindLastRow(1, 17, 1, 0, "config")
        Sheets("config").Cells(EmptRow, 17) = X.getAttribute("ows_ID")
        Sheets("config").Cells(EmptRow, 18) = X.getAttribute("ows_CC")
        Sheets("config").Cells(EmptRow, 19) = X.getAttribute("ows_VariantName")
        Sheets("config").Cells(EmptRow, 20) = X.getAttribute("ows_Currency")
        Sheets("config").Cells(EmptRow, 21) = X.getAttribute("ows_FCUSD")
    Next
End With

'----------------------------------------------------------------
'update PC table
'----------------------------------------------------------------
'import current table
LastRow = FindLastRow(1, 11, 0, 0, "config")
If LastRow > 1 Then Sheets("config").Range("K2", Sheets("config").Cells(LastRow, 15)).ClearContents

List = "ProfitCenters"

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

With CreateObject("Microsoft.XMLHTTP")
    .Open "Get", Url, False, "", ""
    .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    .setRequestHeader "SOAPAction", "http://schemas.microsoft.com/sharepoint/soap/GetListItems"
    .send request

    xmlDoc.LoadXML (.responsetext)
        
    For Each X In xmlDoc.getElementsByTagName("z:row")
        EmptRow = FindLastRow(1, 11, 1, 0, "config")
        Sheets("config").Cells(EmptRow, 11) = X.getAttribute("ows_ID")
        Sheets("config").Cells(EmptRow, 12) = X.getAttribute("ows_CC")
        Sheets("config").Cells(EmptRow, 13) = X.getAttribute("ows_PC")
        Sheets("config").Cells(EmptRow, 14) = X.getAttribute("ows_Transmit")
        Sheets("config").Cells(EmptRow, 15) = X.getAttribute("ows_GAAPindicator")
    Next
End With

'update table if new item
LastRow = FindLastRow(1, 11, 0, 0, "config")
ArrCur = Sheets("config").Range("K2", Sheets("config").Cells(LastRow, 15))

LastRow = FindLastRow(1, 8, 0, 0, "Changes")
ArrUp = Sheets("Changes").Range("H2", Sheets("Changes").Cells(LastRow, 11))

For i = 1 To UBound(ArrUp, 1)
    found = False
    For j = 1 To UBound(ArrCur, 1)
        If CStr(ArrCur(j, 3)) = CStr(ArrUp(i, 2)) Then
            found = True
            Exit For
        End If
    Next j
    If found = False Then
        updates = "<Batch> <Method ID='1' Cmd='New'>" & _
                    "<Field Name='CC'>" & ArrUp(i, 1) & "</Field>" & _
                    "<Field Name='PC'>" & ArrUp(i, 2) & "</Field>" & _
                    "<Field Name='Transmit'>" & ArrUp(i, 3) & "</Field>" & _
                    "<Field Name='GAAPindicator'>" & ArrUp(i, 4) & "</Field>" & _
                        "</Method></Batch>"
        
        Call spAddToList(updates, List)
        
    End If
Next i

'update table if changed item
For i = 1 To UBound(ArrUp, 1)
    For j = 1 To UBound(ArrCur, 1)
        If CStr(ArrUp(i, 2)) = CStr(ArrCur(j, 3)) Then
            If CStr(ArrUp(i, 3)) <> CStr(ArrCur(j, 4)) Or CStr(ArrUp(i, 4)) <> CStr(ArrCur(j, 5)) Then
                updates = "<Batch> <Method ID='1' Cmd='Update'>" & _
                            "<Field Name='ID'>" & ArrCur(j, 1) & "</Field>" & _
                            "<Field Name='Transmit'>" & ArrUp(i, 3) & "</Field>" & _
                            "<Field Name='GAAPindicator'>" & ArrUp(i, 4) & "</Field>" & _
                            "</Method></Batch>"
                            
                Call spUpdateList(updates, List)
                
            End If
        End If
    Next j
Next i

'import table after changes
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

With CreateObject("Microsoft.XMLHTTP")
    .Open "Get", Url, False, "", ""
    .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    .setRequestHeader "SOAPAction", "http://schemas.microsoft.com/sharepoint/soap/GetListItems"
    .send request
    
    xmlDoc.LoadXML (.responsetext)
    
    For Each X In xmlDoc.getElementsByTagName("z:row")
        EmptRow = FindLastRow(1, 11, 1, 0, "config")
        Sheets("config").Cells(EmptRow, 11) = X.getAttribute("ows_ID")
        Sheets("config").Cells(EmptRow, 12) = X.getAttribute("ows_CC")
        Sheets("config").Cells(EmptRow, 13) = X.getAttribute("ows_PC")
        Sheets("config").Cells(EmptRow, 14) = X.getAttribute("ows_Transmit")
        Sheets("config").Cells(EmptRow, 15) = X.getAttribute("ows_GAAPindicator")
    Next
End With

'----------------------------------------------------------------
'update posting blocks for 5100001 account
'----------------------------------------------------------------
LastRow = FindLastRow(1, 3, 0, 0, "config")
If LastRow > 1 Then Sheets("config").Range("C2", Sheets("config").Cells(LastRow, 9)).ClearContents

Url = CM_SP_BASE
List = "CCCrossList"

Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
xmlDoc.async = False

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

With CreateObject("Microsoft.XMLHTTP")
    .Open "Get", Url, False, "", ""
    .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    .setRequestHeader "SOAPAction", "http://schemas.microsoft.com/sharepoint/soap/GetListItems"
    .send request
    
    xmlDoc.LoadXML (.responsetext)
    
    For Each X In xmlDoc.getElementsByTagName("z:row")
        EmptRow = FindLastRow(1, 3, 1, 0, "config")
        Sheets("config").Cells(EmptRow, 3) = X.getAttribute("ows_ID")
        Sheets("config").Cells(EmptRow, 4) = X.getAttribute("ows_CC")
        Sheets("config").Cells(EmptRow, 5) = X.getAttribute("ows_PostingBlock")
        Sheets("config").Cells(EmptRow, 6) = X.getAttribute("ows_LocationClosed")
        Sheets("config").Cells(EmptRow, 7) = X.getAttribute("ows_CPCAllowedGAAP")
        Sheets("config").Cells(EmptRow, 8) = X.getAttribute("ows_CPCNotAllowed")
        Sheets("config").Cells(EmptRow, 9) = X.getAttribute("ows_Comment")
        
    Next
End With

'import from SAP
LastRow = FindLastRow(1, 3, 0, 0, "config")
ArrUp = Sheets("config").Range("D2", Sheets("config").Cells(LastRow, 9))

Set alloutput = New ADODB.Stream
        
alloutput.Charset = "utf-8"
alloutput.Open

For i = 1 To UBound(ArrUp, 1)
    
    alloutput.writetext (ArrUp(i, 1) & vbCrLf)
                                            
Next i

alloutput.SaveToFile FPath & "\CC.csv", 2
alloutput.Close

With sess
    .findById("wnd[0]").maximize
    .findById("wnd[0]/tbar[0]/okcd").Text = "/nse16"
    .findById("wnd[0]").sendVKey 0
    .findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").Text = "SKB1"
    .findById("wnd[0]").sendVKey 0
    On Error Resume Next
    .findById("wnd[1]/tbar[0]/btn[0]").press
    On Error GoTo 0
    .findById("wnd[0]/mbar/menu[3]/menu[2]").Select
    .findById("wnd[1]/tbar[0]/btn[14]").press
    
    k = 6
    Call SAPSelectFields(sess, k)
    
    .findById("wnd[1]/tbar[0]/btn[0]").press
    .findById("wnd[0]/mbar/menu[3]/menu[1]").Select
    .findById("wnd[1]/usr/tabsG_TABSTRIP/tabp0400/ssubTOOLAREA:SAPLWB_CUSTOMIZING:0400/ctxtRSEUMOD-TBLISTBR").Text = "250"
    .findById("wnd[1]/usr/tabsG_TABSTRIP/tabp0400/ssubTOOLAREA:SAPLWB_CUSTOMIZING:0400/txtRSEUMOD-TBMAXSEL").Text = "999999999"
    .findById("wnd[1]/usr/tabsG_TABSTRIP/tabp0400/ssubTOOLAREA:SAPLWB_CUSTOMIZING:0400/radRSEUMOD-TBALV_GRID").Select
    .findById("wnd[1]/usr/tabsG_TABSTRIP/tabp0400/ssubTOOLAREA:SAPLWB_CUSTOMIZING:0400/radSEUCUSTOM-FIELDNAME").Select
    .findById("wnd[1]/tbar[0]/btn[0]").press
    .findById("wnd[0]/mbar/menu[3]/menu[0]/menu[1]").Select
    .findById("wnd[1]/tbar[0]/btn[14]").press
    
    LastRow = FindLastRow(1, 5, 0, 0, "SAP config")
    j = 1
    Do Until LastRow = j
        Set Area = .findById("wnd[1]/usr")
        Set Children = Area.Children()
        For i = 0 To Children.Count() - 1
            Set obj = Children(CInt(i))
            If obj.Type = "GuiLabel" And obj.Text <> "" Then
                w = 2
                Do Until Sheets("SAP config").Cells(w, 5) = ""
                    FieldName = Sheets("SAP config").Cells(w, 5)
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
    .findById("wnd[0]/usr/btn%_I1_%_APP_%-VALU_PUSH").press
    .findById("wnd[1]/tbar[0]/btn[23]").press
    .findById("wnd[2]/usr/ctxtDY_PATH").Text = FPath
    .findById("wnd[2]/usr/ctxtDY_FILENAME").Text = "CC.csv"
    .findById("wnd[2]/tbar[0]/btn[0]").press
    .findById("wnd[1]/tbar[0]/btn[8]").press
    .findById("wnd[0]/usr/ctxtI2-LOW").Text = "5100001"
    
    .findById("wnd[0]/tbar[1]/btn[8]").press
    .findById("wnd[0]/tbar[1]/btn[32]").press
    
    sess.findById("wnd[0]/tbar[1]/btn[32]").press
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
    .findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "skb1.txt"
    .findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").Text = "4120"
    .findById("wnd[1]/tbar[0]/btn[11]").press
    .findById("wnd[0]/tbar[0]/btn[3]").press
    .findById("wnd[0]/mbar/menu[3]/menu[1]").Select
    .findById("wnd[1]/usr/tabsG_TABSTRIP/tabp0400/ssubTOOLAREA:SAPLWB_CUSTOMIZING:0400/radSEUCUSTOM-FIELDTEXT").Select
    .findById("wnd[1]").sendVKey 0
    .findById("wnd[0]/tbar[0]/btn[3]").press
    .findById("wnd[0]/tbar[0]/btn[3]").press
    
End With

Call CreateArray(FPath & "skb1.txt")
strim.LoadFromFile (FPath & "skb1.txt")

Do Until strim.EOS
    line = strim.ReadText(-2)

    If Left(Trim(line), 1) = "|" And Trim(getLineData(line, FirstColumn, 1)) <> FirstColumn Then
        
        For i = 1 To UBound(ArrUp, 1)
            If ArrUp(i, 1) = getLineData(line, "BUKRS", 1) Then
                If getLineData(line, "XSPEB", 1) = "X" Then
                    ArrUp(i, 2) = 1
                Else
                    ArrUp(i, 2) = 0
                End If
                Exit For
            End If
        Next i
        
    End If
Loop

'update table if changed item
LastRow = FindLastRow(1, 3, 0, 0, "config")
ArrCur = Sheets("config").Range("C2", Sheets("config").Cells(LastRow, 9))

For i = 1 To UBound(ArrUp, 1)
    For j = 1 To UBound(ArrCur, 1)
        If CStr(ArrUp(i, 1)) = CStr(ArrCur(j, 2)) Then
            updates = "<Batch> <Method ID='1' Cmd='Update'>" & _
                        "<Field Name='ID'>" & ArrCur(j, 1) & "</Field>" & _
                        "<Field Name='PostingBlock'>" & ArrUp(i, 3) & "</Field>" & _
                        "</Method></Batch>"
                        
            Call spUpdateList(updates, List)
            
        End If
    Next j
Next i

'import table after changes
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

With CreateObject("Microsoft.XMLHTTP")
    .Open "Get", Url, False, "", ""
    .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    .setRequestHeader "SOAPAction", "http://schemas.microsoft.com/sharepoint/soap/GetListItems"
    .send request
    
    xmlDoc.LoadXML (.responsetext)
    
    For Each X In xmlDoc.getElementsByTagName("z:row")
        EmptRow = FindLastRow(1, 3, 1, 0, "config")
        Sheets("config").Cells(EmptRow, 3) = X.getAttribute("ows_ID")
        Sheets("config").Cells(EmptRow, 4) = X.getAttribute("ows_CC")
        Sheets("config").Cells(EmptRow, 5) = X.getAttribute("ows_PostingBlock")
        Sheets("config").Cells(EmptRow, 6) = X.getAttribute("ows_LocationClosed")
        Sheets("config").Cells(EmptRow, 7) = X.getAttribute("ows_CPCAllowedGAAP")
        Sheets("config").Cells(EmptRow, 8) = X.getAttribute("ows_CPCNotAllowed")
        Sheets("config").Cells(EmptRow, 9) = X.getAttribute("ows_Comment")
        
    Next
End With

'----------------------------------------------------------------
'update tracker
'----------------------------------------------------------------
List = "ClosingTracker"
Date1 = Format(Now(), "yyyy-MM-ddTHH:mm:ssZ")

updates = "<Batch> <Method ID='1' Cmd='New'>" & _
            "<Field Name='UName'>" & UCase(Environ("username")) & "</Field>" & _
            "<Field Name='LogType'>" & "Update Data" & "</Field>" & _
            "<Field Name='Periodx'>" & Month(LastDay) & "</Field>" & _
            "<Field Name='Yearx'>" & Year(LastDay) & "</Field>" & _
            "<Field Name='Success'>1</Field>" & _
            "<Field Name='Timestamp'>" & Date1 & "</Field>" & _
            "</Method></Batch>"
        
Call spAddToList(updates, List)

Kill FPath & "t001.txt"
Kill FPath & "zgxmit.txt"
Kill FPath & "t001z.txt"
Kill FPath & "skb1.txt"
Kill FPath & "cc.csv"

Unload UF_Run
With Application
    .ScreenUpdating = True
    .DisplayAlerts = True
End With
strInfo = "Data updated."
With UF_Info
    .Lbl_Info.Caption = strInfo
    .StartUpPosition = 0
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    .Show
End With

End Sub
Sub ShowTracker()

Sheets("Tracker").Visible = True
Sheets("Tracker").Select

End Sub
Sub CloseTracker()

Sheets("Tracker").Visible = xlVeryHidden
Sheets("START").Select

End Sub
Sub ChangeCPC()

With UF_CPCChange
    .StartUpPosition = 0
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    .Show
End With

End Sub
Sub AddCPC()

With UF_CPCAdd
    .StartUpPosition = 0
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    .Show
End With

End Sub
Sub RunReport()

'import data abount Company Codes
'LastRow = FindLastRow(1, 3, 0, 0, "config")
'If LastRow > 1 Then Sheets("config").Range("C2", Sheets("config").Cells(LastRow, 9)).ClearContents

'Url = CM_SP_BASE
'List = "CCCrossList"
'
'Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
'xmlDoc.async = False
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
'
'    For Each X In xmlDoc.getElementsByTagName("z:row")
'        EmptRow = FindLastRow(1, 3, 1, 0, "config")
'        Sheets("config").Cells(EmptRow, 3) = X.getAttribute("ows_ID")
'        Sheets("config").Cells(EmptRow, 4) = X.getAttribute("ows_CC")
'        Sheets("config").Cells(EmptRow, 5) = X.getAttribute("ows_PostingBlock")
'        Sheets("config").Cells(EmptRow, 6) = X.getAttribute("ows_LocationClosed")
'        Sheets("config").Cells(EmptRow, 7) = X.getAttribute("ows_CPCAllowedGAAP")
'        Sheets("config").Cells(EmptRow, 8) = X.getAttribute("ows_CPCNotAllowed")
'        Sheets("config").Cells(EmptRow, 9) = X.getAttribute("ows_Comment")
'    Next
'End With

With UF_Report
    .StartUpPosition = 0
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    .Show
End With

End Sub
