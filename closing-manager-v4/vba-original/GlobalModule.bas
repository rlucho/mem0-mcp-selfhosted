Global FPath As String, FTemp As String, Fmerger As String, Fmerged As String, Fprinted As String, FFinal As String
Global LastDay As Date
Global Monthx, Yearx
Dim araj()
Dim datax
Dim Data
Sub CreatePaths()

Set fso = CreateObject("Scripting.FileSystemObject")

FPath = ThisWorkbook.Path
If Right(FPath, 1) <> "\" Then FPath = FPath & "\"

If fso.DriveExists("D:") Then
    DiscN = "D:\"
Else
    DiscN = "C:\"
End If

FTemp = DiscN & "pdf\temp"
Fmerger = DiscN & "pdf\merger"
Fmerged = DiscN & "pdf\mergedFiles"
Fprinted = DiscN & "pdf\printed"
FFinal = DiscN & "pdf\final"
FShared = "\\pl-krabpo-fsc01\ipa$\R2R\R2R - IP EU\MONTH-END\CLOSING REPORTS\" & Year(LastDay) & "\" & Right("0" & Month(LastDay), 2)

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

Url = "https://troom-x.capgemini.com/sites/InternationalPaper/r2a/_vti_bin/Lists.asmx"
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

Url = "https://troom-x.capgemini.com/sites/InternationalPaper/r2a/_vti_bin/Lists.asmx"
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

Url = "https://troom-x.capgemini.com/sites/InternationalPaper/r2a/_vti_bin/Lists.asmx"

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

Url = "https://troom-x.capgemini.com/sites/InternationalPaper/r2a/_vti_bin/Lists.asmx"

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
