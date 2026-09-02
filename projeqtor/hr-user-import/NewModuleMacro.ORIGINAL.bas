Attribute VB_Name = "NewModuleMacro"
Option Explicit

'==========================================================
' Generate ProjeQtOr CSVs
' - CREATE Resources (with linked User)
' - UPDATE Resources (by NAME, mandatory fields included)
' - Project Allocations
' - Folder picker to avoid path errors
'==========================================================
Public Sub GenerateProjeQtOrCSVs()

    Const SHEET_NAME As String = "Users Import"
    Const HOURS_PER_DAY As Double = 8#
    Const COLOR_NOFILL As Long = -4142

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long

    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then lastRow = 2

    '======================================================
    ' CSV HEADERS
    '======================================================
    Dim headerCreate As String
    Dim headerUpdate As String
    Dim headerAlloc As String

    ' CREATE Resource (creates linked User)
    headerCreate = "name;userName;email;initials;idTeam;startDate;maxWeeklyWork;capacity;idProfile;idRole;isUser;isResource;idCalendarDefinition;isLdap"

    ' UPDATE Resource (mandatory fields included)
    headerUpdate = "name;idTeam;capacity;maxWeeklyWork;idRole;idCalendarDefinition"

    ' Allocation
    headerAlloc = "Resource;idProject;rate;idProfile"

    Dim strNew As String, strUpdate As String, strAlloc As String
    strNew = headerCreate & vbCrLf
    strUpdate = headerUpdate & vbCrLf
    strAlloc = headerAlloc & vbCrLf

    Dim countNew As Long, countUpdate As Long, countAlloc As Long
    countNew = 0: countUpdate = 0: countAlloc = 0

    '======================================================
    ' PROJECT MAPS
    '======================================================
    Dim projs_AP As Variant, projs_Banking As Variant
    Dim projs_Q2C As Variant, projs_R2R As Variant

    projs_AP = Array(28, 3, 4, 5, 7, 10, 8, 6, 68, 70, 72, 73)
    projs_Banking = Array(29, 9, 31, 32, 33, 34, 35, 36, 37, 69, 14, 16, 38)
    projs_Q2C = Array(30, 18, 24, 25, 26, 27, 52, 71, 19, 17, 20, 21, 22, 23, 53)
    projs_R2R = Array(40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50)

    '======================================================
    ' MAIN LOOP
    '======================================================
    For i = 2 To lastRow

        If ws.Cells(i, 2).Interior.ColorIndex = COLOR_NOFILL Then

            Dim action As String
            Dim resourceName As String
            Dim userName As String
            Dim email As String
            Dim team As String
            Dim startDate As String

            action = Trim(CStr(ws.Cells(i, 1).Value))
            resourceName = StrConv(Trim(CStr(ws.Cells(i, 2).Value)), vbProperCase)
            userName = LCase(Trim(CStr(ws.Cells(i, 3).Value)))
            email = LCase(Trim(CStr(ws.Cells(i, 4).Value)))
            team = Trim(CStr(ws.Cells(i, 6).Value))
            startDate = Trim(ws.Cells(i, 7).Text)

            If action <> "" And resourceName <> "" Then

                Dim weeklyHours As Double
                Dim capacityVal As String
                Dim convertedDays As String

                weeklyHours = CDbl(ws.Cells(i, 8).Value)
                convertedDays = Replace(Format$(weeklyHours / HOURS_PER_DAY, "0.####"), ",", ".")
                capacityVal = Replace(Format$(CDbl(ws.Cells(i, 9).Value), "0.####"), ",", ".")

                '------------------------------
                ' CREATE RESOURCE
                '------------------------------
                If action = "Create" Then

                    Dim createFields As Variant
                    createFields = Array( _
                        resourceName, _
                        userName, _
                        email, _
                        userName, _
                        ws.Cells(i, 6).Value, _
                        startDate, _
                        convertedDays, _
                        capacityVal, _
                        ws.Cells(i, 10).Value, _
                        ws.Cells(i, 11).Value, _
                        ws.Cells(i, 12).Value, _
                        ws.Cells(i, 13).Value, _
                        ws.Cells(i, 14).Value, _
                        ws.Cells(i, 15).Value _
                    )

                    strNew = strNew & Join(createFields, ";") & vbCrLf
                    countNew = countNew + 1

                '------------------------------
                ' UPDATE RESOURCE (BY NAME)
                '------------------------------
                ElseIf action = "Update" Then

    ' UPDATE Resource (by name)
    ' name, idTeam, capacity, maxWeeklyWork, idRole, idCalendarDefinition

    Dim updateFields As Variant
    updateFields = Array( _
        resourceName, _
        ws.Cells(i, 6).Value, _
        capacityVal, _
        convertedDays, _
        ws.Cells(i, 11).Value, _
        ws.Cells(i, 14).Value _
    )

    strUpdate = strUpdate & Join(updateFields, ";") & vbCrLf
    countUpdate = countUpdate + 1

End If

                '------------------------------
                ' ALLOCATIONS
                '------------------------------
                If action = "Create" Or action = "Update" Then

                    Dim currentProjs As Variant
                    Dim hasProjs As Boolean
                    hasProjs = False

                    If team = "AP" Then
                        currentProjs = projs_AP: hasProjs = True
                    ElseIf team = "Banking" Then
                        currentProjs = projs_Banking: hasProjs = True
                    ElseIf team = "Q2C" Then
                        currentProjs = projs_Q2C: hasProjs = True
                    ElseIf team = "R2R" Then
                        currentProjs = projs_R2R: hasProjs = True
                    End If

                    If hasProjs Then
                        For j = LBound(currentProjs) To UBound(currentProjs)
                            strAlloc = strAlloc & resourceName & ";" & currentProjs(j) & ";100;4" & vbCrLf
                            countAlloc = countAlloc + 1
                        Next j
                    End If

                End If

            End If
        End If
    Next i

    '======================================================
    ' SELECT EXPORT FOLDER (OPTION A)
    '======================================================
    Dim fDialog As FileDialog
    Dim exportPath As String

    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With fDialog
        .Title = "Select folder to save ProjeQtOr CSV files"
        If .Show <> -1 Then
            MsgBox "Export cancelled.", vbInformation
            Exit Sub
        End If
        exportPath = .SelectedItems(1)
    End With

    '======================================================
    ' WRITE FILES
    '======================================================
    Dim fso As Object, fileOut As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If countNew > 0 Then
        Set fileOut = fso.CreateTextFile(exportPath & "\import_hr_new_resources.csv", True, False)
        fileOut.Write strNew
        fileOut.Close
    End If

    If countUpdate > 0 Then
        Set fileOut = fso.CreateTextFile(exportPath & "\import_hr_update_resources.csv", True, False)
        fileOut.Write strUpdate
        fileOut.Close
    End If

    If countAlloc > 0 Then
        Set fileOut = fso.CreateTextFile(exportPath & "\import_project_allocations.csv", True, False)
        fileOut.Write strAlloc
        fileOut.Close
    End If

    MsgBox "Export complete!" & vbCrLf & _
           "Created: " & countNew & vbCrLf & _
           "Updated: " & countUpdate & vbCrLf & _
           "Allocations: " & countAlloc, _
           vbInformation, "ProjeQtOr Export"

End Sub
