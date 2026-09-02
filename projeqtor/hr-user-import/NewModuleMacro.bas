Attribute VB_Name = "NewModuleMacro"
Option Explicit

'==========================================================
' Generate ProjeQtOr CSVs   (rev. 2026-08-19)
'
' WHAT CHANGED AND WHY -- read this before editing.
'
' 1. UPDATE ROWS NOW REQUIRE THE PROJEQTOR RESOURCE ID (column P).
'    ProjeQtor's import rule is: row has an `id` -> UPDATE the record;
'    no `id` -> INSERT a new one. The previous UPDATE file was keyed by
'    NAME with no id, so it did not update anything -- it created a
'    SECOND resource with the same name. A row with no id is now SKIPPED
'    and reported, never written.
'
' 2. THE UPDATE FILE NO LONGER SENDS `name`.
'    With the id present the name is not needed to find the record, and
'    sending it would RENAME the resource to whatever the sheet says.
'    That is a live hazard here: of 214 resources, 55 have a DOUBLE SPACE
'    in the real name ("Francisco  Manzanilla") which HTML collapses on
'    screen, so a retyped name looks identical and is not.
'
' 3. NAMES ARE NO LONGER BLINDLY PROPER-CASED.
'    StrConv(vbProperCase) turned "Silvia De la Fuente Peña" into
'    "De La", and "Cristina Garcia-Rojo Martorell" into "Garcia-rojo" --
'    11 of 214 names were being mangled. SmartName() now normalises only
'    input that arrives ALL CAPS or all lowercase (the HR-export case it
'    was there for) and leaves anything already mixed-case alone.
'
' 4. ALLOCATIONS ARE KEYED BY idResource WHERE THE ID IS KNOWN.
'    Update rows have an id, so their allocations use it. Create rows do
'    not exist yet, so theirs still go by name -- hence two files. Import
'    the by-name one straight after the resources, while the names still
'    match exactly what was just created.
'
' 5. ALLOCATION IMPORTS ALWAYS INSERT -- THERE IS NO UPSERT.
'    Re-running this export and re-importing DUPLICATES every allocation
'    it contains. Confirmed on this instance. Import each file once.
'
' 6. A TEAM CHANGE DOES NOT REMOVE THE OLD ALLOCATIONS.
'    Nothing in the import can delete. Fill column Q with the resource's
'    PREVIOUS team and this writes a checklist naming the allocations to
'    remove by hand. Note that closing an allocation also auto-closes
'    that resource's activity assignments on the project.
'
' 7. `idProfile` for allocations now comes from column J instead of being
'    hardcoded to 4, so a Project Leader (3) can be allocated.
'
' SHEET COLUMNS
'   A action (Create/Update)      I  capacity (FTE)
'   B resource name  <- the       J  idProfile
'     no-fill cell gates the row  K  idRole
'   C userName                    L  isUser
'   D email                       M  isResource
'   E (unused)                    N  idCalendarDefinition
'   F team                        O  isLdap
'   G startDate                   P  ProjeQtor resource id  <- NEW, required for Update
'   H weekly hours                Q  previous team          <- NEW, optional
'==========================================================
Public Sub GenerateProjeQtOrCSVs()

    Const SHEET_NAME As String = "Users Import"
    Const HOURS_PER_DAY As Double = 8#
    Const COLOR_NOFILL As Long = -4142
    Const ALLOC_RATE As String = "100"

    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, j As Long

    Dim headerCreate As String, headerUpdate As String
    Dim headerAllocById As String, headerAllocByName As String
    Dim strNew As String, strUpdate As String
    Dim strAllocById As String, strAllocByName As String, strRemove As String
    Dim countNew As Long, countUpdate As Long
    Dim countAllocById As Long, countAllocByName As Long
    Dim countSkipped As Long, countRemove As Long

    Dim projs_AP As Variant, projs_Banking As Variant
    Dim projs_Q2C As Variant, projs_R2R As Variant
    Dim currentProjs As Variant, oldProjs As Variant
    Dim hasProjs As Boolean, hasOld As Boolean

    Dim action As String, resourceName As String, userName As String
    Dim email As String, team As String, oldTeam As String
    Dim startDate As String, resourceId As String
    Dim weeklyHours As Double, capacityVal As String, convertedDays As String
    Dim allocProfile As String
    Dim createFields As Variant, updateFields As Variant
    Dim skippedList As String

    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then lastRow = 2

    '======================================================
    ' CSV HEADERS
    '======================================================
    headerCreate = "name;userName;email;initials;idTeam;startDate;maxWeeklyWork;capacity;idProfile;idRole;isUser;isResource;idCalendarDefinition;isLdap"

    ' id FIRST and no `name` -- see notes 1 and 2 at the top.
    headerUpdate = "id;idTeam;capacity;maxWeeklyWork;idRole;idCalendarDefinition"

    headerAllocById = "idResource;idProject;rate;idProfile"
    headerAllocByName = "Resource;idProject;rate;idProfile"

    strNew = headerCreate & vbCrLf
    strUpdate = headerUpdate & vbCrLf
    strAllocById = headerAllocById & vbCrLf
    strAllocByName = headerAllocByName & vbCrLf
    strRemove = "ALLOCATIONS TO REMOVE BY HAND (the import cannot delete)" & vbCrLf & _
                "Export element type Allocation to get each row's own id," & vbCrLf & _
                "then delete or close those rows in the UI." & vbCrLf & _
                "Closing an allocation also auto-closes that resource's" & vbCrLf & _
                "activity assignments on the project." & vbCrLf & vbCrLf

    countNew = 0: countUpdate = 0
    countAllocById = 0: countAllocByName = 0
    countSkipped = 0: countRemove = 0
    skippedList = ""

    '======================================================
    ' PROJECT MAPS
    ' Management (39 AP / 51 Banking / 54 Q2C / 55 R2R) is deliberately
    ' EXCLUDED -- regular users are not allocated to it.
    ' Banking gained 74 (PL & Others) and 75 (Marruecos) in the 2026-08
    ' restructure; they were missing here until this revision.
    '======================================================
    projs_AP = Array(28, 3, 4, 5, 7, 10, 8, 6, 68, 70, 72, 73)
    projs_Banking = Array(29, 9, 31, 32, 33, 34, 35, 36, 37, 69, 74, 75, 14, 16, 38)
    projs_Q2C = Array(30, 18, 24, 25, 26, 27, 52, 71, 19, 17, 20, 21, 22, 23, 53)
    projs_R2R = Array(40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50)

    '======================================================
    ' MAIN LOOP
    '======================================================
    For i = 2 To lastRow

        If ws.Cells(i, 2).Interior.ColorIndex = COLOR_NOFILL Then

            action = Trim(CStr(ws.Cells(i, 1).Value))
            resourceName = SmartName(Trim(CStr(ws.Cells(i, 2).Value)))
            userName = LCase(Trim(CStr(ws.Cells(i, 3).Value)))
            email = LCase(Trim(CStr(ws.Cells(i, 4).Value)))
            team = Trim(CStr(ws.Cells(i, 6).Value))
            startDate = Trim(ws.Cells(i, 7).Text)
            resourceId = Trim(CStr(ws.Cells(i, 16).Value))     ' column P
            oldTeam = Trim(CStr(ws.Cells(i, 17).Value))        ' column Q

            If action <> "" And resourceName <> "" Then

                weeklyHours = CDbl(ws.Cells(i, 8).Value)
                convertedDays = Replace(Format$(weeklyHours / HOURS_PER_DAY, "0.####"), ",", ".")
                capacityVal = Replace(Format$(CDbl(ws.Cells(i, 9).Value), "0.####"), ",", ".")

                allocProfile = Trim(CStr(ws.Cells(i, 10).Value))
                If allocProfile = "" Then allocProfile = "4"    ' 4 = Project Member

                hasProjs = False: hasOld = False
                If team = "AP" Then
                    currentProjs = projs_AP: hasProjs = True
                ElseIf team = "Banking" Then
                    currentProjs = projs_Banking: hasProjs = True
                ElseIf team = "Q2C" Then
                    currentProjs = projs_Q2C: hasProjs = True
                ElseIf team = "R2R" Then
                    currentProjs = projs_R2R: hasProjs = True
                End If

                '------------------------------
                ' CREATE RESOURCE
                '------------------------------
                If action = "Create" Then

                    createFields = Array( _
                        resourceName, userName, email, userName, _
                        ws.Cells(i, 6).Value, startDate, _
                        convertedDays, capacityVal, _
                        ws.Cells(i, 10).Value, ws.Cells(i, 11).Value, _
                        ws.Cells(i, 12).Value, ws.Cells(i, 13).Value, _
                        ws.Cells(i, 14).Value, ws.Cells(i, 15).Value)

                    strNew = strNew & Join(createFields, ";") & vbCrLf
                    countNew = countNew + 1

                    ' No id yet -- allocations must go by name, and the name
                    ' must match what the create file just wrote.
                    If hasProjs Then
                        For j = LBound(currentProjs) To UBound(currentProjs)
                            strAllocByName = strAllocByName & resourceName & ";" & _
                                currentProjs(j) & ";" & ALLOC_RATE & ";" & allocProfile & vbCrLf
                            countAllocByName = countAllocByName + 1
                        Next j
                    End If

                '------------------------------
                ' UPDATE RESOURCE (BY ID)
                '------------------------------
                ElseIf action = "Update" Then

                    If resourceId = "" Then
                        ' Without an id this row would INSERT a duplicate
                        ' resource instead of updating. Refuse it.
                        countSkipped = countSkipped + 1
                        skippedList = skippedList & "  row " & i & ": " & resourceName & vbCrLf
                    Else
                        updateFields = Array( _
                            resourceId, _
                            ws.Cells(i, 6).Value, _
                            capacityVal, _
                            convertedDays, _
                            ws.Cells(i, 11).Value, _
                            ws.Cells(i, 14).Value)

                        strUpdate = strUpdate & Join(updateFields, ";") & vbCrLf
                        countUpdate = countUpdate + 1

                        If hasProjs Then
                            For j = LBound(currentProjs) To UBound(currentProjs)
                                strAllocById = strAllocById & resourceId & ";" & _
                                    currentProjs(j) & ";" & ALLOC_RATE & ";" & allocProfile & vbCrLf
                                countAllocById = countAllocById + 1
                            Next j
                        End If

                        ' Team change -> the OLD allocations are still there.
                        If oldTeam <> "" And oldTeam <> team Then
                            If oldTeam = "AP" Then
                                oldProjs = projs_AP: hasOld = True
                            ElseIf oldTeam = "Banking" Then
                                oldProjs = projs_Banking: hasOld = True
                            ElseIf oldTeam = "Q2C" Then
                                oldProjs = projs_Q2C: hasOld = True
                            ElseIf oldTeam = "R2R" Then
                                oldProjs = projs_R2R: hasOld = True
                            End If
                            If hasOld Then
                                strRemove = strRemove & "Resource #" & resourceId & " " & _
                                    resourceName & "  --  moved " & oldTeam & " -> " & team & vbCrLf & _
                                    "  remove its allocations on projects: "
                                For j = LBound(oldProjs) To UBound(oldProjs)
                                    strRemove = strRemove & oldProjs(j)
                                    If j < UBound(oldProjs) Then strRemove = strRemove & ", "
                                    countRemove = countRemove + 1
                                Next j
                                strRemove = strRemove & vbCrLf & vbCrLf
                            End If
                        End If
                    End If

                End If

            End If
        End If
    Next i

    '======================================================
    ' SELECT EXPORT FOLDER
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
        fileOut.Write strNew: fileOut.Close
    End If

    If countUpdate > 0 Then
        Set fileOut = fso.CreateTextFile(exportPath & "\import_hr_update_resources.csv", True, False)
        fileOut.Write strUpdate: fileOut.Close
    End If

    If countAllocById > 0 Then
        Set fileOut = fso.CreateTextFile(exportPath & "\import_allocations_by_id.csv", True, False)
        fileOut.Write strAllocById: fileOut.Close
    End If

    If countAllocByName > 0 Then
        Set fileOut = fso.CreateTextFile(exportPath & "\import_allocations_by_name.csv", True, False)
        fileOut.Write strAllocByName: fileOut.Close
    End If

    If countRemove > 0 Then
        Set fileOut = fso.CreateTextFile(exportPath & "\CHECKLIST_allocations_to_remove.txt", True, False)
        fileOut.Write strRemove: fileOut.Close
    End If

    '======================================================
    ' REPORT
    '======================================================
    Dim msg As String
    msg = "Export complete!" & vbCrLf & vbCrLf & _
          "Created resources:     " & countNew & vbCrLf & _
          "Updated resources:     " & countUpdate & vbCrLf & _
          "Allocations (by id):   " & countAllocById & vbCrLf & _
          "Allocations (by name): " & countAllocByName & vbCrLf

    If countRemove > 0 Then
        msg = msg & "Allocations to REMOVE: " & countRemove & _
              "  (see CHECKLIST_allocations_to_remove.txt)" & vbCrLf
    End If

    If countSkipped > 0 Then
        msg = msg & vbCrLf & "SKIPPED " & countSkipped & " Update row(s) with no " & _
              "ProjeQtor id in column P." & vbCrLf & _
              "Without an id the import would CREATE A DUPLICATE resource " & _
              "instead of updating." & vbCrLf & _
              "Get the ids from an Encoding > Resources export." & vbCrLf & vbCrLf & _
              skippedList
    End If

    msg = msg & vbCrLf & "Import each allocation file ONCE -- allocation rows " & _
          "always INSERT, so re-importing duplicates them."

    MsgBox msg, vbInformation, "ProjeQtOr Export"

End Sub


'==========================================================
' Normalise a name ONLY when it arrives in a single case.
'
' StrConv(vbProperCase) was being applied unconditionally, which
' mangled 11 of the 214 real names on this instance:
'   "Silvia De la Fuente Peña"      -> "Silvia De La Fuente Peña"
'   "Cristina Garcia-Rojo Martorell"-> "Cristina Garcia-rojo Martorell"
'   "Alvaro Prieto de la Torre"     -> "Alvaro Prieto De La Torre"
' Anything already mixed-case is the HR system's own spelling and is
' left exactly as typed.
'==========================================================
Private Function SmartName(ByVal s As String) As String
    If s = "" Then
        SmartName = s
    ElseIf s = UCase$(s) Or s = LCase$(s) Then
        SmartName = StrConv(s, vbProperCase)
    Else
        SmartName = s
    End If
End Function
