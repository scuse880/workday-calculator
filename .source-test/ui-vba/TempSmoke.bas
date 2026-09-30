Attribute VB_Name = "TempSmoke"
Option Explicit

Private Function EvidenceSheet() As Worksheet
    On Error Resume Next
    Set EvidenceSheet = ThisWorkbook.Worksheets("TestEvidence")
    On Error GoTo 0
    If EvidenceSheet Is Nothing Then
        Set EvidenceSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EvidenceSheet.Name = "TestEvidence"
    End If
End Function

Private Function CountRecords(ByVal recordKind As String) As Long
    Dim item As Variant, employee As CEmployee, total As Long
    For Each item In GetEmployeesInOrder()
        Set employee = item
        Select Case recordKind
            Case "WorkStatus": total = total + employee.WorkStatusRecords.Count
            Case "Trip": total = total + employee.TripRecords.Count
            Case "VacationWork": total = total + employee.VacationWorkRecords.Count
        End Select
    Next item
    CountRecords = total
End Function

Public Sub SeedEmployees(ByVal sourcePath As String)
    Dim book As Workbook, sheet As Worksheet
    Dim staged As Object, group As Collection, employee As CEmployee
    Dim employeeName As String, birthdate As String
    Dim r As Long, count As Long, item As Variant

    Set book = Application.Workbooks.Open(sourcePath, 0, True)
    Set sheet = book.Worksheets(1)
    Set staged = NewDictionary()
    r = 5
    Do While r <= sheet.Rows.Count
        If Len(CellText(sheet.Cells(r, 3))) = 0 Then
            r = r + 3
            If Len(CellText(sheet.Cells(r, 3))) = 0 Then Exit Do
        End If
        ParseNameBirth CellText(sheet.Cells(r, 3)), employeeName, birthdate, "seed"
        Set employee = New CEmployee
        count = count + 1
        employee.Order = count
        employee.Name = employeeName
        employee.Birthdate = birthdate
        If Not staged.Exists(employeeName) Then
            Set group = New Collection
            staged.Add employeeName, group
        End If
        Set group = staged(employeeName)
        group.Add employee
        r = r + 1
    Loop
    book.Close False
    For Each item In GetEmployeesInOrder(staged)
        Set employee = item
        Set group = staged(employee.Name)
        If group.Count > 1 Then employee.NeisPersonId = "C" & Right$("000000000" & CStr(employee.Order), 9)
    Next item
    Set gEmployeesByName = staged
    gEmployeesMonth = DateSerial(2026, 8, 1)
    gWorkStatusLoaded = False
    gTripLoaded = False
    gVacationWorkLoaded = False
End Sub

Public Sub TestRealSourceLoaders(ByVal sourceFolder As String)
    Dim evidence As Worksheet, employeeCount As Long
    Dim workStatusCount As Long, tripCount As Long, vacationCount As Long
    Dim item As Variant
    On Error GoTo Failed
    Set evidence = EvidenceSheet()
    evidence.Cells.Clear
    evidence.Range("A1:B1").Value = Array("check", "result")

    gTestSourcePath = sourceFolder & "초과근무 월별집계(생년월일 미표시).xlsx"
    LoadOvertimeEmployees
    evidence.Range("A2").Value = "missing_birth_rejected"
    evidence.Range("B2").Value = (gEmployeesByName Is Nothing)
    If Not gEmployeesByName Is Nothing Then Err.Raise vbObjectError + 901, "TempSmoke", "No-birth input unexpectedly accepted"

    SeedEmployees sourceFolder & "초과근무 월별집계(생년월일 표시).xlsx"
    For Each item In GetEmployeesInOrder()
        employeeCount = employeeCount + 1
    Next item
    evidence.Range("A3").Value = "seed_employee_count"
    evidence.Range("B3").Value = employeeCount
    If employeeCount <> 77 Then Err.Raise vbObjectError + 902, "TempSmoke", "Employee count mismatch"

    gTestSourcePath = sourceFolder & "근무상황목록.xlsx"
    LoadWorkStatusRecords
    workStatusCount = CountRecords("WorkStatus")
    evidence.Range("A4").Value = "work_status_loaded"
    evidence.Range("B4").Value = gWorkStatusLoaded
    evidence.Range("A5").Value = "work_status_records"
    evidence.Range("B5").Value = workStatusCount
    If Not gWorkStatusLoaded Or workStatusCount <> 200 Then Err.Raise vbObjectError + 903, "TempSmoke", "Work status count mismatch"

    gTestSourcePath = sourceFolder & "출장 근무상황부 개인별.xlsx"
    gTestAutoSelect = True
    LoadTripRecords
    tripCount = CountRecords("Trip")
    evidence.Range("A6").Value = "trip_loaded"
    evidence.Range("B6").Value = gTripLoaded
    evidence.Range("A7").Value = "trip_records"
    evidence.Range("B7").Value = tripCount
    If Not gTripLoaded Or tripCount <> 110 Then Err.Raise vbObjectError + 904, "TempSmoke", "Trip count mismatch"

    gTestSourcePath = vbNullString
    LoadVacationWorkRecords
    vacationCount = CountRecords("VacationWork")
    evidence.Range("A8").Value = "vacation_loaded"
    evidence.Range("B8").Value = gVacationWorkLoaded
    evidence.Range("A9").Value = "vacation_records"
    evidence.Range("B9").Value = vacationCount
    If Not gVacationWorkLoaded Or vacationCount <> 12 Then Err.Raise vbObjectError + 905, "TempSmoke", "Vacation count mismatch"

    evidence.Range("A10").Value = "result"
    evidence.Range("B10").Value = "PASS"
    Exit Sub
Failed:
    evidence.Range("A10").Value = "result"
    evidence.Range("B10").Value = "FAIL " & CStr(Err.Number) & " " & Err.Description
    Err.Raise Err.Number, "TempSmoke", Err.Description
End Sub
