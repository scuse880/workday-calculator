Attribute VB_Name = "main_방학중근무일지"
Option Explicit

' Button entry point. The input sheet is read only; all records are staged.
Public Sub LoadVacationWorkRecords()
    Dim workMonth As Date, arrivalMinutes As Long, departureMinutes As Long
    Dim sourceSheet As Worksheet, staged As Object, records As Collection
    Dim employee As CEmployee, record As CVacationWorkRecord
    Dim r As Long, recordCount As Long
    Dim nameText As String, employeeName As String, birthdate As String
    Dim workDate As Date, startTime As Date, endTime As Date
    Dim errorText As String

    On Error GoTo Failed
    ReadBaseSettings workMonth, arrivalMinutes, departureMinutes, False
    RequireEmployeeState workMonth
    Set sourceSheet = ThisWorkbook.Worksheets("방학중근무")
    Set staged = NewRecordStage()

    r = 3
    Do While r <= sourceSheet.Rows.Count
        nameText = CellText(sourceSheet.Cells(r, "A"))
        If Len(nameText) = 0 Then Exit Do
        ParseNameBirth nameText, employeeName, birthdate, sourceSheet.Name & "!A" & CStr(r)
        Set employee = ResolveVacationEmployee(employeeName, birthdate, r)
        workDate = ParseDateCell(sourceSheet.Cells(r, "B"))
        If Year(workDate) <> Year(workMonth) Or Month(workDate) <> Month(workMonth) Then
            RaiseValidation sourceSheet.Name & "!B" & CStr(r), "날짜가 현재 작업년월에 속해야 합니다."
        End If
        startTime = ParseTimeCell(sourceSheet.Cells(r, "C"))
        endTime = ParseTimeCell(sourceSheet.Cells(r, "D"))
        If startTime >= endTime Then
            RaiseValidation "방학중근무 " & CStr(r) & "행 (C:D)", _
                            "시작 시각은 종료 시각보다 빨라야 하며 날짜를 넘길 수 없습니다."
        End If

        Set record = New CVacationWorkRecord
        record.StartAt = workDate + startTime
        record.EndAt = workDate + endTime
        Set records = staged(CStr(employee.Order))
        records.Add record
        recordCount = recordCount + 1
        r = r + 1
    Loop

    CommitRecordStage staged, "VacationWork"
    Debug.Print "방학중근무 자료 " & CStr(recordCount) & "건을 불러왔습니다.", vbInformation, "방학중근무"
    Exit Sub

Failed:
    errorText = Err.Description
    If r >= 3 Then errorText = "방학중근무 " & CStr(r) & "행: " & errorText
    Debug.Print errorText & vbCrLf & "기존 방학중근무 자료는 유지됩니다.", vbExclamation, "방학중근무 불러오기 오류"
End Sub

Private Function ResolveVacationEmployee(ByVal employeeName As String, ByVal birthdate As String, _
                                         ByVal sourceRow As Long) As CEmployee
    Dim candidates As Collection, employee As CEmployee, matched As CEmployee
    Dim context As String
    context = "방학중근무!A" & CStr(sourceRow)
    If Not gEmployeesByName.Exists(employeeName) Then
        RaiseValidation context, employeeName & "은(는) 근무일수 산정 대상 목록에 없습니다."
    End If
    Set candidates = gEmployeesByName(employeeName)
    If candidates.Count = 1 Then
        Set employee = candidates(1)
        If Len(birthdate) > 0 And Len(employee.Birthdate) > 0 Then
            If birthdate <> employee.Birthdate Then
                RaiseValidation context, "입력한 생년월일이 등록된 대상자의 생년월일과 다릅니다."
            End If
        End If
        Set ResolveVacationEmployee = employee
        Exit Function
    End If

    If Len(birthdate) = 0 Then
        RaiseValidation context, "동명이인인 경우 성명 열에 ""이름(생년월일)"" 형식으로 입력하세요."
    End If
    For Each employee In candidates
        If employee.Birthdate = birthdate Then
            If Not matched Is Nothing Then
                RaiseValidation context, "같은 이름과 생년월일의 대상자가 여러 명입니다. 대상자를 다시 불러오세요."
            End If
            Set matched = employee
        End If
    Next employee
    If matched Is Nothing Then RaiseValidation context, "동명이인의 생년월일을 확인하세요."
    Set ResolveVacationEmployee = matched
End Function
