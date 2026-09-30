Attribute VB_Name = "main_인사변동자등록"
Option Explicit

Public Sub 인사변동자등록()
    Dim workMonth As Date, arrivalMinutes As Long, departureMinutes As Long
    Dim employees As Collection, dialog As frmPersonnelChange
    Dim failure As String

    On Error GoTo Failed
    ReadBaseSettings workMonth, arrivalMinutes, departureMinutes, False
    RequireEmployeeState workMonth
    Set employees = GetEmployeesInOrder()
    Set dialog = New frmPersonnelChange
    dialog.Configure workMonth, employees
    dialog.Show vbModal
CleanUp:
    On Error Resume Next
    If Not dialog Is Nothing Then
        dialog.DisposeHandlers
        Unload dialog
    End If
    Set dialog = Nothing
    On Error GoTo 0
    Exit Sub
Failed:
    failure = Err.Description
    Debug.Print failure, vbExclamation, "인사변동자 등록"
    Resume CleanUp
End Sub

Public Sub SavePersonnelChange(ByVal workMonth As Date, ByVal employeeOrder As Long, _
                               ByVal startDay As Date, ByVal endDay As Date)
    Dim employee As CEmployee

    RequireEmployeeState workMonth
    If startDay = 0 Or endDay = 0 Then
        RaiseValidation "인사변동 기간", "시작일과 종료일을 차례로 선택하세요."
    End If
    If CDbl(startDay) <> Fix(CDbl(startDay)) Or CDbl(endDay) <> Fix(CDbl(endDay)) Then
        RaiseValidation "인사변동 기간", "날짜에는 시각을 포함할 수 없습니다."
    End If
    If Year(startDay) <> Year(workMonth) Or Month(startDay) <> Month(workMonth) _
       Or Year(endDay) <> Year(workMonth) Or Month(endDay) <> Month(workMonth) Then
        RaiseValidation "인사변동 기간", "현재 작업년월의 날짜만 선택하세요."
    End If
    If startDay >= endDay Then
        RaiseValidation "인사변동 기간", "종료일은 시작일보다 늦어야 합니다."
    End If
    Set employee = PersonnelEmployee(employeeOrder)
    ' 결과를 지울 수 없는 경우 등록 상태도 그대로 둔다.
    ClearCalculationResults
    employee.PersonnelChangeStart = startDay
    employee.PersonnelChangeEnd = endDay
End Sub

Public Sub ClearPersonnelChange(ByVal workMonth As Date, ByVal employeeOrder As Long)
    Dim employee As CEmployee

    RequireEmployeeState workMonth
    Set employee = PersonnelEmployee(employeeOrder)
    If employee.PersonnelChangeStart = 0 And employee.PersonnelChangeEnd = 0 Then Exit Sub
    ClearCalculationResults
    employee.PersonnelChangeStart = 0
    employee.PersonnelChangeEnd = 0
End Sub

Public Function IsPersonnelChangeDay(ByVal employee As CEmployee, ByVal dayValue As Date) As Boolean
    If employee.PersonnelChangeStart = 0 Or employee.PersonnelChangeEnd = 0 Then Exit Function
    IsPersonnelChangeDay = (dayValue >= employee.PersonnelChangeStart And dayValue <= employee.PersonnelChangeEnd)
End Function

Private Function PersonnelEmployee(ByVal employeeOrder As Long) As CEmployee
    Dim employees As Collection
    Set employees = GetEmployeesInOrder()
    If employeeOrder < 1 Or employeeOrder > employees.Count Then
        RaiseValidation "인사변동 기간", "선택한 교직원을 확인할 수 없습니다. 대상자를 다시 불러오세요."
    End If
    Set PersonnelEmployee = employees(employeeOrder)
End Function
