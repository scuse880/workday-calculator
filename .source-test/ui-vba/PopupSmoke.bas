Attribute VB_Name = "PopupSmoke"
Option Explicit

Private Function LogSheet() As Worksheet
    On Error Resume Next
    Set LogSheet = ThisWorkbook.Worksheets("PopupEvidence")
    On Error GoTo 0
    If LogSheet Is Nothing Then
        Set LogSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        LogSheet.Name = "PopupEvidence"
    End If
End Function

Public Sub TestMissingBirthPopup(ByVal sourceFolder As String)
    Dim ws As Worksheet
    Set ws = LogSheet()
    gTestSourcePath = sourceFolder & "초과근무 월별집계(생년월일 미표시).xlsx"
    LoadOvertimeEmployees
    ws.Range("A1").Value = "missing_birth_preserved"
    ws.Range("B1").Value = (gEmployeesByName Is Nothing)
End Sub

Public Sub TestWorkStatusWarning(ByVal sourceFolder As String)
    Dim ws As Worksheet
    Set ws = LogSheet()
    SeedEmployees sourceFolder & "초과근무 월별집계(생년월일 표시).xlsx"
    gTestSourcePath = sourceFolder & "근무상황목록.xlsx"
    LoadWorkStatusRecords
    ws.Range("A2").Value = "work_status_loaded"
    ws.Range("B2").Value = gWorkStatusLoaded
End Sub

Public Sub TestTripFormCancelAndSelect(ByVal sourceFolder As String)
    Dim ws As Worksheet, key As Variant, group As Collection
    Dim selector As frmTripEmployee
    Set ws = LogSheet()
    SeedEmployees sourceFolder & "초과근무 월별집계(생년월일 표시).xlsx"
    For Each key In gEmployeesByName.Keys
        Set group = gEmployeesByName(key)
        If group.Count > 1 Then Exit For
    Next key
    Set selector = New frmTripEmployee
    selector.Configure group, "source", CStr(key), "abc", 5
    selector.Show vbModeless
    ws.Range("A3").Value = "trip_form_cancel_initial_visible"
    ws.Range("B3").Value = selector.Visible
    selector.TestClickCancel
    ws.Range("A4").Value = "trip_form_cancelled"
    ws.Range("B4").Value = selector.Cancelled
    ws.Range("A5").Value = "trip_form_cancel_hidden"
    ws.Range("B5").Value = Not selector.Visible
    Unload selector
    Set selector = New frmTripEmployee
    selector.Configure group, "source", CStr(key), "abc", 5
    selector.Show vbModeless
    selector.AutoSelectFirst
    ws.Range("A6").Value = "trip_form_selected"
    ws.Range("B6").Value = Not selector.Cancelled And Not selector.SelectedEmployee Is Nothing
    ws.Range("A7").Value = "trip_form_select_hidden"
    ws.Range("B7").Value = Not selector.Visible
    Unload selector
End Sub

Public Sub TestCalculationPopups(ByVal sourceFolder As String)
    Dim ws As Worksheet, item As Variant, employee As CEmployee
    Dim trip As CTripRecord, vacation As CVacationWorkRecord
    Set ws = LogSheet()
    SeedEmployees sourceFolder & "초과근무 월별집계(생년월일 표시).xlsx"
    gTestSourcePath = sourceFolder & "근무상황목록.xlsx"
    LoadWorkStatusRecords
    gTestSourcePath = sourceFolder & "출장 근무상황부 개인별.xlsx"
    gTestAutoSelect = True
    LoadTripRecords
    gTestSourcePath = vbNullString
    LoadVacationWorkRecords
    SaveHolidays DateSerial(2026, 8, 1), NewDictionary()
    근무일수계산
    ws.Range("A8").Value = "actual_calculation_output"
    ws.Range("B8").Value = ThisWorkbook.Worksheets("작업결과").Cells(1, 1).Value2
    For Each item In GetEmployeesInOrder()
        Set employee = item
        Exit For
    Next item
    Set trip = New CTripRecord
    trip.StartAt = DateSerial(2026, 8, 3) + TimeSerial(10, 0, 0)
    trip.EndAt = DateSerial(2026, 8, 3) + TimeSerial(11, 0, 0)
    employee.TripRecords.Add trip
    Set vacation = New CVacationWorkRecord
    vacation.StartAt = DateSerial(2026, 8, 3) + TimeSerial(12, 0, 0)
    vacation.EndAt = DateSerial(2026, 8, 3) + TimeSerial(13, 0, 0)
    employee.VacationWorkRecords.Add vacation
    근무일수계산
    ws.Range("A9").Value = "forced_pending_output"
    ws.Range("B9").Value = ThisWorkbook.Worksheets("작업결과").Cells(1, 1).Value2
End Sub
