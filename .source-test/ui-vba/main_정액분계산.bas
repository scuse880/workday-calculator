Attribute VB_Name = "main_정액분계산"
Option Explicit

' 결과: 1=근무일, 0=근무일 아님, -1=판정 보류. 휴일은 판정 자체를 하지 않는다.
Public Sub 정액분계산()
    근무일수계산
End Sub

Public Sub 근무일수계산()
    Dim workMonth As Date, arrivalMinutes As Long, departureMinutes As Long
    Dim holidays As Object, saved As Boolean, employees As Collection, employee As CEmployee
    Dim dailyResults As Object, result As Variant, x As Variant, y As Variant, z As Variant
    Dim decision As Variant, outputValues() As Variant, outputColors() As Long
    Dim dayValue As Date, numberOfDays As Long, i As Long, d As Long, total As Long
    Dim pendingPeople As Long, pendingDays As Long, personPending As Boolean
    Dim cellTextValue As String, rowName As String, group As Collection
    Dim oldScreenUpdating As Boolean, screenChanged As Boolean
    On Error GoTo Failed
    ReadBaseSettings workMonth, arrivalMinutes, departureMinutes
    RequireEmployeeState workMonth
    If Not gWorkStatusLoaded Then RaiseValidation "근무일수 계산", "근무상황목록을 먼저 불러오세요."
    If Not gTripLoaded Then RaiseValidation "근무일수 계산", "출장 근무상황부를 먼저 불러오세요."
    Set holidays = LoadHolidays(workMonth, saved)
    If Not saved Then
        RaiseValidation "근무일수 계산", "현재 작업년월의 공휴일을 확인한 후 저장 및 닫기를 눌러주세요."
    End If
    Set employees = GetEmployeesInOrder()
    Set dailyResults = NewDictionary()
    numberOfDays = Day(DateSerial(Year(workMonth), Month(workMonth) + 1, 0))
    ReDim outputValues(1 To employees.Count + 1, 1 To numberOfDays + 2)
    ReDim outputColors(1 To employees.Count, 1 To numberOfDays)
    outputValues(1, 1) = Space$(22) & "날짜" & vbLf & "성명"
    outputValues(1, numberOfDays + 2) = "근무일수"
    For d = 1 To numberOfDays
        dayValue = DateSerial(Year(workMonth), Month(workMonth), d)
        outputValues(1, d + 1) = CStr(Month(dayValue)) & "/" & Right$("0" & CStr(Day(dayValue)), 2) & _
                                 "(" & KoreanWeekday(dayValue) & ")"
    Next d
    ' 모든 계산과 표시문자열을 먼저 완성한다. 계산 실패 시 이전 출력은 유지된다.
    For i = 1 To employees.Count
        Set employee = employees(i)
        Set group = gEmployeesByName(employee.Name)
        rowName = employee.Name
        If group.Count > 1 Then rowName = rowName & "(" & employee.Birthdate & ")"
        outputValues(i + 1, 1) = rowName
        total = 0
        personPending = False
        For d = 1 To numberOfDays
            outputColors(i, d) = -1
            dayValue = DateSerial(Year(workMonth), Month(workMonth), d)
            If Weekday(dayValue, vbMonday) <= 5 And Not holidays.Exists(HolidayKey(dayValue)) _
               And Not IsPersonnelChangeDay(employee, dayValue) Then
                x = RecordDayIntervals(employee.WorkStatusRecords, "X", dayValue, arrivalMinutes, departureMinutes)
                y = RecordDayIntervals(employee.TripRecords, "Y", dayValue, arrivalMinutes, departureMinutes)
                If gVacationWorkLoaded Then
                    z = RecordDayIntervals(employee.VacationWorkRecords, "Z", dayValue, arrivalMinutes, departureMinutes)
                Else
                    z = Empty
                End If
                decision = JudgeIntervals(x, y, z, arrivalMinutes, departureMinutes)
                cellTextValue = DayResultText(x, y, z, CLng(decision(0)))
                dailyResults.Add CStr(i) & ":" & CStr(d), Array(x, y, z, decision, cellTextValue)
                If decision(0) = 1 Then total = total + 1
                If decision(0) = -1 Then
                    personPending = True
                    pendingDays = pendingDays + 1
                End If
            End If
        Next d
        If personPending Then pendingPeople = pendingPeople + 1
        outputValues(i + 1, numberOfDays + 2) = total
    Next i
    ' 저장한 판정 결과를 그대로 사용한다. 출력 단계에서 재판정하지 않는다.
    For i = 1 To employees.Count
        For d = 1 To numberOfDays
            If dailyResults.Exists(CStr(i) & ":" & CStr(d)) Then
                result = dailyResults(CStr(i) & ":" & CStr(d))
                decision = result(3)
                outputValues(i + 1, d + 1) = result(4)
                outputColors(i, d) = CLng(decision(2))
            End If
        Next d
    Next i
    oldScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    screenChanged = True
    WriteCalculationResults workMonth, holidays, outputValues, outputColors, employees.Count, numberOfDays
    Application.ScreenUpdating = oldScreenUpdating
    screenChanged = False
    If pendingDays > 0 Then
        MsgBox "계산이 완료되었습니다." & vbCrLf & "판정 보류 인원: " & pendingPeople & "명" & vbCrLf & _
               "날짜별 보류 건수: " & pendingDays & "건 (교직원 1명·날짜 1개 기준)" & vbCrLf & _
               "판정 보류는 근무일수 합계에 포함하지 않았습니다.", vbInformation, "근무일수 계산"
    Else
        MsgBox "근무일수 계산이 완료되었습니다.", vbInformation, "근무일수 계산"
    End If
    Exit Sub
Failed:
    If screenChanged Then Application.ScreenUpdating = oldScreenUpdating
    MsgBox Err.Description, vbExclamation, "근무일수 계산"
End Sub

Private Sub WriteCalculationResults(ByVal workMonth As Date, ByVal holidays As Object, _
                                    ByRef values() As Variant, ByRef colors() As Long, _
                                    ByVal employeeCount As Long, ByVal numberOfDays As Long)
    Dim ws As Worksheet, table As Range, cell As Range
    Dim i As Long, d As Long, noteRow As Long, dayValue As Date, headerText As String
    Set ws = GetOrCreateSheet("작업결과")
    ws.UsedRange.UnMerge
    ws.UsedRange.Clear
    Set table = ws.Range(ws.Cells(1, 1), ws.Cells(employeeCount + 1, numberOfDays + 2))
    table.Value2 = values
    With table
        .Font.Name = "맑은 고딕"
        .Font.Size = 9
        .Font.Color = RGB(0, 0, 0)
        .Interior.Pattern = xlNone
        .WrapText = True
        .VerticalAlignment = xlTop
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(190, 190, 190)
    End With
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, numberOfDays + 2))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(230, 233, 237)
        .RowHeight = 36
    End With
    With ws.Cells(1, 1)
        .HorizontalAlignment = xlLeft
        .Borders(xlDiagonalUp).LineStyle = xlNone
        With .Borders(xlDiagonalDown)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(120, 130, 145)
        End With
    End With
    ws.Columns(1).ColumnWidth = 19
    ws.Range(ws.Columns(2), ws.Columns(numberOfDays + 1)).ColumnWidth = 23
    ws.Columns(numberOfDays + 2).ColumnWidth = 11
    For d = 1 To numberOfDays
        dayValue = DateSerial(Year(workMonth), Month(workMonth), d)
        Set cell = ws.Cells(1, d + 1)
        headerText = CStr(values(1, d + 1))
        If Weekday(dayValue, vbSunday) = 1 Or holidays.Exists(HolidayKey(dayValue)) Then
            cell.Characters(Len(headerText) - 1, 1).Font.Color = RGB(255, 0, 0)
        ElseIf Weekday(dayValue, vbSunday) = 7 Then
            cell.Characters(Len(headerText) - 1, 1).Font.Color = RGB(0, 0, 255)
        End If
        For i = 1 To employeeCount
            If colors(i, d) <> -1 Then ws.Cells(i + 1, d + 1).Interior.Color = colors(i, d)
        Next i
    Next d
    With ws.Range(ws.Cells(2, numberOfDays + 2), ws.Cells(employeeCount + 1, numberOfDays + 2))
        .NumberFormat = "0"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    ws.Range(ws.Rows(2), ws.Rows(employeeCount + 1)).AutoFit
    noteRow = employeeCount + 3
    ws.Cells(noteRow, 1).Value2 = "색상 안내"
    ws.Cells(noteRow, 1).Font.Bold = True
    ws.Cells(noteRow, 2).Value2 = "노랑: 기본 근무 / 출장"
    ws.Cells(noteRow, 2).Interior.Color = RGB(255, 255, 0)
    ws.Cells(noteRow + 1, 2).Value2 = "주황: 방학 근무 / 복무 구간 충족"
    ws.Cells(noteRow + 1, 2).Interior.Color = RGB(255, 192, 0)
    ws.Cells(noteRow + 2, 2).Value2 = "회색: 판정 보류"
    ws.Cells(noteRow + 2, 2).Interior.Color = RGB(217, 217, 217)
    ws.Cells(noteRow + 3, 2).Value2 = "채우기 없음: 근무일수 아님 / 주말·공휴일"
    ws.Cells(noteRow + 4, 2).Value2 = "판정 보류는 근무일수 합계에 포함하지 않음"
    With ws.Range(ws.Cells(noteRow, 1), ws.Cells(noteRow + 4, 4))
        .Font.Name = "맑은 고딕"
        .Font.Size = 9
        .VerticalAlignment = xlTop
    End With
    ws.PageSetup.PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(noteRow + 4, numberOfDays + 2)).Address
    ws.PageSetup.Orientation = xlLandscape
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = False
    ws.PageSetup.PrintTitleRows = "$1:$1"
    ws.Activate
    With ActiveWindow
        .FreezePanes = False
        .SplitColumn = 1
        .SplitRow = 1
        .FreezePanes = True
    End With
End Sub

Private Function RecordDayIntervals(ByVal records As Collection, ByVal recordKind As String, _
                                    ByVal dayValue As Date, ByVal arrivalMinutes As Long, _
                                    ByVal departureMinutes As Long) As Variant
    Dim record As Object, raw As Variant, span As Variant
    For Each record In records
        span = RecordDaySpan(record.StartAt, record.EndAt, recordKind, dayValue, arrivalMinutes, departureMinutes)
        AppendInterval raw, span
    Next record
    RecordDayIntervals = MergeIntervals(raw)
End Function

' BEGIN PURE INTERVAL LOGIC -- tests/test_calculation.vbs.py executes this exact source.
' Intervals are Empty or Array(Array(startMinute, endMinute), ...), all half-open.
Public Function RecordDaySpan(ByVal startAt As Date, ByVal endAt As Date, ByVal recordKind As String, _
                              ByVal dayValue As Date, ByVal arrivalMinutes As Long, _
                              ByVal departureMinutes As Long) As Variant
    Dim dayStart As Double, dayEnd As Double, spanStart As Double, spanEnd As Double
    Dim recordStartDay As Double, recordEndDay As Double, minuteStart As Long, minuteEnd As Long
    dayStart = Int(CDbl(dayValue))
    dayEnd = dayStart + 1
    If startAt >= endAt Then Err.Raise 5, "RecordDaySpan", "기록의 시작은 종료보다 빨라야 합니다."
    If CDbl(startAt) >= dayEnd Or CDbl(endAt) <= dayStart Then Exit Function
    recordStartDay = Int(CDbl(startAt))
    recordEndDay = Int(CDbl(endAt))
    spanStart = CDbl(startAt)
    spanEnd = CDbl(endAt)
    If recordKind <> "Z" And recordStartDay <> recordEndDay Then
        If dayStart = recordStartDay Then
            spanEnd = dayStart + departureMinutes / 1440#
            If recordKind = "Y" Then
                If CLng((CDbl(startAt) - recordStartDay) * 1440#) > departureMinutes Then spanEnd = dayEnd
            End If
        ElseIf dayStart = recordEndDay Then
            spanStart = dayStart + arrivalMinutes / 1440#
            If recordKind = "Y" Then
                If CLng((CDbl(endAt) - recordEndDay) * 1440#) < arrivalMinutes Then spanStart = dayStart
            End If
        Else
            spanStart = dayStart + arrivalMinutes / 1440#
            spanEnd = dayStart + departureMinutes / 1440#
        End If
    End If
    If spanStart < CDbl(startAt) Then spanStart = CDbl(startAt)
    If spanEnd > CDbl(endAt) Then spanEnd = CDbl(endAt)
    If spanStart < dayStart Then spanStart = dayStart
    If spanEnd > dayEnd Then spanEnd = dayEnd
    If spanStart >= spanEnd Then Exit Function
    minuteStart = CLng((spanStart - dayStart) * 1440#)
    minuteEnd = CLng((spanEnd - dayStart) * 1440#)
    If minuteStart < minuteEnd Then RecordDaySpan = Array(minuteStart, minuteEnd)
End Function

Public Sub AppendInterval(ByRef intervals As Variant, ByVal span As Variant)
    Dim n As Long
    If IsEmpty(span) Then Exit Sub
    If CLng(span(0)) >= CLng(span(1)) Then Exit Sub
    If IsEmpty(intervals) Then
        intervals = Array(span)
    Else
        n = UBound(intervals) + 1
        ReDim Preserve intervals(n)
        intervals(n) = span
    End If
End Sub

Public Function MergeIntervals(ByVal intervals As Variant) As Variant
    Dim sorted As Variant, result As Variant, current As Variant, nextSpan As Variant
    Dim i As Long, j As Long, minimum As Long, temporary As Variant
    If IsEmpty(intervals) Then Exit Function
    sorted = intervals
    ' 최댓값은 하루 1,440분이며 입력 건수에 관계없이 같은 날의 중복은 합친다.
    For i = LBound(sorted) To UBound(sorted) - 1
        minimum = i
        For j = i + 1 To UBound(sorted)
            If sorted(j)(0) < sorted(minimum)(0) Then minimum = j
        Next j
        If minimum <> i Then
            temporary = sorted(i)
            sorted(i) = sorted(minimum)
            sorted(minimum) = temporary
        End If
    Next i
    current = sorted(LBound(sorted))
    For i = LBound(sorted) + 1 To UBound(sorted)
        nextSpan = sorted(i)
        If nextSpan(0) <= current(1) Then
            If nextSpan(1) > current(1) Then current(1) = nextSpan(1)
        Else
            AppendInterval result, current
            current = nextSpan
        End If
    Next i
    AppendInterval result, current
    MergeIntervals = result
End Function

' A \ B를 구한다. 포함 판정은 길이 비교 대신 이 차집합이 비었는지 확인한다.
Public Function SubtractIntervals(ByVal minuend As Variant, ByVal subtrahend As Variant) As Variant
    Dim leftSpans As Variant, rightSpans As Variant, result As Variant
    Dim leftSpan As Variant, rightSpan As Variant, i As Long, j As Long
    Dim cursor As Long, finish As Long, gapEnd As Long
    leftSpans = MergeIntervals(minuend)
    If IsEmpty(leftSpans) Then Exit Function
    rightSpans = MergeIntervals(subtrahend)
    If IsEmpty(rightSpans) Then
        SubtractIntervals = leftSpans
        Exit Function
    End If
    For i = LBound(leftSpans) To UBound(leftSpans)
        leftSpan = leftSpans(i)
        cursor = CLng(leftSpan(0))
        finish = CLng(leftSpan(1))
        For j = LBound(rightSpans) To UBound(rightSpans)
            rightSpan = rightSpans(j)
            If rightSpan(0) >= finish Then Exit For
            If rightSpan(1) > cursor Then
                If rightSpan(0) > cursor Then
                    gapEnd = CLng(rightSpan(0))
                    If gapEnd > finish Then gapEnd = finish
                    AppendInterval result, Array(cursor, gapEnd)
                End If
                If rightSpan(1) > cursor Then cursor = CLng(rightSpan(1))
                If cursor >= finish Then Exit For
            End If
        Next j
        If cursor < finish Then AppendInterval result, Array(cursor, finish)
    Next i
    SubtractIntervals = result
End Function

Public Function IntervalsContain(ByVal outerIntervals As Variant, ByVal innerIntervals As Variant) As Boolean
    Dim remainder As Variant
    remainder = SubtractIntervals(innerIntervals, outerIntervals)
    IntervalsContain = IsEmpty(remainder)
End Function

Public Function IntervalsOutsideWork(ByVal intervals As Variant, ByVal arrivalMinutes As Long, _
                                    ByVal departureMinutes As Long) As Boolean
    Dim i As Long, span As Variant
    If IsEmpty(intervals) Then Exit Function
    For i = LBound(intervals) To UBound(intervals)
        span = intervals(i)
        If span(0) < arrivalMinutes Or span(1) > departureMinutes Then
            IntervalsOutsideWork = True
            Exit Function
        End If
    Next i
End Function

Public Function JudgeIntervals(ByVal x As Variant, ByVal y As Variant, ByVal z As Variant, _
                               ByVal arrivalMinutes As Long, ByVal departureMinutes As Long) As Variant
    Dim mask As Long, status As Long, caseNumber As Long, color As Long, reason As String
    If Not IsEmpty(x) Then mask = mask + 1
    If Not IsEmpty(y) Then mask = mask + 2
    If Not IsEmpty(z) Then mask = mask + 4
    status = 0
    color = -1
    Select Case mask
        Case 0
            status = 1: caseNumber = 1: color = RGB(255, 255, 0)
            reason = "복무·출장·방학 근무 기록 없음"
        Case 1
            caseNumber = 2
            reason = "복무 기록만 있음"
        Case 2
            status = 1: caseNumber = 3: color = RGB(255, 255, 0)
            reason = "출장 기록만 있음"
        Case 4
            status = 1: caseNumber = 4: color = RGB(255, 192, 0)
            reason = "방학 근무 기록만 있음"
        Case 3
            caseNumber = 5
            If IntervalsContain(y, x) Then
                status = 1: color = RGB(255, 192, 0)
                reason = "출장 구간이 모든 복무 구간을 포함함"
            ElseIf IntervalsOutsideWork(y, arrivalMinutes, departureMinutes) Then
                status = -1: color = RGB(217, 217, 217)
                reason = "출장이 복무 구간을 포함하지 않고 정규 근무시간 밖의 시간을 포함함"
            Else
                reason = "출장이 복무 구간을 포함하지 않고 정규 근무시간 안에 있음"
            End If
        Case 5
            caseNumber = 6
            If IntervalsContain(z, x) Then
                status = 1: color = RGB(255, 192, 0)
                reason = "방학 근무 구간이 모든 복무 구간을 포함함"
            Else
                reason = "방학 근무가 모든 복무 구간을 포함하지 않음"
            End If
        Case Else
            status = -1: color = RGB(217, 217, 217)
            reason = "출장과 방학 근무가 함께 있어 확인이 필요함"
    End Select
    JudgeIntervals = Array(status, caseNumber, color, reason)
End Function

Public Function MinuteText(ByVal minutes As Long) As String
    MinuteText = Right$("0" & CStr(minutes \ 60), 2) & ":" & Right$("0" & CStr(minutes Mod 60), 2)
End Function

Public Function IntervalText(ByVal intervals As Variant, ByVal prefix As String) As String
    Dim i As Long, span As Variant, result As String
    If IsEmpty(intervals) Then Exit Function
    For i = LBound(intervals) To UBound(intervals)
        span = intervals(i)
        If Len(result) > 0 Then result = result & vbLf
        If i = LBound(intervals) Then result = prefix
        result = result & MinuteText(CLng(span(0))) & "~" & MinuteText(CLng(span(1)))
    Next i
    IntervalText = result
End Function

Public Function DayResultText(ByVal x As Variant, ByVal y As Variant, ByVal z As Variant, ByVal status As Long) As String
    Dim result As String, part As String
    result = IntervalText(x, "복무: ")
    part = IntervalText(y, "출장: ")
    If Len(part) > 0 Then
        If Len(result) > 0 Then result = result & vbLf
        result = result & part
    End If
    part = IntervalText(z, "방학: ")
    If Len(part) > 0 Then
        If Len(result) > 0 Then result = result & vbLf
        result = result & part
    End If
    If status = -1 Then
        If Len(result) > 0 Then result = result & vbLf
        result = result & "판정 보류"
    End If
    DayResultText = result
End Function
' END PURE INTERVAL LOGIC
