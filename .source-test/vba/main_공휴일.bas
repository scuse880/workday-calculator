Attribute VB_Name = "main_공휴일"
Option Explicit

Public Sub 공휴일설정()
    Dim workMonth As Date, arrivalMinutes As Long, departureMinutes As Long
    Dim holidays As Object, saved As Boolean, dialog As frmHoliday
    On Error GoTo Failed
    ReadBaseSettings workMonth, arrivalMinutes, departureMinutes, False
    Set holidays = LoadHolidays(workMonth, saved)
    Set dialog = New frmHoliday
    dialog.Configure workMonth, holidays
    dialog.Show vbModal
CleanUp:
    If Not dialog Is Nothing Then
        dialog.DisposeHandlers
        Unload dialog
    End If
    Set dialog = Nothing
    Exit Sub
Failed:
    Debug.Print Err.Description, vbExclamation, "공휴일 설정"
    Resume CleanUp
End Sub

' Date 키는 로캘에 의존하지 않는 정수 일련번호 문자열로 통일한다.
Public Function HolidayKey(ByVal value As Date) As String
    HolidayKey = CStr(CLng(Int(CDbl(value))))
End Function

Public Function KoreanWeekday(ByVal value As Date) As String
    KoreanWeekday = Mid$("일월화수목금토", Weekday(value, vbSunday), 1)
End Function

' saved=False와 빈 Dictionary를 구별하여 공휴일 0개 저장 여부를 보존한다.
Public Function LoadHolidays(ByVal workMonth As Date, ByRef saved As Boolean) As Object
    Dim ws As Worksheet, targetRow As Long, lastColumn As Long, c As Long
    Dim holiday As Date, result As Object
    Set ws = GetOrCreateSheet("공휴일")
    Set result = NewDictionary()
    targetRow = FindHolidayMonthRow(ws, workMonth)
    saved = (targetRow > 0)
    If saved Then
        lastColumn = ws.Cells(targetRow, ws.Columns.Count).End(xlToLeft).Column
        For c = 2 To lastColumn
            If Not IsBlankValue(ws.Cells(targetRow, c).Value2) Then
                holiday = ParseDateCell(ws.Cells(targetRow, c))
                If Year(holiday) <> Year(workMonth) Or Month(holiday) <> Month(workMonth) Then
                    RaiseValidation "공휴일!" & ws.Cells(targetRow, c).Address(False, False), _
                                    "현재 작업년월에 속하는 날짜만 저장할 수 있습니다."
                End If
                If Not result.Exists(HolidayKey(holiday)) Then result.Add HolidayKey(holiday), holiday
            End If
        Next c
    End If
    Set LoadHolidays = result
End Function

Private Function FindHolidayMonthRow(ByVal ws As Worksheet, ByVal workMonth As Date) As Long
    Dim lastRow As Long, r As Long, rowMonth As Date
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If Not IsBlankValue(ws.Cells(r, 1).Value2) Then
            rowMonth = ParseDateCell(ws.Cells(r, 1))
            If Year(rowMonth) = Year(workMonth) And Month(rowMonth) = Month(workMonth) Then
                If Day(rowMonth) <> 1 Then
                    RaiseValidation "공휴일!A" & r, "작업년월에는 해당 월 1일을 입력하세요."
                End If
                If FindHolidayMonthRow <> 0 Then
                    RaiseValidation "공휴일!A" & r, "같은 작업년월이 여러 행에 저장되어 있습니다."
                End If
                FindHolidayMonthRow = r
            End If
        End If
    Next r
End Function

Public Sub SaveHolidays(ByVal workMonth As Date, ByVal holidays As Object)
    Dim ws As Worksheet, existing As Object, saved As Boolean, targetRow As Long
    Dim lastColumn As Long, writeColumn As Long, d As Long, dayValue As Date
    Dim key As Variant, rawValue As Variant, oldFormula As Variant, backupRange As Range
    Dim errorNumber As Long, errorText As String, hasBackup As Boolean
    ' 저장 직전에도 시트 상태와 현재 선택을 검증한다. 실패 시 기존 행을 복원한다.
    Set existing = LoadHolidays(workMonth, saved)
    For Each key In holidays.Keys
        rawValue = holidays(key)
        If Not IsDate(rawValue) Then RaiseValidation "공휴일", "날짜가 아닌 선택값이 있습니다."
        dayValue = CDate(rawValue)
        If CDbl(dayValue) <> Int(CDbl(dayValue)) Then RaiseValidation "공휴일", "공휴일에는 시각을 입력할 수 없습니다."
        If Year(dayValue) <> Year(workMonth) Or Month(dayValue) <> Month(workMonth) Then
            RaiseValidation "공휴일", "현재 작업년월에 속하는 날짜만 선택하세요."
        End If
        If CStr(key) <> HolidayKey(dayValue) Then RaiseValidation "공휴일", "선택한 날짜의 식별값이 올바르지 않습니다."
    Next key
    Set ws = GetOrCreateSheet("공휴일")
    targetRow = FindHolidayMonthRow(ws, workMonth)
    If targetRow = 0 Then
        targetRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        If targetRow < 2 Then targetRow = 2
    End If
    lastColumn = ws.Cells(targetRow, ws.Columns.Count).End(xlToLeft).Column
    If lastColumn < 32 Then lastColumn = 32
    Set backupRange = ws.Range(ws.Cells(targetRow, 1), ws.Cells(targetRow, lastColumn))
    oldFormula = backupRange.Formula
    hasBackup = True
    On Error GoTo Rollback
    ws.Range("A1").Value2 = "작업년월"
    ws.Range("B1").Value2 = "공휴일"
    backupRange.ClearContents
    ws.Cells(targetRow, 1).Value = DateSerial(Year(workMonth), Month(workMonth), 1)
    ws.Cells(targetRow, 1).NumberFormat = "yyyy-mm"
    writeColumn = 2
    For d = 1 To Day(DateSerial(Year(workMonth), Month(workMonth) + 1, 0))
        dayValue = DateSerial(Year(workMonth), Month(workMonth), d)
        If holidays.Exists(HolidayKey(dayValue)) Then
            ws.Cells(targetRow, writeColumn).Value = dayValue
            ws.Cells(targetRow, writeColumn).NumberFormat = "yyyy-mm-dd"
            writeColumn = writeColumn + 1
        End If
    Next d
    ClearCalculationResults
    Exit Sub
Rollback:
    errorNumber = Err.Number
    errorText = Err.Description
    On Error Resume Next
    If hasBackup Then backupRange.Formula = oldFormula
    On Error GoTo 0
    Err.Raise errorNumber, "SaveHolidays", errorText
End Sub
