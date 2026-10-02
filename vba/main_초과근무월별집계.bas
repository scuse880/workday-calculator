Attribute VB_Name = "main_초과근무월별집계"
Option Explicit

' 자료는 메모리에만 보관하며 재실행 성공 시 한 번에 교체한다.
Public gEmployeesByName As Object
Public gEmployeesMonth As Date
Public gWorkStatusLoaded As Boolean
Public gTripLoaded As Boolean
Public gVacationWorkLoaded As Boolean

Public Sub LoadOvertimeEmployees()
    Dim workMonth As Date, startMinute As Long, endMinute As Long
    Dim sourceBook As Workbook, sourceSheet As Worksheet, openedByCode As Boolean
    Dim staged As Object, group As Collection, employee As CEmployee
    Dim employeeName As String, birthdate As String, r As Long, count As Long
    Dim key As Variant, births As Object, item As Variant, duplicates As Collection
    Dim dialog As frmNeisPersonId, failure As String
    On Error GoTo Failed
    ReadBaseSettings workMonth, startMinute, endMinute, False
    Set sourceBook = SelectInputWorkbook(openedByCode)
    If sourceBook Is Nothing Then Exit Sub
    Set sourceSheet = sourceBook.Worksheets(1)
    Set staged = NewDictionary()
    r = 5
    Do While r <= sourceSheet.Rows.Count
        If Len(CellText(sourceSheet.Cells(r, 3))) = 0 Then
            r = r + 3
            If r > sourceSheet.Rows.Count Then Exit Do
            If Len(CellText(sourceSheet.Cells(r, 3))) = 0 Then Exit Do
        End If
        ParseNameBirth CellText(sourceSheet.Cells(r, 3)), employeeName, birthdate, CellContext(sourceSheet.Cells(r, 3))
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
    If count = 0 Then RaiseValidation "초과근무월별집계 C5", "대상자가 없습니다. 입력 파일을 확인하세요."
    For Each key In staged.Keys
        Set group = staged(key)
        If group.count > 1 Then
            Set births = NewDictionary()
            For Each item In group
                Set employee = item
                If Len(employee.Birthdate) = 0 Then
                    RaiseValidation employee.Name, "동명이인이 있는 경우 생년월일을 포함한 파일을 업로드해주시기 바랍니다."
                End If
                If births.Exists(employee.Birthdate) Then
                    RaiseValidation employee.Name & "(" & employee.Birthdate & ")", "이름과 생년월일이 같은 대상자가 중복되어 구분할 수 없습니다."
                End If
                births.Add employee.Birthdate, True
            Next item
        End If
    Next key
    Set duplicates = New Collection
    For Each item In GetEmployeesInOrder(staged)
        Set employee = item
        Set group = staged(employee.Name)
        If group.count > 1 Then duplicates.Add employee
    Next item
    CloseInputWorkbook sourceBook, openedByCode
    Set sourceBook = Nothing
    If duplicates.count > 0 Then
        Set dialog = New frmNeisPersonId
        dialog.Configure duplicates
        dialog.Show vbModal
        If Not dialog.Accepted Then GoTo Cancelled
        Unload dialog
        Set dialog = Nothing
    End If
    ' 결과를 지우지 못하면 기존 메모리 자료도 유지한다.
    ClearCalculationResults
    Set gEmployeesByName = staged
    gEmployeesMonth = workMonth
    gWorkStatusLoaded = False
    gTripLoaded = False
    gVacationWorkLoaded = False
    MsgBox Format$(workMonth, "yyyy년 m월") & " 대상자 " & count & "명을 불러왔습니다." & vbCrLf & _
           "근무상황목록과 출장 자료를 이어서 불러오세요.", vbInformation, "대상자 생성 완료"
    Exit Sub
Cancelled:
    On Error Resume Next
    If Not dialog Is Nothing Then Unload dialog
    CloseInputWorkbook sourceBook, openedByCode
    Exit Sub
Failed:
    failure = Err.Description
    On Error Resume Next
    If Not dialog Is Nothing Then Unload dialog
    CloseInputWorkbook sourceBook, openedByCode
    MsgBox failure & vbCrLf & "기존 대상자와 근무 기록은 유지됩니다.", vbExclamation, "대상자 생성 중단"
End Sub

Public Function NewDictionary() As Object
    Set NewDictionary = CreateObject("Scripting.Dictionary")
    NewDictionary.CompareMode = vbBinaryCompare
End Function

Public Sub RaiseValidation(ByVal context As String, ByVal message As String)
    Err.Raise vbObjectError + 2100, "근무일수계산", context & vbCrLf & message
End Sub

Public Function CellContext(ByVal cell As Range) As String
    CellContext = "[" & cell.Parent.Parent.Name & "] " & cell.Parent.Name & "!" & cell.Address(False, False) & " (" & cell.Row & "행)"
End Function

Public Function TrimInput(ByVal value As String) As String
    Dim code As Long
    Do While Len(value) > 0
        code = AscW(Left$(value, 1))
        If code <> 32 And code <> 9 And code <> 10 And code <> 13 And code <> 160 And code <> 12288 Then Exit Do
        value = Mid$(value, 2)
    Loop
    Do While Len(value) > 0
        code = AscW(Right$(value, 1))
        If code <> 32 And code <> 9 And code <> 10 And code <> 13 And code <> 160 And code <> 12288 Then Exit Do
        value = Left$(value, Len(value) - 1)
    Loop
    TrimInput = value
End Function

Public Function CellText(ByVal cell As Range) As String
    Dim value As Variant
    value = cell.Value2
    If IsError(value) Then RaiseValidation CellContext(cell), "셀에 Excel 오류 값이 있습니다."
    If IsEmpty(value) Or IsNull(value) Then Exit Function
    CellText = TrimInput(CStr(value))
End Function

Public Function IsBlankValue(ByVal value As Variant) As Boolean
    If IsError(value) Then Exit Function
    If IsEmpty(value) Or IsNull(value) Then
        IsBlankValue = True
    Else
        IsBlankValue = (Len(TrimInput(CStr(value))) = 0)
    End If
End Function

Public Function MatchPattern(ByVal value As String, ByVal pattern As String) As Object
    Dim expression As Object, matches As Object
    Set expression = CreateObject("VBScript.RegExp")
    expression.pattern = pattern
    expression.Global = False
    expression.IgnoreCase = False
    Set matches = expression.Execute(value)
    If matches.count > 0 Then Set MatchPattern = matches(0)
End Function

Public Sub ParseNameBirth(ByVal text As String, ByRef employeeName As String, ByRef birthdate As String, ByVal context As String)
    Dim match As Object
    text = TrimInput(text)
    employeeName = vbNullString
    birthdate = vbNullString
    If InStr(text, "(") > 0 Or InStr(text, ")") > 0 Then
        Set match = MatchPattern(text, "^([^()]*)\(([0-9]{6})\)$")
        If match Is Nothing Then RaiseValidation context, "이름(생년월일) 형식을 확인하세요. 생년월일은 숫자 6자리입니다."
        employeeName = TrimInput(CStr(match.SubMatches(0)))
        birthdate = CStr(match.SubMatches(1))
    Else
        employeeName = text
    End If
    If Len(employeeName) = 0 Then RaiseValidation context, "이름이 비어 있습니다."
End Sub

Public Function TryNormalizeNeisId(ByVal text As String, ByRef normalized As String) As Boolean
    Dim match As Object
    normalized = UCase$(TrimInput(text))
    Set match = MatchPattern(normalized, "^C[0-9]{9}$")
    TryNormalizeNeisId = Not match Is Nothing
    If Not TryNormalizeNeisId Then normalized = vbNullString
End Function

Public Sub ReadBaseSettings(ByRef workMonth As Date, ByRef arrivalMinutes As Long, ByRef departureMinutes As Long, Optional ByVal needHours As Boolean = True)
    Dim ws As Worksheet, y As Long, m As Long, h1 As Long, h2 As Long, n1 As Long, n2 As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("작업")
    On Error GoTo 0
    If ws Is Nothing Then RaiseValidation "작업", "작업 시트가 없습니다."
    y = ReadIntegerCell(ws.Range("B3"), "년", 1000, 9999, True)
    m = ReadIntegerCell(ws.Range("D3"), "월", 1, 12, False)
    workMonth = DateSerial(y, m, 1)
    arrivalMinutes = 0
    departureMinutes = 0
    If Not needHours Then Exit Sub
    h1 = ReadIntegerCell(ws.Range("B5"), "시", 0, 23, False)
    n1 = ReadIntegerCell(ws.Range("D5"), "분", 0, 59, False)
    h2 = ReadIntegerCell(ws.Range("B6"), "시", 0, 23, False)
    n2 = ReadIntegerCell(ws.Range("D6"), "분", 0, 59, False)
    arrivalMinutes = h1 * 60 + n1
    departureMinutes = h2 * 60 + n2
    If arrivalMinutes >= departureMinutes Then
        RaiseValidation "작업!B5/D5/B6/D6", "출근시각은 같은 날의 퇴근시각보다 빨라야 합니다."
    End If
End Sub

Private Function ReadIntegerCell(ByVal cell As Range, ByVal suffix As String, ByVal minimum As Long, ByVal maximum As Long, ByVal fourDigits As Boolean) As Long
    Dim text As String, match As Object, pattern As String, number As Double
    text = CellText(cell)
    If text = "선택" Or Len(text) = 0 Then RaiseValidation CellContext(cell), "드롭다운에서 값을 선택하세요."
    If fourDigits Then
        pattern = "^([0-9]{4})" & suffix & "?$"
    ElseIf Len(suffix) > 0 Then
        pattern = "^([0-9]{1,2})" & suffix & "?$"
    Else
        pattern = "^([0-9]{1,2})$"
    End If
    Set match = MatchPattern(text, pattern)
    If match Is Nothing Then RaiseValidation CellContext(cell), "정수 입력 형식을 확인하세요. 허용 범위: " & minimum & "~" & maximum
    number = CDbl(match.SubMatches(0))
    If number < minimum Or number > maximum Then RaiseValidation CellContext(cell), "허용 범위는 " & minimum & "~" & maximum & "입니다."
    ReadIntegerCell = CLng(number)
End Function

Public Sub RequireEmployeeState(ByVal workMonth As Date)
    If gEmployeesByName Is Nothing Then
        RaiseValidation "실행 순서", "먼저 초과근무월별집계를 불러와 대상자를 생성하세요."
    End If
    If gEmployeesByName.count = 0 Then RaiseValidation "실행 순서", "먼저 초과근무월별집계를 불러와 대상자를 생성하세요."
    If gEmployeesMonth <> workMonth Then
        RaiseValidation "작업년월 변경", "작업 시트의 작업년월이 변경되었습니다. 초과근무월별집계부터 다시 불러오세요."
    End If
End Sub

Public Function SelectInputWorkbook(ByRef openedByCode As Boolean) As Workbook
    Dim selected As Variant, wb As Workbook, oldSecurity As Long, oldEvents As Boolean
    Dim previousWindow As Window, previousSheet As Object
    Dim savedNumber As Long, savedDescription As String
    openedByCode = False
    selected = Application.GetOpenFilename("Excel 통합 문서 (*.xlsx),*.xlsx", , "입력할 xlsx 파일 선택")
    If VarType(selected) = vbBoolean Then Exit Function
    If LCase$(Right$(CStr(selected), 5)) <> ".xlsx" Then RaiseValidation "입력 파일", "xlsx 파일을 선택하세요."
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, CStr(selected), vbTextCompare) = 0 Then
            Set SelectInputWorkbook = wb
            Exit Function
        End If
    Next wb
    Set previousWindow = Application.ActiveWindow
    Set previousSheet = Application.ActiveSheet
    oldSecurity = Application.AutomationSecurity
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.AutomationSecurity = 3
    Application.EnableEvents = False
    Set wb = Application.Workbooks.Open(Filename:=CStr(selected), UpdateLinks:=0, ReadOnly:=True, AddToMru:=False, IgnoreReadOnlyRecommended:=True)
    openedByCode = True
    ' 입력 파일은 뒤에서 읽는다. 열기/닫기가 호출한 창과 시트의 포커스를 바꾸지 않게 한다.
    If Not previousWindow Is Nothing Then previousWindow.Activate
    If Not previousSheet Is Nothing Then previousSheet.Activate
    Application.AutomationSecurity = oldSecurity
    Application.EnableEvents = oldEvents
    Set SelectInputWorkbook = wb
    Exit Function
Failed:
    savedNumber = Err.Number
    savedDescription = Err.Description
    On Error Resume Next
    If openedByCode Then wb.Close SaveChanges:=False
    If Not previousWindow Is Nothing Then previousWindow.Activate
    If Not previousSheet Is Nothing Then previousSheet.Activate
    Application.AutomationSecurity = oldSecurity
    Application.EnableEvents = oldEvents
    On Error GoTo 0
    Err.Raise savedNumber, "입력 파일 열기", savedDescription
End Function

Public Sub CloseInputWorkbook(ByVal wb As Workbook, ByVal openedByCode As Boolean)
    If wb Is Nothing Then Exit Sub
    If Not openedByCode Then Exit Sub
    ' 원래 오류를 가리지 않고, 읽기 전용 원본은 저장하지 않는다.
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Public Function ParseDateTimeCell(ByVal cell As Range, ByVal dateSeparator As String) As Date
    ParseDateTimeCell = ParseDateTimeValue(cell.Value2, dateSeparator, CellContext(cell), CBool(cell.Parent.Parent.Date1904))
End Function

Public Function ParseDateTimeValue(ByVal value As Variant, ByVal dateSeparator As String, ByVal context As String, Optional ByVal date1904 As Boolean = False) As Date
    Dim text As String, match As Object, escaped As String, d As Date
    If IsError(value) Then RaiseValidation context, "날짜·시간 셀에 Excel 오류가 있습니다."
    If IsBlankValue(value) Then RaiseValidation context, "날짜·시간이 비어 있습니다."
    If VarType(value) = vbString Then
        text = TrimInput(CStr(value))
        escaped = dateSeparator
        If escaped = "." Then escaped = "\."
        Set match = MatchPattern(text, "^([0-9]{4})" & escaped & "([0-9]{2})" & escaped & "([0-9]{2}) ([0-9]{2}):([0-9]{2})$")
        If match Is Nothing Then RaiseValidation context, "날짜·시간은 yyyy" & dateSeparator & "mm" & dateSeparator & "dd hh:mm 형식으로 입력하세요."
        d = CheckedDate(CLng(match.SubMatches(0)), CLng(match.SubMatches(1)), CLng(match.SubMatches(2)), context)
        ParseDateTimeValue = d + CheckedTime(CLng(match.SubMatches(3)), CLng(match.SubMatches(4)), context)
    ElseIf VarType(value) = vbDate Then
        ParseDateTimeValue = ValidateWholeMinute(CDate(value), context)
    ElseIf IsNumeric(value) And VarType(value) <> vbBoolean Then
        ParseDateTimeValue = ExcelSerialDate(CDbl(value), date1904, context)
    Else
        RaiseValidation context, "유효한 날짜·시간을 입력하세요."
    End If
End Function

Public Function ParseDateCell(ByVal cell As Range) As Date
    Dim value As Variant, match As Object, d As Date
    value = cell.Value2
    If IsError(value) Then RaiseValidation CellContext(cell), "날짜 셀에 Excel 오류가 있습니다."
    If IsBlankValue(value) Then RaiseValidation CellContext(cell), "날짜가 비어 있습니다."
    If VarType(value) = vbString Then
        Set match = MatchPattern(TrimInput(CStr(value)), "^([0-9]{4})-([0-9]{2})-([0-9]{2})$")
        If match Is Nothing Then RaiseValidation CellContext(cell), "날짜는 yyyy-mm-dd 형식으로 입력하세요."
        d = CheckedDate(CLng(match.SubMatches(0)), CLng(match.SubMatches(1)), CLng(match.SubMatches(2)), CellContext(cell))
    Else
        d = ParseDateTimeValue(value, "-", CellContext(cell), CBool(cell.Parent.Parent.Date1904))
    End If
    If CDbl(d) <> Fix(CDbl(d)) Then RaiseValidation CellContext(cell), "날짜에는 시각을 포함할 수 없습니다."
    ParseDateCell = d
End Function

Public Function ParseTimeCell(ByVal cell As Range) As Date
    Dim value As Variant, match As Object, number As Double, minutes As Double
    value = cell.Value2
    If IsError(value) Then RaiseValidation CellContext(cell), "시각 셀에 Excel 오류가 있습니다."
    If IsBlankValue(value) Then RaiseValidation CellContext(cell), "시각이 비어 있습니다."
    If VarType(value) = vbString Then
        Set match = MatchPattern(TrimInput(CStr(value)), "^([0-9]{2}):([0-9]{2})$")
        If match Is Nothing Then RaiseValidation CellContext(cell), "시각은 hh:mm 형식으로 입력하세요."
        ParseTimeCell = CheckedTime(CLng(match.SubMatches(0)), CLng(match.SubMatches(1)), CellContext(cell))
    ElseIf IsNumeric(value) And VarType(value) <> vbBoolean Then
        number = CDbl(value)
        If number < 0 Or number >= 1 Then RaiseValidation CellContext(cell), "같은 날의 시각만 입력하세요. 24:00 및 날짜를 포함한 값은 사용할 수 없습니다."
        minutes = number * 1440#
        If Abs(minutes - Round(minutes, 0)) > 0.00001 Then RaiseValidation CellContext(cell), "초가 0이 아닌 시각은 사용할 수 없습니다."
        If Round(minutes, 0) >= 1440 Then RaiseValidation CellContext(cell), "24:00은 사용할 수 없습니다."
        ParseTimeCell = CDate(Round(minutes, 0) / 1440#)
    Else
        RaiseValidation CellContext(cell), "유효한 시각을 입력하세요."
    End If
End Function

Private Function CheckedDate(ByVal y As Long, ByVal m As Long, ByVal d As Long, ByVal context As String) As Date
    Dim result As Date
    If y < 1000 Or y > 9999 Or m < 1 Or m > 12 Or d < 1 Or d > 31 Then RaiseValidation context, "실제로 존재하는 연·월·일을 입력하세요."
    On Error GoTo InvalidDate
    result = DateSerial(y, m, d)
    On Error GoTo 0
    If Year(result) <> y Or Month(result) <> m Or Day(result) <> d Then GoTo InvalidDate
    CheckedDate = result
    Exit Function
InvalidDate:
    On Error GoTo 0
    RaiseValidation context, "실제로 존재하는 연·월·일을 입력하세요."
End Function

Private Function CheckedTime(ByVal h As Long, ByVal n As Long, ByVal context As String) As Date
    If h < 0 Or h > 23 Or n < 0 Or n > 59 Then RaiseValidation context, "시는 00~23, 분은 00~59여야 합니다."
    CheckedTime = TimeSerial(h, n, 0)
End Function

Private Function ExcelSerialDate(ByVal serial As Double, ByVal date1904 As Boolean, ByVal context As String) As Date
    Dim adjusted As Double
    If date1904 Then
        If serial < 0 Then RaiseValidation context, "Excel 날짜 값이 허용 범위를 벗어났습니다."
        adjusted = serial + 1462#
    Else
        If serial < 1 Then RaiseValidation context, "날짜가 없는 시각 값입니다."
        If Fix(serial) = 60 Then RaiseValidation context, "1900-02-29는 실제로 존재하지 않는 날짜입니다."
        adjusted = serial
        If serial < 60 Then adjusted = serial + 1#
    End If
    If adjusted >= 2958466# Then RaiseValidation context, "Excel 날짜 값이 허용 범위를 벗어났습니다."
    ExcelSerialDate = ValidateWholeMinute(CDate(adjusted), context)
End Function

Private Function ValidateWholeMinute(ByVal value As Date, ByVal context As String) As Date
    Dim minutes As Double
    minutes = CDbl(value) * 1440#
    If Abs(minutes - Round(minutes, 0)) > 0.00001 Then RaiseValidation context, "초가 0이 아닌 날짜·시간은 사용할 수 없습니다."
    ValidateWholeMinute = CDate(Round(minutes, 0) / 1440#)
End Function

Public Sub ValidateInterval(ByVal startAt As Date, ByVal endAt As Date, ByVal context As String)
    If startAt >= endAt Then RaiseValidation context, "시작 시각은 종료 시각보다 빨라야 합니다."
End Sub

Public Function GetEmployeesInOrder(Optional ByVal employeeMap As Object = Nothing) As Collection
    Dim result As Collection, indexed() As CEmployee, key As Variant, group As Collection
    Dim item As Variant, employee As CEmployee, count As Long, i As Long
    Set result = New Collection
    If employeeMap Is Nothing Then Set employeeMap = gEmployeesByName
    If employeeMap Is Nothing Then
        Set GetEmployeesInOrder = result
        Exit Function
    End If
    For Each key In employeeMap.Keys
        Set group = employeeMap(key)
        count = count + group.count
    Next key
    If count > 0 Then
        ReDim indexed(1 To count)
        For Each key In employeeMap.Keys
            Set group = employeeMap(key)
            For Each item In group
                Set employee = item
                If employee.Order < 1 Or employee.Order > count Then RaiseValidation "대상자 순서", "Order가 올바르지 않습니다. 대상자를 다시 불러오세요."
                If Not indexed(employee.Order) Is Nothing Then RaiseValidation "대상자 순서", "Order가 중복되었습니다. 대상자를 다시 불러오세요."
                Set indexed(employee.Order) = employee
            Next item
        Next key
        For i = 1 To count
            result.Add indexed(i)
        Next i
    End If
    Set GetEmployeesInOrder = result
End Function

Public Function NewRecordStage() As Object
    Dim staged As Object, item As Variant, records As Collection
    Set staged = NewDictionary()
    For Each item In GetEmployeesInOrder()
        Set records = New Collection
        staged.Add CStr(item.Order), records
    Next item
    Set NewRecordStage = staged
End Function

Public Sub CommitRecordStage(ByVal staged As Object, ByVal recordKind As String)
    Dim replacement As Object, item As Variant, oldEmployee As CEmployee, employee As CEmployee
    Dim group As Collection, records As Collection
    If recordKind <> "WorkStatus" And recordKind <> "Trip" And recordKind <> "VacationWork" Then RaiseValidation "자료 반영", "알 수 없는 기록 종류입니다."
    Set replacement = NewDictionary()
    ' 변경 대상 외의 Collection은 그대로 보존하고 새 CEmployee들에 연결한다.
    For Each item In GetEmployeesInOrder()
        Set oldEmployee = item
        If Not staged.Exists(CStr(oldEmployee.Order)) Then RaiseValidation "자료 반영", "임시 기록에 대상자가 누락되었습니다."
        Set records = staged(CStr(oldEmployee.Order))
        If records Is Nothing Then RaiseValidation "자료 반영", "임시 기록이 초기화되지 않았습니다."
        Set employee = New CEmployee
        employee.Order = oldEmployee.Order
        employee.Name = oldEmployee.Name
        employee.Birthdate = oldEmployee.Birthdate
        employee.NeisPersonId = oldEmployee.NeisPersonId
        employee.PersonnelChangeStart = oldEmployee.PersonnelChangeStart
        employee.PersonnelChangeEnd = oldEmployee.PersonnelChangeEnd
        Set employee.WorkStatusRecords = oldEmployee.WorkStatusRecords
        Set employee.TripRecords = oldEmployee.TripRecords
        Set employee.VacationWorkRecords = oldEmployee.VacationWorkRecords
        Select Case recordKind
            Case "WorkStatus": Set employee.WorkStatusRecords = records
            Case "Trip": Set employee.TripRecords = records
            Case "VacationWork": Set employee.VacationWorkRecords = records
        End Select
        If Not replacement.Exists(employee.Name) Then
            Set group = New Collection
            replacement.Add employee.Name, group
        End If
        Set group = replacement(employee.Name)
        group.Add employee
    Next item
    ClearCalculationResults
    Set gEmployeesByName = replacement
    Select Case recordKind
        Case "WorkStatus": gWorkStatusLoaded = True
        Case "Trip": gTripLoaded = True
        Case "VacationWork": gVacationWorkLoaded = True
    End Select
End Sub

Public Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = CreateProtectedWorkbookSheet(sheetName)
    End If
    Set GetOrCreateSheet = ws
End Function

Public Sub ClearCalculationResults()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("작업결과")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    If ws.ProtectContents And Not ws.ProtectionMode Then RaiseValidation "작업결과", "시트 보호를 해제한 후 다시 실행하세요."
    ws.UsedRange.Clear
End Sub
