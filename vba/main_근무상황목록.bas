Attribute VB_Name = "main_근무상황목록"
Option Explicit

Public Sub LoadWorkStatusRecords()
    Dim workMonth As Date, startMinute As Long, endMinute As Long
    Dim sourceBook As Workbook, sourceSheet As Worksheet, openedByCode As Boolean
    Dim staged As Object, warnings As Object, group As Collection, records As Collection
    Dim employee As CEmployee, record As CWorkStatusRecord
    Dim r As Long, count As Long, rawName As String, name As String, rawId As String
    Dim normalizedId As String, validId As Boolean, reason As String, warning As Variant
    Dim startAt As Date, endAt As Date, failure As String
    On Error GoTo Failed
    ReadBaseSettings workMonth, startMinute, endMinute, False
    RequireEmployeeState workMonth
    Set sourceBook = SelectInputWorkbook(openedByCode)
    If sourceBook Is Nothing Then Exit Sub
    Set sourceSheet = sourceBook.Worksheets(1)
    Set staged = NewRecordStage()
    Set warnings = NewDictionary()
    r = 2
    Do While r <= sourceSheet.Rows.count
        If IsBlankValue(sourceSheet.Cells(r, 3).Value2) Then Exit Do
        If Not IsCompletedWorkStatus(sourceSheet.Cells(r, 11).Value2) Then GoTo NextRow
        rawName = CellText(sourceSheet.Cells(r, 3))
        reason = CellText(sourceSheet.Cells(r, 7))
        If InStr(reason, "육아시간") = 0 And InStr(reason, "모성보호시간") = 0 Then
            ReadWorkStatusIdentity rawName, name, rawId, validId, normalizedId
            If Len(name) = 0 Then RaiseValidation CellContext(sourceSheet.Cells(r, 3)), "이름이 비어 있습니다."
            If gEmployeesByName.Exists(name) Then
                Set employee = Nothing
                Set group = gEmployeesByName(name)
                If group.count = 1 Then
                    Set employee = group(1)
                Else
                    Set employee = ResolveWorkStatusDuplicate(group, validId, normalizedId, sourceSheet.Cells(r, 3))
                    If employee Is Nothing Then
                        If Not IsConfirmedNonTargetOccupation(CellText(sourceSheet.Cells(r, 4)), group) Then
                            AddIdentityWarning warnings, sourceSheet, r, name, rawId, validId, normalizedId, group
                        End If
                    End If
                End If
                If Not employee Is Nothing Then
                    ReadWorkStatusPeriod sourceSheet.Cells(r, 5), startAt, endAt
                    Set record = New CWorkStatusRecord
                    record.StartAt = startAt
                    record.EndAt = endAt
                    Set records = staged(CStr(employee.Order))
                    records.Add record
                    count = count + 1
                End If
            End If
        End If
NextRow:
        r = r + 1
    Loop
    CloseInputWorkbook sourceBook, openedByCode
    Set sourceBook = Nothing
    CommitRecordStage staged, "WorkStatus"
    If warnings.count > 0 Then
        For Each warning In warnings.Keys
            ShowIdentityWarning CStr(warnings(warning)), "제외한 근무상황 기록"
        Next warning
    End If
    MsgBox "근무상황목록 " & count & "건을 불러왔습니다.", vbInformation, "근무상황목록 완료"
    Exit Sub
Failed:
    failure = Err.Description
    CloseInputWorkbook sourceBook, openedByCode
    MsgBox failure & vbCrLf & "기존 근무상황 기록은 유지됩니다.", vbExclamation, "근무상황목록 중단"
End Sub

' CellText의 공백 제거를 사용하지 않는다. 원본 값이 정확히 "완결"인 행만 읽는다.
Public Function IsCompletedWorkStatus(ByVal value As Variant) As Boolean
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function
    If VarType(value) <> vbString Then Exit Function
    IsCompletedWorkStatus = (StrComp(CStr(value), "완결", vbBinaryCompare) = 0)
End Function

Private Function ResolveWorkStatusDuplicate(ByVal candidates As Collection, ByVal validId As Boolean, _
                                           ByVal normalizedId As String, ByVal sourceCell As Range) As CEmployee
    Dim candidate As CEmployee, matched As CEmployee, registeredIds As Object
    Dim registeredId As String
    If Not validId Then
        RaiseValidation CellContext(sourceCell), _
                        "동명이인 중 어느 작업 대상자의 기록인지 구분할 수 없습니다. 나이스 개인번호를 확인한 뒤 다시 불러오세요."
    End If
    Set registeredIds = NewDictionary()
    For Each candidate In candidates
        If Not TryNormalizeNeisId(candidate.NeisPersonId, registeredId) Then
            RaiseValidation CellContext(sourceCell), "동명이인의 등록된 나이스 개인번호를 확인하고 대상자를 다시 생성하세요."
        End If
        If registeredIds.Exists(registeredId) Then
            RaiseValidation CellContext(sourceCell), "동명이인의 등록된 나이스 개인번호가 중복되어 구분할 수 없습니다. 대상자를 다시 생성하세요."
        End If
        registeredIds.Add registeredId, True
        If registeredId = normalizedId Then Set matched = candidate
    Next candidate
    Set ResolveWorkStatusDuplicate = matched
End Function

Private Function IsConfirmedNonTargetOccupation(ByVal sourceJobTitle As String, ByVal candidates As Collection) As Boolean
    Dim candidate As CEmployee
    ' 번호가 유효하지만 대상자 전원과 다를 때만 호출한다. 미일치 번호만으로
    ' 비대상으로 단정하지 않으며 확인된 비대상 직종에 한해서만 경고를 생략한다.
    If sourceJobTitle <> "영양사" Then Exit Function
    For Each candidate In candidates
        If Len(candidate.JobTitle) = 0 Then Exit Function
        If candidate.JobTitle = sourceJobTitle Then Exit Function
    Next candidate
    IsConfirmedNonTargetOccupation = True
End Function

Private Sub ReadWorkStatusIdentity(ByVal raw As String, ByRef name As String, ByRef rawId As String, ByRef validId As Boolean, ByRef normalizedId As String)
    Dim lines As Variant, candidateId As String
    raw = Replace(raw, vbCrLf, vbLf)
    raw = Replace(raw, vbCr, vbLf)
    lines = Split(raw, vbLf)
    name = TrimInput(CStr(lines(0)))
    rawId = vbNullString
    normalizedId = vbNullString
    validId = False
    If UBound(lines) = 0 Then Exit Sub
    rawId = TrimInput(CStr(lines(1)))
    If UBound(lines) > 1 Then
        rawId = Mid$(raw, InStr(raw, vbLf) + 1)
        Exit Sub
    End If
    candidateId = rawId
    If Len(candidateId) >= 2 Then
        If Left$(candidateId, 1) = "(" And Right$(candidateId, 1) = ")" Then
            candidateId = TrimInput(Mid$(candidateId, 2, Len(candidateId) - 2))
        End If
    End If
    validId = TryNormalizeNeisId(candidateId, normalizedId)
End Sub

Private Sub ReadWorkStatusPeriod(ByVal cell As Range, ByRef startAt As Date, ByRef endAt As Date)
    Dim parts As Variant, text As String
    text = CellText(cell)
    parts = Split(text, "~")
    If UBound(parts) <> 1 Then RaiseValidation CellContext(cell), "기간에는 ~가 정확히 1개 있어야 합니다."
    startAt = ParseDateTimeValue(TrimInput(CStr(parts(0))), "-", CellContext(cell) & " 시작")
    endAt = ParseDateTimeValue(TrimInput(CStr(parts(1))), "-", CellContext(cell) & " 종료")
    ValidateInterval startAt, endAt, CellContext(cell)
End Sub

Private Sub AddIdentityWarning(ByVal warnings As Object, ByVal ws As Worksheet, ByVal rowNumber As Long, _
                               ByVal name As String, ByVal rawId As String, ByVal validId As Boolean, _
                               ByVal normalizedId As String, ByVal candidates As Collection)
    Dim key As String, identity As String, detail As String
    identity = rawId
    If validId Then identity = normalizedId
    key = name & vbNullChar & identity
    If Not warnings.Exists(key) Then
        warnings.Add key, IdentityWarning(name, rawId, candidates) & vbCrLf & vbCrLf & "제외한 행 / 직종:"
    End If
    detail = CellContext(ws.Cells(rowNumber, 3)) & " / " & CellText(ws.Cells(rowNumber, 4))
    If Len(rawId) = 0 Then
        detail = detail & " / 나이스 개인번호: (없음)"
    Else
        detail = detail & " / 나이스 개인번호: " & rawId
    End If
    warnings(key) = CStr(warnings(key)) & vbCrLf & detail
End Sub

Private Function IdentityWarning(ByVal name As String, ByVal rawId As String, ByVal candidates As Collection) As String
    Dim text As String
    text = "근무상황목록에서 나이스 개인번호가 등록되지 않았거나 일치하지 않는 교직원이 있습니다. 해당 기록은 제외되었습니다." & _
           vbCrLf & vbCrLf & "이름: " & name & vbCrLf & "나이스 개인번호: "
    If Len(rawId) = 0 Then
        text = text & "(없음)"
    Else
        text = text & rawId
    End If
    IdentityWarning = text
End Function

Private Sub ShowIdentityWarning(ByVal text As String, ByVal title As String)
    Dim dialog As frmWorkStatusWarning, errorNumber As Long, errorText As String
    If Len(text) <= 800 Then
        MsgBox text, vbExclamation, title
        Exit Sub
    End If
    ' MsgBox 길이 한도로 긴 경고가 잘리거나 같은 사람 팝업이 여러 번 뜨지 않게 한다.
    On Error GoTo Failed
    Set dialog = New frmWorkStatusWarning
    dialog.Configure text, title
    dialog.Show vbModal
    Unload dialog
    Set dialog = Nothing
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    On Error Resume Next
    If Not dialog Is Nothing Then Unload dialog
    Set dialog = Nothing
    On Error GoTo 0
    Err.Raise errorNumber, "ShowIdentityWarning", errorText
End Sub
