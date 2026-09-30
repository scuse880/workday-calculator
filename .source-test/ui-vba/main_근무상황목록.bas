Attribute VB_Name = "main_근무상황목록"
Option Explicit

Public Sub LoadWorkStatusRecords()
    Dim workMonth As Date, startMinute As Long, endMinute As Long
    Dim sourceBook As Workbook, sourceSheet As Worksheet, openedByCode As Boolean
    Dim staged As Object, warnings As Collection, group As Collection, records As Collection
    Dim employee As CEmployee, candidate As Variant, record As CWorkStatusRecord
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
    Set warnings = New Collection
    r = 2
    Do While r <= sourceSheet.Rows.count
        rawName = CellText(sourceSheet.Cells(r, 3))
        If Len(rawName) = 0 Then Exit Do
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
                    If validId Then
                        For Each candidate In group
                            If candidate.NeisPersonId = normalizedId Then
                                Set employee = candidate
                                Exit For
                            End If
                        Next candidate
                    End If
                    If employee Is Nothing Then
                        warnings.Add IdentityWarning(sourceSheet, r, name, rawId, group)
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
        r = r + 1
    Loop
    CloseInputWorkbook sourceBook, openedByCode
    Set sourceBook = Nothing
    CommitRecordStage staged, "WorkStatus"
    If warnings.count > 0 Then
        MsgBox "근무상황목록에서 나이스 개인번호가 등록되지 않았거나 일치하지 않는 교직원이 있습니다. 해당 기록은 제외되었습니다." & _
               vbCrLf & "제외한 " & warnings.count & "개 행의 내용을 차례로 표시합니다.", vbExclamation, "근무상황목록 확인"
        For Each warning In warnings
            ShowTextPages CStr(warning), "제외한 근무상황 기록"
        Next warning
    End If
    MsgBox "근무상황목록 " & count & "건을 불러왔습니다.", vbInformation, "근무상황목록 완료"
    Exit Sub
Failed:
    failure = Err.Description
    CloseInputWorkbook sourceBook, openedByCode
    MsgBox failure & vbCrLf & "기존 근무상황 기록은 유지됩니다.", vbExclamation, "근무상황목록 중단"
End Sub

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

Private Function IdentityWarning(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal name As String, ByVal rawId As String, ByVal candidates As Collection) As String
    Dim text As String, item As Variant
    text = CellContext(ws.Cells(rowNumber, 3)) & vbCrLf & "이름: " & name & vbCrLf & "원본 개인번호: "
    If Len(rawId) = 0 Then
        text = text & "(없음)"
    Else
        text = text & rawId
    End If
    text = text & vbCrLf & "직종: " & CellText(ws.Cells(rowNumber, 4)) & vbCrLf & "등록된 동명이인:"
    For Each item In candidates
        text = text & vbCrLf & item.Name & "(" & item.Birthdate & ") / " & item.NeisPersonId
    Next item
    IdentityWarning = text
End Function

Private Sub ShowTextPages(ByVal text As String, ByVal title As String)
    Dim position As Long, page As Long, total As Long
    total = (Len(text) + 799) \ 800
    If total = 0 Then total = 1
    position = 1
    For page = 1 To total
        MsgBox Mid$(text, position, 800), vbExclamation, title & " (" & page & "/" & total & ")"
        position = position + 800
    Next page
End Sub
