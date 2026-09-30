Attribute VB_Name = "main_출장목록개인별"
Option Explicit

' Button entry point. A cancelled selection discards the entire staged import.
Public Sub LoadTripRecords()
    Dim workMonth As Date, arrivalMinutes As Long, departureMinutes As Long
    Dim sourceBook As Workbook, sourceSheet As Worksheet, openedByCode As Boolean
    Dim staged As Object, candidates As Collection, records As Collection
    Dim employee As CEmployee, record As CTripRecord
    Dim selector As frmTripEmployee
    Dim r As Long, consecutiveNonTargets As Long, recordCount As Long
    Dim originalName As String, employeeName As String, idPrefix As String
    Dim startText As String, endText As String, hasMaskedId As Boolean
    Dim startAt As Date, endAt As Date, errorText As String
    Dim skipRecord As Boolean

    On Error GoTo Failed
    ReadBaseSettings workMonth, arrivalMinutes, departureMinutes, False
    RequireEmployeeState workMonth
    Set sourceBook = SelectInputWorkbook(openedByCode)
    If sourceBook Is Nothing Then Exit Sub
    Set sourceSheet = sourceBook.Worksheets(1)
    Set staged = NewRecordStage()

    r = 5
    Do While r <= sourceSheet.Rows.Count
        originalName = CellText(sourceSheet.Cells(r, "B"))
        startText = CellText(sourceSheet.Cells(r, "C"))
        endText = CellText(sourceSheet.Cells(r, "D"))
        If Len(originalName) > 0 And Len(startText) > 0 And Len(endText) > 0 _
           And InStr(1, startText, "직급", vbBinaryCompare) = 0 Then
            consecutiveNonTargets = 0
            ParseTripDisplayName originalName, employeeName, idPrefix, hasMaskedId, _
                                 sourceSheet.Name & "!B" & CStr(r)
            If gEmployeesByName.Exists(employeeName) Then
                Set candidates = gEmployeesByName(employeeName)
                Set employee = Nothing
                skipRecord = False
                If Not hasMaskedId And candidates.Count = 1 Then
                    Set employee = candidates(1)
                Else
                    ' Always ask for this row, even when another row looked identical.
                    Set selector = New frmTripEmployee
                    selector.Configure candidates, originalName, employeeName, idPrefix, r
                    If gTestAutoSelect Then
                        selector.AutoSelectFirst
                    Else
                        selector.Show vbModal
                    End If
                    If selector.Cancelled Then GoTo Cancelled
                    skipRecord = selector.Skipped
                    Set employee = selector.SelectedEmployee
                    Unload selector
                    Set selector = Nothing
                End If

                If Not skipRecord Then
                    If employee Is Nothing Then
                        RaiseValidation "출장 자료 " & CStr(r) & "행", "선택한 대상자를 확인할 수 없습니다."
                    End If
                    startAt = ParseDateTimeCell(sourceSheet.Cells(r, "C"), ".")
                    endAt = ParseDateTimeCell(sourceSheet.Cells(r, "D"), ".")
                    If startAt >= endAt Then
                        RaiseValidation "출장 자료 " & CStr(r) & "행 (C:D)", _
                                        "시작일시는 종료일시보다 빨라야 합니다."
                    End If
                    Set record = New CTripRecord
                    record.StartAt = startAt
                    record.EndAt = endAt
                    Set records = staged(CStr(employee.Order))
                    records.Add record
                    recordCount = recordCount + 1
                End If
            End If
        Else
            consecutiveNonTargets = consecutiveNonTargets + 1
            If consecutiveNonTargets = 7 Then Exit Do
        End If
        r = r + 1
    Loop

    ' Close only our own input workbook, before touching the existing model.
    CloseInputWorkbook sourceBook, openedByCode
    Set sourceBook = Nothing
    CommitRecordStage staged, "Trip"
    Debug.Print "출장 자료 " & CStr(recordCount) & "건을 불러왔습니다.", vbInformation, "출장 근무상황부"
    Exit Sub

Cancelled:
    On Error Resume Next
    If Not selector Is Nothing Then Unload selector
    Set selector = Nothing
    CloseInputWorkbook sourceBook, openedByCode
    On Error GoTo 0
    Exit Sub

Failed:
    errorText = Err.Description
    If r >= 5 Then errorText = "출장 자료 " & CStr(r) & "행: " & errorText
    On Error Resume Next
    If Not selector Is Nothing Then Unload selector
    Set selector = Nothing
    CloseInputWorkbook sourceBook, openedByCode
    On Error GoTo 0
    Debug.Print errorText & vbCrLf & "기존 출장 자료는 유지됩니다.", vbExclamation, "출장 자료 불러오기 오류"
End Sub

Private Sub ParseTripDisplayName(ByVal originalName As String, ByRef employeeName As String, _
                                 ByRef idPrefix As String, ByRef hasMaskedId As Boolean, _
                                 ByVal context As String)
    Dim openAt As Long, front As String, maskedId As String, i As Long
    Dim character As String
    originalName = TrimInput(originalName)
    employeeName = vbNullString
    idPrefix = vbNullString
    hasMaskedId = False
    If Len(originalName) = 0 Then RaiseValidation context, "이름이 비어 있습니다."

    If InStr(1, originalName, "(", vbBinaryCompare) = 0 _
       And InStr(1, originalName, ")", vbBinaryCompare) = 0 Then
        employeeName = originalName
        Exit Sub
    End If

    openAt = InStrRev(originalName, "(", -1, vbBinaryCompare)
    If openAt <= 1 Or Right$(originalName, 1) <> ")" Then GoTo InvalidFormat
    front = TrimInput(Left$(originalName, openAt - 1))
    maskedId = Mid$(originalName, openAt + 1, Len(originalName) - openAt - 1)
    If InStr(front, "(") > 0 Or InStr(front, ")") > 0 Then GoTo InvalidFormat
    If Len(front) < 3 Then GoTo InvalidFormat
    If Not Right$(front, 2) Like "##" Then GoTo InvalidFormat
    employeeName = TrimInput(Left$(front, Len(front) - 2))
    If Len(employeeName) = 0 Or Len(maskedId) < 4 Then GoTo InvalidFormat

    idPrefix = Left$(maskedId, 3)
    For i = 1 To 3
        character = Mid$(idPrefix, i, 1)
        If character = "*" Or character = "(" Or character = ")" _
           Or Len(TrimInput(character)) = 0 Then GoTo InvalidFormat
    Next i
    For i = 4 To Len(maskedId)
        If Mid$(maskedId, i, 1) <> "*" Then GoTo InvalidFormat
    Next i
    hasMaskedId = True
    Exit Sub

InvalidFormat:
    RaiseValidation context, "이름의 구분 번호·괄호·아이디 마스킹 형식을 확인하세요. " & _
                             "예: 홍길동01 (abc****)"
End Sub
