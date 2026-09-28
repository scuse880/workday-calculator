Attribute VB_Name = "main_근무상황목록"

Option Explicit

Private gMissingNames As Collection

Public Function func_name(ByVal v As Variant) As String
    Dim s As String
    s = Trim$(GetText(v))
    
    If Len(s) <= 3 Then
        func_name = s
    ElseIf IsAsciiAlpha(Mid$(s, 4, 1)) Then
        func_name = Left$(s, 4)
    Else
        func_name = Left$(s, 3)
    End If
End Function

Public Function func_date(ByVal v As Variant) As Collection
    Dim raw As String
    Dim leftPart As String, rightPart As String
    Dim p As Long
    
    Dim sy As Long, sm As Long, sd As Long, sh As Long, sn As Long
    Dim ey As Long, em As Long, ed As Long, EH As Long, en As Long
    
    Dim totalDays As Long
    Dim cy As Long, cm As Long, cd As Long
    Dim i As Long
    
    Dim st As String, et As String
    Dim col As Collection
    
    ' out!B1에서 현재 대상 월 자동 취득
    Dim targetMonth As Long
    targetMonth = GetTargetMonthFromOutHeader()
    
    raw = Trim$(GetText(v))
    p = InStr(raw, "~")
    If p = 0 Then
        Err.Raise vbObjectError + 1001, "func_date", _
                  "날짜 문자열 형식이 올바르지 않습니다: " & raw
    End If
    
    leftPart = Trim$(Left$(raw, p - 1))
    rightPart = Trim$(Mid$(raw, p + 1))
    
    ParseDateTimeText leftPart, sy, sm, sd, sh, sn
    ParseDateTimeText rightPart, ey, em, ed, EH, en
    
    If IsDateEarlier(ey, em, ed, sy, sm, sd) Then
        Err.Raise vbObjectError + 1002, "func_date", _
                  "종료일이 시작일보다 빠릅니다: " & raw
    End If
    
    totalDays = DaysBetweenInclusiveSimple(sy, sm, sd, ey, em, ed)
    
    Set col = New Collection
    
    cy = sy
    cm = sm
    cd = sd
    
    For i = 1 To totalDays
        If totalDays = 1 Then
            st = FormatTimeText(sh, sn)
            et = FormatTimeText(EH, en)
            
        ElseIf i = 1 Then
            st = FormatTimeText(sh, sn)
            et = "16:30"
            
        ElseIf i = totalDays Then
            st = "08:30"
            et = FormatTimeText(EH, en)
            
        Else
            st = "08:30"
            et = "16:30"
        End If
        
        ' 기존: If cm = 7 Then
        ' 변경: out!B1의 월과 일치할 때만 추가
        If cm = targetMonth Then
            col.Add Array(Format$(cm, "00") & "-" & _
                          Format$(cd, "00"), st, et)
        End If
        
        AddOneDaySimple cy, cm, cd
    Next i
    
    Set func_date = col
End Function


'========================================================
' out!B1에서 대상 월 자동 취득
' 예:
'   8/1(토) -> 8
'   실제 Excel 날짜값 -> Month()로 8
'========================================================
Private Function GetTargetMonthFromOutHeader() As Long
    Dim ws As Worksheet
    Dim cell As Range
    Dim txt As String
    Dim p As Long
    Dim m As Long
    
    Set ws = ThisWorkbook.Worksheets("out")
    Set cell = ws.Cells(1, 2)   ' B1
    
    ' B1이 실제 Excel 날짜값인 경우
    If IsDate(cell.Value) Then
        m = Month(CDate(cell.Value))
    Else
        ' B1이 "8/1(토)" 같은 문자열인 경우
        txt = Trim$(cell.Text)
        If Len(txt) = 0 Then txt = Trim$(CStr(cell.Value))
        
        p = InStr(1, txt, "/", vbTextCompare)
        If p = 0 Then
            Err.Raise vbObjectError + 1503, _
                      "GetTargetMonthFromOutHeader", _
                      """out""!B1에서 월을 읽을 수 없습니다: " & txt
        End If
        
        m = CLng(Val(Left$(txt, p - 1)))
    End If
    
    If m < 1 Or m > 12 Then
        Err.Raise vbObjectError + 1504, _
                  "GetTargetMonthFromOutHeader", _
                  """out""!B1의 월 값이 올바르지 않습니다: " & m
    End If
    
    GetTargetMonthFromOutHeader = m
End Function

Public Sub func_move(ByVal y_ As Collection, ByVal x_ As String)
    Dim ws As Worksheet
    Dim nameCell As Range
    Dim dateCell As Range
    Dim item As Variant
    Dim outText As String
    Dim targetCell As Range
    
    Set ws = ThisWorkbook.Worksheets("out")
    Set nameCell = FindUniqueNameCell(ws, x_)
    
    If nameCell Is Nothing Then Exit Sub
    
    For Each item In y_
        Set dateCell = FindUniqueHeaderCell(ws, CStr(item(0)))
        
        If IsWeekendHeader(CStr(dateCell.Value)) Then
            GoTo ContinueLoop
        End If
        
        outText = CStr(item(1)) & "~" & CStr(item(2))
        Set targetCell = ws.Cells(nameCell.Row, dateCell.Column)
        
        If Len(CStr(targetCell.Value)) > 0 Then
            targetCell.Value = CStr(targetCell.Value) & vbLf & outText
        Else
            targetCell.Value = outText
        End If
        
        targetCell.WrapText = True
        targetCell.Font.color = RGB(0, 0, 255)
        
ContinueLoop:
    Next item
End Sub

Public Sub main_근무상황목록()
    Dim wsSrc As Worksheet
    Dim r As Long
    Dim x_ As String
    Dim y_ As Collection
    Dim msg As String
    Dim i As Long
    
    On Error GoTo EH
    
    Set wsSrc = ThisWorkbook.Worksheets("근무상황목록")
    Set gMissingNames = New Collection
    
    Application.ScreenUpdating = False
    
    r = 2
    Do While Len(Trim$(CStr(wsSrc.Cells(r, "C").Value))) > 0
        x_ = func_name(wsSrc.Cells(r, "C").Value)
        Set y_ = func_date(wsSrc.Cells(r, "E").Value)
        func_move y_, x_
        r = r + 1
    Loop
    
    Application.ScreenUpdating = True
    
    If gMissingNames.Count > 0 Then
        msg = """out"" 시트 A열에서 찾지 못한 이름:" & vbCrLf & vbCrLf
        For i = 1 To gMissingNames.Count
            msg = msg & "- " & CStr(gMissingNames(i)) & vbCrLf
        Next i
        MsgBox msg, vbExclamation
    End If
    
    Exit Sub
    
EH:
    Application.ScreenUpdating = True
    MsgBox "오류 발생 (행 " & r & "): " & Err.Description, vbCritical
End Sub

Private Sub AddMissingName(ByVal targetName As String)
    If gMissingNames Is Nothing Then Set gMissingNames = New Collection
    
    On Error Resume Next
    gMissingNames.Add targetName, CStr(targetName)
    On Error GoTo 0
End Sub

Private Function GetText(ByVal v As Variant) As String
    If IsObject(v) Then
        If TypeName(v) = "Range" Then
            GetText = CStr(v.Value)
        Else
            GetText = CStr(v)
        End If
    Else
        GetText = CStr(v)
    End If
End Function

Private Function IsAsciiAlpha(ByVal ch As String) As Boolean
    Dim code As Long
    
    If Len(ch) <> 1 Then Exit Function
    
    code = Asc(UCase$(ch))
    IsAsciiAlpha = (code >= 65 And code <= 90)
End Function

Private Sub ParseDateTimeText(ByVal s As String, _
                              ByRef yy As Long, ByRef mm As Long, ByRef dd As Long, _
                              ByRef hh As Long, ByRef nn As Long)
    Dim arr() As String
    Dim dPart As String, tPart As String
    Dim dArr() As String, tArr() As String
    
    s = NormalizeSpaces(Trim$(s))
    arr = Split(s, " ")
    
    If UBound(arr) <> 1 Then
        Err.Raise vbObjectError + 1101, "ParseDateTimeText", "날짜/시간 형식이 올바르지 않습니다: " & s
    End If
    
    dPart = arr(0)
    tPart = arr(1)
    
    dArr = Split(dPart, "-")
    tArr = Split(tPart, ":")
    
    If UBound(dArr) <> 2 Or UBound(tArr) <> 1 Then
        Err.Raise vbObjectError + 1102, "ParseDateTimeText", "날짜/시간 형식이 올바르지 않습니다: " & s
    End If
    
    yy = CLng(dArr(0))
    mm = CLng(dArr(1))
    dd = CLng(dArr(2))
    hh = CLng(tArr(0))
    nn = CLng(tArr(1))
    
    If mm < 1 Or mm > 12 Then
        Err.Raise vbObjectError + 1103, "ParseDateTimeText", "월 값이 올바르지 않습니다: " & s
    End If
    
    If dd < 1 Or dd > DaysInMonthSimple(yy, mm) Then
        Err.Raise vbObjectError + 1104, "ParseDateTimeText", "일 값이 올바르지 않습니다: " & s
    End If
    
    If hh < 0 Or hh > 23 Or nn < 0 Or nn > 59 Then
        Err.Raise vbObjectError + 1105, "ParseDateTimeText", "시간 값이 올바르지 않습니다: " & s
    End If
End Sub

Private Function NormalizeSpaces(ByVal s As String) As String
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormalizeSpaces = s
End Function

Private Function FormatTimeText(ByVal hh As Long, ByVal nn As Long) As String
    FormatTimeText = Format$(hh, "00") & ":" & Format$(nn, "00")
End Function

Private Function DaysInMonthSimple(ByVal yy As Long, ByVal mm As Long) As Long
    Select Case mm
        Case 1, 3, 5, 7, 8, 10, 12
            DaysInMonthSimple = 31
        Case 4, 6, 9, 11
            DaysInMonthSimple = 30
        Case 2
            If yy Mod 4 = 0 Then
                DaysInMonthSimple = 29
            Else
                DaysInMonthSimple = 28
            End If
        Case Else
            Err.Raise vbObjectError + 1201, "DaysInMonthSimple", "월 값이 올바르지 않습니다: " & mm
    End Select
End Function

Private Sub AddOneDaySimple(ByRef yy As Long, ByRef mm As Long, ByRef dd As Long)
    dd = dd + 1
    If dd > DaysInMonthSimple(yy, mm) Then
        dd = 1
        mm = mm + 1
        If mm > 12 Then
            mm = 1
            yy = yy + 1
        End If
    End If
End Sub

Private Function IsDateEarlier(ByVal y1 As Long, ByVal m1 As Long, ByVal d1 As Long, _
                               ByVal y2 As Long, ByVal m2 As Long, ByVal d2 As Long) As Boolean
    If y1 <> y2 Then
        IsDateEarlier = (y1 < y2)
    ElseIf m1 <> m2 Then
        IsDateEarlier = (m1 < m2)
    Else
        IsDateEarlier = (d1 < d2)
    End If
End Function

Private Function IsSameDate(ByVal y1 As Long, ByVal m1 As Long, ByVal d1 As Long, _
                            ByVal y2 As Long, ByVal m2 As Long, ByVal d2 As Long) As Boolean
    IsSameDate = (y1 = y2 And m1 = m2 And d1 = d2)
End Function

Private Function DaysBetweenInclusiveSimple(ByVal sy As Long, ByVal sm As Long, ByVal sd As Long, _
                                            ByVal ey As Long, ByVal em As Long, ByVal ed As Long) As Long
    Dim cy As Long, cm As Long, cd As Long
    Dim cnt As Long
    
    If IsDateEarlier(ey, em, ed, sy, sm, sd) Then
        Err.Raise vbObjectError + 1301, "DaysBetweenInclusiveSimple", "종료일이 시작일보다 빠릅니다."
    End If
    
    cy = sy: cm = sm: cd = sd
    cnt = 1
    
    Do While Not IsSameDate(cy, cm, cd, ey, em, ed)
        AddOneDaySimple cy, cm, cd
        cnt = cnt + 1
    Loop
    
    DaysBetweenInclusiveSimple = cnt
End Function

Private Function FindUniqueNameCell(ByVal ws As Worksheet, ByVal targetName As String) As Range
    Dim r As Long
    Dim cnt As Long
    Dim hit As Range
    
    r = 2
    Do While Len(Trim$(CStr(ws.Cells(r, "A").Value))) > 0
        If CStr(ws.Cells(r, "A").Value) = targetName Then
            cnt = cnt + 1
            If cnt = 1 Then Set hit = ws.Cells(r, "A")
        End If
        r = r + 1
    Loop
    
    If cnt = 0 Then
        AddMissingName targetName
        Set FindUniqueNameCell = Nothing
        Exit Function
    ElseIf cnt > 1 Then
        Err.Raise vbObjectError + 1402, "FindUniqueNameCell", """out"" 시트 A열에서 """ & targetName & """ 이(가) 중복됩니다."
    End If
    
    Set FindUniqueNameCell = hit
End Function

Private Function FindUniqueHeaderCell(ByVal ws As Worksheet, ByVal mmdd As String) As Range
    Dim c As Long
    Dim cnt As Long
    Dim hit As Range
    
    c = 2
    Do While Len(Trim$(CStr(ws.Cells(1, c).Value))) > 0
        If HeaderMatchesMMDD(CStr(ws.Cells(1, c).Value), mmdd) Then
            cnt = cnt + 1
            If cnt = 1 Then Set hit = ws.Cells(1, c)
        End If
        c = c + 1
    Loop
    
    If cnt = 0 Then
        Err.Raise vbObjectError + 1501, "FindUniqueHeaderCell", """out"" 시트 1행에서 날짜 """ & mmdd & """ 을(를) 찾지 못했습니다."
    ElseIf cnt > 1 Then
        Err.Raise vbObjectError + 1502, "FindUniqueHeaderCell", """out"" 시트 1행에서 날짜 """ & mmdd & """ 이(가) 중복됩니다."
    End If
    
    Set FindUniqueHeaderCell = hit
End Function

Private Function HeaderMatchesMMDD(ByVal headerText As String, ByVal mmdd As String) As Boolean
    Dim s As String
    Dim p As Long
    Dim datePart As String
    Dim arr() As String
    Dim m As Long, d As Long
    
    s = Trim$(headerText)
    p = InStr(s, "(")
    
    If p > 0 Then
        datePart = Trim$(Left$(s, p - 1))
    Else
        datePart = s
    End If
    
    arr = Split(datePart, "/")
    If UBound(arr) <> 1 Then Exit Function
    
    m = CLng(arr(0))
    d = CLng(arr(1))
    
    HeaderMatchesMMDD = (Format$(m, "00") & "-" & Format$(d, "00") = mmdd)
End Function

Private Function IsWeekendHeader(ByVal headerText As String) As Boolean
    Dim s As String
    Dim p1 As Long, p2 As Long
    Dim wd As String
    
    s = CStr(headerText)
    p1 = InStr(s, "(")
    p2 = InStr(s, ")")
    
    If p1 > 0 And p2 > p1 Then
        wd = Mid$(s, p1 + 1, p2 - p1 - 1)
    Else
        wd = s
    End If
    
    IsWeekendHeader = (wd = "토" Or wd = "일")
End Function

