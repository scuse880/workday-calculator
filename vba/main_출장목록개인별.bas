Attribute VB_Name = "main_출장목록개인별"
Option Explicit

' main에서 초기화해서 쓰는 "out A열에 없는 대상자" 누적용(중복 제거)
Private gMissing As Object  ' Scripting.Dictionary (late binding)

'========================================================
' is_valid(x)
' - startCell부터 같은 열에서 아래로 내려가며 탐색
' - 값이 "부산공업고등학교" / "성명" / "" 이면 skip
' - ""(빈셀)만 "연속 20개"면 False 반환
' 반환: Array(found As Boolean, value As String, row As Long, col As Long, nextCell As Range)
'========================================================
Public Function is_valid(ByVal startCell As Range) As Variant
    Dim ws As Worksheet: Set ws = startCell.Worksheet
    Dim r As Long, c As Long
    r = startCell.Row
    c = startCell.Column

    Dim blankStreak As Long
    blankStreak = 0

    Do While r <= ws.Rows.Count
        Dim v As String
        v = Trim$(CStr(ws.Cells(r, c).Value))

        If v = "" Then
            blankStreak = blankStreak + 1
            If blankStreak >= 20 Then
                is_valid = Array(False)
                Exit Function
            End If
        Else
            blankStreak = 0
            If v <> "부산공업고등학교" And v <> "성명" Then
                Dim found As Range
                Set found = ws.Cells(r, c)
                is_valid = Array(True, v, found.Row, found.Column, found.Offset(1, 0))
                Exit Function
            End If
        End If

        r = r + 1
    Loop

    is_valid = Array(False)
End Function

'========================================================
' func_name(var)
'========================================================
Public Function func_name(ByVal v As Variant) As String
    Dim s As String
    s = Trim$(GetValAsString(v))

    If Len(s) <= 3 Then
        func_name = s
    ElseIf Len(s) >= 4 And IsAlpha(Mid$(s, 4, 1)) Then
        func_name = Left$(s, 4)
    Else
        func_name = Left$(s, 3)
    End If
End Function

'========================================================
' func_date(var)
' 입력: "yyyy.mm.dd hh:mm ~ yyyy.mm.dd hh:mm"
' 출력: Collection of Array(mm.dd, startHH:MM, endHH:MM)
'       중 out!B1과 같은 월만
'
' 규칙:
' - 기간이 이틀 이상이어도 매일 시작/종료시간은 타겟 셀에 적힌 시간 그대로 사용
'========================================================
Public Function func_date(ByVal v As Variant) As Collection
    Dim s As String
    s = Trim$(GetValAsString(v))
    
    Dim parts() As String
    parts = Split(s, "~")
    If UBound(parts) <> 1 Then ErrStop "기간 문자열 형식 오류: " & s
    
    Dim startDT As Date, endDT As Date
    startDT = ParseDateTimeDot(Trim$(parts(0)))
    endDT = ParseDateTimeDot(Trim$(parts(1)))
    
    If endDT < startDT Then
        Dim tmp As Date
        tmp = startDT
        startDT = endDT
        endDT = tmp
    End If
    
    Dim startDate As Date, endDate As Date
    startDate = DateSerial(Year(startDT), Month(startDT), day(startDT))
    endDate = DateSerial(Year(endDT), Month(endDT), day(endDT))
    
    Dim st As String, en As String
    st = Format$(startDT, "hh:nn")
    en = Format$(endDT, "hh:nn")
    
    ' out!B1에서 대상 월 자동 취득
    Dim targetMonth As Long
    Dim dummyDay As Long
    
    ExtractMonthDayFromHeaderCell _
        ThisWorkbook.Worksheets("out").Cells(1, 2), _
        targetMonth, _
        dummyDay
    
    Dim col As New Collection
    Dim d As Date
    
    For d = startDate To endDate
        Dim mmdd As String
        mmdd = Format$(d, "mm.dd")
        
        ' 기존:
        ' If Left$(mmdd, 2) = "07" Then
        '
        ' 변경:
        If Month(d) = targetMonth Then
            col.Add Array(mmdd, st, en)
        End If
    Next d
    
    Set func_date = col
End Function

'========================================================
' func_move(y_, x_)
' - x_가 out A열(out_a)에 없으면: continue + 누적 저장(중복 제거)
' - 주말(토/일)이면 continue
' - 입력은 빨간 글씨, 기존 글씨 색은 건드리지 않음
'========================================================
Public Sub func_move(ByVal y_ As Collection, ByVal x_ As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("out")

    ' out_b: B1부터 오른쪽으로 빈 셀 전까지
    Dim headerLastCol As Long
    headerLastCol = GetHeaderLastCol(ws, 1, 2)
    If headerLastCol < 2 Then ErrStop "'out' 시트 B1부터 헤더가 없습니다."

    ' out_a: A2부터 아래로 빈 셀 전까지에서 x_ 찾기
    Dim nameRow As Long
    nameRow = FindNameRowInOutA(ws, Trim$(x_), 2)

    ' 이름 못 찾으면 누적만 하고 종료
    If nameRow = 0 Then
        If gMissing Is Nothing Then Set gMissing = CreateObject("Scripting.Dictionary")
        If Not gMissing.Exists(Trim$(x_)) Then gMissing.Add Trim$(x_), True
        Exit Sub
    End If

    Dim i As Long
    For i = 1 To y_.Count
        Dim item As Variant
        item = y_(i) ' Array(mm.dd, start, end)

        Dim mmdd As String, st As String, en As String
        mmdd = CStr(item(0))
        st = CStr(item(1))
        en = CStr(item(2))

        Dim m As Long, d As Long
        m = CLng(Left$(mmdd, 2))
        d = CLng(Right$(mmdd, 2))

        ' 헤더에서 월/일 유일 매칭
        Dim headerCell As Range
        Set headerCell = FindUniqueHeaderByMonthDay(ws, 1, 2, headerLastCol, m, d)

        ' 토/일이면 continue
        Dim wk As String
        wk = ExtractWeekdayChar(headerCell)
        If wk = "토" Or wk = "일" Then GoTo NextItem

        ' 교차셀에 추가
        Dim target As Range
        Set target = ws.Cells(nameRow, headerCell.Column)

        Dim newText As String
        newText = st & "~" & en

        AppendColoredLine target, newText, vbRed

NextItem:
    Next i
End Sub

'========================================================
' main_출장목록_개인별
' ※ VBA 프로시저명에는 괄호를 넣을 수 없어서 이렇게 사용
'========================================================
Public Sub main_출장목록_개인별()
    Dim wsSrc As Worksheet
    Set wsSrc = ThisWorkbook.Worksheets("출장목록(개인별)")

    Set gMissing = CreateObject("Scripting.Dictionary")

    Dim cur As Range
    Set cur = wsSrc.Range("B5")

    Do
        Dim res As Variant
        res = is_valid(cur)
        If Not IsArray(res) Then Exit Do
        If res(0) = False Then Exit Do

        Dim srcRow As Long
        Dim rawName As String
        Dim x_ As String
        rawName = CStr(res(1))
        srcRow = CLng(res(2))

        ' 이름 가공
        x_ = func_name(rawName)

        ' "x_행 c열 값" + " ~ " + "x_행 d열 값"
        Dim s As String
        s = GetCellAsDateTimeDot(wsSrc.Cells(srcRow, "C")) & " ~ " & GetCellAsDateTimeDot(wsSrc.Cells(srcRow, "D"))

        Dim y_ As Collection
        Set y_ = func_date(s)

        func_move y_, x_

        Set cur = res(4) ' 다음 탐색 시작 셀
    Loop

    ' out_a에 없던 대상자 알림(중복 제거)
    If Not gMissing Is Nothing Then
        If gMissing.Count > 0 Then
            Dim k As Variant, msg As String
            msg = "out 시트 A열(out_a)에서 찾지 못해 건너뛴 대상자:" & vbCrLf & vbCrLf
            For Each k In gMissing.Keys
                msg = msg & CStr(k) & vbCrLf
            Next k
            MsgBox msg, vbExclamation
        End If
    End If
End Sub

'========================================================
' Helpers
'========================================================
Private Function GetValAsString(ByVal v As Variant) As String
    If TypeName(v) = "Range" Then
        GetValAsString = CStr(v.Value)
    Else
        GetValAsString = CStr(v)
    End If
End Function

Private Function IsAlpha(ByVal ch As String) As Boolean
    If Len(ch) <> 1 Then
        IsAlpha = False
    Else
        Dim code As Long
        code = AscW(ch)
        IsAlpha = (code >= 65 And code <= 90) Or (code >= 97 And code <= 122)
    End If
End Function

Private Sub ErrStop(ByVal msg As String)
    MsgBox msg, vbCritical
    End
End Sub

' yyyy.mm.dd hh:mm -> Date
Private Function ParseDateTimeDot(ByVal s As String) As Date
    s = Trim$(s)

    Dim a() As String
    a = Split(s, " ")
    If UBound(a) <> 1 Then ErrStop "날짜시간 파싱 실패: " & s

    Dim dtPart As String, tmPart As String
    dtPart = a(0)
    tmPart = a(1)

    Dim d() As String
    d = Split(dtPart, ".")
    If UBound(d) <> 2 Then ErrStop "날짜 파싱 실패: " & dtPart

    Dim t() As String
    t = Split(tmPart, ":")
    If UBound(t) <> 1 Then ErrStop "시간 파싱 실패: " & tmPart

    Dim yy As Long, mm As Long, dd As Long, hh As Long, nn As Long
    yy = CLng(d(0))
    mm = CLng(d(1))
    dd = CLng(d(2))
    hh = CLng(t(0))
    nn = CLng(t(1))

    ParseDateTimeDot = DateSerial(yy, mm, dd) + TimeSerial(hh, nn, 0)
End Function

' 셀 값이 날짜/시간 값일 수도 있으니 "yyyy.mm.dd hh:mm" 문자열로 안전 변환
Private Function GetCellAsDateTimeDot(ByVal cell As Range) As String
    If IsDate(cell.Value) Then
        GetCellAsDateTimeDot = Format$(CDate(cell.Value), "yyyy.mm.dd hh:nn")
    Else
        GetCellAsDateTimeDot = Trim$(CStr(cell.Value))
    End If
End Function

' out 시트 1행에서 B부터 "첫 빈셀 직전" 열 반환
Private Function GetHeaderLastCol(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal startCol As Long) As Long
    Dim c As Long
    c = startCol

    If Trim$(CStr(ws.Cells(headerRow, c).Value)) = "" Then
        GetHeaderLastCol = startCol - 1
        Exit Function
    End If

    Do While Trim$(CStr(ws.Cells(headerRow, c).Value)) <> ""
        c = c + 1
    Loop
    GetHeaderLastCol = c - 1
End Function

' out A열(AstartRow부터, 빈 셀 전)에서 keyVal 찾기
' - 0개: 0 반환
' - 1개: 해당 행 반환
' - 2개 이상: 에러 종료
Private Function FindNameRowInOutA(ByVal ws As Worksheet, ByVal keyVal As String, ByVal startRow As Long) As Long
    Dim endRow As Long
    endRow = startRow
    Do While Trim$(CStr(ws.Cells(endRow, "A").Value)) <> ""
        endRow = endRow + 1
    Loop
    endRow = endRow - 1
    If endRow < startRow Then ErrStop "'out' 시트 A" & startRow & "부터 데이터가 없습니다."

    Dim rng As Range
    Set rng = ws.Range("A" & startRow & ":A" & endRow)

    Dim cnt As Long
    cnt = Application.WorksheetFunction.CountIf(rng, keyVal)

    If cnt = 0 Then
        FindNameRowInOutA = 0
        Exit Function
    End If
    If cnt <> 1 Then
        ErrStop "out_a(A" & startRow & "~)에서 '" & keyVal & "' 일치 개수=" & cnt & " (2개 이상)"
    End If

    Dim f As Range
    Set f = rng.Find(What:=keyVal, LookIn:=xlValues, LookAt:=xlWhole)
    If f Is Nothing Then
        FindNameRowInOutA = 0
    Else
        FindNameRowInOutA = f.Row
    End If
End Function

' 헤더(1행 B~lastCol)에서 "m/d(요일)"의 월/일이 (m,d)와 일치하는 셀을 유일하게 찾기
Private Function FindUniqueHeaderByMonthDay(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal startCol As Long, ByVal lastCol As Long, ByVal m As Long, ByVal d As Long) As Range
    Dim c As Long, cnt As Long
    Dim hit As Range

    For c = startCol To lastCol
        Dim cell As Range
        Set cell = ws.Cells(headerRow, c)

        Dim hm As Long, hd As Long
        ExtractMonthDayFromHeaderCell cell, hm, hd

        If hm = m And hd = d Then
            cnt = cnt + 1
            Set hit = cell
        End If
    Next c

    If cnt <> 1 Then
        ErrStop "out_b(1행)에서 " & m & "/" & d & " 일치 개수=" & cnt & " (없거나 2개 이상)"
    End If

    Set FindUniqueHeaderByMonthDay = hit
End Function

' 헤더 셀에서 월/일 추출 (표시 텍스트 "1/4(목)" 기준)
Private Sub ExtractMonthDayFromHeaderCell(ByVal cell As Range, ByRef outM As Long, ByRef outD As Long)
    Dim txt As String
    txt = cell.Text
    If Len(Trim$(txt)) = 0 Then txt = CStr(cell.Value)

    If IsDate(cell.Value) Then
        outM = Month(CDate(cell.Value))
        outD = day(CDate(cell.Value))
        Exit Sub
    End If

    Dim p As Long
    p = InStr(1, txt, "/", vbTextCompare)
    If p = 0 Then ErrStop "헤더 파싱 실패(슬래시 없음): " & cell.Address(0, 0) & " [" & txt & "]"

    outM = CLng(Val(Left$(txt, p - 1)))
    outD = CLng(Val(Mid$(txt, p + 1))) ' 뒤에 "(요일)" 있어도 Val로 잘림
End Sub

' 헤더 셀에서 요일 한 글자("토","일",...) 추출
Private Function ExtractWeekdayChar(ByVal cell As Range) As String
    Dim txt As String
    txt = cell.Text
    If Len(Trim$(txt)) = 0 Then txt = CStr(cell.Value)

    Dim p1 As Long, p2 As Long
    p1 = InStr(1, txt, "(", vbTextCompare)
    p2 = InStr(1, txt, ")", vbTextCompare)

    If p1 > 0 And p2 > p1 Then
        ExtractWeekdayChar = Mid$(txt, p1 + 1, 1)
    Else
        ExtractWeekdayChar = ""
    End If
End Function

' 기존 텍스트 색 유지 + 새로 추가된 newText 부분만 color로 지정
Private Sub AppendColoredLine(ByVal target As Range, ByVal newText As String, ByVal color As Long)
    Dim oldLen As Long, startPos As Long, addLen As Long
    oldLen = Len(CStr(target.Value))

    If Len(Trim$(CStr(target.Value))) > 0 Then
        target.Value = CStr(target.Value) & vbLf & newText
        startPos = oldLen + 2 ' vbLf(1) + 1-based 보정
    Else
        target.Value = newText
        startPos = 1
    End If

    addLen = Len(newText)
    target.WrapText = True
    target.Characters(startPos, addLen).Font.color = color
End Sub
