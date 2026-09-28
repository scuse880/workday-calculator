Attribute VB_Name = "main_정액분계산"
Option Explicit

' 현재 처리 중인 셀(요구사항대로 global)
Private gX As Range

' 시간 역전 오류 누적용(중복 제거)
' key: Sheet!A1|08:30~07:30
' val: 표시 문자열
Private gTimeReverse As Object  ' Scripting.Dictionary (late binding)

'========================================================
' get_times(X)
' X 값: hh:mm~hh:mm [whitespace] hh:mm~hh:mm ... (반복)
' 1) 각 구간을 (text, color, diffMin, startMin, endMin)로 수집
' 2) 같은 color 내에서 겹치는 구간은 merge
' 3) 시간 역전 오류는 즉시 중단하지 않고 누적
' 4) timesColl(Collection) 반환
'========================================================
Public Function get_times(ByVal srcCell As Range) As Collection
    Dim timesColl As New Collection

    Dim s As String
    s = CStr(srcCell.Value2)
    If Len(Trim$(s)) = 0 Then
        Set get_times = timesColl
        Exit Function
    End If

    Dim re As Object, matches As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "(\d{1,2}:\d{2})~(\d{1,2}:\d{2})"

    Set matches = re.Execute(s)
    If matches.Count = 0 Then
        Set get_times = timesColl
        Exit Function
    End If

    ' color(Long) -> Collection(Array(startMin, endMin))
    Dim byColor As Object
    Set byColor = CreateObject("Scripting.Dictionary")

    For Each m In matches
        Dim stStr As String, enStr As String
        stStr = m.SubMatches(0)
        enStr = m.SubMatches(1)

        Dim stMin As Long, enMin As Long
        stMin = TimeStrToMin(stStr)
        enMin = TimeStrToMin(enStr)

        ' 시간 역전 오류는 누적만 하고 이 구간은 건너뜀
        If enMin < stMin Then
            AddTimeReverseError srcCell, stStr, enStr
            GoTo NextMatch
        End If

        ' 리치텍스트 색: 해당 구간 첫 글자의 색으로 판단
        Dim startPos As Long
        startPos = CLng(m.FirstIndex) + 1 ' Characters는 1-based

        Dim col As Long
        col = srcCell.Characters(startPos, 1).Font.color

        If Not byColor.Exists(col) Then byColor.Add col, New Collection
        byColor(col).Add Array(stMin, enMin)

NextMatch:
    Next m

    ' 같은 color 내 겹치는 구간 merge 후 timesColl 생성
    Dim key As Variant
    For Each key In byColor.Keys
        Dim merged As Collection
        Set merged = MergeIntervals(byColor(key))

        Dim j As Long
        For j = 1 To merged.Count
            Dim it As Variant
            it = merged(j) ' Array(stMin, endMin)

            Dim txt As String
            txt = MinToTimeStr(CLng(it(0))) & "~" & MinToTimeStr(CLng(it(1)))

            Dim diff As Long
            diff = CLng(it(1)) - CLng(it(0))

            ' (text, color, diffMin, startMin, endMin)
            timesColl.Add Array(txt, CLng(key), diff, CLng(it(0)), CLng(it(1)))
        Next j
    Next key

    Set get_times = timesColl
End Function

'========================================================
' calc_정액분(timesColl)  [요구 로직 반영]
' 1) 주말(헤더가 토/일)이면 continue
' 2) 빈 셀이면 노란 하이라이트
' 3) vbBlue만 -> continue
' 4) vbRed만 -> 노란
' 5) vbMagenta만 -> 노란
' 6) vbBlue+vbMagenta & blue=08:30~16:30:
'    - magenta 합=480 -> 노란
'    - magenta 합<480 -> continue
' 7) vbBlue+vbRed & blue=08:30~16:30:
'    - red 합=480 -> 노란
'    - red 합<480 -> continue
' 8) 그 외 -> 회색
'========================================================
Public Sub calc_정액분(ByVal timesColl As Collection)
    Dim ws As Worksheet
    Set ws = gX.Worksheet

    ' 1) 주말이면 continue (같은 열 1행 헤더의 요일)
    Dim headerCell As Range
    Set headerCell = ws.Cells(1, gX.Column)

    Dim wk As String
    wk = ExtractWeekdayChar(headerCell)
    If wk = "토" Or wk = "일" Then Exit Sub

    ' 2) 빈 셀이면 노란 하이라이트
    If Len(Trim$(CStr(gX.Value2))) = 0 Then
        SetFill gX, vbYellow
        Exit Sub
    End If

    ' 시간구간이 아예 없으면(형식 불일치/전부 역전 등) -> 9번(그 외)로 회색 처리
    If (timesColl Is Nothing) Or (timesColl.Count = 0) Then
        SetFill gX, RGB(200, 200, 200)
        Exit Sub
    End If

    Dim colors As Object
    Set colors = CreateObject("Scripting.Dictionary")

    Dim blueHasFull As Boolean
    blueHasFull = False

    Dim redTotal As Long
    redTotal = 0

    Dim magentaTotal As Long
    magentaTotal = 0

    ' 8번 규칙용: blue/red 구간 수집
    Dim blueIntervals As New Collection ' Array(stMin, enMin)
    Dim redIntervals As New Collection  ' Array(stMin, enMin)

    Dim i As Long
    For i = 1 To timesColl.Count
        Dim t As Variant
        t = timesColl(i) ' Array(text, color, diff, startMin, endMin)

        Dim col As Long
        col = CLng(t(1))
        If Not colors.Exists(col) Then colors.Add col, True

        Dim stMin As Long, enMin As Long, diff As Long
        stMin = CLng(t(3))
        enMin = CLng(t(4))
        diff = CLng(t(2))

        If col = vbBlue Then
            blueIntervals.Add Array(stMin, enMin)
            If stMin = (8 * 60 + 30) And enMin = (16 * 60 + 30) Then
                blueHasFull = True
            End If
        ElseIf col = vbMagenta Then
            magentaTotal = magentaTotal + diff
        ElseIf col = vbRed Then
            redTotal = redTotal + diff
            redIntervals.Add Array(stMin, enMin)
        End If
    Next i

    ' 3) vbBlue만 -> continue
    If colors.Count = 1 And colors.Exists(vbBlue) Then Exit Sub

    ' 4) vbRed만 -> 노란
    If colors.Count = 1 And colors.Exists(vbRed) Then
        SetFill gX, vbYellow
        Exit Sub
    End If

    ' 5) vbMagenta만 -> 노란
    If colors.Count = 1 And colors.Exists(vbMagenta) Then
        SetFill gX, vbYellow
        Exit Sub
    End If

    ' 6) vbBlue + vbMagenta
    If colors.Count = 2 And colors.Exists(vbBlue) And colors.Exists(vbMagenta) Then
        If blueHasFull Then
            If magentaTotal = 480 Then
                SetFill gX, vbYellow
                Exit Sub
            ElseIf magentaTotal < 480 Then
                Exit Sub
            End If
        End If
        SetFill gX, RGB(200, 200, 200)
        Exit Sub
    End If

    ' 7) vbBlue + vbRed  (+ 8번 추가 규칙)
    If colors.Count = 2 And colors.Exists(vbBlue) And colors.Exists(vbRed) Then

        ' 7) blue=08:30~16:30이면 red 합계로 판정
        If blueHasFull Then
            If redTotal = 480 Then
                SetFill gX, vbYellow
                Exit Sub
            ElseIf redTotal < 480 Then
                Exit Sub
            End If
        End If

        ' 8) blue/red가 서로 안 겹치고, 둘 다 08:30~16:30 범위 "안"이면 continue
        Dim windowStart As Long, windowEnd As Long
        windowStart = 8 * 60 + 30   ' 08:30
        windowEnd = 16 * 60 + 30    ' 16:30

        Dim withinOK As Boolean
        withinOK = True

        Dim b As Variant, r As Variant

        ' blue 범위 체크
        For Each b In blueIntervals
            If CLng(b(0)) < windowStart Or CLng(b(1)) > windowEnd Then
                withinOK = False
                Exit For
            End If
        Next b

        ' red 범위 체크
        If withinOK Then
            For Each r In redIntervals
                If CLng(r(0)) < windowStart Or CLng(r(1)) > windowEnd Then
                    withinOK = False
                    Exit For
                End If
            Next r
        End If

        ' 겹침 체크(끝=시작은 겹침 아님)
        If withinOK Then
            Dim overlap As Boolean
            overlap = False

            For Each b In blueIntervals
                For Each r In redIntervals
                    If Not (CLng(b(1)) <= CLng(r(0)) Or CLng(r(1)) <= CLng(b(0))) Then
                        overlap = True
                        Exit For
                    End If
                Next r
                If overlap Then Exit For
            Next b

            If Not overlap Then Exit Sub ' 8번 만족 -> continue
        End If

        ' 9) 그 외 -> 회색
        SetFill gX, RGB(200, 200, 200)
        Exit Sub
    End If

    ' 9) 그 외 -> 회색
    SetFill gX, RGB(200, 200, 200)
End Sub

'========================================================
' main_정액분계산
' out!B2:AF80 지그재그(행 단위 왕복) 순회
'========================================================
Public Sub main_정액분계산()
    Dim ws As Worksheet
    Dim rng As Range
    Dim r As Long, c As Long
    Dim r0 As Long, r1 As Long, c0 As Long, c1 As Long
    Dim timesColl As Collection

    Set ws = ThisWorkbook.Worksheets("out")
    Set rng = ws.Range("B2:AF80")

    r0 = rng.Row
    r1 = rng.Row + rng.Rows.Count - 1
    c0 = rng.Column
    c1 = rng.Column + rng.Columns.Count - 1

    Set gTimeReverse = CreateObject("Scripting.Dictionary")

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo FIN

    For r = r0 To r1
        If ((r - r0) Mod 2) = 0 Then
            For c = c0 To c1
                Set gX = ws.Cells(r, c)
                Set timesColl = get_times(gX)
                calc_정액분 timesColl
            Next c
        Else
            For c = c1 To c0 Step -1
                Set gX = ws.Cells(r, c)
                Set timesColl = get_times(gX)
                calc_정액분 timesColl
            Next c
        End If
    Next r

FIN:
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical
        Exit Sub
    End If

    ShowTimeReverseErrors
End Sub

'========================================================
' Helpers
'========================================================
Private Sub SetFill(ByVal cell As Range, ByVal fillColor As Long)
    With cell.Interior
        .Pattern = xlSolid
        .color = fillColor
    End With
End Sub

Private Function TimeStrToMin(ByVal hhmm As String) As Long
    Dim p() As String
    p = Split(hhmm, ":")
    If UBound(p) <> 1 Then ErrStop "시간 파싱 실패: " & hhmm
    TimeStrToMin = CLng(p(0)) * 60 + CLng(p(1))
End Function

Private Function MinToTimeStr(ByVal mins As Long) As String
    Dim hh As Long, nn As Long
    hh = mins \ 60
    nn = mins Mod 60
    MinToTimeStr = Right$("0" & CStr(hh), 2) & ":" & Right$("0" & CStr(nn), 2)
End Function

Private Sub AddTimeReverseError(ByVal srcCell As Range, ByVal stStr As String, ByVal enStr As String)
    If gTimeReverse Is Nothing Then Set gTimeReverse = CreateObject("Scripting.Dictionary")

    Dim key As String
    key = srcCell.Worksheet.Name & "!" & srcCell.Address(False, False) & "|" & stStr & "~" & enStr

    If Not gTimeReverse.Exists(key) Then
        gTimeReverse.Add key, srcCell.Worksheet.Name & "!" & srcCell.Address(False, False) & " : " & stStr & "~" & enStr
    End If
End Sub

Private Sub ShowTimeReverseErrors()
    If gTimeReverse Is Nothing Then Exit Sub
    If gTimeReverse.Count = 0 Then Exit Sub

    Dim msg As String
    Dim v As Variant

    msg = "시간 역전 오류 셀:" & vbCrLf & vbCrLf
    For Each v In gTimeReverse.Items
        msg = msg & CStr(v) & vbCrLf
    Next v

    MsgBox msg, vbExclamation
End Sub

' 같은 색상 내 구간들(시작/끝 분)을 정렬 후 겹치면 merge
Private Function MergeIntervals(ByVal coll As Collection) As Collection
    Dim out As New Collection
    If coll.Count = 0 Then
        Set MergeIntervals = out
        Exit Function
    End If

    Dim n As Long: n = coll.Count
    Dim s() As Long, e() As Long
    ReDim s(1 To n)
    ReDim e(1 To n)

    Dim i As Long, j As Long
    For i = 1 To n
        Dim it As Variant
        it = coll(i)
        s(i) = CLng(it(0))
        e(i) = CLng(it(1))
    Next i

    ' sort by start (버블 정렬)
    For i = 1 To n - 1
        For j = i + 1 To n
            If s(j) < s(i) Then
                Dim ts As Long, te As Long
                ts = s(i): te = e(i)
                s(i) = s(j): e(i) = e(j)
                s(j) = ts: e(j) = te
            End If
        Next j
    Next i

    ' merge
    Dim curS As Long, curE As Long
    Dim k As Long
    curS = s(1): curE = e(1)

    For k = 2 To n
        If s(k) <= curE Then
            If e(k) > curE Then curE = e(k)
        Else
            out.Add Array(curS, curE)
            curS = s(k): curE = e(k)
        End If
    Next k
    out.Add Array(curS, curE)

    Set MergeIntervals = out
End Function

' 헤더 셀에서 요일 한 글자("토","일",...) 추출 (표시 텍스트 기준)
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

Private Sub ErrStop(ByVal msg As String)
    Err.Raise vbObjectError + 513, "VBA", msg
End Sub
