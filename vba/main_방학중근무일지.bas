Attribute VB_Name = "main_방학중근무일지"
Option Explicit

Private gMissingNames As Collection

'========================
' func_name(var)
'========================
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

'========================
' func_date(var)
' 전제: "yyyymmdd hhmm~hhmm"
' 반환: Collection(Array(mm-dd, startHHMM, endHHMM))  (단일 원소)
'========================
Public Function func_date(ByVal v As Variant) As Collection
    Dim s As String
    s = Trim$(GetValAsString(v))
    s = Replace(s, " ", "") ' 혹시 중간 공백 들어가도 흡수

    ' 예: 20260104 0830~1630  -> 202601040830~1630
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = "^(\d{8})(\d{4})~(\d{4})$"

    If Not re.Test(s) Then ErrStop "func_date 형식 오류(yyyymmdd hhmm~hhmm): " & GetValAsString(v)

    Set m = re.Execute(s)(0)

    Dim ymd As String, st As String, en As String
    ymd = m.SubMatches(0)    ' yyyymmdd
    st = m.SubMatches(1)     ' hhmm
    en = m.SubMatches(2)     ' hhmm

    Dim mmdd As String
    mmdd = Mid$(ymd, 5, 2) & "-" & Mid$(ymd, 7, 2)  ' mm-dd

    Dim stFmt As String, enFmt As String
    stFmt = Left$(st, 2) & ":" & Right$(st, 2)      ' hh:mm
    enFmt = Left$(en, 2) & ":" & Right$(en, 2)

    Dim col As New Collection
    col.Add Array(mmdd, stFmt, enFmt)
    Set func_date = col
End Function

'========================
' func_move(y_, x_)
' y_: Collection of Array(mm-dd, start, end)
' 주황 글씨로 hh:mm~hh:mm (기존 글씨색 유지)
'========================
Public Sub func_move(ByVal y_ As Collection, ByVal x_ As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("out")

    ' out_b: B1부터 오른쪽으로 빈 셀 전까지
    Dim headerLastCol As Long
    headerLastCol = GetHeaderLastCol(ws, 1, 2)
    If headerLastCol < 2 Then ErrStop "'out' 시트 B1부터 헤더가 없습니다."

    ' out_a: A2부터 아래로 빈 셀 전까지 (x_ 유일매칭)
    Dim nameRow As Long
    nameRow = FindUniqueRowInColA_UntilBlank(ws, Trim$(x_), 2)

    ' 이름을 못 찾은 경우: 기록만 하고 종료하지 않음
    If nameRow = 0 Then Exit Sub

    Dim i As Long
    For i = 1 To y_.Count
        Dim item As Variant
        item = y_(i)

        Dim mmdd As String, st As String, en As String
        mmdd = CStr(item(0))   ' mm-dd
        st = CStr(item(1))     ' hh:mm
        en = CStr(item(2))

        Dim m As Long, d As Long
        m = CLng(Left$(mmdd, 2))
        d = CLng(Right$(mmdd, 2))

        ' 1) 월/일 일치 헤더 셀 유일 찾기
        Dim headerCell As Range
        Set headerCell = FindUniqueHeaderByMonthDay(ws, 1, 2, headerLastCol, m, d)

        ' 4) 토/일이면 continue
        Dim wk As String
        wk = ExtractWeekdayChar(headerCell)
        If wk = "토" Or wk = "일" Then GoTo NextItem

        ' 4) 교차셀에 주황 글씨로 추가
        Dim target As Range
        Set target = ws.Cells(nameRow, headerCell.Column)

        Dim newText As String
        newText = st & "~" & en

        AppendColoredLine target, newText, vbMagenta

NextItem:
    Next i
End Sub

'========================
' main_방학중근무일지
'========================
Public Sub main_방학중근무일지()
    Dim wsSrc As Worksheet
    Set wsSrc = ThisWorkbook.Worksheets("방학중근무일지")

    Set gMissingNames = New Collection

    Dim r As Long
    r = 1 ' B1부터

    Do While Trim$(CStr(wsSrc.Cells(r, "B").Value)) <> ""
        Dim X As Range
        Set X = wsSrc.Cells(r, "B")

        ' y = (같은행 A열) + " " + (같은행 C열)
        Dim y As String
        y = Trim$(CStr(wsSrc.Cells(r, "A").Value)) & " " & Trim$(CStr(wsSrc.Cells(r, "C").Value))

        Dim y_ As Collection
        Set y_ = func_date(y)

        Dim x_ As String
        x_ = func_name(X.Value)

        func_move y_, x_

        r = r + 1
    Loop

    ShowMissingNames
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

Private Sub AddMissingName(ByVal nameVal As String)
    If gMissingNames Is Nothing Then Set gMissingNames = New Collection

    On Error Resume Next
    gMissingNames.Add nameVal, CStr(nameVal) ' 중복 제거
    On Error GoTo 0
End Sub

Private Sub ShowMissingNames()
    If gMissingNames Is Nothing Then Exit Sub
    If gMissingNames.Count = 0 Then Exit Sub

    Dim msg As String
    Dim i As Long

    msg = "out 시트 A열에서 찾지 못한 이름 목록:" & vbCrLf & vbCrLf
    For i = 1 To gMissingNames.Count
        msg = msg & "- " & CStr(gMissingNames(i)) & vbCrLf
    Next i

    MsgBox msg, vbExclamation
End Sub

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

' out 시트 A열: startRow부터 아래로, 빈 셀 전까지만 보고 keyVal 유일 행 찾기
' 찾지 못한 경우: 0 반환 + 이름 기록
Private Function FindUniqueRowInColA_UntilBlank(ByVal ws As Worksheet, ByVal keyVal As String, ByVal startRow As Long) As Long
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
        AddMissingName keyVal
        FindUniqueRowInColA_UntilBlank = 0
        Exit Function
    End If

    If cnt > 1 Then
        ErrStop "out_a(A" & startRow & "~)에서 '" & keyVal & "' 일치 개수=" & cnt & " (2개 이상)"
    End If

    Dim f As Range
    Set f = rng.Find(What:=keyVal, LookIn:=xlValues, LookAt:=xlWhole)
    If f Is Nothing Then
        AddMissingName keyVal
        FindUniqueRowInColA_UntilBlank = 0
        Exit Function
    End If

    FindUniqueRowInColA_UntilBlank = f.Row
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

' 기존 텍스트(여러 색 포함) 유지 + 새로 추가된 newText만 color로 지정
Private Sub AppendColoredLine(ByVal target As Range, ByVal newText As String, ByVal color As Long)
    Dim baseText As String
    baseText = CStr(target.Value2)

    Dim pos As Long
    Dim startNew As Long

    If Len(baseText) > 0 Then
        pos = Len(baseText) + 1
        target.Characters(pos, 0).Insert vbLf & newText
        startNew = pos + 1
    Else
        target.Value2 = newText
        startNew = 1
    End If

    target.WrapText = True
    target.Characters(startNew, Len(newText)).Font.color = color
End Sub
