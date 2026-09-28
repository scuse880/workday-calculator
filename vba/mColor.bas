Attribute VB_Name = "mColor"
Option Explicit

'====================================================
' CountFillColor
'  - rng: 색을 셀 범위
'  - colorOrSample: (옵션) 색상 Long 값 또는 "샘플 셀" (기본=노랑)
'  - includeConditional: True면 조건부서식으로 보이는 색(DisplayFormat)까지 포함
'
' 사용 예)
'  =CountFillColor(A2:AF2)                      ' 노랑(실제 채우기색) 개수
'  =CountFillColor(A2:AF2, RGB(255,255,0))      ' 노랑 지정
'  =CountFillColor(A2:AF2, $Z$1)                ' Z1 셀의 채우기색과 같은 셀 개수
'  =CountFillColor(A2:AF2, vbYellow, TRUE)      ' 조건부서식 포함하여 노랑 개수
'====================================================
Public Function CountFillColor( _
    ByVal rng As Range, _
    Optional ByVal colorOrSample As Variant = vbYellow, _
    Optional ByVal includeConditional As Boolean = False _
) As Long

    Application.Volatile  ' 계산(F9 등) 발생 시 다시 계산되게

    Dim targetColor As Long
    targetColor = ResolveColor(colorOrSample, includeConditional)

    Dim cnt As Long
    Dim c As Range
    For Each c In rng.Cells
        If GetFillColorValue(c, includeConditional) = targetColor Then
            cnt = cnt + 1
        End If
    Next c

    CountFillColor = cnt
End Function

'--- (편의) 노랑만 세기 ---
Public Function CountYellow(ByVal rng As Range) As Long
    CountYellow = CountFillColor(rng, vbYellow, False)
End Function

Public Function CountYellow_Display(ByVal rng As Range) As Long
    CountYellow_Display = CountFillColor(rng, vbYellow, True)
End Function

'====================================================
' 내부 헬퍼들
'====================================================
Private Function ResolveColor(ByVal colorOrSample As Variant, ByVal includeConditional As Boolean) As Long
    ' 샘플 셀이 들어오면 그 셀의 채우기색을 기준으로
    If TypeName(colorOrSample) = "Range" Then
        ResolveColor = GetFillColorValue(colorOrSample, includeConditional)
    Else
        ResolveColor = CLng(colorOrSample)
    End If
End Function

Private Function GetFillColorValue(ByVal cell As Range, ByVal includeConditional As Boolean) As Long
    ' 채우기색이 없으면 -1 반환
    On Error GoTo EH

    If includeConditional Then
        If cell.DisplayFormat.Interior.Pattern = xlNone Then
            GetFillColorValue = -1
        Else
            GetFillColorValue = cell.DisplayFormat.Interior.color
        End If
    Else
        If cell.Interior.Pattern = xlNone Then
            GetFillColorValue = -1
        Else
            GetFillColorValue = cell.Interior.color
        End If
    End If
    Exit Function

EH:
    ' DisplayFormat 접근이 막히는 환경 등
    GetFillColorValue = -1
End Function


