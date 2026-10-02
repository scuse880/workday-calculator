Attribute VB_Name = "main_시트보호"
Option Explicit

' 사용자 편집만 제한한다. VBA는 UserInterfaceOnly 보호를 통해 기존 셀에 쓴다.
Private Const SHEET_PASSWORD As String = "workday-ui"
Private mApplying As Boolean
Private mVacationBorderLastRow As Long
Private mVacationPreviousEdges As Object

Public Sub InitializeWorkbookProtection()
    mVacationBorderLastRow = 0
    Set mVacationPreviousEdges = Nothing
    ApplyWorkbookProtection True
End Sub

Public Sub ApplyWorkbookProtection(Optional ByVal resetWorkDefaults As Boolean = False)
    Dim ws As Worksheet, previousEvents As Boolean
    Dim previousWindow As Window, previousSheet As Object, previousSelection As Range
    Dim previousScrollRow As Long, previousScrollColumn As Long, previousZoom As Variant
    Dim errorNumber As Long, errorText As String
    If mApplying Then Exit Sub
    mApplying = True
    previousEvents = Application.EnableEvents
    On Error GoTo Failed
    Set previousWindow = Application.ActiveWindow
    Set previousSheet = Application.ActiveSheet
    If TypeOf Application.Selection Is Range Then Set previousSelection = Application.Selection
    If Not previousWindow Is Nothing Then
        previousScrollRow = previousWindow.ScrollRow
        previousScrollColumn = previousWindow.ScrollColumn
        previousZoom = previousWindow.Zoom
    End If
    Application.EnableEvents = False
    ThisWorkbook.Unprotect Password:=SHEET_PASSWORD
    Set ws = ThisWorkbook.Worksheets("작업")
    ws.Unprotect Password:=SHEET_PASSWORD
    ws.Cells.Locked = True
    PrepareWorkSelectorSources ws
    If resetWorkDefaults Then ws.Range("B3,B5,B6,D3,D5,D6").Value2 = "선택"
    ' Forms 드롭다운의 최소 높이에 입력 행 전체를 맞춘다.
    ws.Rows(3).RowHeight = 18.75
    ws.Rows(5).RowHeight = 18.75
    ws.Rows(6).RowHeight = 18.75
    ConfigureWorkSelectors ws
    ProtectUserSheet ws

    Set ws = ThisWorkbook.Worksheets("방학중근무")
    ws.Unprotect Password:=SHEET_PASSWORD
    ws.Cells.Locked = True
    ws.Range("A3:D" & ws.Rows.Count).Locked = False
    ws.Range("B2").Value2 = "날짜" & vbLf & "(yyyy-mm-dd)"
    ws.Range("C2").Value2 = "시작 시각" & vbLf & "(hh:mm)"
    ws.Range("D2").Value2 = "종료 시각" & vbLf & "(hh:mm)"
    ws.Range("B2:D2").WrapText = True
    ws.Rows(2).AutoFit
    RefreshVacationBorders ws, ws.UsedRange
    ProtectUserSheet ws

    Set ws = ThisWorkbook.Worksheets("사용방법 및 주의사항")
    ws.Unprotect Password:=SHEET_PASSWORD
    ws.Cells.Locked = True
    ProtectUserSheet ws
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "설정" Or ws.Name = "공휴일" Then
            ws.Unprotect Password:=SHEET_PASSWORD
            ws.Cells.Locked = True
            ProtectUserSheet ws
            If ws.Visible <> xlSheetVeryHidden Then ws.Visible = xlSheetVeryHidden
        End If
    Next ws
    ThisWorkbook.Protect Password:=SHEET_PASSWORD, Structure:=True, Windows:=False
Done:
    ' 보이는 Excel에서 외부 파일을 닫아 재활성화될 때 보호 재적용이 다른 시트의
    ' 화면 상태를 선택할 수 있다. 이벤트를 다시 켜기 전에 원래 화면까지 복원한다.
    On Error Resume Next
    If Not previousWindow Is Nothing Then previousWindow.Activate
    If Not previousSheet Is Nothing Then
        If previousSheet.Visible = xlSheetVisible Then
            previousSheet.Activate
            If Not previousSelection Is Nothing Then previousSelection.Select
        End If
    End If
    If Not previousWindow Is Nothing Then
        previousWindow.Zoom = previousZoom
        previousWindow.ScrollRow = previousScrollRow
        previousWindow.ScrollColumn = previousScrollColumn
    End If
    If Err.Number <> 0 And errorNumber = 0 Then
        errorNumber = Err.Number
        errorText = Err.Description
    End If
    Application.EnableEvents = previousEvents
    mApplying = False
    On Error GoTo 0
    If errorNumber <> 0 Then Err.Raise errorNumber, "ApplyWorkbookProtection", errorText
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    ' 초기화 도중 오류가 나도 이미 해제한 사용자 보호를 복구한다.
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "작업" Or ws.Name = "방학중근무" Or ws.Name = "사용방법 및 주의사항" _
           Or ws.Name = "설정" Or ws.Name = "공휴일" Then ProtectUserSheet ws
    Next ws
    ThisWorkbook.Protect Password:=SHEET_PASSWORD, Structure:=True, Windows:=False
    On Error GoTo 0
    Resume Done
End Sub

Private Sub ProtectUserSheet(ByVal ws As Worksheet)
    ws.Protect Password:=SHEET_PASSWORD, DrawingObjects:=True, Contents:=True, _
               Scenarios:=True, UserInterfaceOnly:=True
End Sub

' 통합문서 구조 보호 중에도 기존의 누락 시트 생성 경로를 유지한다.
Public Function CreateProtectedWorkbookSheet(ByVal sheetName As String) As Worksheet
    Dim wasProtected As Boolean, ws As Worksheet
    Dim errorNumber As Long, errorText As String
    wasProtected = ThisWorkbook.ProtectStructure
    On Error GoTo Failed
    If wasProtected Then ThisWorkbook.Unprotect Password:=SHEET_PASSWORD
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = sheetName
    If sheetName = "공휴일" Then
        ProtectUserSheet ws
        ws.Visible = xlSheetVeryHidden
    End If
    Set CreateProtectedWorkbookSheet = ws
Done:
    If wasProtected Then ThisWorkbook.Protect Password:=SHEET_PASSWORD, Structure:=True, Windows:=False
    If errorNumber <> 0 Then Err.Raise errorNumber, "CreateProtectedWorkbookSheet", errorText
    Exit Function
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    Resume Done
End Function

Private Sub PrepareWorkSelectorSources(ByVal ws As Worksheet)
    Dim settings As Worksheet, columnNumber As Long, address As Variant
    Dim source As Range, formula As String
    Set settings = ThisWorkbook.Worksheets("설정")
    settings.Unprotect Password:=SHEET_PASSWORD
    For columnNumber = 1 To 4
        If CStr(settings.Cells(2, columnNumber).Value2) <> "선택" Then
            settings.Cells(2, columnNumber).Insert Shift:=xlShiftDown
            settings.Cells(2, columnNumber).Value2 = "선택"
        End If
    Next columnNumber
    For Each address In Array("B3", "B5", "B6", "D3", "D5", "D6")
        formula = ws.Range(CStr(address)).Validation.Formula1
        Set source = settings.Range(Split(formula, "!")(1))
        Set source = settings.Range(settings.Cells(2, source.Column), source.Cells(source.Rows.Count, 1))
        ws.Range(CStr(address)).Validation.Modify Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                                                   Operator:=xlBetween, Formula1:="=설정!" & source.Address
    Next address
End Sub

Private Sub ConfigureWorkSelectors(ByVal ws As Worksheet)
    Dim address As Variant, cell As Range, selector As Shape, source As Range
    Dim formula As String, entry As Range, selectedIndex As Long, itemIndex As Long
    For Each address In Array("B3", "B5", "B6", "D3", "D5", "D6")
        Set cell = ws.Range(CStr(address))
        formula = cell.Validation.Formula1
        Set source = ThisWorkbook.Worksheets("설정").Range(Split(formula, "!")(1))
        Set selector = Nothing
        On Error Resume Next
        Set selector = ws.Shapes("select_" & CStr(address))
        On Error GoTo 0
        If selector Is Nothing Then
            Set selector = ws.Shapes.AddFormControl(xlDropDown, cell.Left, cell.Top, cell.Width, cell.Height)
            selector.Name = "select_" & CStr(address)
        End If
        With selector
            .Left = cell.Left
            .Top = cell.Top
            .Width = cell.Width
            .Height = cell.Height
            .Placement = xlMoveAndSize
            .Locked = True
            .OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!WorkSettingSelected"
        End With
        With selector.ControlFormat
            .LinkedCell = vbNullString
            .ListFillRange = vbNullString
            .RemoveAllItems
            selectedIndex = 0
            itemIndex = 0
            For Each entry In source.Cells
                If Len(Trim$(CStr(entry.Value2))) > 0 Then
                    .AddItem CStr(entry.Value2)
                    itemIndex = itemIndex + 1
                    If CStr(entry.Value2) = CStr(cell.Value2) Then selectedIndex = itemIndex
                End If
            Next entry
            .DropDownLines = 25
            .ListIndex = selectedIndex
        End With
    Next address
End Sub

Public Sub WorkSettingSelected()
    Dim ws As Worksheet, selector As Shape, address As String, selection As Long
    Set ws = ThisWorkbook.Worksheets("작업")
    Set selector = ws.Shapes(CStr(Application.Caller))
    address = Mid$(selector.Name, Len("select_") + 1)
    Select Case address
        Case "B3", "B5", "B6", "D3", "D5", "D6"
            selection = selector.ControlFormat.ListIndex
            If selection > 0 Then ws.Range(address).Value2 = selector.ControlFormat.List(selection)
    End Select
End Sub

Public Sub RefreshVacationBorders(ByVal ws As Worksheet, ByVal changed As Range)
    Dim affected As Range, area As Range, rowRange As Range, rowNumber As Long
    Dim seen As Object, clearedRows As Object, usedBottom As Long, lastCell As Range, key As Variant
    If mVacationBorderLastRow = 0 Then mVacationBorderLastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    Set lastCell = ws.Range("A:D").Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
                                    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    usedBottom = mVacationBorderLastRow
    If Not lastCell Is Nothing Then
        If lastCell.Row > usedBottom Then usedBottom = lastCell.Row
    End If
    If usedBottom < 13 Then Exit Sub
    Set affected = Application.Intersect(changed, ws.Range("A13:D" & usedBottom))
    If affected Is Nothing Then Exit Sub
    Set seen = CreateObject("Scripting.Dictionary")
    Set clearedRows = CreateObject("Scripting.Dictionary")
    For Each area In affected.Areas
        For Each rowRange In area.Rows
            rowNumber = rowRange.Row
            If Not seen.Exists(CStr(rowNumber)) Then
                seen.Add CStr(rowNumber), True
                If Application.CountA(ws.Range(ws.Cells(rowNumber, 1), ws.Cells(rowNumber, 4))) = 0 Then
                    clearedRows.Add CStr(rowNumber), True
                End If
            End If
        Next rowRange
    Next area
    For Each key In seen.Keys
        UpdateVacationRowBorders ws, CLng(key), clearedRows
    Next key
    mVacationBorderLastRow = usedBottom
End Sub

Private Sub UpdateVacationRowBorders(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal clearedRows As Object)
    Dim rowRange As Range, savedBorders As Collection, saved As Variant
    Dim edge As Border, neighbor As Range, previousEdges As Collection, original As Variant
    Dim hasName As Boolean
    Set rowRange = ws.Range(ws.Cells(rowNumber, 1), ws.Cells(rowNumber, 4))
    If mVacationPreviousEdges Is Nothing Then Set mVacationPreviousEdges = CreateObject("Scripting.Dictionary")
    hasName = IsError(ws.Cells(rowNumber, 1).Value2)
    If Not hasName Then hasName = (Len(Trim$(CStr(ws.Cells(rowNumber, 1).Value2))) > 0)
    If hasName Then
        If Not mVacationPreviousEdges.Exists(CStr(rowNumber)) Then
            Set previousEdges = CaptureVacationNeighbors(ws, rowNumber)
            mVacationPreviousEdges.Add CStr(rowNumber), previousEdges
        End If
        With rowRange.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    ElseIf Application.CountA(rowRange) = 0 Then
        Set savedBorders = New Collection
        If mVacationPreviousEdges.Exists(CStr(rowNumber)) Then
            Set previousEdges = mVacationPreviousEdges(CStr(rowNumber))
        Else
            Set previousEdges = CaptureVacationNeighbors(ws, rowNumber)
        End If
        For Each original In previousEdges
            Set neighbor = ws.Range(CStr(original(0)))
            If Not clearedRows.Exists(CStr(neighbor.Row)) Then
                If neighbor.Row = 12 Or Application.CountA(ws.Range(ws.Cells(neighbor.Row, 1), ws.Cells(neighbor.Row, 4))) > 0 _
                   Or VacationRowHasOwnBorder(ws, neighbor.Row, CLng(original(1))) Then
                    Set edge = neighbor.Borders(CLng(original(1)))
                    savedBorders.Add Array(neighbor.Address, original(1), edge.LineStyle, edge.Weight, edge.Color)
                ElseIf original(5) Then
                    ' 성명 입력 전 빈 이웃 행에 이미 있던 경계만 복원한다.
                    savedBorders.Add original
                End If
            End If
        Next original
        rowRange.Borders.LineStyle = xlLineStyleNone
        For Each saved In savedBorders
            Set edge = ws.Range(CStr(saved(0))).Borders(CLng(saved(1)))
            edge.LineStyle = saved(2)
            If saved(2) <> xlLineStyleNone Then
                edge.Weight = saved(3)
                edge.Color = saved(4)
            End If
        Next saved
        If mVacationPreviousEdges.Exists(CStr(rowNumber)) Then mVacationPreviousEdges.Remove CStr(rowNumber)
    End If
End Sub

Private Function CaptureVacationNeighbors(ByVal ws As Worksheet, ByVal rowNumber As Long) As Collection
    Dim result As New Collection, c As Long, neighborRow As Variant, edgeIndex As Long
    Dim neighbor As Range, edge As Border, emptyNeighbor As Boolean, lineStyle As Variant
    For Each neighborRow In Array(rowNumber - 1, rowNumber + 1)
        If CLng(neighborRow) <= ws.Rows.Count Then
            edgeIndex = xlEdgeTop
            If CLng(neighborRow) < rowNumber Then edgeIndex = xlEdgeBottom
            emptyNeighbor = (Application.CountA(ws.Range(ws.Cells(CLng(neighborRow), 1), ws.Cells(CLng(neighborRow), 4))) = 0)
            For c = 1 To 4
                Set neighbor = ws.Cells(CLng(neighborRow), c)
                Set edge = neighbor.Borders(edgeIndex)
                lineStyle = edge.LineStyle
                If mApplying And emptyNeighbor And CLng(neighborRow) >= 13 Then
                    If Not VacationRowHasOwnBorder(ws, CLng(neighborRow), edgeIndex) Then lineStyle = xlLineStyleNone
                End If
                result.Add Array(neighbor.Address, edgeIndex, lineStyle, edge.Weight, edge.Color, emptyNeighbor)
            Next c
        End If
    Next neighborRow
    Set CaptureVacationNeighbors = result
End Function

Private Function VacationRowHasOwnBorder(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal sharedEdge As Long) As Boolean
    Dim c As Long, edge As Border
    For c = 1 To 4
        ' 위·아래 공유 선과 E열에서 보이는 바깥 선을 이 행의 입력 테두리로 오인하지 않는다.
        Set edge = ws.Cells(rowNumber, c).Borders(xlEdgeLeft)
        If edge.LineStyle <> xlLineStyleNone Then
            VacationRowHasOwnBorder = True
            Exit Function
        End If
    Next c
End Function
