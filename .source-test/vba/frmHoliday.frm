VERSION 5.00
Begin VB.UserForm frmHoliday
   Caption = "공휴일 설정"
   ClientHeight = 3900
   ClientWidth = 6900
   StartUpPosition = 1
End
Attribute VB_Name = "frmHoliday"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' @RuntimeForm

Private mWorkMonth As Date
Private mSelected As Object
Public DayButtonHandlers As Collection
Private WithEvents mSaveButton As MSForms.CommandButton
Private mSelectedList As MSForms.ListBox

Private Sub UserForm_Initialize()
    Me.Caption = "공휴일 설정"
    Me.BackColor = RGB(245, 245, 242)
    Me.Font.Name = "맑은 고딕"
    Me.Font.Size = 9
    Set DayButtonHandlers = New Collection
End Sub

Public Sub Configure(ByVal workMonth As Date, ByVal holidays As Object)
    Dim zoneWidth As Single, margin As Single, gap As Single, contentTop As Single
    Dim rightLeft As Single, cellWidth As Single, cellHeight As Single
    Dim d As Long, slot As Long, firstSlot As Long, columnIndex As Long
    Dim rowIndex As Long, dayValue As Date, dayButton As MSForms.CommandButton
    Dim handler As CHolidayDayButton, key As Variant, title As MSForms.Label
    mWorkMonth = workMonth
    Set mSelected = NewDictionary()
    For Each key In holidays.Keys
        mSelected.Add CStr(key), holidays(key)
    Next key
    margin = 12
    gap = 12
    zoneWidth = Application.Width * 0.125
    If zoneWidth < 210 Then zoneWidth = 210
    contentTop = 42
    rightLeft = margin + zoneWidth + gap
    Me.Width = margin * 2 + zoneWidth * 2 + gap + 6
    Me.Height = contentTop + zoneWidth + margin + 26
    Set title = AddLabel("lblTitle", "공휴일 설정  |  " & Format$(workMonth, "yyyy년 m월"), margin, 10, zoneWidth * 2 + gap, 23)
    title.Font.Size = 13
    title.Font.Bold = True
    AddLabel "lblMonth", Format$(workMonth, "yyyy년 m월"), margin + 5, contentTop + 5, zoneWidth - 10, 20
    cellWidth = (zoneWidth - 10) / 7
    cellHeight = (zoneWidth - 50) / 6
    For columnIndex = 0 To 6
        Set title = AddLabel("lblWeek" & columnIndex, Mid$("일월화수목금토", columnIndex + 1, 1), _
                             margin + 5 + columnIndex * cellWidth, contentTop + 27, cellWidth, 16)
        title.TextAlign = fmTextAlignCenter
        If columnIndex = 0 Then title.ForeColor = RGB(160, 40, 40)
        If columnIndex = 6 Then title.ForeColor = RGB(45, 70, 130)
    Next columnIndex
    firstSlot = Weekday(workMonth, vbSunday) - 1
    For d = 1 To Day(DateSerial(Year(workMonth), Month(workMonth) + 1, 0))
        dayValue = DateSerial(Year(workMonth), Month(workMonth), d)
        slot = firstSlot + d - 1
        rowIndex = slot \ 7
        columnIndex = slot Mod 7
        Set dayButton = Me.Controls.Add("Forms.CommandButton.1", "day" & CStr(d), True)
        With dayButton
            .Caption = CStr(d)
            .Left = margin + 5 + columnIndex * cellWidth
            .Top = contentTop + 46 + rowIndex * cellHeight
            .Width = cellWidth - 2
            .Height = cellHeight - 2
            .Font.Name = "맑은 고딕"
            .Font.Size = 9
            .ControlTipText = Format$(dayValue, "yyyy-mm-dd") & " 공휴일 선택/해제"
        End With
        Set handler = New CHolidayDayButton
        Set handler.Button = dayButton
        handler.HolidayDate = dayValue
        Set handler.ParentForm = Me
        DayButtonHandlers.Add handler
    Next d
    AddLabel "lblSelected", "선택한 공휴일", rightLeft + 5, contentTop + 5, zoneWidth - 10, 18
    Set mSelectedList = Me.Controls.Add("Forms.ListBox.1", "lstSelected", True)
    With mSelectedList
        .Left = rightLeft + 5
        .Top = contentTop + 28
        .Width = zoneWidth - 10
        .Height = zoneWidth * 0.9 - 33
        .Font.Name = "맑은 고딕"
        .Font.Size = 9
        .TabStop = False
        .BackColor = RGB(255, 255, 255)
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
    End With
    Set mSaveButton = Me.Controls.Add("Forms.CommandButton.1", "cmdSave", True)
    With mSaveButton
        .Caption = "저장 및 닫기"
        .Left = rightLeft + 5
        .Top = contentTop + zoneWidth * 0.9 + 1
        .Width = zoneWidth - 10
        .Height = zoneWidth * 0.1 - 2
        .Default = True
        .Font.Name = "맑은 고딕"
        .Font.Size = 9
        .BackColor = RGB(223, 229, 233)
    End With
    RefreshSelection
End Sub

Private Function AddLabel(ByVal controlName As String, ByVal captionText As String, _
                          ByVal leftValue As Single, ByVal topValue As Single, _
                          ByVal widthValue As Single, ByVal heightValue As Single) As MSForms.Label
    Dim label As MSForms.Label
    Set label = Me.Controls.Add("Forms.Label.1", controlName, True)
    With label
        .Caption = captionText
        .Left = leftValue
        .Top = topValue
        .Width = widthValue
        .Height = heightValue
        .BackStyle = fmBackStyleTransparent
        .Font.Name = "맑은 고딕"
        .Font.Size = 9
    End With
    Set AddLabel = label
End Function

Public Sub ToggleHoliday(ByVal dayValue As Date)
    Dim key As String
    If Year(dayValue) <> Year(mWorkMonth) Or Month(dayValue) <> Month(mWorkMonth) Then Exit Sub
    key = HolidayKey(dayValue)
    If mSelected.Exists(key) Then
        mSelected.Remove key
    Else
        mSelected.Add key, dayValue
    End If
    RefreshSelection
End Sub

Private Sub RefreshSelection()
    Dim d As Long, dayValue As Date, n As Long, handler As CHolidayDayButton
    mSelectedList.Clear
    For d = 1 To Day(DateSerial(Year(mWorkMonth), Month(mWorkMonth) + 1, 0))
        dayValue = DateSerial(Year(mWorkMonth), Month(mWorkMonth), d)
        If mSelected.Exists(HolidayKey(dayValue)) Then
            n = n + 1
            mSelectedList.AddItem CStr(n) & ". " & Format$(dayValue, "yyyy-mm-dd") & "(" & KoreanWeekday(dayValue) & ")"
        End If
    Next d
    For Each handler In DayButtonHandlers
        If mSelected.Exists(HolidayKey(handler.HolidayDate)) Then
            handler.Button.BackColor = RGB(65, 85, 105)
            handler.Button.ForeColor = RGB(255, 255, 255)
        Else
            handler.Button.BackColor = RGB(239, 239, 235)
            handler.Button.ForeColor = RGB(0, 0, 0)
            If Weekday(handler.HolidayDate, vbSunday) = 1 Then handler.Button.ForeColor = RGB(160, 40, 40)
            If Weekday(handler.HolidayDate, vbSunday) = 7 Then handler.Button.ForeColor = RGB(45, 70, 130)
        End If
    Next handler
End Sub

Private Sub mSaveButton_Click()
    On Error GoTo Failed
    SaveHolidays mWorkMonth, mSelected
    Me.Hide
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "공휴일 저장"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' X는 저장하지 않고 숨긴다. 호출자가 이벤트 종료 후 핸들러를 해제한다.
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub

Public Sub DisposeHandlers()
    Dim handler As CHolidayDayButton
    If Not DayButtonHandlers Is Nothing Then
        For Each handler In DayButtonHandlers
            handler.Detach
        Next handler
    End If
    Set DayButtonHandlers = Nothing
    Set mSaveButton = Nothing
End Sub

Private Sub UserForm_Terminate()
    DisposeHandlers
    Set mSelectedList = Nothing
    Set mSelected = Nothing
End Sub
