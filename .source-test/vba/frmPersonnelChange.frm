VERSION 5.00
Begin VB.UserForm frmPersonnelChange
   Caption = "인사변동자 등록"
   ClientHeight = 5700
   ClientWidth = 7800
   StartUpPosition = 1
End
Attribute VB_Name = "frmPersonnelChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' @RuntimeForm

Private mWorkMonth As Date
Private mEmployees As Collection
Private mStartDay As Date
Private mEndDay As Date
Private mSavedStartDay As Date
Private mSavedEndDay As Date
Private mLoading As Boolean
Private mStatus As MSForms.Label
Private WithEvents mEmployeeList As MSForms.ComboBox
Private WithEvents mSaveButton As MSForms.CommandButton
Private WithEvents mClearButton As MSForms.CommandButton
Private WithEvents mCloseButton As MSForms.CommandButton
Public DayButtonHandlers As Collection

Private Sub UserForm_Initialize()
    Me.Caption = "인사변동자 등록"
    Me.BackColor = RGB(245, 245, 242)
    Me.Font.Name = "맑은 고딕"
    Me.Font.Size = 9
    Set DayButtonHandlers = New Collection
End Sub

Public Sub Configure(ByVal workMonth As Date, ByVal employees As Collection)
    Dim title As MSForms.Label, employee As CEmployee
    Dim dayButton As MSForms.CommandButton, handler As CPersonnelDayButton
    Dim d As Long, firstSlot As Long, slot As Long, rowIndex As Long, columnIndex As Long
    Dim dayValue As Date

    mLoading = True
    mWorkMonth = workMonth
    Set mEmployees = employees
    Me.Width = 520
    Me.Height = 405

    Set title = AddLabel("lblTitle", "인사변동자 등록  |  " & Format$(workMonth, "yyyy년 m월"), 17, 10, 480, 24)
    title.Font.Size = 13
    title.Font.Bold = True
    AddLabel "lblEmployee", "교직원", 17, 41, 72, 18
    Set mEmployeeList = Me.Controls.Add("Forms.ComboBox.1", "cmbEmployee", True)
    With mEmployeeList
        .Left = 96
        .Top = 39
        .Width = 250
        .Height = 22
        .Style = fmStyleDropDownList
        .Font.Name = "맑은 고딕"
        .Font.Size = 9
    End With
    For Each employee In mEmployees
        If gEmployeesByName(employee.Name).Count > 1 Then
            mEmployeeList.AddItem employee.Name & "(" & employee.Birthdate & ")"
        Else
            mEmployeeList.AddItem employee.Name
        End If
    Next employee

    Set title = AddLabel("lblNotice", "등록한 기간의 날짜는 해당 교직원의 근무일수 계산에서 제외됩니다.", 17, 69, 480, 19)
    title.Font.Bold = True
    title.ForeColor = RGB(115, 45, 35)
    AddLabel "lblInstructions", "시작일과 종료일을 차례로 클릭하세요. 종료일은 시작일보다 늦어야 합니다. 교직원 변경·닫기 시 저장하지 않은 선택은 취소됩니다.", 17, 90, 480, 32
    AddLabel "lblMonth", Format$(workMonth, "yyyy년 m월"), 19, 130, 264, 19
    For columnIndex = 0 To 6
        Set title = AddLabel("lblWeek" & columnIndex, Mid$("일월화수목금토", columnIndex + 1, 1), _
                             19 + columnIndex * 38, 151, 36, 17)
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
            .Left = 19 + columnIndex * 38
            .Top = 173 + rowIndex * 29
            .Width = 36
            .Height = 27
            .Font.Name = "맑은 고딕"
            .Font.Size = 9
            .ControlTipText = Format$(dayValue, "yyyy-mm-dd") & " 선택"
        End With
        Set handler = New CPersonnelDayButton
        Set handler.Button = dayButton
        handler.DayValue = dayValue
        Set handler.ParentForm = Me
        DayButtonHandlers.Add handler
    Next d

    AddLabel "lblSelection", "선택한 기간", 307, 130, 185, 20
    Set mStatus = AddLabel("lblStatus", vbNullString, 307, 154, 185, 69)
    mStatus.BackStyle = fmBackStyleOpaque
    mStatus.BackColor = RGB(255, 255, 255)
    mStatus.BorderStyle = fmBorderStyleSingle
    mStatus.WordWrap = True
    Set mSaveButton = AddActionButton("cmdSave", "저장", 307, 232)
    Set mClearButton = AddActionButton("cmdClear", "등록 해제", 307, 269)
    Set mCloseButton = AddActionButton("cmdClose", "닫기", 307, 306)

    If mEmployeeList.ListCount > 0 Then mEmployeeList.ListIndex = 0
    mLoading = False
    LoadEmployeeSelection
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

Private Function AddActionButton(ByVal controlName As String, ByVal captionText As String, _
                                 ByVal leftValue As Single, ByVal topValue As Single) As MSForms.CommandButton
    Dim button As MSForms.CommandButton
    Set button = Me.Controls.Add("Forms.CommandButton.1", controlName, True)
    With button
        .Caption = captionText
        .Left = leftValue
        .Top = topValue
        .Width = 185
        .Height = 29
        .Font.Name = "맑은 고딕"
        .Font.Size = 9
        .BackColor = RGB(223, 229, 233)
    End With
    Set AddActionButton = button
End Function

Private Sub mEmployeeList_Change()
    If Not mLoading Then LoadEmployeeSelection
End Sub

Private Sub LoadEmployeeSelection()
    Dim employee As CEmployee
    mStartDay = 0
    mEndDay = 0
    mSavedStartDay = 0
    mSavedEndDay = 0
    If Not mEmployeeList Is Nothing Then
        If mEmployeeList.ListIndex >= 0 Then
            Set employee = mEmployees(mEmployeeList.ListIndex + 1)
            mStartDay = employee.PersonnelChangeStart
            mEndDay = employee.PersonnelChangeEnd
            mSavedStartDay = mStartDay
            mSavedEndDay = mEndDay
        End If
    End If
    RefreshSelection
End Sub

Public Sub SelectDay(ByVal dayValue As Date)
    If mEmployeeList.ListIndex < 0 Then Exit Sub
    If Year(dayValue) <> Year(mWorkMonth) Or Month(dayValue) <> Month(mWorkMonth) Then Exit Sub
    If mStartDay = 0 Or mEndDay <> 0 Then
        mStartDay = dayValue
        mEndDay = 0
    ElseIf dayValue <= mStartDay Then
        MsgBox "종료일은 시작일보다 늦어야 합니다.", vbExclamation, "인사변동 기간"
        Exit Sub
    Else
        mEndDay = dayValue
    End If
    RefreshSelection
End Sub

Private Sub RefreshSelection()
    Dim handler As CPersonnelDayButton, dayValue As Date
    If mStatus Is Nothing Then Exit Sub
    If mEmployeeList.ListIndex < 0 Then
        mStatus.Caption = "교직원을 선택하세요."
    ElseIf mStartDay = 0 Then
        mStatus.Caption = "등록된 기간 없음" & vbCrLf & "날짜를 차례로 두 번 클릭하세요."
    ElseIf mEndDay = 0 Then
        mStatus.Caption = "시작: " & Format$(mStartDay, "yyyy-mm-dd") & vbCrLf & "종료 날짜를 선택하세요."
    Else
        mStatus.Caption = Format$(mStartDay, "yyyy-mm-dd") & " ~ " & Format$(mEndDay, "yyyy-mm-dd") & vbCrLf & _
                          CStr(DateDiff("d", mStartDay, mEndDay) + 1) & "일" & _
                          IIf(mStartDay = mSavedStartDay And mEndDay = mSavedEndDay, " · 저장됨", " · 저장 전")
    End If
    For Each handler In DayButtonHandlers
        dayValue = handler.DayValue
        With handler.Button
            .BackColor = RGB(239, 239, 235)
            .ForeColor = RGB(0, 0, 0)
            If Weekday(dayValue, vbSunday) = 1 Then .ForeColor = RGB(160, 40, 40)
            If Weekday(dayValue, vbSunday) = 7 Then .ForeColor = RGB(45, 70, 130)
            If mStartDay <> 0 Then
                If dayValue = mStartDay Or (mEndDay <> 0 And dayValue = mEndDay) Then
                    .BackColor = RGB(65, 85, 105)
                    .ForeColor = RGB(255, 255, 255)
                ElseIf mEndDay <> 0 And dayValue > mStartDay And dayValue < mEndDay Then
                    .BackColor = RGB(204, 217, 226)
                    .ForeColor = RGB(0, 0, 0)
                End If
            End If
        End With
    Next handler
End Sub

Private Sub mSaveButton_Click()
    Dim employee As CEmployee
    On Error GoTo Failed
    If mEmployeeList.ListIndex < 0 Then RaiseValidation "인사변동 기간", "교직원을 선택하세요."
    Set employee = mEmployees(mEmployeeList.ListIndex + 1)
    SavePersonnelChange mWorkMonth, employee.Order, mStartDay, mEndDay
    mSavedStartDay = mStartDay
    mSavedEndDay = mEndDay
    RefreshSelection
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "인사변동자 등록"
End Sub

Private Sub mClearButton_Click()
    Dim employee As CEmployee
    On Error GoTo Failed
    If mEmployeeList.ListIndex < 0 Then RaiseValidation "인사변동 기간", "교직원을 선택하세요."
    Set employee = mEmployees(mEmployeeList.ListIndex + 1)
    ClearPersonnelChange mWorkMonth, employee.Order
    mStartDay = 0
    mEndDay = 0
    mSavedStartDay = 0
    mSavedEndDay = 0
    RefreshSelection
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "인사변동자 등록"
End Sub

Private Sub mCloseButton_Click()
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub

Public Sub DisposeHandlers()
    Dim handler As CPersonnelDayButton
    If Not DayButtonHandlers Is Nothing Then
        For Each handler In DayButtonHandlers
            handler.Detach
        Next handler
    End If
    Set DayButtonHandlers = Nothing
    Set mEmployeeList = Nothing
    Set mSaveButton = Nothing
    Set mClearButton = Nothing
    Set mCloseButton = Nothing
End Sub

Private Sub UserForm_Terminate()
    DisposeHandlers
    Set mStatus = Nothing
    Set mEmployees = Nothing
End Sub
