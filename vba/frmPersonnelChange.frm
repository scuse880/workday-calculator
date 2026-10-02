VERSION 5.00
Begin VB.UserForm frmPersonnelChange
   Caption = "인사변동자 등록"
   ClientHeight = 5700
   ClientWidth = 11400
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
Private mSelectedEmployeeIndex As Long
Private mDisplayedEmployees As Collection
Private mRegisteredEmployees As Collection
Private mStatus As MSForms.label
Private WithEvents mEmployeeList As MSForms.ComboBox
Private WithEvents mRegisteredList As MSForms.ListBox
Private WithEvents mSaveButton As MSForms.CommandButton
Private WithEvents mClearButton As MSForms.CommandButton
Private WithEvents mCloseButton As MSForms.CommandButton
Public DayButtonHandlers As Collection

Private Sub UserForm_Initialize()
    Me.Caption = "인사변동자 등록"
    Me.BackColor = RGB(245, 245, 242)
    Me.Font.name = "맑은 고딕"
    Me.Font.Size = 9
    Set DayButtonHandlers = New Collection
End Sub

Public Sub Configure(ByVal workMonth As Date, ByVal employees As Collection)
    Dim title As MSForms.label
    Dim dayButton As MSForms.CommandButton, handler As CPersonnelDayButton
    Dim d As Long, firstSlot As Long, slot As Long, rowIndex As Long, columnIndex As Long
    Dim dayValue As Date

    mLoading = True
    mWorkMonth = workMonth
    Set mEmployees = employees
    Me.Width = 760
    Me.Height = 405

    Set title = AddLabel("lblTitle", Format$(workMonth, "yyyy년 m월"), 17, 10, 720, 24)
    title.Font.Size = 13
    title.Font.Bold = True
    AddLabel "lblEmployee", "교직원", 17, 41, 72, 18
    Set mEmployeeList = Me.Controls.Add("Forms.ComboBox.1", "cmbEmployee", True)
    With mEmployeeList
        .Left = 96
        .Top = 39
        .Width = 420
        .Height = 22
        .Style = fmStyleDropDownCombo
        .MatchEntry = fmMatchEntryNone
        .ListRows = 25
        .Font.name = "맑은 고딕"
        .Font.Size = 9
    End With
    PopulateEmployeeList vbNullString

    Set title = AddLabel("lblNotice", "등록한 기간의 날짜는 해당 교직원의 근무일수 계산에서 제외됩니다.", 17, 69, 720, 19)
    title.Font.Bold = True
    title.ForeColor = RGB(115, 45, 35)
    AddLabel "lblInstructions", "시작일과 종료일을 차례로 클릭하세요. 종료일은 시작일보다 늦어야 합니다. 교직원 변경·닫기 시 저장하지 않은 선택은 취소됩니다.", 17, 90, 720, 32
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
            .Font.name = "맑은 고딕"
            .Font.Size = 9
            .ControlTipText = Format$(dayValue, "yyyy-mm-dd") & " 선택"
        End With
        Set handler = New CPersonnelDayButton
        Set handler.button = dayButton
        handler.dayValue = dayValue
        Set handler.ParentForm = Me
        DayButtonHandlers.Add handler
    Next d

    AddLabel "lblRegistered", "등록된 교직원 및 기간", 307, 130, 420, 20
    Set mRegisteredList = Me.Controls.Add("Forms.ListBox.1", "lstRegistered", True)
    With mRegisteredList
        .Left = 307
        .Top = 154
        .Width = 420
        .Height = 125
        .ColumnCount = 2
        .ColumnWidths = "165 pt;235 pt"
        .Font.name = "맑은 고딕"
        .Font.Size = 9
        .IntegralHeight = False
    End With
    RefreshRegisteredEmployees
    AddLabel "lblSelection", "선택한 기간", 307, 287, 420, 18
    Set mStatus = AddLabel("lblStatus", vbNullString, 307, 307, 420, 35)
    mStatus.BackStyle = fmBackStyleOpaque
    mStatus.BackColor = RGB(255, 255, 255)
    mStatus.BorderStyle = fmBorderStyleSingle
    mStatus.WordWrap = True
    Set mSaveButton = AddActionButton("cmdSave", "저장", 445, 350)
    Set mClearButton = AddActionButton("cmdClear", "등록 해제", 541, 350)
    Set mCloseButton = AddActionButton("cmdClose", "닫기", 637, 350)
    ' Height에는 제목 표시줄과 테두리가 포함되므로 실제 내부 높이로 하단 여백을 확보한다.
    If Me.InsideHeight < mCloseButton.Top + mCloseButton.Height + 12 Then
        ' 화면에 표시할 때 픽셀 단위로 반올림되는 높이를 고려해 1pt를 더 확보한다.
        Me.Height = Me.Height + mCloseButton.Top + mCloseButton.Height + 13 - Me.InsideHeight
    End If

    If mEmployeeList.ListCount > 0 Then mEmployeeList.ListIndex = 0
    If mEmployeeList.ListIndex >= 0 Then mSelectedEmployeeIndex = CLng(mDisplayedEmployees(1))
    mLoading = False
    LoadEmployeeSelection
End Sub

Private Function AddLabel(ByVal controlName As String, ByVal captionText As String, _
                          ByVal leftValue As Single, ByVal topValue As Single, _
                          ByVal widthValue As Single, ByVal heightValue As Single) As MSForms.label
    Dim label As MSForms.label
    Set label = Me.Controls.Add("Forms.Label.1", controlName, True)
    With label
        .Caption = captionText
        .Left = leftValue
        .Top = topValue
        .Width = widthValue
        .Height = heightValue
        .BackStyle = fmBackStyleTransparent
        .Font.name = "맑은 고딕"
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
        .Width = 90
        .Height = 29
        .Font.name = "맑은 고딕"
        .Font.Size = 9
        .BackColor = RGB(223, 229, 233)
    End With
    Set AddActionButton = button
End Function

Private Sub mEmployeeList_Change()
    Dim searchText As String, caretPosition As Long
    If mLoading Then Exit Sub
    If mEmployeeList.ListIndex >= 0 Then
        mSelectedEmployeeIndex = CLng(mDisplayedEmployees(mEmployeeList.ListIndex + 1))
    Else
        searchText = mEmployeeList.text
        caretPosition = mEmployeeList.SelStart
        mSelectedEmployeeIndex = 0
        mLoading = True
        PopulateEmployeeList searchText
        mEmployeeList.text = searchText
        mEmployeeList.SelStart = caretPosition
        mEmployeeList.SelLength = 0
        If mEmployeeList.ListCount > 0 Then mEmployeeList.DropDown
        If mEmployeeList.ListIndex >= 0 Then
            mSelectedEmployeeIndex = CLng(mDisplayedEmployees(mEmployeeList.ListIndex + 1))
        End If
        mLoading = False
    End If
    LoadEmployeeSelection
End Sub

Private Sub PopulateEmployeeList(ByVal searchText As String)
    Dim employee As CEmployee, employeeIndex As Long, displayName As String
    Set mDisplayedEmployees = New Collection
    mEmployeeList.Clear
    For employeeIndex = 1 To mEmployees.count
        Set employee = mEmployees(employeeIndex)
        displayName = EmployeeDisplayName(employee)
        If EmployeePrefixMatches(displayName, searchText) Then
            mEmployeeList.AddItem displayName
            mDisplayedEmployees.Add employeeIndex
        End If
    Next employeeIndex
End Sub

Private Function EmployeeDisplayName(ByVal employee As CEmployee) As String
    EmployeeDisplayName = employee.name
    If gEmployeesByName(employee.name).count > 1 Then
        EmployeeDisplayName = employee.name & "(" & employee.birthdate & ")"
    End If
End Function

Private Function EmployeePrefixMatches(ByVal displayName As String, ByVal searchText As String) As Boolean
    Dim textLength As Long, searchCode As Long, candidateCode As Long, initialIndex As Long
    textLength = Len(searchText)
    If textLength = 0 Then
        EmployeePrefixMatches = True
        Exit Function
    End If
    If Len(displayName) < textLength Then Exit Function
    If StrComp(Left$(displayName, textLength - 1), Left$(searchText, textLength - 1), vbTextCompare) <> 0 Then Exit Function
    If StrComp(Mid$(displayName, textLength, 1), Right$(searchText, 1), vbTextCompare) = 0 Then
        EmployeePrefixMatches = True
        Exit Function
    End If
    searchCode = AscW(Right$(searchText, 1)) And &HFFFF&
    candidateCode = AscW(Mid$(displayName, textLength, 1)) And &HFFFF&
    If candidateCode < &HAC00& Or candidateCode > &HD7A3& Then Exit Function
    candidateCode = candidateCode - &HAC00&
    ' 마지막 열린 음절은 아직 받침을 조합하는 검색어로 취급한다. 기 -> 기/길/김.
    If searchCode >= &HAC00& And searchCode <= &HD7A3& Then
        searchCode = searchCode - &HAC00&
        If searchCode Mod 28 = 0 Then EmployeePrefixMatches = (candidateCode \ 28 = searchCode \ 28)
    Else
        initialIndex = InStr(1, "ㄱㄲㄴㄷㄸㄹㅁㅂㅃㅅㅆㅇㅈㅉㅊㅋㅌㅍㅎ", Right$(searchText, 1), vbBinaryCompare)
        If initialIndex > 0 Then EmployeePrefixMatches = (candidateCode \ 588 = initialIndex - 1)
    End If
End Function

Private Sub mEmployeeList_DropButtonClick()
    If mLoading Or mSelectedEmployeeIndex = 0 Then Exit Sub
    ' 이미 선택한 이름을 다시 펼칠 때에는 전체 교직원 목록을 보여 준다.
    mLoading = True
    PopulateEmployeeList vbNullString
    mEmployeeList.ListIndex = mSelectedEmployeeIndex - 1
    mLoading = False
End Sub

Private Sub mEmployeeList_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn And mSelectedEmployeeIndex = 0 And mEmployeeList.ListCount > 0 Then
        mLoading = True
        mEmployeeList.ListIndex = 0
        mSelectedEmployeeIndex = CLng(mDisplayedEmployees(1))
        mLoading = False
        LoadEmployeeSelection
        KeyCode = 0
    End If
End Sub

Private Sub RefreshRegisteredEmployees()
    Dim employee As CEmployee, employeeIndex As Long, savedLoading As Boolean
    Dim selectedRow As Long, previousTop As Long
    savedLoading = mLoading
    mLoading = True
    previousTop = mRegisteredList.topIndex
    selectedRow = -1
    Set mRegisteredEmployees = New Collection
    mRegisteredList.Clear
    For employeeIndex = 1 To mEmployees.count
        Set employee = mEmployees(employeeIndex)
        If employee.PersonnelChangeStart <> 0 And employee.PersonnelChangeEnd <> 0 Then
            mRegisteredList.AddItem EmployeeDisplayName(employee)
            mRegisteredList.List(mRegisteredList.ListCount - 1, 1) = _
                Format$(employee.PersonnelChangeStart, "yyyy-mm-dd") & "~" & Format$(employee.PersonnelChangeEnd, "yyyy-mm-dd")
            mRegisteredEmployees.Add employeeIndex
            If employeeIndex = mSelectedEmployeeIndex Then selectedRow = mRegisteredList.ListCount - 1
        End If
    Next employeeIndex
    If selectedRow >= 0 Then mRegisteredList.ListIndex = selectedRow
    If mRegisteredList.ListCount > 0 Then
        If previousTop >= mRegisteredList.ListCount Then previousTop = mRegisteredList.ListCount - 1
        If previousTop >= 0 Then mRegisteredList.topIndex = previousTop
    End If
    mLoading = savedLoading
End Sub

Private Sub mRegisteredList_Click()
    If mLoading Or mRegisteredList.ListIndex < 0 Then Exit Sub
    mSelectedEmployeeIndex = CLng(mRegisteredEmployees(mRegisteredList.ListIndex + 1))
    mLoading = True
    PopulateEmployeeList vbNullString
    mEmployeeList.ListIndex = mSelectedEmployeeIndex - 1
    mLoading = False
    LoadEmployeeSelection
End Sub

Private Sub LoadEmployeeSelection()
    Dim employee As CEmployee
    mStartDay = 0
    mEndDay = 0
    mSavedStartDay = 0
    mSavedEndDay = 0
    If Not mEmployeeList Is Nothing Then
        If mSelectedEmployeeIndex > 0 Then
            Set employee = mEmployees(mSelectedEmployeeIndex)
            mStartDay = employee.PersonnelChangeStart
            mEndDay = employee.PersonnelChangeEnd
            mSavedStartDay = mStartDay
            mSavedEndDay = mEndDay
        End If
    End If
    RefreshSelection
End Sub

Public Sub SelectDay(ByVal dayValue As Date)
    If mSelectedEmployeeIndex = 0 Then Exit Sub
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
    If mSelectedEmployeeIndex = 0 Then
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
        dayValue = handler.dayValue
        With handler.button
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
    If mSelectedEmployeeIndex = 0 Then RaiseValidation "인사변동 기간", "교직원을 선택하세요."
    Set employee = mEmployees(mSelectedEmployeeIndex)
    SavePersonnelChange mWorkMonth, employee.Order, mStartDay, mEndDay
    mSavedStartDay = mStartDay
    mSavedEndDay = mEndDay
    RefreshRegisteredEmployees
    RefreshSelection
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "인사변동자 등록"
End Sub

Private Sub mClearButton_Click()
    Dim employee As CEmployee
    On Error GoTo Failed
    If mSelectedEmployeeIndex = 0 Then RaiseValidation "인사변동 기간", "교직원을 선택하세요."
    Set employee = mEmployees(mSelectedEmployeeIndex)
    ClearPersonnelChange mWorkMonth, employee.Order
    mStartDay = 0
    mEndDay = 0
    mSavedStartDay = 0
    mSavedEndDay = 0
    RefreshRegisteredEmployees
    RefreshSelection
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "인사변동자 등록"
End Sub

Private Sub mCloseButton_Click()
    StopPersonnelWheel
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        StopPersonnelWheel
        Me.Hide
    End If
End Sub

Public Sub DisposeHandlers()
    Dim handler As CPersonnelDayButton
    StopPersonnelWheel
    If Not DayButtonHandlers Is Nothing Then
        For Each handler In DayButtonHandlers
            handler.Detach
        Next handler
    End If
    Set DayButtonHandlers = Nothing
    Set mEmployeeList = Nothing
    Set mRegisteredList = Nothing
    Set mDisplayedEmployees = Nothing
    Set mRegisteredEmployees = Nothing
    Set mSaveButton = Nothing
    Set mClearButton = Nothing
    Set mCloseButton = Nothing
End Sub

Private Sub UserForm_Activate()
    StartPersonnelWheel Me
End Sub

Private Sub UserForm_Deactivate()
    StopPersonnelWheel
End Sub

Public Function HandlePersonnelWheel(ByVal x As Single, ByVal y As Single, ByVal delta As Long, _
                                     ByVal overPopup As Boolean) As Boolean
    Dim rowStep As Long, nextIndex As Long, topIndex As Long, popupRows As Long
    If mLoading Or delta = 0 Or mEmployeeList Is Nothing Or mRegisteredList Is Nothing Then Exit Function
    rowStep = -(delta \ 120) * 3
    If rowStep = 0 Then rowStep = -Sgn(delta)
    topIndex = mEmployeeList.topIndex
    popupRows = mEmployeeList.ListCount
    If popupRows > mEmployeeList.ListRows Then popupRows = mEmployeeList.ListRows
    If x >= mEmployeeList.Left And x < mEmployeeList.Left + mEmployeeList.Width Then
        If (y >= mEmployeeList.Top And y < mEmployeeList.Top + mEmployeeList.Height) _
           Or (topIndex >= 0 And overPopup And popupRows > 0) Then
            If mEmployeeList.ListCount > 0 Then
                If topIndex >= 0 Then
                    nextIndex = topIndex + rowStep
                    If nextIndex < 0 Then nextIndex = 0
                    If nextIndex >= mEmployeeList.ListCount Then nextIndex = mEmployeeList.ListCount - 1
                    mEmployeeList.topIndex = nextIndex
                Else
                    nextIndex = mEmployeeList.ListIndex + rowStep
                    If nextIndex < 0 Then nextIndex = 0
                    If nextIndex >= mEmployeeList.ListCount Then nextIndex = mEmployeeList.ListCount - 1
                    mEmployeeList.ListIndex = nextIndex
                End If
            End If
            HandlePersonnelWheel = True
            Exit Function
        End If
    End If
    If overPopup Then Exit Function
    If x >= mRegisteredList.Left And x < mRegisteredList.Left + mRegisteredList.Width _
       And y >= mRegisteredList.Top And y < mRegisteredList.Top + mRegisteredList.Height Then
        If mRegisteredList.ListCount > 0 Then
            nextIndex = mRegisteredList.topIndex + rowStep
            If nextIndex < 0 Then nextIndex = 0
            If nextIndex >= mRegisteredList.ListCount Then nextIndex = mRegisteredList.ListCount - 1
            mRegisteredList.topIndex = nextIndex
        End If
        HandlePersonnelWheel = True
    End If
End Function

Private Sub UserForm_Terminate()
    DisposeHandlers
    Set mStatus = Nothing
    Set mEmployees = Nothing
End Sub
