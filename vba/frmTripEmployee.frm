VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTripEmployee
   Caption         =   "출장 기록 대상자 확인"
   ClientHeight    =   5880
   ClientWidth     =   8700
   StartUpPosition =   1
End
Attribute VB_Name = "frmTripEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' @RuntimeForm
' Controls are created at run time; this source does not require an .frx file.
Private WithEvents mCandidates As MSForms.ListBox
Private WithEvents mExclude As MSForms.CheckBox
Private WithEvents mConfirm As MSForms.CommandButton
Private WithEvents mCancel As MSForms.CommandButton
Private mOriginal As MSForms.Label
Private mParsed As MSForms.Label
Private mPrefix As MSForms.Label
Private mRow As MSForms.Label
Private mEmployees As Collection
Private mSelected As CEmployee
Private mCancelled As Boolean
Private mSkipped As Boolean

Private Sub UserForm_Initialize()
    mCancelled = True
    Me.Caption = "출장 기록 대상자 확인"
    Me.Width = 580
    Me.Height = 398
    Me.BackColor = RGB(246, 246, 246)
    Me.Font.Name = "맑은 고딕"
    Me.Font.Size = 10

    AddLabel "lblTitle", "출장 근무상황부 · 대상자 확인", 18, 14, 536, 25, True
    Set mRow = AddLabel("lblRow", "", 18, 44, 536, 20)
    Set mOriginal = AddLabel("lblOriginal", "", 18, 68, 536, 38)
    Set mParsed = AddLabel("lblParsed", "", 18, 108, 536, 20)
    Set mPrefix = AddLabel("lblPrefix", "", 18, 132, 536, 20)
    AddLabel "lblName", "성명", 22, 161, 164, 19, True
    AddLabel "lblBirth", "생년월일", 190, 161, 102, 19, True
    AddLabel "lblNeis", "나이스 개인번호", 302, 161, 238, 19, True

    Set mCandidates = Me.Controls.Add("Forms.ListBox.1", "lstCandidates", True)
    With mCandidates
        .Left = 18
        .Top = 182
        .Width = 536
        .Height = 98
        .ColumnCount = 3
        .ColumnWidths = "168 pt;112 pt;230 pt"
        .MultiSelect = fmMultiSelectSingle
        .ListStyle = fmListStylePlain
        .TabIndex = 0
    End With

    Set mExclude = Me.Controls.Add("Forms.CheckBox.1", "chkExclude", True)
    With mExclude
        .Caption = "근무일수를 계산하지 않는 직종인 경우"
        .Left = 18
        .Top = 292
        .Width = 410
        .Height = 22
        .BackColor = Me.BackColor
        .TripleState = False
        .Value = False
        .TabIndex = 1
    End With

    Set mConfirm = Me.Controls.Add("Forms.CommandButton.1", "cmdConfirm", True)
    With mConfirm
        .Caption = "확인"
        .Left = 366
        .Top = 329
        .Width = 90
        .Height = 27
        .Default = True
        .TabIndex = 2
    End With
    Set mCancel = Me.Controls.Add("Forms.CommandButton.1", "cmdCancel", True)
    With mCancel
        .Caption = "취소"
        .Left = 464
        .Top = 329
        .Width = 90
        .Height = 27
        .Cancel = True
        .TabIndex = 3
    End With
End Sub

Private Function AddLabel(ByVal controlName As String, ByVal captionText As String, _
                          ByVal leftPoint As Single, ByVal topPoint As Single, _
                          ByVal widthPoint As Single, ByVal heightPoint As Single, _
                          Optional ByVal boldText As Boolean = False) As MSForms.Label
    Dim label As MSForms.Label
    Set label = Me.Controls.Add("Forms.Label.1", controlName, True)
    With label
        .Caption = captionText
        .Left = leftPoint
        .Top = topPoint
        .Width = widthPoint
        .Height = heightPoint
        .BackStyle = fmBackStyleTransparent
        .WordWrap = True
        .Font.Bold = boldText
        If boldText Then .ForeColor = RGB(42, 56, 71)
    End With
    Set AddLabel = label
End Function

Public Sub Configure(ByVal candidates As Collection, ByVal originalName As String, _
                     ByVal parsedName As String, ByVal idPrefix As String, ByVal sourceRow As Long)
    Dim employee As CEmployee
    If candidates Is Nothing Then Err.Raise vbObjectError + 250, Me.Name, "대상자 목록이 없습니다."
    If candidates.Count = 0 Then Err.Raise vbObjectError + 251, Me.Name, "대상자 목록이 비어 있습니다."
    Set mEmployees = candidates
    Set mSelected = Nothing
    mCancelled = True
    mSkipped = False
    mRow.Caption = "입력 파일 " & CStr(sourceRow) & "행 · 이 기록의 대상자를 선택하십시오."
    mOriginal.Caption = "원본 표시 이름: " & originalName
    mParsed.Caption = "추출한 이름: " & parsedName
    If Len(idPrefix) = 0 Then
        mPrefix.Caption = "나이스 아이디 앞 3자리: 아이디 정보 없음"
    Else
        mPrefix.Caption = "나이스 아이디 앞 3자리: " & idPrefix & " (선택 참고 정보)"
    End If
    mCandidates.Clear
    For Each employee In candidates
        mCandidates.AddItem employee.Name
        mCandidates.List(mCandidates.ListCount - 1, 1) = employee.Birthdate
        mCandidates.List(mCandidates.ListCount - 1, 2) = employee.NeisPersonId
    Next employee
    mExclude.Value = False
    mCandidates.Enabled = True
    mCandidates.ListIndex = -1
End Sub

Public Property Get Cancelled() As Boolean
    Cancelled = mCancelled
End Property

Public Property Get Skipped() As Boolean
    Skipped = mSkipped
End Property

Public Property Get SelectedEmployee() As CEmployee
    Set SelectedEmployee = mSelected
End Property

Private Sub mExclude_Click()
    If mExclude.Value Then mCandidates.ListIndex = -1
    mCandidates.Enabled = Not CBool(mExclude.Value)
End Sub

Private Sub mConfirm_Click()
    If mExclude.Value Then
        Set mSelected = Nothing
        mSkipped = True
    Else
        If mCandidates.ListIndex < 0 Then
            MsgBox "목록에서 대상자 1명을 선택하세요.", vbExclamation, Me.Caption
            Exit Sub
        End If
        Set mSelected = mEmployees(mCandidates.ListIndex + 1)
        mSkipped = False
    End If
    mCancelled = False
    Me.Hide
End Sub

Private Sub mCancel_Click()
    CancelSelection
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        CancelSelection
    End If
End Sub

Private Sub CancelSelection()
    mCancelled = True
    mSkipped = False
    Set mSelected = Nothing
    Me.Hide
End Sub

Private Sub UserForm_Terminate()
    Set mEmployees = Nothing
    Set mSelected = Nothing
    Set mCandidates = Nothing
    Set mExclude = Nothing
    Set mConfirm = Nothing
    Set mCancel = Nothing
End Sub
