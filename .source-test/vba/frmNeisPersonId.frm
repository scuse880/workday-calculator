VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNeisPersonId
   Caption         =   "나이스 개인번호 등록"
   ClientHeight    =   7200
   ClientWidth     =   7800
   StartUpPosition =   1
End
Attribute VB_Name = "frmNeisPersonId"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' @RuntimeForm
' 컨트롤은 코드로 생성한다. import-vba.ps1은 VBComponents.Add(3) 후 이 코드 본문을 넣는다.

Public Accepted As Boolean
Private mEmployees As Collection
Private mInputs As Collection
Private mRows As MSForms.Frame
Private WithEvents mConfirm As MSForms.CommandButton
Private WithEvents mCancel As MSForms.CommandButton

Private Sub UserForm_Initialize()
    Dim label As MSForms.label
    Me.Caption = "나이스 개인번호 등록"
    Me.Width = 530
    Me.Height = 480
    Me.StartUpPosition = 1
    Me.BackColor = RGB(245, 245, 245)
    Me.Font.Name = "맑은 고딕"
    Me.Font.Size = 10
    Set label = Me.Controls.Add("Forms.Label.1", "lblTitle", True)
    With label
        .Caption = "동명이인 식별정보 등록"
        .Left = 16
        .Top = 14
        .Width = 470
        .Height = 22
        .Font.Bold = True
        .Font.Size = 13
        .BackStyle = 0
    End With
    Set label = Me.Controls.Add("Forms.Label.1", "lblNameHeader", True)
    label.Caption = "성명(생년월일)"
    label.Left = 24
    label.Top = 47
    label.Width = 220
    label.Height = 18
    label.BackStyle = 0
    Set label = Me.Controls.Add("Forms.Label.1", "lblIdHeader", True)
    label.Caption = "나이스 개인번호 (C + 숫자 9자리)"
    label.Left = 258
    label.Top = 47
    label.Width = 230
    label.Height = 18
    label.BackStyle = 0
    Set mRows = Me.Controls.Add("Forms.Frame.1", "fraRows", True)
    With mRows
        .Caption = vbNullString
        .Left = 16
        .Top = 67
        .Width = 488
        .Height = 290
        .ScrollBars = 2
        .KeepScrollBarsVisible = 2
        .BackColor = vbWhite
        .SpecialEffect = 0
    End With
    Set label = Me.Controls.Add("Forms.Label.1", "lblGuide", True)
    With label
        .Caption = "동명이인의 나이스 개인번호를 입력하세요" & vbCrLf & "개인번호는 대상자마다 달라야 합니다. 취소하면 기존 자료가 유지됩니다."
        .Left = 16
        .Top = 368
        .Width = 488
        .Height = 36
        .BackStyle = 0
        .Font.Size = 9
    End With
    Set mConfirm = Me.Controls.Add("Forms.CommandButton.1", "cmdConfirm", True)
    With mConfirm
        .Caption = "확인"
        .Left = 314
        .Top = 414
        .Width = 90
        .Height = 26
        .Default = True
    End With
    Set mCancel = Me.Controls.Add("Forms.CommandButton.1", "cmdCancel", True)
    With mCancel
        .Caption = "취소"
        .Left = 414
        .Top = 414
        .Width = 90
        .Height = 26
        .Cancel = True
    End With
End Sub

Public Sub Configure(ByVal employees As Collection)
    Dim employee As CEmployee, i As Long, label As MSForms.label, idBox As MSForms.TextBox
    Set mEmployees = employees
    Set mInputs = New Collection
    Accepted = False
    For i = 1 To mEmployees.count
        Set employee = mEmployees(i)
        Set label = mRows.Controls.Add("Forms.Label.1", "lblEmployee" & i, True)
        With label
            .Caption = employee.Name & "(" & employee.Birthdate & ")"
            .Left = 8
            .Top = 8 + (i - 1) * 30
            .Width = 222
            .Height = 22
            .BackStyle = 0
        End With
        Set idBox = mRows.Controls.Add("Forms.TextBox.1", "txtPersonId" & i, True)
        With idBox
            .Left = 242
            .Top = 5 + (i - 1) * 30
            .Width = 212
            .Height = 24
            .MaxLength = 40
            .Text = employee.NeisPersonId
            .TabIndex = i - 1
        End With
        mInputs.Add idBox
    Next i
    mRows.ScrollHeight = mEmployees.count * 30 + 12
    If mRows.ScrollHeight < mRows.Height Then mRows.ScrollHeight = mRows.Height
End Sub

Private Sub mConfirm_Click()
    Dim i As Long, idBox As MSForms.TextBox, normalized As String
    Dim values As Collection, used As Object, employee As CEmployee
    Set values = New Collection
    Set used = NewDictionary()
    For i = 1 To mInputs.count
        Set idBox = mInputs(i)
        If Not TryNormalizeNeisId(idBox.Text, normalized) Then
            FocusInvalid idBox, i, "나이스 개인번호를 확인하세요"
            Exit Sub
        End If
        If used.Exists(normalized) Then
            FocusInvalid idBox, i, "나이스 개인번호를 확인하세요" & vbCrLf & "서로 다른 대상자에게 같은 개인번호를 입력할 수 없습니다."
            Exit Sub
        End If
        used.Add normalized, True
        values.Add normalized
    Next i
    For i = 1 To mEmployees.count
        Set employee = mEmployees(i)
        employee.NeisPersonId = CStr(values(i))
    Next i
    Accepted = True
    Me.Hide
End Sub

Private Sub FocusInvalid(ByVal idBox As MSForms.TextBox, ByVal rowIndex As Long, ByVal message As String)
    Dim offset As Single, maximum As Single
    MsgBox message, vbExclamation, "입력 확인"
    offset = (rowIndex - 1) * 30
    maximum = mRows.ScrollHeight - mRows.InsideHeight
    If maximum < 0 Then maximum = 0
    If offset > maximum Then offset = maximum
    mRows.ScrollTop = offset
    idBox.SetFocus
    idBox.SelStart = 0
    idBox.SelLength = Len(idBox.Text)
End Sub

Private Sub mCancel_Click()
    Accepted = False
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        Accepted = False
        Me.Hide
    End If
End Sub
