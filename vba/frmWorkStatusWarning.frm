VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmWorkStatusWarning
   Caption         =   "제외한 근무상황 기록"
   ClientHeight    =   6600
   ClientWidth     =   8700
   StartUpPosition =   1
End
Attribute VB_Name = "frmWorkStatusWarning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' @RuntimeForm
' 긴 불일치 경고도 한 사람당 한 번만 표시하며 상세 내용 전체를 스크롤로 확인한다.

Private mDetails As MSForms.TextBox
Private WithEvents mConfirm As MSForms.CommandButton

Private Sub UserForm_Initialize()
    Me.Caption = "제외한 근무상황 기록"
    Me.Width = 580
    Me.Height = 440
    Me.StartUpPosition = 1
    Me.BackColor = RGB(245, 245, 245)
    Me.Font.Name = "맑은 고딕"
    Me.Font.Size = 10
    Set mDetails = Me.Controls.Add("Forms.TextBox.1", "txtDetails", True)
    With mDetails
        .Left = 16
        .Top = 16
        .Width = 536
        .Height = 338
        .MultiLine = True
        .WordWrap = False
        .ScrollBars = fmScrollBarsBoth
        .Locked = True
        .MaxLength = 0
        .EnterKeyBehavior = False
        .BackColor = vbWhite
        .TabIndex = 0
    End With
    Set mConfirm = Me.Controls.Add("Forms.CommandButton.1", "cmdConfirm", True)
    With mConfirm
        .Caption = "확인"
        .Left = 462
        .Top = 369
        .Width = 90
        .Height = 27
        .Default = True
        .Cancel = True
        .TabIndex = 1
    End With
End Sub

Public Sub Configure(ByVal text As String, ByVal title As String)
    Me.Caption = title
    mDetails.Value = text
    mDetails.SelStart = 0
    mDetails.SelLength = 0
End Sub

Private Sub mConfirm_Click()
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        Me.Hide
    End If
End Sub
