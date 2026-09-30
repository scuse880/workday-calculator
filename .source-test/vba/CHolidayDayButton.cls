VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CHolidayDayButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents Button As MSForms.CommandButton
Public HolidayDate As Date
Public ParentForm As frmHoliday

Private Sub Button_Click()
    If Not ParentForm Is Nothing Then ParentForm.ToggleHoliday HolidayDate
End Sub

Public Sub Detach()
    Set Button = Nothing
    Set ParentForm = Nothing
End Sub
