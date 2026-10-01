VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "현재_통합_문서"
Option Explicit
' @DocumentModule:Workbook

Private Sub Workbook_Open()
    InitializeWorkbookProtection
End Sub

Private Sub Workbook_Activate()
    ' VBA를 재가져온 경우에도 저장되지 않는 UserInterfaceOnly 설정을 복원한다.
    ApplyWorkbookProtection
End Sub
