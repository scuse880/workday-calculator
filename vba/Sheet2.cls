VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "Sheet2"
Option Explicit
' @DocumentModule:방학중근무

Private Sub Worksheet_Change(ByVal Target As Range)
    RefreshVacationBorders Me, Target
End Sub
