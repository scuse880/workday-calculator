Attribute VB_Name = "PersonnelMouseWheel"
Option Explicit

' 현재 Excel UI thread에만 설치하며 폼 종료/비활성화 시 해제한다.
Private Const WH_MOUSE As Long = 7
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const GA_ROOT As Long = 2
Private Const GW_OWNER As Long = 4

Private Type WheelPoint
    x As Long
    y As Long
End Type

Private Type WheelRectangle
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

#If VBA7 Then
    Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExW" _
        (ByVal idHook As Long, ByVal hookProcedure As LongPtr, ByVal instance As LongPtr, ByVal threadId As Long) As LongPtr
    Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hookHandle As LongPtr) As Long
    Private Declare PtrSafe Function CallNextHookEx Lib "user32" _
        (ByVal hookHandle As LongPtr, ByVal code As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long
    Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function GetAncestor Lib "user32" (ByVal windowHandle As LongPtr, ByVal flags As Long) As LongPtr
    Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal windowHandle As LongPtr, ByVal command As Long) As LongPtr
    Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal windowHandle As LongPtr) As Long
    Private Declare PtrSafe Function GetClientRect Lib "user32" _
        (ByVal windowHandle As LongPtr, ByRef rectangle As WheelRectangle) As Long
    Private Declare PtrSafe Function GetWindowRect Lib "user32" _
        (ByVal windowHandle As LongPtr, ByRef rectangle As WheelRectangle) As Long
    Private Declare PtrSafe Function ScreenToClient Lib "user32" _
        (ByVal windowHandle As LongPtr, ByRef point As WheelPoint) As Long
#If Win64 Then
    Private Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal pointValue As LongLong) As LongPtr
#Else
    Private Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal pointX As Long, ByVal pointY As Long) As LongPtr
#End If
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (ByRef destination As Any, ByVal source As LongPtr, ByVal length As LongPtr)
    Private mHook As LongPtr
    Private mFormWindow As LongPtr
#Else
    Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExW" _
        (ByVal idHook As Long, ByVal hookProcedure As Long, ByVal instance As Long, ByVal threadId As Long) As Long
    Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hookHandle As Long) As Long
    Private Declare Function CallNextHookEx Lib "user32" _
        (ByVal hookHandle As Long, ByVal code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
    Private Declare Function GetActiveWindow Lib "user32" () As Long
    Private Declare Function GetAncestor Lib "user32" (ByVal windowHandle As Long, ByVal flags As Long) As Long
    Private Declare Function GetWindow Lib "user32" (ByVal windowHandle As Long, ByVal command As Long) As Long
    Private Declare Function IsWindowVisible Lib "user32" (ByVal windowHandle As Long) As Long
    Private Declare Function GetClientRect Lib "user32" _
        (ByVal windowHandle As Long, ByRef rectangle As WheelRectangle) As Long
    Private Declare Function GetWindowRect Lib "user32" _
        (ByVal windowHandle As Long, ByRef rectangle As WheelRectangle) As Long
    Private Declare Function ScreenToClient Lib "user32" _
        (ByVal windowHandle As Long, ByRef point As WheelPoint) As Long
    Private Declare Function WindowFromPoint Lib "user32" (ByVal pointX As Long, ByVal pointY As Long) As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (ByRef destination As Any, ByVal source As Long, ByVal length As Long)
    Private mHook As Long
    Private mFormWindow As Long
#End If

Private mWheelForm As frmPersonnelChange
Private mHandlingWheel As Boolean

Public Sub StartPersonnelWheel(ByVal dialog As frmPersonnelChange)
    StopPersonnelWheel
    mFormWindow = GetActiveWindow()
    If mFormWindow = 0 Then Exit Sub
    Set mWheelForm = dialog
    mHook = SetWindowsHookEx(WH_MOUSE, AddressOf PersonnelMouseProcedure, 0, GetCurrentThreadId())
    If mHook = 0 Then
        StopPersonnelWheel
        Err.Raise vbObjectError + 2201, "PersonnelMouseWheel", "인사변동자 목록의 마우스 휠을 준비할 수 없습니다."
    End If
End Sub

Public Sub StopPersonnelWheel()
    If mHook <> 0 Then UnhookWindowsHookEx mHook
    mHook = 0
    mFormWindow = 0
    mHandlingWheel = False
    Set mWheelForm = Nothing
End Sub

#If VBA7 Then
Public Function PersonnelMouseProcedure(ByVal code As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Dim eventWindow As LongPtr, ownerWindow As LongPtr
#Else
Public Function PersonnelMouseProcedure(ByVal code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim eventWindow As Long, ownerWindow As Long
#End If
    Dim point As WheelPoint, bounds As WheelRectangle, popupBounds As WheelRectangle
    Dim mouseData As Long, delta As Long, overPopup As Boolean, ownerDepth As Long
    Dim x As Single, y As Single
#If Win64 Then
    Dim packedPoint As LongLong
#End If

    On Error GoTo Failed
    If code < 0 Or wParam <> WM_MOUSEWHEEL Or mHandlingWheel Then GoTo PassToNext
    If mWheelForm Is Nothing Or mFormWindow = 0 Then GoTo PassToNext
    If GetActiveWindow() <> mFormWindow Or IsWindowVisible(mFormWindow) = 0 Then GoTo PassToNext
    CopyMemory point, lParam, 8
#If Win64 Then
    CopyMemory packedPoint, VarPtr(point), 8
    eventWindow = WindowFromPoint(packedPoint)
#Else
    eventWindow = WindowFromPoint(point.x, point.y)
#End If
    ' MOUSEHOOKSTRUCTEX의 mouseData는 Win64에서 32, Win32에서 20 byte 뒤에 있다.
#If Win64 Then
    CopyMemory mouseData, lParam + 32, 4
#Else
    CopyMemory mouseData, lParam + 20, 4
#End If
    delta = (mouseData And &H7FFF0000) \ 65536
    If mouseData < 0 Then delta = delta - 32768

    eventWindow = GetAncestor(eventWindow, GA_ROOT)
    If eventWindow <> mFormWindow Then
        ' 펼친 ComboBox 팝업만 폼 바깥 좌표를 허용한다. 다른 창에는 전달한다.
        ownerWindow = eventWindow
        For ownerDepth = 1 To 8
            ownerWindow = GetWindow(ownerWindow, GW_OWNER)
            If ownerWindow = mFormWindow Then Exit For
            If ownerWindow = 0 Then GoTo PassToNext
        Next ownerDepth
        If ownerWindow <> mFormWindow Then GoTo PassToNext
        If GetWindowRect(eventWindow, popupBounds) = 0 Then GoTo PassToNext
        overPopup = (point.x >= popupBounds.Left And point.x < popupBounds.Right _
                    And point.y >= popupBounds.Top And point.y < popupBounds.Bottom)
        If Not overPopup Then GoTo PassToNext
    End If
    If GetClientRect(mFormWindow, bounds) = 0 Then GoTo PassToNext
    If bounds.Right <= 0 Or bounds.Bottom <= 0 Then GoTo PassToNext
    If ScreenToClient(mFormWindow, point) = 0 Then GoTo PassToNext
    ' 창의 실제 client 크기로 환산해 Windows 배율이 달라도 control 위치와 맞춘다.
    x = CSng(point.x) * mWheelForm.InsideWidth / bounds.Right
    y = CSng(point.y) * mWheelForm.InsideHeight / bounds.Bottom
    mHandlingWheel = True
    If mWheelForm.HandlePersonnelWheel(x, y, delta, overPopup) Then
        mHandlingWheel = False
        PersonnelMouseProcedure = 1
        Exit Function
    End If
    mHandlingWheel = False
PassToNext:
    PersonnelMouseProcedure = CallNextHookEx(mHook, code, wParam, lParam)
    Exit Function
Failed:
    StopPersonnelWheel
    Resume PassToNext
End Function
