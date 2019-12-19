Attribute VB_Name = "全局热键"
Private Declare Function CallNextHookEx Lib "user32" _
(ByVal hHook As Long, _
ByVal nCode As Long, _
ByVal wParam As Long, _
lParam As Any) As Long
Private Declare Function SetWindowsHookEx Lib "user32" _
Alias "SetWindowsHookExA" _
(ByVal idHook As Long, _
ByVal lpfn As Long, _
ByVal hmod As Long, _
ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" _
(ByVal hHook As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" _
Alias "RtlMoveMemory" _
(Destination As Any, _
Source As Any, _
ByVal Length As Long)
Private Type PKBDLLHOOKSTRUCT
vkCode As Long
scanCode As Long
flags As Long
time As Long
dwExtraInfo As Long
End Type
Private Const WM_KEYDOWN = &H100
Private Const WM_SYSKEYDOWN = &H104
Private Const WM_KEYUP = &H101
Private Const WM_SYSKEYUP = &H105
Private Const HC_ACTION = 0
Private Const WH_KEYBOARD_LL = 13
Private lngHook As Long
Public Function HotKey(ByVal nCode As Long, _
ByVal wParam As Long, _
ByVal lParam As Long) As Long
Dim p As PKBDLLHOOKSTRUCT
If nCode = HC_ACTION Then
Select Case wParam
Case WM_KEYDOWN, WM_SYSKEYDOWN
Call CopyMemory(p, ByVal lParam, Len(p))
Select Case p.vkCode
Case vbKeySpace '空格
Call 设置窗口.开始演奏_Click
End Select
Case Else
End Select
End If
Call CallNextHookEx(WH_KEYBOARD_LL, nCode, wParam, lParam)
End Function
Public Sub HooK()
lngHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf HotKey, App.hInstance, 0)
End Sub
Public Sub UnHooK()
Call UnhookWindowsHookEx(lngHook)
End Sub

