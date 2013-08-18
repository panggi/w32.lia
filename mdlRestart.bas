Attribute VB_Name = "mdlRestart"
'Contoh penggunaan fungsi SendKeys
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds _
    As Long)
Public Declare Sub keybd_event Lib "user32" (ByVal bVk _
    As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal _
    dwExtraInfo As Long)
Public Declare Function MapVirtualKey Lib "user32" Alias _
    "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType _
    As Long) As Long
Public Const KEYEVENTF_KEYUP = &H2
Public Const VK_CONTROL = &H11
Public Const VK_ESCAPE = &H1B
Public Const VK_R = &H52
Public Const VK_TAB = &H9
Public Const VK_SPACE = &H20
Public Const VK_UP = &H26
Public Const VK_RETURN = &HD

Public Function Restart()
Dim aba, baa, aab As Integer
Dim bab, bba, abb As Integer
Dim abc As Integer
aab = MapVirtualKey(VK_UP, 0)
aba = MapVirtualKey(VK_CONTROL, 0)
baa = MapVirtualKey(VK_ESCAPE, 0)
bab = MapVirtualKey(VK_R, 0)
bba = MapVirtualKey(VK_TAB, 0)
abb = MapVirtualKey(VK_SPACE, 0)
abc = MapVirtualKey(VK_RETURN, 0)
keybd_event VK_CONTROL, aba, 0, 0
keybd_event VK_ESCAPE, baa, 0, 0
keybd_event VK_CONTROL, aba, KEYEVENTF_KEYUP, 0
keybd_event VK_ESCAPE, baa, KEYEVENTF_KEYUP, 0
Sleep (0)
keybd_event VK_UP, aab, 0, 0
keybd_event VK_UP, aab, KEYEVENTF_KEYUP, 0
Sleep (0)
keybd_event VK_RETURN, abc, 0, 0
keybd_event VK_RETURN, abc, KEYEVENTF_KEYUP, 0
Sleep (0)
keybd_event VK_R, bab, 0, 0
keybd_event VK_R, bab, KEYEVENTF_KEYUP, 0
Sleep (0)
keybd_event VK_TAB, bba, 0, 0
keybd_event VK_TAB, bba, KEYEVENTF_KEYUP, 0
Sleep (0)
keybd_event VK_SPACE, abb, 0, 0
keybd_event VK_SPACE, abb, KEYEVENTF_KEYUP, 0
Sleep (0)
End Function
