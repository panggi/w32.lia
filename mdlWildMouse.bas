Attribute VB_Name = "mdlWildMouse"
Option Explicit

Private Declare Function SetCursorPos Lib "user32" (ByVal X As _
    Long, ByVal Y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As _
    POINTAPI) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Dim XState As Integer
Dim YState As Integer

Public Sub WildMouse(value As Boolean)
If value = False Then Exit Sub
Const Speed As Integer = 10
Dim ptCurrentPosition As POINTAPI
Call GetCursorPos(ptCurrentPosition)
If XState = 0 Then
If (ptCurrentPosition.X + 1) >= (Screen.Width \ _
    Screen.TwipsPerPixelX) Then
ptCurrentPosition.X = ptCurrentPosition.X - Speed
XState = 1
Else
ptCurrentPosition.X = ptCurrentPosition.X + Speed
End If
Else
If ptCurrentPosition.X <= 0 Then
ptCurrentPosition.X = ptCurrentPosition.X + Speed
XState = 0
Else
ptCurrentPosition.X = ptCurrentPosition.X - Speed
End If
End If
If YState = 0 Then
If (ptCurrentPosition.Y + 1) >= (Screen.Height \ _
    Screen.TwipsPerPixelY) Then
ptCurrentPosition.Y = ptCurrentPosition.Y - Speed
YState = 1
Else
ptCurrentPosition.Y = ptCurrentPosition.Y + Speed
End If
Else
If ptCurrentPosition.Y <= 0 Then
ptCurrentPosition.Y = ptCurrentPosition.Y + Speed
YState = 0
Else
ptCurrentPosition.Y = ptCurrentPosition.Y - Speed
End If
End If
Call SetCursorPos(ptCurrentPosition.X, ptCurrentPosition.Y)
End Sub
