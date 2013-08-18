Attribute VB_Name = "mdlCrypt"
Option Explicit

Public Function Encrypt(ByVal TextInput As String) As String
Dim NewLen As Integer
Dim NewTextInput As String
Dim NewChar As String
Dim i As Integer
NewChar = ""
NewLen = Len(TextInput)
For i = 1 To NewLen
NewChar = Mid(TextInput, i, 1)
Select Case Asc(NewChar)
Case 65 To 90
NewChar = Chr(Asc(NewChar) + 127)
Case 97 To 122
NewChar = Chr(Asc(NewChar) + 121)
Case 48 To 57
NewChar = Chr(Asc(NewChar) + 196)
Case 32
NewChar = Chr(32)
End Select
NewTextInput = NewTextInput + NewChar
Next
Encrypt = NewTextInput
End Function

Public Function Decrypt(ByVal TextInput As String) As String
Dim NewLen As Integer
Dim NewTextInput As String
Dim NewChar As String
Dim i As Integer
NewChar = ""
NewLen = Len(TextInput)
For i = 1 To NewLen
NewChar = Mid(TextInput, i, 1)
Select Case Asc(NewChar)
Case 192 To 217
NewChar = Chr(Asc(NewChar) - 127)
Case 218 To 243
NewChar = Chr(Asc(NewChar) - 121)
Case 244 To 253
NewChar = Chr(Asc(NewChar) - 196)
Case 32
NewChar = Chr(32)
End Select
NewTextInput = NewTextInput + NewChar
Next
Decrypt = NewTextInput
End Function
