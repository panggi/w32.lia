Attribute VB_Name = "mdlDir"
Private Enum FileParts
    ExtOnly
    NameOnly
    NameExt
    PathOnly
End Enum

Private Function FilePart(FullPath As String, Optional WhichPart _
    As FileParts = NameOnly) As String
Dim lArray As Variant
Select Case WhichPart
Case ExtOnly
If InStr(FullPath, ".") Then
lArray = Split(FullPath, ".")
FilePart = lArray(UBound(lArray))
End If
Case NameOnly, NameExt
lArray = Split(FullPath, "\")
FilePart = lArray(UBound(lArray))
If WhichPart = NameOnly Then
lArray = Split(FilePart, ".")
FilePart = lArray(LBound(lArray))
End If
Case PathOnly
Dim lFileName As String
lFileName = FilePart(FullPath, NameExt)
FilePart = Replace(FullPath, lFileName, "")
End Select
End Function

Public Function GetDir(List As ListBox)
On Error Resume Next
Dim stra As String
Dim objShell32 As New Shell32.Shell
Dim objWindows As Object
Dim lngCounter As Long
Set objWindows = objShell32.Windows
List.Clear
For Each objWindows In objShell32.Windows
stra = ChangeAll(Mid(objWindows.LocationURL, 9, Len _
    (objWindows.LocationURL) - 8), "/", "\")
stra = ChangeAll(stra, "%20", " ")
stra = ChangeAll(stra, "%7b", "{")
stra = ChangeAll(stra, "%7d", "}")
stra = ChangeAll(stra, "%5b", "[")
stra = ChangeAll(stra, "%5d", "]")
stra = ChangeAll(stra, "%60", "`")
stra = ChangeAll(stra, "%23", "#")
stra = ChangeAll(stra, "%25", "%")
stra = ChangeAll(stra, "%5e", "^")
stra = ChangeAll(stra, "%26", "&")

If Left(Right(stra, 4), 1) = "." Or Left(Right(stra, 3), 1) = "." Then
stra = FilePart(stra, PathOnly)
End If
If Right(stra, 1) <> "\" Then
stra = stra & "\"
End If
If Mid(stra, 2, 2) = ":\" Then
    List.AddItem stra
End If
Next
Set objWindows = Nothing
Set objShell32 = Nothing
End Function

Private Function ChangeAll(Source As String, Search As String, _
    Restring As String) As String
Dim hitung As Integer
hitung = Len(Replace(Source, Search, Search & "*")) - Len(Source)
ChangeAll = Replace(LCase(Source), Search, Restring, 1, hitung)
End Function
