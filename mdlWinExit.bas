Attribute VB_Name = "mdlWinExit"
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" _
    (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal _
    hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal _
    hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal _
    dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal _
    hProcess As Long, ByVal uExitCode As Long) As Long
Private Const TH32CS_SNAPPROCESS As Long = 2&
Private Const MAX_PATH As Integer = 260
Private Type PROCESSENTRY32
dwSize As Long
cntUsage As Long
th32ProcessID As Long
th32DefaultHeapID As Long
th32ModuleID As Long
cntThreads As Long
th32ParentProcessID As Long
pcPriClassBase As Long
dwFlags As Long
szExeFile As String * MAX_PATH
End Type

Public Function WinExit(sExeNam As String)
Dim lLng As Long, lA As Long, lExCode As Long
Dim procObj As PROCESSENTRY32
Dim hSnap As Long
Dim lRet As Long
hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
procObj.dwSize = Len(procObj)
lRet = Process32First(hSnap, procObj)
Do While Process32Next(hSnap, procObj)
If InStr(1, LCase(procObj.szExeFile), LCase(sExeNam$)) > 0 Then
lLng = OpenProcess(&H1, ByVal 0&, procObj.th32ProcessID)
lA = TerminateProcess(lLng, lExCode)
Exit Do
End If
Loop
End Function
