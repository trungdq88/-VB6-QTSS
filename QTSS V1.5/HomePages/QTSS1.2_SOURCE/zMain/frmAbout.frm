VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H80000007&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   3465
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5925
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2391.604
   ScaleMode       =   0  'User
   ScaleWidth      =   5563.881
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4560
      TabIndex        =   0
      Top             =   2880
      Width           =   1185
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000007&
      Caption         =   "Tester : Amen0111@Yahoo.Com"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1875
      Left            =   120
      Picture         =   "frmAbout.frx":6852
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2340
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Caption         =   "Enjoy !"
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   2040
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "Writer : DinhQuangTrung90@Yahoo.Com"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   1440
      Width           =   3975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   6310.427
      Y1              =   1770.408
      Y2              =   1770.408
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H80000007&
      Caption         =   "This Program Is Free. Build In 2009."
      ForeColor       =   &H0000FF00&
      Height          =   210
      Left            =   2610
      TabIndex        =   1
      Top             =   1125
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H80000007&
      Caption         =   "QTSS V1.2 - SPY FOR HACKER"
      ForeColor       =   &H0000FF00&
      Height          =   480
      Left            =   2610
      TabIndex        =   3
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   6310.427
      Y1              =   1780.761
      Y2              =   1780.761
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H80000007&
      Caption         =   "Version 1.2"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   2610
      TabIndex        =   4
      Top             =   780
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      BackColor       =   &H80000007&
      Caption         =   "Warning: Anti Virus May Be Detected And Delete This Program!"
      ForeColor       =   &H0000FF00&
      Height          =   825
      Left            =   255
      TabIndex        =   2
      Top             =   2745
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
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

Private Sub cmdOK_Click()
Form1.Show
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "About " & App.Title
End Sub

Private Sub lblTitle_Click()

End Sub
