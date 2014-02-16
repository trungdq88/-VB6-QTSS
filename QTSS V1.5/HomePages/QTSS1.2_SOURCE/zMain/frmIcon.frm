VERSION 5.00
Begin VB.Form frmIcon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Icon"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3600
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   3600
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Op 
      Height          =   255
      Index           =   10
      Left            =   360
      TabIndex        =   9
      Top             =   4440
      Width           =   255
   End
   Begin VB.OptionButton Op 
      Height          =   255
      Index           =   8
      Left            =   2040
      TabIndex        =   8
      Top             =   3240
      Width           =   255
   End
   Begin VB.OptionButton Op 
      Height          =   255
      Index           =   7
      Left            =   2040
      TabIndex        =   7
      Top             =   2160
      Width           =   255
   End
   Begin VB.OptionButton Op 
      Height          =   255
      Index           =   6
      Left            =   2040
      TabIndex        =   6
      Top             =   1080
      Width           =   255
   End
   Begin VB.OptionButton Op 
      Height          =   255
      Index           =   5
      Left            =   1200
      TabIndex        =   5
      Top             =   3240
      Width           =   255
   End
   Begin VB.OptionButton Op 
      Height          =   255
      Index           =   4
      Left            =   1200
      TabIndex        =   4
      Top             =   2160
      Width           =   255
   End
   Begin VB.OptionButton Op 
      Height          =   255
      Index           =   3
      Left            =   1200
      TabIndex        =   3
      Top             =   1080
      Width           =   255
   End
   Begin VB.OptionButton Op 
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   3240
      Width           =   255
   End
   Begin VB.OptionButton Op 
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   2160
      Width           =   255
   End
   Begin VB.OptionButton Op 
      Height          =   255
      Index           =   9
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image IconI 
      Height          =   615
      Index           =   10
      Left            =   240
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image IconI 
      Height          =   615
      Index           =   8
      Left            =   1920
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image IconI 
      Height          =   615
      Index           =   7
      Left            =   1920
      Top             =   1440
      Width           =   615
   End
   Begin VB.Image IconI 
      Height          =   615
      Index           =   6
      Left            =   1920
      Top             =   360
      Width           =   615
   End
   Begin VB.Image IconI 
      Height          =   615
      Index           =   5
      Left            =   1080
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image IconI 
      Height          =   615
      Index           =   4
      Left            =   1080
      Top             =   1440
      Width           =   615
   End
   Begin VB.Image IconI 
      Height          =   615
      Index           =   3
      Left            =   1080
      Top             =   360
      Width           =   615
   End
   Begin VB.Image IconI 
      Height          =   615
      Index           =   2
      Left            =   240
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image IconI 
      Height          =   615
      Index           =   1
      Left            =   240
      Top             =   1440
      Width           =   615
   End
   Begin VB.Image IconI 
      Height          =   615
      Index           =   9
      Left            =   240
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "frmIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Left = Form1.Left + Form1.Width
Me.Top = Form1.Top
Me.Height = Form1.Height

For i = 1 To 10
IconI(i).Picture = LoadResPicture(100 + i, 1)
Next i

End Sub

Private Sub Op_Click(Index As Integer)
Form1.L.Caption = 100 + Index
Form1.txtIcon.Text = "C:\WINDOWS\I.ICO"
End Sub

