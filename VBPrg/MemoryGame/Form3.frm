VERSION 5.00
Begin VB.Form Form3 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   4215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   4215
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3555
      Width           =   960
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      Height          =   600
      Left            =   405
      Top             =   90
      Width           =   3885
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1590
      Left            =   270
      TabIndex        =   3
      Top             =   1620
      Width           =   4155
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(c) 1999 by Swertvaegher Stephan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   225
      TabIndex        =   1
      Top             =   945
      Width           =   4290
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Memory"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   645
      Left            =   540
      TabIndex        =   0
      Top             =   90
      Width           =   3615
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form3.Hide
Form2.Show
End Sub

Private Sub Form_Load()
Call BoxGradient(Form3, 64, 64, 32, 1, 4, 3, True)
Label3.Caption = "Coded in about 4 hours on a lazy afternoon..." & vbCr
Label3.Caption = Label3.Caption & "The purpose of this game is to find 2 matching cards. "
Label3.Caption = Label3.Caption & "The computer does the same. The one with the most pair of matching cards has won !"

End Sub

