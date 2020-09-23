VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   Caption         =   "Memory - 1999 by Swertvaegher Stephan"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5040
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   322
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   336
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4410
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Difficulty"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2355
      Index           =   2
      Left            =   2655
      TabIndex        =   10
      Top             =   180
      Width           =   2220
      Begin VB.OptionButton Option3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Are you NUTS !?!?"
         ForeColor       =   &H0080FF80&
         Height          =   240
         Index           =   4
         Left            =   225
         TabIndex        =   15
         Top             =   2025
         Width           =   1860
      End
      Begin VB.OptionButton Option3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Impossible"
         ForeColor       =   &H0080FF80&
         Height          =   240
         Index           =   3
         Left            =   225
         TabIndex        =   14
         Top             =   1620
         Width           =   1860
      End
      Begin VB.OptionButton Option3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Heavy"
         ForeColor       =   &H0080FF80&
         Height          =   240
         Index           =   2
         Left            =   225
         TabIndex        =   13
         Top             =   1215
         Width           =   1860
      End
      Begin VB.OptionButton Option3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Not so easy"
         ForeColor       =   &H0080FF80&
         Height          =   240
         Index           =   1
         Left            =   225
         TabIndex        =   12
         Top             =   855
         Value           =   -1  'True
         Width           =   1860
      End
      Begin VB.OptionButton Option3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Easy"
         ForeColor       =   &H0080FF80&
         Height          =   240
         Index           =   0
         Left            =   225
         TabIndex        =   11
         Top             =   495
         Width           =   1860
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Backstyle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1410
      Index           =   1
      Left            =   180
      TabIndex        =   4
      Top             =   2655
      Width           =   4695
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Index           =   4
         Left            =   3825
         TabIndex        =   9
         Top             =   495
         Width           =   210
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Index           =   3
         Left            =   3015
         TabIndex        =   8
         Top             =   495
         Width           =   210
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Index           =   2
         Left            =   2205
         TabIndex        =   7
         Top             =   495
         Value           =   -1  'True
         Width           =   210
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Index           =   1
         Left            =   1395
         TabIndex        =   6
         Top             =   495
         Width           =   210
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Index           =   0
         Left            =   630
         TabIndex        =   5
         Top             =   495
         Width           =   210
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   4
         Left            =   3690
         Picture         =   "Form2.frx":030A
         Top             =   765
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   3
         Left            =   2880
         Picture         =   "Form2.frx":0614
         Top             =   765
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   2070
         Picture         =   "Form2.frx":091E
         Top             =   765
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   1260
         Picture         =   "Form2.frx":0C28
         Top             =   765
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   495
         Picture         =   "Form2.frx":0F32
         Top             =   765
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Style"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2355
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   180
      Width           =   2220
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "Traffic"
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Index           =   2
         Left            =   225
         TabIndex        =   3
         Top             =   1755
         Width           =   1200
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "Science"
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Index           =   1
         Left            =   225
         TabIndex        =   2
         Top             =   1080
         Width           =   1200
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "Cartoons"
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Index           =   0
         Left            =   225
         TabIndex        =   1
         Top             =   405
         Value           =   -1  'True
         Width           =   1200
      End
      Begin VB.Image Image2 
         Height          =   480
         Index           =   2
         Left            =   1485
         Picture         =   "Form2.frx":123C
         Top             =   1575
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   480
         Index           =   1
         Left            =   1485
         Picture         =   "Form2.frx":1546
         Top             =   945
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   480
         Index           =   0
         Left            =   1530
         Picture         =   "Form2.frx":1850
         Top             =   315
         Width           =   480
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Start Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1575
      TabIndex        =   16
      Top             =   4230
      Width           =   1770
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim col(31), tt, tt0, zq As Integer

Private Sub Form_Load()
tt0 = 15
col(0) = 0
col(1) = 16
col(2) = 32
col(3) = 48
col(4) = 64
col(5) = 80
col(6) = 96
col(7) = 112
col(8) = 128
col(9) = 144
col(10) = 160
col(11) = 176
col(12) = 192
col(13) = 208
col(14) = 224
col(15) = 240
col(16) = 255
col(17) = 240
col(18) = 224
col(19) = 208
col(20) = 192
col(21) = 176
col(22) = 160
col(23) = 144
col(24) = 128
col(25) = 112
col(26) = 96
col(27) = 80
col(28) = 64
col(29) = 48
col(30) = 32
col(31) = 16
Option3(1).ForeColor = RGB(255, 255, 0)
Option3(1).Value = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub Image1_Click(Index As Integer)
Option2(Index).Value = True
End Sub

Private Sub Image2_Click(Index As Integer)
Option1(Index).Value = True
End Sub

Private Sub Label1_Click()
Form2.Hide
Form1.Enabled = True
Form1.Show
End Sub

'Private Sub Label2_Click()
'End
'End Sub

Private Sub Option3_Click(Index As Integer)
For zq = 0 To 4
Option3(zq).ForeColor = &H80FF80
Next zq
Option3(Index).ForeColor = RGB(255, 255, 0)
End Sub

Private Sub Timer1_Timer()
Label1.ForeColor = RGB(col(tt), 64, 64)
'Label2.ForeColor = RGB(col(tt0), 64, 64)
tt = tt + 1
If tt = 32 Then tt = 0
tt0 = tt0 + 1
If tt0 = 32 Then tt0 = 0
End Sub
