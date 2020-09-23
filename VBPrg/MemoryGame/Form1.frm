VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Memory - 1999 by Swertvaegher Stephan"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8325
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   448
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   555
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   45
      Top             =   3690
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   45
      Top             =   3195
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   45
      Top             =   2700
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   45
      Top             =   2205
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   45
      Top             =   1665
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Score"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1365
      Left            =   585
      TabIndex        =   2
      Top             =   5040
      Width           =   7170
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Index           =   1
         Left            =   5760
         TabIndex        =   11
         Top             =   810
         Width           =   510
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Index           =   0
         Left            =   5760
         TabIndex        =   10
         Top             =   405
         Width           =   510
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Pairs:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   1
         Left            =   4950
         TabIndex        =   9
         Top             =   810
         Width           =   690
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Pairs:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   0
         Left            =   4950
         TabIndex        =   8
         Top             =   405
         Width           =   690
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "000"
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
         Height          =   285
         Left            =   3105
         TabIndex        =   6
         Top             =   810
         Width           =   750
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "000"
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
         Height          =   285
         Left            =   3105
         TabIndex        =   5
         Top             =   405
         Width           =   750
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Computer points:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   1
         Left            =   675
         TabIndex        =   4
         Top             =   810
         Width           =   2310
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Player points:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   0
         Left            =   675
         TabIndex        =   3
         Top             =   405
         Width           =   2310
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Memory"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   4470
      Left            =   540
      TabIndex        =   1
      Top             =   45
      Width           =   7170
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Start new game"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   465
         Left            =   1575
         TabIndex        =   7
         Top             =   2430
         Visible         =   0   'False
         Width           =   4020
      End
      Begin VB.Image Image2 
         Height          =   600
         Left            =   1665
         Picture         =   "Form1.frx":030A
         Top             =   945
         Visible         =   0   'False
         Width           =   3750
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   135
         Top             =   315
         Width           =   480
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   135
      Top             =   5985
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   90
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1361
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":167D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1999
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1CB5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1FD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":22ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2609
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2925
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2C41
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2F5D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3279
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3595
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":38B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3BCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3EE9
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4205
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4521
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":483D
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4B59
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4E75
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5191
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":54AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":57C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5AE5
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5E01
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":611D
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6439
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6755
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6A71
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6D8D
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":70A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":73C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":76E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":79FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7D19
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8035
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8351
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":866D
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8989
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8CA5
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8FC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":92DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":95F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9915
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9C31
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9F4D
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A269
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A585
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A8A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":ABBD
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":AED9
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B1F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B511
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B82D
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BB49
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BE65
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C181
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C49D
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C7B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":CAD5
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":CDF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D10D
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D429
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D745
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":DA61
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":DD7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":E099
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":E3B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":E6D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":E9ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":ED09
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":F025
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":F341
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":F65D
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":F979
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":FC95
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":FFB1
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":102CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":105E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":10905
            Key             =   ""
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":10C21
            Key             =   ""
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":10F3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":11259
            Key             =   ""
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":11575
            Key             =   ""
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":11891
            Key             =   ""
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":11BAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":11EC9
            Key             =   ""
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":121E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":12501
            Key             =   ""
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1281D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "You are as powerfull as the computer !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1395
      TabIndex        =   0
      Top             =   4590
      Width           =   5505
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuStart 
         Caption         =   "Start new game"
      End
      Begin VB.Menu mnubar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit Game"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X, xx, Y, Pic(59), Pic1(59), dummy, pp As Integer
Dim PLcount, CPcount, PlPic(1), CpPic(1) As Integer
Dim PlCor(1), CpCor(1) As Integer
Dim PlPnt, CpPnt, CpPick, PlCards, CpCards, Cp As Integer
Dim aa(4) As String

Private Sub Form_Activate()
Image2.Visible = False
Label5.Visible = False
SetImages
PLcount = 0
CPcount = 0
PlPnt = 0
CpPnt = 0
PlCards = 0
CpCards = 0
Label7(0).Caption = "00"
Label7(1).Caption = "00"
Label3.Caption = "000"
Label4.Caption = "000"
Label1.Caption = "Player, pick two cards"
For X = 0 To 59
Image1(X).Visible = True
Image1(X).Tag = ""
Next X
' difficulty level
For X = 0 To 4
    If Form2.Option3(X).Value = True Then
    Cp = (X * 5) + 2
    Frame1.Caption = "Memory     Difficulty: " & aa(X)
    End If
Next X
End Sub

Private Sub Form_Load()
aa(0) = "Easy"
aa(1) = "Not so easy"
aa(2) = "Heavy"
aa(3) = "Impossible"
aa(4) = "NUTS !"

Form1.BackColor = &H0
For X = 1 To 59
Load Image1(X)
Image1(X).Visible = True
Next X
For X = 0 To 9
Image1(X).Top = 600
Image1(X + 10).Top = Image1(X).Top + (40 * 15)
Image1(X + 20).Top = Image1(X).Top + (80 * 15)
Image1(X + 30).Top = Image1(X).Top + (120 * 15)
Image1(X + 40).Top = Image1(X).Top + (160 * 15)
Image1(X + 50).Top = Image1(X).Top + (200 * 15)
Image1(X).Left = (X * 600) + (40 * 15)
Image1(X + 10).Left = (X * 600) + (40 * 15)
Image1(X + 20).Left = (X * 600) + (40 * 15)
Image1(X + 30).Left = (X * 600) + (40 * 15)
Image1(X + 40).Left = (X * 600) + (40 * 15)
Image1(X + 50).Left = (X * 600) + (40 * 15)
Next X
End Sub

Private Sub SetImages()
For X = 0 To 29
NotGood:
Randomize
Pic(X) = Int(Rnd * 30) + 1
If X > 0 Then
    For Y = 0 To X - 1
    If Pic(Y) = Pic(X) Then
    'Exit For
    GoTo NotGood
    End If
    Next Y
End If
Pic1(X) = 255
Next X

For X = 30 To 59
NotGood1:
Randomize
Pic(X) = Int(Rnd * 30) + 1
If X > 30 Then
    For Y = 30 To X - 1
    If Pic(Y) = Pic(X) Then
    'Exit For
    GoTo NotGood1
    End If
    Next Y
End If
Pic1(X) = 255
Next X

For X = 0 To 59
NotGood2:
Randomize
dummy = Int(Rnd * 60)
If Pic1(dummy) <> 255 Then
GoTo NotGood2
End If
Pic1(dummy) = Pic(X)
Next X

For X = 0 To 59
If Form2.Option1(1).Value = True Then Pic1(X) = Pic1(X) + 30
If Form2.Option1(2).Value = True Then Pic1(X) = Pic1(X) + 60
Next X

For X = 0 To 4
If Form2.Option2(X).Value = True Then
pp = X
Exit For
End If
Next X

For X = 0 To 59
Image1(X).Picture = Form2.Image1(pp).Picture
Next X

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub Image1_Click(Index As Integer) 'player picks
If PLcount = 2 Then Exit Sub
If Image1(Index).Tag = "1" Then Exit Sub
Image1(Index).Picture = ImageList1.ListImages(Pic1(Index)).Picture
PlPic(PLcount) = Pic1(Index)
PlCor(PLcount) = Index
Image1(Index).Tag = "1"
PLcount = PLcount + 1
If PLcount = 2 Then PlControl
End Sub

Private Sub PlControl()
Image1(PlCor(0)).Tag = ""
Image1(PlCor(1)).Tag = ""
If PlPic(0) = PlPic(1) Then
Label1.Caption = "They match !"
Timer1.Enabled = True
Else
Label1.Caption = "They don't match..."
Timer2.Enabled = True
End If
End Sub

Private Sub Label5_Click()
Form1.Enabled = False
Form2.Show
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuStart_Click()
Form1.Enabled = False
Form2.Show
End Sub

Private Sub Timer1_Timer() 'player match
Timer1.Enabled = False
Image1(PlCor(0)).Visible = False
Image1(PlCor(1)).Visible = False
PlPnt = PlPnt + 10
PlCards = PlCards + 1
Label7(0).Caption = Format(PlCards, "00")
Label3.Caption = Format(PlPnt, "000")
If PlPnt + CpPnt = 300 Then
Image2.Visible = True
Label5.Visible = True
    If PlPnt = CpPnt Then
    Label1.Caption = "You are as powerfull as the computer !"
    End If
    If PlPnt < CpPnt Then
    Label1.Caption = "You lost the game !"
    End If
    If PlPnt > CpPnt Then
    Label1.Caption = "Your victory !"
    End If
Else
Label1.Caption = "Player, pick two cards"
PLcount = 0
End If
End Sub

Private Sub Timer2_Timer() 'player don't match
Timer2.Enabled = False
Image1(PlCor(0)).Picture = Form2.Image1(pp).Picture
Image1(PlCor(1)).Picture = Form2.Image1(pp).Picture
Label3.Caption = Format(PlPnt, "000")
Label1.Caption = "Computer, pick two cards"
Screen.MousePointer = 11: DoEvents
Timer3.Enabled = True
End Sub

Private Sub Timer3_Timer() 'computer picks 2 cards
Timer3.Enabled = False
Again0:
Randomize
CpPick = Int(Rnd * 60)
If Image1(CpPick).Visible = False Then GoTo Again0
If Image1(CpPick).Tag = "1" Then GoTo Again0
Image1(CpPick).Tag = "1"
Image1(CpPick).Picture = ImageList1.ListImages(Pic1(CpPick)).Picture
CpCor(0) = CpPick
CpPic(0) = Pic1(CpPick)
DoEvents

For xx = 0 To Cp 'number of second picks
Again1:
Randomize
CpPick = Int(Rnd * 60)
If Image1(CpPick).Visible = False Then GoTo Again1
If Image1(CpPick).Tag = "1" Then GoTo Again1
Image1(CpPick).Tag = "1"
CpCor(1) = CpPick
CpPic(1) = Pic1(CpPick)
    If CpPic(0) = CpPic(1) Then GoTo Further
Image1(CpPick).Tag = ""
Image1(CpPick) = Form2.Image1(pp).Picture
Next xx
Further:
Image1(CpPick).Picture = ImageList1.ListImages(Pic1(CpPick)).Picture
Image1(CpCor(0)).Tag = ""
Image1(CpCor(1)).Tag = ""
Screen.MousePointer = 1
CpControl
End Sub

Private Sub CpControl()
If CpPic(0) = CpPic(1) Then
Label1.Caption = "They match !"
Timer4.Enabled = True
Else
Label1.Caption = "They don't match..."
Timer5.Enabled = True
End If
End Sub

Private Sub Timer4_Timer() 'computer match
Timer4.Enabled = False
Image1(CpCor(0)).Visible = False
Image1(CpCor(1)).Visible = False
CpPnt = CpPnt + 10
CpCards = CpCards + 1
Label7(1).Caption = Format(CpCards, "00")
Label4.Caption = Format(CpPnt, "000")
If PlPnt + CpPnt = 300 Then
Image2.Visible = True
Label5.Visible = True
    If PlPnt = CpPnt Then
    Label1.Caption = "You are as powerfull as the player !"
    End If
    If PlPnt < CpPnt Then
    Label1.Caption = "The computer did win the game !"
    End If
    If PlPnt > CpPnt Then
    Label1.Caption = "The computer lost the game !"
    End If
Else
Label1.Caption = "Computer, pick two cards"
Timer3.Enabled = True
End If
End Sub

Private Sub Timer5_Timer() 'computer no match
Timer5.Enabled = False
Image1(CpCor(0)).Picture = Form2.Image1(pp).Picture
Image1(CpCor(1)).Picture = Form2.Image1(pp).Picture
Label4.Caption = Format(CpPnt, "000")
PLcount = 0
Label1.Caption = "Player, pick two cards"
End Sub
