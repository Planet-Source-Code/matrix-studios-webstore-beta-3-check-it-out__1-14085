VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTable 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Webstore Table"
   ClientHeight    =   1695
   ClientLeft      =   2775
   ClientTop       =   1500
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   6390
   Begin TabDlg.SSTab SSTab1 
      Height          =   1455
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   2566
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Tab"
      TabPicture(0)   =   "frmTable.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Command26"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Command36"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command34"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command10"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command31"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Fonts"
      TabPicture(1)   =   "frmTable.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command30"
      Tab(1).Control(1)=   "Command27"
      Tab(1).Control(2)=   "Command25"
      Tab(1).Control(3)=   "Command24"
      Tab(1).Control(4)=   "Command23"
      Tab(1).Control(5)=   "Command22"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Links"
      TabPicture(2)   =   "frmTable.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Text1"
      Tab(2).Control(1)=   "Command1"
      Tab(2).ControlCount=   2
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -71160
         TabIndex        =   15
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   -74760
         TabIndex        =   14
         Text            =   "www.matrixstudios-online.com"
         Top             =   600
         Width           =   3495
      End
      Begin VB.CommandButton Command22 
         Caption         =   "H1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -74760
         TabIndex        =   13
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton Command23 
         Caption         =   "H2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -74160
         TabIndex        =   12
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton Command24 
         Caption         =   "H3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -73560
         TabIndex        =   11
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton Command25 
         Caption         =   "H4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -72960
         TabIndex        =   10
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton Command27 
         Caption         =   "H5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -72360
         TabIndex        =   9
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton Command30 
         Caption         =   "H6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -71760
         TabIndex        =   8
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton Command31 
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1200
         TabIndex        =   7
         ToolTipText     =   "Italics"
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton Command10 
         Caption         =   "U"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   720
         TabIndex        =   6
         ToolTipText     =   "Underline"
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton Command34 
         Caption         =   "Tit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2280
         TabIndex        =   5
         ToolTipText     =   "Title"
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton Command36 
         Caption         =   "BR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1800
         TabIndex        =   4
         ToolTipText     =   "Break"
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton Command26 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "Bolds"
         Top             =   600
         Width           =   495
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "Webstore"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6495
   End
End
Attribute VB_Name = "frmTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command10_Click()
frmMenu.Text1.SelText = "<U> Insert Caption </U>"
End Sub

Private Sub Command22_Click()
frmMenu.Text1.SelText = "<H1> Insert Caption </H1>"
End Sub

Private Sub Command23_Click()
frmMenu.Text1.SelText = "<H2> Insert Caption </H2>"
End Sub

Private Sub Command24_Click()
frmMenu.Text1.SelText = "<H3> Insert Caption </H3>"
End Sub

Private Sub Command25_Click()
frmMenu.Text1.SelText = "<H4> Insert Caption </H4>"
End Sub

Private Sub Command26_Click()
frmMenu.Text1.SelText = "<B> Insert Caption</B>"
End Sub

Private Sub Command27_Click()
frmMenu.Text1.SelText = "<H5> Insert Caption </H5>"
End Sub

Private Sub Command31_Click()
frmMenu.Text1.SelText = "<I> Insert Caption </I>"
End Sub

Private Sub Command34_Click()
frmHeadline.Show
End Sub

Private Sub Command36_Click()
frmMenu.Text1.SelText = "<BR>"
End Sub

