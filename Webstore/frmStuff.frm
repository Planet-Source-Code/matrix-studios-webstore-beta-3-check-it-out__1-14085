VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmStuff 
   BorderStyle     =   0  'None
   Caption         =   "Webstore Stuff"
   ClientHeight    =   7710
   ClientLeft      =   9825
   ClientTop       =   2010
   ClientWidth     =   3645
   Icon            =   "frmStuff.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   3645
   Begin TabDlg.SSTab SSTab1 
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   13150
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      BackColor       =   255
      ForeColor       =   255
      TabCaption(0)   =   "Coding"
      TabPicture(0)   =   "frmStuff.frx":324A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(1)=   "Label2(0)"
      Tab(0).Control(2)=   "Label9"
      Tab(0).Control(3)=   "List2"
      Tab(0).Control(4)=   "List4"
      Tab(0).Control(5)=   "List3"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Options"
      TabPicture(1)   =   "frmStuff.frx":3266
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label4(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label8"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label10"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Check1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Text1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Command1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Command2"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Command3"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Command4"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Command5"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Command6"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Command7"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Command8"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Command9"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Command10"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).ControlCount=   15
      TabCaption(2)   =   "Info"
      TabPicture(2)   =   "frmStuff.frx":3282
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label5"
      Tab(2).Control(1)=   "Label6"
      Tab(2).Control(2)=   "Label7"
      Tab(2).ControlCount=   3
      Begin VB.CommandButton Command10 
         Caption         =   "Insert Image"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Add Headline"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   1920
         Width           =   1335
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Add Link"
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Seek for HTML"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Exit Program"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   6000
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Open File"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   5640
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Save File As"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   5280
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Save File"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   4920
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "New File"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   4560
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Flash Player"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   9
         Text            =   "You!"
         Top             =   6840
         Width           =   3135
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Autosave"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H000040C0&
         Height          =   1785
         ItemData        =   "frmStuff.frx":329E
         Left            =   -74880
         List            =   "frmStuff.frx":3446
         TabIndex        =   5
         Top             =   5400
         Width           =   3375
      End
      Begin VB.ListBox List4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H000040C0&
         Height          =   1785
         ItemData        =   "frmStuff.frx":3AD5
         Left            =   -74880
         List            =   "frmStuff.frx":3D07
         TabIndex        =   4
         Top             =   3120
         Width           =   3375
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H000040C0&
         Height          =   1785
         ItemData        =   "frmStuff.frx":468B
         Left            =   -74880
         List            =   "frmStuff.frx":47D3
         TabIndex        =   1
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label Label10 
         Caption         =   "BETA #3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   25
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "[ ASP ]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73680
         TabIndex        =   24
         Top             =   5040
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Webstore is still in BETA. Full version are soon avaible at Matrix Studios Website. Visit!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         TabIndex        =   19
         Top             =   3120
         Width           =   3135
      End
      Begin VB.Label Label7 
         Caption         =   $"frmStuff.frx":4E77
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   -74880
         TabIndex        =   13
         Top             =   5040
         Width           =   3375
      End
      Begin VB.Label Label6 
         Caption         =   $"frmStuff.frx":5011
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   -74880
         TabIndex        =   12
         Top             =   2280
         Width           =   3255
      End
      Begin VB.Label Label5 
         Caption         =   $"frmStuff.frx":51EA
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   -74880
         TabIndex        =   11
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label4 
         Caption         =   "Registerd to:"
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
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   6480
         Width           =   3015
      End
      Begin VB.Label Label3 
         Caption         =   "[ ASP ]"
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
         Index           =   0
         Left            =   -73560
         TabIndex        =   6
         Top             =   5040
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "[ IHTML ]"
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
         Index           =   0
         Left            =   -73800
         TabIndex        =   3
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "[ HTML ]"
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
         Index           =   0
         Left            =   -73680
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000FF&
      Caption         =   "_"
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
      Index           =   1
      Left            =   3240
      TabIndex        =   23
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Caption         =   "X"
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
      Index           =   1
      Left            =   3480
      TabIndex        =   22
      ToolTipText     =   "Exit Webstore"
      Top             =   0
      Width           =   135
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
      Index           =   1
      Left            =   1320
      TabIndex        =   21
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   9375
   End
End
Attribute VB_Name = "frmStuff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmFlash.Show
End Sub

Private Sub Command10_Click()
frminsimage.Show
End Sub

Private Sub Command2_Click()

Dim X%
X% = MsgBox("Are you sure you want to create a new file?", vbYesNo, "New File")
If X% = vbYes Then frmMenu.Text1.Text = frmTemp.Text1.Item(1).Text
Cancel = 1
Exit Sub

End Sub

Private Sub Command3_Click()
frmMenu.AppDoc1.FileFilter = "HTML Files (*.html)|*.html|HTM Files (*.html)|*.html|Webstore Files (*.web)|*.web|ASP Files (*.asp)|*.asp|Prodon Files (*.prd)|*.prd|All Files (*.*)|*.*|"
frmMenu.AppDoc1.FileSave
End Sub

Private Sub Command4_Click()
frmMenu.AppDoc1.FileFilter = "HTML Files (*.html)|*.html|HTM Files (*.html)|*.html|Webstore Files (*.web)|*.web|ASP Files (*.asp)|*.asp|Prodon Files (*.prd)|*.prd|All Files (*.*)|*.*|"
frmMenu.AppDoc1.FileSaveAs
End Sub

Private Sub Command5_Click()

Dim X%
X% = MsgBox("Are you sure you want to open another file?", vbYesNo, "Opening file")
If X% = vbYes Then frmMenu.AppDoc1.FileFilter = "HTML Files (*.html)|*.html|HTM Files (*.html)|*.html|Webstore Files (*.web)|*.web|ASP Files (*.asp)|*.asp|Prodon Files (*.prd)|*.prd|All Files (*.*)|*.*|"
If X% = vbYes Then frmMenu.AppDoc1.FileOpen
Cancel = 1
Exit Sub
End Sub

Private Sub Command6_Click()
End
End Sub

Private Sub Command7_Click()
WinSeek.Show
End Sub

Private Sub Command9_Click()
frmHeadline.Show
End Sub

Private Sub Label3_Click(Index As Integer)
Me.Hide
frmMenu.Check1.Value = 0
End Sub

Private Sub Label4_Click(Index As Integer)
frmStuff.WindowState = 1
End Sub

Private Sub List2_Click()
frmMenu.Text1.SelText = List2.Text & " "
End Sub

Private Sub List3_Click()
frmMenu.Text1.SelText = List3.Text & " "
End Sub

Private Sub List4_Click()
frmMenu.Text1.SelText = List4.Text & " "
End Sub

