VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMenu 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Webstore"
   ClientHeight    =   6450
   ClientLeft      =   2775
   ClientTop       =   3510
   ClientWidth     =   6720
   ForeColor       =   &H8000000C&
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   6720
   Begin VB.CheckBox Check2 
      BackColor       =   &H80000007&
      Caption         =   "Table"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   6000
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2040
      TabIndex        =   6
      Top             =   6000
      Width           =   4455
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000007&
      Caption         =   "Stuff"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   6000
      Value           =   1  'Checked
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6600
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Project1.AppDoc AppDoc1 
      Left            =   6600
      Top             =   5880
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000005&
      Height          =   5175
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frmMenu.frx":324A
      Top             =   600
      Width           =   6015
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
      Left            =   6240
      TabIndex        =   4
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
      Left            =   6480
      TabIndex        =   3
      ToolTipText     =   "Exit Webstore"
      Top             =   0
      Width           =   135
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      X1              =   6480
      X2              =   6480
      Y1              =   480
      Y2              =   5880
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      X1              =   240
      X2              =   6480
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      X1              =   240
      X2              =   6480
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   240
      X2              =   240
      Y1              =   5880
      Y2              =   480
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
      Width           =   9375
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
If Check2.Value = 1 Then frmStuff.Show
If Check2.Value = 0 Then frmStuff.Hide
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then frmTable.Show
If Check2.Value = 0 Then frmTable.Hide
End Sub

Private Sub Form_Click()
frmMenu.PopupMenu frmPopup.mnuForm1.Item(0)
End Sub

Private Sub Form_Load()
AppDoc1.FileSave
Text2.Text = AppDoc1.GetFilename
frmMenu.Show
frmTable.Show
frmStuff.Show
If Check1.Value = 1 Then frmStuff.Show
If Check1.Value = 0 Then frmStuff.Hide
If Check2.Value = 1 Then frmTable.Show
If Check2.Value = 0 Then frmTable.Hide
End Sub

Private Sub Label2_Click()
MsgBox "This is BETA #3 Version 1.00 of this new great web creator...", vbInformation
End Sub

Private Sub Label3_Click()

Dim X%
X% = MsgBox("Are you sure you want to quit?", vbYesNo, "Exit program")
If X% = vbYes Then End
Cancel = 1
Exit Sub

End Sub

Private Sub Label4_Click()
frmMenu.WindowState = 1
End Sub

Private Sub Text1_Change()
If frmStuff.Check1.Value = 1 Then AppDoc1.FileSave
End Sub

Private Sub AppDoc1_NewFile(bSuccess As Boolean)
    'Reset current data
    On Error GoTo NewFileErr
    Text1 = ""
NewFileEnd:
    Exit Sub
NewFileErr:
    MsgBox "Error creating new file : " & Err.Description
    bSuccess = False
    Resume NewFileEnd
End Sub

Private Sub AppDoc1_LoadFile(sFileName As String, bSuccess As Boolean)
    'Load file
    On Error GoTo LoadFileErr
    Open sFileName For Input As #1
    Text1 = Input$(LOF(1), 1)
    AppDoc1.FileFilter = "Text Files(*.txt)|*.txt|"
LoadFileEnd:
    On Error Resume Next
    Close #1
    Exit Sub
LoadFileErr:
    MsgBox "Error loading " & sFileName & " : " & Err.Description
    bSuccess = False
    Resume LoadFileErr
End Sub

Private Sub AppDoc1_SaveFile(sFileName As String, bSuccess As Boolean)
    'Save file
    On Error GoTo SaveFileErr
    Open sFileName For Output As #1
    Print #1, Text1;
SaveFileEnd:
    On Error Resume Next
    Close #1
    Exit Sub
SaveFileErr:
    MsgBox "Error saving " & sFileName & " : " & Err.Description
    bSuccess = False
    Resume SaveFileEnd
End Sub

Private Sub Text1_DblClick()
frmMenu.PopupMenu frmPopup.mnuForm2.Item(0)
End Sub
