VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "SWFLASH.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmFlash 
   Caption         =   "Flash Video Player"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "frmFlash.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   2160
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   2760
      Width           =   4695
      Visible         =   0   'False
      _ExtentX        =   8281
      _ExtentY        =   661
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   240
      Top             =   3120
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   3105
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4480
            MinWidth        =   4480
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3069
            MinWidth        =   3069
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog openfile 
      Left            =   120
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash flash 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      _cx             =   4202585
      _cy             =   4199622
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu open 
         Caption         =   "&Open"
      End
      Begin VB.Menu closee 
         Caption         =   "&Close"
      End
      Begin VB.Menu line8 
         Caption         =   "-"
      End
      Begin VB.Menu print 
         Caption         =   "&Print"
         Enabled         =   0   'False
      End
      Begin VB.Menu line 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu view 
      Caption         =   "&View"
      Begin VB.Menu full 
         Caption         =   "&Full Screen"
      End
      Begin VB.Menu min 
         Caption         =   "&Minimize"
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu zin 
         Caption         =   "&Zoom In"
      End
      Begin VB.Menu zout 
         Caption         =   "&Zoom Out"
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu qual 
         Caption         =   "Quality"
         Begin VB.Menu low 
            Caption         =   "&Low"
         End
         Begin VB.Menu med 
            Caption         =   "&Med"
         End
         Begin VB.Menu high 
            Caption         =   "&High"
         End
      End
   End
   Begin VB.Menu control 
      Caption         =   "&Control"
      Begin VB.Menu play 
         Caption         =   "&Play"
      End
      Begin VB.Menu stopp 
         Caption         =   "&Stop"
      End
      Begin VB.Menu rewind 
         Caption         =   "&Rewind"
      End
      Begin VB.Menu line4 
         Caption         =   "-"
      End
      Begin VB.Menu sf 
         Caption         =   "&Step Forward"
      End
      Begin VB.Menu sb 
         Caption         =   "&Step Back"
      End
      Begin VB.Menu gotoo 
         Caption         =   "&Goto Frame"
      End
      Begin VB.Menu line5 
         Caption         =   "-"
      End
      Begin VB.Menu loopp 
         Caption         =   "&Loop"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmFlash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub closee_Click()
On Error Resume Next
flash.Movie = ""
StatusBar1.Panels(1).Text = "Stoped"
bout.Enabled = True
Timer1.Enabled = False
Timer2.Enabled = False
full.Enabled = False
zin.Enabled = False
zout.Enabled = False
play.Enabled = False
stopp.Enabled = False
sf.Enabled = False
sb.Enabled = False
rewind.Enabled = False
closee.Enabled = False
qual.Enabled = False
gotoo.Enabled = False
min.Enabled = False
End Sub
Private Sub exit_Click()
Dim a, b As String
a = MsgBox("Are you sure you want to quit?", vbYesNo Or vbQuestion, "Quit?")
If a = vbYes Then Unload Me
If a = vbYes Then frmMenu.Show

Exit Sub

End Sub
Private Sub Form_Load()
On Error Resume Next
StatusBar1.Panels(1).Text = "Stoped"
full.Enabled = False
zin.Enabled = False
zout.Enabled = False
play.Enabled = False
stopp.Enabled = False
sf.Enabled = False
sb.Enabled = False
rewind.Enabled = False
closee.Enabled = False
qual.Enabled = False
gotoo.Enabled = False
min.Enabled = False
End Sub
Private Sub Form_Resize()
On Error Resume Next
ProgressBar1.Width = Form1.Width - 5
ProgressBar1.Top = Form1.Top - Form1.Top + 10
StatusBar1.Panels(1).Width = Form1.Width - 3000
StatusBar1.Panels(2).Width = Form1.Width - 600
flash.Width = Form1.Width
flash.Height = Form1.Height
End Sub

Private Sub full_Click()
On Error Resume Next
Form1.WindowState = vbMaximized


End Sub

Private Sub gotoo_Click()
On Error Resume Next
Dim aa, bb As String
aa = InputBox("What frame would you like to jump to?", "Frames")
If aa > flash.TotalFrames Then
MsgBox "Error you cannot jump farther than then" & Chr$(32) & flash.TotalFrames, vbCritical, "Error"
Else:
flash.GotoFrame (aa)
End If

End Sub

Private Sub high_Click()
flash.Quality2 = high

End Sub

Private Sub loopp_Click()
If loopp.Checked = True Then
loopp.Checked = False
Else: loopp.Checked = True
End If
End Sub

Private Sub low_Click()
flash.Quality2 = low
End Sub

Private Sub med_Click()
flash.Quality2 = med

End Sub

Private Sub min_Click()
Form1.WindowState = vbNormal
End Sub

Private Sub open_Click()
On Error Resume Next
flash.Movie = ""
openfile.Filter = "Swf Files (*.swf) | *.swf|All Files (*.*) | *.*"
openfile.DialogTitle = "Open swf file"
openfile.ShowOpen
If loopp.Checked = False Then
flash.Movie = openfile.Filename
ProgressBar1.Visible = True
Timer2.Enabled = True
StatusBar1.Panels(1).Text = "Playing" & Chr$(32) & openfile.FileTitle
flash.Loop = False
Timer1.Enabled = True
Else:
flash.Movie = ""
flash.Movie = openfile.Filename
ProgressBar1.Visible = True
Timer2.Enabled = True
StatusBar1.Panels(1).Text = "Playing" & Chr$(32) & openfile.FileTitle
flash.Loop = True
Timer1.Enabled = True
End If
End Sub
Private Sub play_Click()
On Error Resume Next
flash.play
StatusBar1.Panels(1).Text = "Playing"
End Sub
Private Sub rewind_Click()

On Error Resume Next
flash.rewind
End Sub
Private Sub sb_Click()
On Error Resume Next
flash.Back
End Sub

Private Sub sf_Click()
On Error Resume Next
flash.Forward
End Sub


Private Sub stopp_Click()
On Error Resume Next
flash.Stop
bout.Enabled = True

End Sub

Private Sub Timer1_Timer()
On Error Resume Next
StatusBar1.Panels(2).Text = "Frames:" & Chr$(32) & flash.FrameNum & "/" & flash.TotalFrames
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
ProgressBar1.Max = flash.TotalFrames
If ProgressBar1.Value = flash.TotalFrames Then
ProgressBar1.Visible = False
Timer2.Enabled = False
full.Enabled = True
full.Enabled = True
zin.Enabled = True
zout.Enabled = True
play.Enabled = True
stopp.Enabled = True
sf.Enabled = True
sb.Enabled = True
rewind.Enabled = True
closee.Enabled = True
qual.Enabled = True
gotoo.Enabled = True
min.Enabled = True
Else:
ProgressBar1.Value = ProgressBar1.Value + 1
End If

End Sub

Private Sub zin_Click()
On Error Resume Next
flash.Zoom (3)
End Sub

Private Sub zout_Click()
On Error Resume Next
flash.Movie = ""
End Sub
