VERSION 5.00
Begin VB.Form frmHeadline 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Headline..."
   ClientHeight    =   495
   ClientLeft      =   2955
   ClientTop       =   3405
   ClientWidth     =   4860
   Icon            =   "frmHeadline.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   495
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command5 
      Caption         =   "Insert Headline..."
      Height          =   285
      Left            =   2760
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   3
      Text            =   "</title>"
      Top             =   480
      Width           =   615
      Visible         =   0   'False
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "Matrix Studios"
      Top             =   120
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "<title>"
      Top             =   480
      Width           =   615
      Visible         =   0   'False
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "<title>Matrix Studios</title>"
      Top             =   120
      Width           =   6135
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmHeadline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)
            Clipboard.Clear
            Clipboard.SetText Text2.SelText
End Sub

Private Sub Command2_Click()
Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)
            Clipboard.Clear
            Clipboard.SetText Text3.SelText
End Sub

Private Sub Command3_Click()
Text4.SelStart = 0
Text4.SelLength = Len(Text4.Text)

            Clipboard.Clear
            Clipboard.SetText Text4.SelText
End Sub

Private Sub Command4_Click()
            frmMenu.Text1.SelText = Clipboard.GetText()
            frmMenu.Show
            Me.Hide
End Sub

Private Sub Command5_Click()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
            Clipboard.Clear
            Clipboard.SetText Text1.SelText
            
            frmMenu.Text1.SelText = Clipboard.GetText()
            frmMenu.Show
            Me.Hide
End Sub

Private Sub Command6_Click()
Text1.Text = ""
End Sub

Private Sub Form_Load()
Text1.Text = Text2.Text & Text3.Text & Text4.Text
End Sub

Private Sub Text1_Change()
Text1.Text = Text2.Text & Text3.Text & Text4.Text
End Sub

Private Sub Text3_Change()
Text1.Text = Text2.Text & Text3.Text & Text4.Text
End Sub
