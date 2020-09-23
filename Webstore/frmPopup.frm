VERSION 5.00
Begin VB.Form frmPopup 
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   Icon            =   "frmPopup.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuForm1 
      Caption         =   "Form1"
      Index           =   0
      Begin VB.Menu mnuNewFile 
         Caption         =   "New File"
         Index           =   0
      End
      Begin VB.Menu mnu01 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuOpenFile 
         Caption         =   "Open File"
         Index           =   0
      End
      Begin VB.Menu mnuSaveFile 
         Caption         =   "Save File"
         Index           =   0
      End
      Begin VB.Menu mnu02 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Index           =   0
      End
   End
   Begin VB.Menu mnuForm2 
      Caption         =   "Form2"
      Index           =   0
      Begin VB.Menu mnuSource 
         Caption         =   "Source"
         Index           =   0
         Begin VB.Menu mnuAdd 
            Caption         =   "Add"
            Index           =   0
            Begin VB.Menu mnuCheckbox 
               Caption         =   "Checkbox"
               Index           =   0
            End
            Begin VB.Menu frmAddLink1 
               Caption         =   "Link"
               Index           =   0
            End
         End
         Begin VB.Menu mnuInsert 
            Caption         =   "Insert"
            Index           =   0
            Begin VB.Menu frmInsertImage 
               Caption         =   "Image"
               Index           =   0
            End
            Begin VB.Menu frmInsertLine 
               Caption         =   "Line"
               Index           =   0
            End
         End
         Begin VB.Menu mnu 
            Caption         =   "New"
            Index           =   0
            Begin VB.Menu mnuNewRow 
               Caption         =   "Row"
               Index           =   0
            End
         End
         Begin VB.Menu mnuText 
            Caption         =   "Text"
            Index           =   0
            Begin VB.Menu frmSquare 
               Caption         =   "Square"
               Index           =   0
            End
         End
         Begin VB.Menu mnuSpace 
            Caption         =   "Space"
            Index           =   0
            Begin VB.Menu mnuAdd1 
               Caption         =   "Add"
            End
            Begin VB.Menu mnuNew 
               Caption         =   "New"
               Index           =   0
            End
         End
         Begin VB.Menu mnuIncreaseIndent 
            Caption         =   "Increase Indent"
            Index           =   0
            Begin VB.Menu mnu1 
               Caption         =   "1"
               Index           =   0
            End
            Begin VB.Menu mnu2 
               Caption         =   "2"
               Index           =   0
            End
            Begin VB.Menu mnu3 
               Caption         =   "3"
               Index           =   0
            End
            Begin VB.Menu mnu4 
               Caption         =   "4"
               Index           =   0
            End
            Begin VB.Menu mnu5 
               Caption         =   "5"
               Index           =   0
            End
            Begin VB.Menu mnu6 
               Caption         =   "6"
               Index           =   0
            End
            Begin VB.Menu mnu7 
               Caption         =   "7"
               Index           =   0
            End
            Begin VB.Menu mnu8 
               Caption         =   "8"
               Index           =   0
            End
         End
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit"
         Index           =   0
         Begin VB.Menu mnuDeleteText 
            Caption         =   "Delete Text"
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu mnuUndo 
            Caption         =   "Undo"
            Index           =   0
         End
      End
      Begin VB.Menu mnuRequire 
         Caption         =   "Require"
         Index           =   0
         Begin VB.Menu frmBackround 
            Caption         =   "Backround"
            Begin VB.Menu frmAddBC 
               Caption         =   "Colour"
               Index           =   0
            End
            Begin VB.Menu frmInsertBP 
               Caption         =   "Picture"
               Index           =   0
            End
            Begin VB.Menu frmInsertBS 
               Caption         =   "Sound"
               Index           =   0
            End
         End
         Begin VB.Menu mnuOther 
            Caption         =   "Other"
            Index           =   0
            Begin VB.Menu frmAddHeadline 
               Caption         =   "Headline"
               Index           =   0
            End
         End
      End
      Begin VB.Menu mnuVarious 
         Caption         =   "Various"
         Index           =   0
         Begin VB.Menu mnuDots 
            Caption         =   "Dots"
            Index           =   0
         End
         Begin VB.Menu mnuNumbering 
            Caption         =   "Numbering"
            Index           =   0
         End
      End
   End
End
Attribute VB_Name = "frmPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub frmAddBC_Click(Index As Integer)
frmBColour.Show
End Sub

Private Sub frmAddHeadline_Click(Index As Integer)
frmHeadline.Show
End Sub

Private Sub frmAddLink_Click(Index As Integer)

End Sub

Private Sub frmAddLink1_Click(Index As Integer)
frmaddlink.Show
End Sub

Private Sub frmInsertImage_Click(Index As Integer)
frminsimage.Show
End Sub

Private Sub frmInsertLine_Click(Index As Integer)
frmMenu.Text1.SelText = "<hr>"
End Sub

Private Sub frmSpace_Click(Index As Integer)

End Sub

Private Sub frmSquare_Click(Index As Integer)
frmMenu.Text1.SelText = frmTemp.Text1(0).Text
End Sub

Private Sub mnu1_Click(Index As Integer)
frmMenu.Text1.SelText = frmIndent.Text1.Text
End Sub

Private Sub mnu2_Click(Index As Integer)
frmMenu.Text1.SelText = frmIndent.Text2.Text
End Sub

Private Sub mnu3_Click(Index As Integer)
frmMenu.Text1.SelText = frmIndent.Text3.Text
End Sub

Private Sub mnu4_Click(Index As Integer)
frmMenu.Text1.SelText = frmIndent.Text4.Text
End Sub

Private Sub mnu5_Click(Index As Integer)
frmMenu.Text1.SelText = frmIndent.Text5.Text
End Sub

Private Sub mnu6_Click(Index As Integer)
frmMenu.Text1.SelText = frmIndent.Text6.Text
End Sub

Private Sub mnu7_Click(Index As Integer)
frmMenu.Text1.SelText = frmIndent.Text7.Text
End Sub

Private Sub mnu8_Click(Index As Integer)
frmMenu.Text1.SelText = frmIndent.Text8.Text
End Sub

Private Sub mnuAdd1_Click()
frmMenu.Text1.SelText = "&nbsp;"
End Sub

Private Sub mnuCheckbox_Click(Index As Integer)
frmMenu.Text1.SelText = frmTemp.Text2.Text
End Sub

Private Sub mnuDots_Click(Index As Integer)
frmMenu.Text1.SelText = frmVarious.Text1.Text
End Sub

Private Sub mnuExit_Click(Index As Integer)
Dim X%
X% = MsgBox("Are you sure you want to quit?", vbYesNo, "Exit program")
If X% = vbYes Then End
Cancel = 1
Exit Sub
End Sub

Private Sub mnuNew_Click(Index As Integer)
frmMenu.Text1.SelText = "<p>&nbsp;</p>"
End Sub

Private Sub mnuNewFile_Click(Index As Integer)

Dim X%
X% = MsgBox("Are you sure you want to create a new file?", vbYesNo, "New File")
If X% = vbYes Then frmMenu.Text1.Text = frmTemp.Text1.Item(1).Text
Cancel = 1
Exit Sub

End Sub

Private Sub mnuNewRow_Click(Index As Integer)
frmMenu.Text1.SelText = "<br>"
End Sub

Private Sub mnuNumbering_Click(Index As Integer)
frmMenu.Text1.SelText = frmVarious.Text2.Text
End Sub

Private Sub mnuOpenFile_Click(Index As Integer)

Dim X%
X% = MsgBox("Are you sure you want to open another file?", vbYesNo, "Opening file")
If X% = vbYes Then frmMenu.AppDoc1.FileOpen
If X% = vbYes Then frmMenu.Text2.Text = AppDoc1.GetFilename
Cancel = 1
Exit Sub
End Sub

Private Sub mnuSaveFile_Click(Index As Integer)
frmMenu.AppDoc1.FileFilter = "HTML Files (*.html)|*.html|HTM Files (*.html)|*.html|Webstore Files (*.web)|*.web|ASP Files (*.asp)|*.asp|Prodon Files (*.prd)|*.prd|All Files (*.*)|*.*|"
frmMenu.AppDoc1.FileSave
frmMenu.Text2.Text = AppDoc1.GetFilename
End Sub

Private Sub mnuUndo_Click(Index As Integer)
Undo
End Sub
