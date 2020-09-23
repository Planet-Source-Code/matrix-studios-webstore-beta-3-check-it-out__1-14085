VERSION 5.00
Begin VB.Form frmaddlink 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Link"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   Icon            =   "frmaddlink.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   4500
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Make link responce write"
      Height          =   270
      Left            =   120
      TabIndex        =   11
      Top             =   1695
      Width           =   2205
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Insert"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   840
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Link"
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   4335
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmaddlink.frx":324A
         Left            =   2880
         List            =   "frmaddlink.frx":325D
         TabIndex        =   6
         Text            =   "_Default"
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmaddlink.frx":3284
         Left            =   960
         List            =   "frmaddlink.frx":3291
         TabIndex        =   5
         Text            =   "http://"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   1200
         Width           =   3255
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Descrption"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Actual link"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Target"
         Height          =   255
         Left            =   2280
         TabIndex        =   8
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Link type"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmaddlink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Dim Link
If Check1.Value = False Then
Link = "<a href=""" & Combo1.Text & Text1.Text & """ target=""" & Combo2.Text & """ >" & Text2.Text & "</a>"
Else
Link = "Responce.write(" & Chr(34) & "<a href=""" & Chr(34) & Combo1.Text & Text1.Text & Chr(34) & """ target=""" & Chr(34) & Combo2.Text & Chr(34) & """ >" & Text2.Text & "</a>)" & Chr(34)
End If
frmMenu.Text1.SelText = Link

Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
