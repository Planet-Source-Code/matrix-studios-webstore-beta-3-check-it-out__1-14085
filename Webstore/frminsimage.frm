VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frminsimage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Image"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5865
   Icon            =   "frminsimage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check2 
      Caption         =   "Make image responce write"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   2760
      Width           =   2400
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   2205
      Left            =   4440
      ScaleHeight     =   2145
      ScaleWidth      =   1275
      TabIndex        =   20
      Top             =   120
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   2700
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Image"
      Height          =   1095
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   4335
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2400
         TabIndex        =   15
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   840
         TabIndex        =   14
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         Height          =   255
         Left            =   3855
         TabIndex        =   12
         Top             =   255
         Width           =   345
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Height"
         Height          =   255
         Left            =   1800
         TabIndex        =   17
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Width"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   495
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   1800
      Width           =   3225
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Insert"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   2760
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Link"
      Height          =   1500
      Left            =   0
      TabIndex        =   3
      Top             =   1155
      Width           =   4335
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   1486
         TabIndex        =   19
         Top             =   1065
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         OrigLeft        =   1486
         OrigTop         =   1065
         OrigRight       =   1726
         OrigBottom      =   1350
         Max             =   99
         Enabled         =   -1  'True
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   960
         TabIndex        =   18
         Text            =   "0"
         Top             =   1065
         Width           =   525
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Make into link"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frminsimage.frx":324A
         Left            =   3120
         List            =   "frminsimage.frx":3257
         TabIndex        =   5
         Text            =   "http://"
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   2880
         TabIndex        =   4
         Text            =   "_Self"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Link type"
         Height          =   255
         Left            =   2190
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Target"
         Height          =   255
         Left            =   2235
         TabIndex        =   8
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Actual link"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Border"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   855
      End
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Preview"
      Height          =   255
      Left            =   4800
      TabIndex        =   22
      Top             =   2400
      Width           =   735
   End
End
Attribute VB_Name = "frminsimage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type BITMAP 'standard WIN API32 structure
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
    End Type
    Private m_Picture As Picture 'the VB object Picture (not PictureBox)
    Private m_bm As BITMAP
    'Aliased since GetObject() is also a VB
    '     OLE function


Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" ( _
    ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long


Public Function ImageReadOK(Filename As String) As Boolean
    On Error Resume Next
    Set m_Picture = LoadPicture(Filename)


    If Err Then
        ImageReadOK = False
        Exit Function
    End If
    'm_bm is filled here with the BITMAP str
    '     ucture...

    ImageReadOK = (GetObjectAPI(m_Picture.Handle, Len(m_bm), m_bm) = Len(m_bm))

End Function
'---------------------------------------
'     ---------------------


Public Property Get WidthPixels() As Long
    WidthPixels = m_bm.bmWidth
End Property


Public Property Get HeightPixels() As Long
    HeightPixels = m_bm.bmHeight
End Property


Public Property Get WidthHiMetric() As Long
    WidthHiMetric = m_Picture.Width
End Property


Public Property Get HeightHiMetric() As Long
    HeightHiMetric = m_Picture.Height
End Property




Private Sub Command1_Click()
Dim Image As String
Dim Link As String
If Check2.Value = False Then
Image = "<IMG SRC=""" & Text3.Text & """ border=""" & Text2.Text & """ width=""" & Text4.Text & """ height=""" & Text5.Text & """>"
Link = "<a href=""" & Combo1.Text & Text1.Text & """ target=""" & Combo2.Text & """ >" & Image & "</a>"
Else
If Check1.Value = False Then
Image = "Responce.write" & Chr(34) & "<IMG SRC=""" & Chr(34) & Text3.Text & Chr(34) & """ border=""" & Chr(34) & Text2.Text & Chr(34) & """ width=""" & Chr(34) & Text4.Text & Chr(34) & """ height=""" & Chr(34) & Text5.Text & Chr(34) & """>" & Chr(34)
Else
Image = "<IMG SRC=""" & Chr(34) & Text3.Text & Chr(34) & """ border=""" & Chr(34) & Text2.Text & Chr(34) & """ width=""" & Chr(34) & Text4.Text & Chr(34) & """ height=""" & Chr(34) & Text5.Text & Chr(34) & """>"
Link = "Response.Write" & Chr(34) & "<a href=""" & Chr(34) & Combo1.Text & Text1.Text & Chr(34) & """ target=""" & Chr(34) & Combo2.Text & Chr(34) & """ >" & Image & "</a>" & Chr(34)
End If
End If
If Check1.Value = False Then
frmMenu.Text1.SelText = Image
Else
frmMenu.Text1.SelText = Link
End If
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
    With CommonDialog1
        .DialogTitle = "Open"
        .CancelError = False
        .Filter = "All Images(*.jpg *.gif)|*.jpg;*.gif|Jpeg(*.jpg)|*.jpg|Gif(*.gif)|*.gif|All Files(*.*)|*.*"
        .DefaultExt = "JPEG"
        .ShowOpen
        If Len(.Filename) = 0 Then
            Exit Sub
        End If
        sFile = .Filename
        Text3.Text = sFile
       Picture1.Picture = LoadPicture(CommonDialog1.Filename)
       ImageReadOK (CommonDialog1.Filename)
        Text4.Text = WidthPixels
        Text5.Text = HeightPixels
        End With
End Sub

