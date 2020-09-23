VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl AppDoc 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin MSComDlg.CommonDialog CmnDlg 
      Left            =   3960
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   0
      Picture         =   "AppDoc.ctx":0000
      Top             =   0
      Width           =   420
   End
End
Attribute VB_Name = "AppDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'AppDoc - VB5 control demo
'Copyright (c) 1997 SoftCircuits
'Redistributed by Permission.
'
'This Visual Basic 5.0 example program demonstrates how you can use
'VB5 controls to encapsulate the logic for keeping track of File New,
'Open, Save and Save As commands. Simply call the appropriate methods
'to implement each of these commands.
'
'The control saves you time by keeping track of if the current file
'has been modified and if it has yet been named. All you need to do
'is implement the code that actually initializes, loads and saves your
'data. This is done via the NewFile, LoadFile and SaveFile methods
'that the control calls at the appropriate times.
'
'The control is not visible at run time. A sample program is provided
'to demonstrate use of the control.
'
'This program may be distributed on the condition that it is
'distributed in full and unchanged, and that no fee is charged for
'such distribution with the exception of reasonable shipping and media
'charged. In addition, the code in this program may be incorporated
'into your own programs and the resulting programs may be distributed
'without payment of royalties.
'
'This example program was provided by:
' SoftCircuits Programming
' http://www.softcircuits.com
' P.O. Box 16262
' Irvine, CA 92623
Option Explicit

Const DEF_FILEFILTER = "HTML Files (*.html)|*.html|HTM Files (*.html)|*.html|iWeb Files (*.iwd)|*.iwd|ASP Files (*.asp)|*.asp|All Files (*.*)|*.*|"
Const DEF_DEFEXTENSION = "100000kb"

Event NewFile(bSuccess As Boolean)
Event LoadFile(sFileName As String, bSuccess As Boolean)
Event SaveFile(sFileName As String, bSuccess As Boolean)

'Indicates supported file types
Private m_sFileFilter As String
'Default extension appended to files saved with no extension
Private m_sDefExtension As String
'File name and title
Private m_sFileName As String
Private m_sFileTitle As String
'Public property to get/set modified status
Private m_bModified As Boolean


Public Property Let FileFilter(sFileFilter As String)
Attribute FileFilter.VB_Description = "Returns/sets the file filters for common file dialog"
    m_sFileFilter = sFileFilter
    PropertyChanged "FileFilter"
End Property

Public Property Get FileFilter() As String
    FileFilter = m_sFileFilter
End Property

Public Property Let DefExtension(sDefExtension As String)
Attribute DefExtension.VB_Description = "Returns/sets the default extension appended to files with no extension during Save As"
    m_sDefExtension = sDefExtension
    PropertyChanged "DefExtension"
End Property

Public Property Get DefExtension() As String
    DefExtension = m_sDefExtension
End Property

'Returns the current filename
Public Function GetFileTitle() As String
    GetFileTitle = m_sFileTitle
End Function

'Returns the current filename (with path)
Public Function GetFilename() As String
    GetFilename = m_sFileName
End Function

'Sets the "dirty" flag
Public Sub SetModified(Optional bModified As Boolean = True)
    m_bModified = bModified
End Sub

'Returns the "dirty" flag
Public Function GetModified() As Boolean
    GetModified = m_bModified
End Function

'Creates a new file
'Call to implement the New command
Public Function FileNew() As Boolean
    If FileSaveIfModified() Then
        If DoFileNew Then
            m_sFileTitle = "Untitled"
            m_sFileName = ""
            m_bModified = False
            FileNew = True
        End If
    End If
End Function

'Lets the user select and open a new file
'Call to implement the Open command
Public Function FileOpen() As Boolean
    If FileSaveIfModified() Then
        CmnDlg.Filename = ""
        CmnDlg.Filter = m_sFileFilter
        CmnDlg.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames 'Or cdlOFNHelpButton
        CmnDlg.CancelError = True
        On Error GoTo FileOpenErr
        CmnDlg.ShowOpen
        If DoFileLoad(CmnDlg.Filename) Then
            m_sFileTitle = CmnDlg.FileTitle
            m_sFileName = CmnDlg.Filename
            m_bModified = False
            FileOpen = True
        End If
    End If
EndFileOpen:
    Exit Function
FileOpenErr:
    If Err <> cdlCancel Then
        MsgBox "Error opening file : " & Err.Description
    End If
    Resume EndFileOpen
End Function

'Lets the user save the current file
'Call to implement the Save command
Public Function FileSave() As Boolean
    If m_sFileName = "" Then
        FileSave = FileSaveAs()
    Else
        If DoFileSave(m_sFileName) Then
            m_bModified = False
            FileSave = True
        End If
    End If
End Function

'Lets the user save the current file with a specified filename
'Call to implement the Save As command
Public Function FileSaveAs() As Boolean
    CmnDlg.Filename = m_sFileName
    CmnDlg.Filter = m_sFileFilter
    CmnDlg.DefaultExt = m_sDefExtension
    CmnDlg.Flags = cdlOFNOverwritePrompt Or cdlOFNHideReadOnly Or cdlOFNLongNames 'Or cdlOFNHelpButton
    CmnDlg.CancelError = True
    On Error GoTo FileSaveAsErr
    CmnDlg.ShowSave
    If DoFileSave(CmnDlg.Filename) Then
        m_sFileTitle = CmnDlg.FileTitle
        m_sFileName = CmnDlg.Filename
        m_bModified = False
        FileSaveAs = True
    End If
EndFileSaveAs:
    Exit Function
FileSaveAsErr:
    If Err <> cdlCancel Then
    End If
    Resume EndFileSaveAs
End Function

'Allows the user to save the current file if it is modified
'Call just before closing a file
Public Function FileSaveIfModified() As Boolean
    Dim bResult As Boolean, I As Integer

    bResult = True  'Assume success for now
    If m_bModified Then
        I = MsgBox("Save changes to '" & m_sFileTitle & "'?", vbYesNoCancel)
        If I = vbYes Then
            bResult = FileSave()
        ElseIf I = vbCancel Then
            bResult = False
        End If
    End If
    FileSaveIfModified = bResult
End Function

'Perform file new
Private Function DoFileNew() As Boolean
    Dim bSuccess As Boolean
    Screen.MousePointer = vbHourglass
    On Error Resume Next
    bSuccess = True 'Return success if no event
    RaiseEvent NewFile(bSuccess)
    Screen.MousePointer = vbDefault
    DoFileNew = bSuccess
    Screen.MousePointer = vbDefault
End Function

'Perform file load
Private Function DoFileLoad(sFileName As String) As Boolean
    Dim bSuccess As Boolean
    Screen.MousePointer = vbHourglass
    bSuccess = True 'Return success if no event
    RaiseEvent LoadFile(sFileName, bSuccess)
    DoFileLoad = bSuccess
    Screen.MousePointer = vbDefault
End Function

'Perform file save
Private Function DoFileSave(sFileName As String) As Boolean
    Dim bSuccess As Boolean
    Screen.MousePointer = vbHourglass
    On Error Resume Next
    bSuccess = True 'Return success if no event
    RaiseEvent SaveFile(sFileName, bSuccess)
    DoFileSave = bSuccess
    Screen.MousePointer = vbDefault
End Function

'Initialize control properties on first use
Private Sub UserControl_InitProperties()
    m_sFileFilter = DEF_FILEFILTER
    m_sDefExtension = DEF_DEFEXTENSION
End Sub

'Load control properties
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error GoTo ReadPropErr
    m_sFileFilter = PropBag.ReadProperty("FileFilter", DEF_FILEFILTER)
    m_sDefExtension = PropBag.ReadProperty("DefExtension", DEF_DEFEXTENSION)
EndReadProp:
    Exit Sub
ReadPropErr:
    'Use default property settings
    UserControl_InitProperties
    Resume EndReadProp
End Sub

'Save control properties
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "FileFilter", m_sFileFilter, DEF_FILEFILTER
    PropBag.WriteProperty "DefExtension", m_sDefExtension, DEF_DEFEXTENSION
End Sub

'Restrict design-time size to image size
Private Sub UserControl_Resize()
    Size Image1.Width, Image1.Height
End Sub
