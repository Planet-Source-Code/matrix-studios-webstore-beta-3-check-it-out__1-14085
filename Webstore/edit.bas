Attribute VB_Name = "Edit"
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_UNDO = &H304

Public Sub Copy()
Clipboard.SetText frmMenu.Text1.SelText
End Sub
Public Sub Copy2()
Clipboard.SetText frmMenu.Text1.SelText
End Sub


Public Sub Cut2()
Clipboard.SetText (frmMenu.Text1.SelText = "")
End Sub

Public Sub Paste()
frmMenu.Text1.SelText = Clipboard.GetText(1)
End Sub
Public Sub Paste2()
frmMenu.Text1.SelText = Clipboard.GetText(1)
End Sub

Public Sub Undo()
Call SendMessage(frmMenu.Text1.hwnd, WM_UNDO, 0&, 0&)
End Sub


Public Sub Undo2()
Call SendMessage(frmMenu.Text1.hwnd, WM_UNDO, 0&, 0&)
End Sub

