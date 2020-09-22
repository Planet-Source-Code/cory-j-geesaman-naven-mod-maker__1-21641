Attribute VB_Name = "modMisc"
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Public SelectDD As String

Public Sub DragForm(Frm As Form)
If Frm.WindowState <> vbNormal Then Exit Sub
On Local Error Resume Next
Call ReleaseCapture
Call SendMessage(Frm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub
