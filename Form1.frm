VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00929A93&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   30
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   870
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   58
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_GotFocus()
frmMain.SetFocus
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbNormal Then
frmMain.Show
Else
frmMain.Hide
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
