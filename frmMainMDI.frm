VERSION 5.00
Begin VB.MDIForm frmMainMDI 
   BackColor       =   &H00929A93&
   Caption         =   "MDIForm1"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   9615
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMainMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
frmMain.Show
End Sub
