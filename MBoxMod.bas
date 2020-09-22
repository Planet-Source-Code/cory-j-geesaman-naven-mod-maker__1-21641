Attribute VB_Name = "MBoxMod"
Public MBReturn%

Public Function Msbox(Message As Variant, Optional Title As Variant)
On Error Resume Next
Load MBox
MBox.Label1.Caption = Title
MBox.Label2.Caption = Message
MBReturn = 0
MBox.Show 1
Do Until MBReturn <> 0
DoEvents
Loop
End Function
