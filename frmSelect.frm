VERSION 5.00
Begin VB.Form frmSelect 
   BackColor       =   &H00929A93&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7500
   Icon            =   "frmSelect.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSelect.frx":08CA
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   3120
      Top             =   2520
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   3600
      Top             =   2040
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3120
      Top             =   2040
   End
   Begin VB.PictureBox Picture4 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   4440
      Picture         =   "frmSelect.frx":2442
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   4
      Top             =   3840
      Width           =   1425
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "&Ok"
         BeginProperty Font 
            Name            =   "Calisto MT"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         TabIndex        =   5
         Top             =   60
         Width           =   1440
      End
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   5880
      Picture         =   "frmSelect.frx":2DE0
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   2
      Top             =   3840
      Width           =   1425
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Calisto MT"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         TabIndex        =   3
         Top             =   60
         Width           =   1440
      End
   End
   Begin VB.ListBox ListBox 
      BackColor       =   &H00929A93&
      ForeColor       =   &H00400000&
      Height          =   2985
      ItemData        =   "frmSelect.frx":377E
      Left            =   330
      List            =   "frmSelect.frx":3780
      MultiSelect     =   1  'Simple
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   720
      Width           =   6975
   End
   Begin VB.Image BUp 
      Height          =   345
      Left            =   6120
      Picture         =   "frmSelect.frx":3782
      Top             =   360
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Image BDown 
      Height          =   345
      Left            =   6120
      Picture         =   "frmSelect.frx":4120
      Top             =   0
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label Label1 
      BackColor       =   &H00929A93&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1020
      TabIndex        =   0
      Top             =   330
      Width           =   5850
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public OT As Boolean
Public sL As Boolean

Public Function InList(Text)
If InStr(1, Text, "-", vbTextCompare) > 0 Then
a = Mid(Text, 1, InStr(1, Text, "-", vbTextCompare) - 1)
Else
a = Text
End If
j = -1
For i = 0 To ListBox.ListCount - 1 Step 1
If ListBox.List(i) = a Then
j = i
If a <> Text Then ListBox.List(j) = Text
End If
Next i
InList = j
End Function

Public Sub Init()
Setting = frmMain.SelectSettingV
Text = frmMain.SelectTextV
Dim TStr
Select Case Setting
Case 0 'Abilities
Label1.Caption = frmMain.lCaption.Caption & " - Select Abilities"
ListBox.AddItem "Blend With Enemy"
ListBox.AddItem "Burrow"
ListBox.AddItem "Cloak"
ListBox.AddItem "Fire Weapon"
ListBox.AddItem "Generate Power"
ListBox.AddItem "Give Energy"
ListBox.AddItem "Healing"
ListBox.AddItem "Parasite"
ListBox.AddItem "Poison"
ListBox.AddItem "Repairing"
ListBox.AddItem "Recharge Shields"
ListBox.AddItem "Sink"
ListBox.AddItem "Take Energy"
ListBox.AddItem "Teleport Self"
ListBox.AddItem "Teleport Self And Surrounding Units"
ListBox.AddItem "Teleport Units"
ListBox.AddItem "Unit Help Cloud"
ListBox.AddItem "Use Armor"
ListBox.AddItem "Use Shield"
Case 1 'Multiple Units
Label1.Caption = frmMain.lCaption.Caption & " - Select Units"
For i = 1 To frmMain.DataGrid.Rows - 1 Step 1
ListBox.AddItem "Infantry - " & frmMain.DataGrid.TextMatrix(i, 1)
Next i
For i = 1 To frmMain.Buildings.Rows - 1 Step 1
ListBox.AddItem "Building - " & frmMain.Buildings.TextMatrix(i, 1)
Next i
For i = 1 To frmMain.Weapons.Rows - 1 Step 1
ListBox.AddItem "Weapon - " & frmMain.Weapons.TextMatrix(i, 1)
Next i
For i = 1 To frmMain.Armor.Rows - 1 Step 1
ListBox.AddItem "Armor - " & frmMain.Armor.TextMatrix(i, 1)
Next i
For i = 1 To frmMain.Shields.Rows - 1 Step 1
ListBox.AddItem "Shield - " & frmMain.Shields.TextMatrix(i, 1)
Next i
Case 2 'Upgrades
Label1.Caption = frmMain.lCaption.Caption & " - Select Properties To Upgrade"
ListBox.AddItem "Damage"
ListBox.AddItem "Energy"
ListBox.AddItem "Energy Recharge Rate"
ListBox.AddItem "Fireing Speed"
ListBox.AddItem "HP"
ListBox.AddItem "HP Heal"
ListBox.AddItem "Range"
ListBox.AddItem "Sight"
ListBox.AddItem "Speed"
Case 3 'Multiple Abilities
Label1.Caption = frmMain.lCaption.Caption & " - Select Abilities"
For i = 1 To frmMain.Abilities.Rows - 1 Step 1
ListBox.AddItem frmMain.Abilities.TextMatrix(i, 1)
Next i
Case 4 'Multiple Upgrades
Label1.Caption = frmMain.lCaption.Caption & " - Select Upgrades"
For i = 1 To frmMain.Upgrades.Rows - 1 Step 1
ListBox.AddItem frmMain.Upgrades.TextMatrix(i, 1)
Next i
Case 5 'Fields Emitted
Label1.Caption = frmMain.lCaption.Caption & " - Select Fields Emitted"
ListBox.AddItem "Cloak"
ListBox.AddItem "Healing"
ListBox.AddItem "Power"
ListBox.AddItem "Quick Energy"
ListBox.AddItem "Repairing"
Case 6 'Multiple Infantry/Workers
Label1.Caption = frmMain.lCaption.Caption & " - Select Infantry"
For i = 1 To frmMain.DataGrid.Rows - 1 Step 1
If frmMain.DataGrid.TextMatrix(i, 46) = "1-Infantry" Or frmMain.DataGrid.TextMatrix(i, 46) = "4-Worker" Then
ListBox.AddItem frmMain.DataGrid.TextMatrix(i, 1)
End If
Next i
Case 7 'Multiple Buildings
Label1.Caption = frmMain.lCaption.Caption & " - Select Buildings"
For i = 1 To frmMain.Buildings.Rows - 1 Step 1
ListBox.AddItem frmMain.Buildings.TextMatrix(i, 1)
Next i
Case 8 'Multiple AddOn Select
Label1.Caption = frmMain.lCaption.Caption & " - Select AddOns"
For i = 1 To frmMain.Buildings.Rows - 1 Step 1
If frmMain.Buildings.TextMatrix(i, 89) = "1-True" Then
ListBox.AddItem frmMain.Buildings.TextMatrix(i, 1)
End If
Next i
End Select
sL = True
If Len(Text) > 0 Then
TStr = ""
For i = 1 To Len(Text) Step 1
If Mid(Text, i, 1) = ";" Then
j = InList(TStr)
If j > -1 Then ListBox.Selected(j) = True
TStr = ""
Else
TStr = TStr & Mid(Text, i, 1)
End If
Next i
End If
sL = False
End Sub

Private Sub Form_Load()
Timer1.Enabled = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then DragForm Me
End Sub

Private Sub Label4_Click()
Unload Me
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set Picture3.Picture = BDown.Picture
End Sub

Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set Picture3.Picture = BUp.Picture
End Sub

Private Sub Label5_Click()
tdata = ""
For i = 0 To ListBox.ListCount - 1 Step 1
If ListBox.Selected(i) = True Then tdata = tdata & ListBox.List(i) & ";"
Next i
frmMain.CControl.TextMatrix(frmMain.CControl.Row, frmMain.CControl.Col) = tdata
Unload Me
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set Picture4.Picture = BDown.Picture
End Sub

Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set Picture4.Picture = BUp.Picture
End Sub

Private Sub ListBox_Click()
If sL = True Then Exit Sub
If ListBox.Selected(ListBox.ListIndex) = False Then Exit Sub
If Mid(ListBox.List(ListBox.ListIndex), 1, Len("Fire Weapon")) = "Fire Weapon" Then
SelectDD = ""
Load frmSelectDD
frmSelectDD.List1.Clear
For i = 1 To frmMain.Weapons.Rows - 1 Step 1
frmSelectDD.List1.AddItem frmMain.Weapons.TextMatrix(i, 1)
Next i
frmSelectDD.Show vbModal, Me
ListBox.List(ListBox.ListIndex) = "Fire Weapon" & "-" & SelectDD
ElseIf Mid(ListBox.List(ListBox.ListIndex), 1, Len("Use Armor")) = "Use Armor" Then
SelectDD = ""
Load frmSelectDD
frmSelectDD.List1.Clear
For i = 1 To frmMain.Armor.Rows - 1 Step 1
frmSelectDD.List1.AddItem frmMain.Armor.TextMatrix(i, 1)
Next i
frmSelectDD.Show vbModal, Me
ListBox.List(ListBox.ListIndex) = "Use Armor" & "-" & SelectDD
ElseIf Mid(ListBox.List(ListBox.ListIndex), 1, Len("Use Shield")) = "Use Shield" Then
SelectDD = ""
Load frmSelectDD
frmSelectDD.List1.Clear
For i = 1 To frmMain.Shields.Rows - 1 Step 1
frmSelectDD.List1.AddItem frmMain.Shields.TextMatrix(i, 1)
Next i
frmSelectDD.Show vbModal, Me
ListBox.List(ListBox.ListIndex) = "Use Shield" & "-" & SelectDD
End If
End Sub

Private Sub Timer1_Timer()
Init
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
If OT = True Then
If frmSelectDD.Visible = True Then
frmSelectDD.ZOrder 0
Else
Me.ZOrder 0
End If
Timer3.Enabled = False
Else
Timer2.Enabled = False
Timer3.Enabled = True
End If
End Sub

Private Sub Timer3_Timer()
If OT = True Then Timer2.Enabled = True
End Sub
