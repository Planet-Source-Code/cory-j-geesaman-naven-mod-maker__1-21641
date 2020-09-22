VERSION 5.00
Begin VB.Form MBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00929A93&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4590
   FillColor       =   &H00FFFFFF&
   Icon            =   "MBox.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   Picture         =   "MBox.frx":08CA
   ScaleHeight     =   161
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   306
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1440
      Top             =   720
   End
   Begin VB.PictureBox But1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   345
      Index           =   1
      Left            =   1635
      Picture         =   "MBox.frx":24BA4
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   95
      TabIndex        =   2
      Top             =   1800
      Width           =   1425
      Begin VB.Label Label3 
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
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   0
         TabIndex        =   3
         Top             =   45
         Width           =   1425
      End
   End
   Begin VB.Image Image3 
      Height          =   345
      Left            =   360
      Picture         =   "MBox.frx":25542
      Top             =   600
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Image Image2 
      Height          =   345
      Left            =   360
      Picture         =   "MBox.frx":25D21
      Top             =   960
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1275
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   4215
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   120
      Left            =   630
      TabIndex        =   0
      Top             =   210
      Width           =   3570
   End
End
Attribute VB_Name = "MBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub But1_KeyPress(Index As Integer, KeyAscii As Integer)
Form_KeyPress KeyAscii
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 10 Or KeyAscii = 13 Then Label3_Click 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
MBReturn = 1
End Sub

Private Sub Label3_Click(Index As Integer)
MBReturn = 1
Unload Me
End Sub

Private Sub Label3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Set But1(1).Picture = Image3.Picture
End Sub

Private Sub Label3_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Set But1(1).Picture = Image2.Picture
End Sub

Private Sub Timer1_Timer()
Me.ZOrder 0
End Sub
