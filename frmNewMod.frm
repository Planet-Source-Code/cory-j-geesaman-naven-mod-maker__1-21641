VERSION 5.00
Begin VB.Form frmNewMod 
   BackColor       =   &H00929A93&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4590
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmNewMod.frx":0000
   ScaleHeight     =   183
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   306
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   3000
      Picture         =   "frmNewMod.frx":30AC
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   2
      Top             =   2160
      Width           =   1425
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         TabIndex        =   4
         Top             =   60
         Width           =   1440
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   1440
      Picture         =   "frmNewMod.frx":3A4A
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   1
      Top             =   2160
      Width           =   1425
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "&Ok"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   60
         Width           =   1455
      End
   End
   Begin VB.Image Radio 
      Height          =   225
      Index           =   4
      Left            =   240
      Picture         =   "frmNewMod.frx":43E8
      Top             =   960
      Width           =   225
   End
   Begin VB.Label Label 
      BackColor       =   &H00929A93&
      Caption         =   "&Ability Mod"
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   10
      Top             =   960
      Width           =   3975
   End
   Begin VB.Image Radio 
      Height          =   225
      Index           =   3
      Left            =   240
      Picture         =   "frmNewMod.frx":477B
      Top             =   1440
      Width           =   225
   End
   Begin VB.Label Label 
      BackColor       =   &H00929A93&
      Caption         =   "&Sound Mod"
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   9
      Top             =   1440
      Width           =   3975
   End
   Begin VB.Image CheckOff 
      Height          =   225
      Left            =   4320
      Picture         =   "frmNewMod.frx":4B0E
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image CheckOn 
      Height          =   225
      Left            =   4080
      Picture         =   "frmNewMod.frx":4EC4
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image RadioOn 
      Height          =   225
      Left            =   4080
      Picture         =   "frmNewMod.frx":52FA
      Top             =   960
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image RadioOff 
      Height          =   225
      Left            =   4320
      Picture         =   "frmNewMod.frx":5700
      Top             =   960
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image BUp 
      Height          =   345
      Left            =   3120
      Picture         =   "frmNewMod.frx":5A93
      Top             =   480
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Image BDown 
      Height          =   345
      Left            =   3120
      Picture         =   "frmNewMod.frx":6431
      Top             =   120
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Image Check 
      Height          =   225
      Left            =   360
      Picture         =   "frmNewMod.frx":6C10
      Top             =   1800
      Width           =   225
   End
   Begin VB.Label Label4 
      BackColor       =   &H00929A93&
      Caption         =   "&Lock Mod After Compiling"
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Label Label 
      BackColor       =   &H00929A93&
      Caption         =   "&Image Mod"
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   7
      Top             =   1200
      Width           =   3975
   End
   Begin VB.Label Label 
      BackColor       =   &H00929A93&
      Caption         =   "&Upgrade Mod"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   6
      Top             =   720
      Width           =   3975
   End
   Begin VB.Label Label 
      BackColor       =   &H00929A93&
      Caption         =   "&Unit Mod"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   5
      Top             =   480
      Width           =   3975
   End
   Begin VB.Image Radio 
      Height          =   225
      Index           =   2
      Left            =   240
      Picture         =   "frmNewMod.frx":7046
      Top             =   1200
      Width           =   225
   End
   Begin VB.Image Radio 
      Height          =   225
      Index           =   0
      Left            =   240
      Picture         =   "frmNewMod.frx":73D9
      Top             =   480
      Width           =   225
   End
   Begin VB.Image Radio 
      Height          =   225
      Index           =   1
      Left            =   240
      Picture         =   "frmNewMod.frx":77DF
      Top             =   720
      Width           =   225
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
      Height          =   120
      Left            =   630
      TabIndex        =   0
      Top             =   210
      Width           =   3570
   End
End
Attribute VB_Name = "frmNewMod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ModType As Integer
Public ModLock As Boolean

Private Sub Check_Click()
Label4_Click
End Sub

Private Sub Form_Load()
Label1.Caption = frmMain.lCaption.Caption & " - New Mod"
ModLock = True
ModType = 0
End Sub

Private Sub Label_Click(Index As Integer)
For i = 0 To 2 Step 1
If i = Index Then Set Radio(i).Picture = RadioOn.Picture Else Set Radio(i).Picture = RadioOff.Picture
Next i
ModType = Index
End Sub

Private Sub Label2_Click()
frmMain.ModLock = modloack
frmMain.CreateMod ModType
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set Picture1.Picture = BDown.Picture
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set Picture1.Picture = BUp.Picture
End Sub

Private Sub Label3_Click()
Unload Me
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set Picture2.Picture = BDown.Picture
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set Picture2.Picture = BUp.Picture
End Sub

Private Sub Label4_Click()
ModLock = Not ModLock
If ModLock = True Then
Set Check.Picture = CheckOn.Picture
Else
Set Check.Picture = CheckOff.Picture
End If
End Sub

Private Sub Radio_Click(Index As Integer)
Label_Click Index
End Sub
