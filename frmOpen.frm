VERSION 5.00
Begin VB.Form frmOpen 
   BackColor       =   &H00929A93&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7500
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmOpen.frx":0000
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.FileListBox File1 
      Height          =   2820
      Left            =   3840
      Pattern         =   "*.mod"
      TabIndex        =   9
      Top             =   1050
      Width           =   3495
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   330
      TabIndex        =   8
      Top             =   1050
      Width           =   3495
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   930
      TabIndex        =   7
      Top             =   720
      Width           =   6405
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00929A93&
      Height          =   345
      Left            =   330
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   3900
      Width           =   4050
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   4440
      Picture         =   "frmOpen.frx":1B78
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   3
      Top             =   3900
      Width           =   1425
      Begin VB.Label Label2 
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
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   60
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   5880
      Picture         =   "frmOpen.frx":2357
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   1
      Top             =   3900
      Width           =   1425
      Begin VB.Label Label3 
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
         Height          =   195
         Left            =   0
         TabIndex        =   2
         Top             =   60
         Width           =   1440
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Look &in:"
      Height          =   195
      Left            =   330
      TabIndex        =   6
      Top             =   810
      Width           =   570
   End
   Begin VB.Image BDown 
      Height          =   345
      Left            =   6120
      Picture         =   "frmOpen.frx":2CF5
      Top             =   0
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Image BUp 
      Height          =   345
      Left            =   6120
      Picture         =   "frmOpen.frx":34D4
      Top             =   360
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
Attribute VB_Name = "frmOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Dir1_KeyDown(KeyCode As Integer, Shift As Integer)
Dir1_Change
End Sub

Private Sub Dir1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dir1_Change
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Drive1_KeyDown(KeyCode As Integer, Shift As Integer)
Drive1_Change
End Sub

Private Sub File1_KeyDown(KeyCode As Integer, Shift As Integer)
'File1_Click
End Sub

Private Sub Form_Load()
Dir1.Path = App.Path
Label1.Caption = frmMain.lCaption.Caption & " - Open Mod"
ShapeForm Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
End Sub

Private Sub Label2_Click()
If File1.FileName <> "" Then
If Right(File1.Path, 1) = "\" Or Right(File1.Path, 1) = "/" Then
A = File1.Path & File1.FileName
Else
A = File1.Path & "\" & File1.FileName
End If
End If
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

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If File1.FileName <> "" Then Set Picture1.Picture = BDown.Picture
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If File1.FileName <> "" Then Set Picture1.Picture = BUp.Picture
End Sub
