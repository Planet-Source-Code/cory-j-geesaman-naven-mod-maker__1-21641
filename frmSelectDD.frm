VERSION 5.00
Begin VB.Form frmSelectDD 
   BackColor       =   &H00929A93&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4590
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSelectDD.frx":0000
   ScaleHeight     =   161
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   306
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List1 
      BackColor       =   &H00929A93&
      Height          =   1815
      ItemData        =   "frmSelectDD.frx":242DA
      Left            =   240
      List            =   "frmSelectDD.frx":242DC
      TabIndex        =   1
      Top             =   405
      Width           =   4215
   End
   Begin VB.Label lCaption 
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
   Begin VB.Image BU 
      Height          =   345
      Left            =   4440
      Picture         =   "frmSelectDD.frx":242DE
      Top             =   390
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Image BD 
      Height          =   345
      Left            =   4440
      Picture         =   "frmSelectDD.frx":24C7C
      Top             =   30
      Visible         =   0   'False
      Width           =   1425
   End
End
Attribute VB_Name = "frmSelectDD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label3_Click(Index As Integer)
Me.Visible = False
Unload Me
End Sub

Private Sub List1_Click()
SelectDD = List1.List(List1.ListIndex)
Unload Me
End Sub
