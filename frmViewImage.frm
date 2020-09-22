VERSION 5.00
Begin VB.Form frmViewImage 
   BackColor       =   &H00929A93&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4590
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmViewImage.frx":0000
   ScaleHeight     =   326
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   306
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00929A93&
      Height          =   3735
      Left            =   360
      ScaleHeight     =   245
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   261
      TabIndex        =   2
      Top             =   480
      Width           =   3975
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00929A93&
         BorderStyle     =   0  'None
         Height          =   1815
         Left            =   0
         ScaleHeight     =   121
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   129
         TabIndex        =   4
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.PictureBox BOk 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   1560
      Picture         =   "frmViewImage.frx":107F
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   95
      TabIndex        =   0
      Top             =   4320
      Width           =   1425
      Begin VB.Label LOk 
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
         Left            =   0
         TabIndex        =   1
         Top             =   45
         Width           =   1425
      End
   End
   Begin VB.Label lCaption 
      BackColor       =   &H00929A93&
      BackStyle       =   0  'Transparent
      Caption         =   "View Image"
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
      TabIndex        =   3
      Top             =   210
      Width           =   3570
   End
   Begin VB.Image BD 
      Height          =   345
      Left            =   4440
      Picture         =   "frmViewImage.frx":1A1D
      Top             =   0
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Image BU 
      Height          =   345
      Left            =   4440
      Picture         =   "frmViewImage.frx":21FC
      Top             =   360
      Visible         =   0   'False
      Width           =   1425
   End
End
Attribute VB_Name = "frmViewImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LOk_Click()
Unload Me
End Sub

Private Sub LOk_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set BOk.Picture = BD.Picture
End Sub

Private Sub LOk_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set BOk.Picture = BU.Picture
End Sub
