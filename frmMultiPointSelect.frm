VERSION 5.00
Begin VB.Form frmMultiPointSelect 
   BackColor       =   &H00929A93&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4590
   Icon            =   "frmMultiPointSelect.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMultiPointSelect.frx":08CA
   ScaleHeight     =   161
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   306
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox YBox2 
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2385
      TabIndex        =   10
      Text            =   "0"
      Top             =   1320
      Width           =   1800
   End
   Begin VB.TextBox XBox2 
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      TabIndex        =   9
      Text            =   "0"
      Top             =   1320
      Width           =   1800
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox BOk 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   840
      Picture         =   "frmMultiPointSelect.frx":24BA4
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   95
      TabIndex        =   7
      Top             =   1800
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
         TabIndex        =   8
         Top             =   45
         Width           =   1425
      End
   End
   Begin VB.PictureBox BCancel 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   2400
      Picture         =   "frmMultiPointSelect.frx":25542
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   95
      TabIndex        =   5
      Top             =   1800
      Width           =   1425
      Begin VB.Label LCancel 
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   0
         TabIndex        =   6
         Top             =   45
         Width           =   1425
      End
   End
   Begin VB.TextBox YBox 
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2415
      TabIndex        =   2
      Text            =   "0"
      Top             =   720
      Width           =   1800
   End
   Begin VB.TextBox XBox 
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   630
      TabIndex        =   1
      Text            =   "0"
      Top             =   720
      Width           =   1800
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00929A93&
      BackStyle       =   0  'Transparent
      Caption         =   "Y2:"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2385
      TabIndex        =   12
      Top             =   1080
      Width           =   285
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00929A93&
      BackStyle       =   0  'Transparent
      Caption         =   "X2:"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   600
      TabIndex        =   11
      Top             =   1080
      Width           =   285
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00929A93&
      BackStyle       =   0  'Transparent
      Caption         =   "Y1:"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2415
      TabIndex        =   4
      Top             =   480
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00929A93&
      BackStyle       =   0  'Transparent
      Caption         =   "X1:"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   630
      TabIndex        =   3
      Top             =   480
      Width           =   300
   End
   Begin VB.Image BU 
      Height          =   345
      Left            =   4440
      Picture         =   "frmMultiPointSelect.frx":25EE0
      Top             =   360
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Image BD 
      Height          =   345
      Left            =   4440
      Picture         =   "frmMultiPointSelect.frx":2687E
      Top             =   0
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label lCaption 
      BackColor       =   &H00929A93&
      BackStyle       =   0  'Transparent
      Caption         =   "Point Select"
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
Attribute VB_Name = "frmMultiPointSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LCancel_Click()
Unload Me
End Sub

Private Sub LCancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set BCancel.Picture = BD.Picture
End Sub

Private Sub LCancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set BCancel.Picture = BU.Picture
End Sub

Private Sub LOk_Click()
If frmMain.IsNumber(XBox.Text) = True And frmMain.IsNumber(YBox.Text) = True And frmMain.IsNumber(XBox2.Text) = True And frmMain.IsNumber(YBox2.Text) = True Then
frmMain.CControl.TextMatrix(frmMain.CControl.Row, frmMain.CControl.Col) = XBox.Text & "," & YBox.Text & "," & XBox2.Text & "," & YBox2.Text
Unload Me
Else
Msbox "The fields must be numbers!", "Invalid Field(s)"
End If
End Sub

Private Sub LOk_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set BOk.Picture = BD.Picture
End Sub

Private Sub LOk_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set BOk.Picture = BU.Picture
End Sub

Private Sub Timer1_Timer()
If Len(frmMain.SelectTextV) > 0 Then
A = InStr(1, frmMain.SelectTextV, ",", vbTextCompare)
b = Mid(frmMain.SelectTextV, 1, A - 1) ''''''''
C = Mid(frmMain.SelectTextV, A + 1)
d = InStr(1, C, ",", vbTextCompare)
e = Mid(C, 1, d - 1) ''''''''
f = Mid(C, d + 1)
g = InStr(1, f, ",", vbTextCompare)
h = Mid(f, 1, g - 1) ''''''''
i = Mid(f, g + 1) '''''''''''''
XBox.Text = b
YBox.Text = e
XBox2.Text = h
YBox2.Text = i
End If
Timer1.Enabled = False
End Sub
