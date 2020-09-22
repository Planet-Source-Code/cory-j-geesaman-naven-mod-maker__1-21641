VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00929A93&
   BorderStyle     =   0  'None
   ClientHeight    =   2760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4590
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":08CA
   ScaleHeight     =   184
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   306
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox PB 
      BackColor       =   &H00929A93&
      Height          =   1815
      Index           =   1
      Left            =   360
      ScaleHeight     =   1755
      ScaleWidth      =   3915
      TabIndex        =   14
      Top             =   480
      Visible         =   0   'False
      Width           =   3975
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00929A93&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Calisto MT"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00929A93&
         Caption         =   "Graphics By:"
         BeginProperty Font 
            Name            =   "Calisto MT"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   3495
      End
   End
   Begin VB.PictureBox PB 
      BackColor       =   &H00929A93&
      Height          =   1815
      Index           =   4
      Left            =   360
      ScaleHeight     =   1755
      ScaleWidth      =   3915
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   3975
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00929A93&
         Caption         =   "Cory J. Geesaman"
         BeginProperty Font 
            Name            =   "Calisto MT"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00929A93&
         Caption         =   "Encryption Routines By:"
         BeginProperty Font 
            Name            =   "Calisto MT"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   3495
      End
   End
   Begin VB.PictureBox PB 
      BackColor       =   &H00929A93&
      Height          =   1815
      Index           =   3
      Left            =   360
      ScaleHeight     =   1755
      ScaleWidth      =   3915
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   3975
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00929A93&
         Caption         =   "Cory J. Geesaman"
         BeginProperty Font 
            Name            =   "Calisto MT"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00929A93&
         Caption         =   "File Format By:"
         BeginProperty Font 
            Name            =   "Calisto MT"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   3495
      End
   End
   Begin VB.PictureBox PB 
      BackColor       =   &H00929A93&
      Height          =   1815
      Index           =   2
      Left            =   360
      ScaleHeight     =   1755
      ScaleWidth      =   3915
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   3975
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00929A93&
         Caption         =   "Coded By:"
         BeginProperty Font 
            Name            =   "Calisto MT"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00929A93&
         Caption         =   "Cory J. Geesaman"
         BeginProperty Font 
            Name            =   "Calisto MT"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   3495
      End
   End
   Begin VB.PictureBox PB 
      BackColor       =   &H00929A93&
      Height          =   1815
      Index           =   0
      Left            =   360
      ScaleHeight     =   1755
      ScaleWidth      =   3915
      TabIndex        =   1
      Top             =   480
      Width           =   3975
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00929A93&
         Caption         =   "©2001 Moonson Entertainment"
         BeginProperty Font 
            Name            =   "Calisto MT"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   3495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00929A93&
         BeginProperty Font 
            Name            =   "Calisto MT"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00929A93&
         BeginProperty Font 
            Name            =   "Calisto MT"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   120
         Width           =   3495
      End
   End
   Begin VB.Image l 
      Height          =   180
      Left            =   1905
      Picture         =   "frmAbout.frx":3976
      Top             =   2400
      Width           =   270
   End
   Begin VB.Image m 
      Height          =   180
      Left            =   2175
      Picture         =   "frmAbout.frx":3C58
      Top             =   2400
      Width           =   240
   End
   Begin VB.Image r 
      Height          =   195
      Left            =   2415
      Picture         =   "frmAbout.frx":3EDA
      Top             =   2400
      Width           =   270
   End
   Begin VB.Image ld 
      Height          =   180
      Left            =   0
      Picture         =   "frmAbout.frx":41F4
      Top             =   0
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image md 
      Height          =   180
      Left            =   240
      Picture         =   "frmAbout.frx":44D6
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image rd 
      Height          =   195
      Left            =   480
      Picture         =   "frmAbout.frx":4758
      Top             =   0
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image lu 
      Height          =   180
      Left            =   0
      Picture         =   "frmAbout.frx":4A72
      Top             =   240
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image mu 
      Height          =   180
      Left            =   255
      Picture         =   "frmAbout.frx":4D54
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image ru 
      Height          =   195
      Left            =   480
      Picture         =   "frmAbout.frx":4FD6
      Top             =   240
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label lCaption 
      BackColor       =   &H00929A93&
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
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cP As Integer

Private Sub Form_Load()
cP = 0
lCaption = frmMain.lCaption.Caption & " - About"
Label1.Caption = App.Title
If App.Minor < 10 And App.Revision < 10 Then
Label2.Caption = "v" & App.Major & "." & App.Minor & App.Revision
Else
Label2.Caption = "v" & App.Major & "." & App.Minor & "." & App.Revision
End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
End Sub

Private Sub l_Click()
cP = cP - 1
If cP < 0 Then
cP = 4
End If
For i = 0 To 4 Step 1
If i = cP Then
PB(i).Visible = True
Else
PB(i).Visible = False
End If
Next i
End Sub

Private Sub l_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set l.Picture = ld.Picture
End Sub

Private Sub l_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set l.Picture = lu.Picture
End Sub

Private Sub lCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
End Sub

Private Sub m_Click()
Unload Me
End Sub

Private Sub m_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set m.Picture = md.Picture
End Sub

Private Sub m_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set m.Picture = mu.Picture
End Sub

Private Sub r_Click()
cP = cP + 1
If cP > 4 Then
cP = 0
End If
For i = 0 To 4 Step 1
If i = cP Then
PB(i).Visible = True
Else
PB(i).Visible = False
End If
Next i
End Sub

Private Sub r_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set r.Picture = rd.Picture
End Sub

Private Sub r_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set r.Picture = ru.Picture
End Sub
