VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00929A93&
   BorderStyle     =   0  'None
   ClientHeight    =   8160
   ClientLeft      =   -45
   ClientTop       =   -360
   ClientWidth     =   10440
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   544
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   696
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList 
      Left            =   1440
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   9607827
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   7299152
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   "Ability"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15A6
            Key             =   "Animation"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2282
            Key             =   "Armor"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F5E
            Key             =   "Building"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3C3A
            Key             =   "Image"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4916
            Key             =   "Infantry"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":55F2
            Key             =   "Race"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":62CE
            Key             =   "Shield"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6FAA
            Key             =   "Sound"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7C86
            Key             =   "Upgrade"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8962
            Key             =   "Weapon"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton PlayView 
      Height          =   240
      Left            =   4920
      Picture         =   "frmMain.frx":963E
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   3240
      Visible         =   0   'False
      Width           =   255
   End
   Begin prjModMaker.ctlProgBar PBar 
      Height          =   735
      Left            =   2160
      TabIndex        =   55
      Top             =   360
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1296
      BGC             =   9607827
      C1              =   9607827
      C2              =   4194304
      vP              =   0
      lfC             =   0
      lbC             =   -2147483633
      lV              =   -1  'True
      lS              =   0
   End
   Begin RichTextLib.RichTextBox TRtf 
      Height          =   420
      Left            =   5640
      TabIndex        =   54
      Top             =   4320
      Visible         =   0   'False
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393217
      BackColor       =   -2147483633
      BorderStyle     =   0
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":980C
   End
   Begin VB.PictureBox TPic 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   5160
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   53
      Top             =   4320
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox TBox 
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   4680
      TabIndex        =   52
      Top             =   4320
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Timer Timer3 
      Interval        =   25
      Left            =   5640
      Top             =   3840
   End
   Begin VB.PictureBox VM 
      Height          =   6645
      Left            =   345
      ScaleHeight     =   439
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   45
      Top             =   1185
      Width           =   960
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   8580
         Left            =   30
         TabIndex        =   46
         Top             =   60
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   15134
         ButtonWidth     =   1535
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImageList2"
         HotImageList    =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   11
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Abilities"
               Key             =   "Ability"
               ImageIndex      =   1
               Style           =   2
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Animations"
               Key             =   "Animation"
               ImageIndex      =   2
               Style           =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Armor"
               Key             =   "Armor"
               ImageIndex      =   3
               Style           =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Buildings"
               Key             =   "Buildings"
               ImageIndex      =   4
               Style           =   2
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Images"
               Key             =   "Image"
               Object.ToolTipText     =   "Images"
               ImageIndex      =   5
               Style           =   2
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Infantry"
               Key             =   "Infantry"
               ImageIndex      =   6
               Style           =   2
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Races"
               Key             =   "Race"
               ImageIndex      =   7
               Style           =   2
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Sheilds"
               Key             =   "Sheild"
               ImageIndex      =   8
               Style           =   2
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Sounds"
               Key             =   "Sound"
               Object.ToolTipText     =   "Sounds"
               ImageIndex      =   9
               Style           =   2
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Upgrades"
               Key             =   "Upgrade"
               Object.ToolTipText     =   "Upgrades"
               ImageIndex      =   10
               Style           =   2
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Weapons"
               Key             =   "Weapon"
               Object.ToolTipText     =   "Weapons"
               ImageIndex      =   11
               Style           =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox Image29 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   960
      Picture         =   "frmMain.frx":98BA
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   8
      Top             =   1050
      Visible         =   0   'False
      Width           =   615
      Begin VB.Line Line4 
         X1              =   0
         X2              =   30
         Y1              =   9
         Y2              =   9
      End
   End
   Begin VB.PictureBox Image18 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1080
      Left            =   1410
      Picture         =   "frmMain.frx":A8F8
      ScaleHeight     =   1080
      ScaleWidth      =   165
      TabIndex        =   3
      Top             =   1395
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox Image17 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1080
      Left            =   1410
      Picture         =   "frmMain.frx":B35A
      ScaleHeight     =   1080
      ScaleWidth      =   165
      TabIndex        =   11
      Top             =   2040
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox Image25 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1080
      Left            =   1410
      Picture         =   "frmMain.frx":BDBC
      ScaleHeight     =   1080
      ScaleWidth      =   165
      TabIndex        =   10
      Top             =   2790
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox Image24 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   1455
      Picture         =   "frmMain.frx":C81E
      ScaleHeight     =   450
      ScaleWidth      =   1170
      TabIndex        =   19
      Top             =   3480
      Visible         =   0   'False
      Width           =   1170
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "&Exit      Esc"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.PictureBox Image23 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   1440
      Picture         =   "frmMain.frx":E408
      ScaleHeight     =   450
      ScaleWidth      =   1710
      TabIndex        =   17
      Top             =   2760
      Visible         =   0   'False
      Width           =   1710
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Save &As...       F12"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.PictureBox Image22 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   1440
      Picture         =   "frmMain.frx":10C9A
      ScaleHeight     =   450
      ScaleWidth      =   1710
      TabIndex        =   14
      Top             =   2250
      Visible         =   0   'False
      Width           =   1710
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "&Save...         Ctrl+S"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.PictureBox Image21 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   1440
      Picture         =   "frmMain.frx":1352C
      ScaleHeight     =   450
      ScaleWidth      =   1710
      TabIndex        =   12
      Top             =   1800
      Visible         =   0   'False
      Width           =   1710
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "&Open...         Ctrl+O"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.PictureBox Image20 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   1440
      Picture         =   "frmMain.frx":15DBE
      ScaleHeight     =   450
      ScaleWidth      =   1710
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   1710
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&New             Ctrl+N"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2580
      Picture         =   "frmMain.frx":18650
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   25
      Top             =   1050
      Visible         =   0   'False
      Width           =   615
      Begin VB.Line Line2 
         X1              =   0
         X2              =   30
         Y1              =   9
         Y2              =   9
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   1770
      Picture         =   "frmMain.frx":1968E
      ScaleHeight     =   450
      ScaleWidth      =   1605
      TabIndex        =   23
      Top             =   720
      Width           =   1605
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "DataGrid"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   220
         TabIndex        =   24
         Top             =   120
         Width           =   1170
      End
   End
   Begin VB.PictureBox Picture16 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1080
      Left            =   3030
      Picture         =   "frmMain.frx":1BCC8
      ScaleHeight     =   1080
      ScaleWidth      =   165
      TabIndex        =   42
      Top             =   1095
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox Picture15 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   3075
      Picture         =   "frmMain.frx":1C72A
      ScaleHeight     =   450
      ScaleWidth      =   2115
      TabIndex        =   41
      Top             =   1800
      Visible         =   0   'False
      Width           =   2115
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "&Remove Infantry"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   240
         TabIndex        =   44
         Top             =   120
         Width           =   1770
      End
   End
   Begin VB.PictureBox Picture13 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   3075
      Picture         =   "frmMain.frx":1F91C
      ScaleHeight     =   450
      ScaleWidth      =   2115
      TabIndex        =   40
      Top             =   1320
      Visible         =   0   'False
      Width           =   2115
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Add Infantry"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   240
         TabIndex        =   43
         Top             =   120
         Width           =   1800
      End
   End
   Begin VB.PictureBox Image15 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3735
      Picture         =   "frmMain.frx":22B0E
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   9
      Top             =   1050
      Visible         =   0   'False
      Width           =   615
      Begin VB.Line Line3 
         X1              =   0
         X2              =   30
         Y1              =   9
         Y2              =   9
      End
   End
   Begin VB.PictureBox Image12 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   3375
      Picture         =   "frmMain.frx":23B4C
      ScaleHeight     =   450
      ScaleWidth      =   1170
      TabIndex        =   4
      Top             =   720
      Width           =   1170
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "&Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   210
         TabIndex        =   6
         Top             =   120
         Width           =   750
      End
   End
   Begin VB.PictureBox Picture12 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   2760
      Picture         =   "frmMain.frx":25736
      ScaleHeight     =   315
      ScaleWidth      =   4125
      TabIndex        =   27
      Top             =   720
      Width           =   4125
   End
   Begin VB.PictureBox Picture11 
      BackColor       =   &H00929A93&
      BorderStyle     =   0  'None
      Height          =   120
      Left            =   3720
      ScaleHeight     =   120
      ScaleWidth      =   1335
      TabIndex        =   26
      Top             =   600
      Width           =   1335
   End
   Begin VB.PictureBox Image14 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1080
      Left            =   4185
      Picture         =   "frmMain.frx":29B64
      ScaleHeight     =   1080
      ScaleWidth      =   165
      TabIndex        =   16
      Top             =   630
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox Image16 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   4230
      Picture         =   "frmMain.frx":2A5C6
      ScaleHeight     =   450
      ScaleWidth      =   1170
      TabIndex        =   21
      Top             =   1320
      Visible         =   0   'False
      Width           =   1170
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&About    F1"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5160
      Top             =   3840
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1500
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   9607827
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   7299152
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C1B0
            Key             =   "Ability"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2CE8C
            Key             =   "Animation"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2DB68
            Key             =   "Armor"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E844
            Key             =   "Building"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F520
            Key             =   "Image"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":301FC
            Key             =   "Infantry"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":30ED8
            Key             =   "Race"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":31BB4
            Key             =   "Shield"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":32890
            Key             =   "Sound"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3356C
            Key             =   "Upgrade"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":34248
            Key             =   "Weapon"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1500
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   9607827
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   -2147483633
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":34F24
            Key             =   "Ability"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3537C
            Key             =   "Animation"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":36058
            Key             =   "Armor"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":36D34
            Key             =   "Building"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":37A10
            Key             =   "Image"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":37E68
            Key             =   "Infantry"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":38B44
            Key             =   "Race"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":39820
            Key             =   "Shield"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3A4FC
            Key             =   "Sound"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3B1D8
            Key             =   "Upgrade"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3BEB4
            Key             =   "Weapon"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Line1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   1320
      ScaleHeight     =   15
      ScaleWidth      =   8895
      TabIndex        =   39
      Top             =   1440
      Visible         =   0   'False
      Width           =   8895
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4680
      Top             =   3840
   End
   Begin VB.ComboBox LBox 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      IntegralHeight  =   0   'False
      Left            =   5640
      Style           =   2  'Dropdown List
      TabIndex        =   38
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox TextEdit 
      BackColor       =   &H00929A93&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   240
      Left            =   3720
      TabIndex        =   31
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton FileSelect 
      Caption         =   "..."
      Height          =   240
      Left            =   4560
      TabIndex        =   30
      Top             =   3240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ListBox ChrList 
      Height          =   255
      ItemData        =   "frmMain.frx":3C30C
      Left            =   7920
      List            =   "frmMain.frx":3C60D
      TabIndex        =   28
      Top             =   360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.PictureBox Image11 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   600
      Picture         =   "frmMain.frx":3CBC6
      ScaleHeight     =   450
      ScaleWidth      =   1170
      TabIndex        =   5
      Top             =   720
      Width           =   1170
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "&File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   215
         TabIndex        =   7
         Top             =   120
         Width           =   750
      End
   End
   Begin MSFlexGridLib.MSFlexGrid DataGrid 
      Height          =   6375
      Left            =   1320
      TabIndex        =   29
      Top             =   1440
      Visible         =   0   'False
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   11245
      _Version        =   393216
      BackColor       =   9607827
      ForeColor       =   4194304
      BackColorSel    =   9607827
      ForeColorSel    =   4194304
      BackColorBkg    =   9607827
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calisto MT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid Upgrades 
      Height          =   6375
      Left            =   1320
      TabIndex        =   32
      Top             =   1440
      Visible         =   0   'False
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   11245
      _Version        =   393216
      BackColor       =   9607827
      ForeColor       =   4194304
      BackColorSel    =   9607827
      ForeColorSel    =   4194304
      BackColorBkg    =   9607827
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calisto MT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid Sounds 
      Height          =   6375
      Left            =   1320
      TabIndex        =   33
      Top             =   1440
      Visible         =   0   'False
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   11245
      _Version        =   393216
      BackColor       =   9607827
      ForeColor       =   4194304
      BackColorSel    =   9607827
      ForeColorSel    =   4194304
      BackColorBkg    =   9607827
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calisto MT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid Images 
      Height          =   6375
      Left            =   1320
      TabIndex        =   34
      Top             =   1440
      Visible         =   0   'False
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   11245
      _Version        =   393216
      BackColor       =   9607827
      ForeColor       =   4194304
      BackColorSel    =   9607827
      ForeColorSel    =   4194304
      BackColorBkg    =   9607827
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calisto MT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid Abilities 
      Height          =   6375
      Left            =   1320
      TabIndex        =   35
      Top             =   1440
      Visible         =   0   'False
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   11245
      _Version        =   393216
      BackColor       =   9607827
      ForeColor       =   4194304
      BackColorSel    =   9607827
      ForeColorSel    =   4194304
      BackColorBkg    =   9607827
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calisto MT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid Weapons 
      Height          =   6375
      Left            =   1320
      TabIndex        =   36
      Top             =   1440
      Visible         =   0   'False
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   11245
      _Version        =   393216
      BackColor       =   9607827
      ForeColor       =   4194304
      BackColorSel    =   9607827
      ForeColorSel    =   4194304
      BackColorBkg    =   9607827
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calisto MT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid Armor 
      Height          =   6375
      Left            =   1320
      TabIndex        =   47
      Top             =   1440
      Visible         =   0   'False
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   11245
      _Version        =   393216
      BackColor       =   9607827
      ForeColor       =   4194304
      BackColorSel    =   9607827
      ForeColorSel    =   4194304
      BackColorBkg    =   9607827
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calisto MT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid Animations 
      Height          =   6375
      Left            =   1320
      TabIndex        =   48
      Top             =   1440
      Visible         =   0   'False
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   11245
      _Version        =   393216
      BackColor       =   9607827
      ForeColor       =   4194304
      BackColorSel    =   9607827
      ForeColorSel    =   4194304
      BackColorBkg    =   9607827
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calisto MT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid Shields 
      Height          =   6375
      Left            =   1320
      TabIndex        =   49
      Top             =   1440
      Visible         =   0   'False
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   11245
      _Version        =   393216
      BackColor       =   9607827
      ForeColor       =   4194304
      BackColorSel    =   9607827
      ForeColorSel    =   4194304
      BackColorBkg    =   9607827
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calisto MT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid Buildings 
      Height          =   6375
      Left            =   1320
      TabIndex        =   50
      Top             =   1440
      Visible         =   0   'False
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   11245
      _Version        =   393216
      BackColor       =   9607827
      ForeColor       =   4194304
      BackColorSel    =   9607827
      ForeColorSel    =   4194304
      BackColorBkg    =   9607827
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calisto MT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid Races 
      Height          =   6375
      Left            =   1320
      TabIndex        =   51
      Top             =   1440
      Visible         =   0   'False
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   11245
      _Version        =   393216
      BackColor       =   9607827
      ForeColor       =   4194304
      BackColorSel    =   9607827
      ForeColorSel    =   4194304
      BackColorBkg    =   9607827
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calisto MT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lCaption 
      AutoSize        =   -1  'True
      BackColor       =   &H00929A93&
      Caption         =   "NAVEN - Mode Maker"
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
      Left            =   1020
      TabIndex        =   0
      Top             =   330
      Width           =   1875
   End
   Begin VB.Image Image5 
      Height          =   2100
      Index           =   0
      Left            =   0
      Picture         =   "frmMain.frx":3E7B0
      Top             =   1200
      Width           =   1185
   End
   Begin VB.Label lSection 
      Alignment       =   2  'Center
      BackColor       =   &H00929A93&
      Caption         =   "Units"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   1320
      TabIndex        =   37
      Top             =   1200
      Visible         =   0   'False
      Width           =   8895
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   6630
      Left            =   1320
      Top             =   1200
      Width           =   8910
   End
   Begin VB.Shape Shape1 
      Height          =   6630
      Left            =   1305
      Top             =   1185
      Width           =   8925
   End
   Begin VB.Image Image3 
      Height          =   1200
      Left            =   0
      Picture         =   "frmMain.frx":46B32
      Top             =   6960
      Width           =   1200
   End
   Begin VB.Image Image4 
      Height          =   1200
      Left            =   9240
      Picture         =   "frmMain.frx":4B674
      Top             =   6960
      Width           =   1200
   End
   Begin VB.Image Image6 
      Height          =   555
      Left            =   8490
      Picture         =   "frmMain.frx":501B6
      Top             =   720
      Width           =   1950
   End
   Begin VB.Image Image10 
      Height          =   315
      Index           =   2
      Left            =   15000
      Picture         =   "frmMain.frx":53AA0
      Top             =   720
      Width           =   4125
   End
   Begin VB.Image MaxRes 
      Height          =   180
      Left            =   9915
      Picture         =   "frmMain.frx":57ECE
      ToolTipText     =   "Maximize"
      Top             =   60
      Width           =   240
   End
   Begin VB.Image MaxResU 
      Height          =   180
      Left            =   510
      Picture         =   "frmMain.frx":58261
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image8 
      Height          =   1080
      Index           =   6
      Left            =   10275
      Picture         =   "frmMain.frx":585F4
      Top             =   1200
      Width           =   165
   End
   Begin VB.Image Image19 
      Height          =   1080
      Left            =   1410
      Picture         =   "frmMain.frx":59056
      Top             =   2610
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image Image9 
      Height          =   360
      Left            =   0
      Picture         =   "frmMain.frx":59AB8
      Top             =   720
      Width           =   2760
   End
   Begin VB.Image Image10 
      Height          =   315
      Index           =   1
      Left            =   6840
      Picture         =   "frmMain.frx":5CEBA
      Top             =   720
      Width           =   4125
   End
   Begin VB.Image CloseD 
      Height          =   195
      Left            =   750
      Picture         =   "frmMain.frx":612E8
      Top             =   0
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image MinD 
      Height          =   180
      Left            =   240
      Picture         =   "frmMain.frx":616D5
      Top             =   0
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image CloseU 
      Height          =   195
      Left            =   750
      Picture         =   "frmMain.frx":61ADC
      Top             =   240
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image MinU 
      Height          =   180
      Left            =   240
      Picture         =   "frmMain.frx":61F10
      Top             =   240
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image Image5 
      Height          =   2100
      Index           =   5
      Left            =   0
      Picture         =   "frmMain.frx":62363
      Top             =   11400
      Width           =   1185
   End
   Begin VB.Image Image8 
      Height          =   2100
      Index           =   5
      Left            =   9240
      Picture         =   "frmMain.frx":6A6E5
      Top             =   11400
      Width           =   1200
   End
   Begin VB.Image Min 
      Height          =   180
      Left            =   9645
      Picture         =   "frmMain.frx":72A67
      ToolTipText     =   "Minimize"
      Top             =   60
      Width           =   270
   End
   Begin VB.Image CloseP 
      Height          =   195
      Left            =   10155
      Picture         =   "frmMain.frx":72EBA
      ToolTipText     =   "Close"
      Top             =   60
      Width           =   270
   End
   Begin VB.Image Image7 
      Height          =   1200
      Index           =   2
      Left            =   11280
      Picture         =   "frmMain.frx":732EE
      Top             =   6960
      Width           =   5130
   End
   Begin VB.Image Image7 
      Height          =   1200
      Index           =   1
      Left            =   6240
      Picture         =   "frmMain.frx":87470
      Top             =   6960
      Width           =   5130
   End
   Begin VB.Image Image7 
      Height          =   1200
      Index           =   0
      Left            =   1200
      Picture         =   "frmMain.frx":9B5F2
      Top             =   6960
      Width           =   5130
   End
   Begin VB.Image Image5 
      Height          =   2100
      Index           =   4
      Left            =   0
      Picture         =   "frmMain.frx":AF774
      Top             =   9360
      Width           =   1185
   End
   Begin VB.Image Image5 
      Height          =   2100
      Index           =   3
      Left            =   0
      Picture         =   "frmMain.frx":B7AF6
      Top             =   7320
      Width           =   1185
   End
   Begin VB.Image Image5 
      Height          =   2100
      Index           =   2
      Left            =   0
      Picture         =   "frmMain.frx":BFE78
      Top             =   5280
      Width           =   1185
   End
   Begin VB.Image Image5 
      Height          =   2100
      Index           =   1
      Left            =   0
      Picture         =   "frmMain.frx":C81FA
      Top             =   3240
      Width           =   1185
   End
   Begin VB.Image Image8 
      Height          =   2100
      Index           =   4
      Left            =   9240
      Picture         =   "frmMain.frx":D057C
      Top             =   9360
      Width           =   1200
   End
   Begin VB.Image Image8 
      Height          =   2100
      Index           =   3
      Left            =   9240
      Picture         =   "frmMain.frx":D88FE
      Top             =   7320
      Width           =   1200
   End
   Begin VB.Image Image8 
      Height          =   2100
      Index           =   2
      Left            =   9240
      Picture         =   "frmMain.frx":E0C80
      Top             =   5280
      Width           =   1200
   End
   Begin VB.Image Image8 
      Height          =   2100
      Index           =   1
      Left            =   9240
      Picture         =   "frmMain.frx":E9002
      Top             =   3240
      Width           =   1200
   End
   Begin VB.Image Image8 
      Height          =   1080
      Index           =   0
      Left            =   10275
      Picture         =   "frmMain.frx":F1384
      Top             =   2280
      Width           =   165
   End
   Begin VB.Image MaxResD 
      Height          =   180
      Left            =   510
      Picture         =   "frmMain.frx":F1DE6
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   1200
      Left            =   0
      Picture         =   "frmMain.frx":F213D
      Top             =   0
      Width           =   1200
   End
   Begin VB.Image T 
      Height          =   1200
      Index           =   0
      Left            =   1200
      Picture         =   "frmMain.frx":F6C7F
      Top             =   0
      Width           =   5085
   End
   Begin VB.Image Image10 
      Height          =   315
      Index           =   3
      Left            =   10920
      Picture         =   "frmMain.frx":10AB81
      Top             =   720
      Width           =   4125
   End
   Begin VB.Image Image2 
      Height          =   1200
      Left            =   9240
      Picture         =   "frmMain.frx":10EFAF
      Top             =   0
      Width           =   1200
   End
   Begin VB.Image T 
      Height          =   1200
      Index           =   2
      Left            =   11280
      Picture         =   "frmMain.frx":113AF1
      Top             =   0
      Width           =   5085
   End
   Begin VB.Image T 
      Height          =   1200
      Index           =   1
      Left            =   6240
      Picture         =   "frmMain.frx":1279F3
      Top             =   0
      Width           =   5085
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public HF As Boolean
Public FF As Boolean
Public DF As Boolean
Public ModLock As Boolean
Public LastState As FormWindowStateConstants
Public TextIsNumber As Boolean
Public CControl As MSFlexGrid
Public GScrolling As Boolean
Public DGN As Integer
'''''''''''VERY IMPORTANT!!!!!!!!!!!!!!
Public SelectSettingV
Public SelectTextV
Public ModFileName As String
Public Saved As Boolean
Private GridDataArray()
Public TSv As Integer
Public OnTimer As Boolean
Private GArray As Variant
Private GArrayL2 As Variant
Private GArrayL3 As Variant
Private GArrayCount(0 To 10) As Integer
Private GArrayCountL2() As Integer
Private GArrayCountL3() As Integer
Private SoundArray(0 To 1024) As String
Public SoundArrayCount As Integer
Private aCount(1 To 11) As Long

Public Function GetSound(sRef As String) As String
a = Mid(sRef, 19)
GetSound = SoundArray(a)
End Function

Public Function AddSound(Sound As String, Optional CurSpot As Integer)
If Len(CurSpot) > 0 Then
b = CurSpot
Else
SoundArrayCount = SoundArrayCount + 1
b = SoundArrayCount
End If
SoundArray(b) = Sound
AddSound = b
End Function

Public Sub ClearSoundArray()
For i = 0 To 1024 Step 1
SoundArray(i) = ""
Next i
SoundArrayCount = 0
End Sub

Public Sub RemoveSound(Spot As Integer)
If Len(Spot) < 1 Then Exit Sub
If Spot < SoundArrayCount Then
For i = Spot To SoundArrayCount Step 1
If i + 1 < 1025 Then
SoundArray(i) = SoundArray(i + 1)
Else
SoundArray(i) = ""
End If
Next i
Else
SoundArray(SoundArrayCount) = ""
End If
SoundArrayCount = SoundArrayCount - 1
End Sub

Public Function N2G(Number) As MSFlexGrid
Select Case Number
Case 1
Set N2G = Abilities
Case 2
Set N2G = Animations
Case 3
Set N2G = Armor
Case 4
Set N2G = Buildings
Case 5
Set N2G = Images
Case 6
Set N2G = DataGrid
Case 7
Set N2G = Races
Case 8
Set N2G = Shields
Case 9
Set N2G = Sounds
Case 10
Set N2G = Upgrades
Case 11
Set N2G = Weapons
End Select
End Function

Public Function MaxRows()
Dim DG As MSFlexGrid, mR
mR = 0
For i = 1 To 11 Step 1
Set DG = N2G(i)
r = DG.Rows - 1
If mR < r Then mR = r
Next i
MaxRows = mR
End Function

Public Function MaxCols()
Dim DG As MSFlexGrid, mR
mR = 0
For i = 1 To 11 Step 1
Set DG = N2G(i)
r = DG.Cols - 1
If mR < r Then mR = r
Next i
MaxCols = mR
End Function

Public Function ChangePrefix(Text As String) As String
If Mid(Text, 2, 1) = "-" Then
ChangePrefix = Mid(Text, 1, 1)
ElseIf Mid(Text, 3, 1) = "-" Then
ChangePrefix = Mid(Text, 1, 2)
Else
ChangePrefix = Text
End If
End Function

Public Function BufferOut(Text) As String
Text = Replace(Text, "%", "%3")
Text = Replace(Text, Chr(180), "%0")
Text = Replace(Text, Chr(181), "%1")
Text = Replace(Text, Chr(182), "%2")
BufferOut = Text
End Function

Public Function SaveMod(FileName) As Boolean
'On Error GoTo ErrSH
Dim r, Rm, C, Cm, g, Gm, DG As MSFlexGrid, TStr As String
ReDim GridDataArray(11, MaxRows, MaxCols)
g = 1
Gm = 12
Do 'Until G >= Gm
Set DG = N2G(g)
r = 1
Rm = DG.Rows - 1
Do 'Until R >= Rm
C = 1
Cm = DG.Cols - 1
Do 'Until C >= Cm
If Left(DG.TextMatrix(r, C), Len("SoundArrayVariable")) = "SoundArrayVariable" Then
GridDataArray(g, r, C) = GetSound(DG.TextMatrix(r, C))
Else
GridDataArray(g, r, C) = BufferOut(DG.TextMatrix(r, C))
End If
C = C + 1
Loop Until C >= Cm ''''''''''''
r = r + 1
Loop Until r >= Rm ''''''''''''
g = g + 1
Loop Until g >= Gm ''''''''''''
'######################################################################################'
'########################################trtf BS#######################################'
'######################################################################################'
TRtf.Text = ""
PBar.Visible = True
a = 0
b = 0
For i = 1 To 11 Step 1
Set DG = N2G(i)
a = a + (DG.Rows - 1)
Next i
g = 1
Do 'Until G >= Gm
Set DG = N2G(g)
r = 1
Rm = DG.Rows
Do 'Until R >= Rm
C = 1
Cm = DG.Cols
Do 'Until C >= Cm
DoEvents
If Left(DG.TextMatrix(r, C), Len("SoundArrayVariable")) = "SoundArrayVariable" And LCase(Mid(DG.TextMatrix(0, C), 1, 5)) <> "[name" Then
TStr = TStr & BufferOut("SoundArrayVariable" & GetSound(DG.TextMatrix(r, C))) & Chr(180)
Else
TStr = TStr & BufferOut(DG.TextMatrix(r, C)) & Chr(180)
End If
C = C + 1
Loop Until C >= Cm
b = b + 1
PBar.Percent = (b / a) * 100
TStr = TStr & Chr(181)
r = r + 1
Loop Until r >= Rm
TStr = TStr & Chr(182)
g = g + 1
Loop Until g >= Gm
TRtf.Text = TStr
Open FileName For Output As 1
Print #1, TRtf.Text
Close 1
PBar.Visible = False
SaveMod = True
Exit Function
ErrSH:
PBar.Visible = False
SaveMod = False
End Function

Public Function BufferIn(Text) As String
Text = Replace(Text, "%2", Chr(182))
Text = Replace(Text, "%1", Chr(181))
Text = Replace(Text, "%0", Chr(180))
Text = Replace(Text, "%3", "%")
BufferIn = Text
End Function

Public Function ChrCount(StrIn, ToCount)
i = 1
j = 0
Do Until i >= Len(StrIn)
a = InStr(i, StrIn, ToCount, vbTextCompare)
If a <> 0 Then
j = j + 1
i = a + 1
End If
Loop
ChrCount = j
End Function

Public Function OpenMod(FileName) As Boolean
On Error GoTo NF
ClearSoundArray
If FileLen(FileName) < 1 Then
NF:
Msbox "The selected file does not exist!", "Invalid File"
OpenMod = False
Exit Function
End If
On Error GoTo ErrOH
Dim r, Rm, C, Cm, g, Gm, DG As MSFlexGrid, FileIn As String, mRows, mCols, GmR, GmC, tS, i, j, k, l, m
TRtf.LoadFile FileName
FileIn = TRtf.Text
TRtf.Text = ""

ReDim GArrayL2(10)

GArray = Split(FileIn, Chr(182))

For i = 0 To 10 Step 1
GArrayL2(i) = Split(GArray(i), Chr(181))
GArrayCount(i) = UBound(GArrayL2(i))
Next i

j = 0

For i = 0 To 10 Step 1
If GArrayCount(i) > j Then j = GArrayCount(i)
Next i

ReDim GArrayL3(10, j)
ReDim GArrayCountL3(10, j)

Y = 0

For i = 0 To 10 Step 1
Set DG = N2G(i + 1)
DG.Rows = 1
DG.Rows = GArrayCount(i) + 1
For j = 0 To GArrayCount(i) - 1 Step 1
GArrayL3(i, j) = Split(GArrayL2(i)(j), Chr(180))
Next j
Y = Y + (DG.Rows - 1)
Next i

z = 0
PBar.Visible = True
For i = 0 To 10 Step 1
Set DG = N2G(i + 1)
For j = 0 To GArrayCount(i) - 1 Step 1
For k = 0 To UBound(GArrayL3(i, j)) - 1 Step 1
a = BufferIn(GArrayL3(i, j)(k))
If Left(a, 18) = "SoundArrayVariable" And k > 0 Then
b = AddSound(Right(a, Len(a) - 18))
DG.TextMatrix(j + 1, k + 1) = "SoundArrayVariable" & b
Else
If k = 0 Then
DG.TextMatrix(j + 1, 0) = a
End If
DG.TextMatrix(j + 1, k + 1) = a
End If
DoEvents
Next k
z = z + 1
PBar.Percent = (z / Y) * 100
Next j
Next i

tS = ""
PBar.Visible = False
For i = 1 To 11 Step 1
aCount(i) = 0
Next i
OpenMod = True
Exit Function
ErrOH:
OpenMod = False
End Function

Public Sub subKeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then 'no added key
If KeyCode = 27 Then Label9_Click 'Esc
If KeyCode = 112 Then Label4_Click ' F1
If KeyCode = 123 Then Label8_Click 'F12
ElseIf Shift = 1 Then 'shift
ElseIf Shift = 2 Then 'ctrl
If KeyCode = 78 Then Label5_Click 'n
If KeyCode = 79 Then Label6_Click 'o
If KeyCode = 83 Then Label7_Click 's
If KeyCode = 65 Then Label3_Click 'a
If KeyCode = 82 Then Label16_Click 'r
ElseIf Shift = 4 Then 'alt
If KeyCode = 115 Then Label9_Click 'Alt+F4
End If
End Sub

Public Function GetTextWidth(sIn As String) As Integer
j = 0
k = 0
If Len(sIn) > 0 Then
For i = 1 To Len(sIn) Step 1
j = j + CInt(ChrList.List(Asc(Mid(sIn, i, 1))))
Next i
End If
GetTextWidth = (j \ 12)
End Function

Public Sub CreateMod()
ClearSoundArray
For i = 1 To 11 Step 1
aCount(i) = 0
Next i
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''---Abilities---'''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Abilities.Clear
Abilities.Cols = 10
Abilities.TextMatrix(0, 0) = "                                                                         "
Abilities.TextMatrix(0, 1) = "[Name]                                                       "
Abilities.TextMatrix(0, 2) = "Abilities Added          "
Abilities.TextMatrix(0, 3) = "Energy Required - Start          "
Abilities.TextMatrix(0, 4) = "Energy Required - Per Second          "
Abilities.TextMatrix(0, 5) = "Food Cost Added To Affected Units          "
Abilities.TextMatrix(0, 6) = "Mineral Cost          "
Abilities.TextMatrix(0, 7) = "Sub-Mineral Cost          "
Abilities.TextMatrix(0, 8) = "Time Cost          "
Abilities.TextMatrix(0, 9) = "Units Affected - When Researched          "
For i = 0 To 9 Step 1
Abilities.ColWidth(i) = GetTextWidth(Abilities.TextMatrix(0, i))
Next i
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''---Animations---'''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Animations.Clear
Animations.Cols = 19
Animations.TextMatrix(0, 0) = "                                                                       "
Animations.TextMatrix(0, 1) = "[Name]                                                       "
Animations.TextMatrix(0, 2) = "[Size]              " 'locked value
Animations.TextMatrix(0, 3) = "[Type]                                       "
Animations.TextMatrix(0, 4) = "FPS          "
Animations.TextMatrix(0, 5) = "Image 1          "
Animations.TextMatrix(0, 6) = "Image 2          "
Animations.TextMatrix(0, 7) = "Image 3          "
Animations.TextMatrix(0, 8) = "Image 4          "
Animations.TextMatrix(0, 9) = "Image 5          "
Animations.TextMatrix(0, 10) = "Image 6          "
Animations.TextMatrix(0, 11) = "Image 7          "
Animations.TextMatrix(0, 12) = "Image 8          "
Animations.TextMatrix(0, 13) = "Image 9          "
Animations.TextMatrix(0, 14) = "Image 10          "
Animations.TextMatrix(0, 15) = "Image 11          "
Animations.TextMatrix(0, 16) = "Image 12          "
Animations.TextMatrix(0, 17) = "Pause Time Between Playback          "
Animations.TextMatrix(0, 18) = "Random Frames          " 't/f
For i = 0 To 18 Step 1
Animations.ColWidth(i) = GetTextWidth(Animations.TextMatrix(0, i))
Next i
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''---Armor---''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Armor.Clear
Armor.Cols = 6
Armor.TextMatrix(0, 0) = "                                                                       "
Armor.TextMatrix(0, 1) = "[Name]                                                       "
Armor.TextMatrix(0, 2) = "Armor Effect          "
Armor.TextMatrix(0, 3) = "Armor HP          "
Armor.TextMatrix(0, 4) = "Defence          "
Armor.TextMatrix(0, 5) = "Effect Color          "
For i = 0 To 5 Step 1
Armor.ColWidth(i) = GetTextWidth(Armor.TextMatrix(0, i))
Next i
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''---Buildings---''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Buildings.Clear
Buildings.Cols = 97
Buildings.TextMatrix(0, 0) = "                                                                       "
Buildings.TextMatrix(0, 1) = "[Name]                                                       "
Buildings.TextMatrix(0, 2) = "Abilities          "
Buildings.TextMatrix(0, 3) = "Abilities Reseached At Building          "
Buildings.TextMatrix(0, 4) = "Ability Research Cancel          "
Buildings.TextMatrix(0, 5) = "Air --> Land Animation          "
Buildings.TextMatrix(0, 6) = "Air --> Land Sound          "
Buildings.TextMatrix(0, 7) = "Air --> Sea Animation          "
Buildings.TextMatrix(0, 8) = "Air --> Sea Sound          "
Buildings.TextMatrix(0, 9) = "Air Build Animation          "
Buildings.TextMatrix(0, 10) = "Air Cancel Animation          "
Buildings.TextMatrix(0, 11) = "Air Fire Animation Point          "
Buildings.TextMatrix(0, 12) = "Air Moveing Animation          "
Buildings.TextMatrix(0, 13) = "Air Normal Animation          "
Buildings.TextMatrix(0, 14) = "Air Sight          "
Buildings.TextMatrix(0, 15) = "Air Sound Building - 1          "
Buildings.TextMatrix(0, 16) = "Air Sound Building - 2          "
Buildings.TextMatrix(0, 17) = "Air Sound Normal - 1          "
Buildings.TextMatrix(0, 18) = "Air Sound Normal - 2          "
Buildings.TextMatrix(0, 19) = "Air Weapon          "
Buildings.TextMatrix(0, 20) = "Armor          "
Buildings.TextMatrix(0, 21) = "Burrows          "
Buildings.TextMatrix(0, 22) = "Cancel Construction After Start          "
Buildings.TextMatrix(0, 23) = "Cloaked          "
Buildings.TextMatrix(0, 24) = "Damage Mode          " 'none,fire,sparks,blood shooting,blood around base,glowing - red,glowing - blue, glowing - green
Buildings.TextMatrix(0, 25) = "Food Cost          "
Buildings.TextMatrix(0, 26) = "Fields Emitted          " 'cloak/power
Buildings.TextMatrix(0, 27) = "Floatable          "
Buildings.TextMatrix(0, 28) = "Flyable          "
Buildings.TextMatrix(0, 29) = "Function In Air          "
Buildings.TextMatrix(0, 30) = "Function In Sea          "
Buildings.TextMatrix(0, 31) = "Function On Land          "
Buildings.TextMatrix(0, 32) = "Hanger Unit          "
Buildings.TextMatrix(0, 33) = "Has Hanger          "
Buildings.TextMatrix(0, 34) = "Has Larva          "
Buildings.TextMatrix(0, 35) = "HP          "
Buildings.TextMatrix(0, 36) = "Land --> Air Animation          "
Buildings.TextMatrix(0, 37) = "Land --> Air Sound          "
Buildings.TextMatrix(0, 38) = "Land Build Animation          "
Buildings.TextMatrix(0, 39) = "Land Cancel Animation          "
Buildings.TextMatrix(0, 40) = "Land Fire Animation Point          "
Buildings.TextMatrix(0, 41) = "Land Moveing Animation          "
Buildings.TextMatrix(0, 42) = "Land Normal Animation          "
Buildings.TextMatrix(0, 43) = "Land Sight          "
Buildings.TextMatrix(0, 44) = "Land Sound Building - 1          "
Buildings.TextMatrix(0, 45) = "Land Sound Building - 2          "
Buildings.TextMatrix(0, 46) = "Land Sound Normal - 1          "
Buildings.TextMatrix(0, 47) = "Land Sound Normal - 2          "
Buildings.TextMatrix(0, 48) = "Land Weapon          "
Buildings.TextMatrix(0, 49) = "Larva Add Speed[In Seconds]          "
Buildings.TextMatrix(0, 50) = "Larva Cancel Animation          "
Buildings.TextMatrix(0, 51) = "Larva Cancel Mode          "
Buildings.TextMatrix(0, 52) = "Larva Morph Animation          "
Buildings.TextMatrix(0, 53) = "Larva Morph Done Animation          "
Buildings.TextMatrix(0, 54) = "Larva Transformation(s)          "
Buildings.TextMatrix(0, 55) = "Larva Unit          "
Buildings.TextMatrix(0, 56) = "Max Larva Units          "
Buildings.TextMatrix(0, 57) = "Max Hanger Units          "
Buildings.TextMatrix(0, 58) = "Mineral Cost          "
Buildings.TextMatrix(0, 59) = "Movable - Normal          "
Buildings.TextMatrix(0, 60) = "Power Cost          "
Buildings.TextMatrix(0, 61) = "Repair Animation          "
Buildings.TextMatrix(0, 62) = "Requires Power          "
Buildings.TextMatrix(0, 63) = "Rollable          "
Buildings.TextMatrix(0, 64) = "Sea --> Air Animation          "
Buildings.TextMatrix(0, 65) = "Sea --> Air Sound          "
Buildings.TextMatrix(0, 66) = "Sea Build Animation          "
Buildings.TextMatrix(0, 67) = "Sea Cancel Animation          "
Buildings.TextMatrix(0, 68) = "Sea Fire Animation Point          "
Buildings.TextMatrix(0, 69) = "Sea Moveing Animation          "
Buildings.TextMatrix(0, 70) = "Sea Normal Animation          "
Buildings.TextMatrix(0, 71) = "Sea Sight          "
Buildings.TextMatrix(0, 72) = "Sea Sound Building - 1          "
Buildings.TextMatrix(0, 73) = "Sea Sound Building - 2          "
Buildings.TextMatrix(0, 74) = "Sea Sound Normal - 1          "
Buildings.TextMatrix(0, 75) = "Sea Sound Normal - 2          "
Buildings.TextMatrix(0, 76) = "Sea Weapon          "
Buildings.TextMatrix(0, 77) = "Shield Name          "
Buildings.TextMatrix(0, 78) = "Start Larva Amount          "
Buildings.TextMatrix(0, 79) = "Start Amount In Hanger          "
Buildings.TextMatrix(0, 80) = "Sub-Mineral Cost          "
Buildings.TextMatrix(0, 81) = "Time Cost          "
Buildings.TextMatrix(0, 82) = "Unit Construction Cancel          "
Buildings.TextMatrix(0, 83) = "Unit Energy          "
Buildings.TextMatrix(0, 84) = "Upgrades Researched At Building          "
Buildings.TextMatrix(0, 85) = "Upgrade To Building - 1          "
Buildings.TextMatrix(0, 86) = "Upgrade To Building - 2          "
Buildings.TextMatrix(0, 87) = "Upgrade To Building - 3          "
Buildings.TextMatrix(0, 88) = "Wire Frame Image          "
Buildings.TextMatrix(0, 89) = "{AddOns}          "
Buildings.TextMatrix(0, 90) = "{Is AddOn}          "
Buildings.TextMatrix(0, 91) = "{Morphs}          "
Buildings.TextMatrix(0, 92) = "{Race}          "
Buildings.TextMatrix(0, 93) = "{Required Buildings}          "
Buildings.TextMatrix(0, 94) = "{Stores Units}          "
Buildings.TextMatrix(0, 95) = "{Teleporter}          "
Buildings.TextMatrix(0, 96) = "{Teleport Exit}          "
For i = 0 To 96 Step 1
Buildings.ColWidth(i) = GetTextWidth(Buildings.TextMatrix(0, i))
Next i
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''---Images---'''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Images.Clear
Images.Cols = 3
Images.TextMatrix(0, 0) = "                                                                       "
Images.TextMatrix(0, 1) = "[Name]                                                       "
Images.TextMatrix(0, 2) = "Location                                               "
For i = 0 To 2 Step 1
Images.ColWidth(i) = GetTextWidth(Images.TextMatrix(0, i))
Next i
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''---Infantry---''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
DataGrid.Clear
DataGrid.Cols = 134
DataGrid.TextMatrix(0, 0) = "                                                                       "
DataGrid.TextMatrix(0, 1) = "[Name]                                                       "
DataGrid.TextMatrix(0, 2) = "[Name - Mode 2 - Leave Blank For Single Mode]     "
DataGrid.TextMatrix(0, 3) = "Abilities - Mode 1          "
DataGrid.TextMatrix(0, 4) = "Abilities - Mode 2          "
DataGrid.TextMatrix(0, 5) = "Air Fire Animation Point - Mode 1          " 'x,y
DataGrid.TextMatrix(0, 6) = "Air Fire Animation Point - Mode 2          " 'x,y
DataGrid.TextMatrix(0, 7) = "Air Repair Range - Mode 1          "
DataGrid.TextMatrix(0, 8) = "Air Repair Range - Mode 2          "
DataGrid.TextMatrix(0, 9) = "Air Sight - Mode 1          "
DataGrid.TextMatrix(0, 10) = "Air Sight - Mode 2          "
DataGrid.TextMatrix(0, 11) = "Air Weapon - Mode 1          "
DataGrid.TextMatrix(0, 12) = "Air Weapon - Mode 2          "
DataGrid.TextMatrix(0, 13) = "Armor - Mode 1          "
DataGrid.TextMatrix(0, 14) = "Armor - Mode 2          "
DataGrid.TextMatrix(0, 15) = "Auto-Repair - Mode 1          "
DataGrid.TextMatrix(0, 16) = "Auto-Repair - Mode 2          "
DataGrid.TextMatrix(0, 17) = "Build Location 1[Optional If Hanger Or Larva Unit]          "
DataGrid.TextMatrix(0, 18) = "Build Location 2[Optional]          "
DataGrid.TextMatrix(0, 19) = "Build Location 3[Optional]          "
DataGrid.TextMatrix(0, 20) = "Build Method - Mode 1          " 'start and leave/stay in spot/stay and move/morph/none
DataGrid.TextMatrix(0, 21) = "Build Method - Mode 2          " 'start and leave/stay in spot/stay and move/morph/none
DataGrid.TextMatrix(0, 22) = "Cloaked - Mode 1          "
DataGrid.TextMatrix(0, 23) = "Cloaked - Mode 2          "
DataGrid.TextMatrix(0, 24) = "Covers Area Of Unit Image - Mode 1          " 'x1,y1,x2,y2
DataGrid.TextMatrix(0, 25) = "Covers Area Of Unit Image - Mode 2          " 'x1,y1,x2,y2
DataGrid.TextMatrix(0, 26) = "Death Animation 1 - Mode 1          "
DataGrid.TextMatrix(0, 27) = "Death Animation 1 - Mode 2          "
DataGrid.TextMatrix(0, 28) = "Death Animation 2 - Mode 1          "
DataGrid.TextMatrix(0, 29) = "Death Animation 2 - Mode 2          "
DataGrid.TextMatrix(0, 30) = "Death Animation 3 - Mode 1          "
DataGrid.TextMatrix(0, 31) = "Death Animation 3 - Mode 2          "
DataGrid.TextMatrix(0, 32) = "Death Sound - Mode 1          "
DataGrid.TextMatrix(0, 33) = "Death Sound - Mode 2          "
DataGrid.TextMatrix(0, 34) = "Fields Emitted - Mode 1          " 'cloak/power
DataGrid.TextMatrix(0, 35) = "Fields Emitted - Mode 2          " 'cloak/power
DataGrid.TextMatrix(0, 36) = "Hanger - Mode 1          "
DataGrid.TextMatrix(0, 37) = "Hanger - Mode 2          "
DataGrid.TextMatrix(0, 38) = "Hanger Max - Mode 1          "
DataGrid.TextMatrix(0, 39) = "Hanger Max - Mode 2          "
DataGrid.TextMatrix(0, 40) = "Hanger Start - Mode 1          "
DataGrid.TextMatrix(0, 41) = "Hanger Start - Mode 2          "
DataGrid.TextMatrix(0, 42) = "Hanger Unit - Mode 1          "
DataGrid.TextMatrix(0, 43) = "Hanger Unit - Mode 2          "
DataGrid.TextMatrix(0, 44) = "HP - Mode 1          "
DataGrid.TextMatrix(0, 45) = "HP - Mode 2          "
DataGrid.TextMatrix(0, 46) = "Infantry Type - Mode 1          " 'Infantry/Worker/Hanger Unit/Larva
DataGrid.TextMatrix(0, 47) = "Infantry Type - Mode 2          " 'Infantry/Worker/Hanger Unit/Larva
DataGrid.TextMatrix(0, 48) = "Land Fire Animation Point - Mode 1          " 'x,y
DataGrid.TextMatrix(0, 49) = "Land Fire Animation Point - Mode 2          " 'x,y
DataGrid.TextMatrix(0, 50) = "Land Repair Range - Mode 1          "
DataGrid.TextMatrix(0, 51) = "Land Repair Range - Mode 2          "
DataGrid.TextMatrix(0, 52) = "Land Sight - Mode 1          "
DataGrid.TextMatrix(0, 53) = "Land Sight - Mode 2          "
DataGrid.TextMatrix(0, 54) = "Land Weapon - Mode 1          "
DataGrid.TextMatrix(0, 55) = "Land Weapon - Mode 2          "
DataGrid.TextMatrix(0, 56) = "Mode 1 --> 2 Animation          "
DataGrid.TextMatrix(0, 57) = "Mode 1 --> 2 Sound          "
DataGrid.TextMatrix(0, 58) = "Mode 1 - Food Cost          "
DataGrid.TextMatrix(0, 59) = "Mode 1 - Mineral Cost          "
DataGrid.TextMatrix(0, 60) = "Mode 1 - Power Cost          "
DataGrid.TextMatrix(0, 61) = "Mode 1 - Sub-Mineral Cost          "
DataGrid.TextMatrix(0, 62) = "Mode 1 - Time Cost          "
DataGrid.TextMatrix(0, 63) = "Mode 1 Icon          "
DataGrid.TextMatrix(0, 64) = "Mode 2 --> 1 Animation          "
DataGrid.TextMatrix(0, 65) = "Mode 2 --> 1 Sound          "
DataGrid.TextMatrix(0, 66) = "Mode 2 - Food Cost          "
DataGrid.TextMatrix(0, 67) = "Mode 2 - Mineral Cost          "
DataGrid.TextMatrix(0, 68) = "Mode 2 - Power Cost          "
DataGrid.TextMatrix(0, 69) = "Mode 2 - Sub-Mineral Cost          "
DataGrid.TextMatrix(0, 70) = "Mode 2 - Time Cost          "
DataGrid.TextMatrix(0, 71) = "Mode 2 Icon          "
DataGrid.TextMatrix(0, 72) = "Mode 2 Upgrade Location 1          "
DataGrid.TextMatrix(0, 73) = "Mode 2 Upgrade Location 2[Optional]          "
DataGrid.TextMatrix(0, 74) = "Mode 2 Upgrade Location 3[Optional]          "
DataGrid.TextMatrix(0, 75) = "Mode Switch          " '1-->2|1<->2
DataGrid.TextMatrix(0, 76) = "Repair/Heal Animation - Mode 1          "
DataGrid.TextMatrix(0, 77) = "Repair/Heal Animation - Mode 2          "
DataGrid.TextMatrix(0, 78) = "Repairs Type - Mode 1          " 'Bio/Energy/Mech
DataGrid.TextMatrix(0, 79) = "Repairs Type - Mode 2          " 'Bio/Energy/Mech
DataGrid.TextMatrix(0, 80) = "Sea Fire Animation Point - Mode 1          " 'x,y
DataGrid.TextMatrix(0, 81) = "Sea Fire Animation Point - Mode 2          " 'x,y
DataGrid.TextMatrix(0, 82) = "Sea Repair Range - Mode 1          "
DataGrid.TextMatrix(0, 83) = "Sea Repair Range - Mode 2          "
DataGrid.TextMatrix(0, 84) = "Sea Sight - Mode 1          "
DataGrid.TextMatrix(0, 85) = "Sea Sight - Mode 2          "
DataGrid.TextMatrix(0, 86) = "Sea Weapon - Mode 1          "
DataGrid.TextMatrix(0, 87) = "Sea Weapon - Mode 2          "
DataGrid.TextMatrix(0, 88) = "Shield - Mode 1          "
DataGrid.TextMatrix(0, 89) = "Shield - Mode 2          "
DataGrid.TextMatrix(0, 90) = "Sound Acknowledge - 1 - Mode 1          "
DataGrid.TextMatrix(0, 91) = "Sound Acknowledge - 1 - Mode 2          "
DataGrid.TextMatrix(0, 92) = "Sound Acknowledge - 2 - Mode 1          "
DataGrid.TextMatrix(0, 93) = "Sound Acknowledge - 2 - Mode 2          "
DataGrid.TextMatrix(0, 94) = "Sound Acknowledge - 3 - Mode 1          "
DataGrid.TextMatrix(0, 95) = "Sound Acknowledge - 3 - Mode 2          "
DataGrid.TextMatrix(0, 96) = "Sound Acknowledge - 4 - Mode 1          "
DataGrid.TextMatrix(0, 97) = "Sound Acknowledge - 4 - Mode 2          "
DataGrid.TextMatrix(0, 98) = "Sound Acknowledge - 5 - Mode 1          "
DataGrid.TextMatrix(0, 99) = "Sound Acknowledge - 5 - Mode 2          "
DataGrid.TextMatrix(0, 100) = "Still Animation - Mode 1          "
DataGrid.TextMatrix(0, 101) = "Still Animation - Mode 2          "
DataGrid.TextMatrix(0, 102) = "Terrain - Mode 1          "
DataGrid.TextMatrix(0, 103) = "Terrain - Mode 2          "
DataGrid.TextMatrix(0, 104) = "Unit Energy - Mode 1          "
DataGrid.TextMatrix(0, 105) = "Unit Energy - Mode 2          "
DataGrid.TextMatrix(0, 106) = "Unit Morphs - Mode 1          "
DataGrid.TextMatrix(0, 107) = "Unit Morphs - Mode 2          "
DataGrid.TextMatrix(0, 108) = "Unit Morph Enabled - Mode 1          "
DataGrid.TextMatrix(0, 109) = "Unit Morph Enabled - Mode 2          "
DataGrid.TextMatrix(0, 110) = "Unit Morph Finish Animation - Mode 1          "
DataGrid.TextMatrix(0, 111) = "Unit Morph Finish Animation - Mode 2          "
DataGrid.TextMatrix(0, 112) = "Unit Morph Start Animation - Mode 1          "
DataGrid.TextMatrix(0, 113) = "Unit Morph Start Animation - Mode 2          "
DataGrid.TextMatrix(0, 114) = "Unit Type - Mode 1          " 'Bio/Mech
DataGrid.TextMatrix(0, 115) = "Unit Type - Mode 2          " 'Bio/Mech
DataGrid.TextMatrix(0, 116) = "Walk Animation - Mode 1          "
DataGrid.TextMatrix(0, 117) = "Walk Animation - Mode 2          "
DataGrid.TextMatrix(0, 118) = "Wire Frame Image - Mode 1          "
DataGrid.TextMatrix(0, 119) = "Wire Frame Image - Mode 2          "
DataGrid.TextMatrix(0, 120) = "{Builds Mines - Mode 1}          "
DataGrid.TextMatrix(0, 121) = "{Builds Mines - Mode 2}          "
DataGrid.TextMatrix(0, 122) = "{Mine - Mode 1}          "
DataGrid.TextMatrix(0, 123) = "{Mine - Mode 2}          "
DataGrid.TextMatrix(0, 124) = "{Mine Payload - Mode 1}          "
DataGrid.TextMatrix(0, 125) = "{Mine Payload - Mode 2}          "
DataGrid.TextMatrix(0, 126) = "(Race)          "
DataGrid.TextMatrix(0, 127) = "{Required Buildings}          "
DataGrid.TextMatrix(0, 128) = "{Unit Attacks Burrowed}          "
DataGrid.TextMatrix(0, 129) = "{Unit Attacks Sunken}          "
DataGrid.TextMatrix(0, 130) = "{Unit Burrowable}          "
DataGrid.TextMatrix(0, 131) = "{Unit Is Detector}          "
DataGrid.TextMatrix(0, 132) = "{Unit Responds Over Distance}          "
DataGrid.TextMatrix(0, 133) = "{Unit Sinkable}          "
For i = 0 To 133 Step 1
DataGrid.ColWidth(i) = GetTextWidth(DataGrid.TextMatrix(0, i))
Next i
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''---Races---'''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Races.Clear
Races.Cols = 9
Races.TextMatrix(0, 0) = "                                                                       "
Races.TextMatrix(0, 1) = "[Name]                                                       "
Races.TextMatrix(0, 2) = "Cost Factor - Food          "
Races.TextMatrix(0, 3) = "Cost Factor - Minerals          "
Races.TextMatrix(0, 4) = "Cost Factor - Sub-Minerals          "
Races.TextMatrix(0, 5) = "Cost Factor - Time          "
Races.TextMatrix(0, 6) = "Max Food          "
Races.TextMatrix(0, 7) = "Starting Building          "
Races.TextMatrix(0, 8) = "Starting Worker          "
For i = 0 To 8 Step 1
Races.ColWidth(i) = GetTextWidth(Races.TextMatrix(0, i))
Next i
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''---Shields---'''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Shields.Clear
Shields.Cols = 6
Shields.TextMatrix(0, 0) = "                                                                       "
Shields.TextMatrix(0, 1) = "[Name]                                                       "
Shields.TextMatrix(0, 2) = "Defence          "
Shields.TextMatrix(0, 3) = "Effect Color          "
Shields.TextMatrix(0, 4) = "Shield Effect          "
Shields.TextMatrix(0, 5) = "Shield HP          "
For i = 0 To 5 Step 1
Shields.ColWidth(i) = GetTextWidth(Shields.TextMatrix(0, i))
Next i
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''---Sounds---'''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sounds.Clear
Sounds.Cols = 4
Sounds.TextMatrix(0, 0) = "                                                                       "
Sounds.TextMatrix(0, 1) = "[Name]                                                       "
Sounds.TextMatrix(0, 2) = "Location                                               "
Sounds.TextMatrix(0, 3) = "Pause Time In Seconds Between Playback[-1 To Play Once]          "
For i = 0 To 3 Step 1
Sounds.ColWidth(i) = GetTextWidth(Sounds.TextMatrix(0, i))
Next i
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''---Upgrades---'''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Upgrades.Clear
Upgrades.Cols = 10
Upgrades.TextMatrix(0, 0) = "                                                                       "
Upgrades.TextMatrix(0, 1) = "[Name]                                                       "
Upgrades.TextMatrix(0, 2) = "Amount Added To Properties          "
Upgrades.TextMatrix(0, 3) = "Food Cost Added To Affected Units          "
Upgrades.TextMatrix(0, 4) = "Mineral Cost          "
Upgrades.TextMatrix(0, 5) = "Max Upgrades          "
Upgrades.TextMatrix(0, 6) = "Properties Upgraded          "
Upgrades.TextMatrix(0, 7) = "Sub-Mineral Cost          "
Upgrades.TextMatrix(0, 8) = "Time Cost          "
Upgrades.TextMatrix(0, 9) = "Units Affected          "
For i = 0 To 9 Step 1
Upgrades.ColWidth(i) = GetTextWidth(Upgrades.TextMatrix(0, i))
Next i
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''---Weapons---'''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Weapons.Clear
Weapons.Cols = 40
Weapons.TextMatrix(0, 0) = "                                                                       "
Weapons.TextMatrix(0, 1) = "[Name]                                                       "
Weapons.TextMatrix(0, 2) = "Air Animation          "
Weapons.TextMatrix(0, 3) = "Air End Animation          "
Weapons.TextMatrix(0, 4) = "Air Damage          "
Weapons.TextMatrix(0, 5) = "Air Fire Animation          "
Weapons.TextMatrix(0, 6) = "Air Fireing Sound          "
Weapons.TextMatrix(0, 7) = "Air Fireing Speed          "
Weapons.TextMatrix(0, 8) = "Air Range          "
Weapons.TextMatrix(0, 9) = "Air Sight          "
Weapons.TextMatrix(0, 10) = "Air Speed          "
Weapons.TextMatrix(0, 11) = "Air Splash Damage          "
Weapons.TextMatrix(0, 12) = "Air Splash Damage Range          "
Weapons.TextMatrix(0, 13) = "Animation Direction          " 'still,to target,opposite of target,up,down,left,right,topleft,topright,bottomleft,bottomright,left of target,right of target,up then down on unit,down then up on unit,left then right on unit,right then left on unit
Weapons.TextMatrix(0, 14) = "Land Animation          "
Weapons.TextMatrix(0, 15) = "Land Damage          "
Weapons.TextMatrix(0, 16) = "Land End Animation          "
Weapons.TextMatrix(0, 17) = "Land Fire Animation          "
Weapons.TextMatrix(0, 18) = "Land Fireing Sound          "
Weapons.TextMatrix(0, 19) = "Land Fireing Speed          "
Weapons.TextMatrix(0, 20) = "Land Range          "
Weapons.TextMatrix(0, 21) = "Land Sight          "
Weapons.TextMatrix(0, 22) = "Land Speed          "
Weapons.TextMatrix(0, 23) = "Land Splash Damage          "
Weapons.TextMatrix(0, 24) = "Land Splash Damage Range          "
Weapons.TextMatrix(0, 25) = "Sea Animation          "
Weapons.TextMatrix(0, 26) = "Sea Damage          "
Weapons.TextMatrix(0, 27) = "Sea End Animation          "
Weapons.TextMatrix(0, 28) = "Sea Fire Animation          "
Weapons.TextMatrix(0, 29) = "Sea Fireing Sound          "
Weapons.TextMatrix(0, 30) = "Sea Fireing Speed          "
Weapons.TextMatrix(0, 31) = "Sea Range          "
Weapons.TextMatrix(0, 32) = "Sea Sight          "
Weapons.TextMatrix(0, 33) = "Sea Speed          "
Weapons.TextMatrix(0, 34) = "Sea Splash Damage          "
Weapons.TextMatrix(0, 35) = "Sea Splash Damage Range          "
Weapons.TextMatrix(0, 36) = "Splash Damage Animation          "
Weapons.TextMatrix(0, 37) = "Splash Damage Mode          "
Weapons.TextMatrix(0, 38) = "Travels Over Terrain          "
Weapons.TextMatrix(0, 39) = "Weapon Type          "
For i = 0 To 39 Step 1
Weapons.ColWidth(i) = GetTextWidth(Weapons.TextMatrix(0, i))
Next i
End Sub

Private Sub Abilities_Click()
Abilities_EnterCell
End Sub

Private Sub Abilities_EnterCell()
DGN = 11
Timer2.Enabled = True
End Sub

Private Sub Abilities_KeyDown(KeyCode As Integer, Shift As Integer)
subKeyDown KeyCode, Shift
End Sub

Private Sub Abilities_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If HF = True Then Label2_Click
If FF = True Then Label1_Click
If DF = True Then Label10_Click
End Sub

Private Sub Abilities_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
DataGridUpdate True, 1
End Sub

Private Sub Abilities_Scroll()
DataGridUpdate True, 1
End Sub

Private Sub Animations_Click()
Animations_EnterCell
End Sub

Private Sub Animations_EnterCell()
DGN = 12
Timer2.Enabled = True
End Sub

Private Sub Animations_KeyDown(KeyCode As Integer, Shift As Integer)
subKeyDown KeyCode, Shift
End Sub

Private Sub Animations_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If HF = True Then Label2_Click
If FF = True Then Label1_Click
If DF = True Then Label10_Click
End Sub

Private Sub Animations_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
DataGridUpdate True, 2
End Sub

Private Sub Animations_Scroll()
DataGridUpdate True, 2
End Sub

Private Sub Armor_Click()
Armor_EnterCell
End Sub

Private Sub Armor_EnterCell()
DGN = 13
Timer2.Enabled = True
End Sub

Private Sub Armor_KeyDown(KeyCode As Integer, Shift As Integer)
subKeyDown KeyCode, Shift
End Sub

Private Sub Armor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If HF = True Then Label2_Click
If FF = True Then Label1_Click
If DF = True Then Label10_Click
End Sub

Private Sub Armor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
DataGridUpdate True, 3
End Sub

Private Sub Armor_Scroll()
DataGridUpdate True, 3
End Sub

Private Sub Buildings_Click()
Buildings_EnterCell
End Sub

Private Sub Buildings_EnterCell()
DGN = 14
Timer2.Enabled = True
End Sub

Private Sub Buildings_KeyDown(KeyCode As Integer, Shift As Integer)
subKeyDown KeyCode, Shift
End Sub

Private Sub Buildings_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If HF = True Then Label2_Click
If FF = True Then Label1_Click
If DF = True Then Label10_Click
End Sub

Private Sub Buildings_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
DataGridUpdate True, 4
End Sub

Private Sub Buildings_Scroll()
DataGridUpdate True, 4
End Sub

Private Sub Closep_Click()
Unload Form1
End Sub

Private Sub Closep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set CloseP.Picture = CloseD.Picture
End Sub

Private Sub Closep_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set CloseP.Picture = CloseU.Picture
End Sub

Private Sub DataGrid_Click()
DataGrid_EnterCell
End Sub

Private Sub DataGrid_EnterCell()
DGN = 16
Timer2.Enabled = True
End Sub

Public Sub PosTextEdit(X, Y, X2, Y2, Text, MaxChars, Number As Boolean)
TextEdit.Left = X
TextEdit.Top = Y
TextEdit.Width = X2 - 1
TextEdit.Height = Y2 - 1
If TextEdit.Width + TextEdit.Left > CControl.Left + CControl.Width Then Exit Sub
If GScrolling = True Then Exit Sub

TextEdit.MaxLength = MaxChars

TextEdit.Visible = True
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = False

TextIsNumber = Number

TextEdit.Text = CControl.TextMatrix(CControl.Row, CControl.Col)

TextEdit.SetFocus
End Sub

Public Sub PosImageFileSelect(X, Y, X2, Y2, Text)
PlayView.Left = X
PlayView.Top = Y
PlayView.Height = Y2 - 1

FileSelect.Left = X2 + X - FileSelect.Width
FileSelect.Top = Y
FileSelect.Height = Y2 - 1
If FileSelect.Width + FileSelect.Left > CControl.Left + CControl.Width Then Exit Sub

If GScrolling = True Then Exit Sub

CD.Filter = "Image Files(*.bmp)|*.bmp"

TextEdit.Visible = False
FileSelect.Visible = True: PlayView.Visible = True
LBox.Visible = False
FileSelect.SetFocus
End Sub

Public Sub PosSoundFileSelect(X, Y, X2, Y2, Text)
FileSelect.Left = X2 + X - FileSelect.Width
FileSelect.Top = Y
FileSelect.Height = Y2 - 1
If FileSelect.Width + FileSelect.Left > CControl.Left + CControl.Width Then Exit Sub

If GScrolling = True Then Exit Sub

CD.Filter = "Sound File(*.wav)|*.wav"

TextEdit.Visible = False
FileSelect.Visible = True
LBox.Visible = False
FileSelect.SetFocus
End Sub

Public Sub PosLBox(X, Y, X2, Y2, Text)
LBox.Left = X
LBox.Top = Y
LBox.Width = X2
If LBox.Width + LBox.Left > CControl.Left + CControl.Width Then Exit Sub

If GScrolling = True Then Exit Sub

TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = True

LBox.SetFocus
End Sub

Public Sub PosImageSelect(X, Y, X2, Y2, Text)
LBox.Left = X
LBox.Top = Y
LBox.Width = X2
If LBox.Width + LBox.Left > CControl.Left + CControl.Width Then Exit Sub

If GScrolling = True Then Exit Sub

TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = True

LBox.Clear

LBox.AddItem "[none]"

For i = 1 To Images.Rows - 1
LBox.AddItem Images.TextMatrix(i, 1)
Next i

Hit = False

For i = 0 To LBox.ListCount - 1 Step 1
If LBox.List(i) = Text Then
Hit = True
Exit For
End If
Next i

If Hit = True Then
LBox.ListIndex = i
Else
LBox.ListIndex = 0
End If

LBox.SetFocus
End Sub

Public Sub PosTrueFalse(X, Y, X2, Y2, Text)
LBox.Left = X
LBox.Top = Y
LBox.Width = X2
If LBox.Width + LBox.Left > CControl.Left + CControl.Width Then Exit Sub

If GScrolling = True Then Exit Sub

TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = True

LBox.Clear

LBox.AddItem "0-False"
LBox.AddItem "1-True"

Hit = False

For i = 0 To LBox.ListCount - 1 Step 1
If LBox.List(i) = Text Then
Hit = True
Exit For
End If
Next i

If Hit = True Then
LBox.ListIndex = i
Else
LBox.ListIndex = 0
End If

LBox.SetFocus
End Sub

Public Sub PosModeSwitchSelect(X, Y, X2, Y2, Text)
LBox.Left = X
LBox.Top = Y
LBox.Width = X2
If LBox.Width + LBox.Left > CControl.Left + CControl.Width Then Exit Sub

If GScrolling = True Then Exit Sub

TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = True

LBox.Clear

LBox.AddItem "0-1-->2"
LBox.AddItem "1-1<->2"

Hit = False

For i = 0 To LBox.ListCount - 1 Step 1
If LBox.List(i) = Text Then
Hit = True
Exit For
End If
Next i

If Hit = True Then
LBox.ListIndex = i
Else
LBox.ListIndex = 0
End If

LBox.SetFocus
End Sub

Public Sub PosAnimationType(X, Y, X2, Y2, Text)
LBox.Left = X
LBox.Top = Y
LBox.Width = X2
If LBox.Width + LBox.Left > CControl.Left + CControl.Width Then Exit Sub

If GScrolling = True Then Exit Sub

TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = True

LBox.Clear

LBox.AddItem "0-Building Build"
LBox.AddItem "1-Building Cancel"
LBox.AddItem "2-Building Change Terrain Type"
LBox.AddItem "3-Building Moveing"
LBox.AddItem "4-Building Normal"
LBox.AddItem "5-Building Repair/Heal Animation"
LBox.AddItem "6-Larva Cancel"
LBox.AddItem "7-Larva Done"
LBox.AddItem "8-Larva Morph"
LBox.AddItem "9-Splash Damage"
LBox.AddItem "10-Unit Death"
LBox.AddItem "11-Unit Mode Switch"
LBox.AddItem "12-Unit Repair/Heal Animation"
LBox.AddItem "13-Unit Still"
LBox.AddItem "14-Unit Walk"
LBox.AddItem "15-Weapon Move"
LBox.AddItem "16-Weapon Fire"
LBox.AddItem "17-Weapon End"

Hit = False

For i = 0 To LBox.ListCount - 1 Step 1
If LBox.List(i) = Text Then
Hit = True
Exit For
End If
Next i

If Hit = True Then
LBox.ListIndex = i
Else
LBox.ListIndex = 0
End If

LBox.SetFocus
End Sub

Public Sub PosBuildMethodSelect(X, Y, X2, Y2, Text)
LBox.Left = X
LBox.Top = Y
LBox.Width = X2
If LBox.Width + LBox.Left > CControl.Left + CControl.Width Then Exit Sub

If GScrolling = True Then Exit Sub

TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = True

LBox.Clear

LBox.AddItem "0-[none]"
LBox.AddItem "1-Morph"
LBox.AddItem "2-Start And Leave"
LBox.AddItem "3-Stay And Move"
LBox.AddItem "4-Stay In Spot"

Hit = False

For i = 0 To LBox.ListCount - 1 Step 1
If LBox.List(i) = Text Then
Hit = True
Exit For
End If
Next i

If Hit = True Then
LBox.ListIndex = i
Else
LBox.ListIndex = 0
End If

LBox.SetFocus
End Sub

Public Sub PosDamageMode(X, Y, X2, Y2, Text)
LBox.Left = X
LBox.Top = Y
LBox.Width = X2
If LBox.Width + LBox.Left > CControl.Left + CControl.Width Then Exit Sub

If GScrolling = True Then Exit Sub

TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = True

LBox.Clear

LBox.AddItem "0-[none]"
LBox.AddItem "1-Blood Around Base"
LBox.AddItem "2-Blood Shooting"
LBox.AddItem "3-Fire"
LBox.AddItem "4-Glowing - Black"
LBox.AddItem "5-Glowing - Blue"
LBox.AddItem "6-Glowing - Brown"
LBox.AddItem "7-Glowing - Cyan"
LBox.AddItem "8-Glowing - Green"
LBox.AddItem "9-Glowing - Grey"
LBox.AddItem "10-Glowing - Orange"
LBox.AddItem "11-Glowing - Purple"
LBox.AddItem "12-Glowing - Red"
LBox.AddItem "13-Glowing - Teal"
LBox.AddItem "14-Glowing - White"
LBox.AddItem "15-Glowing - Yellow"
LBox.AddItem "16-Sparks"

Hit = False

For i = 0 To LBox.ListCount - 1 Step 1
If LBox.List(i) = Text Then
Hit = True
Exit For
End If
Next i

If Hit = True Then
LBox.ListIndex = i
Else
LBox.ListIndex = 0
End If

LBox.SetFocus
End Sub

Public Sub PosArmorShieldEffect(X, Y, X2, Y2, Text)
LBox.Left = X
LBox.Top = Y
LBox.Width = X2
If LBox.Width + LBox.Left > CControl.Left + CControl.Width Then Exit Sub

If GScrolling = True Then Exit Sub

TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = True

LBox.Clear

LBox.AddItem "0-Cloak Apperance"
LBox.AddItem "1-Energy Field"
LBox.AddItem "2-Energy Field - Constant"
LBox.AddItem "3-Fire"
LBox.AddItem "4-Fire - Constant"
LBox.AddItem "5-Flame"
LBox.AddItem "6-Flame - Constant"
LBox.AddItem "7-Phase"
LBox.AddItem "8-Reflect"
LBox.AddItem "9-Refract"
LBox.AddItem "10-Sparks"
LBox.AddItem "11-Sparks - Constant"

Hit = False

For i = 0 To LBox.ListCount - 1 Step 1
If LBox.List(i) = Text Then
Hit = True
Exit For
End If
Next i

If Hit = True Then
LBox.ListIndex = i
Else
LBox.ListIndex = 0
End If

LBox.SetFocus
End Sub

Public Sub PosColorSelect(X, Y, X2, Y2, Text)
LBox.Left = X
LBox.Top = Y
LBox.Width = X2
If LBox.Width + LBox.Left > CControl.Left + CControl.Width Then Exit Sub

If GScrolling = True Then Exit Sub

TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = True

LBox.Clear

LBox.AddItem "0-Black"
LBox.AddItem "1-Blue"
LBox.AddItem "2-Brown"
LBox.AddItem "3-Cyan"
LBox.AddItem "4-Green"
LBox.AddItem "5-Grey"
LBox.AddItem "6-Orange"
LBox.AddItem "7-Red"
LBox.AddItem "8-Teal"
LBox.AddItem "9-Purple"
LBox.AddItem "10-White"
LBox.AddItem "11-Yellow"

Hit = False

For i = 0 To LBox.ListCount - 1 Step 1
If LBox.List(i) = Text Then
Hit = True
Exit For
End If
Next i

If Hit = True Then
LBox.ListIndex = i
Else
LBox.ListIndex = 0
End If

LBox.SetFocus
End Sub

Public Sub PosTerrainSelect(X, Y, X2, Y2, Text)
LBox.Left = X
LBox.Top = Y
LBox.Width = X2
If LBox.Width + LBox.Left > CControl.Left + CControl.Width Then Exit Sub

If GScrolling = True Then Exit Sub

TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = True

LBox.Clear

LBox.AddItem "0-Air"
LBox.AddItem "1-Air-Land"
LBox.AddItem "2-Air-Land-Sea"
LBox.AddItem "3-Air-Sea"
LBox.AddItem "4-Land"
LBox.AddItem "5-Land-Sea"
LBox.AddItem "6-Sea"
LBox.AddItem "7-Sea Coast"

Hit = False

For i = 0 To LBox.ListCount - 1 Step 1
If LBox.List(i) = Text Then
Hit = True
Exit For
End If
Next i

If Hit = True Then
LBox.ListIndex = i
Else
LBox.ListIndex = 0
End If

LBox.SetFocus
End Sub

Public Sub PosSplashDamageModeSelect(X, Y, X2, Y2, Text)
LBox.Left = X
LBox.Top = Y
LBox.Width = X2
If LBox.Width + LBox.Left > CControl.Left + CControl.Width Then Exit Sub

If GScrolling = True Then Exit Sub

TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = True

LBox.Clear

LBox.AddItem "0-[none]"
LBox.AddItem "1-Jump Random"
LBox.AddItem "2-Jump To Closest"
LBox.AddItem "3-Linear"
LBox.AddItem "4-Radial"

Hit = False

For i = 0 To LBox.ListCount - 1 Step 1
If LBox.List(i) = Text Then
Hit = True
Exit For
End If
Next i

If Hit = True Then
LBox.ListIndex = i
Else
LBox.ListIndex = 0
End If

LBox.SetFocus
End Sub

Public Sub PosAnimationDirectionSelect(X, Y, X2, Y2, Text)
LBox.Left = X
LBox.Top = Y
LBox.Width = X2
If LBox.Width + LBox.Left > CControl.Left + CControl.Width Then Exit Sub

If GScrolling = True Then Exit Sub

TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = True

LBox.Clear

LBox.AddItem "0-Bottom Left"
LBox.AddItem "1-Bottom Right"
LBox.AddItem "2-Down"
LBox.AddItem "3-Down Fire, Up End"
LBox.AddItem "4-Left"
LBox.AddItem "5-Left Fire, Right End"
LBox.AddItem "6-Left Of Target"
LBox.AddItem "7-Opposite Of Target"
LBox.AddItem "8-Right"
LBox.AddItem "9-Right Fire, Left End"
LBox.AddItem "10-Right Of Target"
LBox.AddItem "11-Still"
LBox.AddItem "12-Top Left"
LBox.AddItem "13-Top Right"
LBox.AddItem "14-To Target"
LBox.AddItem "15-Up"
LBox.AddItem "16-Up Fire, Down End"

Hit = False

For i = 0 To LBox.ListCount - 1 Step 1
If LBox.List(i) = Text Then
Hit = True
Exit For
End If
Next i

If Hit = True Then
LBox.ListIndex = i
Else
LBox.ListIndex = 0
End If

LBox.SetFocus
End Sub

Public Sub PosInfantryTypeSelect(X, Y, X2, Y2, Text)
LBox.Left = X
LBox.Top = Y
LBox.Width = X2
If LBox.Width + LBox.Left > CControl.Left + CControl.Width Then Exit Sub

If GScrolling = True Then Exit Sub

TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = True

LBox.Clear

LBox.AddItem "0-Hanger Unit"
LBox.AddItem "1-Infantry"
LBox.AddItem "2-Larva"
LBox.AddItem "3-Transport"
LBox.AddItem "4-Worker"

Hit = False

For i = 0 To LBox.ListCount - 1 Step 1
If LBox.List(i) = Text Then
Hit = True
Exit For
End If
Next i

If Hit = True Then
LBox.ListIndex = i
Else
LBox.ListIndex = 0
End If

LBox.SetFocus
End Sub

Public Sub PosWeaponTypeSelect(X, Y, X2, Y2, Text)
LBox.Left = X
LBox.Top = Y
LBox.Width = X2
If LBox.Width + LBox.Left > CControl.Left + CControl.Width Then Exit Sub

If GScrolling = True Then Exit Sub

TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = True

LBox.Clear

LBox.AddItem "0-Normal"
LBox.AddItem "1-Bomb"
LBox.AddItem "2-Mine"
LBox.AddItem "3-Seeking"
LBox.AddItem "4-Sub-Terrainian"

Hit = False

For i = 0 To LBox.ListCount - 1 Step 1
If LBox.List(i) = Text Then
Hit = True
Exit For
End If
Next i

If Hit = True Then
LBox.ListIndex = i
Else
LBox.ListIndex = 0
End If

LBox.SetFocus
End Sub

Public Sub PosUnitTypeSelect(X, Y, X2, Y2, Text)
LBox.Left = X
LBox.Top = Y
LBox.Width = X2
If LBox.Width + LBox.Left > CControl.Left + CControl.Width Then Exit Sub

If GScrolling = True Then Exit Sub

TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = True

LBox.Clear

LBox.AddItem "0-Biological"
LBox.AddItem "1-Energy"
LBox.AddItem "2-Mechanical"

Hit = False

For i = 0 To LBox.ListCount - 1 Step 1
If LBox.List(i) = Text Then
Hit = True
Exit For
End If
Next i

If Hit = True Then
LBox.ListIndex = i
Else
LBox.ListIndex = 0
End If

LBox.SetFocus
End Sub

Public Sub PosLarvaCancelModeSelect(X, Y, X2, Y2, Text)
LBox.Left = X
LBox.Top = Y
LBox.Width = X2
If LBox.Width + LBox.Left > CControl.Left + CControl.Width Then Exit Sub

If GScrolling = True Then Exit Sub

TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = True

LBox.Clear

LBox.AddItem "0-Kill Larva"
LBox.AddItem "1-Restore Larva"

Hit = False

For i = 0 To LBox.ListCount - 1 Step 1
If LBox.List(i) = Text Then
Hit = True
Exit For
End If
Next i

If Hit = True Then
LBox.ListIndex = i
Else
LBox.ListIndex = 0
End If

LBox.SetFocus
End Sub

Public Sub PosAnimationSelect(X, Y, X2, Y2, Text)
LBox.Left = X
LBox.Top = Y
LBox.Width = X2
If LBox.Width + LBox.Left > CControl.Left + CControl.Width Then Exit Sub

If GScrolling = True Then Exit Sub

TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = True

LBox.Clear

LBox.AddItem "[none]"

For i = 1 To Animations.Rows - 1
LBox.AddItem Animations.TextMatrix(i, 1)
Next i

Hit = False

For i = 0 To LBox.ListCount - 1 Step 1
If LBox.List(i) = Text Then
Hit = True
Exit For
End If
Next i

If Hit = True Then
LBox.ListIndex = i
Else
LBox.ListIndex = 0
End If

LBox.SetFocus
End Sub

Public Sub PosBuildingSelectNoNone(X, Y, X2, Y2, Text)
LBox.Left = X
LBox.Top = Y
LBox.Width = X2
If LBox.Width + LBox.Left > CControl.Left + CControl.Width Then Exit Sub

If GScrolling = True Then Exit Sub

TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = True

LBox.Clear

For i = 1 To Buildings.Rows - 1
LBox.AddItem Buildings.TextMatrix(i, 1)
Next i

Hit = False

For i = 0 To LBox.ListCount - 1 Step 1
If LBox.List(i) = Text Then
Hit = True
Exit For
End If
Next i

If Hit = True Then
LBox.ListIndex = i
Else
LBox.ListIndex = 0
End If

LBox.SetFocus
End Sub

Public Sub PosBuildingSelect(X, Y, X2, Y2, Text)
LBox.Left = X
LBox.Top = Y
LBox.Width = X2
If LBox.Width + LBox.Left > CControl.Left + CControl.Width Then Exit Sub

If GScrolling = True Then Exit Sub

TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = True

LBox.Clear

LBox.AddItem "[none]"

For i = 1 To Buildings.Rows - 1
LBox.AddItem Buildings.TextMatrix(i, 1)
Next i

Hit = False

For i = 0 To LBox.ListCount - 1 Step 1
If LBox.List(i) = Text Then
Hit = True
Exit For
End If
Next i

If Hit = True Then
LBox.ListIndex = i
Else
LBox.ListIndex = 0
End If

LBox.SetFocus
End Sub

Public Sub PosSoundSelect(X, Y, X2, Y2, Text)
LBox.Left = X
LBox.Top = Y
LBox.Width = X2
If LBox.Width + LBox.Left > CControl.Left + CControl.Width Then Exit Sub

If GScrolling = True Then Exit Sub

TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = True

LBox.Clear

LBox.AddItem "[none]"

For i = 1 To Sounds.Rows - 1
LBox.AddItem Sounds.TextMatrix(i, 1)
Next i

Hit = False

For i = 0 To LBox.ListCount - 1 Step 1
If LBox.List(i) = Text Then
Hit = True
Exit For
End If
Next i

If Hit = True Then
LBox.ListIndex = i
Else
LBox.ListIndex = 0
End If

LBox.SetFocus
End Sub

Public Sub PosWeaponSelect(X, Y, X2, Y2, Text)
LBox.Left = X
LBox.Top = Y
LBox.Width = X2
If LBox.Width + LBox.Left > CControl.Left + CControl.Width Then Exit Sub

If GScrolling = True Then Exit Sub

TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = True

LBox.Clear

LBox.AddItem "[none]"

For i = 1 To Weapons.Rows - 1
LBox.AddItem Weapons.TextMatrix(i, 1)
Next i

Hit = False

For i = 0 To LBox.ListCount - 1 Step 1
If LBox.List(i) = Text Then
Hit = True
Exit For
End If
Next i

If Hit = True Then
LBox.ListIndex = i
Else
LBox.ListIndex = 0
End If

LBox.SetFocus
End Sub

Public Sub PosArmorSelect(X, Y, X2, Y2, Text)
LBox.Left = X
LBox.Top = Y
LBox.Width = X2
If LBox.Width + LBox.Left > CControl.Left + CControl.Width Then Exit Sub

If GScrolling = True Then Exit Sub

TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = True

LBox.Clear

For i = 1 To Armor.Rows - 1
LBox.AddItem Armor.TextMatrix(i, 1)
Next i

Hit = False

For i = 0 To LBox.ListCount - 1 Step 1
If LBox.List(i) = Text Then
Hit = True
Exit For
End If
Next i

If Hit = True Then
LBox.ListIndex = i
Else
LBox.ListIndex = 0
End If

LBox.SetFocus
End Sub

Public Sub PosShieldSelect(X, Y, X2, Y2, Text)
LBox.Left = X
LBox.Top = Y
LBox.Width = X2
If LBox.Width + LBox.Left > CControl.Left + CControl.Width Then Exit Sub

If GScrolling = True Then Exit Sub

TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = True

LBox.Clear

LBox.AddItem "[none]"

For i = 1 To Shields.Rows - 1
LBox.AddItem Shields.TextMatrix(i, 1)
Next i

Hit = False

For i = 0 To LBox.ListCount - 1 Step 1
If LBox.List(i) = Text Then
Hit = True
Exit For
End If
Next i

If Hit = True Then
LBox.ListIndex = i
Else
LBox.ListIndex = 0
End If

LBox.SetFocus
End Sub

Public Sub posRaceSelectNoNone(X, Y, X2, Y2, Text)
LBox.Left = X
LBox.Top = Y
LBox.Width = X2
If LBox.Width + LBox.Left > CControl.Left + CControl.Width Then Exit Sub

If GScrolling = True Then Exit Sub

TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = True

LBox.Clear

For i = 1 To Races.Rows - 1
LBox.AddItem Races.TextMatrix(i, 1)
Next i

Hit = False

For i = 0 To LBox.ListCount - 1 Step 1
If LBox.List(i) = Text Then
Hit = True
Exit For
End If
Next i

If Hit = True Then
LBox.ListIndex = i
Else
LBox.ListIndex = 0
End If

LBox.SetFocus
End Sub

Public Sub PosWorkerSelect(X, Y, X2, Y2, Text)
LBox.Left = X
LBox.Top = Y
LBox.Width = X2
If LBox.Width + LBox.Left > CControl.Left + CControl.Width Then Exit Sub

If GScrolling = True Then Exit Sub

TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = True

LBox.Clear

For i = 1 To DataGrid.Rows - 1
If LCase(DataGrid.TextMatrix(i, 20)) <> "0-[none]" Then LBox.AddItem DataGrid.TextMatrix(i, 1)
Next i

Hit = False

For i = 0 To LBox.ListCount - 1 Step 1
If LBox.List(i) = Text Then
Hit = True
Exit For
End If
Next i

If Hit = True Then
LBox.ListIndex = i
Else
If LBox.ListCount > 0 Then LBox.ListIndex = 0
End If

If LBox.ListCount > 0 Then LBox.SetFocus Else LBox.Visible = False
End Sub

Public Sub PosInfantryWorkerSelect(X, Y, X2, Y2, Text)
LBox.Left = X
LBox.Top = Y
LBox.Width = X2
If LBox.Width + LBox.Left > CControl.Left + CControl.Width Then Exit Sub

If GScrolling = True Then Exit Sub

TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = True

LBox.Clear

For i = 1 To DataGrid.Rows - 1
If LCase(DataGrid.TextMatrix(i, 46)) = "4-Worker" Or LCase(DataGrid.TextMatrix(i, 46)) = "1-Infantry" Then LBox.AddItem DataGrid.TextMatrix(i, 1)
Next i

Hit = False

For i = 0 To LBox.ListCount - 1 Step 1
If LBox.List(i) = Text Then
Hit = True
Exit For
End If
Next i

If Hit = True Then
LBox.ListIndex = i
Else
LBox.ListIndex = 0
End If

LBox.SetFocus
End Sub

Public Sub PosLarvaSelect(X, Y, X2, Y2, Text)
LBox.Left = X
LBox.Top = Y
LBox.Width = X2
If LBox.Width + LBox.Left > CControl.Left + CControl.Width Then Exit Sub

If GScrolling = True Then Exit Sub

TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = True

LBox.Clear

For i = 1 To DataGrid.Rows - 1
If LCase(DataGrid.TextMatrix(i, 42)) = "2-Larva" Then LBox.AddItem DataGrid.TextMatrix(i, 1)
Next i

Hit = False

For i = 0 To LBox.ListCount - 1 Step 1
If LBox.List(i) = Text Then
Hit = True
Exit For
End If
Next i

If Hit = True Then
LBox.ListIndex = i
Else
LBox.ListIndex = 0
End If

LBox.SetFocus
End Sub

Public Sub PosMineSelect(X, Y, X2, Y2, Text)
LBox.Left = X
LBox.Top = Y
LBox.Width = X2
If LBox.Width + LBox.Left > CControl.Left + CControl.Width Then Exit Sub

If GScrolling = True Then Exit Sub

TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = True

LBox.Clear

For i = 1 To DataGrid.Rows - 1
If LCase(DataGrid.TextMatrix(i, 42)) = "2-Mine" Then LBox.AddItem DataGrid.TextMatrix(i, 1)
Next i

Hit = False

For i = 0 To LBox.ListCount - 1 Step 1
If LBox.List(i) = Text Then
Hit = True
Exit For
End If
Next i

If Hit = True Then
LBox.ListIndex = i
Else
LBox.ListIndex = 0
End If

LBox.SetFocus
End Sub

Public Sub PosHangerUnitSelect(X, Y, X2, Y2, Text)
LBox.Left = X
LBox.Top = Y
LBox.Width = X2
If LBox.Width + LBox.Left > CControl.Left + CControl.Width Then Exit Sub

If GScrolling = True Then Exit Sub

TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = True

LBox.Clear

For i = 1 To DataGrid.Rows - 1
If LCase(DataGrid.TextMatrix(i, 42)) = "Hanger Unit" Then LBox.AddItem DataGrid.TextMatrix(i, 1)
Next i

Hit = False

For i = 0 To LBox.ListCount - 1 Step 1
If LBox.List(i) = Text Then
Hit = True
Exit For
End If
Next i

If Hit = True Then
LBox.ListIndex = i
Else
LBox.ListIndex = 0
End If

LBox.SetFocus
End Sub

Public Sub PosAbilitiesSelect(X, Y, X2, Y2, Text)

If GScrolling = True Then Exit Sub

If frmSelect.Visible = False Then
TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = False
SelectSettingV = 0
SelectTextV = Text
frmSelect.Show , Me
frmSelect.OT = True
End If
End Sub

Public Sub PosUpgradesSelect(X, Y, X2, Y2, Text)

If GScrolling = True Then Exit Sub

If frmSelect.Visible = False Then
TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = False
SelectSettingV = 2
SelectTextV = Text
frmSelect.Show vbModal, Me
End If
End Sub

Public Sub PosMultipleUnitSelect(X, Y, X2, Y2, Text)

If GScrolling = True Then Exit Sub

If frmSelect.Visible = False Then
TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = False
SelectSettingV = 1
SelectTextV = Text
frmSelect.Show vbModal, Me
End If
End Sub

Public Sub PosMultipleUpgradeSelect(X, Y, X2, Y2, Text)

If GScrolling = True Then Exit Sub

If frmSelect.Visible = False Then
TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = False
SelectSettingV = 4
SelectTextV = Text
frmSelect.Show vbModal, Me
End If
End Sub

Public Sub PosMultipleAbilitySelect(X, Y, X2, Y2, Text)

If GScrolling = True Then Exit Sub

If frmSelect.Visible = False Then
TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = False
SelectSettingV = 3
SelectTextV = Text
frmSelect.Show vbModal, Me
End If
End Sub

Public Sub PosFieldsEmittedSelect(X, Y, X2, Y2, Text)

If GScrolling = True Then Exit Sub

If frmSelect.Visible = False Then
TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = False
SelectSettingV = 5
SelectTextV = Text
frmSelect.Show vbModal, Me
End If
End Sub

Public Sub PosMultipleInfantrySelect(X, Y, X2, Y2, Text)

If GScrolling = True Then Exit Sub

If frmSelect.Visible = False Then
TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = False
SelectSettingV = 6
SelectTextV = Text
frmSelect.Show vbModal, Me
End If
End Sub

Public Sub PosMultipleBuildingSelect(X, Y, X2, Y2, Text)

If GScrolling = True Then Exit Sub

If frmSelect.Visible = False Then
TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = False
SelectSettingV = 7
SelectTextV = Text
frmSelect.Show vbModal, Me
End If
End Sub

Public Sub PosPointSelect(X, Y, X2, Y2, Text)

If GScrolling = True Then Exit Sub

If frmPointSelect.Visible = False Then
TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = False
SelectTextV = Text
frmPointSelect.Show vbModal, Me
End If
End Sub

Public Sub PosMultiPointSelect(X, Y, X2, Y2, Text)

If GScrolling = True Then Exit Sub

If frmMultiPointSelect.Visible = False Then
TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = False
SelectTextV = Text
frmMultiPointSelect.Show vbModal, Me
End If
End Sub

Public Sub PosAddOnSelect(X, Y, X2, Y2, Text)

If GScrolling = True Then Exit Sub

If frmSelect.Visible = False Then
TextEdit.Visible = False
FileSelect.Visible = False: PlayView.Visible = False
LBox.Visible = False
SelectSettingV = 8
SelectTextV = Text
frmSelect.Show vbModal, Me
End If
End Sub

Public Sub DataGridUpdate(Scrolling As Boolean, DGrid)
GScrolling = Scrolling
Select Case DGrid
'######################################################################################
Case 1 'abilities######################################################################
'######################################################################################
Set CControl = Abilities
With Abilities
X = .Left + (.ColPos(.Col) \ 15)
Y = .Top + (.RowPos(.Row) \ 15)
X2 = .ColWidth(.Col) \ 15
Y2 = .RowHeight(.Row) \ 15
Select Case .Col
Case 1 '[Name]
PosTextEdit X, Y, X2, Y2, .Text, 255, False
Case 2 'Abilities Added
PosAbilitiesSelect X, Y, X2, Y2, .Text
Case 3 'Energy Required - Start
PosTextEdit X, Y, X2, Y2, .Text, 6, True
Case 4 'Energy Required - Per Second
PosTextEdit X, Y, X2, Y2, .Text, 6, True
Case 5 'Food Cost Added To Affected Units
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 6 'Mineral Cost
PosTextEdit X, Y, X2, Y2, .Text, 6, True
Case 7 'Sub-Mineral Cost
PosTextEdit X, Y, X2, Y2, .Text, 6, True
Case 8 'Time Cost
PosTextEdit X, Y, X2, Y2, .Text, 6, True
Case 9 'Units Affected - When Researched
PosMultipleUnitSelect X, Y, X2, Y2, .Text
End Select
End With
'######################################################################################
Case 2 'animations#####################################################################
'######################################################################################
Set CControl = Animations
With Animations
X = .Left + (.ColPos(.Col) \ 15)
Y = .Top + (.RowPos(.Row) \ 15)
X2 = .ColWidth(.Col) \ 15
Y2 = .RowHeight(.Row) \ 15
Select Case .Col
Case 1 '[Name]
PosTextEdit X, Y, X2, Y2, .Text, 255, False
Case 2 '[Size]'locked value
'locked
Case 3 '[Type]
PosAnimationType X, Y, X2, Y2, .Text
Case 4 'FPS
PosTextEdit X, Y, X2, Y2, .Text, 2, True
Case 5 'Image 1
PosImageSelect X, Y, X2, Y2, .Text
Case 6 'Image 2
PosImageSelect X, Y, X2, Y2, .Text
Case 7 'Image 3
PosImageSelect X, Y, X2, Y2, .Text
Case 8 'Image 4
PosImageSelect X, Y, X2, Y2, .Text
Case 9 'Image 5
PosImageSelect X, Y, X2, Y2, .Text
Case 10 'Image 6
PosImageSelect X, Y, X2, Y2, .Text
Case 11 'Image 7
PosImageSelect X, Y, X2, Y2, .Text
Case 12 'Image 8
PosImageSelect X, Y, X2, Y2, .Text
Case 13 'Image 9
PosImageSelect X, Y, X2, Y2, .Text
Case 14 'Image 10
PosImageSelect X, Y, X2, Y2, .Text
Case 15 'Image 11
PosImageSelect X, Y, X2, Y2, .Text
Case 16 'Image 12
PosImageSelect X, Y, X2, Y2, .Text
Case 17 'Pause Time Between Playback-MicroSeconds
PosTextEdit X, Y, X2, Y2, .Text, 6, True
Case 18 'Random Frames
PosTrueFalse X, Y, X2, Y2, .Text
End Select
End With
'######################################################################################
Case 3 'armor##########################################################################
'######################################################################################
Set CControl = Armor
With Armor
X = .Left + (.ColPos(.Col) \ 15)
Y = .Top + (.RowPos(.Row) \ 15)
X2 = .ColWidth(.Col) \ 15
Y2 = .RowHeight(.Row) \ 15
Select Case .Col
Case 1 '[Name]
PosTextEdit X, Y, X2, Y2, .Text, 255, False
Case 2 'Armor Effect
PosArmorShieldEffect X, Y, X2, Y2, .Text
Case 3 'Armor HP
PosTextEdit X, Y, X2, Y2, .Text, 6, True
Case 4 'Defence
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 5 'Effect Color
PosColorSelect X, Y, X2, Y2, .Text
End Select
End With
'######################################################################################
Case 4 'buildings######################################################################
'######################################################################################
Set CControl = Buildings
With Buildings
X = .Left + (.ColPos(.Col) \ 15)
Y = .Top + (.RowPos(.Row) \ 15)
X2 = .ColWidth(.Col) \ 15
Y2 = .RowHeight(.Row) \ 15
Select Case .Col
Case 1 '[Name]
PosTextEdit X, Y, X2, Y2, .Text, 255, False
Case 2 'Abilities
PosMultipleAbilitySelect X, Y, X2, Y2, .Text
Case 3 'Abilities Reseached At Building
PosMultipleAbilitySelect X, Y, X2, Y2, .Text
Case 4 'Ability Research Cancel
PosTrueFalse X, Y, X2, Y2, .Text
Case 5 'Air --> Land Animation
PosAnimationSelect X, Y, X2, Y2, .Text
Case 6 'Air --> Land Sound
PosSoundSelect X, Y, X2, Y2, .Text
Case 7 'Air --> Sea Animation
PosAnimationSelect X, Y, X2, Y2, .Text
Case 8 'Air --> Sea Sound
PosSoundSelect X, Y, X2, Y2, .Text
Case 9 'Air Build Animation
PosAnimationSelect X, Y, X2, Y2, .Text
Case 10 'Air Cancel Animation
PosAnimationSelect X, Y, X2, Y2, .Text
Case 11 'Air Fire Animation Point
PosPointSelect X, Y, X2, Y2, .Text
Case 12 'Air Moveing Animation
PosAnimationSelect X, Y, X2, Y2, .Text
Case 13 'Air Normal Animation
PosAnimationSelect X, Y, X2, Y2, .Text
Case 14 'Air Sight
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 15 'Air Sound Building - 1
PosSoundSelect X, Y, X2, Y2, .Text
Case 16 'Air Sound Building - 2
PosSoundSelect X, Y, X2, Y2, .Text
Case 17 'Air Sound Normal - 1
PosSoundSelect X, Y, X2, Y2, .Text
Case 18 'Air Sound Normal - 2
PosSoundSelect X, Y, X2, Y2, .Text
Case 19 'Air Weapon
PosWeaponSelect X, Y, X2, Y2, .Text
Case 20 'Armor
PosArmorSelect X, Y, X2, Y2, .Text
Case 21 'Burrows
PosTrueFalse X, Y, X2, Y2, .Text
Case 22 'Cancel Construction After Start
PosTrueFalse X, Y, X2, Y2, .Text
Case 23 'Cloaked
PosTrueFalse X, Y, X2, Y2, .Text
Case 24 'Damage Mode 'none,fire,sparks,blood shooting,blood around base,glowing - red,glowing - blue, glowing - green
PosDamageMode X, Y, X2, Y2, .Text
Case 25 'Food Cost
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 26 'Fields Emitted 'cloak/power
PosFieldsEmittedSelect X, Y, X2, Y2, .Text
Case 27 'Floatable
PosTrueFalse X, Y, X2, Y2, .Text
Case 28 'Flyable
PosTrueFalse X, Y, X2, Y2, .Text
Case 29 'Function In Air
PosTrueFalse X, Y, X2, Y2, .Text
Case 30 'Function In Sea
PosTrueFalse X, Y, X2, Y2, .Text
Case 31 'Function On Land
PosTrueFalse X, Y, X2, Y2, .Text
Case 32 'Hanger Unit
PosHangerUnitSelect X, Y, X2, Y2, .Text
Case 33 'Has Hanger
PosTrueFalse X, Y, X2, Y2, .Text
Case 34 'Has Larva
PosTrueFalse X, Y, X2, Y2, .Text
Case 35 'HP
PosTextEdit X, Y, X2, Y2, .Text, 6, True
Case 36 'Land --> Air Animation
PosAnimationSelect X, Y, X2, Y2, .Text
Case 37 'Land --> Air Sound
PosSoundSelect X, Y, X2, Y2, .Text
Case 38 'Land Build Animation
PosAnimationSelect X, Y, X2, Y2, .Text
Case 39 'Land Cancel Animation
PosAnimationSelect X, Y, X2, Y2, .Text
Case 40 'Land Fire Animation Point
PosPointSelect X, Y, X2, Y2, .Text
Case 41 'Land Moveing Animation
PosAnimationSelect X, Y, X2, Y2, .Text
Case 42 'Land Normal Animation
PosAnimationSelect X, Y, X2, Y2, .Text
Case 43 'Land Sight
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 44 'Land Sound Building - 1
PosSoundSelect X, Y, X2, Y2, .Text
Case 45 'Land Sound Building - 2
PosSoundSelect X, Y, X2, Y2, .Text
Case 46 'Land Sound Normal - 1
PosSoundSelect X, Y, X2, Y2, .Text
Case 47 'Land Sound Normal - 2
PosSoundSelect X, Y, X2, Y2, .Text
Case 48 'Land Weapon
PosWeaponSelect X, Y, X2, Y2, .Text
Case 49 'Larva Add Speed[In Seconds]
PosTextEdit X, Y, X2, Y2, .Text, 3, True
Case 50 'Larva Cancel Animation
PosAnimationSelect X, Y, X2, Y2, .Text
Case 51 'Larva Cancel Mode
PosLarvaCancelModeSelect X, Y, X2, Y2, .Text
Case 52 'Larva Morph Animation
PosAnimationSelect X, Y, X2, Y2, .Text
Case 53 'Larva Morph Done Animation
PosAnimationSelect X, Y, X2, Y2, .Text
Case 54 'Larva Transformation(s)
PosMultipleInfantrySelect X, Y, X2, Y2, .Text
Case 55 'Larva Unit
PosLarvaSelect X, Y, X2, Y2, .Text
Case 56 'Max Larva Units
PosTextEdit X, Y, X2, Y2, .Text, 2, True
Case 57 'Max Hanger Units
PosTextEdit X, Y, X2, Y2, .Text, 2, True
Case 58 'Mineral Cost
PosTextEdit X, Y, X2, Y2, .Text, 6, True
Case 59 'Movable - Normal
PosTrueFalse X, Y, X2, Y2, .Text
Case 60 'Power Cost
PosTextEdit X, Y, X2, Y2, .Text, 1, True
Case 61 'Repair Animation
PosAnimationSelect X, Y, X2, Y2, .Text
Case 62 'Requires Power
PosTrueFalse X, Y, X2, Y2, .Text
Case 63 'Rollable
PosTrueFalse X, Y, X2, Y2, .Text
Case 64 'Sea --> Air Animation
PosAnimationSelect X, Y, X2, Y2, .Text
Case 65 'Sea --> Air Sound
PosSoundSelect X, Y, X2, Y2, .Text
Case 66 'Sea Build Animation
PosAnimationSelect X, Y, X2, Y2, .Text
Case 67 'Sea Cancel Animation
PosAnimationSelect X, Y, X2, Y2, .Text
Case 68 'Sea Fire Animation Point
PosPointSelect X, Y, X2, Y2, .Text
Case 69 'Sea Moveing Animation
PosAnimationSelect X, Y, X2, Y2, .Text
Case 70 'Sea Normal Animation
PosAnimationSelect X, Y, X2, Y2, .Text
Case 71 'Sea Sight
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 72 'Sea Sound Building - 1
PosSoundSelect X, Y, X2, Y2, .Text
Case 73 'Sea Sound Building - 2
PosSoundSelect X, Y, X2, Y2, .Text
Case 74 'Sea Sound Normal - 1
PosSoundSelect X, Y, X2, Y2, .Text
Case 75 'Sea Sound Normal - 2
PosSoundSelect X, Y, X2, Y2, .Text
Case 76 'Sea Weapon
PosWeaponSelect X, Y, X2, Y2, .Text
Case 77 'Shield Name
PosShieldSelect X, Y, X2, Y2, .Text
Case 78 'Start Larva Amount
PosTextEdit X, Y, X2, Y2, .Text, 2, True
Case 79 'Start Amount In Hanger
PosTextEdit X, Y, X2, Y2, .Text, 2, True
Case 80 'Sub-Mineral Cost
PosTextEdit X, Y, X2, Y2, .Text, 6, True
Case 81 'Time Cost
PosTextEdit X, Y, X2, Y2, .Text, 6, True
Case 82 'Unit Construction Cancel
PosTrueFalse X, Y, X2, Y2, .Text
Case 83 'Unit Energy
PosTextEdit X, Y, X2, Y2, .Text, 5, True
Case 84 'Upgrades Researched At Building
PosMultipleUpgradeSelect X, Y, X2, Y2, .Text
Case 85 'Upgrade To Building - 1
PosBuildingSelect X, Y, X2, Y2, .Text
Case 86 'Upgrade To Building - 2
PosBuildingSelect X, Y, X2, Y2, .Text
Case 87 'Upgrade To Building - 3
PosBuildingSelect X, Y, X2, Y2, .Text
Case 88 'Wire Frame Image
PosImageSelect X, Y, X2, Y2, .Text
Case 89 '{AddOns}
PosAddOnSelect X, Y, X2, Y2, .Text
Case 90 '{Is AddOn}
PosTrueFalse X, Y, X2, Y2, .Text
Case 91 '{Morphs}
PosMultipleBuildingSelect X, Y, X2, Y2, .Text
Case 92 '{Race}
posRaceSelectNoNone X, Y, X2, Y2, .Text
Case 93 '{Required Buildings}
PosMultipleBuildingSelect X, Y, X2, Y2, .Text
Case 94 '{Stores Units}
PosMultipleInfantrySelect X, Y, X2, Y2, .Text
Case 95 '{Teleporter}
PosTrueFalse X, Y, X2, Y2, .Text
Case 96 '{Teleport Exit}
PosBuildingSelect X, Y, X2, Y2, .Text
End Select
End With
'######################################################################################
Case 5 'images#########################################################################
'######################################################################################
Set CControl = Images
With Images
X = .Left + (.ColPos(.Col) \ 15)
Y = .Top + (.RowPos(.Row) \ 15)
X2 = .ColWidth(.Col) \ 15
Y2 = .RowHeight(.Row) \ 15
Select Case .Col
Case 1 '[Name]
PosTextEdit X, Y, X2, Y2, .Text, 255, False
Case 2 'Location
PosImageFileSelect X, Y, X2, Y2, .Text
End Select
End With
'######################################################################################
Case 6 'infantry#######################################################################
'######################################################################################
Set CControl = DataGrid
With DataGrid
X = .Left + (.ColPos(.Col) \ 15)
Y = .Top + (.RowPos(.Row) \ 15)
X2 = .ColWidth(.Col) \ 15
Y2 = .RowHeight(.Row) \ 15
Select Case .Col
Case 1 '[Name]
PosTextEdit X, Y, X2, Y2, .Text, 255, False
Case 2 '[Name - Mode 2 - Leave Blank For Single Mode]
PosTextEdit X, Y, X2, Y2, .Text, 255, False
Case 3 'Abilities - Mode 1
PosMultipleAbilitySelect X, Y, X2, Y2, .Text
Case 4 'Abilities - Mode 2
PosMultipleAbilitySelect X, Y, X2, Y2, .Text
Case 5 'Air Fire Animation Point - Mode 1 'x,y
PosPointSelect X, Y, X2, Y2, .Text
Case 6 'Air Fire Animation Point - Mode 2 'x,y
PosPointSelect X, Y, X2, Y2, .Text
Case 7 'Air Repair Range - Mode 1
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 8 'Air Repair Range - Mode 2
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 9 'Air Sight - Mode 1
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 10 'Air Sight - Mode 2
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 11 'Air Weapon - Mode 1
PosWeaponSelect X, Y, X2, Y2, .Text
Case 12 'Air Weapon - Mode 2
PosWeaponSelect X, Y, X2, Y2, .Text
Case 13 'Armor - Mode 1
PosArmorSelect X, Y, X2, Y2, .Text
Case 14 'Armor - Mode 2
PosArmorSelect X, Y, X2, Y2, .Text
Case 15 'Auto-Repair - Mode 1
PosTrueFalse X, Y, X2, Y2, .Text
Case 16 'Auto-Repair - Mode 2
PosTrueFalse X, Y, X2, Y2, .Text
Case 17 'Build Location 1[Optional If Hanger Or Larva Unit]
If CControl.TextMatrix(.Row, 46) = "Larva" Or CControl.TextMatrix(.Row, 46) = "Hanger Unit" Then
PosBuildingSelect X, Y, X2, Y2, .Text
Else
PosBuildingSelectNoNone X, Y, X2, Y2, .Text
End If
Case 18 'Build Location 2[Optional]
PosBuildingSelect X, Y, X2, Y2, .Text
Case 19 'Build Location 3[Optional]
PosBuildingSelect X, Y, X2, Y2, .Text
Case 20 'Build Method - Mode 1 'start and leave/stay in spot/stay and move/morph/none
PosBuildMethodSelect X, Y, X2, Y2, .Text
Case 21 'Build Method - Mode 2 'start and leave/stay in spot/stay and move/morph/none
PosBuildMethodSelect X, Y, X2, Y2, .Text
Case 22 'Cloaked - Mode 1
PosTrueFalse X, Y, X2, Y2, .Text
Case 23 'Cloaked - Mode 2
PosTrueFalse X, Y, X2, Y2, .Text
Case 24 'Covers Area Of Unit Image - Mode 1 'x1,y1,x2,y2
PosMultiPointSelect X, Y, X2, Y2, .Text
Case 25 'Covers Area Of Unit Image - Mode 2 'x1,y1,x2,y2
PosMultiPointSelect X, Y, X2, Y2, .Text
Case 26 'Death Animation 1 - Mode 1
PosAnimationSelect X, Y, X2, Y2, .Text
Case 27 'Death Animation 1 - Mode 2
PosAnimationSelect X, Y, X2, Y2, .Text
Case 28 'Death Animation 2 - Mode 1
PosAnimationSelect X, Y, X2, Y2, .Text
Case 29 'Death Animation 2 - Mode 2
PosAnimationSelect X, Y, X2, Y2, .Text
Case 30 'Death Animation 3 - Mode 1
PosAnimationSelect X, Y, X2, Y2, .Text
Case 31 'Death Animation 3 - Mode 2
PosAnimationSelect X, Y, X2, Y2, .Text
Case 32 'Death Sound - Mode 1
PosSoundSelect X, Y, X2, Y2, .Text
Case 33 'Death Sound - Mode 2
PosSoundSelect X, Y, X2, Y2, .Text
Case 34 'Fields Emitted - Mode 1 'cloak/power
PosFieldsEmittedSelect X, Y, X2, Y2, .Text
Case 35 'Fields Emitted - Mode 2 'cloak/power
PosFieldsEmittedSelect X, Y, X2, Y2, .Text
Case 36 'Hanger - Mode 1
PosTrueFalse X, Y, X2, Y2, .Text
Case 37 'Hanger - Mode 2
PosTrueFalse X, Y, X2, Y2, .Text
Case 38 'Hanger Max - Mode 1
PosTextEdit X, Y, X2, Y2, .Text, 2, True
Case 39 'Hanger Max - Mode 2
PosTextEdit X, Y, X2, Y2, .Text, 2, True
Case 40 'Hanger Start - Mode 1
PosTextEdit X, Y, X2, Y2, .Text, 2, True
Case 41 'Hanger Start - Mode 2
PosTextEdit X, Y, X2, Y2, .Text, 2, True
Case 42 'Hanger Unit - Mode 1
PosHangerUnitSelect X, Y, X2, Y2, .Text
Case 43 'Hanger Unit - Mode 2
PosHangerUnitSelect X, Y, X2, Y2, .Text
Case 44 'HP - Mode 1
PosTextEdit X, Y, X2, Y2, .Text, 6, True
Case 45 'HP - Mode 2
PosTextEdit X, Y, X2, Y2, .Text, 6, True
Case 46 'Infantry Type - Mode 1 'Infantry/Worker/Hanger Unit/Larva
PosInfantryTypeSelect X, Y, X2, Y2, .Text
Case 47 'Infantry Type - Mode 2 'Infantry/Worker/Hanger Unit/Larva
PosInfantryTypeSelect X, Y, X2, Y2, .Text
Case 48 'Land Fire Animation Point - Mode 1 'x,y
PosPointSelect X, Y, X2, Y2, .Text
Case 49 'Land Fire Animation Point - Mode 2 'x,y
PosPointSelect X, Y, X2, Y2, .Text
Case 50 'Land Repair Range - Mode 1
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 51 'Land Repair Range - Mode 2
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 52 'Land Sight - Mode 1
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 53 'Land Sight - Mode 2
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 54 'Land Weapon - Mode 1
PosWeaponSelect X, Y, X2, Y2, .Text
Case 55 'Land Weapon - Mode 2
PosWeaponSelect X, Y, X2, Y2, .Text
Case 56 'Mode 1 --> 2 Animation
PosAnimationSelect X, Y, X2, Y2, .Text
Case 57 'Mode 1 --> 2 Sound
PosSoundSelect X, Y, X2, Y2, .Text
Case 58 'Mode 1 - Food Cost
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 59 'Mode 1 - Mineral Cost
PosTextEdit X, Y, X2, Y2, .Text, 6, True
Case 60 'Mode 1 - Power Cost
PosTextEdit X, Y, X2, Y2, .Text, 1, True
Case 61 'Mode 1 - Sub-Mineral Cost
PosTextEdit X, Y, X2, Y2, .Text, 6, True
Case 62 'Mode 1 - Time Cost
PosTextEdit X, Y, X2, Y2, .Text, 6, True
Case 63 'Mode 1 Icon
PosImageSelect X, Y, X2, Y2, .Text
Case 64 'Mode 2 --> 1 Animation
PosAnimationSelect X, Y, X2, Y2, .Text
Case 65 'Mode 2 --> 1 Sound
PosSoundSelect X, Y, X2, Y2, .Text
Case 66 'Mode 2 - Food Cost
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 67 'Mode 2 - Mineral Cost
PosTextEdit X, Y, X2, Y2, .Text, 6, True
Case 68 'Mode 2 - Power Cost
PosTextEdit X, Y, X2, Y2, .Text, 1, True
Case 69 'Mode 2 - Sub-Mineral Cost
PosTextEdit X, Y, X2, Y2, .Text, 6, True
Case 70 'Mode 2 - Time Cost
PosTextEdit X, Y, X2, Y2, .Text, 6, True
Case 71 'Mode 2 Icon
PosImageSelect X, Y, X2, Y2, .Text
Case 72 'Mode 2 Upgrade Location 1
PosBuildingSelect X, Y, X2, Y2, .Text
Case 73 'Mode 2 Upgrade Location 2[Optional]
PosBuildingSelect X, Y, X2, Y2, .Text
Case 74 'Mode 2 Upgrade Location 3[Optional]
PosBuildingSelect X, Y, X2, Y2, .Text
Case 75 'Mode Switch '1-->2|1<->2
PosModeSwitchSelect X, Y, X2, Y2, .Text
Case 76 'Repair/Heal Animation - Mode 1
PosAnimationSelect X, Y, X2, Y2, .Text
Case 77 'Repair/Heal Animation - Mode 2
PosAnimationSelect X, Y, X2, Y2, .Text
Case 78 'Repairs Type - Mode 1 'Bio/Energy/Mech
PosUnitTypeSelect X, Y, X2, Y2, .Text
Case 79 'Repairs Type - Mode 2 'Bio/Energy/Mech
PosUnitTypeSelect X, Y, X2, Y2, .Text
Case 80 'Sea Fire Animation Point - Mode 1 'x,y
PosPointSelect X, Y, X2, Y2, .Text
Case 81 'Sea Fire Animation Point - Mode 2 'x,y
PosPointSelect X, Y, X2, Y2, .Text
Case 82 'Sea Repair Range - Mode 1
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 83 'Sea Repair Range - Mode 2
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 84 'Sea Sight - Mode 1
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 85 'Sea Sight - Mode 2
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 86 'Sea Weapon - Mode 1
PosWeaponSelect X, Y, X2, Y2, .Text
Case 87 'Sea Weapon - Mode 2
PosWeaponSelect X, Y, X2, Y2, .Text
Case 88 'Shield - Mode 1
PosShieldSelect X, Y, X2, Y2, .Text
Case 89 'Shield - Mode 2
PosShieldSelect X, Y, X2, Y2, .Text
Case 90 'Sound Acknowledge - 1 - Mode 1
PosSoundSelect X, Y, X2, Y2, .Text
Case 91 'Sound Acknowledge - 1 - Mode 2
PosSoundSelect X, Y, X2, Y2, .Text
Case 92 'Sound Acknowledge - 2 - Mode 1
PosSoundSelect X, Y, X2, Y2, .Text
Case 93 'Sound Acknowledge - 2 - Mode 2
PosSoundSelect X, Y, X2, Y2, .Text
Case 94 'Sound Acknowledge - 3 - Mode 1
PosSoundSelect X, Y, X2, Y2, .Text
Case 95 'Sound Acknowledge - 3 - Mode 2
PosSoundSelect X, Y, X2, Y2, .Text
Case 96 'Sound Acknowledge - 4 - Mode 1
PosSoundSelect X, Y, X2, Y2, .Text
Case 97 'Sound Acknowledge - 4 - Mode 2
PosSoundSelect X, Y, X2, Y2, .Text
Case 98 'Sound Acknowledge - 5 - Mode 1
PosSoundSelect X, Y, X2, Y2, .Text
Case 99 'Sound Acknowledge - 5 - Mode 2
PosSoundSelect X, Y, X2, Y2, .Text
Case 100 'Still Animation - Mode 1
PosAnimationSelect X, Y, X2, Y2, .Text
Case 101 'Still Animation - Mode 2
PosAnimationSelect X, Y, X2, Y2, .Text
Case 102 'Terrain - Mode 1
PosTerrainSelect X, Y, X2, Y2, .Text
Case 103 'Terrain - Mode 2
PosTerrainSelect X, Y, X2, Y2, .Text
Case 104 'Unit Energy - Mode 1
PosTextEdit X, Y, X2, Y2, .Text, 6, True
Case 105 'Unit Energy - Mode 2
PosTextEdit X, Y, X2, Y2, .Text, 6, True
Case 106 'Unit Morph - Mode 1
PosMultipleInfantrySelect X, Y, X2, Y2, .Text
Case 107 'Unit Morph - Mode 2
PosMultipleInfantrySelect X, Y, X2, Y2, .Text
Case 108 'Unit Morph Enabled - Mode 1
PosTrueFalse X, Y, X2, Y2, .Text
Case 109 'Unit Morph Enabled - Mode 2
PosTrueFalse X, Y, X2, Y2, .Text
Case 110 'Unit Morph Finish Animation - Mode 1
PosAnimationSelect X, Y, X2, Y2, .Text
Case 111 'Unit Morph Finish Animation - Mode 2
PosAnimationSelect X, Y, X2, Y2, .Text
Case 112 'Unit Morph Start Animation - Mode 1
PosAnimationSelect X, Y, X2, Y2, .Text
Case 113 'Unit Morph Start Animation - Mode 2
PosAnimationSelect X, Y, X2, Y2, .Text
Case 114 'Unit Type - Mode 1 'Bio/Mech
PosUnitTypeSelect X, Y, X2, Y2, .Text
Case 115 'Unit Type - Mode 2 'Bio/Mech
PosUnitTypeSelect X, Y, X2, Y2, .Text
Case 116 'Walk Animation - Mode 1
PosAnimationSelect X, Y, X2, Y2, .Text
Case 117 'Walk Animation - Mode 2
PosAnimationSelect X, Y, X2, Y2, .Text
Case 118 'Wire Frame Image - Mode 1
PosImageSelect X, Y, X2, Y2, .Text
Case 119 'Wire Frame Image - Mode 2
PosImageSelect X, Y, X2, Y2, .Text
Case 120 '{Builds Mines - Mode 1}
PosTrueFalse X, Y, X2, Y2, .Text
Case 121 '{Builds Mines - Mode 2}
PosTrueFalse X, Y, X2, Y2, .Text
Case 122 '{Mine - Mode 1}
PosWeaponSelect X, Y, X2, Y2, .Text
Case 123 '{Mine - Mode 2}
PosWeaponSelect X, Y, X2, Y2, .Text
Case 124 '{Mine Payload - Mode 1}
PosTextEdit X, Y, X2, Y2, .Text, 2, True
Case 125 '{Mine Payload - Mode 2}
PosTextEdit X, Y, X2, Y2, .Text, 2, True
Case 126 '{Race}
posRaceSelectNoNone X, Y, X2, Y2, .Text
Case 127 '{Required Buildings}
PosMultipleBuildingSelect X, Y, X2, Y2, .Text
Case 128 '{Unit Attacks Burrowed}
PosTrueFalse X, Y, X2, Y2, .Text
Case 129 '{Unit Attacks Sunken}
PosTrueFalse X, Y, X2, Y2, .Text
Case 130 '{Unit Burrowable}
PosTrueFalse X, Y, X2, Y2, .Text
Case 131 '{Unit Is Detector}
PosTrueFalse X, Y, X2, Y2, .Text
Case 132 '{Unit Responds Over Distance}
PosTextEdit X, Y, X2, Y2, .Text, 3, True
Case 133 '{Unit Sinkable}
PosTrueFalse X, Y, X2, Y2, .Text
End Select
End With
'######################################################################################
Case 11 'races##########################################################################
'######################################################################################
Set CControl = Races
With Races
X = .Left + (.ColPos(.Col) \ 15)
Y = .Top + (.RowPos(.Row) \ 15)
X2 = .ColWidth(.Col) \ 15
Y2 = .RowHeight(.Row) \ 15
Select Case .Col
Case 1 '[Name]
PosTextEdit X, Y, X2, Y2, .Text, 255, False
Case 2 'Cost Factor - Food
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 3 'Cost Factor - Minerals
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 4 'Cost Factor - Sub-Minerals
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 5 'Cost Factor - Time
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 6 'Max Food
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 7 'Starting Building
PosBuildingSelectNoNone X, Y, X2, Y2, .Text
Case 8 'Starting Worker
PosWorkerSelect X, Y, X2, Y2, .Text
End Select
End With
'######################################################################################
Case 7 'shields########################################################################
'######################################################################################
Set CControl = Shields
With Shields
X = .Left + (.ColPos(.Col) \ 15)
Y = .Top + (.RowPos(.Row) \ 15)
X2 = .ColWidth(.Col) \ 15
Y2 = .RowHeight(.Row) \ 15
Select Case .Col
Case 1 '[Name]
PosTextEdit X, Y, X2, Y2, .Text, 255, False
Case 2 'Defence
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 3 'Effect Color
PosColorSelect X, Y, X2, Y2, .Text
Case 4 'Shield Effect
PosArmorShieldEffect X, Y, X2, Y2, .Text
Case 5 'Shield HP
PosTextEdit X, Y, X2, Y2, .Text, 6, True
End Select
End With
'######################################################################################
Case 8 'sounds#########################################################################
'######################################################################################
Set CControl = Sounds
With Sounds
X = .Left + (.ColPos(.Col) \ 15)
Y = .Top + (.RowPos(.Row) \ 15)
X2 = .ColWidth(.Col) \ 15
Y2 = .RowHeight(.Row) \ 15
Select Case .Col
Case 1 '[Name]
PosTextEdit X, Y, X2, Y2, .Text, 255, False
Case 2 'Location
PosSoundFileSelect X, Y, X2, Y2, .Text
Case 3 'Pause Time In Seconds Between Playback[-1 To Play Once]
PosTextEdit X, Y, X2, Y2, .Text, 6, True
End Select
End With
'######################################################################################
Case 9 'upgrades#######################################################################
'######################################################################################
Set CControl = Upgrades
With Upgrades
X = .Left + (.ColPos(.Col) \ 15)
Y = .Top + (.RowPos(.Row) \ 15)
X2 = .ColWidth(.Col) \ 15
Y2 = .RowHeight(.Row) \ 15
Select Case .Col
Case 1 '[Name]
PosTextEdit X, Y, X2, Y2, .Text, 255, False
Case 2 'Amount Added To Properties
PosTextEdit X, Y, X2, Y2, .Text, 3, True
Case 3 'Food Cost Added To Affected Units
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 4 'Mineral Cost
PosTextEdit X, Y, X2, Y2, .Text, 6, True
Case 5 'Max Upgrades
PosTextEdit X, Y, X2, Y2, .Text, 3, True
Case 6 'Properties Upgraded
PosUpgradesSelect X, Y, X2, Y2, .Text
Case 7 'Sub-Mineral Cost
PosTextEdit X, Y, X2, Y2, .Text, 6, True
Case 8 'Time Cost
PosTextEdit X, Y, X2, Y2, .Text, 6, True
Case 9 'Units Affected
PosMultipleUnitSelect X, Y, X2, Y2, .Text
End Select
End With
'######################################################################################
Case 10 'weapons#######################################################################
'######################################################################################
Set CControl = Weapons
With Weapons
X = .Left + (.ColPos(.Col) \ 15)
Y = .Top + (.RowPos(.Row) \ 15)
X2 = .ColWidth(.Col) \ 15
Y2 = .RowHeight(.Row) \ 15
Select Case .Col
Case 1 '[Name]
PosTextEdit X, Y, X2, Y2, .Text, 255, False
Case 2 'Air Animation
PosAnimationSelect X, Y, X2, Y2, .Text
Case 3 'Air End Animation
PosAnimationSelect X, Y, X2, Y2, .Text
Case 4 'Air Damage
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 5 'Air Fire Animation
PosAnimationSelect X, Y, X2, Y2, .Text
Case 6 'Air Fireing Sound
PosSoundSelect X, Y, X2, Y2, .Text
Case 7 'Air Fireing Speed
PosTextEdit X, Y, X2, Y2, .Text, 3, True
Case 8 'Air Range
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 9 'Air Sight
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 10 'Air Speed
PosTextEdit X, Y, X2, Y2, .Text, 3, True
Case 11 'Air Splash Damage
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 12 'Air Splash Damage Range
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 13 'Animation Direction'still,to target,opposite of target,up,down,left,right,topleft,topright,bottomleft,bottomright,left of target,right of target,up then down on unit,down then up on unit,left then right on unit,right then left on unit
PosAnimationDirectionSelect X, Y, X2, Y2, .Text
Case 14 'Land Animation
PosAnimationSelect X, Y, X2, Y2, .Text
Case 15 'Land Damage
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 16 'Land End Animation
PosAnimationSelect X, Y, X2, Y2, .Text
Case 17 'Land Fire Animation
PosAnimationSelect X, Y, X2, Y2, .Text
Case 18 'Land Fireing Sound
PosSoundSelect X, Y, X2, Y2, .Text
Case 19 'Land Fireing Speed
PosTextEdit X, Y, X2, Y2, .Text, 3, True
Case 20 'Land Range
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 21 'Land Sight
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 22 'Land Speed
PosTextEdit X, Y, X2, Y2, .Text, 3, True
Case 23 'Land Splash Damage
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 24 'Land Splash Damage Range
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 25 'Sea Animation
PosAnimationSelect X, Y, X2, Y2, .Text
Case 26 'Sea Damage
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 27 'Sea End Animation
PosAnimationSelect X, Y, X2, Y2, .Text
Case 28 'Sea Fire Animation
PosAnimationSelect X, Y, X2, Y2, .Text
Case 29 'Sea Fireing Sound
PosSoundSelect X, Y, X2, Y2, .Text
Case 30 'Sea Fireing Speed
PosTextEdit X, Y, X2, Y2, .Text, 3, True
Case 31 'Sea Range
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 32 'Sea Sight
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 33 'Sea Speed
PosTextEdit X, Y, X2, Y2, .Text, 3, True
Case 34 'Sea Splash Damage
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 35 'Sea Splash Damage Range
PosTextEdit X, Y, X2, Y2, .Text, 4, True
Case 36 'Splash Damage Animation
PosAnimationSelect X, Y, X2, Y2, .Text
Case 37 'Splash Damage Mode
PosSplashDamageModeSelect X, Y, X2, Y2, .Text
Case 38 'Travels Over Terrain
PosTerrainSelect X, Y, X2, Y2, .Text
Case 39 'Weapon Type
PosWeaponTypeSelect X, Y, X2, Y2, .Text
End Select
End With
End Select
End Sub

Private Sub DataGrid_KeyDown(KeyCode As Integer, Shift As Integer)
subKeyDown KeyCode, Shift
Select Case KeyCode
Case 37
If DataGrid.Col > 1 Then DataGrid.Col = DataGrid.Col - 1
Case 38
If DataGrid.Row > 1 Then DataGrid.Row = DataGrid.Row - 1
Case 39
If DataGrid.Col < DataGrid.Cols - 2 Then DataGrid.Col = DataGrid.Col + 1
Case 40
If DataGrid.Row < DataGrid.Rows - 2 Then DataGrid.Row = DataGrid.Row + 1
End Select
End Sub

Private Sub DataGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If HF = True Then Label2_Click
If FF = True Then Label1_Click
If DF = True Then Label10_Click
End Sub

Private Sub DataGrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
DataGridUpdate True, 6
End Sub

Private Sub DataGrid_Scroll()
DataGridUpdate True, 6
End Sub

Private Sub FileSelect_Click()
If HF = True Then Label2_Click
If FF = True Then Label1_Click
If DF = True Then Label10_Click
Dim tS As String, tL As String
CD.FileName = ""
CD.FileName = CControl.TextMatrix(CControl.Row, CControl.Col)
CD.ShowOpen
If CD.Flags <> 0 Then
If LCase(Right(CD.FileName, 4)) = ".bmp" Then 'image file
Set TPic = LoadPicture(CD.FileName)
Dim tH As hBitmap
LoadBitmap CD.FileName, tH
BitBlt TPic.hdc, 0, 0, tH.Width, tH.Height, tH.hdc, 0, 0, SRCCOPY
tS = ""
For X = 0 To tH.Width - 1 Step 1
For Y = 0 To tH.Height - 1 Step 1
tS = tS & TPic.Point(X, Y) & Chr(171)
Next Y
tS = tS & Chr(170)
Next X
DeleteBitmap tH
CControl.TextMatrix(CControl.Row, CControl.Col) = tS
ElseIf LCase(Right(CD.FileName, 4)) = ".wav" Then 'sound file
TRtf.LoadFile CD.FileName
TRtf.Text = Replace(TRtf.Text, "$", "$0")
TRtf.Text = Replace(TRtf.Text, Chr(10), "$1")
TRtf.Text = Replace(TRtf.Text, Chr(13), "$2")
TRtf.Text = Replace(TRtf.Text, Chr(0), "$3")
tS = TRtf.Text
If Len(CControl.TextMatrix(CControl.Row, CControl.Col)) > 18 Then
If Mid(CControl.TextMatrix(CControl.Row, CControl.Col), 1, 18) = "SoundArrayVariable" Then
a = AddSound(tS, CInt(Mid(CControl.TextMatrix(CControl.Row, CControl.Col), 19)))
Else
a = AddSound(tS)
End If
Else
a = AddSound(tS)
End If
CControl.TextMatrix(CControl.Row, CControl.Col) = "SoundArrayVariable" & Mid(Str(a), 2)
End If
End If
End Sub

Private Sub FileSelect_KeyDown(KeyCode As Integer, Shift As Integer)
subKeyDown KeyCode, Shift
Select Case KeyCode
Case 37
If DataGrid.Col > 1 Then DataGrid.Col = DataGrid.Col - 1
Case 38
If DataGrid.Row > 1 Then DataGrid.Row = DataGrid.Row - 1
Case 39
If DataGrid.Col < DataGrid.Cols - 2 Then DataGrid.Col = DataGrid.Col + 1
Case 40
If DataGrid.Row < DataGrid.Rows - 2 Then DataGrid.Row = DataGrid.Row + 1
End Select
End Sub

Private Sub Form_Load()
CreateMod
Dim GridDataArray(0, 0, 0)
TSv = 0
LastState = vbNormal
ShapeForm Me
If App.Minor < 10 And App.Revision < 10 Then
lCaption.Caption = App.Title & " - v" & App.Major & "." & App.Minor & App.Revision
Else
lCaption.Caption = App.Title & " - v" & App.Major & "." & App.Minor & "." & App.Revision
End If
Form1.Show
Form1.Top = (Me.Top + Me.Height) * 10
Form1.Left = Form1.Width * 2
Form1.Caption = lCaption.Caption
Me.SetFocus
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If HF = True Then Label2_Click
If FF = True Then Label1_Click
If DF = True Then Label10_Click
If Button = 1 Then DragForm Me
End Sub

Private Sub Form_Resize()
PBar.Width = Me.ScaleWidth - 181 + lCaption.Left + 3 '- lCaption.Width
PBar.Left = lCaption.Left '+ lCaption.Width + 3
PBar.Height = lCaption.Height
PBar.Top = lCaption.Top
DataGrid.Height = Me.ScaleHeight - 104 - lSection.Height
DataGrid.Width = Me.ScaleWidth - DataGrid.Left - 20
VM.Height = DataGrid.Height + 1 + lSection.Height
Abilities.Height = DataGrid.Height
Animations.Height = DataGrid.Height
Armor.Height = DataGrid.Height
Buildings.Height = DataGrid.Height
Images.Height = DataGrid.Height
Races.Height = DataGrid.Height
Shields.Height = DataGrid.Height
Sounds.Height = DataGrid.Height
Upgrades.Height = DataGrid.Height
Weapons.Height = DataGrid.Height
Abilities.Width = DataGrid.Width
Animations.Width = DataGrid.Width
Armor.Width = DataGrid.Width
Buildings.Width = DataGrid.Width
Images.Width = DataGrid.Width
Shields.Width = DataGrid.Width
Races.Width = DataGrid.Width
Sounds.Width = DataGrid.Width
Upgrades.Width = DataGrid.Width
Weapons.Width = DataGrid.Width
lSection.Width = DataGrid.Width
Shape1.Height = DataGrid.Height + 1 + lSection.Height
Shape2.Height = DataGrid.Height + lSection.Height
Shape1.Width = DataGrid.Width + 2
Shape2.Width = DataGrid.Width + 1
Line1.Width = DataGrid.Width
Image3.Top = Me.ScaleHeight - Image3.Height
Image2.Left = Me.ScaleWidth - Image2.Width
Image6.Left = Me.ScaleWidth - Image6.Width
Image4.Left = Me.ScaleWidth - Image4.Width
Image4.Top = Me.ScaleHeight - Image4.Height
For i = Image8.lbound To Image8.UBound Step 1
Image8(i).Left = Me.ScaleWidth - Image8(i).Width
Next i
For i = Image7.lbound To Image7.UBound Step 1
Image7(i).Top = Me.ScaleHeight - Image7(i).Height
Next i
CloseP.Left = Me.ScaleWidth - 19
MaxRes.Left = CloseP.Left - MaxRes.Width
Min.Left = MaxRes.Left - Min.Width
ShapeForm Me
If Me.WindowState = vbMaximized Then MaxRes.ToolTipText = "Restore" Else MaxRes.ToolTipText = "Maximize"
End Sub

Private Sub Image1_DblClick()
MaxRes_Click
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then DragForm Me
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If HF = True Then Label2_Click
If FF = True Then Label1_Click
If DF = True Then Label10_Click
End Sub

Private Sub Image10_DblClick(Index As Integer)
MaxRes_Click
End Sub

Private Sub Image10_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then DragForm Me
End Sub

Private Sub Image11_KeyDown(KeyCode As Integer, Shift As Integer)

subKeyDown KeyCode, Shift
End Sub

Private Sub Image12_KeyDown(KeyCode As Integer, Shift As Integer)

subKeyDown KeyCode, Shift
End Sub

Private Sub Image15_KeyDown(KeyCode As Integer, Shift As Integer)

subKeyDown KeyCode, Shift
End Sub

Private Sub Image16_KeyDown(KeyCode As Integer, Shift As Integer)

subKeyDown KeyCode, Shift
End Sub

Private Sub Image17_KeyDown(KeyCode As Integer, Shift As Integer)

subKeyDown KeyCode, Shift
End Sub

Private Sub Image18_KeyDown(KeyCode As Integer, Shift As Integer)

subKeyDown KeyCode, Shift
End Sub

Private Sub Image2_DblClick()
MaxRes_Click
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then DragForm Me
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If HF = True Then Label2_Click
If FF = True Then Label1_Click
If DF = True Then Label10_Click
End Sub

Private Sub Image20_KeyDown(KeyCode As Integer, Shift As Integer)

subKeyDown KeyCode, Shift
End Sub

Private Sub Image21_KeyDown(KeyCode As Integer, Shift As Integer)

subKeyDown KeyCode, Shift
End Sub

Private Sub Image22_KeyDown(KeyCode As Integer, Shift As Integer)

subKeyDown KeyCode, Shift
End Sub

Private Sub Image23_KeyDown(KeyCode As Integer, Shift As Integer)

subKeyDown KeyCode, Shift
End Sub

Private Sub Image24_KeyDown(KeyCode As Integer, Shift As Integer)

subKeyDown KeyCode, Shift
End Sub

Private Sub Image25_KeyDown(KeyCode As Integer, Shift As Integer)

subKeyDown KeyCode, Shift
End Sub

Private Sub Image29_KeyDown(KeyCode As Integer, Shift As Integer)

subKeyDown KeyCode, Shift
End Sub

Private Sub Image3_DblClick()
MaxRes_Click
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then DragForm Me
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If HF = True Then Label2_Click
If FF = True Then Label1_Click
If DF = True Then Label10_Click
End Sub

Private Sub Image4_DblClick()
MaxRes_Click
End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then DragForm Me
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If HF = True Then Label2_Click
If FF = True Then Label1_Click
If DF = True Then Label10_Click
End Sub

Private Sub Image5_DblClick(Index As Integer)
MaxRes_Click
End Sub

Private Sub Image5_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then DragForm Me
End Sub

Private Sub Image5_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If HF = True Then Label2_Click
If FF = True Then Label1_Click
If DF = True Then Label10_Click
End Sub

Private Sub Image6_DblClick()
MaxRes_Click
End Sub

Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then DragForm Me
End Sub

Private Sub Image7_DblClick(Index As Integer)
MaxRes_Click
End Sub

Private Sub Image7_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then DragForm Me
End Sub

Private Sub Image7_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If HF = True Then Label2_Click
If FF = True Then Label1_Click
If DF = True Then Label10_Click
End Sub

Private Sub Image8_DblClick(Index As Integer)
MaxRes_Click
End Sub

Private Sub Image8_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then DragForm Me
End Sub

Private Sub Image8_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If HF = True Then Label2_Click
If FF = True Then Label1_Click
If DF = True Then Label10_Click
End Sub

Private Sub Image9_DblClick()
MaxRes_Click
End Sub

Private Sub Image9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then DragForm Me
End Sub

Private Sub Images_Click()
Images_EnterCell
End Sub

Private Sub Images_EnterCell()
DGN = 15
Timer2.Enabled = True
End Sub

Private Sub Images_KeyDown(KeyCode As Integer, Shift As Integer)
subKeyDown KeyCode, Shift
End Sub

Private Sub Images_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If HF = True Then Label2_Click
If FF = True Then Label1_Click
If DF = True Then Label10_Click
End Sub

Private Sub Images_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
DataGridUpdate False, 5
End Sub

Private Sub Images_Scroll()
DataGridUpdate True, 5
End Sub

Private Sub Label1_Click()
If HF = True Then Label2_Click
If DF = True Then Label10_Click
If FF = False Then
Label5.Visible = True
Label6.Visible = True
Label7.Visible = True
Label8.Visible = True
Label9.Visible = True
Image17.Visible = True
Image18.Visible = True
Image19.Visible = True
Image20.Visible = True
Image21.Visible = True
Image22.Visible = True
Image23.Visible = True
Image24.Visible = True
Image25.Visible = True
Image29.Visible = True
FF = True
Label1.ForeColor = &H80&
Else
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Image17.Visible = False
Image18.Visible = False
Image19.Visible = False
Image20.Visible = False
Image21.Visible = False
Image22.Visible = False
Image23.Visible = False
Image24.Visible = False
Image25.Visible = False
Image29.Visible = False
FF = False
Label1.ForeColor = &HC0C0C0
End If
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &H80&
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If HF = True Then
Label2_Click
Label1_Click
End If
If DF = True Then
Label10_Click
Label1_Click
End If
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If FF = False Then Label1.ForeColor = &HC0C0C0
End Sub

Private Sub Label10_Click()
If FF = True Then Label1_Click
If HF = True Then Label1_Click
If DF = True Then
Picture2.Visible = False
Picture13.Visible = False
Picture15.Visible = False
Picture16.Visible = False
Label3.Visible = False
Label16.Visible = False
DF = False
Label10.ForeColor = &HC0C0C0
Else
Picture2.Visible = True
Picture13.Visible = True
Picture15.Visible = True
Picture16.Visible = True
Label3.Visible = True
Label16.Visible = True
DF = True
Label10.ForeColor = &H80&
End If
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.ForeColor = &H80&
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If FF = True Then
Label1_Click
Label10_Click
End If
If HF = True Then
Label2_Click
Label10_Click
End If
End Sub

Private Sub Label10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If HF = False Then Label10.ForeColor = &HC0C0C0
End Sub

Public Sub RemoveItem(Index)
Select Case Index
Case 1
If Abilities.Rows > 2 Then Abilities.RemoveItem Abilities.Row
Case 2
If Animations.Rows > 2 Then Animations.RemoveItem Animations.Row
Case 3
If Armor.Rows > 2 Then Armor.RemoveItem Armor.Row
Case 4
If Buildings.Rows > 2 Then Buildings.RemoveItem Buildings.Row
Case 5
If Images.Rows > 2 Then Images.RemoveItem Images.Row
Case 6
If DataGrid.Rows > 2 Then DataGrid.RemoveItem DataGrid.Row
Case 7
If Races.Rows > 2 Then Races.RemoveItem Races.Row
Case 8
If Shields.Rows > 2 Then Shields.RemoveItem Shields.Row
Case 9
If Len(Sounds.TextMatrix(Sounds.Row, 2)) > 18 Then
If Mid(Sounds.TextMatrix(Sounds.Row, 2), 1, 18) = "SoundArrayVariable" Then
RemoveSound CInt(Mid(Sounds.TextMatrix(Sounds.Row, 2), 19))
If Sounds.Rows > 2 Then Sounds.RemoveItem Sounds.Row
Else
If Sounds.Rows > 2 Then Sounds.RemoveItem Sounds.Row
End If
Else
If Sounds.Rows > 2 Then Sounds.RemoveItem Sounds.Row
End If
Case 10
If Upgrades.Rows > 2 Then Upgrades.RemoveItem Upgrades.Row
Case 11
If Weapons.Rows > 2 Then Weapons.RemoveItem Weapons.Row
End Select
End Sub

Private Sub Label16_Click()
Select Case Label16.Caption
Case "&Remove Ability"
RemoveItem 1
Case "&Remove Animation"
RemoveItem 2
Case "&Remove Armor"
RemoveItem 3
Case "&Remove Building"
RemoveItem 4
Case "&Remove Image"
RemoveItem 5
Case "&Remove Infantry"
RemoveItem 6
Case "&Remove Race"
RemoveItem 7
Case "&Remove Shield"
RemoveItem 8
Case "&Remove Sound"
RemoveItem 9
Case "&Remove Upgrade"
RemoveItem 10
Case "&Remove Weapon"
RemoveItem 11
End Select
Picture2.Visible = False
Picture13.Visible = False
Picture15.Visible = False
Picture16.Visible = False
Label3.Visible = False
Label16.Visible = False
DF = False
Label10.ForeColor = &HC0C0C0
End Sub

Private Sub Label16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.ForeColor = &H40C0&
End Sub

Private Sub Label16_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If HF = False Then Label16.ForeColor = &H808080
End Sub

Private Sub Label2_Click()
If FF = True Then Label1_Click
If DF = True Then Label10_Click
If HF = False Then
Image15.Visible = True
Image14.Visible = True
Label4.Visible = True
Image16.Visible = True
HF = True
Label2.ForeColor = &H80&
Else
Image15.Visible = False
Image14.Visible = False
Label4.Visible = False
Image16.Visible = False
HF = False
Label2.ForeColor = &HC0C0C0
End If
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &H80&
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If FF = True Then
Label1_Click
Label2_Click
End If
If DF = True Then
Label10_Click
Label2_Click
End If
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If HF = False Then Label2.ForeColor = &HC0C0C0
End Sub

Public Sub AddItem(Index)
Select Case Index
Case 1 'Ability
Abilities.AddItem "", Abilities.Row + 1
Abilities.TextMatrix(Abilities.Row + 1, 0) = "Untitled Ability" & aCount(Index)
Abilities.TextMatrix(Abilities.Row + 1, 1) = "Untitled Ability" & aCount(Index)
Abilities.TextMatrix(Abilities.Row + 1, 2) = ""
Abilities.TextMatrix(Abilities.Row + 1, 3) = "0"
Abilities.TextMatrix(Abilities.Row + 1, 4) = "0"
Abilities.TextMatrix(Abilities.Row + 1, 5) = "0"
Abilities.TextMatrix(Abilities.Row + 1, 6) = "0"
Abilities.TextMatrix(Abilities.Row + 1, 7) = "0"
Abilities.TextMatrix(Abilities.Row + 1, 8) = "0"
Abilities.TextMatrix(Abilities.Row + 1, 9) = ""
Case 2 'Animation
Animations.AddItem "", Animations.Row + 1
Animations.TextMatrix(Animations.Row + 1, 0) = "Untitled Animation" & aCount(Index)
Animations.TextMatrix(Animations.Row + 1, 1) = "Untitled Animation" & aCount(Index)
Animations.TextMatrix(Animations.Row + 1, 2) = "12"
Animations.TextMatrix(Animations.Row + 1, 3) = "0-Building Build"
Animations.TextMatrix(Animations.Row + 1, 4) = "0"
Animations.TextMatrix(Animations.Row + 1, 5) = "[none]"
Animations.TextMatrix(Animations.Row + 1, 6) = "[none]"
Animations.TextMatrix(Animations.Row + 1, 7) = "[none]"
Animations.TextMatrix(Animations.Row + 1, 8) = "[none]"
Animations.TextMatrix(Animations.Row + 1, 9) = "[none]"
Animations.TextMatrix(Animations.Row + 1, 10) = "[none]"
Animations.TextMatrix(Animations.Row + 1, 11) = "[none]"
Animations.TextMatrix(Animations.Row + 1, 12) = "[none]"
Animations.TextMatrix(Animations.Row + 1, 13) = "[none]"
Animations.TextMatrix(Animations.Row + 1, 14) = "[none]"
Animations.TextMatrix(Animations.Row + 1, 15) = "[none]"
Animations.TextMatrix(Animations.Row + 1, 16) = "[none]"
Animations.TextMatrix(Animations.Row + 1, 17) = "-1"
Animations.TextMatrix(Animations.Row + 1, 18) = "0-False"
Case 3 'Armor
Armor.AddItem "", Armor.Row + 1
Armor.TextMatrix(Armor.Row + 1, 0) = "Untitled Armor" & aCount(Index)
Armor.TextMatrix(Armor.Row + 1, 1) = "Untitled Armor" & aCount(Index)
Armor.TextMatrix(Armor.Row + 1, 2) = "0-Cloak Apperance"
Armor.TextMatrix(Armor.Row + 1, 3) = "0"
Armor.TextMatrix(Armor.Row + 1, 4) = "0"
Armor.TextMatrix(Armor.Row + 1, 5) = "0-Black"
Case 4 'Building
Buildings.AddItem "", Buildings.Row + 1
Buildings.TextMatrix(Buildings.Row + 1, 0) = "Untitled Building" & aCount(Index)
Buildings.TextMatrix(Buildings.Row + 1, 1) = "Untitled Building" & aCount(Index)
Buildings.TextMatrix(Buildings.Row + 1, 2) = ""
Buildings.TextMatrix(Buildings.Row + 1, 3) = ""
Buildings.TextMatrix(Buildings.Row + 1, 4) = "1-True"
Buildings.TextMatrix(Buildings.Row + 1, 5) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 6) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 7) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 8) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 9) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 10) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 11) = "0,0"
Buildings.TextMatrix(Buildings.Row + 1, 12) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 13) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 14) = "0"
Buildings.TextMatrix(Buildings.Row + 1, 15) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 16) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 17) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 18) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 19) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 20) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 21) = "0-False"
Buildings.TextMatrix(Buildings.Row + 1, 22) = "1-True"
Buildings.TextMatrix(Buildings.Row + 1, 23) = "0-False"
Buildings.TextMatrix(Buildings.Row + 1, 24) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 25) = "0"
Buildings.TextMatrix(Buildings.Row + 1, 26) = ""
Buildings.TextMatrix(Buildings.Row + 1, 27) = "0-False"
Buildings.TextMatrix(Buildings.Row + 1, 28) = "0-False"
Buildings.TextMatrix(Buildings.Row + 1, 29) = "0-False"
Buildings.TextMatrix(Buildings.Row + 1, 30) = "0-False"
Buildings.TextMatrix(Buildings.Row + 1, 31) = "0-True"
Buildings.TextMatrix(Buildings.Row + 1, 32) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 33) = "0-False"
Buildings.TextMatrix(Buildings.Row + 1, 34) = "0-False"
Buildings.TextMatrix(Buildings.Row + 1, 35) = "0"
Buildings.TextMatrix(Buildings.Row + 1, 36) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 37) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 38) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 39) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 40) = "0,0"
Buildings.TextMatrix(Buildings.Row + 1, 41) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 42) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 43) = "0"
Buildings.TextMatrix(Buildings.Row + 1, 44) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 45) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 46) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 47) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 48) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 49) = "0"
Buildings.TextMatrix(Buildings.Row + 1, 50) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 51) = "0-Kill Larva"
Buildings.TextMatrix(Buildings.Row + 1, 52) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 53) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 54) = ""
Buildings.TextMatrix(Buildings.Row + 1, 55) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 56) = "0"
Buildings.TextMatrix(Buildings.Row + 1, 57) = "0"
Buildings.TextMatrix(Buildings.Row + 1, 58) = "0"
Buildings.TextMatrix(Buildings.Row + 1, 59) = "0-False"
Buildings.TextMatrix(Buildings.Row + 1, 60) = "0"
Buildings.TextMatrix(Buildings.Row + 1, 61) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 62) = "0-False"
Buildings.TextMatrix(Buildings.Row + 1, 63) = "0-False"
Buildings.TextMatrix(Buildings.Row + 1, 64) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 65) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 66) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 67) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 68) = "0,0"
Buildings.TextMatrix(Buildings.Row + 1, 69) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 70) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 71) = "0"
Buildings.TextMatrix(Buildings.Row + 1, 72) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 73) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 74) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 75) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 76) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 77) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 78) = "0"
Buildings.TextMatrix(Buildings.Row + 1, 79) = "0"
Buildings.TextMatrix(Buildings.Row + 1, 80) = "0"
Buildings.TextMatrix(Buildings.Row + 1, 81) = "0"
Buildings.TextMatrix(Buildings.Row + 1, 82) = "1-True"
Buildings.TextMatrix(Buildings.Row + 1, 83) = "0"
Buildings.TextMatrix(Buildings.Row + 1, 84) = ""
Buildings.TextMatrix(Buildings.Row + 1, 85) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 86) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 87) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 88) = "[none]"
Buildings.TextMatrix(Buildings.Row + 1, 89) = ""
Buildings.TextMatrix(Buildings.Row + 1, 90) = "0-False"
Buildings.TextMatrix(Buildings.Row + 1, 91) = ""
Buildings.TextMatrix(Buildings.Row + 1, 92) = ""
Buildings.TextMatrix(Buildings.Row + 1, 93) = ""
Buildings.TextMatrix(Buildings.Row + 1, 94) = ""
Buildings.TextMatrix(Buildings.Row + 1, 95) = "0-False"
Buildings.TextMatrix(Buildings.Row + 1, 96) = "[none]"
Case 5 'Image
Images.AddItem "", Images.Row + 1
Images.TextMatrix(Images.Row + 1, 0) = "Untitled Image" & aCount(Index)
Images.TextMatrix(Images.Row + 1, 1) = "Untitled Image" & aCount(Index)
Images.TextMatrix(Images.Row + 1, 2) = ""
Case 6 'Infantry
DataGrid.AddItem "", DataGrid.Row + 1
DataGrid.TextMatrix(DataGrid.Row + 1, 0) = "Untitled Infantry Unit" & aCount(Index)
DataGrid.TextMatrix(DataGrid.Row + 1, 1) = "Untitled Infantry Unit" & aCount(Index)
DataGrid.TextMatrix(DataGrid.Row + 1, 2) = ""
DataGrid.TextMatrix(DataGrid.Row + 1, 3) = ""
DataGrid.TextMatrix(DataGrid.Row + 1, 4) = ""
DataGrid.TextMatrix(DataGrid.Row + 1, 5) = "0,0"
DataGrid.TextMatrix(DataGrid.Row + 1, 6) = "0,0"
DataGrid.TextMatrix(DataGrid.Row + 1, 7) = "0"
DataGrid.TextMatrix(DataGrid.Row + 1, 8) = "0"
DataGrid.TextMatrix(DataGrid.Row + 1, 9) = "0"
DataGrid.TextMatrix(DataGrid.Row + 1, 10) = "0"
DataGrid.TextMatrix(DataGrid.Row + 1, 11) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 12) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 13) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 14) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 15) = "0-False"
DataGrid.TextMatrix(DataGrid.Row + 1, 16) = "0-False"
DataGrid.TextMatrix(DataGrid.Row + 1, 17) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 18) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 19) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 20) = "0-[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 21) = "0-[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 22) = "0-False"
DataGrid.TextMatrix(DataGrid.Row + 1, 23) = "0-False"
DataGrid.TextMatrix(DataGrid.Row + 1, 24) = "0,0,1,1"
DataGrid.TextMatrix(DataGrid.Row + 1, 25) = "0,0,1,1"
DataGrid.TextMatrix(DataGrid.Row + 1, 26) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 27) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 28) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 29) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 30) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 31) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 32) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 33) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 34) = ""
DataGrid.TextMatrix(DataGrid.Row + 1, 35) = ""
DataGrid.TextMatrix(DataGrid.Row + 1, 36) = "0-False"
DataGrid.TextMatrix(DataGrid.Row + 1, 37) = "0-False"
DataGrid.TextMatrix(DataGrid.Row + 1, 38) = "0"
DataGrid.TextMatrix(DataGrid.Row + 1, 39) = "0"
DataGrid.TextMatrix(DataGrid.Row + 1, 40) = "0"
DataGrid.TextMatrix(DataGrid.Row + 1, 41) = "0"
DataGrid.TextMatrix(DataGrid.Row + 1, 42) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 43) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 44) = "0"
DataGrid.TextMatrix(DataGrid.Row + 1, 45) = "0"
DataGrid.TextMatrix(DataGrid.Row + 1, 46) = "1-Infantry"
DataGrid.TextMatrix(DataGrid.Row + 1, 47) = "1-Infantry"
DataGrid.TextMatrix(DataGrid.Row + 1, 48) = "0,0"
DataGrid.TextMatrix(DataGrid.Row + 1, 49) = "0,0"
DataGrid.TextMatrix(DataGrid.Row + 1, 50) = "0"
DataGrid.TextMatrix(DataGrid.Row + 1, 51) = "0"
DataGrid.TextMatrix(DataGrid.Row + 1, 52) = "0"
DataGrid.TextMatrix(DataGrid.Row + 1, 53) = "0"
DataGrid.TextMatrix(DataGrid.Row + 1, 54) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 55) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 56) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 57) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 58) = "0"
DataGrid.TextMatrix(DataGrid.Row + 1, 59) = "0"
DataGrid.TextMatrix(DataGrid.Row + 1, 60) = "0"
DataGrid.TextMatrix(DataGrid.Row + 1, 61) = "0"
DataGrid.TextMatrix(DataGrid.Row + 1, 62) = "0"
DataGrid.TextMatrix(DataGrid.Row + 1, 63) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 64) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 65) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 66) = "0"
DataGrid.TextMatrix(DataGrid.Row + 1, 67) = "0"
DataGrid.TextMatrix(DataGrid.Row + 1, 68) = "0"
DataGrid.TextMatrix(DataGrid.Row + 1, 69) = "0"
DataGrid.TextMatrix(DataGrid.Row + 1, 70) = "0"
DataGrid.TextMatrix(DataGrid.Row + 1, 71) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 72) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 73) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 74) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 75) = "0-1-->2"
DataGrid.TextMatrix(DataGrid.Row + 1, 76) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 77) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 78) = "0-Biological"
DataGrid.TextMatrix(DataGrid.Row + 1, 79) = "0-Biological"
DataGrid.TextMatrix(DataGrid.Row + 1, 80) = "0,0"
DataGrid.TextMatrix(DataGrid.Row + 1, 81) = "0,0"
DataGrid.TextMatrix(DataGrid.Row + 1, 82) = "0"
DataGrid.TextMatrix(DataGrid.Row + 1, 83) = "0"
DataGrid.TextMatrix(DataGrid.Row + 1, 84) = "0"
DataGrid.TextMatrix(DataGrid.Row + 1, 85) = "0"
DataGrid.TextMatrix(DataGrid.Row + 1, 86) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 87) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 88) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 89) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 90) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 91) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 92) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 93) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 94) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 95) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 96) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 97) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 98) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 99) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 100) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 101) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 102) = "0-Air"
DataGrid.TextMatrix(DataGrid.Row + 1, 103) = "0-Air"
DataGrid.TextMatrix(DataGrid.Row + 1, 104) = "0"
DataGrid.TextMatrix(DataGrid.Row + 1, 105) = "0"
DataGrid.TextMatrix(DataGrid.Row + 1, 106) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 107) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 108) = "False"
DataGrid.TextMatrix(DataGrid.Row + 1, 109) = "False"
DataGrid.TextMatrix(DataGrid.Row + 1, 110) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 111) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 112) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 113) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 114) = "0-Biological"
DataGrid.TextMatrix(DataGrid.Row + 1, 115) = "0-Biological"
DataGrid.TextMatrix(DataGrid.Row + 1, 116) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 117) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 118) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 119) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 120) = "0-False"
DataGrid.TextMatrix(DataGrid.Row + 1, 121) = "0-False"
DataGrid.TextMatrix(DataGrid.Row + 1, 122) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 123) = "[none]"
DataGrid.TextMatrix(DataGrid.Row + 1, 124) = "0"
DataGrid.TextMatrix(DataGrid.Row + 1, 125) = "0"
DataGrid.TextMatrix(DataGrid.Row + 1, 126) = ""
DataGrid.TextMatrix(DataGrid.Row + 1, 127) = ""
DataGrid.TextMatrix(DataGrid.Row + 1, 128) = "0-False"
DataGrid.TextMatrix(DataGrid.Row + 1, 129) = "0-False"
DataGrid.TextMatrix(DataGrid.Row + 1, 130) = "0-False"
DataGrid.TextMatrix(DataGrid.Row + 1, 131) = "0-False"
DataGrid.TextMatrix(DataGrid.Row + 1, 132) = "0"
DataGrid.TextMatrix(DataGrid.Row + 1, 133) = "0-False"
Case 7 'Race
Races.AddItem "", Races.Row + 1
Races.TextMatrix(Races.Row + 1, 0) = "Untitled Race" & aCount(Index)
Races.TextMatrix(Races.Row + 1, 1) = "Untitled Race" & aCount(Index)
Races.TextMatrix(Races.Row + 1, 2) = "1"
Races.TextMatrix(Races.Row + 1, 3) = "1"
Races.TextMatrix(Races.Row + 1, 4) = "1"
Races.TextMatrix(Races.Row + 1, 5) = "1"
Races.TextMatrix(Races.Row + 1, 6) = "300"
Races.TextMatrix(Races.Row + 1, 7) = ""
Races.TextMatrix(Races.Row + 1, 8) = ""
Case 8 'Shield
Shields.AddItem "", Shields.Row + 1
Shields.TextMatrix(Shields.Row + 1, 0) = "Untitled Shield" & aCount(Index)
Shields.TextMatrix(Shields.Row + 1, 1) = "Untitled Shield" & aCount(Index)
Shields.TextMatrix(Shields.Row + 1, 2) = "0"
Shields.TextMatrix(Shields.Row + 1, 3) = "0-Black"
Shields.TextMatrix(Shields.Row + 1, 4) = "0-Cloak Apperance"
Shields.TextMatrix(Shields.Row + 1, 5) = "0"
Case 9 'Sound
Sounds.AddItem "", Sounds.Row + 1
Sounds.TextMatrix(Sounds.Row + 1, 0) = "Untitled Sound" & aCount(Index)
Sounds.TextMatrix(Sounds.Row + 1, 1) = "Untitled Sound" & aCount(Index)
Sounds.TextMatrix(Sounds.Row + 1, 2) = ""
Sounds.TextMatrix(Sounds.Row + 1, 3) = "-1"
Case 10 'Upgrade
Upgrades.AddItem "", Upgrades.Row + 1
Upgrades.TextMatrix(Upgrades.Row + 1, 0) = "Untitled Upgrade" & aCount(Index)
Upgrades.TextMatrix(Upgrades.Row + 1, 1) = "Untitled Upgrade" & aCount(Index)
Upgrades.TextMatrix(Upgrades.Row + 1, 2) = "0"
Upgrades.TextMatrix(Upgrades.Row + 1, 3) = "0"
Upgrades.TextMatrix(Upgrades.Row + 1, 4) = "0"
Upgrades.TextMatrix(Upgrades.Row + 1, 5) = "0"
Upgrades.TextMatrix(Upgrades.Row + 1, 6) = ""
Upgrades.TextMatrix(Upgrades.Row + 1, 7) = "0"
Upgrades.TextMatrix(Upgrades.Row + 1, 8) = "0"
Upgrades.TextMatrix(Upgrades.Row + 1, 9) = ""
Case 11 'Weapon
Weapons.AddItem "", Weapons.Row + 1
Weapons.TextMatrix(Weapons.Row + 1, 0) = "Untitled Weapon" & aCount(Index)
Weapons.TextMatrix(Weapons.Row + 1, 1) = "Untitled Weapon" & aCount(Index)
Weapons.TextMatrix(Weapons.Row + 1, 2) = "[none]"
Weapons.TextMatrix(Weapons.Row + 1, 3) = "[none]"
Weapons.TextMatrix(Weapons.Row + 1, 4) = "0"
Weapons.TextMatrix(Weapons.Row + 1, 5) = "[none]"
Weapons.TextMatrix(Weapons.Row + 1, 6) = "[none]"
Weapons.TextMatrix(Weapons.Row + 1, 7) = "0"
Weapons.TextMatrix(Weapons.Row + 1, 8) = "0"
Weapons.TextMatrix(Weapons.Row + 1, 9) = "0"
Weapons.TextMatrix(Weapons.Row + 1, 10) = "0"
Weapons.TextMatrix(Weapons.Row + 1, 11) = "0"
Weapons.TextMatrix(Weapons.Row + 1, 12) = "0"
Weapons.TextMatrix(Weapons.Row + 1, 13) = "0-Bottom Left"
Weapons.TextMatrix(Weapons.Row + 1, 14) = "[none]"
Weapons.TextMatrix(Weapons.Row + 1, 15) = "0"
Weapons.TextMatrix(Weapons.Row + 1, 16) = "[none]"
Weapons.TextMatrix(Weapons.Row + 1, 17) = "[none]"
Weapons.TextMatrix(Weapons.Row + 1, 18) = "[none]"
Weapons.TextMatrix(Weapons.Row + 1, 19) = "0"
Weapons.TextMatrix(Weapons.Row + 1, 20) = "0"
Weapons.TextMatrix(Weapons.Row + 1, 21) = "0"
Weapons.TextMatrix(Weapons.Row + 1, 22) = "0"
Weapons.TextMatrix(Weapons.Row + 1, 23) = "0"
Weapons.TextMatrix(Weapons.Row + 1, 24) = "0"
Weapons.TextMatrix(Weapons.Row + 1, 25) = "[none]"
Weapons.TextMatrix(Weapons.Row + 1, 26) = "0"
Weapons.TextMatrix(Weapons.Row + 1, 27) = "[none]"
Weapons.TextMatrix(Weapons.Row + 1, 28) = "[none]"
Weapons.TextMatrix(Weapons.Row + 1, 29) = "[none]"
Weapons.TextMatrix(Weapons.Row + 1, 30) = "0"
Weapons.TextMatrix(Weapons.Row + 1, 31) = "0"
Weapons.TextMatrix(Weapons.Row + 1, 32) = "0"
Weapons.TextMatrix(Weapons.Row + 1, 33) = "0"
Weapons.TextMatrix(Weapons.Row + 1, 34) = "0"
Weapons.TextMatrix(Weapons.Row + 1, 35) = "0"
Weapons.TextMatrix(Weapons.Row + 1, 36) = "[none]"
Weapons.TextMatrix(Weapons.Row + 1, 37) = "0-[none]"
Weapons.TextMatrix(Weapons.Row + 1, 38) = "0-Air"
Weapons.TextMatrix(Weapons.Row + 1, 39) = "0-Normal"
End Select
aCount(Index) = aCount(Index) + 1
End Sub

Private Sub Label3_Click()
Select Case Label3.Caption
Case "&Add Ability"
AddItem 1
Case "&Add Animation"
AddItem 2
Case "&Add Armor"
AddItem 3
Case "&Add Building"
AddItem 4
Case "&Add Image"
AddItem 5
Case "&Add Infantry"
AddItem 6
Case "&Add Race"
AddItem 7
Case "&Add Shield"
AddItem 8
Case "&Add Sound"
If Sounds.Rows > 1024 Then
Msbox "You have reached the maximum number of sounds(1024)", "Thats A Lot Of Sound!"
Else
AddItem 9
End If
Case "&Add Upgrade"
AddItem 10
Case "&Add Weapon"
AddItem 11
End Select
Picture2.Visible = False
Picture13.Visible = False
Picture15.Visible = False
Picture16.Visible = False
Label3.Visible = False
Label16.Visible = False
DF = False
Label10.ForeColor = &HC0C0C0
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = &H40C0&
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = &H808080
End Sub

Private Sub Label4_Click()
Image15.Visible = False
Image14.Visible = False
Label4.Visible = False
Image16.Visible = False
HF = False
Label2.ForeColor = &HC0C0C0
frmAbout.Show vbModal, Me
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &H40C0&
End Sub

Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &H808080
End Sub

Private Sub Label5_Click()
CreateMod
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Image17.Visible = False
Image18.Visible = False
Image19.Visible = False
Image20.Visible = False
Image21.Visible = False
Image22.Visible = False
Image23.Visible = False
Image24.Visible = False
Image25.Visible = False
Image29.Visible = False
FF = False
Label1.ForeColor = &HC0C0C0
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.ForeColor = &H40C0&
End Sub

Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.ForeColor = &H808080
End Sub

Private Sub Label6_Click()
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Image17.Visible = False
Image18.Visible = False
Image19.Visible = False
Image20.Visible = False
Image21.Visible = False
Image22.Visible = False
Image23.Visible = False
Image24.Visible = False
Image25.Visible = False
Image29.Visible = False
FF = False
Label1.ForeColor = &HC0C0C0
FileSelect.Visible = False: PlayView.Visible = False
CD.FileName = ModFileName
CD.Filter = "NAVEN Mod Files(*.mdf)|*.mdf|All Files(*.*)|*.*;*"
CD.ShowOpen
If CD.Flags <> 0 Then
If OpenMod(CD.FileName) = False Then Msbox "An error occoured opening the selected file.", "Invalid File"
End If
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = &H40C0&
End Sub

Private Sub Label6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = &H808080
End Sub

Private Sub Label7_Click()
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Image17.Visible = False
Image18.Visible = False
Image19.Visible = False
Image20.Visible = False
Image21.Visible = False
Image22.Visible = False
Image23.Visible = False
Image24.Visible = False
Image25.Visible = False
Image29.Visible = False
FF = False
Label1.ForeColor = &HC0C0C0
If Len(ModFileName) < 1 Then
FileSelect.Visible = False: PlayView.Visible = False
CD.FileName = ModFileName
CD.Filter = "NAVEN Mod Files(*.mdf)|*.mdf|All Files(*.*)|*.*;*"
CD.ShowSave
If CD.Flags <> 0 Then
ModFileName = CD.FileName
Else
Exit Sub
End If
End If
If SaveMod(ModFileName) = False Then Msbox "Error Saving File.", "File Save Error"
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.ForeColor = &H40C0&
End Sub

Private Sub Label7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.ForeColor = &H808080
End Sub

Private Sub Label8_Click()
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Image17.Visible = False
Image18.Visible = False
Image19.Visible = False
Image20.Visible = False
Image21.Visible = False
Image22.Visible = False
Image23.Visible = False
Image24.Visible = False
Image25.Visible = False
Image29.Visible = False
FF = False
Label1.ForeColor = &HC0C0C0
FileSelect.Visible = False: PlayView.Visible = False
CD.FileName = ModFileName
CD.Filter = "NAVEN Mod Files(*.mdf)|*.mdf|All Files(*.*)|*.*;*"
CD.ShowSave
If CD.Flags <> 0 Then
ModFileName = CD.FileName
Else
Exit Sub
End If
If SaveMod(ModFileName) = False Then Msbox "Error Saving File.", "File Save Error"
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.ForeColor = &H40C0&
End Sub

Private Sub Label8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.ForeColor = &H808080
End Sub

Private Sub Label9_Click()
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Image17.Visible = False
Image18.Visible = False
Image19.Visible = False
Image20.Visible = False
Image21.Visible = False
Image22.Visible = False
Image23.Visible = False
Image24.Visible = False
Image25.Visible = False
Image29.Visible = False
FF = False
Label1.ForeColor = &HC0C0C0
Unload Form1
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.ForeColor = &H40C0&
End Sub

Private Sub Label9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.ForeColor = &H808080
End Sub

Private Sub LBox_Change()
If HF = True Then Label2_Click
If FF = True Then Label1_Click
If DF = True Then Label10_Click
UpdateLBox
End Sub

Private Sub LBox_Click()
If HF = True Then Label2_Click
If FF = True Then Label1_Click
If DF = True Then Label10_Click
UpdateLBox
End Sub

Private Sub LBox_KeyDown(KeyCode As Integer, Shift As Integer)
UpdateLBox
subKeyDown KeyCode, Shift
Select Case KeyCode
Case 37
If DataGrid.Col > 1 Then DataGrid.Col = DataGrid.Col - 1
'Case 38
'If DataGrid.Row > 1 Then DataGrid.Row = DataGrid.Row - 1
Case 39
If DataGrid.Col < DataGrid.Cols - 2 Then DataGrid.Col = DataGrid.Col + 1
'Case 40
'If DataGrid.Row < DataGrid.Rows - 2 Then DataGrid.Row = DataGrid.Row + 1
End Select
End Sub

Private Sub LBox_KeyPress(KeyAscii As Integer)
UpdateLBox
If KeyAscii = 10 Or KeyAscii = 13 Then LBox_KeyDown 40, 0
End Sub

Private Sub LBox_Scroll()
UpdateLBox
End Sub

Private Sub lCaption_DblClick()
MaxRes_Click
End Sub

Private Sub lCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then DragForm Me
End Sub

Private Sub lCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If HF = True Then Label2_Click
If FF = True Then Label1_Click
If DF = True Then Label10_Click
End Sub

Private Sub MaxRes_Click()
If Me.WindowState = vbNormal Then
Me.WindowState = vbMaximized
Else
Me.WindowState = vbNormal
End If
End Sub

Private Sub MaxRes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set MaxRes.Picture = MaxResD.Picture
End Sub

Private Sub MaxRes_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set MaxRes.Picture = MaxResU.Picture
End Sub

Private Sub Min_Click()
Form1.WindowState = vbMinimized
End Sub

Private Sub Min_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set Min.Picture = MinD.Picture
End Sub

Private Sub Min_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set Min.Picture = MinU.Picture
End Sub

Private Sub mnuEditCut_Click()
MsgBox 1
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)

subKeyDown KeyCode, Shift
End Sub

Private Sub Picture12_DblClick()
MaxRes_Click
End Sub

Private Sub Picture12_KeyDown(KeyCode As Integer, Shift As Integer)

subKeyDown KeyCode, Shift
End Sub

Private Sub Picture2_KeyDown(KeyCode As Integer, Shift As Integer)

subKeyDown KeyCode, Shift
End Sub

Private Sub PlayView_Click()
If Len(CControl.Text) < 3 Then Exit Sub
Load frmViewImage
Dim a As Variant
Dim b() As Variant
a = Split(CControl.Text, Chr(170))
ReDim b(UBound(a))
For i = 0 To UBound(a) - 1 Step 1
b(i) = Split(a(i), Chr(171))
Next i
For X = 0 To UBound(a) - 1 Step 1
For Y = 0 To UBound(b(0)) - 1 Step 1
frmViewImage.Picture2.PSet (X, Y), b(X)(Y)
Next Y
Next X
frmViewImage.Picture2.Width = UBound(a) - 1
frmViewImage.Picture2.Height = UBound(b(0)) - 1
frmViewImage.Picture2.Left = (frmViewImage.Picture1.ScaleWidth / 2) - (frmViewImage.Picture2.Width / 2)
frmViewImage.Picture2.Top = (frmViewImage.Picture1.ScaleHeight / 2) - (frmViewImage.Picture2.Height / 2)
frmViewImage.Show vbModal, Me
End Sub

Private Sub Races_Click()
Races_EnterCell
End Sub

Private Sub Races_EnterCell()
DGN = 21
Timer2.Enabled = True
End Sub

Private Sub Races_KeyDown(KeyCode As Integer, Shift As Integer)
subKeyDown KeyCode, Shift
End Sub

Private Sub Races_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If HF = True Then Label2_Click
If FF = True Then Label1_Click
If DF = True Then Label10_Click
End Sub

Private Sub Races_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
DataGridUpdate True, 11
End Sub

Private Sub Races_Scroll()
DataGridUpdate True, 11
End Sub

Private Sub Shields_Click()
Shields_EnterCell
End Sub

Private Sub Shields_EnterCell()
DGN = 17
Timer2.Enabled = True
End Sub

Private Sub Shields_KeyDown(KeyCode As Integer, Shift As Integer)
subKeyDown KeyCode, Shift
End Sub

Private Sub Shields_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If HF = True Then Label2_Click
If FF = True Then Label1_Click
If DF = True Then Label10_Click
End Sub

Private Sub Shields_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
DataGridUpdate True, 7
End Sub

Private Sub Shields_Scroll()
DataGridUpdate True, 7
End Sub

Private Sub Sounds_Click()
Sounds_EnterCell
End Sub

Private Sub Sounds_EnterCell()
DGN = 18
Timer2.Enabled = True
End Sub

Private Sub Sounds_KeyDown(KeyCode As Integer, Shift As Integer)
subKeyDown KeyCode, Shift
End Sub

Private Sub Sounds_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If HF = True Then Label2_Click
If FF = True Then Label1_Click
If DF = True Then Label10_Click
End Sub

Private Sub Sounds_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
DataGridUpdate True, 8
End Sub

Private Sub Sounds_Scroll()
DataGridUpdate True, 8
End Sub

Private Sub T_DblClick(Index As Integer)
MaxRes_Click
End Sub

Private Sub T_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then DragForm Me
End Sub

Private Sub T_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then
If Y < (T(0).Height / 4) * 3 Then
If HF = True Then Label2_Click
If FF = True Then Label1_Click
If DF = True Then Label10_Click
End If
Else
If HF = True Then Label2_Click
If FF = True Then Label1_Click
If DF = True Then Label10_Click
End If
End Sub

Public Function IsNumber(Text) As Boolean
On Error GoTo ItsAString
a = CLng(Text)
IsNumber = True
Exit Function
ItsAString:
IsNumber = False
Exit Function
End Function

Public Sub UpdateLBox()
CControl.TextMatrix(CControl.Row, CControl.Col) = LBox.Text
End Sub

Private Sub TextEdit_Change()
With CControl
If Len(TextEdit.Text) > 0 Then
If TextIsNumber = True Then
If IsNumber(TextEdit.Text) = True Then
.TextMatrix(.Row, .Col) = TextEdit.Text
Else
Msbox "The selected field is suppose to be a number!", "Invaild Field"
End If
Else
If LCase(Mid(CControl.TextMatrix(0, CControl.Col), 1, 5)) = "[name" Then
If LCase(TextEdit.Text) = "[none]" Then
Msbox "The selected field is a name field and cannot be set to ""[none]""", "Invalid Field"
Else
If InStr(1, CControl.TextMatrix(CControl.Row, CControl.Col), ";", vbTextCompare) <> 0 Then
Msbox "The selected field is a name field and cannot contain the charactor "";""!", "Invalid Field"
Else
If IsName(TextEdit.Text) = True Then
Msbox "The name """ & TextEdit.Text & """ Already Exists!", "Field Already Exists"
Else
If .Col = 1 Then .TextMatrix(.Row, 0) = TextEdit.Text
.TextMatrix(.Row, .Col) = TextEdit.Text
End If
End If
End If
Else
.TextMatrix(.Row, .Col) = TextEdit.Text
End If
End If
Else
If TextIsNumber = False Then
.TextMatrix(.Row, .Col) = ""
Else
.TextMatrix(.Row, .Col) = "0"
End If
End If
End With
End Sub

Public Function IsName(Text) As Boolean
Dim a As Boolean
a = False
b = 0
For i = 1 To CControl.Rows - 1 Step 1
If CControl.TextMatrix(i, 1) = Text Then
b = b + 1
End If
Next i
If b > 1 Then a = True
IsName = a
End Function

Private Sub TextEdit_Click()
If HF = True Then Label2_Click
If FF = True Then Label1_Click
If DF = True Then Label10_Click
End Sub

Private Sub TextEdit_KeyDown(KeyCode As Integer, Shift As Integer)
'shift=2=ctrl||shift=4=alt
subKeyDown KeyCode, Shift
Select Case KeyCode
'Case 37
'If DataGrid.Col > 1 Then DataGrid.Col = DataGrid.Col - 1
Case 38
If DataGrid.Row > 1 Then DataGrid.Row = DataGrid.Row - 1
'Case 39
'If DataGrid.Col < DataGrid.Cols - 2 Then DataGrid.Col = DataGrid.Col + 1
Case 40
If DataGrid.Row < DataGrid.Rows - 2 Then DataGrid.Row = DataGrid.Row + 1
End Select
End Sub

Private Sub TextEdit_KeyPress(KeyAscii As Integer)
If KeyAscii = 10 Or KeyAscii = 13 Then TextEdit_KeyDown 40, 0
End Sub

Private Sub Timer1_Timer()
If LastState <> Me.WindowState Then
LastState = Me.WindowState
Form_Resize
End If
'For i = 1 To Toolbar1.Buttons.Count Step 1
'If Toolbar1.Buttons(i).Value = tbrPressed Then Exit For
'Next i
'DataGridUpdate True, i
End Sub

Private Sub Timer2_Timer()
DataGridUpdate False, DGN - 10
Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
OnTimer = True
TSv = TSv + 1
If TSv < 12 Then
Toolbar1.Buttons(TSv).Value = tbrPressed
AddItem TSv
ElseIf TSv < 23 Then
Toolbar1.Buttons(TSv - 11).Value = tbrPressed
RemoveItem TSv - 11
ElseIf TSv < 29 Then
Select Case TSv
Case 24
lSection.Visible = True
Case 26
Line1.Visible = True
Case 28
DataGrid.Visible = True
End Select
Toolbar1.Buttons(TSv - 22).Value = tbrPressed
Else
Toolbar1_ButtonClick Toolbar1.Buttons(6)
Timer3.Enabled = False
End If
OnTimer = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If HF = True Then Label2_Click
If FF = True Then Label1_Click
If DF = True Then Label10_Click
Set Toolbar1.ImageList = Nothing
ImageList.ListImages.Clear
For i = 1 To 11 Step 1
If i = Button.Index Then
ImageList.ListImages.Add , ImageList1.ListImages(i).Key, ImageList1.ListImages(i).Picture
Else
ImageList.ListImages.Add , ImageList2.ListImages(i).Key, ImageList2.ListImages(i).Picture
End If
Next i
Set Toolbar1.ImageList = ImageList
Select Case Button.Caption
Case "Abilities"
If OnTimer = False Then
Abilities.Visible = True
Animations.Visible = False
Armor.Visible = False
Buildings.Visible = False
Images.Visible = False
Shields.Visible = False
Sounds.Visible = False
DataGrid.Visible = False
Races.Visible = False
Upgrades.Visible = False
Weapons.Visible = False
Abilities_EnterCell
lSection = Button.Caption
End If
Label3.Caption = "&Add Ability"
Label16.Caption = "&Remove Ability"
Case "Armor"
If OnTimer = False Then
Abilities.Visible = False
Animations.Visible = False
Armor.Visible = True
Buildings.Visible = False
Images.Visible = False
Shields.Visible = False
Sounds.Visible = False
DataGrid.Visible = False
Races.Visible = False
Upgrades.Visible = False
Weapons.Visible = False
Armor_EnterCell
lSection = Button.Caption
End If
Label3.Caption = "&Add Armor"
Label16.Caption = "&Remove Armor"
Case "Animations"
If OnTimer = False Then
Abilities.Visible = False
Animations.Visible = True
Armor.Visible = False
Buildings.Visible = False
Images.Visible = False
Shields.Visible = False
Sounds.Visible = False
DataGrid.Visible = False
Races.Visible = False
Upgrades.Visible = False
Weapons.Visible = False
Animations_EnterCell
lSection = Button.Caption
End If
Label3.Caption = "&Add Animation"
Label16.Caption = "&Remove Animation"
Case "Buildings"
If OnTimer = False Then
Abilities.Visible = False
Animations.Visible = False
Armor.Visible = False
Buildings.Visible = True
Images.Visible = False
Shields.Visible = False
Sounds.Visible = False
DataGrid.Visible = False
Races.Visible = False
Upgrades.Visible = False
Weapons.Visible = False
Buildings_EnterCell
lSection = Button.Caption
End If
Label3.Caption = "&Add Building"
Label16.Caption = "&Remove Building"
Case "Images"
If OnTimer = False Then
Abilities.Visible = False
Animations.Visible = False
Armor.Visible = False
Buildings.Visible = False
Images.Visible = True
Shields.Visible = False
Sounds.Visible = False
DataGrid.Visible = False
Races.Visible = False
Upgrades.Visible = False
Weapons.Visible = False
Images_EnterCell
lSection = Button.Caption
End If
Label3.Caption = "&Add Image"
Label16.Caption = "&Remove Image"
Case "Infantry"
If OnTimer = False Then
Abilities.Visible = False
Animations.Visible = False
Armor.Visible = False
Buildings.Visible = False
Images.Visible = False
Shields.Visible = False
Sounds.Visible = False
DataGrid.Visible = True
Races.Visible = False
Upgrades.Visible = False
Weapons.Visible = False
DataGrid_EnterCell
lSection = Button.Caption
End If
Label3.Caption = "&Add Infantry"
Label16.Caption = "&Remove Infantry"
Case "Races"
If OnTimer = False Then
Abilities.Visible = False
Animations.Visible = False
Armor.Visible = False
Buildings.Visible = False
Images.Visible = False
Shields.Visible = False
Sounds.Visible = False
DataGrid.Visible = False
Races.Visible = True
Upgrades.Visible = False
Weapons.Visible = False
DataGrid_EnterCell
lSection = Button.Caption
End If
Label3.Caption = "&Add Race"
Label16.Caption = "&Remove Race"
Case "Sheilds"
If OnTimer = False Then
Abilities.Visible = False
Animations.Visible = False
Armor.Visible = False
Buildings.Visible = False
Images.Visible = False
Shields.Visible = True
Sounds.Visible = False
DataGrid.Visible = False
Races.Visible = False
Upgrades.Visible = False
Weapons.Visible = False
Shields_EnterCell
lSection = Button.Caption
End If
Label3.Caption = "&Add Shield"
Label16.Caption = "&Remove Shield"
Case "Sounds"
If OnTimer = False Then
Abilities.Visible = False
Animations.Visible = False
Armor.Visible = False
Buildings.Visible = False
Images.Visible = False
Shields.Visible = False
Sounds.Visible = True
DataGrid.Visible = False
Races.Visible = False
Upgrades.Visible = False
Weapons.Visible = False
Sounds_EnterCell
lSection = Button.Caption
End If
Label3.Caption = "&Add Sound"
Label16.Caption = "&Remove Sound"
Case "Upgrades"
If OnTimer = False Then
Abilities.Visible = False
Animations.Visible = False
Armor.Visible = False
Buildings.Visible = False
Images.Visible = False
Shields.Visible = False
Sounds.Visible = False
DataGrid.Visible = False
Races.Visible = False
Upgrades.Visible = True
Weapons.Visible = False
Upgrades_EnterCell
lSection = Button.Caption
End If
Label3.Caption = "&Add Upgrade"
Label16.Caption = "&Remove Upgrade"
Case "Weapons"
If OnTimer = False Then
Abilities.Visible = False
Animations.Visible = False
Armor.Visible = False
Buildings.Visible = False
Images.Visible = False
Shields.Visible = False
Sounds.Visible = False
DataGrid.Visible = False
Races.Visible = False
Upgrades.Visible = False
Weapons.Visible = True
Weapons_EnterCell
lSection = Button.Caption
End If
Label3.Caption = "&Add Weapon"
Label16.Caption = "&Remove Weapon"
End Select
For i = 1 To 11 Step 1
Toolbar1.Buttons(i).Image = i
Next i
End Sub

Private Sub Toolbar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If HF = True Then Label2_Click
If FF = True Then Label1_Click
If DF = True Then Label10_Click
End Sub

Private Sub Upgrades_Click()
Upgrades_EnterCell
End Sub

Private Sub Upgrades_EnterCell()
DGN = 19
Timer2.Enabled = True
End Sub

Private Sub Upgrades_KeyDown(KeyCode As Integer, Shift As Integer)
subKeyDown KeyCode, Shift
End Sub

Private Sub Upgrades_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If HF = True Then Label2_Click
If FF = True Then Label1_Click
If DF = True Then Label10_Click
End Sub

Private Sub Upgrades_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
DataGridUpdate True, 9
End Sub

Private Sub Upgrades_Scroll()
DataGridUpdate True, 9
End Sub

Private Sub VM_KeyDown(KeyCode As Integer, Shift As Integer)
subKeyDown KeyCode, Shift
End Sub

Private Sub VM_MenuItemClick(MenuNumber As Long, MenuItem As Long)

End Sub

Private Sub VM_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If HF = True Then Label2_Click
If FF = True Then Label1_Click
If DF = True Then Label10_Click
End Sub

Private Sub VM_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If HF = True Then Label2_Click
If FF = True Then Label1_Click
If DF = True Then Label10_Click
End Sub

Private Sub Weapons_Click()
Weapons_EnterCell
End Sub

Private Sub Weapons_EnterCell()
DGN = 20
Timer2.Enabled = True
End Sub

Private Sub Weapons_KeyDown(KeyCode As Integer, Shift As Integer)
subKeyDown KeyCode, Shift
End Sub

Private Sub Weapons_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If HF = True Then Label2_Click
If FF = True Then Label1_Click
If DF = True Then Label10_Click
End Sub

Private Sub Weapons_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
DataGridUpdate True, 10
End Sub

Private Sub Weapons_Scroll()
DataGridUpdate True, 10
End Sub
