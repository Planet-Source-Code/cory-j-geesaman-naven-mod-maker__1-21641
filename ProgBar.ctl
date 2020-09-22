VERSION 5.00
Begin VB.UserControl ctlProgBar 
   Alignable       =   -1  'True
   BackColor       =   &H00929A93&
   BackStyle       =   0  'Transparent
   ClientHeight    =   1290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2445
   ScaleHeight     =   86
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   163
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   1815
      TabIndex        =   0
      Top             =   0
      Width           =   1815
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   0
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   81
         TabIndex        =   1
         Top             =   0
         Width           =   1215
         Begin VB.Label lPercent 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   45
         End
      End
   End
End
Attribute VB_Name = "ctlProgBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private vPercent As Integer

Public Property Get BGColor() As OLE_COLOR
BGColor = UserControl.BackColor
End Property

Public Property Let BGColor(Data As OLE_COLOR)
UserControl.BackColor = Data
PropertyChanged "BGColor"
ReDrawBar
End Property

Public Property Get FColor1() As OLE_COLOR
FColor1 = UserControl.ForeColor
End Property

Public Property Let FColor1(Data As OLE_COLOR)
UserControl.ForeColor = Data
PropertyChanged "FColor1"
ReDrawBar
End Property

Public Property Get FColor2() As OLE_COLOR
FColor2 = UserControl.FillColor
End Property

Public Property Let FColor2(Data As OLE_COLOR)
UserControl.FillColor = Data
PropertyChanged "FColor2"
ReDrawBar
End Property

Public Property Get Percent() As Integer
Percent = vPercent
End Property

Public Property Let Percent(Data As Integer)
vPercent = Data
PropertyChanged "Percent"
ReDrawBar
End Property

Public Property Get LabelVisible() As Boolean
LabelVisible = lPercent.Visible
End Property

Public Property Let LabelVisible(Data As Boolean)
lPercent.Visible = Data
PropertyChanged "LabelVisible"
ReDrawBar
End Property

Public Property Get LabelColor() As OLE_COLOR
LabelColor = lPercent.ForeColor
End Property

Public Property Let LabelColor(Data As OLE_COLOR)
lPercent.ForeColor = Data
PropertyChanged "LabelColor"
ReDrawBar
End Property

Public Property Get LabelBColor() As OLE_COLOR
LabelBColor = lPercent.BackColor
End Property

Public Property Let LabelBColor(Data As OLE_COLOR)
lPercent.BackColor = Data
PropertyChanged "LabelBColor"
ReDrawBar
End Property

Public Property Get LabelBackStyle() As fmBackStyle
LabelBackStyle = lPercent.BackStyle
End Property

Public Property Let LabelBackStyle(Data As fmBackStyle)
lPercent.BackStyle = Data
PropertyChanged "LabelBackStyle"
ReDrawBar
End Property

Public Sub ReDrawBar()
Picture1.Height = UserControl.ScaleHeight
Picture1.Left = 0
Picture2.Width = UserControl.ScaleWidth
Picture2.Top = 0
Picture2.Height = Picture1.ScaleHeight
Picture2.ScaleHeight = 100
Picture1.Width = vPercent
If vPercent > 0 Then
Picture1.ScaleWidth = vPercent
Picture1.Visible = True
Else
Picture1.ScaleWidth = vPercent + 1
Picture1.Visible = False
End If
Picture2.Width = 100
Dim i As Integer
Picture2.ScaleMode = 3
Picture2.DrawWidth = (Picture2.ScaleWidth \ 100) + 1
Picture2.ScaleMode = 0
Picture2.ScaleHeight = 100
Picture2.ScaleWidth = 100
Picture2.Cls
For i = 0 To Picture2.ScaleWidth Step 1
Picture2.Line (i * (Picture2.ScaleWidth / 100), 0)-(i * (Picture2.ScaleWidth / 100), Picture2.ScaleHeight), Blend(FColor1, FColor2, i)
Next i
UserControl.ScaleWidth = 100
lPercent.Caption = vPercent & "%"
lPercent.Left = (vPercent / 2) - (lPercent.Width / 2)
lPercent.Top = (Picture2.ScaleHeight / 2) - (lPercent.Height / 2)
End Sub

Public Function RGBRed(RGBCol As Long) As Integer
    RGBRed = RGBCol And &HFF
End Function

Public Function RGBGreen(RGBCol As Long) As Integer
    RGBGreen = ((RGBCol And &H100FF00) / &H100)
End Function

Public Function RGBBlue(RGBCol As Long) As Integer
    RGBBlue = (RGBCol And &HFF0000) / &H10000
End Function

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
UserControl.BackColor = PropBag.ReadProperty("BGC", 0)
UserControl.ForeColor = PropBag.ReadProperty("C1", 0)
UserControl.FillColor = PropBag.ReadProperty("C2", 0)
vPercent = PropBag.ReadProperty("vP", 0)
lPercent.ForeColor = PropBag.ReadProperty("lfC", 0)
lPercent.BackColor = PropBag.ReadProperty("lbC", 0)
lPercent.Visible = PropBag.ReadProperty("lV", True)
lPercent.BackStyle = PropBag.ReadProperty("lS", 0)
End Sub

Private Sub UserControl_Resize()
ReDrawBar
End Sub

Public Function Blend(Color1 As OLE_COLOR, Color2 As OLE_COLOR, Number As Integer) As OLE_COLOR
r = ((RGBRed(Color1) * (100 - Number)) + (RGBRed(Color2) * (Number))) / 100
g = ((RGBGreen(Color1) * (100 - Number)) + (RGBGreen(Color2) * (Number))) / 100
b = ((RGBBlue(Color1) * (100 - Number)) + (RGBBlue(Color2) * (Number))) / 100
Blend = RGB(r, g, b)
End Function

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "BGC", UserControl.BackColor
PropBag.WriteProperty "C1", UserControl.ForeColor
PropBag.WriteProperty "C2", UserControl.FillColor
PropBag.WriteProperty "vP", vPercent
PropBag.WriteProperty "lfC", lPercent.ForeColor
PropBag.WriteProperty "lbC", lPercent.BackColor
PropBag.WriteProperty "lV", lPercent.Visible
PropBag.WriteProperty "lS", lPercent.BackStyle
End Sub
