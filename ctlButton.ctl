VERSION 5.00
Begin VB.UserControl ctlButton 
   ClientHeight    =   3315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2445
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   221
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   163
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   600
      Top             =   1440
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   2175
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000014&
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   2295
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000010&
      Height          =   1335
      Left            =   0
      Top             =   0
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   600
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "ctlButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private ButtonUp As Boolean
Public Event Click()

Public Property Get ButtonIsUp() As Boolean
ButtonIsUp = ButtonUp
End Property

Public Property Let ButtonIsUp(Data As Boolean)
ButtonUp = Data
PropertyChanged "ButtonIsUp"
UserControl_Resize
End Property

Public Property Get ClearPicture() As Boolean
ClearPicture = False
End Property

Public Property Let ClearPicture(Data As Boolean)
ClearPicture = False
End Property

Public Property Get Caption() As String
Caption = Label1.Caption
End Property

Public Property Let Caption(Data As String)
Label1.Caption = Data
PropertyChanged "Caption"
UserControl_Resize
End Property

Public Property Get bPicture() As IPictureDisp
Set bPicture = Image1.Picture
End Property

Public Property Set bPicture(Data As IPictureDisp)
Set Image1.Picture = Data
PropertyChanged "bPicture"
UserControl_Resize
End Property

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
UserControl_MouseUp Button, Shift, (X / 15) + Image1.Left, (Y / 15) + Image1.Top
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
UserControl_MouseUp Button, Shift, X / 15, (Y / 15) + Label1.Top
End Sub

Private Sub Timer1_Timer()
If ButtonUp = True And Shape1.BorderColor <> &H80000008 Then
UserControl_Resize
End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X >= 0 And Y >= 0 And X <= UserControl.ScaleWidth And Y <= UserControl.ScaleHeight Then
ButtonUp = Not ButtonUp
UserControl_Resize
RaiseEvent Click
End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Label1.Caption = PropBag.ReadProperty("LC", "")
ButtonUp = PropBag.ReadProperty("BU", True)
Set Image1.Picture = PropBag.ReadProperty("bP", Nothing)
End Sub

Public Sub ReDrawButton()
UserControl_Resize
End Sub

Private Sub UserControl_Resize()
Label1.Caption = Caption
If ButtonUp = True Then
Shape1.Left = -1
Shape1.Top = -1
Shape1.Width = UserControl.ScaleWidth + 1
Shape1.Height = UserControl.ScaleHeight + 1
Shape1.BorderColor = &H80000008
Shape2.Width = UserControl.ScaleWidth + 1
Shape2.Height = UserControl.ScaleHeight + 1
Shape2.BorderColor = &H80000014 '&HC0C0C0
Shape3.Left = -1
Shape3.Top = -1
Shape3.Width = Shape1.Width - 1
Shape3.Height = Shape1.Height - 1
Label1.Top = (UserControl.ScaleHeight / 2) - (Label1.Height / 2) + 1
Label1.Width = UserControl.ScaleWidth
Image1.Left = (UserControl.ScaleWidth / 2) - (Image1.Width / 2)
Image1.Top = (UserControl.ScaleHeight / 2) - (Image1.Height / 2)
Else
Shape1.Width = UserControl.ScaleWidth
Shape1.Height = UserControl.ScaleHeight
Shape1.BorderColor = &H80000014 '&HC0C0C0
Shape2.Width = UserControl.ScaleWidth + 1
Shape2.Height = UserControl.ScaleHeight + 1
Shape2.BorderColor = &H80000008
Shape3.Left = 1
Shape3.Top = 1
Shape3.Width = Shape1.Width
Shape3.Height = Shape1.Height
Label1.Top = (UserControl.ScaleHeight / 2) - (Label1.Height / 2) + 1
Label1.Width = UserControl.ScaleWidth
Image1.Left = (UserControl.ScaleWidth / 2) - (Image1.Width / 2)
Image1.Top = (UserControl.ScaleHeight / 2) - (Image1.Height / 2)
End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "LC", Label1.Caption
PropBag.WriteProperty "BU", ButtonUp
PropBag.WriteProperty "bP", Image1.Picture
End Sub
