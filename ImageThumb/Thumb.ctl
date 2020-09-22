VERSION 5.00
Begin VB.UserControl ThumbImg 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   2865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2985
   ScaleHeight     =   2865
   ScaleWidth      =   2985
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFFFF&
      X1              =   1080
      X2              =   120
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   1080
      X2              =   1080
      Y1              =   960
      Y2              =   120
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   120
      Y1              =   960
      Y2              =   120
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00808080&
      X1              =   1080
      X2              =   120
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Text"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   1230
      Width           =   1215
   End
   Begin VB.Line Line4 
      X1              =   1200
      X2              =   1200
      Y1              =   0
      Y2              =   1200
   End
   Begin VB.Line Line3 
      X1              =   1200
      X2              =   0
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   1200
      X2              =   0
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1200
   End
End
Attribute VB_Name = "ThumbImg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'               Thumb Activex Version 1.00
'                 Written by: Behrouz Rad
'                Copyright: Decamber 2003
'    >>> Islamic Azad University Division Of Mahshahr <<<
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

Event Click()
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim EnterF As Long
Dim ExitF As Long

Private Sub Image1_Click()
 Call UserControl_Click
End Sub

Private Sub Image1_DblClick()
 Call UserControl_DblClick
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call UserControl_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call UserControl_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call UserControl_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Label1_Click()
 Call UserControl_Click
End Sub

Private Sub Label1_DblClick()
 Call UserControl_DblClick
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call UserControl_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call UserControl_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call UserControl_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Click()
 RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
 RaiseEvent DblClick
End Sub

Private Sub UserControl_EnterFocus()
 UserControl.BackColor = EnterF
End Sub

Private Sub UserControl_ExitFocus()
 UserControl.BackColor = ExitF
End Sub

Private Sub UserControl_Initialize()
 EnterF = vbBlue
 ExitF = &H8000000F
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
 Label1.BackColor = PropBag.ReadProperty("CaptionBkColor", &HFFFFC0)
 Label1.ForeColor = PropBag.ReadProperty("CaptionFrColor", vbBlack)
 Set Picture = PropBag.ReadProperty("Picture", Nothing)
 Label1.Caption = PropBag.ReadProperty("ImgCaption", "")
 EnterF = PropBag.ReadProperty("EnterFocusColor", EnterF)
 ExitF = PropBag.ReadProperty("ExitFocusColor", ExitF)
End Sub

Private Sub UserControl_Resize()
 If Height <> 1480 Then Height = 1480
 If Width <> 1210 Then Width = 1210
End Sub

Public Property Get Picture() As Picture
 Set Picture = Image1.Picture
End Property

Public Property Set Picture(ByVal New_ButtonIcon As Picture)
  Set Image1.Picture = New_ButtonIcon
PropertyChanged "Picture"
End Property

Public Property Get ImgCaption() As String
 ImgCaption = Label1.Caption
End Property

Public Property Let ImgCaption(ByVal New_ImgCaption As String)
 Label1.Caption = New_ImgCaption
PropertyChanged "ImgCaption"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get EnterFocusColor() As OLE_COLOR
    EnterFocusColor = EnterF
End Property

Public Property Let EnterFocusColor(ByVal New_EnterFocusColor As OLE_COLOR)
    EnterF = New_EnterFocusColor
    PropertyChanged "EnterFocusColor"
End Property

Public Property Get ExitFocusColor() As OLE_COLOR
    ExitFocusColor = ExitF
End Property

Public Property Let ExitFocusColor(ByVal New_ExitFocusColor As OLE_COLOR)
    ExitF = New_ExitFocusColor
    PropertyChanged "ExitFocusColor"
End Property
Public Property Get CaptionBkColor() As OLE_COLOR
    CaptionBkColor = Label1.BackColor
End Property
Public Property Let CaptionBkColor(ByVal New_CaptionBkColor As OLE_COLOR)
 Label1.BackColor = New_CaptionBkColor
 PropertyChanged "CaptionBkColor"
End Property

Public Property Get CaptionFrColor() As OLE_COLOR
 CaptionFrColor = Label1.ForeColor
End Property

Public Property Let CaptionFrColor(ByVal New_CaptionFrColor As OLE_COLOR)
 Label1.ForeColor = New_CaptionFrColor
 PropertyChanged "CaptionFrColor"
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
 Call PropBag.WriteProperty("Picture", Me.Picture, Nothing)
 Call PropBag.WriteProperty("ImgCaption", Label1.Caption, "")
 Call PropBag.WriteProperty("EnterFocusColor", EnterF, vbBlue)
 Call PropBag.WriteProperty("ExitFocusColor", ExitF, &HC0C0C0)
 Call PropBag.WriteProperty("CaptionBkColor", Label1.BackColor, &HFFFFC0)
 Call PropBag.WriteProperty("CaptionFrColor", Label1.ForeColor, vbBlack)
End Sub

Public Sub Refresh()
'UserControl.Refresh
Image1.Refresh
End Sub
