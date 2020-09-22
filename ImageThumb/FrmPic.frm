VERSION 5.00
Begin VB.Form FrmPic 
   BackColor       =   &H00404040&
   Caption         =   "Image Viewer"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FrmPic.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdPer 
      Caption         =   "< Perview"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton CmdNext 
      Caption         =   "Next >"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   4680
      Top             =   3000
      Width           =   1575
   End
End
Attribute VB_Name = "FrmPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CenterImage(FormX As Form, ImgName As Image)
 ImgName.Top = (FormX.Height - ImgName.Height) / 2
 ImgName.Left = (FormX.Width - ImgName.Width) / 2
End Sub

Private Sub CmdNext_Click()
On Error GoTo ErrEnd
 SelImg = SelImg + 1
 Image1.Picture = FrmTest.Thumb(SelImg).Picture
 CenterImage FrmPic, Image1
 Me.Caption = "Image Viewer >>>   " & FrmTest.Thumb(SelImg).ImgCaption
 Exit Sub
ErrEnd:
   MsgBox "Last Image is Available", vbExclamation, "Image Viewer"
End Sub

Private Sub CmdPer_Click()
On Error GoTo ErrEnd
 SelImg = SelImg - 1
 Image1.Picture = FrmTest.Thumb(SelImg).Picture
 CenterImage FrmPic, Image1
 Me.Caption = "Image Viewer >>>   " & FrmTest.Thumb(SelImg).ImgCaption
 Exit Sub
ErrEnd:
   MsgBox "First Image is Available", vbExclamation, "Image Viewer"
End Sub

Private Sub Form_Activate()
 CenterImage FrmPic, Image1
End Sub

Private Sub Form_Click()
 Unload Me
End Sub

Private Sub Form_Load()
 Image1.Picture = FrmTest.Thumb(SelImg).Picture
 Me.Caption = "Image Viewer >>>   " & FrmTest.Thumb(SelImg).ImgCaption
End Sub

Private Sub Image1_Click()
 Unload Me
End Sub
