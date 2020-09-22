VERSION 5.00
Object = "*\AThumb.vbp"
Begin VB.Form FrmTest 
   Caption         =   "Thumb Activex Version 1.00 - Written by: Behrouz Rad"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7185
   Icon            =   "FrmThumb.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   7185
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin Thumb.ThumbImg Thumb 
      Height          =   1485
      Index           =   0
      Left            =   3360
      TabIndex        =   6
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2619
      BackColor       =   12632256
      ImgCaption      =   "Text"
      ExitFocusColor  =   -2147483633
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Your Directory:"
      Height          =   5055
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3135
      Begin VB.DriveListBox Drive1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   2655
      End
      Begin VB.DirListBox Dir1 
         Appearance      =   0  'Flat
         Height          =   1890
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   2655
      End
      Begin VB.FileListBox File1 
         Appearance      =   0  'Flat
         Height          =   2175
         Left            =   240
         Pattern         =   "*.bmp;*.jpg;*.gif;*.dib;*.wmf;*.ico;*.cur"
         TabIndex        =   2
         Top             =   2640
         Width           =   2655
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Picture # of #"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   5760
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Founded:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   5400
      Width           =   810
   End
End
Attribute VB_Name = "FrmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'              Thumbital View Version 1.00
'                Written by: Behrouz Rad
'
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

Private Sub Command1_Click()
Dim S, J As String
Dim I, K, T As Integer
'If a File Has Only a Supported Extention
'But it is not a Picture an Error Will Occure.
On Error Resume Next
K = 3360 'Set Default Left
T = 240  'Set Default Top
I = 0    'Counter
Label2.Visible = True
File1.ListIndex = 0
Do While File1.ListIndex <= File1.ListCount - 1
If Len(Dir1.Path) = 3 Then
   S = File1.Path & File1.FileName
Else
   S = File1.Path & "\" & File1.FileName
End If
  J = File1.FileName
  
  'If Extention Of Your File Is Not a Supported Format an
  'Error Will Occure.
  'I Write a Comment For Check The Extentions.
  If Right$(S, 3) = LCase$("jpg") Or Right$(S, 3) = LCase$("bmp") _
  Or Right$(S, 3) = LCase$("gif") Or Right$(S, 3) = LCase$("dib") _
  Or Right$(S, 3) = LCase$("wmf") Or Right$(S, 3) = LCase$("emf") _
  Or Right$(S, 3) = LCase$("ico") Or Right$(S, 3) = LCase$("cur") _
  Or Right$(S, 4) = LCase$("jpeg") Or Right$(S, 3) = UCase$("jpg") _
  Or Right$(S, 3) = UCase$("bmp") Or Right$(S, 3) = UCase$("gif") _
  Or Right$(S, 3) = UCase$("dib") Or Right$(S, 3) = UCase$("wmf") _
  Or Right$(S, 3) = UCase$("emf") Or Right$(S, 3) = UCase$("ico") _
  Or Right$(S, 3) = UCase$("cur") Or Right$(S, 4) = UCase$("jpeg") Then
  Thumb(0).Visible = False 'Just Because i Don't Wanna See This
  I = I + 1
  Load Thumb(I)
  Thumb(I).Left = K
  Thumb(I).Top = T
  Thumb(I).Visible = True
  K = K + 1270
  If Thumb(I).Left > 10000 Then
     K = 3360
     T = T + 1550
  End If
  
  Set Thumb(I).Picture = LoadPicture(S)
      Thumb(I).ImgCaption = J
      Thumb(I).Refresh 'IMPORTANT For Repainting
  End If
  
  If File1.ListIndex = File1.ListCount - 1 Then Exit Do
  File1.ListIndex = I
  Label2.Caption = "Picture " & CStr(I + 1) & " of " & File1.ListCount
  Label2.Refresh
Loop
End Sub

Private Sub Dir1_Change()
 File1.Path = Dir1.Path
 Label1.Caption = "Founded: " & File1.ListCount
End Sub

Private Sub Drive1_Change()
On Error GoTo ErrNotReady
 Dir1.Path = Drive1.Drive
 Exit Sub
ErrNotReady:
    MsgBox UCase(Drive1.Drive) & "\" & " is not accessible." & _
    vbCr & vbCr & "the device is not ready.", vbCritical, "Image Viewer"
    Drive1.Drive = "C:"
    Dir1.Path = "C:\"
End Sub

Private Sub Form_Load()
 Drive1.Drive = "C:"
 Dir1.Path = "C:"
End Sub

Private Sub Thumb_Click(Index As Integer)
 SelImg = Thumb(Index).Index
 FrmPic.Show
'MsgBox "You Clicked Item: " & Thumb(Index).Index, vbInformation, "Thumb Activex"
End Sub
