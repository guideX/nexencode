VERSION 5.00
Object = "{EE128208-4F73-11D3-83BB-C47C02EE3D01}#1.0#0"; "ControlResizer.ocx"
Begin VB.Form frmMP3Info 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "NexENCODE - Tag Info"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   5985
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMP3-Info.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Wipe Tag"
      Height          =   375
      Left            =   4800
      TabIndex        =   24
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load"
      Height          =   375
      Left            =   2040
      TabIndex        =   23
      ToolTipText     =   "Load an Mpeg Layer 3 File"
      Top             =   120
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   6015
      TabIndex        =   0
      Top             =   2640
      Width           =   6015
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Default         =   -1  'True
         Height          =   315
         Left            =   3840
         TabIndex        =   26
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   315
         Left            =   4920
         TabIndex        =   25
         Top             =   120
         Width           =   975
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         X1              =   6120
         X2              =   0
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Frame fraID3 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1935
      Left            =   2040
      TabIndex        =   10
      Top             =   600
      Width           =   3975
      Begin VB.ComboBox cmbGenre 
         Height          =   315
         ItemData        =   "frmMP3-Info.frx":000C
         Left            =   2160
         List            =   "frmMP3-Info.frx":028B
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtComment 
         Height          =   285
         Left            =   900
         MaxLength       =   30
         TabIndex        =   15
         Top             =   1560
         Width           =   3015
      End
      Begin VB.TextBox txtTitle 
         Height          =   285
         Left            =   900
         MaxLength       =   30
         TabIndex        =   14
         Top             =   480
         Width           =   3015
      End
      Begin VB.TextBox txtArtist 
         Height          =   285
         Left            =   900
         MaxLength       =   30
         TabIndex        =   13
         Top             =   120
         Width           =   3015
      End
      Begin VB.TextBox txtAlbum 
         Height          =   285
         Left            =   900
         MaxLength       =   30
         TabIndex        =   12
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox txtYear 
         Height          =   285
         Left            =   900
         MaxLength       =   4
         TabIndex        =   11
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lblDescribe 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Year:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   0
         TabIndex        =   22
         Top             =   1245
         Width           =   390
      End
      Begin VB.Label lblDescribe 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Comment:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   0
         TabIndex        =   21
         Top             =   1605
         Width           =   735
      End
      Begin VB.Label lblDescribe 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Title:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   20
         Top             =   480
         Width           =   360
      End
      Begin VB.Label lblDescribe 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Album:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   0
         TabIndex        =   19
         Top             =   900
         Width           =   495
      End
      Begin VB.Label lblDescribe 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Artist:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   18
         Top             =   180
         Width           =   450
      End
      Begin VB.Label lblDescribe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Genre"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   1560
         TabIndex        =   17
         Top             =   1260
         Width           =   435
      End
   End
   Begin ControlResizer.AutoResizer AutoResizer1 
      Height          =   495
      Left            =   360
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin VB.Label lblSize 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   0
      TabIndex        =   9
      Top             =   120
      Width           =   1950
   End
   Begin VB.Label lblLength 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   0
      TabIndex        =   8
      Top             =   300
      Width           =   1950
   End
   Begin VB.Label lblFreqChan 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   0
      TabIndex        =   7
      Top             =   900
      Width           =   1950
   End
   Begin VB.Label lblCRC 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   0
      TabIndex        =   6
      Top             =   1095
      Width           =   1950
   End
   Begin VB.Label lblLayer 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   0
      TabIndex        =   5
      Top             =   495
      Width           =   1950
   End
   Begin VB.Label lblBitRate 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   0
      TabIndex        =   4
      Top             =   705
      Width           =   1950
   End
   Begin VB.Label lblCopyRight 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   1305
      Width           =   1950
   End
   Begin VB.Label lblOriginal 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   1500
      Width           =   1950
   End
   Begin VB.Label lblEmphasis 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   1680
      Width           =   1950
   End
End
Attribute VB_Name = "frmMP3Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strArtist As String * 30
Dim strTitle As String * 30
Dim strAlbum As String * 30
Dim strYear As String * 4
Dim strComment As String * 30
Dim strGenre As String * 1
Dim TempString As String

Public Sub RefreshTagInfo()
On Local Error Resume Next
If Len(lTag.tFile) = 0 Then
    PromptGetTag
    If Len(lTag.tFile) = 0 Then Exit Sub
End If
PromptGetTag lTag.tFile
If lTag.tHasTag = True Then
    lblLayer.Caption = "Layer: " & lTag.tLayer
    lblSize.Caption = "Size: " & lTag.tSize
    lblLength.Caption = "Length: " & lTag.tLength
    lblBitrate.Caption = "Bitrate: " & lTag.tBitrate
    lblFreqChan.Caption = "Chan: " & lTag.tFreqChan
    lblCRC.Caption = "CRC: " & lTag.tCRC
    lblCopyRight.Caption = "Copyright: " & lTag.tCopyright
    lblOriginal.Caption = "Origional: " & lTag.tOrigional
    lblEmphasis.Caption = "Emphasis: " & lTag.tEmphisis
    txtArtist.Text = lTag.tArtist
    txtTitle.Text = lTag.tTitle
    txtAlbum.Text = lTag.tAlbum
    txtYear.Text = lTag.tYear
    txtComment.Text = lTag.tComment
    cmdRemove.Visible = True
    cmbGenre.ListIndex = lTag.tGenre
Else
    lblSize.Caption = "Size: " & lTag.tSize
    lblLength.Caption = "Length: " & lTag.tLength
    lblBitrate.Caption = "Bitrate: " & lTag.tBitrate
    lblFreqChan.Caption = "FreqChan: " & lTag.tFreqChan
    lblCRC.Caption = "CRC: " & lTag.tCRC
    lblCopyRight.Caption = "Copyright: " & lTag.tCopyright
    lblOriginal.Caption = "Origional: " & lTag.tOrigional
    lblEmphasis.Caption = "Emphasis: " & lTag.tEmphisis
    cmdRemove.Visible = False
End If
'Caption = lTag.tFile
'Caption = GetFileTitle(Caption)
If Err.Number <> 0 Then SetError "RefreshTagInfo()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdRemove_Click()
On Local Error Resume Next
If Len(lTag.tFile) = 0 Then Exit Sub
MP3FileSize = FileLen(lTag.tFile)
TempString = Space(MP3FileSize - 128)
Open lTag.tFile For Binary As #1
    Get #1, 1, TempString
Close #1
Kill lTag.tFile
Open lTag.tFile For Binary As #1
    Put #1, 1, TempString
Close #1
GetTagInfo
If Err.Number <> 0 Then SetError "cmdRemove_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdSave_Click()
On Local Error Resume Next
If Len(lTag.tFile) <> 0 Then
    lTag.tSize = FileLen(lTag.tFile)
    If cmbGenre.ListIndex = -1 Then
        strGenre = Chr(255)
    Else
        strGenre = Chr(cmbGenre.ItemData(cmbGenre.ListIndex))
    End If
    With lTag
        .tTitle = txtTitle.Text
        .tArtist = txtArtist.Text
        .tAlbum = txtAlbum.Text
        .tYear = txtYear.Text
        .tComment = txtComment
        .tGenre = strGenre
    End With
    SaveTagInfo lTag.tFile
End If
Unload Me
If Err.Number <> 0 Then SetError "cmdSave_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Command1_Click()
On Local Error Resume Next
Unload Me
End Sub

Private Sub Command2_Click()
On Local Error Resume Next
PromptGetTag
DoEvents
RefreshTagInfo
If Err.Number <> 0 Then SetError "Command2()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Load()
On Local Error Resume Next
Dim msg As String
'Image1.Picture = frmGraphics.imgTopper.Picture
'Image2.Picture = frmGraphics.imgSideGradient.Picture
If lEvents.eEncoderBusy = True Or lEvents.eRipperBusy = True Then
    Exit Sub
    Me.Visible = False
End If
Me.Icon = frmMain.Icon
lTag.tFile = ""
'FlashIN Me
If Err.Number <> 0 Then SetError "Form_Load", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Local Error Resume Next
FlashOut frmMP3Info
End Sub

Private Sub Image1_DblClick()
If Me.WindowState = vbMaximized Then
    Me.WindowState = vbNormal
Else
    Me.WindowState = vbMaximized
End If

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub
