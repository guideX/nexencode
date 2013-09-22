VERSION 5.00
Object = "{EE128208-4F73-11D3-83BB-C47C02EE3D01}#1.0#0"; "ControlResizer.ocx"
Begin VB.Form frmEncode 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "NexENCODE - Encode/Decode"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   5970
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEncode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.ComboBox cboFormat 
      Height          =   315
      ItemData        =   "frmEncode.frx":000C
      Left            =   1320
      List            =   "frmEncode.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "< &Remove"
      Height          =   315
      Left            =   2520
      TabIndex        =   5
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add >"
      Height          =   315
      Left            =   2520
      TabIndex        =   4
      Top             =   1800
      Width           =   975
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Change Drive"
      Top             =   120
      Width           =   5775
   End
   Begin VB.ListBox lstQue 
      BackColor       =   &H00FFFFFF&
      Height          =   1035
      Left            =   3600
      TabIndex        =   2
      ToolTipText     =   "Que to Encode/Decode"
      Top             =   1800
      Width           =   2295
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00FFFFFF&
      Height          =   1065
      Left            =   120
      Pattern         =   "*.wav"
      TabIndex        =   1
      ToolTipText     =   "Files in Directory"
      Top             =   1800
      Width           =   2295
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Change Directory"
      Top             =   480
      Width           =   5775
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   6015
      TabIndex        =   8
      Top             =   2880
      Width           =   6015
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   315
         Left            =   4800
         TabIndex        =   11
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   315
         Left            =   3600
         TabIndex        =   10
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdWizard 
         Caption         =   "&Wizard"
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   1095
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
   Begin VB.ComboBox cboDirectory 
      Height          =   315
      ItemData        =   "frmEncode.frx":0036
      Left            =   1200
      List            =   "frmEncode.frx":0046
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3720
      Visible         =   0   'False
      Width           =   3255
   End
   Begin ControlResizer.AutoResizer AutoResizer1 
      Height          =   855
      Left            =   1200
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1508
   End
End
Attribute VB_Name = "frmEncode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboFormat_Click()
'On Local Error Resume Next
If cboFormat.Text = ".Wav to .Mp3" Then
    File1.Pattern = "*.wav"
Else
    File1.Pattern = "*.mp3"
End If
lstQue.Clear
If Err.Number <> 0 Then SetError "cboFormat_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdAdd_Click()
'On Local Error Resume Next
Dim msg As String
msg = Dir1.Path
If Right(Dir1.Path, 1) = "\" Then
    lstQue.AddItem Dir1.Path & File1.List(File1.ListIndex)
Else
    lstQue.AddItem Dir1.Path & "\" & File1.List(File1.ListIndex)
End If
If Err.Number <> 0 Then SetError "cmdAdd_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdCancel_Click()
'On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then SetError "cmdCancel_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdOK_Click()
'On Local Error Resume Next
Dim i As Integer, msg As String, X As Integer, msg2 As String
Dim lPath As String, lWavFile As String, lMp3File As String
Me.Visible = False
If cboFormat.Text = ".Wav to .Mp3" Then
    For i = 1 To lstQue.ListCount
        If Len(lstQue.List(i - 1)) <> 0 Then
            msg = lstQue.List(i - 1)
            msg2 = msg
            lWavFile = GetFileTitle(msg2)
            lMp3File = Left(lWavFile, Len(lWavFile) - 4) & ".mp3"
            lPath = Left(msg, Len(msg) - Len(lWavFile))
            AddEvent Encode, lPath, lWavFile, lEncoderSettings.eOutputDir, lMp3File, 0, ""
        End If
    Next i
ElseIf cboFormat.Text = ".Mp3 to .Wav" Then
    For i = 1 To lstQue.ListCount
        If Len(lstQue.List(i - 1)) <> 0 Then
            msg = lstQue.List(i - 1)
            msg2 = msg
            lMp3File = GetFileTitle(msg2)
            lPath = Left(msg, Len(msg) - Len(lMp3File))
            lWavFile = Left(lMp3File, Len(lMp3File) - 4) & ".wav"
            AddEvent Decode, lPath, lMp3File, App.Path & "\", lWavFile, 0, "AUTODELETE"
        End If
    Next i
End If
ProcessNextEvent
DoEvents
If Err.Number <> 0 Then SetError "cmdOK_Click()", lEvents.eSettings.iErrDescription, Err.Description
Unload Me
End Sub

Private Sub cmdWizard_Click()
'On Local Error Resume Next
Unload Me
frmEncoderWizard.Show
If Err.Number <> 0 Then SetError "cmdWizard_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Command1_Click()
'On Local Error Resume Next
lstQue.RemoveItem lstQue.ListIndex
If Err.Number <> 0 Then SetError "Command1_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Dir1_Change()
'On Local Error Resume Next
File1.Path = Dir1.Path
If Err.Number <> 0 Then SetError "Dir1_Change()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Drive1_Change()
'On Local Error Resume Next
Dir1.Path = Drive1.Drive
If Err.Number <> 0 Then SetError "Drive1_Change()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub File1_DblClick()
'On Local Error Resume Next
cmdAdd_Click
If Err.Number <> 0 Then SetError "File1_DblClick()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Load()
'On Local Error Resume Next
FlashIN frmEncode
cboFormat.ListIndex = 0
Icon = frmMain.Icon
If Err.Number <> 0 Then SetError "Form_Load()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Local Error Resume Next
FlashOut frmEncode
If Err.Number <> 0 Then SetError "Form_Unload()", lEvents.eSettings.iErrDescription, Err.Description
End Sub
