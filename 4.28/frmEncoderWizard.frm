VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{EE128208-4F73-11D3-83BB-C47C02EE3D01}#1.0#0"; "ControlResizer.ocx"
Begin VB.Form frmEncoderWizard 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "NexENCODE - Encoder Wizard"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4140
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picBottom 
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
      Top             =   2760
      Width           =   6015
      Begin VB.CommandButton cmdHelp 
         Caption         =   "Help"
         Height          =   315
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Description of objects"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "&Back"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         ToolTipText     =   "Click to execute"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         Height          =   315
         Left            =   2280
         TabIndex        =   3
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "Close"
         Default         =   -1  'True
         Height          =   315
         Left            =   3240
         TabIndex        =   2
         Top             =   120
         Width           =   855
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         X1              =   0
         X2              =   6000
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Frame fraWizard 
      BorderStyle     =   0  'None
      Height          =   2655
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4095
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmEncoderWizard.frx":0000
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   720
         TabIndex        =   6
         Top             =   120
         Width           =   3255
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   0
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.Frame fraWizard 
      BorderStyle     =   0  'None
      Height          =   2655
      Index           =   2
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   4095
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Select"
         Height          =   375
         Left            =   3120
         TabIndex        =   15
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtFilename 
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   3855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Step 2: Select File:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   13
         Top             =   120
         Width           =   3135
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   2
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.Frame fraWizard 
      BorderStyle     =   0  'None
      Height          =   2655
      Index           =   4
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   4095
      Begin VB.ListBox lstSampleRate 
         Height          =   645
         ItemData        =   "frmEncoderWizard.frx":00AE
         Left            =   120
         List            =   "frmEncoderWizard.frx":00BB
         TabIndex        =   29
         Top             =   1920
         Width           =   1815
      End
      Begin VB.ListBox lstBitrate 
         Height          =   645
         ItemData        =   "frmEncoderWizard.frx":00D4
         Left            =   2040
         List            =   "frmEncoderWizard.frx":0102
         TabIndex        =   28
         Top             =   1920
         Width           =   1935
      End
      Begin VB.ComboBox cboProfile 
         Height          =   315
         ItemData        =   "frmEncoderWizard.frx":0145
         Left            =   120
         List            =   "frmEncoderWizard.frx":0155
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CheckBox chkCreateAlbumFile 
         Appearance      =   0  'Flat
         Caption         =   "Create Album"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2400
         TabIndex        =   26
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CheckBox chkAddTags 
         Appearance      =   0  'Flat
         Caption         =   "Add Tags"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2400
         TabIndex        =   25
         Top             =   840
         Width           =   1455
      End
      Begin VB.CheckBox chkDownsample 
         Appearance      =   0  'Flat
         Caption         =   "Downsample"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2400
         TabIndex        =   24
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblBitrate 
         BackStyle       =   0  'Transparent
         Caption         =   "Bitrate:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2040
         TabIndex        =   32
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Sample Rate:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Current profile:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Step 3: Select Settings:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   23
         Top             =   120
         Width           =   3135
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   4
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.Frame fraWizard 
      BorderStyle     =   0  'None
      Height          =   2655
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   4095
      Begin VB.OptionButton optEncType3 
         Caption         =   "Convert CD Audio files to .mp3 files"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   1440
         Width           =   3255
      End
      Begin VB.OptionButton optEncType2 
         Caption         =   "Convert multiple .wav files to .mp3 files"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   10
         Top             =   1200
         Width           =   3255
      End
      Begin VB.OptionButton optEncType1 
         Caption         =   "Convert single .wav file to .mp3 file"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   9
         Top             =   960
         Value           =   -1  'True
         Width           =   3255
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Step 1: Please select the type of encode you are looking for below ..."
         ForeColor       =   &H00000000&
         Height          =   1335
         Left            =   720
         TabIndex        =   8
         Top             =   120
         Width           =   3135
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   1
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.Frame fraWizard 
      BorderStyle     =   0  'None
      Height          =   2655
      Index           =   3
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   4095
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   255
         Left            =   3240
         TabIndex        =   21
         Top             =   2160
         Width           =   735
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "Delete"
         Height          =   255
         Left            =   2400
         TabIndex        =   20
         Top             =   2160
         Width           =   735
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   255
         Left            =   1560
         TabIndex        =   19
         Top             =   2160
         Width           =   735
      End
      Begin VB.ListBox lstFiles 
         Height          =   1620
         Left            =   720
         TabIndex        =   18
         Top             =   480
         Width           =   3255
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   3
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Step 2: Select Files:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   17
         Top             =   120
         Width           =   3135
      End
   End
   Begin VB.Frame fraWizard 
      BorderStyle     =   0  'None
      Height          =   2655
      Index           =   5
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   4095
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   2280
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label lblFilename 
         BackStyle       =   0  'Transparent
         Caption         =   "Filename:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1680
         Width           =   3975
      End
      Begin VB.Label lblEncProgress 
         BackStyle       =   0  'Transparent
         Caption         =   "Progress: Waiting ..."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1920
         Width           =   3975
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   5
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Step 4: Encoding, please wait ..."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   34
         Top             =   120
         Width           =   3135
      End
   End
   Begin VB.Frame fraWizard 
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   2655
      Index           =   6
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   4095
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "The wizard has completed its task. Click 'Close'"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   720
         TabIndex        =   39
         Top             =   120
         Width           =   3135
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   6
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   480
      End
   End
   Begin ControlResizer.AutoResizer AutoResizer1 
      Height          =   135
      Left            =   1080
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   238
   End
End
Attribute VB_Name = "frmEncoderWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub PreviousWizardFrame()
On Local Error Resume Next
Dim i As Integer
For i = 1 To fraWizard.Count - 1
    fraWizard(i).Visible = False
Next i
lEncWizard.eWizFrame = lEncWizard.eWizFrame - 1
If lEncWizard.eWizFrame = 0 Then cmdBack.Enabled = False
fraWizard(lEncWizard.eWizFrame).Visible = True
cmdNext.Enabled = True
If Err.Number <> 0 Then SetError "PreviousWizardFrame", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub NextEncWizardFrame()
On Local Error Resume Next
Dim i As Integer, mbox As VbMsgBoxResult, msg As String, msg2 As String, msg3 As String
For i = 0 To fraWizard.Count - 1
    fraWizard(i).Visible = False
Next i
Select Case lEncWizard.eWizFrame
Case 0
Case 1
    If optEncType2.Value = True Then
        lEncWizard.eWizFrame = 3
        fraWizard(3).Visible = True
        If fraWizard.Count = lEncWizard.eWizFrame + 1 Then cmdNext.Enabled = False
        cmdBack.Enabled = True
        Exit Sub
    ElseIf optEncType3.Value = True Then
        If lEvents.eSettings.iOverwritePrompts = True Then
            mbox = MsgBox("The Encoder Wizard is strictly for encoding .wav files to .mp3 files. In order to copy CD Audio, you must use the track get dialog. Would you like to switch to this now?", vbYesNo + vbQuestion, "Use Track Get?")
            If mbox = vbYes Then
                Unload Me
                LoadTrackGet True
                Exit Sub
            ElseIf mbox = vbNo Then
                fraWizard(1).Visible = True
                Exit Sub
            End If
        Else
            Unload Me
            LoadTrackGet True
            Exit Sub
        End If
    End If
Case 2
    If Len(txtFilename.Text) = 0 Or DoesFileExist(txtFilename.Text) = False Then
        fraWizard(2).Visible = True
        Beep
        txtFilename.SetFocus
        Exit Sub
    End If
    lEncWizard.eWizFrame = lEncWizard.eWizFrame + 1
    lEncWizard.eType = eSingleWav
Case 3
    If lstFiles.ListCount = 0 Then
        fraWizard(3).Visible = True
        Beep
        lstFiles.SetFocus
        Exit Sub
    End If
    lEncWizard.eType = eMultiWav
Case 4
    If lEncWizard.eType = eMultiWav Then
        lEncWizard.eEnabled = False
        For i = 0 To lstFiles.ListCount
            If Len(lstFiles.List(i)) <> 0 Then
                msg = lstFiles.List(i)
                If DoesFileExist(msg) = True Then
                    msg2 = msg
                    msg2 = GetFileTitle(msg2)
                    msg = Left(msg, Len(msg) - Len(msg2))
                    msg3 = Left(msg2, Len(msg2) - 4) & ".mp3"
                    AddEvent Encode, msg, msg2, msg, msg3, 0, ""
                    lEncWizard.eCount = lEncWizard.eCount + 1
                End If
            End If
        Next i
        lEncWizard.eFinished = True
        Unload Me
        Exit Sub
    ElseIf lEncWizard.eType = eSingleWav Then
        msg = txtFilename.Text
        If DoesFileExist(msg) = True Then
            msg2 = msg
            msg2 = GetFileTitle(msg2)
            msg = Left(msg, Len(msg) - Len(msg2))
            msg3 = Left(msg2, Len(msg2) - 4) & ".mp3"
            AddEvent Encode, msg, msg2, msg, msg3, 0, ""
        End If
    End If
'Case 5
End Select
lEncWizard.eWizFrame = lEncWizard.eWizFrame + 1
fraWizard(lEncWizard.eWizFrame).Visible = True
If fraWizard.Count = lEncWizard.eWizFrame + 1 Then cmdNext.Enabled = False
cmdBack.Enabled = True
If Err.Number <> 0 Then SetError "NextEncWizardFrame", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdAdd_Click()
On Local Error Resume Next
Dim msg As String

msg = OpenDialog(Me, "Wave Audio (*.wav)|*.wav", "Select Wave Audio", CurDir)
If Len(msg) <> 0 Then lstFiles.AddItem msg
If Err.Number <> 0 Then SetError "cmdBack_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdBack_Click()
On Local Error Resume Next
If lEncWizard.eWizFrame = 3 Then lEncWizard.eWizFrame = 2
If lEncWizard.eWizFrame = 4 And lEncWizard.eType = eSingleWav Then lEncWizard.eWizFrame = 3
PreviousWizardFrame
If Err.Number <> 0 Then SetError "cmdBack_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdClear_Click()
On Local Error Resume Next
lstFiles.Clear
If Err.Number <> 0 Then SetError "cmdClear_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdClose_Click()
On Local Error Resume Next
Dim mbox As VbMsgBoxResult

If lEncWizard.eFinished = False And lEvents.eSettings.iOverwritePrompts = True Then
    mbox = MsgBox("Are you sure you wish to close the wizard before finishing?", vbYesNo + vbQuestion)
    If mbox = vbYes Then
        PlayWav App.Path & "\media\done.wav", SND_ASYNC
        Unload Me
        Exit Sub
    Else
        Exit Sub
    End If
ElseIf lEncWizard.eFinished = True Then
    PlayWav App.Path & "\media\done.wav", SND_ASYNC
    Unload Me
    Exit Sub
ElseIf lEvents.eSettings.iOverwritePrompts = False Then
    PlayWav App.Path & "\media\done.wav", SND_ASYNC
    Unload Me
    Exit Sub
End If
If Err.Number <> 0 Then SetError "Form_Load", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdDel_Click()
On Local Error Resume Next
lstFiles.RemoveItem lstFiles.ListIndex
If Err.Number <> 0 Then SetError "cmdNext_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdNext_Click()
On Local Error Resume Next
NextEncWizardFrame
If Err.Number <> 0 Then SetError "cmdNext_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdSelect_Click()
On Local Error Resume Next
txtFilename.Text = OpenDialog(Me, "Wave Audio (*.wav)|*.wav", "Select Wave Audio", CurDir)
If Err.Number <> 0 Then SetError "cmdSelect_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Load()
On Local Error Resume Next
Dim i As Integer
FlashIN Me
If lEncoderSettings.eDownsample = True Then chkDownsample.Value = 1
cboProfile.ListIndex = lEncoderSettings.eProfile
lstBitrate.Text = lEncoderSettings.eBitrate
lstSampleRate.Text = lEncoderSettings.eSampleRate
lEncWizard.eEnabled = True

For i = 0 To imgIcon.Count - 1
    imgIcon(i).Picture = frmGraphics.imgIcon.Picture
    fraWizard(i).Visible = False
Next i
fraWizard(0).Visible = True
If Err.Number <> 0 Then SetError "Form_Load", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Local Error Resume Next
lEncWizard.eCount = 0
lEncWizard.eEnabled = False
lEncWizard.eFinished = False
lEncWizard.eWizFrame = 0
FlashOut Me
If Err.Number <> 0 Then SetError "Form_Unload", lEvents.eSettings.iErrDescription, Err.Description
End Sub
