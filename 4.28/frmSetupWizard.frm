VERSION 5.00
Object = "{9F5F61C6-83A0-11D2-A800-00A0CC20D781}#1.0#0"; "ACD.OCX"
Object = "{EE128208-4F73-11D3-83BB-C47C02EE3D01}#1.0#0"; "ControlResizer.ocx"
Begin VB.Form frmSetupWizard 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "NexENCODE - Setup Wizard"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4125
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSetupWizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Top             =   3240
      Width           =   6015
      Begin VB.CommandButton cmdBack 
         Caption         =   "<< Back"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next >>"
         Default         =   -1  'True
         Height          =   315
         Left            =   2280
         TabIndex        =   4
         Top             =   120
         Width           =   855
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   0
         ScaleHeight     =   615
         ScaleWidth      =   6015
         TabIndex        =   3
         Top             =   3405
         Width           =   6015
      End
      Begin VB.CommandButton cmdFinish 
         Caption         =   "&Finish"
         Enabled         =   0   'False
         Height          =   315
         Left            =   3240
         TabIndex        =   2
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   855
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
   Begin ACDLib.ACD ASPIChecker 
      Height          =   495
      Left            =   2400
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   873
      _StockProps     =   0
   End
   Begin VB.Frame fraSetup 
      BorderStyle     =   0  'None
      Caption         =   "Setup 1"
      Height          =   3255
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   3975
      Begin VB.CommandButton cmdHelp0 
         Caption         =   "Help"
         Height          =   375
         Left            =   3000
         TabIndex        =   21
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Setup - Welcome"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   18
         Top             =   240
         Width           =   2895
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   0
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmSetupWizard.frx":000C
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   3735
      End
   End
   Begin VB.Frame fraSetup 
      BorderStyle     =   0  'None
      Caption         =   "Setup 1"
      Height          =   3255
      Index           =   2
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   3975
      Begin VB.CommandButton Command2 
         Caption         =   "Help"
         Height          =   375
         Left            =   3000
         TabIndex        =   27
         Top             =   2760
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Do not search"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2760
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Add single folder"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2400
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Search Hard Drive(s)"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   2040
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   2
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Setup - Search"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   19
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmSetupWizard.frx":009D
         ForeColor       =   &H00000000&
         Height          =   1695
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   3735
      End
   End
   Begin VB.Frame fraSetup 
      BorderStyle     =   0  'None
      Caption         =   "Setup Complete"
      Height          =   3255
      Index           =   3
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   3975
      Begin VB.CommandButton Command3 
         Caption         =   "Help"
         Height          =   375
         Left            =   3000
         TabIndex        =   28
         Top             =   2760
         Width           =   855
      End
      Begin VB.CheckBox chkRandomMP3 
         Appearance      =   0  'Flat
         Caption         =   "Play an MP3 when I click 'Finish'"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2880
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Click the finish button to complete setup"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   3735
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   3
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Setup - Finish"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   8
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame fraSetup 
      BorderStyle     =   0  'None
      Caption         =   "Setup 1"
      Height          =   3255
      Index           =   1
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   3975
      Begin VB.CommandButton Command1 
         Caption         =   "Help"
         Height          =   375
         Left            =   3000
         TabIndex        =   16
         Top             =   2760
         Width           =   855
      End
      Begin VB.OptionButton optAspi3 
         Caption         =   "My ASPI drivers are out of date, however I do not want them installed (Copy CD Audio Disabled)"
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   15
         Top             =   1800
         Width           =   3735
      End
      Begin VB.OptionButton optAspi2 
         Caption         =   "My ASPI drivers either do not exist, or are out of date. Update them for me."
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   14
         Top             =   1200
         Width           =   3735
      End
      Begin VB.OptionButton optAspi1 
         Caption         =   "My ASPI drivers are up to date, copying CD Audio in NexENCODE is enabled"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   12
         Top             =   720
         Width           =   3735
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   1
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Setup - ASPI Check"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   13
         Top             =   240
         Width           =   3015
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
   Begin ControlResizer.AutoResizer AutoResizer1 
      Height          =   735
      Left            =   1320
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1296
   End
End
Attribute VB_Name = "frmSetupWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ResetWizardFrames(lIndex As Integer)
'On Local Error Resume Next
Dim i As Integer, msg As String, msg2 As String
Select Case lIndex
Case 2
    If optAspi2.Value = True Then
        UpdateASPI True
    Else
        If lUnloadSetupWizardAfterASPI = True Then
            If lEvents.eSettings.iOverwritePrompts = True Then MsgBox "Your ripper cannot continue without ASPI drivers. Please reconsider installing them. They are built for all Windows operating systems (including Windows NT)", vbExclamation
            ResetWizardFrames 1
            lUnloadSetupWizardAfterASPI = False
            Unload Me
            Exit Sub
        End If
    End If
    msg = App.Path & "\programs\NexMediaCDLoader.exe"
    If DoesFileExist(msg) = True Then
        i = AddPlayer("NexMEDIA", App.Path & "\programs\nexmediacdloader.exe", pCDPlayer, "")
        If i <> 0 Then
            lPlayers.pCDPlayerIndex = i
            WriteINI lIniFiles.iPlayers, "Settings", "CDPlayer", lPlayers.pCDPlayerIndex
        Else
            Exit Sub
        End If
    Else
        Exit Sub
    End If
Case 3
    If Option1.Value = True Then
        cmdNext.Enabled = False
        cmdCancel.Enabled = False
        cmdBack.Enabled = False
        lAutoScanHdd = True
        Option1.Visible = False
        Option2.Visible = False
        Option3.Visible = False
        Label7.Caption = "Please wait, searching your hard drive(s) for media"
        frmSearchForMedia.Show 1
        cmdNext.Enabled = True
        cmdCancel.Enabled = True
        cmdBack.Enabled = True
        Label7.Caption = "Done searching"
        Option1.Value = False
    ElseIf Option2.Value = True Then
        frmMain.Show
        frmPlaylist.PromptMediaDir
    End If
End Select
If lIndex = lSetupWizard.sFrameCount Then
    If Playlist.pFileCount = 0 Then
        'Unload frmPlaylist
        chkRandomMP3.Value = 0
        chkRandomMP3.Enabled = False
    End If
    cmdNext.Enabled = False
    cmdFinish.Enabled = True
    For i = 0 To lSetupWizard.sFrameCount
        fraSetup(i).Visible = False
    Next i
    fraSetup(lSetupWizard.sFrameCount).Visible = True
Else
    lSetupWizard.sFrameIndex = lIndex
    For i = 0 To lSetupWizard.sFrameCount
        fraSetup(i).Visible = False
    Next i
    fraSetup(lIndex).Visible = True
End If
If lIndex = 0 Then cmdBack.Enabled = False
If Err.Number <> 0 Then SetError "ResetWizardFrames()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdBack_Click()
'On Local Error Resume Next
If lSetupWizard.sFrameIndex <> 0 Then
    lSetupWizard.sFrameIndex = lSetupWizard.sFrameIndex - 1
    ResetWizardFrames lSetupWizard.sFrameIndex
    cmdNext.Enabled = True
Else
    cmdBack.Enabled = False
End If
If Err.Number <> 0 Then SetError "cmdBack_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdCancel_Click()
'On Local Error Resume Next
Dim msg As VbMsgBoxResult

If lEvents.eSettings.iOverwritePrompts = True Then
    msg = MsgBox("Are you sure you want to cancel?", vbYesNo + vbQuestion)
Else
    msg = vbYes
End If
If msg = vbYes Then
    Unload Me
    frmMain.Show
ElseIf msg = vbNo Then
    Exit Sub
End If
If Err.Number <> 0 Then SetError "cmdCancel_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdFinish_Click()
'On Local Error Resume Next
Dim i As Integer, lFile As String, msg As String

If chkRandomMP3.Value = 1 Then LoadRandomMP3
frmMain.Show
Unload Me
If Err.Number <> 0 Then SetError "cmdFinish_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdHelp0_Click()
'On Local Error Resume Next
MsgBox "NexENCODE will not collect any information for the purpose of spying or user databases. Nor does NexENCODE ever do this, it simply needs to make sure your pc is ready to run NexENCODE, and to search your pc for media for the purpose of cd ripping and mp3 playback", vbInformation + vbMsgBoxHelpButton
End Sub

Private Sub cmdNext_Click()
'On Local Error Resume Next
ResetWizardFrames lSetupWizard.sFrameIndex + 1
cmdBack.Enabled = True
If Err.Number <> 0 Then SetError "cmdNext_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Command1_Click()
'On Local Error Resume Next
MsgBox "Aspi drivers help NexENCODE copy cd audio to your hard drive in a way that is faster than most other processes. Without current up to date ASPI driver, your ripper will not function, meaning you will not be able to copy CD audio to MP3's. Installing these drivers are harmless to your system, however it does require you to reboot.", vbInformation + vbInformation, "Help"
End Sub

Private Sub Command2_Click()
'On Local Error Resume Next
MsgBox "The 'Search Hard Drive(s)' option will take anywhere from 20 seconds to 2 minutes, but will search for all your MP3 files on all your hard drives. The add single folder option, prompts asks you where you keep your MP3 files. The do not search option totally avoids adding MP3 files into your playlists", vbInformation + vbMsgBoxHelpButton
End Sub

Private Sub Command3_Click()
'On Local Error Resume Next
MsgBox "Your just one step away from using NexENCODE Studio! When you click finish, NexENCODE will automatically play an MP3 file (if you have the options to the left checked).", vbMsgBoxHelpButton + vbInformation
End Sub

Private Sub Form_Load()
'On Local Error Resume Next
Dim i As Integer
For i = 0 To 4
    imgIcon(i).Picture = frmGraphics.imgIcon.Picture
Next i
Me.Icon = frmGraphics.Icon
'If DoesFileExist(App.Path & "\skins\inex\inex_top.gif") = True Then Image1.Picture = LoadPicture(App.Path & "\skins\inex\inex_top.gif")
'If DoesFileExist(App.Path & "\skins\inex\inex_sg.gif") = True Then Image2.Picture = LoadPicture(App.Path & "\skins\inex\inex_sg.gif")
lSetupWizard.sFrameCount = 3
ResetWizardFrames 0
DoEvents
ASPIChecker.Init
ASPIChecker.OpenDriveByNumber 1
DoEvents
If ASPIChecker.IsAspiLoaded = False Then
    lRipperSettings.eAspiEnabled = False
Else
    lRipperSettings.eAspiEnabled = True
End If
If lRipperSettings.eAspiEnabled = True Then
    optAspi2.Enabled = False
    optAspi3.Enabled = False
    optAspi1.Value = True
Else
    optAspi1.Enabled = False
    optAspi2.Value = True
End If
If Err.Number <> 0 Then SetError "Form_Load()", lEvents.eSettings.iErrDescription, Err.Description
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
