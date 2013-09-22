VERSION 5.00
Object = "{EE128208-4F73-11D3-83BB-C47C02EE3D01}#1.0#0"; "ControlResizer.ocx"
Begin VB.Form frmLatestVersionCheck 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Update Check"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   2880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLatestVersionCheck.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   2880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkAlwaysCheckForUpdates 
      Caption         =   "Always check for updates"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Check this to check for updates on NexENCODE startup"
      Top             =   3000
      Width           =   2655
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   315
      Left            =   1680
      TabIndex        =   6
      ToolTipText     =   "Cancel/Hide this window"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdDownload 
      Caption         =   "Download"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   315
      Left            =   480
      TabIndex        =   5
      ToolTipText     =   "Get this upgrade"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox txtMyVersion 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      ToolTipText     =   "The Version you are using"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox txtLatestVersion 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      ToolTipText     =   "Latest version of this software available"
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Image picNexENCODE 
      Height          =   825
      Left            =   120
      Top             =   120
      Width           =   2625
   End
   Begin ControlResizer.AutoResizer AutoResizer1 
      Height          =   255
      Left            =   1080
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   4
      X1              =   0
      X2              =   3000
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   4
      X1              =   0
      X2              =   3000
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label lblInfo 
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Information about the upgrade"
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Your Version:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Latest Version:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   4
      X1              =   0
      X2              =   3000
      Y1              =   1100
      Y2              =   1100
   End
End
Attribute VB_Name = "frmLatestVersionCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDownload_Click()
'On Local Error Resume Next
Dim msg As String, lini As String
msg = cmdDownload.Tag
If Len(msg) <> 0 Then
    Surf msg
End If
If Err.Number <> 0 Then SetError "cmdDownload_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdExit_Click()
'On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then SetError "cmdExit_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Load()
'On Local Error Resume Next
picNexENCODE.Picture = frmGraphics.picNS4.Picture
txtMyVersion.Text = App.Major & "." & App.Minor
txtLatestVersion.Text = ReadINI(lIniFiles.iUpdate, "Settings", "Version", "")
If Len(txtLatestVersion.Text) = 0 Then
    lblInfo.Caption = "Unable to define the latest version, if you are not connected to the internet, connect and try again"
    Exit Sub
End If
If txtMyVersion.Text <> txtLatestVersion.Text Then
    lblInfo.Caption = ReadINI(lIniFiles.iUpdate, "Settings", "Description", "")
    cmdDownload.Tag = ReadINI(lIniFiles.iUpdate, "Settings", "Location", "")
    If Len(cmdDownload.Tag) <> 0 Then
        cmdDownload.Enabled = True
    End If
Else
    lblInfo.Caption = "You are running the latest version of NexENCODE Studio"
End If
If lEvents.eSettings.iUpdateCheck = True Then chkAlwaysCheckForUpdates.Value = 1
If Err.Number <> 0 Then SetError "Form_Load", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then SetError "Form_MouseDown()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Local Error Resume Next
frmMain.wskUpdate.Close
If chkAlwaysCheckForUpdates.Value = 0 Then
    WriteINI lIniFiles.iSettings, "Settings", "UpdateCheck", "False"
    lEvents.eSettings.iUpdateCheck = False
Else
    WriteINI lIniFiles.iSettings, "Settings", "UpdateCheck", "True"
    lEvents.eSettings.iUpdateCheck = True
End If
If Err.Number <> 0 Then SetError "Form_Unload", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then SetError "label1_mousedown()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then SetError "tmrDots_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub lblInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then SetError "lblInfo_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub picNexENCODE_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then SetError "picNexENCODE_MouseDOwn()", lEvents.eSettings.iErrDescription, Err.Description
End Sub
