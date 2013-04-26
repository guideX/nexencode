VERSION 5.00
Begin VB.Form frmThanks 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "NexENCODE"
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   3600
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   3600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   2160
      Width           =   3735
      Begin VB.CommandButton cmdRestartNow 
         Caption         =   "Restart NexENCODE Now"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   3375
      End
   End
   Begin VB.Timer tmrCountdown 
      Interval        =   1000
      Left            =   3000
      Top             =   1560
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.Label lblThanks 
         BackStyle       =   0  'Transparent
         Caption         =   "Thanks for registering!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   720
         TabIndex        =   1
         Top             =   240
         Width           =   3135
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Left            =   120
         Picture         =   "frmThanks.frx":0000
         Stretch         =   -1  'True
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.Label lblCountDown 
      Caption         =   "Restarting NexENCODE in ..."
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   3495
   End
   Begin VB.Label lblMessage 
      Caption         =   $"frmThanks.frx":08CA
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   3375
   End
End
Attribute VB_Name = "frmThanks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lCountDown As Integer

Private Sub cmdRestartNow_Click()
''on local error resume next

tmrCountdown.Enabled = False
App.Title = "NS4 BAD INSTANCE"
Shell App.Path & "\nexShell.exe " & App.Path & "\NexENCODE.exe", vbNormalFocus
'Shell App.Path & "\NexENCODE.exe", vbNormalFocus
UnloadMain

If Err.Number <> 0 Then SetError "cmdRestartNow_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Load()
''on local error resume next

AlwaysOnTop Me, True
'frmMain.Visible = False
'If lEvents.ePlaylistVisible = True Then Unload frmPlaylist
lCountDown = 10
tmrCountdown.Enabled = True
lblCountDown.Caption = "Restarting NexENCODE in " & lCountDown

If Err.Number <> 0 Then SetError "Form_Load()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub tmrCountdown_Timer()
''on local error resume next

lCountDown = lCountDown - 1
If lCountDown <> 0 Then
    lblCountDown.Caption = "Restarting NexENCODE in " & lCountDown
Else
    tmrCountdown.Enabled = False
    App.Title = "NS4 BAD INSTANCE"
    'Shell App.Path & "\NexENCODE.exe", vbNormalFocus
    Shell App.Path & "\nexShell.exe " & App.Path & "\NexENCODE.exe", vbNormalFocus
    UnloadMain
End If

If Err.Number <> 0 Then SetError "tmrCountDown_Timer()", lEvents.eSettings.iErrDescription, Err.Description
End Sub
