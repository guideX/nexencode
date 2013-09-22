VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EE128208-4F73-11D3-83BB-C47C02EE3D01}#1.0#0"; "ControlResizer.ocx"
Object = "{7314ED99-8643-4E82-A4F8-5E9F4DEC14BE}#1.0#0"; "VolumeControl.ocx"
Begin VB.Form frmVolume 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "NexENCODE - Volume"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   3930
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
   ScaleHeight     =   3720
   ScaleWidth      =   3930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picBottom 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   6015
      TabIndex        =   7
      Top             =   3240
      Width           =   6015
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Mute"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   315
         Left            =   2640
         TabIndex        =   8
         Top             =   120
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         X1              =   0
         X2              =   6000
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.OptionButton optVolType 
      Caption         =   "Auxiliary"
      Height          =   375
      Index           =   6
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2280
      Width           =   3735
   End
   Begin VB.OptionButton optVolType 
      Caption         =   "Wave"
      Height          =   375
      Index           =   5
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   3735
   End
   Begin VB.OptionButton optVolType 
      Caption         =   "CD Audio"
      Height          =   375
      Index           =   4
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   3735
   End
   Begin VB.OptionButton optVolType 
      Caption         =   "Synthesizer"
      Height          =   375
      Index           =   3
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   3735
   End
   Begin VB.OptionButton optVolType 
      Caption         =   "Microphone"
      Height          =   375
      Index           =   2
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   3735
   End
   Begin VB.OptionButton optVolType 
      Caption         =   "Line In"
      Height          =   375
      Index           =   1
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   3735
   End
   Begin VB.OptionButton optVolType 
      Caption         =   "Master"
      Height          =   375
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin MSComctlLib.Slider sldVolume 
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   661
      _Version        =   393216
      Max             =   100
      TickStyle       =   3
   End
   Begin ControlResizer.AutoResizer AutoResizer1 
      Height          =   135
      Left            =   3000
      Top             =   240
      Visible         =   0   'False
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   238
   End
   Begin VolControl.VolumeControl VolumeControl1 
      Left            =   3000
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      Volume          =   49
   End
End
Attribute VB_Name = "frmVolume"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
'On Local Error Resume Next
If Check1.Value = vbChecked Then
    VolumeControl1.Mute = True
Else
    VolumeControl1.Mute = False
End If
If Err.Number <> 0 Then SetError "Check1_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdClose_Click()
'On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then SetError "cmdClose_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Load()
'On Local Error Resume Next
Icon = frmMain.Icon
FlashIN Me
optVolType(VolumeControl1.DeviceToControl).Value = True
sldVolume.Value = VolumeControl1.Volume
If Err.Number <> 0 Then SetError "Form_Load", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Local Error Resume Next
FlashOut Me
If Err.Number <> 0 Then SetError "Form_Unload", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgTopper_DblClick()
'On Local Error Resume Next
If Me.WindowState = vbMaximized Then
    Me.WindowState = vbNormal
Else
    Me.WindowState = vbMaximized
End If
If Err.Number <> 0 Then SetError "imgTopper_DblClick()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgTopper_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then SetError "imgTopper_MouseDown()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub optVolType_Click(Index As Integer)
'On Local Error Resume Next
VolumeControl1.DeviceToControl = Index
If Err.Number <> 0 Then SetError "OptVolType_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub sldVolume_Scroll()
'On Local Error Resume Next
VolumeControl1.Volume = sldVolume.Value
If Err.Number <> 0 Then SetError "sldVolume_Scroll()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub VolumeControl1_MuteChanged(NewMute As Boolean)
'On Local Error Resume Next
If NewMute Then
    Check1.Value = vbChecked
Else
    Check1.Value = vbUnchecked
End If
If Err.Number <> 0 Then SetError "VolumeControl1_MuteChanged()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub VolumeControl1_VolumeChanged(NewVolume As Long)
'On Local Error Resume Next
sldVolume.Value = NewVolume
If Err.Number <> 0 Then SetError "VolumeControl1_VolumeChanged()", lEvents.eSettings.iErrDescription, Err.Description
End Sub
