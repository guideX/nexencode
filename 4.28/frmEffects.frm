VERSION 5.00
Object = "{EE128208-4F73-11D3-83BB-C47C02EE3D01}#1.0#0"; "ControlResizer.ocx"
Begin VB.Form frmEffects 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "NexENCODE - Effects"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   5940
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
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optEffect 
      Caption         =   "Shifting"
      Height          =   255
      Index           =   10
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   1695
   End
   Begin VB.OptionButton optEffect 
      Caption         =   "Reverb"
      Height          =   255
      Index           =   9
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   1695
   End
   Begin VB.OptionButton optEffect 
      Caption         =   "FadeOut"
      Height          =   255
      Index           =   6
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1560
      Width           =   1695
   End
   Begin VB.OptionButton optEffect 
      Caption         =   "FadeIn"
      Height          =   255
      Index           =   5
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1320
      Width           =   1695
   End
   Begin VB.OptionButton optEffect 
      Caption         =   "Eq"
      Height          =   255
      Index           =   4
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2760
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.OptionButton optEffect 
      Caption         =   "Echo"
      Height          =   255
      Index           =   3
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1080
      Width           =   1695
   End
   Begin VB.OptionButton optEffect 
      Caption         =   "Distortion"
      Height          =   255
      Index           =   2
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   840
      Width           =   1695
   End
   Begin VB.OptionButton optEffect 
      Caption         =   "CFilter"
      Height          =   255
      Index           =   1
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   600
      Width           =   1695
   End
   Begin VB.OptionButton optEffect 
      Caption         =   "Chorus"
      Height          =   255
      Index           =   8
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   360
      Width           =   1695
   End
   Begin VB.OptionButton optEffect 
      Caption         =   "Amplitude"
      Height          =   255
      Index           =   7
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   1695
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
      Top             =   3240
      Width           =   6015
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   315
         Left            =   4680
         TabIndex        =   2
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "Apply Effect"
         Default         =   -1  'True
         Height          =   315
         Left            =   3360
         TabIndex        =   1
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
   Begin VB.Frame fraEffect 
      BorderStyle     =   0  'None
      Caption         =   "Eq"
      Height          =   3015
      Index           =   4
      Left            =   1920
      TabIndex        =   20
      Top             =   120
      Width           =   3975
   End
   Begin VB.Frame fraEffect 
      BorderStyle     =   0  'None
      Caption         =   "FadeOut"
      Height          =   3015
      Index           =   6
      Left            =   1920
      TabIndex        =   18
      Top             =   120
      Width           =   3975
      Begin VB.TextBox txtFadeOutSeconds10th 
         Height          =   285
         Left            =   1800
         TabIndex        =   84
         Text            =   "50"
         Top             =   0
         Width           =   2055
      End
      Begin VB.Label Label35 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Seconds 10th:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   83
         Top             =   0
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Play backwards with Volume 0 and brings the Volume within Seconds10th on the normal Volume."
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   2520
         Width           =   3735
      End
   End
   Begin VB.Frame fraEffect 
      BorderStyle     =   0  'None
      Caption         =   "FadeIn"
      Height          =   3015
      Index           =   5
      Left            =   1920
      TabIndex        =   19
      Top             =   120
      Width           =   3975
      Begin VB.TextBox txtSeconds10th 
         Height          =   285
         Left            =   1800
         TabIndex        =   82
         Text            =   "50"
         Top             =   0
         Width           =   2055
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "Seconds 10th:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   81
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Starts to play with Volume 0 and brings the Volume within Seconds10th on the normal Volume."
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   2520
         Width           =   3735
      End
   End
   Begin VB.Frame fraEffect 
      BorderStyle     =   0  'None
      Caption         =   "Reverb"
      Height          =   3015
      Index           =   9
      Left            =   1920
      TabIndex        =   15
      Top             =   120
      Width           =   3975
      Begin VB.TextBox txtRatio 
         Height          =   285
         Left            =   1920
         TabIndex        =   32
         Text            =   "90"
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtDelay 
         Height          =   285
         Left            =   1920
         TabIndex        =   30
         Text            =   "900"
         Top             =   0
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Ratio (% of effect):"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Delay (in ms):"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Aplies a reverb-effect on the waveform"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   2640
         Width           =   3735
      End
   End
   Begin VB.Frame fraEffect 
      BorderStyle     =   0  'None
      Caption         =   "Amplitude"
      Height          =   3015
      Index           =   7
      Left            =   1920
      TabIndex        =   17
      Top             =   120
      Width           =   3975
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Builds the sum of all points and shifts them, so that the sum will be 0"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   2520
         Width           =   3735
      End
   End
   Begin VB.Frame fraEffect 
      BorderStyle     =   0  'None
      Caption         =   "Shifting"
      Height          =   3015
      Index           =   10
      Left            =   1920
      TabIndex        =   14
      Top             =   120
      Width           =   3975
      Begin VB.TextBox txtLongSize 
         Height          =   285
         Left            =   1800
         TabIndex        =   37
         Text            =   "2048"
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtShortMode 
         Height          =   285
         Left            =   1800
         TabIndex        =   35
         Text            =   "1"
         Top             =   0
         Width           =   2055
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Long Size:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Short Mode:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "mode takes values from 0 ~ 10% UP, 1 ~ 20% UP, ... and 1 ~ -10% DOWN, -2 ~ 20% DOWN"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   2520
         Width           =   3735
      End
   End
   Begin VB.Frame fraEffect 
      BorderStyle     =   0  'None
      Caption         =   "Distortion"
      Height          =   3015
      Index           =   2
      Left            =   1920
      TabIndex        =   13
      Top             =   120
      Width           =   3975
      Begin VB.TextBox txtClamp 
         Height          =   285
         Left            =   1920
         TabIndex        =   46
         Text            =   "0"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtDistorted 
         Height          =   285
         Left            =   1920
         TabIndex        =   42
         Text            =   "560 "
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtGate 
         Height          =   285
         Left            =   1920
         TabIndex        =   48
         Text            =   "0"
         Top             =   960
         Width           =   1935
      End
      Begin VB.ComboBox cboDistortionPresets 
         Height          =   315
         Left            =   1920
         TabIndex        =   51
         Top             =   0
         Width           =   1935
      End
      Begin VB.TextBox txtThreshold 
         Height          =   285
         Left            =   1920
         TabIndex        =   44
         Text            =   "3"
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtDry 
         Height          =   285
         Left            =   1920
         TabIndex        =   40
         Text            =   "1005"
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Preset:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Gate:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Clamp:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Threshold:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Distorted:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Dry:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmEffects.frx":0000
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   120
         TabIndex        =   38
         Top             =   2280
         Width           =   3735
      End
   End
   Begin VB.Frame fraEffect 
      BorderStyle     =   0  'None
      Caption         =   "Echo"
      Height          =   3015
      Index           =   3
      Left            =   1920
      TabIndex        =   21
      Top             =   120
      Width           =   3975
      Begin VB.TextBox txtEchoRatio 
         Height          =   285
         Left            =   2040
         TabIndex        =   76
         Top             =   720
         Width           =   1815
      End
      Begin VB.ComboBox cboEchoPresets 
         Height          =   315
         Left            =   2040
         TabIndex        =   75
         Top             =   0
         Width           =   1815
      End
      Begin VB.TextBox txtEchoDelay 
         Height          =   285
         Left            =   2040
         TabIndex        =   74
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "delay in ms, ratio = Percent of the Echo "
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   80
         Top             =   2640
         Width           =   3735
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Echo Ratio:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   77
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Presets:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   73
         Top             =   0
         Width           =   1935
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Short Delay:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   72
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame fraWelcome 
      BorderStyle     =   0  'None
      Caption         =   "Welcome"
      Height          =   3015
      Left            =   1920
      TabIndex        =   78
      Top             =   120
      Width           =   3975
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Click the buttons to the left to access the effects proporties."
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   120
         TabIndex        =   79
         Top             =   840
         Width           =   3735
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   765
         Left            =   120
         Picture         =   "frmEffects.frx":008B
         Top             =   0
         Width           =   3765
      End
   End
   Begin VB.Frame fraEffect 
      BorderStyle     =   0  'None
      Caption         =   "CFilter"
      Height          =   3015
      Index           =   1
      Left            =   1920
      TabIndex        =   22
      Top             =   120
      Width           =   3975
      Begin VB.TextBox txtCFilterFactor 
         Height          =   285
         Left            =   2160
         TabIndex        =   23
         Text            =   "5"
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmEffects.frx":22AD
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   120
         TabIndex        =   49
         Top             =   2040
         Width           =   3735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Factor:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.Frame fraEffect 
      BorderStyle     =   0  'None
      Caption         =   "Chorus"
      Height          =   3015
      Index           =   8
      Left            =   1920
      TabIndex        =   16
      Top             =   120
      Width           =   3975
      Begin VB.TextBox txtShortFeedback 
         Height          =   285
         Left            =   1920
         TabIndex        =   69
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox txtShortMixing 
         Height          =   285
         Left            =   1920
         TabIndex        =   67
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox txtInvertFeedback 
         Height          =   285
         Left            =   1920
         TabIndex        =   65
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox txtShortWet 
         Height          =   285
         Left            =   1920
         TabIndex        =   71
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtShortDry 
         Height          =   285
         Left            =   1920
         TabIndex        =   63
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtWaveForm 
         Height          =   285
         Left            =   1920
         TabIndex        =   61
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtFloatRate 
         Height          =   285
         Left            =   1920
         TabIndex        =   59
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtShortDepth 
         Height          =   285
         Left            =   1920
         TabIndex        =   56
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtShortDelay 
         Height          =   285
         Left            =   1920
         TabIndex        =   55
         Top             =   480
         Width           =   1935
      End
      Begin VB.ComboBox cboChorusPresets 
         Height          =   315
         Left            =   1920
         TabIndex        =   53
         Top             =   0
         Width           =   1935
      End
      Begin ControlResizer.AutoResizer AutoResizer1 
         Height          =   375
         Left            =   360
         Top             =   1200
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Short Wet:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Short Feedback:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Short Mixing:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Invert Feedback:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Short Dry:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Wave Form:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Float Rate:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Short Depth:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Short Delay:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Presets:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   0
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmEffects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim f As Integer
Option Explicit

Private Sub cboChorusPresets_Click()
On Local Error Resume Next
Dim i As Integer, m As Integer
For i = 1 To lEffectsPresets.eChorus.cCount
    If LCase(lEffectsPresets.eChorus.cChorus(i).cDescription) = LCase(cboChorusPresets.Text) Then
        m = i
    End If
Next i
If m <> 0 Then
    txtShortDelay.Text = lEffectsPresets.eChorus.cChorus(m).cShortDelay
    txtShortDepth.Text = lEffectsPresets.eChorus.cChorus(m).cShortDepth
    txtFloatRate.Text = lEffectsPresets.eChorus.cChorus(m).cFloatRate
    txtWaveForm.Text = lEffectsPresets.eChorus.cChorus(m).cWaveForm
    txtShortDry.Text = lEffectsPresets.eChorus.cChorus(m).cShortDry
    txtInvertFeedback.Text = lEffectsPresets.eChorus.cChorus(m).cInvertFeedback
    txtShortMixing.Text = lEffectsPresets.eChorus.cChorus(m).cShortMixing
    txtShortFeedback.Text = lEffectsPresets.eChorus.cChorus(m).cShortFeedback
    txtShortWet.Text = lEffectsPresets.eChorus.cChorus(m).cShortWet
End If
If Err.Number <> 0 Then SetError "cboChorusPresets_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cboDistortionPresets_Click()
On Local Error Resume Next
Dim i As Integer, m As Integer
For i = 1 To lEffectsPresets.eDistortion.dCount
    If LCase(lEffectsPresets.eDistortion.dDistortion(i).lDescription) = LCase(cboDistortionPresets.Text) Then
        m = i
    End If
Next i
If m <> 0 Then
    txtDry.Text = lEffectsPresets.eDistortion.dDistortion(m).lDry
    txtThreshold.Text = lEffectsPresets.eDistortion.dDistortion(m).lThreshold
    txtGate.Text = lEffectsPresets.eDistortion.dDistortion(m).lGate
    txtDistorted.Text = lEffectsPresets.eDistortion.dDistortion(m).lDistorted
    txtClamp.Text = lEffectsPresets.eDistortion.dDistortion(m).lClamp
End If
If Err.Number <> 0 Then SetError "cboDistortionPresets_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cboEchoPresets_Click()
On Local Error Resume Next
Dim i As Integer, m As Integer
For i = 1 To lEffectsPresets.eEcho.eCount
    If LCase(lEffectsPresets.eEcho.eEcho(i).eDescription) = LCase(cboEchoPresets.Text) Then
        m = i
    End If
Next i
If m <> 0 Then
    txtEchoDelay.Text = lEffectsPresets.eEcho.eEcho(m).eShortDelay
    txtEchoRatio.Text = lEffectsPresets.eEcho.eEcho(m).eShortRatio
End If
If Err.Number <> 0 Then SetError "cboEchoPresets_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdCancel_Click()
On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then SetError "cmdCancel_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdOK_Click()
On Local Error Resume Next
Select Case f
Case 1
    AddCFilter Int(txtCFilterFactor.Text)
Case 2
    AddDistortion Int(txtDry.Text), Int(txtDistorted.Text), Int(txtThreshold.Text), Int(txtClamp.Text), Int(txtGate.Text)
Case 3
    AddEcho Int(txtEchoDelay.Text), Int(txtEchoRatio.Text)
Case 5
    AddFadeIn txtSeconds10th.Text
Case 6
    AddFadeOut txtFadeOutSeconds10th.Text
Case 7
    AddAmplitude
Case 8
    AddChorus Int(txtShortDelay.Text), Int(txtShortDepth.Text), Int(txtFloatRate.Text), Int(txtWaveForm.Text), Int(txtShortDry.Text), Int(txtShortWet.Text), Int(txtInvertFeedback.Text), Int(txtShortMixing.Text), Int(txtShortFeedback.Text)
Case 9
    AddReverb Int(txtDelay.Text), Int(txtRatio.Text)
Case 10
    AddShifting txtShortMode.Text, txtLongSize.Text
End Select
Unload Me
If Err.Number <> 0 Then SetError "cmd_OK_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Load()
On Local Error Resume Next
Dim msg As String, i As Integer
LoadEffectsPresets
f = 7
Me.Icon = frmMain.Icon
For i = 1 To 10
    fraEffect(i).Visible = False
Next i
msg = ReadINI(lIniFiles.iEffects, "Settings", "DefaultEcho", "Default")
For i = 1 To lEffectsPresets.eEcho.eCount
    If lEffectsPresets.eEcho.eEcho(i).eEnabled = True Then
        cboEchoPresets.AddItem lEffectsPresets.eEcho.eEcho(i).eDescription
        If Len(msg) <> 0 And LCase(msg) = LCase(lEffectsPresets.eEcho.eEcho(i).eDescription) Then
            cboEchoPresets.ListIndex = FindComoboxIndex(cboEchoPresets, lEffectsPresets.eEcho.eEcho(i).eDescription)
        End If
    End If
Next i
msg = ReadINI(lIniFiles.iEffects, "Settings", "DefaultDistortion", "Default")
For i = 1 To lEffectsPresets.eDistortion.dCount
    If lEffectsPresets.eDistortion.dDistortion(i).lEnabled = True Then
        cboDistortionPresets.AddItem lEffectsPresets.eDistortion.dDistortion(i).lDescription
        If Len(msg) <> 0 And LCase(msg) = LCase(lEffectsPresets.eDistortion.dDistortion(i).lDescription) Then
            cboDistortionPresets.ListIndex = FindComoboxIndex(cboDistortionPresets, lEffectsPresets.eDistortion.dDistortion(i).lDescription)
        End If
    End If
Next i
msg = ReadINI(lIniFiles.iEffects, "Settings", "DefaultChorus", "Default")
For i = 1 To lEffectsPresets.eChorus.cCount
    If lEffectsPresets.eChorus.cChorus(i).cEnabled = True Then
        cboChorusPresets.AddItem lEffectsPresets.eChorus.cChorus(i).cDescription
        If Len(msg) <> 0 And LCase(msg) = LCase(lEffectsPresets.eChorus.cChorus(i).cDescription) Then
            cboChorusPresets.ListIndex = FindComoboxIndex(cboChorusPresets, lEffectsPresets.eChorus.cChorus(i).cDescription)
        End If
    End If
Next i
If Err.Number <> 0 Then SetError "Form_Load()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub optEffect_Click(Index As Integer)
On Local Error Resume Next
Dim i As Integer, h As Integer
fraWelcome.Visible = False
For i = 1 To 10
    fraEffect(i).Visible = False
Next i
fraEffect(Index).Visible = True
f = Index
If Err.Number <> 0 Then SetError "optEffect_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub
