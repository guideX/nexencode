VERSION 5.00
Object = "{EE128208-4F73-11D3-83BB-C47C02EE3D01}#1.0#0"; "ControlResizer.ocx"
Begin VB.Form frmAbout_NS 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "NexENCODE - About"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout_NS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   325
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      ToolTipText     =   "Hide this window"
      Top             =   4440
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   3015
      TabIndex        =   3
      Top             =   0
      Width           =   3015
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   4
         X1              =   0
         X2              =   3000
         Y1              =   750
         Y2              =   750
      End
      Begin VB.Label lblApp 
         BackStyle       =   0  'Transparent
         Caption         =   "NexENCODE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   720
         TabIndex        =   5
         ToolTipText     =   "Program Name"
         Top             =   120
         Width           =   2055
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Left            =   120
         Stretch         =   -1  'True
         ToolTipText     =   "NexENCODE Icon"
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "Newest version"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   4
         ToolTipText     =   "Program Version"
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   3015
      TabIndex        =   2
      Top             =   4320
      Width           =   3015
      Begin VB.Label lblEmailGuideX 
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail Leon Aiossa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         X1              =   0
         X2              =   3000
         Y1              =   0
         Y2              =   0
      End
   End
   Begin ControlResizer.AutoResizer AutoResizer1 
      Height          =   855
      Left            =   720
      Top             =   2400
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "About credits and information"
      Top             =   960
      Width           =   2775
   End
End
Attribute VB_Name = "frmAbout_NS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
'On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then SetError "cmdClose_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Load()
'On Local Error Resume Next
imgIcon.Picture = frmGraphics.imgIcon.Picture
lblEmailGuideX.MouseIcon = frmMain.imgRip.MouseIcon
'lblEmailGuideX.MousePointer = 99
Icon = frmMain.Icon
lblInfo.Caption = "Programming by Leon J Aiossa (guideX) NexENCODE concept by 'Warlok'. First version of NexENCODE developed by him. Graphics concept originally developed by Colin Foss (KnightFal) in v3.02. Graphics in this version entirely developed by Leon J Aiossa. Thanks to #gnnchat for ruining their pc's in the testing of NexENCODE Studio 4. Thanks to Jason Bird for beta testing. This program could not have been written without all the faithfull users throughout the years. NexENCODE and NexENCODE Studio are registered trademarks of Team Nexgen Inc."
FlashIN frmAbout_NS
lblVersion.Caption = "Version " & App.Major & "." & App.Minor
If Err.Number <> 0 Then SetError "Form_Load", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then SetError "Form_MouseDown()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Local Error Resume Next
FlashOut frmAbout_NS
If Err.Number <> 0 Then SetError "Form_Unload()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgIcon_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then SetError "imgIcon_MouseDown()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub lblApp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then SetError "lblApp_MouseDOwn()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub lblEmailGuideX_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Local Error Resume Next
If Button = 1 Then
    lblEmailGuideX.ForeColor = vbWhite
End If
If Err.Number <> 0 Then SetError "lblEmailGuideX()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub lblEmailGuideX_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Local Error Resume Next
If Button = 1 Then
    PlayWav App.Path & "\media\click.wav", SND_ASYNC
    lblEmailGuideX.ForeColor = vbBlue
    Surf "mailto:guide_X@live.com"
End If
If Err.Number <> 0 Then SetError "lblEmailGuideX_MouseUp()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub lblInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then SetError "lblinfo_MouseDown()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub lblVersion_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then SetError "lblVersion_MouseDown()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then SetError "picture1_mousedown()", lEvents.eSettings.iErrDescription, Err.Description
End Sub
