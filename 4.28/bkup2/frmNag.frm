VERSION 5.00
Object = "{EE128208-4F73-11D3-83BB-C47C02EE3D01}#1.0#0"; "ControlResizer.ocx"
Begin VB.Form frmNag 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "NexENCODE"
   ClientHeight    =   840
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4395
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNag.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   56
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   293
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   4800
      Top             =   1080
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   5895
      TabIndex        =   1
      Top             =   360
      Width           =   5895
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Continue"
         Enabled         =   0   'False
         Height          =   315
         Left            =   3000
         TabIndex        =   4
         ToolTipText     =   "Hide Window, Continue loading NexENCODE"
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         ToolTipText     =   "Close NexENCODE"
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdRegister 
         Caption         =   "&Register"
         Height          =   315
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Show the registration dialog"
         Top             =   120
         Width           =   1335
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
   Begin ControlResizer.AutoResizer AutoResizer1 
      Height          =   375
      Left            =   5280
      Top             =   600
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please register this copy of NexENCODE"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmNag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
On Local Error Resume Next
Timer1.Enabled = False
Unload Me
End Sub

Private Sub cmdExit_Click()
On Local Error Resume Next
Timer1.Enabled = False
End
If Err.Number <> 0 Then SetError "cmdExit_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdRegister_Click()
On Local Error Resume Next
Unload Me
frmRegister.Show 1
If lEvents.eRegistered = False Then frmNag.Show 1
If Err.Number <> 0 Then SetError "cmdRegister_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Load()
On Local Error Resume Next
Icon = frmGraphics.Icon
'If DoesFileExist(App.Path & "\skins\inex\inex_top.gif") = True Then Image1.Picture = LoadPicture(App.Path & "\skins\inex\inex_top.gif")
If Err.Number <> 0 Then SetError "Form_Load()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then SetError "img1_mousedown", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Timer1_Timer()
On Local Error Resume Next
Timer1.Enabled = False
cmdClose.Enabled = True
End Sub
