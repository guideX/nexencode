VERSION 5.00
Object = "{EE128208-4F73-11D3-83BB-C47C02EE3D01}#1.0#0"; "ControlResizer.ocx"
Begin VB.Form frmTextViewer 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "NexENCODE - Text Viewer"
   ClientHeight    =   2745
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   5805
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
   ScaleHeight     =   2745
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.TextBox txtInformation 
      Height          =   2055
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   120
      Width           =   5655
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
      Top             =   2280
      Width           =   6015
      Begin VB.CommandButton cmdOK 
         Caption         =   "Close"
         Default         =   -1  'True
         Height          =   315
         Left            =   4680
         TabIndex        =   1
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
   Begin ControlResizer.AutoResizer AutoResizer1 
      Height          =   30
      Left            =   2040
      Top             =   120
      Visible         =   0   'False
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   53
   End
End
Attribute VB_Name = "frmTextViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
'On Local Error Resume Next
'imgNS4.Picture = frmGraphics.imgIcon.Picture
'Image1.Picture = frmGraphics.imgTopper.Picture
'Image2.Picture = frmGraphics.imgSideGradient.Picture
FlashIN Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Local Error Resume Next
FlashOut Me
End Sub
