VERSION 5.00
Object = "{EE128208-4F73-11D3-83BB-C47C02EE3D01}#1.0#0"; "ControlResizer.ocx"
Begin VB.Form frmErrHandler 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "NexENCODE - Error"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
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
   Icon            =   "frmErrHandler.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   209
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   396
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtInfo 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   120
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1680
      Width           =   5775
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   6015
      TabIndex        =   1
      Top             =   2640
      Width           =   6015
      Begin VB.ComboBox cboErrorReport 
         Height          =   315
         ItemData        =   "frmErrHandler.frx":000C
         Left            =   120
         List            =   "frmErrHandler.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   120
         Width           =   2775
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Close"
         Default         =   -1  'True
         Height          =   315
         Left            =   4080
         TabIndex        =   3
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdEnd 
         Caption         =   "&End"
         Height          =   315
         Left            =   5040
         TabIndex        =   2
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
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   5775
   End
   Begin VB.Label lblSub 
      BackStyle       =   0  'Transparent
      Caption         =   "Sub or Function:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   5775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended information:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   5775
   End
   Begin VB.Label lblErrCount 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2760
      Width           =   5535
   End
   Begin ControlResizer.AutoResizer AutoResizer1 
      Height          =   375
      Left            =   840
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "An error was controled, it is described below"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frmErrHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lCheckForBackgroundSetting As Boolean
Option Explicit

Private Sub cmdEnd_Click()
'On Local Error Resume Next
End
If Err.Number <> 0 Then SetError "cmdEnd_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdOK_Click()
'On Local Error Resume Next
PlayWav App.Path & "\media\done.wav", SND_ASYNC
Select Case cboErrorReport.ListIndex
Case 0
    Dim msg As String, msg2 As String, b As Boolean
    msg = Trim(SaveDialog(frmMain, "Text Files (*.txt)|*.txt", "Save error report as..", App.Path))
    msg2 = "Error Report" & vbCrLf & lblDescription.Caption & vbCrLf & lblSub.Caption & vbCrLf & "Extended: " & txtInfo.Text
    msg = Left(msg, Len(msg) - 1) & ".txt"
    SaveFile msg, msg2
End Select
Unload Me
If Err.Number <> 0 Then SetError "cmdOK_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Load()
'On Local Error Resume Next
lCheckForBackgroundSetting = lEvents.eSettings.iCheckForActiveWindow
lEvents.eSettings.iCheckForActiveWindow = False
Icon = frmMain.Icon
cboErrorReport.ListIndex = 1
frmMain.Picture = frmMain.imgErrorBackground.Picture
FlashIN frmErrHandler
lblErrCount.Caption = "Total number of error generated by NexENCODE: " & lEvents.eErrCount
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Local Error Resume Next
frmMain.Picture = LoadPicture(lSkins.sSkin(lSkins.sSkinIndex).sFilepath & lSkins.sSkin(lSkins.sSkinIndex).sGraphic)
lEvents.eSettings.iCheckForActiveWindow = lCheckForBackgroundSetting
FlashOut Me
If Err.Number <> 0 Then SetError "Form_Unload()", lEvents.eSettings.iErrDescription, Err.Description
End Sub
