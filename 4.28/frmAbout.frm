VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   "NexENCODE"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4410
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   322
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   294
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer tmrUnload 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   600
      Top             =   120
   End
   Begin VB.CheckBox chkAbout 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   600
      TabIndex        =   0
      Top             =   3000
      Width           =   195
   End
   Begin VB.Timer tmrDots 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   120
      Top             =   120
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Shape three 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   90
      Left            =   3120
      Shape           =   2  'Oval
      Top             =   3210
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape two 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   90
      Left            =   3000
      Shape           =   2  'Oval
      Top             =   3210
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape one 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   90
      Left            =   2880
      Shape           =   2  'Oval
      Top             =   3210
      Visible         =   0   'False
      Width           =   90
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public jk As Integer
Dim lCancel As Boolean
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Long) As Long

Public Sub SetCancel(mCancel As Boolean)
On Local Error Resume Next
lCancel = mCancel
If Err.Number <> 0 Then SetError "SetCancel()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub SetAboutShape()
On Local Error Resume Next
Dim i As Integer
Dim rgn As Long, tmp As Long, X As Long, Y As Long
GetWindowSettings hWnd
X = lMainWndSettings.wWindowBorder
Y = lMainWndSettings.wTitleBarHeight
rgn = CreateEllipticRgn(X + 1, Y + 1, X + 292, Y + 288)
tmp = SetWindowRgn(Me.hWnd, rgn, True)
If Err.Number <> 0 Then SetError "SetAboutShape()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_DblClick()
On Local Error Resume Next
Me.Visible = False
If Err.Number <> 0 Then SetError "frmAbout_DBLClick()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Load()
On Local Error Resume Next
Dim i As Integer, X As Integer, msg As String
If DoesFileExist(App.Path & "\skins\inex\about.png") = True Then Me.Picture = LoadPicture(App.Path & "\skins\inex\about.gif")
Icon = frmGraphics.Icon
lEvents.eSettings.iCommand = Command$
PreLoadSettings
LoadDrives
DoEvents
If App.PrevInstance = True Then
    ActivatePrevInstance
    Unload Me
    Exit Sub
End If
If CheckPassword = True Then
    lEvents.eRegistered = True
Else
    frmNag.Show 1
    If lCancel = True Then
        Unload Me
        Exit Sub
    End If
End If
If lEvents.eSettings.iShowAbout = True Then
    'tmrDots.Enabled = True
    chkAbout.Value = 1
    SetAboutShape
    DoEvents
    WindowSize wLoading, Me: DoEvents
    Me.Visible = True
    Me.Height = 5430
    AlwaysOnTop frmAbout, True
    pause 1
    LoadSettings
Else
    LoadSettings
    DoEvents
End If
If Err.Number <> 0 Then SetError "frmAbout_Load()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 Then FormDrag Me
If Err.Number <> 0 Then SetError "frmAbout_MouseDown()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Local Error Resume Next
tmrUnload.Enabled = False
tmrDots.Enabled = False
If chkAbout.Value = 0 Then
    WriteINI lIniFiles.iSettings, "Settings", "ShowAbout", False
Else
    WriteINI lIniFiles.iSettings, "Settings", "ShowAbout", True
End If
If Err.Number <> 0 Then SetError "Form_Unload()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub tmrDots_Timer()
On Local Error Resume Next
jk = jk + 1
Select Case jk
Case 1
    LockWindowUpdate Me.hWnd
    one.BackColor = vbBlack
    two.BackColor = vbWhite
    three.BackColor = vbWhite
    LockWindowUpdate 0
Case 2
    LockWindowUpdate Me.hWnd
    one.BackColor = vbWhite
    two.BackColor = vbBlack
    three.BackColor = vbWhite
    LockWindowUpdate 0
Case 3
    LockWindowUpdate Me.hWnd
    one.BackColor = vbWhite
    two.BackColor = vbWhite
    three.BackColor = vbBlack
    LockWindowUpdate 0
    jk = 0
End Select
If Err.Number <> 0 Then SetError "tmrDots_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub tmrUnload_Timer()
On Local Error Resume Next
tmrUnload.Enabled = False
Unload Me
If Err.Number <> 0 Then SetError "tmrUnload_Timer()", lEvents.eSettings.iErrDescription, Err.Description
End Sub
