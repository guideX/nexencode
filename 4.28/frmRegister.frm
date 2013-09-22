VERSION 5.00
Object = "{EE128208-4F73-11D3-83BB-C47C02EE3D01}#1.0#0"; "ControlResizer.ocx"
Begin VB.Form frmRegister 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "NexENCODE - Register"
   ClientHeight    =   3495
   ClientLeft      =   2025
   ClientTop       =   1650
   ClientWidth     =   3900
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRegister.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   233
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Top             =   2640
      Width           =   2775
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1080
      TabIndex        =   5
      Top             =   2280
      Width           =   2775
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
      Top             =   3000
      Width           =   6015
      Begin VB.CommandButton cmdUnlock 
         Caption         =   "Unlock"
         Height          =   315
         Left            =   1560
         TabIndex        =   9
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   315
         Left            =   2760
         TabIndex        =   8
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "!"
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   4
         X1              =   6120
         X2              =   0
         Y1              =   0
         Y2              =   0
      End
   End
   Begin ControlResizer.AutoResizer AutoResizer1 
      Height          =   495
      Left            =   2640
      Top             =   0
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      X1              =   0
      X2              =   272
      Y1              =   58
      Y2              =   58
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Name: "
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "If you have a registration code, please enter it in here..."
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   3735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   0
      X2              =   272
      Y1              =   56
      Y2              =   56
   End
   Begin VB.Label lbladdress 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "If you wish to register NexENCODE Studio, please send $20 check cash or money order with e-mail address and full name to..."
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lShift As Boolean

Private Sub cmdCancel_Click()
'On Local Error Resume Next
Unload Me
End Sub

Private Sub cmdUnlock_Click()
'On Local Error Resume Next
If Len(txtName.Text) <> 0 And Len(txtPassword.Text) <> 0 And Len(txtName.Text) > 5 Then
    lEvents.eName = txtName.Text
    lEvents.ePassword = txtPassword.Text
    If CheckPassword() = True Then
        lEvents.eRegistered = True
        WriteINI lIniFiles.iSettings, "Settings", "Name", txtName.Text
        WriteINI lIniFiles.iSettings, "Settings", "Password", txtPassword.Text
        If lEvents.eSettings.iOverwritePrompts = True Then MsgBox "Thanks for registering!", vbInformation
        Unload Me
    Else
        If lEvents.eSettings.iOverwritePrompts = True Then MsgBox "Sorry, the information you provided is not valid. If you feel this in error, contact guide_X@live.com with your name and password and the situation will be resolved", vbCritical
        Beep
    End If
Else
    Beep
    Exit Sub
End If
If Err.Number <> 0 Then SetError "cmdUnlock_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Command2_Click()
'On Local Error Resume Next
txtPassword.Text = Crypt(txtName.Text, "pickles", False)
If Err.Number <> 0 Then SetError "Command2_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Load()
'On Local Error Resume Next

'AlwaysOnTop Me, True
'If DoesFileExist(App.Path & "\skins\inex\inex_top.gif") = True Then
'    Image1.Picture = LoadPicture(App.Path & "\skins\inex\inex_top.gif")
'Else
'    Image1.Picture = frmGraphics.imgTopper.Picture
'End If
'If DoesFileExist(App.Path & "\skins\inex\inex_sg.gif") = True Then
'    Image2.Picture = LoadPicture(App.Path & "\skins\inex\inex_sg.gif")
'Else
'    Image2.Picture = frmGraphics.imgTopper.Picture
'End If
lbladdress.Caption = "Leon Aiossa" & vbCrLf & "1056 churchill street #1" & vbCrLf & "St. Paul, Minnesota, 55103"
FlashIN Me
If Err.Number <> 0 Then SetError "Form_Load()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Image1_DblClick()
'On Local Error Resume Next
If Me.WindowState = vbMaximized Then
    Me.WindowState = vbNormal
Else
    Me.WindowState = vbMaximized
End If
If Err.Number <> 0 Then SetError "Image1_DblClikc", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
'On Local Error Resume Next
If Shift = 1 Then lShift = True
If Err.Number <> 0 Then SetError "txtName_KeyDown()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
'On Local Error Resume Next
If KeyAscii = 43 And txtName.Text = "picklespipe" And lShift = True Then
    Command2.Visible = True
End If
If Err.Number <> 0 Then SetError "txtName_KeyPress()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub txtName_KeyUp(KeyCode As Integer, Shift As Integer)
'On Local Error Resume Next
lShift = False
If Err.Number <> 0 Then SetError "txtName_KeyUp()", lEvents.eSettings.iErrDescription, Err.Description
End Sub
