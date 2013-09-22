VERSION 5.00
Begin VB.Form frmWait 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Please wait ..."
   ClientHeight    =   1020
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4620
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
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   1020
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrUnload 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   4080
      Top             =   0
   End
   Begin VB.CheckBox chkWindowToggle 
      Caption         =   "Always show this window"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.CommandButton cmdRaw 
      Caption         =   "Show Raw Data"
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox txtIncoming 
      Height          =   1575
      Left            =   120
      MousePointer    =   11  'Hourglass
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1080
      Width           =   4455
   End
   Begin VB.Label lblExtended 
      Height          =   255
      Left            =   720
      MousePointer    =   11  'Hourglass
      TabIndex        =   1
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label lblDescription 
      Caption         =   "Please wait"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      MousePointer    =   11  'Hourglass
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      MousePointer    =   11  'Hourglass
      Stretch         =   -1  'True
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRaw_Click()
'On Local Error Resume Next
If cmdRaw.Caption = "Show Raw Data" Then
    cmdRaw.Caption = "Hide Raw Data"
    Me.Height = 4800
    tmrUnload.Enabled = True
Else
    cmdRaw.Caption = "Hide Raw Data"
    Me.Height = 1460
    tmrUnload.Enabled = False
End If
If Err.Number <> 0 Then SetError "cmdRaw_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Load()
'On Local Error Resume Next
Dim b As Boolean
imgIcon.Picture = frmGraphics.imgIcon.Picture
Me.Height = 1460
b = lEvents.eSettings.iFreeDB.cShowDialog
If b = False Then
    Me.Visible = False
Else
    b = ReadINI(lIniFiles.iSettings, "WaitWindow", "ShowRaw", False)
    If b = True Then
        cmdRaw_Click
    End If
End If
If Err.Number <> 0 Then SetError "Form_Load", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Local Error Resume Next
Dim b As Boolean

If chkWindowToggle.Value = 0 Then
    WriteINI lIniFiles.iSettings, "CDDB", "ShowDialog", True
Else
    WriteINI lIniFiles.iSettings, "CDDB", "ShowDialog", False
End If

If Me.Height = 4800 Then
    WriteINI lIniFiles.iSettings, "WaitWindow", "ShowRaw", True
Else
    WriteINI lIniFiles.iSettings, "WaitWindow", "ShowRaw", False
End If
If Err.Number <> 0 Then SetError "Form_Unload", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub tmrUnload_Timer()
Unload Me
End Sub
