VERSION 5.00
Object = "{EE128208-4F73-11D3-83BB-C47C02EE3D01}#1.0#0"; "ControlResizer.ocx"
Begin VB.Form frmReport 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "NexENCODE - Report"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5670
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   217
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   378
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.ListBox lstReport 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   2580
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5475
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   5775
      TabIndex        =   1
      Top             =   2760
      Width           =   5775
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Default         =   -1  'True
         Height          =   315
         Left            =   4320
         TabIndex        =   3
         Top             =   120
         Width           =   1215
      End
      Begin VB.ComboBox cboOption 
         Height          =   315
         ItemData        =   "frmReport.frx":000C
         Left            =   120
         List            =   "frmReport.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   120
         Width           =   3255
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
      Height          =   255
      Left            =   3240
      Top             =   0
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboOption_Click()
'On Local Error Resume Next
Select Case cboOption.ListIndex
Case 0
    cmdClose.Caption = "Close"
Case 1
    cmdClose.Caption = "Close"
Case 2
    cmdClose.Caption = "End"
Case 3
    cmdClose.Caption = "Play"
End Select
If Err.Number <> 0 Then SetError "cboOption_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdClose_Click()
'On Local Error Resume Next
Dim i As Integer
Select Case cboOption.ListIndex
Case 0
    ClearReports
    WriteINI lIniFiles.iSettings, "Settings", "ShowEvents", "True"
    lEvents.eSettings.iShowReports = True
    PlayWav App.Path & "\media\done.wav", SND_ASYNC
    Unload Me
Case 1
    ClearReports
    WriteINI lIniFiles.iSettings, "Settings", "ShowEvents", "False"
    lEvents.eSettings.iShowReports = False
    PlayWav App.Path & "\media\done.wav", SND_ASYNC
    Unload Me
Case 2
    PlayWav App.Path & "\media\done.wav", SND_ASYNC
    End
Case 3
    For i = 1 To lReports.rCount
        If lReports.rReport(i).rType = Encode Then
            If DoesFileExist(lReports.rReport(i).rFilepath & lReports.rReport(i).rFilename) = True Then
                AddEvent Play, lReports.rReport(i).rFilepath, lReports.rReport(i).rFilename, "", "", 0, ""
            End If
        End If
    Next i
    ClearReports
    Unload Me
End Select
If Err.Number <> 0 Then SetError "cmdClose_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Load()
'On Local Error Resume Next
Dim i As Integer
If lReports.rCount = 0 Then
    Exit Sub
Else
'    Image1.Picture = frmGraphics.imgTopper.Picture
'    Image2.Picture = frmGraphics.imgSideGradient.Picture
    Icon = frmMain.Icon
    cboOption.ListIndex = 3
    For i = 1 To lReports.rCount
        lstReport.AddItem lReports.rReport(i).rReportString
    Next i
    FlashIN Me
    'ClearReports
End If
If Err.Number <> 0 Then SetError "Form_Load()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Local Error Resume Next
FlashOut Me
If cboOption.ListIndex = 1 Then
    WriteINI lIniFiles.iSettings, "Settings", "ShowReports", "False"
    lEvents.eSettings.iShowReports = False
End If
If Err.Number <> 0 Then SetError "Form_Unload()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Image1_DblClick()
If Me.WindowState = vbMaximized Then
    Me.WindowState = vbNormal
Else
    Me.WindowState = vbMaximized
End If

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub
