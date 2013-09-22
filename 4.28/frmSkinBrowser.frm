VERSION 5.00
Object = "{EE128208-4F73-11D3-83BB-C47C02EE3D01}#1.0#0"; "ControlResizer.ocx"
Begin VB.Form frmSkinBrowser 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "NexENCODE - Skin Browser"
   ClientHeight    =   2760
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   5730
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSkinBrowser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   184
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   382
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.ListBox lstSkins 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1785
      IntegralHeight  =   0   'False
      ItemData        =   "frmSkinBrowser.frx":08CA
      Left            =   1800
      List            =   "frmSkinBrowser.frx":08CC
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "&Edit"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   6015
      TabIndex        =   4
      Top             =   2280
      Width           =   6015
      Begin VB.CommandButton cmdOK 
         Cancel          =   -1  'True
         Caption         =   "Close"
         Default         =   -1  'True
         Height          =   315
         Left            =   4560
         TabIndex        =   7
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
      Begin VB.Label lblAuthor 
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   4455
      End
   End
   Begin ControlResizer.AutoResizer AutoResizer1 
      Height          =   735
      Left            =   2280
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1296
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select a skin below"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmSkinBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
'On Local Error Resume Next
Me.Visible = False
frmSkinEditor.Show
NewSkin
If Err.Number <> 0 Then SetError "cmdAdd_click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdApply_Click()
'On Local Error Resume Next
Dim i As Integer
If Len(lstSkins.Text) <> 0 Then
    i = FindSkinIndex(lstSkins.Text)
    If i <> 0 Then ApplySkin frmMain, i
End If
If Err.Number <> 0 Then SetError "cmdApply_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdModify_Click()
'On Local Error Resume Next
Dim i As Integer
Me.Visible = False
frmSkinEditor.Show
SetSkin lstSkins.ListIndex + 1
Unload Me
If Err.Number <> 0 Then SetError "cmdModify_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdOK_Click()
'On Local Error Resume Next
PlayWav App.Path & "\media\done.wav", SND_ASYNC
Unload Me
End Sub

Private Sub Form_Load()
'On Local Error Resume Next
'imgNS4.Picture = frmGraphics.imgIcon.Picture
'Image1.Picture = frmGraphics.imgTopper.Picture
'Image2.Picture = frmGraphics.imgSideGradient.Picture
Dim i As Integer
For i = 1 To lSkins.sCount
    lstSkins.AddItem lSkins.sSkin(i).sName
Next i
'FlashIN Me
If Err.Number <> 0 Then SetError "frmSkinBrowser_Load()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Local Error Resume Next
FlashOut Me
If Err.Number <> 0 Then SetError "Form_Unload()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub lstSkins_Click()
'On Local Error Resume Next
Dim i As Integer
If Len(lstSkins.Text) <> 0 Then
    cmdModify.Enabled = True
    i = FindSkinIndex(lstSkins.Text)
    If i <> 0 Then
        lblAuthor.Caption = "Author: " & lSkins.sSkin(i).sAuthor
    Else
        lblAuthor.Caption = ""
    End If
Else
    cmdModify.Enabled = False
End If
If Err.Number <> 0 Then SetError "lstSkins_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub lstSkins_DblClick()
'On Local Error Resume Next
cmdApply_Click
If Err.Number <> 0 Then SetError "lstSkins_DblClick()", lEvents.eSettings.iErrDescription, Err.Description
End Sub
