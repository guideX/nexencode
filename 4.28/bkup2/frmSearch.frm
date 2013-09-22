VERSION 5.00
Object = "{EE128208-4F73-11D3-83BB-C47C02EE3D01}#1.0#0"; "ControlResizer.ocx"
Begin VB.Form frmSearch 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "NexENCODE - Search"
   ClientHeight    =   2880
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4095
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.TextBox txtSearch 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   3855
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
      TabIndex        =   1
      Top             =   2400
      Width           =   6015
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   315
         Left            =   2160
         TabIndex        =   5
         Top             =   120
         Width           =   900
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   315
         Left            =   3120
         TabIndex        =   4
         Top             =   120
         Width           =   900
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Search HD"
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   1020
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
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSearch.frx":000C
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3735
   End
   Begin ControlResizer.AutoResizer AutoResizer1 
      Height          =   495
      Left            =   4920
      Top             =   720
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
On Local Error Resume Next
Unload Me
End Sub

Private Sub cmdOK_Click()
On Local Error Resume Next
Dim lSearch As String
lSearch = (Replace(txtSearch.Text, " ", "+"))
Surf "http://audiogalaxy.com/list/searches.php?SID=324da677fda00b2f197cc15b87ca2dee&searchType=0&searchStr=" & lSearch
Unload Me
If Err.Number <> 0 Then SetError "cmdOK_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdPlaylistSeek_Click()
On Local Error Resume Next
frmSearchPlaylists.Show
Unload Me
End Sub

Private Sub Command1_Click()
On Local Error Resume Next
frmSearchForMedia.Show
Unload Me
End Sub

Private Sub Form_Load()
On Local Error Resume Next
'Image1.Picture = frmGraphics.imgTopper.Picture
'Image2.Picture = frmGraphics.imgSideGradient.Picture
FlashIN Me
If Err.Number <> 0 Then SetError "Form_Unload", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Local Error Resume Next
FlashOut Me
If Err.Number <> 0 Then SetError "Form_Unload", lEvents.eSettings.iErrDescription, Err.Description
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

Private Sub Image3_DblClick()
On Local Error Resume Next
Surf "http://audiogalaxy.com"
If Err.Number <> 0 Then SetError "Surf_AudioGalaxy", lEvents.eSettings.iErrDescription, Err.Description
End Sub
