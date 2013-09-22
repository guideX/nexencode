VERSION 5.00
Object = "{EE128208-4F73-11D3-83BB-C47C02EE3D01}#1.0#0"; "ControlResizer.ocx"
Begin VB.Form frmEditTracks 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "NexENCODE - Edit Tracks"
   ClientHeight    =   2760
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4905
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
   ScaleHeight     =   2760
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   720
      TabIndex        =   9
      ToolTipText     =   "Album title (Example: Experience Expanded Disc 2)"
      Top             =   480
      Width           =   4095
   End
   Begin VB.TextBox txtArtist 
      Height          =   285
      Left            =   720
      TabIndex        =   7
      ToolTipText     =   "Artist or band name (Example: The Prodigy)"
      Top             =   120
      Width           =   4095
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   ">>"
      Height          =   255
      Left            =   4080
      TabIndex        =   5
      ToolTipText     =   "Click to change song title"
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox txtTrack 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Enter track name here"
      Top             =   840
      Width           =   3855
   End
   Begin VB.ListBox lstTracks 
      Height          =   1035
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "List of Audio Tracks"
      Top             =   1200
      Width           =   4695
   End
   Begin VB.PictureBox picBottom 
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
      Begin VB.ComboBox cboAutoSubmit 
         Height          =   315
         ItemData        =   "frmEditTracks.frx":0000
         Left            =   120
         List            =   "frmEditTracks.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   "Setting for what to do next time"
         Top             =   120
         Width           =   2055
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Default         =   -1  'True
         Height          =   315
         Left            =   3000
         TabIndex        =   2
         ToolTipText     =   "Save Track list and exit"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   315
         Left            =   3960
         TabIndex        =   1
         ToolTipText     =   "Cancel/Hide Window"
         Top             =   120
         Width           =   855
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   4
         X1              =   0
         X2              =   6000
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Title:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Artist:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
   Begin ControlResizer.AutoResizer AutoResizer1 
      Height          =   615
      Left            =   5280
      Top             =   0
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   1085
   End
End
Attribute VB_Name = "frmEditTracks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
'On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then SetError "cmdCancel_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdChange_Click()
'On Local Error Resume Next
Dim i As Integer
If Len(txtTrack.Text) <> 0 Then
    If i <> -1 Then
        i = lstTracks.ListIndex
        lstTracks.RemoveItem i
        lstTracks.AddItem txtTrack.Text, i
    End If
End If
If Err.Number <> 0 Then SetError "cmdChange_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdSave_Click()
'On Local Error Resume Next
Dim i As Integer
If Len(txtArtist.Text) <> 0 Then lTracks.tArtist = txtArtist.Text
If Len(txtTitle.Text) <> 0 Then lTracks.tTitle = txtTitle.Text
For i = 1 To lTracks.tCount
    lTracks.tTrack(i).tName = lstTracks.List(i - 1)
Next i
SaveCDTracks lRipperSettings.eDiscID
Unload Me
If Err.Number <> 0 Then SetError "cmdSave_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Load()
'On Local Error Resume Next
cboAutoSubmit.ListIndex = 0
If Len(lRipperSettings.eDriveLetter) <> 0 Then
    FlashIN Me
    txtArtist.Text = lTracks.tArtist
    txtTitle.Text = lTracks.tTitle
    Dim i As Integer
    For i = 1 To lTracks.tCount
        If Len(lTracks.tTrack(i).tName) <> 0 Then
            lstTracks.AddItem lTracks.tTrack(i).tName
        Else
            lstTracks.AddItem "Track " & i
        End If
    Next i
Else
    If lEvents.eSettings.iOverwritePrompts = True Then MsgBox "No Disc"
End If
If Err.Number <> 0 Then SetError "Form_Load()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Local Error Resume Next
FlashOut Me
If Err.Number <> 0 Then SetError "Form_Unload()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub lstTracks_Click()
'On Local Error Resume Next
txtTrack.Text = lstTracks.Text
txtTrack.SetFocus
If Err.Number <> 0 Then SetError "lstTracks_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub txtTrack_GotFocus()
'On Local Error Resume Next
txtTrack.SelLength = Len(txtTrack.Text)
If Err.Number <> 0 Then SetError "txtTrack_GotFOcus()", lEvents.eSettings.iErrDescription, Err.Description
End Sub
