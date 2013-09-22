VERSION 5.00
Object = "{EE128208-4F73-11D3-83BB-C47C02EE3D01}#1.0#0"; "ControlResizer.ocx"
Begin VB.Form frmTrackGet 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "NexENCODE - Track Get"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   270
   ClientWidth     =   5595
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
   ScaleHeight     =   248
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   373
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   " "
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   6015
      TabIndex        =   18
      Top             =   3240
      Width           =   6015
      Begin VB.ComboBox cboFormat 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   315
         Left            =   4680
         TabIndex        =   20
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         Default         =   -1  'True
         Height          =   315
         Left            =   3720
         TabIndex        =   19
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
   Begin VB.Timer tmrPreviewCD 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   5160
      Top             =   0
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "&Preview"
      Height          =   255
      Left            =   2160
      TabIndex        =   17
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "< Remove"
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add >"
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdALL 
      Caption         =   "All >"
      Height          =   255
      Left            =   2160
      TabIndex        =   16
      Top             =   1920
      Width           =   1335
   End
   Begin VB.ComboBox cboCDDrive 
      Height          =   315
      ItemData        =   "frmTrackGet.frx":0000
      Left            =   2160
      List            =   "frmTrackGet.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1560
      Width           =   1335
   End
   Begin VB.ListBox lstQue 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   1740
      IntegralHeight  =   0   'False
      Left            =   3600
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   1440
      Width           =   1935
   End
   Begin VB.ListBox lstAvailableTracks 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   1740
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox txtAlbum 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   230
      Left            =   840
      MousePointer    =   1  'Arrow
      TabIndex        =   3
      Top             =   360
      Width           =   4695
   End
   Begin VB.TextBox txtArtist 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   230
      Left            =   840
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Top             =   120
      Width           =   4695
   End
   Begin ControlResizer.AutoResizer AutoResizer1 
      Height          =   375
      Left            =   4680
      Top             =   0
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin VB.Label lblEstimatedSize 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   15
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label lblTrackTime 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   14
      Top             =   720
      Width           =   2535
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      X1              =   190
      X2              =   190
      Y1              =   48
      Y2              =   88
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   192
      X2              =   192
      Y1              =   48
      Y2              =   88
   End
   Begin VB.Label lblPlay 
      BackStyle       =   0  'Transparent
      Caption         =   "Play in CD Player"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   720
      TabIndex        =   13
      Top             =   960
      Width           =   2055
   End
   Begin VB.Image imgDisc 
      Height          =   480
      Left            =   120
      Stretch         =   -1  'True
      Top             =   720
      Width           =   480
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "CD Audio"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   720
      TabIndex        =   12
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "CD Drive:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   1350
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Que:"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   3600
      TabIndex        =   6
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Available Tracks:"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00400040&
      BackStyle       =   0  'Transparent
      Caption         =   "Album:"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblArtist 
      BackColor       =   &H00400040&
      BackStyle       =   0  'Transparent
      Caption         =   "Artist:"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmTrackGet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lReady As Boolean

Public Sub FillWithCdContents()
'On Local Error Resume Next
Dim Z As Integer, i As Integer, X As Integer, lToc As String

lReady = False
DoEvents
txtArtist.Text = ReturnDirCompliant(lTracks.tArtist)
txtAlbum.Text = ReturnDirCompliant(lTracks.tTitle)
cboFormat.AddItem "To Mp3"
cboFormat.AddItem "To Wav"
cboFormat.ListIndex = ReadINI(lIniFiles.iSettings, "Settings", "CopyFormat", 0)
lblInfo.Caption = "Length: " & lTracks.tDiscLen & " / " & lTracks.tGenre
With frmMain
    cboCDDrive.Clear
    FillComboWithDrives cboCDDrive
    cboCDDrive.ListIndex = FindComoboxIndex(cboCDDrive, lRipperSettings.eDriveLetter)
    lstAvailableTracks.Clear
    .Ripper.Init: DoEvents
    .Ripper.OpenDriveByLetter lRipperSettings.eDriveLetter
    
    For i = 1 To lTracks.tCount
        If .Ripper.TrackIsAudio(i) = True Then
            If Len(lTracks.tTrack(i).tName) <> 0 Then
                lstAvailableTracks.AddItem i & ": " & ReturnDirCompliant(lTracks.tTrack(i).tName)
            Else
                lstAvailableTracks.AddItem i & ": Track " & i
            End If
        End If
    Next i
End With
lReady = True
If Err.Number <> 0 Then SetError "FillWithCDContents()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cboCDDrive_Click()
'On Local Error Resume Next
Dim i As Integer, msg As String

If lReady = False Then Exit Sub
If Int(cboCDDrive.Text) <> lRipperSettings.eDriveLetter Then
    lRipperSettings.eDriveLetter = cboCDDrive.Text
    GetSimpleTracks
    Unload Me
    Load Me
    FillWithCdContents
End If
If Err.Number <> 0 Then SetError "cboCdDrive_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdAdd_Click()
'On Local Error Resume Next
Dim msg As String

If Len(lstAvailableTracks.Text) <> 0 Then
    msg = lstAvailableTracks.Text
    lstAvailableTracks.RemoveItem lstAvailableTracks.ListIndex
    lstQue.AddItem msg
End If
If Err.Number <> 0 Then SetError "cmdAdd_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdALL_Click()
'On Local Error Resume Next
Dim i As Integer, msg As String

For i = 0 To lstAvailableTracks.ListCount
     If lstAvailableTracks.ListCount <> 0 Then
        lstAvailableTracks.ListIndex = 0
        msg = lstAvailableTracks.Text
        If Len(msg) <> 0 Then
            lstAvailableTracks.RemoveItem 0
            lstQue.AddItem msg
            DoEvents
        End If
    Else
        Exit Sub
    End If
Next i

If Err.Number <> 0 Then SetError "cmdAll_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdCancel_Click()
'On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then SetError "cmdCancel_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdCopy_Click()
'On Local Error Resume Next
Dim i As Integer, msg As String, lPath As String, lMp3 As String, lWav As String, f As Integer, p As Integer, lefty As String
Dim msg2 As String

If Len(txtAlbum.Text) = 0 Or Len(txtArtist.Text) = 0 Then
    If lEvents.eSettings.iOverwritePrompts = True Then MsgBox "Please enter a value in the Artist/Album feild", vbQuestion
    If Len(txtAlbum.Text) = 0 Then
        txtAlbum.SetFocus
        Beep
    End If
    If Len(txtArtist.Text) = 0 Then
        txtArtist.SetFocus
        Beep
    End If
    Exit Sub
ElseIf lstQue.ListCount = 0 Then
    If lEvents.eSettings.iOverwritePrompts = True Then MsgBox "No tracks selected, aborting rip process", vbExclamation
    Exit Sub
End If

Me.Visible = False

Select Case cboFormat.ListIndex
Case 0
    For i = 1 To lstQue.ListCount
        If Len(lstQue.List(i - 1)) <> 0 Then
            msg = lstQue.List(i - 1)
            lefty = Left(msg, 1)
            f = Int(Trim(lefty & ParseString(msg, Left(msg, 1), ":")))
            If Len(lTracks.tTrack(f).tName) <> 0 Then
                lWav = "(" & ReturnDirCompliant(txtArtist.Text) & ") " & lTracks.tTrack(f).tName & ".wav"
                lMp3 = "(" & ReturnDirCompliant(txtArtist.Text) & ") " & lTracks.tTrack(f).tName & ".mp3"
            Else
                lWav = "(" & ReturnDirCompliant(txtArtist.Text & " - " & txtAlbum.Text) & ") track " & f & ".wav"
                lMp3 = "(" & ReturnDirCompliant(txtArtist.Text & " - " & txtAlbum.Text) & ") track " & f & ".mp3"
            End If
            lPath = lEncoderSettings.eOutputDir & txtArtist.Text & "-" & txtAlbum.Text & "\"
            If f <> 0 Then
                AddEvent Rip, "", "", lPath, lWav, f, ""
                AddEvent Encode, lPath, lWav, lPath, lMp3, 0, "AUTODELETE"
                p = AddPlaylist(txtArtist.Text & "-" & txtAlbum.Text)
                Playlist.pIndex = p
                DoEvents
                AddToPlaylist lPath & lMp3, p
            End If
        End If
    Next i
    If lEvents.eSettings.iCreateAlbumFileOnEncode = True Then AddEvent Merge, "", "", lPath, "(" & txtArtist.Text & ") " & txtAlbum.Text & " [ Whole Album ].mp3", 0, ""
Case 1
    For i = 1 To lstQue.ListCount
        If Len(lstQue.List(i - 1)) <> 0 Then
            msg = lstQue.List(i - 1)
            lefty = Left(msg, 1)
            f = Int(Trim(lefty & ParseString(msg, Left(msg, 1), ":")))
            If Len(lTracks.tTrack(f).tName) <> 0 Then
                lWav = "(" & ReturnDirCompliant(txtArtist.Text) & ") " & lTracks.tTrack(f).tName & ".wav"
            Else
                lWav = "(" & ReturnDirCompliant(txtArtist.Text & " - " & txtAlbum.Text) & ") " & f & ".wav"
            End If
            lPath = lEncoderSettings.eOutputDir & txtArtist.Text & "-" & txtAlbum.Text & "\"
            If f <> 0 Then
                AddEvent Rip, "", "", lPath, lWav, f, ""
            End If
        End If
    Next i
End Select

ProcessNextEvent
DoEvents
If Err.Number <> 0 Then SetError "cmdCopy_Click()", lEvents.eSettings.iErrDescription, Err.Description
Unload Me
End Sub

Private Sub cmdPreview_Click()
'On Local Error Resume Next
Dim msg As String, i As Integer
If cmdPreview.Caption = "&Preview" Then
    i = lstAvailableTracks.ListIndex + 1
    If i <> 0 Then
        frmMain.Ripper.Play i, 0, 10
        cmdPreview.Caption = "&Stop"
    End If
    tmrPreviewCD.Enabled = True
ElseIf cmdPreview.Caption = "&Stop" Then
    cmdPreview.Caption = "&Preview"
    frmMain.Ripper.StopPlaying
    tmrPreviewCD.Enabled = False
End If
If Err.Number <> 0 Then SetError "cmdPreview()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdRemove_Click()
'On Local Error Resume Next
Dim msg As String

If Len(lstQue.Text) <> 0 Then
    msg = lstQue.Text
    lstQue.RemoveItem lstQue.ListIndex
    lstAvailableTracks.AddItem msg
End If
If Err.Number <> 0 Then SetError "cmdRemove_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Load()
'On Local Error Resume Next
'imgNS4.Picture = frmGraphics.imgIcon.Picture
imgDisc.Picture = frmGraphics.imgDisc.Picture
'Image1.Picture = frmGraphics.imgTopper.Picture
'Image2.Picture = frmGraphics.imgSideGradient.Picture
Icon = frmMain.Icon
FillWithCdContents
FlashIN frmTrackGet
If Err.Number <> 0 Then SetError "Form_Load()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Local Error Resume Next

FlashOut frmTrackGet
WriteINI lIniFiles.iSettings, "Settings", "CopyFormat", cboFormat.ListIndex
WriteINI lIniFiles.iSettings, "Settings", "LastArtist", txtArtist.Text
WriteINI lIniFiles.iSettings, "Settings", "LastAlbum", txtAlbum.Text

If Err.Number <> 0 Then SetError "Form_Unload()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Image1_Click()
'On Local Error Resume Next
If Me.WindowState = vbMaximized Then
    Me.WindowState = vbNormal
Else
    Me.WindowState = vbMaximized
End If
If Err.Number <> 0 Then SetError "imgTopper_DblClick()", lEvents.eSettings.iErrDescription, Err.Description
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

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu frmMain.mnuHidden
End If
End Sub

Private Sub lblPlay_Click()
'On Local Error Resume Next
If lPlayers.pCDPlayerIndex <> 0 Then GoCDPlayer
If Err.Number <> 0 Then SetError "lblPlay_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub lstAvailableTracks_Click()
'On Local Error Resume Next
Dim i As Integer, msg As String, f As Long

i = Int(Left(lstAvailableTracks.Text, 1) & ParseString(lstAvailableTracks.Text, Left(lstAvailableTracks.Text, 1), ":"))

If i <> 0 And i < lTracks.tCount + 1 Then
    lblTrackTime.Caption = "Length: " & Format(lTracks.tTrack(i).tLength, "@.@@ Minutes / Seconds")
    lblEstimatedSize.Caption = "Milleseconds: " & frmMain.Ripper.GetTrackLengthMs(i)

Else
    lblTrackTime.Caption = ""
    lblEstimatedSize.Caption = ""
End If


If Err.Number <> 0 Then SetError "lstAvailableTracks_DblClick()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub lstAvailableTracks_DblClick()
'On Local Error Resume Next

lblTrackTime.Caption = ""
lblEstimatedSize.Caption = ""
cmdAdd_Click

If Err.Number <> 0 Then SetError "lstAvailableTracks_DblClick()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub lstQue_DblClick()
'On Local Error Resume Next
cmdRemove_Click
If Err.Number <> 0 Then SetError "lstQue_DblClick()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub tmrPreviewCD_Timer()
'On Local Error Resume Next
cmdPreview.Caption = "&Preview"
lEvents.eRipperBusy = False
tmrPreviewCD.Enabled = False
If Err.Number <> 0 Then SetError "tmrPreviewCD()", lEvents.eSettings.iErrDescription, Err.Description
End Sub
