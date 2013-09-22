VERSION 5.00
Begin VB.Form frmPlaylist 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Playlist Editor"
   ClientHeight    =   4050
   ClientLeft      =   4635
   ClientTop       =   3210
   ClientWidth     =   4485
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPlaylistEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   270
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   299
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.ComboBox cboPlaylists 
      Appearance      =   0  'Flat
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
      Height          =   315
      ItemData        =   "frmPlaylistEditor.frx":000C
      Left            =   1680
      List            =   "frmPlaylistEditor.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Current playlist"
      Top             =   240
      Width           =   1815
   End
   Begin VB.ListBox lstPlaylist 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2820
      IntegralHeight  =   0   'False
      Left            =   225
      MouseIcon       =   "frmPlaylistEditor.frx":0010
      MousePointer    =   99  'Custom
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "List of Files"
      Top             =   675
      Visible         =   0   'False
      Width           =   3285
   End
   Begin VB.Image imgLayout 
      Enabled         =   0   'False
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   855
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   3480
      Width           =   3255
   End
   Begin VB.Image imgExit2 
      Height          =   240
      Left            =   2160
      Picture         =   "frmPlaylistEditor.frx":0162
      Top             =   2880
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Image imgExit1 
      Height          =   240
      Left            =   2160
      Picture         =   "frmPlaylistEditor.frx":0A24
      Top             =   2880
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Image imgExit 
      Height          =   240
      Left            =   3750
      Picture         =   "frmPlaylistEditor.frx":12E6
      ToolTipText     =   "Exit / Hide this window"
      Top             =   3360
      Width           =   675
   End
   Begin VB.Image imgDelete1 
      Height          =   240
      Left            =   1800
      Picture         =   "frmPlaylistEditor.frx":1BA8
      Top             =   2520
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Image imgDelete 
      Height          =   240
      Left            =   3750
      Picture         =   "frmPlaylistEditor.frx":246A
      ToolTipText     =   "Delete Menu"
      Top             =   1950
      Width           =   675
   End
   Begin VB.Image imgAdd1 
      Height          =   240
      Left            =   2640
      Picture         =   "frmPlaylistEditor.frx":2D2C
      Top             =   2760
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Image imgAdd 
      Height          =   240
      Left            =   3750
      Picture         =   "frmPlaylistEditor.frx":35EE
      ToolTipText     =   "Add Menu"
      Top             =   1320
      Width           =   675
   End
   Begin VB.Image imgPlay1 
      Height          =   240
      Left            =   1440
      Picture         =   "frmPlaylistEditor.frx":3EB0
      Top             =   2880
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Image imgPlay2 
      Height          =   240
      Left            =   1440
      Picture         =   "frmPlaylistEditor.frx":4772
      Top             =   2880
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Image imgPlay 
      Height          =   240
      Left            =   3750
      Picture         =   "frmPlaylistEditor.frx":5034
      ToolTipText     =   "Play menu"
      Top             =   720
      Width           =   675
   End
   Begin VB.Image imgAdd2 
      Height          =   240
      Left            =   2640
      Picture         =   "frmPlaylistEditor.frx":58F6
      Top             =   2760
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Image imgDelete2 
      Height          =   240
      Left            =   1800
      Picture         =   "frmPlaylistEditor.frx":61B8
      Top             =   2520
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Menu mnuListbox 
      Caption         =   "Hidden"
      Begin VB.Menu mnuPlayListbox 
         Caption         =   "Play"
      End
      Begin VB.Menu mnuDecodeListbox 
         Caption         =   "Decode"
      End
      Begin VB.Menu mnuMerge 
         Caption         =   "Merge"
      End
      Begin VB.Menu mnuSep9749734 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMovetoPlaylist 
         Caption         =   "Move (to recient)"
      End
      Begin VB.Menu mnuRemoveFromPlaylist 
         Caption         =   "Delete (from playlist)"
      End
      Begin VB.Menu mnuDeleteFromDisc 
         Caption         =   "Delete (from disc)"
      End
      Begin VB.Menu mnuSep93792379 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMoreInfoListbox 
         Caption         =   "Properties"
      End
   End
   Begin VB.Menu mnuPlay 
      Caption         =   "Hidden"
      Begin VB.Menu mnuPlayFiles 
         Caption         =   "Play"
      End
      Begin VB.Menu mnuPlayPlaylist 
         Caption         =   "Playlist"
      End
      Begin VB.Menu mnuPlayHTMLPlaylist 
         Caption         =   "Web Playlist"
      End
   End
   Begin VB.Menu mnuAdd 
      Caption         =   "Hidden"
      Begin VB.Menu mnuAddFiles 
         Caption         =   "Add File"
      End
      Begin VB.Menu mnuAddDirectory 
         Caption         =   "Add Directory"
      End
      Begin VB.Menu mnuSep938729732 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearchForFiles 
         Caption         =   "Search For Files"
      End
   End
   Begin VB.Menu mnuDel 
      Caption         =   "Hidden"
      Begin VB.Menu MNUFROMLIBRARY 
         Caption         =   "From Library"
      End
      Begin VB.Menu mnuDelThisPlaylist 
         Caption         =   "Playlist"
      End
      Begin VB.Menu mnuDelPlaylist 
         Caption         =   "All Playlists"
      End
   End
End
Attribute VB_Name = "frmPlaylist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub PlaylistKeypress(lKeyascii As Integer)
On Local Error Resume Next
If lKeyascii = 27 Then
    'frmMain.Visible = False
    'frmMain.WindowState = vbMinimized
    Unload Me
End If
If Err.Number <> 0 Then SetError "PlaylistKeyPress()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub SortByGenre(lPlaylist As Integer)
On Local Error Resume Next
Dim i As Integer, e As Integer, n As Integer, lGenre As String, lFile As String, lPath As String
For i = 1 To Playlist.pFileCount
    If Playlist.pFiles(i).fPlaylist = lPlaylist Then
        lFile = Playlist.pFiles(i).fFile
        lPath = Playlist.pFiles(i).fPath
        If Right(lPath, 1) <> "\" Then lPath = lPath & "\"
        lTag.tFile = lPath & lFile
        GetTagInfo
        DoEvents
        lGenre = frmMP3Info.cmbGenre.Text
        If IsNumeric(lTag.tGenre) = True And Len(lGenre) <> 0 And Len(lFile) <> 0 And Len(lPath) <> 0 And lTag.tHasTag = True Then
            e = FindPlaylistIndex(lGenre)
            If e = 0 Then e = AddPlaylist(lGenre)
            n = FindMediaIndex(lFile)
            Playlist.pFiles(n).fPlaylist = e
        ElseIf IsNumeric(lTag.tGenre) = False Or Len(lGenre) = 0 Or lTag.tHasTag = False Then
            e = FindPlaylistIndex("NexENCODE")
            If e = 0 Then e = AddPlaylist("NexENCODE")
            n = FindMediaIndex(lFile)
            Playlist.pFiles(n).fPlaylist = e
        End If
    End If
Next i
Playlist.pPlaylists(lPlaylist).pEnabled = False
WriteINI lIniFiles.iPlaylists, lPlaylist, "Enabled", "False"
SavePlaylists
LoadPlaylist
If Err.Number <> 0 Then SetError "SortByGenre()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub PromptMediaDir()
On Local Error Resume Next
Dim i As Integer
If Len(frmPlaylist.cboPlaylists.Text) = 0 Or Left(frmPlaylist.cboPlaylists.Text, 1) = "<" Then
    i = AddPlaylist("Temp")
Else
    i = FindPlaylistIndex(frmPlaylist.cboPlaylists.Text)
End If
AddDirToPlaylist i
frmPlaylist.SortByGenre i
frmPlaylist.cboPlaylists.ListIndex = 1
If Err.Number <> 0 Then SetError "PromptMediaDir()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub SavePlaylists()
On Local Error Resume Next
Dim i As Integer, msg As String
For i = 1 To Playlist.pPlaylistCount
    If Playlist.pPlaylists(i).pEnabled = True Then
        SavePlaylist i
        PlaylistToHTMLFile False, i
    End If
Next i
PlaylistToHTMLFile True
If Err.Number <> 0 Then SetError "SortByGenre()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub LoadPlaylist()
On Local Error Resume Next
Dim i As Integer, msg As String
cboPlaylists.Clear
cboPlaylists.AddItem "<Playlists>"
cboPlaylists.AddItem "<All>"
cboPlaylists.AddItem "<New>"
cboPlaylists.AddItem "<Search>"
For i = 1 To Playlist.pPlaylistCount
    If Len(Playlist.pPlaylists(i).pFile) <> 0 Then msg = Left(Playlist.pPlaylists(i).pFile, Len(Playlist.pPlaylists(i).pFile) - 4)
    If Playlist.pPlaylists(i).pEnabled = True Then cboPlaylists.AddItem msg
Next i
DoEvents
If Err.Number <> 0 Then SetError "LoadPlaylist()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub FillWithPlaylist(lPlaylist As String)
On Local Error Resume Next
Dim i As Integer, lIndex As Integer
lIndex = FindPlaylistIndex(lPlaylist)
Playlist.pIndex = lIndex
If Len(lPlaylist) <> 0 And lIndex <> 0 Then
    lstPlaylist.Clear
    For i = 1 To Playlist.pFileCount
        If Playlist.pFiles(i).fPlaylist = lIndex Then
            lstPlaylist.AddItem Playlist.pFiles(i).fFile
        End If
    Next i
End If
If Err.Number <> 0 Then SetError "FillWithPlayer()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cboPlaylists_Click()
On Local Error Resume Next
Dim msg As String, i As Integer
lblInfo.Caption = ""
If cboPlaylists.Text = "<New>" Then
    If lEvents.eRipperBusy = False And lEvents.eEncoderBusy = False Then
        msg = InputBox("Enter name of playlist:")
        If Len(msg) <> 0 Then
            AddPlaylist msg
            cboPlaylists.AddItem msg
            cboPlaylists.ListIndex = FindComoboxIndex(cboPlaylists, msg)
            lstPlaylist.Clear
        Else
            cboPlaylists.ListIndex = 0
        End If
    End If
ElseIf cboPlaylists.Text = "<Playlists>" Then
    lstPlaylist.Clear
    For i = 1 To Playlist.pPlaylistCount
        If Playlist.pPlaylists(i).pEnabled = True Then lstPlaylist.AddItem Playlist.pPlaylists(i).pFile
    Next i
ElseIf cboPlaylists.Text = "<Search>" Then
    If lEvents.eRipperBusy = False And lEvents.eEncoderBusy = False Then
        cboPlaylists.ListIndex = 1
        frmSearchPlaylists.Show
    End If
ElseIf cboPlaylists.Text = "<All>" Then
    lstPlaylist.Clear
    For i = 1 To Playlist.pFileCount
        If Len(Playlist.pFiles(i).fFile) <> 0 Then lstPlaylist.AddItem Playlist.pFiles(i).fFile
    Next i
Else
    msg = cboPlaylists.Text
    i = FindPlaylistIndex(msg)
    FillWithPlaylist cboPlaylists.Text
End If
If Err.Number <> 0 Then SetError "cboPlaylists_Change()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub DisplayPlaylist(lIndex)
On Local Error Resume Next
If Playlist.pPlaylists(lIndex).pEnabled = True Then
    cboPlaylists.ListIndex = FindComoboxIndex(cboPlaylists, Playlist.pPlaylists(lIndex).pDescription)
    cboPlaylists_Click
Else
    cboPlaylists.ListIndex = 1
End If
If Err.Number <> 0 Then SetError "DisplayPlaylist()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cboPlaylists_KeyPress(KeyAscii As Integer)
On Local Error Resume Next
PlaylistKeypress KeyAscii
If Err.Number <> 0 Then SetError "cboPlaylists_KeyPress()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Local Error Resume Next
PlaylistKeypress KeyAscii
If Err.Number <> 0 Then SetError "Form_KeyPress()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Load()
On Local Error Resume Next
Dim rgn As Long, tmp As Long, X As Long, Y As Long
Me.Picture = frmGraphics.imgPlaylist.Picture
lEvents.ePlaylistVisible = True
WriteINI lIniFiles.iSettings, "PlaylistWind", "Visible", "True"
Icon = frmMain.Icon
GetWindowSettings Me.hwnd
X = lMainWndSettings.wWindowBorder
Y = lMainWndSettings.wTitleBarHeight
rgn = CreateRoundRectRgn(X + 0, Y + 0, X + 298, Y + 245, 30, 30)
tmp = SetWindowRgn(Me.hwnd, rgn, True)
frmMain.mnuPlaylist.Checked = True
LoadPlaylist
DoEvents
FlashIN frmPlaylist
lstPlaylist.Visible = True
cboPlaylists.ListIndex = 1
If Playlist.pIndex <> 0 Then
    cboPlaylists.ListIndex = FindComoboxIndex(cboPlaylists, Playlist.pPlaylists(Playlist.pIndex).pDescription)
Else
    Dim msg As String
    msg = ReadINI(lIniFiles.iSettings, "Settings", "CurrentPlaylist", "")
    If Len(msg) <> 0 Then
        cboPlaylists.ListIndex = FindComoboxIndex(cboPlaylists, msg)
    End If
End If
If Err.Number <> 0 Then SetError "Form_Load()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 Then
    FormDrag Me
End If
If Err.Number <> 0 Then SetError "Form_MouseDown()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
DragDrop Data
If Err.Number <> 0 Then SetError "Form_OleDragDrop()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Local Error Resume Next
frmMain.mnuPlaylist.Checked = False
FlashOut frmPlaylist
WriteINI lIniFiles.iSettings, "Settings", "CurrentPlaylist", cboPlaylists.Text
WriteINI lIniFiles.iSettings, "PlaylistWind", "Visible", False
lEvents.ePlaylistVisible = False
If Err.Number <> 0 Then SetError "Form_Unload()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgAdd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 Then
    imgAdd.Picture = imgAdd2.Picture
End If
If Err.Number <> 0 Then SetError "imgAdd_MouseDown()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgAdd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 Then
    PlayWav App.Path & "\media\click2.wav", SND_ASYNC
    PopupMenu mnuAdd
    imgAdd.Picture = imgAdd1.Picture
End If
If Err.Number <> 0 Then SetError "imgAdd_MouseUp()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgDelete_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 Then
    imgDelete.Picture = imgDelete2.Picture
End If
If Err.Number <> 0 Then SetError "imgDelete_MouseDown()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgDelete_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 Then
    PlayWav App.Path & "\media\click2.wav", SND_ASYNC
    PopupMenu mnuDel
    imgDelete.Picture = imgDelete1.Picture
End If
If Err.Number <> 0 Then SetError "imgDelete_MouseUp()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 Then
    imgExit.Picture = imgExit2.Picture
End If
If Err.Number <> 0 Then SetError "imgExit_MouseDown()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 Then
    WriteINI lIniFiles.iSettings, "PlaylistWind", "Visible", "False"
    PlayWav App.Path & "\media\click.wav", SND_ASYNC
    imgExit.Picture = imgExit1.Picture
    Unload Me
End If
If Err.Number <> 0 Then SetError "imgExit_MouseUp()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgPlay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 Then
    imgPlay.Picture = imgPlay2.Picture
End If
If Err.Number <> 0 Then SetError "imgPlay_MouseDown()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgPlay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 Then
    PlayWav App.Path & "\media\click2.wav", SND_ASYNC
    PopupMenu mnuPlay
    imgPlay.Picture = imgPlay1.Picture
End If
If Err.Number <> 0 Then SetError "imgPlay_MouseUp()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub lblInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then SetError "lblInfo()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub lstPlaylist_Click()
On Local Error Resume Next
Dim msg As String, msg2 As String, lFile As String, i As Integer, l As Integer, s As Integer
If lEvents.eEncoderBusy = True Or lPlayer.pStatus = sPlaying Or lPlayer.pStatus = sPaused Or lPlayer.pStatus = sSeekingBackward Or lPlayer.pStatus = sSeekingForward Or lEvents.eRipperBusy = True Then
    lblInfo.Caption = ""
    Exit Sub
End If
If Right(LCase(lstPlaylist.Text), 3) = "m3u" Then
    i = FindPlaylistIndexByFile(lstPlaylist.Text)
    If i <> 0 Then
        SetFileLabels "", "", Playlist.pPlaylists(i).pFile, Playlist.pPlaylists(i).pPath
        msg = Playlist.pPlaylists(i).pPath & Playlist.pPlaylists(i).pFile
        If Len(msg) <> 0 And DoesFileExist(msg) = True Then
            For l = 1 To Playlist.pFileCount
                If Playlist.pFiles(l).fEnabled = True And Playlist.pFiles(l).fPlaylist = i Then
                    s = s + 1
                End If
            Next l
            lblInfo.Caption = s & " file(s)"
        Else
            lblInfo.Caption = ""
        End If
    End If
ElseIf Right(LCase(lstPlaylist.Text), 3) = "mp3" Then
    i = FindMediaIndex(lstPlaylist.Text)
    If i <> 0 Then
        SetFileLabels "", "", Playlist.pFiles(i).fFile, Playlist.pFiles(i).fPath
        msg = Playlist.pFiles(i).fPath & Playlist.pFiles(i).fFile
        If Len(msg) <> 0 And DoesFileExist(msg) = True And Playlist.pFiles(i).fEnabled = True Then
            lTag.tFile = msg
            GetTagInfo
            DoEvents
            lblInfo.Caption = lTag.tAlbum
        End If
    End If
Else
    lblInfo.Caption = ""
End If
If Err.Number <> 0 Then SetError "lstPlaylist_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub lstPlaylist_DblClick()
On Local Error Resume Next
Dim i As Integer, msg As String
If Right(lstPlaylist.Text, 4) = ".m3u" Then
    If lEvents.eSettings.iPlayMp3sInNexENCODE = True Then
        i = FindPlaylistIndex(Left(lstPlaylist.Text, Len(lstPlaylist.Text) - 4))
        PlayPlaylist "", "", i
    Else
        i = FindPlaylistIndex(Left(lstPlaylist.Text, Len(lstPlaylist.Text) - 4))
        If i <> 0 Then
            GoMp3Player Playlist.pPlaylists(i).pPath & Playlist.pPlaylists(i).pFile
            Exit Sub
        End If
    End If
ElseIf Right(lstPlaylist.Text, 4) = ".mp3" Then
    i = FindMediaIndex(lstPlaylist.Text)
    If i <> 0 Then
        If lEvents.eSettings.iPlayMp3sInNexENCODE = True Then
            If lPlayer.pStatus = sPlaying Then
                StopMp3
            End If
            DoEvents
            AddEvent Play, Playlist.pFiles(i).fPath, Playlist.pFiles(i).fFile, "", "", 0, ""
            Exit Sub
        Else
            GoMp3Player Trim(Playlist.pFiles(i).fPath & Playlist.pFiles(i).fFile)
            Exit Sub
        End If
    End If
End If
If Err.Number <> 0 Then SetError "lstPlaylist_DblClick()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub lstPlaylist_KeyPress(KeyAscii As Integer)
On Local Error Resume Next
PlaylistKeypress KeyAscii
If Err.Number <> 0 Then SetError "lstPlaylist_Keypress()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub lstPlaylist_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 2 Then
    If Len(lstPlaylist.Text) <> 0 Then PopupMenu mnuListbox
End If
If Err.Number <> 0 Then SetError "lstPlaylist_MouseDown()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub lstPlaylist_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 Then FormDrag Me
If Err.Number <> 0 Then SetError "lstPlaylist_MouseMove()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub lstPlaylist_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
DragDrop Data
If Err.Number <> 0 Then SetError "lstPlaylist_OleDragDrop()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuAddDirectory_Click()
On Local Error Resume Next
PromptMediaDir
If Err.Number <> 0 Then SetError "mnuAddDirectory_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuAddFiles_Click()
On Local Error Resume Next
Dim i As Integer

If Len(cboPlaylists.Text) <> 0 Then
    i = FindPlaylistIndex(cboPlaylists.Text)
    If i <> 0 Then
        PromptAddToPlaylist i
    Else
        PromptAddToPlaylist 0
    End If
End If
If Err.Number <> 0 Then SetError "mnuAddFiles_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuDecodeListbox_Click()
On Local Error Resume Next
Dim msg As String, i As Integer, msg2 As String, lPath As String, lWav As String
i = FindMediaIndex(lstPlaylist.Text)
If i <> 0 Then
    msg = Playlist.pFiles(i).fPath & Playlist.pFiles(i).fFile
    If Len(msg) <> 0 Then
        If DoesFileExist(msg) = True Then
            msg2 = msg
            msg2 = GetFileTitle(msg2)
            lPath = Left(msg, Len(msg) - Len(msg2))
            lWav = Left(msg2, Len(msg2) - 3) & "wav"
            AddEvent Decode, lPath, msg2, lPath, lWav, 0, ""
        End If
    End If
End If
If Err.Number <> 0 Then SetError "mnuDecode_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuDeleteFromDisc_Click()
On Local Error Resume Next
Dim mbox As VbMsgBoxResult, msg As String, i As Integer

i = FindMediaIndex(lstPlaylist.Text)
If i = 0 Then Exit Sub
If lEvents.eSettings.iOverwritePrompts = True Then
    mbox = MsgBox("Are you sure you wish to perminantly delete this file from disc?", vbQuestion + vbYesNo)
    If mbox = vbYes Then
        Kill Playlist.pFiles(i).fPath & Playlist.pFiles(i).fFile
        RemoveFromPlaylist i
    ElseIf mbox = vbNo Then
        Exit Sub
    End If
Else
    Kill Playlist.pFiles(i).fPath & Playlist.pFiles(i).fFile
    RemoveFromPlaylist i
End If
If Err.Number <> 0 Then SetError "mnuDeleteFromDisc_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuDELPlaylist_Click()
On Local Error Resume Next
Dim i As Integer, msg As String, msg2 As String
msg = MsgBox("Are you sure you would like to delete all playlist information from your disc?", vbYesNo + vbExclamation, "Delete Playlists?")
If msg = vbYes Then
    For i = 1 To Playlist.pFileCount
        Playlist.pFiles(i).fEnabled = False
        Playlist.pFiles(i).fFile = ""
        Playlist.pFiles(i).fPath = ""
        Playlist.pFiles(i).fPlaylist = 0
    Next i
    For i = 1 To Playlist.pPlaylistCount
        If DoesFileExist(Playlist.pPlaylists(i).pPath & "\" & Playlist.pPlaylists(i).pFile) = True Then
            Kill Playlist.pPlaylists(i).pPath & "\" & Playlist.pPlaylists(i).pFile
        End If
        Playlist.pPlaylists(i).pEnabled = False
        Playlist.pPlaylists(i).pDescription = ""
        Playlist.pPlaylists(i).pFile = ""
        Playlist.pPlaylists(i).pPath = ""
    Next i
    Playlist.pFileCount = 0
    Playlist.pPlaylistCount = 0
    If DoesFileExist(lIniFiles.iPlaylists) = True Then
        msg2 = lIniFiles.iPlaylists
        msg2 = GetFileTitle(msg2)
        Kill lIniFiles.iPlaylists
    End If
    lstPlaylist.Clear
End If
If Err.Number <> 0 Then SetError "mnuDelPlaylist_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuDelThisPlaylist_Click()
On Local Error Resume Next
Dim i As Integer, msg As String
i = FindPlaylistIndex(cboPlaylists.Text)
If i <> 0 Then
    If DoesFileExist(Playlist.pPlaylists(i).pPath & "\" & Playlist.pPlaylists(i).pFile) = True Then
        Kill Playlist.pPlaylists(i).pPath & "\" & Playlist.pPlaylists(i).pFile
        Playlist.pPlaylists(i).pDescription = ""
        Playlist.pPlaylists(i).pEnabled = False
        Playlist.pPlaylists(i).pFile = ""
        Playlist.pPlaylists(i).pPath = ""
    End If
End If
Unload frmPlaylist
Load frmPlaylist
If Err.Number <> 0 Then SetError "mnuDelThisPlaylist_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub MNUFROMLIBRARY_Click()
On Local Error Resume Next
Dim i As Integer, f As Integer
f = FindPlaylistIndex(cboPlaylists.Text)
i = FindMediaIndex(lstPlaylist.Text)
If i <> 0 Then
    With Playlist.pFiles(i)
        .fEnabled = False
        .fFile = ""
        .fPath = ""
        f = .fPlaylist
        .fPlaylist = 0
    End With
    'SavePlaylists
    RemoveFromPlaylist f
    SavePlaylist f
    'FillWithPlaylist Playlist.pPlaylists(i).pDescription
    lstPlaylist.RemoveItem lstPlaylist.ListIndex
End If
If Err.Number <> 0 Then SetError "mnuFromLibrary_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuMerge_Click()
On Local Error Resume Next
Dim i As Integer
i = FindMediaIndex(lstPlaylist.Text)
If i <> 0 Then
    frmFileMerger.AddToMergeList Playlist.pFiles(i).fPath & Playlist.pFiles(i).fFile
End If
If Err.Number <> 0 Then SetError "mnuMerge_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuMoreInfoListbox_Click()
On Local Error Resume Next
Dim i As Integer, msg As String
msg = lstPlaylist.Text
If Len(msg) <> 0 Then
    i = FindMediaIndex(msg)
    If i <> 0 Then
        frmMP3Info.Show
        
        
        PromptGetTag Playlist.pFiles(i).fPath & Playlist.pFiles(i).fFile
        frmMP3Info.RefreshTagInfo
    End If
End If
If Err.Number <> 0 Then SetError "mnuMoreInfoListBox_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuMovetoPlaylist_Click()
Dim i As Integer
i = FindPlaylistIndex("Recient")
If i <> 0 Then
    MoveMediaToPlaylist FindMediaIndex(lstPlaylist.Text), i
Else
    i = AddPlaylist("Recient", "Recient.m3u")
    MoveMediaToPlaylist FindMediaIndex(lstPlaylist.Text), i
End If
End Sub

Private Sub mnuPlayFiles_Click()
On Local Error Resume Next
Dim i As Integer, msg As String
If Len(lstPlaylist.Text) <> 0 Then
    lstPlaylist_DblClick
End If
If Err.Number <> 0 Then SetError "mnuPlayFiles_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuPlayHTMLPlaylist_Click()
On Local Error Resume Next
If cboPlaylists.Text = "<All>" Or cboPlaylists.Text = "<Playlists>" Or cboPlaylists.Text = "<New>" Or cboPlaylists.Text = "<Search>" Then
    PlaylistToHTMLFile True, 0, True
    Exit Sub
End If
If Playlist.pIndex <> 0 Then
    If Playlist.pPlaylists(Playlist.pIndex).pEnabled = True Then
        PlaylistToHTMLFile False, Playlist.pIndex, True
    End If
End If
If Err.Number <> 0 Then SetError "mnuPlayHtmlPlaylist_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuPlayListbox_Click()
On Local Error Resume Next
Dim i As Integer, msg As String
If Len(lstPlaylist.Text) <> 0 Then
    lstPlaylist_DblClick
End If
If Err.Number <> 0 Then SetError "mnuPlayFiles_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuPlayPlaylist_Click()
On Local Error Resume Next
Dim msg As String, i As Integer
If Playlist.pIndex <> 0 Then
    With Playlist.pPlaylists(Playlist.pIndex)
        For i = 1 To Playlist.pFileCount
            If Playlist.pFiles(i).fPlaylist = Playlist.pIndex Then
                AddEvent Play, Playlist.pFiles(i).fPath, Playlist.pFiles(i).fFile, "", "", 0, ""
                pause 0.2
            End If
        Next i
    End With
End If
If Err.Number <> 0 Then SetError "mnuPlayPlaylist_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuRemoveFromPlaylist_Click()
RemoveFromPlaylist FindMediaIndex(lstPlaylist.Text)
End Sub
