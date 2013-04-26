Attribute VB_Name = "mdlPlayer"
Enum ePlayerType
    pVFMp3Player = 1
    pWMP = 2
    pDialogMedian = 3
End Enum
Enum eStatusTypes
    sNotPlaying = 0
    sPlaying = 1
    sPaused = 2
    sSeekingForward = 3
    sSeekingBackward = 4
End Enum
Private Type gCurrentFile
    cBitrate As Integer
    cSamplerate As Long
    cTotalFrame As Long
    cArtist As String
    cAlbum As String
    cTitle As String
    cGenre As String
End Type
Private Type gLabels
    lWavFile As String
    lWavPath As String
    lMp3File As String
    lMp3Path As String
End Type
Private Type gPlayer
    pPlayerType As ePlayerType
    pCurrentFile As gCurrentFile
    pTotalTime As Long
    pStatus As eStatusTypes
    pPosition As Integer
    pPlaylistIndex As Integer
    pMediaIndex As Integer
    pStatusString As String
    pTime As String
    pContinuous As Boolean
    pPlayCanceled As Boolean
    pLabels As gLabels
End Type
Global lPlayer As gPlayer

Public Sub StopPlayQue()
On Local Error Resume Next
Dim i As Integer, m As Integer

m = 100
If m <> 0 Then
    For i = 1 To m
        If lEvents.eEvent(i).eEventType = Play Then
            RemoveEvent i
            DoEvents
        End If
    Next i
End If

If Err.Number <> 0 Then SetError "StopPlayQue()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub AddPlayEvent(lPath As String, lFile As String, Optional lPlaylist As Integer)
On Local Error Resume Next
Dim i As Integer, m As Integer

If lPlayer.pStatus = sPlaying Or lPlayer.pStatus = sPaused Then
    StopMp3
    DoEvents
    pause 0.2
End If
If Len(lPath) <> 0 And Len(lFile) <> 0 Then
    If DoesFileExist(lPath & lFile) = False Then
        If lEvents.eSettings.iOverwritePrompts = True Then MsgBox "File Doesn't Exist!", vbExclamation
        Exit Sub
    End If
    i = FindMediaIndex(lFile)
    If i <> 0 Then
        DoEvents
        AddEvent Play, lPath, lFile, "", "", 0, ""
    Else
        If lPlaylist <> 0 And Playlist.pPlaylists(lPlaylist).pEnabled = True Then
            If AddToPlaylist(lPath & lFile, lPlaylist) = True Then
                AddEvent Play, lPath, lFile, "", "", 0, ""
            Else
                AddEvent Play, lPath, lFile, "", "", 0, ""
            End If
        Else
            i = AddPlaylist("Recient", "Recient.m3u")
            If i <> 0 Then
                AddToPlaylist lPath & lFile, i
                DoEvents
                SavePlaylist i
                If lEvents.ePlaylistVisible = True Then
                    Unload frmPlaylist
                    frmPlaylist.Show
                End If
            End If
            AddEvent Play, lPath, lFile, "", "", 0, ""
        End If
    End If
End If

If Err.Number <> 0 Then SetError "AddPlayEvent()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub GoBackward()
On Local Error Resume Next
Dim i As Integer

i = Playlist.pIndex
If i = 0 Then
    i = Playlist.pFileCount
ElseIf i = Playlist.pFileCount Then
    i = i - 1
Else
    i = i - 1
End If

If Playlist.pFiles(i).fEnabled = True Then
    If lPlayer.pStatus = sPlaying Then StopMp3
'    PlayMp3 i
    AddEvent Play, Playlist.pFiles(i).fPath, Playlist.pFiles(i).fFile, "", "", 0, ""
    ProcessNextEvent
End If

If Err.Number <> 0 Then SetError "GoBackward()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub PauseMp3()
On Local Error Resume Next

If lPlayer.pStatus = sPaused Then
    lPlayer.pStatus = sPlaying
    frmMain.SimpleMP3.pause
Else
    lPlayer.pStatus = sPaused
    frmMain.SimpleMP3.pause
End If
If Err.Number <> 0 Then SetError "PauseMp3()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub GoForward()
On Local Error Resume Next
Dim i As Integer, b As Boolean

For i = 1 To lEvents.eEventCount
    If lEvents.eEvent(i).eEventType = Play Then
        b = True
    End If
Next i

If b = False Then
    i = Playlist.pIndex
    If i = 0 Then
        i = 1
    ElseIf i = Playlist.pFileCount Then
        i = 1
    Else
        i = i + 1
    End If
ElseIf b = True Then
    If lPlayer.pStatus = sPlaying Then StopMp3
    ProcessNextEvent
    Exit Sub
End If
If Playlist.pFiles(i).fEnabled = True Then
    If lPlayer.pStatus = sPlaying Then StopMp3
    AddEvent Play, Playlist.pFiles(i).fPath, Playlist.pFiles(i).fFile, "", "", 0, ""
    ProcessNextEvent
End If

If Err.Number <> 0 Then SetError "GoForward()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub StopMp3()
On Local Error Resume Next

lPlayer.pPlayCanceled = True
lPlayer.pStatus = sNotPlaying
lPlayer.pMediaIndex = 0
frmMain.tmrScrollStatus.Enabled = False
frmMain.SimpleMP3.sTop
frmMain.SimpleMP3.Close
frmMain.lblWavFile.Caption = ""

If Err.Number <> 0 Then SetError "StopMp3()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub PlayerDone()
On Local Error Resume Next

lPlayer.pMediaIndex = 0
lPlayer.pStatus = sNotPlaying
lPlayer.pCurrentFile.cAlbum = ""
lPlayer.pCurrentFile.cArtist = ""
lPlayer.pCurrentFile.cTitle = ""

'frmMain.Scope.Visible = False
frmMain.lblMp3File.Caption = ""
frmMain.lblWavFile.Caption = ""

ResetEncoderCircles
lEvents.eEncoderBusy = False
ConvertCaption oIdle
ToggleButtons oIdle
frmMain.tmrScrollStatus.Enabled = False
If Err.Number <> 0 Then SetError "PlayerDone()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub PlayMp3(lIndex As Integer)
On Local Error Resume Next
Dim i As Integer, msg As String
If lEvents.eSettings.iPlayMp3sInNexENCODE = False Then
    GoMp3Player Playlist.pFiles(lIndex).fPath & Playlist.pFiles(lIndex).fFile
    Exit Sub
End If
Playlist.pIndex = Playlist.pFiles(lIndex).fPlaylist
With Playlist.pFiles(lIndex)
    If Len(.fFile) <> 0 And Len(.fPath) <> 0 And DoesFileExist(.fPath & .fFile) = True Then
        frmMain.mnu1xFast.Checked = False
        frmMain.mnu2xFast.Checked = False
        frmMain.mnuNormal.Checked = True
        frmMain.mnu1xSlow.Checked = False
        frmMain.mnu2xSlow.Checked = False
        frmMain.SimpleMP3.SetSpeed 100
        lEvents.eTimeType = 1
        lPlayer.pMediaIndex = lIndex
        ResetEncoderCircles
        frmMain.tmrShowEncoderCircles.interval = 40
        frmMain.tmrShowEncoderCircles.Enabled = True
        lEvents.eEncoderBusy = True
        frmMain.SimpleMP3.Authorize "Leon J Aiossa", "812144397"
        FindOpenDevice frmMain.SimpleMP3
        i = frmMain.SimpleMP3.Open(.fPath & .fFile, "")
        frmMain.SimpleMP3.Play
        ToggleButtons oPlaying
        lArtist = frmMain.SimpleMP3.GetArtist
        lTitle = frmMain.SimpleMP3.GetTitle
        frmMain.lblMp3File.Caption = .fFile
        lPlayer.pStatus = sPlaying
        frmMain.tmrScrollStatus.Enabled = True
        lTag.tFile = .fPath & .fFile
        GetTagInfo
        If frmMain.SimpleMP3.BitRate <> -1 Then
            msg = Str(frmMain.SimpleMP3.BitRate)
            msg = Left(msg, Len(msg) - 3)
        End If
        If lTag.tHasTag = True Then
            If Len(Trim(lTag.tArtist)) <> 0 And Len(Trim(lTag.tTitle)) <> 0 Then
                lPlayer.pStatusString = "Play: " & Trim(lTag.tTitle) & " - by - " & Trim(lTag.tArtist) & " ... bitrate: " & msg & " ... Sample: " & frmMain.SimpleMP3.SampleFrequency & " ... "
            Else
                lPlayer.pStatusString = "Play: " & Left(.fFile, Len(.fFile) - 4) & " ... bitrate: " & msg & " ... Sample: " & frmMain.SimpleMP3.SampleFrequency & " ... "
            End If
        Else
            lPlayer.pStatusString = "Play: " & Left(.fFile, Len(.fFile) - 4) & " ... bitrate: " & msg & " ... Sample: " & frmMain.SimpleMP3.SampleFrequency & " ... "
        End If
    Else
        If lEvents.eSettings.iOverwritePrompts = True Then MsgBox "This file could not be found"
        frmMain.lblInfo.Caption = "File not found"
        SetError "PlayMP3", "File does not exist", "NexENCODE could not find the file specified"
        StopMp3
        PlayerDone
        Playlist.pIndex = 0
        Exit Sub
    End If
End With
If Err.Number <> 0 Then SetError "PlayMp3()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Function ResetFileCount() As Integer
On Local Error Resume Next
Dim X As Integer

For X = 1 To Playlist.pFileCount
    If Playlist.pFiles(X).fEnabled = True Then i = i + 1
Next X

ResetFileCount = i
If Err.Number <> 0 Then SetError "ResetFileCount()", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Sub LoadRandomMP3()
On Local Error Resume Next
Dim i As Integer

i = ResetFileCount
TryAgain:
If Playlist.pFileCount <> 0 Then
    i = GetRnd(i)
    If Len(Playlist.pFiles(i).fFile) <> 0 Then
        If lPlayer.pStatus = sPlaying Then StopMp3
        DoEvents
        'PlayMp3 i
        AddEvent Play, Playlist.pFiles(i).fPath, Playlist.pFiles(i).fFile, "", "", 0, ""
        ProcessNextEvent
    Else
        GoTo TryAgain
    End If
End If

If Err.Number <> 0 Then SetError "LoadRandomMP3()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub PlayPlaylist(Optional lPath As String, Optional lFile As String, Optional lIndex As Integer)
On Local Error Resume Next
Dim i As Integer, msg As String, t As Integer, s As Integer
If lIndex <> 0 Then
    If Playlist.pPlaylists(lIndex).pEnabled = True Then
        For i = 1 To Playlist.pFileCount
            If Playlist.pFiles(i).fPlaylist = lIndex Then
                If Len(Playlist.pFiles(i).fFile) <> 0 And Playlist.pFiles(i).fEnabled = True Then
                    AddEvent Play, Playlist.pFiles(i).fPath, Playlist.pFiles(i).fFile, "", "", 0, ""
'                    pause 0.2
                End If
            End If
        Next i
    End If
Else
    If Right(lPath, 1) <> "\" Then lPath = lPath & "\"
    If DoesFileExist(App.Path & "\playlists\" & lFile) = False Then FileCopy lPath & lFile, App.Path & "\playlists\" & lFile
    i = AddPlaylist(Left(lFile, Len(lFile) - 4), lFile)
    LoadPlaylists
    If lEvents.ePlaylistVisible = True Then
        Unload frmPlaylist
        DoEvents
        Load frmPlaylist
        Playlist.pIndex = i
        frmPlaylist.DisplayPlaylist i
    End If
    For t = 1 To Playlist.pFileCount
        If Playlist.pFiles(t).fPlaylist = i Then
            If Playlist.pFiles(t).fEnabled = True And Len(Playlist.pFiles(t).fFile) <> 0 And Len(Playlist.pFiles(t).fPath) <> 0 Then
                AddEvent Play, Playlist.pFiles(t).fPath, Playlist.pFiles(t).fFile, "", "", 0, ""
                pause 0.2: DoEvents
            End If
        End If
    Next t
    Exit Sub
End If
If Err.Number <> 0 Then SetError "PlayPlaylist()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub Progress(lProgress As Long, lMax As Long)
On Local Error Resume Next
Dim i As Integer, e As Integer, msg As String, msg2 As String, r As Long

If lMax <> 0 And lProgress <> 0 And lPlayer.pStatus = sPlaying Then
    i = lProgress / lMax * 9
    e = lProgress / lMax * 100
    If e = 97 Or e = 98 Or e = 100 Or e = 99 Then
        PlayerDone
        lPlayer.pStatus = sNotPlaying
        Exit Sub
    End If
    If e <> lPlayer.pPosition Then
        r = e
        EncodeCircleEffect i
        lPlayer.pPosition = e
        lPlayer.pTime = Format(lProgress, "00:00")
        SetCaption Play, r, LCase(lPlayer.pCurrentFile.cTitle)
    End If
End If
If Err.Number <> 0 Then SetError "Progress()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub PromptToPlay()
On Local Error Resume Next
Dim msg As String, lFilename As String, lFile As String, i As Integer, f As Integer, lPath As String, CNumber

msg = OpenDialog(frmMain, "MP3 Files (*.mp3)|*.mp3|Playlists (*.m3u)|*.m3u|Wave Audio (*.wav)|*.wav|All Files (*.*)|*.*", "Play NexENCODE Media", CurDir)
If Len(msg) <> 0 Then
    lFile = msg
    lFile = GetFileTitle(lFile)
    lPath = Left(msg, Len(msg) - Len(lFile))
    If Right(LCase(msg), 4) = ".wav" Then
        PlayWav msg, SND_ASYNC
    ElseIf Right(LCase(msg), 4) = ".m3u" Then
        PlayPlaylist lPath, lFile, FindPlaylistIndexByFile(msg)
    ElseIf Right(LCase(lFile), 4) = ".mp3" Then
        AddPlayEvent lPath, lFile
    End If
End If
If Err.Number <> 0 Then SetError "PromptToPlay()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub SetBallColor(lProgress As Integer)
Dim i As Integer
For i = 0 To 9
    'frmMain.shpEncoder(i).BackColor = vbRed
Next i
For i = 0 To lProgress
    frmMain.shpEncoder(i).BackColor = vbYellow
Next i
End Sub
