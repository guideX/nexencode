Attribute VB_Name = "mdlPlaylist"
Option Explicit
Private Type gPlaylists
    pEnabled As Boolean
    pPath As String
    pFile As String
    pDescription As String
End Type
Private Type gFiles
    fEnabled As Boolean
    fPath As String
    fFile As String
    fPlaylist As Integer
End Type
Private Type gPlaylist
    pFiles(2464) As gFiles
    pPlaylists(500) As gPlaylists
    pFileCount As Integer
    pPlaylistCount As Integer
    pIndex As Integer
End Type
Global Playlist As gPlaylist

Public Sub PlaylistToHTMLFile(lAllPlaylists As Boolean, Optional lPlaylist As Integer, Optional lSurf As Boolean)
On Local Error Resume Next
Dim msg As String, msg2 As String, i As Integer, m As Integer, msg3 As String

If lAllPlaylists = True Then
    If Playlist.pFileCount <> 0 Then
        For i = 1 To Playlist.pFileCount
            lTag.tFile = Playlist.pFiles(i).fPath & Playlist.pFiles(i).fFile
            GetTagInfo
            DoEvents
            If lTag.tHasTag = True And Len(Trim(lTag.tArtist)) <> 0 And Len(Trim(lTag.tTitle)) <> 0 Then
                msg2 = Trim(lTag.tArtist) & " - " & Trim(lTag.tTitle)
                If Playlist.pFiles(i).fEnabled = True And Len(Playlist.pFiles(i).fFile) <> 0 Then msg = msg & "<a href=" & Chr(34) & "file://" & Playlist.pFiles(i).fPath & Playlist.pFiles(i).fFile & Chr(34) & "><font face=Tahoma size=2>" & msg2 & "</a> - " & Playlist.pPlaylists(Playlist.pFiles(i).fPlaylist).pDescription & "</font><br>" & vbCrLf
            Else
                If Playlist.pFiles(i).fEnabled = True And Len(Playlist.pFiles(i).fFile) <> 0 Then msg = msg & "<a href=" & Chr(34) & "file://" & Playlist.pFiles(i).fPath & Playlist.pFiles(i).fFile & Chr(34) & "><font face=Tahoma size=2>" & Left(Playlist.pFiles(i).fFile, Len(Playlist.pFiles(i).fFile) - 4) & "</a> - " & Playlist.pPlaylists(Playlist.pFiles(i).fPlaylist).pDescription & "</font><br>" & vbCrLf
            End If
            m = m + 1
        Next i
        msg = "<html>" & vbCrLf & "<body vlink=Blue alink=White bgcolor=#000000 text=#FFFFFF link=Blue>" & vbCrLf & "<font face=Tahoma size=5>" & vbCrLf & "<b><p>NexENCODE Studio Playlist</b><br>" & "<font face=Tahoma size=2>Total Files: " & m & "</font></p>" & vbCrLf & "<hr>" & msg & vbCrLf & vbCrLf & "</font></body></html>"
        msg3 = App.Path & "\html\All.htm"
        SaveFile msg3, msg
        If lSurf = True Then Surf msg3
    Else
        If lEvents.eSettings.iOverwritePrompts = True Then MsgBox "Error Printing to html file"
        Exit Sub
    End If
Else
    If Playlist.pFileCount <> 0 Then
        For i = 1 To Playlist.pFileCount
            If Playlist.pFiles(i).fPlaylist = lPlaylist Then
                lTag.tFile = Playlist.pFiles(i).fPath & Playlist.pFiles(i).fFile
                GetTagInfo
                DoEvents
                If lTag.tHasTag = True And Len(Trim(lTag.tArtist)) <> 0 And Len(Trim(lTag.tTitle)) <> 0 Then
                    msg2 = Trim(lTag.tArtist) & " - " & Trim(lTag.tTitle)
                    If Playlist.pFiles(i).fEnabled = True And Len(Playlist.pFiles(i).fFile) <> 0 Then msg = msg & "<a href=" & Chr(34) & "file://" & Playlist.pFiles(i).fPath & Playlist.pFiles(i).fFile & Chr(34) & "><font face=Tahoma size=2>" & msg2 & "</a> - " & Playlist.pPlaylists(Playlist.pFiles(i).fPlaylist).pDescription & "</font><br>" & vbCrLf
                Else
                    If Playlist.pFiles(i).fEnabled = True And Len(Playlist.pFiles(i).fFile) <> 0 Then msg = msg & "<a href=" & Chr(34) & "file://" & Playlist.pFiles(i).fPath & Playlist.pFiles(i).fFile & Chr(34) & "><font face=Tahoma size=2>" & Left(Playlist.pFiles(i).fFile, Len(Playlist.pFiles(i).fFile) - 4) & "</a> - " & Playlist.pPlaylists(Playlist.pFiles(i).fPlaylist).pDescription & "</font><br>" & vbCrLf
                End If
                m = m + 1
            End If
        Next i
        msg = "<html>" & vbCrLf & "<body vlink=Blue alink=White bgcolor=#000000 text=#FFFFFF link=Blue>" & vbCrLf & "<font face=Tahoma size=5>" & vbCrLf & "<b><p>NexENCODE Studio Playlist</b><br>" & "<font face=Tahoma size=2>Total Files: " & m & "</font></p>" & vbCrLf & "<hr>" & msg & vbCrLf & vbCrLf & "</font></body></html>"
        msg3 = App.Path & "\html\" & Playlist.pPlaylists(lPlaylist).pDescription & ".htm"
        SaveFile msg3, msg
        If lSurf = True Then Surf msg3
    Else
        If lEvents.eSettings.iOverwritePrompts = True Then MsgBox "Error Printing to html file"
        Exit Sub
    End If
End If

If Err.Number <> 0 Then SetError "PlaylistToHTMLFile", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub ResetPlayButtons()
On Local Error Resume Next
With frmMain
If lEvents.eSettings.iPlayMp3sInNexENCODE = False Then
    .mnuAudica.Visible = True
    .imgPlay.Visible = False
    .imgStop.Visible = False
    .imgBackward.Visible = False
    .imgForward.Visible = False
Else
    .mnuAudica.Visible = False
    .imgPlay.Visible = True
    .imgStop.Visible = True
    .imgBackward.Visible = True
    .imgForward.Visible = True
End If
End With
End Sub

Public Sub AddDirToPlaylist(lPlaylist As Integer)
On Local Error Resume Next
Dim msg As String, i As Integer

lEvents.eRetStr = ""
frmSelectDir.Show 1
If Len(lEvents.eRetStr) <> 0 Then
    For i = 0 To frmSelectDir.File1.ListCount
        msg = frmSelectDir.File1.List(i)
        If Len(msg) <> 0 Then AddToPlaylist lEvents.eRetStr & "\" & msg, lPlaylist
    Next i
End If

If Err.Number <> 0 Then SetError "AddDirToPlaylist", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub LoadM3uIntoPlaylist(lIndex As Integer)
On Local Error Resume Next
Dim k As Integer, lFilename As String, msg As String, msg2 As String, i As Integer, lTitle As String, lPath As String

If lIndex = 0 Then
    SetError "LoadM3uIntoPlaylist()", "Integer value unspecified", "lIndex=0"
    Exit Sub
End If
lFilename = Playlist.pPlaylists(lIndex).pPath & Playlist.pPlaylists(lIndex).pFile
If DoesFileExist(lFilename) = True Then
    msg = ReadFile(lFilename)
    msg = Trim(msg)
    Dim lefty As String
Again:
    If Len(msg) <> 0 Then
        If InStr(msg, Chr(10)) Then
            If Left(msg, 1) = Chr(10) Or Left(msg, 1) = Chr(13) Then
                msg = Right(msg, Len(msg) - 1)
                GoTo Again
            ElseIf Right(msg, 1) = Chr(10) Or Right(msg, 1) = Chr(13) Then
                msg = Left(msg, Len(msg) - 1)
                GoTo Again
            Else
                lefty = Left(msg, 1)
                msg2 = lefty & ParseString(msg, Left(msg, 1), Chr(10))
                If Len(msg) = Len(msg2) Then
                    msg = ""
                Else
                    If Len(msg) > Len(msg2) Then
                        msg = Right(msg, Len(msg) - Len(msg2))
                    ElseIf Len(msg) < Len(msg2) Then
                        msg = ""
                    End If
                End If
            End If
        Else
            If Len(msg) > 1 Then
                msg2 = Trim(msg)
                msg = ""
            Else
                Exit Sub
            End If
        End If
CheckFile:
        If Asc(Left(msg2, 1)) = 10 Or Asc(Left(msg2, 1)) = 13 Then msg2 = Right(msg2, Len(msg2) - 1)
        If Asc(Right(msg2, 1)) = 10 Or Asc(Right(msg2, 1)) = 13 Then msg2 = Left(msg2, Len(msg2) - 1)
        If Asc(Left(msg2, 1)) = 10 Or Asc(Left(msg2, 1)) = 13 Or Asc(Right(msg2, 1)) = 10 Or Asc(Right(msg2, 1)) = 13 Then GoTo CheckFile
        lTitle = msg2
        lTitle = GetFileTitle(lTitle)
        If FindMediaIndex(lTitle) = 0 Then
            With Playlist.pFiles(Playlist.pFileCount)
                If DoesFileExist(msg2) = True Then
                    Playlist.pFileCount = Playlist.pFileCount + 1
                    .fEnabled = True
                    .fFile = lTitle
                    .fPath = Left(msg2, Len(msg2) - Len(lTitle))
                    .fPlaylist = lIndex
                End If
            End With
            DoEvents
            GoTo Again
        End If
    Else
        Exit Sub
    End If
End If

If Err.Number <> 0 Then SetError "LoadM3uIntoPlaylist", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub CountFiles()
On Local Error Resume Next
Dim i As Integer
Playlist.pFileCount = 0
For i = 1 To 2464
    If Playlist.pFiles(i).fEnabled = True Then
        Playlist.pFileCount = Playlist.pFileCount + 1
    End If
Next i
Playlist.pPlaylistCount = 0
For i = 1 To 500
    If Playlist.pPlaylists(i).pEnabled = True Then
        Playlist.pPlaylistCount = Playlist.pPlaylistCount + 1
    End If
Next i
If Err.Number <> 0 Then SetError "CountFiles()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub LoadPlaylists()
On Local Error Resume Next
Dim i As Integer, X As Integer
Dim msg As String, msg2 As String, msg3 As String

Playlist.pPlaylistCount = ReadINI(lIniFiles.iPlaylists, "Settings", "Count", 0)
For i = 1 To Playlist.pPlaylistCount
    Playlist.pPlaylists(i).pEnabled = ReadINI(lIniFiles.iPlaylists, i, "Enabled", "False")
    If Playlist.pPlaylists(i).pEnabled = True Then
        Playlist.pPlaylists(i).pDescription = ReadINI(lIniFiles.iPlaylists, i, "Description", "")
        Playlist.pPlaylists(i).pFile = ReadINI(lIniFiles.iPlaylists, i, "File", "")
        Playlist.pPlaylists(i).pPath = ReadINI(lIniFiles.iPlaylists, i, "Path", "")
        If Right(Playlist.pPlaylists(i).pPath, 1) <> "\" Then Playlist.pPlaylists(i).pPath = Playlist.pPlaylists(i).pPath & "\"
        For X = 1 To Playlist.pPlaylistCount
            LoadM3uIntoPlaylist i
        Next X
    End If
Next i

If Err.Number <> 0 Then SetError "LoadPlaylists", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Function AddPlaylist(lName As String, Optional lFile As String) As Integer
On Local Error Resume Next
Dim i As Integer, f As Integer
f = FindPlaylistIndex(lName)
If f <> 0 Then
    AddPlaylist = f
    Exit Function
End If
If Len(lName) <> 0 Then
    If Len(lFile) = 0 Then lFile = ReturnDirCompliant(lName & ".m3u")
    i = Playlist.pPlaylistCount + 1
    Playlist.pPlaylistCount = i
    With Playlist.pPlaylists(i)
        .pDescription = lName
        .pEnabled = True
        .pFile = lFile
        .pPath = App.Path & "\playlists\"
        WriteINI lIniFiles.iPlaylists, i, "Enabled", "True"
        WriteINI lIniFiles.iPlaylists, i, "Description", lName
        WriteINI lIniFiles.iPlaylists, i, "File", .pFile
        WriteINI lIniFiles.iPlaylists, i, "Path", .pPath
        WriteINI lIniFiles.iPlaylists, "Settings", "Count", Playlist.pPlaylistCount
    End With
    AddPlaylist = i
End If
If Err.Number <> 0 Then SetError "AddPlaylist", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Function RemoveCharReturns(lText As String) As String
On Local Error Resume Next
Dim lAgain As Boolean

lAgain = True
Again:
If lAgain = False Then
    RemoveCharReturns = lText
    Exit Function
End If
If Len(lText) <> 0 Then
    If InStr(lText, Chr(13)) Then
        If Left(lText, 1) = Chr(13) Then
            lText = Right(lText, Len(lText) - 1)
            lAgain = True
        ElseIf Right(lText, 1) = Chr(13) Then
            lText = Left(lText, Len(lText) - 1)
            lAgain = True
        ElseIf Right(lText, 1) <> Chr(13) And Left(lText, 1) <> Chr(13) Then
            lAgain = False
        End If
        GoTo Again
    End If
End If

If Err.Number <> 0 Then SetError "RemoveCharReturns", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Function DoesMediaExistInPlaylist(lFile As String, lPlaylist As Integer) As Boolean
On Local Error Resume Next
Dim i As Integer, s As Integer

If Len(lFile) <> 0 And lPlaylist <> 0 And Playlist.pPlaylists(lPlaylist).pEnabled = True Then
    i = FindMediaIndex(lFile)
    If i <> 0 Then
        If Playlist.pFiles(i).fPlaylist = lPlaylist Then
            DoesMediaExistInPlaylist = True
        Else
            DoesMediaExistInPlaylist = False
        End If
    End If
End If
If Err.Number <> 0 Then SetError "DoesMediaExistInPlaylist", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Sub MoveMediaToPlaylist(lMediaIndex As Integer, lPlaylistIndex As Integer)
On Local Error Resume Next
Dim lFile As String, lPath As String
If Playlist.pFiles(lMediaIndex).fEnabled = True And Playlist.pPlaylists(lPlaylistIndex).pEnabled = True Then
    lFile = Playlist.pFiles(lMediaIndex).fFile
    lPath = Playlist.pFiles(lMediaIndex).fPath
    RemoveFromPlaylist lMediaIndex
    If AddToPlaylist(lPath & lFile, lPlaylistIndex) = True Then SavePlaylist lPlaylistIndex
End If
If Err.Number <> 0 Then SetError "MoveMediaToPlaylist()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub PromptAddToPlaylist(lPlaylist As Integer)
On Local Error Resume Next
Dim lFile As String, lPath As String, i As Integer, mbox As VbMsgBoxResult
lFile = OpenDialog(frmPlaylist, "Mpeg Layer 3 Files (*.mp3)|*.mp3", "Select MP3 File ...", CurDir)
If Len(lFile) <> 0 Then
    lPath = lFile
    lFile = GetFileTitle(lFile)
    lPath = Left(lPath, Len(lPath) - Len(lFile))
    i = FindMediaIndex(lFile)
    If i <> 0 Then
        If lEvents.eSettings.iOverwritePrompts = True Then
            mbox = MsgBox("This file already exists in another playlist. Would you like to remove it from this playlist before adding it to another?", vbYesNo + vbQuestion)
            If mbox = vbYes Then
                RemoveFromPlaylist i
                DoEvents
            ElseIf mbox = vbNo Then
                Exit Sub
            End If
        Else
            RemoveFromPlaylist i
            DoEvents
        End If
    End If
    If Playlist.pPlaylists(lPlaylist).pEnabled = True Then
        If AddToPlaylist(lPath & lFile, lPlaylist) = True Then
            DoEvents
            SavePlaylist FindPlaylistIndexByFile(lFile)
        End If
    Else
        lPlaylist = AddPlaylist("Recient", "Recient.m3u")
        If AddToPlaylist(lPath & lFile, lPlaylist) = True Then
            DoEvents
            SavePlaylist FindPlaylistIndexByFile(lFile)
        End If
        If lEvents.ePlaylistVisible = True Then
            Unload frmPlaylist
            frmPlaylist.Show
        End If
    End If
End If

If Err.Number <> 0 Then SetError "PromptPlaylist()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Function AddToPlaylist(lFilename As String, lPlaylist As Integer) As Boolean
On Local Error Resume Next
Dim i As Integer, msg As String
If DoesMediaExistInPlaylist(lFilename, lPlaylist) = True Then
    AddToPlaylist = True
    Exit Function
End If
If Len(lFilename) <> 0 And lPlaylist <> 0 Then
    i = Playlist.pFileCount + 1
    Playlist.pFileCount = i
    msg = lFilename
    msg = GetFileTitle(msg)
    With Playlist.pFiles(i)
        .fEnabled = True
        .fFile = msg
        .fPath = Left(lFilename, Len(lFilename) - Len(.fFile))
        .fPlaylist = lPlaylist
        If Right(.fPath, 1) <> "\" Then .fPath = .fPath & "\"
    End With
    If Playlist.pPlaylists(lPlaylist).pEnabled = False Then
        lPlaylist = AddPlaylist("Recient", "recient.m3u")
        DoEvents
    End If
    SavePlaylist lPlaylist
    AddToPlaylist = True
    If lEvents.ePlaylistVisible = True Then
        i = FindComoboxIndex(frmPlaylist, Playlist.pPlaylists(lPlaylist).pDescription)
        If frmPlaylist.cboPlaylists.ListIndex <> i Then
            If i <> 2 And i <> 3 Then
                frmPlaylist.cboPlaylists.ListIndex = i
            Else
                frmPlaylist.cboPlaylists.ListIndex = 1
            End If
        End If
    End If
Else
    AddToPlaylist = False
End If
If Err.Number <> 0 Then SetError "AddToPlaylist", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Sub SaveAsPlaylist(lData As String, Optional lPath As String, Optional lFile As String)
On Local Error Resume Next
Dim lFilename As String
If Right(lPath, 1) <> "\" Then lPath = lPath & "\"
lFilename = lPath & lFile
If Len(lFilename) = 0 Or Len(lFilename) = 1 Then
    lFilename = ""
    lFilename = SaveDialog(frmMain, "Playlist Files (*.m3u)|*.m3u|All Files (*.*)|*.*", "NexENCODE - Save as?", CurDir)
    If Len(lFilename) = 0 Then Exit Sub
    lFilename = Left(lFilename, Len(lFilename) - 1)
End If
If Right(LCase(lFilename), 4) <> ".m3u" Then lFilename = lFilename & ".m3u"
SaveFile lFilename, lData
If Err.Number <> 0 Then SetError "SaveAsPlaylist()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub SavePlaylist(lIndex As Integer)
On Local Error Resume Next
Dim lFilename As String, i As Integer, msg2 As String

lFilename = Playlist.pPlaylists(lIndex).pPath & Playlist.pPlaylists(lIndex).pFile
Start:
If Len(lFilename) = 0 Or Len(lFilename) = 1 Then
    lFilename = SaveDialog(frmMain, "M3u Files (*.m3u)|*.m3u|All Files (*.*)|*.*", "NexENCODE - Save as?", CurDir)
    If Len(lFilename) = 0 Then Exit Sub
    lFilename = Left(lFilename, Len(lFilename) - 1)
End If
If Right(LCase(lFilename), 4) <> ".m3u" Then lFilename = lFilename & ".m3u"
For i = 1 To Playlist.pFileCount
    If Playlist.pFiles(i).fPlaylist = lIndex Then
        If Len(msg2) = 0 Then
            msg2 = Playlist.pFiles(i).fPath & Playlist.pFiles(i).fFile
        Else
            msg2 = msg2 & vbCrLf & Playlist.pFiles(i).fPath & Playlist.pFiles(i).fFile
        End If
    End If
Next i
SaveFile lFilename, msg2
If Err.Number <> 0 Then SetError "SavePlaylist", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub RemoveFromPlaylist(lIndex As Integer)
On Local Error Resume Next

If lIndex <> 0 Then
    With Playlist.pFiles(lIndex)
        .fEnabled = False
        .fFile = ""
        .fPath = ""
        SavePlaylist .fPlaylist
        DoEvents
        .fPlaylist = 0
    End With
    If lEvents.ePlaylistVisible = True Then
        Unload frmPlaylist
        frmPlaylist.Show
    End If
End If
If Err.Number <> 0 Then SetError "RemoveFromPlaylist", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Function FindPlaylistIndexByFile(lFile As String) As Integer
On Local Error Resume Next
Dim i As Integer

If Len(lFile) <> 0 Then
    For i = 1 To Playlist.pPlaylistCount
        If LCase(lFile) = LCase(Playlist.pPlaylists(i).pFile) Then
            FindPlaylistIndexByFile = i
            Exit For
        End If
    Next i
End If

If Err.Number <> 0 Then SetError "FindPlaylistIndexByFile", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Function FindPlaylistIndex(lDescription As String) As Integer
On Local Error Resume Next
Dim i As Integer
If Len(lDescription) <> 0 Then
    For i = 1 To Playlist.pPlaylistCount
        If i <> 500 And i <> 501 Then
            If LCase(lDescription) = LCase(Playlist.pPlaylists(i).pDescription) Then
                FindPlaylistIndex = i
                Exit For
            End If
        Else
            Exit For
        End If
    Next i
End If
If Err.Number <> 0 Then SetError "FindPlaylistIndex", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Function FindNextIndexInPlaylist(lIndex As Integer) As Integer
On Local Error Resume Next
Dim i As Integer

If lIndex <> 0 Then
    For i = 1 To Playlist.pFileCount
        If Playlist.pFiles(i).fPlaylist = lIndex Then
            FindNextIndexInPlaylist = i
            Exit For
        End If
    Next i
End If

If Err.Number <> 0 Then SetError "FindPlaylistIndex", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Function FindMediaIndex(lFilename As String) As Integer
On Local Error Resume Next
Dim i As Integer
If Len(lFilename) <> 0 Then
    For i = 1 To Playlist.pFileCount
        If i < 2464 Then
            If LCase(lFilename) = LCase(Playlist.pFiles(i).fFile) Then
                FindMediaIndex = i
                Exit For
            End If
        End If
    Next i
End If
If Err.Number <> 0 Then SetError "FindMediaIndex", lEvents.eSettings.iErrDescription, Err.Description
End Function
