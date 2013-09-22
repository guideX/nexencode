Attribute VB_Name = "mdlSubs"
Option Explicit
Enum eWindowSize
    wLoading = 1
    wUnloading = 2
End Enum

Public Sub CheckMainButtonsOver(lObjectOver As eObjectTypes)
'On Local Error Resume Next
With frmMain
    If lObjectOver <> oBackwardButton And .imgBackward.Picture <> .imgBackward1.Picture Then .imgBackward.Picture = .imgBackward1.Picture
    If lObjectOver <> oStopRipping And .imgCancelRip.Picture <> .imgStopRipping1.Picture Then .imgCancelRip.Picture = .imgStopRipping1.Picture
    If lObjectOver <> oEncode And .imgEncode.Picture <> .imgEncode1.Picture Then .imgEncode.Picture = .imgEncode1.Picture
    If lObjectOver <> oEnd And .imgEnd.Picture <> .imgEnd1.Picture Then .imgEnd.Picture = .imgEnd1.Picture
    If lObjectOver <> oForwardButton And .imgForward.Picture <> .imgForward1.Picture Then .imgForward.Picture = .imgForward1.Picture
    If lObjectOver <> oTag And .imgId3.Picture <> .imgId31.Picture Then .imgId3.Picture = .imgId31.Picture
    If lObjectOver <> oMinimize And .imgMinimize.Picture <> .imgMinimize1.Picture Then .imgMinimize.Picture = .imgMinimize1.Picture
    If lObjectOver <> oCDAudio And .imgNexMedia.Picture <> .imgNexMedia1.Picture Then .imgNexMedia.Picture = .imgNexMedia1.Picture
    If lObjectOver <> oOptions And .imgOptions.Picture <> .imgOptions1.Picture Then .imgOptions.Picture = .imgOptions1.Picture
    If lObjectOver <> oPlayButton And .imgPlay.Picture <> .imgPlay1.Picture Then .imgPlay.Picture = .imgPlay1.Picture
    If lObjectOver <> oPlayMp3 And .imgPlayMp3.Picture <> .imgPlayMp31.Picture Then .imgPlayMp3.Picture = .imgPlayMp31.Picture
    If lObjectOver <> oPlayWav And .imgPlayWav.Picture <> .imgPlayWav1.Picture Then .imgPlayWav.Picture = .imgPlayWav1.Picture
    If lObjectOver <> oRip And .imgRip.Picture <> .imgRip1.Picture Then .imgRip.Picture = .imgRip1.Picture
    If lObjectOver <> oSkinEdit And .imgSkinEdit.Picture <> .imgSkinEdit1.Picture Then .imgSkinEdit.Picture = .imgSkinEdit1.Picture
    If lObjectOver <> oStopButton And .imgStop.Picture <> .imgStop1.Picture Then .imgStop.Picture = .imgStop1.Picture
    If lObjectOver <> oStopEncoding And .imgStopEncoding.Picture <> .imgStopEncoding1.Picture Then .imgStopEncoding.Picture = .imgStopEncoding1.Picture
End With
If Err.Number <> 0 Then SetError "SetFileLabels()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub SetFileLabels(Optional lWavPath As String, Optional lWavFile As String, Optional lMp3File As String, Optional lMp3Path As String)
'On Local Error Resume Next

lPlayer.pLabels.lMp3File = lMp3File
lPlayer.pLabels.lMp3Path = lMp3Path
lPlayer.pLabels.lWavFile = lWavFile
lPlayer.pLabels.lWavPath = lWavPath

If lPlayer.pStatus = sPlaying Then Exit Sub
If lEvents.eEncoderBusy = True Then Exit Sub
If lEvents.eRipperBusy = True Then Exit Sub

If Len(lWavFile) <> 0 Then
    frmMain.lblWavFile.Caption = lWavFile
End If
If Len(lMp3File) <> 0 Then
    frmMain.lblMp3File.Caption = lMp3File
End If

If Err.Number <> 0 Then SetError "SetFileLabels()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub StopAllEvents()
'On Local Error Resume Next
Dim i As Integer

For i = 1 To 100
    RemoveEvent i
Next i
lEvents.eEventCount = 0

If Err.Number <> 0 Then SetError "StopAllEvents()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub DragDrop(lData As DataObject)
'On Local Error Resume Next
Dim i As Integer, c As Integer, msg As String, msg2 As String

c = lData.Files.Count
If c <> 0 Then
    For i = 1 To c
        If Len(lData.Files(i)) <> 0 Then
            If LCase(Right(lData.Files(i), 3)) = "mp3" Then
                msg = lData.Files(i)
                If DoesFileExist(msg) = True Then
                    msg2 = msg
                    msg2 = GetFileTitle(msg2)
                    msg = Left(msg, Len(msg) - Len(msg2))
                    AddEvent Play, msg, msg2, "", "", 0, ""
                End If
            ElseIf LCase(Right(lData.Files(i), 3)) = "wav" Then
                msg = lData.Files(i)
                If DoesFileExist(msg) = True Then
                    msg2 = msg
                    msg2 = GetFileTitle(msg2)
                    msg = Left(msg, Len(msg) - Len(msg2))
                    SetFileLabels msg, msg2
                    OpenEffects msg & msg2
                End If
            End If
        End If
    Next i
End If

If Err.Number <> 0 Then SetError "DragDrop()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub CheckMouseOver()
'On Local Error Resume Next

With frmMain
    If .imgPlay.Picture <> .imgPlay1.Picture Then .imgPlay.Picture = .imgPlay1.Picture
    If .imgStop.Picture <> .imgStop1.Picture Then .imgStop.Picture = .imgStop1.Picture
    If .imgBackward.Picture <> .imgBackward1.Picture Then .imgBackward.Picture = .imgBackward1.Picture
    If .imgForward.Picture <> .imgForward1.Picture Then .imgForward.Picture = .imgForward1.Picture
    If .imgRip.Picture <> .imgRip1.Picture Then .imgRip.Picture = .imgRip1.Picture
    If .imgEncode.Picture <> .imgEncode1.Picture Then .imgEncode.Picture = .imgEncode1.Picture
    If .imgEnd.Picture <> .imgEnd1.Picture Then .imgEnd.Picture = .imgEnd1.Picture
    If .imgMinimize.Picture <> .imgMinimize1.Picture Then .imgMinimize.Picture = .imgMinimize1.Picture
    If .imgId3.Picture <> .imgId31.Picture Then .imgId3.Picture = .imgId31.Picture
    If .imgStopEncoding.Picture <> .imgStopEncoding1.Picture Then .imgStopEncoding.Picture = .imgStopEncoding1.Picture
    If .imgPlayMp3.Picture <> .imgPlayMp31.Picture Then .imgPlayMp3.Picture = .imgPlayMp31.Picture
    If .imgPlayWav.Picture <> .imgPlayWav1.Picture Then .imgPlayWav.Picture = .imgPlayWav1.Picture
    If .imgOptions.Picture <> .imgOptions1.Picture Then .imgOptions.Picture = .imgOptions1.Picture
    If .imgNexMedia.Picture <> .imgNexMedia1.Picture Then .imgNexMedia.Picture = .imgNexMedia1.Picture
    If .imgSkinEdit.Picture <> .imgSkinEdit1.Picture Then .imgSkinEdit.Picture = .imgSkinEdit1.Picture
    If .imgCancelRip.Picture <> .imgStopRipping1.Picture Then .imgCancelRip.Picture = .imgStopRipping1.Picture
End With
If Err.Number <> 0 Then SetError "CheckMouseOver()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub UnloadMain()
'On Local Error Resume Next

If DoesFileExist(lIniFiles.iUpdate) = True Then Kill lIniFiles.iUpdate
lEvents.eSettings.iEnding = True
Call Shell_NotifyIcon(NIM_DELETE, try)
frmMain.ns4Effects.sTop
If lEvents.eRipperBusy = True Then frmMain.Ripper.sTop
If lEvents.eEncoderBusy = True Then
    If lPlayer.pStatus = sPlaying Or lPlayer.pStatus = sPaused Then
        StopMp3
    Else
        frmMain.Encoder.sTop
    End If
End If
DoEvents
End
End Sub

Public Sub FillListboxWithDrives(lListbox As ListBox)
'On Local Error Resume Next
Dim i As Integer

If lDrives.dCount <> 0 Then
    For i = 1 To lDrives.dCount
        lListbox.AddItem lDrives.dDrive(i).dLetter
    Next i
Else
    lListbox.AddItem "<Select..>"
End If

If Err.Number <> 0 Then SetError "FillListBoxWithDrives()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub FillComboWithDrives(lComboBox As ComboBox)
'On Local Error Resume Next
Dim i As Integer

If lDrives.dCount <> 0 Then
    For i = 1 To lDrives.dCount
        lComboBox.AddItem lDrives.dDrive(i).dLetter
    Next i
Else
    lComboBox.AddItem "<Select..>"
End If

If Err.Number <> 0 Then SetError "FillListBoxWithDrives()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub AlwaysOnTop(myfrm As Form, SetOnTop As Boolean)
'On Local Error Resume Next
Dim lFlag As Integer

If SetOnTop Then
    lFlag = HWND_TOPMOST
Else
    lFlag = HWND_NOTOPMOST
End If
SetWindowPos myfrm.hwnd, lFlag, myfrm.Left / Screen.TwipsPerPixelX, myfrm.Top / Screen.TwipsPerPixelY, myfrm.Width / Screen.TwipsPerPixelX, myfrm.Height / Screen.TwipsPerPixelY, SWP_NOACTIVATE Or SWP_SHOWWINDOW

If Err.Number <> 0 Then SetError "AlwaysOnTop", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub WindowSize(lType As eWindowSize, lForm As Form)
'On Local Error Resume Next
Dim msg As String

msg = lForm.Name
If Len(msg) <> 0 Then
    If lType = wLoading Then
        lForm.Width = ReadINI(lIniFiles.iWindowPos, msg, "Width", lForm.Width)
        lForm.Height = ReadINI(lIniFiles.iWindowPos, msg, "Height", lForm.Height)
    Else
        WriteINI lIniFiles.iWindowPos, msg, "Width", lForm.Width
        WriteINI lIniFiles.iWindowPos, msg, "Height", lForm.Height
    End If
End If

If Err.Number <> 0 Then SetError "ResizeWindow", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub ShowMainWindow()
'On Local Error Resume Next

If lPlayers.pCDPlayerIndex = 0 And lPlayers.pMp3PlayerIndex = 0 Then
    ShowWizard
    If lEvents.eSettings.iShowAbout = True Then
        frmAbout.tmrDots.Enabled = False
        frmAbout.tmrUnload.Enabled = False
        frmAbout.Visible = False
    End If
    Exit Sub
Else
    frmMain.Show
    If lEvents.eSettings.iCheckForActiveWindow = True Then
        frmMain.tmrCheckActive.Enabled = True
    Else
        frmMain.tmrCheckActive.Enabled = False
    End If
End If
If Err.Number <> 0 Then SetError "ShowMainWindow", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub SetCaption(lStatus As eEventTypes, Optional lProgress As Long, Optional lFile As String)
'On Local Error Resume Next
Dim msg As String
If lProgress <> 0 And Len(lFile) <> 0 Then
    Select Case lStatus
    Case Play
        msg = "NS4 - Play: " & lFile & " " & lProgress & "%"
    Case Encode
        msg = "NS4 - Encode: " & lFile & " " & lProgress & "%"
    Case Rip
        msg = "NS4 - Rip: " & lFile & " " & lProgress & "%"
    End Select
Else
    If lEvents.eRegistered = True Then
        msg = "NexENCODE Studio v4." & App.Minor
    Else
        msg = "NexENCODE Studio v4." & App.Minor & " (Unregistered)"
    End If
End If
frmMain.Caption = msg
If Err.Number <> 0 Then SetError "SetCaption", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub UpdateASPI(Optional lDoNotPrompt As Boolean)
'On Local Error Resume Next
Dim msg As String
Dim i As Integer
If lDoNotPrompt = False Then
    lUnloadSetupWizardAfterASPI = True
    frmSetupWizard.Show
    Exit Sub
Else
    msg = vbYes
End If
If msg = vbYes Then
    If DoesFileExist(App.Path & "\programs\aspiupd.exe") = True Then
        Shell App.Path & "\programs\aspiupd.exe", vbNormalFocus
        End
    Else
        If lEvents.eSettings.iOverwritePrompts = True Then MsgBox "Unable to locate the WINASPI update. This file can be downloaded off of the Team Nexgen Website. Try one of the following urls..." & vbCrLf & "http://team-nexgen.org" & vbCrLf & "http://www.team-nexgen.org"
    End If
ElseIf msg = vbNo Then
    ToggleButtons oIdle
End If
If Err.Number <> 0 Then SetError "UpdateASPI", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub SetWait(lDescription As String, lExtended As String)
'On Local Error Resume Next

If lEvents.eSettings.iFreeDB.cShowDialog = False Then Exit Sub
frmWait.lblDescription.Caption = lDescription
frmWait.lblExtended.Caption = lExtended
End Sub

Public Sub ShowWait(lDescription As String, lExtended As String)
'On Local Error Resume Next
If lEvents.eSettings.iFreeDB.cShowDialog = False Then Exit Sub
AlwaysOnTop frmWait, True
frmWait.lblDescription.Caption = lDescription
frmWait.lblExtended.Caption = lExtended
frmWait.Show: DoEvents
End Sub

Public Sub PreLoadSettings()
'On Local Error Resume Next
lEvents.eSettings.iErrDescription = "An error was raised internally"
With lIniFiles
    .iEffects = App.Path & "\settings\effects.ini"
    .iWindowPos = App.Path & "\settings\nexwind.ini"
    .iErrors = App.Path & "\settings\nexerrors.ini"
    .iSettings = App.Path & "\settings\nexsettings.ini"
    .iPlayers = App.Path & "\settings\nexplayers.ini"
    .iPlaylists = App.Path & "\settings\nexplaylists.ini"
    .iCD = App.Path & "\settings\cd.ini"
    .iCDDBServers = App.Path & "\settings\cddbserv.ini"
    .iUpdate = App.Path & "\settings\ns4update.txt"
End With
lEvents.eName = ReadINI(lIniFiles.iSettings, "Settings", "Name", "")
lEvents.ePassword = ReadINI(lIniFiles.iSettings, "Settings", "Password", "")
lEvents.eSettings.iPlayWavs = ReadINI(lIniFiles.iSettings, "Settings", "PlayWavs", False)
lEvents.eSettings.iShowAbout = ReadINI(lIniFiles.iSettings, "Settings", "ShowAbout", True)
lEvents.eSettings.iRememberWindowSizes = ReadINI(lIniFiles.iSettings, "Settings", "RememberWindowSizes", True)
If Err.Number <> 0 Then SetError "PreLoadSettings", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub ShowEditDisc()
'On Local Error Resume Next
Dim msg As Boolean

If lEvents.eSettings.iFreeDB.cEnabled = True And lEvents.eSettings.iFreeDB.cAutoSubmit = True Then
    frmEditTracks.Show
End If

If Err.Number <> 0 Then SetError "ShowEditDisc", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Function GetSimpleTracks()
'On Local Error Resume Next
Dim i As Integer, X As Integer, Z As Integer, h As Boolean, msg As String

SelectCDDrive
frmMain.Ripper.Init: DoEvents
frmMain.Ripper.OpenDriveByLetter lRipperSettings.eDriveLetter
DoEvents
Z = frmMain.Ripper.TrackCount

lTracks.tCount = Z
If Z <> 0 Then
    lRipperSettings.eAvailable = True
    InitMediaToc lRipperSettings.eDriveLetter
    msg = GetTOC
    lRipperSettings.eDiscID = msg
    If Len(lRipperSettings.eDiscID) <> 0 Then
        If GetCDTracks(lRipperSettings.eDiscID) = False Then
            If lEvents.eSettings.iFreeDB.cEnabled = True Then
                frmMain.wskFreeDB.Close
                frmMain.wskFreeDB.LocalPort = GetRnd(1500)
                frmMain.wskFreeDB.Connect lEvents.eSettings.iFreeDB.cServer, 8880
            End If
        Else
            frmMain.lblInfo.Caption = lTracks.tArtist & " \ " & lTracks.tTitle
            frmMain.mnuArtist.Caption = "Artist: " & lTracks.tArtist
            frmMain.mnuAlbum.Caption = "Album: " & lTracks.tTitle
            frmMain.mnuArtist.Enabled = True
            frmMain.mnuAlbum.Enabled = True
        End If
    End If
Else
    frmMain.mnuArtist.Caption = "Artist: <None>"
    frmMain.mnuAlbum.Caption = "Album: <None>"
    frmMain.mnuArtist.Enabled = False
    frmMain.mnuAlbum.Enabled = False
    lRipperSettings.eAvailable = False
End If
If Err.Number <> 0 Then SetError "GetSimpleTracks()", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Function ReturnDirCompliant(lText As String) As String
'On Local Error Resume Next
Again:
If InStr(lText, "/") Or InStr(lText, "\") Or InStr(lText, "*") Or InStr(lText, ":") Or InStr(lText, Chr(34)) Or InStr(lText, "<") Or InStr(lText, ">") Or InStr(lText, "|") Or InStr(lText, "?") Then
    If InStr(lText, "/") Then
        lText = Replace(lText, "/", "_")
    ElseIf InStr(lText, "\") Then
        lText = Replace(lText, "\", "_")
    ElseIf InStr(lText, "*") Then
        lText = Replace(lText, "*", "_")
    ElseIf InStr(lText, ":") Then
        lText = Replace(lText, ":", "_")
    ElseIf InStr(lText, Chr(34)) Then
        lText = Replace(lText, Chr(34), "_")
    ElseIf InStr(lText, "<") Then
        lText = Replace(lText, "<", "_")
    ElseIf InStr(lText, ">") Then
        lText = Replace(lText, ">", "_")
    ElseIf InStr(lText, "?") Then
        lText = Replace(lText, "?", "_")
    ElseIf InStr(lText, "|") Then
        lText = Replace(lText, "|", "_")
    End If
Else
    ReturnDirCompliant = lText
    Exit Function
End If
GoTo Again
If Err.Number <> 0 Then SetError "ReturnDirCompliant()", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Sub MainButtonsEnabled(lEnabled As Boolean)
'On Local Error Resume Next
If lEnabled = True Then
    With frmMain
        .mnuEncode.Enabled = True
        .mnuRip.Enabled = True
        .mnuPlayer.Enabled = True
        .mnuNexMedia.Enabled = True
        .mnuAudica.Enabled = True
        .imgId3.Enabled = True
        .imgSkinEdit.Enabled = True
        .imgNexMedia.Enabled = True
        .imgOptions.Enabled = True
        .imgEncode.Enabled = True
        .imgCancelRip.Enabled = True
        .imgRip.Enabled = True
        .imgEnd.Enabled = True
        .imgMinimize.Enabled = True
        .imgStopEncoding.Enabled = True
        .imgPlayMp3.Enabled = True
        .imgPlayWav.Enabled = True
        .imgPlay.Enabled = True
        .imgStop.Enabled = True
        .imgForward.Enabled = True
        .imgBackward.Enabled = True
    End With
Else
    With frmMain
        .imgId3.Enabled = False
        .imgSkinEdit.Enabled = False
        .imgNexMedia.Enabled = False
        .imgOptions.Enabled = False
        .imgEncode.Enabled = False
        .imgCancelRip.Enabled = False
        .imgRip.Enabled = False
        .imgEnd.Enabled = False
        .imgMinimize.Enabled = False
        .imgStopEncoding.Enabled = False
        .imgPlayMp3.Enabled = False
        .imgPlayWav.Enabled = False
        .mnuEncode.Enabled = False
        .mnuRip.Enabled = False
        .mnuPlayer.Enabled = False
        .mnuNexMedia.Enabled = False
        .mnuAudica.Enabled = False
        .imgPlay.Enabled = False
        .imgStop.Enabled = False
        .imgForward.Enabled = False
        .imgBackward.Enabled = False
    End With
End If
If Err.Number <> 0 Then SetError "MainButtonsEnabled", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub ToggleButtons(lType As eObjectTypes)
'On Local Error Resume Next
With frmMain
    .imgId3.Visible = True
    .imgSkinEdit.Visible = True
    .imgNexMedia.Visible = True
    .imgOptions.Visible = True
    .imgEncode.Visible = True
    .imgCancelRip.Visible = True
    .imgRip.Visible = True
    .imgEnd.Visible = True
    .imgMinimize.Visible = True
    .imgStopEncoding.Visible = True
    .imgPlayMp3.Visible = True
    .imgPlayWav.Visible = True
    .mnuPlayMpeg.Enabled = True
    .mnuPauseMpeg.Enabled = True
    .mnuOpenMpeg.Enabled = True
    .mnuStopMpeg.Enabled = True
    .mnuForwardMpeg.Enabled = True
    .mnuBackwardMpeg.Enabled = True
    .mnuRandom2.Enabled = True
    .mnuOpenWave.Enabled = True
    .mnuEffectsEditor2.Enabled = True
    .mnuConvertToMp3.Enabled = True
    .mnuConvertToWav.Enabled = True
    .mnuOpenInSoundRec.Enabled = True
    .mnuDelete.Enabled = True
    .mnuCloseWav.Enabled = True
    .mnuSettings.Enabled = True
    .mnuMore.Enabled = True
    .mnuBatch.Enabled = True
    .mnuSettings.Enabled = True
    .mnuEncode.Enabled = True
    .mnuRip.Enabled = True
    .mnuPlayer.Enabled = True
    .mnuPlaylist.Enabled = True
    .mnuNexMedia.Enabled = True
    .mnuAudica.Enabled = True
    .mnuOpenFile.Enabled = True
    .imgOptions.Visible = True
    .imgSkinEdit.Visible = True
    If lEvents.eSettings.iPlayMp3sInNexENCODE = True Then
        .imgPlay.Visible = True
        .imgStop.Visible = True
        .imgForward.Visible = True
        .imgBackward.Visible = True
    Else
        .imgPlay.Visible = False
        .imgStop.Visible = False
        .imgForward.Visible = False
        .imgBackward.Visible = False
    End If
    Select Case lType

    Case oCDDB
        .imgStopEncoding.Visible = False
        .imgEncode.Visible = False
        .imgRip.Visible = False
        .imgPlayMp3.Visible = False
        .imgCancelRip.Visible = False
        .imgPlayWav.Visible = False
        .imgNexMedia.Visible = False
        .imgId3.Visible = False
        .imgPlay.Visible = False
        .imgStop.Visible = False
        .imgForward.Visible = False
        .imgBackward.Visible = False
        .imgOptions.Visible = False
        .imgSkinEdit.Visible = False
        .mnuOpenFile.Enabled = False
        .mnuEncode.Enabled = False
        .mnuRip.Enabled = False
        .mnuPlayer.Enabled = False
        .mnuNexMedia.Enabled = False
        .mnuAudica.Enabled = False
        .mnuPlaylist.Enabled = False
        .mnuMore.Enabled = False
        .mnuBatch.Enabled = False
        .mnuSettings.Enabled = False
        .mnuEffectsEditor2.Enabled = False
    Case oEncode
        If lEvents.ePlaylistVisible = True Then frmPlaylist.Hide
        .imgStopEncoding.Visible = True
        .imgEncode.Visible = False
        .imgRip.Visible = False
        .imgPlayMp3.Visible = False
        .imgCancelRip.Visible = False
        .imgPlayWav.Visible = False
        .imgNexMedia.Visible = False
        .imgId3.Visible = False
        .imgPlay.Visible = False
        .imgStop.Visible = False
        .imgForward.Visible = False
        .imgBackward.Visible = False
        .imgOptions.Visible = False
        .imgSkinEdit.Visible = False
        .mnuOpenFile.Enabled = False
        .mnuPlaylist.Checked = False
        .mnuEncode.Enabled = False
        .mnuRip.Enabled = False
        .mnuPlayer.Enabled = False
        .mnuNexMedia.Enabled = False
        .mnuAudica.Enabled = False
        .mnuPlaylist.Enabled = False
        .mnuMore.Enabled = False
        .mnuBatch.Enabled = False
        .mnuSettings.Enabled = False
        .mnuPlayMpeg.Enabled = False
        .mnuPauseMpeg.Enabled = False
        .mnuOpenMpeg.Enabled = False
        .mnuStopMpeg.Enabled = False
        .mnuForwardMpeg.Enabled = False
        .mnuBackwardMpeg.Enabled = False
        .mnuRandom2.Enabled = False
        .mnuConvertToWav.Enabled = False
        .mnuEffectsEditor2.Enabled = False
    Case oRip
        If lEvents.ePlaylistVisible = True Then frmPlaylist.Hide
        .imgStopEncoding.Visible = False
        .imgEncode.Visible = False
        .imgRip.Visible = False
        .imgPlayMp3.Visible = False
        .imgCancelRip.Visible = True
        .imgPlayWav.Visible = False
        .imgOptions.Visible = False
        .imgSkinEdit.Visible = False
        .imgNexMedia.Visible = False
        .imgId3.Visible = False
        .imgPlay.Visible = False
        .imgStop.Visible = False
        .imgForward.Visible = False
        .imgBackward.Visible = False
        .mnuPlaylist.Checked = False
        .mnuOpenFile.Enabled = False
        .mnuPlaylist.Checked = False
        .mnuPlaylist.Enabled = False
        .mnuEncode.Enabled = False
        .mnuRip.Enabled = False
        .mnuPlayer.Enabled = False
        .mnuNexMedia.Enabled = False
        .mnuAudica.Enabled = False
        .mnuMore.Enabled = False
        .mnuBatch.Enabled = False
        .mnuSettings.Enabled = False
        .mnuOpenWave.Enabled = False
        .mnuConvertToMp3.Enabled = False
        .mnuOpenInSoundRec.Enabled = False
        .mnuDelete.Enabled = False
        .mnuCloseWav.Enabled = False
        .mnuEffectsEditor2.Enabled = False
    Case oStopRipping
        .imgStopEncoding.Visible = False
        .imgCancelRip.Visible = False
    Case oStopEncoding
        .imgStopEncoding.Visible = False
        .imgCancelRip.Visible = False
    Case oPlayMp3
        .imgStopEncoding.Visible = False
        .imgCancelRip.Visible = False
    Case oPlayWav
        .imgStopEncoding.Visible = False
        .imgCancelRip.Visible = False
    Case oEnd
        .imgStopEncoding.Visible = False
        .imgCancelRip.Visible = False
    Case oMinimize
        .imgStopEncoding.Visible = False
        .imgCancelRip.Visible = False
    Case oOptions
        .imgStopEncoding.Visible = False
        .imgCancelRip.Visible = False
    Case oIdle
        .imgStopEncoding.Visible = False
        .imgCancelRip.Visible = False
    Case oHideAll
        If lEvents.ePlaylistVisible = True Then frmPlaylist.Hide
        .mnuPlaylist.Checked = False
        .imgPlay.Visible = False
        .imgStop.Visible = False
        .imgBackward.Visible = False
        .imgForward.Visible = False
        .imgOptions.Visible = False
        .imgSkinEdit.Visible = False
        .imgEncode.Visible = False
        .imgCancelRip.Visible = False
        .imgRip.Visible = False
        .imgEnd.Visible = False
        .imgMinimize.Visible = False
        .imgStopEncoding.Visible = False
        .imgPlayMp3.Visible = False
        .imgPlayWav.Visible = False
        .imgSkinEdit.Visible = False
        .imgNexMedia.Visible = False
        .imgOptions.Visible = False
        .imgId3.Visible = False
        .mnuEncode.Enabled = False
        .mnuRip.Enabled = False
        .mnuPlayer.Enabled = False
        .mnuNexMedia.Enabled = False
        .mnuAudica.Enabled = False
        .mnuPlaylist.Enabled = False
        .mnuMore.Enabled = False
        .mnuBatch.Enabled = False
        .mnuSettings.Enabled = False
        .mnuOpenFile.Enabled = False
        .imgPlay.Visible = False
        .imgStop.Visible = False
        .imgForward.Visible = False
        .imgBackward.Visible = False
        .mnuPlayMpeg.Enabled = False
        .mnuPauseMpeg.Enabled = False
        .mnuOpenMpeg.Enabled = False
        .mnuStopMpeg.Enabled = False
        .mnuForwardMpeg.Enabled = False
        .mnuBackwardMpeg.Enabled = False
        .mnuRandom2.Enabled = False
        .mnuConvertToWav.Enabled = False
        .mnuConvertToMp3.Enabled = False
        .mnuEffectsEditor2.Enabled = False
    Case 26
        .mnuBatch.Enabled = False
        .mnuEncode.Enabled = False
        .mnuRip.Enabled = False
        .mnuOpenFile.Enabled = False
        .imgStopEncoding.Visible = True
        .imgEncode.Visible = False
        .imgOptions.Visible = False
        .imgSkinEdit.Visible = False
        .imgRip.Visible = False
        .imgPlayMp3.Visible = False
        .imgCancelRip.Visible = False
        .imgPlayWav.Visible = False
        .imgNexMedia.Visible = False
        .imgId3.Visible = False
        .mnuPlayMpeg.Enabled = False
        .mnuForwardMpeg.Enabled = False
        .mnuBackwardMpeg.Enabled = False
        .mnuRandom2.Enabled = False
        .mnuConvertToWav.Enabled = False
        .mnuOpenMpeg.Enabled = False
        .mnuEffectsEditor2.Enabled = False
    End Select
End With
If Err.Number <> 0 Then SetError "ToggleButtons()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub LoadSettings()
'On Local Error Resume Next
Dim i As Integer

lEvents.eGlobalPass = "pickles"
With lEvents.eSettings.iFreeDB
    .cShowDialog = ReadINI(lIniFiles.iSettings, "CDDB", "ShowDialog", False)
    .cSaveTracksToDisk = ReadINI(lIniFiles.iSettings, "CDDB", "SaveTracksToDisk", True)
    .cAutoSubmit = ReadINI(lIniFiles.iSettings, "CDDB", "AutoSubmit", False)
    .cEmailAddress = ReadINI(lIniFiles.iSettings, "CDDB", "EmailAddress", "guide_X@live.com")
    .cServer = ReadINI(lIniFiles.iSettings, "CDDB", "Server", "freedb.freedb.org")
    .cUseFirstMatch = ReadINI(lIniFiles.iSettings, "CDDB", "UseFirstMatch", True)
    .cEnabled = ReadINI(lIniFiles.iSettings, "CDDB", "Enabled", True)
End With
With lCDDBServ
    .cCount = ReadINI(lIniFiles.iCDDBServers, "Settings", "Count", 0)
    If .cCount <> 0 Then
        For i = 1 To .cCount
            lCDDBServ.cServer(i).sIp = ReadINI(lIniFiles.iCDDBServers, Str(i), "Ip", "")
            lCDDBServ.cServer(i).sLocation = ReadINI(lIniFiles.iCDDBServers, Str(i), "Location", "")
        Next i
    End If
End With
lEvents.eSettings.iAlwaysOnTop = ReadINI(lIniFiles.iSettings, "Settings", "AlwaysOnTop", False)
lEvents.eSettings.iAutoPlay = ReadINI(lIniFiles.iSettings, "Settings", "AutoPlay", "False")
lEvents.eSettings.iOverwritePrompts = ReadINI(lIniFiles.iSettings, "Settings", "OverwritePrompts", True)
lEvents.eSettings.iShowErrors = ReadINI(lIniFiles.iSettings, "Settings", "ShowErrors", False)
lEvents.eSettings.iShowReports = ReadINI(lIniFiles.iSettings, "Settings", "ShowReports", "True")
lEvents.eSettings.iPlayWavs = ReadINI(lIniFiles.iSettings, "Settings", "PlayWavs", True)
lEvents.eSettings.iPlayMp3sInNexENCODE = ReadINI(lIniFiles.iSettings, "Settings", "PlayMp3sInNexENCODE", True)
lEvents.eSettings.iUpdateCheck = ReadINI(lIniFiles.iSettings, "Settings", "UpdateCheck", True)
lEvents.eSettings.iCheckForActiveWindow = ReadINI(lIniFiles.iSettings, "Settings", "CheckForActiveWindow", True)
lEvents.eSettings.iCreateAlbumFileOnEncode = ReadINI(lIniFiles.iSettings, "Settings", "CreateAlbumFileOnEncode", False)
lEvents.eErrCount = ReadINI(lIniFiles.iErrors, "Settings", "Count", 0)
lPlayers.pCount = ReadINI(lIniFiles.iPlayers, "Settings", "Count", 0)
For i = 1 To lPlayers.pCount
    lPlayers.pPlayer(i).pFile = ReadINI(lIniFiles.iPlayers, Str(i), "File", "")
    lPlayers.pPlayer(i).pName = ReadINI(lIniFiles.iPlayers, Str(i), "Name", "")
    lPlayers.pPlayer(i).pPath = ReadINI(lIniFiles.iPlayers, Str(i), "Path", "")
    lPlayers.pPlayer(i).pPlaylist = ReadINI(lIniFiles.iPlayers, Str(i), "Playlist", "")
    lPlayers.pPlayer(i).pType = ReadINI(lIniFiles.iPlayers, Str(i), "Type", 0)
    If UCase(lPlayers.pPlayer(i).pPath) = "NS4PATH" Then lPlayers.pPlayer(i).pPath = App.Path
Next i
lRipperSettings.eDriveLetter = ReadINI(lIniFiles.iSettings, "Settings", "DriveLetter", "")
lRipperSettings.eAutoDeleteRipedFiles = ReadINI(lIniFiles.iSettings, "Settings", "AutoDeleteRipedFiles", "True")
lRipperSettings.eLockCDTrayDuringRip = ReadINI(lIniFiles.iSettings, "Settings", "LockCDTrayDuringRip", False)
lRipperSettings.eCopyMode = ReadINI(lIniFiles.iSettings, "Settings", "CopyMode", 2)
lPlayer.pContinuous = ReadINI(lIniFiles.iPlayers, "Settings", "Continuous", True)
lPlayer.pPlayerType = pVFMp3Player
lPlayers.pMp3PlayerIndex = ReadINI(lIniFiles.iPlayers, "Settings", "Mp3Player", 0)
lPlayers.pCDPlayerIndex = ReadINI(lIniFiles.iPlayers, "Settings", "CDPlayer", 0)
lEncoderSettings.eOutputDir = ReadINI(lIniFiles.iSettings, "Settings", "OutputDir", App.Path & "\library\")
If Right(lEncoderSettings.eOutputDir, 1) <> "\" Then lEncoderSettings.eOutputDir = lEncoderSettings.eOutputDir & "\"
lEncoderSettings.eAutoAddTags = ReadINI(lIniFiles.iSettings, "Settings", "AutoAddTags", True)
lEncoderSettings.eDownmix = ReadINI(lIniFiles.iSettings, "Settings", "Downmix", False)
lEncoderSettings.eBitrate = ReadINI(lIniFiles.iSettings, "Settings", "Bitrate", 192)
lEncoderSettings.eSampleRate = ReadINI(lIniFiles.iSettings, "Settings", "SampleRate", 44100)
lEncoderSettings.eDownsample = ReadINI(lIniFiles.iSettings, "Settings", "Downsample", True)
lEncoderSettings.eCopyrighted = ReadINI(lIniFiles.iSettings, "Settings", "Copyrighted", "False")
lEncoderSettings.eOrigionalWork = ReadINI(lIniFiles.iSettings, "Settings", "Origional", "False")
lEncoderSettings.eProfile = ReadINI(lIniFiles.iSettings, "Settings", "Profile", 3)

ShowMainWindow

If Err.Number <> 0 Then SetError "LoadRipperSettings()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub FindOpenDevice(lPlayer As Mp3Play)
'On Local Error Resume Next
Dim i As Integer, p As Integer, b As Integer

With lPlayer
    p = .GetNumDevs
    For i = 0 To p
        b = .SetOutDevice(i)
        If b = 0 Then Exit For
    Next i
End With

If Err.Number <> 0 Then SetError "FindOpenDevice", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Function GetCDTracks(lToc As String) As Boolean
'On Local Error Resume Next
Dim i As Integer
If Len(lToc) <> 0 Then
    If ReadINI(lIniFiles.iCD, lToc, "Enabled", False) = True Then
        lTracks.tDiscLen = ReadINI(lIniFiles.iCD, lToc, "DiscLen", "")
        lTracks.tArtist = ReadINI(lIniFiles.iCD, lToc, "Artist", "")
        lTracks.tTitle = ReadINI(lIniFiles.iCD, lToc, "Title", "")
        lTracks.tGenre = ReadINI(lIniFiles.iCD, lToc, "Genre", "")
        lTracks.tLabel = ReadINI(lIniFiles.iCD, lToc, "Label", "")
        lTracks.tYear = ReadINI(lIniFiles.iCD, lToc, "Year", "")
        If lTracks.tCount > 300 Then lTracks.tCount = 300
        If Len(lTracks.tArtist) <> 0 And Len(lTracks.tTitle) <> 0 Then
            For i = 1 To lTracks.tCount
                lTracks.tTrack(i).tName = ReadINI(lIniFiles.iCD, lToc, Str(i), "")
                lTracks.tTrack(i).tLength = ReadINI(lIniFiles.iCD, lToc, Str(i) & "L", "")
            Next i
        End If
        GetCDTracks = True
    Else
        GetCDTracks = False
        Exit Function
    End If
End If
If Err.Number <> 0 Then SetError "LoadRipperSettings()", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Sub SaveCDTracks(lToc As String)
'On Local Error Resume Next
Dim i As Integer
If Len(lToc) <> 0 Then
    WriteINI lIniFiles.iCD, lToc, "DiscLen", lTracks.tDiscLen
    WriteINI lIniFiles.iCD, lToc, "Enabled", True
    WriteINI lIniFiles.iCD, lToc, "Artist", lTracks.tArtist
    WriteINI lIniFiles.iCD, lToc, "Title", lTracks.tTitle
    WriteINI lIniFiles.iCD, lToc, "Genre", lTracks.tGenre
    WriteINI lIniFiles.iCD, lToc, "Label", lTracks.tLabel
    WriteINI lIniFiles.iCD, lToc, "Year", lTracks.tYear
    For i = 1 To lTracks.tCount
        WriteINI lIniFiles.iCD, lToc, Str(i), lTracks.tTrack(i).tName
        WriteINI lIniFiles.iCD, lToc, Str(i) & "L", lTracks.tTrack(i).tLength
    Next i
Else
    Exit Sub
End If
If Err.Number <> 0 Then SetError "SaveCDTracks()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub Surf(lUrl As String)
'On Local Error Resume Next
Dim msg As Long
msg = ShellExecute(frmMain.hwnd, vbNullString, lUrl, vbNullString, "c:\", SW_SHOWNORMAL)
If Err.Number <> 0 Then SetError "Surf()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub SetError(lSub As String, lDescription As String, lExtended As String)
'On Local Error Resume Next
Dim k As Integer
k = lEvents.eErrCount + 1
lEvents.eErrCount = k
WriteINI lIniFiles.iErrors, "Settings", "Count", Str(k)
WriteINI lIniFiles.iErrors, Str(k), "Sub", lSub
WriteINI lIniFiles.iErrors, Str(k), "Description", lDescription
WriteINI lIniFiles.iErrors, Str(k), "Extended", lExtended
If lEvents.eSettings.iShowErrors = False Then Exit Sub
frmErrHandler.txtInfo.Text = lExtended
frmErrHandler.lblSub.Caption = "Sub or function: " & lSub
frmErrHandler.lblDescription.Caption = "Description: " & lDescription
frmErrHandler.WindowState = vbNormal
frmErrHandler.Visible = True
End Sub

Public Sub GoMp3Player(Optional lFilename As String)
'On Local Error Resume Next
Dim msg As String, i As Integer
lFilename = Trim(lFilename)
If lEvents.eSettings.iPlayMp3sInNexENCODE = True Then
    msg = LCase(Right(lFilename, 3))
    Select Case msg
    Case "mp3"
        msg = lFilename
        i = FindMediaIndex(GetFileTitle(msg))
        If i <> 0 Then PlayMp3 i
        Exit Sub
    Case "m3u"
        Dim lFile As String, lPath As String
        lFile = lFilename
        lFile = GetFileTitle(lFilename)
        lPath = Left(lFilename, Len(lFilename) - Len(lFile))
        PlayPlaylist lPath, lFile, FindPlaylistIndexByFile(lFile)
        Exit Sub
    End Select
End If
msg = lPlayers.pPlayer(lPlayers.pMp3PlayerIndex).pPath & lPlayers.pPlayer(lPlayers.pMp3PlayerIndex).pFile
If Len(msg) <> 0 And DoesFileExist(msg) = True Then
    If Len(lFilename) <> 0 Then
        If DoesFileExist(lFilename) = True Then
            Shell msg & " " & Chr(34) & lFilename & Chr(34), vbNormalFocus
        Else
            Shell msg, vbNormalFocus
        End If
    Else
        Shell msg, vbNormalFocus
    End If
    ToggleButtons oIdle
Else
    SetError "GoMp3Player()", "Null Value", "Unable to find Mp3 player"
End If
If Err.Number <> 0 Then SetError "GoAudica()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub LoadId3Editor()
'On Local Error Resume Next
frmMP3Info.Show
If Err.Number <> 0 Then SetError "LoadId3Editor()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub ShowText(lFiletype As String, lFilename As String)
'On Local Error Resume Next
Dim msg As String
If Len(lFilename) <> 0 Then
    msg = ReadFile(lFilename)
    If Len(msg) <> 0 Then
        frmTextViewer.txtInformation = msg
    Else
        Shell "notepad " & lFilename, vbNormalFocus
    End If
End If
If Err.Number <> 0 Then SetError "ShowText()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub GoCDPlayer()
'On Local Error Resume Next
Dim msg As String, msg2 As VbMsgBoxResult
If lEvents.eSettings.iOverwritePrompts = True Then
    msg2 = MsgBox("Loading your cd player will require you to shut down NexENCODE, are you sure you want to do this?", vbYesNo + vbQuestion)
    If msg2 = vbNo Then Exit Sub
    msg = lPlayers.pPlayer(lPlayers.pCDPlayerIndex).pPath & lPlayers.pPlayer(lPlayers.pCDPlayerIndex).pFile
    If DoesFileExist(msg) = True Then
        Shell msg, vbNormalFocus
        End
    Else
        MsgBox "Unable to find CD Player", vbExclamation
    End If
Else
    msg = lPlayers.pPlayer(lPlayers.pCDPlayerIndex).pPath & lPlayers.pPlayer(lPlayers.pCDPlayerIndex).pFile
    If DoesFileExist(msg) = True Then
        Shell msg, vbNormalFocus
        End
    End If
End If
If Err.Number <> 0 Then SetError "GoCDPlayer()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub ShowMp3PlayError(lError As Integer)
'On Local Error Resume Next
Dim msg As String
If lError <> 0 Then
    Select Case lError
    Case 714
        msg = "ErrWrongPathSpecification"
    Case 643
        msg = "Cannot read mpeg header"
    Case 712
        msg = "Free Disc Space Error, try decoding to different drive or partition"
    Case 713
        msg = "Device not found"
    End Select
    If lEvents.eSettings.iOverwritePrompts = True Then MsgBox "The MP3 Decoder raised an error of # " & lError & vbCrLf & msg, vbExclamation
End If
If Err.Number <> 0 Then SetError "ShowMp3PlayError()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub LoadTrackGet(lCDAToMP3 As Boolean)
'On Local Error Resume Next
Again:
GetSimpleTracks
DoEvents
If lRipperSettings.eAvailable = False Then
    frmInsertDisc.Show 1
    Exit Sub
ElseIf lTracks.tCount = 0 Then
ElseIf lTracks.tCount = 1 And frmMain.Ripper.TrackIsAudio(1) = False Then
    frmInsertDisc.Show 1
Else
    frmTrackGet.Show
    If lCDAToMP3 = True Then
        frmTrackGet.cboFormat.ListIndex = 0
    Else
        frmTrackGet.cboFormat.ListIndex = 1
    End If
End If
If Err.Number <> 0 Then SetError "LoadTrackGet()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub FadeOut(lHwnd As Long)
'On Local Error Resume Next
Dim X As Integer, i As Integer
X = 100
PlayWav App.Path & "\media\minimize.wav", SND_ASYNC
For i = 1 To 5
    X = X - 20
    MakeTransparent lHwnd, X
    DoEvents
Next i
If Err.Number <> 0 Then SetError "FadeOut()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub FadeIn(lHwnd As Long)
'On Local Error Resume Next
Dim i As Integer, X As Integer
X = 0
PlayWav App.Path & "\media\maximize.wav", SND_ASYNC
For i = 1 To 5
    X = X + 20
    MakeTransparent lHwnd, X
    DoEvents
Next i
MakeOpaque lHwnd
If Err.Number <> 0 Then SetError "FadeIn()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub pause(interval)
'On Local Error Resume Next
Dim Current

Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub

Public Sub ShowWizard()
'On Local Error Resume Next
frmSetupWizard.Show
End Sub

Public Sub ResetButtons()
'On Local Error Resume Next
With frmMain
    .imgRip.Picture = Nothing
    .imgRip1.Picture = Nothing
    .imgRip2.Picture = Nothing
    .imgCancelRip.Picture = Nothing
    .imgStopEncoding.Picture = Nothing
    .imgStopEncoding1.Picture = Nothing
    .imgStopEncoding2.Picture = Nothing
    .imgEncode.Picture = Nothing
    .imgEncode1.Picture = Nothing
    .imgEncode2.Picture = Nothing
    .imgEncode1.Picture = Nothing
    .imgEnd.Picture = Nothing
    .imgEnd1.Picture = Nothing
    .imgEnd2.Picture = Nothing
    .imgMinimize.Picture = Nothing
    .imgMinimize1.Picture = Nothing
    .imgMinimize2.Picture = Nothing
    .imgId3.Picture = Nothing
    .imgId31.Picture = Nothing
    .imgId32.Picture = Nothing
    .imgSkinEdit.Picture = Nothing
    .imgSkinEdit1.Picture = Nothing
    .imgSkinEdit2.Picture = Nothing
    .imgPlayMp3.Picture = Nothing
    .imgPlayMp31.Picture = Nothing
    .imgPlayWav.Picture = Nothing
    .imgPlayWav2.Picture = Nothing
    .imgPlayWav1.Picture = Nothing
    .imgPlay.Picture = Nothing
    .imgPlay1.Picture = Nothing
    .imgPlay2.Picture = Nothing
    .imgStop.Picture = Nothing
    .imgStop1.Picture = Nothing
    .imgStop2.Picture = Nothing
End With
If Err.Number <> 0 Then SetError "ResetButtons()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub RipCircleEffect(lPercent As Integer)
'On Local Error Resume Next
Dim i As Integer, m As Integer
With frmMain
    For i = 0 To 9
        m = m + 10
        If lPercent + 1 > m Then
            If .shpRipper(i).BackColor <> vbWhite Then .shpRipper(i).BackColor = vbWhite
        Else
            .shpRipper(i).BackColor = .shpRipperColor.BackColor
        End If
    Next i
End With
If Err.Number <> 0 Then SetError "RipCircleEffect()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub EncodeCircleEffect(lPercent As Integer)
'On Local Error Resume Next
Dim i As Integer, m As Integer
With frmMain
    For i = 0 To 9
        m = m + 10
        If lPercent + 1 > m Then
            If .shpEncoder(i).BackColor <> vbWhite Then .shpEncoder(i).BackColor = vbWhite
        Else
            .shpEncoder(i).BackColor = .shpEncoderColor.BackColor
        End If
    Next i
End With
If Err.Number <> 0 Then SetError "EncodeCircleEffect()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub FormDrag(lFormname As Form)
'On Local Error Resume Next
ReleaseCapture
Call SendMessage(lFormname.hwnd, &HA1, 2, 0&)
If Err.Number <> 0 Then SetError "FormDrag()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub AddEvent(lEventType As eEventTypes, lInputFilePath As String, lInputFilename As String, lOutputFilepath As String, lOutputFilename As String, lTrack As Integer, lExtendedParameters As String)
'On Local Error Resume Next
Dim i As Integer
lEvents.eEventCount = lEvents.eEventCount + 1
i = lEvents.eEventCount
If i > 100 Then Exit Sub
With lEvents.eEvent(i)
    .eInputFilepath = lInputFilePath
    .eInputFilename = lInputFilename
    .eOutputFilepath = lOutputFilepath
    .eOutputFilename = lOutputFilename
    If Len(.eOutputFilepath) <> 0 And Right(.eOutputFilepath, 1) <> "\" Then .eOutputFilepath = .eOutputFilepath & "\"
    If Len(.eInputFilepath) <> 0 And Right(.eInputFilepath, 1) <> "\" Then .eInputFilepath = .eInputFilepath & "\"
    If Len(lExtendedParameters) <> 0 Then
        If lExtendedParameters = "AUTODELETE" Then
            .eAutoDeleteThisWav = True
        Else
            .eAutoDeleteThisWav = False
        End If
    End If
    .eEnabled = True
    .eEventType = lEventType
    .eTrack = lTrack
    .eExtendedParameters = lExtendedParameters
End With
If lEvents.eRipperBusy = False And lEvents.eEncoderBusy = False And lPlayer.pStatus <> sPlaying Then
    ProcessNextEvent
End If
If Err.Number <> 0 Then SetError "AddEvent()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub EncodeFile(lInputWavFile As String, lOutputMp3File As String)
'On Local Error Resume Next
Dim msg As String, msg2 As String, lPath As String, lFile As String
If lPlayer.pStatus = sPlaying Then Exit Sub
If lEvents.eRipperBusy = True Then Exit Sub
If lEvents.eEncoderBusy = True Then Exit Sub
If DoesFileExist(lInputWavFile) = False Then Exit Sub
ResetEncoderCircles
frmMain.tmrShowEncoderCircles.interval = 40
frmMain.tmrShowEncoderCircles.Enabled = True
If Len(lInputWavFile) = 0 Then
    lInputWavFile = OpenDialog(frmMain, "Wav Audio Files (*.wav)|*.wav|All Files(*.*)|*.*", "Select .WAV File ...", CurDir)
    If Len(lInputWavFile) = 0 Then Exit Sub
End If
If Len(lOutputMp3File) = 0 Then
    lOutputMp3File = SaveDialog(frmMain, "Mpeg Layer 3 (*.mp3)|*.mp3|All Files(*.*)|*.*", "Save As ...", CurDir)
    If Len(lInputWavFile) = 0 Then Exit Sub
End If
lFile = lOutputMp3File
lFile = GetFileTitle(lFile)
lPath = Left(lOutputMp3File, Len(lOutputMp3File) - Len(lFile))
If CheckFileConstants(lPath, lOutputMp3File) = True Then
    With frmMain
        If lEncWizard.eEnabled = True Then frmEncoderWizard.lblFilename.Caption = lFile
        msg2 = lOutputMp3File
        msg2 = GetFileTitle(msg2)
        .lblWavFile.Caption = ""
        .lblMp3File.Caption = lFile
        .lblMp3File.ToolTipText = lFile
        .Encoder.BitRate = lEncoderSettings.eBitrate
        .Encoder.SampleRate = lEncoderSettings.eSampleRate
        .Encoder.AllowDownSample = lEncoderSettings.eDownsample
        AddFinishedEvent Encode, "Starting", lOutputMp3File
        .Encoder.Open lInputWavFile, lOutputMp3File
        .Encoder.Encode
        ToggleButtons oEncode
        DoEvents
    End With
End If
If Err.Number <> 0 Then SetError "EncodeFile()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub ProcessNextEvent()
'On Local Error Resume Next
Dim i As Integer, bad As Integer, f As Integer, msg As String, k As Integer
If lEvents.eEventBusy = True Then Exit Sub
If lEvents.eRipperBusy = True Then Exit Sub
If lEvents.eEncoderBusy = True Then Exit Sub
If lEvents.eSettings.iEnding = True Then Exit Sub
If lEvents.eEventCount = 0 Then Exit Sub
lEvents.eEventBusy = True
pause 1
For i = 1 To lEvents.eEventCount
    If lEvents.eEvent(i).eEnabled = True Then
        Select Case lEvents.eEvent(i).eEventType
        Case Play
            If Len(lEvents.eEvent(i).eInputFilename) <> 0 Then
                With lEvents.eEvent(i)
                    k = FindMediaIndex(.eInputFilename)
                    If k <> 0 Then
                        RemoveEvent i
                        PlayMp3 k
                        Exit For
                    Else
                        If AddToPlaylist(.eInputFilepath & .eInputFilename, 1) = True Then
                            k = FindMediaIndex(.eInputFilename)
                            If k <> 0 Then
                                StopMp3
                                DoEvents
                                RemoveEvent i
                                PlayMp3 k
                            End If
                        End If
                    End If
                End With
            Else
                RemoveEvent i
                ProcessNextEvent
            End If
        Case Merge
            For k = 1 To lReports.rCount
                If lReports.rReport(k).rType = Encode Then
                    f = f + 1
                End If
            Next k
            If f > 2 Then
                For k = 1 To lReports.rCount
                    frmFileMerger.AddToMergeList lReports.rReport(k).rFilepath & lReports.rReport(k).rFilename
                Next k
                DoEvents
                frmFileMerger.MergeFiles lEvents.eEvent(i).eOutputFilepath & lEvents.eEvent(i).eOutputFilename, True
                DoEvents
                RemoveEvent i
            End If
        Case Decode
            If Len(lEvents.eEvent(i).eInputFilename) <> 0 Then
                With lEvents.eEvent(i)
                    DecodeMPEG .eInputFilename, .eInputFilepath, .eOutputFilename
                    DoEvents
                    RemoveEvent i
                    Exit For
                End With
            End If
        Case Rip
            If lEvents.eEvent(i).eTrack <> 0 Then
                RecordCDTrack lEvents.eEvent(i).eTrack, lEvents.eEvent(i).eOutputFilepath & lEvents.eEvent(i).eOutputFilename
                DoEvents
                RemoveEvent i
                Exit For
            End If
        Case Encode
            If lEvents.eEvent(i).eAutoDeleteThisWav = True Then
                lEvents.eAutoDel = lEvents.eEvent(i).eInputFilepath & lEvents.eEvent(i).eInputFilename
            Else
                lEvents.eAutoDel = ""
            End If
            EncodeFile lEvents.eEvent(i).eInputFilepath & lEvents.eEvent(i).eInputFilename, lEvents.eEvent(i).eOutputFilepath & lEvents.eEvent(i).eOutputFilename
            msg = lEvents.eEvent(i).eOutputFilename
            msg = Left(msg, Len(msg) - 4)
            f = FindTrackIndex(msg)
            SetId3Info lEvents.eEvent(i).eOutputFilepath & lEvents.eEvent(i).eOutputFilename, lTracks.tTrack(f).tName, lTracks.tArtist, lTracks.tTitle, lTracks.tYear, lTracks.tGenre
            RemoveEvent i
            Exit For
        End Select
    Else
        bad = bad + 1
        If bad = lEvents.eEventCount Then
            If lEvents.eSettings.iShowReports = True Then
                If lReports.rCount <> 0 Then
                    frmReport.Show
                    lEvents.eEventCount = 0
                End If
            End If
            If lRipperSettings.eAutoEject = True Then frmMain.Ripper.Eject
        End If
    End If
Next i
lEvents.eEventBusy = False
If Err.Number <> 0 Then SetError "ProcessNextEvent()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub ResetRipperCircles()
'On Local Error Resume Next
Dim i As Integer
For i = 0 To 9
    frmMain.shpRipper(i).Visible = False
Next i
If Err.Number <> 0 Then SetError "ResetRipperCircles()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub ResetEncoderCircles()
'On Local Error Resume Next
Dim i As Integer
For i = 0 To 9
    frmMain.shpEncoder(i).Visible = False
Next i
lEvents.eCircleNum = 0
If Err.Number <> 0 Then SetError "ResetEncoderCircles()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub MakeDir(lDirectory As String)
'On Local Error Resume Next

If Len(lDirectory) <> 0 Then
    If Right(lDirectory, 1) <> "\" Then lDirectory = lDirectory & "\"
    MkDir lDirectory
End If

If Err.Number <> 0 Then SetError "MkDir()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub InitEffects()
'On Local Error Resume Next
With frmMain
    DisableEffects
    .mnuStopEffectWav.Enabled = True
    lEvents.eRipperBusy = True
    ToggleButtons oRip
    .tmrShowRipperCircles.Enabled = True
    pause 1
    DoEvents
    lEffectsPresets.eStatus = eAddingEffect
End With

If Err.Number <> 0 Then SetError "RecordCDTrack()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub RecordCDTrack(lTrack As Integer, lOutputFilename As String)
'On Local Error Resume Next
Dim lDir As String, lFile As String, msg As String

If lTrack = 0 Then Exit Sub
If Len(lOutputFilename) = 0 Then
    lOutputFilename = SaveDialog(frmMain, "Wav Files (*.wav)|*.wav|All Files (*.*)|*.*", "Save As ..", CurDir)
    If Len(lOutputFilename) = 0 Then Exit Sub
End If

msg = lOutputFilename
lFile = GetFileTitle(msg)
lDir = Left(lOutputFilename, Len(lOutputFilename) - Len(lFile))
If CheckFileConstants(lDir, lOutputFilename) = True Then
    With frmMain
        lEvents.eRipperBusy = True
        ToggleButtons oRip
        AddFinishedEvent Rip, "Starting", lOutputFilename
        .lblWavFile.Caption = lFile
        .tmrShowRipperCircles.Enabled = True
        If lRipperSettings.eLockCDTrayDuringRip = True Then .Ripper.LockTray
        .Ripper.SetCopyMode lRipperSettings.eCopyMode
        .Ripper.OpenDriveByLetter lRipperSettings.eDriveLetter
        DoEvents
        .Ripper.ReadTrack lTrack, lOutputFilename
    End With
End If

If Err.Number <> 0 Then SetError "RecordCDTrack()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub ClearReports()
'On Local Error Resume Next
Dim i As Integer

For i = 1 To lReports.rCount
    With lReports.rReport(i)
        .rReportString = ""
        .rFilename = ""
        .rFilepath = ""
        .rType = 0
    End With
Next i
lReports.rCount = 0
If Err.Number <> 0 Then SetError "ClearReports()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub AddReport(lReporttring As String, Optional lFilePath As String, Optional lFilename As String, Optional lType As eEventTypes)
'On Local Error Resume Next
Dim i As Integer

i = lReports.rCount + 1
lReports.rCount = i
With lReports.rReport(i)
    .rReportString = lReporttring
    .rType = lType
    .rFilename = lFilename
    .rFilepath = lFilePath
End With

If Err.Number <> 0 Then SetError "AddReport()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub AddFinishedEvent(lFunction As eEventTypes, lAction As String, Optional lFilename As String)
'On Local Error Resume Next
Dim msg As String, msg2 As String
If lFunction = Play Then
    Exit Sub
End If
If lEvents.eSettings.iShowReports = False Then Exit Sub
Select Case lFunction
Case Encode
    If Len(lFilename) <> 0 Then
        msg = lFilename
        msg2 = GetFileTitle(msg)
        msg = Left(lFilename, Len(lFilename) - Len(msg2))
        AddReport "> Encode (" & msg2 & ") " & lAction, msg, msg2, Encode
    Else
        AddReport "> Encode " & lAction
    End If
Case Rip
    If Len(lFilename) <> 0 Then
        msg = lFilename
        msg = GetFileTitle(msg)
        AddReport "> Rip (" & msg & ") " & lAction
    Else
        AddReport "> Rip " & lAction
    End If
Case Decode
    If Len(lFilename) <> 0 Then
        msg = lFilename
        msg = GetFileTitle(msg)
        AddReport "> Decode (" & msg & ") " & lAction
    Else
        AddReport "> Decode " & lAction
    End If
End Select
If Err.Number <> 0 Then SetError "AddFinishedEvent()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub RemoveEvent(lIndex As Integer)
'On Local Error Resume Next

With lEvents.eEvent(lIndex)
    .eEnabled = False
    .eEventType = 0
    .eInputFilename = vbNullString
    .eOutputFilename = vbNullString
    .eInputFilepath = vbNullString
    .eOutputFilepath = vbNullString
    .eTrack = 0
    .eExtendedParameters = vbNullString
End With

If Err.Number <> 0 Then SetError "RemoveEvent()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub ConvertCaption(lType As eObjectTypes)
'On Local Error Resume Next

lEvents.eSettings.iOverEncoder = False
lEvents.eSettings.iOverRipper = False
With frmMain
    If lEvents.eEncoderBusy = True Or lEvents.eRipperBusy = True Then
        Exit Sub
    Else
        lEvents.eEventType = lType
        Select Case LCase(lType)
            Case oIdle
                If lEffectsPresets.eStatus = eOpening Then
                    .lblInfo.Caption = "Please wait ..."
                ElseIf lEffectsPresets.eStatus = eOpen Then
                    .lblInfo.Caption = "Editor Ready"
                Else
                    .lblInfo.Caption = "Ready"
                End If
            Case oMinimize
                .lblInfo.Caption = "Hide - Send to taskbar"
            Case oEnd
                .lblInfo.Caption = "Power - End Program"
            Case oEncode
                frmMain.lblInfo.Caption = "Encode - Wav to MP3"
                lEvents.eSettings.iOverEncoder = True
            Case oRip
                frmMain.lblInfo.Caption = "Rip - CDA to MP3"
                lEvents.eSettings.iOverRipper = True
            Case oStopEncoding
                frmMain.lblInfo.Caption = "Cancel Encode"
                lEvents.eSettings.iOverEncoder = True
            Case oPlayWav
                .lblInfo.Caption = "Search for mp3's"
                lEvents.eSettings.iOverRipper = True
            Case oPlayMp3
                lEvents.eSettings.iOverEncoder = True
                .lblInfo.Caption = "MP3 - Load Player"
            Case oTag
                .lblInfo.Caption = "Wave Audio Effects"
            Case oStopRipping
                .lblInfo.Caption = "Cancel Rip"
                lEvents.eSettings.iOverRipper = True
            Case oSkinEdit
                .lblInfo.Caption = "Skins - Edit look"
            Case oCDAudio
                .lblInfo.Caption = "CDAudio - Load Player"
            Case oOptions
                .lblInfo.Caption = "Settings"
            Case oPlayButton
                .lblInfo.Caption = "Play"
            Case oStopButton
                .lblInfo.Caption = "Stop"
            Case oForwardButton
                .lblInfo.Caption = "Go Forward"
            Case oBackwardButton
                .lblInfo.Caption = "Go Backward"
            Case oCDDB
                .lblInfo.Caption = "Please wait ..."
            Case oDecode
                .lblInfo.Caption = "Decode MP3's"
            Case oPlayWav
                .lblInfo.Caption = "Play Wave Audio"
        End Select
    End If
End With

If Err.Number <> 0 Then SetError "ConvertCaption()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub RegisterComponents()
'On Local Error Resume Next

frmMain.Ripper.Authorize "Leon Aiossa", "698070606"
frmMain.Encoder.Authorize "Leon Aiossa", "680665552"

If Err.Number <> 0 Then SetError "RegisterComponents()", lEvents.eSettings.iErrDescription, Err.Description
End Sub
