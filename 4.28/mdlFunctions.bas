Attribute VB_Name = "mdlFunctions"
Option Explicit
Private blnFinished As Boolean
Enum eCDDBInfoTypes
    eArtist = 1
    eTitle = 2
    ecYear = 3
    eGenre = 4
    eLabel = 5
    eYear = 6
    eTrackname = 7
    eTrackcount = 8
End Enum

Private Function LeftZeroPad(s As String, n As Integer) As String
'On Local Error Resume Next
If Len(s) < n Then
    LeftZeroPad = String$(n - Len(s), "0") & s
Else
    LeftZeroPad = s
End If
If Err.Number <> 0 Then SetError "LeftZeroPad", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Function SaveAs(lForm As Form, lFilter As String, lTitle As String, lDirectory As String, lExtension As String)
'On Local Error Resume Next
Dim msg As String
If Left(lExtension, 1) <> "." Then lExtension = "." & lExtension
msg = SaveDialog(lForm, lFilter, lTitle, lDirectory)
If Len(msg) <> 0 Then
    msg = Left(msg, Len(msg) - 1)
    If LCase(Right(msg, 4)) <> LCase(lExtension) Then
        msg = msg & LCase(lExtension)
    End If
    SaveAs = msg
End If
If Err.Number <> 0 Then SetError "SaveAs", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Function ReturnFreeDBQueryString(lToc As String) As String
Dim strTocData() As String, sum As Long, tmp As Long, idx As Integer, msg As String, msg2 As String, lTrackNum As String, mediaID As String
On Error GoTo errChk
msg = Trim$(lToc)
If (msg = "" Or InStr(1, msg, " ") = 0) Then Exit Function
strTocData = Split(msg, " ", 100, vbTextCompare)
lTrackNum = UBound(strTocData)
For idx = 1 To lTrackNum
    lTracks.tTrack(idx).tLength = (Val(strTocData(idx)) - Val(strTocData(idx - 1))) \ 75
Next idx
lTracks.tDiscLen = (Val(strTocData(lTrackNum)) \ 75) - (Val(strTocData(0)) \ 75)
For idx = 0 To lTrackNum - 1
    tmp = Val(strTocData(idx)) \ 75
    Do While tmp > 0
        sum = sum + (tmp Mod 10)
        tmp = tmp \ 10
    Loop
Next idx
mediaID = LCase$(LeftZeroPad(Hex$(sum Mod &HFF), 2) & LeftZeroPad(Hex$(lTracks.tDiscLen), 4) & LeftZeroPad(Hex$(lTrackNum), 2))
msg2 = mediaID & "+" & lTrackNum
For idx = 0 To lTrackNum - 1
    msg2 = msg2 & "+" & strTocData(idx)
Next
msg2 = msg2 & "+" & (Val(strTocData(lTrackNum)) \ 75)
ReturnFreeDBQueryString = msg2
Exit Function

errChk:
    If Err.Number <> 0 Then SetError "ReturnFreeDBQueryString", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Function SelectCDDrive() As String
'On Local Error Resume Next
Dim lResult As VbMsgBoxResult
If Len(lRipperSettings.eDriveLetter) = 0 Then
    If lDrives.dCount <> 0 Then
        lRipperSettings.eDriveLetter = lDrives.dDrive(1).dLetter
    Else
        If lEvents.eSettings.iOverwritePrompts = True Then
            lResult = MsgBox("No CD-ROM Drive(s) were detected. Would you like to input one manually?", vbYesNo + vbExclamation, App.Title)
            If lResult = vbYes Then
                lRipperSettings.eDriveLetter = InputBox("Unable to detect CD Drive. Enter drive letter (example: E:)", "NexENCODE Setup", "D:")
                If Len(lRipperSettings.eDriveLetter) <> 0 Then WriteINI lIniFiles.iSettings, "Settings", "DriveLetter", lRipperSettings.eDriveLetter
            End If
        Else
            lRipperSettings.eDriveLetter = InputBox("Unable to detect CD Drive. Enter drive letter (example: E:)", "NexENCODE Setup", "D:")
            If Len(lRipperSettings.eDriveLetter) <> 0 Then WriteINI lIniFiles.iSettings, "Settings", "DriveLetter", lRipperSettings.eDriveLetter
        End If
    End If
End If
SelectCDDrive = lRipperSettings.eDriveLetter
If Err.Number <> 0 Then SetError "SelectCDDrive", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Function DecodeMPEG(lInputFilename As String, lInputFilePath As String, lOutputFilename As String) As Integer
'On Local Error Resume Next
Dim i As Integer, msg As String, msg2 As String
Dim mbox As VbMsgBoxResult
If Len(lInputFilename) <> 0 And Len(lOutputFilename) <> 0 Then
    If Right(lInputFilePath, 1) <> "\" Then lInputFilePath = lInputFilePath & "\"
    With frmMain.Decoder
        If DoesFileExist(lInputFilePath & lInputFilename) = True Then
            If DoesFileExist(lInputFilePath & lOutputFilename) = True Then
                If lEvents.eSettings.iOverwritePrompts = True Then
                    mbox = MsgBox("File exists, overwrite?", vbYesNo, "NexENCODE")
                    If mbox = vbYes Then
                        Kill lInputFilePath & lOutputFilename
                    ElseIf mbox = vbNo Then
                        MsgBox "Decode function canceled", vbExclamation
                        Exit Function
                    End If
                Else
                    Kill lInputFilePath & lOutputFilename
                End If
            End If
Start:
            frmMain.Decoder.Authorize "Leon J Aiossa", "812144397"
            frmMain.Decoder.Close: DoEvents
            FindOpenDevice frmMain.Decoder
            i = .Open(lInputFilePath & lInputFilename, lInputFilePath & lOutputFilename)
            If i <> 0 Then
                If i = 712 Then
                    msg2 = MsgBox("NexENCODE could not decode '" & lInputFilename & "' on drive '" & UCase(Left(lInputFilePath, 1)) & "'" & vbCrLf & vbCrLf & "Would you like to copy the mp3 to a different drive?", vbYesNo + vbQuestion)
                    If msg2 = vbYes Then
                        frmSelectDir.Show 1
                        If Len(lEvents.eRetStr) <> 0 And Len(lEvents.eRetStr) > 3 Then
                            msg2 = lEvents.eRetStr & "\"
                            If DoesFileExist(msg2 & lInputFilename) = False Then
                                FileCopy lInputFilePath & lInputFilename, msg2 & lInputFilename
                                DoEvents
                                pause 0.2
                                lInputFilePath = msg2
                                GoTo Start
                            Else
                                lInputFilePath = msg2
                                GoTo Start
                            End If
                        Else
                            Exit Function
                        End If
                    ElseIf msg2 = vbCancel Then
                        Exit Function
                    End If
                End If
                ShowMp3PlayError i
                Exit Function
            Else
                frmMain.lblMp3File.Caption = lInputFilename
                ResetEncoderCircles
                frmMain.tmrShowEncoderCircles.interval = 40
                frmMain.tmrShowEncoderCircles.Enabled = True
                .Play
                ToggleButtons oEncode
                DecodeMPEG = i
            End If
        End If
    End With
End If
If Err.Number <> 0 Then SetError "DecodeMP3", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Function GetCDDBServerIndexByLocation(lLocation As String) As Integer
Dim i As Integer
For i = 1 To lCDDBServ.cCount
    If LCase(lCDDBServ.cServer(i).sLocation) = LCase(lLocation) Then
        GetCDDBServerIndexByLocation = i
        Exit For
    End If
Next i
If Err.Number <> 0 Then SetError "GetCDDBServerIndexByLocation", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Function GetCDDBServerIndexByIp(lIp As String) As Integer
Dim i As Integer
For i = 1 To lCDDBServ.cCount
    If LCase(lCDDBServ.cServer(i).sIp) = lIp Then
        GetCDDBServerIndexByIp = i
        Exit For
    End If
Next i
If Err.Number <> 0 Then SetError "GetCDDBServerIndexByIp", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Function AddCDDBServer(lLocation As String, lIp As String) As Integer
Dim i As Integer
If Len(lLocation) <> 0 And Len(lIp) <> 0 Then
    i = lCDDBServ.cCount + 1
    lCDDBServ.cServer(i).sIp = lIp
    lCDDBServ.cServer(i).sLocation = lLocation
    lCDDBServ.cCount = i
    WriteINI lIniFiles.iCDDBServers, Str(i), "Ip", lIp
    WriteINI lIniFiles.iCDDBServers, Str(i), "Location", lLocation
    WriteINI lIniFiles.iCDDBServers, "Settings", "Count", i
    AddCDDBServer = i
End If
If Err.Number <> 0 Then SetError "DoesFileExist()", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Function DoesFileExist(lFilename As String) As Boolean
'On Local Error Resume Next
Dim msg As String
msg = Dir(lFilename)
If msg <> "" Then
    DoesFileExist = True
Else
    DoesFileExist = False
End If
If Err.Number <> 0 Then SetError "DoesFileExist()", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Function GetRnd(Num As Integer) As Integer
'On Local Error Resume Next
Randomize Timer
GetRnd = Int((Num * Rnd) + 1)
If Err.Number <> 0 Then SetError "GetRnd", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Function IsAlphaNum(sData As String) As Boolean
'On Local Error Resume Next
If sData = "" Then Exit Function
sData = Mid(sData, 1, 1)
If Asc(sData) >= 65 And Asc(sData) <= 90 Then
    IsAlphaNum = True
    Exit Function
ElseIf Asc(sData) >= 97 And Asc(sData) <= 122 Then
    IsAlphaNum = True
    Exit Function
ElseIf Asc(sData) >= 48 And Asc(sData) <= 57 Then
    IsAlphaNum = True
    Exit Function
End If
IsAlphaNum = False
If Err.Number <> 0 Then SetError "IsAlphaNum", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Function GetCheckboxValue(lCheckbox As CheckBox) As Boolean
'On Local Error Resume Next
If lCheckbox.Value = 1 Then
    GetCheckboxValue = True
Else
    GetCheckboxValue = False
End If
If Err.Number <> 0 Then SetError "GetCheckboxValue()", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Function CheckPassword() As Boolean
'On Local Error Resume Next
Dim msg As String
If Len(lEvents.eName) <> 0 And Len(lEvents.ePassword) <> 0 Then
    msg = Crypt(lEvents.ePassword, "pickles", True)
    If msg = lEvents.eName Then
        CheckPassword = True
        lEvents.eRegistered = True
    Else
        CheckPassword = False
    End If
Else
    CheckPassword = False
End If
If Err.Number <> 0 Then SetError "CheckPassword()", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Function Crypt(Source As String, strPassword As String, EnDeCrypt As Boolean) As String
'On Local Error Resume Next
Dim intPassword As Long, X As Integer, intCrypt As Long
For X = 1 To Len(strPassword)
    intPassword = intPassword + Asc(Mid$(strPassword, X, 1))
Next X
For X = 1 To Len(Source)
    If EnDeCrypt = True Then
        intCrypt = Asc(Mid$(Source, X, 1)) + intPassword + X
        Do Until intCrypt <= 255
            intCrypt = intCrypt - 255
        Loop
    Else
        intCrypt = Asc(Mid$(Source, X, 1)) - intPassword - X
        Do Until intCrypt > 0
            intCrypt = intCrypt + 255
        Loop
    End If
    Crypt = Crypt & Chr(intCrypt)
Next X
If Err.Number <> 0 Then SetError "Crypt()", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Function ReadFile(lFile As String) As String
'On Local Error Resume Next
Dim o As Integer, msg As String
o = FreeFile
Open lFile For Input As #o
    'ReadFile = Input(LOF(o), o)
    msg = StrConv(InputB(LOF(o), o), vbUnicode)
    ReadFile = Left(msg, Len(msg) - 2)
Close #o
End Function

Public Function GetDriveNumberByLetter(lLetter As String) As String
'On Local Error Resume Next
Dim i As Integer
With frmMain.Ripper
    For i = 1 To .DriveCount
        If .DriveStringByNumber(i) = .DriveStringByLetter(lLetter) Then
            GetDriveNumberByLetter = i
            Exit For
        End If
    Next i
End With
If Err.Number <> 0 Then SetError "GetDriveLetterByDriveNumber", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Function ParseString(lWhole As String, lStart As String, lEnd As String)
On Local Error GoTo ErrHandler
Dim len1 As Integer, len2 As Integer, Str1 As String, Str2 As String
len1 = InStr(lWhole, lStart)
len2 = InStr(lWhole, lEnd)
Str1 = Right(lWhole, Len(lWhole) - len1)
Str2 = Right(lWhole, Len(lWhole) - len2)
ParseString = Left(Str1, Len(Str1) - Len(Str2) - 1)
Err = 0
ErrHandler:
End Function

Public Function TimeToString(currTime As Single) As String
'On Local Error Resume Next
Dim sMinutes As String, sSeconds As String, sMilliseconds As String, sHours As String, iMinutes As Integer, iSeconds As Integer, iHours As Integer
iHours = Int(currTime / 3600)
iMinutes = Int((currTime - iHours * 3600) / 60)
iSeconds = Int(currTime - iHours * 3600 - iMinutes * 60)
If iHours > 0 Then
    sHours = Format(Str(iHours), "00")
    TimeToString = sHours & "-"
End If
sMinutes = Format(Str(Int(iMinutes)), "00")
sSeconds = Format(Str(Int(iSeconds)), "00")
sMilliseconds = Format(Str(Int((currTime - iHours * 3600 - iMinutes * 60 - iSeconds) * 100)), "00")
TimeToString = TimeToString & sMinutes & ":" & sSeconds & "." & sMilliseconds
Err = 0
If Err.Number <> 0 Then SetError "TimeToString", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Function FindComoboxIndex(lCombo As ComboBox, lText As String) As Integer
'On Local Error Resume Next
Dim i As Integer
If Len(lText) <> 0 Then
    For i = 0 To lCombo.ListCount
        If LCase(lCombo.List(i)) = LCase(lText) Then
            FindComoboxIndex = i
            Exit For
            Exit Function
        End If
    Next i
End If
If Err.Number <> 0 Then SetError "FindComboIndex()", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Function FindListboxIndex(lListbox As ListBox, lText As String) As Integer
'On Local Error Resume Next
Dim i As Integer
If Len(lText) <> 0 Then
    For i = 0 To lListbox.ListCount
        If LCase(lListbox.List(i)) = LCase(lText) Then
            FindListboxIndex = i
            Exit For
            Exit Function
        End If
    Next i
End If
If Err.Number <> 0 Then SetError "FindListboxIndex()", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Function AddPlayer(lPlayername As String, lPlayerLocation As String, lType As ePlayerTypes, lPlaylist As String) As Integer
'On Local Error Resume Next
Dim lFile As String, lPath As String, i As Integer, f As Integer
lFile = lPlayerLocation
lFile = GetFileTitle(lFile)
lPath = Left(lPlayerLocation, Len(lPlayerLocation) - Len(lFile))
If Len(lPlayername) <> 0 And Len(lPlayerLocation) <> 0 Then
    f = FindPlayerIndex(lPlayername)
    If f <> 0 Then
        AddPlayer = f
        Exit Function
    End If
    lPlayers.pCount = lPlayers.pCount + 1
    WriteINI lIniFiles.iPlayers, "Settings", "Count", lPlayers.pCount
    With lPlayers.pPlayer(lPlayers.pCount)
        .pFile = lFile
        .pPath = lPath
        .pName = lPlayername
        .pType = lType
        .pPlaylist = lPlaylist
        WriteINI lIniFiles.iPlayers, lPlayers.pCount, "File", .pFile
        WriteINI lIniFiles.iPlayers, lPlayers.pCount, "Name", .pName
        WriteINI lIniFiles.iPlayers, lPlayers.pCount, "Path", .pPath
        WriteINI lIniFiles.iPlayers, lPlayers.pCount, "Playlist", .pPlaylist
        WriteINI lIniFiles.iPlayers, lPlayers.pCount, "Type", .pType
        AddPlayer = lPlayers.pCount
    End With
End If
If Err.Number <> 0 Then SetError "AddPlayer()", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Sub RemovePlayer(lName As String)
'On Local Error Resume Next
Dim i As Integer

i = FindPlayerIndex(lName)
If i <> 0 Then
    With lPlayers.pPlayer(i)
        .pFile = ""
        .pName = ""
        .pPath = ""
        .pPlaylist = ""
        .pType = 0
        WriteINI lIniFiles.iPlayers, Str(i), "File", ""
        WriteINI lIniFiles.iPlayers, Str(i), "Name", ""
        WriteINI lIniFiles.iPlayers, Str(i), "Path", ""
        WriteINI lIniFiles.iPlayers, Str(i), "Playlist", ""
        WriteINI lIniFiles.iPlayers, Str(i), "Type", ""
    End With
End If

If Err.Number <> 0 Then SetError "RemovePlayer()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Function FindPlayerIndex(lName As String) As Integer
'On Local Error Resume Next
Dim i As Integer

For i = 1 To lPlayers.pCount
    If LCase(lPlayers.pPlayer(i).pName) = LCase(lName) Then
        FindPlayerIndex = i
        Exit For
    End If
Next i

If Err.Number <> 0 Then SetError "FindPlayerIndex()", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Function SaveFile(lFilename As String, lText As String) As Boolean
'On Local Error Resume Next

If Len(lFilename) <> 0 And Len(lText) <> 0 Then
    If InStr(lFilename, "\\") Then
        SetError "SaveFile", "Coding error", "Double \\'s were used"
        Exit Function
    End If
    Open lFilename For Output As #10
    If Err.Number <> 0 Then
        Close #10
        SetError "SaveFile", lEvents.eSettings.iErrDescription, Err.Description
        Exit Function
    End If
    Print #10, lText
    Close #10
End If

If Err.Number <> 0 Then SetError "SaveFile()", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Function FindTrackIndex(lSearch As String) As Integer
'On Local Error Resume Next
Dim i As Integer

For i = 1 To lTracks.tCount
    If InStr(lSearch, lTracks.tTrack(i).tName) Then
        FindTrackIndex = i
        Exit For
    End If
Next i

If Err.Number <> 0 Then SetError "FindTrackIndex()", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Function CheckFileConstants(lDirectory As String, lFilename As String) As Boolean
'On Local Error Resume Next
Dim msg As String

If DoesDirExist(lDirectory) = False Then MakeDir lDirectory
If DoesFileExist(lFilename) = True Then
    If lEvents.eSettings.iOverwritePrompts = True Then
        msg = MsgBox("File already exists (" & lFilename & ") , delete existing file and continue?", vbYesNoCancel + vbExclamation, "Overwrite?")
    Else
        msg = vbYes
    End If
    If msg = vbYes Then
        Kill lFilename
        If DoesFileExist(lFilename) Then
            MsgBox "File is in use, aborting!", vbExclamation
            Exit Function
        End If
        If Err.Number = 70 Then
            Err = 0
            SetError "RecordCDTrack()", lEvents.eSettings.iErrDescription, "File " & lFilename & " is in use. Try exiting NexENCODE Studio, deleteing the existing wav file(s), then trying your process again"
            Exit Function
        End If
    ElseIf msg = vbNo Then
        msg = InputBox("File already exists, please rename:", "Error", lFilename)
        If msg = lFilename Or Len(msg) = 0 Then
            Exit Function
        End If
    ElseIf msg = vbCancel Then
        Exit Function
    End If
End If
CheckFileConstants = True

If Err.Number <> 0 Then SetError "CheckFileConstants()", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Function PlayWav(strPath As String, sndVal As sndConst)
'On Local Error Resume Next

If lPlayer.pStatus = sPlaying Then Exit Function
If lEvents.eSettings.iPlayWavs = True Then sndPlaySound strPath, sndVal

If Err.Number <> 0 Then SetError "Playwav()", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Function FindRipEvent() As Integer
'On Local Error Resume Next
Dim i As Integer

For i = 1 To lEvents.eEventCount
    If lEvents.eEvent(i).eEventType = Rip Then
        FindRipEvent = i
        Exit Function
    End If
Next i

If Err.Number <> 0 Then SetError "FindRipEvent()", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Function MakeTransparent(ByVal hwnd As Long, Perc As Integer) As Long
'On Local Error Resume Next
Dim msg As Long

If Perc < 0 Or Perc > 255 Then
    MakeTransparent = 1
Else
    msg = GetWindowLong(hwnd, GWL_EXSTYLE)
    msg = msg Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, msg
    SetLayeredWindowAttributes hwnd, 0, Perc, LWA_ALPHA
    MakeTransparent = 0
End If
If Err Then MakeTransparent = 2

If Err.Number <> 0 Then SetError "MakeTransparent()", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Function MakeOpaque(ByVal hwnd As Long) As Long
'On Local Error Resume Next
Dim msg As Long

msg = GetWindowLong(hwnd, GWL_EXSTYLE)
msg = msg And Not WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, msg
SetLayeredWindowAttributes hwnd, 0, 0, LWA_ALPHA
MakeOpaque = 0
If Err Then MakeOpaque = 2

If Err.Number <> 0 Then SetError "MakeOpaque()", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Function GetFileTitle(lFilename As String) As String
'On Local Error Resume Next

If Len(lFilename) <> 0 Then
Again:
    If InStr(lFilename, "\") Then
        lFilename = Right(lFilename, Len(lFilename) - InStr(lFilename, "\"))
        If InStr(lFilename, "\") Then
            GoTo Again
        Else
            GetFileTitle = lFilename
        End If
    Else
        GetFileTitle = lFilename
    End If
Else
    Exit Function
End If

If Err.Number <> 0 Then SetError "GetFileTitle()", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Function DoesDirExist(lDirectory As String) As Boolean
'On Local Error Resume Next
Dim msg As String

msg = Dir(lDirectory, vbDirectory)
If Len(msg) <> 0 Then
    DoesDirExist = True
Else
    DoesDirExist = False
End If

If Err.Number <> 0 Then SetError "DoesDirExist()", lEvents.eSettings.iErrDescription, Err.Description
End Function
