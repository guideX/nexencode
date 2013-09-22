Attribute VB_Name = "mdlId3"
Option Explicit
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Type gTag
    tHasTag As Boolean
    tFile As String
    tArtist As String * 30
    tTitle As String * 30
    tAlbum As String * 30
    tYear As String * 4
    tComment As String * 30
    tGenre As String * 1
    tCRC As String
    tLabel As String
    tFreqChan As String
    tBitrate As String
    tLength As String
    tCopyright As String
    tOrigional As String
    tSize As Single
    tEmphisis As String
    tLayer As String
End Type
Global lTag As gTag
Global MP3FileSize As Single
Global MP3FileLength As Single

Public Sub SaveTagInfo(lFilename As String)
'On Local Error Resume Next

If Len(lFilename) = 0 Then Exit Sub
Open lFilename For Binary As #1
lTag.tSize = FileLen(lFilename)
Put #1, lTag.tSize, "TAG" & lTag.tTitle & lTag.tArtist & lTag.tAlbum & lTag.tYear & lTag.tComment & lTag.tGenre
Close #1

If Err.Number <> 0 Then SetError "SaveTagInfo()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub PromptGetTag(Optional lFilename As String)
'On Local Error Resume Next
Dim msg As String

If Len(lFilename) = 0 Then
    lFilename = OpenDialog(frmMP3Info, "MP3 Files (*.mp3)|*.mp3|All Files (*.*)|*.*", "Select File", CurDir)
    If Len(lFilename) = 0 Then Exit Sub
End If
ClearAll
lTag.tFile = lFilename
GetTagInfo
GetMP3Info

If Err.Number <> 0 Then SetError "PromptGetTag()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Function Between(ByVal accNum As Byte, ByVal accDown As Byte, ByVal accUp As Byte) As Boolean
'On Local Error Resume Next

If accNum >= accDown And accNum <= accUp Then
    Between = True
Else
    Between = False
End If
If Err.Number <> 0 Then SetError "Between()", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Sub ClearInputs()
'On Local Error Resume Next

With lTag
    .tAlbum = ""
    .tArtist = ""
    .tBitrate = ""
    .tComment = ""
    .tCopyright = ""
    .tCRC = ""
    .tFreqChan = ""
    .tGenre = ""
    .tLabel = ""
    .tLength = ""
    .tOrigional = ""
    .tSize = 0
    .tTitle = ""
    .tYear = ""
End With
With frmMP3Info
    .txtArtist = ""
    .txtTitle = ""
    .txtAlbum = ""
    .txtYear = ""
    .cmbGenre.ListIndex = -1
    .txtComment = ""
End With
If Err.Number <> 0 Then SetError "ClearInputs()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub ClearLabels()
'On Local Error Resume Next
With frmMP3Info
    .lblBitrate.Caption = ""
    .lblCopyRight.Caption = ""
    .lblCRC.Caption = ""
    .lblEmphasis.Caption = ""
    .lblFreqChan.Caption = ""
    .lblBitrate.Caption = ""
    .lblLayer.Caption = ""
    .lblLength.Caption = ""
    .lblOriginal.Caption = ""
    .lblSize.Caption = ""
End With
If Err.Number <> 0 Then SetError "ClearLabels()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub ClearAll()
'On Local Error Resume Next
ClearInputs
ClearLabels
If Err.Number <> 0 Then SetError "ClearAll()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub GetMP3Info()
On Local Error GoTo ErrHandler
Dim Buffer As String * 4096, infoStr As String * 3, tmpByte As Byte, tmpNum As Byte, i As Integer, designator As Byte, baseFreq As Single

If Len(lTag.tFile) <> 0 Then
    Open lTag.tFile For Binary As #1
        Get #1, 1, Buffer
    Close #1
    For i = 1 To 4092
        If Asc(Mid(Buffer, i, 1)) = &HFF Then
            tmpByte = Asc(Mid(Buffer, i + 1, 1))
            If Between(tmpByte, &HF2, &HF7) Or Between(tmpByte, &HFA, &HFF) Then Exit For
        End If
    Next i
    If i = 4093 Then Exit Sub
    
    infoStr = Mid(Buffer, i + 1, 3)
    tmpByte = Asc(Mid(infoStr, 1, 1))

    If ((tmpByte Mod 16) Mod 2) = 0 Then
        lTag.tCRC = "Yes"
    Else
        lTag.tCRC = "No"
    End If

    If Between(tmpByte, &HF2, &HF7) Then
        lTag.tLayer = "MPEG 2.0"
        designator = 1
    Else
        lTag.tLayer = "MPEG 1.0"
        designator = 2
    End If

    If Between(tmpByte, &HF2, &HF3) Or Between(tmpByte, &HFA, &HFB) Then
        lTag.tLabel = lTag.tLayer & " layer 3"
    Else
        If Between(tmpByte, &HF4, &HF5) Or Between(tmpByte, &HFC, &HFD) Then
            lTag.tLabel = lTag.tLayer & " layer 2"
        Else
            lTag.tLabel = lTag.tLayer & " layer 1"
        End If
    End If
    
    tmpByte = Asc(Mid(infoStr, 2, 1))
    
    If Between(tmpByte Mod 16, &H0, &H3) Then
        baseFreq = 22.05
    Else
        If Between(tmpByte Mod 16, &H4, &H7) Then
            baseFreq = 24
        Else
            baseFreq = 16
        End If
    End If
    lTag.tFreqChan = baseFreq * designator * 1000 & " Hz"
    
    tmpNum = tmpByte \ 16 Mod 16
    If designator = 1 Then
        If tmpNum < &H8 Then
            lTag.tBitrate = tmpNum * 8
        Else
            lTag.tBitrate = 64 + (tmpNum - 8) * 16
        End If
    Else
        If tmpNum <= &H5 Then
            lTag.tBitrate = (tmpNum + 3) * 8
        Else
            If tmpNum <= &H9 Then
                lTag.tBitrate = 64 + (tmpNum - 5) * 16
            Else
                If tmpNum <= &HD Then
                    lTag.tBitrate = 128 + (tmpNum - 9) * 32
                Else
                    lTag.tBitrate = 320
                End If
            End If
        End If
    End If
    
    lTag.tLength = lTag.tSize / (lTag.tBitrate / 8) / 1000

    tmpByte = Asc(Mid(infoStr, 3, 1))
    tmpNum = tmpByte Mod 16
    
    lTag.tCopyright = ""
    If tmpNum \ 8 = 1 Then
        lTag.tCopyright = "Yes"
        tmpNum = tmpNum - 8
    Else
        lTag.tCopyright = "No"
    End If
    
    lTag.tOrigional = ""
    If (tmpNum \ 4) Mod 2 Then
        lTag.tOrigional = "Yes"
        tmpNum = tmpNum - 4
    Else
        lTag.tOrigional = "No"
    End If
    
    lTag.tEmphisis = ""
    Select Case tmpNum
    Case 0
        lTag.tEmphisis = "None"
    Case 1
        lTag.tEmphisis = "50/15 microsec"
    Case 2
        lTag.tEmphisis = "invalid"
    Case 3
        lTag.tEmphisis = "CITT j. 17"
    End Select
    
    tmpNum = (tmpByte \ 16) \ 4
    Select Case tmpNum
    Case 0
        lTag.tFreqChan = lTag.tFreqChan & " Stereo"
    Case 1
        lTag.tFreqChan = lTag.tFreqChan & " Joint Stereo"
    Case 2
        lTag.tFreqChan = lTag.tFreqChan & " 2 Channel"
    Case 3
        lTag.tFreqChan = lTag.tFreqChan & " Mono"
    End Select
End If

ErrHandler:
If Err.Number <> 0 Then SetError "GetMp3Info()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub SetId3Info(lFile As String, lTitle As String, lArtist As String, lAlbum As String, lYear As String, lGenre As String)
'On Local Error Resume Next

With lTag
    lTag.tFile = lFile
    lTag.tArtist = lArtist
    lTag.tAlbum = lAlbum
    lTag.tYear = lYear
    lTag.tComment = "Created with NexENCODE v" & App.Major & "." & App.Minor
    lTag.tGenre = lGenre
    lTag.tTitle = lTitle
    If DoesFileExist(lTag.tFile) = True Then lTag.tSize = FileLen(lTag.tFile)
End With

If Err.Number <> 0 Then SetError "SetId3Info()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub GetTagInfo()
On Local Error GoTo ErrHandler
Dim Buffer As String * 128, TempString As String, i As Byte, lSize As Single

With frmMP3Info
    If Len(lTag.tFile) = 0 Or DoesFileExist(lTag.tFile) = False Then Exit Sub
    lTag.tSize = FileLen(lTag.tFile)
    If lTag.tSize = 0 Then Exit Sub
    Open lTag.tFile For Binary As #1
    Get #1, lTag.tSize - 127, Buffer
    If Format(Left(Buffer, 3), "<") <> "tag" Then
        lTag.tHasTag = False
        lTag.tAlbum = ""
        lTag.tArtist = ""
        lTag.tGenre = ""
        lTag.tLabel = ""
        lTag.tOrigional = ""
        lTag.tTitle = ""
        lTag.tYear = ""
    Else
        lTag.tHasTag = True

        lTag.tTitle = Trim(Mid(Buffer, 4, 30))
        lTag.tArtist = Trim(Mid(Buffer, 34, 30))
        lTag.tAlbum = Trim(Mid(Buffer, 64, 30))
        lTag.tYear = Trim(Mid(Buffer, 94, 4))
        lTag.tComment = Trim(Mid(Buffer, 98, 30))
        For i = 0 To 148
            If .cmbGenre.ItemData(i) = Trim(Asc(Mid$(Buffer, 128, 1))) Then Exit For
        Next i
        If i < 149 Then
            .cmbGenre.ListIndex = i
            lTag.tGenre = i
        End If
    End If
    Close #1
End With

ErrHandler:
If Err.Number <> 0 Then SetError "GetMp3Info()", lEvents.eSettings.iErrDescription, Err.Description
End Sub
