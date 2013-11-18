Attribute VB_Name = "mdlSubs"
Option Explicit

Enum eEventTypes
    Rip = 1
    Encode = 2
    Id3 = 3
    Effects = 4
    Download = 5
    Upload = 6
End Enum
Private Type gEvent
    eEnabled As Boolean
    eDescription As String
    eEventType As eEventTypes
    eInputFilepath As String
    eInputFilename As String
    eOutputFilepath As String
    eOutputFilename As String
    eTrack As Integer
    eExtendedParameters As String
End Type
Private Type gEvents
    eDescription As String
    ePercent As Integer
    eEnabled As Boolean
    eEventBusy As Boolean
    eEventCount As Integer
    eEvent(100) As gEvent
    eMaxEvents As Integer
    eCircleNum As Integer
    eCurrentFilename As String
End Type
Public lEvents As gEvents

Public Sub Pause(interval)
Dim Current
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub

Public Function GetFileTitle(lFilename As String) As String
'on local error resume next

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
End Function

Public Sub RipCircleEffect(lPercent As Integer)
With frmMain
Select Case lEvents.ePercent
Case 1
    .shpRipper(0).BackColor = vbGreen
Case 10
    .shpRipper(1).BackColor = vbGreen
Case 20
    .shpRipper(2).BackColor = vbGreen
Case 30
    .shpRipper(3).BackColor = vbGreen
Case 40
    .shpRipper(4).BackColor = vbGreen
Case 50
    .shpRipper(5).BackColor = vbGreen
Case 60
    .shpRipper(6).BackColor = vbGreen
Case 70
    .shpRipper(7).BackColor = vbGreen
Case 80
    .shpRipper(8).BackColor = vbGreen
Case 90
    .shpRipper(9).BackColor = vbGreen
Case 95
    .shpRipper(10).BackColor = vbGreen
End Select
End With
End Sub

Public Sub EncodeCircleEffect(lPercent As Integer)
With frmMain
Select Case lEvents.ePercent
Case Is > 95
    .shpEncoder(10).BackColor = vbGreen
Case Is > 90
    .shpEncoder(9).BackColor = vbGreen
Case Is > 80
    .shpEncoder(8).BackColor = vbGreen
Case Is > 70
    .shpEncoder(7).BackColor = vbGreen
Case Is > 60
    .shpEncoder(6).BackColor = vbGreen
Case Is > 50
    .shpEncoder(5).BackColor = vbGreen
Case Is > 40
    .shpEncoder(4).BackColor = vbGreen
Case Is > 30
    .shpEncoder(3).BackColor = vbGreen
Case Is > 20
    .shpEncoder(2).BackColor = vbGreen
Case Is > 10
    .shpEncoder(1).BackColor = vbGreen
Case Is > 1
    .shpEncoder(0).BackColor = vbGreen
End Select
End With
End Sub

Public Sub FormDrag(lFormname As Form)
    ReleaseCapture
    Call SendMessage(lFormname.hwnd, &HA1, 2, 0&)
End Sub

Public Sub AddEvent(lEventType As eEventTypes, lInputFilePath As String, lInputFilename As String, lOutputFilepath As String, lOutputFilename As String, lTrack As Integer, lExtendedParameters As String)
Dim i As Integer
lEvents.eMaxEvents = 100
lEvents.eEnabled = True
lEvents.eEventCount = lEvents.eEventCount + 1
i = lEvents.eEventCount
With lEvents.eEvent(i)
    .eEnabled = True
    .eEventType = lEventType
    .eInputFilepath = lInputFilePath
    .eInputFilename = lInputFilename
    .eOutputFilepath = lOutputFilepath
    .eOutputFilename = lOutputFilename
    .eTrack = lTrack
    .eExtendedParameters = lExtendedParameters
End With
If lEvents.eEventBusy = False Then
    ProcessNextEvent
    DoEvents
    frmMain.tmrProcessEvent.Enabled = True
Else
    Exit Sub
End If
End Sub

Public Sub EncodeFile(lInputWavFile As String, lOutputMp3File As String)
Dim msg As String
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

frmMain.Encoder.Open lInputWavFile, lOutputMp3File
frmMain.Encoder.Encode
DoEvents
End Sub

Public Sub ProcessNextEvent()
frmMain.tmrProcessEvent.Enabled = False
Dim i As Integer
Start:
i = lEvents.eEventCount
If i = 0 Then Exit Sub
If lEvents.eEvent(i).eEnabled = False Then
    RemoveEvent i
    GoTo Start
Else
    Select Case lEvents.eEvent(i).eEventType
    Case Encode
        If Len(lEvents.eEvent(i).eInputFilename) <> 0 Then
            lEvents.eCurrentFilename = lEvents.eEvent(i).eOutputFilepath & "\" & lEvents.eEvent(i).eOutputFilename
            EncodeFile lEvents.eEvent(i).eInputFilepath & "\" & lEvents.eEvent(i).eInputFilename, lEvents.eEvent(i).eOutputFilepath & "\" & lEvents.eEvent(i).eOutputFilename
            frmMain.lblMp3File.Caption = lEvents.eEvent(i).eOutputFilename
            RemoveEvent i
        End If
    Case Rip
            If lEvents.eEvent(i).eTrack = 0 Then
                RemoveEvent i
                Exit Sub
            End If
            lEvents.eCurrentFilename = lEvents.eEvent(i).eOutputFilepath & "\" & lEvents.eEvent(i).eOutputFilename
            RecordCDTrack lEvents.eEvent(i).eTrack, lEvents.eEvent(i).eOutputFilepath & "\" & lEvents.eEvent(i).eOutputFilename
            frmMain.lblWavFile.Caption = lEvents.eEvent(i).eOutputFilename
            RemoveEvent i
            lEvents.eEventBusy = True
            frmMain.tmrShowRipperCircles.Enabled = True
    End Select
End If
End Sub

Public Sub ResetRipperCircles()
Dim i As Integer
For i = 0 To 11
    frmMain.shpRipper(i).BackColor = frmMain.shpRipper(11).BackColor
    frmMain.shpRipper(i).Visible = False
Next i
End Sub

Public Sub ResetEncoderCircles()
Dim i As Integer
For i = 0 To 11
    frmMain.shpEncoder(i).BackColor = frmMain.shpEncoder(11).BackColor
    frmMain.shpEncoder(i).Visible = False
Next i
lEvents.eCircleNum = 0
End Sub

Public Sub RecordCDTrack(lTrack As Integer, lOutputFilename As String)
If lTrack = 0 Then Exit Sub
If Len(lOutputFilename) = 0 Then
    lOutputFilename = SaveDialog(frmMain, "Wav Files (*.wav)|*.wav|All Files (*.*)|*.*", "Save As ..", CurDir)
    If Len(lOutputFilename) = 0 Then Exit Sub
End If
frmMain.Ripper.OpenDriveByNumber (1)
frmMain.Ripper.ReadTrack lTrack, lOutputFilename
End Sub

Public Sub RemoveEvent(lIndex As Integer)
lEvents.eMaxEvents = 100
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
lEvents.eEventCount = lEvents.eEventCount - 1
End Sub

Public Function FindEventIndex(lDescription As String)
Dim i As Integer
For i = 1 To lEvents.eEventCount
    If LCase(lEvents.eEvent(i).eDescription) = LCase(lDescription) Then
        FindEventIndex = i
        Exit For
    End If
Next i
End Function

Public Sub ConvertCaption()
With lEvents
If .eEventBusy = True Then
    frmMain.lblInfo.Caption = lEvents.eDescription & " " & lEvents.ePercent
Else
    Select Case LCase(.eDescription)
        Case "minimize"
            frmMain.lblInfo.Caption = "minimize"
        Case "end"
            frmMain.lblInfo.Caption = "power (exit)"
        Case "idle"
            If frmMain.imgFlashEncode.Picture <> frmMain.imgFlashEncode1.Picture Then frmMain.imgFlashEncode.Picture = frmMain.imgFlashEncode1.Picture
            If frmMain.imgRipper.Picture <> frmMain.imgRipper1.Picture Then frmMain.imgRipper.Picture = frmMain.imgRipper1.Picture
            frmMain.lblInfo.Caption = "Idle"
            frmMain.tmrFlash.Enabled = False
        Case "encode"
            frmMain.tmrFlash.Enabled = True
            frmMain.lblInfo.Caption = "encode .wav files"
        Case "rip"
            frmMain.tmrFlash.Enabled = True
            frmMain.lblInfo.Caption = "copy cd audio"
        Case "stopencode"
            frmMain.lblInfo.Caption = "stop encoding"
            frmMain.tmrFlash.Enabled = True
        Case "playwav"
            frmMain.tmrFlash.Enabled = True
            frmMain.lblInfo.Caption = "play wav"
        Case "playmp3"
            frmMain.lblInfo.Caption = "play mp3"
            frmMain.tmrFlash.Enabled = True
        Case "id3"
            frmMain.lblInfo.Caption = "set .mp3 file info"
        Case "stoprip"
            frmMain.tmrFlash.Enabled = True
            frmMain.lblInfo.Caption = "stop ripping"
    End Select
End If
End With
End Sub

Public Sub RegisterComponents()
frmMain.Ripper.Authorize "Leon Aiossa", "698070606"
frmMain.Encoder.Authorize "Leon Aiossa", "680665552"
End Sub
