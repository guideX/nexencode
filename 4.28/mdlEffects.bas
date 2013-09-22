Attribute VB_Name = "mdlEffects"
Option Explicit
Private Type gEcho
    eEnabled As Boolean
    eShortDelay As Integer
    eShortRatio As Integer
    eDescription As String
End Type
Private Type gEchoPresets
    eEcho(100) As gEcho
    eCount As Integer
End Type
Private Type gChorus
    cDescription As String
    cEnabled As Boolean
    cShortDelay As Integer
    cShortDepth As Integer
    cFloatRate As Integer
    cWaveForm As Integer
    cShortDry As Integer
    cInvertFeedback As Integer
    cShortMixing As Integer
    cShortFeedback As Integer
    cShortWet As Integer
End Type
Private Type gChorusPresets
    cCount As Integer
    cChorus(100) As gChorus
End Type
Private Type gDistortion
    lDescription As String
    lEnabled As Boolean
    lDry As Integer
    lDistorted As Integer
    lThreshold As Integer
    lClamp As Integer
    lGate As Integer
End Type
Private Type gDistortionPresets
    dCount As Integer
    dDistortion(100) As gDistortion
End Type
Enum eEffectsStatus
    ePlaying = 1
    eStopped = 2
    ePaused = 3
    eAddingEffect = 4
    eClosed = 5
    eOpen = 6
    eOpening = 7
End Enum
Private Type gEffects
    eSaved As Boolean
    eStatus As eEffectsStatus
    eDistortion As gDistortionPresets
    eChorus As gChorusPresets
    eEcho As gEchoPresets
End Type
Global lEffectsPresets As gEffects

Public Sub AddFadeIn(lSeconds10th As Integer)
'On Local Error Resume Next
InitEffects
frmMain.ns4Effects.FadeIn lSeconds10th
If Err.Number <> 0 Then SetError "AddFadeIn", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub AddFadeOut(lSeconds10th As Integer)
'On Local Error Resume Next
InitEffects
frmMain.ns4Effects.FadeOut lSeconds10th
If Err.Number <> 0 Then SetError "AddFadeOut", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub AddEcho(lShortDelay As Integer, lShortRatio As Integer)
'On Local Error Resume Next
InitEffects
frmMain.ns4Effects.Echo lShortDelay, lShortRatio
If Err.Number <> 0 Then SetError "AddCFilter", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub LoadEffectsPresets()
'On Local Error Resume Next
Dim i As Integer, msg As String
lEffectsPresets.eDistortion.dCount = ReadINI(lIniFiles.iEffects, "Settings", "DistortionCount", 0)
lEffectsPresets.eChorus.cCount = ReadINI(lIniFiles.iEffects, "Settings", "ChorusCount", 0)
lEffectsPresets.eEcho.eCount = ReadINI(lIniFiles.iEffects, "Settings", "EchoCount", 0)
If lEffectsPresets.eEcho.eCount <> 0 Then
    For i = 1 To lEffectsPresets.eEcho.eCount
        lEffectsPresets.eEcho.eEcho(i).eEnabled = ReadINI(lIniFiles.iEffects, "Echo" & i, "Enabled", False)
        If lEffectsPresets.eEcho.eEcho(i).eEnabled = True Then
            lEffectsPresets.eEcho.eEcho(i).eDescription = ReadINI(lIniFiles.iEffects, "Echo" & i, "Description", "")
            lEffectsPresets.eEcho.eEcho(i).eShortDelay = ReadINI(lIniFiles.iEffects, "Echo" & i, "ShortDelay", "")
            lEffectsPresets.eEcho.eEcho(i).eShortRatio = ReadINI(lIniFiles.iEffects, "Echo" & i, "ShortRatio", "")
        End If
    Next i
End If
If lEffectsPresets.eChorus.cCount <> 0 Then
    For i = 1 To lEffectsPresets.eChorus.cCount
        lEffectsPresets.eChorus.cChorus(i).cEnabled = ReadINI(lIniFiles.iEffects, "Chorus" & i, "Enabled", False)
        lEffectsPresets.eChorus.cChorus(i).cDescription = ReadINI(lIniFiles.iEffects, "Chorus" & i, "Description", "")
        lEffectsPresets.eChorus.cChorus(i).cShortDelay = ReadINI(lIniFiles.iEffects, "Chorus" & i, "ShortDelay", 0)
        lEffectsPresets.eChorus.cChorus(i).cShortDepth = ReadINI(lIniFiles.iEffects, "Chorus" & i, "ShortDepth", 0)
        lEffectsPresets.eChorus.cChorus(i).cShortDry = ReadINI(lIniFiles.iEffects, "Chorus" & i, "ShortDry", 0)
        lEffectsPresets.eChorus.cChorus(i).cShortFeedback = ReadINI(lIniFiles.iEffects, "Chorus" & i, "ShortFeedback", 0)
        lEffectsPresets.eChorus.cChorus(i).cFloatRate = ReadINI(lIniFiles.iEffects, "Chorus" & i, "FloatRate", 0)
        lEffectsPresets.eChorus.cChorus(i).cInvertFeedback = ReadINI(lIniFiles.iEffects, "Chorus" & i, "InvertFeedback", 0)
        lEffectsPresets.eChorus.cChorus(i).cShortMixing = ReadINI(lIniFiles.iEffects, "Chorus" & i, "ShortMixing", 0)
        lEffectsPresets.eChorus.cChorus(i).cWaveForm = ReadINI(lIniFiles.iEffects, "Chorus" & i, "ShortWaveForm", 0)
        lEffectsPresets.eChorus.cChorus(i).cShortWet = ReadINI(lIniFiles.iEffects, "Chorus" & i, "ShortWet", 0)
    Next i
End If
If lEffectsPresets.eDistortion.dCount <> 0 Then
    For i = 1 To lEffectsPresets.eDistortion.dCount
        lEffectsPresets.eDistortion.dDistortion(i).lDescription = ReadINI(lIniFiles.iEffects, "Distortion" & i, "Description", "")
        lEffectsPresets.eDistortion.dDistortion(i).lEnabled = ReadINI(lIniFiles.iEffects, "Distortion" & i, "Enabled", False)
        lEffectsPresets.eDistortion.dDistortion(i).lDry = ReadINI(lIniFiles.iEffects, "Distortion" & i, "Dry", 0)
        lEffectsPresets.eDistortion.dDistortion(i).lClamp = ReadINI(lIniFiles.iEffects, "Distortion" & i, "Clamp", 0)
        lEffectsPresets.eDistortion.dDistortion(i).lGate = ReadINI(lIniFiles.iEffects, "Distortion" & i, "Gate", 0)
        lEffectsPresets.eDistortion.dDistortion(i).lThreshold = ReadINI(lIniFiles.iEffects, "Distortion" & i, "Threshold", 0)
        lEffectsPresets.eDistortion.dDistortion(i).lDistorted = ReadINI(lIniFiles.iEffects, "Distortion" & i, "Distorted", 0)
        lEffectsPresets.eDistortion.dDistortion(i).lClamp = ReadINI(lIniFiles.iEffects, "Distortion" & i, "Clamp", 0)
    Next i
End If
If Err.Number <> 0 Then SetError "LoadEffectsPresets", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub EnableEffects()
'On Local Error Resume Next
With frmMain
    .mnuReverb.Enabled = True
    .mnuChorus.Enabled = True
    .mnuDistortion.Enabled = True
    .mnuEcho.Enabled = True
    .mnuFadeIN.Enabled = True
    .mnuFadeOut.Enabled = True
    .mnuAmplitude.Enabled = True
    .mnuCFilter.Enabled = True
    .mnuSaveWavAs.Enabled = True
    .mnuStopEffectWav.Enabled = True
    .mnuCloseWavFile.Enabled = True
    .mnuPlayEffectWav.Enabled = True
    .mnuInvert.Enabled = True
    .mnuShifting.Enabled = True
    .mnuEffectsEditor.Enabled = True
End With
If Err.Number <> 0 Then SetError "EnableEffects", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub DisableEffects()
'On Local Error Resume Next
With frmMain
    .mnuChorus.Enabled = False
    .mnuDistortion.Enabled = False
    .mnuEcho.Enabled = False
    .mnuFadeIN.Enabled = False
    .mnuFadeOut.Enabled = False
    .mnuAmplitude.Enabled = False
    .mnuCFilter.Enabled = False
    .mnuSaveWavAs.Enabled = False
    .mnuPlayEffectWav.Enabled = False
    .mnuStopEffectWav.Enabled = False
    .mnuCloseWavFile.Enabled = False
    .mnuInvert.Enabled = False
    .mnuShifting.Enabled = False
    .mnuReverb.Enabled = False
    .mnuEffectsEditor.Enabled = False
End With
If Err.Number <> 0 Then SetError "DisableEffects", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub CloseEffects()
'On Local Error Resume Next
Dim msg As String

With frmMain
    If lEffectsPresets.eSaved = False And lEvents.eSettings.iOverwritePrompts = True Then SaveEffects
    .ns4Effects.sTop
    .ns4Effects.StopEffect
    DisableEffects
    .mnuOpenCurrentFile.Enabled = True
End With
DoEvents
lEffectsPresets.eStatus = eClosed
If Err.Number <> 0 Then SetError "CloseEffects", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub SaveEffects()
'On Local Error Resume Next
Dim msg As String
If lEvents.eSettings.iOverwritePrompts = True Then
    msg = MsgBox("Save changes to '" & frmMain.lblWavFile.Caption & "'?", vbQuestion + vbYesNoCancel)
    If msg = vbYes Then
        msg = ""
        msg = SaveAs(frmMain, "Wave Audio (*.wav)|*.wav", "Save as ...", CurDir, "wav")
        If Len(msg) <> 0 Then
            InitEffects
            DoEvents
            pause 1
            frmMain.ns4Effects.InputFileSave msg
            pause 0.2
            lEffectsPresets.eSaved = True
        Else
            Exit Sub
        End If
    ElseIf msg = vbCancel Then
        Exit Sub
    End If
Else
    msg = ""
    msg = SaveAs(frmMain, "Wave Audio (*.wav)|*.wav", "Save as ...", CurDir, "wav")
    If Len(msg) <> 0 Then
        InitEffects
        DoEvents
        pause 1
        frmMain.ns4Effects.InputFileSave msg
        pause 0.2
        lEffectsPresets.eSaved = True
    Else
        Exit Sub
    End If
End If
If Err.Number <> 0 Then SetError "SaveEffects", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub OpenEffects(lFile As String)
'On Local Error Resume Next

If DoesFileExist(lFile) = True Then
    With frmMain
        EnableEffects
        InitEffects
        DoEvents
        pause 1
        frmMain.tmrScrollStatus.Enabled = True
        lEffectsPresets.eStatus = eOpening
        .ns4Effects.InputFileOpen lFile
        .mnuOpenCurrentFile.Enabled = False
        lEffectsPresets.eSaved = False
    End With
End If
lEffectsPresets.eStatus = eOpen
If Err.Number <> 0 Then SetError "OpenEffects", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub PlayEffect()
'On Local Error Resume Next
InitEffects
frmMain.ns4Effects.Play
lEffectsPresets.eStatus = ePlaying
lPlayer.pStatusString = "Playing Wave Audio: " & lPlayer.pLabels.lWavFile & " ... "
frmMain.tmrScrollStatus.Enabled = True
If Err.Number <> 0 Then SetError "PlayEffect()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub AddCFilter(lShortFactor As Integer)
'On Local Error Resume Next
InitEffects
frmMain.ns4Effects.CFilter lShortFactor
If Err.Number <> 0 Then SetError "AddCFilter", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub AddDistortion(lDryOut As Integer, lDistortedOut As Integer, lThreshholdLevel As Integer, lClampLevel As Integer, lGate As Integer)
'On Local Error Resume Next
InitEffects
frmMain.ns4Effects.Distortion lDryOut, lDistortedOut, lThreshholdLevel, lClampLevel, lGate
If Err.Number <> 0 Then SetError "AddDistortion", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub AddInvert()
'On Local Error Resume Next
InitEffects
frmMain.ns4Effects.Invert
If Err.Number <> 0 Then SetError "AddInvert", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub AddAmplitude()
'On Local Error Resume Next
InitEffects
frmMain.ns4Effects.Amplitude
If Err.Number <> 0 Then SetError "AddAmplitude", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub AddChorus(lDelay As Integer, lDepth As Integer, lRate As Single, lWavForm As Integer, lDry As Integer, lWet As Integer, lInvertFeedback As Integer, lMixing As Integer, lFeedback As Integer)
'On Local Error Resume Next
InitEffects
frmMain.ns4Effects.Chorus lDelay, lDepth, lRate, lWavForm, lDry, lWet, lInvertFeedback, lMixing, lFeedback
If Err.Number <> 0 Then SetError "AddChorus", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub AddShifting(lMode As Integer, lSize As Long)
'On Local Error Resume Next
InitEffects
frmMain.ns4Effects.Shifting lMode, lSize
If Err.Number <> 0 Then SetError "AddShifting", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub AddReverb(lShortDelay As Integer, lShortRatio As Integer)
'On Local Error Resume Next
InitEffects
frmMain.ns4Effects.Reverb lShortDelay, lShortRatio
If Err.Number <> 0 Then SetError "AddReverb", lEvents.eSettings.iErrDescription, Err.Description
End Sub
