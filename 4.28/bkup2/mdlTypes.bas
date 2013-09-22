Attribute VB_Name = "mdlTypes"
Option Explicit
Enum eEventTypes
    Rip = 1
    Encode = 2
    Id3 = 3
    Effects = 4
    Download = 5
    Upload = 6
    Play = 7
    Decode = 8
    Merge = 9
    'PlayWav = 9
End Enum
Private Type gTrack
    tName As String
    tLength As String
End Type
Private Type gTracks
    tDiscLen As String
    tTitle As String
    tArtist As String
    tYear As String
    tLabel As String
    tGenre As String
    tCount As Integer
    tTrack(300) As gTrack
End Type
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
    eAutoDeleteThisWav As Boolean
End Type
Private Type gCDDBSettings
    cEmailAddress As String
    cAutoSubmit As Boolean
    cUseFirstMatch As Boolean
    cServer As String
    cSaveTracksToDisk As Boolean
    cEnabled As Boolean
    cLoaded As Boolean
    cShowDialog As Boolean
End Type
Private Type gInterfaceSettings
    iCreateAlbumFileOnEncode As Boolean
    iLoading As Boolean
    iCommand As String
    iAlwaysOnTop As Boolean
    iPlayMp3sInNexENCODE As Boolean
    iOverRipper As Boolean
    iOverEncoder As Boolean
    iEnding As Boolean
    iErrDescription As String
    iShowReports As Boolean
    iOverwritePrompts As Boolean
    iShowErrors As Boolean
    iAutoPlay As Boolean
    iPlayWavs As Boolean
    iFreeDB As gCDDBSettings
    iHandleErrors As Boolean
    iShowAbout As Boolean
    iUpdateCheck As Boolean
    iCheckForActiveWindow As Boolean
    iRememberWindowSizes As Boolean
End Type
Private Type gEvents
    eErrCount As Long
    ePlaylistVisible As Boolean
    eMainHWND As String
    ePercent As Integer
    eEventCount As Integer
    eEvent(200) As gEvent
    eCircleNum As Integer
    eEventIndex As Integer
    eEncoderBusy As Boolean
    eRipperBusy As Boolean
    eAudicaLocation As String
    eNexMediaLocation As String
    eAutoDel As String
    eSettings As gInterfaceSettings
    eRetStr As String
    eName As String
    ePassword As String
    eGlobalPass As String
    eRegistered As Boolean
    eTimeType As Integer
    eEventType As Integer
    eEventBusy As Boolean
End Type
Private Type gRipperSettings
    eAvailable As Boolean
    eAutoDeleteRipedFiles As Boolean
    eLockCDTrayDuringRip As Boolean
    eCopyMode As Integer
'    eTempFiles As String
    eDiscID As String
    eDriveLetter As String
    eAutoEject As Boolean
    eAspiEnabled As Boolean
End Type
Private Type gEncoderSettings
    eProfile As Integer
    eSampleRate As Long
    eBitrate As Long
    eDownsample As Boolean
    eCopyrighted As Boolean
    eOrigionalWork As Boolean
    eDownmix As Boolean
    eAutoAddTags As Boolean
    eOutputDir As String
End Type
Enum ePlayerTypes
    pMp3Player = 1
    pWavPlayer = 2
    pCDPlayer = 3
End Enum
Private Type gPlayer
    pName As String
    pPath As String
    pFile As String
    pPlaylist As String
    pType As ePlayerTypes
End Type
Private Type gPlayers
    pPlayer(100) As gPlayer
    pCount As Integer
    pMp3PlayerIndex As Integer
'    pWavPlayerIndex As Integer
    pCDPlayerIndex As Integer
End Type
Private Type gIniFiles
    iEffects As String
    iSettings As String
    iErrors As String
    iPlayers As String
    iPlaylists As String
    iWindowPos As String
    iCD As String
    iCDDBServers As String
    'iWritter As String
    iUpdate As String
End Type
Private Type gSetupWizard
    sFrameCount As Integer
    sFrameIndex As Integer
End Type
Private Type gDrive
    dType As Integer
    dEnabled As Boolean
    dLetter As String
End Type
Private Type gDrives
    dDrive(20) As gDrive
    dCount As Integer
    dHardDrives As String
End Type
Private Type gServer
    sIp As String
    sLocation As String
End Type
Private Type gCDDBServers
    cCount As Integer
    cServer(20) As gServer
End Type
Private Type gReport
    rReportString As String
    rType As eEventTypes
    rFilename As String
    rFilepath As String
End Type
Private Type gReports
    rCount As Integer
    rReport(100) As gReport
End Type
Private Type gWritter
    wEnabled As Boolean
    wDrive As String
End Type
Enum eEncWizardTypes
    eSingleWav = 1
    eMultiWav = 2
End Enum
Private Type gEncWizard
    eFinished As Boolean
    eWizFrame As Integer
    eType As eEncWizardTypes
    eEnabled As Boolean
    eCount As Integer
End Type
Public lEncWizard As gEncWizard

Public lWritter As gWritter
Public lReports As gReports
Public lCDDBServ As gCDDBServers
Public lSetupWizard As gSetupWizard
Public lEncoderSettings As gEncoderSettings
Public lEvents As gEvents
Public lRipperSettings As gRipperSettings
Public lTracks As gTracks
Global lPlayers As gPlayers
Global lIniFiles As gIniFiles
Global lUnloadSetupWizardAfterASPI As Boolean
Global lAutoScanHdd As Boolean

