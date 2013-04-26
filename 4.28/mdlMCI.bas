Attribute VB_Name = "mdlMCI"
Option Explicit
Private Const MCI_OPEN = &H803
Private Const MCI_CLOSE = &H804
Private Const MCI_FORMAT_MSF = 2
Private Const MCI_OPEN_ELEMENT = &H200&
Private Const MCI_OPEN_TYPE = &H2000&
Private Const MCI_SET = &H80D
Private Const MCI_SET_TIME_FORMAT = &H400&
Private Const MCI_STATUS = &H814
Private Const MCI_STATUS_ITEM = &H100&
Private Const MCI_STATUS_LENGTH = &H1&
Private Const MCI_STATUS_NUMBER_OF_TRACKS = &H3&
Private Const MCI_STATUS_POSITION = &H2&
Private Const MCI_TRACK = &H10&
Private Type MCI_OPEN_PARMS
    dwCallback As Long
    wDeviceID As Long
    lpstrDeviceType As String
    lpstrElementName As String
    lpstrAlias As String
End Type
Private Type MCI_SET_PARMS
    dwCallback As Long
    dwTimeFormat As Long
    dwAudio As Long
End Type
Private Type MCI_STATUS_PARMS
    dwCallback As Long
    dwReturn As Long
    dwItem As Long
    dwTrack As Integer
End Type
Private Declare Function mciSendCommand Lib "winmm.dll" Alias "mciSendCommandA" (ByVal wDeviceID As Long, ByVal uMessage As Long, ByVal dwParam1 As Long, ByRef dwParam2 As Any) As Long
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciSendCommandA" (ByVal fdwError As Long, ByRef lpszErrorText As String, ByVal cchErrorText As Integer)
Private mciOpenParms As MCI_OPEN_PARMS
Private mciSetParms As MCI_SET_PARMS
Private mciStatusParms As MCI_STATUS_PARMS
Private m_TOC   As String
Private m_DevID As Long

Public Property Get GetTOC() As String
GetTOC = m_TOC
End Property

Public Function InitMediaToc(ByVal strDrive As String) As Boolean
On Local Error Resume Next
strDrive = Trim(strDrive)
If (VerifyCD(strDrive) = True) Then
    If (OpenCD(strDrive) = True) Then
        InitMediaToc = ReadTOC
        CloseCD
        Exit Function
    End If
End If
InitMediaToc = False
If Err.Number <> 0 Then SetError "InitMediaToc", lEvents.eSettings.iErrDescription, Err.Description
End Function

Private Function VerifyCD(ByVal strDrive As String) As Boolean
On Error GoTo errChk
Dim fso As FileSystemObject, fsoDrive As Drive
Set fso = New FileSystemObject
If (fso.DriveExists(strDrive) = False) Then GoTo errChk
Set fsoDrive = fso.GetDrive(strDrive)
strDrive = fsoDrive.Path
If (fso.GetDrive(strDrive).DriveType <> CDRom) Then GoTo errChk
If (fso.GetDrive(strDrive).IsReady = False) Then GoTo errChk
Set fso = Nothing
Set fsoDrive = Nothing
VerifyCD = True
Exit Function

errChk:
    If Err.Number <> 0 Then SetError "VerifyCD", lEvents.eSettings.iErrDescription, Err.Description
    Set fso = Nothing
    Set fsoDrive = Nothing
    VerifyCD = False
End Function

Private Function OpenCD(ByVal strDrive As String) As Boolean
On Error GoTo errChk
Dim idx As Integer, sts As Long
mciOpenParms.lpstrDeviceType = "cdaudio"
mciOpenParms.lpstrElementName = strDrive
sts = mciSendCommand(0, MCI_OPEN, (MCI_OPEN_TYPE Or MCI_OPEN_ELEMENT), mciOpenParms)
If (sts <> 0) Then GoTo errChk
m_DevID = mciOpenParms.wDeviceID
mciSetParms.dwTimeFormat = MCI_FORMAT_MSF
sts = mciSendCommand(m_DevID, MCI_SET, MCI_SET_TIME_FORMAT, mciSetParms)
If (sts <> 0) Then GoTo errChk
OpenCD = True
Exit Function

errChk:
    If Err.Number <> 0 Then SetError "OpenCD", lEvents.eSettings.iErrDescription, Err.Description
    OpenCD = False
End Function

Private Function ReadTOC() As Boolean
On Error GoTo errChk
Dim idx As Integer, trks As Integer, mins As Long, secs As Long, frms As Long, sts As Long, offst As Long, s As String
mciStatusParms.dwItem = MCI_STATUS_NUMBER_OF_TRACKS
sts = mciSendCommand(m_DevID, MCI_STATUS, MCI_STATUS_ITEM, mciStatusParms)
If (sts <> 0) Then GoTo errChk
trks = mciStatusParms.dwReturn
For idx = 1 To trks
    mciStatusParms.dwItem = MCI_STATUS_POSITION
    mciStatusParms.dwTrack = idx
    sts = mciSendCommand(m_DevID, MCI_STATUS, MCI_STATUS_ITEM Or MCI_TRACK, mciStatusParms)
    If (sts <> 0) Then GoTo errChk
    mins = (mciStatusParms.dwReturn) And &HFF
    secs = (mciStatusParms.dwReturn \ 256) And &HFF
    frms = (mciStatusParms.dwReturn \ 65536) And &HFF
    offst = (mins * 60 * 75) + (secs * 75) + frms
    s = s & " " & Format$(offst)
Next idx
mciStatusParms.dwItem = MCI_STATUS_LENGTH
mciStatusParms.dwTrack = trks
sts = mciSendCommand(m_DevID, MCI_STATUS, MCI_STATUS_ITEM Or MCI_TRACK, mciStatusParms)
If (sts <> 0) Then GoTo errChk
mins = (mciStatusParms.dwReturn) And &HFF
secs = (mciStatusParms.dwReturn \ 256) And &HFF
frms = ((mciStatusParms.dwReturn \ 65536) And &HFF) + 1
offst = offst + (mins * 60 * 75) + (secs * 75) + frms
s = s & " " & Format$(offst)
m_TOC = Trim$(s)
ReadTOC = True
Exit Function

errChk:
    If Err.Number <> 0 Then SetError "ReadToc", lEvents.eSettings.iErrDescription, Err.Description
    ReadTOC = False
End Function

Private Sub CloseCD()
On Local Error Resume Next
Dim sts As Long
sts = mciSendCommand(m_DevID, MCI_CLOSE, 0, 0)
If Err.Number <> 0 Then SetError "CloseCD", lEvents.eSettings.iErrDescription, Err.Description
End Sub
