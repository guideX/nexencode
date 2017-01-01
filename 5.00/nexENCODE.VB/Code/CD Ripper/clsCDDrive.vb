'nexENCODE Studio 5.0 Alpha 1.3
'October 6th, 2013
Option Explicit On
Option Strict Off
Imports System
Imports System.Runtime.InteropServices

Namespace nexENCODE.CDRipper
    Public Class clsCDBufferFiller
        Private _bufferArray As Byte()
        Private _writePosition As Integer = 0

        Public Sub New(buffer As Byte())
            _bufferArray = buffer
        End Sub

        Public Sub OnCdDataRead(sender As Object, e As CDDriveEvents.DataReadEventArgs)
            Buffer.BlockCopy(e.Data, 0, _bufferArray, _writePosition, CInt(e.DataSize))
            _writePosition = _writePosition + CInt(e.DataSize)
        End Sub
    End Class

    Public Class clsCDDrive
        Public Event CDInserted As EventHandler
        Public Event CDRemoved As EventHandler
        Private _handle As IntPtr
        Private lTocValid As Boolean = False
        Private lToc As clsWin32Functions.CDROM_TOC = Nothing
        Private lDrive As Char = ControlChars.NullChar
        Private lNotWnd As CDDriveEvents.DeviceChangeNotificationWindow = Nothing
        Protected Const NSECTORS As Integer = 13
        Protected Const UNDERSAMPLING As Integer = 1
        Protected Const CB_CDDASECTOR As Integer = 2368
        Protected Const CB_QSUBCHANNEL As Integer = 16
        Protected Const CB_CDROMSECTOR As Integer = 2048
        Protected Const CB_AUDIO As Integer = (CB_CDDASECTOR - CB_QSUBCHANNEL)

        Public Sub New()
            lToc = New clsWin32Functions.CDROM_TOC()
            _handle = IntPtr.Zero
        End Sub

        Public Function Open(_Drive As Char) As Boolean
            Close()
            If clsWin32Functions.GetDriveType(_Drive + ":\") = clsWin32Functions.DriveTypes.DRIVE_CDROM Then
                _handle = clsWin32Functions.CreateFile("\\.\" + _Drive + ":"c, clsWin32Functions.GENERIC_READ, clsWin32Functions.FILE_SHARE_READ, IntPtr.Zero, clsWin32Functions.OPEN_EXISTING, 0, _
                 IntPtr.Zero)
                If (CInt(_handle) <> -1) AndAlso (CInt(_handle) <> 0) Then
                    lDrive = _Drive
                    lNotWnd = New CDDriveEvents.DeviceChangeNotificationWindow()
                    Return True
                Else
                    Return True
                End If
            Else
                Return False
            End If
        End Function

        Public Sub Close()
            UnLockCD()
            If lNotWnd IsNot Nothing Then
                lNotWnd.DestroyHandle()
                lNotWnd = Nothing
            End If
            If (CInt(_handle) <> -1) AndAlso (CInt(_handle) <> 0) Then
                clsWin32Functions.CloseHandle(_handle)
            End If
            _handle = IntPtr.Zero
            lDrive = ControlChars.NullChar
            lTocValid = False
        End Sub

        Public ReadOnly Property IsOpened() As Boolean
            Get
                Return (CInt(_handle) <> -1) AndAlso (CInt(_handle) <> 0)
            End Get
        End Property

        Public Sub Dispose()
            Close()
            GC.SuppressFinalize(Me)
        End Sub

        Protected Overrides Sub Finalize()
            Dispose()
        End Sub

        Protected Function ReadTOC() As Boolean
            If (CInt(_handle) <> -1) AndAlso (CInt(_handle) <> 0) Then
                Dim _BytesRead As UInteger = 0
                lTocValid = clsWin32Functions.DeviceIoControl(_handle, clsWin32Functions.IOCTL_CDROM_READ_TOC, IntPtr.Zero, 0, lToc, CUInt(Marshal.SizeOf(lToc)), _BytesRead, IntPtr.Zero) <> 0
            Else
                lTocValid = False
            End If
            Return lTocValid
        End Function

        Protected Function GetStartSector(_Track As Integer) As Integer
            If lTocValid AndAlso (_Track >= lToc.FirstTrack) AndAlso (_Track <= lToc.LastTrack) Then
                Dim _TD As clsWin32Functions.TRACK_DATA = lToc.TrackData(_Track - 1)
                Return (_TD.Address_1 * 60 * 75 + _TD.Address_2 * 75 + _TD.Address_3) - 150
            Else
                Return -1
            End If
        End Function

        Protected Function GetEndSector(_Track As Integer) As Integer
            If lTocValid AndAlso (_Track >= lToc.FirstTrack) AndAlso (_Track <= lToc.LastTrack) Then
                Dim _TD As clsWin32Functions.TRACK_DATA = lToc.TrackData(_Track)
                Return (_TD.Address_1 * 60 * 75 + _TD.Address_2 * 75 + _TD.Address_3) - 151
            Else
                Return -1
            End If
        End Function

        Protected Function ReadSector(sector As Integer, Buffer As Byte(), NumSectors As Integer) As Boolean
            If lTocValid AndAlso ((sector + NumSectors) <= GetEndSector(lToc.LastTrack)) AndAlso (Buffer.Length >= CB_AUDIO * NumSectors) Then
                Dim rri As New clsWin32Functions.RAW_READ_INFO()
                rri.TrackMode = clsWin32Functions.TRACK_MODE_TYPE.CDDA
                rri.SectorCount = CUInt(NumSectors)
                rri.DiskOffset = sector * CB_CDROMSECTOR
                Dim BytesRead As UInteger = 0
                If clsWin32Functions.DeviceIoControl(_handle, clsWin32Functions.IOCTL_CDROM_RAW_READ, rri, CUInt(Marshal.SizeOf(rri)), Buffer, CUInt(NumSectors) * CB_AUDIO, BytesRead, IntPtr.Zero) <> 0 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        End Function

        Public Function LockCD() As Boolean
            If (CInt(_handle) <> -1) AndAlso (CInt(_handle) <> 0) Then
                Dim Dummy As UInteger = 0
                Dim pmr As New clsWin32Functions.PREVENT_MEDIA_REMOVAL()
                pmr.PreventMediaRemoval = 1
                Return clsWin32Functions.DeviceIoControl(_handle, clsWin32Functions.IOCTL_STORAGE_MEDIA_REMOVAL, pmr, CUInt(Marshal.SizeOf(pmr)), IntPtr.Zero, 0, _
                 Dummy, IntPtr.Zero) <> 0
            Else
                Return False
            End If
        End Function

        Public Function UnLockCD() As Boolean
            If (CInt(_handle) <> -1) AndAlso (CInt(_handle) <> 0) Then
                Dim Dummy As UInteger = 0
                Dim pmr As New clsWin32Functions.PREVENT_MEDIA_REMOVAL()
                pmr.PreventMediaRemoval = 0
                Return clsWin32Functions.DeviceIoControl(_handle, clsWin32Functions.IOCTL_STORAGE_MEDIA_REMOVAL, pmr, CUInt(Marshal.SizeOf(pmr)), IntPtr.Zero, 0, _
                 Dummy, IntPtr.Zero) <> 0
            Else
                Return False
            End If
        End Function

        Public Function LoadCD() As Boolean
            lTocValid = False
            If (CInt(_handle) <> -1) AndAlso (CInt(_handle) <> 0) Then
                Dim Dummy As UInteger = 0
                Return clsWin32Functions.DeviceIoControl(_handle, clsWin32Functions.IOCTL_STORAGE_LOAD_MEDIA, IntPtr.Zero, 0, IntPtr.Zero, 0, _
                 Dummy, IntPtr.Zero) <> 0
            Else
                Return False
            End If
        End Function

        Public Function EjectCD() As Boolean
            lTocValid = False
            If (CInt(_handle) <> -1) AndAlso (CInt(_handle) <> 0) Then
                Dim Dummy As UInteger = 0
                Return clsWin32Functions.DeviceIoControl(_handle, clsWin32Functions.IOCTL_STORAGE_EJECT_MEDIA, IntPtr.Zero, 0, IntPtr.Zero, 0, _
                 Dummy, IntPtr.Zero) <> 0
            Else
                Return False
            End If
        End Function

        Public Function IsCDReady() As Boolean
            If (CInt(_handle) <> -1) AndAlso (CInt(_handle) <> 0) Then
                Dim Dummy As UInteger = 0
                If clsWin32Functions.DeviceIoControl(_handle, clsWin32Functions.IOCTL_STORAGE_CHECK_VERIFY, IntPtr.Zero, 0, IntPtr.Zero, 0, _
                 Dummy, IntPtr.Zero) <> 0 Then
                    Return True
                Else
                    lTocValid = False
                    Return False
                End If
            Else
                lTocValid = False
                Return False
            End If
        End Function

        Public Function Refresh() As Boolean
            If IsCDReady() Then
                Return ReadTOC()
            Else
                Return False
            End If
        End Function

        Public Function GetNumTracks() As Integer
            If lTocValid Then
                Return lToc.LastTrack - lToc.FirstTrack + 1
            Else
                Return -1
            End If
        End Function

        Public Function GetNumAudioTracks() As Integer
            If lTocValid Then
                Dim tracks As Integer = 0
                For i As Integer = lToc.FirstTrack - 1 To lToc.LastTrack - 1
                    If lToc.TrackData(i).Control = 0 Then
                        tracks += 1
                    End If
                Next
                Return tracks
            Else
                Return -1
            End If

        End Function

        Public Function ReadTrack(track As Integer, Data As Byte(), ByRef DataSize As UInteger, StartSecond As UInteger, Seconds2Read As UInteger, ProgressEvent As CDDriveEvents.CdReadProgressEventHandler) As Integer
            If lTocValid AndAlso (track >= lToc.FirstTrack) AndAlso (track <= lToc.LastTrack) Then
                Dim StartSect As Integer = GetStartSector(track)
                Dim EndSect As Integer = GetEndSector(track)
                If CInt(StartSect = StartSect + CInt(StartSecond) * 75) >= EndSect Then
                    StartSect -= CInt(StartSecond) * 75
                End If
                If (Seconds2Read > 0) AndAlso (CInt(StartSect + Seconds2Read * 75) < EndSect) Then
                    EndSect = StartSect + CInt(Seconds2Read) * 75
                End If
                DataSize = CUInt(CUInt(EndSect - StartSect) * CB_AUDIO)
                If Data IsNot Nothing Then
                    If Data.Length >= DataSize Then
                        Dim BufferFiller As New clsCDBufferFiller(Data)
                        Return ReadTrack(track, New CDDriveEvents.CdDataReadEventHandler(AddressOf BufferFiller.OnCdDataRead), StartSecond, Seconds2Read, ProgressEvent)
                    Else
                        Return 0
                    End If
                Else
                    Return 0
                End If
            Else
                Return -1
            End If
        End Function

        Public Function ReadTrack(track As Integer, Data As Byte(), ByRef DataSize As UInteger, ProgressEvent As CDDriveEvents.CdReadProgressEventHandler) As Integer
            Return ReadTrack(track, Data, DataSize, 0, 0, ProgressEvent)
        End Function

        Public Function ReadTrack(track As Integer, DataReadEvent As CDDriveEvents.CdDataReadEventHandler, StartSecond As UInteger, Seconds2Read As UInteger, ProgressEvent As CDDriveEvents.CdReadProgressEventHandler) As Integer
            If lTocValid AndAlso (track >= lToc.FirstTrack) AndAlso (track <= lToc.LastTrack) AndAlso (DataReadEvent IsNot Nothing) Then
                Dim StartSect As Integer = GetStartSector(track)
                Dim EndSect As Integer = GetEndSector(track)
                If CInt(StartSect = StartSect + CInt(StartSecond) * 75) >= EndSect Then
                    StartSect -= CInt(StartSecond) * 75
                End If
                If (Seconds2Read > 0) AndAlso (CInt(StartSect + Seconds2Read * 75) < EndSect) Then
                    EndSect = StartSect + CInt(Seconds2Read) * 75
                End If
                Dim Bytes2Read As UInteger = CUInt(CUInt(EndSect - StartSect) * CB_AUDIO)
                Dim BytesRead As UInteger = 0
                Dim Data As Byte() = New Byte(CB_AUDIO * NSECTORS - 1) {}
                Dim Cont As Boolean = True
                Dim ReadOk As Boolean = True
                If ProgressEvent IsNot Nothing Then
                    Dim rpa As New CDDriveEvents.ReadProgressEventArgs(Bytes2Read, 0)
                    ProgressEvent(Me, rpa)
                    Cont = Not rpa.CancelRead
                End If
                Dim sector As Integer = StartSect
                While (sector < EndSect) AndAlso (Cont) AndAlso (ReadOk)
                    Dim Sectors2Read As Integer = If(((sector + NSECTORS) < EndSect), NSECTORS, (EndSect - sector))
                    ReadOk = ReadSector(sector, Data, Sectors2Read)
                    If ReadOk Then
                        Dim dra As New CDDriveEvents.DataReadEventArgs(Data, CUInt(CB_AUDIO * Sectors2Read))
                        DataReadEvent(Me, dra)
                        BytesRead += CUInt(CB_AUDIO * Sectors2Read)
                        If ProgressEvent IsNot Nothing Then
                            Dim rpa As New CDDriveEvents.ReadProgressEventArgs(Bytes2Read, BytesRead)
                            ProgressEvent(Me, rpa)
                            Cont = Not rpa.CancelRead
                        End If
                    End If
                    sector += NSECTORS
                End While
                If ReadOk Then
                    Return CInt(BytesRead)
                Else
                    Return -1
                End If
            Else
                Return -1
            End If
        End Function

        Public Function ReadTrack(track As Integer, DataReadEvent As CDDriveEvents.CdDataReadEventHandler, ProgressEvent As CDDriveEvents.CdReadProgressEventHandler) As Integer
            Return ReadTrack(track, DataReadEvent, 0, 0, ProgressEvent)
        End Function

        Public Function TrackSize(track As Integer) As UInteger
            Dim Size As UInteger = 0
            ReadTrack(track, Nothing, Size, Nothing)
            Return Size
        End Function

        Public Function IsAudioTrack(track As Integer) As Boolean
            If (lTocValid) AndAlso (track >= lToc.FirstTrack) AndAlso (track <= lToc.LastTrack) Then
                Return (lToc.TrackData(track - 1).Control And 4) = 0
            Else
                Return False
            End If
        End Function

        Public Shared Function GetCDDriveLetters() As Char()
            Dim _Res As String = ""
            For Each c In "DEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray()
                If clsWin32Functions.GetDriveType(c + ":") = clsWin32Functions.DriveTypes.DRIVE_CDROM Then
                    _Res += c
                End If
            Next
            Return _Res.ToCharArray()
        End Function

        Private Sub OnCDInserted()
            RaiseEvent CDInserted(Me, EventArgs.Empty)
        End Sub

        Private Sub OnCDRemoved()
            RaiseEvent CDRemoved(Me, EventArgs.Empty)
        End Sub

        Private Sub NotWnd_DeviceChange(sender As Object, e As CDDriveEvents.DeviceChangeEventArgs)
            If e.Drive = lDrive Then
                lTocValid = False
                Select Case e.ChangeType
                    Case CDDriveEvents.DeviceChangeEventType.DeviceInserted
                        OnCDInserted()
                        Exit Select
                    Case CDDriveEvents.DeviceChangeEventType.DeviceRemoved
                        OnCDRemoved()
                        Exit Select
                End Select
            End If
        End Sub
    End Class
End Namespace