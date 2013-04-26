'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict Off
Imports System.IO
Imports System.Runtime.InteropServices

Public Class clsWaveReader
    Inherits Stream
    Private waveFile As String
    Private hMmio As IntPtr = IntPtr.Zero
    Private disposed As Boolean = False
    Private format As clsWinMMInterop.WAVEFORMATEX
    Private dataOffset As Integer = 0
    Private audioLength As Integer = 0

    Public Overloads Sub Dispose()
        DisposeResources(True)
        GC.SuppressFinalize(Me)
    End Sub

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal file As String)
        Me.New()
        Filename = file
    End Sub

    Protected Overrides Sub Finalize()
        DisposeResources(False)
    End Sub


    Protected Overridable ReadOnly Property Handle() As IntPtr
        Get
            Return hMmio
        End Get
    End Property

    ''' <summary>
    ''' Clears up resources associated with this class.
    ''' </summary>
    ''' <param name="disposing"><code>true</code> if disposing from the <c>Dispose</c>
    ''' method, otherwise <c>false</c>.</param>
    Protected Overridable Sub DisposeResources(ByVal disposing As Boolean)
        If Not (disposed) Then
            If (disposing) Then
                '// nothing to do
            End If
            CloseWaveFile()
            disposed = True
        End If
    End Sub

    ''' <summary>
    ''' Gets/sets the wave file name.
    ''' </summary>
    Public Property Filename() As String
        Get
            Return waveFile
        End Get
        Set(ByVal Value As String)
            If Not (hMmio.Equals(IntPtr.Zero)) Then
                CloseWaveFile()
            End If
            waveFile = Value
            OpenWaveFile()
        End Set
    End Property

    ''' <summary>
    ''' Gets the number of audio channels in the file.
    ''' </summary>
    Public ReadOnly Property Channels() As Int16
        Get
            Return format.nChannels
        End Get
    End Property

    ''' <summary>
    ''' Gets the sample frequency of the file.
    ''' </summary>
    Public ReadOnly Property SamplingFrequency() As Integer
        Get
            Return format.nSamplesPerSec
        End Get
    End Property

    ''' <summary>
    ''' Gets the number of bits per sample in the wave file.
    ''' </summary>
    Public ReadOnly Property BitsPerSample() As Int16
        Get
            Return format.wBitsPerSample
        End Get
    End Property


    ''' <summary>
    ''' Gets whether the stream can be read or not (true whenever a wave file
    ''' is open).
    ''' </summary>
    Public Overrides ReadOnly Property CanRead() As Boolean
        Get
            Return Not (hMmio.Equals(IntPtr.Zero))
        End Get
    End Property

    ''' <summary>
    ''' Gets whether the stream is seekable or not (true whenever a wave file
    ''' is open)
    ''' </summary>
    Public Overrides ReadOnly Property CanSeek() As Boolean
        Get
            Return Not (hMmio.Equals(IntPtr.Zero))
        End Get
    End Property

    ''' <summary>
    ''' Returns false; this is a read-only stream
    ''' </summary>
    Public Overrides ReadOnly Property CanWrite() As Boolean
        Get
            Return False
        End Get
    End Property

    ''' <summary>
    ''' Throws an exception; this is a read-only stream
    ''' </summary>
    ''' <exception cref="InvalidOperationException">Thrown exception</exception>
    Public Overrides Sub Flush()
        Throw New InvalidOperationException( _
            "This class can only read files.  Use the WaveStreamWriter class to write files.")
    End Sub

    ''' <summary>
    ''' Gets the length of this wave file, in bytes.
    ''' </summary>
    Public Overrides ReadOnly Property Length() As Long
        Get
            Return audioLength
        End Get
    End Property

    ''' <summary>
    ''' Throws an exception; this is a read-only stream
    ''' </summary>
    ''' <exception cref="InvalidOperationException">Thrown exception</exception>
    Public Overrides Sub SetLength(ByVal length As Long)
        Throw New InvalidOperationException( _
         "This class can only read files.  Use the WaveStreamWriter class to write files.")
    End Sub

    ''' <summary>
    ''' Gets/sets the position within the wave file.
    ''' </summary>
    Public Overrides Property Position() As Long
        Get
            Return 0
        End Get
        Set(ByVal Value As Long)
            Seek(Value, SeekOrigin.Begin)
        End Set
    End Property

    ''' <summary>
    ''' Reads <c>count</c> bytes into the buffer.
    ''' </summary>
    ''' <param name="buffer">Buffer to read into</param>
    ''' <param name="count">Number of bytes to read</param>
    ''' <returns>Number of bytes read.</returns>
    Public Overridable Overloads Function Read(ByVal buffer As Byte(), ByVal count As Integer) As Integer
        Return Read(buffer, 0, count)
    End Function

    ''' <summary>
    ''' Reads <c>count</c> bytes into the buffer.
    ''' </summary>
    ''' <param name="buffer">Buffer to read into</param>
    ''' <param name="count">Number of bytes to read</param>
    ''' <param name="offset">Offset from the current file position to start reading from</param>
    ''' <returns>Number of bytes read.</returns>
    Public Overloads Overrides Function Read(ByVal buffer As Byte(), ByVal offset As Integer, ByVal count As Integer) As Integer

        If (hMmio.Equals(IntPtr.Zero)) Then
            Throw New InvalidOperationException("No wave data is open")
        End If

        If (offset <> 0) Then
            Seek(offset, SeekOrigin.Current)
        End If

        Dim handle As GCHandle = GCHandle.Alloc(buffer, GCHandleType.Pinned)
        Dim ptrBuffer As IntPtr = handle.AddrOfPinnedObject()

        Dim dataRemaining As Integer = (dataOffset + audioLength - _
            clsWinMMInterop.mmioSeek(hMmio, 0, clsWinMMInterop.SEEK_CUR))
        Dim amtRead As Integer = 0
        If (count < dataRemaining) Then
            amtRead = clsWinMMInterop.mmioRead(hMmio, ptrBuffer, count)
        ElseIf (dataRemaining > 0) Then
            amtRead = clsWinMMInterop.mmioRead(hMmio, ptrBuffer, dataRemaining)
        End If

        If (handle.IsAllocated) Then
            handle.Free()
        End If
        Return amtRead
    End Function

    ''' <summary>
    ''' Reads <c>count</c> shorts into the buffer.
    ''' </summary>
    ''' <param name="buffer">Buffer to read into</param>
    ''' <param name="count">Number of shorts to read</param>
    ''' <returns>Number of bytes read.</returns>
    Public Overridable Function Read16bit(ByVal buffer As Short(), ByVal count As Integer) As Integer
        Return Read16bit(buffer, 0, count)
    End Function

    ''' <summary>
    ''' Reads <c>count</c> shorts into the buffer.
    ''' </summary>
    ''' <param name="buffer">Buffer to read into</param>
    ''' <param name="count">Number of shorts to read</param>
    ''' <param name="offset">Offset in shorts (2 bytes) from the current file position to start 
    ''' reading from</param>
    ''' <returns>Number of bytes read.</returns>
    Public Overridable Function Read16bit(ByVal buffer As Short(), ByVal offset As Integer, ByVal count As Integer) As Integer
        If (hMmio.Equals(IntPtr.Zero)) Then
            Throw New InvalidOperationException("No wave data is open")
        End If

        If (offset <> 0) Then
            Seek((offset / 2), SeekOrigin.Current)
        End If

        Dim handle As GCHandle = GCHandle.Alloc(buffer, GCHandleType.Pinned)
        Dim ptrBuffer As IntPtr = handle.AddrOfPinnedObject()

        Dim dataRemaining As Integer = (dataOffset + audioLength - _
          clsWinMMInterop.mmioSeek(hMmio, 0, clsWinMMInterop.SEEK_CUR)) / 2
        Dim read = 0
        If (count < dataRemaining) Then
            read = clsWinMMInterop.mmioRead(hMmio, ptrBuffer, count * 2)
        ElseIf (dataRemaining > 0) Then
            read = clsWinMMInterop.mmioRead(hMmio, ptrBuffer, dataRemaining * 2)
        End If

        If (handle.IsAllocated) Then
            handle.Free()
        End If
        Return read / 2
    End Function

    ''' <summary>
    ''' Throws an exception; this is a read-only stream
    ''' </summary>
    ''' <exception cref="InvalidOperationException">Thrown exception</exception>
    Public Overrides Sub Write(ByVal buffer As Byte(), ByVal offset As Integer, ByVal count As Integer)
        Throw New InvalidOperationException( _
         "This class can only read files.  Use the WaveStreamWriter class to write wave files.")
    End Sub

    ''' <summary>
    ''' Seeks to the specified position in the stream, in bytes
    ''' </summary>
    ''' <param name="position">Position to seek to</param>
    ''' <param name="origin">Specifies the starting postion of the seek</param>
    Public Overrides Function Seek(ByVal position As Long, ByVal origin As SeekOrigin) As Long
        If (hMmio.Equals(IntPtr.Zero)) Then
            Throw New InvalidOperationException("No wave data is open")
        End If

        Dim offset As Integer = position
        Dim mmOrigin As Integer = clsWinMMInterop.SEEK_CUR
        If (origin = SeekOrigin.Begin) Then
            mmOrigin = clsWinMMInterop.SEEK_SET
        ElseIf (origin = SeekOrigin.End) Then
            mmOrigin = clsWinMMInterop.SEEK_END
        End If
        Dim result As Integer = clsWinMMInterop.mmioSeek(hMmio, offset, mmOrigin)
        If (result = -1) Then
            Throw New clsWaveException( _
             String.Format("Failed to seek to position {0} in file", position))
        End If
        Return result
    End Function

    Private Sub OpenWaveFile()
        CloseWaveFile()

        If Not (File.Exists(waveFile)) Then
            Throw New FileNotFoundException( _
                String.Format("The file {0} does not exist", waveFile))
        End If

        hMmio = clsWinMMInterop.mmioOpen(waveFile, IntPtr.Zero, clsWinMMInterop.MMIO_READ)
        If (hMmio.Equals(IntPtr.Zero)) Then
            Throw New IOException(String.Format("Could not open file {0}", waveFile))
        End If
        GetWaveData()
    End Sub

    Private Sub GetWaveData()
        Dim result As Integer = 0
        Dim mmckInfoParent As clsWinMMInterop.MMCKINFO = New clsWinMMInterop.MMCKINFO()
        mmckInfoParent.fccType = clsWinMMInterop.mmioStringToFOURCC("WAVE", 0)
        result = clsWinMMInterop.mmioDescendParent(hMmio, mmckInfoParent, 0, _
            clsWinMMInterop.MMIO_FINDRIFF)
        If (result <> clsWinMMInterop.MMSYSERR_NOERROR) Then
            CloseWaveFile()
            Throw New clsWaveException( _
             String.Format("The file {0} is not a wave file", waveFile))
        End If
        Dim mmckSubChunkIn As clsWinMMInterop.MMCKINFO = New clsWinMMInterop.MMCKINFO()
        mmckSubChunkIn.ckid = clsWinMMInterop.mmioStringToFOURCC("fmt", 0)
        result = clsWinMMInterop.mmioDescend(hMmio, mmckSubChunkIn, mmckInfoParent, clsWinMMInterop.MMIO_FINDCHUNK)
        If (result <> clsWinMMInterop.MMSYSERR_NOERROR) Then
            CloseWaveFile()
            Throw New clsWaveException( _
             String.Format("Unable to locate the format chunk in file {0}", waveFile))
        End If
        format = New clsWinMMInterop.WAVEFORMATEX()
        result = clsWinMMInterop.mmioReadWaveFormat(hMmio, format, mmckSubChunkIn.ckSize)
        If (result = -1) Then
            CloseWaveFile()
            Throw New clsWaveException( _
                String.Format("Unable to read the wave format from file {0}", waveFile))
        End If
        result = clsWinMMInterop.mmioAscend(hMmio, mmckSubChunkIn, 0)
        mmckSubChunkIn.ckid = clsWinMMInterop.mmioStringToFOURCC("data", 0)
        result = clsWinMMInterop.mmioDescend(hMmio, mmckSubChunkIn, mmckInfoParent, _
          clsWinMMInterop.MMIO_FINDCHUNK)
        If (result <> clsWinMMInterop.MMSYSERR_NOERROR) Then
            CloseWaveFile()
            Throw New clsWaveException( _
             String.Format("Unable to locate the data chunk in file {0}", waveFile))
        End If
        dataOffset = clsWinMMInterop.mmioSeek(hMmio, 0, clsWinMMInterop.SEEK_CUR)
        audioLength = mmckSubChunkIn.ckSize
    End Sub

    Private Sub CloseWaveFile()
        If Not (hMmio.Equals(IntPtr.Zero)) Then
            clsWinMMInterop.mmioClose(hMmio, 0)
            hMmio = IntPtr.Zero
            audioLength = 0
        End If
    End Sub
End Class