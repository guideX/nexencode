'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Imports System.Threading
Imports NAudio.CoreAudioApi.Interfaces
Imports System.Runtime.InteropServices
Imports NAudio.Wave

Namespace NAudio.CoreAudioApi
    Public Class AudioClient
        Implements IDisposable
        Private audioClientInterface As IAudioClient
        Private m_mixFormat As WaveFormat
        Private m_audioRenderClient As AudioRenderClient
        Private m_audioCaptureClient As AudioCaptureClient

        Friend Sub New(audioClientInterface As IAudioClient)
            Me.audioClientInterface = audioClientInterface
        End Sub

        Public ReadOnly Property MixFormat() As WaveFormat
            Get
                If m_mixFormat Is Nothing Then
                    Dim waveFormatPointer As IntPtr
                    Marshal.ThrowExceptionForHR(audioClientInterface.GetMixFormat(waveFormatPointer))
                    'WaveFormatExtensible waveFormat = new WaveFormatExtensible(44100,32,2);
                    'Marshal.PtrToStructure(waveFormatPointer, waveFormat);
                    Dim waveFormat__1 As WaveFormat = WaveFormat.MarshalFromPtr(waveFormatPointer)
                    Marshal.FreeCoTaskMem(waveFormatPointer)
                    m_mixFormat = waveFormat__1
                    Return waveFormat__1
                Else
                    Return m_mixFormat
                End If
            End Get
        End Property

        Public Sub Initialize(shareMode As AudioClientShareMode, streamFlags As AudioClientStreamFlags, bufferDuration As Long, periodicity As Long, waveFormat As WaveFormat, audioSessionGuid As Guid)
            Dim hresult As Integer = audioClientInterface.Initialize(shareMode, streamFlags, bufferDuration, periodicity, waveFormat, audioSessionGuid)
            Marshal.ThrowExceptionForHR(hresult)
            ' may have changed the mix format so reset it
            m_mixFormat = Nothing
        End Sub

        Public ReadOnly Property BufferSize() As Integer
            Get
                Dim bufferSize__1 As UInteger
                Marshal.ThrowExceptionForHR(audioClientInterface.GetBufferSize(bufferSize__1))
                Return CInt(bufferSize__1)
            End Get
        End Property

        Public ReadOnly Property StreamLatency() As Long
            Get
                Return audioClientInterface.GetStreamLatency()
            End Get
        End Property

        Public ReadOnly Property CurrentPadding() As Integer
            Get
                Dim currentPadding__1 As Integer
                Marshal.ThrowExceptionForHR(audioClientInterface.GetCurrentPadding(currentPadding__1))
                Return currentPadding__1
            End Get
        End Property

        Public ReadOnly Property DefaultDevicePeriod() As Long
            Get
                Dim defaultDevicePeriod__1 As Long
                Dim minimumDevicePeriod As Long
                Marshal.ThrowExceptionForHR(audioClientInterface.GetDevicePeriod(defaultDevicePeriod__1, minimumDevicePeriod))
                Return defaultDevicePeriod__1
            End Get
        End Property

        Public ReadOnly Property MinimumDevicePeriod() As Long
            Get
                Dim defaultDevicePeriod As Long
                Dim minimumDevicePeriod__1 As Long
                Marshal.ThrowExceptionForHR(audioClientInterface.GetDevicePeriod(defaultDevicePeriod, minimumDevicePeriod__1))
                Return minimumDevicePeriod__1
            End Get
        End Property

        ' TODO: GetService:
        ' IID_IAudioCaptureClient
        ' IID_IAudioClock
        ' IID_IAudioSessionControl
        ' IID_IAudioStreamVolume
        ' IID_IChannelAudioVolume
        ' IID_ISimpleAudioVolume

        Public ReadOnly Property AudioRenderClient() As AudioRenderClient
            Get
                If m_audioRenderClient Is Nothing Then
                    Dim audioRenderClientInterface As Object
                    Dim audioRenderClientGuid As New Guid("F294ACFC-3146-4483-A7BF-ADDCA7C260E2")
                    Marshal.ThrowExceptionForHR(audioClientInterface.GetService(audioRenderClientGuid, audioRenderClientInterface))
                    m_audioRenderClient = New AudioRenderClient(DirectCast(audioRenderClientInterface, IAudioRenderClient))
                End If
                Return m_audioRenderClient
            End Get
        End Property

        Public ReadOnly Property AudioCaptureClient() As AudioCaptureClient
            Get
                If m_audioCaptureClient Is Nothing Then
                    Dim audioCaptureClientInterface As Object
                    Dim audioCaptureClientGuid As New Guid("c8adbd64-e71e-48a0-a4de-185c395cd317")
                    Marshal.ThrowExceptionForHR(audioClientInterface.GetService(audioCaptureClientGuid, audioCaptureClientInterface))
                    m_audioCaptureClient = New AudioCaptureClient(DirectCast(audioCaptureClientInterface, IAudioCaptureClient))
                End If
                Return m_audioCaptureClient
            End Get
        End Property

        Public Function IsFormatSupported(shareMode As AudioClientShareMode, desiredFormat As WaveFormat) As Boolean
            Dim closestMatchFormat As WaveFormatExtensible
            Return IsFormatSupported(shareMode, desiredFormat, closestMatchFormat)
        End Function

        Public Function IsFormatSupported(shareMode As AudioClientShareMode, desiredFormat As WaveFormat, ByRef closestMatchFormat As WaveFormatExtensible) As Boolean
            Dim hresult As Integer = audioClientInterface.IsFormatSupported(shareMode, desiredFormat, closestMatchFormat)
            ' S_OK is 0, S_FALSE = 1
            If hresult = 0 Then
                ' directly supported
                Return True
            End If
            If hresult = 1 Then
                Return False
            ElseIf hresult = CInt(AudioClientErrors.UnsupportedFormat) Then
                Return False
            Else
                Marshal.ThrowExceptionForHR(hresult)
            End If
            ' shouldn't get here
            Throw New NotSupportedException("Unknown hresult " & hresult.ToString())
        End Function

        Public Sub Start()
            audioClientInterface.Start()
        End Sub

        Public Sub [Stop]()
            audioClientInterface.[Stop]()
        End Sub

        Public Sub SetEventHandle(eventWaitHandle As EventWaitHandle)
            audioClientInterface.SetEventHandle(eventWaitHandle.SafeWaitHandle.DangerousGetHandle())
        End Sub

        Public Sub Reset()
            audioClientInterface.Reset()
        End Sub

#Region "IDisposable Members"
        Public Sub Dispose() Implements IDisposable.Dispose
            If audioClientInterface IsNot Nothing Then
                If m_audioRenderClient IsNot Nothing Then
                    m_audioRenderClient.Dispose()
                    m_audioRenderClient = Nothing
                End If
                If m_audioCaptureClient IsNot Nothing Then
                    m_audioCaptureClient.Dispose()
                    m_audioCaptureClient = Nothing
                End If
                Marshal.ReleaseComObject(audioClientInterface)
                audioClientInterface = Nothing
                GC.SuppressFinalize(Me)
            End If
        End Sub
#End Region
    End Class
End Namespace