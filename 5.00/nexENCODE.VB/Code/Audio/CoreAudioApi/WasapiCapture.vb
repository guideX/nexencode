'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Imports System.Collections.Generic
Imports System.Text
Imports NAudio.Wave
Imports System.Threading
Imports System.Diagnostics
Imports System.Runtime.InteropServices

Namespace NAudio.CoreAudioApi
    Public Class WasapiCapture
        Implements IWaveIn
        Private Const REFTIMES_PER_SEC As Long = 10000000
        Private Const REFTIMES_PER_MILLISEC As Long = 10000
        Private [stop] As Boolean
        Private recordBuffer As Byte()
        Private captureThread As Thread
        Private audioClient As AudioClient
        Private bytesPerFrame As Integer
        Public Event DataAvailable As EventHandler(Of WaveInEventArgs)
        Public Event RecordingStopped As EventHandler

        Public Sub New()
            Me.New(GetDefaultCaptureDevice())
        End Sub

        Public Sub New(captureDevice As MMDevice)
            Me.audioClient = captureDevice.AudioClient
            WaveFormat = audioClient.MixFormat
        End Sub

        Public Property WaveFormat() As WaveFormat
            Get
                Return m_WaveFormat
            End Get
            Set(value As WaveFormat)
                m_WaveFormat = value
            End Set
        End Property

        Private m_WaveFormat As WaveFormat

        Public Shared Function GetDefaultCaptureDevice() As MMDevice
            Dim devices As New MMDeviceEnumerator()
            Return devices.GetDefaultAudioEndpoint(DataFlow.Capture, Role.Console)
        End Function

        Private Sub InitializeCaptureDevice()
            Dim requestedDuration As Long = REFTIMES_PER_MILLISEC * 100
            If Not audioClient.IsFormatSupported(AudioClientShareMode.[Shared], WaveFormat) Then
                Throw New ArgumentException("Unsupported Wave Format")
            End If
            audioClient.Initialize(AudioClientShareMode.[Shared], AudioClientStreamFlags.None, requestedDuration, 0, WaveFormat, Guid.Empty)
            Dim bufferFrameCount As Integer = audioClient.BufferSize
            bytesPerFrame = WaveFormat.Channels * WaveFormat.BitsPerSample / 8
            recordBuffer = New Byte(bufferFrameCount * bytesPerFrame - 1) {}
            Debug.WriteLine(String.Format("record buffer size = {0}", recordBuffer.Length))
        End Sub

        Public Sub StartRecording()
            InitializeCaptureDevice()
			Dim start As ThreadStart = Function() Do
            Me.captureThread(Me.audioClient)
        End Sub
			Me.captureThread = New Thread(start)

			Debug.WriteLine("Thread starting...")
			Me.[stop] = False
			Me.captureThread.Start()
		End Sub

        Public Sub StopRecording()
            If Me.captureThread IsNot Nothing Then
                Me.[stop] = True
                Debug.WriteLine("Thread ending...")
                Me.captureThread.Join()
                Me.captureThread = Nothing
                Debug.WriteLine("Done.")
                Me.[stop] = False
            End If
        End Sub

        Private Sub CaptureThread(client As AudioClient)
            Debug.WriteLine(client.BufferSize)
            Dim bufferFrameCount As Integer = audioClient.BufferSize
            Dim actualDuration As Long = CLng(CDbl(REFTIMES_PER_SEC) * bufferFrameCount / WaveFormat.SampleRate)
            Dim sleepMilliseconds As Integer = CInt(actualDuration \ REFTIMES_PER_MILLISEC \ 2)
            Dim capture As AudioCaptureClient = client.AudioCaptureClient
            client.Start()
            Try
                Debug.WriteLine(String.Format("sleep: {0} ms", sleepMilliseconds))
                While Not Me.[stop]
                    Thread.Sleep(sleepMilliseconds)
                    ReadNextPacket(capture)
                End While
            Finally
                client.[Stop]()
                RaiseEvent RecordingStopped(Me, EventArgs.Empty)
            End Try
            System.Diagnostics.Debug.WriteLine("stop wasapi")
        End Sub

        Private Sub ReadNextPacket(capture As AudioCaptureClient)
            Dim buffer As IntPtr
            Dim framesAvailable As Integer
            Dim flags As AudioClientBufferFlags
            Dim packetSize As Integer = capture.GetNextPacketSize()
            Dim recordBufferOffset As Integer = 0
            'Debug.WriteLine(string.Format("packet size: {0} samples", packetSize / 4));
            While packetSize <> 0
                buffer = capture.GetBuffer(framesAvailable, flags)
                Dim bytesAvailable As Integer = framesAvailable * bytesPerFrame
                Dim spaceRemaining As Integer = Math.Max(0, recordBuffer.Length - recordBufferOffset)
                If spaceRemaining < bytesAvailable AndAlso recordBufferOffset > 0 Then
                    RaiseEvent DataAvailable(Me, New WaveInEventArgs(recordBuffer, recordBufferOffset))
                    recordBufferOffset = 0
                End If
                If (flags And AudioClientBufferFlags.Silent) <> AudioClientBufferFlags.Silent Then
                    Marshal.Copy(buffer, recordBuffer, recordBufferOffset, bytesAvailable)
                Else
                    Array.Clear(recordBuffer, recordBufferOffset, bytesAvailable)
                End If
                recordBufferOffset += bytesAvailable
                capture.ReleaseBuffer(framesAvailable)
                packetSize = capture.GetNextPacketSize()
            End While
            RaiseEvent DataAvailable(Me, New WaveInEventArgs(recordBuffer, recordBufferOffset))
        End Sub

        Public Sub Dispose()
            StopRecording()
            If audioClient IsNot Nothing Then
                audioClient.Dispose()
                audioClient = Nothing
            End If
        End Sub
    End Class
End Namespace