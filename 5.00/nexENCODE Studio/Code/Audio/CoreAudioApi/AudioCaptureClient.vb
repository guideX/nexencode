'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Imports System.Collections.Generic
Imports System.Text
Imports System.Runtime.InteropServices
Imports NAudio.CoreAudioApi.Interfaces
Imports NAudio.Utils

Namespace NAudio.CoreAudioApi
    Public Class AudioCaptureClient
        Implements IDisposable
        Private audioCaptureClientInterface As IAudioCaptureClient

        Friend Sub New(audioCaptureClientInterface As IAudioCaptureClient)
            Me.audioCaptureClientInterface = audioCaptureClientInterface
        End Sub

        Public Function GetBuffer(ByRef numFramesToRead As Integer, ByRef bufferFlags As AudioClientBufferFlags, ByRef devicePosition As Long, ByRef qpcPosition As Long) As IntPtr
            Dim bufferPointer As IntPtr
            Marshal.ThrowExceptionForHR(audioCaptureClientInterface.GetBuffer(bufferPointer, numFramesToRead, bufferFlags, devicePosition, qpcPosition))
            Return bufferPointer
        End Function

        Public Function GetBuffer(ByRef numFramesToRead As Integer, ByRef bufferFlags As AudioClientBufferFlags) As IntPtr
            Dim bufferPointer As IntPtr
            Dim devicePosition As Long
            Dim qpcPosition As Long
            Marshal.ThrowExceptionForHR(audioCaptureClientInterface.GetBuffer(bufferPointer, numFramesToRead, bufferFlags, devicePosition, qpcPosition))
            Return bufferPointer
        End Function

        Public Function GetNextPacketSize() As Integer
            Dim numFramesInNextPacket As Integer
            Marshal.ThrowExceptionForHR(audioCaptureClientInterface.GetNextPacketSize(numFramesInNextPacket))
            Return numFramesInNextPacket
        End Function

        Public Sub ReleaseBuffer(numFramesWritten As Integer)
            Marshal.ThrowExceptionForHR(audioCaptureClientInterface.ReleaseBuffer(numFramesWritten))
        End Sub

#Region "IDisposable Members"
        Public Sub Dispose() Implements IDisposable.Dispose
            If audioCaptureClientInterface IsNot Nothing Then
                Marshal.ReleaseComObject(audioCaptureClientInterface)
                audioCaptureClientInterface = Nothing
                GC.SuppressFinalize(Me)
            End If
        End Sub
#End Region
    End Class
End Namespace