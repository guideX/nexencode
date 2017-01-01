'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Imports System.Collections.Generic
Imports System.Text
Imports NAudio.CoreAudioApi.Interfaces
Imports System.Runtime.InteropServices

Namespace NAudio.CoreAudioApi
    Public Class AudioRenderClient
        Implements IDisposable
        Private audioRenderClientInterface As IAudioRenderClient

        Friend Sub New(audioRenderClientInterface As IAudioRenderClient)
            Me.audioRenderClientInterface = audioRenderClientInterface
        End Sub

        Public Function GetBuffer(numFramesRequested As Integer) As IntPtr
            Dim bufferPointer As IntPtr
            Marshal.ThrowExceptionForHR(audioRenderClientInterface.GetBuffer(numFramesRequested, bufferPointer))
            Return bufferPointer
        End Function

        Public Sub ReleaseBuffer(numFramesWritten As Integer, bufferFlags As AudioClientBufferFlags)
            Marshal.ThrowExceptionForHR(audioRenderClientInterface.ReleaseBuffer(numFramesWritten, bufferFlags))
        End Sub

#Region "IDisposable Members"
        Public Sub Dispose() Implements IDisposable.Dispose
            If audioRenderClientInterface IsNot Nothing Then
                Marshal.ReleaseComObject(audioRenderClientInterface)
                audioRenderClientInterface = Nothing
                GC.SuppressFinalize(Me)
            End If
        End Sub
#End Region
    End Class
End Namespace