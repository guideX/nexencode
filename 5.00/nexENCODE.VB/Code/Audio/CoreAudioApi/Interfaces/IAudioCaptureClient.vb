'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Imports System.Collections.Generic
Imports System.Text
Imports System.Runtime.InteropServices

Namespace NAudio.CoreAudioApi.Interfaces
    <Guid("C8ADBD64-E71E-48a0-A4DE-185C395CD317"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)> Interface IAudioCaptureClient
        Function GetBuffer(ByRef dataBuffer As IntPtr, ByRef numFramesToRead As Integer, ByRef bufferFlags As AudioClientBufferFlags, ByRef devicePosition As Long, ByRef qpcPosition As Long) As Integer
        Function ReleaseBuffer(numFramesRead As Integer) As Integer
        Function GetNextPacketSize(ByRef numFramesInNextPacket As Integer) As Integer
    End Interface
End Namespace