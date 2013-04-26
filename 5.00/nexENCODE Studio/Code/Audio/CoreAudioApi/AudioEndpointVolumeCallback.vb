'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Imports System.Collections.Generic
Imports System.Text
Imports NAudio.CoreAudioApi.Interfaces
Imports System.Runtime.InteropServices

Namespace NAudio.CoreAudioApi
    Friend Class AudioEndpointVolumeCallback
        Implements IAudioEndpointVolumeCallback
        Private _Parent As AudioEndpointVolume

        Friend Sub New(parent As AudioEndpointVolume)
            _Parent = parent
        End Sub

        Public Function OnNotify(NotifyData As IntPtr) As Integer
            Dim data As AudioVolumeNotificationDataStruct = DirectCast(Marshal.PtrToStructure(NotifyData, GetType(AudioVolumeNotificationDataStruct)), AudioVolumeNotificationDataStruct)
            Dim Offset As IntPtr = Marshal.OffsetOf(GetType(AudioVolumeNotificationDataStruct), "ChannelVolume")
            Dim FirstFloatPtr As IntPtr = CType(CLng(NotifyData) + CLng(Offset), IntPtr)
            Dim voldata As Single() = New Single(data.nChannels - 1) {}
            For i As Integer = 0 To data.nChannels - 1
                voldata(i) = CSng(Marshal.PtrToStructure(FirstFloatPtr, GetType(Single)))
            Next
            Dim NotificationData As New AudioVolumeNotificationData(data.guidEventContext, data.bMuted, data.fMasterVolume, voldata)
            _Parent.FireNotification(NotificationData)
            Return 0
        End Function
    End Class
End Namespace