'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Imports System.Collections.Generic
Imports System.Text
Imports NAudio.CoreAudioApi.Interfaces
Imports System.Runtime.InteropServices

Namespace NAudio.CoreAudioApi
    Public Class AudioEndpointVolumeChannels
        Private _AudioEndPointVolume As IAudioEndpointVolume
        Private _Channels As AudioEndpointVolumeChannel()

        Public ReadOnly Property Count() As Integer
            Get
                Dim result As Integer
                Marshal.ThrowExceptionForHR(_AudioEndPointVolume.GetChannelCount(result))
                Return result
            End Get
        End Property

        Default Public ReadOnly Property Item(index As Integer) As AudioEndpointVolumeChannel
            Get
                Return _Channels(index)
            End Get
        End Property

        Friend Sub New(parent As IAudioEndpointVolume)
            Dim ChannelCount As Integer
            _AudioEndPointVolume = parent
            ChannelCount = Count
            _Channels = New AudioEndpointVolumeChannel(ChannelCount - 1) {}
            For i As Integer = 0 To ChannelCount - 1
                _Channels(i) = New AudioEndpointVolumeChannel(_AudioEndPointVolume, i)
            Next
        End Sub
    End Class
End Namespace