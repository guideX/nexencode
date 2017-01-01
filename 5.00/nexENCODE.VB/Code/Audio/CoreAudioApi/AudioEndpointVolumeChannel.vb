'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Imports System.Collections.Generic
Imports System.Text
Imports NAudio.CoreAudioApi.Interfaces
Imports System.Runtime.InteropServices

Namespace NAudio.CoreAudioApi
    Public Class AudioEndpointVolumeChannel
        Private _Channel As UInteger
        Private _AudioEndpointVolume As IAudioEndpointVolume

        Friend Sub New(parent As IAudioEndpointVolume, channel As Integer)
            _Channel = CUInt(channel)
            _AudioEndpointVolume = parent
        End Sub

        Public Property VolumeLevel() As Single
            Get
                Dim result As Single
                Marshal.ThrowExceptionForHR(_AudioEndpointVolume.GetChannelVolumeLevel(_Channel, result))
                Return result
            End Get
            Set(value As Single)
                Marshal.ThrowExceptionForHR(_AudioEndpointVolume.SetChannelVolumeLevel(_Channel, value, Guid.Empty))
            End Set
        End Property

        Public Property VolumeLevelScalar() As Single
            Get
                Dim result As Single
                Marshal.ThrowExceptionForHR(_AudioEndpointVolume.GetChannelVolumeLevelScalar(_Channel, result))
                Return result
            End Get
            Set(value As Single)
                Marshal.ThrowExceptionForHR(_AudioEndpointVolume.SetChannelVolumeLevelScalar(_Channel, value, Guid.Empty))
            End Set
        End Property
    End Class
End Namespace