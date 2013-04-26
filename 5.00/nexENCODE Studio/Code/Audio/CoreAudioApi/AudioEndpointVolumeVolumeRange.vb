'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Imports System.Collections.Generic
Imports System.Text
Imports NAudio.CoreAudioApi.Interfaces
Imports System.Runtime.InteropServices

Namespace NAudio.CoreAudioApi
    Public Class AudioEndpointVolumeVolumeRange
        Private _VolumeMindB As Single
        Private _VolumeMaxdB As Single
        Private _VolumeIncrementdB As Single

        Friend Sub New(parent As IAudioEndpointVolume)
            Marshal.ThrowExceptionForHR(parent.GetVolumeRange(_VolumeMindB, _VolumeMaxdB, _VolumeIncrementdB))
        End Sub

        Public ReadOnly Property MinDecibels() As Single
            Get
                Return _VolumeMindB
            End Get
        End Property

        Public ReadOnly Property MaxDecibels() As Single
            Get
                Return _VolumeMaxdB
            End Get
        End Property

        Public ReadOnly Property IncrementDecibels() As Single
            Get
                Return _VolumeIncrementdB
            End Get
        End Property
    End Class
End Namespace
