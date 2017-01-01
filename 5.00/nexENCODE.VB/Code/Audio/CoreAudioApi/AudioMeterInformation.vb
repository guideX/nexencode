'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Imports System.Collections.Generic
Imports System.Text
Imports System.Runtime.InteropServices
Imports NAudio.CoreAudioApi.Interfaces

Namespace NAudio.CoreAudioApi
    Public Class AudioMeterInformation
        Private _AudioMeterInformation As IAudioMeterInformation
        Private _HardwareSupport As EEndpointHardwareSupport
        Private _Channels As AudioMeterInformationChannels

        Friend Sub New(realInterface As IAudioMeterInformation)
            Dim HardwareSupp As Integer
            _AudioMeterInformation = realInterface
            Marshal.ThrowExceptionForHR(_AudioMeterInformation.QueryHardwareSupport(HardwareSupp))
            _HardwareSupport = DirectCast(HardwareSupp, EEndpointHardwareSupport)

            _Channels = New AudioMeterInformationChannels(_AudioMeterInformation)
        End Sub

        Public ReadOnly Property PeakValues() As AudioMeterInformationChannels
            Get
                Return _Channels
            End Get
        End Property

        Public ReadOnly Property HardwareSupport() As EEndpointHardwareSupport
            Get
                Return _HardwareSupport
            End Get
        End Property

        Public ReadOnly Property MasterPeakValue() As Single
            Get
                Dim result As Single
                Marshal.ThrowExceptionForHR(_AudioMeterInformation.GetPeakValue(result))
                Return result
            End Get
        End Property
    End Class
End Namespace