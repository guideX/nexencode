'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Imports System.Collections.Generic
Imports System.Text
Imports NAudio.CoreAudioApi.Interfaces
Imports System.Runtime.InteropServices

Namespace NAudio.CoreAudioApi
    Public Class AudioMeterInformationChannels
        Private _AudioMeterInformation As IAudioMeterInformation

        Public ReadOnly Property Count() As Integer
            Get
                Dim result As Integer
                Marshal.ThrowExceptionForHR(_AudioMeterInformation.GetMeteringChannelCount(result))
                Return result
            End Get
        End Property

        Default Public ReadOnly Property Item(index As Integer) As Single
            Get
                Dim peakValues As Single() = New Single(Count - 1) {}
                Dim Params As GCHandle = GCHandle.Alloc(peakValues, GCHandleType.Pinned)
                Marshal.ThrowExceptionForHR(_AudioMeterInformation.GetChannelsPeakValues(peakValues.Length, Params.AddrOfPinnedObject()))
                Params.Free()
                Return peakValues(index)
            End Get
        End Property

        Friend Sub New(parent As IAudioMeterInformation)
            _AudioMeterInformation = parent
        End Sub
    End Class
End Namespace