'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Imports System.Collections.Generic
Imports System.Text
Imports NAudio.CoreAudioApi.Interfaces
Imports System.Runtime.InteropServices

Namespace NAudio.CoreAudioApi
    Public Class AudioEndpointVolumeStepInformation
        Private _Step As UInteger
        Private _StepCount As UInteger

        Friend Sub New(parent As IAudioEndpointVolume)
            Marshal.ThrowExceptionForHR(parent.GetVolumeStepInfo(_Step, _StepCount))
        End Sub

        Public ReadOnly Property [Step]() As UInteger
            Get
                Return _Step
            End Get
        End Property

        Public ReadOnly Property StepCount() As UInteger
            Get
                Return _StepCount
            End Get
        End Property
    End Class
End Namespace