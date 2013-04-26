'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Imports System.Collections.Generic
Imports System.Text
Imports System.Runtime.InteropServices
Imports NAudio.CoreAudioApi.Interfaces

Namespace NAudio.CoreAudioApi
    Public Class MMDeviceEnumerator
        Private _realEnumerator As IMMDeviceEnumerator

        Public Function EnumerateAudioEndPoints(dataFlow As DataFlow, dwStateMask As DeviceState) As MMDeviceCollection
            Dim result As IMMDeviceCollection
            Marshal.ThrowExceptionForHR(_realEnumerator.EnumAudioEndpoints(dataFlow, dwStateMask, result))
            Return New MMDeviceCollection(result)
        End Function

        Public Function GetDefaultAudioEndpoint(dataFlow As DataFlow, role As Role) As MMDevice
            Dim _Device As IMMDevice = Nothing
            Marshal.ThrowExceptionForHR(DirectCast(_realEnumerator, IMMDeviceEnumerator).GetDefaultAudioEndpoint(dataFlow, role, _Device))
            Return New MMDevice(_Device)
        End Function

        Public Function GetDevice(ID As String) As MMDevice
            Dim _Device As IMMDevice = Nothing
            Marshal.ThrowExceptionForHR(DirectCast(_realEnumerator, IMMDeviceEnumerator).GetDevice(ID, _Device))
            Return New MMDevice(_Device)
        End Function

        Public Sub New()
            If System.Environment.OSVersion.Version.Major < 6 Then
                Throw New NotSupportedException("This functionality is only supported on Windows Vista or newer.")
            End If
            _realEnumerator = TryCast(New MMDeviceEnumeratorComObject(), IMMDeviceEnumerator)
        End Sub
    End Class
End Namespace