'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Imports System.Collections.Generic
Imports System.Text
Imports NAudio.CoreAudioApi.Interfaces
Imports System.Runtime.InteropServices

Namespace NAudio.CoreAudioApi
    Public Class AudioEndpointVolume
        Implements IDisposable
        Private _AudioEndPointVolume As IAudioEndpointVolume
        Private _Channels As AudioEndpointVolumeChannels
        Private _StepInformation As AudioEndpointVolumeStepInformation
        Private _VolumeRange As AudioEndpointVolumeVolumeRange
        Private _HardwareSupport As EEndpointHardwareSupport
        Private _CallBack As AudioEndpointVolumeCallback
        Public Event OnVolumeNotification As AudioEndpointVolumeNotificationDelegate

        Public ReadOnly Property VolumeRange() As AudioEndpointVolumeVolumeRange
            Get
                Return _VolumeRange
            End Get
        End Property

        Public ReadOnly Property HardwareSupport() As EEndpointHardwareSupport
            Get
                Return _HardwareSupport
            End Get
        End Property

        Public ReadOnly Property StepInformation() As AudioEndpointVolumeStepInformation
            Get
                Return _StepInformation
            End Get
        End Property

        Public ReadOnly Property Channels() As AudioEndpointVolumeChannels
            Get
                Return _Channels
            End Get
        End Property

        Public Property MasterVolumeLevel() As Single
            Get
                Dim result As Single
                Marshal.ThrowExceptionForHR(_AudioEndPointVolume.GetMasterVolumeLevel(result))
                Return result
            End Get
            Set(value As Single)
                Marshal.ThrowExceptionForHR(_AudioEndPointVolume.SetMasterVolumeLevel(value, Guid.Empty))
            End Set
        End Property

        Public Property MasterVolumeLevelScalar() As Single
            Get
                Dim result As Single
                Marshal.ThrowExceptionForHR(_AudioEndPointVolume.GetMasterVolumeLevelScalar(result))
                Return result
            End Get
            Set(value As Single)
                Marshal.ThrowExceptionForHR(_AudioEndPointVolume.SetMasterVolumeLevelScalar(value, Guid.Empty))
            End Set
        End Property

        Public Property Mute() As Boolean
            Get
                Dim result As Boolean
                Marshal.ThrowExceptionForHR(_AudioEndPointVolume.GetMute(result))
                Return result
            End Get
            Set(value As Boolean)
                Marshal.ThrowExceptionForHR(_AudioEndPointVolume.SetMute(value, Guid.Empty))
            End Set
        End Property

        Public Sub VolumeStepUp()
            Marshal.ThrowExceptionForHR(_AudioEndPointVolume.VolumeStepUp(Guid.Empty))
        End Sub

        Public Sub VolumeStepDown()
            Marshal.ThrowExceptionForHR(_AudioEndPointVolume.VolumeStepDown(Guid.Empty))
        End Sub

        Friend Sub New(realEndpointVolume As IAudioEndpointVolume)
            Dim HardwareSupp As UInteger
            _AudioEndPointVolume = realEndpointVolume
            _Channels = New AudioEndpointVolumeChannels(_AudioEndPointVolume)
            _StepInformation = New AudioEndpointVolumeStepInformation(_AudioEndPointVolume)
            Marshal.ThrowExceptionForHR(_AudioEndPointVolume.QueryHardwareSupport(HardwareSupp))
            _HardwareSupport = DirectCast(HardwareSupp, EEndpointHardwareSupport)
            _VolumeRange = New AudioEndpointVolumeVolumeRange(_AudioEndPointVolume)
            _CallBack = New AudioEndpointVolumeCallback(Me)
            Marshal.ThrowExceptionForHR(_AudioEndPointVolume.RegisterControlChangeNotify(_CallBack))
        End Sub
        Friend Sub FireNotification(NotificationData As AudioVolumeNotificationData)
            Dim del As AudioEndpointVolumeNotificationDelegate = OnVolumeNotification
            RaiseEvent del(NotificationData)
        End Sub
#Region "IDisposable Members"
        Public Sub Dispose() Implements IDisposable.Dispose
            If _CallBack IsNot Nothing Then
                Marshal.ThrowExceptionForHR(_AudioEndPointVolume.UnregisterControlChangeNotify(_CallBack))
                _CallBack = Nothing
            End If
            GC.SuppressFinalize(Me)
        End Sub

        Protected Overrides Sub Finalize()
            Try
                Dispose()
            Finally
                MyBase.Finalize()
            End Try
        End Sub
#End Region
    End Class
End Namespace