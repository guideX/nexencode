'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Imports System.Collections.Generic
Imports System.Text
Imports NAudio.CoreAudioApi.Interfaces
Imports System.Runtime.InteropServices

Namespace NAudio.CoreAudioApi
    Public Class MMDevice
#Region "Variables"
        Private deviceInterface As IMMDevice
        Private _PropertyStore As PropertyStore
        Private _AudioMeterInformation As AudioMeterInformation
        Private _AudioEndpointVolume As AudioEndpointVolume
#End Region
#Region "Guids"
        Private Shared IID_IAudioMeterInformation As New Guid("C02216F6-8C67-4B5B-9D00-D008E73E0064")
        Private Shared IID_IAudioEndpointVolume As New Guid("5CDF2C82-841E-4546-9722-0CF74078229A")
        Private Shared IID_IAudioClient As New Guid("1CB9AD4C-DBFA-4c32-B178-C2F568A703B2")
#End Region
#Region "Init"
        Private Sub GetPropertyInformation()
            Dim propstore As IPropertyStore
            Marshal.ThrowExceptionForHR(deviceInterface.OpenPropertyStore(StorageAccessMode.Read, propstore))
            _PropertyStore = New PropertyStore(propstore)
        End Sub

        Private Function GetAudioClient() As AudioClient
            Dim result As Object
            Marshal.ThrowExceptionForHR(deviceInterface.Activate(IID_IAudioClient, ClsCtx.ALL, IntPtr.Zero, result))
            Return New AudioClient(TryCast(result, IAudioClient))
        End Function

        Private Sub GetAudioMeterInformation()
            Dim result As Object
            Marshal.ThrowExceptionForHR(deviceInterface.Activate(IID_IAudioMeterInformation, ClsCtx.ALL, IntPtr.Zero, result))
            _AudioMeterInformation = New AudioMeterInformation(TryCast(result, IAudioMeterInformation))
        End Sub

        Private Sub GetAudioEndpointVolume()
            Dim result As Object
            Marshal.ThrowExceptionForHR(deviceInterface.Activate(IID_IAudioEndpointVolume, ClsCtx.ALL, IntPtr.Zero, result))
            _AudioEndpointVolume = New AudioEndpointVolume(TryCast(result, IAudioEndpointVolume))
        End Sub
#End Region
#Region "Properties"
        Public ReadOnly Property AudioClient() As AudioClient
            Get
                Return GetAudioClient()
            End Get
        End Property

        Public ReadOnly Property AudioMeterInformation() As AudioMeterInformation
            Get
                If _AudioMeterInformation Is Nothing Then
                    GetAudioMeterInformation()
                End If

                Return _AudioMeterInformation
            End Get
        End Property

        Public ReadOnly Property AudioEndpointVolume() As AudioEndpointVolume
            Get
                If _AudioEndpointVolume Is Nothing Then
                    GetAudioEndpointVolume()
                End If

                Return _AudioEndpointVolume
            End Get
        End Property

        Public ReadOnly Property Properties() As PropertyStore
            Get
                If _PropertyStore Is Nothing Then
                    GetPropertyInformation()
                End If
                Return _PropertyStore
            End Get
        End Property

        Public ReadOnly Property FriendlyName() As String
            Get
                If _PropertyStore Is Nothing Then
                    GetPropertyInformation()
                End If
                If _PropertyStore.Contains(PropertyKeys.PKEY_DeviceInterface_FriendlyName) Then
                    Return DirectCast(_PropertyStore(PropertyKeys.PKEY_DeviceInterface_FriendlyName).Value, String)
                Else
                    Return "Unknown"
                End If
            End Get
        End Property

        Public ReadOnly Property DeviceFriendlyName() As String
            Get
                If _PropertyStore Is Nothing Then
                    GetPropertyInformation()
                End If
                If _PropertyStore.Contains(PropertyKeys.PKEY_Device_FriendlyName) Then
                    Return DirectCast(_PropertyStore(PropertyKeys.PKEY_Device_FriendlyName).Value, String)
                Else
                    Return "Unknown"
                End If
            End Get
        End Property

        Public ReadOnly Property ID() As String
            Get
                Dim Result As String
                Marshal.ThrowExceptionForHR(deviceInterface.GetId(Result))
                Return Result
            End Get
        End Property

        Public ReadOnly Property DataFlow() As DataFlow
            Get
                Dim Result As DataFlow
                Dim ep As IMMEndpoint = TryCast(deviceInterface, IMMEndpoint)
                ep.GetDataFlow(Result)
                Return Result
            End Get
        End Property

        Public ReadOnly Property State() As DeviceState
            Get
                Dim Result As DeviceState
                Marshal.ThrowExceptionForHR(deviceInterface.GetState(Result))
                Return Result
            End Get
        End Property
#End Region
#Region "Constructor"
        Friend Sub New(realDevice As IMMDevice)
            deviceInterface = realDevice
        End Sub
#End Region
        Public Overrides Function ToString() As String
            Return FriendlyName
        End Function
    End Class
End Namespace
