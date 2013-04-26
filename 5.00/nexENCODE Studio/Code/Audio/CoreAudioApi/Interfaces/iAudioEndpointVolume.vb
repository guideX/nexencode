Imports System.Collections.Generic
Imports System.Text
Imports System.Runtime.InteropServices

Namespace NAudio.CoreAudioApi.Interfaces
    <Guid("5CDF2C82-841E-4546-9722-0CF74078229A"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)> Friend Interface IAudioEndpointVolume
        Function RegisterControlChangeNotify(pNotify As IAudioEndpointVolumeCallback) As Integer
        Function UnregisterControlChangeNotify(pNotify As IAudioEndpointVolumeCallback) As Integer
        Function GetChannelCount(ByRef pnChannelCount As Integer) As Integer
        Function SetMasterVolumeLevel(fLevelDB As Single, pguidEventContext As Guid) As Integer
        Function SetMasterVolumeLevelScalar(fLevel As Single, pguidEventContext As Guid) As Integer
        Function GetMasterVolumeLevel(ByRef pfLevelDB As Single) As Integer
        Function GetMasterVolumeLevelScalar(ByRef pfLevel As Single) As Integer
        Function SetChannelVolumeLevel(nChannel As UInteger, fLevelDB As Single, pguidEventContext As Guid) As Integer
        Function SetChannelVolumeLevelScalar(nChannel As UInteger, fLevel As Single, pguidEventContext As Guid) As Integer
        Function GetChannelVolumeLevel(nChannel As UInteger, ByRef pfLevelDB As Single) As Integer
        Function GetChannelVolumeLevelScalar(nChannel As UInteger, ByRef pfLevel As Single) As Integer
        Function SetMute(<MarshalAs(UnmanagedType.Bool)> bMute As [Boolean], pguidEventContext As Guid) As Integer
        Function GetMute(ByRef pbMute As Boolean) As Integer
        Function GetVolumeStepInfo(ByRef pnStep As UInteger, ByRef pnStepCount As UInteger) As Integer
        Function VolumeStepUp(pguidEventContext As Guid) As Integer
        Function VolumeStepDown(pguidEventContext As Guid) As Integer
        Function QueryHardwareSupport(ByRef pdwHardwareSupportMask As UInteger) As Integer
        Function GetVolumeRange(ByRef pflVolumeMindB As Single, ByRef pflVolumeMaxdB As Single, ByRef pflVolumeIncrementdB As Single) As Integer
    End Interface
End Namespace