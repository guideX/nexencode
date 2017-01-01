'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Imports System.Collections.Generic
Imports System.Text

Namespace NAudio.CoreAudioApi.Interfaces
    Friend Structure AudioVolumeNotificationDataStruct
        Public guidEventContext As Guid
        Public bMuted As Boolean
        Public fMasterVolume As Single
        Public nChannels As UInteger
        Public ChannelVolume As Single

        Private Sub FixCS0649()
            guidEventContext = Guid.Empty
            bMuted = False
            fMasterVolume = 0
            nChannels = 0
            ChannelVolume = 0
        End Sub
    End Structure
End Namespace