'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Imports System.Collections.Generic
Imports System.Text

Namespace NAudio.CoreAudioApi
    Public Delegate Sub AudioEndpointVolumeNotificationDelegate(data As AudioVolumeNotificationData)
End Namespace