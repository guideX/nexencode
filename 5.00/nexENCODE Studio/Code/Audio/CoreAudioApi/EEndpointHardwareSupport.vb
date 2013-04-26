'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Imports System.Collections.Generic
Imports System.Text

Namespace NAudio.CoreAudioApi
    <Flags()> Public Enum EEndpointHardwareSupport
        Volume = &H1
        Mute = &H2
        Meter = &H4
    End Enum
End Namespace