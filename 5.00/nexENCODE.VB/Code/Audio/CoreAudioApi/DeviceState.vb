'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Imports System.Collections.Generic
Imports System.Text

Namespace NAudio.CoreAudioApi
    <Flags()> Public Enum DeviceState
        Active = &H1
        Disabled = &H2
        NotPresent = &H4
        Unplugged = &H8
        All = &HF
    End Enum
End Namespace
