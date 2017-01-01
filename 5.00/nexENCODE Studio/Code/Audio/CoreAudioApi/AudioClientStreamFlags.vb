'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Imports System.Collections.Generic
Imports System.Text

Namespace NAudio.CoreAudioApi
    <Flags()> Public Enum AudioClientStreamFlags
        None
        CrossProcess = &H10000
        Loopback = &H20000
        EventCallback = &H40000
        NoPersist = &H80000
    End Enum
End Namespace
