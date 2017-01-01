'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Imports System.Collections.Generic
Imports System.Text

Namespace NAudio.CoreAudioApi
    <Flags()> Public Enum AudioClientBufferFlags
        None
        DataDiscontinuity = &H1
        Silent = &H2
        TimestampError = &H4
    End Enum
End Namespace