﻿'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Imports System.Collections.Generic
Imports System.Text

Namespace NAudio.CoreAudioApi
    Public Enum DataFlow
        Render
        Capture
        All
    End Enum
End Namespace