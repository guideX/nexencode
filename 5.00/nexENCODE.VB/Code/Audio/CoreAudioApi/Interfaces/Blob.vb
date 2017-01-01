'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Imports System.Collections.Generic
Imports System.Text
Imports System.Runtime.InteropServices

Namespace NAudio.CoreAudioApi.Interfaces
    Friend Structure Blob
        Public Length As Integer
        Public Data As IntPtr

        Private Sub FixCS0649()
            Length = 0
            Data = IntPtr.Zero
        End Sub
    End Structure
End Namespace