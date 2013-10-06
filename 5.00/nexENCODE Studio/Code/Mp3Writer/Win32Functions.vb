'nexENCODE Studio 5.0 Alpha 1.3
'October 6th, 2013
Option Explicit On
Option Strict On
Imports System.Runtime.InteropServices

Namespace Yeti.Sys
    Public Enum BeepType
        SimpleBeep = -1
        SystemAsterisk = &H40
        SystemExclamation = &H30
        SystemHand = &H10
        SystemQuestion = &H20
        SystemDefault = 0
    End Enum

    Public NotInheritable Class Win32
        <DllImport("User32.dll", SetLastError:=True)> _
        Public Shared Function MessageBeep(Type As BeepType) As Boolean
        End Function
    End Class
End Namespace
