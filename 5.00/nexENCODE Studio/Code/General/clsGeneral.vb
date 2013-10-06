'nexENCODE Studio 5.0 Alpha 1.3
'October 6th, 2013
Option Explicit On
Option Strict On
Imports System.Runtime.InteropServices

Public Class clsGeneral
    Public Enum eBeepType
        bSimpleBeep = -1
        bSystemAsterisk = &H40
        bSystemExclamation = &H30
        bSystemHand = &H10
        bSystemQuestion = &H20
        bSystemDefault = 0
    End Enum

    Public NotInheritable Class Win32
        <DllImport("User32.dll", SetLastError:=True)> Public Shared Function MessageBeep(Type As eBeepType) As Boolean
        End Function
    End Class
End Class
