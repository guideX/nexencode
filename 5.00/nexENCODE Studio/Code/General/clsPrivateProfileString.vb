'nexENCODE Studio 5.0 Alpha 1.3
'October 6th, 2013
Option Explicit On
Option Strict On
Imports System.Text

Public Class clsPrivateProfileString
    Public Event ProcessError(ByVal lError As String, ByVal lSub As String)
    Private Declare Auto Function GetPrivateProfileString Lib "kernel32" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As StringBuilder, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    Private Declare Auto Function WritePrivateProfileString Lib "kernel32" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Boolean

    Public Function ReadINI(ByVal lFile As String, ByVal lSection As String, ByVal lKey As String, Optional ByVal lDefault As String = "") As String
        Try
            Dim i As Integer, msg As StringBuilder, msg2 As String
            msg = New StringBuilder(500)
            i = GetPrivateProfileString(lSection, lKey, "", msg, msg.Capacity, lFile)
            msg2 = msg.ToString()
            If i = 0 Then
                ReadINI = lDefault
            Else
                ReadINI = Trim(msg2)
            End If
        Catch Ex As Exception
            RaiseEvent ProcessError(Ex.Message, "Public Function ReadINI(ByVal lFile As String, ByVal lSection As String, ByVal lKey As String, Optional ByVal lDefault As String = "") As String")
            Return lDefault
        End Try
    End Function

    Public Sub WriteINI(ByVal lFile As String, ByVal lSection As String, ByVal lKey As String, ByVal lValue As String)
        Try
            WritePrivateProfileString(lSection, lKey, lValue, lFile)
        Catch Ex As Exception
            RaiseEvent ProcessError(Ex.Message, "Public Sub WriteINI(ByVal lFile As String, ByVal lSection As String, ByVal lKey As String, ByVal lValue As String)")
        End Try
    End Sub
End Class

Public Class clsIniFiles
    Public Event ProcessError(ByVal lError As String, ByVal lSub As String)
    Private Const lSettingsINI As String = "settings.ini"
    Private Const lWindowPosINI As String = "windowpos.ini"
    Private Const lSkinsINI As String = "skins.ini"

    Public ReadOnly Property SettingsINI() As String
        Get
            Try
                Return Application.StartupPath & "\" & lSettingsINI
            Catch ex As Exception
                RaiseEvent ProcessError(ex.Message, "Public ReadOnly Property SettingsINI() As String")
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property WindowPosINI() As String
        Get
            Try
                Return Application.StartupPath & "\data\config\" & lWindowPosINI
            Catch ex As Exception
                RaiseEvent ProcessError(ex.Message, "Public ReadOnly Property WindowPosINI() As String")
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property SkinsINI() As String
        Get
            Try
                Return Application.StartupPath & "\data\config\" & lSkinsINI
            Catch ex As Exception
                RaiseEvent ProcessError(ex.Message, "Public ReadOnly Property SkinsINI() As String")
                Return Nothing
            End Try
        End Get
    End Property
End Class
