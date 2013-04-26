'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Public Class clsFiles
    Public Event ProcessError(ByVal lError As String, ByVal lSub As String)

    Public Function DoesDirectoryExist(lDirectory As String) As Boolean
        Try
            Return System.IO.Directory.Exists(lDirectory)
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Function DoesDirectoryExist(lDirectory As String) As Boolean")
            Return Nothing
        End Try
    End Function

    Public Function DoesFileExist(ByVal lFile As String) As Boolean
        Try
            Return System.IO.File.Exists(lFile)
        Catch Ex As Exception
            RaiseEvent ProcessError(Ex.Message, "Public Function DoesFileExist(ByVal lFile As String) As Boolean")
            Return False
        End Try
    End Function

    Public Function ReturnDirectoryFromFilePath(lFilePath As String) As String
        Try
            Dim splt() As String = Split(lFilePath, "\"), msg As String = "", n As Integer
            n = UBound(splt)
            For i As Integer = 0 To n
                If Len(splt(i)) <> 0 Then
                    If (i) <> n Then
                        If Len(msg) <> 0 Then
                            msg = msg & "\" & splt(i)
                        Else
                            msg = splt(i)
                        End If
                    Else
                        Exit For
                    End If
                End If
            Next i
            Return msg
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Function ReturnDirectoryFromFilePath(lFilePath As String) As String")
            Return Nothing
        End Try
    End Function
End Class