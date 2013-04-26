'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict Off
Imports nexENCODE.WaveLib
Imports nexENCODE.Media.Mp3
Imports System.IO

Public Class clsMp3WriterHandler
#Region "DECLARATIONS"
    Public Event ProcessError(lError As String, lSub As String)
    Public Event EncodeStarted()
    Public Event EncodeComplete()
    Public Event EncodeCanceled()
    Public Event EncodeProgress(lProgress As Integer)
    Private Delegate Sub EmptyDelegate()
    Private Delegate Sub IntegerDelegate(_Value As Integer)
    Private WithEvents lFiles As New clsFiles
    Private WithEvents lMp3Writer As Media.Mp3.Mp3Writer
    Private WithEvents lWaveStream As WaveStream
    'Private WithEvents lWaveReader As clsWaveReader
    'Private WithEvents lWaveWriter As clsWaveWriter
    Private WithEvents lMp3WriterConfig As Mp3WriterConfig
    Private lInvokeForm As Form
#End Region
#Region "SUBS"
    Private Shared Function InlineAssignHelper(Of T)(ByRef _Target As T, _Value As T) As T
        Try
            _Target = _Value
            Return _Value
        Catch ex As Exception
            'RaiseEvent ProcessError(ex.Message, "Private Shared Function InlineAssignHelper(Of T)(ByRef _Target As T, _Value As T) As T")
            Return Nothing
        End Try
    End Function

    Private Sub EncodeStartedProc()
        Try
            RaiseEvent EncodeStarted()
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Private Sub EncodeStartedProc()")
        End Try
    End Sub

    Private Sub EncodeStartedSub()
        Try
            Dim _EncodeStartedProc As New EmptyDelegate(AddressOf EncodeStartedProc)
            lInvokeForm.Invoke(_EncodeStartedProc)
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Private Sub EncodeStartedSub()")
        End Try
    End Sub

    Private Sub EncodeProgressProc(_Value As Integer)
        Try
            RaiseEvent EncodeProgress(_Value)
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Private Sub EncodeProgressProc(lValue As Integer)")
        End Try
    End Sub

    Private Sub EncodeProgressSub(_Value As Integer)
        Try
            Dim _EncodeProgressProc As New IntegerDelegate(AddressOf EncodeProgressProc)
            lInvokeForm.Invoke(_EncodeProgressProc, _Value)
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Private Sub EncodeProgressSub()")
        End Try
    End Sub

    Private Sub EncodeCompleteSub()
        Try
            RaiseEvent EncodeComplete()
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Private Sub EncodeCompleteSub()")
        End Try
    End Sub

    Private Sub EncodeCompleteProc()
        Try
            Dim _EncodeCompleteProc As New EmptyDelegate(AddressOf EncodeCompleteProc)
            lInvokeForm.Invoke(_EncodeCompleteProc)
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Private Sub EncodeCompleteProc()")
        End Try
    End Sub

    Private Sub EncodeCanceledSub()
        Try
            RaiseEvent EncodeCanceled()
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Private Sub EncodeCanceledSub()")
        End Try
    End Sub

    Private Sub EncodeCanceledProc()
        Try
            Dim _EncodeCanceledProc As New EmptyDelegate(AddressOf EncodeCanceledProc)
            lInvokeForm.Invoke(_EncodeCanceledProc)
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Private Sub EncodeCanceledProc()")
        End Try
    End Sub

    Public Sub New(_InvokeForm As Form)
        Try
            lInvokeForm = _InvokeForm
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Sub New(_InvokeForm As Form)")
        End Try
    End Sub

    'Public Sub ConvertMp3ToWav(_Mp3File As String, _WaveFile As String)
    'If lFiles.DoesFileExist(_Mp3File) = True Then
    'If lFiles.DoesFileExist(_WaveFile) = True Then
    'Dim mbox As MsgBoxResult = MsgBox("The file '" & _WaveFile & "' already exists, do you wish to replace?", vbYesNoCancel + vbQuestion)
    'Select Case mbox
    'Case MsgBoxResult.Yes
    'Try
    'Kill(_WaveFile)
    'Catch ex As Exception
    'If lFiles.DoesFileExist(_WaveFile) = True Then
    'RaiseEvent ProcessError("Unable to delete the file '" & _WaveFile & "'.", "ConvertWavToMp3")
    'Exit Sub
    'End If
    'End Try
    'Case MsgBoxResult.No
    'Exit Sub
    'Case MsgBoxResult.Cancel
    'Exit Sub
    'End Select
    'End If
    'Try
    'Catch ex As Exception
    'RaiseEvent ProcessError(ex.Message, "Public Sub ConvertMp3ToWav(_Mp3File As String, _WaveFile As String)")
    'End Try
    'End If
    'End Sub

    Public Sub ConvertWavToMp3(_WaveFile As String, _Mp3File As String)
        If lFiles.DoesFileExist(_WaveFile) = True Then
            lWaveStream = New WaveStream(_WaveFile)
            Try
                If lFiles.DoesFileExist(_Mp3File) = True Then
                    Dim mbox As MsgBoxResult = MsgBox("The file '" & _Mp3File & "' already exists, do you wish to replace?", vbYesNoCancel + vbQuestion)
                    Select Case mbox
                        Case MsgBoxResult.Yes
                            Kill(_Mp3File)
                            If lFiles.DoesFileExist(_Mp3File) = True Then
                                RaiseEvent ProcessError("Unable to delete the file '" & _Mp3File & "'.", "ConvertWavToMp3")
                                Exit Sub
                            End If
                        Case MsgBoxResult.No
                            Exit Sub
                        Case MsgBoxResult.Cancel
                            Exit Sub
                    End Select
                End If
                lMp3WriterConfig = New Mp3WriterConfig()
                lMp3Writer = New Mp3Writer(New FileStream(_Mp3File, FileMode.Create), lMp3WriterConfig)
                Dim _Buff As Byte() = New Byte(lMp3Writer.OptimalBufferSize - 1) {}, _Read As Integer = 0, _Actual As Integer = 0, _Total As Long = lWaveStream.Length
                Cursor.Current = Cursors.WaitCursor
                EncodeStartedSub()
                While (InlineAssignHelper(_Read, lWaveStream.Read(_Buff, 0, _Buff.Length))) > 0
                    Application.DoEvents()
                    lMp3Writer.Write(_Buff, 0, _Read)
                    _Actual += _Read
                    EncodeProgressSub(CInt((CLng(_Actual) * 100) \ _Total))
                    Application.DoEvents()
                End While
            Catch ex As Exception
                RaiseEvent ProcessError(ex.Message, "Public Sub ConvertWavToMp3(_WaveFile As String, _Mp3File As String)")
            Finally
                Cursor.Current = Cursors.Default
                lMp3Writer.Close()
                lWaveStream.Close()
                EncodeCompleteSub()
            End Try
        End If
    End Sub
#End Region
#Region "HANDLERS"
    Private Sub lFiles_ProcessError(lError As String, lSub As String) Handles lFiles.ProcessError
        Try
            RaiseEvent ProcessError(lError, lSub)
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Private Sub lFiles_ProcessError(lError As String, lSub As String) Handles lFiles.ProcessError")
        End Try
    End Sub
#End Region
End Class