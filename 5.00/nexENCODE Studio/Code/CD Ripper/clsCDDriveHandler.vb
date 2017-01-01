'nexENCODE Studio 5.0 Alpha 1.3
'October 6th, 2013
Option Explicit On
Option Strict On
Imports nexENCODE.WaveLib
Imports nexENCODE.Media.Mp3
Imports System.IO
Namespace nexENCODE.CDRipper
    Public Class clsCDDriveHandler
#Region "DECLARATIONS"
        Public Event ProcessError(lError As String, lSub As String)
        Public Event CDReadStarted()
        Public Event CDReadCanceled()
        Public Event CDReadCompleted(_SecondsEllapsed As Integer)
        Public Event CDReadProgress(_Percent As Integer, _SecondsRemaining As Integer, _BytesRead As Long, _BytesRemaining As Long)
        Public Event DisplayLabel(lData As String)
        Public Event CDInserted()
        Public Event CDRemoved()
        Private Delegate Sub EmptyDelegate()
        Private Delegate Sub StringDelegate(lString As String)
        Private Delegate Sub IntegerDelegate(lInteger As Integer)
        Private Delegate Sub CDDriveProgressDelegate(_Percent As Integer, _SecondsRemaining As Integer, _BytesRead As Long, _BytesRemaining As Long)
        Private WithEvents lCDDrive As New nexENCODE.CDRipper.clsCDDrive
        Private WithEvents lFiles As New clsFiles
        Private WithEvents lMp3WriterConfig As Mp3WriterConfig
        Private WithEvents lWaveStreamWriter As clsWaveWriter
        Private WithEvents lMp3Writer As Media.Mp3.Mp3Writer
        Private lInvokeForm As Form
        Private lCurrentTrack As Integer, lCurrentDrive As Char, lCurrentFile As String, lCurrentMp3File As String
        Private lRipSingleTrackProc As Threading.Thread
        Private lRipToWaveFile As Boolean, lRipToMp3File As Boolean
        Private lStartTime As New DateTime
        Private lItemCount As Long
        Private lItemsSoFar As Long
        Private lItemsRemaining As Long
        Private lTimeSoFar As TimeSpan
        Private lPercent As Integer
        Private lDisplayInterval As Long
#End Region
#Region "EVENTHANDLERS"
        Private Sub CDDataRead_EventHandler(sender As Object, e As nexENCODE.CDRipper.CDDriveEvents.DataReadEventArgs)
            Try
                lItemsSoFar = lItemsSoFar + e.DataSize
                If lWaveStreamWriter IsNot Nothing Then
                    lWaveStreamWriter.Write(e.Data, CInt(e.DataSize))
                End If
                If lMp3Writer IsNot Nothing Then
                    lMp3Writer.Write(e.Data, 0, CInt(e.DataSize))
                End If
            Catch ex As Exception
                RaiseEvent ProcessError(ex.Message, "Sub CDDataRead_EventHandler(sender As Object, e As Ripper.DataReadEventArgs)")
            End Try
        End Sub

        Private Sub CDReadProgress_EventHandler(sender As Object, e As nexENCODE.CDRipper.CDDriveEvents.ReadProgressEventArgs)
            Try
                If lItemCount = 0 Then
                    lItemCount = CLng(e.Bytes2Read)
                    lStartTime = Now
                    lDisplayInterval = 0
                    lItemsSoFar = 0
                    lItemsRemaining = CLng(e.Bytes2Read)
                End If
                lDisplayInterval = lDisplayInterval + 1
                If lDisplayInterval = 10 Then
                    lDisplayInterval = 0
                    lTimeSoFar = lStartTime - Now
                    lItemsRemaining = lItemCount - lItemsSoFar
                    CDReadProgressSub(CInt(e.BytesRead / e.Bytes2Read * 100), CInt(lTimeSoFar.Seconds * lItemsRemaining / lItemsSoFar) * -1, lItemsSoFar, lItemsRemaining)
                End If
            Catch ex As Exception
                RaiseEvent ProcessError(ex.Message, "Private Sub CDReadProgress_EventHandler(sender As Object, e As Ripper.CdReadProgressEventHandler)")
            End Try
        End Sub

        Private Sub lCDDrive_CDInserted(sender As Object, e As System.EventArgs) Handles lCDDrive.CDInserted
            Try
                RaiseEvent CDInserted()
            Catch ex As Exception
                RaiseEvent ProcessError(ex.Message, "Private Sub lCDDrive_CDInserted(sender As Object, e As System.EventArgs) Handles lCDDrive.CDInserted")
            End Try
        End Sub

        Private Sub lCDDrive_CDRemoved(sender As Object, e As System.EventArgs) Handles lCDDrive.CDRemoved
            Try
                RaiseEvent CDRemoved()
            Catch ex As Exception
                RaiseEvent ProcessError(ex.Message, "Private Sub lCDDrive_CDRemoved(sender As Object, e As System.EventArgs) Handles lCDDrive.CDRemoved")
            End Try
        End Sub

        Public Sub New(_InvokeForm As Form, _Drive As Char)
            Try
                lInvokeForm = _InvokeForm
                lCurrentDrive = _Drive
            Catch ex As Exception
                RaiseEvent ProcessError(ex.Message, "Public Sub New(_InvokeForm As Form)")
            End Try
        End Sub
#Region "ERROR_HANDLERS"
        Private Sub lFiles_ProcessError(lError As String, lSub As String) Handles lFiles.ProcessError
            Try
                RaiseEvent ProcessError(lError, lSub)
            Catch ex As Exception
                RaiseEvent ProcessError(ex.Message, "Private Sub lFiles_ProcessError(lError As String, lSub As String) Handles lFiles.ProcessError")
            End Try
        End Sub

        Private Sub lCDDrive_ProcessError(lError As String, lSub As String) Handles lCDDrive.ProcessError
            Try
                RaiseEvent ProcessError(lError, lSub)
            Catch ex As Exception
                RaiseEvent ProcessError(ex.Message, "Private Sub lCDDrive_ProcessError(lError As String, lSub As String) Handles lCDDrive.ProcessError")
            End Try
        End Sub
#End Region
#End Region
#Region "DELEGATE_SUBS"
        Private Sub CDReadCanceledProc()
            Try
                RaiseEvent CDReadCanceled()
            Catch ex As Exception
                RaiseEvent ProcessError(ex.Message, "Private Sub CDReadCanceledProc()")
            End Try
        End Sub

        Private Sub CDReadCanceledSub()
            Try
                Dim _CDReadCanceledProc As New EmptyDelegate(AddressOf CDReadCanceledProc)
                lInvokeForm.Invoke(_CDReadCanceledProc)
            Catch ex As Exception
                RaiseEvent ProcessError(ex.Message, "Private Sub CDReadCanceledSub()")
            End Try
        End Sub

        Private Sub CDReadProgressProc(_Percent As Integer, _SecondsRemaining As Integer, _BytesRead As Long, _BytesRemaining As Long)
            Try
                RaiseEvent CDReadProgress(_Percent, _SecondsRemaining, _BytesRead, _BytesRemaining)
            Catch ex As Exception
                RaiseEvent ProcessError(ex.Message, "Private Sub CDReadProgressProc(_Percent As Integer, _SecondsRemaining As Integer, _BytesRead As Long, _BytesRemaining As Long)")
            End Try
        End Sub

        Private Sub CDReadProgressSub(_Percent As Integer, _SecondsRemaining As Integer, _BytesRead As Long, _BytesRemaining As Long)
            Try
                Dim _CDReadProgressProc As New CDDriveProgressDelegate(AddressOf CDReadProgressProc)
                lInvokeForm.Invoke(_CDReadProgressProc, _Percent, _SecondsRemaining, _BytesRead, _BytesRemaining)
            Catch ex As Exception
                RaiseEvent ProcessError(ex.Message, "Private Sub CDReadProgressSub(_Percent As Integer, _SecondsRemaining As Integer, _BytesRead As Long, _BytesRemaining As Long)")
            End Try
        End Sub

        Private Sub CDReadCompleteProc(_SecondsEllapsed As Integer)
            Try
                RaiseEvent CDReadCompleted(_SecondsEllapsed)
            Catch ex As Exception
                RaiseEvent ProcessError(ex.Message, "Private Sub CDReadCompleteProc()")
            End Try
        End Sub

        Private Sub CDReadCompleteSub(_SecondsEllapsed As Integer)
            Try
                Dim _CDReadCompleteProc As New IntegerDelegate(AddressOf CDReadCompleteProc)
                lInvokeForm.Invoke(_CDReadCompleteProc, _SecondsEllapsed)
            Catch ex As Exception
                RaiseEvent ProcessError(ex.Message, "Private Sub CDReadCompleteSub(_SecondsEllapsed As Integer)")
            End Try
        End Sub
#End Region
#Region "SUBS"
        Public Sub RipCurrentTrack(_Track As Integer, _File As String, _Mp3File As String, Optional _RipToWaveFile As Boolean = False, Optional _RipToMp3File As Boolean = True)
            Try
                lRipToMp3File = _RipToMp3File
                lRipToWaveFile = _RipToWaveFile
                lCurrentTrack = _Track
                lCurrentFile = _File
                lCurrentMp3File = _Mp3File
                lRipSingleTrackProc = New Threading.Thread(AddressOf RipCurrentTrackProc)
                lRipSingleTrackProc.Start()
            Catch ex As Exception
                RaiseEvent ProcessError(ex.Message, "Public Sub RipSingleTrack(lForm As Form, lDrive As Char, lTrack As Integer, lFile As String)")
            Finally
            End Try
        End Sub

        Private Sub RipCurrentTrackProc()
            Try
                Dim mBox As MsgBoxResult, b As Boolean = True
                lDisplayInterval = 0
                If lCurrentTrack <> 0 And Len(lCurrentDrive) <> 0 Then
                    If lCDDrive.Open(lCurrentDrive) = True Then
                        If lCDDrive.IsOpened() = True Then
                            If lCDDrive.Refresh() = True Then
                                If lCDDrive.IsAudioTrack(lCurrentTrack) = True Then
                                    If lRipToWaveFile = True Then
                                        If Len(lCurrentFile) = 0 Then
                                            lWaveStreamWriter = Nothing
                                            lRipToWaveFile = False
                                            b = False
                                        End If
                                        If lFiles.DoesFileExist(lCurrentFile) = True Then
                                            mBox = MsgBox("File already exists, overright existing file?", CType(MsgBoxStyle.Question + MsgBoxStyle.YesNo, MsgBoxStyle), "Destination File Conflict")
                                            If mBox = MsgBoxResult.Yes Then
                                                Kill(lCurrentFile)
                                                If lFiles.DoesFileExist(lCurrentFile) = True Then
                                                    b = False
                                                    lWaveStreamWriter = Nothing
                                                Else
                                                    b = True
                                                End If
                                            Else
                                                lWaveStreamWriter = Nothing
                                                lRipToWaveFile = False
                                                b = False
                                            End If
                                        End If
                                        If b = True Then
                                            lWaveStreamWriter = New clsWaveWriter(lCurrentFile, 44100, 2, 16)
                                        End If
                                    Else
                                        lWaveStreamWriter = Nothing
                                    End If
                                    If lRipToMp3File = True Then
                                        b = True
                                        If Len(lCurrentMp3File) = 0 Then
                                            lMp3Writer = Nothing
                                            lRipToMp3File = False
                                            b = False
                                        End If
                                        If lFiles.DoesFileExist(lCurrentMp3File) = True Then
                                            mBox = MsgBox("File already exists, overright existing file?", CType(MsgBoxStyle.Question + MsgBoxStyle.YesNo, MsgBoxStyle), "Destination File Conflict")
                                            If mBox = MsgBoxResult.Yes Then
                                                Kill(lCurrentMp3File)
                                                If lFiles.DoesFileExist(lCurrentMp3File) = True Then
                                                    b = False
                                                    lMp3Writer = Nothing
                                                Else
                                                    b = True
                                                End If
                                            Else
                                                lMp3Writer = Nothing
                                                lRipToMp3File = False
                                                b = False
                                            End If
                                        End If
                                        If b = True Then
                                            lMp3WriterConfig = New Mp3WriterConfig()
                                            lMp3Writer = New Mp3Writer(New FileStream(lCurrentMp3File, FileMode.Create), lMp3WriterConfig)
                                        End If
                                    Else
                                        lMp3Writer = Nothing
                                    End If
                                    If lRipToMp3File = True Or lRipToWaveFile = True Then
                                        lItemCount = 0
                                        lItemsSoFar = 0
                                        Dim n As Integer = lCDDrive.ReadTrack(lCurrentTrack, AddressOf CDDataRead_EventHandler, AddressOf CDReadProgress_EventHandler)
                                    Else
                                        CDReadCanceledSub()
                                        Exit Sub
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                CDReadCompleteSub(lTimeSoFar.Seconds * -1)
            Catch ex As Exception
                RaiseEvent ProcessError(ex.Message, "Private Sub RipSingleTrackProc()")
            Finally
                lDisplayInterval = 0
                lItemCount = 0
                lItemsSoFar = 0
                lStartTime = Nothing
                lTimeSoFar = Nothing
                lItemsRemaining = 0
                lCDDrive.Close()
                lCDDrive.UnLockCD()
                If lRipToMp3File = True Then lMp3Writer.Close()
                If lRipToWaveFile = True Then
                    If (lWaveStreamWriter IsNot Nothing) Then
                        lWaveStreamWriter.CloseWaveFile()
                        lWaveStreamWriter.Close()
                    End If
                End If
            End Try
        End Sub

        Private Shared Function InlineAssignHelper(Of T)(ByRef _Target As T, _Value As T) As T
            Try
                _Target = _Value
                Return _Value
            Catch ex As Exception
                'RaiseEvent ProcessError(ex.Message, "Private Shared Function InlineAssignHelper(Of T)(ByRef _Target As T, _Value As T) As T")
                Return Nothing
            End Try
        End Function
#End Region
    End Class
End Namespace