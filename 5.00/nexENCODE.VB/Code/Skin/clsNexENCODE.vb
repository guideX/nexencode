'nexENCODE Studio 5.0 Alpha 1.3
'October 6th, 2013
Option Explicit On
Option Strict On
Imports System.IO
Imports nexENCODE.Business.Controllers
Imports nexENCODE.Enum
Imports nexENCODE.Enum.Skin

Public Class clsNexENCODE
    Public Event ProcessError(lError As String, lSub As String)
    Public Event StatusLabel_MouseDown(sender As System.Object, e As System.Windows.Forms.MouseEventArgs)
    Public Event StatusLabel_MouseMove(sender As System.Object, e As System.Windows.Forms.MouseEventArgs)
    Public Event DisplayLabel(lData As String)
    Public Event Progress(lPercent As Integer)
    Public WithEvents lSkins As New clsSkin
    Public WithEvents lObjectHandler As New clsObjectHandler
    Public WithEvents lCDDriveHandler As nexENCODE.CDRipper.clsCDDriveHandler
    Public WithEvents lMp3Handler As clsMp3WriterHandler
    Private WithEvents lLoading As clsLoading
    Private WithEvents lScripting As clsScripting
    Private lCodeFile As String
    Public GlobalController As New GlobalController(Application.StartupPath)

    Public Sub UnloadProgram(lForm As frmMain, Optional lAnimationTime As Integer = 300, Optional lAnimationFlags As clsAPI.AnimateWindowFlags = clsAPI.AnimateWindowFlags.AW_VER_NEGATIVE Or clsAPI.AnimateWindowFlags.AW_BLEND Or clsAPI.AnimateWindowFlags.AW_HIDE)
        Try
            lSkins.AnimateWindow(lAnimationTime, lForm, lAnimationFlags)
            GlobalController.Skins.WindowSize(WindowSizes.Unloading, lForm, GlobalController.Ini.WindowPos)
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Sub UnloadProgram(lForm As frmMain)")
        End Try
    End Sub

    Private Sub lLoading_ProcessError(lError As String, lSub As String) Handles lLoading.ProcessError
        Try
            RaiseEvent ProcessError(lError, lSub)
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Private Sub lLoading_ProcessError(lError As String, lSub As String) Handles lLoading.ProcessError")
        End Try
    End Sub

    Private Sub lSkins_ProcessError(lError As String, lSub As String) Handles lSkins.ProcessError
        Try
            RaiseEvent ProcessError(lError, lSub)
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Private Sub lSkins_ProcessError(lError As String, lSub As String) Handles lSkins.ProcessError")
        End Try
    End Sub

    Private Sub RunScriptPrimitiveClick(name As String)
        lScripting.ProcessPrimitive(name & "_Click()")
    End Sub

    Private Sub lObjectHandler_ImageButton_Click(lType As ButtonTypes, lName As String) Handles lObjectHandler.ImageButton_Click
        Try
            Select Case lType
                Case ButtonTypes.Rip
                    RunScriptPrimitiveClick(lName)
                    'lCDDriveHandler.RipCurrentTrack(2, "C:\TEST\test2.wav", "C:\TEST\test2.mp3", True, False)
                Case ButtonTypes.Burn
                    RunScriptPrimitiveClick(lName)
                    'lMp3Handler.ConvertWavToMp3("C:\TEST\test2.wav", "C:\TEST\test2.mp3")
                Case ButtonTypes.RipCancel
                    RunScriptPrimitiveClick(lName)
                Case ButtonTypes.Minimize
                    RunScriptPrimitiveClick(lName)
                Case ButtonTypes.Maximize
                    RunScriptPrimitiveClick(lName)
                Case ButtonTypes.Exit
                    RunScriptPrimitiveClick(lName)
                Case ButtonTypes.Encode
                    RunScriptPrimitiveClick(lName)
                Case ButtonTypes.EncodeCancel
                    RunScriptPrimitiveClick(lName)
                Case ButtonTypes.Decode
                    RunScriptPrimitiveClick(lName)
                Case ButtonTypes.Video
                    RunScriptPrimitiveClick(lName)
            End Select
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Private Sub lObjectHandler_ImageButton_Click(lType As clsSkin.eButtonTypes) Handles lObjectHandler.ImageButton_Click")
        End Try
    End Sub

    Private Sub lObjectHandler_ProcessError(lError As String, lSub As String) Handles lObjectHandler.ProcessError
        Try
            RaiseEvent ProcessError(lError, lSub)
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Private Sub lObjectHandler_ProcessError(lError As String, lSub As String) Handles lObjectHandler.ProcessError")
        End Try
    End Sub

    Private Sub lScripting_ProcessError(lError As String, lSub As String) Handles lScripting.ProcessError
        Try
            RaiseEvent ProcessError(lError, lSub)
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Private Sub lScripting_ProcessError(lError As String, lSub As String) Handles lScripting.ProcessError")
        End Try
    End Sub

    Public Sub New(lForm As Form, Optional lAnimationTime As Integer = 300, Optional lAnimationFlags As clsAPI.AnimateWindowFlags = clsAPI.AnimateWindowFlags.AW_BLEND)
        Try
            lLoading = New clsLoading()
            lLoading.ShowLoadingForm("Loading Configuration Settings", "Initializing nexENCODE Studio")
            lLoading.SetPercent(10, "Setting Window Size")
            GlobalController.Skins.WindowSize(WindowSizes.Loading, lForm, GlobalController.Ini.WindowPos)
            lLoading.SetPercent(40, "Loading Skins")
            lSkins.LoadSkins()
            lLoading.SetPercent(70, "Applying Skin")
            lSkins.ApplySkin(lForm, GlobalController.Skins.ReadIndex(), lObjectHandler)
            lSkins.AnimateWindow(lAnimationTime, lForm, lAnimationFlags)
            lForm.Show()
            lLoading.SetPercent(80, "Initializing Scripting Object")
            lScripting = New clsScripting(GlobalController.Skins.MainWnd_CodeFile(GlobalController.Skins.ReadIndex()))
            lLoading.SetPercent(85, "Initializing CD Drives")
            lCDDriveHandler = New nexENCODE.CDRipper.clsCDDriveHandler(lForm, CChar("D"))
            lLoading.SetPercent(88, "Initializing Media Write Handler")
            lMp3Handler = New clsMp3WriterHandler(lForm)
            lLoading.SetPercent(90, "Authorizing Components")
            'lRipper.Authorize("Leon Aiossa", "698070606")
            'lEncoder.Authorize("Leon Aiossa", "680665552")
            lLoading.SetPercent(100, "Loading Completed")
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Sub New()")
        End Try
    End Sub

    Private Sub lCDDriveHandler_CDReadCanceled() Handles lCDDriveHandler.CDReadCanceled
        Try
            RaiseEvent DisplayLabel("Read Canceled")
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Private Sub lCDDriveHandler_CDReadCanceled() Handles lCDDriveHandler.CDReadCanceled")
        End Try
    End Sub

    Private Sub lCDDriveHandler_CDReadCompleted(_SecondsEllapsed As Integer) Handles lCDDriveHandler.CDReadCompleted
        Try
            RaiseEvent DisplayLabel("CD Read Completed, " & _SecondsEllapsed & " seconds ellapsed.")
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Private Sub lCDDriveHandler_CDReadCompleted(_SecondsEllapsed As Integer) Handles lCDDriveHandler.CDReadCompleted")
        End Try
    End Sub

    Private Sub lCDDriveHandler_CDReadProgress(_Percent As Integer, _SecondsRemaining As Integer, _BytesRead As Long, _BytesRemaining As Long) Handles lCDDriveHandler.CDReadProgress
        Try
            RaiseEvent DisplayLabel(_Percent & " - " & _SecondsRemaining & " - " & _BytesRead & " - " & _BytesRemaining)
            RaiseEvent Progress(_Percent)
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Private Sub lCDDriveHandler_CDReadProgress(lPercent As Integer, lSecondsRemaining As Integer, _BytesRead As Long, _BytesRemaining As Long) Handles lCDDriveHandler.CDReadProgress")
        End Try
    End Sub

    Private Sub lCDDriveHandler_CDReadStarted() Handles lCDDriveHandler.CDReadStarted
        Try
            RaiseEvent DisplayLabel("Read Started")
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Private Sub lCDDriveHandler_CDReadStarted() Handles lCDDriveHandler.CDReadStarted")
        End Try
    End Sub

    Private Sub lCDDriveHandler_DisplayLabel(lData As String) Handles lCDDriveHandler.DisplayLabel
        Try
            RaiseEvent DisplayLabel(lData)
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Private Sub lCDDriveHandler_DisplayLabel(lData As String) Handles lCDDriveHandler.DisplayLabel")
        End Try
    End Sub

    Private Sub lCDDriveHandler_ProcessError(lError As String, lSub As String) Handles lCDDriveHandler.ProcessError
        Try
            RaiseEvent ProcessError(lError, lSub)
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Private Sub lCDDriveHandler_ProcessError(lError As String, lSub As String) Handles lCDDriveHandler.ProcessError")
        End Try
    End Sub

    Private Sub lMp3Handler_EncodeCanceled() Handles lMp3Handler.EncodeCanceled
        Try
            RaiseEvent DisplayLabel("Encode Canceled")
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Private Sub lMp3Handler_EncodeCanceled() Handles lMp3Handler.EncodeCanceled")
        End Try
    End Sub

    Private Sub lMp3Handler_EncodeComplete() Handles lMp3Handler.EncodeComplete
        Try
            RaiseEvent DisplayLabel("Encode Complete")
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Private Sub lMp3Handler_EncodeComplete() Handles lMp3Handler.EncodeComplete")
        End Try
    End Sub

    Private Sub lMp3Handler_EncodeProgress(lProgress As Integer) Handles lMp3Handler.EncodeProgress
        Try
            RaiseEvent Progress(lProgress)
            RaiseEvent DisplayLabel("Progress: " & lProgress & "%")
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Private Sub lMp3Handler_EncodeProgress(lProgress As Integer) Handles lMp3Handler.EncodeProgress")
        End Try
    End Sub

    Private Sub lMp3Handler_EncodeStarted() Handles lMp3Handler.EncodeStarted
        Try
            RaiseEvent DisplayLabel("Encode Started")
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Private Sub lMp3Handler_EncodeStarted() Handles lMp3Handler.EncodeStarted")
        End Try
    End Sub

    Private Sub lMp3Handler_ProcessError(lError As String, lSub As String) Handles lMp3Handler.ProcessError
        Try

        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Private Sub lMp3Handler_ProcessError(lError As String, lSub As String) Handles lMp3Handler.ProcessError")
        End Try
    End Sub

    Private Sub lObjectHandler_StatusLabel_Click(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles lObjectHandler.StatusLabel_MouseDown
        RaiseEvent StatusLabel_MouseDown(sender, e)
    End Sub

    Private Sub lObjectHandler_StatusLabel_MouseMove(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles lObjectHandler.StatusLabel_MouseMove
        RaiseEvent StatusLabel_MouseMove(sender, e)
    End Sub
End Class