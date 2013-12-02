'nexENCODE Studio 5.0 Alpha 1.3
'October 6th, 2013
Option Explicit On
Option Strict On
Public Class frmMain
    Private WithEvents lnexENCODE As clsNexENCODE, lFormDrag As New clsFormDrag
#Region "FORM_EVENTS"
    Private Sub frmMain_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            lnexENCODE.UnloadProgram(Me)
        Catch ex As Exception
            ProcessError(ex.Message, "Private Sub frmMain_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing")
        End Try
    End Sub

    Private Sub frmMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Try
            lnexENCODE = New clsNexENCODE(Me)
        Catch ex As Exception
            ProcessError(ex.Message, "Private Sub frmMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load")
        End Try
    End Sub

    Private Sub frmMain_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDown
        Try
            lFormDrag.Form_MouseDown(Me, MousePosition, sender, e)
        Catch ex As Exception
            ProcessError(ex.Message, "Private Sub frmMain_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDown")
        End Try
    End Sub

    Private Sub frmMain_MouseMove(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseMove
        Try
            lFormDrag.Form_MouseMove(Me, MousePosition, sender, e)
        Catch ex As Exception
            ProcessError(ex.Message, "Private Sub frmMain_MouseMove(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseMove")
        End Try
    End Sub

    Private Sub lnexENCODE_Progress(lPercent As Integer) Handles lnexENCODE.Progress
        Try
            'ProgressBar1.V alue = lPercent
        Catch ex As Exception
            ProcessError(ex.Message, "Private Sub lnexENCODE_Progress(lPercent As Integer) Handles lnexENCODE.Progress")
        End Try
    End Sub

    Private Sub lnexENCODE_DisplayLabel(lData As String) Handles lnexENCODE.DisplayLabel
        Try
            lnexENCODE.lObjectHandler.StatusLabelText = lData
            Me.Refresh()
        Catch ex As Exception
            ProcessError(ex.Message, "Private Sub lnexENCODE_DisplayLabel(lData As String) Handles lnexENCODE.DisplayLabel")
        End Try
    End Sub

    Private Sub lnexENCODE_StatusLabel_Click(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles lnexENCODE.StatusLabel_MouseDown
        lFormDrag.Form_MouseDown(Me, MousePosition, sender, e)
    End Sub

    Private Sub lnexENCODE_StatusLabel_MouseMove(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles lnexENCODE.StatusLabel_MouseMove
        lFormDrag.Form_MouseMove(Me, MousePosition, sender, e)
    End Sub
#End Region
#Region "ERROR_HANDLING"
    Private Sub ProcessError(lError As String, lSub As String)
        Try
            MsgBox(lSub & " - " & lError)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub lnexENCODE_ProcessError(lError As String, lSub As String) Handles lnexENCODE.ProcessError
        Try
            ProcessError(lError, lSub)
        Catch ex As Exception
            ProcessError(ex.Message, "Private Sub lnexENCODE_ProcessError(lError As String, lSub As String) Handles lnexENCODE.ProcessError")
        End Try
    End Sub

    Private Sub lFormDrag_ProcessError(lError As String, lSub As String) Handles lFormDrag.ProcessError
        Try
            ProcessError(lError, lSub)
        Catch ex As Exception
            ProcessError(ex.Message, "Private Sub lFormDrag_ProcessError(lError As String, lSub As String) Handles lFormDrag.ProcessError")
        End Try
    End Sub
#End Region
End Class