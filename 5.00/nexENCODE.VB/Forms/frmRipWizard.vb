Option Explicit On
Option Strict On
Public Class frmRipWizard
#Region "Private Variables"
    Private WithEvents ripWizard As clsRipCDWizard
#End Region
#Region "Form Events"
    Private Sub FrmRip_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim wizardPanels As clsWizardPanels
        Try
            wizardPanels = New clsWizardPanels()
            wizardPanels.wPanel1 = pnlStep1
            wizardPanels.wPanel2 = pnlStep2
            wizardPanels.wPanel3 = pnlStep3
            wizardPanels.wPanel4 = pnlStep4
            wizardPanels.wPanel5 = pnlStep5
            wizardPanels.wPanel6 = pnlStep6
            ripWizard = New clsRipCDWizard()
            ripWizard.wizardPanels = wizardPanels
            Me.Size = New System.Drawing.Size(577, 345)
            ripWizard.Initialize()
        Catch ex As Exception
            ProcessError(ex.Message, "Private Sub FrmRip_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load")
        End Try
    End Sub
    'Private Sub cmdCancel_Click(sender As System.Object, e As System.EventArgs) Handles cmdCancel.Click
    'Try
    'ripWizard.Cancel()
    'Catch ex As Exception
    'ProcessError(ex.Message, "Private Sub cmdCancel_Click(sender As System.Object, e As System.EventArgs) Handles cmdCancel.Click")
    'End Try
    'End Sub
    Private Sub cmdNext_Click(sender As System.Object, e As System.EventArgs) Handles cmdNext.Click
        Try
            ripWizard.Initialize()
        Catch ex As Exception
            ProcessError(ex.Message, "Private Sub cmdNext_Click(sender As System.Object, e As System.EventArgs) Handles cmdNext.Click")
        End Try
    End Sub
    Private Sub cmdFinish_Click(sender As System.Object, e As System.EventArgs) Handles cmdFinish.Click
        Try

        Catch ex As Exception
            ProcessError(ex.Message, "Private Sub cmdFinish_Click(sender As System.Object, e As System.EventArgs) Handles cmdFinish.Click")
        End Try
    End Sub
    Private Sub cmdPrevious_Click(sender As System.Object, e As System.EventArgs) Handles cmdPrevious.Click
        Try
            ripWizard.PreviousStep()
        Catch ex As Exception
            ProcessError(ex.Message, "Private Sub cmdPrevious_Click(sender As System.Object, e As System.EventArgs) Handles cmdPrevious.Click")
        End Try
    End Sub
#End Region
#Region "Rip Wizard Events"
    Private Sub ripWizard_CloseWizard() Handles ripWizard.CloseWizard
        Try
            ripWizard.CloseNow()
        Catch ex As Exception
            ProcessError(ex.Message, "Private Sub ripWizard_CloseWizard() Handles ripWizard.CloseWizard")
        End Try
    End Sub
#End Region
#Region "Error Handling"
    Private Sub ProcessError(lError As String, lSub As String)
        Try
            MsgBox(lSub & " - " & lError)
        Catch ex As Exception
        End Try
    End Sub
#End Region
End Class