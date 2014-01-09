Option Explicit On
Option Strict On
Public Enum eRipCDStep
    rStep0_Uninitialized = 0
    rStep1_Tracks = 1
    rStep2_Format = 2
    rStep3_Location = 3
    rStep4_Ready = 4
    rStep5_RipProgress = 5
    rStep6_EncodeProgress = 6
    rStep7_Playback = 7
End Enum
Public Class clsWizardPanels
    Public Property wPanel1 As Telerik.WinControls.UI.RadPanel
    Public Property wPanel2 As Telerik.WinControls.UI.RadPanel
    Public Property wPanel3 As Telerik.WinControls.UI.RadPanel
    Public Property wPanel4 As Telerik.WinControls.UI.RadPanel
    Public Property wPanel5 As Telerik.WinControls.UI.RadPanel
    Public Property wPanel6 As Telerik.WinControls.UI.RadPanel
    Public Property wPanel7 As Telerik.WinControls.UI.RadPanel
End Class
Public Class clsRipCDWizard
#Region "Public Variables"
    Public ripStep As eRipCDStep
    Public wizardPanels As clsWizardPanels
    Public Event CloseWizard()
#End Region
#Region "Public Subs and Functions"
    Public Sub New()
        Try
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Sub CloseNow()
        Try
            RaiseEvent CloseWizard()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Sub Cancel()
        Dim dialogResult As DialogResult
        Try
            dialogResult = MessageBox.Show(
                "Are you sure you wish to close the wizard?",
                "nexENCODE Studio - Rip CD Wizard",
                 MessageBoxButtons.OKCancel,
                 MessageBoxIcon.Question, MessageBoxDefaultButton.Button1)
            If (dialogResult = MessageBoxButtons.OK) Then
                RaiseEvent CloseWizard()
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Function Initialize() As Boolean
        Try
            ripStep = eRipCDStep.rStep1_Tracks
            HidePanels()
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Sub PreviousStep()
        Try
            HidePanels()
            Select Case ripStep
                Case eRipCDStep.rStep2_Format
                    wizardPanels.wPanel1.Visible = True
                    ripStep = eRipCDStep.rStep1_Tracks
                Case eRipCDStep.rStep3_Location
                    wizardPanels.wPanel2.Visible = True
                    ripStep = eRipCDStep.rStep2_Format
                Case eRipCDStep.rStep4_Ready
                    wizardPanels.wPanel3.Visible = True
                    ripStep = eRipCDStep.rStep3_Location
                Case eRipCDStep.rStep5_RipProgress
                    wizardPanels.wPanel4.Visible = True
                    ripStep = eRipCDStep.rStep4_Ready
                Case eRipCDStep.rStep5_RipProgress
                    wizardPanels.wPanel5.Visible = True
                    ripStep = eRipCDStep.rStep5_RipProgress
                Case eRipCDStep.rStep7_Playback
                    wizardPanels.wPanel6.Visible = True
                    ripStep = eRipCDStep.rStep6_EncodeProgress
            End Select
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Sub NextStep()
        Try
            HidePanels()
            Select Case ripStep
                Case eRipCDStep.rStep0_Uninitialized
                    wizardPanels.wPanel1.Visible = True
                    ripStep = eRipCDStep.rStep1_Tracks
                Case eRipCDStep.rStep1_Tracks
                    wizardPanels.wPanel2.Visible = True
                    ripStep = eRipCDStep.rStep2_Format
                Case eRipCDStep.rStep2_Format
                    wizardPanels.wPanel3.Visible = True
                    ripStep = eRipCDStep.rStep3_Location
                Case eRipCDStep.rStep3_Location
                    wizardPanels.wPanel4.Visible = True
                    ripStep = eRipCDStep.rStep4_Ready
                Case eRipCDStep.rStep4_Ready
                    wizardPanels.wPanel5.Visible = True
                    ripStep = eRipCDStep.rStep5_RipProgress
                Case eRipCDStep.rStep5_RipProgress
                    wizardPanels.wPanel6.Visible = True
                    ripStep = eRipCDStep.rStep6_EncodeProgress
                Case eRipCDStep.rStep6_EncodeProgress
                    wizardPanels.wPanel7.Visible = True
                    ripStep = eRipCDStep.rStep7_Playback
                    'Case eRipCDStep.rStep7_Playback
            End Select
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region
#Region "Private Helpers"
    Private Sub HidePanels()
        Try
            wizardPanels.wPanel1.Visible = False
            wizardPanels.wPanel2.Visible = False
            wizardPanels.wPanel3.Visible = False
            wizardPanels.wPanel4.Visible = False
            wizardPanels.wPanel5.Visible = False
            wizardPanels.wPanel6.Visible = False
            'wizardPanels.wPanel7.Visible = False
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region
End Class