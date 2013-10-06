'nexENCODE Studio 5.0 Alpha 1.3
'October 6th, 2013
Option Explicit On
Option Strict On
Public Class clsLoading
    Public Event ProcessError(lError As String, lSub As String)
    Private lLoadingForm As frmLoading, lShowLoadingForm As Boolean

    Public Sub CloseLoadingForm()
        Try
            If lShowLoadingForm = True Then
                lLoadingForm.Close()
                lLoadingForm = Nothing
            End If
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Sub CloseLoadingForm()")
        End Try
    End Sub

    Public Sub SetPercent(lPercent As Integer, lReason As String)
        Try
            If lShowLoadingForm = True Then
                lLoadingForm.lblLoadingMessage.Text = lReason
                lLoadingForm.Refresh()
                Threading.Thread.Sleep(200)
                lLoadingForm.ProgressBar1.Value = lPercent
                If lPercent = 100 Then
                    CloseLoadingForm()
                End If
            End If
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Sub SetPercent(lPercent As Integer)")
        End Try
    End Sub

    Public Sub ShowLoadingForm(lCaption As String, lTitle As String)
        Try
            If lShowLoadingForm = True Then
                lLoadingForm = New frmLoading
                lLoadingForm.Text = lCaption
                lLoadingForm.lblLoadingMessage.Text = lTitle
                lLoadingForm.ProgressBar1.Maximum = 100
                lLoadingForm.ProgressBar1.Value = 1
                lLoadingForm.Show()
            End If
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Sub ShowLoadingForm(lTitle As String)")
        End Try
    End Sub

    Public Sub New(Optional _ShowLoadingForm As Boolean = True)
        Try
            lShowLoadingForm = _ShowLoadingForm
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Sub New(Optional lShowLoadingForm As Boolean = True)")
        End Try
    End Sub
End Class