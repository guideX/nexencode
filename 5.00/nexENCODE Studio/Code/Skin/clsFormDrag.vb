'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Public Class clsFormDrag
    Public Event ProcessError(lError As String, lSub As String)
    Private lNewPoint As New System.Drawing.Point(), lDragPointA As Integer, lDragPointB As Integer

    Public Sub Form_MouseDown(lForm As Form, lMousePosition As Point, sender As Object, e As System.Windows.Forms.MouseEventArgs)
        Try
            lDragPointA = lMousePosition.X - lForm.Location.X
            lDragPointB = lMousePosition.Y - lForm.Location.Y
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Sub Form_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs)")
        End Try
    End Sub

    Public Sub Form_MouseMove(lForm As Form, lMousePosition As Point, sender As Object, e As System.Windows.Forms.MouseEventArgs)
        Try
            If e.Button = MouseButtons.Left Then 'Left Mouse Button Drag
                lNewPoint = lMousePosition
                lNewPoint.X = lNewPoint.X - (lDragPointA)
                lNewPoint.Y = lNewPoint.Y - (lDragPointB)
                lForm.Location = lNewPoint
            ElseIf e.Button = MouseButtons.Right Then 'Right Mouse Button Drag
                'Do Nothing
            End If
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Private Sub frmMain_MouseMove(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseMove")
        End Try
    End Sub
End Class