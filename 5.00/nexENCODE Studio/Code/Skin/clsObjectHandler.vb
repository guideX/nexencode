'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Public Class clsImageButtonEvents
    Public Event ProcessError(lError As String, lSub As String)
    Public Event ImageButton_Click(lType As clsSkin.eButtonTypes, lName As String)

    Public Sub ImageButton_MouseMove(ByVal sender As System.Object, e As System.Windows.Forms.MouseEventArgs)
        Try
            Dim lImageButton As PictureBox = CType(sender, PictureBox), lImage As System.Drawing.Image, lTag As clsSkin.gImageButtonTag = CType(lImageButton.Tag, clsSkin.gImageButtonTag)
            If e.Button = 0 Then
                lImage = Image.FromFile(lTag.iFileName3)
                If lImage IsNot lImageButton.Image Then lImageButton.Image = lImage
            End If
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Sub ImageBox_MouseMove(ByVal sender As System.Object, ByVal e As System.EventArgs)")
        End Try
    End Sub

    Public Sub ImageButton_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Dim lImageButton As PictureBox = CType(sender, PictureBox), lImage As System.Drawing.Image, lTag As clsSkin.gImageButtonTag = CType(lImageButton.Tag, clsSkin.gImageButtonTag)
            lImage = Image.FromFile(lTag.iFileName1)
            If lImage IsNot lImageButton.Image Then lImageButton.Image = lImage
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Sub ImageButton_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs)")
        End Try
    End Sub

    Public Sub ImageButton_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        Try
            Dim lImageButton As PictureBox = CType(sender, PictureBox), lImage As System.Drawing.Image, lTag As clsSkin.gImageButtonTag = CType(lImageButton.Tag, clsSkin.gImageButtonTag)
            lImage = Image.FromFile(lTag.iFileName2)
            If lImage IsNot lImageButton.Image Then lImageButton.Image = lImage
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Sub ImageButton_MouseDown(ByVal sender As System.Object, ByVal e As System.EventArgs)")
        End Try
    End Sub

    Public Sub ImageButton_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        Try
            Dim lImageButton As PictureBox = CType(sender, PictureBox), lImage As System.Drawing.Image, lTag As clsSkin.gImageButtonTag = CType(lImageButton.Tag, clsSkin.gImageButtonTag)
            lImage = Image.FromFile(lTag.iFileName1)
            If lImage IsNot lImageButton.Image Then lImageButton.Image = lImage
            RaiseEvent ImageButton_Click(lTag.iButtonType, lTag.iName)
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Sub ImageButton_MouseUp(ByVal sender As System.Object, ByVal e As System.EventArgs)")
        End Try
    End Sub
End Class

Public Class clsObjectHandler
    Public Event ProcessError(lError As String, lSub As String)
    Public Event ImageButton_Click(lType As clsSkin.eButtonTypes, lName As String)
    Private WithEvents lImageButtonEvents As New clsImageButtonEvents

    Public Function CreateImageButton(lType As clsSkin.eButtonTypes, lName As String, lFileName1 As String, lFileName2 As String, lFileName3 As String, lImageLeft As Integer, lImageTop As Integer, lImageWidth As Integer, lImageHeight As Integer, lVisible As Boolean, lForm As Form) As Boolean
        Try
            Dim lImageButton As New PictureBox, lTag As New clsSkin.gImageButtonTag
            With lImageButton
                .Name = lName
                .Image = Image.FromFile(lFileName1)
                .Left = lImageLeft
                .Top = lImageTop
                .Width = lImageWidth
                .Height = lImageHeight
                .BorderStyle = BorderStyle.None
                .Visible = lVisible
                lTag.iName = lName
                lTag.iButtonType = lType
                lTag.iFileName1 = lFileName1
                lTag.iFileName2 = lFileName2
                lTag.iFileName3 = lFileName3
                .Tag = lTag
            End With
            lForm.Controls.Add(lImageButton)
            AddHandler lImageButton.MouseDown, AddressOf lImageButtonEvents.ImageButton_MouseDown
            AddHandler lImageButton.MouseMove, AddressOf lImageButtonEvents.ImageButton_MouseMove
            AddHandler lImageButton.MouseUp, AddressOf lImageButtonEvents.ImageButton_MouseUp
            AddHandler lImageButton.MouseLeave, AddressOf lImageButtonEvents.ImageButton_MouseLeave
            Return True
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Function CreatePictureBox(lForm As Form) As Boolean")
            Return Nothing
        End Try
    End Function

    Private Sub lImageButtonEvents_ImageButton_Click(lType As clsSkin.eButtonTypes, lName As String) Handles lImageButtonEvents.ImageButton_Click
        Try
            RaiseEvent ImageButton_Click(lType, lName)
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Private Sub lImageButtonEvents_ImageButton_Click(lType As clsSkin.eButtonTypes) Handles lImageButtonEvents.ImageButton_Click")
        End Try
    End Sub

    Private Sub lImageButtonEvents_ProcessError(lError As String, lSub As String) Handles lImageButtonEvents.ProcessError
        Try
            RaiseEvent ProcessError(lError, lSub)
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Private Sub lImageButtonEvents_ProcessError(lError As String, lSub As String) Handles lImageButtonEvents.ProcessError")
        End Try
    End Sub
End Class