Option Explicit On
Option Strict On
'nexENCODE Studio 5.0 Alpha 1.3
'October 6th, 2013
Imports nexENCODE.Enum
Imports nexENCODE.Models

Public Class clsImageButtonEvents
    Public Event ProcessError(lError As String, lSub As String)
    Public Event ImageButton_Click(lType As ButtonTypes, lName As String)

    Public Sub ImageButton_MouseMove(ByVal sender As System.Object, e As System.Windows.Forms.MouseEventArgs)
        Try
            Dim lImageButton As PictureBox = CType(sender, PictureBox), lImage As System.Drawing.Image, lTag As ImageButtonTagModel = CType(lImageButton.Tag, ImageButtonTagModel)
            If e.Button = 0 Then
                lImage = Image.FromFile(lTag.FileName3)
                If lImage IsNot lImageButton.Image Then lImageButton.Image = lImage
            End If
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Sub ImageBox_MouseMove(ByVal sender As System.Object, ByVal e As System.EventArgs)")
        End Try
    End Sub

    Public Sub ImageButton_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Dim lImageButton As PictureBox = CType(sender, PictureBox), lImage As System.Drawing.Image, lTag As ImageButtonTagModel = CType(lImageButton.Tag, ImageButtonTagModel)
            lImage = Image.FromFile(lTag.FileName1)
            If lImage IsNot lImageButton.Image Then lImageButton.Image = lImage
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Sub ImageButton_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs)")
        End Try
    End Sub

    Public Sub ImageButton_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        Try
            Dim lImageButton As PictureBox = CType(sender, PictureBox), lImage As System.Drawing.Image, lTag As ImageButtonTagModel = CType(lImageButton.Tag, ImageButtonTagModel)
            lImage = Image.FromFile(lTag.FileName2)
            If lImage IsNot lImageButton.Image Then lImageButton.Image = lImage
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Sub ImageButton_MouseDown(ByVal sender As System.Object, ByVal e As System.EventArgs)")
        End Try
    End Sub

    Public Sub ImageButton_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        Try
            Dim lImageButton As PictureBox = CType(sender, PictureBox), lImage As System.Drawing.Image, lTag As ImageButtonTagModel = CType(lImageButton.Tag, ImageButtonTagModel)
            lImage = Image.FromFile(lTag.FileName1)
            If lImage IsNot lImageButton.Image Then lImageButton.Image = lImage
            RaiseEvent ImageButton_Click(lTag.ButtonType, lTag.Name)
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Sub ImageButton_MouseUp(ByVal sender As System.Object, ByVal e As System.EventArgs)")
        End Try
    End Sub
End Class

Public Class clsObjectHandler
    Public Event ProcessError(lError As String, lSub As String)
    Public Event ImageButton_Click(lType As ButtonTypes, lName As String)
    Public Event StatusLabel_MouseDown(sender As System.Object, e As System.Windows.Forms.MouseEventArgs)
    Public Event StatusLabel_MouseMove(sender As System.Object, e As System.Windows.Forms.MouseEventArgs)
    Private WithEvents lImageButtonEvents As New clsImageButtonEvents
    Private WithEvents lStatusLabel As Label

    Public Property StatusLabelText() As String
        Get
            Try
                Return lStatusLabel.Text
            Catch ex As Exception
                RaiseEvent ProcessError(ex.Message, "Public Property StatusLabelText() As String")
                Return Nothing
            End Try
        End Get
        Set(value As String)
            Try
                lStatusLabel.Text = value
            Catch ex As Exception
                RaiseEvent ProcessError(ex.Message, "Public Property StatusLabelText() As String")
            End Try
        End Set
    End Property

    Public Function CreateStatusLabel(width As Integer, height As Integer, left As Integer, top As Integer, form As Form) As Boolean
        Try
            lStatusLabel = New Label()
            lStatusLabel.Width = width
            lStatusLabel.Height = height
            lStatusLabel.Left = left
            lStatusLabel.Top = top
            lStatusLabel.BorderStyle = BorderStyle.None
            lStatusLabel.BackColor = Color.Transparent
            lStatusLabel.BringToFront()
            lStatusLabel.Text = "(Uninitialized)"
            form.Controls.Add(lStatusLabel)
            Return True
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Function CreateStatusLabel(width As Integer, height As Integer, left As Integer, top As Integer) As Boolean")
            Return Nothing
        End Try
    End Function

    Public Function CreateImageButton(lType As ButtonTypes, lName As String, lFileName1 As String, lFileName2 As String, lFileName3 As String, lImageLeft As Integer, lImageTop As Integer, lImageWidth As Integer, lImageHeight As Integer, lVisible As Boolean, lForm As Form) As Boolean
        Try
            Dim lImageButton As New PictureBox, lTag As New ImageButtonTagModel
            With lImageButton
                .Name = lName
                .Image = Image.FromFile(lFileName1)
                .Left = lImageLeft
                .Top = lImageTop
                .Width = lImageWidth
                .Height = lImageHeight
                .BorderStyle = BorderStyle.None
                .Visible = lVisible
                lTag.Name = lName
                lTag.ButtonType = lType
                lTag.FileName1 = lFileName1
                lTag.FileName2 = lFileName2
                lTag.FileName3 = lFileName3
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

    Private Sub lImageButtonEvents_ImageButton_Click(lType As ButtonTypes, lName As String) Handles lImageButtonEvents.ImageButton_Click
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

    Private Sub lStatusLabel_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles lStatusLabel.MouseDown
        RaiseEvent StatusLabel_MouseDown(sender, e)
    End Sub

    Private Sub lStatusLabel_MouseMove(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles lStatusLabel.MouseMove
        RaiseEvent StatusLabel_MouseMove(sender, e)
    End Sub
End Class