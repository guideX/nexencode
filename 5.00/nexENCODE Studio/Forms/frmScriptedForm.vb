Option Explicit On
Option Strict On
Public Class frmScriptedForm
    Private lFormName As String = ""
    Public Property FormTitle As String
        Get
            Return Me.Text
        End Get
        Set(value As String)
            Me.Text = value
        End Set
    End Property
    Public Property FormName As String
        Get
            Return lFormName
        End Get
        Set(value As String)
            lFormName = value
        End Set
    End Property
    Private Sub frmScriptedForm_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
    End Sub
End Class