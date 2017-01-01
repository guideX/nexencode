Option Explicit On
Option Strict On
Public Class frmScriptedForm
    Private lFormName As String = ""
    Private lButtons As New List(Of gButton)
    Private Structure gButton
        Public bVariableIdentifier As String
        Public bName As String
        Public bButton As Button
    End Structure
    Public Sub AddButton(lVariableIdentifier As String)
        Dim lButton As New gButton
        lButton.bVariableIdentifier = lVariableIdentifier
        lButton.bButton = New Button()
        lButtons.Add(lButton)
    End Sub
    Public Property ButtonName(variableIdentifier As String) As String
        Get
            For Each b As gButton In lButtons
                If (b.bVariableIdentifier = variableIdentifier) Then
                    Return b.bName
                End If
            Next b
            Return Nothing
        End Get
        Set(value As String)
            For i As Integer = 0 To lButtons.Count() - 1
                If (lButtons(i).bVariableIdentifier = variableIdentifier) Then
                    lButtons(i).bName.Equals(value)
                End If
            Next i
        End Set
    End Property
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