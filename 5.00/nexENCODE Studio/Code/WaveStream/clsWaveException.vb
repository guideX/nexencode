'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Public Class clsWaveException
    Inherits Exception

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal _Message As String)
        MyBase.New(_Message)
    End Sub

    Public Sub New(ByVal _Message As String, ByVal _innerException As Exception)
        MyBase.New(_Message, _innerException)
    End Sub
End Class