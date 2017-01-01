'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Imports System.Collections.Generic
Imports System.Text
Imports NAudio.CoreAudioApi.Interfaces

Namespace NAudio.CoreAudioApi
    Public Class PropertyStoreProperty
        Private propertyKey As PropertyKey
        Private propertyValue As PropVariant

        Friend Sub New(key As PropertyKey, value As PropVariant)
            propertyKey = key
            propertyValue = value
        End Sub

        Public ReadOnly Property Key() As PropertyKey
            Get
                Return propertyKey
            End Get
        End Property

        Public ReadOnly Property Value() As Object
            Get
                Return propertyValue.Value
            End Get
        End Property
    End Class
End Namespace