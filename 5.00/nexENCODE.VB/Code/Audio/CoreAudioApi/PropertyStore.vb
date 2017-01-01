'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Imports System.Collections.Generic
Imports System.Text
Imports System.Runtime.InteropServices
Imports NAudio.CoreAudioApi.Interfaces

Namespace NAudio.CoreAudioApi
    Public Class PropertyStore
        Private storeInterface As IPropertyStore

        Public ReadOnly Property Count() As Integer
            Get
                Dim result As Integer
                Marshal.ThrowExceptionForHR(storeInterface.GetCount(result))
                Return result
            End Get
        End Property

        Default Public ReadOnly Property Item(index As Integer) As PropertyStoreProperty
            Get
                Dim result As PropVariant
                Dim key As PropertyKey = [Get](index)
                Marshal.ThrowExceptionForHR(storeInterface.GetValue(key, result))
                Return New PropertyStoreProperty(key, result)
            End Get
        End Property

        Public Function Contains(guid As Guid) As Boolean
            For i As Integer = 0 To Count - 1
                Dim key As PropertyKey = [Get](i)
                If key.formatId = guid Then
                    Return True
                End If
            Next
            Return False
        End Function

        Default Public ReadOnly Property Item(guid As Guid) As PropertyStoreProperty
            Get
                Dim result As PropVariant
                For i As Integer = 0 To Count - 1
                    Dim key As PropertyKey = [Get](i)
                    If key.formatId = guid Then
                        Marshal.ThrowExceptionForHR(storeInterface.GetValue(key, result))
                        Return New PropertyStoreProperty(key, result)
                    End If
                Next
                Return Nothing
            End Get
        End Property

        Public Function [Get](index As Integer) As PropertyKey
            Dim key As PropertyKey
            Marshal.ThrowExceptionForHR(storeInterface.GetAt(index, key))
            Return key
        End Function

        Public Function GetValue(index As Integer) As PropVariant
            Dim result As PropVariant
            Dim key As PropertyKey = [Get](index)
            Marshal.ThrowExceptionForHR(storeInterface.GetValue(key, result))
            Return result
        End Function

        Friend Sub New(store As IPropertyStore)
            Me.storeInterface = store
        End Sub
    End Class
End Namespace