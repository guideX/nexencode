'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Imports System.Collections.Generic
Imports System.Text
Imports System.Runtime.InteropServices
Imports NAudio.CoreAudioApi.Interfaces

Namespace NAudio.CoreAudioApi
    Public Class MMDeviceCollection
        Implements IEnumerable(Of MMDevice)
        Private _MMDeviceCollection As IMMDeviceCollection

        Public ReadOnly Property Count() As Integer
            Get
                Dim result As Integer
                Marshal.ThrowExceptionForHR(_MMDeviceCollection.GetCount(result))
                Return result
            End Get
        End Property

        Default Public ReadOnly Property Item(index As Integer) As MMDevice
            Get
                Dim result As IMMDevice
                _MMDeviceCollection.Item(index, result)
                Return New MMDevice(result)
            End Get
        End Property

        Friend Sub New(parent As IMMDeviceCollection)
            _MMDeviceCollection = parent
        End Sub

#Region "IEnumerable<MMDevice> Members"
        Public Function GetEnumerator() As IEnumerator(Of MMDevice) Implements IEnumerable(Of MMDevice).GetEnumerator
            For index As Integer = 0 To Count - 1
                yield Return Me(index)
            Next
        End Function
#End Region
#Region "IEnumerable Members"
        Private Function System_Collections_IEnumerable_GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
            Return GetEnumerator()
        End Function
#End Region
    End Class
End Namespace