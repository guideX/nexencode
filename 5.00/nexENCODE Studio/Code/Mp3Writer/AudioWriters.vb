'nexENCODE Studio 5.0 Alpha 1.3
'October 6th, 2013
Option Explicit On
Option Strict On
Imports System.IO
Imports System.Windows.Forms
Imports nexENCODE.WaveLib

Namespace Media
    Public MustInherit Class AudioWriter
        Inherits BinaryWriter
        Protected m_InputDataFormat As WaveFormat

        Public Sub New(Output As Stream, InputDataFormat As WaveFormat)
            MyBase.New(Output, System.Text.Encoding.ASCII)
            m_InputDataFormat = InputDataFormat
        End Sub

        Public Sub New(Output As Stream, Config As AudioWriterConfig)
            Me.New(Output, Config.Format)
        End Sub

        Protected MustOverride Function GetOptimalBufferSize() As Integer
        Private Shared m_ConfigWidth As Integer = 368
        Private Shared m_ConfigHeight As Integer = 264

        Protected Overridable Function GetWriterConfig() As AudioWriterConfig
            Return New AudioWriterConfig(m_InputDataFormat)
        End Function

        Public ReadOnly Property WriterConfig() As AudioWriterConfig
            Get
                Return GetWriterConfig()
            End Get
        End Property

        Public Shared Property ConfigWidth() As Integer
            Get
                Return m_ConfigWidth
            End Get
            Set(value As Integer)
                m_ConfigWidth = value
            End Set
        End Property

        Public Shared Property ConfigHeight() As Integer
            Get
                Return m_ConfigHeight
            End Get
            Set(value As Integer)
                m_ConfigHeight = value
            End Set
        End Property

        Public ReadOnly Property OptimalBufferSize() As Integer
            Get
                Return GetOptimalBufferSize()
            End Get
        End Property

        Public Overrides Sub Write(value As String)
            Throw New NotSupportedException("Write(string value) is not supported")
        End Sub

        Public Overrides Sub Write(value As Single)
            Throw New NotSupportedException("Write(float value) is not supported")
        End Sub

        Public Overrides Sub Write(value As ULong)
            Throw New NotSupportedException("Write(ulong value) is not supported")
        End Sub

        Public Overrides Sub Write(value As Long)
            Throw New NotSupportedException("Write(long value) is not supported")
        End Sub

        Public Overrides Sub Write(value As UInteger)
            Throw New NotSupportedException("Write(uint value) is not supported")
        End Sub

        Public Overrides Sub Write(value As Integer)
            Throw New NotSupportedException("Write(int value) is not supported")
        End Sub

        Public Overrides Sub Write(value As UShort)
            Throw New NotSupportedException("Write(ushort value) is not supported")
        End Sub

        Public Overrides Sub Write(value As Short)
            Throw New NotSupportedException("Write(short value) is not supported")
        End Sub

        Public Overrides Sub Write(value As Decimal)
            Throw New NotSupportedException("Write(decimal value) is not supported")
        End Sub

        Public Overrides Sub Write(value As Double)
            Throw New NotSupportedException("Write(double value) is not supported")
        End Sub

        Public Overrides Sub Write(chars As Char(), index As Integer, count As Integer)
            Throw New NotSupportedException("Write(char[] chars, int index, int count) is not supported")
        End Sub

        Public Overrides Sub Write(chars As Char())
            Throw New NotSupportedException("Write(char[] chars) is not supported")
        End Sub

        Public Overrides Sub Write(ch As Char)
            Throw New NotSupportedException("Write(char ch) is not supported")
        End Sub

        Public Overrides Sub Write(value As SByte)
            Throw New NotSupportedException("Write(sbyte value) is not supported")
        End Sub

        Public Overrides Sub Write(value As Byte)
            Throw New NotSupportedException("Write(byte value) is not supported")
        End Sub

        Public Overrides Sub Write(value As Boolean)
            Throw New NotSupportedException("Write(bool value) is not supported")
        End Sub
    End Class
End Namespace