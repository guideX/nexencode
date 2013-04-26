'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Imports System.IO

Namespace WaveLib
    Public Class WaveStream
        Inherits Stream
        Implements IDisposable
        Private m_Stream As Stream
        Private m_DataPos As Long
        Private m_Length As Long

        Private m_Format As WaveFormat

        Public ReadOnly Property Format() As WaveFormat
            Get
                Return m_Format
            End Get
        End Property

        Private Function ReadChunk(reader As BinaryReader) As String
            Dim ch As Byte() = New Byte(3) {}
            reader.Read(ch, 0, ch.Length)
            Return System.Text.Encoding.ASCII.GetString(ch)
        End Function

        Private Sub ReadHeader()
            Dim Reader As New BinaryReader(m_Stream)
            If ReadChunk(Reader) <> "RIFF" Then
                Throw New Exception("Invalid file format")
            End If

            Reader.ReadInt32()
            ' File length minus first 8 bytes of RIFF description, we don't use it
            If ReadChunk(Reader) <> "WAVE" Then
                Throw New Exception("Invalid file format")
            End If

            If ReadChunk(Reader) <> "fmt " Then
                Throw New Exception("Invalid file format")
            End If

            Dim FormatLength As Integer = Reader.ReadInt32()
            If FormatLength < 16 Then
                ' bad format chunk length
                Throw New Exception("Invalid file format")
            End If

            m_Format = New WaveFormat(22050, 16, 2)
            ' initialize to any format
            m_Format.wFormatTag = Reader.ReadInt16()
            m_Format.nChannels = Reader.ReadInt16()
            m_Format.nSamplesPerSec = Reader.ReadInt32()
            m_Format.nAvgBytesPerSec = Reader.ReadInt32()
            m_Format.nBlockAlign = Reader.ReadInt16()
            m_Format.wBitsPerSample = Reader.ReadInt16()
            If FormatLength > 16 Then
                m_Stream.Position += (FormatLength - 16)
            End If
            ' assume the data chunk is aligned
            While m_Stream.Position < m_Stream.Length AndAlso ReadChunk(Reader) <> "data"


            End While

            If m_Stream.Position >= m_Stream.Length Then
                Throw New Exception("Invalid file format")
            End If

            m_Length = Reader.ReadInt32()
            m_DataPos = m_Stream.Position

            Position = 0
        End Sub

        Public Sub New(fileName As String)
            Me.New(New FileStream(fileName, FileMode.Open))
        End Sub
        Public Sub New(S As Stream)
            m_Stream = S
            ReadHeader()
        End Sub
        Protected Overrides Sub Finalize()
            Try
                Dispose()
            Finally
                MyBase.Finalize()
            End Try
        End Sub
        Public Overloads Sub Dispose() ' Implements IDisposable.Dispose
            If m_Stream IsNot Nothing Then
                m_Stream.Close()
            End If
            GC.SuppressFinalize(Me)
        End Sub

        Public Overrides ReadOnly Property CanRead() As Boolean
            Get
                Return True
            End Get
        End Property
        Public Overrides ReadOnly Property CanSeek() As Boolean
            Get
                Return True
            End Get
        End Property
        Public Overrides ReadOnly Property CanWrite() As Boolean
            Get
                Return False
            End Get
        End Property
        Public Overrides ReadOnly Property Length() As Long
            Get
                Return m_Length
            End Get
        End Property
        Public Overrides Property Position() As Long
            Get
                Return m_Stream.Position - m_DataPos
            End Get
            Set(value As Long)
                Seek(value, SeekOrigin.Begin)
            End Set
        End Property
        Public Overrides Sub Close()
            Dispose()
        End Sub
        Public Overrides Sub Flush()
        End Sub
        Public Overrides Sub SetLength(len As Long)
            Throw New InvalidOperationException()
        End Sub
        Public Overrides Function Seek(pos As Long, o As SeekOrigin) As Long
            Select Case o
                Case SeekOrigin.Begin
                    m_Stream.Position = pos + m_DataPos
                    Exit Select
                Case SeekOrigin.Current
                    m_Stream.Seek(pos, SeekOrigin.Current)
                    Exit Select
                Case SeekOrigin.[End]
                    m_Stream.Position = m_DataPos + m_Length - pos
                    Exit Select
            End Select
            Return Me.Position
        End Function
        Public Overrides Function Read(buf As Byte(), ofs As Integer, count As Integer) As Integer
            Dim toread As Integer = CInt(Math.Min(count, m_Length - Position))
            Return m_Stream.Read(buf, ofs, toread)
        End Function
        Public Overrides Sub Write(buf As Byte(), ofs As Integer, count As Integer)
            Throw New InvalidOperationException()
        End Sub
    End Class
End Namespace
