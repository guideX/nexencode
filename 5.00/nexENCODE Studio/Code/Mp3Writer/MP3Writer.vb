'nexENCODE Studio 5.0 Alpha 1.3
'October 6th, 2013
Option Explicit On
Option Strict On
Imports System.IO
Imports nexENCODE.LAME
Imports nexENCODE.Media
Imports nexENCODE.WaveLib

Namespace Media.Mp3
    Public Class Mp3Writer
        Inherits AudioWriter
        Private closed As Boolean = False
        Private m_Mp3Config As BE_CONFIG = Nothing
        Private m_hLameStream As UInteger = 0
        Private m_InputSamples As UInteger = 0
        Private m_OutBufferSize As UInteger = 0
        Private m_InBuffer As Byte() = Nothing
        Private m_InBufferPos As Integer = 0
        Private m_OutBuffer As Byte() = Nothing

        Public Sub New(Output As Stream, InputDataFormat As WaveFormat)
            Me.New(Output, InputDataFormat, New BE_CONFIG(InputDataFormat))
        End Sub

        Public Sub New(Output As Stream, cfg As Mp3WriterConfig)
            Me.New(Output, cfg.Format, cfg.Mp3Config)
        End Sub

        Public Sub New(Output As Stream, InputDataFormat As WaveFormat, Mp3Config As BE_CONFIG)
            MyBase.New(Output, InputDataFormat)
            Try
                m_Mp3Config = Mp3Config
                Dim LameResult As UInteger = Lame_encDll.beInitStream(m_Mp3Config, m_InputSamples, m_OutBufferSize, m_hLameStream)
                If LameResult <> Lame_encDll.BE_ERR_SUCCESSFUL Then
                    Throw New ApplicationException(String.Format("Lame_encDll.beInitStream failed with the error code {0}", LameResult))
                End If
                m_InBuffer = New Byte(CInt(m_InputSamples * 2 - 1)) {}
                'Input buffer is expected as short[]
                m_OutBuffer = New Byte(CInt(m_OutBufferSize - 1)) {}
            Catch
                MyBase.Close()
                Throw
            End Try
        End Sub

        Public ReadOnly Property Mp3Config() As BE_CONFIG
            Get
                Return m_Mp3Config
            End Get
        End Property

        Protected Overrides Function GetOptimalBufferSize() As Integer
            Return m_InBuffer.Length
        End Function

        Public Overrides Sub Close()
            If Not closed Then
                Try
                    Dim EncodedSize As UInteger = 0
                    If m_InBufferPos > 0 Then
                        If Lame_encDll.EncodeChunk(m_hLameStream, m_InBuffer, 0, CUInt(m_InBufferPos), m_OutBuffer, EncodedSize) = Lame_encDll.BE_ERR_SUCCESSFUL Then
                            If EncodedSize > 0 Then
                                MyBase.Write(m_OutBuffer, 0, CInt(EncodedSize))
                            End If
                        End If
                    End If
                    EncodedSize = 0
                    If Lame_encDll.beDeinitStream(m_hLameStream, m_OutBuffer, EncodedSize) = Lame_encDll.BE_ERR_SUCCESSFUL Then
                        If EncodedSize > 0 Then
                            MyBase.Write(m_OutBuffer, 0, CInt(EncodedSize))
                        End If
                    End If
                Finally
                    Lame_encDll.beCloseStream(m_hLameStream)
                End Try
            End If
            closed = True
            MyBase.Close()
        End Sub

        Public Overrides Sub Write(buffer__1 As Byte(), index As Integer, count As Integer)
            Dim ToCopy As Integer = 0
            Dim EncodedSize As UInteger = 0
            Dim LameResult As UInteger
            While count > 0
                If m_InBufferPos > 0 Then
                    ToCopy = Math.Min(count, m_InBuffer.Length - m_InBufferPos)
                    Buffer.BlockCopy(buffer__1, index, m_InBuffer, m_InBufferPos, ToCopy)
                    m_InBufferPos += ToCopy
                    index += ToCopy
                    count -= ToCopy
                    If m_InBufferPos >= m_InBuffer.Length Then
                        m_InBufferPos = 0
                        If (InlineAssignHelper(LameResult, Lame_encDll.EncodeChunk(m_hLameStream, m_InBuffer, m_OutBuffer, EncodedSize))) = Lame_encDll.BE_ERR_SUCCESSFUL Then
                            If EncodedSize > 0 Then
                                MyBase.Write(m_OutBuffer, 0, CInt(EncodedSize))
                            End If
                        Else
                            Throw New ApplicationException(String.Format("Lame_encDll.EncodeChunk failed with the error code {0}", LameResult))
                        End If
                    End If
                Else
                    If count >= m_InBuffer.Length Then
                        If (InlineAssignHelper(LameResult, Lame_encDll.EncodeChunk(m_hLameStream, buffer__1, index, CUInt(m_InBuffer.Length), m_OutBuffer, EncodedSize))) = Lame_encDll.BE_ERR_SUCCESSFUL Then
                            If EncodedSize > 0 Then
                                MyBase.Write(m_OutBuffer, 0, CInt(EncodedSize))
                            End If
                        Else
                            Throw New ApplicationException(String.Format("Lame_encDll.EncodeChunk failed with the error code {0}", LameResult))
                        End If
                        count -= m_InBuffer.Length
                        index += m_InBuffer.Length
                    Else
                        Buffer.BlockCopy(buffer__1, index, m_InBuffer, 0, count)
                        m_InBufferPos = count
                        index += count
                        count = 0
                    End If
                End If
            End While
        End Sub

        Public Overrides Sub Write(buffer As Byte())
            Me.Write(buffer, 0, buffer.Length)
        End Sub

        Protected Overrides Function GetWriterConfig() As AudioWriterConfig
            Return New Mp3WriterConfig(m_InputDataFormat, Mp3Config)
        End Function

        Private Shared Function InlineAssignHelper(Of T)(ByRef target As T, value As T) As T
            target = value
            Return value
        End Function
    End Class
End Namespace