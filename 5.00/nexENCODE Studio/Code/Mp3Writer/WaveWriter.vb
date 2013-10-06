'nexENCODE Studio 5.0 Alpha 1.3
'October 6th, 2013
Option Explicit On
Option Strict On
Imports System.IO
Imports nexENCODE.WaveLib
Imports nexENCODE.Media

Namespace Yeti.MMedia
    Public Class WaveWriter
        Inherits AudioWriter
        Private Const WaveHeaderSize As UInteger = 38
        Private Const WaveFormatSize As UInteger = 18
        Private m_AudioDataSize As UInteger = 0
        Private m_WrittenBytes As UInteger = 0
        Private closed As Boolean = False

        Public Sub New(Output As Stream, Format As WaveFormat, AudioDataSize As UInteger)
            MyBase.New(Output, Format)
            m_AudioDataSize = AudioDataSize
            WriteWaveHeader()
        End Sub

        Public Sub New(Output As Stream, Format As WaveFormat)
            MyBase.New(Output, Format)
            If Not OutStream.CanSeek Then
                Throw New ArgumentException("The stream must supports seeking if AudioDataSize is not supported", "Output")
            End If
            OutStream.Seek(WaveHeaderSize + 8, SeekOrigin.Current)
        End Sub

        Private Function Int2ByteArr(val As UInteger) As Byte()
            Dim res As Byte() = New Byte(3) {}
            For i As Integer = 0 To 3
                res(i) = CByte(val >> (i * 8))
            Next
            Return res
        End Function

        Private Function Int2ByteArr(val As Short) As Byte()
            Dim res As Byte() = New Byte(1) {}
            For i As Integer = 0 To 1
                res(i) = CByte(val >> (i * 8))
            Next
            Return res
        End Function

        Protected Sub WriteWaveHeader()
            Write(New Byte() {CByte(AscW("R"c)), CByte(AscW("I"c)), CByte(AscW("F"c)), CByte(AscW("F"c))})
            Write(Int2ByteArr(m_AudioDataSize + WaveHeaderSize))
            Write(New Byte() {CByte(AscW("W"c)), CByte(AscW("A"c)), CByte(AscW("V"c)), CByte(AscW("E"c))})
            Write(New Byte() {CByte(AscW("f"c)), CByte(AscW("m"c)), CByte(AscW("t"c)), CByte(AscW(" "c))})
            Write(Int2ByteArr(WaveFormatSize))
            Write(Int2ByteArr(m_InputDataFormat.wFormatTag))
            Write(Int2ByteArr(m_InputDataFormat.nChannels))
            Write(Int2ByteArr(CUInt(m_InputDataFormat.nSamplesPerSec)))
            Write(Int2ByteArr(CUInt(m_InputDataFormat.nAvgBytesPerSec)))
            Write(Int2ByteArr(m_InputDataFormat.nBlockAlign))
            Write(Int2ByteArr(m_InputDataFormat.wBitsPerSample))
            Write(Int2ByteArr(m_InputDataFormat.cbSize))
            Write(New Byte() {CByte(AscW("d"c)), CByte(AscW("a"c)), CByte(AscW("t"c)), CByte(AscW("a"c))})
            Write(Int2ByteArr(m_AudioDataSize))
            m_WrittenBytes = CUInt(m_WrittenBytes - (WaveHeaderSize + 8))
        End Sub

        Public Overrides Sub Close()
            If Not closed Then
                If m_AudioDataSize = 0 Then
                    Seek(-CInt(m_WrittenBytes) - CInt(WaveHeaderSize) - 8, SeekOrigin.Current)
                    m_AudioDataSize = m_WrittenBytes
                    WriteWaveHeader()
                End If
            End If
            closed = True
            MyBase.Close()
        End Sub

        Public Overrides Sub Write(buffer As Byte(), index As Integer, count As Integer)
            MyBase.Write(buffer, index, count)
            m_WrittenBytes += CUInt(count)
        End Sub

        Public Overrides Sub Write(buffer As Byte())
            MyBase.Write(buffer)
            m_WrittenBytes += CUInt(buffer.Length)
        End Sub

        Protected Overrides Function GetOptimalBufferSize() As Integer
            Return CInt(m_InputDataFormat.nAvgBytesPerSec / 10)
        End Function
    End Class
End Namespace
