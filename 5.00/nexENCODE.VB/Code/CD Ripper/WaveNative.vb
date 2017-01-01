'nexENCODE Studio 5.0 Alpha 1.2
'December 30th, 2011
Option Explicit On
Option Strict On
Imports System
Imports System.Runtime.InteropServices

Namespace WaveLib
    Public Enum WaveFormats
        Pcm = 1
        Float = 3
    End Enum

    <StructLayout(LayoutKind.Sequential)> Public Class WaveFormat
        Public wFormatTag As Short
        Public nChannels As Short
        Public nSamplesPerSec As Integer
        Public nAvgBytesPerSec As Integer
        Public nBlockAlign As Short
        Public wBitsPerSample As Short
        Public cbSize As Short

        Public Sub New(rate As Integer, bits As Integer, channels As Integer)
            wFormatTag = CShort(WaveFormats.Pcm)
            nChannels = CShort(channels)
            nSamplesPerSec = rate
            wBitsPerSample = CShort(bits)
            cbSize = 0
            nBlockAlign = CShort(channels * (bits / 8))
            nAvgBytesPerSec = nSamplesPerSec * nBlockAlign
        End Sub
    End Class

    Friend Class WaveNative
        Public Const MMSYSERR_NOERROR As Integer = 0
        Public Const MM_WOM_OPEN As Integer = &H3BB
        Public Const MM_WOM_CLOSE As Integer = &H3BC
        Public Const MM_WOM_DONE As Integer = &H3BD
        Public Const CALLBACK_FUNCTION As Integer = &H30000
        Public Const TIME_MS As Integer = &H1
        Public Const TIME_SAMPLES As Integer = &H2
        Public Const TIME_BYTES As Integer = &H4
        Public Delegate Sub WaveDelegate(hdrvr As IntPtr, uMsg As Integer, dwUser As Integer, ByRef wavhdr As WaveHdr, dwParam2 As Integer)

        <StructLayout(LayoutKind.Sequential)> Public Structure WaveHdr
            Public lpData As IntPtr
            Public dwBufferLength As Integer
            Public dwBytesRecorded As Integer
            Public dwUser As IntPtr
            Public dwFlags As Integer
            Public dwLoops As Integer
            Public lpNext As IntPtr
            Public reserved As Integer
        End Structure

        Private Const mmdll As String = "winmm.dll"

        <DllImport(mmdll)> _
        Public Shared Function waveOutGetNumDevs() As Integer
        End Function
        <DllImport(mmdll)> _
        Public Shared Function waveOutPrepareHeader(hWaveOut As IntPtr, ByRef lpWaveOutHdr As WaveHdr, uSize As Integer) As Integer
        End Function
        <DllImport(mmdll)> _
        Public Shared Function waveOutUnprepareHeader(hWaveOut As IntPtr, ByRef lpWaveOutHdr As WaveHdr, uSize As Integer) As Integer
        End Function
        <DllImport(mmdll)> _
        Public Shared Function waveOutWrite(hWaveOut As IntPtr, ByRef lpWaveOutHdr As WaveHdr, uSize As Integer) As Integer
        End Function
        <DllImport(mmdll)> _
        Public Shared Function waveOutOpen(hWaveOut As IntPtr, uDeviceID As Integer, lpFormat As WaveFormat, dwCallback As WaveDelegate, dwInstance As Integer, dwFlags As Integer) As Integer
        End Function
        <DllImport(mmdll)> _
        Public Shared Function waveOutReset(hWaveOut As IntPtr) As Integer
        End Function
        <DllImport(mmdll)> _
        Public Shared Function waveOutClose(hWaveOut As IntPtr) As Integer
        End Function
        <DllImport(mmdll)> _
        Public Shared Function waveOutPause(hWaveOut As IntPtr) As Integer
        End Function
        <DllImport(mmdll)> _
        Public Shared Function waveOutRestart(hWaveOut As IntPtr) As Integer
        End Function
        <DllImport(mmdll)> _
        Public Shared Function waveOutGetPosition(hWaveOut As IntPtr, lpInfo As Integer, uSize As Integer) As Integer
        End Function
        <DllImport(mmdll)> _
        Public Shared Function waveOutSetVolume(hWaveOut As IntPtr, dwVolume As Integer) As Integer
        End Function
        <DllImport(mmdll)> _
        Public Shared Function waveOutGetVolume(hWaveOut As IntPtr, dwVolume As Integer) As Integer
        End Function
    End Class
End Namespace