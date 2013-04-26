'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict Off
Imports System.Runtime.InteropServices
Imports System.Runtime.Serialization
Imports nexENCODE.WaveLib

Namespace LAME
    Public Enum VBRMETHOD As Integer
        VBR_METHOD_NONE = -1
        VBR_METHOD_DEFAULT = 0
        VBR_METHOD_OLD = 1
        VBR_METHOD_NEW = 2
        VBR_METHOD_MTRH = 3
        VBR_METHOD_ABR = 4
    End Enum

    Public Enum MpegMode As UInteger
        STEREO = 0
        JOINT_STEREO
        DUAL_CHANNEL
        ' LAME doesn't supports this! 
        MONO
        NOT_SET
        MAX_INDICATOR
        ' Don't use this! It's used for sanity checks. 
    End Enum

    Public Enum LAME_QUALITY_PRESET As Integer
        LQP_NOPRESET = -1
        ' QUALITY PRESETS
        LQP_NORMAL_QUALITY = 0
        LQP_LOW_QUALITY = 1
        LQP_HIGH_QUALITY = 2
        LQP_VOICE_QUALITY = 3
        LQP_R3MIX = 4
        LQP_VERYHIGH_QUALITY = 5
        LQP_STANDARD = 6
        LQP_FAST_STANDARD = 7
        LQP_EXTREME = 8
        LQP_FAST_EXTREME = 9
        LQP_INSANE = 10
        LQP_ABR = 11
        LQP_CBR = 12
        LQP_MEDIUM = 13
        LQP_FAST_MEDIUM = 14
        ' NEW PRESET VALUES
        LQP_PHONE = 1000
        LQP_SW = 2000
        LQP_AM = 3000
        LQP_FM = 4000
        LQP_VOICE = 5000
        LQP_RADIO = 6000
        LQP_TAPE = 7000
        LQP_HIFI = 8000
        LQP_CD = 9000
        LQP_STUDIO = 10000
    End Enum

    <StructLayout(LayoutKind.Sequential), Serializable()> _
    Public Structure MP3
        'BE_CONFIG_MP3
        Public dwSampleRate As UInteger
        ' 48000, 44100 and 32000 allowed
        Public byMode As Byte
        ' BE_MP3_MODE_STEREO, BE_MP3_MODE_DUALCHANNEL, BE_MP3_MODE_MONO
        Public wBitrate As UShort
        ' 32, 40, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256 and 320 allowed
        Public bPrivate As Integer
        Public bCRC As Integer
        Public bCopyright As Integer
        Public bOriginal As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential, Size:=327), Serializable()> _
    Public Structure LHV1
        ' BE_CONFIG_LAME LAME header version 1
        Public Const MPEG1 As UInteger = 1
        Public Const MPEG2 As UInteger = 0

        ' STRUCTURE INFORMATION
        Public dwStructVersion As UInteger
        Public dwStructSize As UInteger
        ' BASIC ENCODER SETTINGS
        Public dwSampleRate As UInteger
        ' SAMPLERATE OF INPUT FILE
        Public dwReSampleRate As UInteger
        ' DOWNSAMPLERATE, 0=ENCODER DECIDES  
        Public nMode As MpegMode
        ' STEREO, MONO
        Public dwBitrate As UInteger
        ' CBR bitrate, VBR min bitrate
        Public dwMaxBitrate As UInteger
        ' CBR ignored, VBR Max bitrate
        Public nPreset As LAME_QUALITY_PRESET
        ' Quality preset
        Public dwMpegVersion As UInteger
        ' MPEG-1 OR MPEG-2
        Public dwPsyModel As UInteger
        ' FUTURE USE, SET TO 0
        Public dwEmphasis As UInteger
        ' FUTURE USE, SET TO 0
        ' BIT STREAM SETTINGS
        Public bPrivate As Integer
        ' Set Private Bit (TRUE/FALSE)
        Public bCRC As Integer
        ' Insert CRC (TRUE/FALSE)
        Public bCopyright As Integer
        ' Set Copyright Bit (TRUE/FALSE)
        Public bOriginal As Integer
        ' Set Original Bit (TRUE/FALSE)
        ' VBR STUFF
        Public bWriteVBRHeader As Integer
        ' WRITE XING VBR HEADER (TRUE/FALSE)
        Public bEnableVBR As Integer
        ' USE VBR ENCODING (TRUE/FALSE)
        Public nVBRQuality As Integer
        ' VBR QUALITY 0..9
        Public dwVbrAbr_bps As UInteger
        ' Use ABR in stead of nVBRQuality
        Public nVbrMethod As VBRMETHOD
        Public bNoRes As Integer
        ' Disable Bit resorvoir (TRUE/FALSE)
        ' MISC SETTINGS
        Public bStrictIso As Integer
        ' Use strict ISO encoding rules (TRUE/FALSE)
        Public nQuality As UShort
        ' Quality Setting, HIGH BYTE should be NOT LOW byte, otherwhise quality=5
        ' FUTURE USE, SET TO 0, align strucutre to 331 bytes
        '[ MarshalAs( UnmanagedType.ByValArray, SizeConst=255-4*4-2 )]
        'public byte[]   btReserved;//[255-4*sizeof(DWORD) - sizeof( WORD )];
        Public Sub New(format As WaveFormat, MpeBitRate As UInteger)
            If format.wFormatTag <> CShort(WaveFormats.Pcm) Then
                Throw New ArgumentOutOfRangeException("format", "Only PCM format supported")
            End If
            If format.wBitsPerSample <> 16 Then
                Throw New ArgumentOutOfRangeException("format", "Only 16 bits samples supported")
            End If
            dwStructVersion = 1
            dwStructSize = CUInt(Marshal.SizeOf(GetType(BE_CONFIG)))
            Select Case format.nSamplesPerSec
                Case 16000, 22050, 24000
                    dwMpegVersion = MPEG2
                    Exit Select
                Case 32000, 44100, 48000
                    dwMpegVersion = MPEG1
                    Exit Select
                Case Else
                    Throw New ArgumentOutOfRangeException("format", "Unsupported sample rate")
            End Select
            dwSampleRate = CUInt(format.nSamplesPerSec)
            ' INPUT FREQUENCY
            dwReSampleRate = 0
            ' DON'T RESAMPLE
            Select Case format.nChannels
                Case 1
                    nMode = MpegMode.MONO
                    Exit Select
                Case 2
                    nMode = MpegMode.STEREO
                    Exit Select
                Case Else
                    Throw New ArgumentOutOfRangeException("format", "Invalid number of channels")
            End Select
            Select Case MpeBitRate
                Case 32, 40, 48, 56, 64, 80, _
                 96, 112, 128, 160
                    'Allowed bit rates in MPEG1 and MPEG2
                    Exit Select
                Case 192, 224, 256, 320
                    'Allowed only in MPEG1
                    If dwMpegVersion <> MPEG1 Then
                        Throw New ArgumentOutOfRangeException("MpsBitRate", "Bit rate not compatible with input format")
                    End If
                    Exit Select
                Case 8, 16, 24, 144
                    'Allowed only in MPEG2
                    If dwMpegVersion <> MPEG2 Then
                        Throw New ArgumentOutOfRangeException("MpsBitRate", "Bit rate not compatible with input format")
                    End If
                    Exit Select
                Case Else
                    Throw New ArgumentOutOfRangeException("MpsBitRate", "Unsupported bit rate")
            End Select
            dwBitrate = MpeBitRate
            ' MINIMUM BIT RATE
            nPreset = LAME_QUALITY_PRESET.LQP_NORMAL_QUALITY
            ' QUALITY PRESET SETTING
            dwPsyModel = 0
            ' USE DEFAULT PSYCHOACOUSTIC MODEL 
            dwEmphasis = 0
            ' NO EMPHASIS TURNED ON
            bOriginal = 1
            ' SET ORIGINAL FLAG
            bWriteVBRHeader = 0
            bNoRes = 0
            ' No Bit resorvoir
            bCopyright = 0
            bCRC = 0
            bEnableVBR = 0
            bPrivate = 0
            bStrictIso = 0
            dwMaxBitrate = 0
            dwVbrAbr_bps = 0
            nQuality = 0
            nVbrMethod = VBRMETHOD.VBR_METHOD_NONE
            nVBRQuality = 0
        End Sub
    End Structure


    <StructLayout(LayoutKind.Sequential), Serializable()> Public Structure ACC
        Public dwSampleRate As UInteger
        Public byMode As Byte
        Public wBitrate As UShort
        Public byEncodingMethod As Byte
    End Structure

    <StructLayout(LayoutKind.Explicit), Serializable()> Public Class Format
        <FieldOffset(0)> Public mp3 As MP3
        <FieldOffset(0)> Public lhv1 As LHV1
        <FieldOffset(0)> Public acc As ACC

        Public Sub New(format__1 As WaveFormat, MpeBitRate As UInteger)
            lhv1 = New LHV1(format__1, MpeBitRate)
        End Sub
    End Class

    <StructLayout(LayoutKind.Sequential), Serializable()> Public Class BE_CONFIG
        ' encoding formats
        Public Const BE_CONFIG_MP3 As UInteger = 0
        Public Const BE_CONFIG_LAME As UInteger = 256
        Public dwConfig As UInteger
        Public format As Format

        Public Sub New(format As WaveFormat, MpeBitRate As UInteger)
            Me.dwConfig = BE_CONFIG_LAME
            Me.format = New Format(format, MpeBitRate)
        End Sub

        Public Sub New(format As WaveFormat)
            Me.New(format, 128)
        End Sub
    End Class

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Ansi)> Public Class BE_VERSION
        Public Const BE_MAX_HOMEPAGE As UInteger = 256
        Public byDLLMajorVersion As Byte
        Public byDLLMinorVersion As Byte
        Public byMajorVersion As Byte
        Public byMinorVersion As Byte
        ' DLL Release date
        Public byDay As Byte
        Public byMonth As Byte
        Public wYear As UShort
        'Homepage URL
        'BE_MAX_HOMEPAGE+1
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=257)> Public zHomepage As String
        Public byAlphaLevel As Byte
        Public byBetaLevel As Byte
        Public byMMXEnabled As Byte
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=125)> Public btReserved As Byte()

        Public Sub New()
            btReserved = New Byte(124) {}
        End Sub
    End Class

    Public Class Lame_encDll
        Public Const BE_ERR_SUCCESSFUL As UInteger = 0
        Public Const BE_ERR_INVALID_FORMAT As UInteger = 1
        Public Const BE_ERR_INVALID_FORMAT_PARAMETERS As UInteger = 2
        Public Const BE_ERR_NO_MORE_HANDLES As UInteger = 3
        Public Const BE_ERR_INVALID_HANDLE As UInteger = 4

        <DllImport("Lame_enc.dll")> _
        Public Shared Function beInitStream(pbeConfig As BE_CONFIG, ByRef dwSamples As UInteger, ByRef dwBufferSize As UInteger, ByRef phbeStream As UInteger) As UInteger
        End Function

        <DllImport("Lame_enc.dll")> _
        Public Shared Function beEncodeChunk(hbeStream As UInteger, nSamples As UInteger, pInSamples As Short(), <[In](), Out()> pOutput As Byte(), ByRef pdwOutput As UInteger) As UInteger
        End Function

        <DllImport("Lame_enc.dll")> _
        Protected Shared Function beEncodeChunk(hbeStream As UInteger, nSamples As UInteger, pSamples As IntPtr, <[In](), Out()> pOutput As Byte(), ByRef pdwOutput As UInteger) As UInteger
        End Function


        Public Shared Function EncodeChunk(hbeStream As UInteger, buffer As Byte(), index As Integer, nBytes As UInteger, pOutput As Byte(), ByRef pdwOutput As UInteger) As UInteger
            Dim res As UInteger
            Dim handle As GCHandle = GCHandle.Alloc(buffer, GCHandleType.Pinned)
            Try
                Dim ptr As IntPtr = CType(handle.AddrOfPinnedObject().ToInt32() + index, IntPtr)
                res = beEncodeChunk(hbeStream, nBytes \ 2, ptr, pOutput, pdwOutput)
            Finally
                handle.Free()
            End Try
            Return res
        End Function

        Public Shared Function EncodeChunk(hbeStream As UInteger, buffer As Byte(), pOutput As Byte(), ByRef pdwOutput As UInteger) As UInteger
            Return EncodeChunk(hbeStream, buffer, 0, CUInt(buffer.Length), pOutput, pdwOutput)
        End Function

        <DllImport("Lame_enc.dll")> _
        Public Shared Function beDeinitStream(hbeStream As UInteger, <[In](), Out()> pOutput As Byte(), ByRef pdwOutput As UInteger) As UInteger
        End Function

        <DllImport("Lame_enc.dll")> _
        Public Shared Function beCloseStream(hbeStream As UInteger) As UInteger
        End Function

        <DllImport("Lame_enc.dll")> _
        Public Shared Sub beVersion(<Out()> pbeVersion As BE_VERSION)
        End Sub

        <DllImport("Lame_enc.dll", CharSet:=CharSet.Ansi)> Public Shared Sub beWriteVBRHeader(pszMP3FileName As String)
        End Sub

        <DllImport("Lame_enc.dll")> Public Shared Function beEncodeChunkFloatS16NI(hbeStream As UInteger, nSamples As UInteger, <[In]()> buffer_l As Single(), <[In]()> buffer_r As Single(), <[In](), Out()> pOutput As Byte(), ByRef pdwOutput As UInteger) As UInteger
        End Function

        <DllImport("Lame_enc.dll")> Public Shared Function beFlushNoGap(hbeStream As UInteger, <[In](), Out()> pOutput As Byte(), ByRef pdwOutput As UInteger) As UInteger
        End Function

        <DllImport("Lame_enc.dll", CharSet:=CharSet.Ansi)> Public Shared Function beWriteInfoTag(hbeStream As UInteger, lpszFileName As String) As UInteger
        End Function
    End Class
End Namespace