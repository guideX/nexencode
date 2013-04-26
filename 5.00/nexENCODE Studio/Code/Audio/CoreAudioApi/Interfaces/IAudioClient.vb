'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Imports System.Runtime.InteropServices
Imports NAudio.Wave

Namespace NAudio.CoreAudioApi.Interfaces
    <Guid("1CB9AD4C-DBFA-4c32-B178-C2F568A703B2"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)> _
    Friend Interface IAudioClient
        <PreserveSig()> Function Initialize(shareMode As AudioClientShareMode, StreamFlags As AudioClientStreamFlags, hnsBufferDuration As Long, hnsPeriodicity As Long, <[In]()> pFormat As WaveFormat, <[In]()> ByRef AudioSessionGuid As Guid) As Integer
        Function GetBufferSize(ByRef bufferSize As UInteger) As Integer
        Function GetStreamLatency() As <MarshalAs(UnmanagedType.I8)> Long
        Function GetCurrentPadding(ByRef currentPadding As Integer) As Integer
        <PreserveSig()> Function IsFormatSupported(shareMode As AudioClientShareMode, <[In]()> pFormat As WaveFormat, <Out(), MarshalAs(UnmanagedType.LPStruct)> ByRef closestMatchFormat As WaveFormatExtensible) As Integer
        Function GetMixFormat(ByRef deviceFormatPointer As IntPtr) As Integer
        Function GetDevicePeriod(ByRef defaultDevicePeriod As Long, ByRef minimumDevicePeriod As Long) As Integer
        Function Start() As Integer
        Function [Stop]() As Integer
        Function Reset() As Integer
        Function SetEventHandle(eventHandle As IntPtr) As Integer
        Function GetService(ByRef interfaceId As Guid, <MarshalAs(UnmanagedType.IUnknown)> ByRef interfacePointer As Object) As Integer
    End Interface
End Namespace