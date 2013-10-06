'nexENCODE Studio 5.0 Alpha 1.3
'October 6th, 2013
Option Explicit On
Option Strict On
Imports System
Imports System.Runtime.InteropServices

Namespace nexENCODE
    Friend Class clsWin32Functions
        Public Enum DriveTypes As UInteger
            DRIVE_UNKNOWN = 0
            DRIVE_NO_ROOT_DIR
            DRIVE_REMOVABLE
            DRIVE_FIXED
            DRIVE_REMOTE
            DRIVE_CDROM
            DRIVE_RAMDISK
        End Enum

        <System.Runtime.InteropServices.DllImport("Kernel32.dll")> Public Shared Function GetDriveType(drive As String) As DriveTypes
        End Function

        Public Const GENERIC_READ As UInteger = &H80000000UI
        Public Const GENERIC_WRITE As UInteger = &H40000000
        Public Const GENERIC_EXECUTE As UInteger = &H20000000
        Public Const GENERIC_ALL As UInteger = &H10000000
        Public Const FILE_SHARE_READ As UInteger = &H1
        Public Const FILE_SHARE_WRITE As UInteger = &H2
        Public Const FILE_SHARE_DELETE As UInteger = &H4
        Public Const CREATE_NEW As UInteger = 1
        Public Const CREATE_ALWAYS As UInteger = 2
        Public Const OPEN_EXISTING As UInteger = 3
        Public Const OPEN_ALWAYS As UInteger = 4
        Public Const TRUNCATE_EXISTING As UInteger = 5

        <System.Runtime.InteropServices.DllImport("Kernel32.dll", SetLastError:=True)> Public Shared Function CreateFile(FileName As String, DesiredAccess As UInteger, ShareMode As UInteger, lpSecurityAttributes As IntPtr, CreationDisposition As UInteger, dwFlagsAndAttributes As UInteger, hTemplateFile As IntPtr) As IntPtr
        End Function

        <System.Runtime.InteropServices.DllImport("Kernel32.dll", SetLastError:=True)> Public Shared Function CloseHandle(hObject As IntPtr) As Integer
        End Function

        Public Const IOCTL_CDROM_READ_TOC As UInteger = &H24000
        Public Const IOCTL_STORAGE_CHECK_VERIFY As UInteger = &H2D4800
        Public Const IOCTL_CDROM_RAW_READ As UInteger = &H2403E
        Public Const IOCTL_STORAGE_MEDIA_REMOVAL As UInteger = &H2D4804
        Public Const IOCTL_STORAGE_EJECT_MEDIA As UInteger = &H2D4808
        Public Const IOCTL_STORAGE_LOAD_MEDIA As UInteger = &H2D480C

        <System.Runtime.InteropServices.DllImport("Kernel32.dll", SetLastError:=True)> Public Shared Function DeviceIoControl(hDevice As IntPtr, IoControlCode As UInteger, lpInBuffer As IntPtr, InBufferSize As UInteger, lpOutBuffer As IntPtr, nOutBufferSize As UInteger, ByRef lpBytesReturned As UInteger, lpOverlapped As IntPtr) As Integer
        End Function

        <StructLayout(LayoutKind.Sequential)> Public Structure TRACK_DATA
            Public Reserved As Byte
            Private BitMapped As Byte

            Public Property Control() As Byte
                Get
                    Return CByte(BitMapped And &HF)
                End Get
                Set(value As Byte)
                    BitMapped = CByte((BitMapped And &HF0) Or (value And CByte(&HF)))
                End Set
            End Property

            Public Property Adr() As Byte
                Get
                    Return CByte((BitMapped And CByte(&HF0)) >> 4)
                End Get
                Set(value As Byte)
                    BitMapped = CByte((BitMapped And CByte(&HF)) Or (value << 4))
                End Set
            End Property

            Public TrackNumber As Byte
            Public Reserved1 As Byte
            Public Address_0 As Byte
            Public Address_1 As Byte
            Public Address_2 As Byte
            Public Address_3 As Byte
        End Structure

        Public Const MAXIMUM_NUMBER_TRACKS As Integer = 100

        <StructLayout(LayoutKind.Sequential)> Public Class TrackDataList
            <MarshalAs(UnmanagedType.ByValArray, SizeConst:=MAXIMUM_NUMBER_TRACKS * 8)> _
            Private Data As Byte()

            Default Public ReadOnly Property Item(Index As Integer) As TRACK_DATA
                Get
                    If (Index < 0) Or (Index >= MAXIMUM_NUMBER_TRACKS) Then
                        Throw New IndexOutOfRangeException()
                    End If
                    Dim res As TRACK_DATA
                    Dim handle As GCHandle = GCHandle.Alloc(Data, GCHandleType.Pinned)
                    Try
                        Dim buffer As IntPtr = handle.AddrOfPinnedObject()
                        buffer = New IntPtr(buffer.ToInt32() + (Index * Marshal.SizeOf(GetType(TRACK_DATA))))
                        res = DirectCast(Marshal.PtrToStructure(buffer, GetType(TRACK_DATA)), TRACK_DATA)
                    Finally
                        handle.Free()
                    End Try
                    Return res
                End Get
            End Property

            Public Sub New()
                Data = New Byte(MAXIMUM_NUMBER_TRACKS * Marshal.SizeOf(GetType(TRACK_DATA)) - 1) {}
            End Sub
        End Class

        <StructLayout(LayoutKind.Sequential)> Public Class CDROM_TOC
            Public Length As UShort
            Public FirstTrack As Byte = 0
            Public LastTrack As Byte = 0

            Public TrackData As TrackDataList

            Public Sub New()
                TrackData = New TrackDataList()
                Length = CUShort(Marshal.SizeOf(Me))
            End Sub
        End Class

        <StructLayout(LayoutKind.Sequential)> Public Class PREVENT_MEDIA_REMOVAL
            Public PreventMediaRemoval As Byte = 0
        End Class

        Public Enum TRACK_MODE_TYPE
            YellowMode2
            XAForm2
            CDDA
        End Enum

        <StructLayout(LayoutKind.Sequential)> Public Class RAW_READ_INFO
            Public DiskOffset As Long = 0
            Public SectorCount As UInteger = 0
            Public TrackMode As TRACK_MODE_TYPE = TRACK_MODE_TYPE.CDDA
        End Class

        <System.Runtime.InteropServices.DllImport("Kernel32.dll", SetLastError:=True)> Public Shared Function DeviceIoControl(hDevice As IntPtr, IoControlCode As UInteger, InBuffer As IntPtr, InBufferSize As UInteger, <Out()> OutTOC As CDROM_TOC, OutBufferSize As UInteger, ByRef BytesReturned As UInteger, Overlapped As IntPtr) As Integer
        End Function

        <System.Runtime.InteropServices.DllImport("Kernel32.dll", SetLastError:=True)> Public Shared Function DeviceIoControl(hDevice As IntPtr, IoControlCode As UInteger, <[In]()> InMediaRemoval As PREVENT_MEDIA_REMOVAL, InBufferSize As UInteger, OutBuffer As IntPtr, OutBufferSize As UInteger, ByRef BytesReturned As UInteger, Overlapped As IntPtr) As Integer
        End Function

        <System.Runtime.InteropServices.DllImport("Kernel32.dll", SetLastError:=True)> Public Shared Function DeviceIoControl(hDevice As IntPtr, IoControlCode As UInteger, <[In]()> rri As RAW_READ_INFO, InBufferSize As UInteger, <[In](), Out()> OutBuffer As Byte(), OutBufferSize As UInteger, ByRef BytesReturned As UInteger, Overlapped As IntPtr) As Integer
        End Function
    End Class
End Namespace