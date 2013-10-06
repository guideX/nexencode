'nexENCODE Studio 5.0 Alpha 1.3
'October 6th, 2013
Option Explicit On
Option Strict On
Imports System
Imports System.Runtime.InteropServices

Namespace nexENCODE.CDDriveEvents
    Public Delegate Sub CdDataReadEventHandler(sender As Object, ea As DataReadEventArgs)
    Public Delegate Sub CdReadProgressEventHandler(sender As Object, ea As ReadProgressEventArgs)
    Friend Delegate Sub DeviceChangeEventHandler(sender As Object, ea As DeviceChangeEventArgs)

    Public Class DataReadEventArgs
        Inherits EventArgs
        Public Event ProcessError(lError As String, lSub As String)
        Private m_Data As Byte()
        Private m_DataSize As UInteger

        Public Sub New(_Data As Byte(), _Size As UInteger)
            Try
                m_Data = _Data
                m_DataSize = _Size
            Catch ex As Exception
                RaiseEvent ProcessError(ex.Message, "Public Sub New(_Data As Byte(), _Size As UInteger)")
            End Try
        End Sub

        Public ReadOnly Property Data() As Byte()
            Get
                Try
                    Return m_Data
                Catch ex As Exception
                    RaiseEvent ProcessError(ex.Message, "Public ReadOnly Property Data() As Byte()")
                    Return Nothing
                End Try
            End Get
        End Property

        Public ReadOnly Property DataSize() As UInteger
            Get
                Try
                    Return m_DataSize
                Catch ex As Exception
                    RaiseEvent ProcessError(ex.Message, "Public ReadOnly Property DataSize() As UInteger")
                    Return Nothing
                End Try
            End Get
        End Property
    End Class

    Public Class ReadProgressEventArgs
        Inherits EventArgs
        Public Event ProcessError(lError As String, lSub As String)
        Private m_Bytes2Read As UInteger
        Private m_BytesRead As UInteger
        Private m_CancelRead As Boolean = False

        Public Sub New(bytes2read As UInteger, bytesread As UInteger)
            m_Bytes2Read = bytes2read
            m_BytesRead = bytesread
        End Sub

        Public ReadOnly Property Bytes2Read() As UInteger
            Get
                Return m_Bytes2Read
            End Get
        End Property

        Public ReadOnly Property BytesRead() As UInteger
            Get
                Return m_BytesRead
            End Get
        End Property

        Public Property CancelRead() As Boolean
            Get
                Return m_CancelRead
            End Get
            Set(value As Boolean)
                m_CancelRead = value
            End Set
        End Property
    End Class

    Friend Enum DeviceChangeEventType
        DeviceInserted
        DeviceRemoved
    End Enum

    Friend Class DeviceChangeEventArgs
        Inherits EventArgs
        Private m_Type As DeviceChangeEventType
        Private m_Drive As Char

        Public Sub New(drive As Char, type As DeviceChangeEventType)
            m_Drive = drive
            m_Type = type
        End Sub

        Public ReadOnly Property Drive() As Char
            Get
                Return m_Drive
            End Get
        End Property

        Public ReadOnly Property ChangeType() As DeviceChangeEventType
            Get
                Return m_Type
            End Get
        End Property
    End Class

    Friend Enum DeviceType As UInteger
        DBT_DEVTYP_OEM = &H0
        DBT_DEVTYP_DEVNODE = &H1
        DBT_DEVTYP_VOLUME = &H2
        DBT_DEVTYP_PORT = &H3
        DBT_DEVTYP_NET = &H4
    End Enum

    Friend Enum VolumeChangeFlags As UShort
        DBTF_MEDIA = &H1
        DBTF_NET = &H2
    End Enum

    <StructLayout(LayoutKind.Sequential)> Friend Structure DEV_BROADCAST_HDR
        Public dbch_size As UInteger
        Public dbch_devicetype As DeviceType
        Private dbch_reserved As UInteger
    End Structure

    <StructLayout(LayoutKind.Sequential)> Friend Structure DEV_BROADCAST_VOLUME
        Public dbcv_size As UInteger
        Public dbcv_devicetype As DeviceType
        Private dbcv_reserved As UInteger
        Private dbcv_unitmask As UInteger

        Public ReadOnly Property Drives() As Char()
            Get
                Dim drvs As String = ""
                'For c As Char = "A"c To "Z"c
                For Each c In "DEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray()
                    'If (dbcv_unitmask And (1 << (c - "A"c))) <> 0 Then
                    If CBool(dbcv_unitmask) Then
                        drvs += c
                    End If
                Next
                Return drvs.ToCharArray()
            End Get
        End Property

        Public dbcv_flags As VolumeChangeFlags
    End Structure

    Friend Class DeviceChangeNotificationWindow
        Inherits NativeWindow
        Public Event DeviceChange As DeviceChangeEventHandler
        Const WS_EX_TOOLWINDOW As Integer = &H80
        'Const WS_POPUP As Integer = CInt(&H80000000UI)
        Const WS_POPUP As Integer = &H80000000
        Const WM_DEVICECHANGE As Integer = &H219
        Const DBT_APPYBEGIN As Integer = &H0
        Const DBT_APPYEND As Integer = &H1
        Const DBT_DEVNODES_CHANGED As Integer = &H7
        Const DBT_QUERYCHANGECONFIG As Integer = &H17
        Const DBT_CONFIGCHANGED As Integer = &H18
        Const DBT_CONFIGCHANGECANCELED As Integer = &H19
        Const DBT_MONITORCHANGE As Integer = &H1B
        Const DBT_SHELLLOGGEDON As Integer = &H20
        Const DBT_CONFIGMGAPI32 As Integer = &H22
        Const DBT_VXDINITCOMPLETE As Integer = &H23
        Const DBT_VOLLOCKQUERYLOCK As Integer = &H8041
        Const DBT_VOLLOCKLOCKTAKEN As Integer = &H8042
        Const DBT_VOLLOCKLOCKFAILED As Integer = &H8043
        Const DBT_VOLLOCKQUERYUNLOCK As Integer = &H8044
        Const DBT_VOLLOCKLOCKRELEASED As Integer = &H8045
        Const DBT_VOLLOCKUNLOCKFAILED As Integer = &H8046
        Const DBT_DEVICEARRIVAL As Integer = &H8000
        Const DBT_DEVICEQUERYREMOVE As Integer = &H8001
        Const DBT_DEVICEQUERYREMOVEFAILED As Integer = &H8002
        Const DBT_DEVICEREMOVEPENDING As Integer = &H8003
        Const DBT_DEVICEREMOVECOMPLETE As Integer = &H8004
        Const DBT_DEVICETYPESPECIFIC As Integer = &H8005

        Public Sub New()
            Dim Params As New CreateParams()
            Params.ExStyle = WS_EX_TOOLWINDOW
            Params.Style = WS_POPUP
            CreateHandle(Params)
        End Sub

        Private Sub OnCDChange(ea As DeviceChangeEventArgs)
            RaiseEvent DeviceChange(Me, ea)
        End Sub

        'Private Sub OnDeviceChange(DevDesc As DEV_BROADCAST_VOLUME, EventType As DeviceChangeEventType)
        'If DeviceChange IsNot Nothing Then
        'For Each ch As Char In DevDesc.Drives
        'Dim a As New DeviceChangeEventArgs(ch, EventType)
        'DeviceChange(Me, a)
        'Next
        'End If
        'End Sub

        Protected Overrides Sub WndProc(ByRef m As Message)
            If m.Msg = WM_DEVICECHANGE Then
                Dim head As DEV_BROADCAST_HDR
                Select Case m.WParam.ToInt32()
                    'case DBT_DEVNODES_CHANGED :
                    'case DBT_CONFIGCHANGED :
                    Case DBT_DEVICEARRIVAL
                        head = DirectCast(Marshal.PtrToStructure(m.LParam, GetType(DEV_BROADCAST_HDR)), DEV_BROADCAST_HDR)
                        If head.dbch_devicetype = DeviceType.DBT_DEVTYP_VOLUME Then
                            Dim DevDesc As DEV_BROADCAST_VOLUME = DirectCast(Marshal.PtrToStructure(m.LParam, GetType(DEV_BROADCAST_VOLUME)), DEV_BROADCAST_VOLUME)
                            If DevDesc.dbcv_flags = VolumeChangeFlags.DBTF_MEDIA Then
                                'OnDeviceChange(DevDesc, DeviceChangeEventType.DeviceInserted)
                            End If
                        End If
                        Exit Select
                        'case DBT_DEVICEQUERYREMOVE :
                        'case DBT_DEVICEQUERYREMOVEFAILED :
                        'case DBT_DEVICEREMOVEPENDING :
                    Case DBT_DEVICEREMOVECOMPLETE
                        head = DirectCast(Marshal.PtrToStructure(m.LParam, GetType(DEV_BROADCAST_HDR)), DEV_BROADCAST_HDR)
                        If head.dbch_devicetype = DeviceType.DBT_DEVTYP_VOLUME Then
                            Dim DevDesc As DEV_BROADCAST_VOLUME = DirectCast(Marshal.PtrToStructure(m.LParam, GetType(DEV_BROADCAST_VOLUME)), DEV_BROADCAST_VOLUME)
                            If DevDesc.dbcv_flags = VolumeChangeFlags.DBTF_MEDIA Then
                                'OnDeviceChange(DevDesc, DeviceChangeEventType.DeviceRemoved)
                            End If
                        End If
                        Exit Select
                        'case DBT_DEVICETYPESPECIFIC :
                End Select
            End If
            MyBase.WndProc(m)
        End Sub
    End Class
End Namespace