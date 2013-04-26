'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Imports System.Collections.Generic
Imports System.Text

Namespace NAudio.CoreAudioApi.Interfaces
    <Flags()> _
    Enum ClsCtx
        INPROC_SERVER = &H1
        INPROC_HANDLER = &H2
        LOCAL_SERVER = &H4
        INPROC_SERVER16 = &H8
        REMOTE_SERVER = &H10
        INPROC_HANDLER16 = &H20
        'RESERVED1	= 0x40,
        'RESERVED2	= 0x80,
        'RESERVED3	= 0x100,
        'RESERVED4	= 0x200,
        NO_CODE_DOWNLOAD = &H400
        'RESERVED5	= 0x800,
        NO_CUSTOM_MARSHAL = &H1000
        ENABLE_CODE_DOWNLOAD = &H2000
        NO_FAILURE_LOG = &H4000
        DISABLE_AAA = &H8000
        ENABLE_AAA = &H10000
        FROM_DEFAULT_CONTEXT = &H20000
        ACTIVATE_32_BIT_SERVER = &H40000
        ACTIVATE_64_BIT_SERVER = &H80000
        ENABLE_CLOAKING = &H100000
        PS_DLL = CInt(&H80000000UI)
        INPROC = INPROC_SERVER Or INPROC_HANDLER
        SERVER = INPROC_SERVER Or LOCAL_SERVER Or REMOTE_SERVER
        ALL = SERVER Or INPROC_HANDLER
    End Enum
End Namespace