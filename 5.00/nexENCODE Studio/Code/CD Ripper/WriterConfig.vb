'nexENCODE Studio 5.0 Alpha 1.2
'December 30th, 2011
Option Explicit On
Option Strict On
Imports System
Imports System.IO
Imports System.Runtime.Serialization
Imports System.Windows.Forms
Imports WaveLib
Imports nexENCODE.WaveLib

Namespace Yeti.MMedia
    <Serializable()> Public Class AudioWriterConfig
        Protected m_Format As WaveFormat

        Protected Sub New(info As SerializationInfo, context As StreamingContext)
            Dim rate As Integer = info.GetInt32("Format.Rate")
            Dim bits As Integer = info.GetInt32("Format.Bits")
            Dim channels As Integer = info.GetInt32("Format.Channels")
            m_Format = New WaveFormat(rate, bits, channels)
        End Sub

        Public Sub New(f As WaveFormat)
            m_Format = New WaveFormat(f.nSamplesPerSec, f.wBitsPerSample, f.nChannels)
        End Sub

        Public Sub New()
            Me.New(New WaveFormat(44100, 16, 2))
        End Sub

        Public Property Format() As WaveFormat
            Get
                Return m_Format
            End Get
            Set(value As WaveFormat)
                m_Format = value
            End Set
        End Property

#Region "ISerializable Members"
        Public Overridable Sub GetObjectData(info As SerializationInfo, context As StreamingContext)
            info.AddValue("Format.Rate", m_Format.nSamplesPerSec)
            info.AddValue("Format.Bits", m_Format.wBitsPerSample)
            info.AddValue("Format.Channels", m_Format.nChannels)
        End Sub
#End Region
    End Class

    Public Interface IConfigControl
        Sub DoApply()
        Sub DoSetInitialValues()
        ReadOnly Property ConfigControl() As Control
        ReadOnly Property ControlName() As String
        Event ConfigChange As EventHandler
    End Interface

    Public Interface IEditAudioWriterConfig
        Inherits IConfigControl
        Property Config() As AudioWriterConfig
    End Interface

    Public Interface IEditFormat
        Inherits IConfigControl
        Property Format() As WaveFormat
    End Interface
End Namespace