'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Imports System.Runtime.Serialization
Imports nexENCODE.Media
Imports nexENCODE.WaveLib

Namespace Media.Mp3
    <Serializable()> Public Class Mp3WriterConfig
        Inherits Media.AudioWriterConfig
        Private m_BeConfig As LAME.BE_CONFIG

        Protected Sub New(info As SerializationInfo, context As StreamingContext)
            MyBase.New(info, context)
            m_BeConfig = DirectCast(info.GetValue("BE_CONFIG", GetType(LAME.BE_CONFIG)), LAME.BE_CONFIG)
        End Sub

        Public Sub New(InFormat As WaveFormat, beconfig As LAME.BE_CONFIG)
            MyBase.New(InFormat)
            m_BeConfig = beconfig
        End Sub

        Public Sub New(InFormat As WaveFormat)
            Me.New(InFormat, New LAME.BE_CONFIG(InFormat))
        End Sub

        Public Sub New()
            Me.New(New WaveLib.WaveFormat(44100, 16, 2))
        End Sub

        Public Overrides Sub GetObjectData(info As System.Runtime.Serialization.SerializationInfo, context As System.Runtime.Serialization.StreamingContext)
            MyBase.GetObjectData(info, context)
            info.AddValue("BE_CONFIG", m_BeConfig, m_BeConfig.[GetType]())
        End Sub

        Public Property Mp3Config() As LAME.BE_CONFIG
            Get
                Return m_BeConfig
            End Get
            Set(value As LAME.BE_CONFIG)
                m_BeConfig = value
            End Set
        End Property
    End Class
End Namespace
