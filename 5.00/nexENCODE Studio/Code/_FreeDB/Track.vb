Namespace Freedb
    ''' <summary>
    ''' Summary description for Track.
    ''' </summary>
    Public Class Track

        Private m_Title As String
        Private m_ExtendedData As String

#Region "Public Properties"
        ''' <summary>
        ''' Property ExtendedData (string)
        ''' </summary>
        Public Property ExtendedData() As String
            Get
                Return Me.m_ExtendedData
            End Get
            Set(value As String)
                Me.m_ExtendedData = value
            End Set
        End Property

        ''' <summary>
        ''' Property Title (string)
        ''' </summary>
        Public Property Title() As String
            Get
                Return Me.m_Title
            End Get
            Set(value As String)
                Me.m_Title = value
            End Set
        End Property
#End Region

        ''' <summary>
        ''' Create an instance of a Track 
        ''' </summary>
        ''' <param name="title"></param>
        Public Sub New()
        End Sub


        ''' <summary>
        ''' Create an instance of a Track passing in a title
        ''' </summary>
        ''' <param name="title"></param>
        Public Sub New(title As String)
            m_Title = title
        End Sub

        ''' <summary>
        ''' Create an instance of a Track passing in a title and extended data
        ''' </summary>
        ''' <param name="title"></param>
        Public Sub New(title As String, extendedData As String)
            m_Title = title
            m_ExtendedData = extendedData
        End Sub
    End Class
End Namespace