Namespace Freedb
    ''' <summary>
    ''' Summary description for Site.
    ''' </summary>
    Public Class Site
        Private m_SiteAddress As String
        Private m_Protocol As String
        Private m_AdditionalAddressInfo As String
        Private m_Port As String
        Private m_Latitude As String
        Private m_Longitude As String
        Private m_Description As String


        Public Class PROTOCOLS
            Public Const HTTP As String = "http"
            Public Const CDDBP As String = "cddbp"
            Public Const ALL As String = "all"
        End Class


        ''' <summary>
        ''' Property AdditionalAddressInfo (string)
        ''' Any additional addressing information needed to access the server.
        ''' For example, for HTTP protocol servers, this would be the path to the CCDB server CGI script.
        ''' This field will be "-" if no additional addressing information is needed.
        ''' </summary>
        Public Property AdditionalAddressInfo() As String
            Get
                Return Me.m_AdditionalAddressInfo
            End Get
            Set(value As String)
                Me.m_AdditionalAddressInfo = value
            End Set
        End Property


#Region "Public Properties"

        ''' <summary>
        ''' Property Site (string) - Internet address of the remote site 
        ''' </summary>
        Public Property SiteAddress() As String
            Get
                Return Me.m_SiteAddress
            End Get
            Set(value As String)
                Me.m_SiteAddress = value
            End Set
        End Property

        ''' <summary>
        ''' Property Protocol (string)
        ''' The transfer protocol used to access the site
        ''' </summary>
        Public Property Protocol() As String
            Get
                Return Me.m_Protocol
            End Get
            Set(value As String)
                Me.m_Protocol = value
            End Set
        End Property

        ''' <summary>
        ''' Property Port (string)- The port at which the server resides on that site.
        ''' </summary>
        Public Property Port() As String
            Get
                Return Me.m_Port
            End Get
            Set(value As String)
                Me.m_Port = value
            End Set
        End Property



        ''' <summary>
        ''' Property Description (string)
        ''' A short description of the geographical location of the site.
        ''' </summary>
        Public Property Description() As String
            Get
                Return Me.m_Description
            End Get
            Set(value As String)
                Me.m_Description = value
            End Set
        End Property


        ''' <summary>
        ''' Property Latitude (string)
        ''' The latitude of the server site. The format is as follows:
        ''' CDDD.MM
        ''' Where "C" is the compass direction (N, S), "DDD" is the
        ''' degrees, and "MM" is the minutes.
        ''' </summary>
        Public Property Latitude() As String
            Get
                Return Me.m_Latitude
            End Get
            Set(value As String)
                Me.m_Latitude = value
            End Set
        End Property

        ''' <summary>
        ''' Property Longitude (string)
        ''' The longitude of the server site. Format is as above, except
        ''' the compass direction must be one of (E, W).
        ''' </summary>
        Public Property Longitude() As String
            Get
                Return Me.m_Longitude
            End Get
            Set(value As String)
                Me.m_Longitude = value
            End Set
        End Property


#End Region


        Public Sub New(siteFromCDDB As String)
            If Not Parse(siteFromCDDB) Then
                Throw New Exception(Convert.ToString("Unable to Parse Site. Input: ") & siteFromCDDB)
            End If
        End Sub


        ''' <summary>
        ''' Builds a site from an address, protocol and addition info
        ''' </summary>
        ''' <param name="siteAddress"></param>
        ''' <param name="protocol"></param>
        ''' <param name="additionAddressInfo"></param>
        Public Sub New(siteAddress As String, protocol As String, additionAddressInfo As String)
            m_SiteAddress = siteAddress
            m_Protocol = protocol
            m_AdditionalAddressInfo = additionAddressInfo
        End Sub



        Public Function Parse(siteAsString As String) As Boolean
            siteAsString.Trim()
            Dim values As String() = siteAsString.Split(" "c)
            If values.Length < 5 Then
                Return False
            End If
            m_SiteAddress = values(0)
            Me.m_Protocol = values(1)
            m_Port = values(2)
            If values(3).Trim() <> "-" Then
                m_AdditionalAddressInfo = values(3)
            End If
            m_Latitude = values(4)
            m_Longitude = values(5)

            ' description could be split over many because it could have spaces
            For i As Integer = 6 To values.Length - 1
                m_Description += values(i)

                m_Description += " "
            Next
            m_Description.Trim()
            Return True
        End Function

        Public Function GetUrl() As String

            If Me.m_Protocol = Site.PROTOCOLS.HTTP Then
                Return Convert.ToString(Convert.ToString("http://") & Me.m_SiteAddress) & Me.m_AdditionalAddressInfo
            Else
                Return Me.m_SiteAddress
            End If
        End Function

        Public Overrides Function ToString() As String
            Return Convert.ToString(m_SiteAddress & Convert.ToString(", ")) & Me.m_Description
        End Function
    End Class
End Namespace