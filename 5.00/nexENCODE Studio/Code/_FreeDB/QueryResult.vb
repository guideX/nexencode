Namespace Freedb
    ''' <summary>
    ''' Summary description for QueryResult.
    ''' </summary>
    Public Class QueryResult
        Private m_ResponseCode As String
        Private m_Category As String
        Private m_Discid As String
        Private m_Artist As String
        Private m_Title As String


#Region "Public Properties"
        ''' <summary>
        ''' Property ResponseCode (string)
        ''' </summary>
        Public Property ResponseCode() As String
            Get
                Return Me.m_ResponseCode
            End Get
            Set(value As String)
                Me.m_ResponseCode = value
            End Set
        End Property


        ''' <summary>
        ''' Property Category (string)
        ''' </summary>
        Public Property Category() As String
            Get
                Return Me.m_Category
            End Get
            Set(value As String)
                Me.m_Category = value
            End Set
        End Property

        ''' <summary>
        ''' Property Discid (string)
        ''' </summary>
        Public Property Discid() As String
            Get
                Return Me.m_Discid
            End Get
            Set(value As String)
                Me.m_Discid = value
            End Set
        End Property

        ''' <summary>
        ''' Property Artist (string)
        ''' </summary>
        Public Property Artist() As String
            Get
                Return Me.m_Artist
            End Get
            Set(value As String)
                Me.m_Artist = value
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

        Public Sub New(queryResult__1 As String)
            If Not Parse(queryResult__1, False) Then
                Throw New Exception(Convert.ToString("Unable to Parse QueryResult. Input: ") & queryResult__1)

            End If
        End Sub

        ''' <summary>
        ''' The parsing for a queryresult returned as part of a number of matches is slightly different
        ''' There is no response code
        ''' </summary>
        ''' <param name="queryResult"></param>
        ''' <param name="multiMatchInput"> true if the result is part of multi-match which means it will not contain a response code</param>
        Public Sub New(queryResult__1 As String, multiMatchInput As Boolean)
            If Not Parse(queryResult__1, multiMatchInput) Then
                Throw New Exception(Convert.ToString("Unable to Parse QueryResult. Input: ") & queryResult__1)

            End If
        End Sub

        ''' <summary>
        ''' Parse the query result line from the cddb server
        ''' </summary>
        ''' <param name="queryResult"></param>
        Public Function Parse(queryResult As String, match As Boolean) As Boolean

            queryResult.Trim()
            Dim secondIndex As Integer = 0

            ' get first white space
            Dim index As Integer = queryResult.IndexOf(" "c)
            'if we are parsing a matched queryresult there is no responsecode so skip it
            If Not match Then
                m_ResponseCode = queryResult.Substring(0, index)
                index += 1
                secondIndex = queryResult.IndexOf(" "c, index)
            Else
                secondIndex = index
                index = 0
            End If

            m_Category = queryResult.Substring(index, secondIndex - index)
            index = secondIndex
            index += 1
            secondIndex = queryResult.IndexOf(" "c, index)
            m_Discid = queryResult.Substring(index, secondIndex - index)
            index = secondIndex
            index += 1
            secondIndex = queryResult.IndexOf("/"c, index)
            m_Artist = queryResult.Substring(index, secondIndex - index - 1)
            ' -1 because there is a space at the end of artist
            index = secondIndex
            index += 2
            'skip past / and space
            m_Title = queryResult.Substring(index)
            Return True
        End Function

        '		public bool Parse(string queryResult)
        '		{
        '			queryResult.Trim();
        '			string [] values = queryResult.Split(' ');
        '			if (values.Length <6)
        '				return false;
        '			this.m_ResponseCode = values[0];
        '			m_Category = values[1];
        '			m_Discid = values[2];
        '
        '			// now we need to look for a slash
        '			bool artist = true;
        '			for (int i = 3; i < values.Length;i++)
        '			{
        '				if (values[i] == "/")
        '				{
        '					artist = false;
        '					continue;
        '				}
        '				if (artist)
        '					this.m_Artist += values[i];
        '				else
        '					this.m_Title += values[i];
        '
        '			}
        '			return true;
        '		}

        Public Overrides Function ToString() As String
            Return Convert.ToString(Me.m_Artist & Convert.ToString(", ")) & Me.m_Title
        End Function

    End Class
End Namespace