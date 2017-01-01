Imports System.Text
Imports System.Net
Imports System.IO
Imports System.Collections.Specialized
Imports System.Diagnostics

Namespace Freedb
    ''' <summary>
    ''' Summary description for FreedbHelper.
    ''' </summary>
    Public Class FreedbHelper
        Public Const MAIN_FREEDB_ADDRESS As String = "freedb.freedb.org"
        Public Const DEFAULT_ADDITIONAL_URL_INFO As String = "/~cddb/cddb.cgi"
        Private m_mainSite As New Site(MAIN_FREEDB_ADDRESS, "http", DEFAULT_ADDITIONAL_URL_INFO)
        Private m_UserName As String
        Private m_Hostname As String
        Private m_ClientName As String
        Private m_Version As String
        Private m_ProtocolLevel As String = "6"
        ' default to level 6 support
        Private m_CurrentSite As Site = Nothing


#Region "Constants for Freedb commands"
        Public Class Commands
            Public Const CMD_HELLO As String = "hello"
            Public Const CMD_READ As String = "cddb+read"
            Public Const CMD_QUERY As String = "cddb+query"
            Public Const CMD_SITES As String = "sites"
            Public Const CMD_PROTO As String = "proto"
            Public Const CMD_CATEGORIES As String = "cddb+lscat"
            Public Const CMD As String = "cmd="
            ' will never use without the equals so put it here
            Public Const CMD_TERMINATOR As String = "."
        End Class
#End Region

#Region "Constants for Freedb ResponseCodes"
        Public Class ResponseCodes
            Public Const CODE_210 As String = "210"
            ' Okay // or in a query multiple exact matches
            Public Const CODE_401 As String = "401"
            ' sites: no site information available
            Public Const CODE_402 As String = "402"
            ' Server Error
            Public Const CODE_500 As String = "500"
            ' Invalid command, invalid parameters, etc.
            'query codes
            Public Const CODE_200 As String = "200"
            ' Exact match 
            Public Const CODE_211 As String = "211"
            ' InExact matches found - list follows
            Public Const CODE_202 As String = "202"
            ' No match 
            Public Const CODE_403 As String = "403"
            ' Database entry is corrupt
            Public Const CODE_409 As String = "409"
            ' No Handshake
            ' our own code
            Public Const CODE_INVALID As String = "-1"
            ' Invalid code 
        End Class
#End Region


#Region "Public Properties"
        ''' <summary>
        ''' Property Version (string)
        ''' </summary>
        Public Property Version() As String
            Get
                Return Me.m_Version
            End Get
            Set(value As String)
                Me.m_Version = value
            End Set
        End Property






        ''' <summary>
        ''' Property MainSite(string)
        ''' </summary>
        Public ReadOnly Property MainSite() As Site
            Get
                Return Me.m_mainSite
            End Get
        End Property

        ''' <summary>
        ''' Property ClientName (string)
        ''' </summary>
        Public Property ClientName() As String
            Get
                Return Me.m_ClientName
            End Get
            Set(value As String)
                Me.m_ClientName = value
            End Set
        End Property

        ''' <summary>
        ''' Property Hostname (string)
        ''' </summary>
        Public Property Hostname() As String
            Get
                Return Me.m_Hostname
            End Get
            Set(value As String)
                Me.m_Hostname = value
            End Set
        End Property

        ''' <summary>
        ''' Property UserName (string)
        ''' </summary>
        Public Property UserName() As String
            Get
                Return Me.m_UserName
            End Get
            Set(value As String)
                Me.m_UserName = value
            End Set
        End Property

        ''' <summary>
        ''' Property ProtocolLevel (string)
        ''' </summary>
        Public Property ProtocolLevel() As String
            Get
                Return Me.m_ProtocolLevel
            End Get
            Set(value As String)
                Me.m_ProtocolLevel = value
            End Set
        End Property

        ''' <summary>
        ''' Property CurrentSite (Site)
        ''' </summary>
        Public Property CurrentSite() As Site
            Get
                Return Me.m_CurrentSite
            End Get
            Set(value As Site)
                Me.m_CurrentSite = value
            End Set
        End Property




        Public Sub New()
            ' default it
            m_ProtocolLevel = "6"
        End Sub



        ''' <summary>
        ''' Retrieve all Freedb servers from the main server site
        ''' </summary>
        ''' <param name="sites">SiteCollection that is populated with the site information</param>
        ''' <returns>Response Code</returns>
        Public Function GetSites(ByRef sites As SiteCollection) As String
            Return GetSites(Site.PROTOCOLS.ALL, sites)
        End Function

#End Region


        ''' <summary>
        ''' Get the Freedb sites
        ''' </summary>
        ''' <param name="protocol"></param>
        ''' <param name="sites">SiteCollection that is populated with the site information</param>
        ''' <returns>Response Code</returns>
        ''' 
        Public Function GetSites(protocol As String, ByRef sites As SiteCollection) As String
            If protocol <> Site.PROTOCOLS.CDDBP AndAlso protocol <> Site.PROTOCOLS.HTTP Then
                protocol = Site.PROTOCOLS.ALL
            End If

            Dim coll As StringCollection

            Try
                coll = [Call](Commands.CMD_SITES, m_mainSite.GetUrl())

            Catch ex As Exception
                Debug.WriteLine("Error retrieving Sites." + ex.Message)
                Dim newEx As New Exception("FreedbHelper.GetSites: Error retrieving Sites.", ex)
                Throw newEx
            End Try

            sites = Nothing

            ' check if results came back
            If coll.Count < 0 Then
                Dim msg As String = "No results returned from sites request."
                Dim ex As New Exception(msg, Nothing)
                Throw ex
            End If

            Dim code As String = GetCode(coll(0))
            If code = ResponseCodes.CODE_INVALID Then
                Dim msg As String = "Unable to process results Sites Request. Returned Data: " + coll(0)
                Dim ex As New Exception(msg, Nothing)
                Throw ex
            End If

            Select Case code
                Case ResponseCodes.CODE_500
                    Return ResponseCodes.CODE_500

                Case ResponseCodes.CODE_401
                    Return ResponseCodes.CODE_401

                Case ResponseCodes.CODE_210
                    If True Then
                        coll.RemoveAt(0)
                        sites = New SiteCollection()
                        For Each line As [String] In coll
                            Debug.WriteLine("line: " + line)
                            Dim site__1 As New Site(line)
                            If protocol = Site.PROTOCOLS.ALL Then
                                sites.Add(New Site(line))
                            ElseIf site__1.Protocol = protocol Then
                                sites.Add(New Site(line))
                            End If
                        Next

                        Return ResponseCodes.CODE_210
                    End If
                Case Else

                    Return ResponseCodes.CODE_500
            End Select

        End Function


        ''' <summary>
        ''' Read Entry from the database. 
        ''' </summary>
        ''' <param name="qr">A QueryResult object that is created by performing a query</param>
        ''' <param name="cdEntry">out parameter - CDEntry object</param>
        ''' <returns></returns>
        Public Function Read(qr As QueryResult, ByRef cdEntry As CDEntry) As String
            Debug.Assert(qr IsNot Nothing)
            cdEntry = Nothing

            Dim coll As StringCollection = Nothing
            Dim builder As New StringBuilder(FreedbHelper.Commands.CMD_READ)
            builder.Append("+")
            builder.Append(qr.Category)
            builder.Append("+")
            builder.Append(qr.Discid)

            'make call
            Try
                coll = [Call](builder.ToString())

            Catch ex As Exception
                Dim msg As String = "Error performing cddb read."
                Dim newex As New Exception(msg, ex)
                Throw newex
            End Try

            ' check if results came back
            If coll.Count < 0 Then
                Dim msg As String = "No results returned from cddb read."
                Dim ex As New Exception(msg, Nothing)
                Throw ex
            End If


            Dim code As String = GetCode(coll(0))
            If code = ResponseCodes.CODE_INVALID Then
                Dim msg As String = "Unable to process results for cddb read. Returned Data: " + coll(0)
                Dim ex As New Exception(msg, Nothing)
                Throw ex
            End If


            Select Case code
                Case ResponseCodes.CODE_500
                    Return ResponseCodes.CODE_500

                    ' entry not found
                    ' server error
                    ' Database entry is corrupt
                Case ResponseCodes.CODE_401, ResponseCodes.CODE_402, ResponseCodes.CODE_403, ResponseCodes.CODE_409
                    ' No handshake
                    Return code

                Case ResponseCodes.CODE_210
                    ' good 
                    If True Then
                        coll.RemoveAt(0)
                        ' remove the 210
                        cdEntry = New CDEntry(coll)
                        Return ResponseCodes.CODE_210
                    End If
                Case Else
                    Return ResponseCodes.CODE_500
            End Select
        End Function


        ''' <summary>
        ''' Query the freedb server to see if there is information on this cd
        ''' </summary>
        ''' <param name="querystring"></param>
        ''' <param name="queryResult"></param>
        ''' <param name="queryResultsColl"></param>
        ''' <returns></returns>
        Public Function Query(querystring As String, ByRef queryResult As QueryResult, ByRef queryResultsColl As QueryResultCollection) As String
            queryResult = Nothing
            queryResultsColl = Nothing
            Dim coll As StringCollection = Nothing

            Dim builder As New StringBuilder(FreedbHelper.Commands.CMD_QUERY)
            builder.Append("+")
            builder.Append(querystring)

            'make call
            Try
                coll = [Call](builder.ToString())

            Catch ex As Exception
                Dim msg As String = "Unable to perform cddb query."
                Dim newex As New Exception(msg, ex)
                Throw newex
            End Try

            ' check if results came back
            If coll.Count < 0 Then
                Dim msg As String = "No results returned from cddb query."
                Dim ex As New Exception(msg, Nothing)
                Throw ex
            End If

            Dim code As String = GetCode(coll(0))
            If code = ResponseCodes.CODE_INVALID Then
                Dim msg As String = "Unable to process results returned for query: Data returned: " + coll(0)
                Dim ex As New Exception(msg, Nothing)
                Throw ex
            End If


            Select Case code
                Case ResponseCodes.CODE_500
                    Return ResponseCodes.CODE_500

                    ' Multiple results were returned
                    ' Put them into a queryResultCollection object
                Case ResponseCodes.CODE_211, ResponseCodes.CODE_210
                    If True Then
                        queryResultsColl = New QueryResultCollection()
                        'remove the 210 or 211
                        coll.RemoveAt(0)
                        For Each line As String In coll
                            Dim result As New QueryResult(line, True)
                            queryResultsColl.Add(result)
                        Next

                        Return ResponseCodes.CODE_211
                    End If


                    ' exact match 
                Case ResponseCodes.CODE_200
                    If True Then
                        queryResult = New QueryResult(coll(0))
                        Return ResponseCodes.CODE_200
                    End If


                    'not found
                Case ResponseCodes.CODE_202
                    Return ResponseCodes.CODE_202

                    'Database entry is corrupt
                Case ResponseCodes.CODE_403
                    Return ResponseCodes.CODE_403

                    'no handshake
                Case ResponseCodes.CODE_409
                    Return ResponseCodes.CODE_409
                Case Else

                    Return ResponseCodes.CODE_500

            End Select
            ' end of switch

        End Function


        ''' <summary>
        ''' Retrieve the categories
        ''' </summary>
        ''' <param name="strings"></param>
        ''' <returns></returns>
        Public Function GetCategories(ByRef strings As StringCollection) As String

            Dim coll As StringCollection
            strings = Nothing

            Try
                coll = [Call](FreedbHelper.Commands.CMD_CATEGORIES)

            Catch ex As Exception
                Dim msg As String = "Unable to retrieve Categories."
                Dim newex As New Exception(msg, ex)
                Throw newex
            End Try

            ' check if results came back
            If coll.Count < 0 Then
                Dim msg As String = "No results returned from categories request."
                Dim ex As New Exception(msg, Nothing)
                Throw ex
            End If

            Dim code As String = GetCode(coll(0))
            If code = ResponseCodes.CODE_INVALID Then
                Dim msg As String = "Unable to retrieve Categories. Data Returned: " + coll(0)
                Dim ex As New Exception(msg, Nothing)
                Throw ex
            End If

            Select Case code
                Case ResponseCodes.CODE_500
                    Return ResponseCodes.CODE_500

                Case ResponseCodes.CODE_210
                    If True Then
                        strings = coll
                        coll.RemoveAt(0)
                        Return ResponseCodes.CODE_210
                    End If
                Case Else

                    If True Then
                        Dim msg As String = "Unknown code returned from GetCategories: " + coll(0)
                        Dim ex As New Exception(msg, Nothing)
                        Throw ex
                    End If


            End Select

        End Function


        ''' <summary>
        ''' Call the Freedb server using the specified command and the current site
        ''' If the current site is null use the default server
        ''' </summary>
        ''' <param name="command">The command to be exectued</param>
        ''' <returns>StringCollection</returns>
        Private Function [Call](command As String) As StringCollection
            If m_CurrentSite IsNot Nothing Then
                Return [Call](command, m_CurrentSite.GetUrl())
            Else
                Return [Call](command, m_mainSite.GetUrl())
            End If
        End Function

        ''' <summary>
        ''' Call the Freedb server using the specified command and the specified url
        ''' The command should not include the cmd= and hello and proto parameters.
        ''' They will be added automatically
        ''' </summary>
        ''' <param name="command">The command to be exectued</param>
        ''' <param name="url">The Freedb server to use</param>
        ''' <returns>StringCollection</returns>
        Private Function [Call](commandIn As String, url As String) As StringCollection
            Dim reader As StreamReader = Nothing
            Dim response As HttpWebResponse = Nothing
            Dim coll As New StringCollection()

            Try
                'create our HttpWebRequest which we use to call the freedb server
                Dim req As HttpWebRequest = DirectCast(WebRequest.Create(url), HttpWebRequest)
                req.ContentType = "text/plain"
                ' we are using th POST method of calling the http server. We could have also used the GET method
                req.Method = "POST"
                'add the hello and proto commands to the request
                Dim command As String = BuildCommand(Commands.CMD & commandIn)
                'using Unicode
                Dim byteArray As Byte() = Encoding.UTF8.GetBytes(command)
                'get our request stream
                Dim newStream As Stream = req.GetRequestStream()
                'write our command data to it
                newStream.Write(byteArray, 0, byteArray.Length)
                newStream.Close()
                'Make the call. Note this is a synchronous call
                response = DirectCast(req.GetResponse(), HttpWebResponse)
                'put the results into a StreamReader
                reader = New StreamReader(response.GetResponseStream(), System.Text.Encoding.UTF8)
                ' add each line to the StringCollection until we get the terminator
                Dim line As String
                While (InlineAssignHelper(line, reader.ReadLine())) IsNot Nothing
                    If line.StartsWith(Commands.CMD_TERMINATOR) Then
                        Exit While
                    Else
                        coll.Add(line)
                    End If
                End While

            Catch ex As Exception
                Throw ex
            Finally

                If response IsNot Nothing Then
                    response.Close()
                End If
                If reader IsNot Nothing Then
                    reader.Close()
                End If
            End Try

            Return coll
        End Function





        ''' <summary>
        ''' Given a specific command add on the hello and proto which are requied for an http call
        ''' </summary>
        ''' <param name="command"></param>
        ''' <returns></returns>
        Private Function BuildCommand(command As String) As String
            Dim builder As New StringBuilder(command)
            builder.Append("&")
            builder.Append(Hello())
            builder.Append("&")
            builder.Append(Proto())
            Return builder.ToString()
        End Function

        ''' <summary>
        ''' Build the hello part of the command 
        ''' </summary>
        ''' <returns></returns>
        Public Function Hello() As String
            Dim builder As New StringBuilder(Commands.CMD_HELLO)
            builder.Append("=")
            builder.Append(m_UserName)
            builder.Append("+")
            builder.Append(Me.m_Hostname)
            builder.Append("+")
            builder.Append(Me.ClientName)
            builder.Append("+")
            builder.Append(Me.m_Version)
            Return builder.ToString()
        End Function

        ''' <summary>
        ''' Build the Proto part of the command
        ''' </summary>
        ''' <returns></returns>
        Public Function Proto() As String
            Dim builder As New StringBuilder(Commands.CMD_PROTO)
            builder.Append("=")
            builder.Append(m_ProtocolLevel)
            Return builder.ToString()
        End Function


        ''' <summary>
        ''' given the first line of a result set return the CDDB code
        ''' </summary>
        ''' <param name="firstLine"></param>
        ''' <returns></returns>
        Private Function GetCode(firstLine As String) As String
            firstLine = firstLine.Trim()

            'find first white space after start
            Dim index As Integer = firstLine.IndexOf(" "c)
            If index <> -1 Then
                firstLine = firstLine.Substring(0, index)
            Else
                Return ResponseCodes.CODE_INVALID
            End If

            Return firstLine
        End Function



        ''' <summary>
        ''' If a different default site address is preferred over "freedb.freedb.org"
        ''' set it here
        ''' NOTE: Only set the ip address
        ''' </summary>
        ''' <param name="ipAddress"></param>
        Public Sub SetDefaultSiteAddress(siteAddress As String)
            'sanity check on the url
            If siteAddress.IndexOf("http") <> -1 OrElse siteAddress.IndexOf("cgi") <> -1 Then
                Throw New Exception("Invalid Site Address specified")
            End If

            Me.m_mainSite.SiteAddress = siteAddress
        End Sub
        Private Shared Function InlineAssignHelper(Of T)(ByRef target As T, value As T) As T
            target = value
            Return value
        End Function

    End Class
End Namespace