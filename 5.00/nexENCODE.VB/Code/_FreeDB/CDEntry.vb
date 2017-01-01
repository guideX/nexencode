Imports System.Collections.Specialized
Imports System.Diagnostics
Imports System.Text

Namespace Freedb
    ''' <summary>
    ''' Summary description for CDEntry.
    ''' </summary>
    Public Class CDEntry



#Region "Private Member Variables"
        Private m_Discid As String
        Private m_Artist As String
        Private m_Title As String
        Private m_Year As String
        Private m_Genre As String
        Private m_Tracks As New TrackCollection()
        ' 0 based - first track is at 0 last track is at numtracks - 1
        Private m_ExtendedData As String
        Private m_PlayOrder As String
#End Region

#Region "Public Member Variables"
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

        ''' <summary>
        ''' Property Year (string)
        ''' </summary>
        Public Property Year() As String
            Get
                Return Me.m_Year
            End Get
            Set(value As String)
                Me.m_Year = value
            End Set
        End Property

        ''' <summary>
        ''' Property Genre (string)
        ''' </summary>
        Public Property Genre() As String
            Get
                Return Me.m_Genre
            End Get
            Set(value As String)
                Me.m_Genre = value
            End Set
        End Property


        ''' <summary>
        ''' Property Tracks (StringCollection)
        ''' </summary>
        Public Property Tracks() As TrackCollection
            Get
                Return Me.m_Tracks
            End Get
            Set(value As TrackCollection)
                Me.m_Tracks = value
            End Set
        End Property


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
        ''' Property PlayOrder (string)
        ''' </summary>
        Public Property PlayOrder() As String
            Get
                Return Me.m_PlayOrder
            End Get
            Set(value As String)
                Me.m_PlayOrder = value
            End Set
        End Property

        Public ReadOnly Property NumberOfTracks() As Integer
            Get
                Return m_Tracks.Count
            End Get
        End Property

#End Region



        Public Sub New(data As StringCollection)
            If Not Parse(data) Then
                Throw New Exception("Unable to Parse CDEntry.")
            End If
        End Sub


        Private Function Parse(data As StringCollection) As Boolean
            For Each line As String In data

                ' check for comment

                If line(0) = "#"c Then
                    Continue For
                End If

                Dim index As Integer = line.IndexOf("="c)
                If index = -1 Then
                    ' couldn't find equal sign have no clue what the data is
                    Continue For
                End If
                Dim field As String = line.Substring(0, index)
                index += 1
                ' move it past the equal sign
                Select Case field
                    Case "DISCID"
                        If True Then
                            Me.m_Discid = line.Substring(index)
                            'Continue Select
                        End If

                    Case "DTITLE"
                        ' artist / title
                        If True Then
                            Me.m_Artist += line.Substring(index)
                            'Continue Select
                        End If

                    Case "DYEAR"
                        If True Then
                            Me.m_Year = line.Substring(index)
                            'Continue Select
                        End If

                    Case "DGENRE"
                        If True Then
                            Me.m_Genre += line.Substring(index)
                            'Continue Select
                        End If

                    Case "EXTD"
                        If True Then
                            ' may be more than one - just concatenate them
                            Me.m_ExtendedData += line.Substring(index)
                            'Continue Select
                        End If

                    Case "PLAYORDER"
                        If True Then
                            Me.m_PlayOrder += line.Substring(index)
                            'Continue Select
                        End If
                    Case Else



                        'get track info or extended track info
                        If field.StartsWith("TTITLE") Then
                            Dim trackNumber As Integer = -1
                            ' Parse could throw an exception
                            Try
                                trackNumber = Integer.Parse(field.Substring("TTITLE".Length))

                            Catch ex As Exception
                                Debug.WriteLine("Failed to parse track Number. Reason: " + ex.Message)
                                'Continue Try
                            End Try

                            'may need to concatenate track info
                            If trackNumber < m_Tracks.Count Then
                                m_Tracks(trackNumber).Title += line.Substring(index)
                            Else
                                Dim track As New Track(line.Substring(index))
                                Me.m_Tracks.Add(track)
                            End If
                            'Continue Select
                        ElseIf field.StartsWith("EXTT") Then
                            Dim trackNumber As Integer = -1
                            ' Parse could throw an exception
                            Try
                                trackNumber = Integer.Parse(field.Substring("EXTT".Length))

                            Catch ex As Exception
                                Debug.WriteLine("Failed to parse track Number. Reason: " + ex.Message)
                                'Continue Try
                            End Try

                            If trackNumber < 0 OrElse trackNumber > m_Tracks.Count - 1 Then
                                'Continue Select
                            End If




                            m_Tracks(trackNumber).ExtendedData += line.Substring(index)
                        End If




                        'Continue Select
                        'end of switch
                End Select
            Next
            'split the title and artist from DTITLE;
            ' see if we have a slash
            Dim slash As Integer = Me.m_Artist.IndexOf(" / ")
            If slash = -1 Then
                Me.m_Title = m_Artist
            Else
                Dim titleArtist As String = m_Artist
                Me.m_Artist = titleArtist.Substring(0, slash)
                slash += 3
                ' move past " / "
                Me.m_Title = titleArtist.Substring(slash)
            End If
            Return True
        End Function

        Public Overrides Function ToString() As String
            Dim builder As New StringBuilder()
            builder.Append("Title: ")
            builder.Append(Me.m_Title)
            builder.Append(vbLf)
            builder.Append("Artist: ")
            builder.Append(Me.m_Artist)
            builder.Append(vbLf)
            builder.Append("Discid: ")
            builder.Append(Me.m_Discid)
            builder.Append(vbLf)
            builder.Append("Genre: ")
            builder.Append(Me.m_Genre)
            builder.Append(vbLf)
            builder.Append("Year: ")
            builder.Append(Me.m_Year)
            builder.Append(vbLf)
            builder.Append("Tracks:")
            For Each track As Track In Me.m_Tracks
                builder.Append(vbLf)
                builder.Append(track.Title)
            Next
            Return builder.ToString()
        End Function
    End Class
End Namespace