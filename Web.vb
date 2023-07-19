Imports System.Text

Namespace AI_SDK_EXTENSIONS

    Namespace Web

        Public Class GoogleMaps

            'Dataquery is to be returned to sender for use in a WebBrowser
            Private street As String = ""

            Private City As String = ""
            Private State As String = ""
            Private Zip As String = ""
            Public Query As String
            Private GoogleQuery As New StringBuilder

            Public Sub New(_street As String, _city As String, _State As String, _Zip As String)
                GoogleQuery.Append("Http://maps.google.com/maps?q=")
                Try
                    If _street <> "" Then GoogleQuery.Append(_street + "," & "+")
                    If _street <> "" Then GoogleQuery.Append(_city + "," & "+")
                    If _street <> "" Then GoogleQuery.Append(_State + "," & "+")
                    If _street <> "" Then GoogleQuery.Append(_Zip + "," & "+")
                    Query = GoogleQuery.ToString
                Catch ex As Exception
                    MsgBox(ex)
                End Try
            End Sub

        End Class

        Public Module ModuleWebSearch
            Public SearchVideo As String = ""
            Public SearchNews As String = ""
            Public SearchPictures As String = ""
            Public SearchRadio As String = ""
            Public SearchPhone As String = ""
            Public SearchBusiness As String = ""
            Public SearchLocation As String = ""
            Public SearchProduct As String = ""

            Public VideoSearch = "http://video.google.com/videosearch?q=" & SearchVideo & "&hl=en&emb=0&aq=f#"
            Public RadioSearch = "http://www.radioteam.eu/?action=archive&search=" & SearchRadio
            Public NewsSearch = "http://news.google.com/news?hl=en&tab=wn&ned=us&nolr=1&q=" & SearchNews & "&btnG=Search"
            Public PicSearch = "http://www.images.google.com/images?svnum=10&hl=en&lr=&q=" & SearchPictures & "&btnG=Search"
            Public PhoneSearch = "http://search.yahoo.com/search?p=phone%3A%22" & SearchPhone & "%22+&phone=" & SearchPhone & "&meta=pplt%3Dr&fr=php-rplu"
            Public ProductSearch = "http://www.google.com/products?q=" & SearchProduct & "&btnG=Search+Products&hl=en"
            Public BusinessSearch = "http://www.yellowpages.com/name/" & SearchLocation & "/" & SearchBusiness

            Public BbcSearch = "http://www.bbc.co.uk/mediaselector/ondemand/worldservice/meta/tx/live_news?bgc=003399&lang=en-ws&nbram=1&nbwm=1&ms3=2&ms_javascript=true&bbcws=1&size=au"
            Public WebSiteSearch = "www."
            Public Mapsearch = "http://www.google.com/local?q="
            Public FilipinoSearch = "http://www.eradioportal.com/index.php?p=2&aid=1/"
            Public WikiSearch = "http://en.wikipedia.org/wiki/"
            Public GoogleSearch = "http://www.google.com/search?q="
            Public StockSearch = "http://finance.google.com/finance?q="
            Public BookSearch = "http://www.google.com/books?q="
            Public YellowPagesSearch = "http://www.yellowpages.com/nationwide/name_search/"
            Public PersonSearch = "http://www.whitepages.com/5116/"
            Public SearchTextAol As String = "http://search.aol.co.uk/aol/search?"
            Public SearchTextGoogle As String = "https://www.google.co.uk/webhp?sourceid=chrome-instant&ion=1&espv=2&es_th=1&ie=UTF-8#q="
            Public SearchTextBing As String = "http://www.bing.com/search?q="
            Public Searchwikipedia As String = "http://en.wikipedia.org/w/index.php?title=Special%3ASearch&profile=default&search="
            Public SearchYoutube As String = "https://www.youtube.com/results?search_query="
            Public SearchGoogleMaps As String = "https://www.google.co.uk/maps/place/"
            Public InfoWarsSite = "http://www.infowars.com/"
            Public FacebookSite = "http://www.facebook.com/"
            Public PrisonPlanetSite = "http://www.prisonplanet.com/"
            Public VirtualhumansforumSite = "http://www.vrconsulting.it/vhf/"
            Public MsnSite = "http://www.msn.com/"
            Public GoogleSite = "http://www.google.com/"
            Public ZabawareforumSite = "http://www.zabaware.com/forum/"
            Public NationalnewsSite = "http://www.foxnews.com/"

        End Module
        Public Module Internet_News_Searching

            Public Function OpeningSite() As String
                OpeningSite = ""
                Select Case (Int(Rnd() * 4) + 1)
                    Case 1
                        OpeningSite = "Here is the page you asked me for." & vbCrLf
                    Case 2
                        OpeningSite = "Here it is. Now you can access it." & vbCrLf
                    Case 3
                        OpeningSite = "I am accessing it for you." & vbCrLf
                    Case 4
                        OpeningSite = "OK. Here it is." & vbCrLf
                    Case 5
                        OpeningSite = "I am opening the connection." & vbCrLf
                End Select
            End Function

            Dim SearchWebSite As String = ""
            Dim SearchVideo As String = ""
            Dim SearchNews As String = ""
            Dim SearchPictures As String = ""
            Dim SearchMaps As String = ""
            Dim SearchRadio As String = ""
            Dim SearchWiki As String = ""
            Dim UserInput As String = ""
            Dim SearchGoogle As String = ""
            Dim SearchPhone As String = ""
            Dim SearchStock As String = ""
            Dim SearchBusiness As String = ""
            Dim SearchLocation As String = ""
            Dim SearchProduct As String = ""
            Dim SearchBook As String = ""
            Dim SearchNation As String = ""

            Dim BbcSearch = "http://www.bbc.co.uk/mediaselector/ondemand/worldservice/meta/tx/live_news?bgc=003399&lang=en-ws&nbram=1&nbwm=1&ms3=2&ms_javascript=true&bbcws=1&size=au"
            Dim VideoSearch = "http://video.google.com/videosearch?q=" & SearchVideo & "&hl=en&emb=0&aq=f#"
            Dim RadioSearch = "http://www.windowsmedia.com/radiotuner/FindStations.asp?locale=409&search=" & SearchRadio
            Dim NewsSearch = "http://news.google.com/news?hl=en&tab=wn&ned=us&nolr=1&q=" & SearchNews & "&btnG=Search"
            Dim WebSiteSearch = "www." & SearchWebSite
            Dim Mapsearch = "http://www.google.com/local?q=" & SearchMaps
            Dim PicSearch = "http://www.images.google.com/images?svnum=10&hl=en&lr=&q=" & SearchPictures & "&btnG=Search"
            Dim WikiSearch = "http://en.wikipedia.org/wiki/" & SearchWiki
            Dim GoogleSearch = "http://www.google.com/search?q=" & SearchGoogle
            Dim FilipinoSearch = "http://www.eradioportal.com/index.php?p=2&aid=1"
            Dim PhoneSearch = "http://search.yahoo.com/search?p=phone%3A%22" & SearchPhone & "%22+&phone=" & SearchPhone & "&meta=pplt%3Dr&fr=php-rplu"
            Dim StockSearch = "http://finance.google.com/finance?q=" & SearchStock
            Dim BookSearch = "http://www.google.com/books?q=" & SearchBook
            Dim ProductSearch = "http://www.google.com/products?q=" & SearchProduct & "&btnG=Search+Products&hl=en"
            Dim BusinessSearch = "http://www.yellowpages.com/name/" & SearchLocation & "/" & SearchBusiness
            Dim NationSearch = "http://www.yellowpages.com/nationwide/name_search/" & SearchNation
            Dim PersonSearch = "http://www.whitepages.com/5116"

            Public Function NewsReader(ByRef Userinput As String) As String
                Dim NewsTopic As String = ""
                Dim NewsRead As String = Userinput
                If InStr(1, NewsRead, "CURRENT NEWS", 1) > 0 Then NewsTopic = "http://feeds.reuters.com/reuters/topNews"
                If InStr(1, NewsRead, "CURRENT NATIONAL NEWS", 1) > 0 Then NewsTopic = "http://feeds.reuters.com/reuters/domesticNews"
                If InStr(1, NewsRead, "CURRENT FINANCIAL NEWS", 1) > 0 Then NewsTopic = "http://feeds.reuters.com/reuters/businessNews"
                If InStr(1, NewsRead, "CURRENT WORLD NEWS", 1) > 0 Then NewsTopic = "http://feeds.reuters.com/reuters/worldNews"
                If InStr(1, NewsRead, "CURRENT ENTERTAINMENT NEWS", 1) > 0 Then NewsTopic = "http://feeds.reuters.com/reuters/entertainment"
                If InStr(1, NewsRead, "CURRENT SPORTS NEWS", 1) > 0 Then NewsTopic = "http://feeds.reuters.com/reuters/sportsNews"
                If InStr(1, NewsRead, "CURRENT TECHNOLOGY NEWS", 1) > 0 Then NewsTopic = "http://feeds.reuters.com/reuters/technologyNews"
                If InStr(1, NewsRead, "CURRENT POLITICAL NEWS", 1) > 0 Then NewsTopic = "http://feeds.reuters.com/reuters/politicsNews"
                If InStr(1, NewsRead, "CURRENT SCIENCE NEWS", 1) > 0 Then NewsTopic = "http://feeds.reuters.com/reuters/scienceNews"
                If InStr(1, NewsRead, "CURRENT HEALTH NEWS", 1) > 0 Then NewsTopic = "http://feeds.reuters.com/reuters/healthNews"
                If InStr(1, NewsRead, "CURRENT ODD NEWS", 1) > 0 Then NewsTopic = "http://feeds.reuters.com/reuters/oddlyEnoughNews"
                If InStr(1, NewsRead, "LATEST NEWS", 1) > 0 Then NewsTopic = "http://feeds.reuters.com/reuters/topNews"
                If InStr(1, NewsRead, "LATEST NATIONAL NEWS", 1) > 0 Then NewsTopic = "http://feeds.reuters.com/reuters/domesticNews"
                If InStr(1, NewsRead, "LATEST FINANCIAL NEWS", 1) > 0 Then NewsTopic = "http://feeds.reuters.com/reuters/businessNews"
                If InStr(1, NewsRead, "LATEST WORLD NEWS", 1) > 0 Then NewsTopic = "http://feeds.reuters.com/reuters/worldNews"
                If InStr(1, NewsRead, "LATEST ENTERTAINMENT NEWS", 1) > 0 Then NewsTopic = "http://feeds.reuters.com/reuters/entertainment"
                If InStr(1, NewsRead, "LATEST SPORTS NEWS", 1) > 0 Then NewsTopic = "http://feeds.reuters.com/reuters/sportsNews"
                If InStr(1, NewsRead, "LATEST TECHNOLOGY NEWS", 1) > 0 Then NewsTopic = "http://feeds.reuters.com/reuters/technologyNews"
                If InStr(1, NewsRead, "LATEST POLITICAL NEWS", 1) > 0 Then NewsTopic = "http://feeds.reuters.com/reuters/politicsNews"
                If InStr(1, NewsRead, "LATEST SCIENCE NEWS", 1) > 0 Then NewsTopic = "http://feeds.reuters.com/reuters/scienceNews"
                If InStr(1, NewsRead, "LATEST HEALTH NEWS", 1) > 0 Then NewsTopic = "http://feeds.reuters.com/reuters/healthNews"
                If InStr(1, NewsRead, "LATEST ODD NEWS", 1) > 0 Then NewsTopic = "http://feeds.reuters.com/reuters/oddlyEnoughNews"

                Return NewsTopic
            End Function

            Public Function WeatherReader(ByRef UserInput As String, ByRef SearchWeatherzip As String) As String
                Dim NewsTopic As String = ""
                Dim NewsRead As String = UserInput

                If InStr(1, NewsRead, "CURRENT WEATHER", 1) > 0 Then NewsTopic = "http://feeds.weatherbug.com/rss.aspx?zipcode=" & SearchWeatherzip & "&feed=currtxt&zcode=z4641"
                If InStr(1, NewsRead, "CURRENT WEATHER FORECAST", 1) > 0 Then NewsTopic = "http://feeds.weatherbug.com/rss.aspx?zipcode=" & SearchWeatherzip & "&feed=fcsttxt&zcode=z4641"
                If InStr(1, NewsRead, "CURRENT FORECAST", 1) > 0 Then NewsTopic = "http://feeds.weatherbug.com/rss.aspx?zipcode=" & SearchWeatherzip & "&feed=fcsttxt&zcode=z4641"
                If InStr(1, NewsRead, "LATEST WEATHER", 1) > 0 Then NewsTopic = "http://feeds.weatherbug.com/rss.aspx?zipcode=" & SearchWeatherzip & "&feed=currtxt&zcode=z4641"
                If InStr(1, NewsRead, "LATEST WEATHER FORECAST", 1) > 0 Then NewsTopic = "http://feeds.weatherbug.com/rss.aspx?zipcode=" & SearchWeatherzip & "&feed=fcsttxt&zcode=z4641"
                If InStr(1, NewsRead, "LATEST FORECAST", 1) > 0 Then NewsTopic = "http://feeds.weatherbug.com/rss.aspx?zipcode=" & SearchWeatherzip & "&feed=fcsttxt&zcode=z4641"
                If InStr(1, NewsRead, "CURRENT WEATHER FORCAST", 1) > 0 Then NewsTopic = "http://feeds.weatherbug.com/rss.aspx?zipcode=" & SearchWeatherzip & "&feed=fcsttxt&zcode=z4641"
                If InStr(1, NewsRead, "CURRENT FORCAST", 1) > 0 Then NewsTopic = "http://feeds.weatherbug.com/rss.aspx?zipcode=" & SearchWeatherzip & "&feed=fcsttxt&zcode=z4641"
                If InStr(1, NewsRead, "LATEST WEATHER FORCAST", 1) > 0 Then NewsTopic = "http://feeds.weatherbug.com/rss.aspx?zipcode=" & SearchWeatherzip & "&feed=fcsttxt&zcode=z4641"
                If InStr(1, NewsRead, "LATEST FORCAST", 1) > 0 Then NewsTopic = "http://feeds.weatherbug.com/rss.aspx?zipcode=" & SearchWeatherzip & "&feed=fcsttxt&zcode=z4641"
                Return NewsTopic
            End Function

            Public Function AttemptFeed() As String
                AttemptFeed = ""
                Dim Choice As Integer
                Select Case Choice
                    Case 1
                        AttemptFeed = "One moment, linking to news feed."
                    Case 2
                        AttemptFeed = "Buffering news stream."
                    Case 3
                        AttemptFeed = "Awaiting news stream."
                    Case 4
                        AttemptFeed = "Updating news feed."
                    Case 5
                        AttemptFeed = "Just a moment. Parsing news data."
                    Case 6
                        AttemptFeed = "Downloading news data."
                End Select
                Return AttemptFeed
            End Function

            Public Function RespondFeed(ByRef FeedData As String) As String
                Dim Choice As Integer
                RespondFeed = ""
                Select Case Choice
                    Case 1
                        RespondFeed = "Just a moment." & vbCrLf & FeedData & "This concludes the news feed. Terminating connection."
                    Case 2
                        RespondFeed = "One moment." & vbCrLf & FeedData & "End of news feed. Closing data stream."
                    Case 3
                        RespondFeed = "Just a moment." & vbCrLf & FeedData & "That is all the news data at this time. News feed terminated."
                    Case 4
                        RespondFeed = "One moment." & vbCrLf & FeedData & "News feed complete. Closing uplink."
                    Case 5
                        RespondFeed = "Just a moment." & vbCrLf & FeedData & "That is all of the current data. Terminating feed."
                    Case 6
                        RespondFeed = "One moment." & vbCrLf & FeedData & "This concludes the news stream. Closing all stream connections."
                End Select

            End Function

            Public Function RespondBadFeed() As String
                RespondBadFeed = ""
                Dim Choice As Integer
                Select Case Choice
                    Case 1
                        RespondBadFeed = "That stream contains no data at the present time. I'm sorry. Please try again momentarily."
                    Case 2
                        RespondBadFeed = "I can't seem to find a valid source at that location. I'm sorry. The stream must be in the process of updating the data. Please try again in a moment."
                    Case 3
                        RespondBadFeed = "I'm sorry to have to tell you this, but that source is not operational. If the data is being updated then, it will be temporarily unavailable."
                    Case 4
                        RespondBadFeed = "There seems to be a problem. That stream is currently void of any usable data. The data must be updating. The stream should be available in a moment."
                    Case 5
                        RespondBadFeed = "There seems to be a fault in the uplink. Please confirm the source, and try again in a moment."
                    Case 6
                        RespondBadFeed = "There must be a fault in the data stream. Please check the uplink source, and try again."
                End Select
            End Function

        End Module
    End Namespace

End Namespace