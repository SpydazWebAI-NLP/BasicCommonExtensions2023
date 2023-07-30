Imports System.Drawing
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Web.Script.Serialization
Imports System.Windows.Forms
Imports System.Xml.Serialization
Imports SpydazWebAI.SystemExtensions.AI_SDK_EXTENSIONS.MathsExt.Maths
Imports SpydazWebAI.SystemExtensions.AI_SDK_EXTENSIONS.Strings
Imports SpydazWebAI.SystemExtensions.AI_SDK_EXTENSIONS.Strings.Grammar

Namespace DesignPatterns

    Namespace StateMachinePattern

        ''' <summary>
        ''' Can be inherited or Used as is
        ''' </summary>
        Public Class State

            ''' <summary>
            ''' Name of State (also used as Identity)
            ''' </summary>
            ''' <returns></returns>
            Public Property StateName As String = ""

            Public Function ToJson() As String
                Dim Converter As New JavaScriptSerializer
                Return Converter.Serialize(Me)
            End Function

        End Class

        ''' <summary>
        ''' State Machine
        ''' </summary>
        Public Class StateMachine

            ''' <summary>
            ''' Current State of the machine
            ''' </summary>
            Private CurrentState As State

            Private mMachineName As String = ""

            ''' <summary>
            ''' Current states held in the Machine
            ''' </summary>
            Private States As New List(Of State)

            ''' <summary>
            ''' </summary>
            ''' <param name="State">
            ''' IEmotionalState.EmotionalStateName used to set the initial state
            ''' </param>
            ''' <param name="StatesList"></param>
            Public Sub New(ByRef State As String, ByRef StatesList As List(Of State))
                If StatesList IsNot Nothing Then
                    States = StatesList
                    If State IsNot Nothing Then
                        If CheckState(State) = True Then
                            ChangeState(State)
                        End If
                    End If
                Else
                End If

            End Sub

            ''' <summary>
            ''' Informs The Controller that the state as been changed
            ''' </summary>
            ''' <param name="Sender"></param>
            ''' <param name="CurrentState"></param>
            Public Event StateChanged(ByRef Sender As StateMachine, ByRef CurrentState As State)

            Public Property MachineName As String
                Get
                    Return mMachineName
                End Get
                Set(value As String)
                    mMachineName = value
                End Set
            End Property

            ''' <summary>
            ''' Adds a new state to the state machines held possible states
            ''' </summary>
            ''' <param name="State"></param>
            Public Function AddState(ByRef State As State) As Boolean
                Dim Added As Boolean = False
                If State IsNot Nothing Then
                    If State.StateName IsNot Nothing Then

                        If CheckState(State) = False Then
                            States.Add(State)
                            Added = True
                        Else
                        End If
                    Else
                    End If
                Else
                End If
                Return Added
            End Function

            ''' <summary>
            ''' Changes the state to the specified State, which must exist in the state machine
            ''' </summary>
            ''' <param name="State">IEmotionalState.EmotionalStateName</param>
            Public Function ChangeState(ByRef State As String) As Boolean
                Dim StateChange As Boolean = False
                If State IsNot Nothing Then
                    For Each item In States
                        If item.StateName = State Then
                            CurrentState = item
                            StateChange = True
                            RaiseEvent StateChanged(Me, item)
                        Else
                        End If
                    Next
                Else
                End If
                Return StateChange
            End Function

            ''' <summary>
            ''' Returns the current state
            ''' </summary>
            ''' <returns></returns>
            Public Function GetState() As State
                Return CurrentState
            End Function

            ''' <summary>
            ''' checks if the state exists in the the list of states
            ''' </summary>
            ''' <param name="State">State To be checked</param>
            ''' <returns></returns>
            Private Function CheckState(ByRef State As State) As Boolean
                Dim found As Boolean = False
                For Each item In States
                    If item.StateName = State.StateName = True Then
                        found = True
                    Else

                    End If
                Next
                Return found
            End Function

            ''' <summary>
            ''' checks if the state exists in the the list of states
            ''' </summary>
            ''' <param name="State">IEmotionalState.EmotionalStateName</param>
            ''' <returns></returns>
            Private Function CheckState(ByRef State As String) As Boolean
                Dim found As Boolean = False
                For Each item In States
                    If item.StateName = State = True Then
                        found = True
                    Else

                    End If
                Next
                Return found
            End Function

        End Class

        ''' <summary>
        ''' Controller for the state machine, Adds/Removes states / Changes States
        ''' </summary>
        Public Class StateMachineController

            ''' <summary>
            ''' State machine
            ''' </summary>
            Private WithEvents StateMachine As StateMachine

            Public MachineState As String = ""

            Public Sub New(ByRef InitialState As String, ByRef States As List(Of String))
                CreateStateMachine(InitialState, CreateStates(States))
            End Sub

            Public Sub New(ByRef InitialState As String, ByRef States As List(Of State))
                CreateStateMachine(InitialState, States)
            End Sub

            Public Event Err(ByRef Err As String)

            ''' <summary>
            ''' State Machine State Changed
            ''' </summary>
            ''' <param name="Sender"></param>
            ''' <param name="Changed"></param>
            Public Event MachineStateChanged(ByRef Sender As StateMachineController, ByRef Changed As Boolean)

            ''' <summary>
            ''' Current state machine
            ''' </summary>
            ''' <returns></returns>
            Public ReadOnly Property CurrentStateMachine As StateMachine
                Get
                    If StateMachine IsNot Nothing Then
                        Return StateMachine
                    Else
                        Return Nothing
                    End If
                End Get
            End Property

            ''' <summary>
            ''' Creates a List of states for the state machine given a list of StateNames Req. for
            ''' creating a state machine
            ''' </summary>
            ''' <param name="ListOfStates"></param>
            ''' <returns></returns>
            Public Shared Function CreateStates(ByRef ListOfStates As List(Of String)) As List(Of State)
                Dim Lst As New List(Of State)
                For Each item In ListOfStates
                    Dim NewState As New State
                    NewState.StateName = item
                Next
                Return Lst
            End Function

            ''' <summary>
            ''' Adds as new state to the state machine
            ''' </summary>
            ''' <param name="State"></param>
            Public Function AddState(ByRef State As State) As Boolean
                Return StateMachine.AddState(State)
            End Function

            ''' <summary>
            ''' Changes current state of the state machine
            ''' </summary>
            ''' <param name="State"></param>
            Public Function ChangeState(ByRef State As String) As Boolean
                Return StateMachine.ChangeState(State)
            End Function

            ''' <summary>
            ''' Returns current state of the state machine
            ''' </summary>
            ''' <returns></returns>
            Public Function GetCurrentState() As State
                If StateMachine IsNot Nothing Then
                    Return StateMachine.GetState()
                End If
                Return Nothing
            End Function

            ''' <summary>
            ''' Creates a new state machine to be controlled by the state machine controller;
            ''' </summary>
            ''' <param name="InitialState">initial state of the state machine</param>
            ''' <param name="States">list of states held in the state machine rto be created</param>
            Private Sub CreateStateMachine(ByRef InitialState As String, ByRef States As List(Of State))
                StateMachine = New StateMachine(InitialState, States)
            End Sub

            Private Sub StateMachine_StateChanged(ByRef Sender As StateMachine, ByRef CurrentState As State) Handles StateMachine.StateChanged
                MachineState = CurrentState.StateName
                RaiseEvent MachineStateChanged(Me, True)
            End Sub

        End Class

    End Namespace

    Namespace InferenceEngines

        ''' <summary>
        ''' Based on the inference engine paradigm: Under Construction / Research
        ''' </summary>
        Public Class InferenceEngineParadigm

            Public Function StartInferenceEngine() As Boolean
                Dim ResultsFound As Boolean = False
                '1. Enter data into matches: Generate ConflictSet
                '2. Add ConflictSet to Resolve : Generate Results
                '3. Execute Results
                Return ResultsFound
            End Function

            ''' <summary>
            ''' Conflict items (Ruleset - RulesetItems should be the same shape)
            ''' </summary>
            Public Structure ConflictSet
                Public items As List(Of ConflictItem)

                Public Function ToJson() As String
                    Dim Converter As New JavaScriptSerializer
                    Return Converter.Serialize(Me)

                End Function

                ''' <summary>
                ''' items to be in Conflct set (Ruleset - RulesetItems should be the same shape)
                ''' </summary>
                Public Structure ConflictItem
                    Public cItem As String

                    Public Rating As Integer

                    Public Function ToJson() As String
                        Dim Converter As New JavaScriptSerializer
                        Return Converter.Serialize(Me)

                    End Function

                End Structure

            End Structure

            ''' <summary>
            ''' Execute Results
            ''' </summary>
            Public Class Execute
                'Execute result
            End Class

            ''' <summary>
            ''' Enter data into matches: Generate ConflictSet
            ''' </summary>
            Public Class Matches

                Private mGeneratedConflictSet As ConflictSet

                Public Sub New()

                End Sub

                ''' <summary>
                ''' returned ConflictSet
                ''' </summary>
                ''' <returns></returns>
                Public ReadOnly Property GeneratedConflictSet As ConflictSet
                    Get
                        Return mGeneratedConflictSet
                    End Get
                End Property

                ''' <summary>
                ''' Check data vs rules
                ''' </summary>
                Public Sub Data()
                    'Get The data
                End Sub

                ''' <summary>
                ''' Keep items fitting the Rule conditions (add to conflict Set)
                ''' </summary>
                Public Sub GenerateConflictSet()
                    'Check data vs rules
                    'Keep items fitting the Rule conditions (add to conflict Set)
                End Sub

                ''' <summary>
                ''' Check set against ruleset
                ''' </summary>
                Public Sub Rules()
                    'List of rules and conditions
                End Sub

            End Class

            ''' <summary>
            ''' Add ConflictSet to Resolve : Generate Results
            ''' </summary>
            Public Class Resolve

                'Interrogate ConflictSet :
                ':use strategy :
                Private mResults As New ConflictSet.ConflictItem

                Public ReadOnly Property Results As ConflictSet.ConflictItem
                    Get
                        Return mResults
                    End Get
                End Property

                Public Sub Interrogate(ByVal _ConflictSet As ConflictSet, ByVal _Strategy As String)
                    Select Case _Strategy
                        Case "LEX:"
                        'LEX: Pick the rule that passes the most amounts of required tests
                        Case "RECENCY"
                        'RECENCY: Pick the rule that is the most common or Recent
                        Case "MeanEndsAynalasis"
                            'Means End Analysis : Combine rules (Lex & Recency)
                    End Select
                End Sub

            End Class

        End Class

    End Namespace

    ''' <summary>
    ''' Abstract Class: Publisher / Subscriber DesignPattern subscribers implement this
    ''' interface to receive notifications from the publisher
    ''' </summary>
    Public Interface Observer

#Region "Public Methods"

        ''' <summary>
        ''' This is the channel to receive data from the publisher this variable needs to match
        ''' the data being updated from the publisher
        ''' </summary>
        ''' <param name="Data"></param>
        Sub Update(ByVal Data As String)

#End Region

    End Interface

    'Publisher & Subscriber DesignPattern
    'Author : Leroy S Dyer
    'Year : 2017
    ''' <summary>
    ''' Abstract Class: Publisher / Subscriber DesignPattern this interface is used as the main
    ''' task manager for the subscriber classes the subject is the information provider or Publisher
    ''' </summary>
    Public Interface Publisher

#Region "Public Methods"

        Sub NotifyObservers()

        Sub RegisterObserver(ByVal mObserver As Observer)

        Sub RemoveObserver(ByVal mObserver As Observer)

#End Region

    End Interface

    ''' <summary>
    ''' Abstract Class: Mediator Implements Both Subscriber and publisher patterns
    '''
    ''' </summary>
    Public Class Mediator
        Implements Observer
        Implements Publisher

#Region "Public Methods"

        Public Sub NotifyObservers() Implements Publisher.NotifyObservers
            Throw New NotImplementedException()
        End Sub

        Public Sub RegisterObserver(mObserver As Observer) Implements Publisher.RegisterObserver
            Throw New NotImplementedException()
        End Sub

        Public Sub RemoveObserver(mObserver As Observer) Implements Publisher.RemoveObserver
            Throw New NotImplementedException()
        End Sub

        Public Sub Update(Data As String) Implements Observer.Update
            Throw New NotImplementedException()
        End Sub

#End Region

    End Class

End Namespace

Namespace Web

    Public Structure Bookmark
        Dim Name As String
        Dim Url As String
    End Structure

    Public Class GoogleMaps

        'Data query is to be returned to sender for use in a WebBrowser
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

    Public Module WebSearch
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

Namespace AI_SDK_EXTENSIONS

    Namespace Functions
        Public Module iFileExtensions

            Public Sub SaveTextFile(ByRef FileName As String, ByRef Text As String)
                SaveTextFileAs(FileName, Text)
            End Sub

            Public Function GetDirSubDirectorys(ByVal directory As IO.DirectoryInfo, ByVal pattern As String) As List(Of String)
                Dim Files As New List(Of String)
                For Each file In directory.GetFiles(pattern)
                    Console.WriteLine(file.FullName)
                    Files.Add(file.FullName)
                Next
                For Each subDir In directory.GetDirectories
                    GetDirSubDirectorys(subDir, pattern)
                Next
                Return Files
            End Function

            ''' <summary>
            ''' Loads Json file to new string
            ''' </summary>
            ''' <returns></returns>
            Public Function LoadJson() As String

                Dim Scriptfile As String = ""
                Dim Ofile As New OpenFileDialog
                With Ofile
                    .Filter = "Json files (*.Json)|*.Json"

                    If (.ShowDialog() = DialogResult.OK) Then
                        Scriptfile = .FileName
                    End If
                End With
                Dim txt As String = ""
                If Scriptfile IsNot "" Then

                    txt = My.Computer.FileSystem.ReadAllText(Scriptfile)
                    Return txt
                Else
                    Return txt
                End If

            End Function

            ''' <summary>
            ''' Writes the contents of an embedded file resource embedded as Bytes to disk.
            ''' EG: My.Resources.DefBrainConcepts.FileSave(Application.StartupPath and "\DefBrainConcepts.ACCDB")
            ''' </summary>
            ''' <param name="BytesToWrite">Embedded resource Name</param>
            ''' <param name="FileName">Save to file</param>
            ''' <remarks></remarks>
            <System.Runtime.CompilerServices.Extension()>
            Public Sub FileSave(ByVal BytesToWrite() As Byte, ByVal FileName As String)

                If IO.File.Exists(FileName) Then
                    IO.File.Delete(FileName)
                End If

                Dim FileStream As New System.IO.FileStream(FileName, System.IO.FileMode.OpenOrCreate)
                Dim BinaryWriter As New System.IO.BinaryWriter(FileStream)

                BinaryWriter.Write(BytesToWrite)
                BinaryWriter.Close()
                FileStream.Close()
            End Sub

            <Runtime.CompilerServices.Extension()>
            Public Sub AppendTextFile(ByRef Text As String, ByRef FileName As String)

                UpdateTextFileAs(FileName, Text)
            End Sub

            ''' <summary>
            ''' replaces text in file with supplied text
            ''' </summary>
            ''' <param name="PathName">Pathname of file to be updated</param>
            ''' <param name="_text">Text to be inserted</param>
            <System.Runtime.CompilerServices.Extension()>
            Public Sub UpdateTextFileAs(ByRef PathName As String, ByRef _text As String)
                Try
                    If File.Exists(PathName) = True Then File.Create(PathName).Dispose()
                    Dim alltext As String = _text
                    File.AppendAllText(PathName, alltext)
                Catch ex As Exception
                    MsgBox("Error: " & Err.Number & ". " & Err.Description, , NameOf(UpdateTextFileAs))
                End Try
            End Sub

            ''' <summary>
            ''' Creates saves text file as
            ''' </summary>
            ''' <param name="PathName">nfilename and path to file</param>
            ''' <param name="_text">text to be inserted</param>
            <System.Runtime.CompilerServices.Extension()>
            Public Sub SaveTextFileAs(ByRef PathName As String, ByRef _text As String)

                Try
                    If File.Exists(PathName) = True Then File.Create(PathName).Dispose()
                    Dim alltext As String = _text
                    File.WriteAllText(PathName, alltext)
                Catch ex As Exception
                    MsgBox("Error: " & Err.Number & ". " & Err.Description, , NameOf(SaveTextFileAs))
                End Try
            End Sub

            ''' <summary>
            ''' Opens text file and returns text
            ''' </summary>
            ''' <param name="PathName">URL of file</param>
            ''' <returns>text in file</returns>
            <System.Runtime.CompilerServices.Extension()>
            Public Function OpenTextFile(ByRef PathName As String) As String
                Dim _text As String = ""
                Try
                    If File.Exists(PathName) = True Then
                        _text = File.ReadAllText(PathName)
                    End If
                Catch ex As Exception
                    MsgBox("Error: " & Err.Number & ". " & Err.Description, , NameOf(SaveTextFileAs))

                End Try
                Return _text
            End Function

        End Module

        Public Class ObjectSerializer

#Region "Public Fields"

            '<ComClass(SDK_Propositional.ClassId, SDK_Propositional.InterfaceId, SDK_Propositional.EventsId)>
            Public Const ClassId As String = "2828E720-7705-401C-BAB3-38FBA7BC1AC9"

            Public Const EventsId As String = "CDB77327-FA5E-401A-ACAD-3CF80B6BD6F1"
            Public Const InterfaceId As String = "8B7A55F8-5D13-4059-82AB-B53131B14BB5"

#End Region

#Region "Public Methods"

            ''' <summary>
            ''' Used to extract Object from XML
            ''' </summary>
            ''' <param name="FileName">XML FILENAME</param>
            ''' <param name="myObj">OBJECT RETURNED</param>
            ''' <returns>True If Sucessfull</returns>
            Public Shared Function ObjectDeserializer(ByRef FileName As String, ByRef myObj As Object) As Boolean
                Try

                    Dim ObjItem As Object
                    Dim string_reader As New StringReader(FileName)
                    Dim xml_serializer As New XmlSerializer(myObj.GetType)
                    ObjItem = xml_serializer.Deserialize(string_reader)
                    ObjectDeserializer = True
                Catch ex As Exception
                    ObjectDeserializer = False
                End Try
            End Function

            ''' <summary>
            ''' Serializes object to XML
            ''' </summary>
            ''' <param name="myObj">Presented Object</param>
            ''' <returns>True if Sucessfull</returns>
            Public Shared Function ObjectSerializer(ByRef myObj As Object) As Boolean
                Try

                    Dim string_writer As New StringWriter()
                    Dim xml_serializer As New Xml.Serialization.XmlSerializer(myObj.GetType)
                    xml_serializer.Serialize(string_writer, myObj)
                    ObjectSerializer = True
                Catch ex As Exception
                    ObjectSerializer = False
                End Try
            End Function

            ''' <summary>
            ''' Serialize object
            ''' </summary>
            ''' <param name="Item"></param>
            ''' <returns></returns>
            Public Shared Function ToJson(ByRef Item As Object) As String
                Dim Converter As New JavaScriptSerializer
                Return Converter.Serialize(Item)

            End Function

#End Region

        End Class

        ''' <summary>
        ''' Colors Words in RichText Box
        ''' </summary>
        Public Class SyntaxHighlighter

#Region "Public Fields"

            Public Shared SyntaxTerms() As String = ({"SPYDAZ", "ABS", "ACCESS", "ADDITEM", "ADDNEW", "ALIAS", "AND", "ANY", "APP", "APPACTIVATE", "APPEND", "APPENDCHUNK", "ARRANGE", "AS", "ASC", "ATN", "BASE", "BEEP", "BEGINTRANS", "BINARY", "BYVAL", "CALL", "CASE", "CCUR", "CDBL", "CHDIR", "CHDRIVE", "CHR", "CHR$", "CINT", "CIRCLE", "CLEAR", "CLIPBOARD", "CLNG", "CLOSE", "CLS", "COMMAND", "
COMMAND$", "COMMITTRANS", "COMPARE", "CONST", "CONTROL", "CONTROLS", "COS", "CREATEDYNASET", "CSNG", "CSTR", "CURDIR$", "CURRENCY", "CVAR", "CVDATE", "DATA", "DATE", "DATE$", "DATESERIAL", "DATEVALUE", "DAY", "
DEBUG", "DECLARE", "DEFCUR", "CEFDBL", "DEFINT", "DEFLNG", "DEFSNG", "DEFSTR", "DEFVAR", "DELETE", "DIM", "DIR", "DIR$", "DO", "DOEVENTS", "DOUBLE", "DRAG", "DYNASET", "EDIT", "ELSE", "ELSEIF", "END", "ENDDOC", "ENDIF", "
ENVIRON$", "EOF", "EQV", "ERASE", "ERL", "ERR", "ERROR", "ERROR$", "EXECUTESQL", "EXIT", "EXP", "EXPLICIT", "FALSE", "FIELDSIZE", "FILEATTR", "FILECOPY", "FILEDATETIME", "FILELEN", "FIX", "FOR", "FORM", "FORMAT", "
FORMAT$", "FORMS", "FREEFILE", "FUNCTION", "GET", "GETATTR", "GETCHUNK", "GETDATA", "DETFORMAT", "GETTEXT", "GLOBAL", "GOSUB", "GOTO", "HEX", "HEX$", "HIDE", "HOUR", "IF", "IMP", "INPUT", "INPUT$", "INPUTBOX", "INPUTBOX$", "
INSTR", "INT", "INTEGER", "IS", "ISDATE", "ISEMPTY", "ISNULL", "ISNUMERIC", "KILL", "LBOUND", "LCASE", "LCASE$", "LEFT", "LEFT$", "LEN", "LET", "LIB", "LIKE", "LINE", "LINKEXECUTE", "LINKPOKE", "LINKREQUEST", "
LINKSEND", "LOAD", "LOADPICTURE", "LOC", "LOCAL", "LOCK", "LOF", "LOG", "LONG", "LOOP", "LSET", "LTRIM",
                        "LTRIM$", "ME", "MID", "MID$", "MINUTE", "MKDIR", "MOD", "MONTH", "MOVE", "MOVEFIRST", "MOVELAST", "MOVENEXT", "MOVEPREVIOUS",
                        "MOVERELATIVE", "MSGBOX", "NAME", "NEW", "NEWPAGE", "NEXT", "NEXTBLOCK", "NOT", "NOTHING",
                        "NOW", "NULL", "OCT", "OCT$", "ON", "OPEN", "OPENDATABASE", "OPTION", "OR", "OUTPUT", "POINT", "PRESERVE", "PRINT",
                        "PRINTER", "PRINTFORM", "PRIVATE", "PSET", "PUT", "PUBLIC", "QBCOLOR", "RANDOM", "RANDOMIZE", "READ", "REDIM", "REFRESH",
                        "REGISTERDATABASE", "REM", "REMOVEITEM", "RESET", "RESTORE", "RESUME", "RETURN", "RGB", "RIGHT", "RIGHT$", "RMDIR", "RND",
                        "ROLLBACK", "RSET", "RTRIM", "RTRIM$", "SAVEPICTURE", "SCALE", "SECOND", "SEEK", "SELECT", "SENDKEYS", "SET", "SETATTR",
                        "SETDATA", "SETFOCUS", "SETTEXT", "SGN", "SHARED",
                        "SHELL", "SHOW", "SIN", "SINGLE", "SPACE", "SPACE$", "SPC", "SQR",
                        "STATIC", "STEP", "STOP", "STR", "STR$", "STRCOMP", "STRING", "STRING$", "SUB",
                        "SYSTEM", "TAB", "TAN", "TEXT", "TEXTHEIGHT", "TEXTWIDTH", "THEN", "TIME", "TIME$",
                        "TIMER", "TIMESERIAL", "TIMEVALUE", "TO", "TRIM",
                        "TRIM$", "TRUE", "TYPE", "TYPEOF", "UBOUND", "UCASE", "UCASE$", "UNLOAD",
                        "UNLOCK", "UNTIL", "UPDATE", "USING", "VAL", "VARIANT", "VARTYPE", "WEEKDAY", "WEND", "WHILE", "WIDTH",
                        "WRITE", "XOR", "YEAR", "ZORDER", "ADDHANDLER", "ADDRESSOF", "ALIAS", "AND", "ANDALSO", "AS", "BYREF",
                        "BOOLEAN", "BYTE", "BYVAL", "CALL", "CASE", "CATCH", "CBOOL", "CBYTE", "CCHAR", "CDATE",
                        "CDEC", "CDBL", "CHAR", "CINT", "CLASS", "CLNG", "COBJ", "CONST", "CONTINUE", "CSBYTE",
                        "CSHORT", "CSNG", "CSTR", "CTYPE", "CUINT", "CULNG", "CUSHORT", "DATE", "DECIMAL", "DECLARE",
                        "DEFAULT", "DELEGATE", "DIM", "DIRECTCAST", "DOUBLE", "DO", "EACH", "ELSE", "ELSEIF", "END",
                        "ENDIF", "ENUM", "ERASE", "ERROR", "EVENT", "EXIT", "FALSE", "FINALLY", "FOR", "FRIEND",
                        "FUNCTION", "GET", "GETTYPE", "GLOBAL", "GOSUB", "GOTO", "HANDLES", "IF", "IMPLEMENTS",
                        "IMPORTS", "IN", "INHERITS", "INTEGER", "INTERFACE", "IS", "ISNOT", "LET", "LIB", "LIKE",
                        "LONG", "LOOP", "ME", "MOD", "MODULE", "MUSTINHERIT", "MUSTOVERRIDE", "MYBASE", "MYCLASS",
                        "NAMESPACE", "NARROWING", "NEW", "NEXT", "NOT", "NOTHING", "NOTINHERITABLE", "NOTOVERRIDABLE",
                        "OBJECT", "ON", "OF", "OPERATOR", "OPTION", "OPTIONAL", "OR", "ORELSE", "OVERLOADS",
                        "OVERRIDABLE", "OVERRIDES", "PARAMARRAY", "PARTIAL", "PRIVATE", "PROPERTY", "PROTECTED",
                        "PUBLIC", "RAISEEVENT", "READONLY", "REDIM", "REM", "REMOVEHANDLER", "RESUME", "RETURN",
                        "SBYTE", "SELECT", "SET", "SHADOWS", "SHARED", "SHORT", "SINGLE", "STATIC", "STEP", "STOP",
                        "STRING", "STRUCTURE", "SUB", "SYNCLOCK", "THEN", "THROW", "TO", "TRUE", "TRY", "TRYCAST",
                        "TYPEOF", "WEND", "VARIANT", "UINTEGER", "ULONG", "USHORT", "USING", "WHEN", "WHILE", "WIDENING",
                        "WITH", "WITHEVENTS", "WRITEONLY",
                        "XOR", "#CONST", "#ELSE", "#ELSEIF", "#END", "#IF"})

#End Region

#Region "Private Fields"

            Private Shared indexOfSearchText As Integer = 0

            Private Shared start As Integer = 0

            Private mGrammar As New List(Of String)

#End Region

#Region "Public Methods"

            Public Shared Sub ColorSearchTerm(ByRef SearchStr As String, Rtb As RichTextBox)
                ColorSearchTerm(SearchStr, Rtb, Color.CadetBlue)
            End Sub

            Public Shared Sub ColorSearchTerm(ByRef SearchStr As String, Rtb As RichTextBox, ByRef MyColor As Color)
                Dim startindex As Integer = 0
                start = 0
                indexOfSearchText = 0

                If SearchStr <> "" Then

                    SearchStr = SearchStr & " "
                    If SearchStr.Length > 0 Then
                        startindex = FindText(Rtb, ProperCase(SearchStr), start, Rtb.Text.Length)
                    End If
                    If SearchStr.Length > 0 And startindex = 0 Then
                        startindex = FindText(Rtb, LCase(SearchStr), start, Rtb.Text.Length)
                    End If
                    If SearchStr.Length > 0 And startindex = 0 Then
                        startindex = FindText(Rtb, UCase(SearchStr), start, Rtb.Text.Length)
                    End If
                    If SearchStr.Length > 0 And startindex = 0 Then
                        startindex = FindText(Rtb, SearchStr, start, Rtb.Text.Length)
                    End If
                    ' If string was found in the RichTextBox, highlight it
                    If startindex >= 0 Then
                        ' Set the highlight color as red
                        Rtb.SelectionColor = MyColor

                        ' Find the end index. End Index = number of characters in textbox
                        Dim endindex As Integer = SearchStr.Length
                        ' Highlight the search string

                        Rtb.Select(startindex, endindex)
                        Rtb.SelectedText.ToUpper()
                        ' mark the start position after the position of last search string
                        start = startindex + endindex

                    End If
                Else
                End If
                Rtb.Select(Rtb.TextLength, Rtb.TextLength)
            End Sub

            Public Shared Sub HighlightInternalSyntax(ByRef sender As RichTextBox)

                For Each Word As String In SyntaxTerms
                    HighlightWord(sender, Word)
                    HighlightWord(sender, LCase(Word))
                    HighlightWord(sender, ProperCase(Word))
                Next

            End Sub

            Public Shared Sub HighlightWord(ByRef sender As RichTextBox, ByRef word As String)

                Dim index As Integer = 0
                While index < sender.Text.LastIndexOf(word)
                    'find
                    sender.Find(word, index, sender.TextLength, RichTextBoxFinds.WholeWord)
                    'select and color
                    sender.SelectionColor = Color.OrangeRed
                    index = sender.Text.IndexOf(word, index) + 1
                End While
            End Sub

            ''' <summary>
            ''' Searches For Internal Syntax
            ''' </summary>
            ''' <param name="Rtb"></param>
            ''' <remarks></remarks>
            Public Shared Sub SearchSyntax(ByRef Rtb As RichTextBox)
                'Searches Basic Syntax
                For Each wrd As String In SyntaxTerms
                    start = 0
                    indexOfSearchText = 0
                    ColorSearchTerm(wrd, Rtb)
                Next
                For Each wrd As String In SyntaxTerms
                    start = 0
                    indexOfSearchText = 0
                    ColorSearchTerm(wrd, Rtb)
                Next
            End Sub

            ''' <summary>
            ''' Searches For Internal Syntax
            ''' </summary>
            ''' <param name="Rtb"></param>
            ''' <remarks></remarks>
            Public Shared Sub SearchSyntax(ByRef Rtb As RichTextBox, ByRef Terms As List(Of String))
                'Searches Basic Syntax
                For Each wrd As String In SyntaxTerms
                    start = 0
                    indexOfSearchText = 0
                    ColorSearchTerm(wrd, Rtb)
                Next
                For Each wrd As String In Terms
                    start = 0
                    indexOfSearchText = 0
                    ColorSearchTerm(UCase(wrd), Rtb)
                Next
            End Sub

            ''' <summary>
            ''' Searches for Specific Word to colorize (Blue)
            ''' </summary>
            ''' <param name="Rtb"></param>
            ''' <param name="SearchStr"></param>
            ''' <remarks></remarks>
            Public Shared Sub SearchSyntax(ByRef Rtb As RichTextBox, ByRef SearchStr As String)
                start = 0
                indexOfSearchText = 0
                ColorSearchTerm(SearchStr, Rtb)
            End Sub

            ''' <summary>
            ''' Searches for Specfic word to colorize specified color
            ''' </summary>
            ''' <param name="Rtb"></param>
            ''' <param name="SearchStr"></param>
            ''' <param name="MyColor"></param>
            ''' <remarks></remarks>
            Public Shared Sub SearchSyntax(ByRef Rtb As RichTextBox, ByRef SearchStr As String, MyColor As Color)

                start = 0
                indexOfSearchText = 0
                ColorSearchTerm(SearchStr, Rtb, MyColor)

            End Sub

#End Region

#Region "Private Methods"

            Private Shared Function FindText(ByRef Rtb As RichTextBox, SearchStr As String,
                                                                            ByVal searchStart As Integer, ByVal searchEnd As Integer) As Integer

                ' Unselect the previously searched string
                If searchStart > 0 AndAlso searchEnd > 0 AndAlso indexOfSearchText >= 0 Then
                    Rtb.Undo()
                End If

                ' Set the return value to -1 by default.
                Dim retVal As Integer = -1

                ' A valid starting index should be specified. if indexOfSearchText = -1, the end of search
                If searchStart >= 0 AndAlso indexOfSearchText >= 0 Then

                    ' A valid ending index
                    If searchEnd > searchStart OrElse searchEnd = -1 Then

                        ' Find the position of search string in RichTextBox
                        indexOfSearchText = Rtb.Find(SearchStr, searchStart, searchEnd, RichTextBoxFinds.WholeWord)

                        ' Determine whether the text was found in richTextBox1.
                        If indexOfSearchText <> -1 Then
                            ' Return the index to the specified search text.
                            retVal = indexOfSearchText
                        End If

                    End If
                End If
                Return retVal

            End Function

            ''' <summary>
            ''' Returns Propercase Sentence
            ''' </summary>
            ''' <param name="TheString">String to be formatted</param>
            ''' <returns></returns>
            Private Shared Function ProperCase(ByRef TheString As String) As String
                ProperCase = UCase(Left(TheString, 1))

                For i = 2 To Len(TheString)

                    ProperCase = If(Mid(TheString, i - 1, 1) = " ", ProperCase & UCase(Mid(TheString, i, 1)), ProperCase & LCase(Mid(TheString, i, 1)))
                Next i
            End Function

#End Region

        End Class

        <ComClass(VbCompiler.ClassId, VbCompiler.InterfaceId, VbCompiler.EventsId)>
        Public Class VbCompiler

#Region "Public Fields"

            Public Const ClassId As String = "28CCC590-7703-401C-BAB3-38FF97BC1AC9"
            Public Const EventsId As String = "CCB5B677-F55E-401A-ABBD-3CF80C6BF6F1"
            Public Const InterfaceId As String = "8B3A25F8-5D13-4058-CB9B-B531C10C44B5"

#End Region

#Region "Private Fields"

            Private Const Lang As String = "VisualBasic"
            Private Shared _results As CodeDom.Compiler.CompilerResults
            Private Shared Compiler As CodeDom.Compiler.CodeDomProvider = CodeDom.Compiler.CodeDomProvider.CreateProvider(Lang)
            Private Shared InternalRefferences As List(Of String)
            Private Shared Settings As New CodeDom.Compiler.CompilerParameters
            Private _compiled As Boolean
            Private _EmbededFiles As List(Of String)
            Private _Refferences As List(Of String)
            Private _source, _entClass, _entMethod As String

#End Region

#Region "Public Constructors"

            Public Sub New(sourceCode As String, entClass As String, Optional entMethod As String = "Main",
                       Optional Assemblies As List(Of String) = Nothing,
                                         Optional ByVal EmbededFiles As List(Of String) = Nothing)
                _source = sourceCode
                _entClass = entClass
                _entMethod = entMethod
                _Refferences = Assemblies
                _EmbededFiles = EmbededFiles
            End Sub

#End Region

#Region "Public Properties"

            Public Shared ReadOnly Property Errors() As CodeDom.Compiler.CompilerErrorCollection
                Get
                    Return _results.Errors
                End Get
            End Property

#End Region

#Region "Public Methods"

            Public Shared Sub AddHelperRefferences()

                AddInternalReferencesHelper()
                AddProjectSpydazWebAIRefferencesHelper()

            End Sub

            Public Shared Sub AddInternalReferencesHelper()
                InternalRefferences = New List(Of String)
                InternalRefferences.Add(Application.StartupPath & "\System.Web.Extensions.dll")
                InternalRefferences.Add(Application.StartupPath & "\System.Windows.Forms.dll")
                InternalRefferences.Add(Application.StartupPath & "\Microsoft.VisualBasic.dll")
                InternalRefferences.Add(Application.StartupPath & "\System.dll")
                InternalRefferences.Add(Application.StartupPath & "\System.Speech.dll")
                If InternalRefferences IsNot Nothing And InternalRefferences.Count > 0 Then
                    For Each str As String In InternalRefferences
                        Settings.ReferencedAssemblies.Add(str)
                    Next
                End If
            End Sub

            Public Shared Sub AddProjectSpydazWebAIRefferencesHelper()
                InternalRefferences = New List(Of String)
                InternalRefferences.Add(Application.StartupPath & "\SpydazWebAI_ControlLibrary.dll")
                InternalRefferences.Add(Application.StartupPath & "\AI_AGENT.dll")
                InternalRefferences.Add(Application.StartupPath & "\AI_SDK.dll")
                InternalRefferences.Add(Application.StartupPath & "\DirectShowLib-2005.dll")
                If InternalRefferences IsNot Nothing And InternalRefferences.Count > 0 Then
                    For Each str As String In InternalRefferences
                        Settings.ReferencedAssemblies.Add(str)
                    Next
                End If
            End Sub

            Public Shared Function GetFilenameDLL() As String
                GetFilenameDLL = ""
                Dim S As New SaveFileDialog
                With S

                    .Filter = "Executable (*.Dll)|*.Dll"
                    If (.ShowDialog() = DialogResult.OK) Then
                        Return .FileName
                    End If
                End With
            End Function

            Public Shared Function GetFilenameEXE() As String
                GetFilenameEXE = ""
                Dim S As New SaveFileDialog
                With S

                    .Filter = "Executable (*.exe)|*.exe"
                    If (.ShowDialog() = DialogResult.OK) Then
                        Return .FileName
                    End If
                End With
            End Function

            Public Shared Sub RunInteractive(ByRef CodeBlock As String, Optional ByVal iClassName As String = "MainClass",
                                          Optional iMethodName As String = "Execute", Optional Assemblies As List(Of String) = Nothing,
                                         Optional ByVal EmbededFiles As List(Of String) = Nothing)
                Try

                    VB_CodeCompilerAsync(CodeBlock, "MEM", iClassName, iMethodName, Assemblies, EmbededFiles)
                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try

            End Sub

            Public Shared Sub VB_CodeCompiler(CodeBlock As String,
                                          Optional ByVal CompileType As String = "MEM",
                                          Optional ByVal iClassName As String = "MainClass",
                                          Optional iMethodName As String = "Execute",
                                          Optional Assemblies As List(Of String) = Nothing,
                                         Optional ByVal EmbededFiles As List(Of String) = Nothing)

                Dim cmp As New VbCompiler(CodeBlock, iClassName, iMethodName)
                cmp.AddEmbeddedFiles(EmbededFiles)
                cmp.AddRefferences(Assemblies)
                Select Case CompileType
                    Case "EXE"
                        cmp.Compile_EXE(GetFilenameDLL)
                    Case "DLL"
                        cmp.Compile_DLL(GetFilenameDLL)
                    Case "MEM"
                        cmp.Compile_MEM()
                End Select
            End Sub

            Public Shared Async Sub VB_CodeCompilerAsync(CodeBlock As String,
                                          Optional ByVal CompileType As String = "MEM",
                                          Optional ByVal iClassName As String = "MainClass",
                                          Optional iMethodName As String = "Execute",
                                          Optional Assemblies As List(Of String) = Nothing,
                                         Optional ByVal EmbededFiles As List(Of String) = Nothing)
                Try

                    Await Task.Run(Sub() VB_CodeCompiler(CodeBlock, CompileType, iClassName, iMethodName, Assemblies, EmbededFiles))
                Catch ex As Exception
                End Try
            End Sub

            Public Function CompileDll() As Boolean
                AddEmbeddedFiles()
                AddRefferences()
                Compile_DLL(GetFilenameDLL)

                If getErrors.Count > 0 Then
                    _compiled = False
                    Return False
                Else
                    _compiled = True
                    Return True
                End If
            End Function

            Public Function CompileExE() As Boolean
                AddEmbeddedFiles()
                AddRefferences()
                Compile_EXE(GetFilenameEXE)
                If getErrors.Count > 0 Then
                    _compiled = False
                    Return False
                Else
                    _compiled = True
                    Return True
                End If
            End Function

            Public Function CompileMEM() As Boolean
                AddEmbeddedFiles()
                AddRefferences()
                Compile_MEM()
                If getErrors.Count > 0 Then
                    _compiled = False
                    Return False
                Else
                    _compiled = True
                    Return True
                End If
            End Function

            Public Function getErrors() As List(Of String)

#Region "Errors"

                Dim IntErrors As New List(Of String)
                'Determines if we have any errors when compiling if so loops through all of the CompileErrors in the Reults variable and displays their ErrorText property.
                If (_results.Errors.Count <> 0) Then

                    '  MessageBox.Show("Exception Occured!", "Whoops!", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    For Each E As CodeDom.Compiler.CompilerError In _results.Errors

                        IntErrors.Add(E.ErrorText)

                    Next
                Else
                End If

#End Region

                Return IntErrors
            End Function

#End Region

#Region "Private Methods"

            Private Sub AddEmbeddedFiles(Optional EmbededFiles As List(Of String) = Nothing)

                Try
                    'handle Embedded Resources
                    If EmbededFiles IsNot Nothing And EmbededFiles.Count > 0 Then
                        For Each str As String In EmbededFiles
                            Settings.EmbeddedResources.Add(str)
                        Next
                    End If
                    If _EmbededFiles IsNot Nothing And _EmbededFiles.Count > 0 Then
                        For Each str As String In _EmbededFiles
                            Settings.EmbeddedResources.Add(str)
                        Next
                    End If
                Catch ex As Exception

                End Try
            End Sub

            Private Sub AddRefferences(Optional Refferences As List(Of String) = Nothing)
                'No References Held from Ctor
                Try
                    If Refferences IsNot Nothing And Refferences.Count > 0 Then
                        For Each str As String In Refferences
                            Settings.ReferencedAssemblies.Add(str)
                        Next
                    End If
                    If _Refferences IsNot Nothing And _Refferences.Count > 0 Then
                        For Each str As String In _Refferences
                            Settings.ReferencedAssemblies.Add(str)
                        Next
                    End If
                Catch ex As Exception

                End Try

            End Sub

            Private Sub Compile_DLL(ExecuteableName As String)
                'Library type options : /target:library, /target:win, /target:winexe
                'Generates an executable instead of a class library.
                'Compiles in memory.
                Settings.GenerateInMemory = False
                Settings.GenerateExecutable = False
                Settings.CompilerOptions = " /target:library"
                'Set the assembly file name / path
                Settings.OutputAssembly = ExecuteableName
                'Read the documentation for the result again variable.
                'Calls the CompileAssemblyFromSource that will compile the specified source code using the parameters specified in the settings variable.
                _results = Compiler.CompileAssemblyFromSource(Settings, _source)
            End Sub

            Private Sub Compile_EXE(ExecuteableName As String)
                'Library type options : /target:library, /target:win, /target:winexe
                'Generates an executable instead of a class library.
                'Compiles in memory.
                Settings.GenerateInMemory = True
                Settings.GenerateExecutable = True
                Settings.CompilerOptions = " /target:winexe"
                'Set the assembly file name / path
                Settings.OutputAssembly = ExecuteableName
                'Read the documentation for the result again variable.
                'Calls the CompileAssemblyFromSource that will compile the specified source code using the parameters specified in the settings variable.
                _results = Compiler.CompileAssemblyFromSource(Settings, _source)
            End Sub

            Private Sub Compile_MEM()
                'Library type options : /target:library, /target:win, /target:winexe
                'Generates an executable instead of a class library.
                'Compiles in memory.
                Settings.GenerateInMemory = True
                'Read the documentation for the result again variable.
                'Calls the CompileAssemblyFromSource that will compile the specified source code using the parameters specified in the settings variable.
                _results = Compiler.CompileAssemblyFromSource(Settings, _source)
            End Sub

#End Region

        End Class

        Public Class WebRegEx

#Region "Public Fields"

            Public Shared DotALL As String = "+(.|\n)*"
            Public Shared HtmlTagA As String = "<.*?>"
            Public Shared HtmlTagL As String = "<w+>"
            Public Shared HtmlTagR As String = "<\w+>"
            Public Shared HttpUrl As String = "(https?://[0-9a-zA-Z\_/-/.]*)"
            Public Shared QuotedText As String = "("".+ "")"

#End Region

#Region "Private Fields"

            Dim Emoticon_string = "(?:[<>]?[:;=8][\-o\*\']?[\)\]\(\[dDpP/\:\}\{@\|\\]|[\)\]\(\[dDpP/\:\}\{@\|\\][\-o\*\']?[:;=8][<>]?|<[/\\]?3|\(?\(?\#?[>\-\^\*\+o\~][\_\.\|oO\,][<\-\^\*\+o\~][\#\;]?\)?\)?)"
            Dim httpGETinfo As String = "(?:\/\w+\?(?:\;?\w+\=\w+)+)"
            Dim PhoneNumbers As String = "(?:(?:\+?[01][\-\s.]*)?(?:[\(]?\d{3}[\-\s.\)]*)?\d{3}[\-\s.]*\d{4})"
            Dim remainingWordTypes As String = "(?:[a-z][a-z'\-_]+[a-z])|(?:[+\-]?\d+[,/.:-]\d+[+\-]?)|(?:[\w_]+)|(?:\.(?:\s*\.){1,})|(?:\S)"
            Dim TwitterHashTags As String = "(?:\#+[\w_]+[\w\'_\-]*[\w_]+)"
            Dim TwitterUSerName As String = "(?:@[\w_]+)"

#End Region

#Region "Public Methods"

            ''' <summary>
            ''' Searches for
            ''' Searches for (tag...........)
            ''' Searches for (...........Tag)
            ''' Searches for (.....Tag... ..)
            ''' Searches for (tag.........tag)
            ''' </summary>
            ''' <param name="Text"></param>
            ''' <param name="Tag"></param>
            ''' <returns></returns>
            Public Shared Function FindByTag(ByRef Text As String, ByRef Tag As String) As List(Of String)
                Dim result2 = New List(Of String)

                result2.AddRange(RegExFindTag(Text, Tag))
                result2.AddRange(RegExFindTagsBetweenBoth(Text, Tag))
                result2.AddRange(RegExFindTagsBetweenLeft(Text, Tag))
                result2.AddRange(RegExFindTagsBetweenRight(Text, Tag))

                Return result2
            End Function

            ''' <summary>
            ''' Find Between
            ''' Start ........
            ''' ....................
            ''' ....................Stop
            ''' </summary>
            ''' <param name="Text"></param>
            ''' <param name="StartStr"></param>
            ''' <param name="EndStr"></param>
            ''' <returns></returns>
            Public Shared Function RegExFindBetween(ByRef Text As String, ByRef StartStr As String, ByRef EndStr As String) As List(Of String)
                Dim Tag = StartStr & "+(.|\n)*" & EndStr
                Dim Searcher As New Regex(Tag)
                Dim iMatch As Match = Searcher.Match(Text)
                Dim iMatches As New List(Of String)
                Do While iMatch.Success
                    iMatches.Add(iMatch.Value)
                    iMatch = iMatch.NextMatch
                Loop
                Return iMatches
            End Function

            ''' <summary>
            ''' Using RegEx for Searching
            ''' </summary>
            ''' <param name="Text">TobeSearched</param>
            ''' <returns></returns>
            Public Shared Function RegExFindDigits(ByRef Text As String) As List(Of String)
                Dim Tag As String = "\d+"
                Return RegExSearch(Text, Tag)
            End Function

            Public Shared Function RegExFindHttpUrls(ByRef Text As String) As List(Of String)
                Dim Tag As String = "(https?://[0-9a-zA-Z\_/-/.]*)"
                Return RegExSearch(Text, Tag)
            End Function

            Public Shared Function RegExFindQuotedTexts(ByRef Text As String) As List(Of String)
                Dim Tag As String = "("".+ "")"
                Return RegExSearch(Text, Tag)
            End Function

            ''' <summary>
            ''' Searches for (....Tag....)
            ''' </summary>
            ''' <param name="Text"></param>
            ''' <param name="iTag"></param>
            ''' <returns></returns>
            Public Shared Function RegExFindTag(ByRef Text As String, ByRef iTag As String) As List(Of String)
                Dim Tag As String = "<.*" & iTag & ".*>"
                Return RegExSearch(Text, Tag)
            End Function

            ''' <summary>
            ''' Using RegEx for Searching Searches for (tag)
            ''' </summary>
            ''' <param name="Text">TobeSearched</param>
            ''' <returns></returns>
            Public Shared Function RegExFindTagLeft(ByRef Text As String) As List(Of String)
                Dim Tag As String = "<\w+>"
                Return RegExSearch(Text, Tag)
            End Function

            ''' <summary>
            ''' Searches for (tag......
            '''               ........
            '''               ........tag)
            '''
            ''' </summary>
            ''' <param name="Text"></param>
            ''' <param name="iTag"></param>
            ''' <returns></returns>
            Public Shared Function RegExFindTagsBetweenBoth(ByRef Text As String, ByRef iTag As String) As List(Of String)
                '  Dim Tag As String = "<" & iTag & ".*?/" & iTag & ">"
                Dim Tag As String = "<" & iTag & "+(.|\n)*" & iTag & ">"
                Return RegExSearch(Text, Tag)
            End Function

            ''' <summary>
            ''' Searches for (tag......)
            ''' </summary>
            ''' <param name="Text"></param>
            ''' <param name="iTag"></param>
            ''' <returns></returns>
            Public Shared Function RegExFindTagsBetweenLeft(ByRef Text As String, ByRef iTag As String) As List(Of String)
                Dim Tag As String = "<" & iTag & ".*?>"
                Return RegExSearch(Text, Tag)
            End Function

            ''' <summary>
            ''' Searches for (......Tag)
            ''' </summary>
            ''' <param name="Text"></param>
            ''' <param name="iTag"></param>
            ''' <returns></returns>
            Public Shared Function RegExFindTagsBetweenRight(ByRef Text As String, ByRef iTag As String) As List(Of String)
                Dim Tag As String = "<.*?" & iTag & ">"
                Return RegExSearch(Text, Tag)
            End Function

            ''' <summary>
            ''' Using RegEx for Searching Searches for (/tag)
            ''' </summary>
            ''' <param name="Text">TobeSearched</param>
            ''' <returns></returns>
            Public Shared Function RegExFindTagsRight(ByRef Text As String) As List(Of String)
                Dim Tag As String = "<\/\w+>"
                Return RegExSearch(Text, Tag)
            End Function

            ''' <summary>
            ''' Main Searcher
            ''' </summary>
            ''' <param name="Text">to be searched </param>
            ''' <param name="Pattern">RegEx Search String</param>
            ''' <returns></returns>
            Public Shared Function RegExSearch(ByRef Text As String, Pattern As String) As List(Of String)
                Dim Searcher As New Regex(Pattern)
                Dim iMatch As Match = Searcher.Match(Text)
                Dim iMatches As New List(Of String)
                Do While iMatch.Success
                    iMatches.Add(iMatch.Value)
                    iMatch = iMatch.NextMatch
                Loop

                Return iMatches
            End Function

            ''' <summary>
            ''' Removes just the outer braces from text
            ''' </summary>
            ''' <param name="html"></param>
            ''' <returns></returns>
            Public Shared Function RemoveBraces(ByVal html As String) As String
                ' Remove HTML tags.
                html = Regex.Replace(html, "<", "")
                html = Regex.Replace(html, ">", "")
                Return html
            End Function

            ''' <summary>
            ''' Strip HTML tags.Removes tags from text
            ''' </summary>
            Public Shared Function RemoveTags(ByVal html As String) As String
                ' Remove HTML tags.
                Return Regex.Replace(html, "<.*?>", "")
            End Function

            ''' <summary>
            ''' removes both sides of the tag from the text
            ''' </summary>
            ''' <param name="html">text to be cleands</param>
            ''' <param name="TagId"> ie Head</param>
            ''' <returns></returns>
            Public Shared Function RemoveTags(ByRef html As String, ByVal TagId As String) As String
                ' Remove HTML tags.
                html = Regex.Replace(html, "<" & TagId & ">", "")

                Return html
            End Function

#End Region

        End Class

    End Namespace

End Namespace

Public Module SyS

    <System.Runtime.CompilerServices.Extension()>
    Public Function ConvertToDataTable(Of T)(ByVal list As IList(Of T)) As DataTable
        Dim table As New DataTable()
        Dim fields() = GetType(T).GetFields()
        For Each field In fields
            table.Columns.Add(field.Name, field.FieldType)
        Next
        For Each item As T In list
            Dim row As DataRow = table.NewRow()
            For Each field In fields
                row(field.Name) = field.GetValue(item)
            Next
            table.Rows.Add(row)
        Next
        Return table
    End Function

    <System.Runtime.CompilerServices.Extension()>
    Public Function CreateDataTable(ByRef HeaderTitles As List(Of String)) As DataTable
        Dim DT As New DataTable
        For Each item In HeaderTitles
            DT.Columns.Add(item, GetType(String))
        Next
        Return DT
    End Function

    <System.Runtime.CompilerServices.Extension()>
    Public Function ExtractLinks(ByVal url As String) As DataTable
        Dim dt As New DataTable
        dt.Columns.Add("LinkText")
        dt.Columns.Add("LinkUrl")

        Dim wc As New WebClient
        Dim html As String = wc.DownloadString(url)

        Dim links As MatchCollection = Regex.Matches(html, "<a.*?href=""(.*?)"".*?>(.*?)</a>")

        For Each match As Match In links
            Dim dr As DataRow = dt.NewRow
            Dim matchUrl As String = match.Groups(1).Value
            'Ignore all anchor links
            If matchUrl.StartsWith("#") Then
                Continue For
            End If
            'Ignore all javascript calls
            If matchUrl.ToLower.StartsWith("javascript:") Then
                Continue For
            End If
            'Ignore all email links
            If matchUrl.ToLower.StartsWith("mailto:") Then
                Continue For
            End If
            'For internal links, build the url mapped to the base address
            If Not matchUrl.StartsWith("http://") And Not matchUrl.StartsWith("https://") Then
                matchUrl = MapUrl(url, matchUrl)
            End If
            'Add the link data to datatable
            dr("LinkUrl") = matchUrl
            dr("LinkText") = match.Groups(2).Value
            dt.Rows.Add(dr)
        Next

        Return dt
    End Function

    <System.Runtime.CompilerServices.Extension()>
    Public Function MapUrl(ByVal baseAddress As String, ByVal relativePath As String) As String

        Dim u As New System.Uri(baseAddress)

        If relativePath = "./" Then
            relativePath = "/"
        End If

        If relativePath.StartsWith("/") Then
            Return u.Scheme + Uri.SchemeDelimiter + u.Authority + relativePath
        Else
            Dim pathAndQuery As String = u.AbsolutePath
            ' If the baseAddress contains a file name, like ..../Something.aspx
            ' Trim off the file name
            pathAndQuery = pathAndQuery.Split("?")(0).TrimEnd("/")
            If pathAndQuery.Split("/")(pathAndQuery.Split("/").Count - 1).Contains(".") Then
                pathAndQuery = pathAndQuery.Substring(0, pathAndQuery.LastIndexOf("/"))
            End If
            baseAddress = u.Scheme + Uri.SchemeDelimiter + u.Authority + pathAndQuery

            'If the relativePath contains ../ then
            ' adjust the baseAddress accordingly

            While relativePath.StartsWith("../")
                relativePath = relativePath.Substring(3)
                If baseAddress.LastIndexOf("/") > baseAddress.IndexOf("//" + 2) Then
                    baseAddress = baseAddress.Substring(0, baseAddress.LastIndexOf("/")).TrimEnd("/")
                End If
            End While

            Return baseAddress + "/" + relativePath
        End If

    End Function

    <System.Runtime.CompilerServices.Extension()>
    Public Sub browserImageLoad_DocumentCompleted(ByVal sender As Object, ByVal e As WebBrowserDocumentCompletedEventArgs)
        Dim browser As WebBrowser = CType(sender, WebBrowser)
        Dim collection As HtmlElementCollection
        Dim imgListString As List(Of HtmlElement) = New List(Of HtmlElement)()

        If browser IsNot Nothing Then

            If browser.Document IsNot Nothing Then
                collection = browser.Document.GetElementsByTagName("img")

                If collection IsNot Nothing Then

                    For Each element As HtmlElement In collection
                        Dim wClient As WebClient = New WebClient()
                        Dim urlDownload As String = element.GetAttribute("src")
                        wClient.DownloadFile(urlDownload, urlDownload.Substring(urlDownload.LastIndexOf("/"c)))
                    Next
                End If
            End If
        End If
    End Sub

    <System.Runtime.CompilerServices.Extension()>
    Public Sub ExportDatasetToCsv(ByVal MyDataSet As DataSet, ByRef filname As String)
        Dim myWriter As New System.IO.StreamWriter(filname)
        Try
            'Declaration of Variables
            Dim dt As DataTable
            '     Dim dr As DataRow
            Dim myString As String = ""
            Dim bFirstRecord As Boolean = True
            myString = ""
            For Each dt In MyDataSet.Tables
                For Each dr As DataRow In dt.Rows
                    bFirstRecord = True
                    For Each field As Object In dr.ItemArray

                        If Not bFirstRecord Then

                            myString &= ","

                        End If

                        myString &= field.ToString

                        bFirstRecord = False

                    Next
                    'New Line to differentiate next row
                    myString &= Environment.NewLine
                Next
            Next
            'Write the String to the Csv File
            myWriter.WriteLine(myString)
        Catch ex As Exception

            MsgBox(ex.Message)

        End Try
        'Clean up
        myWriter.Close()
    End Sub

    <System.Runtime.CompilerServices.Extension()>
    Public Function GetCheckedListBox(ByRef CB As CheckedListBox) As List(Of String)
        Dim str As New List(Of String)
        For Each item In CB.CheckedItems
            str.Add(item)
        Next
        Return str
    End Function

    <System.Runtime.CompilerServices.Extension()>
    Public Async Function GetCheckedListBoxAsync(ByVal CB As CheckedListBox) As Task(Of List(Of String))
        Dim str As List(Of String) = Await Task.Run(Function() GetCheckedListBox(CB))
        Return str
    End Function

    <System.Runtime.CompilerServices.Extension()>
    Public Function GetListBox(ByRef CB As ListBox) As List(Of String)
        Dim str As New List(Of String)
        For Each item In CB.Items
            str.Add(item)
        Next
        Return str
    End Function

    <System.Runtime.CompilerServices.Extension()>
    Public Async Function GetListBoxAsync(ByVal CB As ListBox) As Task(Of List(Of String))
        Dim str As List(Of String) = Await Task.Run(Function() GetListBox(CB))
        Return str
    End Function

    <System.Runtime.CompilerServices.Extension()>
    Public Sub LoadCheckedListBox(ByRef CB As CheckedListBox, ByRef LstStr As List(Of String))
        For Each item In LstStr
            CB.Items.Add(item)
        Next
    End Sub

    <System.Runtime.CompilerServices.Extension()>
    Public Async Sub LoadCheckedListBoxAsync(ByVal CB As CheckedListBox, ByVal LstStr As List(Of String))
        Await Task.Run(Sub() LoadCheckedListBox(CB, LstStr))
    End Sub

    <System.Runtime.CompilerServices.Extension()>
    Public Sub LoadListBox(ByRef CB As ListBox, ByRef LstStr As List(Of String))
        For Each item In LstStr
            CB.Items.Add(item)
        Next
    End Sub

    <System.Runtime.CompilerServices.Extension()>
    Public Async Sub LoadListBoxAsync(ByVal CB As ListBox, ByVal LstStr As List(Of String))
        Await Task.Run(Sub() LoadListBox(CB, LstStr))
    End Sub

    ''' <summary>
    ''' Writes the contents of an embedded resource embedded as Bytes to disk.
    ''' </summary>
    ''' <param name="BytesToWrite">Embedded resource</param>
    ''' <param name="FileName">    Save to file</param>
    ''' <remarks></remarks>
    <System.Runtime.CompilerServices.Extension()>
    Public Sub ResourceFileSave(ByVal BytesToWrite() As Byte, ByVal FileName As String)

        If IO.File.Exists(FileName) Then
            IO.File.Delete(FileName)
        End If

        Dim FileStream As New System.IO.FileStream(FileName, System.IO.FileMode.OpenOrCreate)
        Dim BinaryWriter As New System.IO.BinaryWriter(FileStream)

        BinaryWriter.Write(BytesToWrite)
        BinaryWriter.Close()
        FileStream.Close()
    End Sub

    Public ReadOnly QuestionTable As String = "Questions"
    Public ReadOnly SemanticPatternTable As String = "SemanticPatterns"
    Public ReadOnly CeptDataTable As String = "Cept_Data"
    Public ReadOnly NymListTable As String = "Nymlist"

    Public InstalledApplicationPath As String = Application.StartupPath

    Public StopWords As New List(Of String)

    'Morse Code
    Private MorseCode() As String = {".", "-"}

    Public Enum TextPreProcessingTasks
        Space_Punctuation
        To_Upper
        To_Lower
        Lemmatize_Text
        Tokenize_Characters
        Remove_Stop_Words
        Tokenize_Words
        Tokenize_Sentences
        Remove_Symbols
        Remove_Brackets
        Remove_Maths_Symbols
        Remove_Punctuation
        AlphaNumeric_Only
    End Enum

    ''' <summary>
    ''' Add full stop to end of String
    ''' </summary>
    ''' <param name="MESSAGE"></param>
    ''' <returns></returns>
    <System.Runtime.CompilerServices.Extension()>
    Public Function AddFullStop(ByRef MESSAGE As String) As String
        AddFullStop = MESSAGE
        If MESSAGE = "" Then Exit Function
        MESSAGE = Trim(MESSAGE)
        If MESSAGE Like "*." Then Exit Function
        AddFullStop = MESSAGE + "."
    End Function

    ''' <summary>
    ''' Adds string to end of string (no spaces)
    ''' </summary>
    ''' <param name="Str">base string</param>
    ''' <param name="Prefix">Add before (no spaces)</param>
    ''' <returns></returns>
    <System.Runtime.CompilerServices.Extension()>
    Public Function AddPrefix(ByRef Str As String, ByVal Prefix As String) As String
        Return Prefix & Str
    End Function

    ''' <summary>
    ''' Adds Suffix to String (No Spaces)
    ''' </summary>
    ''' <param name="Str">Base string</param>
    ''' <param name="Suffix">To be added After</param>
    ''' <returns></returns>
    <System.Runtime.CompilerServices.Extension()>
    Public Function AddSuffix(ByRef Str As String, ByVal Suffix As String) As String
        Return Str & Suffix
    End Function

    ''' <summary>
    ''' GO THROUGH EACH CHARACTER AND ' IF PUNCTUATION IE .!?,:'"; REPLACE WITH A SPACE ' IF ,
    ''' OR . THEN CHECK IF BETWEEN TWO NUMBERS, IF IT IS ' THEN LEAVE IT, ELSE REPLACE IT WITH A
    ''' SPACE '
    ''' </summary>
    ''' <param name="STRINPUT">String to be formatted</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.Runtime.CompilerServices.Extension()>
    Public Function AlphaNumericalOnly(ByRef STRINPUT As String) As String

        Dim A As Short
        For A = 1 To Len(STRINPUT)
            If Mid(STRINPUT, A, 1) = "." Or
Mid(STRINPUT, A, 1) = "!" Or
Mid(STRINPUT, A, 1) = "?" Or
Mid(STRINPUT, A, 1) = "," Or
Mid(STRINPUT, A, 1) = ":" Or
Mid(STRINPUT, A, 1) = "'" Or
Mid(STRINPUT, A, 1) = "[" Or
Mid(STRINPUT, A, 1) = """" Or
Mid(STRINPUT, A, 1) = ";" Then

                ' BEGIN CHECKING PERIODS AND COMMAS THAT ARE IN BETWEEN NUMBERS '
                If Mid(STRINPUT, A, 1) = "." Or Mid(STRINPUT, A, 1) = "," Then
                    If Not (A - 1 = 0 Or A = Len(STRINPUT)) Then
                        If Not (IsNumeric(Mid(STRINPUT, A - 1, 1)) Or IsNumeric(Mid(STRINPUT, A + 1, 1))) Then
                            STRINPUT = Mid(STRINPUT, 1, A - 1) & " " & Mid(STRINPUT, A + 1, Len(STRINPUT) - A)
                        End If
                    Else
                        STRINPUT = Mid(STRINPUT, 1, A - 1) & " " & Mid(STRINPUT, A + 1, Len(STRINPUT) - A)
                    End If
                Else
                    STRINPUT = Mid(STRINPUT, 1, A - 1) & " " & Mid(STRINPUT, A + 1, Len(STRINPUT) - A)
                End If

                ' END CHECKING PERIODS AND COMMAS IN BETWEEN NUMBERS '
            End If
        Next A
        ' RETURN PUNCTUATION STRIPPED STRING '
        AlphaNumericalOnly = STRINPUT.Replace("  ", " ")
    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function AlphaNumericOnly(ByRef txt As String) As String
        Dim NewText As String = ""
        Dim IsLetter As Boolean = False
        Dim IsNumerical As Boolean = False
        For Each chr As Char In txt
            IsNumerical = False
            IsLetter = False
            For Each item In AlphaBet
                If IsLetter = False Then
                    If chr.ToString = item Then
                        IsLetter = True
                    Else
                    End If
                End If
            Next
            'Check Numerical
            If IsLetter = False Then
                For Each item In Numerical
                    If IsNumerical = False Then
                        If chr.ToString = item Then
                            IsNumerical = True
                        Else
                        End If
                    End If
                Next
            Else
            End If
            If IsLetter = True Or IsNumerical = True Then
                NewText &= chr.ToString
            Else
                NewText &= " "
            End If
        Next
        NewText = NewText.Replace("  ", " ")
        Return NewText
    End Function

    <Runtime.CompilerServices.Extension()>
    Public Sub AppendTextFile(ByRef Text As String, ByRef FileName As String)

        UpdateTextFileAs(FileName, Text)
    End Sub

    ''' <summary>
    ''' COMMENTS : RETURNS THE INVERSE COSECANT OF THE SUPPLIED NUMBER '
    ''' PARAMETERS: DBLIN - VALUE TO CALCULATE ' RETURNS : INVERSE COSECANT AS A DOUBLE
    ''' </summary>
    ''' <param name="DBLIN"></param>
    ''' <returns>DBLIN - VALUE TO CALCULATE ' RETURNS : INVERSE COSECANT AS A DOUBLE</returns>
    <Runtime.CompilerServices.Extension()>
    Public Function ARCCOSECANT(ByVal DBLIN As Double) As Double

        ' '
        Const CDBLPI As Double = 3.14159265358979

        On Error GoTo PROC_ERR

        ARCCOSECANT = Math.Atan(DBLIN / Math.Sqrt(DBLIN * DBLIN - 1)) +
    (Math.Sign(DBLIN) - 1) * CDBLPI / 2

PROC_EXIT:
        Exit Function

PROC_ERR:
        MsgBox("ERROR: " & Err.Number & ". " & Err.Description, ,
    NameOf(ARCCOSECANT))
        Resume PROC_EXIT

    End Function

    'Math Functions
    ''' <summary>
    ''' COMMENTS: RETURNS THE ARC COSINE OF THE SUPPLIED NUMBER '
    ''' PARAMETERS: DBLIN -Number TO RUN ON ' RETURNS : ARC COSINE AS A DOUBLE
    ''' </summary>
    ''' <param name="DBLIN"></param>
    ''' <returns>DBLIN -Number TO RUN ON ' RETURNS : ARC COSINE AS A DOUBLE</returns>
    <Runtime.CompilerServices.Extension()>
    Public Function ARCCOSINE(ByVal DBLIN As Double) As Double

        Const CDBLPI As Double = 3.14159265358979

        On Error GoTo PROC_ERR

        Select Case DBLIN
            Case 1
                ARCCOSINE = 0

            Case -1
                ARCCOSINE = -CDBLPI

            Case Else
                ARCCOSINE = Math.Atan(DBLIN / Math.Sqrt(-DBLIN * DBLIN + 1)) + CDBLPI / 2

        End Select

PROC_EXIT:
        Exit Function

PROC_ERR:
        MsgBox("ERROR: " & Err.Number & ". " & Err.Description, ,
    NameOf(ARCCOSINE))
        Resume PROC_EXIT

    End Function

    ''' <summary>
    ''' COMMENTS: RETURNS THE INVERSE COTANGENT Of THE SUPPLIED NUMBER '
    ''' PARAMETERS: DBLIN -VALUE TO CALCULATE ' RETURNS : INVERSE COTANGENT AS A DOUBLE
    ''' </summary>
    ''' <param name="DBLIN"></param>
    ''' <returns>INVERSE COTANGENT AS A DOUBLE</returns>
    <Runtime.CompilerServices.Extension()>
    Public Function ARCCOTANGENT(ByVal DBLIN As Double) As Double

        Const CDBLPI As Double = 3.14159265358979

        On Error GoTo PROC_ERR

        ARCCOTANGENT = Math.Atan(DBLIN) + CDBLPI / 2

PROC_EXIT:
        Exit Function

PROC_ERR:
        MsgBox("ERROR: " & Err.Number & ". " & Err.Description, ,
    NameOf(ARCCOTANGENT))
        Resume PROC_EXIT

    End Function

    ''' <summary>
    ''' COMMENTS : RETURNS THE INVERSE SECANT OF THE SUPPLIED NUMBER '
    ''' PARAMETERS: DBLIN - VALUE TO CALCULATE ' RETURNS : INVERSE SECANT AS A DOUBLE ' '
    ''' </summary>
    ''' <param name="DBLIN"></param>
    ''' <returns>DBLIN - VALUE TO CALCULATE ' RETURNS : INVERSE SECANT AS A DOUBLE</returns>
    <Runtime.CompilerServices.Extension()>
    Public Function ARCSECANT(ByVal DBLIN As Double) As Double

        Const CDBLPI As Double = 3.14159265358979

        On Error GoTo PROC_ERR

        ARCSECANT = Math.Atan(DBLIN / Math.Sqrt(DBLIN * DBLIN - 1)) +
    Math.Sign(Math.Sign(DBLIN) - 1) * CDBLPI / 2

PROC_EXIT:
        Exit Function

PROC_ERR:
        MsgBox("ERROR: " & Err.Number & ". " & Err.Description, ,
    NameOf(ARCSECANT))
        Resume PROC_EXIT

    End Function

    ''' <summary>
    ''' COMMENTS : RETURNS THE INVERSE SINE OF THE SUPPLIED NUMBER '
    ''' PARAMETERS:  ' '
    ''' </summary>
    ''' <param name="DBLIN"></param>
    ''' <returns>DBLIN - VALUE TO CALCULATE ' RETURNS : INVERSE SINE AS A DOUBLE</returns>
    <Runtime.CompilerServices.Extension()>
    Public Function ARCSINE(ByVal DBLIN As Double) As Double

        Const CDBLPI As Double = 3.14159265358979

        On Error GoTo PROC_ERR

        Select Case DBLIN

            Case 1
                ARCSINE = CDBLPI / 2

            Case -1
                ARCSINE = -CDBLPI / 2

            Case Else
                ARCSINE = Math.Atan(DBLIN / Math.Sqrt(-DBLIN ^ 2 + 1))

        End Select

PROC_EXIT:
        Exit Function

PROC_ERR:
        MsgBox("ERROR: " & Err.Number & ". " & Err.Description, ,
    NameOf(ARCSINE))
        Resume PROC_EXIT

    End Function

    ''' <summary>
    ''' COMMENTS : RETURNS THE INVERSE TANGENT OF THE SUPPLIED NUMBERS. ' NOTE THAT BOTH VALUES
    ''' CANNOT BE ZERO. '
    ''' PARAMETERS: DBLIN - FIRST VALUE ' DBLIN2 - SECOND VALUE ' RETURNS : INVERSE TANGENT AS A
    '''             DOUBLE ' '
    ''' </summary>
    ''' <param name="DBLIN"></param>
    ''' <param name="DBLIN2"></param>
    ''' <returns>RETURNS : INVERSE TANGENT AS A DOUBLE</returns>
    <Runtime.CompilerServices.Extension()>
    Public Function ARCTANGENT(ByVal DBLIN As Double, ByVal DBLIN2 As Double) As Double

        Const CDBLPI As Double = 3.14159265358979

        On Error GoTo PROC_ERR

        Select Case DBLIN

            Case 0

                Select Case DBLIN2
                    Case 0
                        ' UNDEFINED '
                        ARCTANGENT = 0
                    Case Is > 0
                        ARCTANGENT = CDBLPI / 2
                    Case Else
                        ARCTANGENT = -CDBLPI / 2
                End Select

            Case Is > 0
                ARCTANGENT = If(DBLIN2 = 0, 0, Math.Atan(DBLIN2 / DBLIN))
            Case Else
                ARCTANGENT = If(DBLIN2 = 0, CDBLPI, (CDBLPI - Math.Atan(Math.Abs(DBLIN2 / DBLIN))) * Math.Sign(DBLIN2))
        End Select

PROC_EXIT:
        Exit Function

PROC_ERR:
        MsgBox("ERROR: " & Err.Number & ". " & Err.Description, ,
    NameOf(ARCTANGENT))
        Resume PROC_EXIT

    End Function

    ''' <summary>
    ''' Comments : Returns the area of a circle
    ''' </summary>
    ''' <param name="dblRadius">dblRadius - radius of circle</param>
    ''' <returns>Returns : area (Double)</returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function AreaOfCircle(ByVal dblRadius As Double) As Double
        Const PI = 3.14159265358979
        On Error GoTo PROC_ERR
        AreaOfCircle = PI * dblRadius ^ 2
PROC_EXIT:
        Exit Function
PROC_ERR:
        MsgBox("Error: " & Err.Number & ". " & Err.Description, ,
        NameOf(AreaOfCircle))
        Resume PROC_EXIT
    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function AreaOfElipse(ByRef Radius1 As Double, ByRef Radius2 As Double) As Double
        'Ellipse Formula :  Area of Ellipse = πr1r2
        'Case 1:
        'Find the area and perimeter of an ellipse with the given radii 3, 4.
        'Step 1:
        'Find the area.
        'Area = πr1r2 = 3.14 * 3 * 4 = 37.68 .
        AreaOfElipse = Math.PI * Radius1 * Radius2

    End Function

    ''' <summary>
    ''' Returns the area of a rectangle
    ''' </summary>
    ''' <param name="dblLength">dblLength - length of rectangle</param>
    ''' <param name="dblWidth">width of rectangle</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function AreaOfRectangle(
    ByVal dblLength As Double,
    ByVal dblWidth As Double) _
    As Double
        On Error GoTo PROC_ERR
        AreaOfRectangle = dblLength * dblWidth
PROC_EXIT:
        Exit Function
PROC_ERR:
        MsgBox("Error: " & Err.Number & ". " & Err.Description, ,
        NameOf(AreaOfRectangle))
        Resume PROC_EXIT
    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function AreaOFRhombusMethod1(ByRef base As Double, ByRef height As Double) As Double

        'Case 1:
        'Find the area of a rhombus with the given base 3 and height 4 using Base Times Height Method.
        'Step 1:
        'Find the area.
        'Area = b * h = 3 * 4 = 12.
        AreaOFRhombusMethod1 = base * height
    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function AreaOFRhombusMethod2(ByRef Diagonal1 As Double, ByRef Diagonal2 As Double) As Double
        'Case 2:

        'Find the area of a rhombus with the given diagonals 2, 4 using Diagonal Method.
        'Step 1:

        'Find the area.
        ' Area = ½ * d1 * d2 = 0.5 * 2 * 4 = 4.

        AreaOFRhombusMethod2 = 0.5 * Diagonal1 * Diagonal2
    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function AreaOFRhombusMethod3(ByRef Side As Double) As Double
        'Case 3:

        'Find the area of a rhombus with the given side 2 using Trigonometry Method.
        'Step 1:

        'Find the area.
        'Area = a² * SinA = 2² * Sin(33) = 4 * 1 = 4.

        AreaOFRhombusMethod3 = (Side * Side) * Math.Sin(33)
    End Function

    ''' <summary>
    ''' Returns the area of a ring
    ''' </summary>
    ''' <param name="dblInnerRadius">dblInnerRadius - inner radius of the ring</param>
    ''' <param name="dblOuterRadius">outer radius of the ring</param>
    ''' <returns>area</returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function AreaOfRing(
    ByVal dblInnerRadius As Double,
    ByVal dblOuterRadius As Double) _
    As Double
        On Error GoTo PROC_ERR

        AreaOfRing = AreaOfCircle(dblOuterRadius) -
        AreaOfCircle(dblInnerRadius)
PROC_EXIT:
        Exit Function
PROC_ERR:
        MsgBox("Error: " & Err.Number & ". " & Err.Description, ,
        NameOf(AreaOfRing))
        Resume PROC_EXIT
    End Function

    ''' <summary>
    ''' Returns the area of a sphere
    ''' </summary>
    ''' <param name="dblRadius">dblRadius - radius of the sphere</param>
    ''' <returns>area</returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function AreaOfSphere(ByVal dblRadius As Double) As Double
        Const cdblPI As Double = 3.14159265358979
        On Error GoTo PROC_ERR
        AreaOfSphere = 4 * cdblPI * dblRadius ^ 2
PROC_EXIT:
        Exit Function
PROC_ERR:
        MsgBox("Error: " & Err.Number & ". " & Err.Description, ,
        NameOf(AreaOfSphere))
        Resume PROC_EXIT
    End Function

    ''' <summary>
    ''' Returns the area of a square
    ''' </summary>
    ''' <param name="dblSide">dblSide - length of a side of the square</param>
    ''' <returns>area</returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function AreaOfSquare(ByVal dblSide As Double) As Double
        On Error GoTo PROC_ERR
        AreaOfSquare = dblSide ^ 2
PROC_EXIT:
        Exit Function
PROC_ERR:
        MsgBox("Error: " & Err.Number & ". " & Err.Description, ,
        NameOf(AreaOfSquare))
        Resume PROC_EXIT
    End Function

    ''' <summary>
    ''' Returns the area of a square
    ''' </summary>
    ''' <param name="dblDiag">dblDiag - length of the square's diagonal</param>
    ''' <returns>area</returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function AreaOfSquareDiag(ByVal dblDiag As Double) As Double
        On Error GoTo PROC_ERR
        AreaOfSquareDiag = (dblDiag ^ 2) / 2
PROC_EXIT:
        Exit Function
PROC_ERR:
        MsgBox("Error: " & Err.Number & ". " & Err.Description, ,
        NameOf(AreaOfSquareDiag))
        Resume PROC_EXIT
    End Function

    ''' <summary>
    ''' Returns the area of a trapezoid
    ''' </summary>
    ''' <param name="dblHeight">dblHeight - height</param>
    ''' <param name="dblLength1">length of first side</param>
    ''' <param name="dblLength2">length of second side</param>
    ''' <returns>area</returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function AreaOfTrapezoid(
    ByVal dblHeight As Double,
    ByVal dblLength1 As Double,
    ByVal dblLength2 As Double) _
    As Double
        On Error GoTo PROC_ERR
        AreaOfTrapezoid = dblHeight * (dblLength1 + dblLength2) / 2
PROC_EXIT:
        Exit Function
PROC_ERR:
        MsgBox("Error: " & Err.Number & ". " & Err.Description, ,
        NameOf(AreaOfTrapezoid))
        Resume PROC_EXIT
    End Function

    ''' <summary>
    ''' returns the area of a triangle
    ''' </summary>
    ''' <param name="dblLength">dblLength - length of a side</param>
    ''' <param name="dblHeight">perpendicular height</param>
    ''' <returns></returns>
    ''' <remarks>area</remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function AreaOfTriangle(
    ByVal dblLength As Double,
    ByVal dblHeight As Double) _
    As Double
        On Error GoTo PROC_ERR
        AreaOfTriangle = dblLength * dblHeight / 2
PROC_EXIT:
        Exit Function
PROC_ERR:
        MsgBox("Error: " & Err.Number & ". " & Err.Description, ,
        NameOf(AreaOfTriangle))
        Resume PROC_EXIT
    End Function

    ''' <summary>
    ''' </summary>
    ''' <param name="dblSideA">dblSideA - length of first side</param>
    ''' <param name="dblSideB">dblSideB - length of second side</param>
    ''' <param name="dblSideC">dblSideC - length of third side</param>
    ''' <returns>the area of a triangle</returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function AreaOfTriangle2(
    ByVal dblSideA As Double,
    ByVal dblSideB As Double,
    ByVal dblSideC As Double) As Double
        Dim dblCosine As Double
        On Error GoTo PROC_ERR
        dblCosine = (dblSideA + dblSideB + dblSideC) / 2
        AreaOfTriangle2 = Math.Sqrt(dblCosine * (dblCosine - dblSideA) *
        (dblCosine - dblSideB) *
        (dblCosine - dblSideC))
PROC_EXIT:
        Exit Function
PROC_ERR:
        MsgBox("Error: " & Err.Number & ". " & Err.Description, ,
        NameOf(AreaOfTriangle2))
        Resume PROC_EXIT
    End Function

    'Numerical
    <Runtime.CompilerServices.Extension()>
    Public Function ArithmeticMean(ByRef Elements() As Integer) As Double
        Dim NumberofElements As Integer = 0
        Dim sum As Integer = 0

        'Formula:
        'Mean = sum of elements / number of elements = a1+a2+a3+.....+an/n
        For Each value As Integer In Elements
            NumberofElements = NumberofElements + 1
            sum = value + value
        Next
        ArithmeticMean = sum / NumberofElements
    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function ArithmeticMedian(ByRef Elements() As Integer) As Double
        Dim NumberofElements As Integer = 0
        Dim Position As Integer = 0
        Dim P1 As Integer = 0
        Dim P2 As Integer = 0

        'Count the total numbers given.
        NumberofElements = Elements.Length
        'Arrange the numbers in ascending order.
        Array.Sort(Elements)

        'Formula:Calculate Middle Position

        'Check Odd Even
        If NumberofElements Mod 2 = 0 Then

            'Even:
            'For even average of number at P1 = n/2 and P2= (n/2)+1
            'Then: (P1+P2) / 2
            P1 = NumberofElements / 2
            P2 = (NumberofElements / 2) + 1
            ArithmeticMedian = (Elements(P1) + Elements(P2)) / 2
        Else

            'Odd:
            'For odd (NumberofElements+1)/2
            Position = (NumberofElements + 1) / 2
            ArithmeticMedian = Elements(Position)
        End If

    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function Average(ByVal x As List(Of Double)) As Double

        'Takes an average in absolute terms

        Dim result As Double

        For i = 0 To x.Count - 1
            result += x(i)
        Next

        Return result / x.Count

    End Function

    ''' <summary>
    ''' Calculating Average Growth Rate Over Regular Time Intervals
    ''' </summary>
    ''' <param name="Series"></param>
    ''' <returns></returns>
    <Runtime.CompilerServices.Extension()>
    Public Function AverageGrowth(ByRef Series As List(Of Double)) As Double
        'GrowthRate = Present / Past / Past
        ' formula: (present) = (past) * (1 + growth rate)n where n = number of time periods.

        'The Average Annual Growth Rate over a number Of years
        'means the average Of the annual growth rates over that number Of years.
        'For example, assume that In 2005, the price has increased over 2004 by 10%, 2004 over 2003 by 15%, And 2003 over 2002 by 5%,
        'then the average annual growth rate from 2002 To 2005 Is (10% + 15% + 5%) / 3 = 10%
        Dim NewSeries As New List(Of Double)
        For i = 1 To Series.Count
            'Calc Each Growth rate
            NewSeries.Add(Growth(Series(i - 1), Series(i)))
        Next
        Return Mean(NewSeries)
    End Function

    'Text
    <Runtime.CompilerServices.Extension()>
    Public Function Capitalise(ByRef MESSAGE As String) As String
        Dim FirstLetter As String
        Capitalise = ""
        If MESSAGE = "" Then Exit Function
        FirstLetter = Left(MESSAGE, 1)
        FirstLetter = UCase(FirstLetter)
        MESSAGE = Right(MESSAGE, Len(MESSAGE) - 1)
        Capitalise = (FirstLetter + MESSAGE)
    End Function

    ''' <summary>
    ''' Capitalizes the text
    ''' </summary>
    ''' <param name="MESSAGE"></param>
    ''' <returns></returns>
    <System.Runtime.CompilerServices.Extension()>
    Public Function CapitaliseTEXT(ByVal MESSAGE As String) As String
        Dim FirstLetter As String = ""
        CapitaliseTEXT = ""
        If MESSAGE = "" Then Exit Function
        FirstLetter = Left(MESSAGE, 1)
        FirstLetter = UCase(FirstLetter)
        MESSAGE = Right(MESSAGE, Len(MESSAGE) - 1)
        CapitaliseTEXT = (FirstLetter + MESSAGE)
    End Function

    ''' <summary>
    ''' Capitalise the first letter of each word / Tilte Case
    ''' </summary>
    ''' <param name="words">A string - paragraph or sentence</param>
    ''' <returns>String</returns>
    <Runtime.CompilerServices.Extension()>
    Public Function CapitalizeWords(ByVal words As String)
        Dim output As System.Text.StringBuilder = New System.Text.StringBuilder()
        Dim exploded = words.Split(" ")
        If (exploded IsNot Nothing) Then
            For Each word As String In exploded
                If word IsNot Nothing Then
                    output.Append(word.Substring(0, 1).ToUpper).Append(word.Substring(1, word.Length - 1)).Append(" ")
                End If

            Next
        End If

        Return output.ToString()

    End Function

    ''' <summary>
    ''' converts charactert to Morse code
    ''' </summary>
    ''' <param name="Ch"></param>
    ''' <returns></returns>
    <Runtime.CompilerServices.Extension()>
    Public Function CharToMorse(ByRef Ch As String) As String
        Select Case Ch
            Case "A", "a"
                CharToMorse = ".-"
            Case "B", "b"
                CharToMorse = "-..."
            Case "C", "c"
                CharToMorse = "-.-."
            Case "D", "d"
                CharToMorse = "-.."
            Case "E", "e"
                CharToMorse = "."
            Case "F", "f"
                CharToMorse = "..-."
            Case "G", "g"
                CharToMorse = "--."
            Case "H", "h"
                CharToMorse = "...."
            Case "I", "i"
                CharToMorse = ".."
            Case "J", "j"
                CharToMorse = ".---"
            Case "K", "k"
                CharToMorse = "-.-"
            Case "L", "l"
                CharToMorse = ".-.."
            Case "M", "m"
                CharToMorse = "--"
            Case "N", "n"
                CharToMorse = "-."
            Case "O", "o"
                CharToMorse = "---"
            Case "P", "p"
                CharToMorse = ".--."
            Case "Q", "q"
                CharToMorse = "--.-"
            Case "R", "r"
                CharToMorse = ".-."
            Case "S", "s"
                CharToMorse = "..."
            Case "T", "t"
                CharToMorse = "-"
            Case "U", "u"
                CharToMorse = "..-"
            Case "V", "v"
                CharToMorse = "...-"
            Case "W", "w"
                CharToMorse = ".--"
            Case "X", "x"
                CharToMorse = "-..-"
            Case "Y", "y"
                CharToMorse = "-.--"
            Case "Z", "z"
                CharToMorse = "--.."
            Case "1"
                CharToMorse = ".----"
            Case "2"
                CharToMorse = "..---"
            Case "3"
                CharToMorse = "...--"
            Case "4"
                CharToMorse = "....-"
            Case "5"
                CharToMorse = "....."
            Case "6"
                CharToMorse = "-...."
            Case "7"
                CharToMorse = "--..."
            Case "8"
                CharToMorse = "---.."
            Case "9"
                CharToMorse = "----."
            Case "0"
                CharToMorse = "-----"
            Case " "
                CharToMorse = "   "
            Case "."
                CharToMorse = "^"
            Case "-"
                CharToMorse = "~"
            Case Else
                CharToMorse = Ch
        End Select
    End Function

    ''' <summary>
    ''' Checks if directory exists If it does not then it is created
    ''' </summary>
    ''' <param name="YourPath"></param>
    Public Sub ChkDir(ByRef YourPath As String)

        If (System.IO.Directory.Exists(YourPath)) = True Then
        Else

            System.IO.Directory.CreateDirectory(YourPath)

        End If
    End Sub

    ''' <summary>
    ''' FUNCTION: CELSIUSTOFAHRENHEIT '
    ''' DESCRIPTION: CONVERTS CELSIUS DEGREES TO FAHRENHEIT DEGREES ' WHERE TO PLACE CODE:
    '''              MODULE '
    ''' NOTES: THE LARGEST NUMBER CELSIUSTOFAHRENHEIT WILL CONVERT IS 32,767
    ''' </summary>
    ''' <param name="intCelsius"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function CnvCelsiusToFahrenheit(intCelsius As Integer) As Integer
        CnvCelsiusToFahrenheit = (9 / 5) * intCelsius + 32
    End Function

    'Temperture
    ''' <summary>
    ''' FUNCTION: FAHRENHEITTOCELSIUS '
    ''' DESCRIPTION: CONVERTS FAHRENHEIT DEGREES TO CELSIUS DEGREES '
    ''' NOTES: THE LARGEST NUMBER FAHRENHEITTOCELSIUS WILL CONVERT IS 32,767 '
    ''' </summary>
    ''' <param name="intFahrenheit"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function CnvFahrenheitToCelsius(intFahrenheit As Integer) As Integer
        CnvFahrenheitToCelsius = (5 / 9) * (intFahrenheit - 32)
    End Function

    ' ***************************** '
    ' **     SPYDAZ AI MATRIX    ** '
    ' ***************************** '
    ':FLIUD VOL:
    <Runtime.CompilerServices.Extension()>
    Public Sub CnvGallonToALL(ByRef GALLON As Integer, ByRef LITRE As Integer, ByRef PINT As Integer)
        LITRE = Val(GALLON * 3.79)
        PINT = Val(GALLON * 8)
    End Sub

    ':WEIGHT:
    <Runtime.CompilerServices.Extension()>
    Public Sub CnvGramsTOALL(ByRef GRAM As Integer, ByRef KILO As Integer, ByRef OUNCE As Integer, ByRef POUNDS As Integer)
        KILO = Val(GRAM * 0.001)
        OUNCE = Val(GRAM * 0.03527337)
        POUNDS = Val(GRAM * 0.002204634)
    End Sub

    <Runtime.CompilerServices.Extension()>
    Public Sub CnvkilosTOALL(ByRef KILO As Integer, ByRef GRAM As Integer, ByRef OUNCE As Integer, ByRef POUNDS As Integer)
        GRAM = Val(KILO * 1000)
        OUNCE = Val(KILO * 35.27337)
        POUNDS = Val(KILO * 2.204634141)
    End Sub

    <Runtime.CompilerServices.Extension()>
    Public Sub CnvLitreToALL(ByRef LITRE As Integer, ByRef PINT As Integer, ByRef GALLON As Integer)
        PINT = Val(LITRE * 2.113427663)
        GALLON = Val(LITRE * 0.263852243)
    End Sub

    <Runtime.CompilerServices.Extension()>
    Public Sub CnvOunceToALL(ByRef OUNCE As Integer, ByRef GRAM As Integer, ByRef KILO As Integer, ByRef POUNDS As Integer)
        GRAM = Val(OUNCE * 28.349)
        KILO = Val(OUNCE * 0.028349)
        POUNDS = Val(OUNCE * 0.0625)
    End Sub

    <Runtime.CompilerServices.Extension()>
    Public Sub CnvPintToALL(ByRef PINT As Integer, ByRef GALLON As Integer, ByRef LITRE As Integer)
        LITRE = Val(PINT * 0.473165)
        GALLON = Val(PINT * 0.1248455)
    End Sub

    ''' <summary>
    ''' Checks if String Contains Letters
    ''' </summary>
    ''' <param name="str"></param>
    ''' <returns></returns>
    <Runtime.CompilerServices.Extension()>
    Public Function ContainsLetters(ByVal str As String) As Boolean

        For i = 0 To str.Length - 1
            If Char.IsLetter(str.Chars(i)) Then
                Return True
            End If
        Next

        Return False

    End Function

    ''' <summary>
    ''' When two sets of data are strongly linked together we say they have a High Correlation.
    ''' Correlation is Positive when the values increase together, and Correlation is Negative
    ''' when one value decreases as the other increases 1 is a perfect positive correlation 0 Is
    ''' no correlation (the values don't seem linked at all)
    ''' -1 Is a perfect negative correlation
    ''' </summary>
    ''' <param name="X_Series"></param>
    ''' <param name="Y_Series"></param>
    <Runtime.CompilerServices.Extension()>
    Public Function Correlation(ByRef X_Series As List(Of Double), ByRef Y_Series As List(Of Double)) As Double

        'Step 1 Find the mean Of x, And the mean of y
        'Step 2: Subtract the mean of x from every x value (call them "a"), do the same for y	(call them "b")
        'Results
        Dim DifferenceX As List(Of Double) = Difference(X_Series)
        Dim DifferenceY As List(Of Double) = Difference(Y_Series)

        'Step 3: Calculate : a*b, XSqu And YSqu for every value
        'Step 4: Sum up ab, sum up a2 And sum up b2
        'Results
        Dim SumXSqu As Double = Sum(Square(DifferenceX))
        Dim SumYSqu As Double = Sum(Square(DifferenceY))
        Dim SumX_Y As Double = Sum(Multiply(X_Series, Y_Series))

        'Step 5: Divide the sum of a*b by the square root of [(SumXSqu) × (SumYSqu)]
        'Results
        Dim Answer As Double = SumX_Y / Math.Sqrt(SumXSqu * SumYSqu)
        Return Answer
    End Function

    ''' <summary>
    ''' COMMENTS : RETURNS THE COSECANT OF THE SUPPLIED NUMBER. ' NOTE THAT SIN(DBLIN) CANNOT
    ''' EQUAL ZERO. THIS CAN ' HAPPEN IF DBLIN IS A MULTIPLE OF CDBLPI. '
    ''' PARAMETERS: DBLIN - VALUE TO CALCULATE ' RETURNS : COSECANT AS A DOUBLE ' '
    ''' </summary>
    ''' <param name="DBLIN"></param>
    ''' <returns>DBLIN - VALUE TO CALCULATE ' RETURNS : COSECANT AS A DOUBLE</returns>
    <Runtime.CompilerServices.Extension()>
    Public Function COSECANT(ByVal DBLIN As Double) As Double

        On Error GoTo PROC_ERR

        COSECANT = If(Math.Sin(DBLIN) = 0, 0, 1 / Math.Sin(DBLIN))

PROC_EXIT:
        Exit Function

PROC_ERR:
        MsgBox("ERROR: " & Err.Number & ". " & Err.Description, ,
    NameOf(COSECANT))
        Resume PROC_EXIT

    End Function

    ''' <summary>
    ''' COMMENTS : RETURNS THE COTANGENT OF THE SUPPLIED NUMBER ' FUNCTION IS UNDEFINED IF INPUT
    ''' VALUE IS A MULTIPLE OF CDBLPI. '
    ''' PARAMETERS: DBLIN - VALUE TO CALCULATE ' RETURNS : COTANGENT AS A DOUBLE
    ''' </summary>
    ''' <param name="DBLIN"></param>
    ''' <returns>COTANGENT AS A DOUBLE</returns>
    <Runtime.CompilerServices.Extension()>
    Public Function COTANGENT(ByVal DBLIN As Double) As Double

        On Error GoTo PROC_ERR

        COTANGENT = If(Math.Tan(DBLIN) = 0, 0, 1 / Math.Tan(DBLIN))

PROC_EXIT:
        Exit Function

PROC_ERR:
        MsgBox("ERROR: " & Err.Number & ". " & Err.Description, ,
    NameOf(COTANGENT))
        Resume PROC_EXIT

    End Function

    ''' <summary>
    ''' Counts the number of elements in the text, useful for declaring arrays when the element
    ''' length is unknown could be used to split sentence on full stop Find Sentences then again
    ''' on comma(conjunctions) "Find Clauses" NumberOfElements = CountElements(Userinput, delimiter)
    ''' </summary>
    ''' <param name="PHRASE"></param>
    ''' <param name="Delimiter"></param>
    ''' <returns>Integer : number of elements found</returns>
    ''' <remarks></remarks>
    <System.Runtime.CompilerServices.Extension()>
    Public Function CountElements(ByVal PHRASE As String, ByVal Delimiter As String) As Integer
        Dim elementcounter As Integer = 0
        Dim PhraseArray As String()
        PhraseArray = PHRASE.Split(Delimiter)
        elementcounter = UBound(PhraseArray)
        Return elementcounter
    End Function

    ''' <summary>
    ''' counts occurrences of a specific phoneme
    ''' </summary>
    ''' <param name="strIn"></param>
    ''' <param name="strFind"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function CountOccurrences(ByRef strIn As String, ByRef strFind As String) As Integer
        '**
        ' Returns: the number of times a string appears in a string
        '
        '@rem           Example code for CountOccurrences()
        '
        '  ' Counts the occurrences of "ow" in the supplied string.
        '
        '    strTmp = "How now, brown cow"
        '    Returns a value of 4
        '
        '
        'Debug.Print "CountOccurrences(): there are " &  CountOccurrences(strTmp, "ow") &
        '" occurrences of 'ow'" &    " in the string '" & strTmp & "'"
        '
        '@param        strIn Required. String.
        '@param        strFind Required. String.
        '@return       Long.

        Dim lngPos As Integer
        Dim lngWordCount As Integer

        On Error GoTo PROC_ERR

        lngWordCount = 1

        ' Find the first occurrence
        lngPos = InStr(strIn, strFind)

        Do While lngPos > 0
            ' Find remaining occurrences
            lngPos = InStr(lngPos + 1, strIn, strFind)
            If lngPos > 0 Then
                ' Increment the hit counter
                lngWordCount = lngWordCount + 1
            End If
        Loop

        ' Return the value
        CountOccurrences = lngWordCount

PROC_EXIT:
        Exit Function

PROC_ERR:
        MsgBox("Error: " & Err.Number & ". " & Err.Description, , NameOf(CountOccurrences))
        Resume PROC_EXIT

    End Function

    ''' <summary>
    ''' Counts tokens in string
    ''' </summary>
    ''' <param name="Str">string to be searched</param>
    ''' <param name="Delimiter">delimiter such as space comma etc</param>
    ''' <returns></returns>
    <System.Runtime.CompilerServices.Extension>
    Public Function CountTokensInString(ByRef Str As String, ByRef Delimiter As String) As Integer
        Dim Words() As String = Split(Str, Delimiter)
        Return Words.Count
    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function CountVowels(ByVal InputString As String) As Integer
        Dim v(9) As String 'Declare an array  of 10 elements 0 to 9
        Dim vcount As Short 'This variable will contain number of vowels
        Dim flag As Integer
        Dim strLen As Integer
        Dim i As Integer
        v(0) = "a" 'First element of array is assigned small a
        v(1) = "i"
        v(2) = "o"
        v(3) = "u"
        v(4) = "e"
        v(5) = "A" 'Sixth element is assigned Capital A
        v(6) = "I"
        v(7) = "O"
        v(8) = "U"
        v(9) = "E"
        strLen = Len(InputString)

        For flag = 1 To strLen 'It will get every letter of entered string and loop
            'will terminate when all letters have been examined

            For i = 0 To 9 'Takes every element of v(9) one by one
                'Check if current letter is a vowel
                If Mid(InputString, flag, 1) = v(i) Then
                    vcount = vcount + 1 ' If letter is equal to vowel
                    'then increment vcount by 1
                End If
            Next i 'Consider next value of v(i)
        Next flag 'Consider next letter of the enterd string

        CountVowels = vcount

    End Function

    ''' <summary>
    ''' Counts the words in a given text
    ''' </summary>
    ''' <param name="NewText"></param>
    ''' <returns>integer: number of words</returns>
    ''' <remarks></remarks>
    <System.Runtime.CompilerServices.Extension()>
    Public Function CountWords(NewText As String) As Integer
        Dim TempArray() As String = NewText.Split(" ")
        CountWords = UBound(TempArray)
        Return CountWords
    End Function

    ''' <summary>
    ''' checks Str contains keyword regardless of case
    ''' </summary>
    ''' <param name="Userinput"></param>
    ''' <param name="Keyword"></param>
    ''' <returns></returns>
    <Runtime.CompilerServices.Extension()>
    Public Function DetectKeyWord(ByRef Userinput As String, ByRef Keyword As String) As Boolean
        Dim mfound As Boolean = False
        If UCase(Userinput).Contains(UCase(Keyword)) = True Or
    InStr(Userinput, Keyword) > 1 Then
            mfound = True
        End If

        Return mfound
    End Function

    ''' <summary>
    ''' DETECT IF STATMENT IS AN IF/THEN DETECT IF STATMENT IS AN IF/THEN -- -RETURNS PARTS DETIFTHEN
    ''' = DETECTLOGIC(USERINPUT, "IF", "THEN", IFPART, THENPART)
    ''' </summary>
    ''' <param name="userinput"></param>
    ''' <param name="LOGICA">"IF", can also be replace by "IT CAN BE SAID THAT</param>
    ''' <param name="LOGICB">"THEN" can also be replaced by "it must follow that"</param>
    ''' <param name="IFPART">supply empty string to be used to hold part</param>
    ''' <param name="THENPART">supply empty string to be used to hold part</param>
    ''' <returns>true/false</returns>
    ''' <remarks></remarks>
    <System.Runtime.CompilerServices.Extension()>
    Public Function DetectLOGIC(ByRef userinput As String, ByRef LOGICA As String, ByRef LOGICB As String, ByRef IFPART As String, ByRef THENPART As String) As Boolean
        If InStr(1, userinput, LOGICA, 1) > 0 And InStr(1, userinput, " " & LOGICB & " ", 1) > 0 Then
            'SPLIT USER INPUT
            Call SplitPhrase(userinput, " " & LOGICB & " ", IFPART, THENPART)

            IFPART = Replace(IFPART, LOGICA, "", 1, -1, CompareMethod.Text)
            THENPART = Replace(THENPART, " " & LOGICB & " ", "", 1, -1, CompareMethod.Text)
            DetectLOGIC = True
        Else
            DetectLOGIC = False
        End If
    End Function

    ''' <summary>
    ''' Returns the Difference Form the Mean
    ''' </summary>
    ''' <param name="Series"></param>
    ''' <returns></returns>
    <Runtime.CompilerServices.Extension()>
    Public Function Difference(ByRef Series As List(Of Double)) As List(Of Double)
        Dim TheMean As Double = Mean(Series)
        Dim NewSeries As New List(Of Double)
        For Each item In Series
            NewSeries.Add(item - TheMean)
        Next
        Return NewSeries
    End Function

    ''' <summary>
    ''' Expand a string such as a field name by inserting a space ahead of each capitalized
    ''' letter (where none exists).
    ''' </summary>
    ''' <param name="inputString"></param>
    ''' <returns>Expanded string</returns>
    ''' <remarks></remarks>
    <System.Runtime.CompilerServices.Extension()>
    Public Function ExpandToWords(ByVal inputString As String) As String
        If inputString Is Nothing Then Return Nothing
        Dim charArray = inputString.ToCharArray
        Dim outStringBuilder As New System.Text.StringBuilder(inputString.Length + 10)
        For index = 0 To charArray.GetUpperBound(0)
            If Char.IsUpper(charArray(index)) Then
                'If previous character is also uppercase, don't expand as this may be an acronym.
                If (index > 0) AndAlso Char.IsUpper(charArray(index - 1)) Then
                    outStringBuilder.Append(charArray(index))
                Else
                    outStringBuilder.Append(String.Concat(" ", charArray(index)))
                End If
            Else
                outStringBuilder.Append(charArray(index))
            End If
        Next

        Return outStringBuilder.ToString.Replace("_", " ").Trim

    End Function

    <System.Runtime.CompilerServices.Extension()>
    Public Function ExtractFirstChar(ByRef InputStr As String) As String

        ExtractFirstChar = Left(InputStr, 1)
    End Function

    <System.Runtime.CompilerServices.Extension()>
    Public Function ExtractFirstWord(ByRef Statement As String) As String
        Dim StrArr() As String = Split(Statement, " ")
        Return StrArr(0)
    End Function

    <System.Runtime.CompilerServices.Extension()>
    Public Function ExtractLastChar(ByRef InputStr As String) As String

        ExtractLastChar = Right(InputStr, 1)
    End Function

    ''' <summary>
    ''' Returns The last word in String
    ''' NOTE: String ois converted to Array then the last element is extracted Count-1
    ''' </summary>
    ''' <param name="InputStr"></param>
    ''' <returns>String</returns>
    <System.Runtime.CompilerServices.Extension()>
    Public Function ExtractLastWord(ByRef InputStr As String) As String
        Dim TempArr() As String = Split(InputStr, " ")
        Dim Count As Integer = TempArr.Count - 1
        Return TempArr(Count)
    End Function

    ''' <summary>
    ''' extracts string between defined strings
    ''' </summary>
    ''' <param name="value">base sgtring</param>
    ''' <param name="strStart">Start string</param>
    ''' <param name="strEnd">End string</param>
    ''' <returns></returns>
    <System.Runtime.CompilerServices.Extension()>
    Public Function ExtractStringBetween(ByVal value As String, ByVal strStart As String, ByVal strEnd As String) As String
        If Not String.IsNullOrEmpty(value) Then
            Dim i As Integer = value.IndexOf(strStart)
            Dim j As Integer = value.IndexOf(strEnd)
            Return value.Substring(i, j - i)
        Else
            Return value
        End If
    End Function

    ''' <summary>
    ''' This will extract the words within the (Text) between (word1) and (word2). Example:
    ''' Text="The rain in Spain." Print ExtractWordsBetween("the "," in ",Text) rain
    ''' </summary>
    ''' <param name="Word1"></param>
    ''' <param name="Word2"></param>
    ''' <param name="Text"> </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExtractWordsBetween(ByRef Word1 As String, ByRef Word2 As String, ByRef Text As String) As String

        '
        Dim Position As Short
        Dim Position2 As Short

        If Word1 = "" Then Position = 1 : GoTo Lp1
        Position = InStr(1, Text, Word1)
        If Position = 0 Then Position = 1
        Position = Position + Len(Word1)
Lp1:

        If Word2 = "" Then Position2 = Len(Text) + 1 : GoTo Lp2
        Position2 = InStr(Position, Text, Word2)
        If Position2 = 0 Then Position2 = Len(Text) + 1
Lp2:
        ExtractWordsBetween = Mid(Text, Position, Position2 - Position)
        ExtractWordsBetween = Trim(ExtractWordsBetween)
    End Function

    ''' <summary>
    ''' Extract words Either side of Divider
    ''' </summary>
    ''' <param name="TextStr"></param>
    ''' <param name="Divider"></param>
    ''' <param name="Mode">Front = F Back =B</param>
    ''' <returns></returns>
    <System.Runtime.CompilerServices.Extension>
    Public Function ExtractWordsEitherSide(ByRef TextStr As String, ByRef Divider As String, ByRef Mode As String) As String
        ExtractWordsEitherSide = ""
        Select Case Mode
            Case "F"
                Return ExtractWordsEitherSide(TextStr, Divider, "F")
            Case "B"
                Return ExtractWordsEitherSide(TextStr, Divider, "B")
        End Select

    End Function

    ' Generate a random number based on the upper and lower bounds of the array,
    'then use that to return the item.
    <System.Runtime.CompilerServices.Extension()>
    Public Function FetchRandomItem(Of t)(ByRef theArray() As t) As t

        Dim randNumberGenerator As New Random
        Randomize()
        Dim index As Integer = randNumberGenerator.Next(theArray.GetLowerBound(0),
                                    theArray.GetUpperBound(0) + 1)

        Return theArray(index)

    End Function

    ''' <summary>
    ''' Writes the contents of an embedded file resource embedded as Bytes to disk.
    ''' EG: My.Resources.DefBrainConcepts.FileSave(Application.StartupPath and "\DefBrainConcepts.ACCDB")
    ''' </summary>
    ''' <param name="BytesToWrite">Embedded resource Name</param>
    ''' <param name="FileName">Save to file</param>
    ''' <remarks></remarks>
    <System.Runtime.CompilerServices.Extension()>
    Public Sub FileSave(ByVal BytesToWrite() As Byte, ByVal FileName As String)

        If IO.File.Exists(FileName) Then
            IO.File.Delete(FileName)
        End If

        Dim FileStream As New System.IO.FileStream(FileName, System.IO.FileMode.OpenOrCreate)
        Dim BinaryWriter As New System.IO.BinaryWriter(FileStream)

        BinaryWriter.Write(BytesToWrite)
        BinaryWriter.Close()
        FileStream.Close()
    End Sub

    ''' <summary>
    ''' Define the search terms. This list could also be dynamically populated at runtime Find
    ''' sentences that contain all the terms in the wordsToMatch array Note that the number of
    ''' terms to match is not specified at compile time
    ''' </summary>
    ''' <param name="TextStr1">String to be searched</param>
    ''' <param name="Words">List of Words to be detected</param>
    ''' <returns>Sentences containing words</returns>
    <Runtime.CompilerServices.Extension()>
    Public Function FindSentencesContaining(ByRef TextStr1 As String, ByRef Words As List(Of String)) As List(Of String)
        ' Split the text block into an array of sentences.
        Dim sentences As String() = TextStr1.Split(New Char() {".", "?", "!"})

        Dim wordsToMatch(Words.Count) As String
        Dim I As Integer = 0
        For Each item In Words
            wordsToMatch(I) = item
            I += 1
        Next

        Dim sentenceQuery = From sentence In sentences
                            Let w = sentence.Split(New Char() {" ", ",", ".", ";", ":"},
                                   StringSplitOptions.RemoveEmptyEntries)
                            Where w.Distinct().Intersect(wordsToMatch).Count = wordsToMatch.Count()
                            Select sentence

        ' Execute the query

        Dim StrList As New List(Of String)
        For Each str As String In sentenceQuery
            StrList.Add(str)
        Next
        Return StrList
    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function FormatJsonOutput(ByVal jsonString As String) As String
        Dim stringBuilder = New StringBuilder()
        Dim escaping As Boolean = False
        Dim inQuotes As Boolean = False
        Dim indentation As Integer = 0

        For Each character As Char In jsonString

            If escaping Then
                escaping = False
                stringBuilder.Append(character)
            Else

                If character = "\"c Then
                    escaping = True
                    stringBuilder.Append(character)
                ElseIf character = """"c Then
                    inQuotes = Not inQuotes
                    stringBuilder.Append(character)
                ElseIf Not inQuotes Then

                    If character = ","c Then
                        stringBuilder.Append(character)
                        stringBuilder.Append(vbCrLf)
                        stringBuilder.Append(vbTab, indentation)
                    ElseIf character = "["c OrElse character = "{"c Then
                        stringBuilder.Append(character)
                        stringBuilder.Append(vbCrLf)
                        stringBuilder.Append(vbTab, System.Threading.Interlocked.Increment(indentation))
                    ElseIf character = "]"c OrElse character = "}"c Then
                        stringBuilder.Append(vbCrLf)
                        stringBuilder.Append(vbTab, System.Threading.Interlocked.Decrement(indentation))
                        stringBuilder.Append(character)
                    ElseIf character = ":"c Then
                        stringBuilder.Append(character)
                        stringBuilder.Append(vbTab)
                    ElseIf Not Char.IsWhiteSpace(character) Then
                        stringBuilder.Append(character)
                    End If
                Else
                    stringBuilder.Append(character)
                End If
            End If
        Next

        Return stringBuilder.ToString()
    End Function

    <System.Runtime.CompilerServices.Extension()>
    Public Function FormatText(ByRef Text As String) As String
        Dim FormatTextResponse As String = ""
        'FORMAT USERINPUT
        'turn to uppercase for searching the db
        Text = LTrim(Text)
        Text = RTrim(Text)
        Text = UCase(Text)

        FormatTextResponse = Text
        Return FormatTextResponse
    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function Gaussian(ByRef x As Double) As Double
        Gaussian = Math.Exp((-x * -x) / 2)
    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function GaussianDerivative(ByRef x As Double) As Double
        GaussianDerivative = Gaussian(x) * (-x / (-x * -x))
    End Function

    ''' <summary>
    ''' Gets the string after the given string parameter.
    ''' </summary>
    ''' <param name="value">The default value.</param>
    ''' <param name="x">The given string parameter.</param>
    ''' <returns></returns>
    ''' <remarks>Unlike GetBefore, this method trims the result</remarks>
    <System.Runtime.CompilerServices.Extension>
    Public Function GetAfter(value As String, x As String) As String
        Dim xPos = value.LastIndexOf(x, StringComparison.Ordinal)
        If xPos = -1 Then
            Return [String].Empty
        End If
        Dim startIndex = xPos + x.Length
        Return If(startIndex >= value.Length, [String].Empty, value.Substring(startIndex).Trim())
    End Function

    ''' <summary>
    ''' Gets the string before the given string parameter.
    ''' </summary>
    ''' <param name="value">The default value.</param>
    ''' <param name="x">The given string parameter.</param>
    ''' <returns></returns>
    ''' <remarks>Unlike GetBetween and GetAfter, this does not Trim the result.</remarks>
    <System.Runtime.CompilerServices.Extension>
    Public Function GetBefore(value As String, x As String) As String
        Dim xPos = value.IndexOf(x, StringComparison.Ordinal)
        Return If(xPos = -1, [String].Empty, value.Substring(0, xPos))
    End Function

    ''' <summary>
    ''' Gets the string between the given string parameters.
    ''' </summary>
    ''' <param name="value">The source value.</param>
    ''' <param name="x">The left string sentinel.</param>
    ''' <param name="y">The right string sentinel</param>
    ''' <returns></returns>
    ''' <remarks>Unlike GetBefore, this method trims the result</remarks>
    <System.Runtime.CompilerServices.Extension>
    Public Function GetBetween(value As String, x As String, y As String) As String
        Dim xPos = value.IndexOf(x, StringComparison.Ordinal)
        Dim yPos = value.LastIndexOf(y, StringComparison.Ordinal)
        If xPos = -1 OrElse xPos = -1 Then
            Return [String].Empty
        End If
        Dim startIndex = xPos + x.Length
        Return If(startIndex >= yPos, [String].Empty, value.Substring(startIndex, yPos - startIndex).Trim())
    End Function

    ''' <summary>
    ''' Returns the first Word
    ''' </summary>
    ''' <param name="Statement"></param>
    ''' <returns></returns>
    <System.Runtime.CompilerServices.Extension()>
    Public Function GetPrefix(ByRef Statement As String) As String
        Dim StrArr() As String = Split(Statement, " ")
        Return StrArr(0)
    End Function

    <System.Runtime.CompilerServices.Extension()>
    Public Function GetRandomObjectFromList(ByRef Patterns As List(Of Object)) As Object
        Dim rnd = New Random()
        If Patterns.Count > 0 Then

            Return Patterns(rnd.Next(0, Patterns.Count))
        Else
            Return Nothing
        End If
    End Function

    ''' <summary>
    ''' Returns random character from string given length of the string to choose from
    ''' </summary>
    ''' <param name="Source"></param>
    ''' <param name="Length"></param>
    ''' <returns></returns>
    <Runtime.CompilerServices.Extension()>
    Public Function GetRndChar(ByVal Source As String, ByVal Length As Integer) As String
        Dim rnd As New Random
        If Source Is Nothing Then Throw New ArgumentNullException(NameOf(Source), "Must contain a string,")
        If Length <= 0 Then Throw New ArgumentException("Length must be a least one.", NameOf(Length))
        Dim s As String = ""
        Dim builder As New System.Text.StringBuilder()
        builder.Append(s)
        For i = 1 To Length
            builder.Append(Source(rnd.Next(0, Source.Length)))
        Next
        s = builder.ToString()
        Return s
    End Function

    ''' <summary>
    ''' Returns from index to end of file
    ''' </summary>
    ''' <param name="Str">String</param>
    ''' <param name="indx">Index</param>
    ''' <returns></returns>
    <Runtime.CompilerServices.Extension()>
    Public Function GetSlice(ByRef Str As String, ByRef indx As Integer) As String
        If indx <= Str.Length Then
            Str.Substring(indx, Str.Length)
            Return Str(indx)
        Else
        End If
        Return Nothing
    End Function

    ''' <summary>
    ''' gets the last word
    ''' </summary>
    ''' <param name="InputStr"></param>
    ''' <returns></returns>
    <System.Runtime.CompilerServices.Extension()>
    Public Function GetSuffix(ByRef InputStr As String) As String
        Dim TempArr() As String = Split(InputStr, " ")
        Dim Count As Integer = TempArr.Count - 1
        Return TempArr(Count)
    End Function

    <System.Runtime.CompilerServices.Extension>
    Public Function GetWordsBetween(ByRef InputStr As String, ByRef StartStr As String, ByRef StopStr As String)
        Return InputStr.ExtractStringBetween(StartStr, StopStr)
    End Function

    ''' <summary>
    ''' Basic Growth
    ''' </summary>
    ''' <param name="Past"></param>
    ''' <param name="Present"></param>
    ''' <returns></returns>
    <Runtime.CompilerServices.Extension()>
    Public Function Growth(ByRef Past As Double, ByRef Present As Double) As Double
        Growth = (Present - Past) / Past
    End Function

    ''' <summary>
    ''' COMMENTS : RETURNS THE INVERSE HYPERBOLIC COSECANT OF THE ' SUPPLIED NUMBER '
    ''' PARAMETERS: DBLIN - VALUE TO CALCULATE ' RETURNS : HYPERBOLIC INVERSE COSECANT AS A DOUBLE
    ''' </summary>
    ''' <param name="DBLIN">- VALUE TO CALCULATE</param>
    ''' <returns>HYPERBOLIC INVERSE COSECANT AS A DOUBLE</returns>
    <Runtime.CompilerServices.Extension()>
    Public Function HYPERBOLICARCCOSECANT(ByVal DBLIN As Double) As Double

        On Error GoTo PROC_ERR

        HYPERBOLICARCCOSECANT = Math.Log((Math.Sign(DBLIN) *
    Math.Sqrt(DBLIN * DBLIN + 1) + 1) / DBLIN)

PROC_EXIT:
        Exit Function

PROC_ERR:
        MsgBox("ERROR: " & Err.Number & ". " & Err.Description, ,
    NameOf(HYPERBOLICARCCOSECANT))
        Resume PROC_EXIT

    End Function

    ''' <summary>
    ''' COMMENTS : RETURNS THE INVERSE HYPERBOLIC COSINE OF THE ' SUPPLIED NUMBER '
    ''' PARAMETERS: DBLIN - VALUE TO CALCULATE
    ''' </summary>
    ''' <param name="DBLIN"></param>
    ''' <returns>RETURNS : INVERSE HYPERBOLIC COSINE AS A DOUBLE</returns>
    <Runtime.CompilerServices.Extension()>
    Public Function HYPERBOLICARCCOSINE(ByVal DBLIN As Double) As Double

        On Error GoTo PROC_ERR

        HYPERBOLICARCCOSINE = Math.Log(DBLIN + Math.Sqrt(DBLIN * DBLIN - 1))

PROC_EXIT:
        Exit Function

PROC_ERR:
        MsgBox("ERROR: " & Err.Number & ". " & Err.Description, ,
    NameOf(HYPERBOLICARCCOSINE))
        Resume PROC_EXIT

    End Function

    ''' <summary>
    ''' COMMENTS : RETURNS THE INVERSE HYPERBOLIC TANGENT OF THE ' SUPPLIED NUMBER. UNDEFINED IF
    ''' DBLIN IS 1 '
    ''' PARAMETERS: DBLIN - VALUE TO CALCULATE
    ''' </summary>
    ''' <param name="DBLIN"></param>
    ''' <returns>INVERSE HYPERBOLIC TANGENT AS A DOUBLE</returns>
    <Runtime.CompilerServices.Extension()>
    Public Function HYPERBOLICARCCOTANGENT(ByVal DBLIN As Double) As Double

        On Error GoTo PROC_ERR

        HYPERBOLICARCCOTANGENT = Math.Log((DBLIN + 1) / (DBLIN - 1)) / 2

PROC_EXIT:
        Exit Function

PROC_ERR:
        MsgBox("ERROR: " & Err.Number & ". " & Err.Description, ,
    NameOf(HYPERBOLICARCCOTANGENT))
        Resume PROC_EXIT

    End Function

    ''' <summary>
    ''' COMMENTS : RETURNS THE INVERSE HYPERBOLIC SECANT OF THE ' SUPPLIED NUMBER '
    ''' PARAMETERS: DBLIN - VALUE TO CALCULATE
    ''' </summary>
    ''' <param name="DBLIN"></param>
    ''' <returns>RETURNS : INVERSE HYPERBOLIC SECANT AS A DOUBLE</returns>
    <Runtime.CompilerServices.Extension()>
    Public Function HYPERBOLICARCSECANT(ByVal DBLIN As Double) As Double

        On Error GoTo PROC_ERR

        HYPERBOLICARCSECANT = Math.Log((Math.Sqrt(-DBLIN ^ 2 + 1) + 1) / DBLIN)

PROC_EXIT:
        Exit Function

PROC_ERR:
        MsgBox("ERROR: " & Err.Number & ". " & Err.Description, ,
    NameOf(HYPERBOLICARCSECANT))
        Resume PROC_EXIT

    End Function

    ''' <summary>
    ''' COMMENTS : RETURNS THE INVERSE HYPERBOLIC SINE OF THE SUPPLIED NUMBER '
    ''' PARAMETERS: DBLIN - VALUE TO CALCULATE ' RETURNS : INVERSE HYPERBOLIC SINE AS A DOUBLE
    ''' </summary>
    ''' <param name="DBLIN">VALUE TO CALCULATE</param>
    ''' <returns>INVERSE HYPERBOLIC SINE AS A DOUBLE</returns>
    <Runtime.CompilerServices.Extension()>
    Public Function HYPERBOLICARCSINE(ByVal DBLIN As Double) As Double

        On Error GoTo PROC_ERR

        HYPERBOLICARCSINE = Math.Log(DBLIN + Math.Sqrt(DBLIN ^ 2 + 1))

PROC_EXIT:
        Exit Function

PROC_ERR:
        MsgBox("ERROR: " & Err.Number & ". " & Err.Description, ,
    NameOf(HYPERBOLICARCSINE))
        Resume PROC_EXIT

    End Function

    ''' <summary>
    ''' COMMENTS : RETURNS THE INVERSE HYPERBOLIC TANGENT OF THE ' SUPPLIED NUMBER. THE RETURN
    ''' VALUE IS UNDEFINED IF ' INPUT VALUE IS 1. '
    ''' PARAMETERS: DBLIN - VALUE TO CALCULATE ' RETURNS : INVERSE HYPERBOLIC TANGENT AS A
    '''             DOUBLE '
    ''' </summary>
    ''' <param name="DBLIN">VALUE TO CALCULATE</param>
    ''' <returns>
    ''' DBLIN - VALUE TO CALCULATE ' RETURNS : INVERSE HYPERBOLIC TANGENT AS A DOUBLE '
    ''' </returns>
    <Runtime.CompilerServices.Extension()>
    Public Function HYPERBOLICARCTAN(ByVal DBLIN As Double) As Double

        HYPERBOLICARCTAN = Math.Log((1 + DBLIN) / (1 - DBLIN)) / 2

PROC_EXIT:
        Exit Function

PROC_ERR:
        MsgBox("ERROR: " & Err.Number & ". " & Err.Description, ,
    NameOf(HYPERBOLICARCTAN))
        Resume PROC_EXIT

    End Function

    ''' <summary>
    ''' COMMENTS : RETURNS THE HYPERBOLIC COSECANT OF THE SUPPLIED NUMBER '
    ''' PARAMETERS: DBLIN - VALUE TO CALCULATE ' RETURNS : HYPERBOLIC COSECANT AS A DOUBLE ' '
    ''' </summary>
    ''' <param name="DBLIN"></param>
    ''' <returns>RETURNS : HYPERBOLIC COSECANT AS A DOUBLE</returns>
    <Runtime.CompilerServices.Extension()>
    Public Function HYPERBOLICCOSECANT(ByVal DBLIN As Double) As Double

        On Error GoTo PROC_ERR

        HYPERBOLICCOSECANT = 2 / (Math.Exp(DBLIN) - Math.Exp(-DBLIN))

PROC_EXIT:
        Exit Function

PROC_ERR:
        MsgBox("ERROR: " & Err.Number & ". " & Err.Description, ,
    NameOf(HYPERBOLICCOSECANT))
        Resume PROC_EXIT

    End Function

    ''' <summary>
    ''' COMMENTS : RETURNS THE HYPERBOLIC COSINE OF THE SUPPLIED NUMBER '
    ''' PARAMETERS: DBLIN - VALUE TO CALCULATE ' RETURNS : HYPERBOLIC COSINE AS A DOUBLE ' '
    ''' </summary>
    ''' <param name="DBLIN"></param>
    ''' <returns>RETURNS : HYPERBOLIC COSINE AS A DOUBLE</returns>
    <Runtime.CompilerServices.Extension()>
    Public Function HYPERBOLICCOSINE(ByVal DBLIN As Double) As Double

        On Error GoTo PROC_ERR

        HYPERBOLICCOSINE = (Math.Exp(DBLIN) + Math.Exp(-DBLIN)) / 2

PROC_EXIT:
        Exit Function

PROC_ERR:
        MsgBox("ERROR: " & Err.Number & ". " & Err.Description, ,
    NameOf(HYPERBOLICCOSINE))
        Resume PROC_EXIT

    End Function

    ''' <summary>
    ''' COMMENTS : RETURNS THE HYPERBOLIC COTANGENT OF THE SUPPLIED NUMBER '
    ''' PARAMETERS: DBLIN - VALUE TO CALCULATE ' RETURNS : HYPERBOLIC COTANGENT AS A DOUBLE
    ''' </summary>
    ''' <param name="DBLIN"></param>
    ''' <returns>DBLIN - VALUE TO CALCULATE ' RETURNS : HYPERBOLIC COTANGENT AS A DOUBLE</returns>
    <Runtime.CompilerServices.Extension()>
    Public Function HYPERBOLICCOTANGENT(ByVal DBLIN As Double) As Double

        On Error GoTo PROC_ERR

        HYPERBOLICCOTANGENT = (Math.Exp(DBLIN) + Math.Exp(-DBLIN)) /
    (Math.Exp(DBLIN) - Math.Exp(-DBLIN))

PROC_EXIT:
        Exit Function

PROC_ERR:
        MsgBox("ERROR: " & Err.Number & ". " & Err.Description, ,
    NameOf(HYPERBOLICCOTANGENT))
        Resume PROC_EXIT

    End Function

    ''' <summary>
    ''' COMMENTS : RETURNS THE HYPERBOLIC SECANT OF THE SUPPLIED NUMBER '
    ''' PARAMETERS: DBLIN - VALUE TO CALCULATE ' RETURNS : HYPERBOLIC SECANT AS A DOUBLE
    ''' </summary>
    ''' <param name="DBLIN"></param>
    ''' <returns>DBLIN - VALUE TO CALCULATE ' RETURNS : HYPERBOLIC SECANT AS A DOUBLE</returns>
    <Runtime.CompilerServices.Extension()>
    Public Function HYPERBOLICSECANT(ByVal DBLIN As Double) As Double

        ' COMMENTS : RETURNS THE HYPERBOLIC SECANT OF THE SUPPLIED NUMBER '
        ' PARAMETERS: DBLIN - VALUE TO CALCULATE ' RETURNS : HYPERBOLIC SECANT AS A DOUBLE ' '
        On Error GoTo PROC_ERR

        HYPERBOLICSECANT = 2 / (Math.Exp(DBLIN) + Math.Exp(-DBLIN))

PROC_EXIT:
        Exit Function

PROC_ERR:
        MsgBox("ERROR: " & Err.Number & ". " & Err.Description, ,
    NameOf(HYPERBOLICSECANT))
        Resume PROC_EXIT

    End Function

    ''' <summary>
    ''' COMMENTS : RETURNS THE HYPERBOLIC SINE OF THE SUPPLIED NUMBER
    ''' </summary>
    ''' <param name="DBLIN"></param>
    ''' <returns>DBLIN - VALUE TO CALCULATE ' RETURNS : HYPERBOLIC SINE AS A DOUBLE</returns>
    <Runtime.CompilerServices.Extension()>
    Public Function HYPERBOLICSINE(ByVal DBLIN As Double) As Double

        On Error GoTo PROC_ERR

        HYPERBOLICSINE = (Math.Exp(DBLIN) - Math.Exp(-DBLIN)) / 2

PROC_EXIT:
        Exit Function

PROC_ERR:
        MsgBox("ERROR: " & Err.Number & ". " & Err.Description, ,
    NameOf(HYPERBOLICSINE))
        Resume PROC_EXIT

    End Function

    ''' <summary>
    ''' Returns interpretation of Corelation results
    ''' </summary>
    ''' <param name="Correlation"></param>
    ''' <returns></returns>
    <Runtime.CompilerServices.Extension()>
    Public Function InterpretCorrelationResult(ByRef Correlation As Double) As CorrelationResult
        InterpretCorrelationResult = CorrelationResult.None
        If Correlation >= 1 Then
            InterpretCorrelationResult = CorrelationResult.Positive

        End If
        If Correlation > 0.5 And Correlation <= 0.9 Then
            InterpretCorrelationResult = CorrelationResult.PositiveHigh
        End If
        If Correlation > 0 And Correlation <= 0.5 Then
            InterpretCorrelationResult = CorrelationResult.PositiveLow
        End If
        If Correlation = 0 Then InterpretCorrelationResult = CorrelationResult.None
        If Correlation > -0.5 And Correlation <= 0 Then
            InterpretCorrelationResult = CorrelationResult.NegativeLow
        End If
        If Correlation > -0.9 And Correlation <= -0.5 Then
            InterpretCorrelationResult = CorrelationResult.NegativeHigh
        End If
        If Correlation >= -1 Then
            InterpretCorrelationResult = CorrelationResult.Negative
        End If
    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function IsCubeRoot(ByVal number As Integer) As Boolean
        Dim numberCubeRooted As Double = number ^ (1 / 3)
        Return If(CInt(numberCubeRooted) ^ 3 = number, True, False)
    End Function

    ''' <summary>
    '''     A string extension method that query if '@this' satisfy the specified pattern.
    ''' </summary>
    ''' <param name="this">The @this to act on.</param>
    ''' <param name="pattern">The pattern to use. Use '*' as wild card string.</param>
    ''' <returns>true if '@this' satisfy the specified pattern, false if not.</returns>
    <System.Runtime.CompilerServices.Extension()>
    Public Function IsLike(this As String, pattern As String) As Boolean
        ' Turn the pattern into regex pattern, and match the whole string with ^$
        Dim regexPattern As String = "^" + Regex.Escape(pattern) + "$"

        ' Escape special character ?, #, *, [], and [!]
        regexPattern = regexPattern.Replace("\[!", "[^").Replace("\[", "[").Replace("\]", "]").Replace("\?", ".").Replace("\*", ".*").Replace("\#", "\d")

        Return Regex.IsMatch(this, regexPattern)
    End Function

    ''' <summary>
    ''' Checks if string is a reserved VBscipt Keyword
    ''' </summary>
    ''' <param name="keyword"></param>
    ''' <returns></returns>
    <Runtime.CompilerServices.Extension()>
    Function IsReservedWord(ByVal keyword As String) As Boolean
        Dim IsReserved = False
        Select Case LCase(keyword)
            Case "and" : IsReserved = True
            Case "as" : IsReserved = True
            Case "boolean" : IsReserved = True
            Case "byref" : IsReserved = True
            Case "byte" : IsReserved = True
            Case "byval" : IsReserved = True
            Case "call" : IsReserved = True
            Case "case" : IsReserved = True
            Case "class" : IsReserved = True
            Case "const" : IsReserved = True
            Case "currency" : IsReserved = True
            Case "debug" : IsReserved = True
            Case "dim" : IsReserved = True
            Case "do" : IsReserved = True
            Case "double" : IsReserved = True
            Case "each" : IsReserved = True
            Case "else" : IsReserved = True
            Case "elseif" : IsReserved = True
            Case "empty" : IsReserved = True
            Case "end" : IsReserved = True
            Case "endif" : IsReserved = True
            Case "enum" : IsReserved = True
            Case "eqv" : IsReserved = True
            Case "event" : IsReserved = True
            Case "exit" : IsReserved = True
            Case "false" : IsReserved = True
            Case "for" : IsReserved = True
            Case "function" : IsReserved = True
            Case "get" : IsReserved = True
            Case "goto" : IsReserved = True
            Case "if" : IsReserved = True
            Case "imp" : IsReserved = True
            Case "implements" : IsReserved = True
            Case "in" : IsReserved = True
            Case "integer" : IsReserved = True
            Case "is" : IsReserved = True
            Case "let" : IsReserved = True
            Case "like" : IsReserved = True
            Case "long" : IsReserved = True
            Case "loop" : IsReserved = True
            Case "lset" : IsReserved = True
            Case "me" : IsReserved = True
            Case "mod" : IsReserved = True
            Case "new" : IsReserved = True
            Case "next" : IsReserved = True
            Case "not" : IsReserved = True
            Case "nothing" : IsReserved = True
            Case "null" : IsReserved = True
            Case "on" : IsReserved = True
            Case "option" : IsReserved = True
            Case "optional" : IsReserved = True
            Case "or" : IsReserved = True
            Case "paramarray" : IsReserved = True
            Case "preserve" : IsReserved = True
            Case "private" : IsReserved = True
            Case "public" : IsReserved = True
            Case "raiseevent" : IsReserved = True
            Case "redim" : IsReserved = True
            Case "rem" : IsReserved = True
            Case "resume" : IsReserved = True
            Case "rset" : IsReserved = True
            Case "select" : IsReserved = True
            Case "set" : IsReserved = True
            Case "shared" : IsReserved = True
            Case "single" : IsReserved = True
            Case "static" : IsReserved = True
            Case "stop" : IsReserved = True
            Case "sub" : IsReserved = True
            Case "then" : IsReserved = True
            Case "to" : IsReserved = True
            Case "true" : IsReserved = True
            Case "type" : IsReserved = True
            Case "typeof" : IsReserved = True
            Case "until" : IsReserved = True
            Case "variant" : IsReserved = True
            Case "wend" : IsReserved = True
            Case "while" : IsReserved = True
            Case "with" : IsReserved = True
            Case "xor" : IsReserved = True
        End Select
        Return IsReserved
    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function IsRoot(ByVal number As Integer, power As Integer) As Boolean
        Dim numberNRooted As Double = number ^ (1 / power)
        Return If(CInt(numberNRooted) ^ power = number, True, False)
    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function IsSquareRoot(ByVal number As Integer) As Boolean
        Dim numberSquareRooted As Double = Math.Sqrt(number)
        Return If(CInt(numberSquareRooted) ^ 2 = number, True, False)
    End Function

    ''' <summary>
    ''' Loads Json file to new string
    ''' </summary>
    ''' <returns></returns>
    Public Function LoadJson() As String

        Dim Scriptfile As String = ""
        Dim Ofile As New OpenFileDialog
        With Ofile
            .Filter = "Json files (*.Json)|*.Json"

            If (.ShowDialog() = DialogResult.OK) Then
                Scriptfile = .FileName
            End If
        End With
        Dim txt As String = ""
        If Scriptfile IsNot "" Then

            txt = My.Computer.FileSystem.ReadAllText(Scriptfile)
            Return txt
        Else
            Return txt
        End If

    End Function

    ''' <summary>
    ''' COMMENTS : RETURNS LOG BASE 10. THE POWER 10 MUST BE RAISED ' TO EQUAL A GIVEN NUMBER.
    ''' LOG BASE 10 IS DEFINED AS THIS: ' X = LOG(Y) WHERE Y = 10 ^ X '
    ''' PARAMETERS: DBLDECIMAL - VALUE TO CALCULATE (Y) ' RETURNS : LOG BASE 10 OF THE GIVEN VALUE
    ''' ' '
    ''' </summary>
    ''' <param name="DBLDECIMAL"></param>
    ''' <returns>
    ''' DBLDECIMAL - VALUE TO CALCULATE (Y) ' RETURNS : LOG BASE 10 OF THE GIVEN VALUE
    ''' </returns>
    <Runtime.CompilerServices.Extension()>
    Public Function LOG10(ByVal DBLDECIMAL As Double) As Double

        On Error GoTo PROC_ERR

        LOG10 = Math.Log(DBLDECIMAL) / Math.Log(10)

PROC_EXIT:
        Exit Function

PROC_ERR:
        MsgBox("ERROR: " & Err.Number & ". " & Err.Description, ,
    NameOf(LOG10))
        Resume PROC_EXIT

    End Function

    ''' <summary>
    ''' COMMENTS : RETURNS LOG BASE 2. THE POWER 2 MUST BE RAISED TO EQUAL ' A GIVEN NUMBER. '
    ''' LOG BASE 2 IS DEFINED AS THIS: ' X = LOG(Y) WHERE Y = 2 ^ X '
    ''' PARAMETERS: DBLDECIMAL - VALUE TO CALCULATE (Y) ' RETURNS : LOG BASE 2 OF A GIVEN NUMBER
    '''             ' '
    ''' </summary>
    ''' <param name="DBLDECIMAL"></param>
    ''' <returns>DBLDECIMAL - VALUE TO CALCULATE (Y) ' RETURNS : LOG BASE 2 OF A GIVEN NUMBER</returns>
    <Runtime.CompilerServices.Extension()>
    Public Function LOG2(ByVal DBLDECIMAL As Double) As Double

        On Error GoTo PROC_ERR

        LOG2 = Math.Log(DBLDECIMAL) / Math.Log(2)

PROC_EXIT:
        Exit Function

PROC_ERR:
        MsgBox("ERROR: " & Err.Number & ". " & Err.Description, ,
    NameOf(LOG2))
        Resume PROC_EXIT

    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function Logistic(ByRef Value As Double) As Double
        'z = bias + (sum of all inputs ) * (input*weight)
        'output = Sigmoid(z)
        'derivative input = z/weight
        'derivative Weight = z/input
        'Derivative output = output*(1-Output)
        'learning rule = Sum of total training error* derivative input * derivative output * rootmeansquare of errors

        Return 1 / 1 + Math.Exp(-Value)
    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function LogisticDerivative(ByRef Value As Double) As Double
        'z = bias + (sum of all inputs ) * (input*weight)
        'output = Sigmoid(z)
        'derivative input = z/weight
        'derivative Weight = z/input
        'Derivative output = output*(1-Output)
        'learning rule = Sum of total training error* derivative input * derivative output * rootmeansquare of errors

        Return Logistic(Value) * (1 - Logistic(Value))
    End Function

    ''' <summary>
    ''' COMMENTS : RETURNS LOG BASE N. THE POWER N MUST BE RAISED TO EQUAL ' A GIVEN NUMBER. '
    ''' LOG BASE N IS DEFINED AS THIS: ' X = LOG(Y) WHERE Y = N ^ X ' PARAMETERS:
    ''' </summary>
    ''' <param name="DBLDECIMAL"></param>
    ''' <param name="DBLBASEN"></param>
    ''' <returns>DBLDECIMAL - VALUE TO CALCULATE (Y) ' DBLBASEN - BASE ' RETURNS : LOG BASE</returns>
    <Runtime.CompilerServices.Extension()>
    Public Function LOGN(ByVal DBLDECIMAL As Double, ByVal DBLBASEN As Double) As Double

        ' N OF A GIVEN NUMBER '

        On Error GoTo PROC_ERR

        LOGN = Math.Log(DBLDECIMAL) / Math.Log(DBLBASEN)

PROC_EXIT:
        Exit Function

PROC_ERR:
        MsgBox("ERROR: " & Err.Number & ". " & Err.Description, ,
    NameOf(LOGN))
        Resume PROC_EXIT

    End Function

    ''' <summary>
    ''' The avarage of a Series
    ''' </summary>
    ''' <param name="Series"></param>
    ''' <returns></returns>
    <Runtime.CompilerServices.Extension()>
    Public Function Mean(ByRef Series As List(Of Double)) As Double
        Dim Count = Series.Count
        Dim Sum As Double = 0.0
        For Each item In Series

            Sum += item

        Next
        Mean = Sum / Count
    End Function

    'Morse Code
    ''' <summary>
    ''' Converts Morse code Character to Alphabet
    ''' </summary>
    ''' <param name="Ch"></param>
    ''' <returns></returns>
    <Runtime.CompilerServices.Extension()>
    Public Function MorseToChar(ByRef Ch As String) As String
        Select Case Ch
            Case ".-"
                MorseToChar = "a"
            Case "-..."
                MorseToChar = "b"
            Case "-.-."
                MorseToChar = "c"
            Case "-.."
                MorseToChar = "d"
            Case "."
                MorseToChar = "e"
            Case "..-."
                MorseToChar = "f"
            Case "--."
                MorseToChar = "g"
            Case "...."
                MorseToChar = "h"
            Case ".."
                MorseToChar = "i"
            Case ".---"
                MorseToChar = "j"
            Case "-.-"
                MorseToChar = "k"
            Case ".-.."
                MorseToChar = "l"
            Case "--"
                MorseToChar = "m"
            Case "-."
                MorseToChar = "n"
            Case "---"
                MorseToChar = "o"
            Case ".--."
                MorseToChar = "p"
            Case "--.-"
                MorseToChar = "q"
            Case ".-."
                MorseToChar = "r"
            Case "..."
                MorseToChar = "s"
            Case "-"
                MorseToChar = "t"
            Case "..-"
                MorseToChar = "u"
            Case "...-"
                MorseToChar = "v"
            Case ".--"
                MorseToChar = "w"
            Case "-..-"
                MorseToChar = "x"
            Case "-.--"
                MorseToChar = "y"
            Case "--.."
                MorseToChar = "z"
            Case ".----"
                MorseToChar = "1"
            Case "..---"
                MorseToChar = "2"
            Case "...--"
                MorseToChar = "3"
            Case "....-"
                MorseToChar = "4"
            Case "....."
                MorseToChar = "5"
            Case "-...."
                MorseToChar = "6"
            Case "--..."
                MorseToChar = "7"
            Case "---.."
                MorseToChar = "8"
            Case "----."
                MorseToChar = "9"
            Case "-----"
                MorseToChar = "0"
            Case "   "
                MorseToChar = " "
            Case "^"
                MorseToChar = "."
            Case "~"
                MorseToChar = "-"
            Case Else
                MorseToChar = Ch
        End Select
    End Function

    <Runtime.CompilerServices.Extension()>
    Public Sub MoveFormBottom(ByRef Sender As Form)
        Dim x As Integer
        Dim y As Integer
        x = Screen.PrimaryScreen.WorkingArea.Width - Sender.Width
        y = Screen.PrimaryScreen.WorkingArea.Height - Sender.Height
        'Top
        Sender.Location = New Point(Sender.Location.X, y)
    End Sub

    <Runtime.CompilerServices.Extension()>
    Public Sub MoveFormLeft(ByRef Sender As Form)
        Dim x As Integer
        Dim y As Integer
        x = Screen.PrimaryScreen.WorkingArea.Width - Sender.Width
        y = Screen.PrimaryScreen.WorkingArea.Height - Sender.Height
        'left
        Sender.Location = New Point(0, Sender.Location.Y)
    End Sub

    <Runtime.CompilerServices.Extension()>
    Public Sub MoveFormRight(ByRef Sender As Form)
        Dim x As Integer
        Dim y As Integer
        x = Screen.PrimaryScreen.WorkingArea.Width - Sender.Width
        y = Screen.PrimaryScreen.WorkingArea.Height - Sender.Height
        'left
        Sender.Location = New Point(x, Sender.Location.Y)
    End Sub

    <Runtime.CompilerServices.Extension()>
    Public Sub MoveFormTop(ByRef Sender As Form)
        Dim x As Integer
        Dim y As Integer
        x = Screen.PrimaryScreen.WorkingArea.Width - Sender.Width
        y = Screen.PrimaryScreen.WorkingArea.Height - Sender.Height
        'Top
        Sender.Location = New Point(Sender.Location.X, 0)
    End Sub

    ''' <summary>
    ''' Multiplys X series by Y series
    ''' </summary>
    ''' <param name="X_Series"></param>
    ''' <param name="Y_Series"></param>
    ''' <returns></returns>
    <Runtime.CompilerServices.Extension()>
    Public Function Multiply(ByRef X_Series As List(Of Double), ByRef Y_Series As List(Of Double)) As List(Of Double)
        Dim Count As Integer = X_Series.Count
        Dim Series As New List(Of Double)
        For i = 1 To Count
            Series.Add(X_Series(i) * Y_Series(i))
        Next
        Return Series
    End Function

    ''' <summary>
    ''' Opens text file and returns text
    ''' </summary>
    ''' <param name="PathName">URL of file</param>
    ''' <returns>text in file</returns>
    <System.Runtime.CompilerServices.Extension()>
    Public Function OpenTextFile(ByRef PathName As String) As String
        Dim _text As String = ""
        Try
            If File.Exists(PathName) = True Then
                _text = File.ReadAllText(PathName)
            End If
        Catch ex As Exception
            MsgBox("Error: " & Err.Number & ". " & Err.Description, , NameOf(SaveTextFileAs))

        End Try
        Return _text
    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function PerformTasks(ByRef Txt As String, ByRef Tasks As List(Of String)) As String

        For Each tsk In Tasks
            Select Case tsk

                Case "Space Punctuation"

                    Txt = SpacePunctuation(Txt).Replace("  ", " ")
                Case "To Upper"
                    Txt = Txt.ToUpper.Replace("  ", " ")
                Case "To Lower"
                    Txt = Txt.ToLower.Replace("  ", " ")
                Case "Lemmatize Text"
                Case "Tokenize Characters"
                    Txt = TokenizeChars(Txt)
                    Dim Words() As String = Txt.Split(",")
                    Txt &= vbNewLine & "Total Tokens in doc  -" & Words.Count - 1 & ":" & vbNewLine
                Case "Remove Stop Words"
                    RemoveStopWordsLst(Txt, StopWords)
                Case "Tokenize Words"
                    Txt = TokenizeWords(Txt)
                    Dim Words() As String = Txt.Split(",")
                    Txt &= vbNewLine & "Total Tokens in doc  -" & Words.Count - 1 & ":" & vbNewLine
                Case "Tokenize Sentences"
                    Txt = TokenizeSentences(Txt)
                    Dim Words() As String = Txt.Split(",")
                    Txt &= vbNewLine & "Total Tokens in doc  -" & Words.Count - 2 & ":" & vbNewLine
                Case "Remove Symbols"
                    Txt = RemoveiSymbols(Txt).Replace("  ", " ")
                Case "Remove Brackets"
                    Txt = RemoveBrackets(Txt).Replace("  ", " ")
                Case "Remove Maths Symbols"
                    Txt = RemoveMathsSymbols(Txt).Replace("  ", " ")
                Case "Remove Punctuation"
                    Txt = RemovePunctuation(Txt).Replace("  ", " ")
                Case "AlphaNumeric Only"
                    Txt = AlphaNumericOnly(Txt).Replace("  ", " ")
            End Select
        Next

        Return Txt
    End Function

    ''' <summary>
    ''' Perimeter = 2πSqrt((r1² + r2² / 2)
    ''' = 2 * 3.14 * Sqrt((3² + 4²) / 2)
    ''' = 6.28 * Sqrt((9 + 16) / 2) = 6.28 * Sqrt(25 / 2)
    ''' = 6.28 * Sqrt(12.5) = 6.28 * 3.53 = 22.2. Area = πr1r2 = 3.14 * 3 * 4 = 37.68 .
    ''' </summary>
    ''' <param name="Radius1"></param>
    ''' <param name="Radius2"></param>
    ''' <returns></returns>
    <Runtime.CompilerServices.Extension()>
    Public Function PerimeterOfElipse(ByRef Radius1 As Double, ByRef Radius2 As Double) As Double
        'Perimeter	= 2πSqrt((r1² + r2² / 2)
        '= 2 * 3.14 * Sqrt((3² + 4²) / 2)
        '= 6.28 * Sqrt((9 + 16) / 2) = 6.28 * Sqrt(25 / 2)
        '= 6.28 * Sqrt(12.5) = 6.28 * 3.53 = 22.2.
        'Area = πr1r2 = 3.14 * 3 * 4 = 37.68 .
        PerimeterOfElipse = (2 * Math.PI) * Math.Sqrt(((Radius1 * Radius1) + (Radius2 * Radius2) / 2))

    End Function

    'Phonetics
    ''' <summary>
    ''' returns phonetic character for Letter
    ''' </summary>
    ''' <param name="InputStr"></param>
    ''' <returns></returns>
    <Runtime.CompilerServices.Extension()>
    Public Function Phonetic(ByRef InputStr As String) As String
        Phonetic = ""
        If UCase(InputStr) = "A" Then
            Phonetic = "Alpha"
        End If
        If UCase(InputStr) = "B" Then
            Phonetic = "Bravo"
        End If
        If UCase(InputStr) = "C" Then
            Phonetic = "Charlie"
        End If
        If UCase(InputStr) = "D" Then
            Phonetic = "Delta"
        End If
        If UCase(InputStr) = "E" Then
            Phonetic = "Echo"
        End If
        If UCase(InputStr) = "F" Then
            Phonetic = "Foxtrot"
        End If
        If UCase(InputStr) = "G" Then
            Phonetic = "Golf"
        End If
        If UCase(InputStr) = "H" Then
            Phonetic = "Hotel"
        End If
        If UCase(InputStr) = "I" Then
            Phonetic = "India"
        End If
        If UCase(InputStr) = "J" Then
            Phonetic = "Juliet"
        End If
        If UCase(InputStr) = "K" Then
            Phonetic = "Kilo"
        End If
        If UCase(InputStr) = "L" Then
            Phonetic = "Lima"
        End If
        If UCase(InputStr) = "M" Then
            Phonetic = "Mike"
        End If
        If UCase(InputStr) = "N" Then
            Phonetic = "November"
        End If
        If UCase(InputStr) = "O" Then
            Phonetic = "Oscar"
        End If
        If UCase(InputStr) = "P" Then
            Phonetic = "Papa"
        End If
        If UCase(InputStr) = "Q" Then
            Phonetic = "Quebec"
        End If
        If UCase(InputStr) = "R" Then
            Phonetic = "Romeo"
        End If
        If UCase(InputStr) = "S" Then
            Phonetic = "Sierra"
        End If
        If UCase(InputStr) = "T" Then
            Phonetic = "Tango"
        End If
        If UCase(InputStr) = "U" Then
            Phonetic = "Uniform"
        End If
        If UCase(InputStr) = "V" Then
            Phonetic = "Victor"
        End If
        If UCase(InputStr) = "W" Then
            Phonetic = "Whiskey"
        End If
        If UCase(InputStr) = "X" Then
            Phonetic = "X-Ray"
        End If
        If UCase(InputStr) = "Y" Then
            Phonetic = "Yankee"
        End If
        If UCase(InputStr) = "Z" Then
            Phonetic = "Zulu"
        End If
    End Function

    'Growth
    ''' <summary>
    ''' Given a series of values Predict Value for interval provided based on AverageGrowth
    ''' </summary>
    ''' <param name="Series"></param>
    ''' <param name="Interval"></param>
    ''' <returns></returns>
    <Runtime.CompilerServices.Extension()>
    Public Function PredictGrowth(ByRef Series As List(Of Double), ByRef Interval As Integer) As Double

        Dim GrowthRate As Double = AverageGrowth(Series)
        Dim Past As Double = Series.Last
        Dim Present As Double = Past
        For i = 1 To Interval
            Past = Present
            Present = Past * GrowthRate
        Next
        Return Present
    End Function

    ''' <summary>
    ''' Returns Propercase Sentence
    ''' </summary>
    ''' <param name="TheString">String to be formatted</param>
    ''' <returns></returns>
    <System.Runtime.CompilerServices.Extension()>
    Public Function ProperCase(ByRef TheString As String) As String
        ProperCase = UCase(Left(TheString, 1))

        For i = 2 To Len(TheString)

            ProperCase = If(Mid(TheString, i - 1, 1) = " ", ProperCase & UCase(Mid(TheString, i, 1)), ProperCase & LCase(Mid(TheString, i, 1)))
        Next i
    End Function

    ''' <summary>
    ''' Find Between
    ''' Start ........
    ''' ....................
    ''' ....................Stop
    ''' </summary>
    ''' <param name="Text"></param>
    ''' <param name="StartStr"></param>
    ''' <param name="EndStr"></param>
    ''' <returns></returns>
    <System.Runtime.CompilerServices.Extension()>
    Public Function RegExFindBetween(ByRef Text As String, ByRef StartStr As String, ByRef EndStr As String) As List(Of String)
        Dim Tag = StartStr & "+(.|\n)*" & EndStr
        Dim Searcher As New Regex(Tag)
        Dim iMatch As Match = Searcher.Match(Text)
        Dim iMatches As New List(Of String)
        Do While iMatch.Success
            iMatches.Add(iMatch.Value)
            iMatch = iMatch.NextMatch
        Loop
        Return iMatches
    End Function

    ''' <summary>
    ''' Main Searcher
    ''' </summary>
    ''' <param name="Text">to be searched </param>
    ''' <param name="Pattern">RegEx Search String</param>
    ''' <returns></returns>
    <System.Runtime.CompilerServices.Extension()>
    Public Function RegExSearch(ByRef Text As String, Pattern As String) As List(Of String)
        Dim Searcher As New Regex(Pattern)
        Dim iMatch As Match = Searcher.Match(Text)
        Dim iMatches As New List(Of String)
        Do While iMatch.Success
            iMatches.Add(iMatch.Value)
            iMatch = iMatch.NextMatch
        Loop

        Return iMatches
    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function RemoveBrackets(ByRef Txt As String) As String
        'Brackets
        Txt = Txt.Replace("(", "")
        Txt = Txt.Replace("{", "")
        Txt = Txt.Replace("}", "")
        Txt = Txt.Replace("[", "")
        Txt = Txt.Replace("]", "")
        Return Txt
    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function RemoveFullStop(ByRef MESSAGE As String) As String
Loop1:
        If Right(MESSAGE, 1) = "." Then MESSAGE = Left(MESSAGE, Len(MESSAGE) - 1) : GoTo Loop1
        Return MESSAGE
    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function RemoveiSymbols(ByRef Txt As String) As String
        'Basic Symbols
        Txt = Txt.Replace("£", "")
        Txt = Txt.Replace("$", "")
        Txt = Txt.Replace("^", "")
        Txt = Txt.Replace("@", "")
        Txt = Txt.Replace("#", "")
        Txt = Txt.Replace("~", "")
        Txt = Txt.Replace("\", "")
        Txt = Txt.Replace("\", "")
        Return Txt
    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function RemoveMathsSymbols(ByRef Txt As String) As String
        'Maths Symbols
        Txt = Txt.Replace("+", "")
        Txt = Txt.Replace("=", "")
        Txt = Txt.Replace("-", "")
        Txt = Txt.Replace("/", "")
        Txt = Txt.Replace("*", "")
        Txt = Txt.Replace("<", "")
        Txt = Txt.Replace(">", "")
        Txt = Txt.Replace("%", "")
        Return Txt
    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function RemovePunctuation(ByRef Txt As String) As String
        'Punctuation
        Txt = Txt.Replace(",", "")
        Txt = Txt.Replace(".", "")
        Txt = Txt.Replace(";", "")
        Txt = Txt.Replace("'", "")
        Txt = Txt.Replace("_", "")
        Txt = Txt.Replace("?", "")
        Txt = Txt.Replace("!", "")
        Txt = Txt.Replace("&", "")
        Txt = Txt.Replace(":", "")

        Return Txt
    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function RemoveStopWordsLst(ByRef txt As String, ByRef StopWrds As List(Of String)) As String
        For Each item In StopWrds
            txt = txt.Replace(item, "")
        Next
        Return txt
    End Function

    ''' <summary>
    ''' Creates saves text file as
    ''' </summary>
    ''' <param name="PathName">nfilename and path to file</param>
    ''' <param name="_text">text to be inserted</param>
    <System.Runtime.CompilerServices.Extension()>
    Public Sub SaveTextFileAs(ByRef PathName As String, ByRef _text As String)

        Try
            If File.Exists(PathName) = True Then File.Create(PathName).Dispose()
            Dim alltext As String = _text
            File.WriteAllText(PathName, alltext)
        Catch ex As Exception
            MsgBox("Error: " & Err.Number & ". " & Err.Description, , NameOf(SaveTextFileAs))
        End Try
    End Sub

    ''' <summary>
    ''' Advanced search String pattern Wildcard denotes which position 1st =1 or 2nd =2 Send
    ''' Original input &gt; Search pattern to be used &gt; Wildcard requred SPattern = "WHAT
    ''' COLOUR DO YOU LIKE * OR *" Textstr = "WHAT COLOUR DO YOU LIKE red OR black" ITEM_FOUND =
    ''' = SearchPattern(USERINPUT, SPattern, 1) ---- RETURNS RED ITEM_FOUND = =
    ''' SearchPattern(USERINPUT, SPattern, 1) ---- RETURNS black
    ''' </summary>
    ''' <param name="TextSTR">
    ''' TextStr Required. String.EG: "WHAT COLOUR DO YOU LIKE red OR black"
    ''' </param>
    ''' <param name="SPattern">
    ''' SPattern Required. String.EG: "WHAT COLOUR DO YOU LIKE * OR *"
    ''' </param>
    ''' <param name="Wildcard">Wildcard Required. Integer.EG: 1st =1 or 2nd =2</param>
    ''' <returns></returns>
    ''' <remarks>* in search pattern</remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function SearchPattern(ByRef TextSTR As String, ByRef SPattern As String, ByRef Wildcard As Short) As String
        Dim SearchP2 As String
        Dim SearchP1 As String
        Dim TextStrp3 As String
        Dim TextStrp4 As String
        SearchPattern = ""
        SearchP2 = ""
        SearchP1 = ""
        TextStrp3 = ""
        TextStrp4 = ""
        If TextSTR Like SPattern = True Then
            Select Case Wildcard
                Case 1
                    Call SplitPhrase(SPattern, "*", SearchP1, SearchP2)
                    TextSTR = Replace(TextSTR, SearchP1, "", 1, -1, CompareMethod.Text)

                    SearchP2 = Replace(SearchP2, "*", "", 1, -1, CompareMethod.Text)
                    Call SplitPhrase(TextSTR, SearchP2, TextStrp3, TextStrp4)

                    TextSTR = TextStrp3

                Case 2
                    Call SplitPhrase(SPattern, "*", SearchP1, SearchP2)
                    SPattern = Replace(SPattern, SearchP1, " ", 1, -1, CompareMethod.Text)
                    TextSTR = Replace(TextSTR, SearchP1, " ", 1, -1, CompareMethod.Text)

                    Call SplitPhrase(SearchP2, "*", SearchP1, SearchP2)
                    Call SplitPhrase(TextSTR, SearchP1, TextStrp3, TextStrp4)

                    TextSTR = TextStrp4

            End Select

            SearchPattern = TextSTR
            LTrim(SearchPattern)
            RTrim(SearchPattern)
        Else
        End If

    End Function

    ''' <summary>
    ''' Advanced search String pattern Wildcard denotes which position 1st =1 or 2nd =2 Send
    ''' Original input &gt; Search pattern to be used &gt; Wildcard requred SPattern = "WHAT
    ''' COLOUR DO YOU LIKE * OR *" Textstr = "WHAT COLOUR DO YOU LIKE red OR black" ITEM_FOUND =
    ''' = SearchPattern(USERINPUT, SPattern, 1) ---- RETURNS RED ITEM_FOUND = =
    ''' SearchPattern(USERINPUT, SPattern, 2) ---- RETURNS black
    ''' </summary>
    ''' <param name="TextSTR">TextStr = "Pick Red OR Blue" . String.</param>
    ''' <param name="SPattern">Search String = ("Pick * OR *") String.</param>
    ''' <param name="Wildcard">Wildcard Required. Integer. = 1= Red / 2= Blue</param>
    ''' <returns></returns>
    ''' <remarks>finds the * in search pattern</remarks>
    <System.Runtime.CompilerServices.Extension()>
    Public Function SearchStringbyPattern(ByRef TextSTR As String, ByRef SPattern As String, ByRef Wildcard As Short) As String
        Dim SearchP2 As String
        Dim SearchP1 As String
        Dim TextStrp3 As String
        Dim TextStrp4 As String
        SearchStringbyPattern = ""
        SearchP2 = ""
        SearchP1 = ""
        TextStrp3 = ""
        TextStrp4 = ""
        If TextSTR Like SPattern = True Then
            Select Case Wildcard
                Case 1
                    Call SplitString(SPattern, "*", SearchP1, SearchP2)
                    TextSTR = Replace(TextSTR, SearchP1, "", 1, -1, CompareMethod.Text)

                    SearchP2 = Replace(SearchP2, "*", "", 1, -1, CompareMethod.Text)
                    Call SplitString(TextSTR, SearchP2, TextStrp3, TextStrp4)

                    TextSTR = TextStrp3

                Case 2
                    Call SplitString(SPattern, "*", SearchP1, SearchP2)
                    SPattern = Replace(SPattern, SearchP1, " ", 1, -1, CompareMethod.Text)
                    TextSTR = Replace(TextSTR, SearchP1, " ", 1, -1, CompareMethod.Text)

                    Call SplitString(SearchP2, "*", SearchP1, SearchP2)
                    Call SplitString(TextSTR, SearchP1, TextStrp3, TextStrp4)

                    TextSTR = TextStrp4

            End Select

            SearchStringbyPattern = TextSTR
            LTrim(SearchStringbyPattern)
            RTrim(SearchStringbyPattern)
        Else
        End If

    End Function

    ''' <summary>
    ''' the log-sigmoid function constrains results to the range (0,1), the function is
    ''' sometimes said to be a squashing function in neural network literature. It is the
    ''' non-linear characteristics of the log-sigmoid function (and other similar activation
    ''' functions) that allow neural networks to model complex data.
    ''' </summary>
    ''' <param name="Value"></param>
    ''' <returns></returns>
    ''' <remarks>1 / (1 + Math.Exp(-Value))</remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function Sigmoid(ByRef Value As Integer) As Double
        'z = Bias + (Input*Weight)
        'Output = 1/1+e**z
        Return 1 / (1 + Math.Exp(-Value))
    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function SigmoidDerivitive(ByRef Value As Integer) As Double
        Return Sigmoid(Value) * (1 - Sigmoid(Value))
    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function Signum(ByRef Value As Integer) As Double
        'z = Bias + (Input*Weight)
        'Output = 1/1+e**z
        Return Math.Sign(Value)
    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function SpaceItems(ByRef txt As String, Item As String) As String
        Return txt.Replace(Item, " " & Item & " ")
    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function SpacePunctuation(ByRef Txt As String) As String
        For Each item In Symbols
            Txt = SpaceItems(Txt, item)
        Next
        For Each item In EncapuslationPunctuationEnd
            Txt = SpaceItems(Txt, item)
        Next
        For Each item In EncapuslationPunctuationStart
            Txt = SpaceItems(Txt, item)
        Next
        For Each item In GramaticalPunctuation
            Txt = SpaceItems(Txt, item)
        Next
        For Each item In MathPunctuation
            Txt = SpaceItems(Txt, item)
        Next
        For Each item In MoneyPunctuation
            Txt = SpaceItems(Txt, item)
        Next
        Return Txt
    End Function

    ''' <summary>
    ''' SPLITS THE GIVEN PHRASE UP INTO TWO PARTS by dividing word SplitPhrase(Userinput, "and",
    ''' Firstp, SecondP)
    ''' </summary>
    ''' <param name="PHRASE">Sentence to be divided</param>
    ''' <param name="DIVIDINGWORD">String: Word to divide sentence by</param>
    ''' <param name="FIRSTPART">String: firstpart of sentence to be populated</param>
    ''' <param name="SECONDPART">String: Secondpart of sentence to be populated</param>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Sub SplitPhrase(ByVal PHRASE As String, ByRef DIVIDINGWORD As String, ByRef FIRSTPART As String, ByRef SECONDPART As String)
        Dim POS As Short
        POS = InStr(PHRASE, DIVIDINGWORD)
        If (POS > 0) Then
            FIRSTPART = Trim(Left(PHRASE, POS - 1))
            SECONDPART = Trim(Right(PHRASE, Len(PHRASE) - POS - Len(DIVIDINGWORD) + 1))
        Else
            FIRSTPART = ""
            SECONDPART = PHRASE
        End If
    End Sub

    ''' <summary>
    ''' SPLITS THE GIVEN PHRASE UP INTO TWO PARTS by dividing word SplitPhrase(Userinput, "and",
    ''' Firstp, SecondP)
    ''' </summary>
    ''' <param name="PHRASE">String: Sentence to be divided</param>
    ''' <param name="DIVIDINGWORD">String: Word to divide sentence by</param>
    ''' <param name="FIRSTPART">String-Returned : firstpart of sentence to be populated</param>
    ''' <param name="SECONDPART">String-Returned : Secondpart of sentence to be populated</param>
    ''' <remarks></remarks>
    <System.Runtime.CompilerServices.Extension()>
    Public Sub SplitString(ByVal PHRASE As String, ByRef DIVIDINGWORD As String, ByRef FIRSTPART As String, ByRef SECONDPART As String)
        Dim POS As Short
        'Check Error
        If DIVIDINGWORD IsNot Nothing And PHRASE IsNot Nothing Then

            POS = InStr(PHRASE, DIVIDINGWORD)
            If (POS > 0) Then
                FIRSTPART = Trim(Left(PHRASE, POS - 1))
                SECONDPART = Trim(Right(PHRASE, Len(PHRASE) - POS - Len(DIVIDINGWORD) + 1))
            Else
                FIRSTPART = ""
                SECONDPART = PHRASE
            End If
        Else

        End If
    End Sub

    ''' <summary>
    ''' Split string to List of strings
    ''' </summary>
    ''' <param name="Str">base string</param>
    ''' <param name="Seperator">to be seperated by</param>
    ''' <returns></returns>
    <System.Runtime.CompilerServices.Extension()>
    Public Function SplitToList(ByRef Str As String, ByVal Seperator As String) As List(Of String)
        Dim lst As New List(Of String)
        If Str <> "" = True And Seperator <> "" Then

            Dim Found() As String = Str.Split(Seperator)
            For Each item In Found
                lst.Add(item)
            Next
        Else

        End If
        Return lst
    End Function

    ''' <summary>
    ''' Number multiplied by itself
    ''' </summary>
    ''' <param name="Value"></param>
    ''' <returns></returns>
    <Runtime.CompilerServices.Extension()>
    Public Function Square(ByRef Value As Double) As Double
        Return Value * Value
    End Function

    ''' <summary>
    ''' Square Each value in the series
    ''' </summary>
    ''' <param name="Series"></param>
    ''' <returns></returns>
    <Runtime.CompilerServices.Extension()>
    Public Function Square(ByRef Series As List(Of Double)) As List(Of Double)
        Dim NewSeries As New List(Of Double)
        For Each item In Series
            NewSeries.Add(Square(item))
        Next
        Return NewSeries
    End Function

    ''' <summary>
    ''' squared differences from the Mean.
    ''' </summary>
    ''' <param name="Series"></param>
    ''' <returns></returns>
    <Runtime.CompilerServices.Extension()>
    Public Function SquaredDifferences(ByRef Series As List(Of Double)) As List(Of Double)
        'Results
        Dim NewSeries As New List(Of Double)
        Dim TheMean As Double = Mean(Series)
        For Each item In Series
            'Results
            Dim Difference As Double = 0.0
            Dim NewSum As Double = 0.0
            'For each item Subtract the mean (variance)
            Difference += item - TheMean
            'Square the difference
            NewSum = Square(Difference)
            'Create new series (Squared differences)
            NewSeries.Add(NewSum)
        Next
        Return NewSeries
    End Function

    ''' <summary>
    ''' The Standard Deviation is a measure of how spread out numbers are.
    ''' </summary>
    ''' <param name="Series"></param>
    ''' <returns></returns>
    <Runtime.CompilerServices.Extension()>
    Public Function StandardDeviation(ByRef Series As List(Of Double)) As Double
        'The Square Root of the variance
        Return Math.Sqrt(Variance(Series))
    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function StandardDeviationofSeries(ByVal x As List(Of Double)) As Double

        Dim result As Double
        Dim avg As Double = Average(x)

        For i = 0 To x.Count - 1
            result += Math.Pow((x(i) - avg), 2)
        Next

        result /= x.Count

        Return result

    End Function

    <System.Runtime.CompilerServices.Extension()>
    Public Function StemWords(ByRef txt As List(Of String)) As List(Of String)
        StemWords = New List(Of String)
        For Each wrd In txt
            StemWords.Add((WordStemmer.StemWord(wrd)))
        Next
    End Function

    ''' <summary>
    ''' Sum Series of Values
    ''' </summary>
    ''' <param name="X_Series"></param>
    ''' <returns></returns>
    <Runtime.CompilerServices.Extension()>
    Public Function Sum(ByRef X_Series As List(Of Double)) As Double
        Dim Count As Integer = X_Series.Count
        Dim Ans As Double = 0.0
        For Each i In X_Series
            Ans = +i
        Next
        Return Ans
    End Function

    ''' <summary>
    ''' The Sum of the squared differences from the Mean.
    ''' </summary>
    ''' <param name="Series"></param>
    ''' <returns></returns>
    <Runtime.CompilerServices.Extension()>
    Public Function SumOfSquaredDifferences(ByRef Series As List(Of Double)) As Double
        Dim sum As Double = 0.0
        For Each item In SquaredDifferences(Series)
            sum += item
        Next
        Return sum
    End Function

    ''' <summary>
    ''' Returns a delimited string from the list.
    ''' </summary>
    ''' <param name="ls"></param>
    ''' <param name="delimiter"></param>
    ''' <returns></returns>
    <System.Runtime.CompilerServices.Extension>
    Public Function ToDelimitedString(ls As List(Of String), delimiter As String) As String
        Dim sb As New StringBuilder
        For Each buf As String In ls
            sb.Append(buf)
            sb.Append(delimiter)
        Next
        Return sb.ToString.Trim(CChar(delimiter))
    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function TokenizeBy(ByRef txt As String, ByRef Split As String) As String
        Dim NewTxt As String = ""
        Dim Words() As String = txt.Split(Split)
        For Each item In Words
            NewTxt &= item & ","
        Next
        Return NewTxt
    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function TokenizeChars(ByRef Txt As String) As String

        Dim NewTxt As String = ""
        For Each chr As Char In Txt

            NewTxt &= chr.ToString & ","
        Next

        Return NewTxt
    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function TokenizeSentences(ByRef txt As String) As String
        Dim NewTxt As String = ""
        Dim Words() As String = txt.Split(".")
        For Each item In Words
            NewTxt &= item & ","
        Next
        Return NewTxt
    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function TokenizeWords(ByRef txt As String) As String
        Dim NewTxt As String = ""
        Dim Words() As String = txt.Split(" ")
        For Each item In Words
            NewTxt &= item & ","
        Next
        Return NewTxt
    End Function

    ''' <summary>
    ''' replaces text in file with supplied text
    ''' </summary>
    ''' <param name="PathName">Pathname of file to be updated</param>
    ''' <param name="_text">Text to be inserted</param>
    <System.Runtime.CompilerServices.Extension()>
    Public Sub UpdateTextFileAs(ByRef PathName As String, ByRef _text As String)
        Try
            If File.Exists(PathName) = True Then File.Create(PathName).Dispose()
            Dim alltext As String = _text
            File.AppendAllText(PathName, alltext)
        Catch ex As Exception
            MsgBox("Error: " & Err.Number & ". " & Err.Description, , NameOf(UpdateTextFileAs))
        End Try
    End Sub

    'Statistics
    ''' <summary>
    ''' The average of the squared differences from the Mean.
    ''' </summary>
    ''' <param name="Series"></param>
    ''' <returns></returns>
    <Runtime.CompilerServices.Extension()>
    Public Function Variance(ByRef Series As List(Of Double)) As Double
        Dim TheMean As Double = Mean(Series)

        Dim NewSeries As List(Of Double) = SquaredDifferences(Series)

        'Variance = Average Of the Squared Differences
        Variance = Mean(NewSeries)
    End Function

    ':SHAPES/VOLUMES / Area:
    ''' <summary>
    ''' Comments : Returns the volume of a cone '
    ''' </summary>
    ''' <param name="dblHeight">dblHeight - height of cone</param>
    ''' <param name="dblRadius">radius of cone base</param>
    ''' <returns>volume</returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function VolOfCone(ByVal dblHeight As Double, ByVal dblRadius As Double) As Double
        Const cdblPI As Double = 3.14159265358979
        On Error GoTo PROC_ERR
        VolOfCone = dblHeight * (dblRadius ^ 2) * cdblPI / 3
PROC_EXIT:
        Exit Function
PROC_ERR:
        MsgBox("Error: " & Err.Number & ". " & Err.Description, ,
        NameOf(VolOfCone))
        Resume PROC_EXIT
    End Function

    ''' <summary>
    ''' Comments : Returns the volume of a cylinder
    ''' </summary>
    ''' <param name="dblHeight">dblHeight - height of cylinder</param>
    ''' <param name="dblRadius">radius of cylinder</param>
    ''' <returns>volume</returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function VolOfCylinder(
    ByVal dblHeight As Double,
    ByVal dblRadius As Double) As Double
        Const cdblPI As Double = 3.14159265358979
        On Error GoTo PROC_ERR
        VolOfCylinder = cdblPI * dblHeight * dblRadius ^ 2
PROC_EXIT:
        Exit Function
PROC_ERR:
        MsgBox("Error: " & Err.Number & ". " & Err.Description, ,
        NameOf(VolOfCylinder))
        Resume PROC_EXIT
    End Function

    ''' <summary>
    ''' Returns the volume of a pipe
    ''' </summary>
    ''' <param name="dblHeight">height of pipe</param>
    ''' <param name="dblOuterRadius">outer radius of pipe</param>
    ''' <param name="dblInnerRadius">inner radius of pipe</param>
    ''' <returns>volume</returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function VolOfPipe(
    ByVal dblHeight As Double,
    ByVal dblOuterRadius As Double,
    ByVal dblInnerRadius As Double) _
    As Double
        On Error GoTo PROC_ERR
        VolOfPipe = VolOfCylinder(dblHeight, dblOuterRadius) -
        VolOfCylinder(dblHeight, dblInnerRadius)
PROC_EXIT:
        Exit Function
PROC_ERR:
        MsgBox("Error: " & Err.Number & ". " & Err.Description, ,
        NameOf(VolOfPipe))
        Resume PROC_EXIT
    End Function

    ''' <summary>
    ''' Returns the volume of a pyramid or cone
    ''' </summary>
    ''' <param name="dblHeight">height of pyramid</param>
    ''' <param name="dblBaseArea">area of the base</param>
    ''' <returns>volume</returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function VolOfPyramid(
    ByVal dblHeight As Double,
    ByVal dblBaseArea As Double) As Double
        On Error GoTo PROC_ERR
        VolOfPyramid = dblHeight * dblBaseArea / 3
PROC_EXIT:
        Exit Function
PROC_ERR:
        MsgBox("Error: " & Err.Number & ". " & Err.Description, ,
        NameOf(VolOfPyramid))
        Resume PROC_EXIT
    End Function

    ''' <summary>
    ''' Returns the volume of a sphere
    ''' </summary>
    ''' <param name="dblRadius">dblRadius - radius of the sphere</param>
    ''' <returns>volume</returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function VolOfSphere(ByVal dblRadius As Double) As Double
        Const cdblPI As Double = 3.14159265358979
        On Error GoTo PROC_ERR
        VolOfSphere = cdblPI * (dblRadius ^ 3) * 4 / 3
PROC_EXIT:
        Exit Function
PROC_ERR:
        MsgBox("Error: " & Err.Number & ". " & Err.Description, ,
        NameOf(VolOfSphere))
        Resume PROC_EXIT
    End Function

    ''' <summary>
    ''' Returns the volume of a truncated pyramid
    ''' </summary>
    ''' <param name="dblHeight">dblHeight - height of pyramid</param>
    ''' <param name="dblArea1">area of base</param>
    ''' <param name="dblArea2">area of top</param>
    ''' <returns>volume</returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function VolOfTruncPyramid(
    ByVal dblHeight As Double,
    ByVal dblArea1 As Double,
    ByVal dblArea2 As Double) _
    As Double
        On Error GoTo PROC_ERR
        VolOfTruncPyramid = dblHeight * (dblArea1 + dblArea2 + Math.Sqrt(dblArea1) *
        Math.Sqrt(dblArea2)) / 3
PROC_EXIT:
        Exit Function
PROC_ERR:
        MsgBox("Error: " & Err.Number & ". " & Err.Description, ,
        NameOf(VolOfTruncPyramid))
        Resume PROC_EXIT
    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function VolumeOfElipse(ByRef Radius1 As Double, ByRef Radius2 As Double, ByRef Radius3 As Double) As Double
        ' Case 2:

        ' Find the volume of an ellipse with the given radii 3, 4, 5.
        'Step 1:

        ' Find the volume. Volume = (4/3)πr1r2r3= (4/3) * 3.14 * 3 * 4 * 5 = 1.33 * 188.4 = 251

        VolumeOfElipse = ((4 / 3) * Math.PI) * Radius1 * Radius2 * Radius3

    End Function

    ''' <summary>
    ''' Counts the vowels used (AEIOU)
    ''' </summary>
    ''' <param name="InputString"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Function VowelCount(ByVal InputString As String) As Integer
        Dim v(9) As String 'Declare an array  of 10 elements 0 to 9
        Dim vcount As Short 'This variable will contain number of vowels
        Dim flag As Integer
        Dim strLen As Integer
        Dim i As Integer
        v(0) = "a" 'First element of array is assigned small a
        v(1) = "i"
        v(2) = "o"
        v(3) = "u"
        v(4) = "e"
        v(5) = "A" 'Sixth element is assigned Capital A
        v(6) = "I"
        v(7) = "O"
        v(8) = "U"
        v(9) = "E"
        strLen = Len(InputString)

        For flag = 1 To strLen 'It will get every letter of entered string and loop
            'will terminate when all letters have been examined

            For i = 0 To 9 'Takes every element of v(9) one by one
                'Check if current letter is a vowel
                If Mid(InputString, flag, 1) = v(i) Then
                    vcount = vcount + 1 ' If letter is equal to vowel
                    'then increment vcount by 1
                End If
            Next i 'Consider next value of v(i)
        Next flag 'Consider next letter of the enterd string

        VowelCount = vcount

    End Function

#Region "Remove StopWords"

    Public Function RemoveAlphaBet(ByRef txt As String) As String
        Return RemoveElements(txt, AlphaBet)
    End Function

    ''' <summary>
    ''' Ie: PunctuationList, Numerical List etc
    ''' </summary>
    ''' <param name="txt"></param>
    ''' <param name="StopWrdsElement">PunctuationList, NumericalList</param>
    ''' <returns></returns>
    Public Function RemoveElements(ByRef txt As String, ByRef StopWrdsElement() As String)
        For Each item In StopWrdsElement
            txt = txt.Replace(item, "")
        Next
        Return txt
    End Function

    Public Function RemoveEncapsulationPunctuation(ByRef txt As String) As String

        txt = RemoveElements(txt, EncapuslationPunctuationStart)
        txt = RemoveElements(txt, EncapuslationPunctuationStart)
        Return txt
    End Function

    Public Function RemoveGramaticalPunctuation(ByRef txt As String) As String
        Return RemoveElements(txt, GramaticalPunctuation)
    End Function

    Public Function RemoveMathPunctuation(ByRef txt As String) As String
        Return RemoveElements(txt, MathPunctuation)
    End Function

    Public Function RemoveMoneyPunctuation(ByRef txt As String) As String
        Return RemoveElements(txt, MoneyPunctuation)
    End Function

    Public Function RemoveNumerical(ByRef txt As String) As String
        Return RemoveElements(txt, MoneyPunctuation)
    End Function

    Public Function RemoveStopWords(ByRef txt As String, ByRef StopWrds As List(Of String)) As String
        For Each item In StopWrds
            txt = txt.Replace(item, "")
        Next
        Return txt
    End Function

    Public Function RemoveSymbols(ByRef txt As String) As String
        Return RemoveElements(txt, Symbols)
    End Function

#End Region

    ''' <summary>
    ''' Removes StopWords from sentence
    ''' ARAB/ENG/DUTCH/FRENCH/SPANISH/ITALIAN
    ''' Hopefully leaving just relevant words in the user sentence
    ''' Currently Under Revision (takes too many words)
    ''' </summary>
    ''' <param name="Userinput"></param>
    ''' <returns></returns>
    <Runtime.CompilerServices.Extension()>
    Public Function RemoveStopWords(ByRef Userinput As String) As String
        ' Userinput = LCase(Userinput).Replace("the", "r")
        For Each item In StopWordsENG
            Userinput = LCase(Userinput).Replace(item, "")
        Next
        For Each item In StopWordsArab
            Userinput = Userinput.Replace(item, "")
        Next
        For Each item In StopWordsDutch
            Userinput = Userinput.Replace(item, "")
        Next
        For Each item In StopWordsFrench
            Userinput = Userinput.Replace(item, "")
        Next
        For Each item In StopWordsItalian
            Userinput = Userinput.Replace(item, "")
        Next
        For Each item In StopWordsSpanish
            Userinput = Userinput.Replace(item, "")
        Next
        Return Userinput
    End Function

#Region "Words"

    Public StopWordsENG() As String = {"a", "as", "able", "about", "above", "according", "accordingly", "across", "actually", "after", "afterwards", "again", "against", "aint",
"all", "allow", "allows", "almost", "alone", "along", "already", "also", "although", "always", "am", "among", "amongst", "an", "and", "another", "any",
"anybody", "anyhow", "anyone", "anything", "anyway", "anyways", "anywhere", "apart", "appear", "appreciate", "appropriate", "are", "arent", "around",
"as", "aside", "ask", "asking", "associated", "at", "available", "away", "awfully", "b", "be", "became", "because", "become", "becomes", "becoming",
"been", "before", "beforehand", "behind", "being", "believe", "below", "beside", "besides", "best", "better", "between", "beyond", "both", "brief",
"but", "by", "c", "cmon", "cs", "came", "can", "cant", "cannot", "cant", "cause", "causes", "certain", "certainly", "changes", "clearly", "co", "com",
"come", "comes", "concerning", "consequently", "consider", "considering", "contain", "containing", "contains", "corresponding", "could", "couldnt",
"course", "currently", "d", "definitely", "described", "despite", "did", "didnt", "different", "do", "does", "doesnt", "doing", "dont", "done", "down",
"downwards", "during", "e", "each", "edu", "eg", "eight", "either", "else", "elsewhere", "enough", "entirely", "especially", "et", "etc", "even", "ever",
"every", "everybody", "everyone", "everything", "everywhere", "ex", "exactly", "example", "except", "f", "far", "few", "fifth", "first", "five", "followed",
"following", "follows", "for", "former", "formerly", "forth", "four", "from", "further", "furthermore", "g", "get", "gets", "getting", "given", "gives",
"go", "goes", "going", "gone", "got", "gotten", "greetings", "h", "had", "hadnt", "happens", "hardly", "has", "hasnt", "have", "havent", "having", "he",
"hes", "hello", "help", "hence", "her", "here", "heres", "hereafter", "hereby", "herein", "hereupon", "hers", "herself", "hi", "him", "himself", "his",
"hither", "hopefully", "how", "howbeit", "however", "i", "id", "ill", "im", "ive", "ie", "if", "ignored", "immediate", "in", "inasmuch", "inc", "indeed",
"indicate", "indicated", "indicates", "inner", "insofar", "instead", "into", "inward", "is", "isnt", "it", "itd", "itll", "its", "its", "itself", "j",
"just", "k", "keep", "keeps", "kept", "know", "known", "knows", "l", "last", "lately", "later", "latter", "latterly", "least", "less", "lest", "let", "lets",
"like", "liked", "likely", "little", "look", "looking", "looks", "ltd", "m", "mainly", "many", "may", "maybe", "me", "mean", "meanwhile", "merely", "might",
"more", "moreover", "most", "mostly", "much", "must", "my", "myself", "n", "name", "namely", "nd", "near", "nearly", "necessary", "need", "needs", "neither",
"never", "nevertheless", "new", "next", "nine", "no", "nobody", "non", "none", "noone", "nor", "normally", "not", "nothing", "novel", "now", "nowhere", "o",
"obviously", "of", "off", "often", "oh", "ok", "okay", "old", "on", "once", "one", "ones", "only", "onto", "or", "other", "others", "otherwise", "ought", "our",
"ours", "ourselves", "out", "outside", "over", "overall", "own", "p", "particular", "particularly", "per", "perhaps", "placed", "please", "plus", "possible",
"presumably", "probably", "provides", "q", "que", "quite", "qv", "r", "rather", "rd", "re", "really", "reasonably", "regarding", "regardless", "regards",
"relatively", "respectively", "right", "s", "said", "same", "saw", "say", "saying", "says", "second", "secondly", "see", "seeing", "seem", "seemed", "seeming",
"seems", "seen", "self", "selves", "sensible", "sent", "serious", "seriously", "seven", "several", "shall", "she", "should", "shouldnt", "since", "six", "so",
"some", "somebody", "somehow", "someone", "something", "sometime", "sometimes", "somewhat", "somewhere", "soon", "sorry", "specified", "specify", "specifying",
"still", "sub", "such", "sup", "sure", "t", "ts", "take", "taken", "tell", "tends", "th", "than", "thank", "thanks", "thanx", "that", "thats", "thats", "the",
"their", "theirs", "them", "themselves", "then", "thence", "there", "theres", "thereafter", "thereby", "therefore", "therein", "theres", "thereupon",
"these", "they", "theyd", "theyll", "theyre", "theyve", "think", "third", "this", "thorough", "thoroughly", "those", "though", "three", "through",
"throughout", "thru", "thus", "to", "together", "too", "took", "toward", "towards", "tried", "tries", "truly", "try", "trying", "twice", "two", "u", "un",
"under", "unfortunately", "unless", "unlikely", "until", "unto", "up", "upon", "us", "use", "used", "useful", "uses", "using", "usually", "uucp", "v",
"value", "various", "very", "via", "viz", "vs", "w", "want", "wants", "was", "wasnt", "way", "we", "wed", "well", "were", "weve", "welcome", "well",
"went", "were", "werent", "what", "whats", "whatever", "when", "whence", "whenever", "where", "wheres", "whereafter", "whereas", "whereby", "wherein",
"whereupon", "wherever", "whether", "which", "while", "whither", "who", "whos", "whoever", "whole", "whom", "whose", "why", "will", "willing", "wish",
"with", "within", "without", "wont", "wonder", "would", "wouldnt", "x", "y", "yes", "yet", "you", "youd", "youll", "youre", "youve", "your", "yours",
"yourself", "yourselves", "youll", "z", "zero"}

    Public StopWordsArab() As String = {"،", "آض", "آمينَ", "آه",
    "آهاً", "آي", "أ", "أب", "أجل", "أجمع", "أخ", "أخذ", "أصبح", "أضحى", "أقبل",
    "أقل", "أكثر", "ألا", "أم", "أما", "أمامك", "أمامكَ", "أمسى", "أمّا", "أن", "أنا", "أنت",
    "أنتم", "أنتما", "أنتن", "أنتِ", "أنشأ", "أنّى", "أو", "أوشك", "أولئك", "أولئكم", "أولاء",
    "أولالك", "أوّهْ", "أي", "أيا", "أين", "أينما", "أيّ", "أَنَّ", "أََيُّ", "أُفٍّ", "إذ", "إذا", "إذاً",
    "إذما", "إذن", "إلى", "إليكم", "إليكما", "إليكنّ", "إليكَ", "إلَيْكَ", "إلّا", "إمّا", "إن", "إنّما",
    "إي", "إياك", "إياكم", "إياكما", "إياكن", "إيانا", "إياه", "إياها", "إياهم", "إياهما", "إياهن",
    "إياي", "إيهٍ", "إِنَّ", "ا", "ابتدأ", "اثر", "اجل", "احد", "اخرى", "اخلولق", "اذا", "اربعة", "ارتدّ",
    "استحال", "اطار", "اعادة", "اعلنت", "اف", "اكثر", "اكد", "الألاء", "الألى", "الا", "الاخيرة", "الان", "الاول",
    "الاولى", "التى", "التي", "الثاني", "الثانية", "الذاتي", "الذى", "الذي", "الذين", "السابق", "الف", "اللائي",
    "اللاتي", "اللتان", "اللتيا", "اللتين", "اللذان", "اللذين", "اللواتي", "الماضي", "المقبل", "الوقت", "الى",
    "اليوم", "اما", "امام", "امس", "ان", "انبرى", "انقلب", "انه", "انها", "او", "اول", "اي", "ايار", "ايام",
    "ايضا", "ب", "بات", "باسم", "بان", "بخٍ", "برس", "بسبب", "بسّ", "بشكل", "بضع", "بطآن", "بعد", "بعض", "بك",
    "بكم", "بكما", "بكن", "بل", "بلى", "بما", "بماذا", "بمن", "بن", "بنا", "به", "بها", "بي", "بيد", "بين",
    "بَسْ", "بَلْهَ", "بِئْسَ", "تانِ", "تانِك", "تبدّل", "تجاه", "تحوّل", "تلقاء", "تلك", "تلكم", "تلكما", "تم", "تينك",
    "تَيْنِ", "تِه", "تِي", "ثلاثة", "ثم", "ثمّ", "ثمّة", "ثُمَّ", "جعل", "جلل", "جميع", "جير", "حار", "حاشا", "حاليا",
    "حاي", "حتى", "حرى", "حسب", "حم", "حوالى", "حول", "حيث", "حيثما", "حين", "حيَّ", "حَبَّذَا", "حَتَّى", "حَذارِ", "خلا",
    "خلال", "دون", "دونك", "ذا", "ذات", "ذاك", "ذانك", "ذانِ", "ذلك", "ذلكم", "ذلكما", "ذلكن", "ذو", "ذوا", "ذواتا", "ذواتي", "ذيت", "ذينك", "ذَيْنِ", "ذِه", "ذِي", "راح", "رجع", "رويدك", "ريث", "رُبَّ", "زيارة", "سبحان", "سرعان", "سنة", "سنوات", "سوف", "سوى", "سَاءَ", "سَاءَمَا", "شبه", "شخصا", "شرع", "شَتَّانَ", "صار", "صباح", "صفر", "صهٍ", "صهْ", "ضد", "ضمن", "طاق", "طالما", "طفق", "طَق", "ظلّ", "عاد", "عام", "عاما", "عامة", "عدا", "عدة", "عدد", "عدم", "عسى", "عشر", "عشرة", "علق", "على", "عليك", "عليه", "عليها", "علًّ", "عن", "عند", "عندما", "عوض", "عين", "عَدَسْ", "عَمَّا", "غدا", "غير", "ـ", "ف", "فان", "فلان", "فو", "فى", "في", "فيم", "فيما", "فيه", "فيها", "قال", "قام", "قبل", "قد", "قطّ", "قلما", "قوة", "كأنّما", "كأين", "كأيّ", "كأيّن", "كاد", "كان", "كانت", "كذا", "كذلك", "كرب", "كل", "كلا", "كلاهما", "كلتا", "كلم", "كليكما", "كليهما", "كلّما", "كلَّا", "كم", "كما", "كي", "كيت", "كيف", "كيفما", "كَأَنَّ", "كِخ", "لئن", "لا", "لات", "لاسيما", "لدن", "لدى", "لعمر", "لقاء", "لك", "لكم", "لكما", "لكن", "لكنَّما", "لكي", "لكيلا", "للامم", "لم", "لما", "لمّا", "لن", "لنا", "له", "لها", "لو", "لوكالة", "لولا", "لوما", "لي", "لَسْتَ", "لَسْتُ", "لَسْتُم", "لَسْتُمَا", "لَسْتُنَّ", "لَسْتِ", "لَسْنَ", "لَعَلَّ", "لَكِنَّ", "لَيْتَ", "لَيْسَ", "لَيْسَا", "لَيْسَتَا", "لَيْسَتْ", "لَيْسُوا", "لَِسْنَا", "ما", "ماانفك", "مابرح", "مادام", "ماذا", "مازال", "مافتئ", "مايو", "متى", "مثل", "مذ", "مساء", "مع", "معاذ", "مقابل", "مكانكم", "مكانكما", "مكانكنّ", "مكانَك", "مليار", "مليون", "مما", "ممن", "من", "منذ", "منها", "مه", "مهما", "مَنْ", "مِن", "نحن", "نحو", "نعم", "نفس", "نفسه", "نهاية", "نَخْ", "نِعِمّا", "نِعْمَ", "ها", "هاؤم", "هاكَ", "هاهنا", "هبّ", "هذا", "هذه", "هكذا", "هل", "هلمَّ", "هلّا", "هم", "هما", "هن", "هنا", "هناك", "هنالك", "هو", "هي", "هيا", "هيت", "هيّا", "هَؤلاء", "هَاتانِ", "هَاتَيْنِ", "هَاتِه", "هَاتِي", "هَجْ", "هَذا", "هَذانِ", "هَذَيْنِ", "هَذِه", "هَذِي", "هَيْهَاتَ", "و", "و6", "وا", "واحد", "واضاف", "واضافت", "واكد", "وان", "واهاً", "واوضح", "وراءَك", "وفي", "وقال", "وقالت", "وقد", "وقف", "وكان", "وكانت", "ولا", "ولم",
    "ومن", "وهو", "وهي", "ويكأنّ", "وَيْ", "وُشْكَانََ", "يكون", "يمكن", "يوم", "ّأيّان"}

    Public StopWordsItalian() As String = {"IE", "a", "abbastanza", "abbia", "abbiamo", "abbiano", "abbiate", "accidenti", "ad", "adesso", "affinche", "agl", "agli",
"ahime", "ahimè", "ai", "al", "alcuna", "alcuni", "alcuno", "all", "alla", "alle", "allo", "allora", "altri", "altrimenti", "altro",
"altrove", "altrui", "anche", "ancora", "anni", "anno", "ansa", "anticipo", "assai", "attesa", "attraverso", "avanti", "avemmo",
"avendo", "avente", "aver", "avere", "averlo", "avesse", "avessero", "avessi", "avessimo", "aveste", "avesti", "avete", "aveva",
"avevamo", "avevano", "avevate", "avevi", "avevo", "avrai", "avranno", "avrebbe", "avrebbero", "avrei", "avremmo", "avremo",
"avreste", "avresti", "avrete", "avrà", "avrò", "avuta", "avute", "avuti", "avuto", "basta", "bene", "benissimo", "berlusconi",
"brava", "bravo", "c", "casa", "caso", "cento", "certa", "certe", "certi", "certo", "che", "chi", "chicchessia", "chiunque", "ci",
"ciascuna", "ciascuno", "cima", "cio", "cioe", "cioè", "circa", "citta", "città", "ciò", "co", "codesta", "codesti", "codesto",
"cogli", "coi", "col", "colei", "coll", "coloro", "colui", "come", "cominci", "comunque", "con", "concernente", "conciliarsi",
"conclusione", "consiglio", "contro", "cortesia", "cos", "cosa", "cosi", "così", "cui", "d", "da", "dagl", "dagli", "dai", "dal",
"dall", "dalla", "dalle", "dallo", "dappertutto", "davanti", "degl", "degli", "dei", "del", "dell", "della", "delle", "dello",
"dentro", "detto", "deve", "di", "dice", "dietro", "dire", "dirimpetto", "diventa", "diventare", "diventato", "dopo", "dov", "dove",
"dovra", "dovrà", "dovunque", "due", "dunque", "durante", "e", "ebbe", "ebbero", "ebbi", "ecc", "ecco", "ed", "effettivamente", "egli",
"ella", "entrambi", "eppure", "era", "erano", "eravamo", "eravate", "eri", "ero", "esempio", "esse", "essendo", "esser", "essere",
"essi", "ex", "fa", "faccia", "facciamo", "facciano", "facciate", "faccio", "facemmo", "facendo", "facesse", "facessero", "facessi",
"facessimo", "faceste", "facesti", "faceva", "facevamo", "facevano", "facevate", "facevi", "facevo", "fai", "fanno", "farai",
"faranno", "fare", "farebbe", "farebbero", "farei", "faremmo", "faremo", "fareste", "faresti", "farete", "farà", "farò", "fatto",
"favore", "fece", "fecero", "feci", "fin", "finalmente", "finche", "fine", "fino", "forse", "forza", "fosse", "fossero", "fossi",
"fossimo", "foste", "fosti", "fra", "frattempo", "fu", "fui", "fummo", "fuori", "furono", "futuro", "generale", "gia", "giacche",
"giorni", "giorno", "già", "gli", "gliela", "gliele", "glieli", "glielo", "gliene", "governo", "grande", "grazie", "gruppo", "ha",
"haha", "hai", "hanno", "ho", "i", "ieri", "il", "improvviso", "in", "inc", "infatti", "inoltre", "insieme", "intanto", "intorno",
"invece", "io", "l", "la", "lasciato", "lato", "lavoro", "le", "lei", "li", "lo", "lontano", "loro", "lui", "lungo", "luogo", "là",
"ma", "macche", "magari", "maggior", "mai", "male", "malgrado", "malissimo", "mancanza", "marche", "me", "medesimo", "mediante",
"meglio", "meno", "mentre", "mesi", "mezzo", "mi", "mia", "mie", "miei", "mila", "miliardi", "milioni", "minimi", "ministro",
"mio", "modo", "molti", "moltissimo", "molto", "momento", "mondo", "mosto", "nazionale", "ne", "negl", "negli", "nei", "nel",
"nell", "nella", "nelle", "nello", "nemmeno", "neppure", "nessun", "nessuna", "nessuno", "niente", "no", "noi", "non", "nondimeno",
"nonostante", "nonsia", "nostra", "nostre", "nostri", "nostro", "novanta", "nove", "nulla", "nuovo", "o", "od", "oggi", "ogni",
"ognuna", "ognuno", "oltre", "oppure", "ora", "ore", "osi", "ossia", "ottanta", "otto", "paese", "parecchi", "parecchie",
"parecchio", "parte", "partendo", "peccato", "peggio", "per", "perche", "perchè", "perché", "percio", "perciò", "perfino", "pero",
"persino", "persone", "però", "piedi", "pieno", "piglia", "piu", "piuttosto", "più", "po", "pochissimo", "poco", "poi", "poiche",
"possa", "possedere", "posteriore", "posto", "potrebbe", "preferibilmente", "presa", "press", "prima", "primo", "principalmente",
"probabilmente", "proprio", "puo", "pure", "purtroppo", "può", "qualche", "qualcosa", "qualcuna", "qualcuno", "quale", "quali",
"qualunque", "quando", "quanta", "quante", "quanti", "quanto", "quantunque", "quasi", "quattro", "quel", "quella", "quelle",
"quelli", "quello", "quest", "questa", "queste", "questi", "questo", "qui", "quindi", "realmente", "recente", "recentemente",
"registrazione", "relativo", "riecco", "salvo", "sara", "sarai", "saranno", "sarebbe", "sarebbero", "sarei", "saremmo", "saremo",
"sareste", "saresti", "sarete", "sarà", "sarò", "scola", "scopo", "scorso", "se", "secondo", "seguente", "seguito", "sei", "sembra",
"sembrare", "sembrato", "sembri", "sempre", "senza", "sette", "si", "sia", "siamo", "siano", "siate", "siete", "sig", "solito",
"solo", "soltanto", "sono", "sopra", "sotto", "spesso", "srl", "sta", "stai", "stando", "stanno", "starai", "staranno", "starebbe",
"starebbero", "starei", "staremmo", "staremo", "stareste", "staresti", "starete", "starà", "starò", "stata", "state", "stati",
"stato", "stava", "stavamo", "stavano", "stavate", "stavi", "stavo", "stemmo", "stessa", "stesse", "stessero", "stessi", "stessimo",
"stesso", "steste", "stesti", "stette", "stettero", "stetti", "stia", "stiamo", "stiano", "stiate", "sto", "su", "sua", "subito",
"successivamente", "successivo", "sue", "sugl", "sugli", "sui", "sul", "sull", "sulla", "sulle", "sullo", "suo", "suoi", "tale",
"tali", "talvolta", "tanto", "te", "tempo", "ti", "titolo", "torino", "tra", "tranne", "tre", "trenta", "troppo", "trovato", "tu",
"tua", "tue", "tuo", "tuoi", "tutta", "tuttavia", "tutte", "tutti", "tutto", "uguali", "ulteriore", "ultimo", "un", "una", "uno",
"uomo", "va", "vale", "vari", "varia", "varie", "vario", "verso", "vi", "via", "vicino", "visto", "vita", "voi", "volta", "volte",
"vostra", "vostre", "vostri", "vostro", "è"}

    Public StopWordsSpanish() As String = {"a", "actualmente", "acuerdo", "adelante", "ademas", "además", "adrede", "afirmó", "agregó", "ahi", "ahora",
"ahí", "al", "algo", "alguna", "algunas", "alguno", "algunos", "algún", "alli", "allí", "alrededor", "ambos", "ampleamos",
"antano", "antaño", "ante", "anterior", "antes", "apenas", "aproximadamente", "aquel", "aquella", "aquellas", "aquello",
"aquellos", "aqui", "aquél", "aquélla", "aquéllas", "aquéllos", "aquí", "arriba", "arribaabajo", "aseguró", "asi", "así",
"atras", "aun", "aunque", "ayer", "añadió", "aún", "b", "bajo", "bastante", "bien", "breve", "buen", "buena", "buenas", "bueno",
"buenos", "c", "cada", "casi", "cerca", "cierta", "ciertas", "cierto", "ciertos", "cinco", "claro", "comentó", "como", "con",
"conmigo", "conocer", "conseguimos", "conseguir", "considera", "consideró", "consigo", "consigue", "consiguen", "consigues",
"contigo", "contra", "cosas", "creo", "cual", "cuales", "cualquier", "cuando", "cuanta", "cuantas", "cuanto", "cuantos", "cuatro",
"cuenta", "cuál", "cuáles", "cuándo", "cuánta", "cuántas", "cuánto", "cuántos", "cómo", "d", "da", "dado", "dan", "dar", "de",
"debajo", "debe", "deben", "debido", "decir", "dejó", "del", "delante", "demasiado", "demás", "dentro", "deprisa", "desde",
"despacio", "despues", "después", "detras", "detrás", "dia", "dias", "dice", "dicen", "dicho", "dieron", "diferente", "diferentes",
"dijeron", "dijo", "dio", "donde", "dos", "durante", "día", "días", "dónde", "e", "ejemplo", "el", "ella", "ellas", "ello", "ellos",
"embargo", "empleais", "emplean", "emplear", "empleas", "empleo", "en", "encima", "encuentra", "enfrente", "enseguida", "entonces",
"entre", "era", "eramos", "eran", "eras", "eres", "es", "esa", "esas", "ese", "eso", "esos", "esta", "estaba", "estaban", "estado",
"estados", "estais", "estamos", "estan", "estar", "estará", "estas", "este", "esto", "estos", "estoy", "estuvo", "está", "están", "ex",
"excepto", "existe", "existen", "explicó", "expresó", "f", "fin", "final", "fue", "fuera", "fueron", "fui", "fuimos", "g", "general",
"gran", "grandes", "gueno", "h", "ha", "haber", "habia", "habla", "hablan", "habrá", "había", "habían", "hace", "haceis", "hacemos",
"hacen", "hacer", "hacerlo", "haces", "hacia", "haciendo", "hago", "han", "hasta", "hay", "haya", "he", "hecho", "hemos", "hicieron",
"hizo", "horas", "hoy", "hubo", "i", "igual", "incluso", "indicó", "informo", "informó", "intenta", "intentais", "intentamos", "intentan",
"intentar", "intentas", "intento", "ir", "j", "junto", "k", "l", "la", "lado", "largo", "las", "le", "lejos", "les", "llegó", "lleva",
"llevar", "lo", "los", "luego", "lugar", "m", "mal", "manera", "manifestó", "mas", "mayor", "me", "mediante", "medio", "mejor", "mencionó",
"menos", "menudo", "mi", "mia", "mias", "mientras", "mio", "mios", "mis", "misma", "mismas", "mismo", "mismos", "modo", "momento", "mucha",
"muchas", "mucho", "muchos", "muy", "más", "mí", "mía", "mías", "mío", "míos", "n", "nada", "nadie", "ni", "ninguna", "ningunas", "ninguno",
"ningunos", "ningún", "no", "nos", "nosotras", "nosotros", "nuestra", "nuestras", "nuestro", "nuestros", "nueva", "nuevas", "nuevo",
"nuevos", "nunca", "o", "ocho", "os", "otra", "otras", "otro", "otros", "p", "pais", "para", "parece", "parte", "partir", "pasada",
"pasado", "paìs", "peor", "pero", "pesar", "poca", "pocas", "poco", "pocos", "podeis", "podemos", "poder", "podria", "podriais",
"podriamos", "podrian", "podrias", "podrá", "podrán", "podría", "podrían", "poner", "por", "porque", "posible", "primer", "primera",
"primero", "primeros", "principalmente", "pronto", "propia", "propias", "propio", "propios", "proximo", "próximo", "próximos", "pudo",
"pueda", "puede", "pueden", "puedo", "pues", "q", "qeu", "que", "quedó", "queremos", "quien", "quienes", "quiere", "quiza", "quizas",
"quizá", "quizás", "quién", "quiénes", "qué", "r", "raras", "realizado", "realizar", "realizó", "repente", "respecto", "s", "sabe",
"sabeis", "sabemos", "saben", "saber", "sabes", "salvo", "se", "sea", "sean", "segun", "segunda", "segundo", "según", "seis", "ser",
"sera", "será", "serán", "sería", "señaló", "si", "sido", "siempre", "siendo", "siete", "sigue", "siguiente", "sin", "sino", "sobre",
"sois", "sola", "solamente", "solas", "solo", "solos", "somos", "son", "soy", "soyos", "su", "supuesto", "sus", "suya", "suyas", "suyo",
"sé", "sí", "sólo", "t", "tal", "tambien", "también", "tampoco", "tan", "tanto", "tarde", "te", "temprano", "tendrá", "tendrán", "teneis",
"tenemos", "tener", "tenga", "tengo", "tenido", "tenía", "tercera", "ti", "tiempo", "tiene", "tienen", "toda", "todas", "todavia",
"todavía", "todo", "todos", "total", "trabaja", "trabajais", "trabajamos", "trabajan", "trabajar", "trabajas", "trabajo", "tras",
"trata", "través", "tres", "tu", "tus", "tuvo", "tuya", "tuyas", "tuyo", "tuyos", "tú", "u", "ultimo", "un", "una", "unas", "uno", "unos",
"usa", "usais", "usamos", "usan", "usar", "usas", "uso", "usted", "ustedes", "v", "va", "vais", "valor", "vamos", "van", "varias", "varios",
"vaya", "veces", "ver", "verdad", "verdadera", "verdadero", "vez", "vosotras", "vosotros", "voy", "vuestra", "vuestras", "vuestro",
"vuestros", "w", "x", "y", "ya", "yo", "z", "él", "ésa", "ésas", "ése", "ésos", "ésta", "éstas", "éste", "éstos", "última", "últimas",
"último", "últimos"}

    Public StopWordsFrench() As String = {"a", "abord", "absolument", "afin", "ah", "ai", "aie", "ailleurs", "ainsi", "ait", "allaient", "allo", "allons",
"allô", "alors", "anterieur", "anterieure", "anterieures", "apres", "après", "as", "assez", "attendu", "au", "aucun", "aucune",
"aujourd", "aujourd'hui", "aupres", "auquel", "aura", "auraient", "aurait", "auront", "aussi", "autre", "autrefois", "autrement",
"autres", "autrui", "aux", "auxquelles", "auxquels", "avaient", "avais", "avait", "avant", "avec", "avoir", "avons", "ayant", "b",
"bah", "bas", "basee", "bat", "beau", "beaucoup", "bien", "bigre", "boum", "bravo", "brrr", "c", "car", "ce", "ceci", "cela", "celle",
"celle-ci", "celle-là", "celles", "celles-ci", "celles-là", "celui", "celui-ci", "celui-là", "cent", "cependant", "certain",
"certaine", "certaines", "certains", "certes", "ces", "cet", "cette", "ceux", "ceux-ci", "ceux-là", "chacun", "chacune", "chaque",
"cher", "chers", "chez", "chiche", "chut", "chère", "chères", "ci", "cinq", "cinquantaine", "cinquante", "cinquantième", "cinquième",
"clac", "clic", "combien", "comme", "comment", "comparable", "comparables", "compris", "concernant", "contre", "couic", "crac", "d",
"da", "dans", "de", "debout", "dedans", "dehors", "deja", "delà", "depuis", "dernier", "derniere", "derriere", "derrière", "des",
"desormais", "desquelles", "desquels", "dessous", "dessus", "deux", "deuxième", "deuxièmement", "devant", "devers", "devra",
"different", "differentes", "differents", "différent", "différente", "différentes", "différents", "dire", "directe", "directement",
"dit", "dite", "dits", "divers", "diverse", "diverses", "dix", "dix-huit", "dix-neuf", "dix-sept", "dixième", "doit", "doivent", "donc",
"dont", "douze", "douzième", "dring", "du", "duquel", "durant", "dès", "désormais", "e", "effet", "egale", "egalement", "egales", "eh",
"elle", "elle-même", "elles", "elles-mêmes", "en", "encore", "enfin", "entre", "envers", "environ", "es", "est", "et", "etant", "etc",
"etre", "eu", "euh", "eux", "eux-mêmes", "exactement", "excepté", "extenso", "exterieur", "f", "fais", "faisaient", "faisant", "fait",
"façon", "feront", "fi", "flac", "floc", "font", "g", "gens", "h", "ha", "hein", "hem", "hep", "hi", "ho", "holà", "hop", "hormis", "hors",
"hou", "houp", "hue", "hui", "huit", "huitième", "hum", "hurrah", "hé", "hélas", "i", "il", "ils", "importe", "j", "je", "jusqu", "jusque",
"juste", "k", "l", "la", "laisser", "laquelle", "las", "le", "lequel", "les", "lesquelles", "lesquels", "leur", "leurs", "longtemps",
"lors", "lorsque", "lui", "lui-meme", "lui-même", "là", "lès", "m", "ma", "maint", "maintenant", "mais", "malgre", "malgré", "maximale",
"me", "meme", "memes", "merci", "mes", "mien", "mienne", "miennes", "miens", "mille", "mince", "minimale", "moi", "moi-meme", "moi-même",
"moindres", "moins", "mon", "moyennant", "multiple", "multiples", "même", "mêmes", "n", "na", "naturel", "naturelle", "naturelles", "ne",
"neanmoins", "necessaire", "necessairement", "neuf", "neuvième", "ni", "nombreuses", "nombreux", "non", "nos", "notamment", "notre",
"nous", "nous-mêmes", "nouveau", "nul", "néanmoins", "nôtre", "nôtres", "o", "oh", "ohé", "ollé", "olé", "on", "ont", "onze", "onzième",
"ore", "ou", "ouf", "ouias", "oust", "ouste", "outre", "ouvert", "ouverte", "ouverts", "o|", "où", "p", "paf", "pan", "par", "parce",
"parfois", "parle", "parlent", "parler", "parmi", "parseme", "partant", "particulier", "particulière", "particulièrement", "pas",
"passé", "pendant", "pense", "permet", "personne", "peu", "peut", "peuvent", "peux", "pff", "pfft", "pfut", "pif", "pire", "plein",
"plouf", "plus", "plusieurs", "plutôt", "possessif", "possessifs", "possible", "possibles", "pouah", "pour", "pourquoi", "pourrais",
"pourrait", "pouvait", "prealable", "precisement", "premier", "première", "premièrement", "pres", "probable", "probante",
"procedant", "proche", "près", "psitt", "pu", "puis", "puisque", "pur", "pure", "q", "qu", "quand", "quant", "quant-à-soi", "quanta",
"quarante", "quatorze", "quatre", "quatre-vingt", "quatrième", "quatrièmement", "que", "quel", "quelconque", "quelle", "quelles",
"quelqu'un", "quelque", "quelques", "quels", "qui", "quiconque", "quinze", "quoi", "quoique", "r", "rare", "rarement", "rares",
"relative", "relativement", "remarquable", "rend", "rendre", "restant", "reste", "restent", "restrictif", "retour", "revoici",
"revoilà", "rien", "s", "sa", "sacrebleu", "sait", "sans", "sapristi", "sauf", "se", "sein", "seize", "selon", "semblable", "semblaient",
"semble", "semblent", "sent", "sept", "septième", "sera", "seraient", "serait", "seront", "ses", "seul", "seule", "seulement", "si",
"sien", "sienne", "siennes", "siens", "sinon", "six", "sixième", "soi", "soi-même", "soit", "soixante", "son", "sont", "sous", "souvent",
"specifique", "specifiques", "speculatif", "stop", "strictement", "subtiles", "suffisant", "suffisante", "suffit", "suis", "suit",
"suivant", "suivante", "suivantes", "suivants", "suivre", "superpose", "sur", "surtout", "t", "ta", "tac", "tant", "tardive", "te",
"tel", "telle", "tellement", "telles", "tels", "tenant", "tend", "tenir", "tente", "tes", "tic", "tien", "tienne", "tiennes", "tiens",
"toc", "toi", "toi-même", "ton", "touchant", "toujours", "tous", "tout", "toute", "toutefois", "toutes", "treize", "trente", "tres",
"trois", "troisième", "troisièmement", "trop", "très", "tsoin", "tsouin", "tu", "té", "u", "un", "une", "unes", "uniformement", "unique",
"uniques", "uns", "v", "va", "vais", "vas", "vers", "via", "vif", "vifs", "vingt", "vivat", "vive", "vives", "vlan", "voici", "voilà",
"vont", "vos", "votre", "vous", "vous-mêmes", "vu", "vé", "vôtre", "vôtres", "w", "x", "y", "z", "zut", "à", "â", "ça", "ès", "étaient",
"étais", "était", "étant", "été", "être", "ô"}

    Public StopWordsDutch() As String = {"aan", "achte", "achter", "af", "al", "alle", "alleen", "alles", "als", "ander", "anders", "beetje",
"behalve", "beide", "beiden", "ben", "beneden", "bent", "bij", "bijna", "bijv", "blijkbaar", "blijken", "boven", "bv",
"daar", "daardoor", "daarin", "daarna", "daarom", "daaruit", "dan", "dat", "de", "deden", "deed", "derde", "derhalve", "dertig",
"deze", "dhr", "die", "dit", "doe", "doen", "doet", "door", "drie", "duizend", "echter", "een", "eens", "eerst", "eerste", "eigen",
"eigenlijk", "elk", "elke", "en", "enige", "er", "erg", "ergens", "etc", "etcetera", "even", "geen", "genoeg", "geweest", "haar",
"haarzelf", "had", "hadden", "heb", "hebben", "hebt", "hedden", "heeft", "heel", "hem", "hemzelf", "hen", "het", "hetzelfde",
"hier", "hierin", "hierna", "hierom", "hij", "hijzelf", "hoe", "honderd", "hun", "ieder", "iedere", "iedereen", "iemand", "iets",
"ik", "in", "inderdaad", "intussen", "is", "ja", "je", "jij", "jijzelf", "jou", "jouw", "jullie", "kan", "kon", "konden", "kun",
"kunnen", "kunt", "laatst", "later", "lijken", "lijkt", "maak", "maakt", "maakte", "maakten", "maar", "mag", "maken", "me", "meer",
"meest", "meestal", "men", "met", "mevr", "mij", "mijn", "minder", "miss", "misschien", "missen", "mits", "mocht", "mochten",
"moest", "moesten", "moet", "moeten", "mogen", "mr", "mrs", "mw", "na", "naar", "nam", "namelijk", "nee", "neem", "negen",
"nemen", "nergens", "niemand", "niet", "niets", "niks", "noch", "nochtans", "nog", "nooit", "nu", "nv", "of", "om", "omdat",
"ondanks", "onder", "ondertussen", "ons", "onze", "onzeker", "ooit", "ook", "op", "over", "overal", "overige", "paar", "per",
"recent", "redelijk", "samen", "sinds", "steeds", "te", "tegen", "tegenover", "thans", "tien", "tiende", "tijdens", "tja", "toch",
"toe", "tot", "totdat", "tussen", "twee", "tweede", "u", "uit", "uw", "vaak", "van", "vanaf", "veel", "veertig", "verder",
"verscheidene", "verschillende", "via", "vier", "vierde", "vijf", "vijfde", "vijftig", "volgend", "volgens", "voor", "voordat",
"voorts", "waar", "waarom", "waarschijnlijk", "wanneer", "waren", "was", "wat", "we", "wederom", "weer", "weinig", "wel", "welk",
"welke", "werd", "werden", "werder", "whatever", "wie", "wij", "wijzelf", "wil", "wilden", "willen", "word", "worden", "wordt", "zal",
"ze", "zei", "zeker", "zelf", "zelfde", "zes", "zeven", "zich", "zij", "zijn", "zijzelf", "zo", "zoals", "zodat", "zou", "zouden",
"zulk", "zullen"}

#End Region

    Public SyntaxTerms() As String = ({"SPYDAZ", "ABS", "ACCESS", "ADDITEM", "ADDNEW", "ALIAS", "AND", "ANY", "APP", "APPACTIVATE", "APPEND", "APPENDCHUNK", "ARRANGE", "AS", "ASC", "ATN", "BASE", "BEEP", "BEGINTRANS", "BINARY", "BYVAL", "CALL", "CASE", "CCUR", "CDBL", "CHDIR", "CHDRIVE", "CHR", "CHR$", "CINT", "CIRCLE", "CLEAR", "CLIPBOARD", "CLNG", "CLOSE", "CLS", "COMMAND", "
COMMAND$", "COMMITTRANS", "COMPARE", "CONST", "CONTROL", "CONTROLS", "COS", "CREATEDYNASET", "CSNG", "CSTR", "CURDIR$", "CURRENCY", "CVAR", "CVDATE", "DATA", "DATE", "DATE$", "DATESERIAL", "DATEVALUE", "DAY", "
DEBUG", "DECLARE", "DEFCUR", "CEFDBL", "DEFINT", "DEFLNG", "DEFSNG", "DEFSTR", "DEFVAR", "DELETE", "DIM", "DIR", "DIR$", "DO", "DOEVENTS", "DOUBLE", "DRAG", "DYNASET", "EDIT", "ELSE", "ELSEIF", "END", "ENDDOC", "ENDIF", "
ENVIRON$", "EOF", "EQV", "ERASE", "ERL", "ERR", "ERROR", "ERROR$", "EXECUTESQL", "EXIT", "EXP", "EXPLICIT", "FALSE", "FIELDSIZE", "FILEATTR", "FILECOPY", "FILEDATETIME", "FILELEN", "FIX", "FOR", "FORM", "FORMAT", "
FORMAT$", "FORMS", "FREEFILE", "FUNCTION", "GET", "GETATTR", "GETCHUNK", "GETDATA", "DETFORMAT", "GETTEXT", "GLOBAL", "GOSUB", "GOTO", "HEX", "HEX$", "HIDE", "HOUR", "IF", "IMP", "INPUT", "INPUT$", "INPUTBOX", "INPUTBOX$", "
INSTR", "INT", "INTEGER", "IS", "ISDATE", "ISEMPTY", "ISNULL", "ISNUMERIC", "KILL", "LBOUND", "LCASE", "LCASE$", "LEFT", "LEFT$", "LEN", "LET", "LIB", "LIKE", "LINE", "LINKEXECUTE", "LINKPOKE", "LINKREQUEST", "
LINKSEND", "LOAD", "LOADPICTURE", "LOC", "LOCAL", "LOCK", "LOF", "LOG", "LONG", "LOOP", "LSET", "LTRIM",
    "LTRIM$", "ME", "MID", "MID$", "MINUTE", "MKDIR", "MOD", "MONTH", "MOVE", "MOVEFIRST", "MOVELAST", "MOVENEXT", "MOVEPREVIOUS",
    "MOVERELATIVE", "MSGBOX", "NAME", "NEW", "NEWPAGE", "NEXT", "NEXTBLOCK", "NOT", "NOTHING",
    "NOW", "NULL", "OCT", "OCT$", "ON", "OPEN", "OPENDATABASE", "OPTION", "OR", "OUTPUT", "POINT", "PRESERVE", "PRINT",
    "PRINTER", "PRINTFORM", "PRIVATE", "PSET", "PUT", "PUBLIC", "QBCOLOR", "RANDOM", "RANDOMIZE", "READ", "REDIM", "REFRESH",
    "REGISTERDATABASE", "REM", "REMOVEITEM", "RESET", "RESTORE", "RESUME", "RETURN", "RGB", "RIGHT", "RIGHT$", "RMDIR", "RND",
    "ROLLBACK", "RSET", "RTRIM", "RTRIM$", "SAVEPICTURE", "SCALE", "SECOND", "SEEK", "SELECT", "SENDKEYS", "SET", "SETATTR",
    "SETDATA", "SETFOCUS", "SETTEXT", "SGN", "SHARED",
    "SHELL", "SHOW", "SIN", "SINGLE", "SPACE", "SPACE$", "SPC", "SQR",
    "STATIC", "STEP", "STOP", "STR", "STR$", "STRCOMP", "STRING", "STRING$", "SUB",
    "SYSTEM", "TAB", "TAN", "TEXT", "TEXTHEIGHT", "TEXTWIDTH", "THEN", "TIME", "TIME$",
    "TIMER", "TIMESERIAL", "TIMEVALUE", "TO", "TRIM",
    "TRIM$", "TRUE", "TYPE", "TYPEOF", "UBOUND", "UCASE", "UCASE$", "UNLOAD",
    "UNLOCK", "UNTIL", "UPDATE", "USING", "VAL", "VARIANT", "VARTYPE", "WEEKDAY", "WEND", "WHILE", "WIDTH",
    "WRITE", "XOR", "YEAR", "ZORDER", "ADDHANDLER", "ADDRESSOF", "ALIAS", "AND", "ANDALSO", "AS", "BYREF",
    "BOOLEAN", "BYTE", "BYVAL", "CALL", "CASE", "CATCH", "CBOOL", "CBYTE", "CCHAR", "CDATE",
    "CDEC", "CDBL", "CHAR", "CINT", "CLASS", "CLNG", "COBJ", "CONST", "CONTINUE", "CSBYTE",
    "CSHORT", "CSNG", "CSTR", "CTYPE", "CUINT", "CULNG", "CUSHORT", "DATE", "DECIMAL", "DECLARE",
    "DEFAULT", "DELEGATE", "DIM", "DIRECTCAST", "DOUBLE", "DO", "EACH", "ELSE", "ELSEIF", "END",
    "ENDIF", "ENUM", "ERASE", "ERROR", "EVENT", "EXIT", "FALSE", "FINALLY", "FOR", "FRIEND",
    "FUNCTION", "GET", "GETTYPE", "GLOBAL", "GOSUB", "GOTO", "HANDLES", "IF", "IMPLEMENTS",
    "IMPORTS", "IN", "INHERITS", "INTEGER", "INTERFACE", "IS", "ISNOT", "LET", "LIB", "LIKE",
    "LONG", "LOOP", "ME", "MOD", "MODULE", "MUSTINHERIT", "MUSTOVERRIDE", "MYBASE", "MYCLASS",
    "NAMESPACE", "NARROWING", "NEW", "NEXT", "NOT", "NOTHING", "NOTINHERITABLE", "NOTOVERRIDABLE",
    "OBJECT", "ON", "OF", "OPERATOR", "OPTION", "OPTIONAL", "OR", "ORELSE", "OVERLOADS",
    "OVERRIDABLE", "OVERRIDES", "PARAMARRAY", "PARTIAL", "PRIVATE", "PROPERTY", "PROTECTED",
    "PUBLIC", "RAISEEVENT", "READONLY", "REDIM", "REM", "REMOVEHANDLER", "RESUME", "RETURN",
    "SBYTE", "SELECT", "SET", "SHADOWS", "SHARED", "SHORT", "SINGLE", "STATIC", "STEP", "STOP",
    "STRING", "STRUCTURE", "SUB", "SYNCLOCK", "THEN", "THROW", "TO", "TRUE", "TRY", "TRYCAST",
    "TYPEOF", "WEND", "VARIANT", "UINTEGER", "ULONG", "USHORT", "USING", "WHEN", "WHILE", "WIDENING",
    "WITH", "WITHEVENTS", "WRITEONLY",
    "XOR", "#CONST", "#ELSE", "#ELSEIF", "#END", "#IF"})

    ''' <summary>
    ''' Highlights selection of text from index to length
    ''' </summary>
    ''' <param name="rtb">RichTextBox</param>
    ''' <param name="index">Start Index</param>
    ''' <param name="length">Charchcters width</param>
    ''' <param name="color">Color</param>
    <Runtime.CompilerServices.Extension()>
    Private Sub HighlightIndex(rtb As RichTextBox, index As Integer, length As Integer, color As Color)
        Dim selectionStartSave As Integer = rtb.SelectionStart 'to return this back to its original position
        rtb.SelectionStart = index
        rtb.SelectionLength = length
        rtb.SelectionColor = color
        rtb.SelectionLength = 0
        rtb.SelectionStart = selectionStartSave
        rtb.SelectionColor = rtb.ForeColor 'return back to the original color
    End Sub

    ''' <summary>
    ''' Uses internal syntax to Recolor VB.NEt Syntax Words
    ''' </summary>
    ''' <param name="sender">RichTextBox</param>
    <Runtime.CompilerServices.Extension()>
    Public Sub HighlightInternalSyntax(ByRef sender As RichTextBox)

        For Each Word As String In SyntaxTerms
            HighlightWord(sender, Word)
            HighlightWord(sender, LCase(Word))
            HighlightWord(sender, UCase(Word))
            HighlightWord(sender, StrConv(Word, vbProperCase))
        Next

    End Sub

    ''' <summary>
    ''' Highlights Specific Word in The textbox
    ''' </summary>
    ''' <param name="sender">RichTextBox</param>
    ''' <param name="word">Word to be foud</param>
    <Runtime.CompilerServices.Extension()>
    Public Sub HighlightWord(ByRef sender As RichTextBox, ByRef word As String)
        'Mark Cursor Point
        Dim selectionStartSave As Integer = sender.SelectionStart 'to return this back to its original position
        Dim index As Integer = 0
        While index < sender.Text.LastIndexOf(word)
            'find
            sender.Find(word, index, sender.TextLength, RichTextBoxFinds.WholeWord)
            'select and color
            sender.SelectionColor = Color.Blue
            index = sender.Text.IndexOf(word, index) + 1
        End While
        'Return Cursor Position
        sender.SelectionStart = selectionStartSave

    End Sub

    ''' <summary>
    ''' Adds a child form to the sent form
    ''' </summary>
    ''' <param name="Sender"></param>
    ''' <param name="form"></param>
    <Runtime.CompilerServices.Extension()>
    Public Sub AddChildForm(ByRef Sender As System.Windows.Forms.Form, ByRef form As System.Windows.Forms.Form)
        Dim MdiChild = form
        MdiChild.MdiParent = Sender
        MdiChild.Show()

    End Sub

End Module

Namespace AI_SDK_EXTENSIONS

    Namespace MathsExt

        ''' <summary>
        ''' Caesar cipher, both encoding And decoding.
        ''' The key Is an Integer from 1 To 25.
        ''' This cipher rotates (either towards left Or right) the letters Of the alphabet (A To Z).
        ''' The encoding replaces Each letter With the 1St To 25th Next letter In the alphabet (wrapping Z To A).
        ''' So key 2 encrypts "HI" To "JK", but key 20 encrypts "HI" To "BC".
        ''' </summary>
        Public Class Caesarcipher

            Public Function Encrypt(ch As Char, code As Integer) As Char
                If Not Char.IsLetter(ch) Then
                    Return ch
                End If

                Dim offset = AscW(If(Char.IsUpper(ch), "A"c, "a"c))
                Dim test = (AscW(ch) + code - offset) Mod 26 + offset
                Return ChrW(test)
            End Function

            Public Function Encrypt(input As String, code As Integer) As String
                Return New String(input.Select(Function(ch) Encrypt(ch, code)).ToArray())
            End Function

            Public Function Decrypt(input As String, code As Integer) As String
                Return Encrypt(input, 26 - code)
            End Function

        End Class

    End Namespace

End Namespace