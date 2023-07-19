Imports System.Web.Script.Serialization
Imports System.Windows.Forms
Imports SpydazWebAI.SystemExtensions.AI_SDK_EXTENSIONS.Strings
Imports SpydazWebAI.SystemExtensions.Datatypes.Syllogism

Namespace Datatypes

    Namespace Intents

        ''' <summary>
        ''' Data Required to full-fill action
        ''' </summary>
        Public Structure Slot

#Region "Fields"

            Public ConfirmationPhrase As String
            Public Data As String
            Public EntityType As String
            Public Filled As Boolean

#End Region

        End Structure

    End Namespace

    ''' <summary>
    ''' QuestionTypes are
    '''place           'Where is X - Returns Answer to Question (Where is becomes the subject) X is at location of Subject
    '''time            'When was - Returns Answer to Question (when was X) X was Subject
    '''person          'Who is - Returns Answer to Question - Object becomes the search term for answer replaced by Subject
    '''reason          'Why is - Returns Answer to Question
    '''description     'What is - Returns Answer to Question
    '''instruction     'How do - Returns Answer to Question
    '''Choice          'Which X, either X or Y - Returns Selected choice from input (Which becomes the Subject X and Y are both the objects)
    '''Possession      'Whose X is? - returns Question about Ownership (Whose is the subject X is the object) X belongs to object
    ''' </summary>
    Public Enum _QuestionType
        place           'Where is - Returns Answer to Question
        time            'When was - Returns Answer to Question
        person          'Who is - Returns Answer to Question
        reason          'Why is - Returns Answer to Question
        description     'What is - Returns Answer to Question
        instruction     'How do - Returns Answer to Question
        Choice          'Which X, either X or Y - Returns Selected choice from input
        Possession      'Whose X is? - returns Question about Ownwership
    End Enum

    ''' <summary>
    ''' These are the standard concept types / The future types will be a list collected fromt the
    ''' database the database will only have dynamic predicate types
    ''' </summary>
    Public Enum ConceptType

        'Thing
        CapeableOf 'Something that A can typically do is B.

        ConceptuallyRelatedTo 'The most general relation. There is some positive relationship between A and B, but ConceptNet can't determine what that relationship is based on the data. This was called "ConceptuallyRelatedTo" in
        CapableOfReceivingAction ' Subject A can be eaten, can be washed
        DefinedAs 'A and B overlap considerably in meaning, and B is a more explanatory version of A. (This is similar to TranslationOf, but within one language.)
        DesireOf 'A is a conscious entity that typically wants B. Many assertions of this type use the appropriate language's word for "person" as A
        DesirousEffectOf 'Jump / is to land
        EffectOf 'Singing / is to Entertain
        IsA 'A is a subtype or a specific instance of B; every A is a B. (We do not make the type-token distinction, because people don't usually make that distinction.) This is the hyponym relation in WordNet.
        MadeOf 'A has composition property's of
        MotivationOf 'Someone does A because they want result B; A is a step toward accomplishing the goal B.
        PartOf 'A is a part of B. This is the part meronym relation in WordNet.
        PropertyOf 'A has B as a property; A can be described as B.
        UsedFor 'A is used for B; the purpose of A is B.

        'Event
        FirstSubEventOf 'A is an event that begins with subevent B.

        LastSubEventOf 'A is an event that concludes with subevent B.
        PreRequisiteOf 'In order for A to happen, B needs to happen; B is a dependency of A.
        SubeventOf 'A and B are events, and B happens as a subevent of A.

        'Place
        LocationOf 'A is a typical location for B, or A is the inherent location of B. Some instances of this would be considered meronyms in WordNet.

    End Enum

    Public Enum CorrelationResult
        Positive = 1
        PositiveHigh = 0.9
        PositiveLow = 0.5
        None = 0
        NegativeLow = -0.5
        NegativeHigh = -0.9
        Negative = -1
    End Enum

    ''' <summary>
    ''' Neural network Layer types
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum LayerType
        Input
        Hidden
        Output
    End Enum

    ''' <summary>
    '''can         '– ability, permission, possibility, request
    '''could       '– ability, permission, possibility, request, suggestion
    '''may         '– permission, probability, request
    '''might       '– possibility, probability, suggestion
    '''must        '– deduction, necessity, obligation, prohibition
    '''shall       '– decision, future, offer, question, suggestion
    '''should      '– advice, necessity, prediction, recommendation
    '''will        '– decision, future, intention, offer, prediction, promise, suggestion
    '''would       '– conditional, habit, invitation, permission, preference, request, question, suggestion
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum ModalMeanings
        Ability
        Request
        Permission
        Possibility
        Obligation
        Prohibition
        Lack_Of_necessity
        Advice
        probability
        Conditional
        Invitation
        Suggestion
        deduction
        Question
        necsessity
        prediction
        offer
        promise
        decision
        Prefference
        intention
        Instruction
    End Enum

    Public Enum ModalType
        can         '– ability, permission, possibility, request
        could       '– ability, permission, possibility, request, suggestion
        may         '– permission, probability, request
        might       '– possibility, probability, suggestion
        must        '– deduction, necessity, obligation, prohibition
        shall       '– decision, future, offer, question, suggestion
        should      '– advice, necessity, prediction, recommendation
        will        '– decision, future, intention, offer, prediction, promise, suggestion
        would       '– conditional, habit, invitation, permission, preference, request, question, suggestion
        Unknown
    End Enum

    ''' <summary>
    ''' These are the different object types available these object types are the highest level
    ''' objects of which most things sub objects or combined types
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum ObjectType
        _Location
        _Event
        _Thing
        _Animal
        _Plant
        _Person
    End Enum

    ''' <summary>
    ''' Used to indicate how the data within should be assessed
    ''' </summary>
    Public Enum ResponseType

        ''' <summary>
        ''' An Object was Detected ; Perhaps nyms were missing;
        ''' A question maybe prepared to be returned to the user to seek knowledge about the subject within.
        ''' It maybe the data within is just to be used as pointers?, So that the Query that can be sent to the user be direct or  random
        ''' </summary>
        SeekKnowledge

        ''' <summary>
        ''' User has asked a Question
        ''' </summary>
        UserQuestion

        ''' <summary>
        ''' User has made a statment
        ''' </summary>
        UserStatement

        ''' <summary>
        ''' Input was undetermined or Null
        ''' </summary>
        Null

    End Enum

    ''' <summary>
    ''' Category of sentence
    ''' </summary>
    Public Enum SentenceCatagory
        SimpleSentence
        CompoundSentence
        ComplexSentence
    End Enum

    ''' <summary>
    ''' Used to denote tense
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum tense
        Past = 0
        Present = 1
        Future = 2
        Unknown = 3
    End Enum

    ''' <summary>
    ''' These are the options of transfer functions available to the network
    ''' This is used to select which function to be used:
    ''' The derivative function can also be selected using this as a marker
    ''' </summary>
    Public Enum TransferFunctionType
        none
        sigmoid
        HyperbolTangent
        BinaryThreshold
        RectifiedLinear
        Logistic
        StochasticBinary
        Gaussian
        Signum
    End Enum

    ''' <summary>
    ''' Used to hold Answers to Questions posed ;
    ''' There may be many answers a random answer can be returned
    ''' </summary>
    Public Structure Answer

#Region "Fields"

        ''' <summary>
        ''' Placeholder for current Question
        ''' </summary>
        Public CurrentQuestion As AssociatedQuestion

        ''' <summary>
        ''' A list of known knowledge for Question(Nym Specific or All Answers)
        ''' </summary>
        Public HeldAnswers As List(Of Cept)

        ''' <summary>
        ''' Current object
        ''' </summary>
        Public ObjectStr As String

#End Region

#Region "Methods"

        ''' <summary>
        ''' Returns Random Answer from
        ''' </summary>
        ''' <returns></returns>
        Public Function GetRandomAnswer() As String
            Dim rnd = New Random()
            If HeldAnswers.Count > 0 Then

                Return HeldAnswers(rnd.Next(0, HeldAnswers.Count)).DefinedStr
            Else
                Return Nothing
            End If
        End Function

        Public Function ToJson() As String
            Dim Converter As New JavaScriptSerializer
            Return Converter.Serialize(Me)
        End Function

#End Region

    End Structure

    ''' <summary>
    ''' Used to hold  Associated Questions;
    ''' these are used to detect questions as well as store questions related to learning information related to a particular nym
    ''' these may also be used to ask questions to gain information from the user.
    ''' Questions are in the format of "What IS A A#","IS A"
    ''' </summary>
    Public Structure AssociatedQuestion

#Region "Fields"

        Public NymStr As String
        Public Question As String

#End Region

#Region "Methods"

        ''' <summary>
        ''' Checks if Nym is in set
        ''' </summary>
        ''' <param name="Questions"></param>
        ''' <param name="Nymstr"></param>
        ''' <returns></returns>
        Public Shared Function CheckIfNymExists(ByRef Questions As List(Of AssociatedQuestion), ByRef Nymstr As String) As Boolean
            CheckIfNymExists = False
            For Each item In Questions
                If item.NymStr = Nymstr Then
                    Return True
                End If
            Next
        End Function

        Public Shared Function GetQuestionByNym(ByRef Questions As List(Of AssociatedQuestion), ByRef NymStr As String) As List(Of AssociatedQuestion)
            Dim lst As New List(Of AssociatedQuestion)
            If Questions.Count > 0 Then
                For Each item In Questions
                    If item.NymStr = NymStr Then
                        lst.Add(item)
                    Else
                    End If
                Next
                Return lst
            Else
                Return Nothing
            End If
        End Function

        Public Shared Function UpdateObjectStr(ByRef Question As AssociatedQuestion, ByRef ObjectStr As String) As AssociatedQuestion
            Question.Question.Replace("A#", ObjectStr)
            Return Question
        End Function

        Public Function FilterQuestionsByNym(ByRef Questions As List(Of AssociatedQuestion), ByRef NymStr As String) As List(Of AssociatedQuestion)
            Dim Lst As New List(Of AssociatedQuestion)
            For Each item In Questions
                If item.NymStr = NymStr Then
                    Lst.Add(item)
                End If
            Next

            If Lst.Count > 0 Then
                Return Lst
            Else
                Return Nothing
            End If
        End Function

        Public Function GetRandomQuestion(ByRef Questions As List(Of AssociatedQuestion)) As AssociatedQuestion
            Dim rnd = New Random()
            If Questions.Count > 0 Then

                Return Questions(rnd.Next(0, Questions.Count))
            Else
                Return Nothing
            End If
        End Function

        Public Function InsertSubjectIntoQuestions(ByRef Questions As List(Of AssociatedQuestion), ByRef Subject As String) As List(Of AssociatedQuestion)

            For Each item In Questions
                item.Question.Replace("A#", Subject)
            Next
            Return Questions
        End Function

        Public Function InsertWildCardIntoQuestions(ByRef Questions As List(Of AssociatedQuestion)) As List(Of AssociatedQuestion)

            For Each item In Questions
                item.Question.Replace("A#", "*")
            Next
            Return Questions
        End Function

        Public Function ToJson() As String
            Dim Converter As New JavaScriptSerializer
            Return Converter.Serialize(Me)
        End Function

        Public Function UpdateObjectStr(ByRef ObjectStr As String) As AssociatedQuestion
            Me.Question.Replace("A#", ObjectStr)
            Return Me
        End Function

#End Region

    End Structure

    ''' <summary>
    ''' Used to retrieve or hold on to concept records
    ''' Object / Nym / Defined
    ''' </summary>
    Public Structure Cept

#Region "Fields"

        Public DefinedStr As String

        ''' <summary>
        ''' Used to hold the semantic pattern which detected it
        ''' </summary>
        Public DetectedSearchPattern As SemanticPattern

        ''' <summary>
        ''' Predicate (item between Subjects)
        ''' </summary>
        Public NymStr As String

        Public ObjectStr As String

        ''' <summary>
        ''' Holds Original Sentences / detected sentence
        ''' </summary>
        Public OriginalSentence As String

#End Region

#Region "Methods"

        ''' <summary>
        ''' Returns all Records Cepts from Given Table
        ''' </summary>
        ''' <param name="TableName">Table Name</param>
        ''' <param name="iConnectionStr">Custome Connection Str</param>
        ''' <returns></returns>
        Public Shared Function GetDBCepts(ByRef iConnectionStr As String, ByRef TableName As String) As List(Of Cept)
            Dim Lst As New List(Of Cept)
            Dim SQL As String = "SELECT * FROM " & TableName

            Using conn = New System.Data.OleDb.OleDbConnection(iConnectionStr)
                Using cmd = New System.Data.OleDb.OleDbCommand(SQL, conn)
                    conn.Open()
                    Try
                        Dim dr = cmd.ExecuteReader()
                        While dr.Read()
                            Dim NewKnowledge As New Cept With {
                                .ObjectStr = dr("Object").ToString(),
                                .NymStr = dr("Nym").ToString(),
                                .DefinedStr = dr("Defined").ToString()
                            }
                            Lst.Add(NewKnowledge)
                        End While
                    Catch e As Exception
                        ' Do some logging or something.
                        MessageBox.Show("There was an error accessing your data. GetDBCepts: " & e.ToString())
                    End Try
                End Using
            End Using

            Return Lst
        End Function

        ''' <summary>
        ''' Get all Cepts table by nym
        ''' </summary>
        ''' <param name="iConnectionStr"></param>
        ''' <param name="TableName"></param>
        ''' <param name="NymStr"></param>
        ''' <returns></returns>
        Public Shared Function GetDBCeptsByNym(ByRef iConnectionStr As String, ByRef TableName As String, ByRef NymStr As String) As List(Of Cept)
            Dim Lst As New List(Of Cept)
            Dim SQL As String = "SELECT * FROM " & TableName & " Where Nym='" & NymStr & "'"
            Using conn = New System.Data.OleDb.OleDbConnection(iConnectionStr)
                Using cmd = New System.Data.OleDb.OleDbCommand(SQL, conn)
                    conn.Open()
                    Try
                        Dim dr = cmd.ExecuteReader()
                        While dr.Read()
                            Dim NewKnowledge As New Cept With {
                                .ObjectStr = dr("Object").ToString(),
                                .NymStr = dr("Nym").ToString(),
                                .DefinedStr = dr("Defined").ToString()
                            }
                            Lst.Add(NewKnowledge)
                        End While
                    Catch e As Exception
                        ' Do some logging or something.
                        MessageBox.Show("There was an error accessing your data. GetDBCeptsByNym: " & e.ToString())
                    End Try
                End Using
            End Using
            Return Lst
        End Function

        ''' <summary>
        ''' Get Cepts from table by object
        ''' </summary>
        ''' <param name="iConnectionStr"></param>
        ''' <param name="TableName"></param>
        ''' <param name="ObjectStr"></param>
        ''' <returns></returns>
        Public Shared Function GetDBCeptsByObject(ByRef iConnectionStr As String, ByRef TableName As String, ByRef ObjectStr As String) As List(Of Cept)
            Dim Lst As New List(Of Cept)
            Dim SQL As String = "SELECT * FROM " & TableName & " Where Object='" & ObjectStr & "'"

            Using conn = New System.Data.OleDb.OleDbConnection(iConnectionStr)
                Using cmd = New System.Data.OleDb.OleDbCommand(SQL, conn)
                    conn.Open()
                    Try
                        Dim dr = cmd.ExecuteReader()
                        While dr.Read()
                            Dim NewKnowledge As New Cept With {
                                .ObjectStr = dr("Object").ToString(),
                                .NymStr = dr("Nym").ToString(),
                                .DefinedStr = dr("Defined").ToString()
                            }
                            Lst.Add(NewKnowledge)
                        End While
                    Catch e As Exception
                        ' Do some logging or something.
                        MessageBox.Show("There was an error accessing your data. GetDBCeptsByObject: " & e.ToString())
                    End Try
                End Using
            End Using

            Return Lst
        End Function

        ''' <summary>
        ''' Get Cepts from table by Query
        ''' </summary>
        ''' <param name="iConnectionStr"></param>
        ''' <param name="Query"></param>
        ''' <returns></returns>
        Public Shared Function GetDBCeptsByQuery(ByRef iConnectionStr As String, ByRef Query As String) As List(Of Cept)
            Dim Lst As New List(Of Cept)
            Dim SQL As String = Query

            Using conn = New System.Data.OleDb.OleDbConnection(iConnectionStr)
                Using cmd = New System.Data.OleDb.OleDbCommand(SQL, conn)
                    conn.Open()
                    Try
                        Dim dr = cmd.ExecuteReader()
                        While dr.Read()
                            Dim NewKnowledge As New Cept
                            NewKnowledge.ObjectStr = dr("Object").ToString()
                            NewKnowledge.NymStr = dr("Nym").ToString()
                            NewKnowledge.DefinedStr = dr("Defined").ToString()
                            Lst.Add(NewKnowledge)
                        End While
                    Catch e As Exception
                        ' Do some logging or something.
                        MessageBox.Show("There was an error accessing your data. GetDBCeptsByQuery: " & e.ToString())
                    End Try
                End Using
            End Using

            Return Lst
        End Function

        ''' <summary>
        ''' Adds new Nym into table
        ''' </summary>
        ''' <param name="NymStr">Must not Exist</param>
        Public Function AddNym(ByRef iConnectionStr As String, ByRef Tablename As String, ByRef NymStr As String) As Boolean
            AddNym = False
            If NymStr IsNot Nothing Then

                Dim SQL = "INSERT INTO " & Tablename & " Nym VALUES ('" & NymStr & "')"
                Using conn = New System.Data.OleDb.OleDbConnection(iConnectionStr)

                    Using cmd = New System.Data.OleDb.OleDbCommand(SQL, conn)
                        conn.Open()

                        Try
                            cmd.ExecuteNonQuery()
                            AddNym = True
                        Catch ex As Exception
                            MessageBox.Show("There was an error accessing your data. AddNym: " & ex.ToString())
                        End Try
                    End Using
                End Using
            Else
            End If

        End Function

        ''' <summary>
        ''' Checks given list to see if nym exists in set
        ''' </summary>
        ''' <param name="Knowledge"></param>
        ''' <param name="NymStr"></param>
        ''' <returns></returns>
        Public Function CheckIfNymExists(ByRef Knowledge As List(Of Cept), ByRef NymStr As String) As Boolean
            Dim lst As New List(Of Cept)
            For Each item In Knowledge
                If item.NymStr = NymStr Then
                    lst.Add(item)
                End If
            Next
            If lst.Count > 0 Then
                Return True
            Else
                Return False

            End If
        End Function

        Public Function FilterCepts(ByRef Cepts As List(Of Cept), ByRef ObjectStr As String, ByRef NymStr As String) As List(Of Cept)
            Dim ObjLst As List(Of Cept) = FilterCeptsByNym(FilterCeptsByObject(Cepts, ObjectStr), NymStr)
            If ObjLst.Count > 0 Then
                Return ObjLst
            Else
                Return Nothing
            End If
        End Function

        Public Function FilterCeptsByNym(ByRef Cepts As List(Of Cept), ByRef NymStr As String) As List(Of Cept)
            Dim Lst As New List(Of Cept)
            For Each item In Cepts
                If item.NymStr = NymStr Then
                    Lst.Add(item)
                End If
            Next
            If Lst.Count > 0 Then
                Return Lst
            Else
                Return Nothing

            End If
        End Function

        Public Function FilterCeptsByObject(ByRef Cepts As List(Of Cept), ByRef ObjectStr As String) As List(Of Cept)
            Dim Lst As New List(Of Cept)
            For Each item In Cepts
                If item.ObjectStr = ObjectStr Then
                    Lst.Add(item)
                End If
            Next
            If Lst.Count > 0 Then
                Return Lst
            Else
                Return Nothing

            End If
        End Function

        ''' <summary>
        ''' Return set from list by nym
        ''' This is used to segment cept Lists
        ''' </summary>
        ''' <param name="Knowledge"></param>
        ''' <param name="NymStr"></param>
        ''' <returns></returns>
        Public Function GetByNym(ByRef Knowledge As List(Of Cept), ByRef NymStr As String) As List(Of Cept)
            Dim lst As New List(Of Cept)
            For Each item In Knowledge
                If item.NymStr = NymStr Then
                    lst.Add(item)
                End If
            Next
            Return lst
        End Function

        ''' <summary>
        ''' Returns a random Cept from List
        ''' </summary>
        ''' <param name="Cepts"></param>
        ''' <returns></returns>
        Public Function GetRandomCept(ByRef Cepts As List(Of Cept)) As Cept
            Dim rnd = New Random()
            If Cepts.Count > 0 Then
                Return Cepts(rnd.Next(0, Cepts.Count))
            Else
                Return Nothing
            End If
        End Function

        Public Function ToJson() As String
            Dim Converter As New JavaScriptSerializer
            Return Converter.Serialize(Me)
        End Function

#End Region

    End Structure

    ''' <summary>
    ''' Concept Structure Detected : With Learning Pattern which detected detected the concept
    ''' Type in the text
    ''' </summary>
    Public Structure DiscoveredConcept

#Region "Fields"

        ''' <summary>
        ''' Detected Learning Pattern (Sentence as Predicate) Enables for Predicate is not Required
        ''' </summary>
        Public DetectedPat As String

        Public DetectedResponseType As ResponseType

        ''' <summary>
        ''' Used to hold the semantic pattern which detected it
        ''' </summary>
        Public DetectedSearchPattern As SemanticPattern

        ''' <summary>
        ''' Nym
        ''' </summary>
        Public NymStr As String

        ''' <summary>
        ''' USed to hold the original sentence state
        ''' </summary>
        Public OriginalSentence As String

        ''' <summary>
        ''' Used to hold Predicate in sentence
        ''' </summary>
        Public Predicate As String

        ''' <summary>
        ''' Subject A detected
        ''' </summary>
        Public SubjectA As String

        ''' <summary>
        ''' Subject B detected
        ''' </summary>
        Public SubjectB As String

#End Region

#Region "Properties"

        ''' <summary>
        ''' Type of response generated by Anaylasis
        ''' Unables for understanding of the results gained
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property ResponseTypeStr As String
            Get
                Select Case DetectedResponseType
                    Case ResponseType.SeekKnowledge
                        Return "SeekKnowledge"
                    Case ResponseType.UserQuestion
                        Return "UserQuestion"
                    Case ResponseType.UserStatement
                        Return "UserStatement"
                    Case ResponseType.Null
                        Return "Null"
                End Select
                Return Nothing
            End Get
        End Property

#End Region

#Region "Methods"

        ''' <summary>
        ''' Checks given list to see if nym exists in set
        ''' </summary>
        ''' <param name="Knowledge"></param>
        ''' <param name="NymStr"></param>
        ''' <returns></returns>
        Public Function CheckIfNymExists(ByRef Knowledge As List(Of DiscoveredConcept), ByRef NymStr As String) As Boolean
            Dim lst As New List(Of DiscoveredConcept)
            For Each item In Knowledge
                If item.NymStr = NymStr Then
                    lst.Add(item)
                End If
            Next
            If lst.Count > 0 Then
                Return True
            Else
                Return False

            End If
        End Function

        ''' <summary>
        ''' Filters list by Object and Nym
        ''' </summary>
        ''' <param name="Cepts">List to be searched</param>
        ''' <param name="ObjectStr">Object </param>
        ''' <param name="NymStr">Nym</param>
        ''' <returns></returns>
        Public Function FilterCepts(ByRef Cepts As List(Of DiscoveredConcept), ByRef ObjectStr As String, ByRef NymStr As String) As List(Of DiscoveredConcept)
            Dim ObjLst As List(Of DiscoveredConcept) = FilterCeptsByNym(FilterCeptsByObject(Cepts, ObjectStr), NymStr)
            If ObjLst.Count > 0 Then
                Return ObjLst
            Else
                Return Nothing
            End If
        End Function

        ''' <summary>
        ''' Filter List by Nym
        ''' </summary>
        ''' <param name="Cepts">List to be filtered</param>
        ''' <param name="NymStr">Nym</param>
        ''' <returns></returns>
        Public Function FilterCeptsByNym(ByRef Cepts As List(Of DiscoveredConcept), ByRef NymStr As String) As List(Of DiscoveredConcept)
            Dim Lst As New List(Of DiscoveredConcept)
            For Each item In Cepts
                If item.NymStr = NymStr Then
                    Lst.Add(item)
                End If
            Next
            If Lst.Count > 0 Then
                Return Lst
            Else
                Return Nothing

            End If
        End Function

        ''' <summary>
        ''' Filter list by object
        ''' </summary>
        ''' <param name="Cepts">List to be searched</param>
        ''' <param name="ObjectStr">object</param>
        ''' <returns></returns>
        Public Function FilterCeptsByObject(ByRef Cepts As List(Of DiscoveredConcept), ByRef ObjectStr As String) As List(Of DiscoveredConcept)
            Dim Lst As New List(Of DiscoveredConcept)
            For Each item In Cepts
                If item.SubjectA = ObjectStr Then
                    Lst.Add(item)
                End If
            Next
            If Lst.Count > 0 Then
                Return Lst
            Else
                Return Nothing

            End If
        End Function

        ''' <summary>
        ''' Return set from list by nym
        ''' This is used to segment cept Lists
        ''' </summary>
        ''' <param name="Knowledge"></param>
        ''' <param name="NymStr"></param>
        ''' <returns></returns>
        Public Function GetByNym(ByRef Knowledge As List(Of DiscoveredConcept), ByRef NymStr As String) As List(Of DiscoveredConcept)
            Dim lst As New List(Of DiscoveredConcept)
            For Each item In Knowledge
                If item.NymStr = NymStr Then
                    lst.Add(item)
                End If
            Next
            Return lst
        End Function

        ''' <summary>
        ''' Returns a random Cept from List
        ''' </summary>
        ''' <param name="Cepts"></param>
        ''' <returns></returns>
        Public Function GetRandomCept(ByRef Cepts As List(Of DiscoveredConcept)) As DiscoveredConcept
            Dim rnd = New Random()
            If Cepts.Count > 0 Then
                Return Cepts(rnd.Next(0, Cepts.Count))
            Else
                Return Nothing
            End If
        End Function

        Public Function ToCept() As Cept
            Dim NewCept As New Cept With {
                .DefinedStr = Me.SubjectB,
                .NymStr = Me.NymStr,
                .ObjectStr = Me.SubjectA
            }
            Return NewCept
        End Function

        Public Function ToJson() As String
            Dim Converter As New JavaScriptSerializer
            Return Converter.Serialize(Me)
        End Function

#End Region

    End Structure

    ''' <summary>
    ''' Extract Entities from text using Named Entity Recognition (NER).
    ''' NER labels sequences of words in a text which are the names of things,
    ''' such as person and company names.
    ''' By creating lists On specific topics such As a list Of names / Locations , Organizations etc.
    ''' These can be used To identify these words In the text, As the entity type.
    ''' The entity's extracted can be extracted or the sentence containing the entity.
    ''' this structure is used to hold extracted Entity's and their associated sentences, and keywords
    ''' </summary>
    Public Structure DiscoveredEntity

#Region "Fields"

        ''' <summary>
        ''' Discovered Entity
        ''' </summary>
        Dim DiscoveredKeyWord As String

        ''' <summary>
        ''' Associated Sentence
        ''' </summary>
        Dim DiscoveredSentence As String

        ''' <summary>
        ''' Theme of Entitys in List / name of associated list
        ''' </summary>
        Public EntityList As String

#End Region

#Region "Methods"

        ''' <summary>
        ''' Outputs Structure to Jason(JavaScriptSerializer)
        ''' </summary>
        ''' <returns></returns>
        Public Function ToJson() As String
            Dim Converter As New JavaScriptSerializer
            Return Converter.Serialize(Me)
        End Function

#End Region

    End Structure

    ''' <summary>
    ''' List of Syllogisms held in database
    ''' </summary>
    Public Structure MasterList
        Public ConnectionStr As String

        ''' <summary>
        ''' Gets all of the premises stored in the table
        ''' </summary>
        Private Shared ReadOnly mMasterList As List(Of Syllogism) = GetDBPremises()

        ''' <summary>
        ''' List of The Stored Premises
        ''' </summary>
        ''' <returns></returns>
        Public Shared ReadOnly Property MasterList As List(Of Syllogism)
            Get
                Return mMasterList
            End Get
        End Property

        ''' <summary>
        ''' Checks the syllogism presented to see if the Consequent and Antecedent Match
        ''' Although copulas may not match
        ''' </summary>
        ''' <param name="Andecedant"></param>
        ''' <param name="Consequent"></param>
        ''' <param name="Premises"></param>
        ''' <returns></returns>
        Public Shared Function CheckSyllogismExists(ByRef Andecedant As String, ByRef Consequent As String,
                            ByVal Premises As List(Of Syllogism)) As Boolean
            CheckSyllogismExists = False
            For Each Premise As Syllogism In Premises
                If Premise.Consequent = Consequent And
                Premise.Antecedant = Andecedant = True Then

                    Return True
                Else
                    CheckSyllogismExists = False
                End If
            Next
            Return CheckSyllogismExists
        End Function

        Public Shared Function CheckSyllogismExists(ByRef Premise As Syllogism,
                                                                                            ByVal Premises As List(Of Syllogism)) As Boolean
            CheckSyllogismExists = False
            For Each item As Syllogism In Premises
                If Compare(item, Premise) = True Then
                    CheckSyllogismExists = True
                End If
            Next
            Return CheckSyllogismExists
        End Function

        ''' <summary>
        ''' Checks if term exists in Premises
        ''' </summary>
        ''' <param name="term"></param>
        ''' <param name="searchlist"></param>
        ''' <returns></returns>
        Public Shared Function CheckTermExists(ByRef term As String, ByVal searchlist As List(Of Syllogism)) As Boolean
            CheckTermExists = False
            If ContainsAndecedant(term, searchlist) = True Or
                    ContainsConsequent(term, searchlist) = True Then
                Return True
            End If
        End Function

        ''' <summary>
        ''' Searches Premise list of syllogisms to see if antecedent is contained within as an antecedent
        ''' </summary>
        ''' <param name="SearchAndecedant"></param>
        ''' <param name="Relations"></param>
        ''' <returns></returns>
        Public Shared Function ContainsAndecedant(ByRef SearchAndecedant As String,
                                   ByVal Relations As List(Of Syllogism)) As Boolean
            Dim found As Boolean = False
            For Each relation As Syllogism In Relations
                If relation.Antecedant = SearchAndecedant = True Then
                    found = True
                Else

                End If
            Next
            Return found
        End Function

        ''' <summary>
        ''' Searches Premise list of syllogisms to see if antecedent is contained within as a consequent
        ''' </summary>
        ''' <param name="SearchConsequent"></param>
        ''' <param name="Relations"></param>
        ''' <returns></returns>
        Public Shared Function ContainsConsequent(ByRef SearchConsequent As String,
                                   ByVal Relations As List(Of Syllogism)) As Boolean
            Dim found As Boolean = False
            For Each relation As Syllogism In Relations
                If relation.Consequent = SearchConsequent = True Then
                    found = True
                Else

                End If
            Next
            Return found
        End Function

        ''' <summary>
        ''' Extracts Only Syllogisms of this type from the list
        ''' </summary>
        ''' <param name="MasterList"></param>
        ''' <returns></returns>
        Public Shared Function GetConjunctiveSyllogismList(ByRef MasterList As List(Of Syllogism)) As List(Of Syllogism)
            Dim NewLst As New List(Of Syllogism)
            For Each item As Syllogism In MasterList
                If item.Type = SyllogismType.Conjunctive Then

                    NewLst.Add(item)
                Else
                End If
            Next
            Return NewLst
        End Function

        ''' <summary>
        ''' Extracts Only Syllogisms of this type from the list
        ''' </summary>
        ''' <param name="MasterList"></param>
        ''' <returns></returns>
        Public Shared Function GetDeductiveSyllogismList(ByRef MasterList As List(Of Syllogism)) As List(Of Syllogism)
            Dim NewLst As New List(Of Syllogism)
            For Each item As Syllogism In MasterList
                If item.Type = SyllogismType.Deductive Then
                    NewLst.Add(item)
                Else
                End If
            Next
            Return NewLst
        End Function

        ''' <summary>
        ''' Extracts Only Syllogisms of this type from the list
        ''' </summary>
        ''' <param name="MasterList"></param>
        ''' <returns></returns>
        Public Shared Function GetDisjunctiveSyllogismList(ByRef MasterList As List(Of Syllogism)) As List(Of Syllogism)
            Dim NewLst As New List(Of Syllogism)
            For Each item As Syllogism In MasterList
                If item.Type = SyllogismType.Disjunctive Then

                    NewLst.Add(item)
                Else
                End If
            Next
            Return NewLst
        End Function

        ''' <summary>
        ''' Extracts Only Syllogisms of this type from the list
        ''' </summary>
        ''' <param name="MasterList"></param>
        ''' <returns></returns>
        Public Shared Function GetHypotheticalSyllogismList(ByRef MasterList As List(Of Syllogism)) As List(Of Syllogism)
            Dim NewLst As New List(Of Syllogism)
            For Each item As Syllogism In MasterList
                If item.Type = SyllogismType.Hypothetical Then

                    NewLst.Add(item)
                Else
                End If
            Next
            Return NewLst
        End Function

        ''' <summary>
        ''' Checks the syllogism presented to see if the Consequent and Antecedent Match returns
        ''' the syllogism if found else a empty syllogism
        ''' </summary>
        ''' <param name="Consequent"></param>
        ''' <param name="Andecedant"></param>
        ''' <param name="Premises"></param>
        ''' <returns></returns>
        Public Shared Function GetSyllogism(ByRef Andecedant As String, ByRef Consequent As String,
                            ByVal Premises As List(Of Syllogism)) As Syllogism
            GetSyllogism = New Syllogism
            If CheckSyllogismExists(Andecedant, Consequent, Premises) = False Then
            Else
                For Each Premise As Syllogism In Premises
                    If Premise.Consequent = Consequent And
                    Premise.Antecedant = Andecedant = True Then

                        Return Premise
                    Else

                    End If
                Next
            End If
        End Function

        ''' <summary>
        ''' Checks if Term has or Consequent Relations indicating a subtree is present
        ''' </summary>
        ''' <param name="Term"></param>
        ''' <param name="SearchList"></param>
        ''' <returns></returns>
        Public Shared Function HasAntecedantRelations(ByRef Term As String,
                                     ByVal SearchList As List(Of Syllogism)) As Boolean
            HasAntecedantRelations = False
            'SearchPredicates
            For Each item In SearchList
                If item.Antecedant = Term = True And
                    HasAntecedantRelations = False Then
                    HasAntecedantRelations = True
                Else
                End If
            Next

        End Function

        ''' <summary>
        ''' Checks if Term has or Consequent Relations indicating a subtree is present
        ''' </summary>
        ''' <param name="Term"></param>
        ''' <param name="SearchList"></param>
        ''' <returns></returns>
        Public Shared Function HasConsequentRelations(ByRef Term As String,
                                     ByVal SearchList As List(Of Syllogism)) As Boolean
            HasConsequentRelations = False
            'SearchPredicates
            For Each item In SearchList
                If item.Consequent = Term = True And
                    HasConsequentRelations = False Then
                    HasConsequentRelations = True
                Else
                End If
            Next

        End Function

        Public Shared Function JoinList(ByRef ListA As List(Of Syllogism),
                                                                 ByRef ListB As List(Of Syllogism)) As List(Of Syllogism)
            Dim ListC As New List(Of Syllogism)
            For Each item In ListA
                ListC.Add(item)
            Next

            For Each item In ListB
                If CheckSyllogismExists(item, ListC) = False Then
                    ListC.Add(item)
                End If

            Next
            Return ListC
        End Function

        ''' <summary>
        ''' Adds a new premise to master List
        ''' </summary>
        ''' <param name="NewPremise"></param>
        Public Sub AddPremise(ByRef NewPremise As Syllogism)
            mMasterList.Add(NewPremise)
        End Sub

        Public Function ToJson() As String
            Dim Converter As New JavaScriptSerializer
            Return Converter.Serialize(Me)

        End Function

        ''' <summary>
        ''' to be implemented
        ''' </summary>
        Private Shared Sub AddDbPremise(ByRef NewPremise As Syllogism)

        End Sub

        ''' <summary>
        ''' to be implemented
        ''' </summary>
        ''' <returns></returns>
        Private Shared Function GetDBPremises() As List(Of Syllogism)
            Dim Lst As New List(Of Syllogism)

            Return Lst
        End Function

        ''' <summary>
        ''' to be implemented
        ''' </summary>
        Private Shared Sub RemoveDbPremise(ByRef NewPremise As Syllogism)

        End Sub

        ''' <summary>
        ''' to be implemented
        ''' </summary>
        Private Shared Sub UpdateDbPremise(ByRef NewPremise As Syllogism)

        End Sub

    End Structure

    ''' <summary>
    ''' Used to output information from the Class
    ''' </summary>
    Public Structure PosSentence

#Region "Fields"

        Dim _Actions As List(Of Word)

        Dim _NounPhrases As List(Of Word)

        Dim _NumberOfWords As Integer

        Dim _People As List(Of Word)

        Dim _Subjects As List(Of Word)

        Dim _tense As tense

        Dim _VerbPhrases As List(Of Word)

        Dim _Words As List(Of Word)

#End Region

#Region "Methods"

        Public Function ToJson() As String
            Dim Converter As New JavaScriptSerializer
            Return Converter.Serialize(Me)

        End Function

        Public Overrides Function ToString() As String
            Dim Str As String = ""

            Str &= vbNewLine & vbNewLine & "ACTIONS :" & vbNewLine & Me._Actions.ToJson
            Str &= vbNewLine & vbNewLine & "NOUN_PHRASES :" & vbNewLine & Me._NounPhrases.ToJson
            Str &= vbNewLine & vbNewLine & "PEOPLE :" & vbNewLine & Me._People.ToJson
            Str &= vbNewLine & vbNewLine & "NUMBER_OF_WORDS :" & vbNewLine & Me._NumberOfWords.ToJson
            Str &= vbNewLine & vbNewLine & "SUBJECTS :" & vbNewLine & Me._Subjects.ToJson
            Str &= vbNewLine & vbNewLine & "VERB_PHRASES :" & vbNewLine & Me._VerbPhrases.ToJson
            Select Case Me._tense
                Case tense.Past
                    Str &= vbNewLine & vbNewLine & "TENSE :" & "Past"
                Case tense.Present
                    Str &= vbNewLine & vbNewLine & "TENSE :" & "Present"
                Case tense.Future
                    Str &= vbNewLine & vbNewLine & "TENSE :" & "Future"

            End Select
            Str &= vbNewLine & vbNewLine & "WORDS :" & vbNewLine & Me._Words.ToJson

            Return Str
        End Function

#End Region

    End Structure

    Public Structure QuestionModel

        ''' <summary>
        ''' Context Containing Data
        ''' </summary>
        Public Context As String

        ''' <summary>
        ''' List of potential questions and answers to be found in text
        ''' </summary>
        Public Questions As List(Of ContextQuestion)

        ''' <summary>
        ''' USeful Encoding for grouping Questions
        ''' </summary>
        Public Enum _QuestionType
            Where_is
            Where_are
            When_is
            When_Was
            When_Are
            Who_is
            Who_was
            Who_are
            Why_Do
            Why_are
            What_is
            What_are
            What_do
            How_do
            How_are
            How_is
            Which
            Whose_is
            Whose_are
            Who_has
        End Enum

        ''' <summary>
        ''' _ResponseTypes are
        '''place           'Where is X - Returns Answer to Question (Where is becomes the subject) X is at location of Subject
        '''time            'When was - Returns Answer to Question (when was X) X was Subject
        '''person          'Who is - Returns Answer to Question - Object becomes the search term for answer replaced by Subject
        '''reason          'Why is - Returns Answer to Question
        '''description     'What is - Returns Answer to Question
        '''instruction     'How do - Returns Answer to Question
        '''Choice          'Which X, either X or Y - Returns Selected choice from input (Which becomes the Subject X and Y are both the objects)
        '''Possession      'Whose X is? - returns Question about Ownership (Whose is the subject X is the object) X belongs to object
        ''' </summary>
        Public Enum _ResponseType
            place           'Where is - Returns Answer to Question
            time            'When was - Returns Answer to Question
            person          'Who is - Returns Answer to Question
            reason          'Why is - Returns Answer to Question
            description     'What is - Returns Answer to Question
            instruction     'How do - Returns Answer to Question
            Choice          'Which X, either X or Y - Returns Selected choice from input
            Possession      'Whose X is? - returns Question about Ownwership (who_has)
        End Enum

        ''' <summary>
        ''' The main difference between extractive and abstractive question answering
        ''' is that extractive question answering involves selecting a span of text from the context that contains the answer,
        ''' while abstractive question answering involves generating a new sentence that answers the question based on the context.
        ''' </summary>
        Public Enum ContextAnswerType

            ''' <summary>
            ''' Here is an example of an extractive question and answer:
            ''' Question: What Is the capital Of France?
            ''' Context: France Is a country located In Western Europe. Its capital Is Paris.
            ''' Answer: Paris
            '''The answer Is extractive because it Is a direct copy Of the context.
            '''In this Case, the answer “Paris” Is a span Of text that was copied directly from the context
            '''“France is a country located in Western Europe. Its capital is Paris.”
            ''' </summary>
            extractive

            ''' <summary>
            ''' Here is an example of an abstractive question and answer:
            ''' Question: What Is the capital Of France?
            ''' Context: France Is a country located In Western Europe. Its capital Is Paris.
            ''' Answer: Paris Is the capital Of France.
            ''' The answer is abstractive because it is not a direct copy of the context.
            ''' Instead, it is a new sentence that answers the question based on the context.
            ''' In this case, the model generated the sentence “Paris is the capital of France”
            ''' based on the context “France is a country located in Western Europe. Its capital is Paris.”
            ''' </summary>
            abstractive

        End Enum

        ''' <summary>
        ''' Used For Holding a Question and Answer
        ''' Position can also be used to store the location of an answer in a COntext
        ''' </summary>
        Public Structure ContextQuestion

            ''' <summary>
            ''' Answer
            ''' </summary>
            Public Answer As String

            ''' <summary>
            ''' Position of answer in text Boundry
            ''' </summary>
            Public PositionEnd As Integer

            ''' <summary>
            ''' Position of answer in context boundry
            ''' </summary>
            Public PositionStart As Integer

            ''' <summary>
            ''' Question
            ''' </summary>
            Public Question As String

            ''' <summary>
            ''' Indication of Question type
            ''' </summary>
            Public QuestionType As _QuestionType

            ''' <summary>
            ''' Indication of Answer Type
            ''' </summary>
            Public Responsetype As _ResponseType

        End Structure

    End Structure

    ''' <summary>
    ''' Used to retrieve Learning PAtterns
    ''' Learning Pattern / Nym
    ''' </summary>
    Public Structure SemanticPattern

#Region "Fields"

        ''' <summary>
        ''' Tablename in db
        ''' </summary>
        Public Shared SemanticPatternTable As String = "SemanticPatterns"

        ''' <summary>
        ''' Used to hold the connection string
        ''' </summary>
        Public ConectionStr As String

        ''' <summary>
        ''' used to identify patterns
        ''' </summary>
        Public NymStr As String

        ''' <summary>
        ''' Search patterns A# is B#
        ''' </summary>
        Public SearchPatternStr As String

#End Region

#Region "Methods"

        ''' <summary>
        ''' filters collection of patterns by nym
        ''' </summary>
        ''' <param name="Patterns">patterns </param>
        ''' <param name="NymStr">nym to be segmented</param>
        ''' <returns></returns>
        Public Shared Function FilterSemanticPatternsbyNym(ByRef Patterns As List(Of SemanticPattern), ByRef NymStr As String) As List(Of SemanticPattern)
            Dim Lst As New List(Of SemanticPattern)
            For Each item In Patterns
                If item.NymStr = NymStr Then
                    Lst.Add(item)
                End If
            Next
            If Lst.Count > 0 Then
                Return Lst
            Else
                Return Nothing
            End If
        End Function

        ''' <summary>
        ''' Gets all Semantic Patterns From Table
        ''' </summary>
        ''' <param name="iConnectionStr"></param>
        ''' <param name="TableName"></param>
        ''' <returns></returns>
        Public Shared Function GetDBSemanticPatterns(ByRef iConnectionStr As String, ByRef TableName As String) As List(Of SemanticPattern)
            Dim DbSubjectLst As New List(Of SemanticPattern)

            Dim SQL As String = "SELECT * FROM " & TableName
            Using conn = New System.Data.OleDb.OleDbConnection(iConnectionStr)
                Using cmd = New System.Data.OleDb.OleDbCommand(SQL, conn)
                    conn.Open()
                    Try
                        Dim dr = cmd.ExecuteReader()
                        While dr.Read()
                            Dim NewKnowledge As New SemanticPattern With {
                                .NymStr = dr("Nym").ToString(),
                                .SearchPatternStr = dr("SemanticPattern").ToString()
                            }
                            DbSubjectLst.Add(NewKnowledge)
                        End While
                    Catch e As Exception
                        ' Do some logging or something.
                        MessageBox.Show("There was an error accessing your data. GetDBSemanticPatterns: " & e.ToString())
                    End Try
                End Using
            End Using
            Return DbSubjectLst
        End Function

        ''' <summary>
        ''' gets semantic patterns from table based on query SQL
        ''' </summary>
        ''' <param name="iConnectionStr"></param>
        ''' <param name="Query"></param>
        ''' <returns></returns>
        Public Shared Function GetDBSemanticPatternsbyQuery(ByRef iConnectionStr As String, ByRef Query As String) As List(Of SemanticPattern)
            Dim DbSubjectLst As New List(Of SemanticPattern)

            Dim SQL As String = Query
            Using conn = New System.Data.OleDb.OleDbConnection(iConnectionStr)
                Using cmd = New System.Data.OleDb.OleDbCommand(SQL, conn)
                    conn.Open()
                    Try
                        Dim dr = cmd.ExecuteReader()
                        While dr.Read()
                            Dim NewKnowledge As New SemanticPattern With {
                                    .NymStr = dr("Nym").ToString(),
                                    .SearchPatternStr = dr("SemanticPattern").ToString()
                                }
                            DbSubjectLst.Add(NewKnowledge)
                        End While
                    Catch e As Exception
                        ' Do some logging or something.
                        MessageBox.Show("There was an error accessing your data. GetDBSemanticPatterns: " & e.ToString())
                    End Try
                End Using
            End Using
            Return DbSubjectLst
        End Function

        ''' <summary>
        ''' gets random pattern from list
        ''' </summary>
        ''' <param name="Patterns"></param>
        ''' <returns></returns>
        Public Shared Function GetRandomPattern(ByRef Patterns As List(Of SemanticPattern)) As SemanticPattern
            Dim rnd = New Random()
            If Patterns.Count > 0 Then

                Return Patterns(rnd.Next(0, Patterns.Count))
            Else
                Return Nothing
            End If
        End Function

        ''' <summary>
        ''' used to generalize patterns into general search patterns
        ''' (a# is b#) to (* is a *)
        ''' </summary>
        ''' <param name="Patterns"></param>
        ''' <returns></returns>
        Public Shared Function InsertWildcardsIntoPatterns(ByRef Patterns As List(Of SemanticPattern)) As List(Of SemanticPattern)
            For Each item In Patterns
                item.SearchPatternStr.Replace("A#", "*")
                item.SearchPatternStr.Replace("B#", "*")
            Next
            Return Patterns
        End Function

        ''' <summary>
        ''' Adds a New Semantic pattern
        ''' </summary>
        ''' <param name="NewSemanticPattern"></param>
        Public Function AddSemanticPattern(ByRef iConnectionStr As String, ByRef Tablename As String, ByRef NewSemanticPattern As SemanticPattern) As Boolean
            AddSemanticPattern = False
            If NewSemanticPattern.NymStr IsNot Nothing And NewSemanticPattern.SearchPatternStr IsNot Nothing Then

                Dim sql As String = "INSERT INTO " & Tablename & " (Nym, SemanticPattern) VALUES ('" & NewSemanticPattern.NymStr & "','" & NewSemanticPattern.SearchPatternStr & "')"

                Using conn = New System.Data.OleDb.OleDbConnection(iConnectionStr)

                    Using cmd = New System.Data.OleDb.OleDbCommand(sql, conn)
                        conn.Open()
                        Try
                            cmd.ExecuteNonQuery()
                            AddSemanticPattern = True
                        Catch ex As Exception
                            MessageBox.Show("There was an error accessing your data. AddSemanticPattern: " & ex.ToString())
                        End Try
                    End Using
                End Using
            Else
            End If
        End Function

        Public Function CheckIfSemanticPatternDetected(ByRef iConnectionStr As String, ByRef TableName As String, ByRef Userinput As String) As Boolean
            CheckIfSemanticPatternDetected = False
            For Each item In InsertWildcardsIntoPatterns(GetDBSemanticPatterns(iConnectionStr, TableName))
                If Userinput Like item.SearchPatternStr Then
                    Return True
                End If
            Next
        End Function

        Public Function GetDetectedSemanticPattern(ByRef iConnectionStr As String, ByRef TableName As String, ByRef Userinput As String) As SemanticPattern
            GetDetectedSemanticPattern = Nothing
            For Each item In InsertWildcardsIntoPatterns(GetDBSemanticPatterns(iConnectionStr, TableName))
                If Userinput Like item.SearchPatternStr Then
                    Return item
                End If
            Next
        End Function

        ''' <summary>
        ''' output in json format
        ''' </summary>
        ''' <returns></returns>
        Public Function ToJson() As String
            Dim Converter As New JavaScriptSerializer
            Return Converter.Serialize(Me)
        End Function

#End Region

    End Structure

    'A) (1) store Premises  /
    Public Structure Syllogism
        Public Shared NegationTerm As String = "It is not the case that"

        Public Antecedant As String

        Public Confidence As Integer

        Public Consequent As String

        Public Copula As String

        Public Determiner As String

        Public Negated As Boolean

        Public PropositionalLogic As LOGIC

        Public Truth As Boolean

        Public Type As SyllogismType

        ''' <summary>
        ''' Logic Of Syllogism
        ''' </summary>
        Public Enum LOGIC

            ''' <summary>
            ''' SOME A ARE B
            ''' </summary>
            _SOME_ARE

            ''' <summary>
            ''' SOME A ARE NOT B
            ''' </summary>
            _SOME_ARENOT

            ''' <summary>
            ''' ALL A ARE B
            ''' </summary>
            _ALL_ARE

            ''' <summary>
            ''' ALL A ARE B
            ''' </summary>
            _ALL_ARENOT

            ''' <summary>
            ''' IF A THEN B
            ''' </summary>
            _IF_THEN

            ''' <summary>
            ''' A AND B
            ''' </summary>
            _AND

            ''' <summary>
            ''' A OR B
            ''' </summary>
            _OR

            ''' <summary>
            ''' EITHER A OR B
            ''' </summary>
            _XOR

            ''' <summary>
            ''' NOT A AND B
            ''' </summary>
            _NAND

            ''' <summary>
            ''' NOT A OR B
            ''' </summary>
            _NOR

            ''' <summary>
            ''' NOT IF A THEN B
            ''' </summary>
            _NOT_IF_THEN

            ''' <summary>
            ''' A IF B
            ''' </summary>
            _IF

            ''' <summary>
            ''' NOT A IF B
            ''' </summary>
            _NOT_IF

            ''' <summary>
            ''' A = TRUE only IF B = TRUE
            ''' </summary>
            _UNLESS_IF

            ''' <summary>
            ''' Used for simple statement Logic
            ''' </summary>
            _CUSTOM

        End Enum

        ''' <summary>
        ''' type of syllogism
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum SyllogismType
            Hypothetical
            Disjunctive
            Conjunctive
            Deductive
            SimpleStatement
        End Enum

        ''' <summary>
        ''' Compares parameters stored in the Syllogism
        ''' </summary>
        ''' <param name="A"></param>
        ''' <param name="B"></param>
        ''' <returns></returns>
        Public Shared Function Compare(ByRef A As Syllogism, ByRef B As Syllogism) As Boolean
            Compare = False
            If A.Determiner = B.Determiner = True And
        A.Antecedant = B.Antecedant = True And
        A.Copula = B.Copula = True And
        A.Consequent = B.Consequent = True And
        A.Negated = B.Negated = True And
        A.PropositionalLogic = B.PropositionalLogic = True And
        A.Confidence = B.Confidence = True And
        A.Truth = B.Truth = True And
        A.Type = B.Type = True Then
                Compare = True
            End If
        End Function

        ''' <summary>
        ''' Displays the contents of the object
        ''' </summary>
        ''' <returns></returns>
        Public Function DataToString() As String
            Dim Str As String = ""
            Str += "Determiner: " & Determiner & vbNewLine
            Str += "Antecedent: " & Antecedant & vbNewLine
            Str += "Copula: " & Copula & vbNewLine
            Str += "Consequent :" & Consequent & vbNewLine
            Str += "Confidence :" & Confidence & vbNewLine
            Str += "Negated :" & Negated & vbNewLine
            Str += "NegationTerm :" & NegationTerm & vbNewLine
            Str += "Type :" & Type & vbNewLine
            Str += "PropositionalLogic :" & PropositionalLogic & vbNewLine
            Str += "Truth :" & Truth & vbNewLine
            Return Str
        End Function

        Public Function ToJson() As String
            Dim Converter As New JavaScriptSerializer
            Return Converter.Serialize(Me)

        End Function

        Public Function ToNode() As System.Windows.Forms.TreeNode
            Dim Node As New System.Windows.Forms.TreeNode

            Dim Node1 As New System.Windows.Forms.TreeNode
            Node1.Text = "Determiner: " & Determiner
            Node.Nodes.Add(Node1)
            Dim Node2 As New System.Windows.Forms.TreeNode
            Node2.Text = "Antecedent: " & Antecedant
            Node.Nodes.Add(Node2)
            Dim Node3 As New System.Windows.Forms.TreeNode
            Node3.Text = "Copula: " & Copula
            Node.Nodes.Add(Node3)
            Dim Node4 As New System.Windows.Forms.TreeNode
            Node4.Text = "Consequent :" & Consequent
            Node.Nodes.Add(Node4)
            Dim Node5 As New System.Windows.Forms.TreeNode
            Node5.Text = "Confidence :" & Confidence
            Node.Nodes.Add(Node5)
            Dim Node6 As New System.Windows.Forms.TreeNode
            Node6.Text = "Negated :" & Negated
            Node.Nodes.Add(Node6)
            Dim Node7 As New System.Windows.Forms.TreeNode
            Node7.Text = "NegationTerm :" & NegationTerm
            Node.Nodes.Add(Node7)
            Dim Node8 As New System.Windows.Forms.TreeNode
            Node8.Text = "Type :" & Type & vbNewLine
            Node.Nodes.Add(Node8)
            Dim Node9 As New System.Windows.Forms.TreeNode
            Node9.Text = "PropositionalLogic :" & PropositionalLogic & vbNewLine
            Node.Nodes.Add(Node9)
            Dim Node10 As New System.Windows.Forms.TreeNode
            Node10.Text = "Truth :" & Truth & vbNewLine
            Node.Nodes.Add(Node10)
            Return Node
        End Function

        ''' <summary>
        ''' Displays a String Output of Determiner / Antecedent / Copula / Consequent Displays
        ''' Case in Its Normalized Form
        ''' </summary>
        ''' <returns></returns>
        Public Overrides Function ToString() As String
            Dim Str As String = ""
            If NegationTerm = "" Then NegationTerm = "IT IS NOT THE CASE THAT"
            Select Case PropositionalLogic
                Case LOGIC._CUSTOM
                    Str = If(Negated = True, NegationTerm & " " &
Determiner & " " &
Antecedant & " " &
Copula & " " &
Consequent & " ", Determiner & " " &
Antecedant & " " &
Copula & " " &
Consequent & " ")

                Case LOGIC._ALL_ARE
                    Str = If(Negated = True, NegationTerm & " ALL " & Antecedant & " " & " ARE " & Consequent & " ", " ALL " & Antecedant & " ARE " & Consequent & " ")
                Case LOGIC._ALL_ARENOT
                    Str = If(Negated = True, NegationTerm & " " & " ALL " & Antecedant & " ARE NOT " & Consequent & " ", " ALL " & Antecedant & " ARE NOT " & Consequent & " ")
                Case LOGIC._AND
                    Str = If(Negated = True, NegationTerm & " " & " BOTH " & Antecedant & " AND " & Consequent & " ", " BOTH " & Antecedant & " AND " & Consequent & " ")
                Case LOGIC._NAND
                    Str = If(Negated = True, NegationTerm & " " & " NOT " & Antecedant & " AND " & Consequent & " ", " NOT " & Antecedant & " AND " & Consequent & " ")
                Case LOGIC._IF_THEN
                    Str = If(Negated = True, NegationTerm & " IF " & Antecedant & " THEN " & Consequent & " ", " IF " & Antecedant & " THEN " & Consequent & " ")
                Case LOGIC._NOT_IF_THEN
                    Str = If(Negated = True, NegationTerm & " " & " NOT-IF " & Antecedant & " THEN " & Consequent & " ", " NOT-IF " & Antecedant & " THEN " & Consequent & " ")
                Case LOGIC._UNLESS_IF
                    Str = If(Negated = True, NegationTerm & " " & Antecedant & " UNLESS IF " & Consequent & " ", Antecedant & " " & " UNLESS IF " & Consequent & " ")
                Case LOGIC._IF
                    Str = If(Negated = True, NegationTerm & " " & Antecedant & " " & " IF " & Consequent & " ", Antecedant & " " & " IF " & Consequent & " ")
                Case LOGIC._NOT_IF
                    Str = If(Negated = True, NegationTerm & " NOT " & Antecedant & " " & " IF " & Consequent & " ", " NOT " & Antecedant & " IF " & Consequent & " ")
                Case LOGIC._OR
                    Str = If(Negated = True, NegationTerm & " " & Antecedant & " " & " OR " & Consequent & " ", Determiner & " " & Antecedant & " " & " OR " & Consequent & " ")
                Case LOGIC._SOME_ARE
                    Str = If(Negated = True, NegationTerm & " " & " SOME " & Antecedant & " " & " ARE " & Consequent & " ", Determiner & " ALL " & Antecedant & " " & " ARE NOT " & Consequent & " ")
                Case LOGIC._SOME_ARENOT
                    Str = If(Negated = True, NegationTerm & " " & " SOME " & Antecedant & " " & " ARE NOT " & Consequent & " ", Determiner & " SOME " & Antecedant & " " & " ARE NOT " & Consequent & " ")
                Case LOGIC._XOR
                    Str = If(Negated = True, NegationTerm & " " & " EITHER " & Antecedant & " " & " OR " & Consequent & " ", " EITHER " & Antecedant & " " & " OR " & Consequent & " ")
                Case LOGIC._NOR
                    Str = If(Negated = True, NegationTerm & " NEITHER " & Antecedant & " " & " OR " & Consequent & " ", " NEITHER " & Antecedant & " " & " OR " & Consequent & " ")
            End Select

            Return Str
        End Function

    End Structure

    ''' <summary>
    ''' This is a list of Known SentenceStructures stored in db ; Each sentence which has
    ''' been successfully analasied for parts of speech; the structure is saved; as a list of string
    ''' </summary>
    Public Structure Taglist

#Region "Fields"

        Public Sentence As String
        Public TagList As List(Of String)

#End Region

#Region "Methods"

        Public Function ToJson() As String
            Dim Converter As New JavaScriptSerializer
            Return Converter.Serialize(Me)

        End Function

#End Region

    End Structure

    ''' <summary>
    ''' Training Data -Consisting of input vector and output vector
    ''' </summary>
    Public Structure TrainingData
        Dim ExpectedOutput As Vector
        Dim input As Vector
    End Structure

    ''' <summary>
    ''' Transfer Function used in the calculation of the following layer
    ''' </summary>
    Public Structure TransferFunction

        ''' <summary>
        ''' Returns a result from the transfer function indicated ; Non Derivative
        ''' </summary>
        ''' <param name="TransferFunct">Indicator for Transfer function selection</param>
        ''' <param name="Input">Input value for node/Neuron</param>
        ''' <returns>result</returns>
        Public Shared Function EvaluateTransferFunct(ByRef TransferFunct As TransferFunctionType, ByRef Input As Double) As Integer
            EvaluateTransferFunct = 0
            Select Case TransferFunct
                Case TransferFunctionType.none
                    Return Input
                Case TransferFunctionType.sigmoid
                    Return Sigmoid(Input)
                Case TransferFunctionType.HyperbolTangent
                    Return HyperbolicTangent(Input)
                Case TransferFunctionType.BinaryThreshold
                    Return BinaryThreshold(Input)
                Case TransferFunctionType.RectifiedLinear
                    Return RectifiedLinear(Input)
                Case TransferFunctionType.Logistic
                    Return Logistic(Input)
                Case TransferFunctionType.Gaussian
                    Return Gaussian(Input)
                Case TransferFunctionType.Signum
                    Return Signum(Input)
            End Select
        End Function

        ''' <summary>
        ''' Returns a result from the transfer function indicated ; Non Derivative
        ''' </summary>
        ''' <param name="TransferFunct">Indicator for Transfer function selection</param>
        ''' <param name="Input">Input value for node/Neuron</param>
        ''' <returns>result</returns>
        Public Shared Function EvaluateTransferFunctionDerivative(ByRef TransferFunct As TransferFunctionType, ByRef Input As Double) As Integer
            EvaluateTransferFunctionDerivative = 0
            Select Case TransferFunct
                Case TransferFunctionType.none
                    Return Input
                Case TransferFunctionType.sigmoid
                    Return SigmoidDerivitive(Input)
                Case TransferFunctionType.HyperbolTangent
                    Return HyperbolicTangentDerivative(Input)
                Case TransferFunctionType.Logistic
                    Return LogisticDerivative(Input)
                Case TransferFunctionType.Gaussian
                    Return GaussianDerivative(Input)
            End Select
        End Function

        ''' <summary>
        ''' the step function rarely performs well except in some rare cases with (0,1)-encoded
        ''' binary data.
        ''' </summary>
        ''' <param name="Value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function BinaryThreshold(ByRef Value As Double) As Double

            ' Z = Bias+ (Input*Weight)
            'TransferFunction
            'If Z > 0 then Y = 1
            'If Z < 0 then y = 0

            Return If(Value < 0 = True, 0, 1)
        End Function

        Private Shared Function Gaussian(ByRef x As Double) As Double
            Gaussian = Math.Exp((-x * -x) / 2)
        End Function

        Private Shared Function GaussianDerivative(ByRef x As Double) As Double
            GaussianDerivative = Gaussian(x) * (-x / (-x * -x))
        End Function

        Private Shared Function HyperbolicTangent(ByRef Value As Double) As Double
            ' TanH(x) = (Math.Exp(x) - Math.Exp(-x)) / (Math.Exp(x) + Math.Exp(-x))

            Return Math.Tanh(Value)
        End Function

        Private Shared Function HyperbolicTangentDerivative(ByRef Value As Double) As Double
            HyperbolicTangentDerivative = 1 - (HyperbolicTangent(Value) * HyperbolicTangent(Value)) * Value
        End Function

        'Linear Neurons
        ''' <summary>
        ''' in a liner neuron the weight(s) represent unknown values to be determined the
        ''' outputs could represent the known values of a meal and the inputs the items in the
        ''' meal and the weights the prices of the individual items There are no hidden layers
        ''' </summary>
        ''' <remarks>
        ''' answers are determined by determining the weights of the linear neurons the delta
        ''' rule is used as the learning rule: Weight = Learning rate * Input * LocalError of neuron
        ''' </remarks>
        Private Shared Function Linear(ByRef value As Double) As Double
            ' Output = Bias + (Input*Weight)
            Return value
        End Function

        'Non Linear neurons
        Private Shared Function Logistic(ByRef Value As Double) As Double
            'z = bias + (sum of all inputs ) * (input*weight)
            'output = Sigmoid(z)
            'derivative input = z/weight
            'derivative Weight = z/input
            'Derivative output = output*(1-Output)
            'learning rule = Sum of total training error* derivative input * derivative output * rootmeansquare of errors

            Return 1 / 1 + Math.Exp(-Value)
        End Function

        Private Shared Function LogisticDerivative(ByRef Value As Double) As Double
            'z = bias + (sum of all inputs ) * (input*weight)
            'output = Sigmoid(z)
            'derivative input = z/weight
            'derivative Weight = z/input
            'Derivative output = output*(1-Output)
            'learning rule = Sum of total training error* derivative input * derivative output * rootmeansquare of errors

            Return Logistic(Value) * (1 - Logistic(Value))
        End Function

        Private Shared Function RectifiedLinear(ByRef Value As Double) As Double
            'z = B + (input*Weight)
            'If Z > 0 then output = z
            'If Z < 0 then output = 0
            If Value < 0 = True Then

                Return 0
            Else
                Return Value
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
        Private Shared Function Sigmoid(ByRef Value As Integer) As Double
            'z = Bias + (Input*Weight)
            'Output = 1/1+e**z
            Return 1 / (1 + Math.Exp(-Value))
        End Function

        Private Shared Function SigmoidDerivitive(ByRef Value As Integer) As Double
            Return Sigmoid(Value) * (1 - Sigmoid(Value))
        End Function

        Private Shared Function Signum(ByRef Value As Integer) As Double
            'z = Bias + (Input*Weight)
            'Output = 1/1+e**z
            Return Math.Sign(Value)
        End Function

        Private Shared Function StochasticBinary(ByRef value As Double) As Double
            'Uncreated
            Return value
        End Function

    End Structure

    'Tree Extensions
    ''' <summary>
    ''' A Tree is actually a List of Lists
    ''' Root with branches and leaves:
    ''' In this tree the children are the branches: Locations to hold data have been provided.
    ''' These are not part of the Tree Structure:
    ''' When initializing the structure Its also prudent to Initialize the ChildList;
    ''' Reducing errors; Due to late initialization of the ChildList
    ''' Sub sequentially : Lists used in this class are not initialized in this structure.
    ''' Strings inserted with the insert Cmd will be treated as a Trie Tree insert!
    ''' if this tree requires data to be stored it needs to be stored inside the dataStorae locations
    ''' </summary>
    Public Structure TrieTree

#Region "Fields"

        ''' <summary>
        ''' Theses are essentailly the branches of the Tree....
        ''' if there are no children then the node is a leaf.
        ''' Denoted by the StopCharacter NodeID. or "StopChar" NodeData.
        ''' </summary>
        Public Children As List(Of TrieTree)

        ''' <summary>
        ''' Available Variable Storage (boolean)
        ''' </summary>
        Public NodeBool As Boolean

        ''' <summary>
        ''' Available Variable Storage (list of Boolean)
        ''' </summary>
        Public NodeBoolList As List(Of Boolean)

        ''' <summary>
        ''' Used To hold CharacterStr (tries) Also Useful for Creating ID for node;
        ''' : (a String) can be used to store a specific Pathway :
        ''' </summary>
        Public NodeData As String

        ''' <summary>
        ''' Each NodeID: is specific to Its level this id is generated by the tree:
        ''' </summary>
        Public NodeID As CharID

        ''' <summary>
        ''' Available Variable Storage(int)
        ''' </summary>
        Public NodeInt As Integer

        ''' <summary>
        ''' Available Variable Storage (list of Int)
        ''' </summary>
        Public NodeIntList As List(Of Integer)

        ''' <summary>
        ''' the level denotes also how many vertices before returning to the root.
        ''' In a Trie Tree the first node is blank and not included therefore the Vertices are level-1
        ''' </summary>
        Public NodeLevel As Integer

        ''' <summary>
        ''' Available Variable Storage(string)
        ''' </summary>
        Public NodeStr As String

        ''' <summary>
        ''' Available Variable Storage(list of Strings)
        ''' </summary>
        Public NodeStrList As List(Of String)

#End Region

#Region "Constructors"

        Public Sub New(ByRef IntializeChildren As Boolean)
            If IntializeChildren = True Then
                Children = New List(Of TrieTree)
            Else
            End If
        End Sub

#End Region

#Region "Enums"

        ''' <summary>
        ''' Each Set of Nodes has only 26 ID's ID 0 = stop
        ''' </summary>
        Enum CharID
            StartChar = 0
            a = 1
            b = 2
            c = 3
            d = 4
            e = 5
            f = 6
            g = 7
            h = 8
            i = 9
            j = 10
            k = 11
            l = 12
            m = 13
            n = 14
            o = 15
            p = 16
            q = 17
            r = 18
            s = 19
            t = 20
            u = 21
            v = 22
            w = 23
            x = 24
            y = 25
            z = 26
            StopChar = 27
        End Enum

#End Region

#Region "Methods"

        ''' <summary>
        ''' Add characters Iteratively
        ''' CAT
        ''' AT
        ''' T
        ''' </summary>
        ''' <param name="Word"></param>
        ''' <param name="TrieTree"></param>
        Public Shared Function AddItterativelyByCharacter(ByRef TrieTree As TrieTree, ByRef Word As String) As TrieTree
            'AddWord
            For i = 1 To Word.Length
                TrieTree.InsertByCharacters(Word)
                Word = Word.Remove(0, 1)
            Next
            Return TrieTree
        End Function

        ''' <summary>
        ''' Add by Sent Iteratively
        ''' The cat sat.
        ''' on the mat.
        ''' </summary>
        ''' <param name="Word"></param>
        ''' <param name="TrieTree"></param>
        Public Shared Function AddItterativelyBySent(ByRef TrieTree As TrieTree, ByRef Word As String) As TrieTree
            'AddWord
            Dim x = Word.Split(".")
            For Each item As String In x
                TrieTree.InsertBySent(Word)
                If Word.Length > item.Length + 1 = True Then
                    Word = Word.Remove(0, item.Length + 1)
                Else
                    Word = ""
                End If
            Next
            Return TrieTree
        End Function

        ''' <summary>
        ''' Add characters Iteratively
        ''' CAT
        ''' AT
        ''' T
        ''' </summary>
        ''' <param name="Word"></param>
        ''' <param name="TrieTree"></param>
        Public Shared Function AddItterativelyByWord(ByRef TrieTree As TrieTree, ByRef Word As String) As TrieTree
            'AddWord
            Dim x = Word.Split(" ")
            For Each item As String In x
                TrieTree.InsertByWord(Word)
                If Word.Length > item.Length + 1 = True Then
                    Word = Word.Remove(0, item.Length + 1)
                Else
                    Word = ""
                End If
            Next
            Return TrieTree
        End Function

        ''' <summary>
        ''' Adds string to given trie
        ''' </summary>
        ''' <param name="Tree"></param>
        ''' <param name="Str"></param>
        ''' <returns></returns>
        Public Shared Function AddStringbyChars(ByRef Tree As TrieTree, Str As String) As TrieTree
            Dim curr As TrieTree = Tree
            Dim Pos As Integer = 0
            For Each chr As Char In Str
                Pos += 1
                curr = TrieTree.AddStringToTrie(curr, chr, Pos)
            Next
            curr = TrieTree.AddStringToTrie(curr, "StopChar", Pos + 1)
            Return Tree
        End Function

        ''' <summary>
        ''' Adds string to given trie
        ''' </summary>
        ''' <param name="Tree"></param>
        ''' <param name="Str"></param>
        ''' <returns></returns>
        Public Shared Function AddStringbyWord(ByRef Tree As TrieTree, Str As String) As TrieTree
            Dim curr As TrieTree = Tree
            Dim Pos As Integer = 0
            For Each chr As String In Str.Split(" ")
                Pos += 1
                curr = TrieTree.AddStringToTrie(curr, chr, Pos)
            Next
            curr = TrieTree.AddStringToTrie(curr, "StopChar", Pos + 1)
            Return Tree
        End Function

        ''' <summary>
        ''' checks if node exists in child nodes (used for trie trees (String added is the key and the data)
        ''' </summary>
        ''' <param name="Nodedata">Char string used as identifier</param>
        ''' <returns></returns>
        Public Shared Function CheckNodeExists(ByRef Children As List(Of TrieTree),
                                                   ByRef Nodedata As String) As Boolean
            'Check node does not exist
            Dim found As Boolean = False
            For Each mNode As TrieTree In Children
                If mNode.NodeData = Nodedata Then
                    found = True
                Else
                End If
            Next
            Return found
        End Function

        ''' <summary>
        ''' Returns true if Word is found in trie
        ''' </summary>
        ''' <param name="tree"></param>
        ''' <param name="Str"></param>
        ''' <returns></returns>
        Public Shared Function CheckWord(ByRef tree As TrieTree, ByRef Str As String) As Boolean
            Str = Str.ToUpper
            Dim CurrentNode As TrieTree = tree
            Dim found As Boolean = False
            'Position in Characterstr
            'Check Chars
            If TrieTree.CheckNodeExists(CurrentNode.Children, Str) = True Then
                ' CurrentNode = TrieTree.GetNode(CurrentNode.Children, chrStr)
                found = True
            Else
                'Terminated before end of Word
                found = False
            End If

            'Check for end of word marker
            Return found

        End Function

        Public Shared Function FIND(ByRef TryTree As TrieTree, ByRef Txt As String) As Boolean
            Dim Found As Boolean = False
            If TryTree.FindWord(UCase(Txt)) = True Then

                Return True
            Else

            End If
            Return False
        End Function

        ''' <summary>
        ''' Returns Matched Node to sender (used to recures children)
        ''' </summary>
        ''' <param name="Tree"></param>
        ''' <param name="NodeData"></param>
        ''' <returns></returns>
        Public Shared Function GetNode(ByRef Tree As List(Of TrieTree),
                                           ByRef NodeData As String) As TrieTree
            Dim Foundnode As New TrieTree
            For Each item In Tree
                If item.NodeData = NodeData Then
                    Foundnode = item
                Else
                End If
            Next
            Return Foundnode
        End Function

        ''' <summary>
        ''' Creates a tree
        ''' </summary>
        ''' <param name="RootName">Name of Root node</param>
        ''' <returns></returns>
        Public Shared Function MakeTree(ByRef RootName As String) As TrieTree
            Dim tree As New TrieTree(True)

            tree.NodeData = RootName
            tree.NodeID = CharID.StartChar
            Return tree
        End Function

        ''' <summary>
        ''' Creates a Tree With an Empty Root node Called Root! Intializing the StartChar
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function MakeTree() As TrieTree
            Dim tree As New TrieTree(True)

            tree.NodeData = "Root"
            tree.NodeID = CharID.StartChar
            Return tree
        End Function

        ''' <summary>
        ''' Returns a TreeViewControl with the Contents of the Trie:
        ''' </summary>
        Public Shared Function ToView(ByRef Node As TrieTree) As System.Windows.Forms.TreeNode
            Dim Nde As New System.Windows.Forms.TreeNode
            Nde.Text = Node.NodeData.ToString.ToUpper &
                    "(" & Node.NodeLevel & ")" & vbNewLine

            For Each child In Node.Children
                Nde.Nodes.Add(child.ToView)

            Next
            Return Nde
        End Function

        ''' <summary>
        ''' AddIteratively characters
        ''' CAT
        ''' AT
        ''' T
        ''' </summary>
        ''' <param name="Word"></param>
        Public Function AddItterativelyByCharacter(ByRef Word As String) As TrieTree
            Return AddItterativelyByCharacter(Me, Word)
        End Function

        Public Function AddItterativelyBySent(ByRef Word As String) As TrieTree
            Return AddItterativelyByWord(Me, Word)
        End Function

        ''' <summary>
        ''' AddIteratively characters
        ''' CAT
        ''' AT
        ''' T
        ''' </summary>
        ''' <param name="Word"></param>
        Public Function AddItterativelyByWord(ByRef Word As String) As TrieTree
            Return AddItterativelyByWord(Me, Word)
        End Function

        ''' <summary>
        ''' Returns true if Word is found in trie
        ''' </summary>
        ''' <param name="Str"></param>
        ''' <returns></returns>
        Public Function CheckWord(ByRef Str As String) As Boolean
            Str = Str.ToUpper
            Dim found As Boolean = False
            'Position in Characterstr

            'Check Chars
            If TrieTree.CheckNodeExists(Me.Children, Str) = True Then
                ' CurrentNode = TrieTree.GetNode(CurrentNode.Children, chrStr)
                found = True
            Else
                'Terminated before end of Word
                found = False
            End If

            'Check for end of word marker
            Return found

        End Function

        ''' <summary>
        ''' Returns number of Nodes in tree
        ''' </summary>
        ''' <returns></returns>
        Public Function CountNodes(ByRef CurrentCount As Integer) As Integer
            Dim count As Integer = CurrentCount
            For Each child In Me.Children
                count += 1
                count = child.CountNodes(count)
            Next
            Return count
        End Function

        ''' <summary>
        ''' Returns number of Nodes in tree (ie Words)
        ''' (Number of NodeID's Until StopChar is Found)
        ''' </summary>
        ''' <returns></returns>
        Public Function CountWords(ByRef CurrentCount As Integer) As Integer
            Dim count As Integer = CurrentCount
            For Each child In Me.Children
                If child.NodeID = CharID.StopChar = True Then
                    count += 1

                End If
                count = child.CountWords(count)
            Next

            Return count
        End Function

        ''' <summary>
        ''' Creates a Tree With an Empty Root node Called Root! Intializing the StartChar
        ''' </summary>
        ''' <returns></returns>
        Public Function Create() As TrieTree
            Return TrieTree.MakeTree()
        End Function

        Public Function Create(ByRef NodeName As String) As TrieTree
            Return TrieTree.MakeTree(NodeName)
        End Function

        ''' <summary>
        ''' Returns Depth of tree
        ''' </summary>
        ''' <returns></returns>
        Public Function Depth() As Integer
            'Gets the level for node
            Dim Level As Integer = Me.NodeLevel
            'Recurses children
            For Each child In Me.Children
                If Level < child.Depth = True Then
                    Level = child.Depth
                End If
            Next
            'The loop should finish at the lowest level
            Return Level
        End Function

        Public Function FIND(ByRef Txt As String) As Boolean
            Dim Found As Boolean = False

            If Me.FindWord(UCase(Txt)) = True Then

                Return True
            Else

            End If
            Return False
        End Function

        ''' <summary>
        ''' Returns true if string is found as word in trie
        ''' </summary>
        ''' <param name="Str"></param>
        ''' <returns></returns>
        Public Function FindWord(ByRef Str As String) As Boolean
            Return TrieTree.CheckWord(Me, Str)
        End Function

        ''' <summary>
        ''' deserialize object from Json
        ''' </summary>
        ''' <param name="Str">json</param>
        ''' <returns></returns>
        Public Function FromJson(ByRef Str As String) As TrieTree
            Try
                Dim Converter As New JavaScriptSerializer
                Dim diag As TrieTree = Converter.Deserialize(Of TrieTree)(Str)
                Return diag
            Catch ex As Exception
                Dim Buttons As MessageBoxButtons = MessageBoxButtons.OK
                MessageBox.Show(ex.Message, "ERROR", Buttons)
            End Try
            Return Nothing
        End Function

        ''' <summary>
        ''' gets node if Word is found in trie
        ''' </summary>
        ''' <param name="Str"></param>
        ''' <returns></returns>
        Public Function GetNode(ByRef Str As String) As TrieTree
            Str = Str.ToUpper
            Dim found As Boolean = False
            'Position in Characterstr

            'Check Chars
            If TrieTree.CheckNodeExists(Me.Children, Str) = True Then
                Return TrieTree.GetNode(Me.Children, Str)
                found = True
            Else
                'Terminated before end of Word
                found = False
            End If

            'Check for end of word marker
            Return Nothing

        End Function

        ''' <summary>
        ''' Checks if current node Has children
        ''' </summary>
        ''' <returns></returns>
        Public Function HasChildren() As Boolean
            Return If(Me.Children.Count > 0 = True, True, False)
        End Function

        ''' <summary>
        ''' Inserts a string into the trie
        ''' </summary>
        ''' <param name="Str"></param>
        ''' <returns></returns>
        Public Function InsertByCharacters(ByRef Str As String) As TrieTree
            Return TrieTree.AddStringbyChars(Me, Str)
        End Function

        Public Function InsertBySent(ByRef Str As String) As TrieTree
            Return TrieTree.AddItterativelyBySent(Me, Str)
        End Function

        ''' <summary>
        ''' Inserts a string into the trie
        ''' </summary>
        ''' <param name="Str"></param>
        ''' <returns></returns>
        Public Function InsertByWord(ByRef Str As String) As TrieTree
            Return TrieTree.AddStringbyWord(Me, Str)
        End Function

        Public Function ToJson() As String
            Dim Converter As New JavaScriptSerializer
            Return Converter.Serialize(Me)
        End Function

        ''' <summary>
        ''' Returns a TreeViewControl with the Contents of the Trie:
        ''' </summary>
        ''' <returns></returns>
        Public Function ToView() As System.Windows.Forms.TreeNode
            Dim nde As New System.Windows.Forms.TreeNode
            nde.Text = Me.NodeData.ToString.ToUpper & vbNewLine &
                    "(" & Me.NodeLevel & ")" & vbNewLine

            For Each child In Me.Children
                nde.Nodes.Add(child.ToView)

            Next

            Return nde

        End Function

        ''' <summary>
        ''' Adds char to Node(children) Returning the child
        ''' </summary>
        ''' <param name="CurrentNode">node containing children</param>
        ''' <param name="ChrStr">String to be added </param>
        ''' <param name="CharPos">this denotes the level of the node</param>
        ''' <returns></returns>
        Private Shared Function AddStringToTrie(ByRef CurrentNode As TrieTree,
                                                    ByRef ChrStr As String, ByRef CharPos As Integer) As TrieTree
            'start of tree
            Dim Text As String = ChrStr
            Dim returnNode As New TrieTree
            Dim NewNode As New TrieTree
            'Goto first node
            'does this node have siblings
            If TrieTree.HasChildren(CurrentNode) = True Then
                'Check children
                If TrieTree.CheckNodeExists(CurrentNode.Children, ChrStr) = False Then
                    'create a new node for char
                    NewNode = TrieTree.CreateNode(ChrStr, CurrentNode.NodeLevel + 1)
                    NewNode.NodeLevel = CharPos
                    'Add childnode
                    CurrentNode.Children.Add(NewNode)
                    returnNode = TrieTree.GetNode(CurrentNode.Children, ChrStr)
                Else
                    returnNode = TrieTree.GetNode(CurrentNode.Children, ChrStr)
                End If
            Else
                'If no silings then Create new node
                'create a new node for char
                NewNode = TrieTree.CreateNode(ChrStr, CurrentNode.NodeLevel + 1)
                NewNode.NodeLevel = CharPos
                'Add childnode
                CurrentNode.Children.Add(NewNode)
                returnNode = TrieTree.GetNode(CurrentNode.Children, ChrStr)
            End If

            Return returnNode
        End Function

        ''' <summary>
        ''' If node does not exist in child node set it is added
        ''' if node already exists then no node is added a node ID is generated
        ''' </summary>
        ''' <param name="NodeData">Character to be added</param>
        Private Shared Function CreateNode(ByRef NodeData As String, ByRef Level As Integer) As TrieTree
            'Create node - 2
            Dim NewNode As New TrieTree(True)
            NewNode.NodeData = NodeData
            'create id
            NewNode.NodeID = GenerateNodeID(NodeData)
            NewNode.NodeLevel = Level
            Return NewNode
        End Function

        ''' <summary>
        ''' Generates an ID from NodeData
        ''' </summary>
        ''' <param name="Nodedata">Character string for node</param>
        ''' <returns></returns>
        Private Shared Function GenerateNodeID(ByRef Nodedata As String) As CharID
            Dim newnode As TrieTree
            'SET ID for node
            Select Case Nodedata.ToUpper
                Case "STOPCHAR"
                    newnode.NodeID = CharID.StopChar
                Case "A"
                    newnode.NodeID = CharID.a
                Case "B"
                    newnode.NodeID = CharID.b
                Case "C"
                    newnode.NodeID = CharID.c
                Case "D"
                    newnode.NodeID = CharID.d
                Case "E"
                    newnode.NodeID = CharID.e
                Case "F"
                    newnode.NodeID = CharID.f
                Case "G"
                    newnode.NodeID = CharID.g
                Case "H"
                    newnode.NodeID = CharID.h
                Case "I"
                    newnode.NodeID = CharID.i
                Case "J"
                    newnode.NodeID = CharID.j
                Case "K"
                    newnode.NodeID = CharID.k
                Case "L"
                    newnode.NodeID = CharID.l
                Case "M"
                    newnode.NodeID = CharID.m
                Case "N"
                    newnode.NodeID = CharID.n
                Case "O"
                    newnode.NodeID = CharID.o
                Case "P"
                    newnode.NodeID = CharID.p
                Case "Q"
                    newnode.NodeID = CharID.q
                Case "R"
                    newnode.NodeID = CharID.r
                Case "S"
                    newnode.NodeID = CharID.s
                Case "T"
                    newnode.NodeID = CharID.t
                Case "U"
                    newnode.NodeID = CharID.u
                Case "V"
                    newnode.NodeID = CharID.v
                Case "W"
                    newnode.NodeID = CharID.w
                Case "X"
                    newnode.NodeID = CharID.x
                Case "Y"
                    newnode.NodeID = CharID.y
                Case "Z"
                    newnode.NodeID = CharID.z
            End Select
            Return newnode.NodeID
        End Function

        ''' <summary>
        ''' Checks if given node has children
        ''' </summary>
        ''' <param name="Node"></param>
        ''' <returns></returns>
        Private Shared Function HasChildren(ByRef Node As TrieTree) As Boolean
            Return If(Node.Children.Count > 0 = True, True, False)
        End Function

#End Region

    End Structure

    ''' <summary>
    ''' Used as value for
    ''' Inputs and Outputs
    ''' </summary>
    Public Structure Vector
        Public Values As List(Of Double)

#Region "Functions (math)"

        ''' <summary>
        ''' Adds Two vectors together
        ''' </summary>
        ''' <param name="nVector"></param>
        ''' <returns></returns>
        Public Shared Function ADD(ByRef vect As Vector, ByRef nVector As Vector) As Vector
            Dim Answer As New Vector

            If nVector.Values.Count > vect.Values.Count Then

                For i = 0 To vect.Values.Count
                    Answer.Values.Add(nVector.Values(i) + vect.Values(i))
                Next
                'Add Excess values
                For i = vect.Values.Count To nVector.Values.Count
                    Answer.Values.Add(nVector.Values(i))
                Next
            Else
                '
                For i = 0 To nVector.Values.Count
                    Answer.Values.Add(nVector.Values(i) + vect.Values(i))
                Next
                'Add Excess values
                For i = nVector.Values.Count To vect.Values.Count
                    Answer.Values.Add(nVector.Values(i))
                Next
            End If

            Return Answer
        End Function

        ''' <summary>
        ''' returns dot product of two vectors
        ''' </summary>
        ''' <param name="Vect"></param>
        ''' <returns></returns>
        Public Shared Function DotProduct(ByRef Vect As Vector, ByRef NewVect As Vector) As Double
            Dim answer As Vector = NewVect.Multiply(Vect)
            Dim DotProd = answer.Sum
            Return DotProd
        End Function

        ''' <summary>
        ''' Mutiplys two vectors ;
        ''' If vectors are not matching size excess values will be 0
        ''' </summary>
        ''' <param name="nVector"></param>
        ''' <returns></returns>
        Public Shared Function Multiply(ByRef Vect As Vector, ByRef nVector As Vector) As Vector
            Dim Answer As New Vector
            If nVector.Values.Count > Vect.Values.Count Then
                For i = 0 To Vect.Values.Count
                    Answer.Values.Add(nVector.Values(i) * Vect.Values(i))
                Next
                'Add Excess values
                For i = Vect.Values.Count To nVector.Values.Count
                    Answer.Values.Add(nVector.Values(i))
                Next
            Else
                For i = 0 To nVector.Values.Count
                    Answer.Values.Add(nVector.Values(i) * Vect.Values(i))
                Next
                'Add Excess values
                For i = nVector.Values.Count To Vect.Values.Count
                    Answer.Values.Add(nVector.Values(i) * 0)
                Next
            End If
            Return Answer
        End Function

        ''' <summary>
        ''' Mutiplys Vector by Scalar ;
        ''' If vectors are not matching size excess values will be 0
        ''' </summary>
        ''' <param name="Scalar"></param>
        ''' <returns></returns>
        Public Shared Function Multiply(ByRef Vect As Vector, ByRef Scalar As Integer) As Vector
            Dim Answer As New Vector
            For i = 0 To Vect.Values.Count
                Answer.Values.Add(Scalar * Vect.Values(i))
            Next
            Return Answer
        End Function

        ''' <summary>
        ''' Square each value
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function Squ(ByRef Vect As Vector) As Vector
            Dim Answer As New Vector
            For Each item In Vect.Values
                Answer.Values.Add(item * item)
            Next
            Return Answer
        End Function

        ''' <summary>
        ''' Returns the Sum of the Squares
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function SquSum(ByRef vect As Vector) As Double
            Return vect.Sum * vect.Sum
        End Function

        Public Shared Function Subtract(ByRef Vect As Vector, ByRef NewVect As Vector)
            Dim answer As New Vector
            For i = 0 To Vect.Values.Count
                answer.Values.Add(Vect.Values(i) - NewVect.Values(i))
            Next
            Return answer
        End Function

        ''' <summary>
        ''' Adds Two vectors together
        ''' </summary>
        ''' <param name="nVector"></param>
        ''' <returns></returns>
        Public Function ADD(ByRef nVector As Vector) As Vector
            Dim Answer As New Vector

            If nVector.Values.Count > Values.Count Then

                For i = 0 To Values.Count
                    Answer.Values.Add(nVector.Values(i) + Values(i))
                Next
                'Add Excess values
                For i = Values.Count To nVector.Values.Count
                    Answer.Values.Add(nVector.Values(i))
                Next
            Else
                '
                For i = 0 To nVector.Values.Count
                    Answer.Values.Add(nVector.Values(i) + Values(i))
                Next
                'Add Excess values
                For i = nVector.Values.Count To Values.Count
                    Answer.Values.Add(nVector.Values(i))
                Next
            End If

            Return Answer
        End Function

        ''' <summary>
        ''' returns dot product of two vectors
        ''' </summary>
        ''' <param name="Vect"></param>
        ''' <returns></returns>
        Public Function DotProduct(ByRef Vect As Vector) As Double
            Dim answer As Vector = Multiply(Vect)
            Dim DotProd As Double = answer.Sum
            Return DotProd
        End Function

        ''' <summary>
        ''' Mutiplys two vectors ;
        ''' If vectors are not matching size excess values will be 0
        ''' </summary>
        ''' <param name="nVector"></param>
        ''' <returns></returns>
        Public Function Multiply(ByRef nVector As Vector) As Vector
            Dim Answer As New Vector
            If nVector.Values.Count > Values.Count Then
                For i = 0 To Values.Count
                    Answer.Values.Add(nVector.Values(i) * Values(i))
                Next
                'Add Excess values
                For i = Values.Count To nVector.Values.Count
                    Answer.Values.Add(nVector.Values(i))
                Next
            Else
                For i = 0 To nVector.Values.Count
                    Answer.Values.Add(nVector.Values(i) * Values(i))
                Next
                'Add Excess values
                For i = nVector.Values.Count To Values.Count
                    Answer.Values.Add(nVector.Values(i) * 0)
                Next
            End If
            Return Answer
        End Function

        ''' <summary>
        ''' Mutiplys Vector by Scalar ;
        ''' If vectors are not matching size excess values will be 0
        ''' </summary>
        ''' <param name="Scalar"></param>
        ''' <returns></returns>
        Public Function Multiply(ByRef Scalar As Integer) As Vector
            Dim Answer As New Vector

            For i = 0 To Values.Count
                Answer.Values.Add(Scalar * Values(i))
            Next

            Return Answer
        End Function

        ''' <summary>
        ''' Sums all numbers :
        ''' a+b+c+d... = Answer
        ''' </summary>
        ''' <returns></returns>
        Public Function Scalar_Add() As Double
            Dim total As Double = 0
            For Each item In Values
                total += item
            Next
            Return total
        End Function

        ''' <summary>
        ''' Multiplys each value by previous to give output
        ''' a*b*c*d..... = final value
        ''' </summary>
        ''' <returns></returns>
        Public Function Scalar_Multiply() As Double
            Dim total As Double = 0
            For Each item In Values
                total *= item
            Next
            Return total
        End Function

        ''' <summary>
        ''' Square each value
        ''' </summary>
        ''' <returns></returns>
        Public Function Squ() As Vector
            Dim Answer As New Vector
            For Each item In Values
                Answer.Values.Add(item * item)
            Next
            Return Answer
        End Function

        ''' <summary>
        ''' Returns the Sum of the Squares
        ''' </summary>
        ''' <returns></returns>
        Public Function SquSum() As Double
            Return Me.Sum * Me.Sum
        End Function

        Public Sub Subtract(ByRef NewVect As Vector)
            Dim answer As New Vector
            For i = 0 To Values.Count
                answer.Values.Add(Values(i) - NewVect.Values(i))
            Next
            Me.Values = answer.Values
        End Sub

        ''' <summary>
        ''' Sums the values to produce as single value
        ''' </summary>
        ''' <returns></returns>
        Public Function Sum() As Double
            Dim total As Double = 0
            For Each item In Values
                total += item
            Next
            Return total
        End Function

#End Region

    End Structure

    ''' <summary>
    ''' Detailed Word information
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure Word

#Region "Fields"

        Public Catagory As String

        Public frequecy As Integer
        ''' <summary>
        ''' Part Of Speech
        ''' </summary>

        Public POS As String

        Public Position As Integer

        Public Word As String

        Public WordType As String

#End Region

#Region "Methods"

        Public Function GetRandomPattern(ByRef Patterns As List(Of Word)) As Word
            Dim rnd = New Random()
            If Patterns.Count > 0 Then

                Return Patterns(rnd.Next(0, Patterns.Count))
            Else
                Return Nothing
            End If
        End Function

        Public Function Join(ByVal Word1 As Word, ByRef Word2 As Word) As Word
            Join = New Word
            Join.POS &= Word1.POS & " " & Word2.POS
            Join.Word &= Word1.Word & " " & Word2.Word
            Return Word1
        End Function

        Public Function ToJson() As String
            Dim Converter As New JavaScriptSerializer
            Return Converter.Serialize(Me)

        End Function

        Public Overrides Function ToString() As String
            Dim Str As String = ""
            Str &= frequecy.ToString() & vbCrLf
            Str &= POS.ToString() & vbCrLf
            Str &= Position.ToString() & vbCrLf
            Str &= Me.Word.ToString() & vbCrLf

            Return Str

        End Function

#End Region

        'Word
        'Position in sentence
        'Part of Speech Tag
        'Frequency of Word
    End Structure

    Public Structure WordVector
        Dim Freq As Integer
        Public NormalizedEncoding As Integer
        Public OneHotEncoding As Integer
        Public PositionalEncoding As Double()
        Dim PositionalEncodingVector As List(Of Double)
        Dim SequenceEncoding As Integer
        Dim Token As String

        ''' <summary>
        ''' adds positional encoding to list of word_vectors (ie encoded document)
        ''' Presumes a dimensional model of 512
        ''' </summary>
        ''' <param name="DccumentStr">Current Document</param>
        ''' <returns></returns>
        Public Shared Function AddPositionalEncoding(ByRef DccumentStr As List(Of WordVector)) As List(Of WordVector)
            ' Define the dimension of the model
            Dim d_model As Integer = 512
            ' Loop through each word in the sentence and apply positional encoding
            Dim i As Integer = 0
            For Each wrd In DccumentStr

                wrd.PositionalEncoding = CalcPositionalEncoding(i, d_model)
                i += 1
            Next
            Return DccumentStr
        End Function

        ''' <summary>
        ''' creates a list of word vectors sorted by frequency, from the text given
        ''' </summary>
        ''' <param name="Sentence"></param> document
        ''' <param name="Vocabulary"></param> based on document
        ''' <returns>vocabulary sorted in order of frequency</returns>
        Public Shared Function CreateSortedVocabulary(ByRef Sentence As String) As List(Of WordVector)
            Dim Vocabulary = WordVector.CreateVocabulary(Sentence)
            Dim NewDict As New List(Of WordVector)
            Dim Words() = Sentence.Split(" ")
            ' Count the frequency of each word
            Dim wordCounts As Dictionary(Of String, Integer) = Words.GroupBy(Function(w) w).ToDictionary(Function(g) g.Key, Function(g) g.Count())
            'Get the top ten words
            Dim TopTen As List(Of KeyValuePair(Of String, Integer)) = wordCounts.OrderByDescending(Function(w) w.Value).Take(10).ToList()

            Dim SortedDict As New List(Of WordVector)
            'Create Sorted List
            Dim Sorted As List(Of KeyValuePair(Of String, Integer)) = wordCounts.OrderByDescending(Function(w) w.Value).ToList()
            'Create Sorted Dictionary
            For Each item In Sorted

                Dim NewToken As New WordVector
                NewToken.Token = item.Key
                NewToken.SequenceEncoding = LookUpSeqEncoding(Vocabulary, item.Key)
                NewToken.Freq = item.Value
                SortedDict.Add(NewToken)

            Next

            Return SortedDict
        End Function

        ''' <summary>
        ''' Creates a unique list of words
        ''' Encodes words by their order of appearance in the text
        ''' </summary>
        ''' <param name="Sentence">document text</param>
        ''' <returns>EncodedWordlist (current vocabulary)</returns>
        Public Shared Function CreateVocabulary(ByRef Sentence As String) As List(Of WordVector)
            Dim inputString As String = "This is a sample sentence."
            If Sentence IsNot Nothing Then
                inputString = Sentence
            End If
            Dim uniqueCharacters As New List(Of String)

            Dim Dictionary As New List(Of WordVector)
            Dim Words() = Sentence.Split(" ")
            'Create unique tokens
            For Each c In Words
                If Not uniqueCharacters.Contains(c) Then
                    uniqueCharacters.Add(c)
                End If
            Next
            'Iterate through unique tokens assigning integers
            For i As Integer = 0 To uniqueCharacters.Count - 1
                'create token entry
                Dim newToken As New WordVector
                newToken.Token = uniqueCharacters(i)
                newToken.SequenceEncoding = i + 1
                'Add to vocab
                Dictionary.Add(newToken)

            Next
            Return UpdateVocabularyFrequencys(Sentence, Dictionary)
        End Function

        ''' <summary>
        ''' Creates embeddings for the sentence provided using the generated vocabulary
        ''' </summary>
        ''' <param name="Sentence"></param>
        ''' <param name="Vocabulary"></param>
        ''' <returns></returns>
        Public Shared Function EncodeWordsToVectors(ByRef Sentence As String, ByRef Vocabulary As List(Of WordVector)) As List(Of WordVector)

            Sentence = Sentence.ToLower
            If Vocabulary Is Nothing Then
                Vocabulary = CreateVocabulary(Sentence)
            End If
            Dim words() As String = Sentence.Split(" ")
            Dim Dict As New List(Of WordVector)
            For Each item In words
                Dim RetSent As New WordVector
                RetSent = GetToken(Vocabulary, item)
                Dict.Add(RetSent)
            Next
            Return NormalizeWords(Sentence, AddPositionalEncoding(Dict))
        End Function

        ''' <summary>
        ''' Decoder
        ''' </summary>
        ''' <param name="Vocabulary">Encoded Wordlist</param>
        ''' <param name="Token">desired token</param>
        ''' <returns></returns>
        Public Shared Function GetToken(ByRef Vocabulary As List(Of WordVector), ByRef Token As String) As WordVector
            For Each item In Vocabulary
                If item.Token = Token Then Return item
            Next
        End Function

        ''' <summary>
        ''' finds the frequency of this token in the sentence
        ''' </summary>
        ''' <param name="Token">token to be defined</param>
        ''' <param name="InputStr">string containing token</param>
        ''' <returns></returns>
        Public Shared Function GetTokenFrequency(ByRef Token As String, ByRef InputStr As String) As Integer
            GetTokenFrequency = 0

            If InputStr.Contains(Token) = True Then
                For Each item In WordVector.GetWordFrequencys(InputStr, " ")
                    If item.Token = Token Then
                        GetTokenFrequency = item.Freq
                    End If
                Next
            End If
        End Function

        ''' <summary>
        ''' Returns frequencys for words
        ''' </summary>
        ''' <param name="_Text"></param>
        ''' <param name="Delimiter"></param>
        ''' <returns></returns>
        Public Shared Function GetTokenFrequencys(ByVal _Text As String, ByVal Delimiter As String) As List(Of WordVector)
            Dim Words As New WordVector
            Dim ListOfWordFrequecys As New List(Of WordVector)
            Dim WordList As List(Of String) = _Text.Split(Delimiter).ToList
            Dim groups = WordList.GroupBy(Function(value) value)
            For Each grp In groups
                Words.Token = grp(0)
                Words.Freq = grp.Count
                ListOfWordFrequecys.Add(Words)
            Next
            Return ListOfWordFrequecys
        End Function

        ''' <summary>
        ''' For Legacy Functionality
        ''' </summary>
        ''' <param name="_Text"></param>
        ''' <param name="Delimiter"></param>
        ''' <returns></returns>
        Public Shared Function GetWordFrequencys(ByVal _Text As String, ByVal Delimiter As String) As List(Of WordVector)
            GetTokenFrequencys(_Text, Delimiter)
        End Function

        ''' <summary>
        ''' Decoder - used to look up a token identity using its sequence encoding
        ''' </summary>
        ''' <param name="EncodedWordlist">Encoded VectorList(vocabulary)</param>
        ''' <param name="EncodingValue">Sequence Encoding Value</param>
        ''' <returns></returns>
        Public Shared Function LookUpBySeqEncoding(ByRef EncodedWordlist As List(Of WordVector), ByRef EncodingValue As Integer) As String
            For Each item In EncodedWordlist
                If item.SequenceEncoding = EncodingValue Then Return item.Token
            Next
        End Function

        ''' <summary>
        ''' Encoder - used to look up a tokens sequence encoding, in a vocabulary
        ''' </summary>
        ''' <param name="EncodedWordlist">Encoded VectorList(vocabulary) </param>
        ''' <param name="Token">Desired Token</param>
        ''' <returns></returns>
        Public Shared Function LookUpSeqEncoding(ByRef EncodedWordlist As List(Of WordVector), ByRef Token As String) As Integer
            For Each item In EncodedWordlist
                If item.Token = Token Then Return item.SequenceEncoding
            Next
        End Function

        ''' <summary>
        ''' Adds Normalization to Vocabulary(Word-based)
        ''' </summary>
        ''' <param name="Sentence">Doc</param>
        ''' <param name="dict">Vocabulary</param>
        ''' <returns></returns>
        Public Shared Function NormalizeWords(ByRef Sentence As String, ByRef dict As List(Of WordVector)) As List(Of WordVector)
            Dim Count = CountWords(Sentence)
            For Each item In dict
                item.NormalizedEncoding = Count / item.Freq
            Next
            Return dict
        End Function

        ''' <summary>
        ''' Encodes a list of word-vector by a list of strings
        ''' If a token is found in the list it is encoded with a binary 1 if false then 0
        ''' This is useful for categorizing and adding context to the word vector
        ''' </summary>
        ''' <param name="WordVectorList">list of tokens to be encoded (categorized)</param>
        ''' <param name="Vocabulary">Categorical List, Such as a list of positive sentiment</param>
        ''' <returns></returns>
        Public Shared Function OneShotEncoding(ByVal WordVectorList As List(Of WordVector),
                                    ByRef Vocabulary As List(Of String)) As List(Of WordVector)
            Dim EncodedList As New List(Of WordVector)
            For Each item In WordVectorList
                Dim Found As Boolean = False
                For Each RefItem In Vocabulary
                    If item.Token = RefItem Then
                        Found = True
                    Else

                    End If
                Next
                If Found = True Then
                    Dim newWordvector As WordVector = item
                    newWordvector.OneHotEncoding = True
                End If
                EncodedList.Add(item)
            Next
            Return EncodedList
        End Function

        ''' <summary>
        ''' Creates a List of Bigram WordVectors Based on the text
        ''' to create the vocabulary file use @ProduceBigramVocabulary
        ''' </summary>
        ''' <param name="Sentence"></param>
        ''' <returns>Encoded list of bigrams and vectors (vocabulary)with frequencies </returns>
        Public Shared Function ProduceBigramDocument(ByRef sentence As String) As List(Of WordVector)

            ' Convert sentence to lowercase and split into words
            Dim words As String() = sentence.ToLower().Split()
            Dim GeneratedBigramsList As New List(Of String)
            Dim bigrams As New Dictionary(Of String, Integer)
            'We start at the first word And go up to the second-to-last word
            For i As Integer = 0 To words.Length - 2
                Dim bigram As String = words(i) & " " & words(i + 1)
                'We check If the bigrams dictionary already contains the bigram.
                'If it does, we increment its frequency by 1.
                'If it doesn't, we add it to the dictionary with a frequency of 1.
                GeneratedBigramsList.Add(bigram)
                If bigrams.ContainsKey(bigram) Then
                    bigrams(bigram) += 1
                Else
                    bigrams.Add(bigram, 1)
                End If
            Next

            'we Loop through the bigrams dictionary(of frequncies) And encode a integer to the bi-gram

            Dim bigramVocab As New List(Of WordVector)
            Dim a As Integer = 1
            For Each kvp As KeyValuePair(Of String, Integer) In bigrams
                Dim newvect As New WordVector
                newvect.Token = kvp.Key
                newvect.Freq = kvp.Value

                bigramVocab.Add(newvect)
            Next
            'create a list from the generated bigrams and
            ''add frequecies from vocabulary of frequecies
            Dim nVocab As New List(Of WordVector)
            Dim z As Integer = 0
            For Each item In GeneratedBigramsList
                'create final token
                Dim NewToken As New WordVector
                NewToken.Token = item
                'add current position in document
                NewToken.SequenceEncoding = GeneratedBigramsList(z)
                'add frequency
                For Each Lookupitem In bigramVocab
                    If item = Lookupitem.Token Then
                        NewToken.Freq = Lookupitem.Freq
                    Else
                    End If
                Next
                'add token
                nVocab.Add(NewToken)
                'update current index
                z += 1
            Next

            'Return bigram document with sequence and frequencys
            Return nVocab
        End Function

        ''' <summary>
        ''' Creates a Vocabulary of unique bigrams from sentence with frequencies adds a sequence vector based on
        ''' its appearence in the text, if item is repeated at multiple locations it is not reflected here
        ''' </summary>
        ''' <param name="Sentence"></param>
        ''' <returns>Encoded list of unique bigrams and vectors (vocabulary)with frequencies </returns>
        Public Shared Function ProduceBigramVocabulary(ByRef sentence As String) As List(Of WordVector)

            ' Convert sentence to lowercase and split into words
            Dim words As String() = sentence.ToLower().Split()
            Dim GeneratedBigramsList As New List(Of String)
            Dim bigrams As New Dictionary(Of String, Integer)
            'We start at the first word And go up to the second-to-last word
            For i As Integer = 0 To words.Length - 2
                Dim bigram As String = words(i) & " " & words(i + 1)
                'We check If the bigrams dictionary already contains the bigram.
                'If it does, we increment its frequency by 1.
                'If it doesn't, we add it to the dictionary with a frequency of 1.
                GeneratedBigramsList.Add(bigram)
                If bigrams.ContainsKey(bigram) Then
                    bigrams(bigram) += 1
                Else
                    bigrams.Add(bigram, 1)
                End If
            Next

            'we Loop through the bigrams dictionary(of frequncies) And encode a integer to the bi-gram

            Dim bigramVocab As New List(Of WordVector)
            Dim a As Integer = 0
            For Each kvp As KeyValuePair(Of String, Integer) In bigrams
                Dim newvect As New WordVector
                newvect.Token = kvp.Key
                newvect.Freq = kvp.Value
                newvect.SequenceEncoding = a + 1
                bigramVocab.Add(newvect)
            Next

            'Return bigram document with sequence and frequencys
            Return bigramVocab
        End Function

        ''' <summary>
        ''' Adds Frequencies to a sequentially encoded word-vector list
        ''' </summary>
        ''' <param name="Sentence">current document</param>
        ''' <param name="EncodedWordlist">Current Vocabulary</param>
        ''' <returns>an encoded word-Vector list with Frequencys attached</returns>
        Public Shared Function UpdateVocabularyFrequencys(ByRef Sentence As String, ByVal EncodedWordlist As List(Of WordVector)) As List(Of WordVector)

            Dim NewDict As New List(Of WordVector)
            Dim Words() = Sentence.Split(" ")
            ' Count the frequency of each word
            Dim wordCounts As Dictionary(Of String, Integer) = Words.GroupBy(Function(w) w).ToDictionary(Function(g) g.Key, Function(g) g.Count())
            'Get the top ten words
            Dim TopTen As List(Of KeyValuePair(Of String, Integer)) = wordCounts.OrderByDescending(Function(w) w.Value).Take(10).ToList()

            'Create Standard Dictionary
            For Each EncodedItem In EncodedWordlist
                For Each item In wordCounts
                    If EncodedItem.Token = item.Key Then
                        Dim NewToken As New WordVector
                        NewToken = EncodedItem
                        NewToken.Freq = item.Value
                        NewDict.Add(NewToken)
                    End If
                Next
            Next

            Return NewDict
        End Function

        ''' <summary>
        ''' Outputs Structure to Jason(JavaScriptSerializer)
        ''' </summary>
        ''' <returns></returns>
        Public Function ToJson() As String
            Dim Converter As New JavaScriptSerializer
            Return Converter.Serialize(Me)
        End Function

#Region "Positional Encoding"

        Private Shared Function CalcPositionalEncoding(ByVal position As Integer, ByVal d_model As Integer) As Double()
            ' Create an empty array to store the encoding
            Dim encoding(d_model - 1) As Double

            ' Loop through each dimension of the model and calculate the encoding value
            For i As Integer = 0 To d_model - 1
                If i Mod 2 = 0 Then
                    encoding(i) = Math.Sin(position / (10000 ^ (i / d_model)))
                Else
                    encoding(i) = Math.Cos(position / (10000 ^ ((i - 1) / d_model)))
                End If
            Next

            ' Return the encoding array
            Return encoding
        End Function

#End Region

#Region "Normalization"

        ''' <summary>
        ''' returns number of Chars in text
        ''' </summary>
        ''' <param name="Sentence">Document</param>
        ''' <returns>number of Chars</returns>
        Private Shared Function CountChars(ByRef Sentence As String) As Integer
            Dim uniqueCharacters As New List(Of String)
            For Each c As Char In Sentence
                uniqueCharacters.Add(c.ToString)
            Next
            Return uniqueCharacters.Count
        End Function

        ''' <summary>
        ''' Returns number of words in text
        ''' </summary>
        ''' <param name="Sentence">Document</param>
        ''' <returns>number of words</returns>
        Private Shared Function CountWords(ByRef Sentence As String) As Integer
            Dim Words() = Sentence.Split(" ")
            Return Words.Count
        End Function

#End Region

    End Structure

    Public Class Actions

#Region "Enums"

        ''' <summary>
        ''' Types of Actions that can be executed
        ''' </summary>
        Public Enum ActionType

            ''' <summary>
            ''' Fills current slot
            ''' </summary>
            FillSlot

            ''' <summary>
            ''' Confirm Slots gathered
            ''' </summary>
            ConfirmSlot

            ''' <summary>
            ''' Responds to user with statement
            ''' </summary>
            Respond

            ''' <summary>
            ''' Opens Website
            ''' </summary>
            OpenWebsite

            ''' <summary>
            ''' Executes Script
            ''' </summary>
            ExecuteScript

            ''' <summary>
            ''' Ends current dialog
            ''' </summary>
            EndDialog

            ''' <summary>
            ''' Begins Dialog path (User Story)
            ''' </summary>
            BeginDialog

        End Enum

#End Region

#Region "Classes"

        ''' <summary>
        ''' used for Executing actions generated by UserRequests / Intents
        ''' </summary>
        Public MustInherit Class Action

#Region "Constructors"

            ''' <summary>
            ''' Minimum Requirement Action Type,
            ''' Class should be inherited and custom constructor created
            ''' </summary>
            ''' <param name="Action">Action Type may determine execution strategy</param>
            Public Sub New(ByRef Action As ActionType)
                Me.Action = Action
            End Sub

#End Region

#Region "Properties"

            ''' <summary>
            ''' Type of Action
            ''' </summary>
            ''' <returns></returns>
            Property Action As ActionType

            MustOverride Property Name As String

            ''' <summary>
            ''' Response Should be passed in from calling Intent
            ''' </summary>
            ''' <returns></returns>
            MustOverride Property Response As String

#End Region

#Region "Methods"

            ''' <summary>
            ''' Called by manager to execute action
            ''' </summary>
            MustOverride Sub Execute()

#End Region

        End Class

        Public Class ActionManager

#Region "Fields"

            Private Shared mActions As ICollection(Of Action)

#End Region

#Region "Properties"

            ''' <summary>
            ''' used to hold the collection of Actions
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Shared ReadOnly Property Actions As ICollection(Of Action)
                Get
                    Return mActions
                End Get
            End Property

#End Region

#Region "Methods"

            Public Shared Function ExecuteAction(ByRef iAction As Action) As Boolean
                Dim responded = False
                iAction.Execute()
                responded = True
                Return responded
            End Function

            Public Shared Function ExecuteActions() As Boolean
                Dim responded = False
                mActions = LoadActions(Application.StartupPath & "\Actions\")

                If mActions IsNot Nothing Then
                    For Each Action As Action In mActions
                        Action.Execute()
                        responded = True
                    Next
                End If
                Return responded
            End Function

            ''' <summary>
            ''' Loads the iActions into the class
            ''' </summary>
            ''' <param name="path">Pathname directory which contains files of type</param>
            ''' <returns>Collection of Actions</returns>
            ''' <remarks></remarks>
            Public Shared Function LoadActions(path As String) As ICollection(Of Action)
                On Error Resume Next
                Dim dllFileNames As String()
                If IO.Directory.Exists(path) Then
                    dllFileNames = IO.Directory.GetFiles(path, "*.dll")
                    Dim assemblies As ICollection(Of Reflection.Assembly) = New List(Of Reflection.Assembly)(dllFileNames.Length)
                    For Each dllFile As String In dllFileNames
                        Dim an As Reflection.AssemblyName = Reflection.AssemblyName.GetAssemblyName(dllFile)
                        Dim assembly As Reflection.Assembly = Reflection.Assembly.Load(an)
                        assemblies.Add(assembly)
                    Next
                    Dim ActionType As Type = GetType(Action)
                    Dim ActionTypes As ICollection(Of Type) = New List(Of Type)
                    For Each assembly As Reflection.Assembly In assemblies
                        If assembly <> Nothing Then
                            Dim types As Type() = assembly.GetTypes()
                            For Each type As Type In types
                                If type.IsInterface Or type.IsAbstract Then
                                    Continue For
                                Else
                                    If type.GetInterfaces(ActionType.FullName) <> Nothing Then
                                        ActionTypes.Add(type)
                                    End If
                                End If
                            Next
                        End If
                    Next
                    Dim iActions As ICollection(Of Action) = New List(Of Action)(ActionTypes.Count)
                    For Each type As Type In ActionTypes
                        Dim iAction As Action = Activator.CreateInstance(type)
                        iActions.Add(iAction)
                    Next
                    Return iActions
                End If
                Return Nothing
            End Function

#End Region

        End Class

        Public Class DefaultActions

#Region "Classes"

            Public Class DialogControl

#Region "Classes"

                Public Class BeginDialog
                    Inherits Action

#Region "Constructors"

                    Public Sub New(ByRef Action As ActionType)
                        MyBase.New(ActionType.BeginDialog)
                    End Sub

#End Region

#Region "Properties"

                    Public Overrides Property Name As String
                        Get
                            Throw New NotImplementedException()
                        End Get
                        Set(value As String)
                            Throw New NotImplementedException()
                        End Set
                    End Property

                    Public Overrides Property Response As String
                        Get
                            Throw New NotImplementedException()
                        End Get
                        Set(value As String)
                            Throw New NotImplementedException()
                        End Set
                    End Property

#End Region

#Region "Methods"

                    Public Overrides Sub Execute()
                        Throw New NotImplementedException()
                    End Sub

#End Region

                End Class

                Public Class EndDialog
                    Inherits Action

#Region "Constructors"

                    Public Sub New(ByRef Action As ActionType)
                        MyBase.New(ActionType.EndDialog)
                    End Sub

#End Region

#Region "Properties"

                    Public Overrides Property Name As String
                        Get
                            Throw New NotImplementedException()
                        End Get
                        Set(value As String)
                            Throw New NotImplementedException()
                        End Set
                    End Property

                    Public Overrides Property Response As String
                        Get
                            Throw New NotImplementedException()
                        End Get
                        Set(value As String)
                            Throw New NotImplementedException()
                        End Set
                    End Property

#End Region

#Region "Methods"

                    Public Overrides Sub Execute()
                        Throw New NotImplementedException()
                    End Sub

#End Region

                End Class

#End Region

            End Class

            Public Class Responses

#Region "Classes"

                ''' <summary>
                ''' Make simple Response
                ''' </summary>
                Public Class Respond
                    Inherits Action

#Region "Fields"

                    Private iName As String

                    Private iResponse As String

#End Region

#Region "Constructors"

                    ''' <summary>
                    ''' Opens Website on command
                    ''' </summary>
                    ''' <param name="Name">URL Required</param>
                    ''' <param name="Response">Response Passed in</param>
                    ''' <param name="Action">Type of Action Required</param>
                    Public Sub New(ByRef Name As String, ByRef Response As String, ByRef Action As ActionType)
                        MyBase.New(ActionType.Respond)
                        Me.Name = Name
                        Me.Response = Response

                    End Sub

#End Region

#Region "Properties"

                    Public Overrides Property Name As String
                        Get
                            Return iName
                        End Get
                        Set(value As String)
                            iName = value
                        End Set
                    End Property

                    Public Overrides Property Response As String
                        Get
                            Return iResponse
                        End Get
                        Set(value As String)
                            iResponse = value
                        End Set
                    End Property

#End Region

#Region "Methods"

                    Public Overrides Sub Execute()

                        Throw New NotImplementedException()
                    End Sub

#End Region

                End Class

#End Region

            End Class

            Public Class Scripts

#Region "Classes"

                ''' <summary>
                ''' ACTION: OpensWebsite Provided
                ''' </summary>
                Public Class OpenWebsite
                    Inherits Action

#Region "Fields"

                    Private iName As String

                    Private iResponse As String

                    ''' <summary>
                    ''' URL to be discovered
                    ''' </summary>
                    Private Url As String = ""

#End Region

#Region "Constructors"

                    ''' <summary>
                    ''' Opens Website on command
                    ''' </summary>
                    ''' <param name="Url">URL Required</param>
                    ''' <param name="Response">Response Passed in</param>
                    ''' <param name="Action">Type of Action Required</param>
                    Public Sub New(ByRef Url As String, ByRef Response As String, ByRef Action As ActionType)
                        MyBase.New(ActionType.OpenWebsite)
                        Me.Name = "Open Website"
                        Me.Response = Response
                        Me.Url = Url
                    End Sub

#End Region

#Region "Properties"

                    Public Overrides Property Name As String
                        Get
                            Return iName
                        End Get
                        Set(value As String)
                            iName = value
                        End Set
                    End Property

                    Public Overrides Property Response As String
                        Get
                            Return iResponse
                        End Get
                        Set(value As String)
                            iResponse = value
                        End Set
                    End Property

#End Region

#Region "Methods"

                    Public Overrides Sub Execute()

                        Throw New NotImplementedException()
                    End Sub

#End Region

                End Class

#End Region

            End Class

            Public Class SlotControl

#Region "Classes"

                Public Class ConfirmSlot
                    Inherits Action

#Region "Constructors"

                    Public Sub New(ByRef Action As ActionType)
                        MyBase.New(ActionType.ConfirmSlot)
                    End Sub

#End Region

#Region "Properties"

                    Public Overrides Property Name As String
                        Get
                            Throw New NotImplementedException()
                        End Get
                        Set(value As String)
                            Throw New NotImplementedException()
                        End Set
                    End Property

                    Public Overrides Property Response As String
                        Get
                            Throw New NotImplementedException()
                        End Get
                        Set(value As String)
                            Throw New NotImplementedException()
                        End Set
                    End Property

#End Region

#Region "Methods"

                    Public Overrides Sub Execute()
                        Throw New NotImplementedException()
                    End Sub

#End Region

                End Class

                Public Class FillSlot
                    Inherits Action

#Region "Constructors"

                    Public Sub New(ByRef Action As ActionType)
                        MyBase.New(ActionType.FillSlot)
                    End Sub

#End Region

#Region "Properties"

                    Public Overrides Property Name As String
                        Get
                            Throw New NotImplementedException()
                        End Get
                        Set(value As String)
                            Throw New NotImplementedException()
                        End Set
                    End Property

                    Public Overrides Property Response As String
                        Get
                            Throw New NotImplementedException()
                        End Get
                        Set(value As String)
                            Throw New NotImplementedException()
                        End Set
                    End Property

#End Region

#Region "Methods"

                    Public Overrides Sub Execute()
                        Throw New NotImplementedException()
                    End Sub

#End Region

                End Class

#End Region

            End Class

#End Region

        End Class

#End Region

    End Class

End Namespace