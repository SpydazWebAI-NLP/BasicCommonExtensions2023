Imports System.Reflection
Imports System.Windows.Forms
Imports Microsoft.Win32

Namespace AI_SDK_EXTENSIONS

    Namespace Types

        Public Class TypeLibraries
            Inherits CollectionBase

            Public Function FindByGUIDAndVersion(ByVal GUID As String, ByVal Version As String) As TypeLibrary
                If Me.InnerList.Count = 0 Then Me.GetAllFromRegistry() ' if no data exists, populate

                For Each t As TypeLibrary In Me.InnerList
                    If GUID = t.GUID And Version = t.Version Then Return t
                Next

                Return Nothing
            End Function

            Public Sub GetAllFromRegistry()

                Const TLB_ROOT As String = "TypeLib"
                Dim rkey As RegistryKey = Registry.ClassesRoot

                Dim regTypeLibRoot As RegistryKey = rkey.OpenSubKey(TLB_ROOT)

                For Each strTypeLib As String In regTypeLibRoot.GetSubKeyNames
                    Dim regTypeLib As RegistryKey = regTypeLibRoot.OpenSubKey(strTypeLib)

                    ' In strVersions we should have list of typelib versions
                    For Each strVersion As String In regTypeLib.GetSubKeyNames
                        Dim regVersion As RegistryKey = regTypeLib.OpenSubKey(strVersion)

                        ' Now we can create a new TypeLibrary object since under this value there should
                        ' be only descriptive info of Type Library
                        Dim typeLib As New TypeLibrary

                        ' Set values we know by now
                        typeLib.GUID = strTypeLib
                        typeLib.Version = strVersion

                        ' Here we should have at least one value for descriptive name of TypeLib - not
                        ' always though...
                        For Each strValueName As String In regVersion.GetValueNames
                            Select Case strValueName
                                Case "PrimaryInteropAssemblyName" ' We got name of PrimaryInteropAssemblyName
                                    typeLib.PrimaryInteropAssemblyName = regVersion.GetValue(strValueName)
                                Case Else ' We got descriptive name of TypeLib
                                    typeLib.FullName = regVersion.GetValue(strValueName)
                            End Select
                        Next

                        ' Iterate through subkeys
                        For Each strSubKey As String In regVersion.GetSubKeyNames

                            Dim regSubKey As RegistryKey = regVersion.OpenSubKey(strSubKey)

                            Select Case strSubKey
                                Case "FLAGS" : typeLib.Flags = regSubKey.GetValue("")
                                Case "HELPDIR" : typeLib.HelpDir = regSubKey.GetValue("")
                                Case Else ' We found propably number with subkey of "win32","win16" or even "win64"

                                    Dim typeLibFiles As New TypeLibraryFiles

                                    For Each strPlatform As String In regSubKey.GetSubKeyNames
                                        Dim tlbFile As New TypeLibraryFile
                                        tlbFile.Number = strSubKey
                                        Select Case UCase(Trim(strPlatform))
                                            Case "WIN16" : tlbFile.Platform = TypeLibraryFile.TypeLibraryPlatformType.Win16
                                            Case "WIN32" : tlbFile.Platform = TypeLibraryFile.TypeLibraryPlatformType.Win32
                                            Case "WIN64" : tlbFile.Platform = TypeLibraryFile.TypeLibraryPlatformType.Win64
                                            Case Else : tlbFile.Platform = TypeLibraryFile.TypeLibraryPlatformType.Unknown
                                        End Select

                                        tlbFile.File = regSubKey.OpenSubKey(strPlatform).GetValue("") ' Get name of file as a default value

                                        typeLibFiles.Add(tlbFile) ' Add file to collection
                                    Next

                                    typeLib.Files = typeLibFiles ' Set files collection to typelib info
                            End Select
                        Next

                        MyBase.List.Add(typeLib)
                    Next
                Next
            End Sub

        End Class

        Public Class TypeLibrary

            Private mFiles As TypeLibraryFiles
            Private mFlags As String
            Private mFullName As String
            Private mGUID As String
            Private mHelpDir As String
            Private mPrimaryInteropAssemblyName As String
            Private mVersion As String

            Public Property Files() As TypeLibraryFiles
                Get
                    Return mFiles
                End Get
                Set(ByVal value As TypeLibraryFiles)
                    mFiles = value
                End Set
            End Property

            Public Property Flags() As String
                Get
                    Return mFlags
                End Get
                Set(ByVal value As String)
                    mFlags = value
                End Set
            End Property

            Public Property FullName() As String
                Get
                    Return mFullName
                End Get
                Set(ByVal value As String)
                    mFullName = value
                End Set
            End Property

            Public Property GUID() As String
                Get
                    Return mGUID
                End Get
                Set(ByVal value As String)
                    mGUID = value
                End Set
            End Property

            Public Property HelpDir() As String
                Get
                    Return mHelpDir
                End Get
                Set(ByVal value As String)
                    mHelpDir = value
                End Set
            End Property

            Public Property PrimaryInteropAssemblyName() As String
                Get
                    Return mPrimaryInteropAssemblyName
                End Get
                Set(ByVal value As String)
                    mPrimaryInteropAssemblyName = value
                End Set
            End Property

            Public Property Version() As String
                Get
                    Return mVersion
                End Get
                Set(ByVal value As String)
                    mVersion = value
                End Set
            End Property

        End Class

        Public Class TypeLibraryFile

            Private mFile As String

            Private mNumber As String

            Private mPlatform As TypeLibraryPlatformType

            Public Sub New()

            End Sub

            Public Sub New(ByVal Platform As TypeLibraryPlatformType, ByVal File As String, ByVal Number As String)
                Me.Platform = Platform
                Me.File = File
                Me.Number = Number
            End Sub

            Public Enum TypeLibraryPlatformType As Integer
                Unknown = 0
                Win16 = 1
                Win32 = 2
                Win64 = 3 ' Don't know is this even possible with COM Classes ?
            End Enum

            Public Property File() As String
                Get
                    Return mFile
                End Get
                Set(ByVal value As String)
                    mFile = value
                End Set
            End Property

            Public Property Number() As String
                Get
                    Return mNumber
                End Get
                Set(ByVal value As String)
                    mNumber = value
                End Set
            End Property

            Public Property Platform() As TypeLibraryPlatformType
                Get
                    Return mPlatform
                End Get
                Set(ByVal value As TypeLibraryPlatformType)
                    mPlatform = value
                End Set
            End Property

        End Class

        Public Class TypeLibraryFiles
            Inherits ArrayList
        End Class

        ''' <summary>
        ''' ReadsInfo from Class Type objects
        ''' </summary>
        ''' <remarks></remarks>
        Public Class TypeReader

            ''' <summary>
            ''' GetType(ClassName(Ref))
            ''' </summary>
            ''' <remarks></remarks>
            Public ObjClassLst As New List(Of Type)

            Private Shared mMasterList As New List(Of String)

            Public Sub New()

            End Sub

            ''' <summary>
            ''' Gets list of names from objlst.
            ''' </summary>
            ''' <value></value>
            ''' <remarks></remarks>
            Public ReadOnly Property MasterList As List(Of String)
                Get
                    Return mMasterList
                End Get

            End Property

            ' Create a class having two public methods and one protected method.
            Public Shared Function AddSyntax(ByRef Syntax As List(Of String), ByRef MasterSyntax As List(Of String)) As List(Of String)
                Dim NewSyn As New List(Of String)

                For Each term As String In Syntax
                    NewSyn.Add(term)
                Next
                For Each term As String In MasterSyntax
                    NewSyn.Add(term)
                Next
                Return NewSyn
            End Function

            ''' <summary>
            ''' In this function Changes eed to be made to refference in Internal Types
            ''' </summary>
            ''' <returns></returns>
            Public Shared Function Addtypes() As List(Of Type)
                ' mObjClassLst.Add(GetType(NewType))
                Dim mObjClassLst As New List(Of Type) From {
                GetType(TypeReader)
            }
                Return mObjClassLst
            End Function

            ''' <summary>
            ''' GetType(ClassName(Ref))
            ''' </summary>
            ''' <param name="MyType"></param>
            ''' <remarks></remarks>
            Public Shared Sub ClassToConsole(ByRef MyType As Type)
                Dim Str As String = ""

                ' Get the public methods.
                Dim myArrayMethodInfo As MethodInfo() = MyType.GetMethods((BindingFlags.Public Or BindingFlags.Instance Or BindingFlags.DeclaredOnly))
                Console.WriteLine((ControlChars.Cr + "The number of public methods is " & myArrayMethodInfo.Length.ToString() & "." & vbNewLine))
                ' Display all the public methods.
                DisplayMethodInfo(myArrayMethodInfo)
                ' Get the nonpublic methods.
                Dim myArrayMethodInfo1 As MethodInfo() = MyType.GetMethods((BindingFlags.NonPublic Or BindingFlags.Instance Or BindingFlags.DeclaredOnly))
                Console.WriteLine((ControlChars.Cr + "The number of protected methods is " & myArrayMethodInfo1.Length.ToString() & "."))
                ' Display all the nonpublic methods.
                DisplayMethodInfo(myArrayMethodInfo1)
            End Sub

            'Main
            ''' <summary>
            ''' GetType(ClassName(Ref))
            ''' </summary>
            ''' <param name="MyType"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Shared Function ClassToString(ByRef MyType As Type) As String
                Dim Str As String = ""

                ' Get the public methods.
                Dim myArrayMethodInfo As MethodInfo() = MyType.GetMethods((BindingFlags.Public Or BindingFlags.Instance Or BindingFlags.DeclaredOnly))
                Str += ControlChars.Cr + "The number of public methods is " & myArrayMethodInfo.Length.ToString() & "." & vbNewLine
                ' Display all the public methods.
                Dim i As Integer
                For i = 0 To myArrayMethodInfo.Length - 1
                    Dim myMethodInfo As MethodInfo = CType(myArrayMethodInfo(i), MethodInfo)
                    Str += ControlChars.Cr + "The name of the method is " & myMethodInfo.Name & "." & vbNewLine
                Next i
                ' Get the nonpublic methods.
                Dim myArrayMethodInfo1 As MethodInfo() = MyType.GetMethods((BindingFlags.NonPublic Or BindingFlags.Instance Or BindingFlags.DeclaredOnly))
                Str += ControlChars.Cr + "The number of protected methods is " & myArrayMethodInfo1.Length.ToString() & "." & vbNewLine
                ' Display all the nonpublic methods.
                Dim j As Integer
                For j = 0 To myArrayMethodInfo1.Length - 1
                    Dim myMethodInfo As MethodInfo = CType(myArrayMethodInfo(i), MethodInfo)
                    Str += ControlChars.Cr + "The name of the method is " & myMethodInfo.Name & "." & vbNewLine
                Next j
                Return Str
            End Function

            ''' <summary>
            ''' Returns Grammar for Tools Namespace
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Shared Function GetInternalSyntax() As List(Of String)
                Dim Mlst As New List(Of String)
                Dim Lst As New List(Of String)
                Dim mObjClassLst As List(Of Type) = Addtypes()
                'Public
                For Each Clss As Type In mObjClassLst
                    Lst = GetPublic(Clss)
                    For Each Detail As String In Lst
                        Mlst.Add(LCase(Detail))
                    Next
                Next
                'Private
                For Each Clss As Type In mObjClassLst
                    Lst = GetPrivate(Clss)
                    For Each Detail As String In Lst
                        Mlst.Add(LCase(Detail))
                    Next
                Next
                Return Mlst
            End Function

            Public Shared Function GetListNoSyntax() As List(Of String)
                Dim Types As List(Of MethodItem) = TypeReader.GetMethodsList(TypeReader.Addtypes)
                Dim Lst As New List(Of String)
                For Each mtype As MethodItem In Types
                    If Lst.Contains(mtype.TypeName) Then
                    Else
                        Lst.Add(mtype.TypeName)
                    End If
                Next
                Dim StrLst As New List(Of String)
                For Each str As String In Lst
                    For Each mytype As TypeReader.MethodItem In Types
                        If mytype.TypeName = str Then

                            StrLst.Add(str & "." & mytype.TypeMethod)

                        End If
                    Next
                Next
                Return StrLst
            End Function

            Public Shared Function GetMethodsList(ByRef Types As List(Of Type)) As List(Of MethodItem)
                Dim Lst As New List(Of MethodItem)
                'Dim Types As List(Of Type) = TypeReader.Addtypes
                For Each MyType As Type In Types
                    Dim NewItem As New MethodItem
                    ' Get the public methods.
                    Dim myArrayMethodInfo As MethodInfo() = MyType.GetMethods((BindingFlags.Public Or
                                                                      BindingFlags.Instance Or
                                                                      BindingFlags.DeclaredOnly Or
                                                                    BindingFlags.Static))

                    Dim i As Integer
                    For i = 0 To myArrayMethodInfo.Length - 1
                        Dim myMethodInfo As MethodInfo = myArrayMethodInfo(i)
                        NewItem.TypeName = MyType.FullName
                        NewItem.TypeMethod = myMethodInfo.Name
                        NewItem.TypeMethodInfo = myMethodInfo
                        Lst.Add(NewItem)
                    Next i
                Next
                Return Lst
            End Function

            ''' <summary>
            ''' GetType(ClassName(Ref))
            ''' </summary>
            ''' <param name="MyType"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Shared Function GetPrivate(ByRef MyType As Type) As List(Of String)
                Dim ClassGrammar As New List(Of String)

                ' Get the public methods.
                Dim myArrayMethodInfo1 As MethodInfo() = MyType.GetMethods((BindingFlags.NonPublic Or BindingFlags.Instance Or BindingFlags.DeclaredOnly))
                ClassGrammar = GetMethods(myArrayMethodInfo1)
                Return ClassGrammar
            End Function

            'Main
            ''' <summary>
            ''' GetType(ClassName(Ref))
            ''' </summary>
            ''' <param name="MyType"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Shared Function GetPublic(ByRef MyType As Type) As List(Of String)
                Dim ClassGrammar As New List(Of String)
                ' Get the public methods.
                Dim myArrayMethodInfo As MethodInfo() = MyType.GetMethods((BindingFlags.Public Or BindingFlags.Instance Or BindingFlags.DeclaredOnly Or BindingFlags.Static))
                ClassGrammar = GetMethods(myArrayMethodInfo)
                Return ClassGrammar
            End Function

            Public Shared Sub UpdateTreeViewControl(ByRef TrContol As TreeView, Types As List(Of Type))

                Dim nTypes As List(Of MethodItem) = TypeReader.GetMethodsList(Types)

                Dim Lst As New List(Of String)
                For Each mtype As MethodItem In nTypes

                    If Lst.Contains(mtype.TypeName) Then
                    Else
                        Lst.Add(mtype.TypeName)
                    End If
                Next
                For Each str As String In Lst
                    Dim node As New System.Windows.Forms.TreeNode
                    node.Text = str
                    For Each mytype As TypeReader.MethodItem In nTypes
                        If mytype.TypeName = str Then
                            node.Nodes.Add("<" & mytype.TypeMethod & ">" & " Syntax:  " & vbNewLine & mytype.TypeMethodInfo.ToString)

                        End If
                    Next
                    TrContol.Nodes.Add(node)
                Next
            End Sub

            Public Shared Sub UpdateTreeViewControl(ByRef TrContol As TreeView)

                Dim Types As List(Of MethodItem) = TypeReader.GetMethodsList(TypeReader.Addtypes)

                Dim Lst As New List(Of String)
                For Each mtype As MethodItem In Types

                    If Lst.Contains(mtype.TypeName) Then
                    Else
                        Lst.Add(mtype.TypeName)
                    End If
                Next
                For Each str As String In Lst
                    Dim node As New System.Windows.Forms.TreeNode
                    node.Text = str
                    For Each mytype As TypeReader.MethodItem In Types
                        If mytype.TypeName = str Then
                            node.Nodes.Add("<" & mytype.TypeMethod & ">" & " Syntax:  " & vbNewLine & mytype.TypeMethodInfo.ToString)

                        End If
                    Next
                    TrContol.Nodes.Add(node)
                Next
            End Sub

            Public Shared Sub UpdateTreeViewControlNoSyntax(ByRef TrContol As TreeView)

                Dim Types As List(Of MethodItem) = TypeReader.GetMethodsList(TypeReader.Addtypes)

                Dim Lst As New List(Of String)
                For Each mtype As MethodItem In Types

                    If Lst.Contains(mtype.TypeName) Then
                    Else
                        Lst.Add(mtype.TypeName)
                    End If
                Next
                For Each str As String In Lst
                    Dim node As New System.Windows.Forms.TreeNode
                    node.Text = str
                    For Each mytype As TypeReader.MethodItem In Types
                        If mytype.TypeName = str Then
                            node.Nodes.Add(mytype.TypeMethod)

                        End If
                    Next
                    TrContol.Nodes.Add(node)
                Next
            End Sub

            Public Shared Sub UpdateTreeViewControlWithSyntax(ByRef TrContol As TreeView)

                Dim Types As List(Of MethodItem) = TypeReader.GetMethodsList(TypeReader.Addtypes)

                Dim Lst As New List(Of String)
                For Each mtype As MethodItem In Types

                    If Lst.Contains(mtype.TypeName) Then
                    Else
                        Lst.Add(mtype.TypeName)
                    End If
                Next
                For Each str As String In Lst
                    Dim node As New System.Windows.Forms.TreeNode
                    node.Text = str
                    For Each mytype As TypeReader.MethodItem In Types
                        If mytype.TypeName = str Then
                            node.Nodes.Add(" Syntax:  " & vbNewLine & mytype.TypeMethodInfo.ToString & " #")
                        End If
                    Next
                    TrContol.Nodes.Add(node)
                Next
            End Sub

            Public Function Addtypes(ByRef mType As Type) As List(Of Type)
                Dim mObjClassLst As New List(Of Type) From {
                GetType(Type)
            }
                Return mObjClassLst
            End Function

            ''' <summary>
            ''' Build list with objClassLst
            ''' </summary>
            ''' <remarks></remarks>
            Public Sub BuildMasterList()
                Dim Lst As New List(Of String)
                For Each Clss As Type In ObjClassLst
                    Lst = GetPublic(Clss)
                    For Each Detail As String In Lst
                        mMasterList.Add(Detail)
                    Next
                Next
            End Sub

            Private Shared Sub DisplayMethodInfo(ByVal myArrayMethodInfo() As MethodInfo)
                ' Display information for all methods.
                Dim i As Integer
                For i = 0 To myArrayMethodInfo.Length - 1
                    Dim myMethodInfo As MethodInfo = CType(myArrayMethodInfo(i), MethodInfo)
                    Console.WriteLine((ControlChars.Cr + "The name of the method is " & myMethodInfo.Name & "."))
                Next i
            End Sub

            'DisplayMethodInfo
            ''' <summary>
            ''' GetType(ClassName(Ref))
            ''' </summary>
            ''' <param name="myArrayMethodInfo"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Private Shared Function GetMethods(ByVal myArrayMethodInfo() As MethodInfo) As List(Of String)
                Dim Lst As New List(Of String)
                ' Display information for all methods.
                Dim i As Integer
                For i = 0 To myArrayMethodInfo.Length - 1
                    Dim myMethodInfo As MethodInfo = CType(myArrayMethodInfo(i), MethodInfo)
                    Lst.Add(myMethodInfo.Name)
                Next i
                Return Lst
            End Function

            'Main
            'Main
            Public Structure MethodItem

                Public TypeMethod As String
                Public TypeMethodInfo As MethodInfo
                Public TypeName As String

            End Structure

        End Class

    End Namespace

End Namespace