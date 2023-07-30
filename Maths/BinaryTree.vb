Namespace AI_SDK_EXTENSIONS

    Namespace MathsExt
        ''' <summary>
        ''' All nodes in a binary Tree have Left and Right nodes Nodes are added to the end of the tree
        ''' organized then the tree can be reorganized the rules are such that the lowest numbers are
        ''' always on the left and the highest numbers are on the right
        ''' </summary>
        Public Class BinaryTree

            ''' <summary>
            ''' Data to be held in Node
            ''' </summary>
            Public Data As Integer

            Public Left As BinaryTree

            ''' <summary>
            ''' String Tree
            ''' </summary>
            Public PRN As String = ""

            Public Right As BinaryTree

            ''' <summary>
            ''' TreeControl Tree (once tree has been printed then the view Will Be populated)
            ''' </summary>
            Public Tree As System.Windows.Forms.TreeView

            Public Sub New(ByVal Data As Integer)
                Me.Data = Data
            End Sub

            Public Property Prnt As String
                Get
                    Return PRN

                End Get
                Set(value As String)
                    PRN = value
                End Set
            End Property

            Public Function Contains(ByRef ExistingData As Integer) As Boolean
                Return If(ExistingData = Me.Data, True, If(ExistingData < Data, If(Left Is Nothing, False, Left.Contains(ExistingData)), If(Right Is Nothing, False, Right.Contains(ExistingData))))
            End Function

            ''' <summary>
            ''' Inserts new data into tree and adds node in appropriate place
            ''' </summary>
            ''' <param name="NewData"></param>
            Public Sub insert(ByVal NewData As Integer)
                If NewData <= Data Then
                    If Left IsNot Nothing Then
                        Left = New BinaryTree(NewData)
                    Else
                        Left.insert(NewData)
                    End If
                Else
                    If Right IsNot Nothing Then
                        Right = New BinaryTree(NewData)
                    Else
                        Right.insert(NewData)
                    End If
                End If

            End Sub

            ''' <summary>
            ''' Prints in order ABC Left then Root then Right B=Root A=Left C=Right
            ''' </summary>
            Public Sub PrintInOrder()
                If Left IsNot Nothing Then
                    Left.PrintInOrder()
                End If

                Prnt &= "Node :" & Me.Data & vbNewLine
                Dim nde As New System.Windows.Forms.TreeNode
                nde.Text = "Node :" & Me.Data & vbNewLine
                Tree.Nodes.Add(nde)
                If Right IsNot Nothing Then
                    Right.PrintInOrder()
                End If

            End Sub

            ''' <summary>
            ''' Prints in order ACB Left then Right Then Root B=Root A=Left C=Right
            ''' </summary>
            Public Sub PrintPostOrder()
                'Left Nodes
                If Left IsNot Nothing Then
                    Left.PrintInOrder()
                End If
                'Right nodes
                If Right IsNot Nothing Then
                    Right.PrintInOrder()
                End If

                'Root
                Prnt &= "Node :" & Me.Data & vbNewLine
                Dim nde As New System.Windows.Forms.TreeNode
                nde.Text = "Node :" & Me.Data & vbNewLine
                Tree.Nodes.Add(nde)
            End Sub

            ''' <summary>
            ''' Prints in order BAC Root then left then right B=Root A=Left C=Right
            ''' </summary>
            Public Sub PrintPreOrder()
                'Root
                Prnt &= "Node :" & Me.Data & vbNewLine
                Dim nde As New System.Windows.Forms.TreeNode
                nde.Text = "Node :" & Me.Data & vbNewLine
                Tree.Nodes.Add(nde)
                'Right nodes
                If Right IsNot Nothing Then
                    Right.PrintInOrder()
                End If
                'Left Nodes
                If Left IsNot Nothing Then
                    Left.PrintInOrder()
                End If
            End Sub

        End Class

    End Namespace

End Namespace