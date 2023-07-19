Namespace AI_SDK_EXTENSIONS

    Namespace Strings

        Public Module TextPreProcessing

            Public ReadOnly AlphaBet() As String = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N",
            "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n",
            "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", " "}

            Public ReadOnly EncapuslationPunctuationEnd() As String = {"}", "]", ">", ")"}
            Public ReadOnly EncapuslationPunctuationStart() As String = {"{", "[", "<", "("}
            Public ReadOnly GramaticalPunctuation() As String = {".", "?", "!", ":", ";", ",", "_", "&"}
            Public ReadOnly MathPunctuation() As String = {"+", "-", "=", "/", "*", "%", "<", "(", ">", ")"}
            Public ReadOnly MoneyPunctuation() As String = {"£", "$"}
            Public ReadOnly Numerical() As String = {"1", "2", "3", "4", "5", "6", "7", "8", "9", "0"}
            Public ReadOnly Symbols() As String = {"£", "$", "^", "@", "#", "~", "\"}
            Public StopWords As New List(Of String)

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
                            RemoveStopWords(Txt, StopWords)
                        Case "Tokenize Words"
                            Txt = TokenizeWords(Txt)
                            Dim Words() As String = Txt.Split(",")
                            Txt &= vbNewLine & "Total Tokens in doc  -" & Words.Count - 1 & ":" & vbNewLine
                        Case "Tokenize Sentences"
                            Txt = TokenizeSentences(Txt)
                            Dim Words() As String = Txt.Split(",")
                            Txt &= vbNewLine & "Total Tokens in doc  -" & Words.Count - 2 & ":" & vbNewLine
                        Case "Remove Symbols"
                            Txt = RemoveSymbols(Txt).Replace("  ", " ")
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
            Public Function RemoveStopWords(ByRef txt As String, ByRef StopWrds As List(Of String)) As String
                For Each item In StopWrds
                    txt = txt.Replace(item, "")
                Next
                Return txt
            End Function

            <Runtime.CompilerServices.Extension()>
            Public Function RemoveSymbols(ByRef Txt As String) As String
                'Basic Symbols
                Txt = Txt.Replace("£", "")
                Txt = Txt.Replace("$", "")
                Txt = Txt.Replace("^", "")
                Txt = Txt.Replace("@", "")
                Txt = Txt.Replace("#", "")
                Txt = Txt.Replace("~", "")
                Txt = Txt.Replace("\", "")
                Return Txt
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

        End Module
    End Namespace

End Namespace