Namespace AI_SDK_EXTENSIONS

    Namespace Strings

        Namespace Tokens

            Public Module TDM_Extensions

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

                Public Function CorpusFrequency(ByRef Term As String, ByRef Documents As List(Of String)) As Integer
                    'Sum Of frequency for term in each document
                    Dim Freq As Integer = 0
                    For Each item In Documents
                        Freq += LCase(Term).GetTokenFrequency(LCase(item))
                    Next
                    Return Freq
                End Function

                Public Function CreateDataTable(ByRef HeaderTitles As List(Of String)) As DataTable
                    Dim DT As New DataTable
                    For Each item In HeaderTitles
                        DT.Columns.Add(item, GetType(String))
                    Next
                    Return DT
                End Function

                Public Function DocumentFrequency(ByRef Term As String, ByRef Documents As List(Of String)) As Integer
                    'Document Frequency how many documents term apeares
                    Dim Freq As Integer = 0
                    For Each item In Documents
                        If item.Contains(UCase(Term)) Or item.Contains(LCase(Term)) = True Then
                            Freq += 1
                        Else
                        End If
                    Next
                    Return Freq
                End Function

                Public Function InverseCorpusFrequency(ByRef Term As String, ByRef Documents As List(Of String)) As Double
                    'LOG(numberOf docs / CorpusFrequency)
                    Return LOG10(CorpusFrequency(Term, Documents) / DocumentFrequency(Term, Documents))
                End Function

                Public Function InverseDocuemntFrequency(ByRef Term As String, ByRef Documents As List(Of String)) As Double
                    'Rarety of term ,  Hi number = rare term, Low number = common term
                    'LOG(numberOf docs / documentFrequency)
                    Return LOG10(DocumentFrequency(Term, Documents) / Documents.Count)
                End Function

                Public Function TermFrequency(ByRef Term As String, ByRef Document As String) As Integer
                    Return LCase(Term).GetTokenFrequency(LCase(Document))
                End Function

                Public Structure Doc
                    Dim DocContents As String

                    Dim DocTitle As String

                    Dim Terms As List(Of Term)

                    Public Structure Term
                        Dim DocNumber As Integer
                        Dim Freq As Integer
                        Dim Term As String
                    End Structure

                End Structure

                Public Structure TermDocument
                    Dim DocContents As String
                    Dim DocTitle As String
                    Dim Freq As Integer
                    Dim Term As String
                End Structure

                Public Class Statistic
                    Public ReadOnly RelativeFrequency As Integer = CalcRelativeFrequency()
                    Public ReadOnly RelativeFrequencyPercentage As Integer = CalcRelativeFrequencyPercentage()
                    Public Frequency As Integer
                    Public ItemData As Object
                    Public ItemName As String
                    Public Total_SetItems As Integer

                    ''' <summary>
                    ''' IE: Frequency = 8 Total = 20 8/20= 0.4
                    ''' </summary>
                    ''' <param name="Frequency"></param>
                    ''' <param name="Total"></param>
                    ''' <returns></returns>
                    Public Shared Function CalcRelativeFrequency(ByRef Frequency As Integer, ByRef Total As Integer) As Integer
                        'IE: Frequency = 8 Total = 20 8/20= 0.4
                        Return Frequency / Total
                    End Function

                    ''' <summary>
                    '''    IE: Frequency = 8 Total = 20 8/20= 0.4 * 100 = 40%
                    ''' </summary>
                    ''' <param name="Frequency"></param>
                    ''' <param name="Total"></param>
                    ''' <returns></returns>
                    Public Function CalcRelativeFrequencyPercentage(ByRef Frequency As Integer, ByRef Total As Integer) As Integer
                        'IE: Frequency = 8 Total = 20 8/20= 0.4 * 100 = 40%
                        Return Frequency / Total * 100
                    End Function

                    Private Function CalcRelativeFrequency() As Integer
                        'IE: Frequency = 8 Total = 20 8/20= 0.4
                        Return Me.Frequency / Me.Total_SetItems
                    End Function

                    Private Function CalcRelativeFrequencyPercentage() As Integer
                        'IE: Frequency = 8 Total = 20 8/20= 0.4 * 100 = 40%
                        Return Me.Frequency / Me.Total_SetItems * 100
                    End Function

                End Class

            End Module
        End Namespace

    End Namespace

End Namespace