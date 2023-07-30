Namespace AI_SDK_EXTENSIONS

    Namespace MathsExt
        Public Class Statistic
            Public Frequency As Integer
            Public ReadOnly RelativeFrequency As Integer = CalcRelativeFrequency()
            Public ReadOnly RelativeFrequencyPercentage As Integer = CalcRelativeFrequencyPercentage()
            Public Total_SetItems As Integer
            Public ItemName As String
            Public ItemData As Object

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

            Private Function CalcRelativeFrequency() As Integer
                'IE: Frequency = 8 Total = 20 8/20= 0.4
                Return Me.Frequency / Me.Total_SetItems
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

            Private Function CalcRelativeFrequencyPercentage() As Integer
                'IE: Frequency = 8 Total = 20 8/20= 0.4 * 100 = 40%
                Return Me.Frequency / Me.Total_SetItems * 100
            End Function

        End Class

    End Namespace

End Namespace