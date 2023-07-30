Namespace AI_SDK_EXTENSIONS

    Namespace MathsExt
        ''' <summary>
        ''' 2015: Testing Required: Linear regression using X n Y data variables
        ''' </summary>
        ''' <remarks>
        ''' if Correlation of data is found then predictions can be made via interpolation of Given X If
        ''' correlation is found then Given the Prediction What is X can be found with extrapolation
        ''' </remarks>
        Public Class BasicRegression

            Public Sample As List(Of DataPair)

            Public Shared Function Extrapolation(ByRef Sample As List(Of DataPair), ByRef Yvalue As Double) As Double

                ' x = y - intercept / slope
                Extrapolation = 0
                Dim RegressionTab As RegressionTable = CreateTable(Sample)
                'Good Prediction

                Extrapolation = If(CoefficiantCorrelation(RegressionTab) = True, Yvalue - RegressionTab.YIntercept / RegressionTab.Slope, 0)

                'Returns X value
            End Function

            Public Shared Function Interpolation(ByRef Sample As List(Of DataPair), ByRef Xvalue As Double) As Double
                'Formulae A^ * x + B^
                ' y = Slope * X + Yintercept
                Interpolation = 0
                Dim RegressionTab As RegressionTable = CreateTable(Sample)
                'Good Prediction
                Interpolation = If(CoefficiantCorrelation(RegressionTab) = True, RegressionTab.Slope * Xvalue + RegressionTab.YIntercept, 0)
                'Returns y Prediction
            End Function

            Private Shared Function CoefficiantCorrelation(ByRef Regression As RegressionTable) As Boolean
                CoefficiantCorrelation = False

                Dim R As Double = Regression.CoefficiantCorrelation

                'Is positive
                'If Slope is positive
                'is negative
                CoefficiantCorrelation = If(Regression.Slope > 0, If(R < 0, False, True), If(R < 0, True, False))
            End Function

            ''' <summary>
            ''' Creates a regression table of results
            ''' </summary>
            ''' <param name="mSample"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Private Shared Function CreateTable(ByRef mSample As List(Of DataPair)) As RegressionTable
                For Each mSampl As DataPair In mSample
                    CreateTable.SquY = mSampl.Yvect * mSampl.Yvect
                    CreateTable.SquX = mSampl.Xvect * mSampl.Xvect
                    CreateTable.XY = mSampl.Xvect * mSampl.Yvect
                    CreateTable.SumXSqu = CreateTable.SumXSqu + CreateTable.SquX
                    CreateTable.SumySqu = CreateTable.SumySqu + CreateTable.SquY
                    CreateTable.SumXvect = CreateTable.SumXvect + mSampl.Xvect
                    CreateTable.SumYVect = CreateTable.SumYVect + mSampl.Yvect
                    CreateTable.SumXY = CreateTable.SumXY + CreateTable.XY
                    CreateTable.NumberofPairs = CreateTable.NumberofPairs + 1
                Next
                'Slope & Intercept
                CreateTable.Slope = FindSlope(CreateTable)
                CreateTable.YIntercept = FindYIntercept(CreateTable)

                'Standard Deviation
                Dim temp As Double = 0
                'Standard deviation X
                temp = CreateTable.SumXSqu - (1 / CreateTable.NumberofPairs) * (CreateTable.SumXvect * CreateTable.SumXvect) /
            (CreateTable.NumberofPairs - 1)
                CreateTable.StandardDevX = Math.Sqrt(temp)
                'Standard deviation y
                temp = CreateTable.SumySqu - (1 / CreateTable.NumberofPairs) * (CreateTable.SumYVect * CreateTable.SumYVect) /
            (CreateTable.NumberofPairs - 1)
                CreateTable.StandardDevY = Math.Sqrt(temp)
                'CoefficiantCorrelation
                'Case 1
                CreateTable.CoefficiantCorrelation = (CreateTable.StandardDevX / CreateTable.StandardDevY) * CreateTable.Slope
                'case 2
                '   CreateTable.CoefficiantCorrelation = CreateTable.SumXY - (1 / CreateTable.NumberofPairs) * CreateTable.SumXvect * CreateTable.SumYVect /
                '(CreateTable.NumberofPairs - 1) * CreateTable.StandardDevX * CreateTable.StandardDevY
            End Function

            ''' <summary>
            ''' Given a table of regression results what is the slope
            ''' </summary>
            ''' <param name="RegressionSample"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Private Shared Function FindSlope(ByRef RegressionSample As RegressionTable) As Double
                FindSlope = RegressionSample.SumXvect * RegressionSample.SumYVect - RegressionSample.NumberofPairs * RegressionSample.SumXY /
                    (RegressionSample.SumXvect * RegressionSample.SumXvect) - RegressionSample.NumberofPairs * RegressionSample.SumXSqu
            End Function

            ''' <summary>
            ''' Given a regression table of results what is the yIntercpt
            ''' </summary>
            ''' <param name="RegressionSample"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Private Shared Function FindYIntercept(ByRef RegressionSample As RegressionTable) As Double
                FindYIntercept = RegressionSample.SumXvect * RegressionSample.SumXY - RegressionSample.SumYVect * RegressionSample.SumXSqu /
                         (RegressionSample.SumXvect * RegressionSample.SumXvect) - RegressionSample.NumberofPairs * RegressionSample.SumXSqu
            End Function

            Public Structure DataPair

                Dim Xvect As Integer
                Dim Yvect As Integer

            End Structure

            Private Structure RegressionTable

                Dim CoefficiantCorrelation As Double

                Dim NumberofPairs As Integer

                Dim Slope As Double

                ''' <summary>
                ''' X*X
                ''' </summary>
                ''' <remarks></remarks>
                Dim SquX As Double

                ''' <summary>
                ''' Y*Y
                ''' </summary>
                ''' <remarks></remarks>
                Dim SquY As Double

                Dim StandardDevX As Double

                Dim StandardDevY As Double

                Dim SumXSqu As Double

                Dim SumXvect As Double

                Dim SumXY As Double

                Dim SumySqu As Double

                Dim SumYVect As Double

                ''' <summary>
                ''' X*Y
                ''' </summary>
                ''' <remarks></remarks>
                Dim XY As Double

                Dim YIntercept As Double

            End Structure

        End Class

    End Namespace

End Namespace