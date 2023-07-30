Namespace AI_SDK_EXTENSIONS

    Namespace MathsExt
        Public Module Maths

            'Math Functions

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

            <Runtime.CompilerServices.Extension()>
            Public Function Gaussian(ByRef x As Double) As Double
                Gaussian = Math.Exp((-x * -x) / 2)
            End Function

            <Runtime.CompilerServices.Extension()>
            Public Function GaussianDerivative(ByRef x As Double) As Double
                GaussianDerivative = Gaussian(x) * (-x / (-x * -x))
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
            Public Function IsSquareRoot(ByVal number As Integer) As Boolean
                Dim numberSquareRooted As Double = Math.Sqrt(number)
                Return If(CInt(numberSquareRooted) ^ 2 = number, True, False)
            End Function

            <Runtime.CompilerServices.Extension()>
            Public Function IsCubeRoot(ByVal number As Integer) As Boolean
                Dim numberCubeRooted As Double = number ^ (1 / 3)
                Return If(CInt(numberCubeRooted) ^ 3 = number, True, False)
            End Function

            <Runtime.CompilerServices.Extension()>
            Public Function IsRoot(ByVal number As Integer, power As Integer) As Boolean
                Dim numberNRooted As Double = number ^ (1 / power)
                Return If(CInt(numberNRooted) ^ power = number, True, False)
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
            ''' Number multiplied by itself
            ''' </summary>
            ''' <param name="Value"></param>
            ''' <returns></returns>
            <Runtime.CompilerServices.Extension()>
            Public Function Square(ByRef Value As Double) As Double
                Return Value * Value
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

            Enum CorrelationResult
                Positive = 1
                PositiveHigh = 0.9
                PositiveLow = 0.5
                None = 0
                NegativeLow = -0.5
                NegativeHigh = -0.9
                Negative = -1
            End Enum

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

            'Growth

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

        End Module

    End Namespace

End Namespace