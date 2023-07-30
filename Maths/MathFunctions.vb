Namespace AI_SDK_EXTENSIONS

    Namespace MathsExt
        <ComClass(MathFunctions.ClassId, MathFunctions.InterfaceId, MathFunctions.EventsId)>
        Public Class MathFunctions

            '
            Public Const ClassId As String = "2822E490-7703-401C-BAB3-38FF97BC1AC9"

            Public Const EventsId As String = "CDB54207-F55E-401A-AC6D-3CF8086FD6F1"
            Public Const InterfaceId As String = "8B3325F8-5D13-4059-829B-B531310144B5"

            ''' <summary>
            ''' Calculates a number between two numbers at a specific increment.
            ''' Amount parameter is the amount to interpolate between the two values
            ''' </summary>
            ''' <param name="Start"> first value </param>
            ''' <param name="[Stop]"> second value </param>
            ''' <param name="InterpAmount"> float between 0.0 and 1.0 </param>
            Public Shared Function Lerp(ByVal Start As Single, ByVal [Stop] As Single, ByVal InterpAmount As Single) As Single
                Return Start + ([Stop] - Start) * InterpAmount
            End Function

            ''' <summary>
            ''' Re-maps a number from one range to another. In the example above,
            ''' </summary>
            ''' <param name="value"> the incoming value to be converted </param>
            ''' <param name="start1"> lower bound of the value's current range </param>
            ''' <param name="stop1"> upper bound of the value's current range </param>
            ''' <param name="start2"> lower bound of the value's target range </param>
            ''' <param name="stop2"> upper bound of the value's target range </param>
            Public Shared Function Map(ByVal value As Single, ByVal start1 As Single, ByVal stop1 As Single, ByVal start2 As Single, ByVal stop2 As Single) As Single
                Dim Output As Single = start2 + (stop2 - start2) * ((value - start1) / (stop1 - start1))
                Dim errMessage As String = Nothing

                If Output <> Output Then
                    errMessage = "NaN (not a number)"
                    Throw New Exception(errMessage)
                ElseIf Output = Single.NegativeInfinity OrElse Output = Single.PositiveInfinity Then
                    errMessage = "infinity"
                    Throw New Exception(errMessage)
                End If
                Return Output
            End Function

            ''' <summary>
            ''' Normalizes a number from another range into a value between 0 and 1.
            ''' Identical to map(value, low, high, 0, 1);
            ''' Numbers outside the range are not clamped to 0 and 1, because
            ''' out-of-range values are often intentional and useful.
            ''' </summary>
            ''' <param name="value"> the incoming value to be converted </param>
            ''' <param name="start"> lower bound of the value's current range </param>
            ''' <param name="stop"> upper bound of the value's current range </param>
            Public Shared Function Norm(ByVal value As Single, ByVal start As Single, ByVal [stop] As Single) As Single
                Return (value - start) / ([stop] - start)
            End Function

            ''' <summary>
            ''' Constrains value between min and max values
            '''   if less than min, return min
            '''   more than max, return max
            '''   otherwise return same value
            ''' </summary>
            ''' <param name="Value"></param>
            ''' <param name="min"></param>
            ''' <param name="max"></param>
            ''' <returns></returns>
            Public Function Constraint(Value As Single, min As Single, max As Single) As Single
                If Value <= min Then
                    Return min
                ElseIf Value >= max Then
                    Return max
                End If
                Return Value
            End Function

            ''' <summary>
            ''' Generates 8 bit array of an integer, value from 0 to 255
            ''' </summary>
            ''' <param name="Value"></param>
            ''' <returns></returns>
            Public Function GetBitArray(Value As Integer) As Integer()
                Dim Result(7) As Integer
                Dim sValue As String
                Dim cValue() As Char

                Value = Constraint(Value, 0, 255)
                sValue = Convert.ToString(Value, 2).PadLeft(8, "0"c)
                cValue = sValue.ToArray
                For i As Integer = 0 To cValue.Count - 1
                    If cValue(i) = "1"c Then
                        Result(i) = 1
                    Else
                        Result(i) = 0
                    End If
                Next
                Return Result
            End Function

            ''' <summary>
            ''' Generates 8 bit array of an integer, value from 0 to 255
            ''' </summary>
            ''' <param name="Value"></param>
            ''' <returns></returns>
            Public Function GetBits(Value As Integer) As Integer()
                Dim _Arr As BitArray
                Dim Result() As Integer = {0, 0, 0, 0, 0, 0, 0, 0}

                _Arr = New BitArray(New Integer() {Value})
                _Arr.CopyTo(Result, 0)
                Return Result
            End Function

            Public Function GetBitsString(Value As Integer) As String
                Dim _Array() As Integer = GetBitArray(Value)
                Dim Result As String = ""

                For I As Integer = 0 To _Array.Length - 1
                    Result += _Array(I).ToString
                Next
                Return Result
            End Function

            ''' <summary>
            ''' Used To Hold Risk Evaluation Data
            ''' </summary>
            Public Structure RiskNode

                Private mCost As Integer

                Private mExpectedMonetaryValue As Integer

                Private mExpectedMonetaryValueWithOutPerfectInformation As Integer

                Private mExpectedMonetaryValueWithPerfectInformation As Integer

                Private mGain As Integer

                Private mProbability As Integer

                Private mRegret As Integer

                Public Property Cost As Integer
                    Get
                        Return mCost
                    End Get
                    Set(value As Integer)
                        mCost = value
                    End Set
                End Property

                Public Property ExpectedMonetaryValue As Integer
                    Get
                        Return mExpectedMonetaryValue
                    End Get
                    Set(value As Integer)
                        mExpectedMonetaryValue = value
                    End Set
                End Property

                Public Property ExpectedMonetaryValueWithOutPerfectInformation As Integer
                    Get
                        Return mExpectedMonetaryValueWithOutPerfectInformation
                    End Get
                    Set(value As Integer)
                        mExpectedMonetaryValueWithOutPerfectInformation = value
                    End Set
                End Property

                Public Property ExpectedMonetaryValueWithPerfectInformation As Integer
                    Get
                        Return mExpectedMonetaryValueWithPerfectInformation
                    End Get
                    Set(value As Integer)
                        mExpectedMonetaryValueWithPerfectInformation = value
                    End Set
                End Property

                Public Property Gain As Integer
                    Get
                        Return mGain
                    End Get
                    Set(value As Integer)
                        mGain = value
                    End Set
                End Property

                Public Property Probability As Integer
                    Get
                        Return mProbability
                    End Get
                    Set(value As Integer)
                        mProbability = value
                    End Set
                End Property

                Public Property Regret As Integer
                    Get
                        Return mRegret
                    End Get
                    Set(value As Integer)
                        mRegret = value
                    End Set
                End Property

            End Structure

            ''' <summary>
            ''' Cartesian Corordinate Functions
            ''' Vector of X,Y,Z
            ''' </summary>
            <ComClass(Cartesian.ClassId, Cartesian.InterfaceId, Cartesian.EventsId)>
            <Serializable>
            Public Class Cartesian

                Public Const ClassId As String = "2828B490-7103-401C-7AB3-38FF97BC1AC9"
                Public Const EventsId As String = "CDB74307-F15E-401A-AC6D-3CF8786FD6F1"
                Public Const InterfaceId As String = "8BB745F8-5113-4059-829B-B531310144B5"

                ''' <summary>
                ''' The x component of the Cartesian Co-ordinate
                ''' </summary>
                Public x As Single

                ''' <summary>
                ''' The y component of the Cartesian Co-ordinate
                ''' </summary>
                Public y As Single

                ''' <summary>
                ''' The z component of the Cartesian Co-ordinate
                ''' </summary>
                Public z As Single

                <NonSerialized>
                Private _RandGenerator As RandomFactory

                ''' <summary>
                ''' Constructor for an empty Cartesian: x, y, and z are set to 0.
                ''' </summary>
                Public Sub New()
                    x = 0
                    y = 0
                    z = 0
                    _RandGenerator = New RandomFactory
                End Sub

                ''' <summary>
                ''' Constructor for a 3D Cartesian.
                ''' </summary>
                ''' <param name="x"> the x coordinate. </param>
                ''' <param name="y"> the y coordinate. </param>
                ''' <param name="z"> the z coordinate. </param>
                Public Sub New(ByVal x As Single, ByVal y As Single, ByVal z As Single)
                    Me.New
                    Me.x = x
                    Me.y = y
                    Me.z = z
                End Sub

                ''' <summary>
                ''' Constructor for a 2D Cartesian: z coordinate is set to 0.
                ''' </summary>
                Public Sub New(ByVal x As Single, ByVal y As Single)
                    Me.New
                    Me.x = x
                    Me.y = y
                End Sub

                ''' <summary>
                ''' Subtract one Cartesian from another and store in another Cartesian </summary>
                Public Shared Function [Sub](ByVal v1 As Cartesian, ByVal v2 As Cartesian) As Cartesian
                    Dim target As New Cartesian

                    target.Set(v1.x - v2.x, v1.y - v2.y, v1.z - v2.z)
                    Return target
                End Function

                ''' <summary>
                ''' Add two Cartesians into a target Cartesian </summary>
                Public Shared Function Add(ByVal v1 As Cartesian, ByVal v2 As Cartesian) As Cartesian
                    Dim target As New Cartesian()

                    target.Set(v1.x + v2.x, v1.y + v2.y, v1.z + v2.z)
                    Return target
                End Function

                ''' <summary>
                ''' Calculates and returns the angle (in radians) between two Cartesians.
                ''' </summary>
                ''' <param name="v1"> 1st Cartesian Co-ordinate </param>
                ''' <param name="v2"> 2nd Cartesian Co-ordinate </param>
                Public Shared Function AngleBetween(ByVal v1 As Cartesian, ByVal v2 As Cartesian) As Single
                    Dim dot As Double = v1.x * v2.x + v1.y * v2.y + v1.z * v2.z
                    Dim V1Mag As Double = v1.Magnitude
                    Dim V2Mag As Double = v2.Magnitude
                    Dim Amount As Double = dot / (V1Mag * V2Mag)

                    If v1.x = 0 AndAlso v1.y = 0 AndAlso v1.z = 0 Then
                        Return 0.0F
                    End If
                    If v2.x = 0 AndAlso v2.y = 0 AndAlso v2.z = 0 Then
                        Return 0.0F
                    End If
                    If Amount <= -1 Then
                        Return Math.PI
                    ElseIf Amount >= 1 Then
                        Return 0
                    End If
                    Return CSng(Math.Acos(Amount))
                End Function

                ''' <param name="v1"> any variable of type Cartesian </param>
                ''' <param name="v2"> any variable of type Cartesian </param>
                Public Shared Function Cross(ByVal v1 As Cartesian, ByVal v2 As Cartesian) As Cartesian
                    Dim target As New Cartesian
                    Dim crossX As Single = v1.y * v2.z - v2.y * v1.z
                    Dim crossY As Single = v1.z * v2.x - v2.z * v1.x
                    Dim crossZ As Single = v1.x * v2.y - v2.x * v1.y

                    target.Set(crossX, crossY, crossZ)
                    Return target
                End Function

                ''' <param name="v1"> any variable of type Cartesian </param>
                ''' <param name="v2"> any variable of type Cartesian </param>
                ''' <returns> the Euclidean distance between v1 and v2 </returns>
                Public Shared Function Distance(ByVal v1 As Cartesian, ByVal v2 As Cartesian) As Single
                    Dim dx As Single = v1.x - v2.x
                    Dim dy As Single = v1.y - v2.y
                    Dim dz As Single = v1.z - v2.z
                    Return CSng(Math.Sqrt(dx * dx + dy * dy + dz * dz))
                End Function

                ''' <summary>
                ''' Divide a Cartesian by a scalar and store the result in another Cartesian. </summary>
                Public Shared Function Div(ByVal v As Cartesian, ByVal n As Single) As Cartesian
                    Dim target As New Cartesian

                    target.Set(v.x / n, v.y / n, v.z / n)

                    Return target
                End Function

                ''' <param name="v1"> any variable of type Cartesian </param>
                ''' <param name="v2"> any variable of type Cartesian </param>
                Public Shared Function Dot(ByVal v1 As Cartesian, ByVal v2 As Cartesian) As Single
                    Return v1.x * v2.x + v1.y * v2.y + v1.z * v2.z
                End Function

                ''' <summary>
                ''' Linear interpolate between two Cartesians (returns a new Cartesian object) </summary>
                ''' <param name="v1"> the Cartesian to start from </param>
                ''' <param name="v2"> the Cartesian to lerp to </param>
                Public Shared Function Lerp(ByVal v1 As Cartesian, ByVal v2 As Cartesian, ByVal Amount As Single) As Cartesian
                    Dim v As Cartesian = v1.Copy()
                    v.Lerp(v2, Amount)
                    Return v
                End Function

                ''' <summary>
                ''' Multiply a Cartesian by a scalar, and write the result into a target Cartesian. </summary>
                Public Shared Function Mult(ByVal v As Cartesian, ByVal n As Single) As Cartesian
                    Dim target As New Cartesian

                    target.Set(v.x * n, v.y * n, v.z * n)
                    Return target
                End Function

                ''' <summary>
                ''' Sets the x, y, and z component of the Cartesian using two or three separate
                ''' variables, the data from a Cartesian, or the values from a float array.
                '''  </summary>
                ''' <param name="x"> the x component of the Cartesian Co-ordinate </param>
                ''' <param name="y"> the y component of the Cartesian Co-ordinate </param>
                ''' <param name="z"> the z component of the Cartesian Co-ordinate </param>
                Public Overridable Function [Set](ByVal x As Single, ByVal y As Single, ByVal z As Single) As Cartesian
                    Me.x = x
                    Me.y = y
                    Me.z = z
                    Return Me
                End Function

                ''' <summary>
                ''' Sets the x, y, and z component of the Cartesian using two or three separate
                ''' variables, the data from a Cartesian, or the values from a float array.
                '''  </summary>
                ''' <param name="x"> the x component of the Cartesian Co-ordinate </param>
                ''' <param name="y"> the y component of the Cartesian Co-ordinate </param>
                Public Overridable Function [Set](ByVal x As Single, ByVal y As Single) As Cartesian
                    Me.x = x
                    Me.y = y
                    Me.z = 0
                    Return Me
                End Function

                ''' <summary>
                ''' Sets the x, y, and z component of the Cartesian Co-ordinate from another Cartesian Co-ordinate
                '''  </summary>
                ''' <param name="v"> Cartesian Co-ordinate to copy from </param>
                Public Overridable Function [Set](ByVal v As Cartesian) As Cartesian
                    x = v.x
                    y = v.y
                    z = v.z
                    Return Me
                End Function

                ''' <summary>
                ''' Set the x, y (and maybe z) coordinates using a float[] array as the source. </summary>
                ''' <param name="source"> array to copy from </param>
                Public Overridable Function [Set](ByVal source() As Single) As Cartesian
                    If source.Length >= 2 Then
                        x = source(0)
                        y = source(1)
                    End If
                    If source.Length >= 3 Then
                        z = source(2)
                    Else
                        z = 0
                    End If
                    Return Me
                End Function

                ''' <summary>
                ''' Subtracts x, y, and z components from a Cartesian, subtracts one Cartesian
                ''' from another, or subtracts two independent Cartesians
                ''' </summary>
                ''' <param name="v"> any variable of type Cartesian </param>
                Public Overridable Function [Sub](ByVal v As Cartesian) As Cartesian
                    x -= v.x
                    y -= v.y
                    z -= v.z
                    Return Me
                End Function

                ''' <param name="x"> the x component of the Cartesian </param>
                ''' <param name="y"> the y component of the Cartesian </param>
                Public Overridable Function [Sub](ByVal x As Single, ByVal y As Single) As Cartesian
                    Me.x -= x
                    Me.y -= y
                    Return Me
                End Function

                ''' <param name="z"> the z component of the Cartesian </param>
                Public Overridable Function [Sub](ByVal x As Single, ByVal y As Single, ByVal z As Single) As Cartesian
                    Me.x -= x
                    Me.y -= y
                    Me.z -= z
                    Return Me
                End Function

                ''' <summary>
                ''' Adds x, y, and z components to a Cartesian, adds one Cartesian to another, or
                ''' adds two independent Cartesians together
                ''' </summary>
                ''' <param name="v"> the Cartesian to be added </param>
                Public Overridable Function Add(ByVal v As Cartesian) As Cartesian
                    x += v.x
                    y += v.y
                    z += v.z
                    Return Me
                End Function

                ''' <param name="x"> x component of the Cartesian </param>
                ''' <param name="y"> y component of the Cartesian </param>
                Public Overridable Function Add(ByVal x As Single, ByVal y As Single) As Cartesian
                    Me.x += x
                    Me.y += y
                    Return Me
                End Function

                ''' <param name="z"> z component of the Cartesian Co-ordinate </param>
                Public Overridable Function Add(ByVal x As Single, ByVal y As Single, ByVal z As Single) As Cartesian
                    Me.x += x
                    Me.y += y
                    Me.z += z
                    Return Me
                End Function

                ''' <summary>
                ''' Gets a copy of the Cartesian, returns a Cartesian object.
                ''' </summary>
                Public Overridable Function Copy() As Cartesian
                    Return New Cartesian(x, y, z)
                End Function

                ''' <summary>
                ''' Calculates and returns a Cartesian composed of the cross product between two Cartesians
                ''' </summary>
                ''' <param name="v"> 2nd Cartesian Cartesian </param>
                Public Overridable Function Cross(ByVal v As Cartesian) As Cartesian
                    Dim target As New Cartesian
                    Dim crossX As Single = y * v.z - v.y * z
                    Dim crossY As Single = z * v.x - v.z * x
                    Dim crossZ As Single = x * v.y - v.x * y

                    target.Set(crossX, crossY, crossZ)
                    Return target
                End Function

                ''' <summary>
                ''' Calculates the Euclidean distance between two Cartesians
                ''' </summary>
                ''' <param name="v"> the x, y, and z coordinates of a Cartesian</param>
                Public Overridable Function Distance(ByVal v As Cartesian) As Single
                    Dim dx As Single = x - v.x
                    Dim dy As Single = y - v.y
                    Dim dz As Single = z - v.z
                    Return CSng(Math.Sqrt(dx * dx + dy * dy + dz * dz))
                End Function

                ''' <summary>
                ''' Divides a Cartesian by a scalar or divides one Cartesian by another.
                ''' </summary>
                ''' <param name="n"> the number by which to divide the Cartesian </param>
                Public Overridable Function Div(ByVal n As Single) As Cartesian
                    x /= n
                    y /= n
                    z /= n
                    Return Me
                End Function

                ''' <summary>
                ''' Calculates the dot product of two Cartesians.
                ''' </summary>
                ''' <param name="v"> any variable of type Cartesian </param>
                ''' <returns> the dot product </returns>
                Public Overridable Function Dot(ByVal v As Cartesian) As Single
                    Return x * v.x + y * v.y + z * v.z
                End Function

                ''' <param name="x"> x component of the Cartesian </param>
                ''' <param name="y"> y component of the Cartesian </param>
                ''' <param name="z"> z component of the Cartesian </param>
                Public Overridable Function Dot(ByVal x As Single, ByVal y As Single, ByVal z As Single) As Single
                    Return Me.x * x + Me.y * y + Me.z * z
                End Function

                Public Overrides Function Equals(ByVal obj As Object) As Boolean
                    If Not (TypeOf obj Is Cartesian) Then
                        Return False
                    End If
                    Dim p As Cartesian = DirectCast(obj, Cartesian)
                    Return x = p.x AndAlso y = p.y AndAlso z = p.z
                End Function

                ''' <summary>
                ''' Make a new 2D unit Cartesian from an angle
                ''' </summary>
                ''' <param name="target"> the target Cartesian (if null, a new Cartesian will be created) </param>
                ''' <returns> the Cartesian </returns>
                Public Function FromAngle(ByVal angle As Single, ByVal target As Cartesian) As Cartesian
                    Dim Output As New Cartesian()

                    Output.Set(CSng(Math.Cos(angle)), CSng(Math.Sin(angle)), 0)
                    Return Output
                End Function

                ''' <summary>
                ''' Make a new 2D unit Cartesian Co-ordinate from an angle.
                ''' </summary>
                ''' <param name="angle"> the angle in radians </param>
                ''' <returns> the new unit Cartesian Co-ordinate </returns>
                Public Function FromAngle(ByVal angle As Single) As Cartesian
                    Return FromAngle(angle, Me)
                End Function

                ''' <summary>
                ''' Return Cartesian values as array
                ''' </summary>
                ''' <returns></returns>
                Public Overridable Function GetArray() As Single()
                    Dim Output(2) As Single

                    Output(0) = x
                    Output(1) = y
                    Output(2) = z

                    Return Output
                End Function

                Public Overrides Function GetHashCode() As Integer
                    Dim result As Integer = 1
                    result = 31 * result
                    result = 31 * result
                    result = 31 * result
                    Return result
                End Function

                ''' <summary>
                ''' Calculate the angle of rotation for this Cartesian (only 2D Cartesians)
                ''' </summary>
                ''' <returns> the angle of rotation </returns>
                Public Overridable Function Heading() As Single
                    Dim angle As Single = CSng(Math.Atan2(y, x))

                    Return angle
                End Function

                ''' <summary>
                ''' Linear interpolate the Cartesian to another Cartesian
                ''' </summary>
                ''' <param name="v"> the Cartesian to lerp to </param>
                ''' <param name="Amount">  The amount of interpolation; some value between 0.0 (old Cartesian) and 1.0 (new Cartesian). 0.1 is very near the old Cartesian; 0.5 is halfway in between. </param>
                Public Overridable Function Lerp(ByVal v As Cartesian, ByVal Amount As Single) As Cartesian
                    x = MathFunctions.Lerp(x, v.x, Amount)
                    y = MathFunctions.Lerp(y, v.y, Amount)
                    z = MathFunctions.Lerp(z, v.z, Amount)
                    Return Me
                End Function

                ''' <summary>
                ''' Linear interpolate the Cartesian Co-ordinate to x,y,z values </summary>
                ''' <param name="x"> the x component to lerp to </param>
                ''' <param name="y"> the y component to lerp to </param>
                ''' <param name="z"> the z component to lerp to </param>
                Public Overridable Function Lerp(ByVal x As Single, ByVal y As Single, ByVal z As Single, ByVal Amount As Single) As Cartesian
                    Me.x = MathFunctions.Lerp(Me.x, x, Amount)
                    Me.y = MathFunctions.Lerp(Me.y, y, Amount)
                    Me.z = MathFunctions.Lerp(Me.z, z, Amount)
                    Return Me
                End Function

                ''' <summary>
                ''' Limit the magnitude of this Cartesian to the value passed as max parameter
                ''' </summary>
                ''' <param name="max"> the maximum magnitude for the Cartesian </param>
                Public Overridable Function Limit(ByVal max As Single) As Cartesian
                    If MagSq() > max * max Then
                        Normalize()
                        Mult(max)
                    End If
                    Return Me
                End Function

                ''' <summary>
                ''' Calculates the magnitude (length) of the Cartesian Co-ordinate and returns the result
                ''' </summary>
                ''' <returns> magnitude (length) of the Cartesian Co-ordinate </returns>
                Public Overridable Function Magnitude() As Single
                    Return CSng(Math.Sqrt(x * x + y * y + z * z))
                End Function

                ''' <summary>
                ''' Calculates the squared magnitude of the Cartesian and returns the result
                ''' </summary>
                ''' <returns> squared magnitude of the Cartesian </returns>
                Public Overridable Function MagSq() As Single
                    Return (x * x + y * y + z * z)
                End Function

                ''' <summary>
                ''' Multiplies a Cartesian by a scalar or multiplies one Cartesian by another.
                ''' </summary>
                ''' <param name="n"> the number to multiply with the Cartesian </param>
                Public Overridable Function Mult(ByVal n As Single) As Cartesian
                    x *= n
                    y *= n
                    z *= n
                    Return Me
                End Function

                ''' <summary>
                ''' Normalize the Cartesian to length 1 (make it a unit Cartesian).
                ''' </summary>
                Public Overridable Function Normalize() As Cartesian
                    Dim m As Single = Magnitude()

                    If m <> 0 AndAlso m <> 1 Then
                        Div(m)
                    End If
                    Return Me
                End Function

                ''' <param name="target"> Set to null to create a new Cartesian </param>
                ''' <returns> a new Cartesian (if target was null), or target </returns>
                Public Overridable Function Normalize(ByVal target As Cartesian) As Cartesian
                    If target Is Nothing Then
                        target = New Cartesian()
                    End If
                    Dim m As Single = Magnitude()
                    If m > 0 Then
                        target.Set(x / m, y / m, z / m)
                    Else
                        target.Set(x, y, z)
                    End If
                    Return target
                End Function

                ''' <summary>
                ''' Randmize X, Y and Z components of Cartesian Co-ordinate between 0 and 1
                ''' </summary>
                Public Sub Randomize()
                    Me.x = _RandGenerator.GetRandomDbl
                    Me.y = _RandGenerator.GetRandomDbl
                    Me.z = _RandGenerator.GetRandomDbl
                End Sub

                ''' <summary>
                ''' Rotate the Cartesian by an angle (only 2D Cartesians), magnitude remains the same
                ''' </summary>
                ''' <param name="theta"> the angle of rotation </param>
                Public Overridable Function Rotate(ByVal theta As Single) As Cartesian
                    Dim temp As Single = x

                    x = x * Math.Cos(theta) - y * Math.Sin(theta)
                    y = temp * Math.Sin(theta) + y * Math.Cos(theta)
                    Return Me
                End Function

                ''' <summary>
                ''' Set the magnitude of this Cartesian to the value passed as len parameter
                ''' </summary>
                ''' <param name="len"> the new length for this Cartesian </param>
                Public Overridable Function SetMag(ByVal len As Single) As Cartesian
                    Normalize()
                    Mult(len)
                    Return Me
                End Function

                ''' <summary>
                ''' Sets the magnitude of this Cartesian, storing the result in another Cartesian. </summary>
                ''' <param name="target"> Set to null to create a new Cartesian </param>
                ''' <param name="len"> the new length for the new Cartesian </param>
                ''' <returns> a new Cartesian (if target was null), or target </returns>
                Public Overridable Function SetMag(ByVal target As Cartesian, ByVal len As Single) As Cartesian
                    target = Normalize(target)
                    target.Mult(len)
                    Return target
                End Function

                ''' <summary>
                ''' Returns coordinates of Cartesian [x,y,z]
                ''' </summary>
                ''' <returns></returns>
                Public Overrides Function ToString() As String
                    Return "[ " & x & ", " & y & ", " & z & " ]"
                End Function

            End Class

            ''' <summary>
            ''' Given a molecule's chemical formula, calculate the   molar mass.
            '''In chemistry, the molar mass of a chemical compound Is defined as the mass of a sample of that compound divided by the amount of substance in that sample, measured in moles.
            '''[1] The molar mass Is a bulk, Not molecular, property of a substance.
            '''The molar mass Is an average of many instances of the compound, which often vary in mass due to the presence of isotopes.
            '''Most commonly, the molar mass Is computed from the standard atomic weights And Is thus a terrestrial average And a function of the relative abundance of the isotopes of the constituent atoms on Earth.
            '''The molar mass Is appropriate for converting between the mass of a substance And the amount of a substance for bulk quantities.
            '''The molecular weight Is very commonly used As a synonym Of molar mass, particularly For molecular compounds; however, the most authoritative sources define it differently (see molecular mass).
            '''The formula weight Is a synonym Of molar mass that Is frequently used For non-molecular compounds, such As ionic salts.
            ''' </summary>
            Public Class ChemicalMass

                ' Given a molecule's chemical formula, calculate the   molar mass.
                'In chemistry, the molar mass of a chemical compound Is defined as the mass of a sample of that compound divided by the amount of substance in that sample, measured in moles.[1] The molar mass Is a bulk, Not molecular, property of a substance. The molar mass Is an average of many instances of the compound, which often vary in mass due to the presence of isotopes. Most commonly, the molar mass Is computed from the standard atomic weights And Is thus a terrestrial average And a function of the relative abundance of the isotopes of the constituent atoms on Earth. The molar mass Is appropriate for converting between the mass of a substance And the amount of a substance for bulk quantities.
                'The molecular weight Is very commonly used As a synonym Of molar mass, particularly For molecular compounds; however, the most authoritative sources define it differently (see molecular mass).
                'The formula weight Is a synonym Of molar mass that Is frequently used For non-molecular compounds, such As ionic salts.
                Dim atomicMass As New Dictionary(Of String, Double) From {
                    {"H", 1.008},
                    {"He", 4.002602},
                    {"Li", 6.94},
                    {"Be", 9.0121831},
                    {"B", 10.81},
                    {"C", 12.011},
                    {"N", 14.007},
                    {"O", 15.999},
                    {"F", 18.998403163},
                    {"Ne", 20.1797},
                    {"Na", 22.98976928},
                    {"Mg", 24.305},
                    {"Al", 26.9815385},
                    {"Si", 28.085},
                    {"P", 30.973761998},
                    {"S", 32.06},
                    {"Cl", 35.45},
                    {"Ar", 39.948},
                    {"K", 39.0983},
                    {"Ca", 40.078},
                    {"Sc", 44.955908},
                    {"Ti", 47.867},
                    {"V", 50.9415},
                    {"Cr", 51.9961},
                    {"Mn", 54.938044},
                    {"Fe", 55.845},
                    {"Co", 58.933194},
                    {"Ni", 58.6934},
                    {"Cu", 63.546},
                    {"Zn", 65.38},
                    {"Ga", 69.723},
                    {"Ge", 72.63},
                    {"As", 74.921595},
                    {"Se", 78.971},
                    {"Br", 79.904},
                    {"Kr", 83.798},
                    {"Rb", 85.4678},
                    {"Sr", 87.62},
                    {"Y", 88.90584},
                    {"Zr", 91.224},
                    {"Nb", 92.90637},
                    {"Mo", 95.95},
                    {"Ru", 101.07},
                    {"Rh", 102.9055},
                    {"Pd", 106.42},
                    {"Ag", 107.8682},
                    {"Cd", 112.414},
                    {"In", 114.818},
                    {"Sn", 118.71},
                    {"Sb", 121.76},
                    {"Te", 127.6},
                    {"I", 126.90447},
                    {"Xe", 131.293},
                    {"Cs", 132.90545196},
                    {"Ba", 137.327},
                    {"La", 138.90547},
                    {"Ce", 140.116},
                    {"Pr", 140.90766},
                    {"Nd", 144.242},
                    {"Pm", 145},
                    {"Sm", 150.36},
                    {"Eu", 151.964},
                    {"Gd", 157.25},
                    {"Tb", 158.92535},
                    {"Dy", 162.5},
                    {"Ho", 164.93033},
                    {"Er", 167.259},
                    {"Tm", 168.93422},
                    {"Yb", 173.054},
                    {"Lu", 174.9668},
                    {"Hf", 178.49},
                    {"Ta", 180.94788},
                    {"W", 183.84},
                    {"Re", 186.207},
                    {"Os", 190.23},
                    {"Ir", 192.217},
                    {"Pt", 195.084},
                    {"Au", 196.966569},
                    {"Hg", 200.592},
                    {"Tl", 204.38},
                    {"Pb", 207.2},
                    {"Bi", 208.9804},
                    {"Po", 209},
                    {"At", 210},
                    {"Rn", 222},
                    {"Fr", 223},
                    {"Ra", 226},
                    {"Ac", 227},
                    {"Th", 232.0377},
                    {"Pa", 231.03588},
                    {"U", 238.02891},
                    {"Np", 237},
                    {"Pu", 244},
                    {"Am", 243},
                    {"Cm", 247},
                    {"Bk", 247},
                    {"Cf", 251},
                    {"Es", 252},
                    {"Fm", 257},
                    {"Uue", 315},
                    {"Ubn", 299}
                }

                Function Evaluate(s As String) As Double
                    s += "["
                    Dim sum = 0.0
                    Dim symbol = ""
                    Dim number = ""
                    For i = 1 To s.Length
                        Dim c = s(i - 1)
                        If "@" <= c AndAlso c <= "[" Then
                            ' @,A-Z
                            Dim n = 1
                            If number <> "" Then
                                n = Integer.Parse(number)
                            End If
                            If symbol <> "" Then
                                sum += atomicMass(symbol) * n
                            End If
                            If c = "[" Then
                                Exit For
                            End If
                            symbol = c.ToString
                            number = ""
                        ElseIf "a" <= c AndAlso c <= "z" Then
                            symbol += c
                        ElseIf "0" <= c AndAlso c <= "9" Then
                            number += c
                        Else
                            Throw New Exception(String.Format("Unexpected symbol {0} in molecule", c))
                        End If
                    Next
                    Return sum
                End Function

                Sub Main()
                    Dim molecules() As String = {
                        "H", "H2", "H2O", "H2O2", "(HO)2", "Na2SO4", "C6H12",
                        "COOH(C(CH3)2)3CH3", "C6H4O2(OH)4", "C27H46O", "Uue"
                    }
                    For Each molecule In molecules
                        Dim mass = Evaluate(ReplaceParens(molecule))
                        Console.WriteLine("{0,17} -> {1,7:0.000}", molecule, mass)
                    Next
                End Sub

                Function ReplaceFirst(text As String, search As String, replace As String) As String
                    Dim pos = text.IndexOf(search)
                    If pos < 0 Then
                        Return text
                    Else
                        Return text.Substring(0, pos) + replace + text.Substring(pos + search.Length)
                    End If
                End Function

                Function ReplaceParens(s As String) As String
                    Dim letter = "s"c
                    While True
                        Dim start = s.IndexOf("(")
                        If start = -1 Then
                            Exit While
                        End If

                        For i = start + 1 To s.Length - 1
                            If s(i) = ")" Then
                                Dim expr = s.Substring(start + 1, i - start - 1)
                                Dim symbol = String.Format("@{0}", letter)
                                s = ReplaceFirst(s, s.Substring(start, i + 1 - start), symbol)
                                atomicMass(symbol) = Evaluate(expr)
                                letter = Chr(Asc(letter) + 1)
                                Exit For
                            End If
                            If s(i) = "(" Then
                                start = i
                                Continue For
                            End If
                        Next
                    End While
                    Return s
                End Function

            End Class

            ''' <summary>
            ''' used to generate a payoff table and provide a set of outcomes (needs to edit inherit leaf
            ''' node as outcome)
            ''' </summary>
            Public Class PayOffTable

                'Decision matrix;

                'Risk and uncertainty,
                'Risk is where there are a number of possible outcomes and the probability of each outcome is known.
                'For example, based on past experience of digging for oil in a particular area,
                'an oil company may estimate that they have a 60% chance Of finding oil And a 40% chance Of Not finding oil.

                'Uncertainty
                'Uncertainty occurs when there are a number of possible outcomes but the probability of each outcome is not known.
                'For example, the same oil company may dig for oil in a previously unexplored area.
                'The company knows that it Is possible For them To either find Or Not find oil,
                'but it does Not know the probabilities Of Each Of these outcomes.

                'Once we know the different possible outcomes.
                'we can identify which decision Is best For a particular decision based On their risk aversion.

                'Decision Alternatives : Bonds, Stocks, Mutual Funds
                'States of nature : Growing, Stable, Declining

                Private mMaxiMaxDecision As New Outcome

                Private mMaxiMinDecision As New Outcome

                Private mMinMaxDecision As New Outcome

                Private mPayoffTable As New List(Of DecisionOption)

                ''' <summary>
                ''' Creates and empty class ready for manual population and interrogation
                ''' </summary>
                Public Sub New()

                End Sub

                ''' <summary>
                ''' Generates a Set of decisions (outcomes) as well as updating the payoff table supplied
                ''' with a regret table
                ''' </summary>
                ''' <param name="PayoffTable">a populated payoff table</param>
                Public Sub New(ByRef PayoffTable As List(Of DecisionOption))
                    mMinMaxDecision = MinMax(PayoffTable)
                    mMaxiMaxDecision = MaxiMax(PayoffTable)
                    mMaxiMinDecision = MaxiMin(PayoffTable)
                    mPayoffTable = DecisionOption.GenerateRegretTable(PayoffTable)
                End Sub

                ''' <summary>
                ''' maximax criterion (the optimistic (aggressive)) Select the course Of action with the
                ''' best return (maximum)
                ''' </summary>
                ''' <returns></returns>
                Public ReadOnly Property MaxiMaxDecision As Outcome
                    Get
                        Return mMaxiMaxDecision
                    End Get
                End Property

                ''' <summary>
                ''' maximin criterion (the pessimistic(conservative)) Select the course Of action whose
                ''' worst (maximum) loss Is better than the least (minimum) loss Of all other courses Of
                ''' action possible
                ''' </summary>
                ''' <returns></returns>
                Public ReadOnly Property MaxiMinDecision As Outcome
                    Get
                        Return mMaxiMinDecision
                    End Get
                End Property

                ''' <summary>
                ''' minimax regret criterion : Minimax (Opportunist) The minimization Of regret that Is
                ''' highest When one decision has been made instead Of another.
                ''' </summary>
                ''' <returns></returns>
                Public ReadOnly Property MinMaxDecision As Outcome
                    Get
                        Return mMinMaxDecision
                    End Get
                End Property

                ''' <summary>
                ''' Payoff table in the class
                ''' </summary>
                ''' <returns></returns>
                Public ReadOnly Property PayoffTable As List(Of DecisionOption)
                    Get
                        Return mPayoffTable
                    End Get

                End Property

                'maximax criterion (the optimistic (aggressive))
                'In decision theory, the optimistic (aggressive) decision making rule under conditions of uncertainty.
                'It states that the decision maker should Select the course Of action whose best (maximum)
                'gain Is better than the best gain Of all other courses Of action possible In given circumstances.
                ''' <summary>
                ''' The maximax payoff criterion: Is an optimistic payoff criterion. Using this criterion,
                ''' you Do the following
                ''' 1. Calculate the maximum payoff for each action.
                ''' 2. Choose the action that has the highest of these maximum payoffs.
                ''' </summary>
                ''' <param name="DecisionMatrix">
                ''' Payoff Table (list of decision options and thier potential outcomes)
                ''' </param>
                ''' <returns></returns>
                Public Function MaxiMax(ByRef DecisionMatrix As List(Of DecisionOption)) As Outcome
                    '1. Find the minimum payoff for each action.
                    Dim MaxOutcomes As New List(Of Outcome)
                    MaxOutcomes = DecisionOption.ReturnMaxOutcomes(DecisionMatrix)

                    '2. Choose the action that has the highest of these maximum payoffs.
                    Dim DeterminedOutcome As New Outcome
                    DeterminedOutcome = Outcome.SelectMaxPayoff(MaxOutcomes)

                    'Returns highest potential outcome of all options
                    Return DeterminedOutcome
                End Function

                'maximin criterion (the pessimistic(conservative))
                'In decision theory, the pessimistic (conservative) decision making rule under conditions of uncertainty.
                'It states that the decision maker should Select the course Of action whose worst (maximum)
                'loss Is better than the least (minimum) loss Of all other courses Of action possible In given circumstances
                ''' <summary>
                ''' The maximin payoff criterion: Is a pessimistic payoff criterion. Using this criterion,
                ''' you Do the following:
                ''' 1. Calculate the minimum payoff for each action.
                ''' 2. Choose the action that has the highest of these minimum payoffs.
                ''' </summary>
                ''' <param name="DecisionMatrix">
                ''' Payoff Table (list of decision options and thier potential outcomes)
                ''' </param>
                ''' <returns>highest minimum payoff</returns>
                Public Function MaxiMin(ByRef DecisionMatrix As List(Of DecisionOption)) As Outcome

                    '1. Find the minimum payoff for each action.
                    Dim MaxOutcomes As New List(Of Outcome)
                    MaxOutcomes = DecisionOption.ReturnMinOutcomes(DecisionMatrix)

                    '2. Choose the action that has the highest of these maximum payoffs.
                    Dim DeterminedOutcome As New Outcome
                    DeterminedOutcome = Outcome.SelectMaxPayoff(MaxOutcomes)

                    'Returns highest potential outcome of all options
                    Return DeterminedOutcome
                End Function

                'minimax regret criterion : Minimax (Opportunist)
                'In decision theory, (Opportunist) The minimization Of regret that Is highest When one decision has been made instead Of another.
                'In a situation In which a decision has been made that causes the expected payoff Of an Event To be less than expected,
                'this criterion encourages the avoidance Of regret. also called opportunity loss.
                ''' <summary>
                ''' minimax regret criterion : Minimax (Opportunist) The minimization Of regret that Is
                ''' highest When one decision has been made instead Of another. In a situation In which a
                ''' decision has been made that causes the expected payoff Of an Event To be less than
                ''' expected, this criterion encourages the avoidance Of regret. also called opportunity loss.
                ''' </summary>
                ''' <param name="DecisionMatrix"></param>
                ''' <returns></returns>
                Public Function MinMax(ByRef DecisionMatrix As List(Of DecisionOption)) As Outcome
                    Dim DeterminedOutcome As Outcome
                    'Generate regret table
                    Dim RegretMatrix As New List(Of DecisionOption)
                    RegretMatrix = DecisionOption.GenerateRegretTable(DecisionMatrix)
                    'Select MaxRegret
                    DeterminedOutcome = DecisionOption.SelectMaxRegret(RegretMatrix)
                    'Returns highest potential outcome of all options
                    Return DeterminedOutcome
                End Function

                ''' <summary>
                ''' Decision option and its list of potential outcomes and payoffs a list of these options
                ''' complete a payoff table
                ''' </summary>
                Public Structure DecisionOption

                    ''' <summary>
                    ''' Decision Alternative
                    ''' </summary>
                    Dim Name As String

                    ''' <summary>
                    ''' states of nature
                    ''' </summary>
                    Dim PotentialOutcomes As List(Of Outcome)

                    ''' <summary>
                    ''' Returns a Regret matrix (opportunity loss matrix)
                    ''' </summary>
                    ''' <param name="DecisionMatrix">Payoff table requiring Regret table</param>
                    ''' <returns>Regret Matrix / Opportunity loss payyoff table</returns>
                    Public Shared Function GenerateRegretTable(ByRef DecisionMatrix As List(Of DecisionOption)) As List(Of DecisionOption)
                        Dim Payoff As Integer = 0
                        Dim MaxPayoff As New Outcome
                        Dim Regret As Integer = 0

                        'Generate Regret Table
                        For Each decOption As DecisionOption In DecisionMatrix
                            'Find Max in row
                            MaxPayoff = Outcome.SelectMaxPayoff(decOption.PotentialOutcomes)
                            'Calculate regret for Row
                            For Each PotOutcome As Outcome In decOption.PotentialOutcomes
                                'Calc regret for outcome
                                Regret = MaxPayoff.Payoff - PotOutcome.Payoff
                                PotOutcome.Regret = Regret
                            Next
                        Next
                        'Return Regret table
                        Return DecisionMatrix
                    End Function

                    ''' <summary>
                    ''' Returns the highest payoffs for each row of table
                    ''' </summary>
                    ''' <param name="DecisionMatrix">Payoff table</param>
                    ''' <returns>Highest payoff outcomes for each row</returns>
                    Public Shared Function ReturnMaxOutcomes(ByRef DecisionMatrix As List(Of DecisionOption)) As List(Of Outcome)
                        Dim MaxOptionFound As New Outcome
                        Dim PotentialOutcomes As New List(Of Outcome)

                        'For each row
                        For Each DecOption As DecisionOption In DecisionMatrix
                            'select highest outcome in row
                            MaxOptionFound = Outcome.SelectMaxPayoff(DecOption.PotentialOutcomes)
                            'Add Row
                            PotentialOutcomes.Add(MaxOptionFound)
                        Next

                        'List of Highest Outcomes
                        Return PotentialOutcomes
                    End Function

                    ''' <summary>
                    ''' Returns the min outcomes for each Row
                    ''' </summary>
                    ''' <param name="DecisionMatrix">PayOff Table</param>
                    ''' <returns>List of lowest payoffs for each row</returns>
                    Public Shared Function ReturnMinOutcomes(ByRef DecisionMatrix As List(Of DecisionOption)) As List(Of Outcome)
                        Dim MinOptionFound As New Outcome
                        Dim PotentialOutcomes As New List(Of Outcome)

                        'For each row
                        For Each DecOption As DecisionOption In DecisionMatrix
                            'select Lowest outcome in row
                            MinOptionFound = Outcome.SelectMinPayoff(DecOption.PotentialOutcomes)
                            'Add Row
                            PotentialOutcomes.Add(MinOptionFound)
                        Next

                        'List of Lowest Outcomes
                        Return PotentialOutcomes
                    End Function

                    ''' <summary>
                    ''' Given a list of options Selects the outcome with the highest payoff Essentially a
                    ''' collection of rows
                    ''' </summary>
                    ''' <param name="DecisionMatrix">
                    ''' a list of decision options(rows), a decision is a row, each decision has a set of
                    ''' options, Each option has payoffs / Value
                    ''' </param>
                    ''' <returns>Single outcome (outcome with the highest payoff in the table of options</returns>
                    Public Shared Function SelectMaxPayoff(ByRef DecisionMatrix As List(Of DecisionOption)) As Outcome
                        'Temp outcome
                        Dim mDeterminedOutcome As New Outcome
                        'Highest outcome
                        Dim DeterminedOutcome As New Outcome
                        'highest payoff
                        Dim Payoff As Integer = 0

                        For Each DecOpt In DecisionMatrix
                            'Select Max payoff for Potential outcomes for option
                            mDeterminedOutcome = Outcome.SelectMaxPayoff(DecOpt.PotentialOutcomes)
                            'Check if maximum payoff is greater than existing payoff found
                            If mDeterminedOutcome.Payoff > Payoff = True Then
                                'Greater ?
                                DeterminedOutcome = mDeterminedOutcome
                            Else
                                'Must be lesser
                            End If
                            'Check next option
                        Next
                        'Returns highest potential outcome of all options
                        Return DeterminedOutcome
                    End Function

                    ''' <summary>
                    ''' Returns the Highest outcome in the regret table
                    ''' </summary>
                    ''' <param name="RegretMatrix"></param>
                    ''' <returns>single outcome (highest regret integer</returns>
                    Public Shared Function SelectMaxRegret(ByRef RegretMatrix As List(Of DecisionOption)) As Outcome
                        'Temp outcome
                        Dim mDeterminedOutcome As New Outcome
                        'Highest outcome
                        Dim DeterminedOutcome As New Outcome
                        'highest payoff
                        Dim Regret As Integer = 0

                        For Each DecOpt In RegretMatrix
                            'Select Max payoff for Potential outcomes for option
                            mDeterminedOutcome = Outcome.SelectMaxRegret(DecOpt.PotentialOutcomes)
                            'Check if maximum payoff is greater than existing payoff found
                            If mDeterminedOutcome.Regret > Regret = True Then
                                'Greater ?
                                DeterminedOutcome = mDeterminedOutcome
                            Else
                                'Must be lesser
                            End If
                            'Check next option
                        Next
                        'Returns highest potential outcome of all options
                        Return DeterminedOutcome
                    End Function

                    ''' <summary>
                    ''' Given a list of options Selects the outcome with the Lowest payoff Esentially a
                    ''' collection of rows
                    ''' </summary>
                    ''' <param name="DecisionMatrix">
                    ''' a list of decision options(rows), a decision is a row, each decision has a set of
                    ''' options, Each option has payoffs / Value
                    ''' </param>
                    ''' <returns>Single outcome (outcome with the Lowest payoff in the table of options</returns>
                    Public Shared Function SelectMinPayoff(ByRef DecisionMatrix As List(Of DecisionOption)) As Outcome
                        'Temp outcome
                        Dim mDeterminedOutcome As New Outcome
                        'Lowest outcome
                        Dim DeterminedOutcome As New Outcome
                        'highest payoff
                        Dim Payoff As Integer = 0

                        For Each DecOpt In DecisionMatrix
                            'Select Max payoff for Potential outcomes for option
                            mDeterminedOutcome = Outcome.SelectMinPayoff(DecOpt.PotentialOutcomes)
                            'Check if maximum payoff is greater than existing payoff found
                            If mDeterminedOutcome.Payoff < Payoff = True Then
                                'Greater ?
                                DeterminedOutcome = mDeterminedOutcome
                            Else
                                'Must be lesser
                            End If
                            'Check next option
                        Next
                        'Returns highest potential outcome of all options
                        Return DeterminedOutcome
                    End Function

                    ''' <summary>
                    ''' Adds Row to Payoff table supplied
                    ''' </summary>
                    ''' <param name="Name">Row name</param>
                    ''' <param name="DecisionOutcomes">Outcomes for row</param>
                    ''' <param name="PayoffTable">Table to be updated</param>
                    ''' <returns>updated payoff Table</returns>
                    Public Function AddRow(ByRef Name As String, ByRef DecisionOutcomes As List(Of Outcome), ByRef PayoffTable As List(Of DecisionOption)) As List(Of DecisionOption)
                        Dim NewDecisionOption As New DecisionOption
                        NewDecisionOption.Name = Name
                        NewDecisionOption.PotentialOutcomes = DecisionOutcomes
                        PayoffTable.Add(NewDecisionOption)

                        Return PayoffTable
                    End Function

                End Structure

                ''' <summary>
                ''' 'State of Nature(outcome) Outcome and its payoff. the amount (value of outcome) is used
                ''' to calculate the payoff
                ''' </summary>
                Public Structure Outcome

                    ''' <summary>
                    ''' Expected amount
                    ''' </summary>
                    Dim Amount As Integer

                    ''' <summary>
                    ''' State of Nature
                    ''' </summary>
                    Dim Name As String

                    ''' <summary>
                    ''' Expected PayOff
                    ''' </summary>
                    Dim Payoff As Integer

                    ''' <summary>
                    ''' Expected regret if chosen and other is more profitable
                    ''' </summary>
                    Dim Regret As Integer

                    ''' <summary>
                    ''' Given a list of potential outcomes select the outcome with the highest payoff
                    ''' Essentially a Single row in the table
                    ''' </summary>
                    ''' <param name="PotentialOutcomes">a list of outcomes</param>
                    ''' <returns>Single outcome (outcome with the highest payoff in the list of outcome</returns>
                    Public Shared Function SelectMaxPayoff(ByRef PotentialOutcomes As List(Of Outcome)) As Outcome
                        Dim DeterminedOutcome As New Outcome
                        Dim Payoff As Integer = 0

                        For Each PotOutcome In PotentialOutcomes

                            'Find Max Payoff
                            If Payoff < PotOutcome.Payoff = True Then
                                Payoff = PotOutcome.Payoff
                                'save payoff
                                DeterminedOutcome = PotOutcome
                            Else
                            End If

                        Next

                        Return DeterminedOutcome
                    End Function

                    Public Shared Function SelectMaxRegret(ByRef PotentialOutcomes As List(Of Outcome)) As Outcome
                        Dim DeterminedOutcome As New Outcome
                        Dim Regret As Integer = 0

                        For Each PotOutcome In PotentialOutcomes

                            'Find Max Regret
                            If Regret < PotOutcome.Regret = True Then
                                Regret = PotOutcome.Regret
                                'save payoff
                                DeterminedOutcome = PotOutcome
                            Else
                            End If

                        Next

                        Return DeterminedOutcome
                    End Function

                    ''' <summary>
                    ''' Given a list of potential outcomes select the outcome with the Lowest payoff
                    ''' Essentially a Single row in the table
                    ''' </summary>
                    ''' <param name="PotentialOutcomes">a list of outcomes</param>
                    ''' <returns>Single outcome (outcome with the Lowest payoff in the list of outcome</returns>
                    Public Shared Function SelectMinPayoff(ByRef PotentialOutcomes As List(Of Outcome)) As Outcome
                        Dim DeterminedOutcome As New Outcome
                        Dim Payoff As Integer = 0

                        For Each PotOutcome In PotentialOutcomes

                            'Find Min Payoff
                            If Payoff < PotOutcome.Payoff = True Then
                                Payoff = PotOutcome.Payoff
                                'save payoff
                                DeterminedOutcome = PotOutcome
                            Else
                            End If

                        Next

                        Return DeterminedOutcome
                    End Function

                    ''' <summary>
                    ''' Adds a new Outcome to the List of Outcomes supplied
                    ''' </summary>
                    ''' <param name="Name">Name of Outcome</param>
                    ''' <param name="PayOff">Payoff of O</param>
                    ''' <param name="DecisionOutcomes">List of Outcomes to have new Outcome added to</param>
                    ''' <returns>Updated List of Outcomes</returns>
                    Public Function AddOutcome(ByRef Name As String, ByRef PayOff As Integer, ByRef DecisionOutcomes As List(Of Outcome)) As List(Of Outcome)
                        Dim mOutcome = New Outcome
                        mOutcome.Name = Name
                        mOutcome.Payoff = PayOff
                        DecisionOutcomes.Add(mOutcome)
                        Return DecisionOutcomes
                    End Function

                End Structure

                'Payoff tables are a simple way of showing the different possible scenarios and their respective payoffs -
                'i.e. the profit or loss of each.
                'Maximax, maximin And minimax regret are different perspectives that can be applied to payoff tables.
            End Class

            Public Class Probability

                Public Shared Function BayesProbabilityOfAGivenB(ByRef ProbA As Integer, ProbB As Integer) As Integer
                    BayesProbabilityOfAGivenB = (ProbabilityofBgivenA(ProbA, ProbB) * ProbA) / ProbabilityOfData(ProbA, ProbB)
                End Function

                Public Shared Function OddsAgainst(ByRef NumberOfFavCases As Integer, ByRef NumberOfUnfavCases As Integer) As Integer
                    OddsAgainst = NumberOfUnfavCases / NumberOfFavCases
                End Function

                Public Shared Function OddsOf(ByRef NumberOfFavCases As Integer, ByRef NumberOfUnfavCases As Integer) As Integer
                    OddsOf = NumberOfFavCases / NumberOfUnfavCases
                End Function

                Public Shared Function ProbabilityOfAnB(ByRef ProbA As Integer, ProbB As Integer) As Integer
                    ProbabilityOfAnB = ProbA * ProbB
                End Function

                Public Shared Function ProbabilityofBgivenA(ByRef ProbA As Integer, ProbB As Integer) As Integer
                    ProbabilityofBgivenA = ProbabilityOfAnB(ProbA, ProbB) / ProbA
                End Function

                Public Shared Function ProbabilityofBgivenNotA(ByRef ProbA As Integer, ProbB As Integer) As Integer
                    ProbabilityofBgivenNotA = 1 - ProbabilityOfAnB(ProbA, ProbB) / ProbA
                End Function

                Public Shared Function ProbabilityOfData(ByRef ProbA As Integer, ProbB As Integer) As Integer
                    'P(D)
                    ProbabilityOfData = (ProbabilityofBgivenA(ProbA, ProbB) * ProbA) + (ProbabilityofBgivenNotA(ProbA, ProbB) * ProbabilityOfNot(ProbA))

                End Function

                Public Shared Function ProbabilityOfNot(ByRef Probablity As Integer) As Integer
                    ProbabilityOfNot = 1 - Probablity
                End Function

                Public Shared Function ProbablitiyOf(ByRef NumberOfFavCases As Integer, ByRef TotalNumberOfCases As Integer) As Integer
                    'P(H)
                    ProbablitiyOf = NumberOfFavCases / TotalNumberOfCases
                End Function

                Public Function FindProbabilityA(ByRef NumberofACases As Integer, ByRef TotalNumberOfCases As Integer) As Integer
                    Return ProbablitiyOf(NumberofACases, TotalNumberOfCases)
                End Function

                Public Function ProbabilityOfAandB(ByRef ProbabilityOFA As Integer, ByRef ProbabilityOFB As Integer) As Integer
                    Return ProbabilityOfAnB(ProbabilityOFA, ProbabilityOFB)
                End Function

                Public Function ProbabilityOfAGivenB(ByRef ProbabilityOFA As Integer, ByRef ProbabilityOFB As Integer) As Integer
                    Return BasicProbability.ProbabilityOfAGivenB(ProbabilityOFA, ProbabilityOFB)
                End Function

                'Author : Leroy S Dyer
                'Year : 2016
                Public Class BasicProbability

                    ''' <summary>
                    ''' P(AnB) probability of both true
                    ''' </summary>
                    ''' <param name="ProbA"></param>
                    ''' <param name="ProbB"></param>
                    ''' <returns></returns>
                    ''' <remarks></remarks>
                    Public Shared Function ProbabilityOfAandB(ByRef ProbA As Integer, ProbB As Integer) As Integer
                        'P(A&B) probability of both true
                        ProbabilityOfAandB = ProbA * ProbB
                    End Function

                    'Basic
                    ''' <summary>
                    ''' P(A | B) = P(B | A) * P(A) / P(B) the posterior probability, is the probability for
                    ''' A after taking into account B for and against A. P(A | B), a conditional
                    ''' probability, is the probability of observing event A given that B is true.
                    ''' </summary>
                    ''' <param name="ProbA"></param>
                    ''' <param name="ProbB"></param>
                    ''' <returns></returns>
                    ''' <remarks></remarks>
                    Public Shared Function ProbabilityOfAGivenB(ByRef ProbA As Integer, ProbB As Integer) As Integer
                        'P(A | B) = P(B | A) * P(A) / P(B)
                        'the posterior probability, is the probability for A after taking into account B for and against A.
                        'P(A | B), a conditional probability, is the probability of observing event A given that B is true.
                        ProbabilityOfAGivenB = ProbabilityofBgivenA(ProbA, ProbB) * ProbA / ProbB
                    End Function

                    ''' <summary>
                    ''' p(B|A) probability of B given A the conditional probability or likelihood, is the
                    ''' degree of belief in B, given that the proposition A is true. P(B | A) is the
                    ''' probability of observing event B given that A is true.
                    ''' </summary>
                    ''' <param name="ProbA"></param>
                    ''' <param name="ProbB"></param>
                    ''' <returns></returns>
                    ''' <remarks></remarks>
                    Public Shared Function ProbabilityofBgivenA(ByRef ProbA As Integer, ProbB As Integer) As Integer
                        'p(B|A) probability of B given A
                        'the conditional probability or likelihood, is the degree of belief in B, given that the proposition A is true.
                        'P(B | A) is the probability of observing event B given that A is true.
                        ProbabilityofBgivenA = ProbabilityOfAandB(ProbA, ProbB) / ProbA
                    End Function

                    ''' <summary>
                    ''' P(B|'A) Probability of Not A Given B the conditional probability or likelihood, is
                    ''' the degree of belief in B, given that the proposition A is false.
                    ''' </summary>
                    ''' <param name="ProbA"></param>
                    ''' <param name="ProbB"></param>
                    ''' <returns></returns>
                    ''' <remarks></remarks>
                    Public Shared Function ProbabilityofBgivenNotA(ByRef ProbA As Integer, ProbB As Integer) As Integer
                        'P(B|'A) Probability of Not A Given B
                        'the conditional probability or likelihood, is the degree of belief in B, given that the proposition A is false.
                        ProbabilityofBgivenNotA = 1 - ProbabilityOfAandB(ProbA, ProbB) / ProbA
                    End Function

                    ''' <summary>
                    ''' P('H) Probability of Not the case is the corresponding probability of the initial
                    ''' degree of belief against A: 1 − P(A) = P(−A)
                    ''' </summary>
                    ''' <param name="Probablity"></param>
                    ''' <returns></returns>
                    ''' <remarks></remarks>
                    Public Shared Function ProbabilityOfNot(ByRef Probablity As Integer) As Integer
                        'P('H) Probability of Not the case
                        'is the corresponding probability of the initial degree of belief against A: 1 − P(A) = P(−A)
                        ProbabilityOfNot = 1 - Probablity
                    End Function

                    ''' <summary>
                    ''' P(H) Probability of case the prior probability, is the initial degree of belief in A.
                    ''' </summary>
                    ''' <param name="NumberOfFavCases"></param>
                    ''' <param name="TotalNumberOfCases"></param>
                    ''' <returns></returns>
                    ''' <remarks></remarks>
                    Public Shared Function ProbablitiyOf(ByRef NumberOfFavCases As Integer, ByRef TotalNumberOfCases As Integer) As Integer
                        'P(H) Probability of case
                        'the prior probability, is the initial degree of belief in A.
                        ProbablitiyOf = NumberOfFavCases / TotalNumberOfCases
                    End Function

                End Class

                ''' <summary>
                ''' Naive Bayes where A and B are events. P(A) and P(B) are the probabilities of A and B
                ''' without regard to each other. P(A | B), a conditional probability, is the probability of
                ''' observing event A given that B is true. P(B | A) is the probability of observing event B
                ''' given that A is true. P(A | B) = P(B | A) * P(A) / P(B)
                ''' </summary>
                ''' <remarks></remarks>
                Public Class Bayes
                    'where A and B are events.
                    'P(A) and P(B) are the probabilities of A and B without regard to each other.
                    'P(A | B), a conditional probability, is the probability of observing event A given that B is true.
                    'P(B | A) is the probability of observing event B given that A is true.
                    'P(A | B) = P(B | A) * P(A) / P(B)

                    ''' <summary>
                    ''' ProbA
                    ''' </summary>
                    ''' <param name="NumberOfFavCases"></param>
                    ''' <param name="TotalNumberOfCases"></param>
                    ''' <returns></returns>
                    ''' <remarks></remarks>
                    Public Shared Function Likelihood(ByRef NumberOfFavCases As Integer, ByRef TotalNumberOfCases As Integer) As Integer
                        'ProbA
                        Likelihood = ProbailityOfData(NumberOfFavCases, TotalNumberOfCases)
                    End Function

                    ''' <summary>
                    ''' P(A | B) (Bayes Theorem) P(A | B) = P(B | A) * P(A) / P(B) the posterior
                    ''' probability, is the probability for A after taking into account B for and against A.
                    ''' P(A | B), a conditional probability, is the probability of observing event A given
                    ''' that B is true.
                    ''' </summary>
                    ''' <param name="prior"></param>
                    ''' <param name="Likelihood"></param>
                    ''' <param name="ProbabilityOfData"></param>
                    ''' <returns></returns>
                    ''' <remarks></remarks>
                    Public Shared Function Posterior(ByRef prior As Integer, Likelihood As Integer, ProbabilityOfData As Integer) As Integer
                        Posterior = prior * Likelihood / ProbabilityOfData
                    End Function

                    ''' <summary>
                    ''' P(B | A) * P(A)
                    ''' </summary>
                    ''' <param name="ProbA"></param>
                    ''' <param name="ProbB"></param>
                    ''' <returns></returns>
                    ''' <remarks></remarks>
                    Public Shared Function Prior(ByRef ProbA As Integer, ProbB As Integer) As Integer
                        'P(B | A) * P(A)
                        Prior = BasicProbability.ProbabilityofBgivenA(ProbA, ProbB)
                    End Function

                    ''' <summary>
                    ''' ProbB
                    ''' </summary>
                    ''' <param name="NumberOfFavCases"></param>
                    ''' <param name="TotalNumberOfCases"></param>
                    ''' <returns></returns>
                    ''' <remarks></remarks>
                    Public Shared Function ProbailityOfData(ByRef NumberOfFavCases As Integer, ByRef TotalNumberOfCases As Integer) As Integer
                        'ProbB
                        ProbailityOfData = BasicProbability.ProbablitiyOf(NumberOfFavCases, TotalNumberOfCases)
                    End Function

                End Class

            End Class

            ''' <summary>
            ''' Venn Diagrams (Set theory)
            ''' </summary>
            Public Class Venn

                Public GroupSet As List(Of VennSet)

                'Three Rules
                Public Function AIntersectB(ByRef A_CircleSet As VennSet, ByRef B_CircleSet As VennSet) As VennSet
                    Dim UnionSet As New VennSet
                    UnionSet.Name = A_CircleSet.Name & "Intersect" & B_CircleSet.Name

                    For Each ItemA As VennSet.VennItem In A_CircleSet.Items
                        For Each ItemB As VennSet.VennItem In B_CircleSet.Items
                            If ItemB.ItemName = ItemA.ItemName = True Then
                                UnionSet.Items.Add(ItemB)
                            End If
                        Next
                    Next
                    Return UnionSet
                End Function

                Public Function ANotB(ByRef A_CircleSet As VennSet, ByRef B_CircleSet As VennSet) As VennSet
                    Dim UnionSet As New VennSet
                    UnionSet.Name = A_CircleSet.Name & "Not" & B_CircleSet.Name
                    Dim found As Boolean
                    For Each ItemA As VennSet.VennItem In A_CircleSet.Items
                        found = False
                        For Each ItemB As VennSet.VennItem In B_CircleSet.Items
                            If ItemB.ItemName = ItemA.ItemName = True Then
                                found = True
                            Else

                            End If
                        Next
                        If found = False Then UnionSet.Items.Add(ItemA)
                    Next
                    Return UnionSet
                End Function

                Public Function AUnionB(ByRef A_CircleSet As VennSet, ByRef B_CircleSet As VennSet) As VennSet
                    Dim UnionSet As New VennSet
                    UnionSet.Name = A_CircleSet.Name & "Union" & B_CircleSet.Name
                    For Each ItemA As VennSet.VennItem In A_CircleSet.Items
                        UnionSet.Items.Add(ItemA)
                    Next
                    For Each Itemb As VennSet.VennItem In B_CircleSet.Items
                        UnionSet.Items.Add(Itemb)
                    Next
                    Return UnionSet
                End Function

                'Probability

                Private Function CountItemsInGroupSet(ByRef mGroupSet As List(Of VennSet)) As Integer
                    Dim count As Integer = 0
                    For Each mCircleset As VennSet In mGroupSet
                        For Each item As VennSet.VennItem In mCircleset.Items
                            count += 1
                        Next
                    Next
                    Return count
                End Function

                Private Function CountItemsInSet(ByRef mCircleSet As VennSet) As Integer
                    Dim count As Integer = 0
                    For Each item As VennSet.VennItem In mCircleSet.Items
                        count += 1
                    Next
                    Return count
                End Function

                Public Structure VennSet

                    Public Items As List(Of VennItem)

                    Public Name As String

                    Public Probability As Integer

                    Public Structure VennItem

                        Public ItemBoolean As Boolean
                        Public ItemName As String
                        Public ItemValue As Integer

                    End Structure

                End Structure

            End Class

        End Class

    End Namespace

End Namespace