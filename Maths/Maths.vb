Imports System.Drawing
Imports System.IO
Imports System.Web.Script.Serialization
Imports System.Windows.Forms
Imports System.Xml.Serialization

Namespace AI_SDK_EXTENSIONS

    Namespace Maths

        Namespace Cypher

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

        Namespace LocationalSpace

            ''' <summary>
            ''' Used as value for
            ''' Inputs and Outputs
            ''' </summary>
            <ComClass(Vector.ClassId, Vector.InterfaceId, Vector.EventsId)>
            Public Class Vector
                Public Const ClassId As String = "2828B490-7103-401C-BAB3-38FF97BC1AC9"
                Public Const EventsId As String = "CDBB4307-F15E-401A-AC6D-3CF8086FD6F1"
                Public Const InterfaceId As String = "8BB345F8-5113-4059-829B-B531310144B5"
                Public Values As New List(Of Double)

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

#Region "Export"

                ''' <summary>
                ''' deserialize object from Json
                ''' </summary>
                ''' <param name="Str">json</param>
                ''' <returns></returns>
                Public Shared Function FromJson(ByRef Str As String) As Vector
                    Try
                        Dim Converter As New JavaScriptSerializer
                        Dim diag As Vector = Converter.Deserialize(Of Vector)(Str)
                        Return diag
                    Catch ex As Exception
                        Dim Buttons As MessageBoxButtons = MessageBoxButtons.OK
                        MessageBox.Show(ex.Message, "ERROR", Buttons)
                    End Try
                    Return Nothing
                End Function

                ''' <summary>
                ''' Serializes object to json
                ''' </summary>
                ''' <returns> </returns>
                Public Function ToJson() As String
                    Dim Converter As New JavaScriptSerializer
                    Return Converter.Serialize(Me)
                End Function

                ''' <summary>
                ''' Serializes Object to XML
                ''' </summary>
                ''' <param name="FileName"></param>
                Public Sub ToXML(ByRef FileName As String)
                    Dim serialWriter As StreamWriter
                    serialWriter = New StreamWriter(FileName)
                    Dim xmlWriter As New XmlSerializer(Me.GetType())
                    xmlWriter.Serialize(serialWriter, Me)
                    serialWriter.Close()
                End Sub

#End Region

            End Class

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

        End Namespace

        <ComClass(MathFunctions.ClassId, MathFunctions.InterfaceId, MathFunctions.EventsId)>
        Public Class MathFunctions

            '
            Public Const ClassId As String = "2822E490-7703-401C-BAB3-38FF97BC1AC9"

            Public Const EventsId As String = "CDB54207-F55E-401A-AC6D-3CF8086FD6F1"
            Public Const InterfaceId As String = "8B3325F8-5D13-4059-829B-B531310144B5"

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

        End Class

        ''' <summary>
        ''' RandomFactory class generates different random numbers, characters, colors, etc..
        ''' </summary>
        <ComClass(RandomFactory.ClassId, RandomFactory.InterfaceId, RandomFactory.EventsId)>
        Public Class RandomFactory
            Public Const ClassId As String = "2828E190-7703-401C-BAB3-38FF97BC1AC9"
            Public Const EventsId As String = "CDB51307-F55E-401A-AC6D-3CF8086FD6F1"
            Public Const InterfaceId As String = "8B3341F8-5D13-4059-829B-B531310144B5"
            Private _Gen As System.Random

            ''' <summary>
            ''' Initializes Random generator variable
            ''' </summary>
            Public Sub New()
                _Gen = New System.Random()
            End Sub

            ''' <summary>
            ''' Generates random integer
            ''' </summary>
            ''' <returns></returns>
            Public Function GetRandomInt() As Integer
                Randomize()
                Return _Gen.Next()
            End Function

            ''' <summary>
            ''' Generates random integer less than Max
            ''' </summary>
            ''' <param name="max"></param>
            ''' <returns></returns>
            Public Function GetRandomInt(max As Integer) As Integer
                Randomize()
                If max < 1 Then max = 1
                Return _Gen.Next(max * 100) / 100
            End Function

            ''' <summary>
            ''' Generates Integer random between min and max
            ''' </summary>
            ''' <param name="min"></param>
            ''' <param name="max"></param>
            ''' <returns></returns>
            Public Function GetRandomInt(min As Integer, max As Integer) As Integer
                Randomize()
                Return _Gen.Next(min * 100, max * 100) / 100
            End Function

            ''' <summary>
            ''' Generates next double random between 0 and 1
            ''' </summary>
            ''' <returns></returns>
            Public Function GetRandomDbl() As Double
                Randomize()
                Return _Gen.NextDouble
            End Function

            ''' <summary>
            ''' Generates next double random between min and max
            ''' </summary>
            ''' <param name="min"></param>
            ''' <param name="max"></param>
            ''' <returns></returns>
            Public Function GetRandomDbl(min As Double, max As Double) As Double
                Randomize()
                Return CDbl(MathFunctions.Map(_Gen.NextDouble(), 0, 1, min, max))
            End Function

            ''' <summary>
            ''' Generates next single random between min and max
            ''' </summary>
            ''' <param name="min"></param>
            ''' <param name="max"></param>
            ''' <returns></returns>
            Public Function GetRandomSngl(min As Double, max As Double) As Double
                Randomize()

                Return MathFunctions.Map(CSng(_Gen.NextDouble()), 0, 1, min, max)
            End Function

            ''' <summary>
            ''' Generates next random color by generating 4 random integers Alpha, red, green and blue
            ''' </summary>
            ''' <returns></returns>
            Public Function GetRandomColor() As Color
                Dim MyAlpha As Integer
                Dim MyRed As Integer
                Dim MyGreen As Integer
                Dim MyBlue As Integer
                ' Initialize the random-number generator.
                Randomize()
                ' Generate random value between 1 and 6.
                MyAlpha = CInt(Int((254 * Rnd()) + 0))
                ' Initialize the random-number generator.
                Randomize()
                ' Generate random value between 1 and 6.
                MyRed = CInt(Int((254 * Rnd()) + 0))
                ' Initialize the random-number generator.
                Randomize()
                ' Generate random value between 1 and 6.
                MyGreen = CInt(Int((254 * Rnd()) + 0))
                ' Initialize the random-number generator.
                Randomize()
                ' Generate random value between 1 and 6.
                MyBlue = CInt(Int((254 * Rnd()) + 0))

                Return Color.FromArgb(MyAlpha, MyRed, MyGreen, MyBlue)
            End Function

            ''' <summary>
            ''' Generates next random character as per ASCII codes, from 32 to 122
            ''' </summary>
            ''' <returns></returns>
            Public Function GetRandomChar() As Char
                Randomize()
                Return GetRandomChar(32, 122)
            End Function

            ''' <summary>
            ''' Generates next random char as per ASCII between min to max
            ''' </summary>
            ''' <param name="min"></param>
            ''' <param name="max"></param>
            ''' <returns></returns>
            Public Function GetRandomChar(min As Integer, max As Integer) As Char
                'Return ChrW(GetRandomInt(min, max))
                Randomize()

                ' Store the numbers 1 to 6 in a list '
                Dim allNumbers As New List(Of Integer)(Enumerable.Range(min, max - min + 1))
                ' Store the randomly selected numbers in this list: '
                Dim selectedNumbers As New List(Of Integer)
                For i As Integer = 0 To allNumbers.Count - 1
                    ' A random index in numbers '
                    Dim index As Integer = _Gen.Next(0, allNumbers.Count)
                    ' Copy the item at index from allNumbers. '
                    Dim selectedNumber As Integer = allNumbers(index)
                    ' And store it in our list of picked numbers. '
                    selectedNumbers.Add(selectedNumber)
                    ' Remove the item from the list so that it cannot be picked again. '
                    allNumbers.RemoveAt(index)
                Next
                Return ChrW(selectedNumbers(0))
            End Function

            ''' <summary>
            '''   Equally likely to return true or false
            ''' </summary>
            ''' <returns></returns>
            Public Function NextBoolean() As Boolean
                Return _Gen.Next(2) > 0
            End Function

            ''' <summary>
            '''   Generates normally distributed numbers using Box-Muller transform by generating 2 random doubles
            '''   Gaussian noise is statistical noise having a probability density function (PDF) equal to that of the normal distribution,
            '''   which is also known as the Gaussian distribution.
            '''   In other words, the values that the noise can take on are Gaussian-distributed.
            ''' </summary>
            ''' <param name = "Mean">Mean of the distribution, default = 0</param>
            ''' <param name = "StdDeviation">Standard deviation, default = 1</param>
            ''' <returns></returns>
            Public Function NextGaussian(Optional ByVal Mean As Double = 0, Optional ByVal StdDeviation As Double = 1) As Double
                Dim X1 = _Gen.NextDouble()
                Dim X2 = _Gen.NextDouble()
                Dim StdDistribution = Math.Sqrt(-2.0 * Math.Log(X1)) * Math.Sin(2.0 * Math.PI * X2)
                Dim GaussianRnd = Mean + StdDeviation * StdDistribution

                Return GaussianRnd
            End Function

            ''' <summary>
            '''   Generates values from a triangular distribution
            '''   Triangular distribution is a continuous probability distribution with:
            '''       lower limit a
            '''       upper limit b
            '''       mode c
            '''   where a less than b
            '''   c is higher than or equal a but lessthan or equal b
            ''' </summary>
            ''' <param name = "min">Minimum</param>
            ''' <param name = "max">Maximum</param>
            ''' <param name = "mode">Mode (most frequent value)</param>
            ''' <returns></returns>
            Public Function NextTriangular(ByVal min As Double, ByVal max As Double, ByVal mode As Double) As Double
                Dim u = _Gen.NextDouble()

                If (u < (mode - min) / (max - min)) Then
                    Return min + Math.Sqrt(u * (max - min) * (mode - min))
                Else
                    Return max - Math.Sqrt((1 - u) * (max - min) * (max - mode))
                End If
            End Function

            ''' <summary>
            '''   Shuffles a list in O(n) time by using the Fisher-Yates/Knuth algorithm
            ''' </summary>
            ''' <param name = "list"></param>
            Public Sub Shuffle(ByVal list As IList)
                For i = 0 To list.Count - 1
                    Dim j = _Gen.Next(0, i + 1)

                    Dim temp = list(j)
                    list(j) = list(i)
                    list(i) = temp
                Next i
            End Sub

        End Class

        Namespace MachineLearning

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

        Namespace Science
            Public Module Conversion_Extensions

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

                'Morse Code
                Public MorseCode() As String = {".", "-"}

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

                'Temperture

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

            End Module

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

            End Class

        End Namespace

        Namespace MachineLearning

            ''' <summary>
            ''' Important Probability functions
            ''' </summary>
            Public Class Probability

                Public Function ProbabilityOfAandB(ByRef ProbabilityOFA As Integer, ByRef ProbabilityOFB As Integer) As Integer
                    Return ProbabilityOfAnB(ProbabilityOFA, ProbabilityOFB)
                End Function

                Public Function ProbabilityOfAGivenB(ByRef ProbabilityOFA As Integer, ByRef ProbabilityOFB As Integer) As Integer
                    Return BasicProbability.ProbabilityOfAGivenB(ProbabilityOFA, ProbabilityOFB)
                End Function

                Public Function FindProbabilityA(ByRef NumberofACases As Integer, ByRef TotalNumberOfCases As Integer) As Integer
                    Return ProbablitiyOf(NumberofACases, TotalNumberOfCases)
                End Function

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

        End Namespace

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

        Namespace Risk

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

        End Namespace

        Namespace Trees

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

            'Tree Extensions
            ''' <summary>
            ''' A Tree is actually a List of Lists
            ''' Root with branches and leaves:
            ''' In this tree the children are the branches: Locations to hold data have been provided.
            ''' These ae not part of the Tree Structure:
            ''' When initializing the structure Its also prudent to Initialize the ChildList;
            ''' Reducing errors; Due to late initialization of the ChildList
            ''' Subsequentially : Lists used in this class are not initialized in this structure.
            ''' Strings inserted with the insert Cmd will be treated as a Trie Tree insert!
            ''' if this tree requires data to be stroed it needs to be stored inside the dataStorae locations
            ''' </summary>
            <ComClass(TreeNode.ClassId, TreeNode.InterfaceId, TreeNode.EventsId)>
            Public Class TreeNode
                Public Const ClassId As String = "2822B490-7703-401C-BAB3-38FF97BC1AC9"
                Public Const EventsId As String = "CDB5B207-F55E-401A-ABBD-3CF8086BB6F1"
                Public Const InterfaceId As String = "8B3325F8-5D13-4059-8B9B-B531310144B5"

                ''' <summary>
                ''' Each Set of Nodes has only 26 ID's ID 0 = stop
                ''' </summary>
                Public Enum CharID
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

                Public Function ToJson() As String
                    Dim Converter As New JavaScriptSerializer
                    Return Converter.Serialize(Me)

                End Function

#Region "Add"

                ''' <summary>
                ''' Adds string to given trie
                ''' </summary>
                ''' <param name="Tree"></param>
                ''' <param name="Str"></param>
                ''' <returns></returns>
                Public Shared Function AddStringbyChars(ByRef Tree As TreeNode, Str As String) As TreeNode
                    Dim curr As TreeNode = Tree
                    Dim Pos As Integer = 0
                    For Each chr As Char In Str
                        Pos += 1
                        curr = TreeNode.AddStringToTrie(curr, chr, Pos)
                    Next
                    curr = TreeNode.AddStringToTrie(curr, "StopChar", Pos + 1)
                    Return Tree
                End Function

                ''' <summary>
                ''' Adds string to given trie
                ''' </summary>
                ''' <param name="Tree"></param>
                ''' <param name="Str"></param>
                ''' <returns></returns>
                Public Shared Function AddStringbyWord(ByRef Tree As TreeNode, Str As String) As TreeNode
                    Dim curr As TreeNode = Tree
                    Dim Pos As Integer = 0
                    For Each chr As String In Str.Split(" ")
                        Pos += 1
                        curr = TreeNode.AddStringToTrie(curr, chr, Pos)
                    Next
                    curr = TreeNode.AddStringToTrie(curr, "StopChar", Pos + 1)
                    Return Tree
                End Function

                ''' <summary>
                ''' Insert String into trie
                ''' </summary>
                ''' <param name="tree"></param>
                ''' <param name="Str"></param>
                ''' <returns></returns>
                Public Shared Function InsertByCharacters(ByRef tree As TreeNode, ByRef Str As String) As TreeNode
                    Return TreeNode.AddStringbyChars(tree, Str)
                End Function

                ''' <summary>
                ''' Insert String into trie
                ''' </summary>
                ''' <param name="tree"></param>
                ''' <param name="Str"></param>
                ''' <returns></returns>
                Public Shared Function InsertByWord(ByRef tree As TreeNode, ByRef Str As String) As TreeNode
                    Return TreeNode.AddStringbyWord(tree, Str)
                End Function

                ''' <summary>
                ''' Add characters Iteratively
                ''' CAT
                ''' AT
                ''' T
                ''' </summary>
                ''' <param name="Word"></param>
                ''' <param name="TrieTree"></param>
                Public Function AddItterativelyByCharacter(ByRef TrieTree As TreeNode, ByRef Word As String) As TreeNode
                    'AddWord
                    For i = 1 To Word.Length
                        TrieTree.InsertByCharacters(Word)
                        Word = Word.Remove(0, 1)
                    Next
                    Return TrieTree
                End Function

                ''' <summary>
                ''' AddIteratively characters
                ''' CAT
                ''' AT
                ''' T
                ''' </summary>
                ''' <param name="Word"></param>
                Public Function AddItterativelyByCharacter(ByRef Word As String) As TreeNode
                    Return AddItterativelyByCharacter(Me, Word)
                End Function

                ''' <summary>
                ''' Add characters Iteratively
                ''' CAT
                ''' AT
                ''' T
                ''' </summary>
                ''' <param name="Word"></param>
                ''' <param name="TrieTree"></param>
                Public Function AddItterativelyByWord(ByRef TrieTree As TreeNode, ByRef Word As String) As TreeNode
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
                ''' AddIteratively characters
                ''' CAT
                ''' AT
                ''' T
                ''' </summary>
                ''' <param name="Word"></param>
                Public Function AddItterativelyByWord(ByRef Word As String) As TreeNode
                    Return AddItterativelyByWord(Me, Word)
                End Function

                ''' <summary>
                ''' Inserts a string into the trie
                ''' </summary>
                ''' <param name="Str"></param>
                ''' <returns></returns>
                Public Function InsertByCharacters(ByRef Str As String) As TreeNode
                    Return TreeNode.AddStringbyChars(Me, Str)
                End Function

                ''' <summary>
                ''' Inserts a string into the trie
                ''' </summary>
                ''' <param name="Str"></param>
                ''' <returns></returns>
                Public Function InsertByWord(ByRef Str As String) As TreeNode
                    Return TreeNode.AddStringbyWord(Me, Str)
                End Function

                ''' <summary>
                ''' Adds char to Node(children) Returning the child
                ''' </summary>
                ''' <param name="CurrentNode">node containing children</param>
                ''' <param name="ChrStr">String to be added </param>
                ''' <param name="CharPos">this denotes the level of the node</param>
                ''' <returns></returns>
                Private Shared Function AddStringToTrie(ByRef CurrentNode As TreeNode, ByRef ChrStr As String, ByRef CharPos As Integer) As TreeNode
                    'start of tree
                    Dim Text As String = ChrStr
                    Dim returnNode As New TreeNode
                    Dim NewNode As New TreeNode
                    'Goto first node
                    'does this node have siblings
                    If TreeNode.HasChildren(CurrentNode) = True Then
                        'Check children
                        If TreeNode.CheckNodeExists(CurrentNode.Children, ChrStr) = False Then
                            'create a new node for char
                            NewNode = TreeNode.CreateNode(ChrStr, CurrentNode.NodeLevel + 1)
                            NewNode.NodeLevel = CharPos
                            'Add childnode
                            CurrentNode.Children.Add(NewNode)
                            returnNode = TreeNode.GetNode(CurrentNode.Children, ChrStr)
                        Else
                            returnNode = TreeNode.GetNode(CurrentNode.Children, ChrStr)
                        End If
                    Else
                        'If no silings then Create new node
                        'create a new node for char
                        NewNode = TreeNode.CreateNode(ChrStr, CurrentNode.NodeLevel + 1)
                        NewNode.NodeLevel = CharPos
                        'Add childnode
                        CurrentNode.Children.Add(NewNode)
                        returnNode = TreeNode.GetNode(CurrentNode.Children, ChrStr)
                    End If

                    Return returnNode
                End Function

#End Region

#Region "DATA"

                ''' <summary>
                ''' Theses are essentailly the branches of the Tree....
                ''' if there are no children then the node is a leaf.
                ''' Denoted by the StopCharacter NodeID. or "StopChar" NodeData.
                ''' </summary>
                Public Children As List(Of TreeNode)

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

                Public Sub New(ByRef IntializeChildren As Boolean)
                    If IntializeChildren = True Then
                        Children = New List(Of TreeNode)
                    Else
                    End If
                End Sub

                Public Sub New()

                End Sub

                ''' <summary>
                ''' Returns Number of Nodes
                ''' </summary>
                ''' <returns></returns>
                Public Shared Function NumberOfNodes(ByRef Tree As TreeNode) As Integer
                    Dim Count As Integer = 0
                    For Each child In Tree.Children
                        Count += child.NumberOfNodes
                    Next
                    Return Count
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
                ''' Checks if current node Has children
                ''' </summary>
                ''' <returns></returns>
                Public Function HasChildren() As Boolean
                    Return If(Me.Children.Count > 0 = True, True, False)
                End Function

                ''' <summary>
                ''' Returns deepest level
                ''' </summary>
                ''' <returns></returns>
                Public Function LowestLevel() As Integer
                    'Gets the level for node
                    Dim Level As Integer = Me.NodeLevel

                    'Recurses children
                    For Each child In Me.Children
                        If Level < child.LowestLevel = True Then
                            Level = child.LowestLevel
                        End If
                    Next
                    'The loop should finish at the lowest level
                    Return Level
                End Function

                'Functions
                ''' <summary>
                ''' Returns Number of Nodes
                ''' </summary>
                ''' <returns></returns>
                Public Function NumberOfNodes() As Integer
                    Dim Count As Integer = 0
                    For Each child In Me.Children
                        Count += child.NumberOfNodes
                    Next
                    Return Count
                End Function

                ''' <summary>
                ''' Checks if given node has children
                ''' </summary>
                ''' <param name="Node"></param>
                ''' <returns></returns>
                Private Shared Function HasChildren(ByRef Node As TreeNode) As Boolean
                    Return If(Node.Children.Count > 0 = True, True, False)
                End Function

#End Region

#Region "Create"

                ''' <summary>
                ''' Creates a Tree With an Empty Root node Called Root! Intializing the StartChar
                ''' </summary>
                ''' <returns></returns>
                Public Function Create() As TreeNode
                    Return TreeNode.MakeTrieTree()
                End Function

                ''' <summary>
                ''' If node does not exist in child node set it is added
                ''' if node already exists then no node is added a node ID is generated
                ''' </summary>
                ''' <param name="NodeData">Character to be added</param>
                Private Shared Function CreateNode(ByRef NodeData As String, ByRef Level As Integer) As TreeNode
                    'Create node - 2
                    Dim NewNode As New TreeNode(True)
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
                    Dim newnode As New TreeNode
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

                Private Shared Function MakeTrieTree() As TreeNode
                    Dim tree As New TreeNode(True)

                    tree.NodeData = "Root"
                    tree.NodeID = CharID.StartChar
                    Return tree
                End Function

#End Region

#Region "Display"

                ''' <summary>
                ''' Displays Contents of trie
                ''' </summary>
                ''' <returns></returns>
                Public Overrides Function ToString() As String

                    'TreeInfo
                    Dim Str As String = "NodeID " & Me.NodeID.ToString & vbNewLine &
            "Data: " & Me.NodeData & vbNewLine &
            "Node Level: " & NodeLevel.ToString & vbNewLine

                    'Data held
                    Dim mData As String = ": Data :" & vbNewLine &
               "Int: " & NodeInt & vbNewLine &
               "Str: " & NodeStr & vbNewLine &
               "Boolean: " & NodeBool & vbNewLine

                    'Lists
                    mData &= "StrLst: " & vbNewLine
                    For Each child In Me.NodeStrList
                        mData &= "Str: " & child.ToString() & vbNewLine
                    Next
                    mData &= "IntLst: " & vbNewLine
                    For Each child In Me.NodeIntList
                        mData &= "Str: " & child.ToString() & vbNewLine
                    Next
                    mData &= "BooleanList: " & vbNewLine
                    For Each child In Me.NodeBoolList
                        mData &= "Bool: " & child.ToString() & vbNewLine
                    Next

                    'Recurse Children
                    For Each child In Me.Children
                        Str &= "Child: " & child.ToString()
                    Next

                    Return Str
                End Function

                ''' <summary>
                ''' Returns a TreeViewControl with the Contents of the Trie:
                ''' </summary>
                Public Function ToView(ByRef Node As TreeNode) As System.Windows.Forms.TreeNode
                    Dim Nde As New System.Windows.Forms.TreeNode
                    Nde.Text = Node.NodeData.ToString.ToUpper &
                "(" & Node.NodeLevel & ")" & vbNewLine

                    For Each child In Me.Children
                        Nde.Nodes.Add(child.ToView)

                    Next
                    Return Nde
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

#End Region

#Region "FIND"

                ''' <summary>
                ''' checks if node exists in child nodes (used for trie trees (String added is the key and the data)
                ''' </summary>
                ''' <param name="Nodedata">Char string used as identifier</param>
                ''' <returns></returns>
                Public Shared Function CheckNodeExists(ByRef Children As List(Of TreeNode), ByRef Nodedata As String) As Boolean
                    'Check node does not exist
                    Dim found As Boolean = False
                    For Each mNode As TreeNode In Children
                        If mNode.NodeData = Nodedata Then
                            found = True
                        Else
                        End If
                    Next
                    Return found
                End Function

                ''' <summary>
                ''' Returns true if string is contined in trie (prefix) not as Word
                ''' </summary>
                ''' <param name="tree"></param>
                ''' <param name="Str"></param>
                ''' <returns></returns>
                Public Shared Function CheckPrefix(ByRef tree As TreeNode, ByRef Str As String) As Boolean
                    Str = Str.ToUpper
                    Dim CurrentNode As TreeNode = tree
                    Dim found As Boolean = False

                    Dim Pos As Integer = 0
                    For Each chrStr As Char In Str
                        Pos += 1

                        'Check Chars
                        If TreeNode.CheckNodeExists(CurrentNode.Children, chrStr) = True Then
                            CurrentNode = TreeNode.GetNode(CurrentNode.Children, chrStr)
                            found = True
                        Else
                            found = False
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
                Public Shared Function CheckSentence(ByRef tree As TreeNode, ByRef Str As String) As Boolean
                    Str = Str.ToUpper
                    Dim CurrentNode As TreeNode = tree
                    Dim found As Boolean = False
                    'Position in Characterstr
                    Dim Pos As Integer = 0
                    For Each chrStr As String In Str.Split(" ")
                        Pos += 1

                        'Check Chars
                        If TreeNode.CheckNodeExists(CurrentNode.Children, chrStr) = True Then
                            CurrentNode = TreeNode.GetNode(CurrentNode.Children, chrStr)
                            found = True
                        Else
                            'Terminated before end of Word
                            found = False
                        End If
                    Next

                    'Check for end of word marker
                    Return If(found = True, TreeNode.CheckNodeExists(CurrentNode.Children, "StopChar") = True, False)

                End Function

                ''' <summary>
                ''' Returns true if string is contined in trie (prefix) not as Word
                ''' </summary>
                ''' <param name="tree"></param>
                ''' <param name="Str"></param>
                ''' <returns></returns>
                Public Shared Function CheckSentPrefix(ByRef tree As TreeNode, ByRef Str As String) As Boolean
                    Dim CurrentNode As TreeNode = tree
                    Dim found As Boolean = False
                    Str = Str.ToUpper
                    Dim Pos As Integer = 0
                    For Each chrStr As String In Str.Split(" ")
                        Pos += 1

                        'Check Chars
                        If TreeNode.CheckNodeExists(CurrentNode.Children, chrStr) = True Then
                            CurrentNode = TreeNode.GetNode(CurrentNode.Children, chrStr)
                            found = True
                        Else
                            found = False
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
                Public Shared Function CheckWord(ByRef tree As TreeNode, ByRef Str As String) As Boolean
                    Str = Str.ToUpper
                    Dim CurrentNode As TreeNode = tree
                    Dim found As Boolean = False
                    'Position in Characterstr
                    Dim Pos As Integer = 0
                    For Each chrStr As Char In Str
                        Pos += 1

                        'Check Chars
                        If TreeNode.CheckNodeExists(CurrentNode.Children, chrStr) = True Then
                            CurrentNode = TreeNode.GetNode(CurrentNode.Children, chrStr)
                            found = True
                        Else
                            'Terminated before end of Word
                            found = False
                        End If
                    Next

                    'Check for end of word marker
                    Return If(found = True, TreeNode.CheckNodeExists(CurrentNode.Children, "StopChar") = True, False)

                End Function

                ''' <summary>
                ''' Returns Matched Node to sender (used to recures children)
                ''' </summary>
                ''' <param name="Tree"></param>
                ''' <param name="NodeData"></param>
                ''' <returns></returns>
                Public Shared Function GetNode(ByRef Tree As List(Of TreeNode), ByRef NodeData As String) As TreeNode
                    Dim Foundnode As New TreeNode
                    For Each item In Tree
                        If item.NodeData = NodeData Then
                            Foundnode = item
                        Else
                        End If
                    Next
                    Return Foundnode
                End Function

                ''' <summary>
                ''' Returns true if String is found as a prefix in trie
                ''' </summary>
                ''' <param name="Str"></param>
                ''' <returns></returns>
                Public Function FindPrefix(ByRef Str As String) As Boolean
                    Return TreeNode.CheckPrefix(Me, Str)
                End Function

                ''' <summary>
                ''' Returns true if string is found as word in trie
                ''' </summary>
                ''' <param name="Str"></param>
                ''' <returns></returns>
                Public Function FindSentence(ByRef Str As String) As Boolean
                    Return TreeNode.CheckSentence(Me, Str)
                End Function

                ''' <summary>
                ''' Returns true if String is found as a prefix in trie
                ''' </summary>
                ''' <param name="Str"></param>
                ''' <returns></returns>
                Public Function FindSentPrefix(ByRef Str As String) As Boolean
                    Return TreeNode.CheckSentPrefix(Me, Str)
                End Function

                ''' <summary>
                ''' Returns true if string is found as word in trie
                ''' </summary>
                ''' <param name="Str"></param>
                ''' <returns></returns>
                Public Function FindWord(ByRef Str As String) As Boolean
                    Return TreeNode.CheckWord(Me, Str)
                End Function

#End Region

            End Class

            'Tree Extensions
            ''' <summary>
            ''' A Tree is actually a List of Lists
            ''' Root with branches and leaves:
            ''' In this tree the children are the branches: Locations to hold data have been provided.
            ''' These ae not part of the Tree Structure:
            ''' When initializing the structure Its also prudent to Initialize the ChildList;
            ''' Reducing errors; Due to late initialization of the ChildList
            ''' Subsequentially : Lists used in this class are not initialized in this structure.
            ''' Strings inserted with the insert Cmd will be treated as a Trie Tree insert!
            ''' if this tree requires data to be stroed it needs to be stored inside the dataStorae locations
            ''' </summary>
            <ComClass(TrieTree.ClassId, TrieTree.InterfaceId, TrieTree.EventsId)>
            Public Class TrieTree
                Public Const ClassId As String = "28CCB490-7703-401C-BAB3-38FF97BC1AC9"
                Public Const EventsId As String = "CCB5B207-F55E-401A-ABBD-3CF80C6BB6F1"
                Public Const InterfaceId As String = "8B3325F8-5D13-4059-CB9B-B531C10C44B5"

                ''' <summary>
                ''' Each Set of Nodes has only 26 ID's ID 0 = stop
                ''' </summary>
                Public Enum CharID

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

                Public Function ToJson() As String
                    Dim Converter As New JavaScriptSerializer
                    Return Converter.Serialize(Me)

                End Function

#Region "Add"

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
                ''' Insert String into trie
                ''' </summary>
                ''' <param name="tree"></param>
                ''' <param name="Str"></param>
                ''' <returns></returns>
                Public Shared Function InsertByCharacters(ByRef tree As TrieTree, ByRef Str As String) As TrieTree
                    Return TrieTree.AddStringbyChars(tree, Str)
                End Function

                ''' <summary>
                ''' Insert String into trie
                ''' </summary>
                ''' <param name="tree"></param>
                ''' <param name="Str"></param>
                ''' <returns></returns>
                Public Shared Function InsertByWord(ByRef tree As TrieTree, ByRef Str As String) As TrieTree
                    Return TrieTree.AddStringbyWord(tree, Str)
                End Function

                ''' <summary>
                ''' Add characters Iteratively
                ''' CAT
                ''' AT
                ''' T
                ''' </summary>
                ''' <param name="Word"></param>
                ''' <param name="TrieTree"></param>
                Public Function AddItterativelyByCharacter(ByRef TrieTree As TrieTree, ByRef Word As String) As TrieTree
                    'AddWord
                    For i = 1 To Word.Length
                        TrieTree.InsertByCharacters(Word)
                        Word = Word.Remove(0, 1)
                    Next
                    Return TrieTree
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

                ''' <summary>
                ''' Add characters Iteratively
                ''' CAT
                ''' AT
                ''' T
                ''' </summary>
                ''' <param name="Word"></param>
                ''' <param name="TrieTree"></param>
                Public Function AddItterativelyByWord(ByRef TrieTree As TrieTree, ByRef Word As String) As TrieTree
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
                ''' Inserts a string into the trie
                ''' </summary>
                ''' <param name="Str"></param>
                ''' <returns></returns>
                Public Function InsertByCharacters(ByRef Str As String) As TrieTree
                    Return TrieTree.AddStringbyChars(Me, Str)
                End Function

                ''' <summary>
                ''' Inserts a string into the trie
                ''' </summary>
                ''' <param name="Str"></param>
                ''' <returns></returns>
                Public Function InsertByWord(ByRef Str As String) As TrieTree
                    Return TrieTree.AddStringbyWord(Me, Str)
                End Function

                ''' <summary>
                ''' Adds char to Node(children) Returning the child
                ''' </summary>
                ''' <param name="CurrentNode">node containing children</param>
                ''' <param name="ChrStr">String to be added </param>
                ''' <param name="CharPos">this denotes the level of the node</param>
                ''' <returns></returns>
                Private Shared Function AddStringToTrie(ByRef CurrentNode As TrieTree, ByRef ChrStr As String, ByRef CharPos As Integer) As TrieTree
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

#End Region

#Region "DATA"

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

                Public Sub New(ByRef IntializeChildren As Boolean)
                    If IntializeChildren = True Then
                        Children = New List(Of TrieTree)
                    Else
                    End If
                End Sub

                Public Sub New()

                End Sub

                ''' <summary>
                ''' Returns Number of Nodes
                ''' </summary>
                ''' <returns></returns>
                Public Shared Function NumberOfNodes(ByRef Tree As TrieTree) As Integer
                    Dim Count As Integer = 0
                    For Each child In Tree.Children
                        Count += child.NumberOfNodes
                    Next
                    Return Count
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
                ''' Checks if current node Has children
                ''' </summary>
                ''' <returns></returns>
                Public Function HasChildren() As Boolean
                    Return If(Me.Children.Count > 0 = True, True, False)
                End Function

                ''' <summary>
                ''' Returns deepest level
                ''' </summary>
                ''' <returns></returns>
                Public Function LowestLevel() As Integer
                    'Gets the level for node
                    Dim Level As Integer = Me.NodeLevel

                    'Recurses children
                    For Each child In Me.Children
                        If Level < child.LowestLevel = True Then
                            Level = child.LowestLevel
                        End If
                    Next
                    'The loop should finish at the lowest level
                    Return Level
                End Function

                'Functions
                ''' <summary>
                ''' Returns Number of Nodes
                ''' </summary>
                ''' <returns></returns>
                Public Function NumberOfNodes() As Integer
                    Dim Count As Integer = 0
                    For Each child In Me.Children
                        Count += child.NumberOfNodes
                    Next
                    Return Count
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

#Region "Create"

                ''' <summary>
                ''' Creates a Tree With an Empty Root node Called Root! Intializing the StartChar
                ''' </summary>
                ''' <returns></returns>
                Public Function Create() As TrieTree
                    Return TrieTree.MakeTrieTree()
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
                    Dim newnode As New TrieTree
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

                Private Shared Function MakeTrieTree() As TrieTree
                    Dim tree As New TrieTree(True)

                    tree.NodeData = "Root"
                    tree.NodeID = CharID.StartChar
                    Return tree
                End Function

#End Region

#Region "Display"

                ''' <summary>
                ''' Displays Contents of trie
                ''' </summary>
                ''' <returns></returns>
                Public Overrides Function ToString() As String

                    'TreeInfo
                    Dim Str As String = "NodeID " & Me.NodeID.ToString & vbNewLine &
                "Data: " & Me.NodeData & vbNewLine &
                "Node Level: " & NodeLevel.ToString & vbNewLine

                    'Data held
                    Dim mData As String = ": Data :" & vbNewLine &
                   "Int: " & NodeInt & vbNewLine &
                   "Str: " & NodeStr & vbNewLine &
                   "Boolean: " & NodeBool & vbNewLine

                    'Lists
                    mData &= "StrLst: " & vbNewLine
                    For Each child In Me.NodeStrList
                        mData &= "Str: " & child.ToString() & vbNewLine
                    Next
                    mData &= "IntLst: " & vbNewLine
                    For Each child In Me.NodeIntList
                        mData &= "Str: " & child.ToString() & vbNewLine
                    Next
                    mData &= "BooleanList: " & vbNewLine
                    For Each child In Me.NodeBoolList
                        mData &= "Bool: " & child.ToString() & vbNewLine
                    Next

                    'Recurse Children
                    For Each child In Me.Children
                        Str &= "Child: " & child.ToString()
                    Next

                    Return Str
                End Function

                ''' <summary>
                ''' Returns a TreeViewControl with the Contents of the Trie:
                ''' </summary>
                Public Function ToView(ByRef Node As TrieTree) As System.Windows.Forms.TreeNode
                    Dim Nde As New System.Windows.Forms.TreeNode
                    Nde.Text = Node.NodeData.ToString.ToUpper &
                    "(" & Node.NodeLevel & ")" & vbNewLine

                    For Each child In Me.Children
                        Nde.Nodes.Add(child.ToView)

                    Next
                    Return Nde
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

#End Region

#Region "FIND"

                ''' <summary>
                ''' checks if node exists in child nodes (used for trie trees (String added is the key and the data)
                ''' </summary>
                ''' <param name="Nodedata">Char string used as identifier</param>
                ''' <returns></returns>
                Public Shared Function CheckNodeExists(ByRef Children As List(Of TrieTree), ByRef Nodedata As String) As Boolean
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
                ''' Returns true if string is contined in trie (prefix) not as Word
                ''' </summary>
                ''' <param name="tree"></param>
                ''' <param name="Str"></param>
                ''' <returns></returns>
                Public Shared Function CheckPrefix(ByRef tree As TrieTree, ByRef Str As String) As Boolean
                    Str = Str.ToUpper
                    Dim CurrentNode As TrieTree = tree
                    Dim found As Boolean = False

                    Dim Pos As Integer = 0
                    For Each chrStr As Char In Str
                        Pos += 1

                        'Check Chars
                        If TrieTree.CheckNodeExists(CurrentNode.Children, chrStr) = True Then
                            CurrentNode = TrieTree.GetNode(CurrentNode.Children, chrStr)
                            found = True
                        Else
                            found = False
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
                Public Shared Function CheckSentence(ByRef tree As TrieTree, ByRef Str As String) As Boolean
                    Str = Str.ToUpper
                    Dim CurrentNode As TrieTree = tree
                    Dim found As Boolean = False
                    'Position in Characterstr
                    Dim Pos As Integer = 0
                    For Each chrStr As String In Str.Split(" ")
                        Pos += 1

                        'Check Chars
                        If TrieTree.CheckNodeExists(CurrentNode.Children, chrStr) = True Then
                            CurrentNode = TrieTree.GetNode(CurrentNode.Children, chrStr)
                            found = True
                        Else
                            'Terminated before end of Word
                            found = False
                        End If
                    Next

                    'Check for end of word marker
                    Return If(found = True, TrieTree.CheckNodeExists(CurrentNode.Children, "StopChar") = True, False)

                End Function

                ''' <summary>
                ''' Returns true if string is contined in trie (prefix) not as Word
                ''' </summary>
                ''' <param name="tree"></param>
                ''' <param name="Str"></param>
                ''' <returns></returns>
                Public Shared Function CheckSentPrefix(ByRef tree As TrieTree, ByRef Str As String) As Boolean
                    Dim CurrentNode As TrieTree = tree
                    Dim found As Boolean = False
                    Str = Str.ToUpper
                    Dim Pos As Integer = 0
                    For Each chrStr As String In Str.Split(" ")
                        Pos += 1

                        'Check Chars
                        If TrieTree.CheckNodeExists(CurrentNode.Children, chrStr) = True Then
                            CurrentNode = TrieTree.GetNode(CurrentNode.Children, chrStr)
                            found = True
                        Else
                            found = False
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
                    Dim Pos As Integer = 0
                    For Each chrStr As Char In Str
                        Pos += 1

                        'Check Chars
                        If TrieTree.CheckNodeExists(CurrentNode.Children, chrStr) = True Then
                            CurrentNode = TrieTree.GetNode(CurrentNode.Children, chrStr)
                            found = True
                        Else
                            'Terminated before end of Word
                            found = False
                        End If
                    Next

                    'Check for end of word marker
                    Return If(found = True, TrieTree.CheckNodeExists(CurrentNode.Children, "StopChar") = True, False)

                End Function

                ''' <summary>
                ''' Returns Matched Node to sender (used to recures children)
                ''' </summary>
                ''' <param name="Tree"></param>
                ''' <param name="NodeData"></param>
                ''' <returns></returns>
                Public Shared Function GetNode(ByRef Tree As List(Of TrieTree), ByRef NodeData As String) As TrieTree
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
                ''' Returns true if String is found as a prefix in trie
                ''' </summary>
                ''' <param name="Str"></param>
                ''' <returns></returns>
                Public Function FindPrefix(ByRef Str As String) As Boolean
                    Return TrieTree.CheckPrefix(Me, Str)
                End Function

                ''' <summary>
                ''' Returns true if string is found as word in trie
                ''' </summary>
                ''' <param name="Str"></param>
                ''' <returns></returns>
                Public Function FindSentence(ByRef Str As String) As Boolean
                    Return TrieTree.CheckSentence(Me, Str)
                End Function

                ''' <summary>
                ''' Returns true if String is found as a prefix in trie
                ''' </summary>
                ''' <param name="Str"></param>
                ''' <returns></returns>
                Public Function FindSentPrefix(ByRef Str As String) As Boolean
                    Return TrieTree.CheckSentPrefix(Me, Str)
                End Function

                ''' <summary>
                ''' Returns true if string is found as word in trie
                ''' </summary>
                ''' <param name="Str"></param>
                ''' <returns></returns>
                Public Function FindWord(ByRef Str As String) As Boolean
                    Return TrieTree.CheckWord(Me, Str)
                End Function

#End Region

            End Class

            ''' <summary>
            ''' Extension methods for Tri Tree Structure
            ''' </summary>
            Public Module TrieTreeExtensions

                ''' <summary>
                ''' Returns true if string is contined in trie (prefix) not as Word
                ''' </summary>
                ''' <param name="tree"></param>
                ''' <param name="Str"></param>
                ''' <returns></returns>
                <System.Runtime.CompilerServices.Extension()>
                Public Function CheckPrefix(ByRef tree As TrieTree, ByRef Str As String) As Boolean
                    Dim CurrentNode As TrieTree = tree
                    Dim found As Boolean = False

                    Dim Pos As Integer = 0
                    For Each chrStr As Char In Str
                        Pos += 1

                        'Check Chars
                        If TrieTree.CheckNodeExists(CurrentNode.Children, chrStr) = True Then
                            CurrentNode = TrieTree.GetNode(CurrentNode.Children, chrStr)
                            found = True
                        Else
                            found = False
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
                <System.Runtime.CompilerServices.Extension()>
                Public Function CheckWord(ByRef tree As TrieTree, ByRef Str As String) As Boolean
                    Dim CurrentNode As TrieTree = tree
                    Dim found As Boolean = False
                    'Position in Characterstr
                    Dim Pos As Integer = 0
                    For Each chrStr As Char In Str
                        Pos += 1

                        'Check Chars
                        If TrieTree.CheckNodeExists(CurrentNode.Children, chrStr) = True Then
                            CurrentNode = TrieTree.GetNode(CurrentNode.Children, chrStr)
                            found = True
                        Else
                            'Terminated before end of Word
                            found = False
                        End If
                    Next

                    'Check for end of word marker
                    Return If(found = True, TrieTree.CheckNodeExists(CurrentNode.Children, "StopChar") = True, False)

                End Function

                ''' <summary>
                ''' Returns Number of Nodes
                ''' </summary>
                ''' <returns></returns>
                <System.Runtime.CompilerServices.Extension()>
                Public Function NumberOfNodes(ByRef Tree As TrieTree) As Integer
                    Dim Count As Integer = 0
                    For Each child In Tree.Children
                        Count += child.NumberOfNodes
                    Next
                    Return Count
                End Function

                ''' <summary>
                ''' Returns Matched Node in children to sender (used to recurse children)
                ''' </summary>
                ''' <param name="Tree"></param>
                ''' <param name="NodeData"></param>
                ''' <returns></returns>
                <System.Runtime.CompilerServices.Extension()>
                Public Function GetNode(ByRef Tree As List(Of TrieTree), ByRef NodeData As String) As TrieTree
                    Dim Foundnode As New TrieTree
                    For Each item In Tree
                        If item.NodeData = NodeData Then
                            Foundnode = item
                        Else
                        End If
                    Next
                    Return Foundnode
                End Function

            End Module

        End Namespace

        Namespace Sets

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

        End Namespace

    End Namespace

End Namespace