Namespace AI_SDK_EXTENSIONS

    Namespace MathsExt
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

    End Namespace

End Namespace