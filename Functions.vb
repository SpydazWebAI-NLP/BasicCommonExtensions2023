Imports System.Drawing
Imports System.IO
Imports System.Text
Imports System.Threading
Imports System.Windows.Forms

Namespace AI_SDK_EXTENSIONS

    Namespace Functions

        ''' <summary>
        ''' Speech handling Extensions
        ''' </summary>
        Public Module Speech_Extensions

            ''' <summary>
            ''' Speaks Text
            ''' </summary>
            ''' <param name="Str"></param>
            <Runtime.CompilerServices.Extension()>
            Public Sub Speak(ByVal Str As String)
                Using synth As New System.Speech.Synthesis.SpeechSynthesizer()
                    synth.Speak(Str)
                End Using
            End Sub

        End Module
        ''' <summary>
        ''' Extensions to handle date and time
        ''' </summary>
        Public Module Time_Extensions

            ''' <summary>
            ''' Gest the elapsed seconds since the input DateTime
            '''</summary>
            ''' <param name="value">Input DateTime</param>
            ''' <returns>Returns a Double value with the elapsed seconds since the input DateTime</returns>
            ''' <example>
            ''' Double elapsed = dtStart.ElapsedSeconds();
            ''' </example>
            <Runtime.CompilerServices.Extension()>
            Public Function ElapsedSeconds(value As DateTime) As Double
                Return DateTime.Now.Subtract(value).TotalSeconds
            End Function

            ''' <summary>
            ''' Calculates the age based on today.
            ''' </summary>
            ''' <param name="dateOfBirth">The date of birth.</param>
            ''' <returns>The calculated age.</returns>
            <Runtime.CompilerServices.Extension()>
            Public Function CalculateAge(dateOfBirth As DateTime) As Integer
                Return dateOfBirth.CalculateAge(Now.[Date])
            End Function

            ''' <summary>
            ''' Calculates the age based on a passed reference date.
            ''' </summary>
            ''' <param name="dateOfBirth">The date of birth.</param>
            ''' <param name="referenceDate">The reference date to calculate on.</param>
            ''' <returns>The calculated age.</returns>
            <Runtime.CompilerServices.Extension()>
            Public Function CalculateAge(dateOfBirth As DateTime, referenceDate As DateTime) As Integer
                Dim years = referenceDate.Year - dateOfBirth.Year
                If referenceDate.Month < dateOfBirth.Month OrElse
            (referenceDate.Month = dateOfBirth.Month AndAlso
             referenceDate.Day < dateOfBirth.Day) Then
                    years -= 1
                End If
                Return years
            End Function

            ''' <summary>
            ''' returns Phase of the moon
            ''' </summary>
            ''' <param name="Day"></param>
            ''' <returns></returns>
            <Runtime.CompilerServices.Extension()>
            Public Function GetMoonPhase(ByVal Day As Date) As String
                Dim Year2 As Integer
                Dim Month2 As Integer
                Dim K1 As Integer
                Dim K2 As Integer
                Dim K3 As Integer
                Dim Julian As Integer
                Dim IP As Double
                Dim Age As Double
                Year2 = Day.Year - Math.Floor((12 - Day.Month) / 10)
                Month2 = Day.Month + 9
                If Month2 >= 12 Then Month2 = Month2 - 12
                K1 = Math.Floor(365.25 * (Year2 + 4712))
                K2 = Math.Floor(30.6 * Month2 + 0.5)
                K3 = Math.Floor(Math.Floor((Year2 / 100) + 49) * 0.75) - 38
                Julian = K1 + K2 + Day.Day + 59
                If (Julian > 2299160) Then
                    Julian -= K3
                End If
                IP = (Julian - 2451550.1) / 29.530588853
                IP -= Math.Floor(IP)
                If IP < 0 Then IP += 1
                Age = IP * 29.53
                If Age < 1.84566 Then
                    Return "New"
                ElseIf Age < 5.53699 Then
                    Return "Waxing Crescent"
                ElseIf Age < 9.22831 Then
                    Return "First Quarter"
                ElseIf Age < 12.91963 Then
                    Return "Waxing Gibbous"
                ElseIf Age < 16.61096 Then
                    Return "Full"
                ElseIf Age < 20.30228 Then
                    Return "Waning Gibbous"
                ElseIf Age < 23.99361 Then
                    Return "Last Quarter"
                ElseIf Age < 27.68493 Then
                    Return "Waning Crescent"
                Else : Return "New"
                End If
            End Function

            ''' <summary>
            ''' Used to convert numerical date to String Monday,tuesday etc.
            ''' </summary>
            ''' <param name="DayOfWeek"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            <Runtime.CompilerServices.Extension()>
            Public Function ExtractDay(ByVal DayOfWeek As Integer) As String
                'Returns String Version of date ie 0 = sun / 1 = mon
                ExtractDay = ""
                Select Case DayOfWeek

                    Case 0
                        ExtractDay = "Sunday"
                    Case 1
                        ExtractDay = "Monday"
                    Case 2
                        ExtractDay = "Tuesday"
                    Case 3
                        ExtractDay = "Wednesday"
                    Case 4
                        ExtractDay = "Thursday"
                    Case 5
                        ExtractDay = "Friday"
                    Case 6
                        ExtractDay = "Saturday"
                End Select

            End Function

            ''' <summary>
            ''' Used to convert numerical date to String January,February
            ''' </summary>
            ''' <param name="Month_Renamed"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            <Runtime.CompilerServices.Extension()>
            Public Function ExtractMonth(ByVal Month_Renamed As Integer) As String
                'Returns String Version of date
                ExtractMonth = ""
                Select Case Month_Renamed
                    Case 1
                        ExtractMonth = "January"
                    Case 2
                        ExtractMonth = "February"
                    Case 3
                        ExtractMonth = "March"
                    Case 4
                        ExtractMonth = "April"
                    Case 5
                        ExtractMonth = "May"
                    Case 6
                        ExtractMonth = "June"
                    Case 7
                        ExtractMonth = "July"
                    Case 8
                        ExtractMonth = "August"
                    Case 9
                        ExtractMonth = "September"
                    Case 10
                        ExtractMonth = "October"
                    Case 11
                        ExtractMonth = "November"
                    Case 12
                        ExtractMonth = "December"
                End Select
            End Function

            ''' <summary>
            ''' Used to convert numerical date to String Spring, Summer
            ''' </summary>
            ''' <param name="Quarter"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            <Runtime.CompilerServices.Extension()>
            Public Function ExtractSeason(ByVal Quarter As Integer) As String
                'Returns String Version of date
                ExtractSeason = ""
                Select Case Quarter
                    Case 1
                        ExtractSeason = "Spring"
                    Case 2
                        ExtractSeason = "Summer"
                    Case 3
                        ExtractSeason = "Autumn"
                    Case 4
                        ExtractSeason = "Winter"

                End Select

            End Function

            Public Function GetDate() As String
                'Call Time Function From WMI
                Dim objWMIService, colItems, objItem
                Const strComputer As String = "."
                objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
                colItems = objWMIService.ExecQuery("Select * from Win32_LocalTime")
                Dim My_DayOfWeek As String = ""
                Dim My_Day As String = ""
                Dim My_Month As String = ""
                Dim My_year As String = ""
                For Each objItem In colItems
                    My_DayOfWeek = ExtractDay(objItem.DayOfWeek) 'real day mon/tues/wed/thurs/fri/sat/sun
                    My_Day = objItem.Day 'day date
                    My_Month = ExtractMonth(objItem.Month) 'real month
                    My_year = objItem.year
                Next
                GetDate = My_DayOfWeek & " " & My_Day & " " & My_Month & " " & My_year
            End Function

            Public Function GetSeason() As String
                'Call Time Function From WMI
                Dim objWMIService, colItems, objItem
                Const strComputer As String = "."
                objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
                colItems = objWMIService.ExecQuery("Select * from Win32_LocalTime")
                Dim Season As String = ""
                For Each objItem In colItems
                    Season = ExtractSeason(objItem.Quarter) 'Seasons
                Next
                GetSeason = Season
            End Function

            Public Function GetTime() As String
                'Call Time Function From WMI
                Dim objWMIService, colItems, objItem
                Const strComputer As String = "."
                objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
                colItems = objWMIService.ExecQuery("Select * from Win32_LocalTime")
                Dim My_Hour As String = ""
                Dim My_Minute As String = ""
                Dim My_Second As String = ""
                For Each objItem In colItems
                    My_Hour = objItem.Hour 'Time
                    My_Minute = objItem.Minute 'Time
                    My_Second = objItem.Second 'time
                Next
                GetTime = My_Hour & ":" & My_Minute & ":" & My_Second
            End Function

            Public GetDateStr As String = GetDayStr & ": " & GetMonthStr & ": " & GetYearStr & " "
            Public GetDayStr As Integer = My.Computer.Clock.LocalTime.Day
            Public GetMonthStr As Integer = My.Computer.Clock.LocalTime.Month
            Public GetYearStr As Integer = My.Computer.Clock.LocalTime.Year
        End Module
        Public Module FileExtensions

            ''' <summary>
            ''' Loads Json file to new string
            ''' </summary>
            ''' <returns></returns>
            Public Function LoadJson() As String

                Dim Scriptfile As String = ""
                Dim Ofile As New OpenFileDialog
                With Ofile
                    .Filter = "Json files (*.Json)|*.Json"

                    If (.ShowDialog() = DialogResult.OK) Then
                        Scriptfile = .FileName
                    End If
                End With
                Dim txt As String = ""
                If Scriptfile IsNot "" Then

                    txt = My.Computer.FileSystem.ReadAllText(Scriptfile)
                    Return txt
                Else
                    Return txt
                End If

            End Function

            ''' <summary>
            ''' Writes the contents of an embedded file resource embedded as Bytes to disk.
            ''' EG: My.Resources.DefBrainConcepts.FileSave(Application.StartupPath and "\DefBrainConcepts.ACCDB")
            ''' </summary>
            ''' <param name="BytesToWrite">Embedded resource Name</param>
            ''' <param name="FileName">Save to file</param>
            ''' <remarks></remarks>
            <System.Runtime.CompilerServices.Extension()>
            Public Sub FileSave(ByVal BytesToWrite() As Byte, ByVal FileName As String)

                If IO.File.Exists(FileName) Then
                    IO.File.Delete(FileName)
                End If

                Dim FileStream As New System.IO.FileStream(FileName, System.IO.FileMode.OpenOrCreate)
                Dim BinaryWriter As New System.IO.BinaryWriter(FileStream)

                BinaryWriter.Write(BytesToWrite)
                BinaryWriter.Close()
                FileStream.Close()
            End Sub

            <Runtime.CompilerServices.Extension()>
            Public Sub AppendTextFile(ByRef Text As String, ByRef FileName As String)

                UpdateTextFileAs(FileName, Text)
            End Sub

            ''' <summary>
            ''' replaces text in file with supplied text
            ''' </summary>
            ''' <param name="PathName">Pathname of file to be updated</param>
            ''' <param name="_text">Text to be inserted</param>
            <System.Runtime.CompilerServices.Extension()>
            Public Sub UpdateTextFileAs(ByRef PathName As String, ByRef _text As String)
                Try
                    If File.Exists(PathName) = True Then File.Create(PathName).Dispose()
                    Dim alltext As String = _text
                    File.AppendAllText(PathName, alltext)
                Catch ex As Exception
                    MsgBox("Error: " & Err.Number & ". " & Err.Description, , NameOf(UpdateTextFileAs))
                End Try
            End Sub

            ''' <summary>
            ''' Creates saves text file as
            ''' </summary>
            ''' <param name="PathName">nfilename and path to file</param>
            ''' <param name="_text">text to be inserted</param>
            <System.Runtime.CompilerServices.Extension()>
            Public Sub SaveTextFileAs(ByRef PathName As String, ByRef _text As String)

                Try
                    If File.Exists(PathName) = True Then File.Create(PathName).Dispose()
                    Dim alltext As String = _text
                    File.WriteAllText(PathName, alltext)
                Catch ex As Exception
                    MsgBox("Error: " & Err.Number & ". " & Err.Description, , NameOf(SaveTextFileAs))
                End Try
            End Sub

            ''' <summary>
            ''' Opens text file and returns text
            ''' </summary>
            ''' <param name="PathName">URL of file</param>
            ''' <returns>text in file</returns>
            <System.Runtime.CompilerServices.Extension()>
            Public Function OpenTextFile(ByRef PathName As String) As String
                Dim _text As String = ""
                Try
                    If File.Exists(PathName) = True Then
                        _text = File.ReadAllText(PathName)
                    End If
                Catch ex As Exception
                    MsgBox("Error: " & Err.Number & ". " & Err.Description, , NameOf(SaveTextFileAs))

                End Try
                Return _text
            End Function

        End Module

        Public Class RES_SyS

            ''' <summary>
            ''' Checks if directory exists If it does not then it is created
            ''' </summary>
            ''' <param name="YourPath"></param>
            Public Sub ChkDir(ByRef YourPath As String)

                If (System.IO.Directory.Exists(YourPath)) = True Then
                Else

                    System.IO.Directory.CreateDirectory(YourPath)

                End If
            End Sub

        End Class

        Public Module ResSys

            ''' <summary>
            ''' Writes the contents of an embedded file resource embedded as Bytes to disk.
            ''' EG: My.Resources.DefBrainConcepts.FileSave(Application.StartupPath   "\DefBrainConcepts.ACCDB")
            ''' OVERWRITES
            ''' </summary>
            ''' <param name="BytesToWrite">Embedded resource Name - My.Resources(NAME)</param>
            ''' <param name="FileName">Save to file</param>
            ''' <remarks></remarks>
            <System.Runtime.CompilerServices.Extension()>
            Public Sub FileSave(ByVal BytesToWrite() As Byte, ByVal FileName As String)

                If IO.File.Exists(FileName) Then
                    IO.File.Delete(FileName)
                End If

                Dim FileStream As New System.IO.FileStream(FileName, System.IO.FileMode.OpenOrCreate)
                Dim BinaryWriter As New System.IO.BinaryWriter(FileStream)

                BinaryWriter.Write(BytesToWrite)
                BinaryWriter.Close()
                FileStream.Close()
            End Sub

        End Module

        ''' <summary>
        ''' Interface to Artificial intelligence SDK Capabilities are centralized here.
        ''' </summary>
        Public Class SySExtensions

            Public Shared Function CountSubstring(str As String, substr As String) As Integer
                Dim count As Integer = 0
                If (Len(str) > 0) And (Len(substr) > 0) Then
                    Dim p As Integer = InStr(str, substr)
                    Do While p <> 0
                        p = InStr(p + Len(substr), str, substr)
                        count += 1
                    Loop
                End If
                Return count
            End Function

            Public Shared Function GetDirSubDirectorys(ByVal directory As IO.DirectoryInfo, ByVal pattern As String) As List(Of String)
                Dim Files As New List(Of String)
                For Each file In directory.GetFiles(pattern)
                    Console.WriteLine(file.FullName)
                    Files.Add(file.FullName)
                Next
                For Each subDir In directory.GetDirectories
                    GetDirSubDirectorys(subDir, pattern)
                Next
                Return Files
            End Function

            Public Shared Function GetRandItemfromList(ByRef li As List(Of String)) As String
                Randomize()
                Return li.Item(Int(Rnd() * (li.Count - 1)))
            End Function

        End Class

        ''' <summary>
        ''' for working with wave forms
        ''' </summary>
        <ComClass(ModWave.ClassId, ModWave.InterfaceId, ModWave.EventsId)>
        Public Class ModWave
            Public Const ClassId As String = "2899E490-7702-401C-BAB3-38FF97BC14C9"
            Public Const EventsId As String = "CD994307-F53E-401A-AC6D-3CFDD86F46F1"
            Public Const InterfaceId As String = "8B4145F1-5D13-4059-829B-B531314144B5"

            Public Shared Sub OpenAndPlayWave()
                Dim OpenFileDialog As New OpenFileDialog
                OpenFileDialog.Filter = "PCM Riff/Wave Files(*.wav)|*.wav"
                If OpenFileDialog.ShowDialog = DialogResult.OK Then
                    Dim WaveFile As ModWave = ModWave.FromFile(OpenFileDialog.FileName)
                    WaveFile.Play(AudioPlayMode.Background)
                End If
            End Sub

            'By Paul Ishak
            'WAVE PCM soundfile format
            'The Canonical WAVE file format
            'As Described Here: https://ccrma.stanford.edu/courses/422/projects/WaveFormat/
            'The File's Header
            Public FileHeader As Header

            'Wave File's Format Sub Chunk
            Public FileFormatSubChunk As FormatSubChunk

            'Data Subchunk
            Public FileDataSubChunk As DataSubChunk

            'This structure is an optional parameter for creating a new wave file
            Public Structure WaveFileOptions
                Public SampleRate As WavSampleRate
                Public AudioFormat As Format
                Public BitsPerSample As BitsPerSample
                Public NumberOfChannels As NumberOfChannels
                Public FormatSize As FormatSize
                Public NumberOfSamples As UInt32
                Public Data As Byte()
            End Structure

            'These are the various structures in the wave file and their description
            '                                               DATATYPE          OFFSET        Endian           Description
            Structure Header
                Public Property ChunkID As Byte() '          Dword              0             Big            Contains the letters "RIFF" in ASCII form(0x52494646 big-endian form).
                Public Property ChunkSize As UInt32 '        Dword              4             Little         36 + SubChunk2Size, or more precisely: 4 + (8 + SubChunk1Size) + (8 + SubChunk2Size)
                Public Property Format As Byte() '           Dword              8             Big            Contains the letters "WAVE" in ASCII form (0x57415645 big-endian form).
            End Structure

            Structure FormatSubChunk
                Public Property Subchunk1ID As Byte() '      Dword              12            Big            Contains the letters "fmt "(0x666d7420 big-endian form).
                Public Property Subchunk1Size As UInt32 '    Dword              16            little         16 for PCM.  This is the size of the rest of the Subchunk which follows this number.
                Public Property AudioFormat As UInt16  '     Word               20            little         PCM = 1 (i.e. Linear quantization)Values other than 1 indicate some form of compression.
                Public Property NumChannels As UInt16 '      Word               22            little         Mono = 1, Stereo = 2, etc.
                Public Property SampleRate As UInt32 '       Dword              24            little         8000, 44100, etc.
                Public Property ByteRate As UInt32 '         Dword              28            little         == SampleRate * NumChannels * BitsPerSample/8
                Public Property BlockAlign As UInt16 '       Word               32            little         == NumChannels * BitsPerSample/8
                Public Property BitsPerSample As UInt16 '    Word               34            little         8 bits = 8, 16 bits = 16, etc.
            End Structure

            Structure DataSubChunk
                Public Property Subchunk2ID As Byte() '      Dword              36            Big            Contains the letters "data"(0x64617461 big-endian form).
                Public Property Subchunk2Size As UInt32 '    Dword              40            little         == NumSamples * NumChannels * BitsPerSample/8     This is the number of bytes in the data.
                Public Property Data As Byte() '             VariableLength     44            little         The actual sound data.
            End Structure

            Public Sub Open(ByVal FileName As String)
                If Not My.Computer.FileSystem.FileExists(FileName) Then Exit Sub
                Dim FileBytes() As Byte = My.Computer.FileSystem.ReadAllBytes(FileName)
                Try
                    Me.FileHeader.ChunkID = GetDataFromByteArray(FileBytes, 0, 0, 4)
                    Me.FileHeader.ChunkSize = BitConverter.ToUInt32(FileBytes, 4)
                    Me.FileHeader.Format = GetDataFromByteArray(FileBytes, 0, 8, 4)
                    Me.FileFormatSubChunk.Subchunk1ID = GetDataFromByteArray(FileBytes, 0, 12, 4)
                    Me.FileFormatSubChunk.Subchunk1Size = BitConverter.ToUInt32(FileBytes, 16)
                    Me.FileFormatSubChunk.AudioFormat = BitConverter.ToUInt16(FileBytes, 20)
                    Me.FileFormatSubChunk.NumChannels = BitConverter.ToUInt16(FileBytes, 22)
                    Me.FileFormatSubChunk.SampleRate = BitConverter.ToUInt32(FileBytes, 24)
                    Me.FileFormatSubChunk.ByteRate = BitConverter.ToUInt32(FileBytes, 28)
                    Me.FileFormatSubChunk.BlockAlign = BitConverter.ToUInt16(FileBytes, 32)
                    Me.FileFormatSubChunk.BitsPerSample = BitConverter.ToUInt16(FileBytes, 34)
                    Me.FileDataSubChunk.Subchunk2ID = GetDataFromByteArray(FileBytes, 0, 36, 4)
                    Me.FileDataSubChunk.Subchunk2Size = BitConverter.ToUInt32(FileBytes, 40)
                    Me.FileDataSubChunk.Data = GetDataFromByteArray(FileBytes, 0, 44, Me.FileDataSubChunk.Subchunk2Size)
                Catch
                    Throw New Exception("File Is Invalid or corrupt!")
                End Try
            End Sub

            Public Shared Function FromFile(ByVal FileName As String) As ModWave
                Dim Result As New ModWave
                If Not My.Computer.FileSystem.FileExists(FileName) Then Return Result
                Dim FileBytes() As Byte = My.Computer.FileSystem.ReadAllBytes(FileName)
                Try
                    Result.FileHeader.ChunkID = GetDataFromByteArray(FileBytes, 0, 0, 4)
                    Result.FileHeader.ChunkSize = BitConverter.ToUInt32(FileBytes, 4)
                    Result.FileHeader.Format = GetDataFromByteArray(FileBytes, 0, 8, 4)
                    Result.FileFormatSubChunk.Subchunk1ID = GetDataFromByteArray(FileBytes, 0, 12, 4)
                    Result.FileFormatSubChunk.Subchunk1Size = BitConverter.ToUInt32(FileBytes, 16)
                    Result.FileFormatSubChunk.AudioFormat = BitConverter.ToUInt16(FileBytes, 20)
                    Result.FileFormatSubChunk.NumChannels = BitConverter.ToUInt16(FileBytes, 22)
                    Result.FileFormatSubChunk.SampleRate = BitConverter.ToUInt32(FileBytes, 24)
                    Result.FileFormatSubChunk.ByteRate = BitConverter.ToUInt32(FileBytes, 28)
                    Result.FileFormatSubChunk.BlockAlign = BitConverter.ToUInt16(FileBytes, 32)
                    Result.FileFormatSubChunk.BitsPerSample = BitConverter.ToUInt16(FileBytes, 34)
                    Result.FileDataSubChunk.Subchunk2ID = GetDataFromByteArray(FileBytes, 0, 36, 4)
                    Result.FileDataSubChunk.Subchunk2Size = BitConverter.ToUInt32(FileBytes, 40)
                    Result.FileDataSubChunk.Data = GetDataFromByteArray(FileBytes, 0, 44, Result.FileDataSubChunk.Subchunk2Size)
                    Return Result
                Catch
                    Result = New ModWave
                    Return Result
                End Try
            End Function

            Public Shared Function FromArray(ByVal Bytes As Byte()) As ModWave
                Dim Result As New ModWave
                Dim FileBytes() As Byte = Bytes
                If Bytes.Count = 0 Then Return Result
                Try
                    Result.FileHeader.ChunkID = GetDataFromByteArray(FileBytes, 0, 0, 4)
                    Result.FileHeader.ChunkSize = BitConverter.ToUInt32(FileBytes, 4)
                    Result.FileHeader.Format = GetDataFromByteArray(FileBytes, 0, 8, 4)
                    Result.FileFormatSubChunk.Subchunk1ID = GetDataFromByteArray(FileBytes, 0, 12, 4)
                    Result.FileFormatSubChunk.Subchunk1Size = BitConverter.ToUInt32(FileBytes, 16)
                    Result.FileFormatSubChunk.AudioFormat = BitConverter.ToUInt16(FileBytes, 20)
                    Result.FileFormatSubChunk.NumChannels = BitConverter.ToUInt16(FileBytes, 22)
                    Result.FileFormatSubChunk.SampleRate = BitConverter.ToUInt32(FileBytes, 24)
                    Result.FileFormatSubChunk.ByteRate = BitConverter.ToUInt32(FileBytes, 28)
                    Result.FileFormatSubChunk.BlockAlign = BitConverter.ToUInt16(FileBytes, 32)
                    Result.FileFormatSubChunk.BitsPerSample = BitConverter.ToUInt16(FileBytes, 34)
                    Result.FileDataSubChunk.Subchunk2ID = GetDataFromByteArray(FileBytes, 0, 36, 4)
                    Result.FileDataSubChunk.Subchunk2Size = BitConverter.ToUInt32(FileBytes, 40)
                    Result.FileDataSubChunk.Data = GetDataFromByteArray(FileBytes, 0, 44, Result.FileDataSubChunk.Subchunk2Size)
                    Return Result
                Catch
                    Result = New ModWave
                    Return Result
                End Try
            End Function

            Public Function ToArray() As Byte()
                Dim Results As Byte() = Nothing
                Results = CombineArrays(FileHeader.ChunkID, BitConverter.GetBytes(FileHeader.ChunkSize))
                Results = CombineArrays(Results, FileHeader.Format)
                Results = CombineArrays(Results, FileFormatSubChunk.Subchunk1ID)
                Results = CombineArrays(Results, BitConverter.GetBytes(FileFormatSubChunk.Subchunk1Size))
                Results = CombineArrays(Results, BitConverter.GetBytes(FileFormatSubChunk.AudioFormat))
                Results = CombineArrays(Results, BitConverter.GetBytes(FileFormatSubChunk.NumChannels))
                Results = CombineArrays(Results, BitConverter.GetBytes(FileFormatSubChunk.SampleRate))
                Results = CombineArrays(Results, BitConverter.GetBytes(FileFormatSubChunk.ByteRate))
                Results = CombineArrays(Results, BitConverter.GetBytes(FileFormatSubChunk.BlockAlign))
                Results = CombineArrays(Results, BitConverter.GetBytes(FileFormatSubChunk.BitsPerSample))
                Results = CombineArrays(Results, FileDataSubChunk.Subchunk2ID)
                Results = CombineArrays(Results, BitConverter.GetBytes(FileDataSubChunk.Subchunk2Size))
                Results = CombineArrays(Results, FileDataSubChunk.Data)
                Return Results
            End Function

            Public Sub Play(ByVal AudioPlayMode As AudioPlayMode)
                Try
                    My.Computer.Audio.Play(ToArray, AudioPlayMode)
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End Sub

            Public Sub Save(ByVal FileName As String)
                My.Computer.FileSystem.WriteAllBytes(FileName, Me.ToArray, False)
            End Sub

            Public Function CombineArrays(ByVal Array1() As Byte, ByVal Array2() As Byte) As Byte()
                Dim AllResults(Array1.Length + Array2.Length - 1) As Byte
                Array1.CopyTo(AllResults, 0)
                Array2.CopyTo(AllResults, Array1.Length)
                Return AllResults
            End Function

            Private Shared Function GetDataFromByteArray(ByVal ByteArray As Byte(), ByVal BlockOffset As Long, ByVal RangeStartOffset As Long, ByVal DataLength As Long) As Byte()
                On Error Resume Next
                Dim AnswerL As New List(Of Byte)
                Dim Answer(0 To CInt((DataLength - 1))) As Byte
                Dim CurrentOffset As Long
                For I = 0 To UBound(ByteArray)
                    CurrentOffset = BlockOffset + I
                    If CurrentOffset >= RangeStartOffset Then
                        If CurrentOffset <= RangeStartOffset + DataLength Then
                            AnswerL.Add(ByteArray(I))
                        End If
                    End If
                Next
                Dim count As Integer = -1
                For Each bt As Byte In AnswerL
                    count = count + 1
                    Answer(count) = bt
                Next
                Return Answer
            End Function

            Sub New(ByVal Options As WaveFileOptions)
                Try
                    FileHeader.ChunkID = Encoding.ASCII.GetBytes("RIFF")
                    FileFormatSubChunk.Subchunk1Size = Options.FormatSize
                    FileFormatSubChunk.NumChannels = Options.NumberOfChannels
                    FileFormatSubChunk.BitsPerSample = Options.BitsPerSample
                    FileDataSubChunk.Subchunk2Size = CUInt(Options.NumberOfSamples * Options.NumberOfChannels * Options.BitsPerSample / 8)
                    FileHeader.ChunkSize = CUInt(4 + (8 + FileFormatSubChunk.Subchunk1Size) + (8 + FileDataSubChunk.Subchunk2Size))
                    FileHeader.Format = Encoding.ASCII.GetBytes("WAVE")
                    FileFormatSubChunk.Subchunk1ID = Encoding.ASCII.GetBytes("fmt ")
                    FileFormatSubChunk.AudioFormat = Options.AudioFormat
                    FileFormatSubChunk.SampleRate = Options.SampleRate
                    FileFormatSubChunk.ByteRate = CUInt(Options.SampleRate * Options.NumberOfChannels * Options.BitsPerSample / 8)
                    FileFormatSubChunk.BlockAlign = CUShort(Options.NumberOfChannels * Options.BitsPerSample / 8)
                    FileDataSubChunk.Subchunk2ID = Encoding.ASCII.GetBytes("data")
                    FileDataSubChunk.Data = Options.Data
                Catch ex As Exception
                End Try
            End Sub

            Sub New()

            End Sub

            Public Enum WavSampleRate As UInt32
                hz8000 = 8000
                hz11025 = 11025
                hz16000 = 16000
                hz22050 = 22050
                hz32000 = 32000
                hz44100 = 44100
                hz48000 = 48000
                hz96000 = 96000
                hz192000 = 192000
            End Enum

            Public Enum Format As UInt16
                Standard = 1
            End Enum

            Public Enum BitsPerSample As UInt16
                bps_8 = 8
                bps_16 = 16
                bps_32 = 32
                bps_64 = 64
                bps_128 = 128
                bps_256 = 256
            End Enum

            Public Enum NumberOfChannels As UInt16
                Mono = 1
                Stereo = 2
            End Enum

            Public Enum FormatSize As UInt32
                PCM = 16
            End Enum

        End Class

        ''' <summary>
        ''' Clipboard monitor
        ''' </summary>
        Public Class ReadClipboard

            ''' <summary>
            ''' on clipboard contents changed the text is sent here to be processed by the consuming class
            ''' </summary>
            ''' <param name="Text"></param>
            Public Event onClipboardChanged(Text As String)

            Public Sub New()
                mCurrentClipboardText = GetClipboardText()
            End Sub

            ''' <summary>
            ''' MONITORS CLIPBOARD FOR CHANGES "TEXT" RAISES THE EVENT
            ''' ONCLIPBOARDCHANGED
            ''' </summary>
            Private Sub StartMonitor()
                Do Until Started = False

                    If My.Computer.Clipboard.ContainsText = True Then
                        If GetClipboardText() <> mCurrentClipboardText Then
                            Dim EventA As String = GetClipboardText()

                            RaiseEvent onClipboardChanged(EventA)
                        End If
                    End If

                Loop
            End Sub

            ''' <summary>
            ''' READ CLIPBOARD: this is called to READ ALL the contents of the clipboard
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function GetClipboardText() As String
                mCurrentClipboardText = My.Computer.Clipboard.GetText
                Return mCurrentClipboardText
            End Function

            ''' <summary>
            ''' CURRENT TEXT HELD IN THE CLIPBOARD
            ''' </summary>
            Private mCurrentClipboardText As String = ""

            ''' <summary>
            ''' Allows for background worker
            ''' </summary>
            Private listening As New Thread(AddressOf StartMonitor)

            ''' <summary>
            ''' On/Off
            ''' </summary>
            Public Started As Boolean = False

            ''' <summary>
            ''' Starts Listener
            ''' </summary>
            Public Sub _start()
                Started = True
                listening.Start()

            End Sub

            ''' <summary>
            ''' Stops Listener
            ''' </summary>
            Public Sub _stop()
                listening.Abort()
                Started = False
            End Sub

        End Class

        Public Module ResourceExtensions
            Public InstalledApplicationPath As String = Application.StartupPath

            ''' <summary>
            ''' Checks if directory exists If it does not then it is created
            ''' </summary>
            ''' <param name="YourPath">Folder Path</param>
            <System.Runtime.CompilerServices.Extension()>
            Public Sub ChkDir(ByRef YourPath As String)

                If (System.IO.Directory.Exists(YourPath)) = True Then
                Else

                    System.IO.Directory.CreateDirectory(YourPath)

                End If
            End Sub

            ''' <summary>
            ''' Writes the contents of an embedded resource embedded as Bytes to disk.
            ''' </summary>
            ''' <param name="BytesToWrite">Embedded resource</param>
            ''' <param name="FileName">    Save to file</param>
            ''' <remarks></remarks>
            <System.Runtime.CompilerServices.Extension()>
            Public Sub FileSave(ByVal BytesToWrite() As Byte, ByVal FileName As String)

                If IO.File.Exists(FileName) Then
                    IO.File.Delete(FileName)
                End If

                Dim FileStream As New System.IO.FileStream(FileName, System.IO.FileMode.OpenOrCreate)
                Dim BinaryWriter As New System.IO.BinaryWriter(FileStream)

                BinaryWriter.Write(BytesToWrite)
                BinaryWriter.Close()
                FileStream.Close()
            End Sub

        End Module
        ''' <summary>
        ''' Various Useful Form Extensions
        ''' </summary>
        Public Module Form_Extensions

            <Runtime.CompilerServices.Extension()>
            Public Sub MoveFormTop(ByRef Sender As Form)
                Dim x As Integer
                Dim y As Integer
                x = Screen.PrimaryScreen.WorkingArea.Width - Sender.Width
                y = Screen.PrimaryScreen.WorkingArea.Height - Sender.Height
                'Top
                Sender.Location = New Point(Sender.Location.X, 0)
            End Sub

            <Runtime.CompilerServices.Extension()>
            Public Sub MoveFormBottom(ByRef Sender As Form)
                Dim x As Integer
                Dim y As Integer
                x = Screen.PrimaryScreen.WorkingArea.Width - Sender.Width
                y = Screen.PrimaryScreen.WorkingArea.Height - Sender.Height
                'Top
                Sender.Location = New Point(Sender.Location.X, y)
            End Sub

            <Runtime.CompilerServices.Extension()>
            Public Sub MoveFormLeft(ByRef Sender As Form)
                Dim x As Integer
                Dim y As Integer
                x = Screen.PrimaryScreen.WorkingArea.Width - Sender.Width
                y = Screen.PrimaryScreen.WorkingArea.Height - Sender.Height
                'left
                Sender.Location = New Point(0, Sender.Location.Y)
            End Sub

            <Runtime.CompilerServices.Extension()>
            Public Sub MoveFormRight(ByRef Sender As Form)
                Dim x As Integer
                Dim y As Integer
                x = Screen.PrimaryScreen.WorkingArea.Width - Sender.Width
                y = Screen.PrimaryScreen.WorkingArea.Height - Sender.Height
                'left
                Sender.Location = New Point(x, Sender.Location.Y)
            End Sub

            Public SyntaxTerms() As String = ({"SPYDAZ", "ABS", "ACCESS", "ADDITEM", "ADDNEW", "ALIAS", "AND", "ANY", "APP", "APPACTIVATE", "APPEND", "APPENDCHUNK", "ARRANGE", "AS", "ASC", "ATN", "BASE", "BEEP", "BEGINTRANS", "BINARY", "BYVAL", "CALL", "CASE", "CCUR", "CDBL", "CHDIR", "CHDRIVE", "CHR", "CHR$", "CINT", "CIRCLE", "CLEAR", "CLIPBOARD", "CLNG", "CLOSE", "CLS", "COMMAND", "
COMMAND$", "COMMITTRANS", "COMPARE", "CONST", "CONTROL", "CONTROLS", "COS", "CREATEDYNASET", "CSNG", "CSTR", "CURDIR$", "CURRENCY", "CVAR", "CVDATE", "DATA", "DATE", "DATE$", "DATESERIAL", "DATEVALUE", "DAY", "
DEBUG", "DECLARE", "DEFCUR", "CEFDBL", "DEFINT", "DEFLNG", "DEFSNG", "DEFSTR", "DEFVAR", "DELETE", "DIM", "DIR", "DIR$", "DO", "DOEVENTS", "DOUBLE", "DRAG", "DYNASET", "EDIT", "ELSE", "ELSEIF", "END", "ENDDOC", "ENDIF", "
ENVIRON$", "EOF", "EQV", "ERASE", "ERL", "ERR", "ERROR", "ERROR$", "EXECUTESQL", "EXIT", "EXP", "EXPLICIT", "FALSE", "FIELDSIZE", "FILEATTR", "FILECOPY", "FILEDATETIME", "FILELEN", "FIX", "FOR", "FORM", "FORMAT", "
FORMAT$", "FORMS", "FREEFILE", "FUNCTION", "GET", "GETATTR", "GETCHUNK", "GETDATA", "DETFORMAT", "GETTEXT", "GLOBAL", "GOSUB", "GOTO", "HEX", "HEX$", "HIDE", "HOUR", "IF", "IMP", "INPUT", "INPUT$", "INPUTBOX", "INPUTBOX$", "
INSTR", "INT", "INTEGER", "IS", "ISDATE", "ISEMPTY", "ISNULL", "ISNUMERIC", "KILL", "LBOUND", "LCASE", "LCASE$", "LEFT", "LEFT$", "LEN", "LET", "LIB", "LIKE", "LINE", "LINKEXECUTE", "LINKPOKE", "LINKREQUEST", "
LINKSEND", "LOAD", "LOADPICTURE", "LOC", "LOCAL", "LOCK", "LOF", "LOG", "LONG", "LOOP", "LSET", "LTRIM",
    "LTRIM$", "ME", "MID", "MID$", "MINUTE", "MKDIR", "MOD", "MONTH", "MOVE", "MOVEFIRST", "MOVELAST", "MOVENEXT", "MOVEPREVIOUS",
    "MOVERELATIVE", "MSGBOX", "NAME", "NEW", "NEWPAGE", "NEXT", "NEXTBLOCK", "NOT", "NOTHING",
    "NOW", "NULL", "OCT", "OCT$", "ON", "OPEN", "OPENDATABASE", "OPTION", "OR", "OUTPUT", "POINT", "PRESERVE", "PRINT",
    "PRINTER", "PRINTFORM", "PRIVATE", "PSET", "PUT", "PUBLIC", "QBCOLOR", "RANDOM", "RANDOMIZE", "READ", "REDIM", "REFRESH",
    "REGISTERDATABASE", "REM", "REMOVEITEM", "RESET", "RESTORE", "RESUME", "RETURN", "RGB", "RIGHT", "RIGHT$", "RMDIR", "RND",
    "ROLLBACK", "RSET", "RTRIM", "RTRIM$", "SAVEPICTURE", "SCALE", "SECOND", "SEEK", "SELECT", "SENDKEYS", "SET", "SETATTR",
    "SETDATA", "SETFOCUS", "SETTEXT", "SGN", "SHARED",
    "SHELL", "SHOW", "SIN", "SINGLE", "SPACE", "SPACE$", "SPC", "SQR",
    "STATIC", "STEP", "STOP", "STR", "STR$", "STRCOMP", "STRING", "STRING$", "SUB",
    "SYSTEM", "TAB", "TAN", "TEXT", "TEXTHEIGHT", "TEXTWIDTH", "THEN", "TIME", "TIME$",
    "TIMER", "TIMESERIAL", "TIMEVALUE", "TO", "TRIM",
    "TRIM$", "TRUE", "TYPE", "TYPEOF", "UBOUND", "UCASE", "UCASE$", "UNLOAD",
    "UNLOCK", "UNTIL", "UPDATE", "USING", "VAL", "VARIANT", "VARTYPE", "WEEKDAY", "WEND", "WHILE", "WIDTH",
    "WRITE", "XOR", "YEAR", "ZORDER", "ADDHANDLER", "ADDRESSOF", "ALIAS", "AND", "ANDALSO", "AS", "BYREF",
    "BOOLEAN", "BYTE", "BYVAL", "CALL", "CASE", "CATCH", "CBOOL", "CBYTE", "CCHAR", "CDATE",
    "CDEC", "CDBL", "CHAR", "CINT", "CLASS", "CLNG", "COBJ", "CONST", "CONTINUE", "CSBYTE",
    "CSHORT", "CSNG", "CSTR", "CTYPE", "CUINT", "CULNG", "CUSHORT", "DATE", "DECIMAL", "DECLARE",
    "DEFAULT", "DELEGATE", "DIM", "DIRECTCAST", "DOUBLE", "DO", "EACH", "ELSE", "ELSEIF", "END",
    "ENDIF", "ENUM", "ERASE", "ERROR", "EVENT", "EXIT", "FALSE", "FINALLY", "FOR", "FRIEND",
    "FUNCTION", "GET", "GETTYPE", "GLOBAL", "GOSUB", "GOTO", "HANDLES", "IF", "IMPLEMENTS",
    "IMPORTS", "IN", "INHERITS", "INTEGER", "INTERFACE", "IS", "ISNOT", "LET", "LIB", "LIKE",
    "LONG", "LOOP", "ME", "MOD", "MODULE", "MUSTINHERIT", "MUSTOVERRIDE", "MYBASE", "MYCLASS",
    "NAMESPACE", "NARROWING", "NEW", "NEXT", "NOT", "NOTHING", "NOTINHERITABLE", "NOTOVERRIDABLE",
    "OBJECT", "ON", "OF", "OPERATOR", "OPTION", "OPTIONAL", "OR", "ORELSE", "OVERLOADS",
    "OVERRIDABLE", "OVERRIDES", "PARAMARRAY", "PARTIAL", "PRIVATE", "PROPERTY", "PROTECTED",
    "PUBLIC", "RAISEEVENT", "READONLY", "REDIM", "REM", "REMOVEHANDLER", "RESUME", "RETURN",
    "SBYTE", "SELECT", "SET", "SHADOWS", "SHARED", "SHORT", "SINGLE", "STATIC", "STEP", "STOP",
    "STRING", "STRUCTURE", "SUB", "SYNCLOCK", "THEN", "THROW", "TO", "TRUE", "TRY", "TRYCAST",
    "TYPEOF", "WEND", "VARIANT", "UINTEGER", "ULONG", "USHORT", "USING", "WHEN", "WHILE", "WIDENING",
    "WITH", "WITHEVENTS", "WRITEONLY",
    "XOR", "#CONST", "#ELSE", "#ELSEIF", "#END", "#IF"})

            ''' <summary>
            ''' Highlights selection of text from index to length
            ''' </summary>
            ''' <param name="rtb">RichTextBox</param>
            ''' <param name="index">Start Index</param>
            ''' <param name="length">Charchcters width</param>
            ''' <param name="color">Color</param>
            <Runtime.CompilerServices.Extension()>
            Private Sub HighlightIndex(rtb As RichTextBox, index As Integer, length As Integer, color As Color)
                Dim selectionStartSave As Integer = rtb.SelectionStart 'to return this back to its original position
                rtb.SelectionStart = index
                rtb.SelectionLength = length
                rtb.SelectionColor = color
                rtb.SelectionLength = 0
                rtb.SelectionStart = selectionStartSave
                rtb.SelectionColor = rtb.ForeColor 'return back to the original color
            End Sub

            ''' <summary>
            ''' Uses internal syntax to Recolor VB.NEt Syntax Words
            ''' </summary>
            ''' <param name="sender">RichTextBox</param>
            <Runtime.CompilerServices.Extension()>
            Public Sub HighlightInternalSyntax(ByRef sender As RichTextBox)

                For Each Word As String In SyntaxTerms
                    HighlightWord(sender, Word)
                    HighlightWord(sender, LCase(Word))
                    HighlightWord(sender, UCase(Word))
                    HighlightWord(sender, StrConv(Word, vbProperCase))
                Next

            End Sub

            ''' <summary>
            ''' Highlights Specific Word in The textbox
            ''' </summary>
            ''' <param name="sender">RichTextBox</param>
            ''' <param name="word">Word to be foud</param>
            <Runtime.CompilerServices.Extension()>
            Public Sub HighlightWord(ByRef sender As RichTextBox, ByRef word As String)
                'Mark Cursor Point
                Dim selectionStartSave As Integer = sender.SelectionStart 'to return this back to its original position
                Dim index As Integer = 0
                While index < sender.Text.LastIndexOf(word)
                    'find
                    sender.Find(word, index, sender.TextLength, RichTextBoxFinds.WholeWord)
                    'select and color
                    sender.SelectionColor = Color.Blue
                    index = sender.Text.IndexOf(word, index) + 1
                End While
                'Return Cursor Position
                sender.SelectionStart = selectionStartSave

            End Sub

            ''' <summary>
            ''' Adds a child form to the sent form
            ''' </summary>
            ''' <param name="Sender"></param>
            ''' <param name="form"></param>
            <Runtime.CompilerServices.Extension()>
            Public Sub AddChildForm(ByRef Sender As System.Windows.Forms.Form, ByRef form As System.Windows.Forms.Form)
                Dim MdiChild = form
                MdiChild.MdiParent = Sender
                MdiChild.Show()

            End Sub

        End Module

    End Namespace

End Namespace