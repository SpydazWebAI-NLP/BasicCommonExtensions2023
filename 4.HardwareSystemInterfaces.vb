Imports System.Drawing
Imports System.IO
Imports System.Net
Imports System.Net.Sockets
Imports System.Runtime.Serialization
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Text
Imports System.Threading
Imports System.Windows.Forms

'You must add a reference to the newest DirectShowLib.dll file if this line has an error

'Imports DirectShowLib 'You must add a reference to the newest DirectShowLib.dll file if this line has an error

Namespace HardwareSystemInterfaces

    Namespace Sound

        ''' <summary>
        ''' for working with wave forms
        ''' </summary>
        <ComClass(ModWave.ClassId, ModWave.InterfaceId, ModWave.EventsId)>
        Public Class ModWave

#Region "Public Fields"

            Public Const ClassId As String = "2899E490-7702-401C-BAB3-38FF97BC14C9"
            Public Const EventsId As String = "CD994307-F53E-401A-AC6D-3CFDD86F46F1"
            Public Const InterfaceId As String = "8B4145F1-5D13-4059-829B-B531314144B5"

            'Data Subchunk
            Public FileDataSubChunk As DataSubChunk

            'Wave File's Format Sub Chunk
            Public FileFormatSubChunk As FormatSubChunk

            'By Paul Ishak
            'WAVE PCM soundfile format
            'The Canonical WAVE file format
            'As Described Here: https://ccrma.stanford.edu/courses/422/projects/WaveFormat/
            'The File's Header
            Public FileHeader As Header

#End Region

#Region "Public Constructors"

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

#End Region

#Region "Public Enums"

            Public Enum BitsPerSample As UInt16
                bps_8 = 8
                bps_16 = 16
                bps_32 = 32
                bps_64 = 64
                bps_128 = 128
                bps_256 = 256
            End Enum

            Public Enum Format As UInt16
                Standard = 1
            End Enum

            Public Enum FormatSize As UInt32
                PCM = 16
            End Enum

            Public Enum NumberOfChannels As UInt16
                Mono = 1
                Stereo = 2
            End Enum

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

#End Region

#Region "Public Methods"

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

            Public Shared Sub OpenAndPlayWave()
                Dim OpenFileDialog As New OpenFileDialog
                OpenFileDialog.Filter = "PCM Riff/Wave Files(*.wav)|*.wav"
                If OpenFileDialog.ShowDialog = DialogResult.OK Then
                    Dim WaveFile As ModWave = ModWave.FromFile(OpenFileDialog.FileName)
                    WaveFile.Play(AudioPlayMode.Background)
                End If
            End Sub

            Public Function CombineArrays(ByVal Array1() As Byte, ByVal Array2() As Byte) As Byte()
                Dim AllResults(Array1.Length + Array2.Length - 1) As Byte
                Array1.CopyTo(AllResults, 0)
                Array2.CopyTo(AllResults, Array1.Length)
                Return AllResults
            End Function

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

#End Region

#Region "Private Methods"

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

#End Region

#Region "Public Structs"

            Structure DataSubChunk

#Region "Public Properties"

                Public Property Data As Byte()
                Public Property Subchunk2ID As Byte() '      Dword              36            Big            Contains the letters "data"(0x64617461 big-endian form).
                Public Property Subchunk2Size As UInt32

#End Region

                '    Dword              40            little         == NumSamples * NumChannels * BitsPerSample/8     This is the number of bytes in the data.
                '             VariableLength     44            little         The actual sound data.
            End Structure

            Structure FormatSubChunk

#Region "Public Properties"

                Public Property AudioFormat As UInt16
                Public Property BitsPerSample As UInt16
                Public Property BlockAlign As UInt16
                Public Property ByteRate As UInt32

                '     Word               20            little         PCM = 1 (i.e. Linear quantization)Values other than 1 indicate some form of compression.
                Public Property NumChannels As UInt16

                '      Word               22            little         Mono = 1, Stereo = 2, etc.
                Public Property SampleRate As UInt32

                Public Property Subchunk1ID As Byte() '      Dword              12            Big            Contains the letters "fmt "(0x666d7420 big-endian form).
                Public Property Subchunk1Size As UInt32

#End Region

                '    Dword              16            little         16 for PCM.  This is the size of the rest of the Subchunk which follows this number.
                '       Dword              24            little         8000, 44100, etc.
                '         Dword              28            little         == SampleRate * NumChannels * BitsPerSample/8
                '       Word               32            little         == NumChannels * BitsPerSample/8
                '    Word               34            little         8 bits = 8, 16 bits = 16, etc.
            End Structure

            'These are the various structures in the wave file and their description
            '                                               DATATYPE          OFFSET        Endian           Description
            Structure Header

#Region "Public Properties"

                Public Property ChunkID As Byte() '          Dword              0             Big            Contains the letters "RIFF" in ASCII form(0x52494646 big-endian form).
                Public Property ChunkSize As UInt32 '        Dword              4             Little         36 + SubChunk2Size, or more precisely: 4 + (8 + SubChunk1Size) + (8 + SubChunk2Size)
                Public Property Format As Byte()

#End Region

                '           Dword              8             Big            Contains the letters "WAVE" in ASCII form (0x57415645 big-endian form).
            End Structure

            'This structure is an optional parameter for creating a new wave file
            Public Structure WaveFileOptions

#Region "Public Fields"

                Public AudioFormat As Format
                Public BitsPerSample As BitsPerSample
                Public Data As Byte()
                Public FormatSize As FormatSize
                Public NumberOfChannels As NumberOfChannels
                Public NumberOfSamples As UInt32
                Public SampleRate As WavSampleRate

#End Region

            End Structure

#End Region

        End Class

    End Namespace

    Namespace Camera

        Namespace TcpCameraInterface

            Public Class Camera

                Private Const WS_CHILD As Integer = &H40000000
                Private Const WS_VISIBLE As Integer = &H10000000
                Private Const SWP_NOMOVE As Short = &H2S
                Private Const SWP_NOZORDER As Short = &H4S
                Private Const SWP_NOSIZE As Short = 1
                Private Const HWND_BOTTOM As Short = 1
                Private Const WM_USER As Short = &H400S
                Private Const WM_CAP_FILE_SAVEAS As Integer = WM_CAP_START + 23
                Private Const WM_CAP_SET_SCALE As Integer = WM_CAP_START + 53
                Private Const WM_CAP_SINGLE_FRAME_OPEN As Integer = WM_CAP_START + 70
                Private Const WM_CAP_SINGLE_FRAME_CLOSE As Integer = WM_CAP_START + 71
                Private Const WM_CAP_SINGLE_FRAME As Integer = WM_CAP_START + 72
                Private Const WM_CAP_GRAB_FRAME As Integer = WM_CAP_START + 60
                Private Const WM_CAP_DLG_VIDEOSOURCE As Integer = WM_USER + 42
                Private Const WM_CAP_SET_OVERLAY As Integer = WM_USER + 51
                Private Const WM_CAP_SET_CALLBACK_CAPCONTROL As Integer = WM_CAP_START + 85
                Private Const WM_CAP_GRAB_FRAME_NOSTOP As Integer = WM_CAP_START + 61
                Private Const WM_CAP_FILE_SET_CAPTURE_FILE As Integer = WM_CAP_START + 20
                Private Const WM_CAP_DRIVER_CONNECT As Integer = WM_USER + 10
                Private Const WM_CAP_DRIVER_DISCONNECT As Integer = WM_USER + 11
                Private Const WM_CAP_SET_VIDEOFORMAT As Integer = WM_USER + 45
                Private Const WM_CAP_SET_PREVIEW As Integer = WM_USER + 50
                Private Const WM_CAP_SET_PREVIEWRATE As Integer = WM_USER + 52
                Private Const WM_CAP_GET_FRAME As Long = 1084
                Private Const WM_CAP_COPY As Long = 1054
                Private Const WM_CAP_START As Long = WM_USER
                Private Const WM_CAP_STOP As Long = (WM_CAP_START + 68)
                Private Const WM_CAP_SEQUENCE As Long = (WM_CAP_START + 62)
                Private Const WM_CAP_SET_SEQUENCE_SETUP As Long = (WM_CAP_START + 64)
                Private Const WM_CAP_FILE_SET_CAPTURE_FILEA As Long = (WM_CAP_START + 20)
                Private Const WM_CAP_SET_CALLBACK_FRAME As Integer = WM_CAP_START + 5
                Private Const WM_CAP_SAVEDIB As Integer = WM_CAP_START + 25
                Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hdcDest As IntPtr, ByVal nXDest As Integer, ByVal nYDest As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hdcSrc As IntPtr, ByVal nXSrc As Integer, ByVal nYSrc As Integer, ByVal dwRop As Int32) As Boolean
                Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Integer) As Boolean
                Private Declare Function SetWindowPos Lib "user32" Alias "SetWindowPos" (ByVal hWnd As Integer, ByVal hWndInsertAfter As Integer, ByVal x As Integer, ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer
                Private Declare Auto Function SendMessage Lib "user32" (ByVal hWnd As Integer, ByVal Msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
                Private Declare Function capCreateCaptureWindowA Lib "avicap32.dll" (ByVal lpszWindowName As String, ByVal dwStyle As Integer, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hWnd As Integer, ByVal nID As Integer) As Integer
                Private Declare Function capGetDriverDescriptionA Lib "avicap32.dll" (ByVal wDriver As Short, ByVal lpszName As String, ByVal cbName As Integer, ByVal lpszVer As String, ByVal cbVer As Integer) As Boolean

                Private VideoSource As Integer 'The camera ID
                Private PreHandle As Int32 'The preview control handle (picturebox usually)
                Private Handle As Integer 'The create preview window handle
                Private lwndC As Integer

                Public Running As Boolean

                Public CamFrameRate As Integer = 30
                Public OutputHeight As Integer = 240
                Public OutputWidth As Integer = 360

                ''' <summary>
                ''' Resets the camera.
                ''' </summary>
                ''' <remarks></remarks>
                Public Sub Reset()
                    If Running Then
                        Close()
                        Application.DoEvents()
                        Handle = capCreateCaptureWindowA(VideoSource.ToString, WS_VISIBLE Or WS_CHILD, 0, 0, 0, 0, PreHandle, 0)
                        If SetCamera() = False Then
                            MessageBox.Show("Error Setting/Re-Setting Camera", "Error Setting/Re-Setting Camera", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        End If
                    End If

                End Sub

                ''' <summary>
                ''' Starts up the camera.
                ''' </summary>
                ''' <param name="PreviewHandle">The handle to to the control to set the image on.</param>
                ''' <remarks></remarks>
                Public Sub Start(ByVal PreviewHandle As Int32)
                    If Running = True Then
                        MessageBox.Show("Camera is already running", "Camera is already running", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Exit Sub
                    Else
                        PreHandle = PreviewHandle
                        Handle = capCreateCaptureWindowA(VideoSource.ToString, WS_VISIBLE Or WS_CHILD, 0, 0, 0, 0, PreviewHandle, 0)
                        If SetCamera() = False Then MessageBox.Show("Error setting up camera", "Error setting up camera", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

                    End If
                End Sub

                ''' <summary>
                ''' Sets the camera frame rate.
                ''' </summary>
                ''' <param name="FrameRate">The rate to set.</param>
                ''' <remarks></remarks>
                Public Sub SetFrameRate(ByVal FrameRate As Integer)
                    CamFrameRate = FrameRate
                    Reset()
                End Sub

                Private Function SetCamera() As Boolean
                    If SendMessage(Handle, WM_CAP_DRIVER_CONNECT, VideoSource, 0) = 1 Then
                        SendMessage(Handle, WM_CAP_SET_SCALE, 1, 0)
                        SendMessage(Handle, WM_CAP_SET_PREVIEWRATE, CamFrameRate, 0)
                        SendMessage(Handle, WM_CAP_SET_PREVIEW, 1, 0)
                        SetWindowPos(Handle, HWND_BOTTOM, 0, 0, OutputWidth, OutputHeight, SWP_NOMOVE Or SWP_NOZORDER)
                        Running = True
                        Return True
                    Else
                        Running = False
                        Return False
                    End If
                End Function

                ''' <summary>
                ''' Closes the camera.
                ''' </summary>
                ''' <returns></returns>
                ''' <remarks></remarks>
                Public Function Close() As Boolean
                    If Running Then
                        Close = CBool(SendMessage(Handle, WM_CAP_DRIVER_DISCONNECT, 0, 0))
                        DestroyWindow(Handle)
                        Running = False
                    End If
                    Return Running
                End Function

                ''' <summary>
                ''' Shows the camera device settings.
                ''' </summary>
                ''' <remarks></remarks>
                Public Sub Settings()
                    SendMessage(Handle, WM_CAP_DRIVER_CONNECT, 0, 0)
                    SendMessage(Handle, WM_CAP_DLG_VIDEOSOURCE, 0, 0)
                End Sub

                ''' <summary>
                ''' Copies the current camera frame.
                ''' </summary>
                ''' <param name="src">The picturebox where the camera is being previewed.</param>
                ''' <param name="rect">The size of the the camera output.</param>
                ''' <returns></returns>
                ''' <remarks></remarks>
                Public Function CopyFrame(ByVal src As PictureBox, ByVal rect As RectangleF) As Bitmap
                    If Running Then
                        Dim srcPic As Graphics = src.CreateGraphics
                        Dim srcBmp As New Bitmap(src.Width, src.Height, srcPic)
                        Dim srcMem As Graphics = Graphics.FromImage(srcBmp)

                        Dim HDC1 As IntPtr = srcPic.GetHdc
                        Dim HDC2 As IntPtr = srcMem.GetHdc

                        BitBlt(HDC2, 0, 0, CInt(rect.Width),
              CInt(rect.Height), HDC1, CInt(rect.X), CInt(rect.Y), 13369376)

                        CopyFrame = CType(srcBmp.Clone(), Bitmap)

                        'Clean Up
                        srcPic.ReleaseHdc(HDC1)
                        srcMem.ReleaseHdc(HDC2)
                        srcPic.Dispose()
                        srcMem.Dispose()
                        Return CopyFrame
                    Else
                        MessageBox.Show("Camera is not running!", "Camera is not running", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    End If
                    Return Nothing
                End Function

            End Class

            Module Structures

                Public Enum Events
                    Connected = 0
                    Disconnected = 1
                    ImageArrived = 2
                    errEncounter = 3
                    OnConnection = 4
                    LostConnection = 5
                End Enum

                Public Structure EventArgs
                    Dim EventT As Events
                    Dim Arg As Object

                    Public Sub New(ByVal e As Events, Optional ByVal a As Object = Nothing)
                        EventT = e
                        Arg = a
                    End Sub

                End Structure

            End Module

            Public Class Host
                Dim serverSocket As TcpListener
                Dim clientSocket As TcpClient
                Dim netStream As NetworkStream
                Dim Context As SynchronizationContext
                Dim Listens As Boolean = False
                Dim PicBox As PictureBox

                Public Event OnConnection()

                Public Event LostConnection()

                Public Event errEncounter(ByVal ex As Exception)

                Public Sub New(ByVal Target As PictureBox, ByVal Port As Integer)
                    serverSocket = New TcpListener(System.Net.IPAddress.Any, Port)
                    Context = SynchronizationContext.Current()
                    PicBox = Target
                End Sub

                Public ReadOnly Property isListening() As Boolean
                    Get
                        Return Listens
                    End Get
                End Property

                Public Property SendBufferSize() As Integer
                    Get
                        Return serverSocket.Server.SendBufferSize
                    End Get
                    Set(ByVal value As Integer)
                        serverSocket.Server.SendBufferSize = value
                    End Set
                End Property

                Public Property ReceiveBufferSize() As Integer
                    Get
                        Return serverSocket.Server.ReceiveBufferSize
                    End Get
                    Set(ByVal value As Integer)
                        serverSocket.Server.ReceiveBufferSize = value
                    End Set
                End Property

                Public Property NoDelay() As Boolean
                    Get
                        Return serverSocket.Server.NoDelay
                    End Get
                    Set(ByVal value As Boolean)
                        serverSocket.Server.NoDelay = value
                    End Set
                End Property

                Public Sub StartConnection()
                    serverSocket.Start()
                    Dim t As New Thread(AddressOf Main)
                    t.Start()
                End Sub

                Public Sub CloseConnection()
                    On Error Resume Next
                    Listens = False
                    If Not clientSocket Is Nothing Then clientSocket.Close()
                    If Not netStream Is Nothing Then netStream.Close()
                    If Not serverSocket Is Nothing Then serverSocket.Stop()
                End Sub

                Private Sub Main()
                    Listens = True
GetConnection:  'Waits for an incoming connection
                    Try
                        clientSocket = serverSocket.AcceptTcpClient()
                        EventRaise(Events.OnConnection)
                        netStream = clientSocket.GetStream()
                    Catch ex As Exception
                        Throw

                    End Try

                    While Listens AndAlso clientSocket.Connected
                        Try
                            Dim Formatter As IFormatter = New BinaryFormatter()
                            Dim img As Image = Formatter.Deserialize(netStream)
                            PicBox.Image = img
                            PicBox.Invalidate()
                        Catch ex As Exception 'If exception had occured a user disconnected
                            EventRaise(Events.LostConnection)
                            GoTo GetConnection 'Wait for another use
                        End Try
                    End While
                    Context.Post(AddressOf CloseConnection, Nothing)
                End Sub

                Private Sub EventRaise(ByVal e As Events, Optional ByVal Arg As Object = Nothing)
                    Context.Post(AddressOf EventHandler, New EventArgs(e, Arg))
                End Sub

                Private Sub EventHandler(ByVal e As EventArgs)
                    Select Case e.EventT
                        Case Events.OnConnection
                            RaiseEvent OnConnection()
                        Case Events.LostConnection
                            RaiseEvent LostConnection()
                    End Select
                End Sub

                Public Function GetImage() As Image
                    Return PicBox.Image.Clone()
                End Function

            End Class

            Public Class Client
                Dim clientSocket As TcpClient
                Dim netStream As NetworkStream
                Dim Context As SynchronizationContext
                Dim ConnectedHost As Boolean = False
                Dim Cam As Camera
                Dim PicBox As PictureBox

                Public Event Connected()

                Public Event Disconnected()

                Public Event errEncounter(ByVal ex As Exception)

                Public Sub New(ByVal Target As PictureBox)
                    clientSocket = New TcpClient()
                    Context = SynchronizationContext.Current()
                    PicBox = Target
                    Cam = New Camera()
                End Sub

                Public Sub StartCamera()
                    Cam.Start(PicBox.Handle.ToInt32)
                End Sub

                Public Sub StopCamera()
                    Cam.Close()
                End Sub

                Public ReadOnly Property isCameraOn() As Boolean
                    Get
                        Return Cam.Running
                    End Get
                End Property

                Public Property CameraFPS() As Integer
                    Get
                        Return Cam.CamFrameRate
                    End Get
                    Set(ByVal value As Integer)
                        Cam.SetFrameRate(value)
                    End Set
                End Property

                Public Property CameraOutputSize() As Size
                    Get
                        Return New Size(Cam.OutputWidth, Cam.OutputHeight)
                    End Get
                    Set(ByVal value As Size)
                        Cam.OutputWidth = value.Width
                        Cam.OutputHeight = value.Height
                        Cam.Reset()
                    End Set
                End Property

                Public Sub CameraSettings()
                    Cam.Settings()
                End Sub

                Public ReadOnly Property isConnected() As Boolean
                    Get
                        Return ConnectedHost
                    End Get
                End Property

                Public Property SendBufferSize() As Integer
                    Get
                        Return clientSocket.SendBufferSize
                    End Get
                    Set(ByVal value As Integer)
                        clientSocket.SendBufferSize = value
                    End Set
                End Property

                Public Property ReceiveBufferSize() As Integer
                    Get
                        Return clientSocket.ReceiveBufferSize
                    End Get
                    Set(ByVal value As Integer)
                        clientSocket.ReceiveBufferSize = value
                    End Set
                End Property

                Public Property NoDelay() As Boolean
                    Get
                        Return clientSocket.NoDelay
                    End Get
                    Set(ByVal value As Boolean)
                        clientSocket.NoDelay = value
                    End Set
                End Property

                Public Sub Connect(ByVal IP As String, ByVal Port As Integer)
                    Try
                        clientSocket.Connect(IP, Port)
                    Catch ex As Exception
                        Throw

                    End Try
                    If clientSocket.Connected Then
                        RaiseEvent Connected()
                        Dim t As New Thread(AddressOf Main)
                        t.Start()
                    End If
                End Sub

                Private Sub Main()
                    ConnectedHost = True
                    netStream = clientSocket.GetStream()
                    While clientSocket.Connected And ConnectedHost
                        If Cam.Running Then
                            Try
                                Dim Bmp As Bitmap
                                Bmp = Cam.CopyFrame(PicBox, New RectangleF(0, 0, CameraOutputSize.Width, CameraOutputSize.Height))
                                Dim ms As New MemoryStream()
                                Bmp.Save(ms, Drawing.Imaging.ImageFormat.Jpeg) 'Change the format so we can pass a smaller image
                                Dim Img As Image = Image.FromStream(ms)
                                If SendImage(Img) = False Then Exit While
                                Thread.Sleep(1000 / CameraFPS) 'Decreasing the time the calculation were taken (So each frame delay will be true)
                            Catch ex As Exception
                                Exit While
                            End Try
                        End If
                    End While
                    Context.Post(AddressOf Disconnect, Nothing)
                End Sub

                Public Function SendImage(ByVal img As Image) As Boolean
                    Try
                        Dim Formatter As IFormatter = New BinaryFormatter()
                        Formatter.Serialize(netStream, img)
                        Return True
                    Catch ex As Exception
                        Return False
                    End Try
                End Function

                Public Sub Disconnect()
                    On Error Resume Next
                    If ConnectedHost Then
                        ConnectedHost = False
                        Dim SBufferSize, RBufferSize As Integer 'Get all of the current client socket properties so we can redeclare it with the previous settings (because it's being disposed)
                        Dim NoDelay As Boolean
                        SBufferSize = clientSocket.SendBufferSize
                        RBufferSize = clientSocket.ReceiveBufferSize
                        NoDelay = clientSocket.NoDelay
                        clientSocket.Client.Disconnect(False)
                        clientSocket.Close()
                        clientSocket = New TcpClient()
                        clientSocket.SendBufferSize = SBufferSize
                        clientSocket.ReceiveBufferSize = RBufferSize
                        clientSocket.NoDelay = NoDelay
                        RaiseEvent Disconnected()
                    End If
                End Sub

                Private Sub EventRaise(ByVal e As Events, ByVal Arg As Object)
                    Context.Post(AddressOf EventHandler, New EventArgs(e, Arg))
                End Sub

                Private Sub EventHandler(ByVal e As EventArgs)
                    Select Case e.EventT
                        Case Events.Connected
                            RaiseEvent Connected()
                        Case Events.Disconnected
                            RaiseEvent Disconnected()
                    End Select
                End Sub

                Public Function GetImage() As Image
                    Return Cam.CopyFrame(PicBox, New RectangleF(0, 0, CameraOutputSize.Width, CameraOutputSize.Height))
                End Function

            End Class

        End Namespace

    End Namespace

    Namespace Network

        Namespace Client

            ''' <summary>
            ''' Chat Client control used to Subscribe to TCP Chatservers operating on Specified port and Ip
            ''' address and send data.
            ''' </summary>
            ''' <remarks></remarks>
            <ComClass(ClientControl.ClassId, ClientControl.InterfaceId, ClientControl.EventsId)>
            Public Class ClientControl

                Private mClientConnected As Boolean = False
                Private mCurrentIP As String
                Private mCurrentPort As Integer

                Public Event ImageRecieved(ByRef Img As Image)

                Private Sub receiveImage()
                    Dim bf As New BinaryFormatter

                    Try
                        While CurrentClient.Connected = True
                            Dim nstream = CurrentClient.GetStream()
                            RaiseEvent ImageRecieved(bf.Deserialize(nstream))
                        End While
                    Catch ex As Exception
                        Throw
                    End Try

                End Sub

                Public Property GetDesktopStream As Boolean = False

                Public Sub RecieveDesktop()
                    Do While GetDesktopStream = True
                        receiveImage()
                    Loop
                End Sub

                ''' <summary>
                ''' Used to recive messages from the server
                ''' </summary>
                ''' <remarks></remarks>
                Private Sub RecieveMessages()
                    Try

                        Do While ClientConnected = True
                            'Capture Data
                            Dim NewClientdata As NetworkStream = CurrentClient.GetStream
                            Dim RecievedBytes(10000) As Byte
                            Try
                                'Decode Bytes
                                NewClientdata.Read(RecievedBytes, 0, 10000)
                                Dim ClientData As String = Encoding.ASCII.GetString(RecievedBytes)
                                RaiseEvent MessageRecieved(Me, ClientData)
                            Catch ex As Exception
                                Throw

                            End Try

                        Loop
                    Catch ex As Exception
                        Throw

                    End Try
                End Sub

                ''' <summary>
                ''' Occurs when THe client cannot connect to host.
                ''' </summary>
                ''' <remarks></remarks>
                Public Event ConnectionError(ByRef Sender As ClientControl, ByRef ex As Exception, ByRef Message As String)

                ''' <summary>
                ''' Occurs when Messages are recived from the server that the control is connected to.
                ''' </summary>
                ''' <remarks></remarks>
                Public Event MessageRecieved(ByRef Sender As ClientControl, ByRef Message As String)

                ''' <summary>
                ''' Occurs when messages are unable to be recived.
                ''' </summary>
                ''' <remarks></remarks>
                Public Event RecieverError(ByRef Sender As ClientControl, ByRef ex As Exception, ByRef Message As String)

                Public Const ClassId As String = "2828F490-7703-499C-BAB3-38FF97BC1AC9"
                Public Const EventsId As String = "05FE0F04-69F4-4AD8-849F-80BD2AFBFCED"
                Public Const InterfaceId As String = "8B3345F8-5D13-4059-829B-B531320144B5"

                ''' <summary>
                ''' Chat Client
                ''' </summary>
                ''' <remarks></remarks>
                Public CurrentClient As TcpClient

                Public Sub Sendbasic(Data As String)

                    Dim dataStream As StreamWriter
                    'Used to Write data
                    dataStream = New StreamWriter(CurrentClient.GetStream)
                    dataStream.Write(Data & vbCrLf)
                    dataStream.Flush()
                End Sub

                ''' <summary>
                ''' Sends data to Chatserver
                ''' </summary>
                ''' <param name="Data">Data string to be sent</param>
                ''' <remarks></remarks>
                Public Sub SendData(ByRef Data As String)
                    Dim dataStream As NetworkStream
                    'Used to Write data
                    dataStream = CurrentClient.GetStream
                    'Write
                    dataStream.Write(Encoding.ASCII.GetBytes(Data & vbCrLf), 0, Data.Length)
                    'Clean
                    dataStream.Flush()
                End Sub

                ''' <summary>
                ''' Closes CLient connection
                ''' </summary>
                ''' <remarks></remarks>
                Public Sub StopClient()
                    mClientConnected = False
                    CurrentClient.Close()

                End Sub

                ''' <summary>
                ''' Initializes a new instance of the <see cref="ClientControl"/> class.
                ''' </summary>
                ''' <param name="HostIP">Ip addres of Server to connected to</param>
                ''' <param name="HostPort">Port Server is listening on</param>
                ''' <remarks></remarks>
                Public Sub New(ByRef HostIP As String, ByRef HostPort As Integer)
                    Try
                        'Client(sender)
                        CurrentClient = New TcpClient(HostIP, HostPort)
                        mClientConnected = True
                    Catch ex As Exception
                        RaiseEvent ConnectionError(Me, ex, "Unable to Connect to Server" & vbNewLine &
                                       "IP: " & HostIP & vbNewLine &
                                       "PORT: " & HostPort)
                    End Try
                    Try
                        If ClientConnected = True Then
                            'Client(Reciever) Begin ServerMessage / Listener
                            Threading.ThreadPool.QueueUserWorkItem(AddressOf RecieveMessages)
                            Threading.ThreadPool.QueueUserWorkItem(AddressOf RecieveDesktop)
                        End If
                    Catch ex As Exception
                        RaiseEvent RecieverError(Me, ex, "Unable to Recive messages")
                    End Try

                End Sub

                ''' <summary>
                ''' Gets a value indicating whether Client is connected to a server.
                ''' </summary>
                ''' <value><see langword="true"/> if ; otherwise, <see langword="false"/>.</value>
                ''' <remarks></remarks>
                Public ReadOnly Property ClientConnected As Boolean
                    Get
                        Return mClientConnected
                    End Get
                End Property

                ''' <summary>
                ''' Gets Current Ip of Control.
                ''' </summary>
                ''' <value></value>
                ''' <remarks></remarks>
                Public ReadOnly Property CurrentIP As String
                    Get
                        Return mCurrentIP
                    End Get
                End Property

                ''' <summary>
                ''' Gets Current Port of Control.
                ''' </summary>
                ''' <value></value>
                ''' <remarks></remarks>
                Public ReadOnly Property CurrentPort As Integer
                    Get
                        Return mCurrentPort
                    End Get
                End Property

            End Class

        End Namespace

        Namespace Server

            <ComClass(ServerControl.ClassId, ServerControl.InterfaceId, ServerControl.EventsId)>
            Public Class ServerControl

                'Generate new server
                Private CurrentServer As TcpListener

                Private mCurrentConnectedClients As New List(Of TcpClient)

                'Server Config Properties
                Private mServerIP As IPAddress = IPAddress.Parse("127.0.0.1")

                Private mServerIsListening As Boolean = False
                Private mServerPort As Integer = "55545"

                'Client Config
                Private SingleClient As TcpClient

                Public Property SendDesktop As Boolean = False

                Private Sub SendImage(ByRef client As TcpClient)

                    Dim bf As New BinaryFormatter
                    Dim nstream As NetworkStream
                    nstream = client.GetStream()
                    bf.Serialize(client.GetStream(), Desktop())
                End Sub

                Public Sub GiveDesktop(ByRef client As TcpClient)
                    Do While SendDesktop = True
                        SendImage(client)
                    Loop

                End Sub

                Private Sub BroadcastMessage(ByRef ClientData As String)
                    For Each client As TcpClient In mCurrentConnectedClients

                        Dim SendStream As NetworkStream = client.GetStream
                        SendStream.Write(Encoding.ASCII.GetBytes(ClientData), 0, ClientData.Length)

                    Next
                End Sub

                Public Function Desktop() As Image
                    Dim bounds As Rectangle = Nothing
                    Dim screenshot As System.Drawing.Bitmap = Nothing
                    Dim graph As Graphics = Nothing
                    bounds = Screen.PrimaryScreen.Bounds
                    screenshot = New Bitmap(bounds.Width, bounds.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
                    graph = Graphics.FromImage(screenshot)
                    graph.CopyFromScreen(bounds.X, bounds.Y, 0, 0, bounds.Size, CopyPixelOperation.SourceCopy)
                    Return screenshot
                End Function

                ''' <summary>
                ''' Listens for newclients then adds them to the list Listens for new client messages and
                ''' returns them to all clients NewListener is started after each new client is added Each
                ''' client as a dedicated listener
                ''' </summary>
                ''' <remarks></remarks>
                Private Sub ClientListener()
                    'Add Incomming Connections
                    Dim NewClient As TcpClient
                    'If CurrentServer.Pending = True Then
                    'add client
                    Try

                        NewClient = CurrentServer.AcceptTcpClient
                        RaiseEvent NewClientAcepted(Me, NewClient)
                        'add to list
                        mCurrentConnectedClients.Add(NewClient)
                        Dim ep As EndPoint = NewClient.Client.RemoteEndPoint
                        Dim ip As String = ep.ToString
                        'Start A newClient Listener for new clients
                        Threading.ThreadPool.QueueUserWorkItem(AddressOf ClientListener)
                        'InformClient And Server
                        RaiseEvent MessageRecieved(Me, "Welcome " & NewClient.ToString & vbNewLine)
                        UpdateClient("Welcome " & NewClient.ToString & vbNewLine, NewClient)
                        BroadcastMessage("New User Connected IP:" & ip & vbNewLine)

                        Try
                            'Listen for messages
                            Do While NewClient.Connected = True And ServerIsListening = True

                                'Capture Data
                                Dim NewClientdata As NetworkStream = NewClient.GetStream
                                Dim RecievedBytes(10000) As Byte
                                If NewClient.GetStream IsNot Nothing Then

                                    'Decode Bytes
                                    NewClientdata.Read(RecievedBytes, 0, 10000)
                                    Dim ClientData As String = Encoding.ASCII.GetString(RecievedBytes)

                                    ' For return to the server control (on the server control a
                                    ' Invoke/Delegate will be required for cross threadding issues)
                                    RaiseEvent MessageRecieved(Me, ClientData)
                                    GiveDesktop(NewClient)
                                    'Update Clients
                                    For Each client As TcpClient In mCurrentConnectedClients
                                        'Do not send to Client sending the data
                                        If client IsNot NewClient Then
                                            UpdateClient(ClientData, client)
                                        End If
                                    Next
                                End If

                            Loop
                        Catch ex As Exception
                            'Remove client if error

                            'Client can try again
                            If mCurrentConnectedClients.Contains(NewClient) = True Then
                                mCurrentConnectedClients.Remove(NewClient)
                            End If

                            'ie:
                            RaiseEvent ClientListenerError(Me, ex, "Unable to Accept incomming Client")
                        End Try

                        ' Else

                        ' End If
                    Catch ex As Exception

                    End Try

                End Sub

                Private Function CreateServer(ByRef IP As IPAddress, ByRef Port As Integer) As Boolean
                    Dim OnOff As Boolean = False
                    Try
                        'Create
                        CurrentServer = New TcpListener(IP, Port)

                        OnOff = True
                        RaiseEvent MessageRecieved(Me, "Created")
                    Catch ex As Exception
                        RaiseEvent ConnectionError(Me, ex, "Unable to Create Server")
                        OnOff = False
                    End Try
                    Return OnOff
                End Function

                Private Sub StartClientListener()
                    Threading.ThreadPool.QueueUserWorkItem(AddressOf ClientListener)
                End Sub

                Private Function StartServer() As Boolean
                    Dim started As Boolean = False
                    Try
                        'Create the server listener

                        If CreateServer(ServerIP, ServerPort) = True Then
                            'Start as Server has been created sucessfully
                            CurrentServer.Start()
                            ' StartListener()
                            started = True
                            RaiseEvent MessageRecieved(Me, "Server Starting")
                        Else

                        End If
                    Catch ex As Exception
                        RaiseEvent ConnectionError(Me, ex, "Unable to Start Server")

                    End Try
                    Return started
                End Function

                Private Sub UpdateClient(ByRef ClientData As String, ByRef Client As TcpClient)
                    'Do not send to Client sending the data

                    Dim SendStream As NetworkStream = Client.GetStream
                    SendStream.Write(Encoding.ASCII.GetBytes(ClientData), 0, ClientData.Length)

                End Sub

                ''' <summary>
                ''' Occurs when Errors are Raised in the NewCLient Listener, Messages are Passed to Consumer.
                ''' </summary>
                ''' <remarks></remarks>
                Public Event ClientListenerError(ByRef Sender As ServerControl, ByRef ErrorMessage As Exception, ByRef Message As String)

                ''' <summary>
                ''' Occurs when the creation of the Chatserver cannot be created.
                ''' </summary>
                ''' <remarks></remarks>
                Public Event ConnectionError(ByRef Sender As ServerControl, ByRef ErrorMessage As Exception, ByRef Message As String)

                ''' <summary>
                ''' Occurs when New messages are needed to be relayed to Server Consumer.
                ''' </summary>
                ''' <remarks></remarks>
                Public Event MessageRecieved(ByRef Sender As ServerControl, ByRef Message As String)

                ''' <summary>
                ''' Occurs when A new Subscriber has connected and been added to the list.
                ''' </summary>
                ''' <remarks></remarks>
                Public Event NewClientAcepted(ByRef Sender As ServerControl, ByRef Client As TcpClient)

                Public Const ClassId As String = "2828F490-9903-499C-BAB3-38FF97BC1AE9"
                Public Const EventsId As String = "05FE0F04-69F4-4FD8-849F-80BD2AFBFCED"
                Public Const InterfaceId As String = "93766451-8D56-422C-B8F2-AD3DA21679B7"

                ''' <summary>
                ''' Clears all Connected users and they will need to subscribe
                ''' </summary>
                ''' <remarks></remarks>
                Public Sub Restart()
                    StopServer()
                    StartServer()
                End Sub

                ''' <summary>
                ''' Stops Server
                ''' </summary>
                ''' <remarks></remarks>
                Public Sub StopServer()

                    Try
                        For Each clin As TcpClient In mCurrentConnectedClients
                            clin.Close()
                        Next
                        'Stops NewClientListener
                        mServerIsListening = False
                        'Stops TCPLISTENER
                        CurrentServer.Stop()
                    Catch ex As Exception
                        Throw

                    End Try

                End Sub

                ''' <summary>
                ''' Initializes a new instance of the <see cref="ServerControl"/> class.
                ''' </summary>
                ''' <param name="IP">Address in string Format</param>
                ''' <param name="Port">Portnumber of Chat Server</param>
                ''' <remarks></remarks>
                Public Sub New(ByRef IP As String, ByRef Port As Integer)

                    Try
                        mServerIP = IPAddress.Parse(IP)
                        mServerPort = Port

                        'Attempt to start server
                        If StartServer() = True Then
                            'After server has been activated we can enable the server is listening var
                            'All Listeners will be deactivated if Server Is Not listening or no listeners will start
                            mServerIsListening = True
                            'We activate the first new client listener
                            StartClientListener()
                            RaiseEvent MessageRecieved(Me, "Server Started" & vbNewLine)
                        Else
                            mServerIsListening = False

                        End If
                    Catch ex As Exception
                        RaiseEvent ConnectionError(Me, ex, "Unable to Initialize Server Services")
                    End Try

                End Sub

                ''' <summary>
                ''' Gets Server CurrentIP Default is loopback.
                ''' </summary>
                ''' <value></value>
                ''' <remarks></remarks>
                Public ReadOnly Property ServerIP As IPAddress
                    Get
                        Return mServerIP
                    End Get
                End Property

                Public ReadOnly Property ServerIsListening As Boolean
                    Get
                        Return mServerIsListening
                    End Get
                End Property

                ''' <summary>
                ''' Gets Port in use, Default is 55545.
                ''' </summary>
                ''' <value></value>
                ''' <remarks></remarks>
                Public ReadOnly Property ServerPort As Integer
                    Get
                        Return mServerPort

                    End Get
                End Property

            End Class

        End Namespace

    End Namespace

End Namespace