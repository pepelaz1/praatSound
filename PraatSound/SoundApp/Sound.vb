Public Class Sound : Inherits Vector
    Public Const Melder_MULAW = 9
    Public Const Melder_ALAW = 10
    Public Const Melder_SHORTEN As Long = 11
    Public Const Melder_POLYPHONE As Long = 12
    Public Const Melder_IEEE_FLOAT_32_LITTLE_ENDIAN = 14
    Public Const Melder_AIFF As Long = 1
    Public Const Melder_AIFC As Long = 2
    Public Const Melder_WAV As Long = 3
    Public Const Melder_NEXT_SUN As Long = 4
    Public Const Melder_NIST As Long = 5
    Public Const Melder_SOUND_DESIGNER_TWO As Long = 6
    Public Const Melder_FLAC As Long = 7
    Public Const Melder_MP3 As Long = 8

    Public Const Melder_LINEAR_8_SIGNED = 1
    Public Const Melder_LINEAR_8_UNSIGNED = 2
    Public Const Melder_LINEAR_16_BIG_ENDIAN = 3
    Public Const Melder_LINEAR_16_LITTLE_ENDIAN = 4
    Public Const Melder_LINEAR_24_BIG_ENDIAN = 5
    Public Const Melder_LINEAR_24_LITTLE_ENDIAN = 6
    Public Const Melder_LINEAR_32_BIG_ENDIAN = 7
    Public Const Melder_LINEAR_32_LITTLE_ENDIAN = 8

    Public Const WAVE_FORMAT_PCM = &H1
    Public Const WAVE_FORMAT_IEEE_FLOAT = &H3
    Public Const WAVE_FORMAT_ALAW = &H6
    Public Const WAVE_FORMAT_MULAW = &H7
    Public Const WAVE_FORMAT_DVI_ADPCM = &H11

    Public Melder_debug As Long = 0

    Private Declare Function GetAddrOf Lib "KERNEL32" Alias "MulDiv" (nNumber As Long, Optional ByVal nNumerator As Long = 1, Optional ByVal nDenominator As Long = 1) As Long
    Private Declare Sub MoveMemory Lib "KERNEL32" Alias "RtlMoveMemory" (hpvDest As Long, hpvSource As Long, ByVal cbCopy As Long)

    Sub Melder_checkWavFile(ByRef reader As System.IO.BinaryReader, ByRef numberOfChannels As Integer, ByRef encoding As Integer, ByRef sampleRate As Double, ByRef startOfData As Long, ByRef numberOfSamples As Long)
        Dim data(7), chunkID(3) As Byte
        Dim formatChunkPresent As Boolean = False, dataChunkPresent = False
        Dim numberOfBitsPerSamplePoint As Integer = -1
        Dim dataChunkSize As Long = -1
        data = reader.ReadBytes(4)
        If data.GetLength(0) < 4 Then
            Console.WriteLine("File too small: no RIFF statement.")
            Return
        End If
        If (Not data.ToString.Substring(0, 4) = "RIFF") Then
            Console.WriteLine("Not a WAV file (RIFF statement expected).")
            Return
        End If
        data = reader.ReadBytes(4)
        If data.GetLength(0) < 4 Then
            Console.WriteLine("File too small: no size of RIFF chunk.")
            Return
        End If
        data = reader.ReadBytes(4)
        If data.GetLength(0) < 4 Then
            Console.WriteLine("File too small: no file type info (expected WAVE statement).")
            Return
        End If
        If (data.ToString.Substring(0, 4) <> "WAVE" And data.ToString.Substring(0, 4) <> "CDDA") Then
            Console.WriteLine("Not a WAVE or CD audio file (wrong file type info).")
            Return
        End If
        chunkID = reader.ReadBytes(4)
        ' /* Search for Format Chunk and Data Chunk. */

        While (chunkID.ToString.Length = 4)
            Dim chunkSize As Long = bingeti4LE(reader)
            ' If (Melder_debug = 23) Then
            '        Melder_warning (Melder_integer (chunkID [0]), L" ", Melder_integer (chunkID [1]), L" ", Melder_integer (chunkID [2]), L" ", Melder_integer (chunkID [3]), L" ", Melder_integer (chunkSize))
            'End If
            If (chunkID.ToString.StartsWith("fmt ")) Then
                '/*
                ' * Found a Format Chunk.
                ' */
                Dim winEncoding As Integer = bingeti2LE(reader)
                formatChunkPresent = True
                numberOfChannels = bingeti2LE(reader)
                If (numberOfChannels < 1) Then
                    Console.WriteLine("Too few sound channels (" + numberOfChannels + ").")
                    Return
                End If
                sampleRate = Convert.ToDouble(bingeti4LE(reader))
                If (sampleRate <= 0.0) Then
                    Console.WriteLine("Wrong sampling frequency (" + sampleRate + " Hz).")
                    Return
                End If
                '// avgBytesPerSec
                bingeti4LE(reader)
                '// blockAlign
                bingeti2LE(reader)
                numberOfBitsPerSamplePoint = bingeti2LE(reader)
                If (numberOfBitsPerSamplePoint = 0) Then
                    numberOfBitsPerSamplePoint = 16
                ElseIf (numberOfBitsPerSamplePoint < 4) Then
                    Console.WriteLine("Too few bits per sample (" + numberOfBitsPerSamplePoint + "; the minimum is 4).")
                    Return
                ElseIf (numberOfBitsPerSamplePoint > 32) Then
                    Console.WriteLine("Too few bits per sample (" + numberOfBitsPerSamplePoint + "; the maximum is 32).")
                    Return
                End If
                    Select winEncoding
                    Case WAVE_FORMAT_PCM
                        If numberOfBitsPerSamplePoint > 24 Then
                            encoding = Melder_LINEAR_32_LITTLE_ENDIAN
                        ElseIf numberOfBitsPerSamplePoint > 16 Then
                            encoding = Melder_LINEAR_24_LITTLE_ENDIAN
                        ElseIf numberOfBitsPerSamplePoint > 8 Then
                            encoding = Melder_LINEAR_16_LITTLE_ENDIAN
                        Else
                            encoding = Melder_LINEAR_8_UNSIGNED
                        End If
                        GoTo endoffor
                    Case WAVE_FORMAT_IEEE_FLOAT
                        encoding = Melder_IEEE_FLOAT_32_LITTLE_ENDIAN
                        GoTo endoffor
                    Case WAVE_FORMAT_ALAW
                        encoding = Melder_ALAW
                        GoTo endoffor
                    Case WAVE_FORMAT_MULAW
                        encoding = Melder_MULAW
                        GoTo endoffor
                    Case WAVE_FORMAT_DVI_ADPCM
                        Console.WriteLine("Cannot read lossy compressed audio files (this one is DVI ADPCM).\n" +
                            "Please use uncompressed audio files. If you must open this file,\n" +
                            "please use an audio converter program to convert it first to normal (PCM) WAV format\n" +
                            "(Praat may have difficulty analysing the poor recording, though).")
                        Return
                    Case Else
                        Console.WriteLine("Unsupported Windows audio encoding %d." + winEncoding)
                        Return
                End Select
                If (chunkSize & 1) Then
                    chunkSize = chunkSize + 1
                End If
                For i As Long = 17 To chunkSize Step 1
                    data = reader.ReadBytes(1)
                    If (data.GetLength(0) < 1) Then
                        Console.WriteLine("File too small: expected " + chunkSize + " bytes in fmt chunk, but found " + i + ".")
                        Return
                    End If
                Next
            Else
                If (chunkID.ToString.StartsWith("data")) Then
                    '/*
                    ' * Found a Data Chunk.
                    ' */
                    dataChunkPresent = True
                    dataChunkSize = chunkSize
                    startOfData = reader.BaseStream.Position
                    If (chunkSize & 1) Then
                        chunkSize = chunkSize + 1
                    End If
                    If (chunkSize < 0) Then
                        '// incorrect data chunk (sometimes -44); assume that the data run till the end of the file
                        reader.BaseStream.Seek(0, System.IO.SeekOrigin.End)
                        Dim endOfData As Long = reader.BaseStream.Position
                        dataChunkSize = chunkSize = endOfData - startOfData
                        reader.BaseStream.Seek(startOfData, IO.SeekOrigin.Begin)
                    End If
                    If (Melder_debug = 23) Then
                        For i As Long = 1 To chunkSize Step 1
                            data = reader.ReadBytes(1)
                            If (data.GetLength(0) < 1) Then
                                Console.WriteLine("File too small: expected " + chunkSize + " bytes of data, but found " + i + ".")
                                Return
                            End If
                        Next
                    Else
                        If (formatChunkPresent) Then
                            GoTo endoffor
                        End If
                    End If
                    '// OPTIMIZATION: do not read the whole data chunk if we have already read the format chunk
                Else
                    '// ignore other chunks
                    If (chunkSize & 1) Then
                        chunkSize = chunkSize + 1
                    End If
                    For i As Long = 1 To chunkSize Step 1
                        data = reader.ReadBytes(1)
                        If (data.GetLength(0) < 1) Then
                            Console.WriteLine("File too small: expected " + chunkSize, " bytes, but found " + i + ".")
                            Return
                        End If
                    Next

                End If
            End If
        End While

endoffor: If (Not formatChunkPresent) Then
            Console.WriteLine("Found no Format Chunk.")
            Return
        End If
        If (Not dataChunkPresent) Then
            Console.WriteLine("Found no Data Chunk.")
            Return
        End If
        If (Not (numberOfBitsPerSamplePoint <> -1 And dataChunkSize <> -1)) Then
            Console.WriteLine("Cannot continue execution: numberOfBitsPerSamplePoint <> -1 And dataChunkSize <> -1")
            Return
        End If
        numberOfSamples = dataChunkSize / numberOfChannels / ((numberOfBitsPerSamplePoint + 7) / 8)
    End Sub
    Function bingeti2LE(ByRef reader As System.IO.BinaryReader) As Integer
        If (Melder_debug <> 18) Then
            Dim s As Short
            s = reader.ReadInt16()
            'if (fread (& s, sizeof (signed short), 1, f) != 1) Then 
            ' readError(f, "a signed short integer.")
            ' End If
            Return Convert.ToInt32(s)
            '// with sign extension if an int is 4 bytes
        Else
            Dim bytes(1) As Byte
            bytes = reader.ReadBytes(2)
            If (bytes.GetLength(0) <> 2) Then
                Console.WriteLine("Cannot read two bytes.")
                Return -1
            End If
            Dim externalValue As UShort = (Convert.ToUInt16(bytes(1)) << 8) Or Convert.ToUInt16(bytes(0))
            Return Convert.ToInt32(Convert.ToInt16(externalValue))
            '// with sign extension if an int is 4 bytes
        End If
    End Function
    Function bingeti4LE(ByRef reader As System.IO.BinaryReader) As Long
        If (Melder_debug <> 18) Then
            Dim l As Long
            l = reader.ReadUInt32()
            'if (fread (& l, sizeof (long), 1, f) != 1) Then
            ' Console.WriteLine("Error reading a signed long integer.")
            'End If
            Return l
        Else
            Dim bytes(3) As Byte
            bytes = reader.ReadBytes(4)
            If (bytes.GetLength(0) <> 4) Then
                Console.WriteLine("Error reading four bytes.")
                Return -1
            End If
            Return ((Convert.ToUInt32(bytes(3)) << 24) Or (Convert.ToUInt32(bytes(2)) << 16) Or (Convert.ToUInt32(bytes(1)) << 8) Or Convert.ToUInt32(bytes(0)))
            ' // 32 or 64 bits
        End If
    End Function

    Function MelderFile_checkSoundFile(ByRef reader As System.IO.BinaryReader, ByRef numberOfChannels As Integer, ByRef encoding As Integer, ByRef sampleRate As Double, ByRef startOfData As Long, ByRef numberOfSamples As Long) As Integer
        Dim data As Byte()
        data = reader.ReadBytes(16)
        If (reader Is Nothing Or data.GetLength(0) < 16) Then
            Return 0
        End If
        reader.BaseStream.Seek(0, System.IO.SeekOrigin.Begin)
        Dim DataStr As String = data.ToString
        'If (strnequ(data, "FORM", 4) And strnequ(data + 8, "AIFF", 4)) Then
        'Melder_checkAiffFile(reader, numberOfChannels, encoding, sampleRate, startOfData, numberOfSamples)
        'Return Melder_AIFF
        'End If
        'If (strnequ(data, "FORM", 4) And strnequ(data + 8, "AIFC", 4)) Then
        '    Melder_checkAiffFile(reader, numberOfChannels, encoding, sampleRate, startOfData, numberOfSamples)
        '    Return Melder_AIFC
        'End If
        If (DataStr.Substring(0, 4) = "RIFF" And (DataStr.Substring(8, 4) = "WAVE" Or DataStr.Substring(8, 4) = "CDDA")) Then
            Melder_checkWavFile(reader, numberOfChannels, encoding, sampleRate, startOfData, numberOfSamples)
            Return Melder_WAV
        End If
        'If (strnequ(data, ".snd", 4)) Then
        '    Melder_checkNextSunFile(reader, numberOfChannels, encoding, sampleRate, startOfData, numberOfSamples)
        '    Return Melder_NEXT_SUN
        'End If
        'If (strnequ(data, "NIST_1A", 7)) Then
        '    Melder_checkNistFile(reader, numberOfChannels, encoding, sampleRate, startOfData, numberOfSamples)
        '    Return Melder_NIST
        'End If
        Return 0
    End Function
    Sub Melder_readAudioToFloat(ByRef reader As System.IO.BinaryReader, ByVal numberOfChannels As Integer, ByVal encoding As Integer, ByRef buffer(,) As Double, ByVal numberOfSamples As Long)
        Select Case encoding
            Case Melder_LINEAR_16_LITTLE_ENDIAN
                Dim numberOfBytesPerSamplePerChannel As Integer = 2
                '(int) sizeof (double) = 8
                If (numberOfChannels > 8 / numberOfBytesPerSamplePerChannel) Then
                    For isamp As Long = 1 To numberOfSamples Step 1
                        For ichan As Long = 1 To numberOfChannels Step 1
                            buffer(ichan, isamp) = bingeti2LE(reader) * (1.0 / 32768)
                        Next
                    Next
                Else
                    Dim numberOfBytes As Long = numberOfChannels * numberOfSamples * numberOfBytesPerSamplePerChannel
                    Dim bytes() As Byte = reader.ReadBytes(numberOfBytes)
                    ' Write bytes to buffer at specific address
                    Dim dstAddr As Long = GetAddrOf(buffer(numberOfChannels, numberOfSamples)) + 8 - numberOfBytes
                    Dim srcAddr As Long = GetAddrOf(bytes(0))
                    MoveMemory(dstAddr, srcAddr, numberOfBytes)
                    'If (fread(bytes, 1, numberOfBytes, f) < numberOfBytes) Then
                    'Console.WriteLine("Error in ReadAudioToFloat")
                    'Return
                    'End If
                    '// read 16-bit data into last quarter of buffer
                    Dim i As Long = 0
                    If (numberOfChannels = 1) Then
                        For isamp As Long = 1 To numberOfSamples Step 1
                            Dim byte1 As Byte = bytes(i)
                            i = i + 1
                            Dim byte2 = bytes(i)
                            i = i + 1
                            Dim value As Long = Convert.ToInt16(((Convert.ToUInt16(byte2) << 8) Or Convert.ToUInt16(byte1)))
                            buffer(1, isamp) = value * (1.0 / 32768)
                        Next
                    Else
                        For isamp As Long = 1 To numberOfSamples Step 1
                            For ichan As Long = 1 To numberOfChannels Step 1
                                Dim byte1 As Byte = bytes(i)
                                i = i + 1
                                Dim byte2 = bytes(i)
                                i = i + 1
                                Dim value As Long = Convert.ToInt16((Convert.ToUInt16(byte2) << 8) Or Convert.ToUInt16(byte1))
                                buffer(ichan, isamp) = value * (1.0 / 32768)
                            Next
                        Next
                    End If
                End If
        End Select

    End Sub
    Sub New(ByVal numberOfChannels As Long, ByVal xmin As Double, ByVal xmax As Double, ByVal nx As Long, ByVal dx As Double, ByVal x1 As Double)
        Try
            init(Me, xmin, xmax, nx, dx, x1, 1, numberOfChannels, numberOfChannels, 1, 1)
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub
    Public Function Sound_createSimple(ByVal numberOfChannels As Long, ByVal duration As Double, ByVal samplingFrequency As Double) As Sound
        Return New Sound(numberOfChannels, 0.0, duration, Math.Floor(duration * samplingFrequency + 0.5), 1 / samplingFrequency, 0.5 / samplingFrequency)
    End Function
    Function Sound_readFromSoundFile(ByVal filePath As String) As Sound
        Try
            ' autoMelderFile mfile = MelderFile_open (file);
            Dim reader = New System.IO.BinaryReader(System.IO.File.Open(filePath, System.IO.FileMode.Open))
            Dim numberOfChannels, encoding As Integer
            Dim sampleRate As Double
            Dim startOfData, numberOfSamples As Long
            Dim fileType As Integer = MelderFile_checkSoundFile(reader, numberOfChannels, encoding, sampleRate, startOfData, numberOfSamples)
            If (fileType = 0) Then
                Console.WriteLine("Not an audio file.")
                Return Nothing
            End If
            '// start from beginning of Data Chunk
            If (reader.BaseStream.Seek(0, IO.SeekOrigin.Begin) = IO.SeekOrigin.End) Then
                Console.WriteLine("No data in audio file.")
                Return Nothing
            End If
            Dim mme As Sound = Sound_createSimple(numberOfChannels, numberOfSamples / sampleRate, sampleRate)
            If (encoding = Melder_SHORTEN Or encoding = Melder_POLYPHONE) Then
                Console.WriteLine("Cannot unshorten. Write to paul.boersma@uva.nl for more information.")
                Return Nothing
            End If
            Melder_readAudioToFloat(reader, numberOfChannels, encoding, mme.z, numberOfSamples)
            reader.Close()
            Return mme
        Catch ex As Exception
            Console.WriteLine("Sound not read from sound file " + filePath)
            Return Nothing
        End Try
    End Function
End Class
