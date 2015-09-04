Public Class Sound : Inherits Vector
    Public Const Melder_SHORTEN As Long = 11
    Public Const Melder_POLYPHONE As Long = 12
    Function MelderFile_checkSoundFile(ByRef reader As System.IO.BinaryReader, ByRef numberOfChannels As Integer, ByRef encoding As Integer, ByRef sampleRate As Double, ByRef startOfData As Long, ByRef numberOfSamples As Long) As Integer
        'TODO
        Return 0
    End Function
    Sub Melder_readAudioToFloat(ByRef reader As System.IO.BinaryReader, ByVal numberOfChannels As Integer, ByVal encoding As Integer, ByRef buffer(,) As Double, ByVal numberOfSamples As Long)
        'TODO
    End Sub
    Sub New(ByVal numberOfChannels As Long, ByVal xmin As Double, ByVal xmax As Double, ByVal nx As Long, ByVal dx As Double, ByVal x1 As Double)
        Try
            init(Me, xmin, xmax, nx, dx, x1, 1, numberOfChannels, numberOfChannels, 1, 1)
        Catch ex As Exception
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
            If (reader.BaseStream.Seek(0, IO.SeekOrigin.Current) = IO.SeekOrigin.End) Then
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
