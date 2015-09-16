Public Class Pitch : Inherits Sampled
    Public ceiling As Double
    Public maxCandidates As Integer
    Public frame(nx) As Pitch_Frame

    Public Sub New(ByVal tmin As Double, ByVal tmax As Double, ByVal nt As Long, dt As Double, ByVal t1 As Double, ByVal ceiling1 As Double, ByVal maxnCandidates1 As Integer)
        MyBase.New(tmin, tmax, nt, dt, t1)

        ceiling = ceiling1
        maxCandidates = maxnCandidates1
        frame = New Pitch_Frame(nt) {}

        '/* Put one candidate in every frame (unvoiced, silent). */
        For it As Long = 0 To nt Step 1
            frame(it) = New Pitch_Frame(1)
        Next
    End Sub
End Class
