Public Class Pitch : Inherits Sampled
    Public ceiling As Double
    Public maxCandidates As Integer
    Public frame(nx) As Pitch_Frame

    Public Sub New(ByVal tmin As Double, ByVal tmax As Double, ByVal nt As Long, dt As Double, ByVal t1 As Double, ByVal ceiling As Double, ByVal maxnCandidates As Integer)
        MyBase.New(tmin, tmax, nt, dt, t1)
        'CONTINUE!!!
    End Sub
End Class
