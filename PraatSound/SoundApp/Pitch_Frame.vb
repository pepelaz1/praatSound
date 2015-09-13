Public Class Pitch_Frame
    Public intensity As Double
    Public nCandidates As Long
    Public candidates(nCandidates) As PitchCandidate
    Public Sub New(ByVal nc As Long)
        candidates = New PitchCandidate() {}
        nCandidates = nc
    End Sub
    Public Sub reinit(ByVal nc As Long)
        nCandidates = nc
        ReDim candidates(nc)
    End Sub
End Class
