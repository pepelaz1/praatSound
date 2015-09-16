Public Class Pitch_Frame
    Public intensity As Double
    Public nCandidates As Long
    Public candidates(nCandidates) As PitchCandidate
    Public Sub New(ByVal nc As Long)
        candidates = New PitchCandidate(nc) {}
        nCandidates = nc
        For i As Long = 0 To nc Step 1
            candidates(i) = New PitchCandidate()
        Next
    End Sub
    Public Sub reinit(ByVal nc As Long)
        nCandidates = nc
        ReDim candidates(nc)
        For i As Long = 0 To nc Step 1
            candidates(i) = New PitchCandidate()
        Next
    End Sub
End Class
