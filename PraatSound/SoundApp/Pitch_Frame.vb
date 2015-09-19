Public Class Pitch_Frame
    Public intensity As Double
    Public nCandidates As Long
    Public candidates(nCandidates) As PitchCandidate
    Public Sub New(ByVal nc As Long)
        nCandidates = nc
        If nc = 0 Then
            Return
        End If
        candidates = New PitchCandidate(nc - 1) {}
        For i As Long = 0 To nc - 1 Step 1
            candidates(i) = New PitchCandidate()
        Next
    End Sub
    Public Sub reinit(ByVal nc As Long)
        nCandidates = nc
        ReDim candidates(nc - 1)
        For i As Long = 0 To nc - 1 Step 1
            candidates(i) = New PitchCandidate()
        Next
    End Sub
End Class
