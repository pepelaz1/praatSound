Public Class Pitch : Inherits Sampled
    Public ceiling As Double
    Public maxCandidates As Integer
    Public frame(nx) As Pitch_Frame

    Public Sub New(ByVal tmin As Double, ByVal tmax As Double, ByVal nt As Long, dt As Double, ByVal t1 As Double, ByVal ceiling1 As Double, ByVal maxnCandidates1 As Integer)
        MyBase.New(tmin, tmax, nt, dt, t1)

        ceiling = ceiling1
        maxCandidates = maxnCandidates1
        frame = New Pitch_Frame(nt - 1) {}

        '/* Put one candidate in every frame (unvoiced, silent). */
        For it As Long = 0 To nt - 1 Step 1
            frame(it) = New Pitch_Frame(0)
        Next
    End Sub
    Function v_convertStandardToSpecialUnit(ByVal value As Double, ByVal ilevel As Long, ByVal unit As Integer) As Double
        If (ilevel = Pitch_LEVEL_FREQUENCY) Then
            Return value
        Else
            Return NUMundefined
        End If
    End Function
    Overrides Function v_getValueAtSample(ByVal iframe As Long, ByVal ilevel As Long, ByVal unit As Long) As Double
        Dim f As Double = frame(iframe).candidates(0).frequency
        If (f <= 0.0 Or f >= ceiling) Then
            Return NUMundefined
        End If
        '// frequency out of range (or NUMundefined)? Voiceless
        Dim d As Double
        If ilevel = Pitch_LEVEL_FREQUENCY Then
            d = f
        Else
            d = frame(iframe).candidates(0).strength
        End If
        Return v_convertStandardToSpecialUnit(d, ilevel, unit)

    End Function
End Class
