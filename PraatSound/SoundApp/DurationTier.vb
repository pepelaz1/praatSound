Public Class DurationTier : Inherits RealTier
    Public Sub New(ByVal _min As Double, _max As Double)
        MyBase.New(_min, _max)
    End Sub
    Public Sub addPoint(ByVal t As Double, ByVal value As Double)
        MyBase.RealTier_addPoint(t, value)
    End Sub
End Class
