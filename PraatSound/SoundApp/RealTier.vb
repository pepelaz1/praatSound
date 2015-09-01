Public Class RealTier : Inherits Func
    Public points() As RealPoint
    Function f_getNumberOfPoints() As Long
        Return points.Length
    End Function
End Class
