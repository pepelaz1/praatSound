Public Class RealTier : Inherits Func
    Public points() As RealPoint

    Public Sub New(ByVal tmin As Double, ByVal tmax As Double)
        MyBase.New(tmin, tmax)
        points = New RealPoint(10) {}
    End Sub
    Public Sub RealTier_addPoint(ByVal t As Double, ByVal value As Double)
        Try
            Dim point As RealPoint = New RealPoint(t, value)
            points(points.Length) = point
        Catch ex As Exception
            'Melder_throw(Me, ": point not added.")
        End Try
    End Sub
    Function f_getNumberOfPoints() As Long
        Return points.Length
    End Function
End Class
