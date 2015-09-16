Public Class RealTier : Inherits Func
    Public points() As RealPoint

    Public Sub New(ByVal tmin As Double, ByVal tmax As Double)
        MyBase.New(tmin, tmax)
        points = New RealPoint() {}
    End Sub
    Public Sub RealTier_addPoint(ByVal t As Double, ByVal value As Double)
        Try
            Dim point As RealPoint = New RealPoint(t, value)
            Dim curL As Long = points.Length
            ReDim Preserve points(curL + 1)
            points(curL) = point
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            'Melder_throw(Me, ": point not added.")
        End Try
    End Sub
    Function f_getNumberOfPoints() As Long
        Return points.Length - 1
    End Function
End Class
