Public Class Vector : Inherits Matrix
    Public Sub New(ByVal _xmin As Double, ByVal _xmax As Double, ByVal _nx As Long, ByVal _dx As Double, ByVal _x1 As Double,
                                              ByVal _ymin As Double, ByVal _ymax As Double, ByVal _ny As Long, ByVal _dy As Double, ByVal _y1 As Double)
        MyBase.New(_xmin, _xmax, _nx, _dx, _x1, _ymin, _ymax, _ny, _dy, _y1)
    End Sub
End Class
