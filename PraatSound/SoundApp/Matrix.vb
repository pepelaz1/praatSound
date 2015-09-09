Public Class Matrix : Inherits Sampled
    Public ny As Long
    Public ymin As Double
    Public ymax As Double
    Public dy As Double
    Public y1 As Double
    Public z(,) As Double

    Public Sub New(ByVal _xmin As Double, ByVal _xmax As Double, ByVal _nx As Long, ByVal _dx As Double, ByVal _x1 As Double,
                                              ByVal _ymin As Double, ByVal _ymax As Double, ByVal _ny As Long, ByVal _dy As Double, ByVal _y1 As Double
                         )
        MyBase.New(_xmin, _xmax, _nx, _dx, _x1)
        ymin = _ymin
        ymax = _ymax
        ny = _ny
        dy = _dy
        y1 = _y1
        z = New Double(1, _nx - _ny + 1) {}
    End Sub
    Public Sub init(ByRef mme As Matrix, ByVal xmin As Double, ByVal xmax As Double, ByVal nx As Long, ByVal dx As Double, ByVal x1 As Double,
                                              ByVal ymin As Double, ByVal ymax As Double, ByVal ny As Long, ByVal dy As Double, ByVal y1 As Double
                         )
        Sampled_Init(mme, xmin, xmax, nx, dx, x1)
        mme.ymin = ymin
        mme.ymax = ymax
        mme.ny = ny
        mme.dy = dy
        mme.y1 = y1
        mme.z = New Double(1, mme.nx - mme.ny + 1) {}
    End Sub
End Class
