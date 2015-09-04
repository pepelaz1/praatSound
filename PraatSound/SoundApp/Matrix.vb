Public Class Matrix : Inherits Sampled
    Public ny As Long
    Public ymin As Double
    Public ymax As Double
    Public dy As Double
    Public y1 As Double
    Public z(,) As Double

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
