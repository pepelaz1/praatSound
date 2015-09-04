Public Class Sampled : Inherits Func
    Public nx As Long
    Public dx As Double
    Public x1 As Double
    Sub Sampled_Init(ByRef mme As Sampled, ByVal xmin As Double, ByVal xmax As Double, ByVal nx As Long, ByVal dx As Double, ByVal x1 As Double)
        mme.xmin = xmin
        mme.xmax = xmax
        mme.nx = nx
        mme.dx = dx
        mme.x1 = x1
    End Sub
End Class
