Public Class Sampled : Inherits Func
    Public nx As Long
    Public dx As Double
    Public x1 As Double
    Public Sub New(ByVal _xmin As Double, ByVal _xmax As Double, ByVal _nx As Long, ByVal _dx As Double, ByVal _x1 As Double)
        MyBase.New(_xmin, _xmax)
        nx = _nx
        dx = _dx
        x1 = _x1
    End Sub
    Sub Sampled_Init(ByRef mme As Sampled, ByVal xmin As Double, ByVal xmax As Double, ByVal nx As Long, ByVal dx As Double, ByVal x1 As Double)
        mme.xmin = xmin
        mme.xmax = xmax
        mme.nx = nx
        mme.dx = dx
        mme.x1 = x1
    End Sub
    ' DEFINE IN REST CLASSES,  WHICH INHERITS SAMPLED!!!
    Overridable Function v_getValueAtSample(ByVal iframe As Long, ByVal ilevel As Long, ByVal unit As Long) As Double
        Return System.Double.PositiveInfinity
    End Function
End Class
