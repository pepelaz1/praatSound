Public Class PointProcess : Inherits Func
    Public maxnt As Long
    Public nt As Long
    Public t() As Double

    Sub New(ByVal mn As Long, ByVal n As Long, ByVal arr() As Double, ByVal f As Func)
        maxnt = mn
        nt = n
        t = arr
        MyBase.xmax = f.xmax
        MyBase.xmin = f.xmin
    End Sub

End Class
