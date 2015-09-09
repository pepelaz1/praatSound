Public Class PointProcess : Inherits Func
    Public maxnt As Long
    Public nt As Long
    Public t() As Double

    'Sub New(ByVal mn As Long, ByVal n As Long, ByVal arr() As Double, ByVal f As Func)
    '    maxnt = mn
    '    nt = n
    '    t = arr
    '    MyBase.xmax = f.xmax
    '    MyBase.xmin = f.xmin
    'End Sub

    Sub New(ByVal tmin As Double, ByVal tmax As Double, ByVal initialMaxnt As Long)
        MyBase.New(tmin, tmax)
        If (initialMaxnt < 1) Then
            initialMaxnt = 1
        End If
        maxnt = initialMaxnt
        nt = 0
        t = New Double(maxnt - 1) {}
    End Sub

    Function PointProcess_create(ByVal tmin As Double, ByVal tmax As Double, ByVal initialMaxnt As Long) As PointProcess
        Try
            Dim pr As PointProcess = New PointProcess(tmin, tmax, initialMaxnt)
            Return pr
        Catch ex As Exception
            Return Nothing
            'Melder_throw("PointProcess not created.")
        End Try
    End Function

End Class
