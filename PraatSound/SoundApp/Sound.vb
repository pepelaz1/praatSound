Public Class Sound : Inherits Vector
    Sub New(ByVal numberOfChannels As Long, ByVal xmin As Double, ByVal xmax As Double, ByVal nx As Long, ByVal dx As Double, ByVal x1 As Double)
        Try
            init(Me, xmin, xmax, nx, dx, x1, 1, numberOfChannels, numberOfChannels, 1, 1)
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub

End Class
