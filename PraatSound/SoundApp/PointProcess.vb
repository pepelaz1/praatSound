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
    Function PointProcess_getLowIndex(ByVal v As Double) As Long
        If (nt = 0 Or v < t(0)) Then
            Return 0
        End If
        '/* Special case that often occurs in practice. */
        If (v >= t(nt - 1)) Then
            Return (nt - 1)
        End If
        '
        '/* May fail if t or my t [1] is NaN. */
        If nt = 0 Then
            Console.WriteLine("cannot process execution:nt = 0")
            Return -1
        End If

        '/* Start binary search. */
        Dim left As Long = 1, right = nt - 1
        While (left < right - 1)
            Dim mid As Long = Math.Floor((left + right) / 2)
            If (v >= t(mid)) Then
                left = mid
            Else
                right = mid
            End If
        End While
        If Not (right = left + 1) Then
            Console.WriteLine("cannot process execution:right = left + 1")
            Return -1
        End If
        Return left
    End Function
    Public Sub addPoint(ByVal v As Double)
        Try
            If (v = NUMundefined) Then
                Console.WriteLine("Cannot add a point at an undefined time.")
                Return
            End If
            If (nt >= maxnt) Then
                '/*
                ' * Create without change.
                ' */
                ReDim Preserve t(2 * maxnt)
                maxnt = maxnt * 2
            End If
            If (nt = 0 Or v >= t(nt - 1)) Then
                '// special case that often occurs in practice
                nt = nt + 1
                t(nt) = v
            Else
                Dim left As Long = PointProcess_getLowIndex(v)
                If (left = 0 Or t(left) <> v) Then
                    For i As Long = nt - 1 To left - 1 Step -1
                        t(i + 1) = t(i)
                    Next
                    nt = nt + 1
                    t(left + 1) = v
                End If
            End If
        Catch ex As Exception
            Console.WriteLine("point not added:" + ex.Message)
            Return
        End Try
    End Sub

End Class
