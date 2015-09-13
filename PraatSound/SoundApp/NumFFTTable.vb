Public Class NumFFTTable
    Public n As Long
    Public trigcache() As Double
    Public splitcache() As Long
    Sub drfti1(ByVal n As Long, ByRef wsave() As Double, ByRef ifac() As Long)
        'TODO
    End Sub
    Sub NUMrffti(ByVal n As Long, ByRef wsave() As Double, ByRef ifac() As Long)

        If (n = 1) Then
            Return
        End If
        'n, wsave + n, iflac
        drfti1(n, wsave, ifac)
    End Sub
    Public Sub New(ByVal n1 As Long)
        n = n
        ReDim trigcache(3 * n - 1)
        ReDim splitcache(32)
        trigcache = New Double() {}
        splitcache = New Long() {}
        NUMrffti(n, trigcache, splitcache)

    End Sub
End Class
