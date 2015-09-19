Public Class NumFFTTable
    Public n As Long
    Public trigcache() As Double
    Public splitcache() As Long
    Sub drfti1(ByVal n As Long, ByRef wa() As Double, ByRef ifac() As Long)
        Dim ntryh() As Long = {4, 2, 3, 5}
        Dim tpi As Double = 2 * Math.PI
        Dim arg, argh, argld, fi As Double
        Dim ntry As Long = 0, i, j
        j = -1
        Dim k1, l1, l2, ib As Long
        Dim ld, ii, ip, is1, nq, nr As Long
        Dim ido, ipm, nfm1 As Long
        Dim nl As Long = n
        Dim nf As Long = 0

L101:
        j += 1
        If (j < 4) Then
            ntry = ntryh(j)
        Else
            ntry += 2
        End If

L104:
        nq = nl / ntry
        nr = nl - ntry * nq
        If (nr <> 0) Then
            GoTo L101
        End If

        nf += 1
        ifac(nf + 1) = ntry
        nl = nq
        If (ntry <> 2) Then
            GoTo L107
        End If
        If (nf = 1) Then
            GoTo L107
        End If

        For i = 1 To nf - 1 Step 1

            ib = nf - i + 1
            ifac(ib + 1) = ifac(ib)
        Next
        ifac(2) = 2

L107:
        If (nl <> 1) Then
            GoTo L104
        End If
        ifac(0) = n
        ifac(1) = nf
        argh = tpi / n
        is1 = 0
        nfm1 = nf - 1
        l1 = 1

        If (nfm1 = 0) Then
            Return
        End If

        For k1 = 0 To nfm1 - 1 Step 1
            ip = ifac(k1 + 2)
            ld = 0
            l2 = l1 * ip
            ido = n / l2
            ipm = ip - 1

            For j = 0 To ipm - 1 Step 1

                ld += l1
                i = is1
                argld = Convert.ToDouble(ld * argh)

                fi = 0.0
                For ii = 2 To ido - 1 Step 2
                    fi += 1.0
                    arg = fi * argld
                    wa(i + n) = Math.Cos(arg)
                    i += 1
                    wa(i + n) = Math.Sin(arg)
                    i += 1
                Next
                is1 += ido
            Next
            l1 = l2
        Next
    End Sub
    Sub NUMrffti(ByVal n As Long, ByRef wsave() As Double, ByRef iac() As Long)

        If (n = 1) Then
            Return
        End If
        'n, wsave + n, iflac
        drfti1(n, wsave, iac)
    End Sub
    Public Sub New(ByVal n1 As Long)
        n = n1
        If n1 = 0 Then
            Return
        End If
        trigcache = New Double(3 * n - 1) {}
        splitcache = New Long(31) {}
        'ReDim trigcache(3 * n - 1)
        'ReDim splitcache(32)
        'trigcache = New Double() {}
        'splitcache = New Long() {}
        NUMrffti(n, trigcache, splitcache)

    End Sub
End Class
