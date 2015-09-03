Module NumRandom
    Public LONG_LAG = 100
    Public SHORT_LAG = 37
    Public LAG_DIFF = (LONG_LAG - SHORT_LAG)
    Public STREAM_SEPARATION = 70
    Public QUALITY = 1009
    ' Arrays in VB.NET are defined from 0 to N, so it has N+1 elements.
    ' In C, arrays are defined from 0 to N, so it has N elements. 
    ' So, to be the same, we should declare array with 1 element less
    Public randomArray(LONG_LAG - 1) As Double
    Public randomInited As Integer = 0
    Public randomArrayPointer1, randomArrayPointer2, iquality As Long
    Sub NUMrandomRestart(seed As Long)
        '/*
        '	Based on Knuth, p. 187,602.

        '	Knuth had:
        '		int s = seed;
        '	This is incorrect (even if an int is 32 bit), since a negative value causes a loop in the test
        '		if (s != 0)
        '			s >>= 1;
        '	because the >> operator on negative integers adds a sign bit to the left.
        ' */
        Dim t, j As Long
        Dim u(2 * LONG_LAG - 2) As Double
        Dim ul(2 * LONG_LAG - 2) As Double
        Dim ulp As Double = 1.0 / (1L << 30) / (1L << 22), ss
        '/* QUESTION: does this work if seed exceeds 2^32 - 3? See Knuth p. 187. */
        ss = 2.0 * ulp * (seed + 2)
        For j = 0 To LONG_LAG - 1 Step 1
            u(j) = ss
            ul(j) = 0.0
            ss += ss
            If (ss >= 1.0) Then
                ss -= 1.0 - 2 * ulp
            End If
        Next

        For j = j To 2 * LONG_LAG - 2 Step 1
            u(j) = ul(j) = 0.0
        Next

        u(1) += ulp
        ul(1) = ulp
        t = STREAM_SEPARATION - 1
        While (t > 0)
            For j = LONG_LAG - 1 To 1 Step -1
                ul(j + j) = ul(j)
                u(j + j) = u(j)
            Next

            For j = 2 * LONG_LAG - 2 To LAG_DIFF - 1 Step -2
                ul(2 * LONG_LAG - 1 - j) = 0.0
                u(2 * LONG_LAG - 1 - j) = u(j) - ul(j)
            Next
            For j = 2 * LONG_LAG - 2 To LONG_LAG Step -1
                If (ul(j) <> 0) Then
                    ul(j - LAG_DIFF) = ulp - ul(j - LAG_DIFF)
                    u(j - LAG_DIFF) += u(j)
                    If (u(j - LAG_DIFF) >= 1.0) Then
                        u(j - LAG_DIFF) -= 1.0
                    End If
                    ul(j - LONG_LAG) = ulp - ul(j - LONG_LAG)
                    u(j - LONG_LAG) += u(j)
                    If (u(j - LONG_LAG) >= 1.0) Then
                        u(j - LONG_LAG) -= 1.0
                    End If
                End If
            Next
            If ((seed & 1) <> 0) Then
                For j = LONG_LAG To 1 Step -1
                    ul(j) = ul(j - 1)
                    u(j) = u(j - 1)
                Next
                ul(0) = ul(LONG_LAG)
                u(0) = u(LONG_LAG)
                If (ul(LONG_LAG) <> 0) Then
                    ul(SHORT_LAG) = ulp - ul(SHORT_LAG)
                    u(SHORT_LAG) += u(LONG_LAG)
                    If (u(SHORT_LAG) >= 1.0) Then
                        u(SHORT_LAG) -= 1.0
                    End If
                End If
            End If
            If (seed <> 0) Then
                seed >>= 1
            Else
                t = t - 1
            End If
        End While
        For j = 0 To SHORT_LAG
            randomArray(j + LAG_DIFF) = u(j)
        Next
        For j = j To LONG_LAG - 1 Step 1
            randomArray(j - SHORT_LAG) = u(j)
        Next

        randomArrayPointer1 = 0
        randomArrayPointer2 = LAG_DIFF
        iquality = 0
        randomInited = 1
    End Sub

    Function NUMrandomFraction() As Double
        '/*
        '		Knuth uses a long random array of length QUALITY to copy values from randomArray.
        'We save 8 kilobytes by using randomArray as a cyclic array (10% speed loss).
        '*/
        Dim p1, p2 As Long
        Dim newValue As Double
        If (randomInited = False) Then
            NUMrandomRestart(DateTime.UtcNow.Millisecond)
        End If
        p1 = randomArrayPointer1
        p2 = randomArrayPointer2
        If (p1 >= LONG_LAG) Then
            p1 = 0
        End If
        If (p2 >= LONG_LAG) Then
            p2 = 0
        End If
        newValue = randomArray(p1) + randomArray(p2)
        If (newValue >= 1.0) Then
            newValue -= 1.0
        End If
        randomArray(p1) = newValue
        p1 += 1
        p2 += 1
        iquality = iquality + 1
        If (iquality = LONG_LAG) Then
            For Index As Long = iquality To QUALITY - 1 Step 1
                Dim newValue2 As Double
                '/*
                '	Possible future minor speed improvement:
                '		the cyclic array is walked down instead of up.
                '		The following tests will then be for 0.
                '*/
                If (p1 >= LONG_LAG) Then
                    p1 = 0
                End If
                If (p2 >= LONG_LAG) Then
                    p2 = 0
                End If
                newValue2 = randomArray(p1) + randomArray(p2)
                If (newValue2 >= 1.0) Then
                    newValue2 -= 1.0
                End If
                randomArray(p1) = newValue2
                p1 += 1
                p2 += 1
            Next


            iquality = 0
        End If

        randomArrayPointer1 = p1
        randomArrayPointer2 = p2
        Return newValue

    End Function

    Function NUMrandomUniform(ByVal lowest As Double, ByVal highest As Double) As Double
        Return lowest + (highest - lowest) * NUMrandomFraction()
    End Function

End Module
