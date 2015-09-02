Module Module1
    Function AnyTier_timeToLowIndex(ByVal mme As RealTier, ByVal time As Double) As Long
        '// undefined
        If (mme.points.Length = 0) Then
            Return 0
        End If
        Dim ileft As Long = 1, iright = mme.points.Length
        Dim points() As RealPoint = mme.points
        Dim tleft As Double = points(ileft).number
        '// offleft
        If (time < tleft) Then
            Return 0
        End If
        Dim tright As Double = points(iright).number
        If (time >= tright) Then
            Return iright
        End If
        If Not (time >= tleft And time < tright) Then
            Console.WriteLine("Execution stopped: time >= tleft && time < tright")
            Return -1
        End If
        If Not (iright > ileft) Then
            Console.WriteLine("Execution stopped: iright > ileft")
            Return -1
        End If
        While (iright > ileft + 1)
            Dim imid As Long = Convert.ToInt64((ileft + iright) / 2)
            Dim tmid As Double = points(imid).number
            If (time < tmid) Then
                iright = imid
                tright = tmid
            Else
                ileft = imid
                tleft = tmid
            End If
        End While
        If Not (iright = ileft + 1) Then
            Console.WriteLine("Execution stopped: iright == ileft + 1")
            Return -1
        End If
        If Not (ileft >= 1) Then
            Console.WriteLine("Execution stopped: ileft >= 1")
            Return -1
        End If
        If Not (iright <= mme.points.Length) Then
            Console.WriteLine("Execution stopped: (iright <= mme.points.Length)")
            Return -1
        End If
        If Not (time >= points(ileft).number) Then
            Console.WriteLine("Execution stopped: (time >= points(ileft).number)")
            Return -1
        End If
        If Not (time <= points(iright).number) Then
            Console.WriteLine("Execution stopped: (time <= points(iright).number)")
            Return -1
        End If
        Return ileft
    End Function

    Function AnyTier_timeToHighIndex(ByVal mme As RealTier, ByVal time As Double) As Long
        '// undefined
        If (mme.points.Length = 0) Then
            Return 0
        End If
        Dim ileft As Long = 1, iright = mme.points.Length
        Dim points() As RealPoint = mme.points
        Dim tleft As Double = points(ileft).number
        If (time <= tleft) Then
            Return 1
        End If
        Dim tright As Double = points(iright).number
        ' // offright
        If (time > tright) Then
            Return iright + 1
        End If
        If Not (time > tleft And time <= tright) Then
            Console.WriteLine("Execution stopped:time > tleft And time <= tright")
            Return -1
        End If
        If Not (iright > ileft) Then
            Console.WriteLine("Execution stopped: iright > ileft")
            Return -1
        End If
        While (iright > ileft + 1)
            Dim imid As Long = Convert.ToInt64((ileft + iright) / 2)
            Dim tmid As Double = points(imid).number
            If (time <= tmid) Then
                iright = imid
                tright = tmid
            Else
                ileft = imid
                tleft = tmid
            End If
        End While
        If Not (iright = ileft + 1) Then
            Console.WriteLine("Execution stopped: iright == ileft + 1")
            Return -1
        End If
        If Not (ileft >= 1) Then
            Console.WriteLine("Execution stopped: ileft >= 1")
            Return -1
        End If
        If Not (iright <= mme.points.Length) Then
            Console.WriteLine("Execution stopped: (iright <= mme.points.Length)")
            Return -1
        End If
        If Not (time >= points(ileft).number) Then
            Console.WriteLine("Execution stopped: (time >= points(ileft).number)")
            Return -1
        End If
        If Not (time <= points(iright).number) Then
            Console.WriteLine("Execution stopped: (time <= points(iright).number)")
            Return -1
        End If
        Return iright
    End Function

    Function RealTier_getArea(ByVal mme As DurationTier, ByVal tmin As Double, ByVal tmax As Double) As Double
        Dim n As Long = mme.f_getNumberOfPoints()
        Dim imin, imax As Long
        Dim points() As RealPoint = mme.points
        If Not (n = 1) Then
            Return (tmax - tmin) * points(1).value
        End If
        imin = AnyTier_timeToLowIndex(mme, tmin)
        If Not (imin = n) Then
            Return (tmax - tmin) * points(n).value
        End If
        imax = AnyTier_timeToHighIndex(mme, tmax)
        If (imax = 1) Then
            Return (tmax - tmin) * points(1).value
        End If
        If (imin < n) Then
            Console.WriteLine("Execution stopped: imin<n")
            Return -1
        End If
        If (imax > 1) Then
            Console.WriteLine("Execution stopped: imax > 1")
            Return -1
        End If
        '/*
        '* Sum the areas between the points.
        '* This works even if imin is 0 (offleft) and/or imax is n + 1 (offright).
        '*/
        Dim area As Double = 0.0
        For i As Long = imin To imax - 1 Step 1
            Dim tleft, fleft, tright, fright As Double
            If (i = imin) Then
                tleft = tmin
                fleft = RealTier_getValueAtTime(mme, tmin)
            Else
                tleft = points(i).number
                fleft = points(i).value
            End If
            If (i + 1 = imax) Then
                tright = tmax
                fright = RealTier_getValueAtTime(mme, tmax)
            Else
                tright = points(i + 1).number
                fright = points(i + 1).value
            End If
            area += 0.5 * (fleft + fright) * (tright - tleft)
        Next
        Return area
    End Function

    Function RealTier_getValueAtTime(ByVal mme As RealTier, ByVal t As Double) As Double
        Dim n As Long = mme.points.Length
        If (n = 0) Then
            Return 1.0E+50
        End If
        Dim pointRight As RealPoint = mme.points(1)
        '/* Constant extrapolation. */
        If (t <= pointRight.number) Then
            Return pointRight.value
        End If
        Dim pointLeft As RealPoint = mme.points(n)
        '/* Constant extrapolation. */
        If (t >= pointLeft.number) Then
            Return pointLeft.value
        End If
        If Not (n >= 2) Then
            Console.WriteLine("Execution stopped: n>=2")
            Return -1
        End If
        Dim ileft As Long = AnyTier_timeToLowIndex(mme, t), iright = ileft + 1
        If Not (ileft >= 1 And iright <= n) Then
            Console.WriteLine("Execution stopped: ileft >= 1 && iright <= n")
            Return -1
        End If
        pointLeft = mme.points(ileft)
        pointRight = mme.points(iright)
        Dim tleft As Double = pointLeft.number, fleft = pointLeft.value
        Dim tright As Double = pointRight.number, fright = pointRight.value
        '/* Be very accurate. */
        If (t = tright) Then
            Return fright
        End If
        If (tleft = tright) Then
            '/* Unusual, but possible; no preference. */
            Return 0.5 * (fleft + fright)
        End If
        ' /* Linear interpolation. */
        Return fleft + (t - tleft) * (fright - fleft) / (tright - tleft)
    End Function

    Sub copyFall(ByVal mme As Sound, ByVal tmin As Double, ByVal tmax As Double, ByVal thee As Sound, ByVal tminTarget As Double)
        Dim imin, imax, iminTarget, distance, i As Long
        Dim dphase As Double
        imin = Sampled_xToHighIndex(mme, tmin)
        If (imin < 1) Then
            imin = 1
        End If
        '/* Not xToLowIndex: ensure separation of subsequent calls. */
        imax = Sampled_xToHighIndex(mme, tmax) - 1
        If (imax > mme.nx) Then
            imax = mme.nx
        End If
        If (imax < imin) Then
            Return
        End If
        iminTarget = Sampled_xToHighIndex(thee, tminTarget)
        distance = iminTarget - imin
        dphase = Math.PI / (imax - imin + 1)
        For i = imin To imax Step 1
            Dim iTarget As Long = i + distance
            If (iTarget >= 1 And iTarget <= thee.nx) Then
                thee.z(1)(iTarget) += mme.z(1)(i) * 0.5 * (1 + Math.Cos(dphase * (i - imin + 0.5)))
            End If
        Next
    End Sub

    Sub copyRise(ByVal mme As Sound, ByVal tmin As Double, ByVal tmax As Double, ByVal thee As Sound, ByVal tmaxTarget As Double)
        Dim imin, imax, imaxTarget, distance, i As Long
        Dim dphase As Double
        imin = Sampled_xToHighIndex(mme, tmin)
        If (imin < 1) Then
            imin = 1
        End If
        '/* Not xToLowIndex: ensure separation of subsequent calls. */
        imax = Sampled_xToHighIndex(mme, tmax) - 1
        If (imax > mme.nx) Then
            imax = mme.nx
        End If
        If (imax < imin) Then
            Return
        End If
        imaxTarget = Sampled_xToHighIndex(thee, tmaxTarget) - 1
        distance = imaxTarget - imax
        dphase = Math.PI / (imax - imin + 1)
        For i = imin To imax Step 1
            Dim iTarget As Long = i + distance
            If (iTarget >= 1 And iTarget <= thee.nx) Then
                thee.z(1)(iTarget) += mme.z(1)(i) * 0.5 * (1 - Math.Cos(dphase * (i - imin + 0.5)))
            End If
        Next
    End Sub

    Sub copyBell(ByVal mme As Sound, ByVal tmid As Double, ByVal leftWidth As Double, ByVal rightWidth As Double, ByVal thee As Sound, ByVal tmidTarget As Double)
        copyRise(mme, tmid - leftWidth, tmid, thee, tmidTarget)
        copyFall(mme, tmid, tmid + rightWidth, thee, tmidTarget)
    End Sub

    Sub copyBell2(ByVal mme As Sound, ByVal source As PointProcess, ByVal isource As Double, ByVal leftWidth As Double,
                   ByVal rightWidth As Double, ByVal thee As Sound, ByVal tmidTarget As Double, ByVal maxT As Double)
        '/*
        '* Replace 'leftWidth' and 'rightWidth' by the lengths of the intervals in the source (instead of target),
        '* if these are shorter.
        '*/
        Dim tmid As Double = source.t(isource)
        If (isource > 1 And tmid - source.t(isource - 1) <= maxT) Then
            Dim sourceLeftWidth As Double = tmid - source.t(isource - 1)
            If (sourceLeftWidth < leftWidth) Then
                leftWidth = sourceLeftWidth
            End If
        End If
        If (isource < source.nt And source.t(isource + 1) - tmid <= maxT) Then
            Dim sourceRightWidth As Double = source.t(isource + 1) - tmid
            If (sourceRightWidth < rightWidth) Then
                rightWidth = sourceRightWidth
            End If
        End If

        copyBell(mme, tmid, leftWidth, rightWidth, thee, tmidTarget)
    End Sub

    Function PointProcess_getNearestIndex(ByVal mme As PointProcess, ByVal t As Double) As Long
        If (mme.nt = 0) Then
            Return 0
        End If
        If (t <= mme.t(1)) Then
            Return 1
        End If
        If (t >= mme.t(mme.nt)) Then
            Return mme.nt
        End If

        '/* Start binary search. */
        Dim left As Long = 1, right = mme.nt
        While (left < right - 1)
            Dim mid As Long = (left + right) / 2
            If (t >= mme.t(mid)) Then
                left = mid
            Else
                right = mid
            End If
        End While
        If Not (right = left + 1) Then
            Console.WriteLine("Cannot continue: right == left + 1")
        End If
        If (t - mme.t(left) < mme.t(right) - t) Then
            Return left
        Else
            Return right
        End If
    End Function

    Function Sampled_xToLowIndex(ByVal mme As Sampled, ByVal x As Double) As Long
        Return Convert.ToInt64(Math.Floor((x - mme.x1) / mme.dx) + 1)
    End Function
    Function Sampled_xToHighIndex(ByVal mme As Sampled, ByVal x As Double) As Long
        Return Convert.ToInt64(Math.Ceiling((x - mme.x1) / mme.dx) + 1)
    End Function

    Function Sound_Point_Pitch_Duration_to_Sound(ByRef mme As Sound, ByVal pulses As PointProcess, ByVal pitch As PitchTier, ByVal duration As DurationTier, ByVal maxT As Double) As Sound
        Try
            Dim ipointleft, ipointright As Long
            Dim deltat As Double = 0
            Dim handledTime As Double = mme.xmin
            Dim startOfSourceNoise, endOfSourceNoise, startOfTargetNoise, endOfTargetNoise As Double

            Dim durationOfSourceNoise, durationOfTargetNoise As Double
            Dim startOfSourceVoice, endOfSourceVoice, startOfTargetVoice, endOfTargetVoice As Double

            Dim durationOfSourceVoice, durationOfTargetVoice As Double

            Dim startingPeriod, finishingPeriod, ttarget, voicelessPeriod As Double

            If (duration.points.Length = 0) Then
                ' Melder_throw("No duration points.")
            End If


            '/*
            ' * Create a Sound long enough to hold the longest possible duration-manipulated sound.
            ' */
            Dim thee As New Sound(1, mme.xmin, mme.xmin + 3 * (mme.xmax - mme.xmin), 3 * mme.nx, mme.dx, mme.x1)

            '/*
            ' * Below, I'll abbreviate the voiced interval as "voice" and the voiceless interval as "noise".
            '*/
            If (Not (pitch Is Nothing) AndAlso pitch.points.Length > 0) Then
                For ipointleft = 1 To pulses.nt
                    '/*
                    ' * Find the beginning of the voice.
                    ' */
                    '/* The first pulse of the voice. */
                    startOfSourceVoice = pulses.t(ipointleft)
                    startingPeriod = 1.0 / RealTier_getValueAtTime(pitch, startOfSourceVoice)
                    '/* The first pulse is in the middle of a period. */
                    startOfSourceVoice -= 0.5 * startingPeriod

                    '/*
                    ' * Measure one noise.
                    ' */
                    startOfSourceNoise = handledTime
                    endOfSourceNoise = startOfSourceVoice
                    durationOfSourceNoise = endOfSourceNoise - startOfSourceNoise
                    startOfTargetNoise = startOfSourceNoise + deltat
                    endOfTargetNoise = startOfTargetNoise + RealTier_getArea(duration, startOfSourceNoise, endOfSourceNoise)
                    durationOfTargetNoise = endOfTargetNoise - startOfTargetNoise

                    '/*
                    ' * Copy the noise.
                    ' */
                    voicelessPeriod = NUMrandomUniform(0.008, 0.012)
                    ttarget = startOfTargetNoise + 0.5 * voicelessPeriod
                    While (ttarget < endOfTargetNoise)
                        Dim tsource As Double
                        Dim tleft As Double = startOfSourceNoise
                        Dim tright As Double = endOfSourceNoise
                        For i As Integer = 1 To 15 Step 1
                            Dim tsourcemid As Double = 0.5 * (tleft + tright)
                            Dim ttargetmid As Double = startOfTargetNoise + RealTier_getArea(duration, startOfSourceNoise, tsourcemid)
                            If ttargetmid < ttarget Then
                                tleft = tsourcemid
                            Else
                                tright = tsourcemid
                            End If
                        Next

                        tsource = 0.5 * (tleft + tright)
                        copyBell(mme, tsource, voicelessPeriod, voicelessPeriod, thee, ttarget)
                        voicelessPeriod = NUMrandomUniform(0.008, 0.012)
                        ttarget += voicelessPeriod
                    End While

                    deltat += durationOfTargetNoise - durationOfSourceNoise

                    '/*
                    ' * Find the end of the voice.
                    ' */
                    For ipointright = ipointleft + 1 To pulses.nt Step 1
                        If pulses.t(ipointright) - pulses.t(ipointright - 1) > maxT Then
                            GoTo endoffor
                        End If
                    Next
endoffor:
                    ipointright = ipointright - 1
                    '/* The last pulse of the voice. */
                    endOfSourceVoice = pulses.t(ipointright)
                    finishingPeriod = 1.0 / RealTier_getValueAtTime(pitch, endOfSourceVoice)
                    '/* The last pulse is in the middle of a period. */
                    endOfSourceVoice += 0.5 * finishingPeriod
                    '/*
                    '* Measure one voice.
                    '*/
                    durationOfSourceVoice = endOfSourceVoice - startOfSourceVoice

                    '/*
                    '* This will be copied to an interval with a different location and duration.
                    '*/
                    startOfTargetVoice = startOfSourceVoice + deltat
                    endOfTargetVoice = startOfTargetVoice + RealTier_getArea(duration, startOfSourceVoice, endOfSourceVoice)
                    durationOfTargetVoice = endOfTargetVoice - startOfTargetVoice

                    '/*
                    '* Copy the voiced part.
                    '*/
                    ttarget = startOfTargetVoice + 0.5 * startingPeriod
                    While (ttarget < endOfTargetVoice)
                        Dim tsource, period As Double
                        Dim isourcepulse As Long
                        Dim tleft As Double = startOfSourceVoice
                        Dim tright As Double = endOfSourceVoice
                        For i As Integer = 1 To 15 Step 1
                            Dim tsourcemid As Double = 0.5 * (tleft + tright)
                            Dim ttargetmid As Double = startOfTargetVoice + RealTier_getArea(duration, startOfSourceVoice, tsourcemid)
                            If (ttargetmid < ttarget) Then
                                tleft = tsourcemid
                            Else : tright = tsourcemid
                            End If
                        Next
                        tsource = 0.5 * (tleft + tright)
                        period = 1.0 / RealTier_getValueAtTime(pitch, tsource)
                        isourcepulse = PointProcess_getNearestIndex(pulses, tsource)
                        copyBell2(mme, pulses, isourcepulse, period, period, thee, ttarget, maxT)
                        ttarget += period
                    End While
                    deltat += durationOfTargetVoice - durationOfSourceVoice
                    handledTime = endOfSourceVoice
                Next
            End If


            '/*
            '* Copy the remaining unvoiced part, if we are at the end.
            '*/
            startOfSourceNoise = handledTime
            endOfSourceNoise = mme.xmax
            durationOfSourceNoise = endOfSourceNoise - startOfSourceNoise
            startOfTargetNoise = startOfSourceNoise + deltat
            endOfTargetNoise = startOfTargetNoise + RealTier_getArea(duration, startOfSourceNoise, endOfSourceNoise)
            durationOfTargetNoise = endOfTargetNoise - startOfTargetNoise
            voicelessPeriod = NUMrandomUniform(0.008, 0.012)
            ttarget = startOfTargetNoise + 0.5 * voicelessPeriod
            While (ttarget < endOfTargetNoise)
                Dim tsource As Double
                Dim tleft As Double = startOfSourceNoise, tright = endOfSourceNoise
                For i As Integer = 1 To 15 Step 1
                    Dim tsourcemid As Double = 0.5 * (tleft + tright)
                    Dim ttargetmid As Double = startOfTargetNoise + RealTier_getArea(duration, startOfSourceNoise, tsourcemid)
                    If (ttargetmid < ttarget) Then
                        tleft = tsourcemid
                    Else : tright = tsourcemid
                    End If
                Next

                tsource = 0.5 * (tleft + tright)
                copyBell(mme, tsource, voicelessPeriod, voicelessPeriod, thee, ttarget)
                voicelessPeriod = NUMrandomUniform(0.008, 0.012)
                ttarget += voicelessPeriod
            End While

            '/*
            '* Find the number of trailing zeroes and hack the sound's time domain.
            '*/
            thee.xmax = thee.xmin + RealTier_getArea(duration, mme.xmin, mme.xmax)
            '/* Common situation. */
            If (Math.Abs(thee.xmax - mme.xmax) < 0.000000000001) Then
                thee.xmax = mme.xmax
            End If
            thee.nx = Sampled_xToLowIndex(thee, thee.xmax)
            If (thee.nx > 3 * mme.nx) Then
                thee.nx = 3 * mme.nx
            End If

            Return thee
        Catch 'MelderError
            'Melder_throw (mme, ": not manipulated.");
        End Try

    End Function

    Sub Main()

    End Sub

End Module
