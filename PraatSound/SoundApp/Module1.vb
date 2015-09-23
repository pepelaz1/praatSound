Module Module1
    Public Const Melder_MULAW = 9
    Public Const Melder_ALAW = 10
    Public Const Melder_SHORTEN As Long = 11
    Public Const Melder_POLYPHONE As Long = 12
    Public Const Melder_IEEE_FLOAT_32_LITTLE_ENDIAN = 14
    Public Const Melder_AIFF As Long = 1
    Public Const Melder_AIFC As Long = 2
    Public Const Melder_WAV As Long = 3
    Public Const Melder_NEXT_SUN As Long = 4
    Public Const Melder_NIST As Long = 5
    Public Const Melder_SOUND_DESIGNER_TWO As Long = 6
    Public Const Melder_FLAC As Long = 7
    Public Const Melder_MP3 As Long = 8

    Public Const Melder_LINEAR_8_SIGNED = 1
    Public Const Melder_LINEAR_8_UNSIGNED = 2
    Public Const Melder_LINEAR_16_BIG_ENDIAN = 3
    Public Const Melder_LINEAR_16_LITTLE_ENDIAN = 4
    Public Const Melder_LINEAR_24_BIG_ENDIAN = 5
    Public Const Melder_LINEAR_24_LITTLE_ENDIAN = 6
    Public Const Melder_LINEAR_32_BIG_ENDIAN = 7
    Public Const Melder_LINEAR_32_LITTLE_ENDIAN = 8

    Public Const WAVE_FORMAT_PCM = &H1
    Public Const WAVE_FORMAT_IEEE_FLOAT = &H3
    Public Const WAVE_FORMAT_ALAW = &H6
    Public Const WAVE_FORMAT_MULAW = &H7
    Public Const WAVE_FORMAT_DVI_ADPCM = &H11

    Public Const Pitch_NEAREST = 0
    Public Const Pitch_LINEAR = 1

    Public Const Pitch_LEVEL_FREQUENCY = 1
    Public Const Pitch_LEVEL_STRENGTH = 2

    Public Const Vector_CHANNEL_AVERAGE = 0
    Public Const Vector_CHANNEL_1 = 1
    Public Const Vector_CHANNEL_2 = 2

    Public Const Vector_VALUE_INTERPOLATION_NEAREST = 0
    Public Const Vector_VALUE_INTERPOLATION_LINEAR = 1
    Public Const Vector_VALUE_INTERPOLATION_CUBIC = 2
    Public Const Vector_VALUE_INTERPOLATION_SINC70 = 3
    Public Const Vector_VALUE_INTERPOLATION_SINC700 = 4

    Public Const NUM_VALUE_INTERPOLATE_NEAREST = 0
    Public Const NUM_VALUE_INTERPOLATE_LINEAR = 1
    Public Const NUM_VALUE_INTERPOLATE_CUBIC = 2
    '// Higher values than 2 yield a true sinc interpolation. Here are some examples:
    Public Const NUM_VALUE_INTERPOLATE_SINC70 = 70
    Public Const NUM_VALUE_INTERPOLATE_SINC700 = 700


    Public Const NUM_PEAK_INTERPOLATE_NONE = 0
    Public Const NUM_PEAK_INTERPOLATE_PARABOLIC = 1
    Public Const NUM_PEAK_INTERPOLATE_CUBIC = 2
    Public Const NUM_PEAK_INTERPOLATE_SINC70 = 3
    Public Const NUM_PEAK_INTERPOLATE_SINC700 = 4

    Public Const MAX_T = 0.02000000001

    Public NUM_goldenSection As Double = 0.6180339887498949
    '0.6180339887498948482045868343656381177203

    Public Const NUMundefined = System.Double.PositiveInfinity
    Public NUMEulersConstant As Double = Math.E
    Public NUMLog2E As Double = Math.Log(NUMEulersConstant, 2)

    Public Const AC_HANNING = 0
    Public Const AC_GAUSS = 1
    Public Const FCC_NORMAL = 2
    Public Const FCC_ACCURATE = 3

    Public Melder_debug As Long = 0

    'Private Declare Function GetAddrOf Lib "KERNEL32" Alias "MulDiv" (nNumber As Integer, Optional ByVal nNumerator As Integer = 1, Optional ByVal nDenominator As Integer = 1) As Long
    Declare Function GetAddrOf Lib "msvbvm50.dll" Alias "VarPtr" (ByVal Var As Integer) As Long
    Private Declare Sub MoveMemory Lib "KERNEL32" Alias "RtlMoveMemory" (hpvDest As Long, hpvSource As Long, ByVal cbCopy As Integer)
    Function AnyTier_timeToLowIndex(ByVal mme As RealTier, ByVal time As Double) As Long
        '// undefined
        If (mme.points.Length = 0) Then
            Return 0
        End If
        Dim ileft As Long = 0, iright = mme.points.Length - 1
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
        Dim imid As Long
        While (iright > ileft + 1)
            imid = (ileft + iright) / 2
            Dim tmid As Double = points(imid).number
            If (time < tmid) Then
                iright = imid
                tright = tmid
            Else
                ileft = imid
                tleft = tmid
            End If
        End While
        ' If Not (iright = ileft + 1) Then
        '            Console.WriteLine("Execution stopped: iright == ileft + 1")
        '            Return -1
        '        End If
        '        If Not (ileft >= 0) Then
        '            Console.WriteLine("Execution stopped: ileft >= 1")
        '            Return -1
        '        End If
        '        If Not (iright <= mme.points.Length) Then
        '            Console.WriteLine("Execution stopped: (iright <= mme.points.Length)")
        '            Return -1
        '        End If
        '        If Not (time >= points(ileft).number) Then
        '            Console.WriteLine("Execution stopped: (time >= points(ileft).number)")
        '            Return -1
        '        End If
        '        If Not (time <= points(iright).number) Then
        '            Console.WriteLine("Execution stopped: (time <= points(iright).number)")
        '            Return -1
        '        End If
Endofwhile: Return ileft
    End Function

    Function AnyTier_timeToHighIndex(ByVal mme As RealTier, ByVal time As Double) As Long
        '// undefined
        If (mme.points.Length = 0) Then
            Return 0
        End If
        Dim ileft As Long = 0, iright = mme.points.Length - 1
        Dim points() As RealPoint = mme.points
        Dim tleft As Double = points(ileft).number
        If (time <= tleft) Then
            Return 0
        End If
        Dim tright As Double = points(iright).number
        ' // offright
        If (time > tright) Then
            Return iright
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

    Function RealTier_getArea(ByRef mme As DurationTier, ByVal tmin As Double, ByVal tmax As Double) As Double
        Dim n As Long = mme.f_getNumberOfPoints()
        Dim imin, imax As Long
        Dim points() As RealPoint = mme.points
        If (n = 1) Then
            Return (tmax - tmin) * points(0).value
        End If
        imin = AnyTier_timeToLowIndex(mme, tmin)
        If Not (imin = n) Then
            Return (tmax - tmin) * points(n - 1).value
        End If
        imax = AnyTier_timeToHighIndex(mme, tmax)
        If (imax = 1) Then
            Return (tmax - tmin) * points(0).value
        End If
        If (imin < n - 1) Then
            Console.WriteLine("Execution stopped: imin<n")
            Return -1
        End If
        If (imax > 0) Then
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
        Dim pointRight As RealPoint = mme.points(0)
        '/* Constant extrapolation. */
        If (t <= pointRight.number) Then
            Return pointRight.value
        End If
        Dim pointLeft As RealPoint = mme.points(n - 1)
        '/* Constant extrapolation. */
        If (t >= pointLeft.number) Then
            Return pointLeft.value
        End If
        If Not (n >= 2) Then
            Console.WriteLine("Execution stopped: n>=2")
            Return -1
        End If
        Dim ileft As Long = AnyTier_timeToLowIndex(mme, t), iright = ileft + 1
        If Not (ileft >= 0 Or iright <= n - 1) Then
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
            '/* Unusual, but possible no preference. */
            Return 0.5 * (fleft + fright)
        End If
        ' /* Linear interpolation. */
        Return fleft + (t - tleft) * (fright - fleft) / (tright - tleft)
    End Function

    Sub copyFall(ByVal mme As Sound, ByVal tmin As Double, ByVal tmax As Double, ByRef thee As Sound, ByVal tminTarget As Double)
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
        For i = imin - 1 To imax - 1 Step 1
            Dim iTarget As Long = i + distance
            If (iTarget >= 0 And iTarget <= thee.nx - 1) Then
                thee.z(0, iTarget) += mme.z(0, i) * 0.5 * (1 + Math.Cos(dphase * (i + 1 - imin + 0.5)))
            End If
        Next
    End Sub

    Sub copyRise(ByVal mme As Sound, ByVal tmin As Double, ByVal tmax As Double, ByRef thee As Sound, ByVal tmaxTarget As Double)
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
        For i = imin - 1 To imax - 1 Step 1
            Dim iTarget As Long = i + distance
            If (iTarget >= 0 And iTarget <= thee.nx - 1) Then
                thee.z(0, iTarget) += mme.z(0, i) * 0.5 * (1 - Math.Cos(dphase * (i + 1 - imin + 0.5)))
            End If
        Next
    End Sub

    Sub copyBell(ByVal mme As Sound, ByVal tmid As Double, ByVal leftWidth As Double, ByVal rightWidth As Double, ByRef thee As Sound, ByVal tmidTarget As Double)
        copyRise(mme, tmid - leftWidth, tmid, thee, tmidTarget)
        copyFall(mme, tmid, tmid + rightWidth, thee, tmidTarget)
    End Sub

    Sub copyBell2(ByRef mme As Sound, ByRef source As PointProcess, ByVal isource As Double, ByRef leftWidth As Double,
                   ByRef rightWidth As Double, ByRef thee As Sound, ByVal tmidTarget As Double, ByVal maxT As Double)
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
        If (t <= mme.t(0)) Then
            Return 1
        End If
        If (t >= mme.t(mme.nt)) Then
            Return mme.nt
        End If

        '/* Start binary search. */
        Dim left As Long = 0, right = mme.nt - 1
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

    Function Sound_Point_Pitch_Duration_to_Sound(ByRef mme As Sound, ByRef pulses As PointProcess, ByRef pitch As PitchTier, ByRef duration As DurationTier, ByVal maxT As Double) As Sound

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
            ipointleft = 1
            While ipointleft <= pulses.nt
                '/*
                ' * Find the beginning of the voice.
                ' */
                '/* The first pulse of the voice. */
                If ipointleft = 321 Then
                    Dim aaaa As Long = 359
                End If
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
                        GoTo endofswitch
                    End If
                Next
endofswitch:
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
                    If isourcepulse = 360 Then
                        Dim ttttt As Long = 9
                    End If
                    period = 1.0 / RealTier_getValueAtTime(pitch, tsource)
                    isourcepulse = PointProcess_getNearestIndex(pulses, tsource)
                    copyBell2(mme, pulses, isourcepulse, period, period, thee, ttarget, maxT)
                    ttarget += period
                End While
                deltat += durationOfTargetVoice - durationOfSourceVoice
                handledTime = endOfSourceVoice
                ipointleft = ipointright + 1
            End While
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

    End Function
    Public Function Sound_createSimple(ByVal numberOfChannels As Long, ByVal duration As Double, ByVal samplingFrequency As Double) As Sound
        Return New Sound(numberOfChannels, 0.0, duration, Math.Floor(duration * samplingFrequency + 0.5), 1 / samplingFrequency, 0.5 / samplingFrequency)
    End Function
    Function bingeti2LE(ByRef reader As System.IO.BinaryReader) As Integer
        If (Melder_debug <> 18) Then
            Dim s As Short
            s = reader.ReadInt16()
            'if (fread (& s, sizeof (signed short), 1, f) <> 1) Then 
            ' readError(f, "a signed short integer.")
            ' End If
            Return Convert.ToInt32(s)
            '// with sign extension if an int is 4 bytes
        Else
            Dim bytes(1) As Byte
            bytes = reader.ReadBytes(2)
            If (bytes.GetLength(0) <> 2) Then
                Console.WriteLine("Cannot read two bytes.")
                Return -1
            End If
            Dim externalValue As UShort = (Convert.ToUInt16(bytes(1)) << 8) Or Convert.ToUInt16(bytes(0))
            Return Convert.ToInt32(Convert.ToInt16(externalValue))
            '// with sign extension if an int is 4 bytes
        End If
    End Function
    Function bingeti4LE(ByRef reader As System.IO.BinaryReader) As Long
        If (Melder_debug <> 18) Then
            Dim l As Long
            l = reader.ReadUInt32()
            'if (fread (& l, sizeof (long), 1, f) <> 1) Then
            ' Console.WriteLine("Error reading a signed long integer.")
            'End If
            Return l
        Else
            Dim bytes(3) As Byte
            bytes = reader.ReadBytes(4)
            If (bytes.GetLength(0) <> 4) Then
                Console.WriteLine("Error reading four bytes.")
                Return -1
            End If
            Return ((Convert.ToUInt32(bytes(3)) << 24) Or (Convert.ToUInt32(bytes(2)) << 16) Or (Convert.ToUInt32(bytes(1)) << 8) Or Convert.ToUInt32(bytes(0)))
            ' // 32 or 64 bits
        End If
    End Function
    Public Function PutSign(ByVal number As UShort) As Short
        If number > 32768 Then 'negative number
            Return (65536 - number) * -1
        Else
            Return number
        End If
    End Function
    Sub Melder_readAudioToFloat(ByRef reader As System.IO.BinaryReader, ByVal numberOfChannels As Integer, ByVal encoding As Integer, ByRef buffer(,) As Double, ByVal numberOfSamples As Long)
        Dim j As Long = 1
        Try
            Select Case encoding
                Case Melder_LINEAR_16_LITTLE_ENDIAN
                    Dim numberOfBytesPerSamplePerChannel As Integer = 2
                    '(int) sizeof (double) = 8
                    If (numberOfChannels > 8 / numberOfBytesPerSamplePerChannel) Then
                        For isamp As Long = 0 To numberOfSamples - 1 Step 1
                            For ichan As Long = 0 To numberOfChannels - 1 Step 1
                                buffer(ichan, isamp) = bingeti2LE(reader) * (1.0 / 32768)
                            Next
                        Next
                    Else
                        Dim numberOfBytes As Long = numberOfChannels * numberOfSamples * numberOfBytesPerSamplePerChannel
                        Dim bytes() As Byte = reader.ReadBytes(numberOfBytes)
                        '  For i As Integer = 0 To numberOfBytes - 1 Step 1

                        '   Next
                        '  buffer(0, ) = reader.ReadBytes(numberOfBytes)
                        ' Write bytes to buffer at specific address
                        ' last quarter of buffer
                        Dim bufStartIndex = Math.Floor(numberOfSamples * 0.75)
                        'number of bytes to be replaced in buffer
                        Dim N As Long = 8 - (((numberOfSamples * 3) Mod 4) * 2)
                        'save double value
                        Dim d As Double = buffer(0, bufStartIndex)
                        ' clear N bytes
                        If (N <> 0) Then
                            'Make needed bytes as 0
                            Dim mult As Double
                            If N = 2 Then
                                mult = &HFFFFFFFFFFFF0000
                            ElseIf N = 4 Then
                                mult = &HFFFFFFFF00000000
                            ElseIf N = 6 Then
                                mult = &HFFFF000000000000
                            End If
                            d = d And mult
                            For counter As Long = 0 To N - 1 Step 1
                                d = d Or (bytes(N - 1 - counter) << (8 * counter))
                            Next
                            'put it to buffer
                            buffer(0, bufStartIndex) = d
                        End If
                        Dim bytesV(numberOfBytes - N) As Byte
                        For ind As Integer = N To numberOfBytes - 1 Step 1
                            bytesV(ind - N) = bytes(ind)
                        Next
                        Dim ms As System.IO.MemoryStream = New System.IO.MemoryStream(bytesV)
                        Dim br As System.IO.BinaryReader = New IO.BinaryReader(ms)
                        Dim curDouble As Double
                        While j < numberOfSamples - bufStartIndex
                            curDouble = br.ReadDouble()
                            buffer(0, bufStartIndex + j) = curDouble
                            j = j + 1
                        End While
                        'Dim tmpAddr As Long = GetAddrOf(buffer(0, 0))
                        'Dim tmpAddr1 As Long = GetAddrOf(buffer(numberOfChannels - 1, numberOfSamples - 1))
                        'Dim dstAddr As Long = tmpAddr + 8 - numberOfBytes
                        'Dim srcAddr As Long = GetAddrOf(bytes(0))
                        'MoveMemory(dstAddr, srcAddr, numberOfBytes)
                        'If (fread(bytes, 1, numberOfBytes, f) < numberOfBytes) Then
                        'Console.WriteLine("Error in ReadAudioToFloat")
                        'Return
                        'End If
                        '// read 16-bit data into last quarter of buffer
                        Dim i As Long = 0
                        If (numberOfChannels = 1) Then
                            For isamp As Long = 0 To numberOfSamples - 1 Step 1
                                Dim byte1 As Byte = bytes(i)
                                i = i + 1
                                Dim byte2 = bytes(i)
                                i = i + 1
                                Dim value As Short = PutSign(((Convert.ToUInt16(byte2) << 8) Or Convert.ToUInt16(byte1)))
                                buffer(0, isamp) = value * (1.0 / 32768)
                            Next
                        Else
                            For isamp As Long = 1 To numberOfSamples Step 1
                                For ichan As Long = 1 To numberOfChannels Step 1
                                    Dim byte1 As Byte = bytes(i)
                                    i = i + 1
                                    Dim byte2 = bytes(i)
                                    i = i + 1
                                    Dim b2 As UShort = Convert.ToUInt16(byte2)
                                    Dim b1 As UShort = Convert.ToUInt16(byte1)
                                    Dim value As Long = Convert.ToInt16((b2 << 8) Or b1)
                                    buffer(ichan, isamp) = value * (1.0 / 32768)
                                Next
                            Next
                        End If
                    End If
            End Select
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub

    Sub MelderFile_writeFloatToAudio(ByRef writer As System.IO.BinaryWriter, ByVal numberOfChannels As Integer, ByVal encoding As Integer, ByRef buffer(,) As Double, ByVal numberOfSamples As Long)
        For isamp As Long = 0 To numberOfSamples - 1 Step 1
            For ichan As Long = 0 To numberOfChannels - 1 Step 1
                Dim value As Long = Math.Round(buffer(ichan, isamp) * 32768)
                If (value < -32768) Then
                    value = -32768
                End If
                If (value > 32767) Then
                    value = 32767
                End If
                Dim r As Short
                r = value
                writer.Write(r)
            Next
        Next
    End Sub
    Sub Melder_checkWavFile(ByRef reader As System.IO.BinaryReader, ByRef numberOfChannels As Integer, ByRef encoding As Integer, ByRef sampleRate As Double, ByRef startOfData As Long, ByRef numberOfSamples As Long)
        Dim data(7), chunkID(3) As Byte
        Dim formatChunkPresent As Boolean = False, dataChunkPresent = False
        Dim numberOfBitsPerSamplePoint As Integer = -1
        Dim dataChunkSize As Long = -1
        data = reader.ReadBytes(4)
        If data.GetLength(0) < 4 Then
            Console.WriteLine("File too small: no RIFF statement.")
            Return
        End If
        Dim dataStr As String = System.Text.Encoding.ASCII.GetString(data)
        If (Not dataStr.Substring(0, 4) = "RIFF") Then
            Console.WriteLine("Not a WAV file (RIFF statement expected).")
            Return
        End If
        data = reader.ReadBytes(4)
        dataStr = System.Text.Encoding.ASCII.GetString(data)
        If data.GetLength(0) < 4 Then
            Console.WriteLine("File too small: no size of RIFF chunk.")
            Return
        End If
        data = reader.ReadBytes(4)
        If data.GetLength(0) < 4 Then
            Console.WriteLine("File too small: no file type info (expected WAVE statement).")
            Return
        End If
        dataStr = System.Text.Encoding.ASCII.GetString(data)
        If (dataStr.Substring(0, 4) <> "WAVE" And dataStr.Substring(0, 4) <> "CDDA") Then
            Console.WriteLine("Not a WAVE or CD audio file (wrong file type info).")
            Return
        End If
        chunkID = reader.ReadBytes(4)
        ' /* Search for Format Chunk and Data Chunk. */
        dataStr = System.Text.Encoding.ASCII.GetString(chunkID)
        While (dataStr.Length = 4)
            Dim chunkSize As Long = bingeti4LE(reader)
            ' If (Melder_debug = 23) Then
            '        Melder_warning (Melder_integer (chunkID [0]), L" ", Melder_integer (chunkID [1]), L" ", Melder_integer (chunkID [2]), L" ", Melder_integer (chunkID [3]), L" ", Melder_integer (chunkSize))
            'End If
            If (dataStr.StartsWith("fmt ")) Then
                '/*
                ' * Found a Format Chunk.
                ' */
                Dim winEncoding As Integer = bingeti2LE(reader)
                formatChunkPresent = True
                numberOfChannels = bingeti2LE(reader)
                If (numberOfChannels < 1) Then
                    Console.WriteLine("Too few sound channels (" + numberOfChannels + ").")
                    Return
                End If
                sampleRate = Convert.ToDouble(bingeti4LE(reader))
                If (sampleRate <= 0.0) Then
                    Console.WriteLine("Wrong sampling frequency (" + sampleRate + " Hz).")
                    Return
                End If
                '// avgBytesPerSec
                bingeti4LE(reader)
                '// blockAlign
                bingeti2LE(reader)
                numberOfBitsPerSamplePoint = bingeti2LE(reader)
                If (numberOfBitsPerSamplePoint = 0) Then
                    numberOfBitsPerSamplePoint = 16
                ElseIf (numberOfBitsPerSamplePoint < 4) Then
                    Console.WriteLine("Too few bits per sample (" + numberOfBitsPerSamplePoint + " the minimum is 4).")
                    Return
                ElseIf (numberOfBitsPerSamplePoint > 32) Then
                    Console.WriteLine("Too few bits per sample (" + numberOfBitsPerSamplePoint + " the maximum is 32).")
                    Return
                End If
                Select Case winEncoding
                    Case WAVE_FORMAT_PCM
                        If numberOfBitsPerSamplePoint > 24 Then
                            encoding = Melder_LINEAR_32_LITTLE_ENDIAN
                        ElseIf numberOfBitsPerSamplePoint > 16 Then
                            encoding = Melder_LINEAR_24_LITTLE_ENDIAN
                        ElseIf numberOfBitsPerSamplePoint > 8 Then
                            encoding = Melder_LINEAR_16_LITTLE_ENDIAN
                        Else
                            encoding = Melder_LINEAR_8_UNSIGNED
                        End If
                        GoTo endofswitch
                    Case WAVE_FORMAT_IEEE_FLOAT
                        encoding = Melder_IEEE_FLOAT_32_LITTLE_ENDIAN
                        GoTo endofswitch
                    Case WAVE_FORMAT_ALAW
                        encoding = Melder_ALAW
                        GoTo endofswitch
                    Case WAVE_FORMAT_MULAW
                        encoding = Melder_MULAW
                        GoTo endofswitch
                    Case WAVE_FORMAT_DVI_ADPCM
                        Console.WriteLine("Cannot read lossy compressed audio files (this one is DVI ADPCM).\n" +
                            "Please use uncompressed audio files. If you must open this file,\n" +
                            "please use an audio converter program to convert it first to normal (PCM) WAV format\n" +
                            "(Praat may have difficulty analysing the poor recording, though).")
                        Return
                    Case Else
                        Console.WriteLine("Unsupported Windows audio encoding %d." + winEncoding)
                        Return
                End Select
endofswitch:    If (chunkSize And 1) Then
                    chunkSize = chunkSize + 1
                End If
                For i As Long = 17 To chunkSize Step 1
                    data = reader.ReadBytes(1)
                    If (data.GetLength(0) < 1) Then
                        Console.WriteLine("File too small: expected " + chunkSize + " bytes in fmt chunk, but found " + i + ".")
                        Return
                    End If
                Next
            Else
                If (dataStr.StartsWith("data")) Then
                    '/*
                    ' * Found a Data Chunk.
                    ' */
                    dataChunkPresent = True
                    dataChunkSize = chunkSize
                    startOfData = reader.BaseStream.Position
                    If (chunkSize And 1) Then
                        chunkSize = chunkSize + 1
                    End If
                    If (chunkSize < 0) Then
                        '// incorrect data chunk (sometimes -44) assume that the data run till the end of the file
                        reader.BaseStream.Seek(0, System.IO.SeekOrigin.End)
                        Dim endOfData As Long = reader.BaseStream.Position
                        dataChunkSize = chunkSize = endOfData - startOfData
                        reader.BaseStream.Seek(startOfData, IO.SeekOrigin.Begin)
                    End If
                    If (Melder_debug = 23) Then
                        For i As Long = 1 To chunkSize Step 1
                            data = reader.ReadBytes(1)
                            If (data.GetLength(0) < 1) Then
                                Console.WriteLine("File too small: expected " + chunkSize + " bytes of data, but found " + i + ".")
                                Return
                            End If
                        Next
                    Else
                        If (formatChunkPresent) Then
                            GoTo endoffor
                        End If
                    End If
                    '// OPTIMIZATION: do not read the whole data chunk if we have already read the format chunk
                Else
                    '// ignore other chunks
                    If (chunkSize And 1) Then
                        chunkSize = chunkSize + 1
                    End If
                    For i As Long = 1 To chunkSize Step 1
                        data = reader.ReadBytes(1)
                        If (data.GetLength(0) < 1) Then
                            Console.WriteLine("File too small: expected " + chunkSize, " bytes, but found " + i + ".")
                            Return
                        End If
                    Next

                End If
            End If
            chunkID = reader.ReadBytes(4)
            ' /* Search for Format Chunk and Data Chunk. */
            dataStr = System.Text.Encoding.ASCII.GetString(chunkID)
        End While

endoffor: If (Not formatChunkPresent) Then
            Console.WriteLine("Found no Format Chunk.")
            Return
        End If
        If (Not dataChunkPresent) Then
            Console.WriteLine("Found no Data Chunk.")
            Return
        End If
        If (Not (numberOfBitsPerSamplePoint <> -1 And dataChunkSize <> -1)) Then
            Console.WriteLine("Cannot continue execution: numberOfBitsPerSamplePoint <> -1 And dataChunkSize <> -1")
            Return
        End If
        Dim d As Long = Math.Floor((numberOfBitsPerSamplePoint + 7) / 8)
        numberOfSamples = dataChunkSize / numberOfChannels / d
    End Sub

    Function MelderFile_checkSoundFile(ByRef reader As System.IO.BinaryReader, ByRef numberOfChannels As Integer, ByRef encoding As Integer, ByRef sampleRate As Double, ByRef startOfData As Long, ByRef numberOfSamples As Long) As Integer
        Dim data As Byte()
        data = reader.ReadBytes(16)
        If (reader Is Nothing Or data.GetLength(0) < 16) Then
            Return 0
        End If
        reader.BaseStream.Seek(0, System.IO.SeekOrigin.Begin)
        Dim DataStr As String = System.Text.Encoding.ASCII.GetString(data)
        'If (strnequ(data, "FORM", 4) And strnequ(data + 8, "AIFF", 4)) Then
        'Melder_checkAiffFile(reader, numberOfChannels, encoding, sampleRate, startOfData, numberOfSamples)
        'Return Melder_AIFF
        'End If
        'If (strnequ(data, "FORM", 4) And strnequ(data + 8, "AIFC", 4)) Then
        '    Melder_checkAiffFile(reader, numberOfChannels, encoding, sampleRate, startOfData, numberOfSamples)
        '    Return Melder_AIFC
        'End If
        If (DataStr.Substring(0, 4) = "RIFF" And (DataStr.Substring(8, 4) = "WAVE" Or DataStr.Substring(8, 4) = "CDDA")) Then
            Melder_checkWavFile(reader, numberOfChannels, encoding, sampleRate, startOfData, numberOfSamples)
            Return Melder_WAV
        End If
        'If (strnequ(data, ".snd", 4)) Then
        '    Melder_checkNextSunFile(reader, numberOfChannels, encoding, sampleRate, startOfData, numberOfSamples)
        '    Return Melder_NEXT_SUN
        'End If
        'If (strnequ(data, "NIST_1A", 7)) Then
        '    Melder_checkNistFile(reader, numberOfChannels, encoding, sampleRate, startOfData, numberOfSamples)
        '    Return Melder_NIST
        'End If
        Return 0
    End Function
    Public Function Sound_readFromSoundFile(ByVal filePath As String) As Sound
        Try
            ' autoMelderFile mfile = MelderFile_open (file)
            Dim reader = New System.IO.BinaryReader(System.IO.File.Open(filePath, System.IO.FileMode.Open))
            Dim numberOfChannels, encoding As Integer
            Dim sampleRate As Double
            Dim startOfData, numberOfSamples As Long
            Dim fileType As Integer = MelderFile_checkSoundFile(reader, numberOfChannels, encoding, sampleRate, startOfData, numberOfSamples)
            If (fileType = 0) Then
                Console.WriteLine("Not an audio file.")
                Return Nothing
            End If
            '// start from beginning of Data Chunk
            If (reader.BaseStream.Seek(startOfData, IO.SeekOrigin.Begin) = IO.SeekOrigin.End) Then
                Console.WriteLine("No data in audio file.")
                Return Nothing
            End If
            Dim mme As Sound = Sound_createSimple(numberOfChannels, numberOfSamples / sampleRate, sampleRate)
            If (encoding = Melder_SHORTEN Or encoding = Melder_POLYPHONE) Then
                Console.WriteLine("Cannot unshorten. Write to paul.boersma@uva.nl for more information.")
                Return Nothing
            End If
            Melder_readAudioToFloat(reader, numberOfChannels, encoding, mme.z, numberOfSamples)
            reader.Close()
            Return mme
        Catch ex As Exception
            Console.WriteLine("Sound not read from sound file " + filePath + ": " + ex.Message)
            Return Nothing
        End Try
    End Function

    Public Function Log2X(ByVal x As Double) As Double
        Dim u As Double = NUMLog2E
        Dim rr As Double = Math.Log10(x)
        Dim rr1 As Double = Math.Log(x, 2)
        Return rr1
    End Function
    Public Function exp(ByVal x As Double) As Double
        Return Math.Pow(NUMEulersConstant, x)
    End Function

    Sub Sampled_shortTermAnalysis(ByRef mme As Sampled, ByVal windowDuration As Double, ByVal timeStep As Double, ByRef numberOfFrames As Long, ByRef firstTime As Double)
        '!!!Melder_assert (windowDuration > 0.0)
        '!!!Melder_assert (timeStep > 0.0)
        Dim myDuration As Double = mme.dx * mme.nx
        If (windowDuration > myDuration) Then
            Console.WriteLine(": shorter than window length.")
        End If
        numberOfFrames = Math.Floor((myDuration - windowDuration) / timeStep) + 1
        '!!!Melder_assert (*numberOfFrames >= 1)
        Dim ourMidTime As Double = mme.x1 - 0.5 * mme.dx + 0.5 * myDuration
        Dim thyDuration As Double = numberOfFrames * timeStep
        firstTime = ourMidTime - 0.5 * thyDuration + 0.5 * timeStep
    End Sub

    Sub Pitch_pathFinder(ByRef mme As Pitch, ByVal silenceThreshold As Double, ByVal voicingThreshold As Double, ByVal octaveCost As Double, ByRef octaveJumpCost As Double, ByRef voicedUnvoicedCost As Double, ByVal ceiling As Double, ByVal pullFormants As Integer)
        '      if (Melder_debug == 33) Melder_casual ("Pitch path finder:\nSilence threshold = %g\nVoicing threshold = %g\nOctave cost = %g\nOctave jump cost = %g\n"
        '"Voiced/unvoiced cost = %g\nCeiling = %g\nPull formants = %d", silenceThreshold, voicingThreshold, octaveCost, octaveJumpCost, voicedUnvoicedCost,
        'ceiling, pullFormants);
        Try
            Dim maxnCandidates As Long = mme.maxCandidates
            Dim place As Long
            Dim maximum, value As Double
            Dim ceiling2 As Double
            If pullFormants Then
                ceiling2 = 2 * ceiling
            Else
                ceiling2 = ceiling
            End If
            '/* Next three lines 20011015 */
            Dim timeStepCorrection As Double = 0.01 / mme.dx
            octaveJumpCost *= timeStepCorrection
            voicedUnvoicedCost *= timeStepCorrection

            mme.ceiling = ceiling
            Dim delta(mme.nx - 1, maxnCandidates - 1) As Double
            Dim psi(mme.nx - 1, maxnCandidates - 1) As Long

            For iframe As Long = 0 To mme.nx - 1 Step 1
                Dim frame As Pitch_Frame = mme.frame(iframe)
                Dim unvoicedStrength As Double
                If silenceThreshold <= 0 Then
                    unvoicedStrength = 0
                Else
                    unvoicedStrength = 2 - frame.intensity / (silenceThreshold / (1 + voicingThreshold))
                End If
                If unvoicedStrength > 0 Then
                    unvoicedStrength = voicingThreshold + unvoicedStrength
                Else
                    unvoicedStrength = voicingThreshold
                End If
                For icand As Long = 0 To frame.nCandidates - 1 Step 1
                    Dim candidate As PitchCandidate = frame.candidates(icand)
                    Dim voiceless As Integer
                    If candidate.frequency = 0 Or candidate.frequency > ceiling2 Then
                        voiceless = 1
                    Else
                        voiceless = 0
                    End If
                    If voiceless <> 0 Then
                        delta(iframe, icand) = unvoicedStrength
                    Else
                        delta(iframe, icand) = candidate.strength - octaveCost * Log2X(ceiling / candidate.frequency)
                    End If
                Next
            Next

            '/* Look for the most probable path through the maxima. */
            '/* There is a cost for the voiced/unvoiced transition, */
            '/* and a cost for a frequency jump. */

            For iframe As Long = 1 To mme.nx - 1 Step 1
                Dim prevFrame As Pitch_Frame = mme.frame(iframe - 1), curFrame = mme.frame(iframe)
                'double *prevDelta = delta [iframe - 1], *curDelta = delta [iframe]
                'long *curPsi = psi [iframe]
                For icand2 As Long = 0 To curFrame.nCandidates - 1 Step 1
                    Dim f2 As Double = curFrame.candidates(icand2).frequency
                    maximum = -1.0E+30
                    place = 0
                    For icand1 As Long = 0 To prevFrame.nCandidates - 1 Step 1
                        Dim f1 As Double = prevFrame.candidates(icand1).frequency
                        Dim transitionCost As Double
                        Dim previousVoiceless As Boolean = f1 <= 0 Or f1 >= ceiling2
                        Dim currentVoiceless As Boolean = f2 <= 0 Or f2 >= ceiling2
                        If (currentVoiceless) Then
                            If (previousVoiceless) Then
                                transitionCost = 0
                                '// both voiceless
                            Else
                                transitionCost = voicedUnvoicedCost
                                '// voiced-to-unvoiced transition
                            End If
                        Else
                            If (previousVoiceless) Then
                                transitionCost = voicedUnvoicedCost
                                '// unvoiced-to-voiced transition
                            Else
                                transitionCost = octaveJumpCost * Math.Abs(Log2X(f1 / f2))
                                '   // both voiced
                            End If
                        End If
                        value = delta(iframe - 1, icand1) - transitionCost + delta(iframe, icand2)
                        '//if (Melder_debug == 33) Melder_casual ("Frame %ld, current candidate %ld (delta %g), previous candidate %ld (delta %g), "
                        '//	"transition cost %g, value %g, maximum %g", iframe, icand2, curDelta [icand2], icand1, prevDelta [icand1], transitionCost, value, maximum);
                        If (value > maximum) Then
                            maximum = value
                            place = icand1

                            'Else
                            'if (value == maximum) {
                            'if (Melder_debug == 33) Melder_casual ("A tie in frame %ld, current candidate %ld, previous candidate %ld", iframe, icand2, icand1);
                            '}
                        End If
                    Next
                    delta(iframe, icand2) = maximum
                    psi(iframe, icand2) = place
                Next
            Next

            '/* Find the end of the most probable path. */

            place = 0
            maximum = delta(mme.nx - 1, place)
            For icand As Long = 1 To mme.frame(mme.nx - 1).nCandidates - 1 Step 1
                If (delta(mme.nx - 1, icand) > maximum) Then
                    place = icand
                    maximum = delta(mme.nx - 1, place)
                End If
            Next

            '/* Backtracking: follow the path backwards. */

            For iframe = mme.nx - 1 To 0 Step -1
                'if (Melder_debug == 33) Melder_casual ("Frame %ld: swapping candidates 1 and %ld", iframe, place);
                Dim frame As Pitch_Frame = mme.frame(iframe)
                Dim help As PitchCandidate = frame.candidates(0)
                frame.candidates(0) = frame.candidates(place)
                frame.candidates(place) = help
                place = psi(iframe, place)
                '// This assignment is challenging to CodeWarrior 11.
            Next

            '/* Pull formants: devoice frames with frequencies between ceiling and ceiling2. */

            If (ceiling2 > ceiling) Then
                'if (Melder_debug == 33) Melder_casual ("Pulling formants...");
                For iframe As Long = mme.nx - 1 To 0 Step -1
                    Dim frame As Pitch_Frame = mme.frame(iframe)
                    Dim winner As PitchCandidate = frame.candidates(0)
                    Dim f As Double = winner.frequency
                    If (f > ceiling And f <= ceiling2) Then
                        For icand As Long = 1 To frame.nCandidates - 1 Step 1
                            Dim loser As PitchCandidate = frame.candidates(icand)
                            If (loser.frequency = 0.0) Then
                                Dim help As PitchCandidate = winner
                                winner = loser
                                loser = help
                                GoTo endoffor
                            End If
                        Next
endoffor:           End If
                Next
            End If
        Catch ex As Exception
            'Melder_throw (me, ": path not found.");
        End Try
    End Sub
    Sub dradb2(ByVal ido As Long, ByVal l1 As Long, ByRef cc() As Double, ByRef ch() As Double, ByRef wa1() As Double, ByVal wa1Index As Long)
        'wa1Index - index at wa1
        Dim i, k, t0, t1, t2, t3, t4, t5, t6 As Long
        Dim ti2, tr2 As Double

        t0 = l1 * ido

        t1 = 0
        t2 = 0
        t3 = (ido << 1) - 1
        For k = 0 To l1 - 1 Step 1

            ch(t1) = cc(t2) + cc(t3 + t2)
            ch(t1 + t0) = cc(t2) - cc(t3 + t2)
            t1 += ido
            t2 = t1 << 1
        Next

        If (ido < 2) Then
            Return
        End If
        If (ido = 2) Then
            GoTo L105
        End If

        t1 = 0
        t2 = 0
        For k = 0 To l1 - 1 Step 1
            t3 = t1
            t5 = (t4 = t2) + (ido << 1)
            t6 = t0 + t1
            For i = 2 To ido - 1 Step 2

                t3 += 2
                t4 += 2
                t5 -= 2
                t6 += 2
                ch(t3 - 1) = cc(t4 - 1) + cc(t5 - 1)
                tr2 = cc(t4 - 1) - cc(t5 - 1)
                ch(t3) = cc(t4) - cc(t5)
                ti2 = cc(t4) + cc(t5)
                ch(t6 - 1) = wa1(wa1Index + i - 2) * tr2 - wa1(wa1Index + i - 1) * ti2
                ch(t6) = wa1(wa1Index + i - 2) * ti2 + wa1(wa1Index + i - 1) * tr2
            Next
            t1 += ido
            t2 = t1 << 1

        Next

        If (ido Mod 2 = 1) Then
            Return
        End If

L105:
        t1 = ido - 1
        t2 = ido - 1
        For k = 0 To l1 - 1 Step 1

            ch(t1) = cc(t2) + cc(t2)
            ch(t1 + t0) = -(cc(t2 + 1) + cc(t2 + 1))
            t1 += ido
            t2 += ido << 1
        Next
    End Sub
    Sub dradf2(ByVal ido As Long, ByVal l1 As Long, ByRef cc() As Double, ByRef ch() As Double, ByRef wa1() As Double, ByVal wa1Index As Long)

        'wa1Index - index at wa1

        Dim i, k As Long
        Dim ti2, tr2 As Double
        Dim t0, t1, t2, t3, t4, t5, t6 As Long

        t1 = 0
        t2 = l1 * ido
        t0 = t2
        t3 = ido << 1
        For k = 0 To l1 - 1 Step 1

            ch(t1 << 1) = cc(t1) + cc(t2)
            ch((t1 << 1) + t3 - 1) = cc(t1) - cc(t2)
            t1 += ido
            t2 += ido
        Next

        If (ido < 2) Then
            Return
        End If
        If (ido = 2) Then
            GoTo L105
        End If

        t1 = 0
        t2 = t0
        For k = 0 To l1 - 1 Step 1

            t3 = t2
            t4 = (t1 << 1) + (ido << 1)
            t5 = t1
            t6 = t1 + t1
            For i = 2 To ido - 1 Step 2

                t3 += 2
                t4 -= 2
                t5 += 2
                t6 += 2
                tr2 = wa1(wa1Index + i - 2) * cc(t3 - 1) + wa1(wa1Index + i - 1) * cc(t3)
                ti2 = wa1(wa1Index + i - 2) * cc(t3) - wa1(wa1Index + i - 1) * cc(t3 - 1)
                ch(t6) = cc(t5) + ti2
                ch(t4) = ti2 - cc(t5)
                ch(t6 - 1) = cc(t5 - 1) + tr2
                ch(t4 - 1) = cc(t5 - 1) - tr2
            Next
            t1 += ido
            t2 += ido
        Next

        If (ido Mod 2 = 1) Then
            Return
        End If

L105:
        t1 = ido
        t2 = t1 - 1
        t3 = t2
        t2 += t0
        For k = 0 To l1 - 1 Step 1

            ch(t1) = -cc(t2)
            ch(t1 - 1) = cc(t3)
            t1 += ido << 1
            t2 += ido
            t3 += ido
        Next
    End Sub
    Sub dradb3(ByVal ido As Long, ByVal l1 As Long, ByRef cc() As Double, ByRef ch() As Double, ByRef wa1() As Double, ByVal wa1Index As Long, ByVal wa2Index As Long)

        'wa1Index, wa2Index = indexes at wa1
        Dim taur As Double = -0.5
        Dim taui As Double = 0.8660254037844386
        Dim i, k, t0, t1, t2, t3, t4, t5, t6, t7, t8, t9, t10 As Long
        Dim ci2, ci3, di2, di3, cr2, cr3, dr2, dr3, ti2, tr2 As Double

        t0 = l1 * ido

        t1 = 0
        t2 = t0 << 1
        t3 = ido << 1
        t4 = ido + (ido << 1)
        t5 = 0
        For k = 0 To l1 - 1 Step 1

            tr2 = cc(t3 - 1) + cc(t3 - 1)
            cr2 = cc(t5) + (taur * tr2)
            ch(t1) = cc(t5) + tr2
            ci3 = taui * (cc(t3) + cc(t3))
            ch(t1 + t0) = cr2 - ci3
            ch(t1 + t2) = cr2 + ci3
            t1 += ido
            t3 += t4
            t5 += t4
        Next

        If (ido = 1) Then
            Return
        End If

        t1 = 0
        t3 = ido << 1
        For k = 0 To l1 - 1 Step 1

            t7 = t1 + (t1 << 1)
            t5 = t7 + t3
            t6 = t5
            t8 = t1
            t9 = t1 + t0
            t10 = t9 + t0

            For i = 2 To ido - 1 Step 2

                t5 += 2
                t6 -= 2
                t7 += 2
                t8 += 2
                t9 += 2
                t10 += 2
                tr2 = cc(t5 - 1) + cc(t6 - 1)
                cr2 = cc(t7 - 1) + (taur * tr2)
                ch(t8 - 1) = cc(t7 - 1) + tr2
                ti2 = cc(t5) - cc(t6)
                ci2 = cc(t7) + (taur * ti2)
                ch(t8) = cc(t7) + ti2
                cr3 = taui * (cc(t5 - 1) - cc(t6 - 1))
                ci3 = taui * (cc(t5) + cc(t6))
                dr2 = cr2 - ci3
                dr3 = cr2 + ci3
                di2 = ci2 + cr3
                di3 = ci2 - cr3
                ch(t9 - 1) = wa1(wa1Index + i - 2) * dr2 - wa1(wa1Index + i - 1) * di2
                ch(t9) = wa1(wa1Index + i - 2) * di2 + wa1(wa1Index + i - 1) * dr2
                ch(t10 - 1) = wa1(wa2Index + i - 2) * dr3 - wa1(wa2Index + i - 1) * di3
                ch(t10) = wa1(wa2Index + i - 2) * di3 + wa1(wa2Index + i - 1) * dr3
            Next
            t1 += ido
        Next
    End Sub
    Sub dradb4(ByVal ido As Long, ByVal l1 As Long, ByRef cc() As Double, ByRef ch() As Double, ByRef wa1() As Double, ByVal wa1Index As Long, ByVal wa2Index As Long, ByVal wa3Index As Long)
        'wa1Index,wa2Index,wa3Index - indexes at wa1
        Dim sqrt2 As Double = Math.Sqrt(2)
        Dim i, k, t0, t1, t2, t3, t4, t5, t6, t7, t8 As Long
        Dim ci2, ci3, ci4, cr2, cr3, cr4, ti1, ti2, ti3, ti4, tr1, tr2, tr3, tr4 As Double

        t0 = l1 * ido

        t1 = 0
        t2 = ido << 2
        t3 = 0
        t6 = ido << 1
        For k = 0 To l1 - 1 Step 1

            t4 = t3 + t6
            t5 = t1
            tr3 = cc(t4 - 1) + cc(t4 - 1)
            tr4 = cc(t4) + cc(t4)
            t4 += t6
            tr1 = cc(t3) - cc(t4 - 1)
            tr2 = cc(t3) + cc(t4 - 1)
            ch(t5) = tr2 + tr3
            t5 += t0
            ch(t5) = tr1 - tr4
            t5 += t0
            ch(t5) = tr2 - tr3
            t5 += t0
            ch(t5) = tr1 + tr4
            t1 += ido
            t3 += t2
        Next

        If (ido < 2) Then
            Return
        End If
        If (ido = 2) Then
            GoTo L105
        End If

        t1 = 0
        For k = 0 To l1 - 1 Step 1
            t2 = t1 << 2
            t3 = t2 + t6
            t4 = t3
            t5 = t4 + t6
            t7 = t1
            For i = 2 To ido - 1 Step 2
                t2 += 2
                t3 += 2
                t4 -= 2
                t5 -= 2
                t7 += 2
                ti1 = cc(t2) + cc(t5)
                ti2 = cc(t2) - cc(t5)
                ti3 = cc(t3) - cc(t4)
                tr4 = cc(t3) + cc(t4)
                tr1 = cc(t2 - 1) - cc(t5 - 1)
                tr2 = cc(t2 - 1) + cc(t5 - 1)
                ti4 = cc(t3 - 1) - cc(t4 - 1)
                tr3 = cc(t3 - 1) + cc(t4 - 1)
                ch(t7 - 1) = tr2 + tr3
                cr3 = tr2 - tr3
                ch(t7) = ti2 + ti3
                ci3 = ti2 - ti3
                cr2 = tr1 - tr4
                cr4 = tr1 + tr4
                ci2 = ti1 + ti4
                ci4 = ti1 - ti4
                t8 = t7 + t0
                ch(t8 - 1) = wa1(wa1Index + i - 2) * cr2 - wa1(wa1Index + i - 1) * ci2
                ch(t8) = wa1(wa1Index + i - 2) * ci2 + wa1(wa1Index + i - 1) * cr2
                t8 += t0
                ch(t8 - 1) = wa1(wa2Index + i - 2) * cr3 - wa1(wa2Index + i - 1) * ci3
                ch(t8) = wa1(wa2Index + i - 2) * ci3 + wa1(wa2Index + i - 1) * cr3
                t8 += t0
                ch(t8 - 1) = wa1(wa3Index + i - 2) * cr4 - wa1(wa3Index + i - 1) * ci4
                ch(t8) = wa1(wa3Index + i - 2) * ci4 + wa1(wa3Index + i - 1) * cr4
            Next
            t1 += ido
        Next

        If (ido Mod 2 = 1) Then
            Return
        End If

L105:

        t1 = ido
        t2 = ido << 2
        t3 = ido - 1
        t4 = ido + (ido << 1)
        For k = 0 To l1 - 1 Step 1

            t5 = t3
            ti1 = cc(t1) + cc(t4)
            ti2 = cc(t4) - cc(t1)
            tr1 = cc(t1 - 1) - cc(t4 - 1)
            tr2 = cc(t1 - 1) + cc(t4 - 1)
            ch(t5) = tr2 + tr2
            t5 += t0
            ch(t5) = sqrt2 * (tr1 - ti1)
            t5 += t0
            ch(t5) = ti2 + ti2
            t5 += t0
            ch(t5) = -sqrt2 * (tr1 + ti1)
            t3 += ido
            t1 += t2
            t4 += t2
        Next
    End Sub
    Sub dradf4(ByVal ido As Long, ByVal l1 As Long, ByRef cc() As Double, ByRef ch() As Double, ByRef wa1() As Double, ByVal wa1Index As Long, ByVal wa2Index As Long, ByVal wa3Index As Long)
        'wa1Index,wa2Index,wa3Index - indexes at wa1

        Dim hsqt2 As Double = Math.Sqrt(2) / 2
        Dim i, k, t0, t1, t2, t3, t4, t5, t6 As Long
        Dim ci2, ci3, ci4, cr2, cr3, cr4, ti1, ti2, ti3, ti4, tr1, tr2, tr3, tr4 As Double

        t0 = l1 * ido

        t1 = t0
        t4 = t1 << 1
        t2 = t1 + (t1 << 1)
        t3 = 0

        For k = 0 To l1 - 1 Step 1

            tr1 = cc(t1) + cc(t2)
            tr2 = cc(t3) + cc(t4)
            t5 = t3 << 2
            ch(t5) = tr1 + tr2
            ch((ido << 2) + t5 - 1) = tr2 - tr1
            t5 += (ido << 1)
            ch(t5 - 1) = cc(t3) - cc(t4)
            ch(t5) = cc(t2) - cc(t1)

            t1 += ido
            t2 += ido
            t3 += ido
            t4 += ido
        Next

        If (ido < 2) Then
            Return
        End If
        If (ido = 2) Then
            GoTo L105
        End If

        t1 = 0
        For k = 0 To l1 - 1 Step 1

            t2 = t1
            t4 = t1 << 2
            t6 = ido << 1
            t5 = t6 + t4
            For i = 2 To ido - 1 Step 2
                t2 += 2
                t3 = t2
                t4 += 2
                t5 -= 2

                t3 += t0
                cr2 = wa1(wa1Index + i - 2) * cc(t3 - 1) + wa1(wa1Index + i - 1) * cc(t3)
                ci2 = wa1(wa1Index + i - 2) * cc(t3) - wa1(wa1Index + i - 1) * cc(t3 - 1)
                t3 += t0
                cr3 = wa1(wa2Index + i - 2) * cc(t3 - 1) + wa1(wa2Index + i - 1) * cc(t3)
                ci3 = wa1(wa2Index + i - 2) * cc(t3) - wa1(wa2Index + i - 1) * cc(t3 - 1)
                t3 += t0
                cr4 = wa1(wa3Index + i - 2) * cc(t3 - 1) + wa1(wa3Index + i - 1) * cc(t3)
                ci4 = wa1(wa3Index + i - 2) * cc(t3) - wa1(wa3Index + i - 1) * cc(t3 - 1)

                tr1 = cr2 + cr4
                tr4 = cr4 - cr2
                ti1 = ci2 + ci4
                ti4 = ci2 - ci4
                ti2 = cc(t2) + ci3
                ti3 = cc(t2) - ci3
                tr2 = cc(t2 - 1) + cr3
                tr3 = cc(t2 - 1) - cr3

                ch(t4 - 1) = tr1 + tr2
                ch(t4) = ti1 + ti2

                ch(t5 - 1) = tr3 - ti4
                ch(t5) = tr4 - ti3

                ch(t4 + t6 - 1) = ti4 + tr3
                ch(t4 + t6) = tr4 + ti3

                ch(t5 + t6 - 1) = tr2 - tr1
                ch(t5 + t6) = ti1 - ti2
            Next
            t1 += ido
        Next
        If (ido Mod 2 = 1) Then
            Return
        End If

L105:
        t1 = t0 + ido - 1
        t2 = t1 + (t0 << 1)
        t3 = ido << 2
        t4 = ido
        t5 = ido << 1
        t6 = ido

        For k = 0 To l1 - 1 Step 1

            ti1 = -hsqt2 * (cc(t1) + cc(t2))
            tr1 = hsqt2 * (cc(t1) - cc(t2))
            ch(t4 - 1) = tr1 + cc(t6 - 1)
            ch(t4 + t5 - 1) = cc(t6 - 1) - tr1
            ch(t4) = ti1 - cc(t1 + t0)
            ch(t4 + t5) = ti1 + cc(t1 + t0)
            t1 += ido
            t2 += ido
            t4 += t3
            t6 += ido
        Next
    End Sub
    Sub dradbg(ByVal ido As Long, ByVal ip As Long, ByVal l1 As Long, ByVal idl1 As Long, ByRef cc() As Double, ByRef c1() As Double, ByRef c2() As Double, ByRef ch() As Double, ByRef ch2() As Double, ByRef wa() As Double, ByVal wa1Index As Long)

        Dim tpi As Double = 2 * Math.PI
        Dim idij, ipph, i, j, k, l, ik, is1, t0, t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12 As Long
        Dim dc2, ai1, ai2, ar1, ar2, ds2 As Double
        Dim nbd As Long
        Dim dcp, arg, dsp, ar1h, ar2h As Double
        Dim ipp2 As Long

        t10 = ip * ido
        t0 = l1 * ido
        arg = tpi / Convert.ToDouble(ip)
        dcp = Math.Cos(arg)
        dsp = Math.Sin(arg)
        nbd = (ido - 1) >> 1
        ipp2 = ip
        ipph = (ip + 1) >> 1
        If (ido < l1) Then
            GoTo L103
        End If

        t1 = 0
        t2 = 0
        For k = 0 To l1 - 1 Step 1
            t3 = t1
            t4 = t2
            For i = 0 To ido - 1 Step 1

                ch(t3) = cc(t4)
                t3 += 1
                t4 += 1
            Next
            t1 += ido
            t2 += t10
        Next
        GoTo L106

L103:
        t1 = 0
        For i = 0 To ido - 1 Step 1

            t2 = t1
            t3 = t1
            For k = 0 To l1 Step 1

                ch(t2) = cc(t3)
                t2 += ido
                t3 += t10
            Next
            t1 += 1
        Next
L106:
        t1 = 0
        t2 = ipp2 * t0
        t5 = ido << 1
        t7 = t5
        For j = 1 To ipph - 1 Step 1
            t1 += t0
            t2 -= t0
            t3 = t1
            t4 = t2
            t6 = t5
            For k = 0 To l1 - 1 Step 1
                ch(t3) = cc(t6 - 1) + cc(t6 - 1)
                ch(t4) = cc(t6) + cc(t6)
                t3 += ido
                t4 += ido
                t6 += t10
            Next
            t5 += t7
        Next

        If (ido = 1) Then
            GoTo L116
        End If
        If (nbd < l1) Then
            GoTo L112
        End If

        t1 = 0
        t2 = ipp2 * t0
        t7 = 0
        For j = 1 To ipph - 1 Step 1
            t1 += t0
            t2 -= t0
            t3 = t1
            t4 = t2

            t7 += (ido << 1)
            t8 = t7
            For k = 0 To l1 - 1 Step 1
                t5 = t3
                t6 = t4
                t9 = t8
                t11 = t8
                For i = 2 To ido - 1 Step 2
                    t5 += 2
                    t6 += 2
                    t9 += 2
                    t11 -= 2
                    ch(t5 - 1) = cc(t9 - 1) + cc(t11 - 1)
                    ch(t6 - 1) = cc(t9 - 1) - cc(t11 - 1)
                    ch(t5) = cc(t9) - cc(t11)
                    ch(t6) = cc(t9) + cc(t11)
                Next
                t3 += ido
                t4 += ido
                t8 += t10
            Next
        Next
        GoTo L116

L112:
        t1 = 0
        t2 = ipp2 * t0
        t7 = 0
        For j = 1 To ipph - 1 Step 1
            t1 += t0
            t2 -= t0
            t3 = t1
            t4 = t2
            t7 += (ido << 1)
            t8 = t7
            t9 = t7
            For i = 2 To ido - 1 Step 2
                t3 += 2
                t4 += 2
                t8 += 2
                t9 -= 2
                t5 = t3
                t6 = t4
                t11 = t8
                t12 = t9
                For k = 0 To l1 - 1 Step 1
                    ch(t5 - 1) = cc(t11 - 1) + cc(t12 - 1)
                    ch(t6 - 1) = cc(t11 - 1) - cc(t12 - 1)
                    ch(t5) = cc(t11) - cc(t12)
                    ch(t6) = cc(t11) + cc(t12)
                    t5 += ido
                    t6 += ido
                    t11 += t10
                    t12 += t10
                Next
            Next
        Next

L116:
        ar1 = 1.0
        ai1 = 0.0
        t1 = 0
        t2 = ipp2 * idl1
        t9 = t2
        t3 = (ip - 1) * idl1
        For l = 1 To ipph - 1 Step 1
            t1 += idl1
            t2 -= idl1

            ar1h = dcp * ar1 - dsp * ai1
            ai1 = dcp * ai1 + dsp * ar1
            ar1 = ar1h
            t4 = t1
            t5 = t2
            t6 = 0
            t7 = idl1
            t8 = t3
            For ik = 0 To idl1 - 1 Step 1
                c2(t4) = ch2(t6) + ar1 * ch2(t7)
                c2(t5) = ai1 * ch2(t8)
                t4 += 1
                t5 += 1
                t6 += 1
                t7 += 1
                t8 += 1
            Next
            dc2 = ar1
            ds2 = ai1
            ar2 = ar1
            ai2 = ai1

            t6 = idl1
            t7 = t9 - idl1
            For j = 2 To ipph - 1 Step 1

                t6 += idl1
                t7 -= idl1
                ar2h = dc2 * ar2 - ds2 * ai2
                ai2 = dc2 * ai2 + ds2 * ar2
                ar2 = ar2h
                t4 = t1
                t5 = t2
                t11 = t6
                t12 = t7
                For ik = 0 To idl1 - 1 Step 1
                    c2(t4) += ar2 * ch2(t11)
                    c2(t5) += ai2 * ch2(t12)
                    t4 += 1
                    t5 += 1
                    t11 += 1
                    t12 += 1
                Next
            Next
        Next

        t1 = 0
        For j = 1 To ipph - 1 Step 1

            t1 += idl1
            t2 = t1
            For ik = 0 To idl1 Step 1
                ch2(ik) += ch2(t2)
                t2 += 1
            Next
        Next

        t1 = 0
        t2 = ipp2 * t0
        For j = 1 To ipph - 1 Step 1

            t1 += t0
            t2 -= t0
            t3 = t1
            t4 = t2
            For k = 0 To l1 - 1 Step 1
                ch(t3) = c1(t3) - c1(t4)
                ch(t4) = c1(t3) + c1(t4)
                t3 += ido
                t4 += ido
            Next
        Next

        If (ido = 1) Then
            GoTo L132
        End If
        If (nbd < l1) Then
            GoTo L128
        End If

        t1 = 0
        t2 = ipp2 * t0
        For j = 1 To ipph - 1 Step 1
            t1 += t0
            t2 -= t0
            t3 = t1
            t4 = t2
            For k = 0 To l1 - 1 Step 1
                t5 = t3
                t6 = t4
                For i = 2 To ido - 1 Step 2
                    t5 += 2
                    t6 += 2
                    ch(t5 - 1) = c1(t5 - 1) - c1(t6)
                    ch(t6 - 1) = c1(t5 - 1) + c1(t6)
                    ch(t5) = c1(t5) + c1(t6 - 1)
                    ch(t6) = c1(t5) - c1(t6 - 1)
                Next
                t3 += ido
                t4 += ido
            Next
        Next
        GoTo L132

L128:
        t1 = 0
        t2 = ipp2 * t0
        For j = 1 To ipph - 1 Step 1
            t1 += t0
            t2 -= t0
            t3 = t1
            t4 = t2
            For i = 2 To ido - 1 Step 2
                t3 += 2
                t4 += 2
                t5 = t3
                t6 = t4
                For k = 0 To l1 - 1 Step 1
                    ch(t5 - 1) = c1(t5 - 1) - c1(t6)
                    ch(t6 - 1) = c1(t5 - 1) + c1(t6)
                    ch(t5) = c1(t5) + c1(t6 - 1)
                    ch(t6) = c1(t5) - c1(t6 - 1)
                    t5 += ido
                    t6 += ido
                Next
            Next
        Next

L132:
        If (ido = 1) Then
            Return
        End If

        For ik = 0 To idl1 - 1 Step 1
            c2(ik) = ch2(ik)
        Next

        t1 = 0
        For j = 1 To ip - 1 Step 1
            t2 = t1
            t1 += t0
            For k = 0 To l1 - 1 Step 1
                c1(t2) = ch(t2)
                t2 += ido
            Next
        Next

        If (nbd > l1) Then
            GoTo L139
        End If

        is1 = -ido - 1
        t1 = 0
        For j = 1 To ip - 1 Step 1

            is1 += ido
            t1 += t0
            idij = is1
            t2 = t1
            For i = 2 To ido - 1 Step 2

                t2 += 2
                idij += 2
                t3 = t2
                For k = 0 To l1 - 1 Step 1

                    c1(t3 - 1) = wa(wa1Index + idij - 1) * ch(t3 - 1) - wa(wa1Index + idij) * ch(t3)
                    c1(t3) = wa(wa1Index + idij - 1) * ch(t3) + wa(wa1Index + idij) * ch(t3 - 1)
                    t3 += ido
                Next
            Next
        Next
        Return

L139:
        is1 = -ido - 1
        t1 = 0
        For j = 1 To ip - 1 Step 1

            is1 += ido
            t1 += t0
            t2 = t1
            For k = 0 To l1 - 1 Step 1

                idij = is1
                t3 = t2
                For i = 2 To ido - 1 Step 2

                    idij += 2
                    t3 += 2
                    c1(t3 - 1) = wa(wa1Index + idij - 1) * ch(t3 - 1) - wa(wa1Index + idij) * ch(t3)
                    c1(t3) = wa(wa1Index + idij - 1) * ch(t3) + wa(wa1Index + idij) * ch(t3 - 1)
                Next
                t2 += ido
            Next
        Next
    End Sub
    Sub dradfg(ByVal ido As Long, ByVal ip As Long, ByVal l1 As Long, ByVal idl1 As Long, ByRef cc() As Double, ByRef c1() As Double, ByRef c2() As Double, ByRef ch() As Double, ByRef ch2() As Double, ByRef wa() As Double, ByVal wa1Index As Long)
        Dim tpi As Double = 2 * Math.PI
        Dim idij, ipph, i, j, k, l, ic, ik, is1 As Long
        Dim t0, t1, t2, t3, t4, t5, t6, t7, t8, t9, t10 As Long
        Dim dc2, ai1, ai2, ar1, ar2, ds2 As Double
        Dim nbd As Long
        Dim dcp, arg, dsp, ar1h, ar2h As Double
        Dim idp2, ipp2 As Long

        arg = tpi / Convert.ToDouble(ip)
        dcp = Math.Cos(arg)
        dsp = Math.Sin(arg)
        ipph = (ip + 1) >> 1
        ipp2 = ip
        idp2 = ido
        nbd = (ido - 1) >> 1
        t0 = l1 * ido
        t10 = ip * ido

        If (ido = 1) Then
            GoTo L119
        End If
        For ik = 0 To idl1 - 1 Step 1
            ch2(ik) = c2(ik)
        Next

        t1 = 0
        For j = 1 To ip - 1 Step 1

            t1 += t0
            t2 = t1
            For k = 0 To l1 - 1 Step 1
                ch(t2) = c1(t2)
                t2 += ido
            Next
        Next

        is1 = -ido
        t1 = 0
        If (nbd > l1) Then
            For j = 1 To ip - 1 Step 1

                t1 += t0
                is1 += ido
                t2 = -ido + t1
                For k = 0 To l1 - 1 Step 1

                    idij = is1 - 1
                    t2 += ido
                    t3 = t2
                    For i = 2 To ido - 1 Step 2

                        idij += 2
                        t3 += 2
                        ch(t3 - 1) = wa(wa1Index + idij - 1) * c1(t3 - 1) + wa(wa1Index + idij) * c1(t3)
                        ch(t3) = wa(wa1Index + idij - 1) * c1(t3) - wa(wa1Index + idij) * c1(t3 - 1)
                    Next
                Next
            Next

        Else

            For j = 1 To ip - 1 Step 1
                is1 += ido
                idij = is1 - 1
                t1 += t0
                t2 = t1
                For i = 2 To ido - 1 Step 2

                    idij += 2
                    t2 += 2
                    t3 = t2
                    For k = 0 To l1 - 1 Step 1
                        ch(t3 - 1) = wa(wa1Index + idij - 1) * c1(t3 - 1) + wa(wa1Index + idij) * c1(t3)
                        ch(t3) = wa(wa1Index + idij - 1) * c1(t3) - wa(wa1Index + idij) * c1(t3 - 1)
                        t3 += ido
                    Next
                Next
            Next
        End If

        t1 = 0
        t2 = ipp2 * t0
        If (nbd < l1) Then
            For j = 1 To ipph - 1 Step 1
                t1 += t0
                t2 -= t0
                t3 = t1
                t4 = t2
                For i = 2 To ido - 1 Step 2
                    t3 += 2
                    t4 += 2
                    t5 = t3 - ido
                    t6 = t4 - ido
                    For k = 0 To l1 - 1 Step 1
                        t5 += ido
                        t6 += ido
                        c1(t5 - 1) = ch(t5 - 1) + ch(t6 - 1)
                        c1(t6 - 1) = ch(t5) - ch(t6)
                        c1(t5) = ch(t5) + ch(t6)
                        c1(t6) = ch(t6 - 1) - ch(t5 - 1)
                    Next
                Next
            Next
        Else
            For j = 1 To ipph - 1 Step 1

                t1 += t0
                t2 -= t0
                t3 = t1
                t4 = t2
                For k = 0 To l1 - 1 Step 1
                    t5 = t3
                    t6 = t4
                    For i = 2 To ido - 1 Step 2
                        t5 += 2
                        t6 += 2
                        c1(t5 - 1) = ch(t5 - 1) + ch(t6 - 1)
                        c1(t6 - 1) = ch(t5) - ch(t6)
                        c1(t5) = ch(t5) + ch(t6)
                        c1(t6) = ch(t6 - 1) - ch(t5 - 1)
                    Next
                    t3 += ido
                    t4 += ido
                Next
            Next
        End If

L119:
        For ik = 0 To idl1 - 1 Step 1
            c2(ik) = ch2(ik)
        Next

        t1 = 0
        t2 = ipp2 * idl1
        For j = 1 To ipph - 1 Step 1
            t1 += t0
            t2 -= t0
            t3 = t1 - ido
            t4 = t2 - ido
            For k = 0 To l1 - 1 Step 1
                t3 += ido
                t4 += ido
                c1(t3) = ch(t3) + ch(t4)
                c1(t4) = ch(t4) - ch(t3)
            Next
        Next

        ar1 = 1.0
        ai1 = 0.0
        t1 = 0
        t2 = ipp2 * idl1
        t3 = (ip - 1) * idl1
        For l = 1 To ipph - 1 Step 1
            t1 += idl1
            t2 -= idl1
            ar1h = dcp * ar1 - dsp * ai1
            ai1 = dcp * ai1 + dsp * ar1
            ar1 = ar1h
            t4 = t1
            t5 = t2
            t6 = t3
            t7 = idl1

            For ik = 0 To idl1 - 1 Step 1

                ch2(t4) = c2(ik) + ar1 * c2(t7)
                t4 += 1
                t7 += 1
                ch2(t5) = ai1 * c2(t6)
                t5 += 1
                t6 += 1
            Next

            dc2 = ar1
            ds2 = ai1
            ar2 = ar1
            ai2 = ai1

            t4 = idl1
            t5 = (ipp2 - 1) * idl1
            For j = 2 To ipph - 1 Step 1
                t4 += idl1
                t5 -= idl1

                ar2h = dc2 * ar2 - ds2 * ai2
                ai2 = dc2 * ai2 + ds2 * ar2
                ar2 = ar2h

                t6 = t1
                t7 = t2
                t8 = t4
                t9 = t5
                For ik = 0 To idl1 - 1 Step 1

                    ch2(t6) += ar2 * c2(t8)
                    t6 += 1
                    t8 += 1
                    ch2(t7) += ai2 * c2(t9)
                    t7 += 1
                    t9 += 1
                    t6 += 1
                    t7 += 1
                    t8 += 1
                    t9 += 1
                Next
            Next
        Next

        t1 = 0
        For j = 1 To ipph - 1 Step 1
            t1 += idl1
            t2 = t1
            For ik = 0 To idl1 - 1 Step 1
                ch2(ik) += c2(t2)
                t2 += 1
            Next
        Next

        If (ido < l1) Then
            GoTo L132
        End If

        t1 = 0
        t2 = 0
        For k = 0 To l1 - 1 Step 1
            t3 = t1
            t4 = t2
            For i = 0 To ido - 1 Step 1
                cc(t4) = ch(t3)
                t4 += 1
                t3 += 1
            Next
            t1 += ido
            t2 += t10
        Next

        GoTo L135

L132:
        For i = 0 To ido - 1 Step 1
            t1 = i
            t2 = i
            For k = 0 To l1 - 1 Step 1

                cc(t2) = ch(t1)
                t1 += ido
                t2 += t10
            Next
        Next

L135:
        t1 = 0
        t2 = ido << 1
        t3 = 0
        t4 = ipp2 * t0
        For j = 1 To ipph - 1 Step 1

            t1 += t2
            t3 += t0
            t4 -= t0

            t5 = t1
            t6 = t3
            t7 = t4

            For k = 0 To l1 - 1 Step 1
                cc(t5 - 1) = ch(t6)
                cc(t5) = ch(t7)
                t5 += t10
                t6 += ido
                t7 += ido
            Next
        Next

        If (ido = 1) Then
            Return
        End If
        If (nbd < l1) Then
            GoTo L141
        End If

        t1 = -ido
        t3 = 0
        t4 = 0
        t5 = ipp2 * t0
        For j = 1 To ipph - 1 Step 1
            t1 += t2
            t3 += t2
            t4 += t0
            t5 -= t0
            t6 = t1
            t7 = t3
            t8 = t4
            t9 = t5
            For k = 0 To l1 - 1 Step 1
                For i = 2 To ido - 1 Step 2
                    ic = idp2 - i
                    cc(i + t7 - 1) = ch(i + t8 - 1) + ch(i + t9 - 1)
                    cc(ic + t6 - 1) = ch(i + t8 - 1) - ch(i + t9 - 1)
                    cc(i + t7) = ch(i + t8) + ch(i + t9)
                    cc(ic + t6) = ch(i + t9) - ch(i + t8)
                Next
                t6 += t10
                t7 += t10
                t8 += ido
                t9 += ido
            Next
        Next
        Return

L141:

        t1 = -ido
        t3 = 0
        t4 = 0
        t5 = ipp2 * t0
        For j = 1 To ipph - 1 Step 1
            t1 += t2
            t3 += t2
            t4 += t0
            t5 -= t0
            For i = 2 To ido - 1 Step 2
                t6 = idp2 + t1 - i
                t7 = i + t3
                t8 = i + t4
                t9 = i + t5
                For k = 0 To l1 - 1 Step 1
                    cc(t7 - 1) = ch(t8 - 1) + ch(t9 - 1)
                    cc(t6 - 1) = ch(t8 - 1) - ch(t9 - 1)
                    cc(t7) = ch(t8) + ch(t9)
                    cc(t6) = ch(t9) - ch(t8)
                    t6 += t10
                    t7 += t10
                    t8 += ido
                    t9 += ido
                Next
            Next
        Next
    End Sub

    Sub drftb1(ByVal n As Long, ByRef c() As Double, ByRef ch() As Double, ByVal waIndex As Long, ByRef ifac() As Long)
        'wa = ch + waIndex
        'wa + ix2 - 1 = ch + waIndex + ix2 - 1
        'wa + iw - 1 = ch(waIndex + iw - 1)
        Dim i, k1, l1, l2 As Long
        Dim na As Long
        Dim nf, ip, iw, ix2, ix3, ido, idl1 As Long

        nf = ifac(1)
        na = 0
        l1 = 1
        iw = 1

        For k1 = 0 To nf - 1 Step 1
            ip = ifac(k1 + 2)
            l2 = ip * l1
            ido = n / l2
            idl1 = ido * l1
            If (ip <> 4) Then
                GoTo L103
            End If
            ix2 = iw + ido
            ix3 = ix2 + ido

            If (na <> 0) Then
                dradb4(ido, l1, ch, c, ch, waIndex + iw - 1, waIndex + ix2 - 1, waIndex + ix3 - 1)
            Else
                dradb4(ido, l1, c, ch, ch, waIndex + iw - 1, waIndex + ix2 - 1, waIndex + ix3 - 1)
            End If
            na = 1 - na
            GoTo L115

L103:
            If (ip <> 2) Then
                GoTo L106
            End If

            If (na <> 0) Then
                dradb2(ido, l1, ch, c, ch, waIndex + iw - 1)
            Else
                dradb2(ido, l1, c, ch, ch, waIndex + iw - 1)
            End If
            na = 1 - na
            GoTo L115

L106:
            If (ip <> 3) Then
                GoTo L109
            End If

            ix2 = iw + ido
            If (na <> 0) Then
                dradb3(ido, l1, ch, c, ch, waIndex + iw - 1, waIndex + ix2 - 1)
            Else
                dradb3(ido, l1, c, ch, ch, waIndex + iw - 1, waIndex + ix2 - 1)
            End If
            na = 1 - na
            GoTo L115

L109:
            '/* The radix five case can be translated later..... */
            '/* if(ip<>5)goto L112

            '  ix2=iw+ido ix3=ix2+ido ix4=ix3+ido if(na<>0)
            '  dradb5(ido,l1,ch,c,wa+iw-1,wa+ix2-1,wa+ix3-1,wa+ix4-1) else
            '  dradb5(ido,l1,c,ch,wa+iw-1,wa+ix2-1,wa+ix3-1,wa+ix4-1) na=1-na
            '   goto L115

            '   L112: */
            If (na <> 0) Then
                dradbg(ido, ip, l1, idl1, ch, ch, ch, c, c, ch, waIndex + iw - 1)
            Else
                dradbg(ido, ip, l1, idl1, c, c, c, ch, ch, ch, waIndex + iw - 1)
            End If
            If (ido = 1) Then
                na = 1 - na
            End If

L115:
            l1 = l2
            iw += (ip - 1) * ido
        Next

        If (na = 0) Then
            Return
        End If

        For i = 0 To n - 1 Step 1
            c(i) = ch(i)
        Next
    End Sub
    Sub drftf1(ByVal n As Long, ByRef c() As Double, ByRef ch() As Double, ByVal waIndex As Long, ByRef ifac() As Long)
        Dim i, k1, l1, l2 As Long
        Dim na, kh, nf As Long
        Dim ip, iw, ido, idl1, ix2, ix3 As Long

        nf = ifac(1)
        na = 1
        l2 = n
        iw = n

        For k1 = 0 To nf - 1 Step 1

            kh = nf - k1
            ip = ifac(kh + 1)
            l1 = l2 / ip
            ido = n / l2
            idl1 = ido * l1
            iw -= (ip - 1) * ido
            na = 1 - na

            If (ip <> 4) Then
                GoTo L102
            End If

            ix2 = iw + ido
            ix3 = ix2 + ido
            If (na <> 0) Then
                dradf4(ido, l1, ch, c, ch, waIndex + iw - 1, waIndex + ix2 - 1, waIndex + ix3 - 1)
            Else
                dradf4(ido, l1, c, ch, ch, waIndex + iw - 1, waIndex + ix2 - 1, waIndex + ix3 - 1)
            End If
            GoTo L110

L102:
            If (ip <> 2) Then
                GoTo L104
            End If
            If (na <> 0) Then
                GoTo L103
            End If

            dradf2(ido, l1, c, ch, ch, waIndex + iw - 1)
            GoTo L110

L103:
            dradf2(ido, l1, ch, c, ch, waIndex + iw - 1)
            GoTo L110

L104:
            If (ido = 1) Then
                na = 1 - na
            End If
            If (na <> 0) Then
                GoTo L109
            End If

            dradfg(ido, ip, l1, idl1, c, c, c, ch, ch, ch, waIndex + iw - 1)
            na = 1
            GoTo L110

L109:
            dradfg(ido, ip, l1, idl1, ch, ch, ch, c, c, ch, waIndex + iw - 1)
            na = 0

L110:
            l2 = l1
        Next

        If (na = 1) Then
            Return
        End If

        For i = 0 To n - 1 Step 1
            c(i) = ch(i)
        Next
    End Sub
    Sub NUMfft_backward(ByRef mme As NumFFTTable, ByRef data() As Double)
        If (mme.n = 1) Then
            Return
        End If
        drftb1(mme.n, data, mme.trigcache, mme.n, mme.splitcache)
    End Sub
    Sub NUMfft_forward(ByRef mme As NumFFTTable, ByRef data() As Double)
        If (mme.n = 1) Then
            Return
        End If
        drftf1(mme.n, data, mme.trigcache, mme.n, mme.splitcache)
    End Sub
    Function Sound_to_Pitch_any(ByRef mme As Sound,
    ByVal dt As Double, ByVal minimumPitch As Double, ByVal periodsPerWindow As Double, ByVal maxnCandidates As Integer,
    ByVal method As Integer,
    ByVal silenceThreshold As Double, ByVal voicingThreshold As Double,
    ByVal octaveCost As Double, ByRef octaveJumpCost As Double, ByRef voicedUnvoicedCost As Double, ByVal ceiling As Double) As Pitch
        Try
            Dim fftTable As NumFFTTable = New NumFFTTable(0)
            Dim duration, t1 As Double
            '/* Window length in seconds. */
            Dim dt_window As Double
            '/* Number of samples per window. */
            Dim nsamp_window, halfnsamp_window As Long
            Dim nFrames, minimumLag, maximumLag As Long
            Dim iframe, nsampFFT As Long
            Dim interpolation_depth As Double
            '/* Number of samples in longest period. */
            Dim nsamp_period, halfnsamp_period As Long
            Dim brent_ixmax, brent_depth As Long
            '/* Obsolete. */
            Dim brent_accuracy As Double
            Dim globalPeak As Double
            'nFrames = 0

            '!!!Melder_assert (maxnCandidates >= 2)
            '!!!Melder_assert (method >= AC_HANNING && method <= FCC_ACCURATE)

            If (maxnCandidates < ceiling / minimumPitch) Then
                maxnCandidates = ceiling / minimumPitch
            End If
            '   /* e.g. 3 periods, 75 Hz: 10 milliseconds. */
            If (dt <= 0.0) Then
                dt = periodsPerWindow / minimumPitch / 4.0
            End If

            Select Case method
                Case AC_HANNING
                    brent_depth = NUM_PEAK_INTERPOLATE_SINC70
                    brent_accuracy = 0.0000001
                    interpolation_depth = 0.5
                    GoTo enofselect
                Case AC_GAUSS
                    '   /* Because Gaussian window is twice as long. */
                    periodsPerWindow *= 2
                    brent_depth = NUM_PEAK_INTERPOLATE_SINC700
                    brent_accuracy = 0.00000000001
                    '   /* Because Gaussian window is twice as long. */
                    interpolation_depth = 0.25
                    GoTo enofselect
                Case FCC_NORMAL
                    brent_depth = NUM_PEAK_INTERPOLATE_SINC70
                    brent_accuracy = 0.0000001
                    interpolation_depth = 1.0
                    GoTo enofselect
                Case FCC_ACCURATE
                    brent_depth = NUM_PEAK_INTERPOLATE_SINC700
                    brent_accuracy = 0.00000000001
                    interpolation_depth = 1.0
                    GoTo enofselect
            End Select
enofselect: duration = mme.dx * mme.nx
            If (minimumPitch < periodsPerWindow / duration) Then
                'Melder_throw ("To analyse this Sound, ", L_LEFT_DOUBLE_QUOTE, "minimum pitch", L_RIGHT_DOUBLE_QUOTE, " must not be less than ", periodsPerWindow / duration, " Hz.")
                Console.WriteLine("To analyse this Sound minimum pitch must not be less than" + periodsPerWindow / duration + " Hz.")
                Return Nothing
            End If

            '/*
            ' * Determine the number of samples in the longest period.
            ' * We need this to compute the local mean of the sound (looking one period in both directions),
            ' * and to compute the local peak of the sound (looking half a period in both directions).
            ' */
            nsamp_period = Math.Floor(1 / mme.dx / minimumPitch)
            halfnsamp_period = nsamp_period / 2 + 1

            If (ceiling > 0.5 / mme.dx) Then
                ceiling = 0.5 / mme.dx
            End If

            '/*
            ' * Determine window length in seconds and in samples.
            ' */
            dt_window = periodsPerWindow / minimumPitch
            nsamp_window = Math.Floor(dt_window / mme.dx)
            halfnsamp_window = nsamp_window / 2 - 1
            If (halfnsamp_window < 2) Then
                'Melder_throw("Analysis window too short.")
                Console.WriteLine("Analysis window too short.")
                Return Nothing
            End If
            nsamp_window = halfnsamp_window * 2

            '/*
            ' * Determine the minimum and maximum lags.
            ' */
            minimumLag = Math.Floor(1 / mme.dx / ceiling)
            If (minimumLag < 2) Then
                minimumLag = 2
            End If
            maximumLag = Math.Floor(nsamp_window / periodsPerWindow) + 2
            If (maximumLag > nsamp_window) Then
                maximumLag = nsamp_window
            End If

            '/*
            ' * Determine the number of frames.
            ' * Fit as many frames as possible symmetrically in the total duration.
            ' * We do this even for the forward cross-correlation method,
            ' * because that allows us to compare the two methods.
            ' */
            Try
                Dim v As Double
                If method >= FCC_NORMAL Then
                    v = 1 / minimumPitch + dt_window
                Else
                    v = dt_window
                End If
                Sampled_shortTermAnalysis(mme, v, dt, nFrames, t1)
            Catch ex As Exception
                'Melder_throw ("The pitch analysis would give zero pitch frames.")
            End Try

            '/*
            ' * Create the resulting pitch contour.
            ' */
            Dim thee As Pitch = New Pitch(mme.xmin, mme.xmax, nFrames, dt, t1, ceiling, maxnCandidates)

            '/*
            ' * Compute the global absolute peak for determination of silence threshold.
            ' */
            globalPeak = 0.0
            For channel As Long = 0 To mme.ny - 1 Step 1
                Dim mean As Double = 0.0
                For i As Long = 0 To mme.nx - 1 Step 1
                    mean += mme.z(channel, i)
                Next
                mean /= mme.nx
                For i As Long = 0 To mme.nx - 1 Step 1
                    Dim value As Double = Math.Abs(mme.z(channel, i) - mean)
                    If (value > globalPeak) Then
                        globalPeak = value
                    End If
                Next
            Next
            If (globalPeak = 0.0) Then
                Return thee
            End If

            Dim frame(,) As Double
            Dim ac(), window(), windowR() As Double
            ac = New Double() {}
            window = New Double() {}
            windowR = New Double() {}
            Dim secondRes As Integer
            If (method >= FCC_NORMAL) Then
                '/* For cross-correlation analysis. */
                '/*
                '* Create buffer for cross-correlation analysis.
                '*/
                frame = New Double(mme.ny, nsamp_window) {}
                'frame.reset(1, mme.ny, 1, nsamp_window)
                secondRes = nsamp_window
                brent_ixmax = nsamp_window * interpolation_depth

            Else
                '/* For autocorrelation analysis. */

                '/*
                '* Compute the number of samples needed for doing FFT.
                '* To avoid edge effects, we have to append zeroes to the window.
                '* The maximum lag considered for maxima is maximumLag.
                '* The maximum lag used in interpolation is nsamp_window * interpolation_depth.
                '*/
                nsampFFT = 1
                While (nsampFFT < nsamp_window * (1 + interpolation_depth))
                    nsampFFT *= 2
                End While

                '/*
                '* Create buffers for autocorrelation analysis.
                '*/
                frame = New Double(mme.ny - 1, nsampFFT - 1) {}
                secondRes = nsampFFT
                ReDim windowR(nsampFFT - 1)
                ReDim window(nsamp_window - 1)
                ReDim ac(nsampFFT - 1)
                fftTable = New NumFFTTable(nsampFFT)


                '/*
                '* A Gaussian or Hanning window is applied against phase effects.
                '* The Hanning window is 2 to 5 dB better for 3 periods/window.
                '* The Gaussian window is 25 to 29 dB better for 6 periods/window.
                '*/
                If (method = AC_GAUSS) Then
                    '/* Gaussian window. */
                    Dim imid As Double = 0.5 * (nsamp_window + 1), edge = exp(-12.0)
                    For i As Long = 0 To nsamp_window - 1 Step 1
                        window(i) = (exp(-48.0 * (i - imid) * (i - imid) / (nsamp_window + 1) / (nsamp_window + 1)) - edge) / (1 - edge)
                    Next
                Else
                    '// Hanning window
                    For i As Long = 1 To nsamp_window Step 1
                        window(i - 1) = 0.5 - 0.5 * Math.Cos(i * 2 * Math.PI / (nsamp_window + 1))
                    Next
                End If

                '/*
                '* Compute the normalized autocorrelation of the window.
                '*/
                For i As Long = 0 To nsamp_window - 1 Step 1
                    windowR(i) = window(i)
                Next
                NUMfft_forward(fftTable, windowR)
                '// DC component
                windowR(0) *= windowR(0)
                For i As Long = 1 To nsampFFT - 2 Step 2
                    windowR(i) = windowR(i) * windowR(i) + windowR(i + 1) * windowR(i + 1)
                    '// power spectrum: square and zero
                    windowR(i + 1) = 0.0
                Next
                '   // Nyquist frequency
                windowR(nsampFFT - 1) *= windowR(nsampFFT - 1)
                '// autocorrelation
                NUMfft_backward(fftTable, windowR)
                '// normalize
                For i As Long = 1 To nsamp_window - 1 Step 1
                    windowR(i) /= windowR(0)
                Next

                '   // normalize
                windowR(0) = 1.0

                brent_ixmax = nsamp_window * interpolation_depth
            End If

            'autoNUMvector <double> r (- nsamp_window, nsamp_window)
            Dim r(2 * nsamp_window) As Double
            Dim imax(maxnCandidates - 1) As Long
            'autoNUMvector <double> localMean (1, mme.ny)
            Dim localMean(mme.ny - 1) As Double

            'autoMelderProgress progress (L"Sound to Pitch...")

            For iframe = 0 To nFrames - 1 Step 1
                Dim pitchFrame As Pitch_Frame = thee.frame(iframe)
                Dim t As Double = Sampled_indexToX(thee, iframe), localPeak
                Dim leftSample As Long = Sampled_xToLowIndex(mme, t), rightSample = leftSample + 1
                Dim startSample, endSample As Long
                'Melder_progress (0.1 + (0.8 * iframe) / (nFrames + 1),
                '	L"Sound to Pitch: analysis of frame ", Melder_integer (iframe), L" out of ", Melder_integer (nFrames))

                For channel As Long = 0 To mme.ny - 1 Step 1
                    '/*
                    ' * Compute the local mean look one longest period to both sides.
                    ' */
                    startSample = rightSample - nsamp_period
                    endSample = leftSample + nsamp_period
                    '!!!Melder_assert(startSample >= 1)
                    '!!!Melder_assert(endSample <= mme.nx)
                    localMean(channel) = 0.0
                    For i As Long = startSample - 1 To endSample - 1 Step 1
                        localMean(channel) += mme.z(channel, i)
                    Next
                    localMean(channel) /= 2 * nsamp_period

                    '/*
                    ' * Copy a window to a frame and subtract the local mean.
                    ' * We are going to kill the DC component before windowing.
                    ' */
                    startSample = rightSample - halfnsamp_window
                    endSample = leftSample + halfnsamp_window
                    'Melder_assert(startSample >= 1)
                    'Melder_assert(endSample <= mme.nx)
                    If (method < FCC_NORMAL) Then
                        Dim ii As Long = startSample - 1
                        For j As Long = 0 To nsamp_window - 1 Step 1
                            frame(channel, j) = (mme.z(channel, ii) - localMean(channel)) * window(j)
                            ii += 1
                        Next
                        For j As Long = nsamp_window To nsampFFT - 1 Step 1
                            frame(channel, j) = 0.0
                        Next
                    Else
                        Dim ii As Long = startSample - 1
                        For j As Long = 0 To nsamp_window - 1 Step 1
                            frame(channel, j) = mme.z(channel, ii) - localMean(channel)
                            ii += 1
                        Next
                    End If
                Next

                '/*
                ' * Compute the local peak look half a longest period to both sides.
                ' */
                localPeak = 0.0
                startSample = halfnsamp_window + 1 - halfnsamp_period
                If (startSample < 1) Then
                    startSample = 0
                End If
                endSample = halfnsamp_window + halfnsamp_period
                If (endSample > nsamp_window) Then
                    endSample = nsamp_window - 1
                End If
                For channel As Long = 0 To mme.ny - 1 Step 1
                    For j As Long = startSample - 1 To endSample - 1 Step 1
                        Dim value As Double = Math.Abs(frame(channel, j))
                        If (value > localPeak) Then
                            localPeak = value
                        End If
                    Next
                Next
                If localPeak > globalPeak Then
                    pitchFrame.intensity = 1.0
                Else
                    pitchFrame.intensity = localPeak / globalPeak
                End If
                '/*
                ' * Compute the correlation into the array 'r'.
                ' */
                If (method >= FCC_NORMAL) Then
                    Dim startTime As Double = t - 0.5 * (1.0 / minimumPitch + dt_window)
                    Dim localSpan As Long = maximumLag + nsamp_window, localMaximumLag, offset
                    If ((startSample = Sampled_xToLowIndex(mme, startTime)) < 1) Then
                        startSample = 1
                    End If
                    If (localSpan > mme.nx + 1 - startSample) Then
                        localSpan = mme.nx + 1 - startSample
                    End If
                    localMaximumLag = localSpan - nsamp_window
                    offset = startSample - 1
                    Dim sumx2 As Double = 0
                    '/* Sum of squares. */
                    For channel As Long = 0 To mme.ny - 1 Step 1
                        'double *amp = my z [channel] + offset
                        Dim ampIndex As Double = channel + offset
                        For i As Long = 0 To nsamp_window Step 1
                            Dim x As Double = mme.z(ampIndex + i, 0) - localMean(channel)
                            sumx2 += x * x
                        Next
                    Next
                    '/* At zero lag, these are still equal. */
                    Dim sumy2 As Double = sumx2
                    r(nsamp_window) = 1.0
                    For i As Long = 1 To localMaximumLag Step 1
                        Dim product As Double = 0.0
                        For channel As Long = 0 To mme.ny - 1 Step 1
                            Dim ampIndex As Long = channel + offset
                            Dim y0 As Double = mme.z(ampIndex + i, 0) - localMean(channel)
                            Dim yZ As Double = mme.z(ampIndex + i + nsamp_window, 0) - localMean(channel)
                            sumy2 += yZ * yZ - y0 * y0
                            For j As Long = 1 To nsamp_window Step 1
                                Dim x As Double = mme.z(ampIndex + j, 0) - localMean(channel)
                                Dim y As Double = mme.z(ampIndex + i + j, 0) - localMean(channel)
                                product += x * y
                            Next
                        Next
                        r(nsamp_window + i) = product / Math.Sqrt(sumx2 * sumy2)
                        r(nsamp_window - i) = r(nsamp_window + i)
                    Next
                Else

                    '/*
                    ' * The FFT of the autocorrelation is the power spectrum.
                    ' */
                    For i As Long = 0 To nsampFFT - 1 Step 1
                        ac(i) = 0.0
                    Next
                    For channel As Long = 0 To mme.ny - 1 Step 1
                        '/* Complex spectrum. */

                        Dim chd() As Double = New Double(secondRes - 1) {}
                        Dim ttt As Integer = 0
                        While ttt < secondRes
                            chd(ttt) = frame(channel, ttt)
                            ttt += 1
                        End While

                        NUMfft_forward(fftTable, chd)
                        ttt = 0
                        While ttt < secondRes
                            frame(channel, ttt) = chd(ttt)
                            ttt += 1
                        End While


                        ' /* DC component. */
                        ac(0) += frame(channel, 0) * frame(channel, 0)
                        For i As Long = 1 To nsampFFT - 2 Step 2
                            '/* Power spectrum. */
                            ac(i) += frame(channel, i) * frame(channel, i) + frame(channel, i + 1) * frame(channel, i + 1)
                        Next
                        ' /* Nyquist frequency. */
                        ac(nsampFFT - 1) += frame(channel, nsampFFT - 1) * frame(channel, nsampFFT - 1)
                    Next
                    '/* Autocorrelation. */
                    NUMfft_backward(fftTable, ac)

                    '/*
                    ' * Normalize the autocorrelation to the value with zero lag,
                    ' * and divide it by the normalized autocorrelation of the window.
                    ' */
                    r(nsamp_window) = 1.0
                    For i As Long = 0 To brent_ixmax - 1 Step 1
                        r(nsamp_window + i + 1) = ac(i + 1) / (ac(0) * windowR(i + 1))
                        r(nsamp_window - i - 1) = r(nsamp_window + i + 1)
                    Next
                End If
                '/*
                ' * Create (too much) space for candidates.
                ' */
                pitchFrame.reinit(maxnCandidates)

                '/*
                ' * Register the first candidate, which is always present: voicelessness.
                ' */
                pitchFrame.nCandidates = 1
                pitchFrame.candidates(0).frequency = 0.0
                '/* Voiceless: always present. */
                pitchFrame.candidates(0).strength = 0.0

                '/*
                ' * Shortcut: absolute silence is always voiceless.
                ' * Go to next frame.
                ' */
                If (localPeak = 0) Then
                    Continue For
                End If

                '/*
                ' * Find the strongest maxima of the correlation of this frame, 
                ' * and register them as candidates.
                ' */
                imax(0) = 0
                Dim mini As Long = Math.Min(maximumLag - 1, brent_ixmax - 1)
                For i As Long = 2 To mini Step 1
                    '/* Not too unvoiced? */
                    ' /* Maximum? */
                    If i = 147 Then
                        Dim hhhh As Long = 0
                    End If
                    If (r(nsamp_window + i) > 0.5 * voicingThreshold And r(nsamp_window + i) > r(nsamp_window + i - 1) And r(nsamp_window + i) >= r(nsamp_window + i + 1)) Then
                        Dim place As Integer = 0

                        '/*
                        ' * Use parabolic interpolation for first estimate of frequency,
                        ' * and sin(x)/x interpolation to compute the strength of this frequency.
                        ' */
                        Dim dr As Double = 0.5 * (r(nsamp_window + i + 1) - r(nsamp_window + i - 1)), d2r = 2 * r(nsamp_window + i) - r(nsamp_window + i - 1) - r(nsamp_window + i + 1)
                        Dim frequencyOfMaximum As Double = 1 / mme.dx / (i + dr / d2r)
                        Dim offset As Long = -brent_ixmax - 1
                        '/* method & 1 ? */
                        Dim strengthOfMaximum As Double = NUM_interpolate_sinc(r, nsamp_window + offset, brent_ixmax - offset, 1 / mme.dx / frequencyOfMaximum - offset, 30)
                        '/* : r [i] + 0.5 * dr * dr / d2r */
                        '	/* High values due to short windows are to be reflected around 1. */
                        If (strengthOfMaximum > 1.0) Then
                            strengthOfMaximum = 1.0 / strengthOfMaximum
                        End If

                        '/*
                        ' * Find a place for this maximum.
                        ' */
                        If (pitchFrame.nCandidates < thee.maxCandidates) Then
                            '/* Is there still a free place? */
                            place = pitchFrame.nCandidates
                            pitchFrame.nCandidates += 1
                        Else
                            '/* Try the place of the weakest candidate so far. */
                            Dim weakest As Double = 1
                            For iweak As Integer = 1 To thee.maxCandidates - 1 Step 1
                                '/* High frequencies are to be favoured */
                                '/* if we want to analyze a perfectly periodic signal correctly. */
                                Dim localStrength As Double = pitchFrame.candidates(iweak).strength - octaveCost * Log2X(minimumPitch / pitchFrame.candidates(iweak).frequency)
                                If (localStrength < weakest) Then
                                    weakest = localStrength
                                    place = iweak
                                End If
                            Next
                            '/* If this maximum is weaker than the weakest candidate so far, give it no place. */
                            If (strengthOfMaximum - octaveCost * Log2X(minimumPitch / frequencyOfMaximum) <= weakest) Then
                                place = 0
                            End If
                        End If
                        If (place <> 0) Then
                            '/* Have we found a place for this candidate? */
                            pitchFrame.candidates(place).frequency = frequencyOfMaximum
                            pitchFrame.candidates(place).strength = strengthOfMaximum
                            imax(place) = i

                        End If
                    End If
                Next

                '/*
                ' * Second pass: for extra precision, maximize sin(x)/x interpolation ('sinc').
                ' */
                For i As Long = 1 To pitchFrame.nCandidates - 1 Step 1
                    If (method <> AC_HANNING Or pitchFrame.candidates(i).frequency > 0.0 / mme.dx) Then
                        Dim xmid, ymid As Double
                        Dim offset As Long = -brent_ixmax - 1
                        Dim vv As Double
                        If pitchFrame.candidates(i).frequency > 0.3 / mme.dx Then
                            vv = NUM_PEAK_INTERPOLATE_SINC700
                        Else
                            vv = brent_depth
                        End If
                        ymid = NUMimproveMaximum(r, nsamp_window + offset, brent_ixmax - offset, imax(i) - offset, vv, xmid)
                        xmid += offset
                        pitchFrame.candidates(i).frequency = 1.0 / mme.dx / xmid
                        If (ymid > 1.0) Then
                            ymid = 1.0 / ymid
                        End If
                        pitchFrame.candidates(i).strength = ymid
                    End If
                Next
            Next
            ' /* Next frame. */
            '// progress (0.95, L"Sound to Pitch: path finder")
            'Melder_progress (0.95, L"Sound to Pitch: path finder")   
            Pitch_pathFinder(thee, silenceThreshold, voicingThreshold, octaveCost, octaveJumpCost, voicedUnvoicedCost, ceiling, False)

            Return thee
        Catch ex As Exception
            'Melder_throw (me, ": pitch analysis not performed.")
            Return Nothing
        End Try
    End Function

    Function Sound_to_Pitch_ac(ByRef mme As Sound,
    ByVal dt As Double, ByVal minimumPitch As Double, ByVal periodsPerWindow As Double, ByVal maxnCandidates As Integer, ByVal accurate As Integer,
    ByVal silenceThreshold As Double, ByVal voicingThreshold As Double,
    ByVal octaveCost As Double, ByVal octaveJumpCost As Double, ByVal voicedUnvoicedCost As Double, ByVal ceiling As Double) As Pitch

        Return Sound_to_Pitch_any(mme, dt, minimumPitch, periodsPerWindow, maxnCandidates, accurate,
            silenceThreshold, voicingThreshold, octaveCost, octaveJumpCost, voicedUnvoicedCost, ceiling)
    End Function

    Function Sound_to_Pitch(ByRef mme As Sound, ByVal timeStep As Double, ByVal minimumPitch As Double, ByVal maximumPitch As Double) As Pitch
        Return Sound_to_Pitch_ac(mme, timeStep, minimumPitch, 3.0, 15, False, 0.03, 0.45, 0.01, 0.35, 0.14, maximumPitch)
    End Function


    Function Sampled_xToIndex(ByRef mme As Sampled, ByVal x As Double) As Double
        Return (x - mme.x1) / mme.dx
    End Function

    Function Sampled_getValueAtSample(ByRef mme As Sampled, ByVal isamp As Long, ByVal ilevel As Long, ByVal unit As Integer) As Double
        If (isamp < 0 Or isamp > mme.nx - 1) Then
            Return NUMundefined
        End If
        Return mme.v_getValueAtSample(isamp, ilevel, unit)
    End Function


    Function Sampled_xToNearestIndex(ByRef mme As Sampled, ByVal x As Double) As Long
        Return Convert.ToInt32(Math.Floor((x - mme.x1) / mme.dx + 1.5))
    End Function

    Function Sampled_getValueAtX(ByRef mme As Sampled, ByVal x As Double, ByVal ilevel As Long, ByVal unit As Integer, ByVal interpolate As Integer) As Double
        If (x < mme.xmin) Then
            Return NUMundefined
        End If
        If x > mme.xmax Then
            Return NUMundefined
        End If
        If (interpolate) Then
            Dim ireal As Double = Sampled_xToIndex(mme, x)
            Dim ileft As Long = Math.Floor(ireal), inear, ifar
            Dim phase As Double = ireal - ileft
            If (phase < 0.5) Then
                inear = ileft
                ifar = ileft + 1
            Else
                ifar = ileft
                inear = ileft + 1
                phase = 1.0 - phase
            End If
            If (inear < 0 Or inear >= mme.nx) Then
                Return NUMundefined
            End If
            '// x out of range?
            Dim fnear As Double = mme.v_getValueAtSample(inear, ilevel, unit)
            '   // function value not defined?
            If (fnear = NUMundefined) Then
                Return NUMundefined
            End If
            '// at edge? Extrapolate
            If (ifar < 0 Or ifar > mme.nx - 1) Then
                Return fnear
            End If
            Dim ffar As Double = mme.v_getValueAtSample(ifar, ilevel, unit)
            '// neighbour undefined? Extrapolate
            If (ffar = NUMundefined) Then
                Return fnear
            End If
            '   // interpolate
            Return fnear + phase * (ffar - fnear)
        End If
        Return Sampled_getValueAtSample(mme, Sampled_xToNearestIndex(mme, x), ilevel, unit)
    End Function

    Function Pitch_getValueAtTime(ByRef mme As Pitch, ByVal time As Double, ByVal unit As Integer, ByVal interpolate As Integer) As Double
        Return Sampled_getValueAtX(mme, time, Pitch_LEVEL_FREQUENCY, unit, interpolate)
    End Function

    Function Sound_findMaximumCorrelation(ByRef mme As Sound, ByVal t1 As Double, ByVal windowLength As Double, ByVal tmin2 As Double, ByVal tmax2 As Double, ByRef tout As Double, ByRef peak As Double) As Double
        Dim maximumCorrelation = -1.0, r1 = 0.0, r2 = 0.0, r3 = 0.0, r1_best, r3_best, ir As Double
        Dim halfWindowLength As Double = 0.5 * windowLength
        Dim ileft1 As Long = Sampled_xToNearestIndex(mme, t1 - halfWindowLength)
        Dim iright1 As Long = Sampled_xToNearestIndex(mme, t1 + halfWindowLength)
        Dim ileft2min As Long = Sampled_xToLowIndex(mme, tmin2 - halfWindowLength)
        Dim ileft2max As Long = Sampled_xToHighIndex(mme, tmax2 - halfWindowLength)
        Dim i2 As Long
        '   /* Default. */
        peak = 0.0
        For ileft2 As Long = ileft2min - 1 To ileft2max - 1 Step 1
            Dim norm1 As Double = 0.0, norm2 = 0.0, product = 0.0, localPeak = 0.0
            If (mme.ny = 1) Then
                i2 = ileft2
                For i1 As Long = ileft1 - 1 To iright1 - 1 Step 1
                    If (i1 < 1 Or i1 > mme.nx Or i2 < 1 Or i2 > mme.nx) Then
                        Continue For
                    End If
                    Dim amp1 As Double = mme.z(0, i1), amp2 = mme.z(0, i2)
                    norm1 += amp1 * amp1
                    norm2 += amp2 * amp2
                    product += amp1 * amp2
                    If (Math.Abs(amp2) > localPeak) Then
                        localPeak = Math.Abs(amp2)
                    End If
                    i2 = i2 + 1
                Next
            Else
                i2 = ileft2
                For i1 As Long = ileft1 To iright1 Step 1
                    If (i1 < 1 Or i1 > mme.nx Or i2 < 1 Or i2 > mme.nx) Then
                        Continue For
                    End If
                    Dim amp1 As Double = 0.5 * (mme.z(0, i1) + mme.z(1, i1)), amp2 = 0.5 * (mme.z(0, i2) + mme.z(1, i2))
                    norm1 += amp1 * amp1
                    norm2 += amp2 * amp2
                    product += amp1 * amp2
                    If (Math.Abs(amp2) > localPeak) Then
                        localPeak = Math.Abs(amp2)
                    End If
                    i2 += 1
                Next
            End If
            r1 = r2
            r2 = r3
            If (product <> 0) Then
                r3 = product / (Math.Sqrt(norm1 * norm2))
            Else
                r3 = 0
            End If
            If (r2 > maximumCorrelation And r2 >= r1 And r2 >= r3) Then
                r1_best = r1
                maximumCorrelation = r2
                r3_best = r3
                ir = ileft2 - 1
                peak = localPeak
            End If
        Next
        '/*
        '* Improve the result by means of parabolic interpolation.
        '*/
        If (maximumCorrelation > -1.0) Then
            Dim d2r As Double = 2 * maximumCorrelation - r1_best - r3_best
            If (d2r <> 0.0) Then
                Dim dr As Double = 0.5 * (r3_best - r1_best)
                maximumCorrelation += 0.5 * dr * dr / d2r
                ir += dr / d2r
            End If
            tout = t1 + (ir - ileft1 + 1) * mme.dx
        End If
        Return maximumCorrelation
    End Function

    Function findExtremum_3(ByRef mme As Sound, ByVal channel1_index As Long, ByVal channel2_index As Long, ByVal d As Long, ByVal n As Long, ByVal includeMaxima As Integer, ByVal includeMinima As Integer) As Double
        Dim channel1 As Long = channel1_index + d
        Dim channel2 As Long
        If channel2_index < 0 Then
            channel2 = channel2_index
        Else
            channel2 = channel2_index + d
        End If
        Dim includeAll As Integer
        If includeMaxima = includeMinima Then
            includeAll = 1
        Else
            includeAll = 0
        End If
        Dim imin As Long = 1, imax = 1, i, iextr
        Dim minimum, maximum As Double
        If (n < 3) Then
            '/* Outside. */
            If (n <= 0) Then
                Return 0.0
            ElseIf (n = 1) Then
                Return 1.0
            Else
                '/* n == 2 */
                Dim x1 As Double = mme.z(channel1, 0)
                Dim x2 As Double = mme.z(channel1 + 1, 0)
                Dim xleft As Double
                Dim xright As Double
                If includeAll Then
                    xleft = Math.Abs(x1)
                    xright = Math.Abs(x2)
                Else
                    If includeMaxima Then
                        xleft = x1
                        xright = x2
                    Else
                        xleft = -x1
                        xright = -x2
                    End If
                End If

                If (xleft > xright) Then
                    Return 1.0
                ElseIf (xleft < xright) Then
                    Return 2.0
                Else
                    Return 1.5
                End If
            End If
        End If
        maximum = mme.z(0, channel1)
        minimum = maximum
        For i = 1 To n - 1 Step 1
            Dim value As Double = mme.z(0, channel1 + i)
            If (value < minimum) Then
                minimum = value
                imin = i
            End If
            If (value > maximum) Then
                maximum = value
                imax = i
            End If
        Next
        If (minimum = maximum) Then
            '/* All equal. */
            Return 0.5 * (n + 1.0)
        End If
        If includeAll Then
            If Math.Abs(minimum) > Math.Abs(maximum) Then
                iextr = imin
            Else
                iextr = imax
            End If
        Else
            If includeMaxima Then
                iextr = imax
            Else
                iextr = imin
            End If
        End If
        'iextr = includeAll ? ( fabs (minimum) > fabs (maximum) ? imin : imax ) : includeMaxima ? imax : imin
        If (iextr = 1) Then
            Return 1.0
        End If
        If (iextr = n) Then
            Return Convert.ToDouble(n)
        End If

        '/* Parabolic interpolation. */
        '/* We do NOT need fabs here: we look for a genuine extremum. */
        Dim valueMid As Double = mme.z(0, channel1 + iextr)
        Dim valueLeft As Double = mme.z(0, channel1 + iextr - 1)
        Dim valueRight As Double = mme.z(0, channel1 + iextr + 1)
        Return iextr + 0.5 * (valueRight - valueLeft) / (2 * valueMid - valueLeft - valueRight)
    End Function

    Function Sound_findExtremum(ByRef mme As Sound, ByVal tmin As Double, ByVal tmax As Double, ByVal includeMaxima As Integer, ByVal includeMinima As Integer) As Double
        Dim imin As Long = Sampled_xToLowIndex(mme, tmin), imax = Sampled_xToHighIndex(mme, tmax)
        If (tmin = NUMundefined) Then
            Console.WriteLine("Error in Find Extremum: tmin is undefined")
            Return -1
        End If
        If (tmax = NUMundefined) Then
            Console.WriteLine("Error in Find Extremum: tmax is undefined")
            Return -1
        End If
        If (imin < 0) Then
            imin = 0
        End If
        If (imax > mme.nx) Then
            imax = mme.nx
        End If
        Dim g As Double
        If mme.ny > 1 Then
            g = 1
        Else
            g = -1
        End If
        ' 0 = index in 1 dimension of z = double *channel1_base
        ' g - index in 1 dimension of z = double *channel2_base
        Dim iextremum As Double = findExtremum_3(mme, 0, g, imin - 1, imax - imin + 1, includeMaxima, includeMinima)
        If (iextremum <> 0) Then
            Return mme.x1 + (imin - 1 + iextremum) * mme.dx
        Else
            Return (tmin + tmax) / 2
        End If
    End Function

    Function Pitch_isVoiced_i(ByRef mme As Pitch, ByVal iframe As Long) As Boolean
        Return Sampled_getValueAtSample(mme, iframe, Pitch_LEVEL_FREQUENCY, 0) <> NUMundefined
    End Function

    Function Pitch_getVoicedIntervalAfter(ByRef mme As Pitch, ByVal after As Double, ByRef tleft As Double, ByRef tright As Double) As Integer
        Dim ileft As Long = Sampled_xToHighIndex(mme, after), iright
        '/* Offright. */
        If (ileft > mme.nx - 1) Then
            Return 0
        End If
        '/* Offleft. */
        If (ileft < 0) Then
            ileft = 0
        End If


        '/* Search for first voiced frame. */
        For ileft = ileft To mme.nx - 1 Step 1
            If (Pitch_isVoiced_i(mme, ileft)) Then
                GoTo exitfor
            End If
        Next
        ' /* Offright. */
exitfor: If (ileft > mme.nx - 1) Then
            Return 0
        End If

        '/* Search for last voiced frame. */
        For iright = ileft To mme.nx - 1 Step 1
            If (Not Pitch_isVoiced_i(mme, iright)) Then
                GoTo exitfor1
            End If
        Next
exitfor1: iright -= 1

        '/* The whole frame is considered voiced. */
        tleft = Sampled_indexToX(mme, ileft) - 0.5 * mme.dx
        tright = Sampled_indexToX(mme, iright) + 0.5 * mme.dx
        If (tleft >= mme.xmax - 0.5 * mme.dx) Then
            Return 0
        End If
        If (tleft < mme.xmin) Then
            tleft = mme.xmin
        End If
        If (tright > mme.xmax) Then
            tright = mme.xmax
        End If
        Return 1
    End Function

    Function improve_evaluate(ByVal x As Double, ByRef r() As Double, ByVal offset As Long, ByVal depth As Long, ByVal nx As Long, ByVal isMaximum As Integer) As Double
        'struct improve_params *me = (struct improve_params *) closure
        Dim y As Double = NUM_interpolate_sinc(r, offset, nx, x, depth)
        If isMaximum Then
            Return -y
        Else
            Return y
        End If
    End Function

    Function NUMminimize_brent(ByVal a As Double, ByVal b As Double, ByRef r() As Double, ByRef offset As Long, ByVal depth As Long, ByVal nx As Long, ByVal isMaximum As Integer, ByVal tol As Double, ByRef fx As Double) As Double
        Dim x, v, fv, w, fw As Double
        Dim golden As Double = 1 - NUM_goldenSection
        Dim sqrt_epsilon As Double = 0.000000010536712127723509
        'Math.Sqrt(Double.Epsilon)(NUMfpp -> eps)

        Dim itermax As Long = 60

        '!!!Melder_assert (tol > 0 && a < b)
        If Not (tol > 0 And a < b) Then
            Console.WriteLine("Cannot continue process: (tol > 0 And a < b) ")
            Return NUMundefined
        End If

        '/* First step - golden section */

        v = a + golden * (b - a)
        fv = improve_evaluate(v, r, offset, depth, nx, isMaximum)
        x = v
        w = v
        fx = fv
        fw = fv

        For iter As Long = 1 To itermax Step 1
            Dim range As Double = b - a
            Dim middle_range As Double = (a + b) / 2
            Dim tol_act As Double = sqrt_epsilon * Math.Abs(x) + tol / 3
            Dim new_step As Double '/* Step at this iteration */



            If (Math.Abs(x - middle_range) + range / 2 <= 2 * tol_act) Then
                Return x
            End If

            '/* Obtain the golden section step */

            If (x < middle_range) Then
                new_step = b - x
            Else
                new_step = a - x
            End If
            new_step = golden * new_step

            '/* Decide if the parabolic interpolation can be tried	*/

            If (Math.Abs(x - w) >= tol_act) Then
                '/*
                '	Interpolation step is calculated as p/q
                '	division operation is delayed until last moment.
                '*/

                Dim p, q, t1 As Double

                t1 = (x - w) * (fx - fv)
                q = (x - v) * (fx - fw)
                p = (x - v) * q - (x - w) * t1
                q = 2 * (q - t1)

                If (q > 0) Then
                    p = -p
                Else
                    q = -q
                End If

                '/*
                '	If x+p/q falls in [a,b], not too close to a and b,
                '	and isn't too large, it is accepted.
                '	If p/q is too large then the golden section procedure can
                '	reduce [a,b] range.
                '*/

                If (Math.Abs(p) < Math.Abs(new_step * q) And p > q * (a - x + 2 * tol_act) And p < q * (b - x - 2 * tol_act)) Then
                    new_step = p / q
                End If
            End If

            '/* Adjust the step to be not less than tolerance. */

            If (Math.Abs(new_step) < tol_act) Then
                If new_step > 0 Then
                    new_step = tol_act
                Else
                    new_step = -tol_act
                End If
            End If

            '/* Obtain the next approximation to min	and reduce the enveloping range */

            Dim t As Double = x + new_step '	/* Tentative point for the min	*/
            Dim ft As Double
            ft = improve_evaluate(t, r, offset, depth, nx, isMaximum)

            '/*
            '	If t is a better approximation, reduce the range so that
            '	t would fall within it. If x remains the best, reduce the range
            '	so that x falls within it.
            '*/

            If (ft <= fx) Then
                If (t < x) Then
                    b = x
                Else
                    a = x
                End If

                v = w
                w = x
                x = t

                fv = fw
                fw = fx
                fx = ft
            Else
                If (t < x) Then
                    a = t
                Else
                    b = t
                End If

                If (ft <= fw Or w = x) Then
                    v = w
                    w = t
                    fv = fw
                    fw = ft
                ElseIf (ft <= fv Or v = x Or v = w) Then
                    v = t
                    fv = ft
                End If
            End If
        Next
        'Melder_warning (L"NUMminimize_brent: maximum number of iterations (", Melder_integer (itermax), L") exceeded.")
        Return x
    End Function

    Function NUMimproveExtremum(ByRef y() As Double, ByVal offset As Long, ByVal nx As Long, ByVal ixmid As Long, ByVal interpolation As Integer, ByRef ixmid_real As Double, ByVal isMaximum As Integer) As Double
        '!!! struct improve_params params
        Dim result As Double
        If (ixmid <= 1) Then
            ixmid_real = 1
            Return y(offset)
        End If

        If (ixmid >= nx) Then
            ixmid_real = nx
            Return y(offset + nx - 1)
        End If
        If (interpolation <= NUM_PEAK_INTERPOLATE_NONE) Then
            ixmid_real = ixmid
            Return y(offset + ixmid)
        End If
        If (interpolation = NUM_PEAK_INTERPOLATE_PARABOLIC) Then
            Dim dy As Double = 0.5 * (y(offset + ixmid + 1) - y(offset + ixmid - 1))
            Dim d2y As Double = 2 * y(offset + ixmid) - y(offset + ixmid - 1) - y(offset + ixmid + 1)
            ixmid_real = ixmid + dy / d2y
            Return y(offset + ixmid) + 0.5 * dy * dy / d2y
        End If
        '/* Sinc interpolation. */
        'params.y = y
        Dim depth As Long
        If interpolation = NUM_PEAK_INTERPOLATE_SINC70 Then
            depth = 70
        Else
            depth = 700
        End If
        'params.ixmax = nx
        'params.isMaximum = isMaximum
        '!!! improve_evaluate
        ixmid_real = NUMminimize_brent(ixmid - 1, ixmid + 1, y, offset, depth, nx, isMaximum, 0.0000000001, result)
        If isMaximum Then
            Return -result
        Else
            Return result
        End If
    End Function

    Function NUMimproveMaximum(ByRef y() As Double, ByVal offset As Long, ByVal nx As Long, ByVal ixmid As Long, ByVal interpolation As Integer, ByRef ixmid_real As Double) As Double
        Return NUMimproveExtremum(y, offset, nx, ixmid, interpolation, ixmid_real, 1)
    End Function
    Function NUMimproveMinimum(ByRef y() As Double, ByVal offset As Long, ByVal nx As Long, ByVal ixmid As Long, ByVal interpolation As Integer, ByRef ixmid_real As Double) As Double
        Return NUMimproveExtremum(y, offset, nx, ixmid, interpolation, ixmid_real, 0)
    End Function
    Function Sampled_getWindowSamples(ByRef mme As Sampled, ByVal xmin As Double, ByVal xmax As Double, ByRef ixmin As Long, ByRef ixmax As Long) As Long
        Dim rixmin As Double = Math.Ceiling((xmin - mme.x1) / mme.dx)
        Dim rixmax As Double = Math.Floor((xmax - mme.x1) / mme.dx)
        'ixmin = rixmin < 1.0 ? 1 : (long) rixmin
        If rixmin < 0.0 Then
            ixmin = 0
        Else
            ixmin = Convert.ToInt32(rixmin)
        End If
        If ixmax > Convert.ToDouble(mme.nx) Then
            ixmax = mme.nx - 1
        Else
            ixmax = Convert.ToInt32(rixmax)
        End If
        '= rixmax > (double) my nx ? my nx : (long) rixmax
        If (ixmin > ixmax) Then
            Return 0
        End If
        Return ixmax - ixmin + 1
    End Function

    Function NUM_interpolate_sinc(ByRef y() As Double, ByRef offset As Long, ByVal nx As Long, ByVal x As Double, ByRef maxDepth As Long) As Double
        Dim ix, midleft As Long, midright, left, right
        midleft = Math.Floor(x)
        midright = midleft + 1
        Dim result As Double = 0.0, a, halfsina, aa, daa

        If (nx < 0) Then
            Return NUMundefined
        End If
        If (x > nx - 1) Then
            Return y(offset + nx - 1)
        End If
        If (x < 0) Then
            Return y(offset + 0)
        End If
        If (x = midleft - 1) Then
            Return y(offset + midleft - 1)
        End If
        If (maxDepth > midright - 1) Then
            maxDepth = midright - 1
        End If
        If (maxDepth > nx - midleft) Then
            maxDepth = nx - midleft
        End If
        If (maxDepth <= NUM_VALUE_INTERPOLATE_NEAREST) Then
            Return y(offset + Convert.ToUInt32(Math.Floor(x + 0.5)))
        End If
        If (maxDepth = NUM_VALUE_INTERPOLATE_LINEAR) Then
            Return y(offset + midleft) + (x - midleft) * (y(offset + midright) - y(offset + midleft))
        End If
        If (maxDepth = NUM_VALUE_INTERPOLATE_CUBIC) Then
            Dim yl As Double = y(offset + midleft), yr = y(offset + midright)
            Dim dyl As Double = 0.5 * (yr - y(offset + midleft - 1)), dyr = 0.5 * (y(offset + midright + 1) - yl)
            Dim fil As Double = x - midleft, fir = midright - x
            Return yl * fir + yr * fil - fil * fir * (0.5 * (dyr - dyl) + (fil - 0.5) * (dyl + dyr - 2 * (yr - yl)))
        End If
        left = midright - maxDepth
        right = midleft + maxDepth
        a = Math.PI * (x - midleft)
        halfsina = 0.5 * Math.Sin(a)
        aa = a / (x - left + 1)
        daa = Math.PI / (x - left + 1)
        For ix = midleft To left Step -1
            Dim d As Double = halfsina / a * (1.0 + Math.Cos(aa))
            result += y(offset + ix) * d
            a += Math.PI
            aa += daa
            halfsina = -halfsina
        Next
        a = Math.PI * (midright - x)
        halfsina = 0.5 * Math.Sin(a)
        aa = a / (right - x + 1)
        daa = Math.PI / (right - x + 1)
        For ix = midright To right Step 1
            Dim d As Double = halfsina / a * (1.0 + Math.Cos(aa))
            result += y(offset + ix) * d
            a += Math.PI
            aa += daa
            halfsina = -halfsina
        Next
        Return result
    End Function

    Function Vector_getValueAtX(ByRef mme As Vector, ByVal x As Double, ByVal ilevel As Long, ByVal interpolation As Integer) As Double
        Dim leftEdge As Double = mme.x1 - 0.5 * mme.dx, rightEdge = leftEdge + mme.nx * mme.dx
        If (x < leftEdge Or x > rightEdge) Then
            Return NUMundefined
        End If
        If (ilevel > Vector_CHANNEL_AVERAGE) Then
            '!!!Melder_assert (ilevel <= my ny)
            Dim interp As Long
            If interpolation = Vector_VALUE_INTERPOLATION_SINC70 Then
                interp = NUM_VALUE_INTERPOLATE_SINC70
            ElseIf interpolation = Vector_VALUE_INTERPOLATION_SINC700 Then
                interp = NUM_VALUE_INTERPOLATE_SINC700
            Else
                interp = interpolation
            End If
            Dim ch1(mme.nx - 1) As Double
            For i As Long = 0 To mme.nx - 1
                ch1(i) = mme.z(ilevel, i)
            Next

            Return NUM_interpolate_sinc(ch1, 0, mme.nx, Sampled_xToIndex(mme, x), interp)
        End If
        Dim sum As Double = 0.0
        For channel As Long = 1 To mme.ny Step 1
            Dim interp As Long
            If interpolation = Vector_VALUE_INTERPOLATION_SINC70 Then
                interp = NUM_VALUE_INTERPOLATE_SINC70
            ElseIf interpolation = Vector_VALUE_INTERPOLATION_SINC700 Then
                interp = NUM_VALUE_INTERPOLATE_SINC700
            Else
                interp = interpolation
            End If
            Dim ch2(mme.nx - 1) As Double
            For i As Long = 0 To mme.nx - 1
                ch2(i) = mme.z(channel, i)
            Next
            sum += NUM_interpolate_sinc(ch2, 0, mme.nx, Sampled_xToIndex(mme, x), interp)
        Next
        Return sum / mme.ny
    End Function
    Sub Vector_getMaximumAndX(ByRef mme As Vector, xmin As Double, xmax As Double, channel As Long, interpolation As Integer, ByRef return_maximum As Double, ByRef return_xOfMaximum As Double)
        Dim n As Long = mme.nx, imin = 0, imax = 0, i
        '!!!Melder_assert (channel >= 1 && channel <= my ny)
        'double *y = my z [channel]
        Dim maximum, x As Double
        If (xmax <= xmin) Then
            xmin = mme.xmin
            xmax = mme.xmax
        End If
        Dim nS As Long = Sampled_getWindowSamples(mme, xmin, xmax, imin, imax)
        If (nS = 0) Then
            '/*
            ' * No samples between xmin and xmax.
            ' * Try to return the greater of the values at these two points.
            ' */
            Dim k As Long
            If (interpolation > Vector_VALUE_INTERPOLATION_NEAREST) Then
                k = Vector_VALUE_INTERPOLATION_LINEAR
            Else
                k = Vector_VALUE_INTERPOLATION_NEAREST
            End If
            Dim yleft As Double = Vector_getValueAtX(mme, xmin, channel, k)
            Dim yright As Double = Vector_getValueAtX(mme, xmax, channel, k)
            If yleft > yright Then
                maximum = yleft
            Else
                maximum = yright
            End If
            If yleft = yright Then
                x = (xmin + xmax) / 2
            ElseIf (yleft > yright) Then
                x = xmin
            Else
                x = xmax
            End If
        Else
            maximum = mme.z(channel, imin)
            x = imin
            If (mme.z(channel, imax) > maximum) Then
                maximum = mme.z(channel, imax)
                x = imax
            End If
            If (imin = 0) Then
                imin += 1
            End If
            If (imax = mme.nx - 1) Then
                imax -= 1
            End If
            Dim ch1(n) As Double
            For ii = 0 To n - 1
                ch1(ii) = mme.z(channel, ii)
            Next
            For i = imin To imax - 1 Step 1
                If (mme.z(channel, i) > mme.z(channel, i - 1) And mme.z(channel, i) >= mme.z(channel, i + 1)) Then
                    Dim i_real As Double = 0
                    Dim localMaximum As Double = NUMimproveMaximum(ch1, 0, n, i, interpolation, i_real)
                    If (localMaximum > maximum) Then
                        maximum = localMaximum
                        x = i_real
                    End If
                End If

            Next
            '/* Convert sample to x. */
            x = mme.x1 + x * mme.dx
            If (x < xmin) Then
                x = xmin
            ElseIf (x > xmax) Then
                x = xmax
            End If
        End If
        If (Not IsNothing(return_maximum)) Then
            return_maximum = maximum
        End If
        'If (return_xOfMaximum) Then
        'return_xOfMaximum = x
        'End If
    End Sub

    Sub Vector_getMinimumAndX(ByRef mme As Vector, xmin As Double, xmax As Double, channel As Long, interpolation As Integer, ByRef return_minimum As Double, return_xOfMinimum As Double)
        Dim n As Long = mme.nx, imin = 0, imax = 0
        '!!!Melder_assert (channel >= 1 && channel <= my ny)
        'double *y = my z [channel]
        Dim minimum, x As Double
        If (xmax <= xmin) Then
            xmin = mme.xmin
            xmax = mme.xmax
        End If
        Dim nS As Long = Sampled_getWindowSamples(mme, xmin, xmax, imin, imax)
        If (nS = 0) Then
            '/*
            ' * No samples between xmin and xmax.
            ' * Try to return the greater of the values at these two points.
            ' */
            Dim k As Long
            If (interpolation > Vector_VALUE_INTERPOLATION_NEAREST) Then
                k = Vector_VALUE_INTERPOLATION_LINEAR
            Else
                k = Vector_VALUE_INTERPOLATION_NEAREST
            End If
            Dim yleft As Double = Vector_getValueAtX(mme, xmin, channel, k)
            Dim yright As Double = Vector_getValueAtX(mme, xmax, channel, k)
            If yleft < yright Then
                minimum = yleft
            Else
                minimum = yright
            End If
            If yleft = yright Then
                x = (xmin + xmax) / 2
            ElseIf (yleft < yright) Then
                x = xmin
            Else
                x = xmax
            End If
        Else
            minimum = mme.z(channel, imin)
            x = imin
            If (mme.z(channel, imax) < minimum) Then
                minimum = mme.z(channel, imax)
                x = imax
            End If
            If (imin = 0) Then
                imin += 1
            End If
            If (imax = mme.nx - 1) Then
                imax -= 1
            End If
            Dim ch1(n) As Double
            For ii = 0 To n - 1
                ch1(ii) = mme.z(channel, ii)
            Next
            For i = imin To imax - 1 Step 1
                If (mme.z(channel, i) < mme.z(channel, i - 1) And mme.z(channel, i) <= mme.z(channel, i + 1)) Then
                    Dim i_real As Double = 0
                    Dim localminimum As Double = NUMimproveMinimum(ch1, 0, n, i, interpolation, i_real)
                    If (localminimum < minimum) Then
                        minimum = localminimum
                        x = i_real
                    End If
                End If

            Next
            '/* Convert sample to x. */
            x = mme.x1 + x * mme.dx
            If (x < xmin) Then
                x = xmin
            ElseIf (x > xmax) Then
                x = xmax
            End If
        End If
        If (Not IsNothing(return_minimum)) Then
            return_minimum = minimum
        End If
        'If (return_xOfminimum) Then
        'return_xOfminimum = x
        'End If
    End Sub

    Sub Vector_getMaximumAndXAndChannel(ByRef mme As Vector, xmin As Double, xmax As Double, interpolation As Integer, ByRef return_maximum As Double)
        Dim maximum, xOfMaximum As Double
        Dim channelOfMaximum As Long = 1
        Vector_getMaximumAndX(mme, xmin, xmax, 0, interpolation, maximum, xOfMaximum)
        For channel As Long = 2 To mme.ny Step 1
            Dim maximumOfChannel, xOfMaximumOfChannel As Double
            Vector_getMaximumAndX(mme, xmin, xmax, channel, interpolation, maximumOfChannel, xOfMaximumOfChannel)
            If (maximumOfChannel > maximum) Then
                maximum = maximumOfChannel
                xOfMaximum = xOfMaximumOfChannel
                channelOfMaximum = channel
            End If
        Next
        If (Not IsNothing(return_maximum)) Then
            return_maximum = maximum
        End If
        '  If (return_xOfMaximum) Then
        'return_xOfMaximum = xOfMaximum
        '  End If
        ' If (return_channelOfMaximum) Then
        '      return_channelOfMaximum = channelOfMaximum
        '  End If
    End Sub
    Function Vector_getMaximum(ByRef mme As Vector, xmin As Double, xmax As Double, interpolation As Integer) As Double
        Dim maximum As Double
        Vector_getMaximumAndXAndChannel(mme, xmin, xmax, interpolation, maximum)
        Return maximum
    End Function
    Sub Vector_getMinimumAndXAndChannel(ByRef mme As Vector, xmin As Double, xmax As Double, interpolation As Integer, ByRef return_minimum As Double)
        Dim minimum, xOfMinimum As Double
        Dim channelOfMinimum As Long = 1
        Vector_getMinimumAndX(mme, xmin, xmax, 0, interpolation, minimum, xOfMinimum)
        For channel As Long = 2 To mme.ny Step 1
            Dim minimumOfChannel, xOfMinimumOfChannel As Double
            Vector_getMinimumAndX(mme, xmin, xmax, channel, interpolation, minimumOfChannel, xOfMinimumOfChannel)
            If (minimumOfChannel < minimum) Then
                minimum = minimumOfChannel
                xOfMinimum = xOfMinimumOfChannel
                channelOfMinimum = channel
            End If
        Next
        If (Not IsNothing(return_minimum)) Then
            return_minimum = minimum
        End If

    End Sub
    Function Vector_getMinimum(ByRef mme As Vector, xmin As Double, xmax As Double, interpolation As Integer) As Double
        Dim minimum As Double
        Vector_getMinimumAndXAndChannel(mme, xmin, xmax, interpolation, minimum)
        Return minimum
    End Function
    Private Function Vector_getAbsoluteExtremum(ByRef mme As Vector, xmin As Double, xmax As Double, interpolation As Integer) As Double
        Dim minimum As Double = Math.Abs(Vector_getMinimum(mme, xmin, xmax, interpolation))
        Dim maximum As Double = Math.Abs(Vector_getMaximum(mme, xmin, xmax, interpolation))
        If (minimum > maximum) Then
            Return minimum
        Else
            Return maximum
        End If
    End Function
    Function Sound_Pitch_to_PointProcess_cc(ByRef sound As Sound, ByRef pitch As Pitch) As PointProcess
        Try
            Dim point As PointProcess = New PointProcess(sound.xmin, sound.xmax, 10)
            Dim t As Double = pitch.xmin
            Dim addedRight As Double = -1.0E+300
            Dim globalPeak As Double = Vector_getAbsoluteExtremum(sound, sound.xmin, sound.xmax, 0), peak = 0

            '/*
            ' * Cycle over all voiced intervals.
            '*/
            'autoMelderProgress progress (L"Sound & Pitch: To PointProcess...")
            While 1 = 1
                Dim tleft, tright As Double
                Dim nS As Integer = Pitch_getVoicedIntervalAfter(pitch, t, tleft, tright)
                If (nS = 0) Then
                    GoTo endofglobalwhile
                End If
                '/*
                ' * Go to the middle of the voice stretch.
                '*/
                Dim tmiddle As Double = (tleft + tright) / 2
                'Melder_progress ((tmiddle - sound -> xmin) / (sound -> xmax - sound -> xmin), L"Sound & Pitch to PointProcess")
                Dim f0middle As Double = Pitch_getValueAtTime(pitch, tmiddle, 0, Pitch_LINEAR)

                '/*
                '* Our first point is near this middle.
                '*/
                If (f0middle = NUMundefined) Then
                    'Melder_fatal("Sound_Pitch_to_PointProcess_cc: tleft %ls, tright %ls, f0middle %ls", Melder_double(tleft), Melder_double(tright), Melder_double(f0middle))
                    Console.WriteLine("Sound_Pitch_to_PointProcess_cc: f0middle = NUMundefined")
                    Return Nothing
                End If
                Dim tmax As Double = Sound_findExtremum(sound, tmiddle - 0.5 / f0middle, tmiddle + 0.5 / f0middle, True, True)
                If (tmax = NUMundefined) Then
                    Console.WriteLine("Error in Find Extremum: tmax is undefined")
                    Return Nothing
                End If
                point.addPoint(tmax)

                Dim tsave As Double = tmax
                While 1 = 1
                    Dim f0 As Double = Pitch_getValueAtTime(pitch, tmax, 0, Pitch_LINEAR), correlation
                    If (f0 = NUMundefined) Then
                        GoTo endofwhile1
                    End If
                    correlation = Sound_findMaximumCorrelation(sound, tmax, 1.0 / f0, tmax - 1.25 / f0, tmax - 0.8 / f0, tmax, peak)
                    If (correlation = -1) Then
                        tmax -= 1.0 / f0
                    End If
                    '/* This one period will drop out. */
                    If (tmax < tleft) Then
                        If (correlation > 0.7 And peak > 0.023333 * globalPeak And (tmax - addedRight) > (0.8 / f0)) Then
                            point.addPoint(tmax)
                        End If
                        GoTo endofwhile1
                    End If
                    If (correlation > 0.3 And (peak = 0.0 Or peak > 0.01 * globalPeak)) Then
                        If (tmax - addedRight > 0.8 / f0) Then
                            '// do not fill in a short originally unvoiced interval twice
                            point.addPoint(tmax)
                        End If
                    End If
                End While
endofwhile1:    tmax = tsave
                While 1 = 1
                    Dim f0 As Double = Pitch_getValueAtTime(pitch, tmax, 0, Pitch_LINEAR), correlation
                    If (f0 = NUMundefined) Then
                        GoTo endofwhile2
                    End If
                    correlation = Sound_findMaximumCorrelation(sound, tmax, 1.0 / f0, tmax + 0.8 / f0, tmax + 1.25 / f0, tmax, peak)
                    If (correlation = -1) Then
                        tmax = tmax + 1.0 / f0
                    End If
                    If (tmax > tright) Then
                        If (correlation > 0.7 And peak > 0.023333 * globalPeak) Then
                            point.addPoint(tmax)
                            addedRight = tmax
                        End If
                        GoTo endofwhile2
                    End If
                    If (correlation > 0.3 And (peak = 0.0 Or peak > 0.01 * globalPeak)) Then
                        point.addPoint(tmax)
                        addedRight = tmax
                    End If
                End While
endofwhile2:    t = tright
            End While
endofglobalwhile: Return point
        Catch ex As Exception
            Console.WriteLine("sound and pitch,  not converted to PointProcess (cc).")
            Return Nothing
        End Try

    End Function

    Function Sampled_indexToX(ByRef mme As Sampled, ByVal i As Long) As Double
        Return mme.x1 + i * mme.dx
    End Function

    Function Pitch_to_PitchTier(ByRef mme As Pitch) As PitchTier
        Try
            Dim thee As PitchTier = New PitchTier(mme.xmin, mme.xmax)
            For i As Long = 0 To mme.nx - 1 Step 1
                Dim frequency As Double = mme.frame(i).candidates(0).frequency

                '/*
                ' * Count only voiced frames.
                '*/
                If (frequency > 0.0 And frequency < mme.ceiling) Then
                    Dim time As Double = Sampled_indexToX(mme, i)
                    thee.RealTier_addPoint(time, frequency)
                End If
            Next
            Return thee
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Sub Vector_subtractMean(ByRef mme As Vector)
        For channel As Long = 0 To mme.ny - 1 Step 1
            Dim sum As Double = 0.0
            For i As Long = 0 To mme.nx - 1 Step 1
                sum += mme.z(channel, i)
            Next
            Dim mean As Double = sum / mme.nx
            For i As Long = 0 To mme.nx - 1 Step 1
                mme.z(channel, i) -= mean
            Next
        Next
    End Sub
    Sub Main()
        Dim s As Sound = Sound_readFromSoundFile("C:\Users\fae1987\Documents\LOL\PraatSound\melanie_PraatSound.wav")
        Vector_subtractMean(s)
        Dim timestep As Double = 0.01
        Dim minimumPitch As Double = 75
        Dim maximumPitch As Double = 600
        Dim pit As Pitch = Sound_to_Pitch(s, timestep, minimumPitch, maximumPitch)

        Dim pulses As PointProcess = Sound_Pitch_to_PointProcess_cc(s, pit)
        Dim pitch As PitchTier = Pitch_to_PitchTier(pit)
        Dim duration As DurationTier = New DurationTier(pit.xmin, pit.xmax)
        pitch.points(12).value = 270
        pitch.points(13).value = 280
        pitch.points(14).value = 290
        pitch.points(15).value = 295
        pitch.points(16).value = 270
        pitch.points(17).value = 280
        pitch.points(18).value = 290
        pitch.points(19).value = 295
        pitch.points(20).value = 290
        pitch.points(21).value = 290
        pitch.points(22).value = 290
        pitch.points(23).value = 290
        pitch.points(24).value = 290
        pitch.points(25).value = 290
        pitch.points(26).value = 290
        pitch.points(27).value = 295
        duration.addPoint(0.5 * (s.xmin + s.xmax), 1)
        Dim sProcessed As Sound = Sound_Point_Pitch_Duration_to_Sound(s, pulses, pitch, duration, MAX_T)
        Dim outFilePath As String = "C:\Users\fae1987\Documents\LOL\PraatSound\melanie_outfile.pcm"
        Dim writer = New System.IO.BinaryWriter(System.IO.File.Open(outFilePath, System.IO.FileMode.Create))
        MelderFile_writeFloatToAudio(writer, sProcessed.ny, 0, sProcessed.z, sProcessed.nx)
        writer.Close()

    End Sub



End Module
