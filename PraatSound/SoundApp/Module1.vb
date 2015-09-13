﻿Module Module1
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
                thee.z(1, iTarget) += mme.z(1, i) * 0.5 * (1 + Math.Cos(dphase * (i - imin + 0.5)))
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
                thee.z(1, iTarget) += mme.z(1, i) * 0.5 * (1 - Math.Cos(dphase * (i - imin + 0.5)))
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

    Function Sound_Point_Pitch_Duration_to_Sound(ByRef mme As Sound, ByVal pulses As PointProcess, ByVal pitch As PitchTier, ByVal duration As DurationTier, ByVal maxT As Double) As Sound

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

    End Function
    Public Function Sound_createSimple(ByVal numberOfChannels As Long, ByVal duration As Double, ByVal samplingFrequency As Double) As Sound
        Return New Sound(numberOfChannels, 0.0, duration, Math.Floor(duration * samplingFrequency + 0.5), 1 / samplingFrequency, 0.5 / samplingFrequency)
    End Function
    Function bingeti2LE(ByRef reader As System.IO.BinaryReader) As Integer
        If (Melder_debug <> 18) Then
            Dim s As Short
            s = reader.ReadInt16()
            'if (fread (& s, sizeof (signed short), 1, f) != 1) Then 
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
            'if (fread (& l, sizeof (long), 1, f) != 1) Then
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
                        For isamp As Long = 1 To numberOfSamples Step 1
                            For ichan As Long = 1 To numberOfChannels Step 1
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
                    Console.WriteLine("Too few bits per sample (" + numberOfBitsPerSamplePoint + "; the minimum is 4).")
                    Return
                ElseIf (numberOfBitsPerSamplePoint > 32) Then
                    Console.WriteLine("Too few bits per sample (" + numberOfBitsPerSamplePoint + "; the maximum is 32).")
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
                        '// incorrect data chunk (sometimes -44); assume that the data run till the end of the file
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
            ' autoMelderFile mfile = MelderFile_open (file);
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
            If (reader.BaseStream.Seek(0, IO.SeekOrigin.Begin) = IO.SeekOrigin.End) Then
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
        Return Math.Log10(x) * NUMLog2E
    End Function
    Public Function exp(ByVal x As Double) As Double
        Return Math.Pow(NUMEulersConstant, x)
    End Function

    Sub Sampled_shortTermAnalysis(ByRef mme As Sampled, ByVal windowDuration As Double, ByVal timeStep As Double, ByVal numberOfFrames As Long, ByVal firstTime As Double)
        '!!!Melder_assert (windowDuration > 0.0);
        '!!!Melder_assert (timeStep > 0.0);
        Dim myDuration As Double = mme.dx * mme.nx
        If (windowDuration > myDuration) Then
            Console.WriteLine(": shorter than window length.")
        End If
        numberOfFrames = Math.Floor((myDuration - windowDuration) / timeStep) + 1
        '!!!Melder_assert (*numberOfFrames >= 1);
        Dim ourMidTime As Double = mme.x1 - 0.5 * mme.dx + 0.5 * myDuration
        Dim thyDuration As Double = numberOfFrames * timeStep
        firstTime = ourMidTime - 0.5 * thyDuration + 0.5 * timeStep
    End Sub

    Sub Pitch_pathFinder(ByRef mme As Pitch, ByVal silenceThreshold As Double, ByVal voicingThreshold As Double, ByVal octaveCost As Double, ByVal octaveJumpCost As Double, ByVal voicedUnvoicedCost As Double, ByVal ceiling As Double, ByVal pullFormants As Integer)
        'TODO
    End Sub

    Sub drftb1(ByVal n As Long, ByRef c() As Double, ByRef ch() As Double, ByRef wa() As Double, ByRef ifac() As Long)
        'TODO
    End Sub
    Sub drftf1(ByVal n As Long, ByRef c() As Double, ByRef ch() As Double, ByRef wa() As Double, ByRef ifac() As Long)
        'TODO
    End Sub
    Sub NUMfft_backward(ByRef mme As NumFFTTable, ByRef data() As Double)
        If (mme.n = 1) Then
            Return
        End If
        Dim tr2() As Double = New Double(mme.trigcache.Length - mme.n + 1) {}
        Dim i As Integer = mme.n
        Dim j As Integer = 0
        While i < mme.trigcache.Length
            tr2(j) = mme.trigcache(i)
            j += 1
            i += 1
        End While

        drftb1(mme.n, data, mme.trigcache, tr2, mme.splitcache)
    End Sub
    Sub NUMfft_forward(ByRef mme As NumFFTTable, ByRef data() As Double)
        If (mme.n = 1) Then
            Return
        End If
        Dim tr2() As Double = New Double(mme.trigcache.Length - mme.n + 1) {}
        Dim i As Integer = mme.n
        Dim j As Integer = 0
        While i < mme.trigcache.Length
            tr2(j) = mme.trigcache(i)
            j += 1
            i += 1
        End While

        drftf1(mme.n, data, mme.trigcache, tr2, mme.splitcache)
    End Sub
    Function Sound_to_Pitch_any(ByRef mme As Sound,
    ByVal dt As Double, ByVal minimumPitch As Double, ByVal periodsPerWindow As Double, ByVal maxnCandidates As Integer,
    ByVal method As Integer,
    ByVal silenceThreshold As Double, ByVal voicingThreshold As Double,
    ByVal octaveCost As Double, ByVal octaveJumpCost As Double, ByVal voicedUnvoicedCost As Double, ByVal ceiling As Double) As Pitch
        Try
            Dim fftTable As NumFFTTable = New NumFFTTable(1)
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

            '!!!Melder_assert (maxnCandidates >= 2);
            '!!!Melder_assert (method >= AC_HANNING && method <= FCC_ACCURATE);

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
                    ';   /* Because Gaussian window is twice as long. */
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
                'Melder_throw ("To analyse this Sound, ", L_LEFT_DOUBLE_QUOTE, "minimum pitch", L_RIGHT_DOUBLE_QUOTE, " must not be less than ", periodsPerWindow / duration, " Hz.");
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
                'Melder_throw ("The pitch analysis would give zero pitch frames.");
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
                frame = New Double(mme.ny, nsampFFT) {}
                secondRes = nsampFFT
                ReDim windowR(nsampFFT)
                ReDim window(nsamp_window)
                ReDim ac(nsampFFT)
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
                    For i As Long = 0 To nsamp_window - 1 Step 1
                        window(i) = 0.5 - 0.5 * Math.Cos(i * 2 * Math.PI / (nsamp_window + 1))
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
                For i As Long = 2 To nsampFFT - 1 Step 2
                    windowR(i) = windowR(i) * windowR(i) + windowR(i + 1) * windowR(i + 1)
                    '// power spectrum: square and zero
                    windowR(i + 1) = 0.0
                Next
                '   // Nyquist frequency
                windowR(nsampFFT) *= windowR(nsampFFT)
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
            Dim imax(maxnCandidates) As Long
            'autoNUMvector <double> localMean (1, mme.ny)
            Dim localMean(mme.ny) As Double

            'autoMelderProgress progress (L"Sound to Pitch...");

            For iframe = 0 To nFrames - 1 Step 1
                Dim pitchFrame As Pitch_Frame = thee.frame(iframe)
                Dim t As Double = Sampled_indexToX(thee, iframe), localPeak
                Dim leftSample As Long = Sampled_xToLowIndex(mme, t), rightSample = leftSample + 1
                Dim startSample, endSample As Long
                'Melder_progress (0.1 + (0.8 * iframe) / (nFrames + 1),
                '	L"Sound to Pitch: analysis of frame ", Melder_integer (iframe), L" out of ", Melder_integer (nFrames));

                For channel As Long = 0 To mme.ny - 1 Step 1
                    '/*
                    ' * Compute the local mean; look one longest period to both sides.
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
                ' * Compute the local peak; look half a longest period to both sides.
                ' */
                localPeak = 0.0
                If ((startSample = halfnsamp_window + 1 - halfnsamp_period) < 1) Then
                    startSample = 0
                End If
                If ((endSample = halfnsamp_window + halfnsamp_period) > nsamp_window) Then
                    endSample = nsamp_window - 1
                End If
                For channel As Long = 0 To mme.ny - 1 Step 1
                    For j As Long = startSample To endSample Step 1
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
                        'double *amp = my z [channel] + offset;
                        Dim ampIndex As Double = channel + offset
                        For i As Long = 0 To nsamp_window Step 1
                            Dim x As Double = mme.z(ampIndex + i, 0) - localMean(channel)
                            sumx2 += x * x
                        Next
                    Next
                    '/* At zero lag, these are still equal. */
                    Dim sumy2 As Double = sumx2
                    r(0) = 1.0
                    For i As Long = 1 To localMaximumLag Step 1
                        Dim product As Double = 0.0
                        For channel As Long = 0 To mme.ny Step 1
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
                        r(-i) = r(i) = product / Math.Sqrt(sumx2 * sumy2)
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
                        Dim chd() As Double = New Double(secondRes) {}
                        Dim ttt As Integer = 0
                        While ttt < secondRes
                            chd(ttt) = frame(channel, ttt)
                            ttt += 1
                        End While

                        NUMfft_forward(fftTable, chd)
                        ' /* DC component. */
                        ac(0) += frame(channel, 0) * frame(channel, 0)
                        For i As Long = 1 To nsampFFT - 2 Step 2
                            '/* Power spectrum. */
                            ac(i) += frame(channel, i) * frame(channel, i) + frame(channel, i + 1) * frame(channel, i + 1)
                        Next
                        ' /* Nyquist frequency. */
                        ac(nsampFFT) += frame(channel, nsampFFT) * frame(channel, nsampFFT)
                    Next
                    '/* Autocorrelation. */
                    NUMfft_backward(fftTable, ac)

                    '/*
                    ' * Normalize the autocorrelation to the value with zero lag,
                    ' * and divide it by the normalized autocorrelation of the window.
                    ' */
                    r(0) = 1.0
                    For i As Long = 0 To brent_ixmax - 1 Step 1
                        r(-i) = r(i) = ac(i + 1) / (ac(0) * windowR(i + 1))
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
                For i As Long = 2 To maximumLag - 1 And i < brent_ixmax Step 1
                    '/* Not too unvoiced? */
                    ' /* Maximum? */
                    If (r(i) > 0.5 * voicingThreshold And r(i) > r(i - 1) And r(i) >= r(i + 1)) Then
                        Dim place As Integer = 0

                        '/*
                        ' * Use parabolic interpolation for first estimate of frequency,
                        ' * and sin(x)/x interpolation to compute the strength of this frequency.
                        ' */
                        Dim dr As Double = 0.5 * (r(i + 1) - r(i - 1)), d2r = 2 * r(i) - r(i - 1) - r(i + 1)
                        Dim frequencyOfMaximum As Double = 1 / mme.dx / (i + dr / d2r)
                        Dim offset As Long = -brent_ixmax - 1
                        '/* method & 1 ? */
                        Dim strengthOfMaximum As Double = NUM_interpolate_sinc(mme.z, r(offset), brent_ixmax - offset, 1 / mme.dx / frequencyOfMaximum - offset, 30)
                        '/* : r [i] + 0.5 * dr * dr / d2r */;
                        '	/* High values due to short windows are to be reflected around 1. */
                        If (strengthOfMaximum > 1.0) Then
                            strengthOfMaximum = 1.0 / strengthOfMaximum
                        End If

                        '/*
                        ' * Find a place for this maximum.
                        ' */
                        If (pitchFrame.nCandidates < thee.maxCandidates) Then
                            '/* Is there still a free place? */
                            place = ++pitchFrame.nCandidates
                        Else
                            '/* Try the place of the weakest candidate so far. */
                            Dim weakest As Double = 2
                            For iweak As Integer = 2 To thee.maxCandidates Step 1
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
                        If (place) Then
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
                For i As Long = 2 To pitchFrame.nCandidates Step 1
                    If (method <> AC_HANNING Or pitchFrame.candidates(i).frequency > 0.0 / mme.dx) Then
                        Dim xmid, ymid As Double
                        Dim offset As Long = -brent_ixmax - 1
                        Dim vv As Double
                        If pitchFrame.candidates(i).frequency > 0.3 / mme.dx Then
                            vv = NUM_PEAK_INTERPOLATE_SINC700
                        Else
                            vv = brent_depth
                        End If
                        ymid = NUMimproveMaximum(mme.z, r(offset), brent_ixmax - offset, imax(i) - offset, vv, xmid)
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
            '// progress (0.95, L"Sound to Pitch: path finder");
            'Melder_progress (0.95, L"Sound to Pitch: path finder");   
            Pitch_pathFinder(thee, silenceThreshold, voicingThreshold, octaveCost, octaveJumpCost, voicedUnvoicedCost, ceiling, False)

            Return thee
        Catch ex As Exception
            'Melder_throw (me, ": pitch analysis not performed.");
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
        Return (x - mme.x1) / mme.dx + 1
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
        If (x < mme.xmin Or x > mme.xmax) Then
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
            If (inear < 1 Or inear > mme.nx) Then
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

    Function Sound_findMaximumCorrelation(ByRef mme As Sound, ByVal t1 As Double, ByVal windowLength As Double, ByVal tmin2 As Double, ByVal tmax2 As Double, ByVal tout As Double, ByVal peak As Double) As Double
        Dim maximumCorrelation = -1.0, r1 = 0.0, r2 = 0.0, r3 = 0.0, r1_best, r3_best, ir As Double
        Dim halfWindowLength As Double = 0.5 * windowLength
        Dim ileft1 As Long = Sampled_xToNearestIndex(mme, t1 - halfWindowLength)
        Dim iright1 As Long = Sampled_xToNearestIndex(mme, t1 + halfWindowLength)
        Dim ileft2min As Long = Sampled_xToLowIndex(mme, tmin2 - halfWindowLength)
        Dim ileft2max As Long = Sampled_xToHighIndex(mme, tmax2 - halfWindowLength)
        Dim i2 As Long
        '   /* Default. */
        peak = 0.0
        For ileft2 As Long = ileft2min To ileft2max Step 1
            Dim norm1 As Double = 0.0, norm2 = 0.0, product = 0.0, localPeak = 0.0
            If (mme.ny = 1) Then
                i2 = ileft2
                For i1 As Long = ileft1 To iright1 Step 1
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
            tout = t1 + (ir - ileft1) * mme.dx
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
            End If
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
        minimum = maximum = mme.z(channel1, 0)
        For i = 2 To n Step 1
            Dim value As Double = mme.z(channel1 + i, 0)
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
        Dim valueMid As Double = mme.z(channel1 + iextr, 0)
        Dim valueLeft As Double = mme.z(channel1 + iextr - 1, 0)
        Dim valueRight As Double = mme.z(channel1 + iextr + 1, 0)
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
        If (True) Then
            Return mme.x1 + (imin - 1 + iextremum - 1) * mme.dx
        Else
            Return (tmin + tmax) / 2
        End If
    End Function

    Function Pitch_isVoiced_i(ByRef mme As Pitch, ByVal iframe As Long) As Boolean
        Return Sampled_getValueAtSample(mme, iframe, Pitch_LEVEL_FREQUENCY, 0) <> NUMundefined
    End Function

    Function Pitch_getVoicedIntervalAfter(ByRef mme As Pitch, ByVal after As Double, ByVal tleft As Double, ByVal tright As Double) As Integer
        Dim ileft As Long = Sampled_xToHighIndex(mme, after), iright
        '/* Offright. */
        If (ileft > mme.nx - 1) Then
            Return 0
        End If
        '/* Offleft. */
        If (ileft < 0) Then
            ileft = 1
        End If


        '/* Search for first voiced frame. */
        For ileft = ileft To mme.nx Step 1
            If (Pitch_isVoiced_i(mme, ileft)) Then
                GoTo exitfor
            End If
        Next
        ' /* Offright. */
exitfor: If (ileft > mme.nx) Then
            Return 0
        End If

        '/* Search for last voiced frame. */
        For iright = ileft To mme.nx - 1 Step 1
            If (Not Pitch_isVoiced_i(mme, iright)) Then
                GoTo exitfor1
            End If
        Next
        iright -= 1

        '/* The whole frame is considered voiced. */
exitfor1: tleft = Sampled_indexToX(mme, ileft) - 0.5 * mme.dx
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

    Function improve_evaluate(ByVal x As Double, ByRef z(,) As Double, ByVal channel As Long, ByVal depth As Long, ByVal nx As Long, ByVal isMaximum As Integer) As Double
        'struct improve_params *me = (struct improve_params *) closure;
        Dim y As Double = NUM_interpolate_sinc(z, channel, nx, x, depth)
        If isMaximum Then
            Return -y
        Else
            Return y
        End If
    End Function

    Function NUMminimize_brent(ByVal a As Double, ByVal b As Double, ByRef z(,) As Double, ByVal channel As Long, ByVal depth As Long, ByVal nx As Long, ByVal isMaximum As Integer, ByVal tol As Double, ByVal fx As Double) As Double
        Dim x, v, fv, w, fw As Double
        Dim golden As Double = 1 - NUM_goldenSection
        Dim sqrt_epsilon As Double = Math.Sqrt(Double.Epsilon) '(NUMfpp -> eps)
        Dim itermax As Long = 60

        '!!!Melder_assert (tol > 0 && a < b);
        If Not (tol > 0 And a < b) Then
            Console.WriteLine("Cannot continue process: (tol > 0 And a < b) ")
            Return NUMundefined
        End If

        '/* First step - golden section */

        v = a + golden * (b - a)
        fv = improve_evaluate(v, z, channel, depth, nx, isMaximum)
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
                '	Interpolation step is calculated as p/q;
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
            ft = fv = improve_evaluate(t, z, channel, depth, nx, isMaximum)

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
        'Melder_warning (L"NUMminimize_brent: maximum number of iterations (", Melder_integer (itermax), L") exceeded.");
        Return x
    End Function

    Function NUMimproveExtremum(ByRef z(,) As Double, ByVal channel As Long, ByVal nx As Long, ByVal ixmid As Long, ByVal interpolation As Integer, ByVal ixmid_real As Double, ByVal isMaximum As Integer) As Double
        '!!! struct improve_params params;
        Dim result As Double
        If (ixmid <= 1) Then
            ixmid_real = 1
            Return z(channel, 0)
        End If

        If (ixmid >= nx) Then
            ixmid_real = nx
            Return z(channel, nx - 1)
        End If
        If (interpolation <= NUM_PEAK_INTERPOLATE_NONE) Then
            ixmid_real = ixmid
            Return z(channel, ixmid)
        End If
        If (interpolation = NUM_PEAK_INTERPOLATE_PARABOLIC) Then
            Dim dy As Double = 0.5 * (z(channel, ixmid + 1) - z(channel, ixmid - 1))
            Dim d2y As Double = 2 * z(channel, ixmid) - z(channel, ixmid - 1) - z(channel, ixmid + 1)
            ixmid_real = ixmid + dy / d2y
            Return z(channel, ixmid) + 0.5 * dy * dy / d2y
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
        ixmid_real = NUMminimize_brent(ixmid - 1, ixmid + 1, z, channel, depth, nx, isMaximum, 0.0000000001, result)
        If isMaximum Then
            Return -result
        Else
            Return result
        End If
    End Function

    Function NUMimproveMaximum(ByRef z(,) As Double, ByVal channel As Long, ByVal nx As Long, ByVal ixmid As Long, ByVal interpolation As Integer, ByVal ixmid_real As Double) As Double
        Return NUMimproveExtremum(z, channel, nx, ixmid, interpolation, ixmid_real, 1)
    End Function
    Function NUMimproveMinimum(ByRef z(,) As Double, ByVal channel As Long, ByVal nx As Long, ByVal ixmid As Long, ByVal interpolation As Integer, ByVal ixmid_real As Double) As Double
        Return NUMimproveExtremum(z, channel, nx, ixmid, interpolation, ixmid_real, 0)
    End Function
    Function Sampled_getWindowSamples(ByRef mme As Sampled, ByVal xmin As Double, ByVal xmax As Double, ByRef ixmin As Long, ByRef ixmax As Long) As Long
        Dim rixmin As Double = 1.0 + Math.Ceiling((xmin - mme.x1) / mme.dx)
        Dim rixmax As Double = 1.0 + Math.Floor((xmax - mme.x1) / mme.dx)
        'ixmin = rixmin < 1.0 ? 1 : (long) rixmin
        If rixmin < 1.0 Then
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

    Function NUM_interpolate_sinc(ByRef z(,) As Double, ByVal channel As Long, ByVal nx As Long, ByVal x As Double, ByVal maxDepth As Long) As Double
        Dim ix, midleft As Long, midright = midleft + 1, left, right
        midleft = Math.Floor(x)
        Dim result As Double = 0.0, a, halfsina, aa, daa

        If (nx < 0) Then
            Return NUMundefined
        End If
        If (x > nx - 1) Then
            Return z(channel, nx)
        End If
        If (x < 0) Then
            Return z(channel, 0)
        End If
        If (x = midleft) Then
            Return z(channel, midleft)
        End If
        If (maxDepth > midright - 1) Then
            maxDepth = midright - 1
        End If
        If (maxDepth > nx - midleft) Then
            maxDepth = nx - midleft
        End If
        If (maxDepth <= NUM_VALUE_INTERPOLATE_NEAREST) Then
            Return z(channel, Convert.ToUInt32(Math.Floor(x + 0.5)))
        End If
        If (maxDepth = NUM_VALUE_INTERPOLATE_LINEAR) Then
            Return z(channel, midleft) + (x - midleft) * (z(channel, midright) - z(channel, midleft))
        End If
        If (maxDepth = NUM_VALUE_INTERPOLATE_CUBIC) Then
            Dim yl As Double = z(channel, midleft), yr = z(channel, midright)
            Dim dyl As Double = 0.5 * (yr - z(channel, midleft - 1)), dyr = 0.5 * (z(channel, midright + 1) - yl)
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
            result += z(channel, ix) * d
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
            result += z(channel, ix) * d
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
            Return NUM_interpolate_sinc(mme.z, ilevel, mme.nx, Sampled_xToIndex(mme, x), interp)
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
            sum += NUM_interpolate_sinc(mme.z, channel, mme.nx, Sampled_xToIndex(mme, x), interp)
        Next
        Return sum / mme.ny
    End Function
    Sub Vector_getMaximumAndX(ByRef mme As Vector, xmin As Double, xmax As Double, channel As Long, interpolation As Integer, ByVal return_maximum As Double, return_xOfMaximum As Double)
        Dim n As Long = mme.nx, imin = 0, imax = 0, i
        '!!!Melder_assert (channel >= 1 && channel <= my ny);
        'double *y = my z [channel]
        Dim maximum, x As Double
        If (xmax <= xmin) Then
            xmin = mme.xmin
            xmax = mme.xmax
        End If
        If (Not Sampled_getWindowSamples(mme, xmin, xmax, imin, imax)) Then
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
            For i = imin To imax - 1 Step 1
                If (mme.z(channel, i) > mme.z(channel, i - 1) And mme.z(channel, i) >= mme.z(channel, i + 1)) Then
                    Dim i_real As Double = 0
                    Dim localMaximum As Double = NUMimproveMaximum(mme.z, channel, n, i, interpolation, i_real)
                    If (localMaximum > maximum) Then
                        maximum = localMaximum
                        x = i_real
                    End If
                End If

            Next
            '/* Convert sample to x. */
            x = mme.x1 + (x - 1) * mme.dx
            If (x < xmin) Then
                x = xmin
            ElseIf (x > xmax) Then
                x = xmax
            End If
        End If
        If (return_maximum) Then
            return_maximum = maximum
        End If
        'If (return_xOfMaximum) Then
        'return_xOfMaximum = x
        'End If
    End Sub

    Sub Vector_getMinimumAndX(ByRef mme As Vector, xmin As Double, xmax As Double, channel As Long, interpolation As Integer, ByVal return_minimum As Double, return_xOfMinimum As Double)
        Dim n As Long = mme.nx, imin = 0, imax = 0
        '!!!Melder_assert (channel >= 1 && channel <= my ny);
        'double *y = my z [channel]
        Dim minimum, x As Double
        If (xmax <= xmin) Then
            xmin = mme.xmin
            xmax = mme.xmax
        End If
        If (Not Sampled_getWindowSamples(mme, xmin, xmax, imin, imax)) Then
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
            For i = imin To imax - 1 Step 1
                If (mme.z(channel, i) < mme.z(channel, i - 1) And mme.z(channel, i) <= mme.z(channel, i + 1)) Then
                    Dim i_real As Double = 0
                    Dim localminimum As Double = NUMimproveMinimum(mme.z, channel, n, i, interpolation, i_real)
                    If (localminimum < minimum) Then
                        minimum = localminimum
                        x = i_real
                    End If
                End If

            Next
            '/* Convert sample to x. */
            x = mme.x1 + (x - 1) * mme.dx
            If (x < xmin) Then
                x = xmin
            ElseIf (x > xmax) Then
                x = xmax
            End If
        End If
        If (return_minimum) Then
            return_minimum = minimum
        End If
        'If (return_xOfminimum) Then
        'return_xOfminimum = x
        'End If
    End Sub

    Sub Vector_getMaximumAndXAndChannel(ByRef mme As Vector, xmin As Double, xmax As Double, interpolation As Integer, return_maximum As Double)
        Dim maximum, xOfMaximum As Double
        Dim channelOfMaximum As Long = 1
        Vector_getMaximumAndX(mme, xmin, xmax, 1, interpolation, maximum, xOfMaximum)
        For channel As Long = 2 To mme.ny Step 1
            Dim maximumOfChannel, xOfMaximumOfChannel As Double
            Vector_getMaximumAndX(mme, xmin, xmax, channel, interpolation, maximumOfChannel, xOfMaximumOfChannel)
            If (maximumOfChannel > maximum) Then
                maximum = maximumOfChannel
                xOfMaximum = xOfMaximumOfChannel
                channelOfMaximum = channel
            End If
        Next
        If (return_maximum) Then
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
    Sub Vector_getMinimumAndXAndChannel(ByRef mme As Vector, xmin As Double, xmax As Double, interpolation As Integer, ByVal return_minimum As Double)
        Dim minimum, xOfMinimum As Double
        Dim channelOfMinimum As Long = 1
        Vector_getMinimumAndX(mme, xmin, xmax, 1, interpolation, minimum, xOfMinimum)
        For channel As Long = 2 To mme.ny Step 1
            Dim minimumOfChannel, xOfMinimumOfChannel As Double
            Vector_getMinimumAndX(mme, xmin, xmax, channel, interpolation, minimumOfChannel, xOfMinimumOfChannel)
            If (minimumOfChannel < minimum) Then
                minimum = minimumOfChannel
                xOfMinimum = xOfMinimumOfChannel
                channelOfMinimum = channel
            End If
        Next
        If (return_minimum) Then
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
            'autoMelderProgress progress (L"Sound & Pitch: To PointProcess...");
            While 1 = 1
                Dim tleft, tright As Double
                If (Not Pitch_getVoicedIntervalAfter(pitch, t, tleft, tright)) Then
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
        Return mme.x1 + (i - 1) * mme.dx
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
            For i As Long = 1 To mme.nx - 1 Step 1
                mme.z(channel, i) -= mean
            Next
        Next
    End Sub
    Sub Main()
        Dim s As Sound = Sound_readFromSoundFile("C:\Users\fae1987\Documents\LOL\PraatSound\melanie.wav")
        Vector_subtractMean(s)
        Dim timestep As Double = 0.01
        Dim minimumPitch As Double = 75
        Dim maximumPitch As Double = 600
        Dim pit As Pitch = Sound_to_Pitch(s, timestep, minimumPitch, maximumPitch)

        Dim pulses As PointProcess = Sound_Pitch_to_PointProcess_cc(s, pit)
        Dim pitch As PitchTier = Pitch_to_PitchTier(pit)
        Dim duration As DurationTier = New DurationTier(pit.xmin, pit.xmax)
        Dim sProcessed As Sound = Sound_Point_Pitch_Duration_to_Sound(s, pulses, pitch, duration, MAX_T)
    End Sub



End Module
