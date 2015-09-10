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

    Public Const NUMundefined = System.Double.PositiveInfinity

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

    Function Sound_to_Pitch_any(ByRef mme As Sound,
    ByVal dt As Double, ByVal minimumPitch As Double, ByVal periodsPerWindow As Double, ByVal maxnCandidates As Integer,
    ByVal method As Integer,
    ByVal silenceThreshold As Double, ByVal voicingThreshold As Double,
    ByVal octaveCost As Double, ByVal octaveJumpCost As Double, ByVal voicedUnvoicedCost As Double, ByVal ceiling As Double) As Pitch
        'TODO
        Return Nothing
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
        'TODO
        Return 0
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


    Sub Vector_getMaximumAndXAndChannel(ByRef mme As Vector, xmin As Double, xmax As Double, interpolation As Integer, ByVal return_maximum As Double)
        'TODO
    End Sub
    Function Vector_getMaximum(ByRef mme As Vector, xmin As Double, xmax As Double, interpolation As Integer) As Double
        Dim maximum As Double
        Vector_getMaximumAndXAndChannel(mme, xmin, xmax, interpolation, maximum)
        Return maximum
    End Function
    Sub Vector_getMinimumAndXAndChannel(ByRef mme As Vector, xmin As Double, xmax As Double, interpolation As Integer, ByVal return_minimum As Double)
        'TODO
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
        'TODO
    End Sub
    Sub Main()
        Dim s As Sound = Sound_readFromSoundFile("C:\Users\fae1987\Documents\LOL\PraatSound\melanie.wav")
        Vector_subtractMean(s)
        Dim timestep As Double
        Dim minimumPitch As Double
        Dim maximumPitch As Double
        Dim pit As Pitch = Sound_to_Pitch(s, timestep, minimumPitch, maximumPitch)

        Dim pulses As PointProcess = Sound_Pitch_to_PointProcess_cc(s, pit)
        Dim pitch As PitchTier = Pitch_to_PitchTier(pit)
    End Sub



End Module
