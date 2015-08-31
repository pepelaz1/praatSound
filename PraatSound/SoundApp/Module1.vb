Module Module1
    Function RealTier_getArea(ByVal mme As DurationTier, ByVal tmin As Double, ByVal tmax As Double) As Double
        Return 0
    End Function

    Function RealTier_getValueAtTime(ByVal mme As RealTier, ByVal t As Double) As Double
        Return 0
    End Function

    Sub copyFall(ByVal mme As Sound, ByVal tmin As Double, ByVal tmax As Double, ByVal thee As Sound, ByVal tminTarget As Double)
    End Sub

    Sub copyRise (ByVal mme As Sound, ByVal tmin As Double,  ByVal tmax As Double, ByVal thee As Sound, ByVal tminTarget As Double)
    End Sub

    Sub copyBell(ByVal mme As Sound, ByVal tmid As Double, ByVal leftWidth As Double, ByVal rightWidth As Double, ByVal thee As Sound, ByVal tmidTarget As Double)
        copyRise(mme, tmid - leftWidth, tmid, thee, tmidTarget)
        copyFall(mme, tmid, tmid + rightWidth, thee, tmidTarget)
    End Sub

    Sub copyBell2(ByVal mme As Sound, ByVal source As PointProcess, ByVal isource As Double, ByVal leftWidth As Double,
                   ByVal rightWidth As Double, ByVal thee As Sound, ByVal tmidTarget As Double, ByVal maxT As Double)
    End Sub

    Function PointProcess_getNearestIndex(ByVal mme As PointProcess, ByVal t As Double) As Long
        Return 0
    End Function

    Function Sampled_xToLowIndex (ByVal mme As Sampled ,ByVal x As Double) As Long
        Return 0
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

            If (duration.points.size = 0) Then
                ' Melder_throw("No duration points.")
            End If


            '/*
            ' * Create a Sound long enough to hold the longest possible duration-manipulated sound.
            ' */
            Dim thee As New Sound(1, mme, xmin, mme.xmin + 3 * (mme.xmax - mme.xmin), 3 * mme.nx, mme.dx, mme.x1)

            '/*
            ' * Below, I'll abbreviate the voiced interval as "voice" and the voiceless interval as "noise".
            '*/
            If (pitch And pitch.points.size) Then
                For ipointleft As Integer = 1 To pulses.nt
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
