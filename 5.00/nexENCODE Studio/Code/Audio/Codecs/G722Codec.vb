'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Imports System.Collections.Generic
Imports System.Text

Namespace NAudio.Codecs
    Public Class G722Codec
        Private Shared Function Saturate(_Amp As Integer) As Short
            Dim _Amp16 As Short = CShort(_Amp)
            If _Amp = _Amp16 Then
                Return _Amp16
            End If
            If _Amp > Int16.MaxValue Then
                Return Int16.MaxValue
            End If
            Return Int16.MinValue
        End Function

        Private Shared Sub Block4(s As G722CodecState, band As Integer, d As Integer)
            Dim wd1 As Integer
            Dim wd2 As Integer
            Dim wd3 As Integer
            Dim i As Integer
            s.Band(band).d(0) = d
            s.Band(band).r(0) = Saturate(s.Band(band).s + d)
            s.Band(band).p(0) = Saturate(s.Band(band).sz + d)
            For i = 0 To 2
                s.Band(band).sg(i) = s.Band(band).p(i) >> 15
            Next
            wd1 = Saturate(s.Band(band).a(1) << 2)
            wd2 = If((s.Band(band).sg(0) = s.Band(band).sg(1)), -wd1, wd1)
            If wd2 > 32767 Then
                wd2 = 32767
            End If
            wd3 = If((s.Band(band).sg(0) = s.Band(band).sg(2)), 128, -128)
            wd3 += (wd2 >> 7)
            wd3 += (s.Band(band).a(2) * 32512) >> 15
            If wd3 > 12288 Then
                wd3 = 12288
            ElseIf wd3 < -12288 Then
                wd3 = -12288
            End If
            s.Band(band).ap(2) = wd3
            s.Band(band).sg(0) = s.Band(band).p(0) >> 15
            s.Band(band).sg(1) = s.Band(band).p(1) >> 15
            wd1 = If((s.Band(band).sg(0) = s.Band(band).sg(1)), 192, -192)
            wd2 = (s.Band(band).a(1) * 32640) >> 15
            s.Band(band).ap(1) = Saturate(wd1 + wd2)
            wd3 = Saturate(15360 - s.Band(band).ap(2))
            If s.Band(band).ap(1) > wd3 Then
                s.Band(band).ap(1) = wd3
            ElseIf s.Band(band).ap(1) < -wd3 Then
                s.Band(band).ap(1) = -wd3
            End If
            wd1 = If((d = 0), 0, 128)
            s.Band(band).sg(0) = d >> 15
            For i = 1 To 6
                s.Band(band).sg(i) = s.Band(band).d(i) >> 15
                wd2 = If((s.Band(band).sg(i) = s.Band(band).sg(0)), wd1, -wd1)
                wd3 = (s.Band(band).b(i) * 32640) >> 15
                s.Band(band).bp(i) = Saturate(wd2 + wd3)
            Next
            For i = 6 To 1 Step -1
                s.Band(band).d(i) = s.Band(band).d(i - 1)
                s.Band(band).b(i) = s.Band(band).bp(i)
            Next
            For i = 2 To 1 Step -1
                s.Band(band).r(i) = s.Band(band).r(i - 1)
                s.Band(band).p(i) = s.Band(band).p(i - 1)
                s.Band(band).a(i) = s.Band(band).ap(i)
            Next
            wd1 = Saturate(s.Band(band).r(1) + s.Band(band).r(1))
            wd1 = (s.Band(band).a(1) * wd1) >> 15
            wd2 = Saturate(s.Band(band).r(2) + s.Band(band).r(2))
            wd2 = (s.Band(band).a(2) * wd2) >> 15
            s.Band(band).sp = Saturate(wd1 + wd2)
            s.Band(band).sz = 0
            For i = 6 To 1 Step -1
                wd1 = Saturate(s.Band(band).d(i) + s.Band(band).d(i))
                s.Band(band).sz += (s.Band(band).b(i) * wd1) >> 15
            Next
            s.Band(band).sz = Saturate(s.Band(band).sz)
            s.Band(band).s = Saturate(s.Band(band).sp + s.Band(band).sz)
        End Sub

        Shared ReadOnly wl As Integer() = {-60, -30, 58, 172, 334, 538, _
         1198, 3042}
        Shared ReadOnly rl42 As Integer() = {0, 7, 6, 5, 4, 3, _
         2, 1, 7, 6, 5, 4, _
         3, 2, 1, 0}
        Shared ReadOnly ilb As Integer() = {2048, 2093, 2139, 2186, 2233, 2282, _
         2332, 2383, 2435, 2489, 2543, 2599, _
         2656, 2714, 2774, 2834, 2896, 2960, _
         3025, 3091, 3158, 3228, 3298, 3371, _
         3444, 3520, 3597, 3676, 3756, 3838, _
         3922, 4008}
        Shared ReadOnly wh As Integer() = {0, -214, 798}
        Shared ReadOnly rh2 As Integer() = {2, 1, 2, 1}
        Shared ReadOnly qm2 As Integer() = {-7408, -1616, 7408, 1616}
        Shared ReadOnly qm4 As Integer() = {0, -20456, -12896, -8968, -6288, -4240, _
         -2584, -1200, 20456, 12896, 8968, 6288, _
         4240, 2584, 1200, 0}
        Shared ReadOnly qm5 As Integer() = {-280, -280, -23352, -17560, -14120, -11664, _
         -9752, -8184, -6864, -5712, -4696, -3784, _
         -2960, -2208, -1520, -880, 23352, 17560, _
         14120, 11664, 9752, 8184, 6864, 5712, _
         4696, 3784, 2960, 2208, 1520, 880, _
         280, -280}
        Shared ReadOnly qm6 As Integer() = {-136, -136, -136, -136, -24808, -21904, _
         -19008, -16704, -14984, -13512, -12280, -11192, _
         -10232, -9360, -8576, -7856, -7192, -6576, _
         -6000, -5456, -4944, -4464, -4008, -3576, _
         -3168, -2776, -2400, -2032, -1688, -1360, _
         -1040, -728, 24808, 21904, 19008, 16704, _
         14984, 13512, 12280, 11192, 10232, 9360, _
         8576, 7856, 7192, 6576, 6000, 5456, _
         4944, 4464, 4008, 3576, 3168, 2776, _
         2400, 2032, 1688, 1360, 1040, 728, _
         432, 136, -432, -136}
        Shared ReadOnly qmf_coeffs As Integer() = {3, -11, 12, 32, -210, 951, _
         3876, -805, 362, -156, 53, -11}
        Shared ReadOnly q6 As Integer() = {0, 35, 72, 110, 150, 190, _
         233, 276, 323, 370, 422, 473, _
         530, 587, 650, 714, 786, 858, _
         940, 1023, 1121, 1219, 1339, 1458, _
         1612, 1765, 1980, 2195, 2557, 2919, _
         0, 0}
        Shared ReadOnly iln As Integer() = {0, 63, 62, 31, 30, 29, _
         28, 27, 26, 25, 24, 23, _
         22, 21, 20, 19, 18, 17, _
         16, 15, 14, 13, 12, 11, _
         10, 9, 8, 7, 6, 5, _
         4, 0}
        Shared ReadOnly ilp As Integer() = {0, 61, 60, 59, 58, 57, _
         56, 55, 54, 53, 52, 51, _
         50, 49, 48, 47, 46, 45, _
         44, 43, 42, 41, 40, 39, _
         38, 37, 36, 35, 34, 33, _
         32, 0}
        Shared ReadOnly ihn As Integer() = {0, 1, 0}
        Shared ReadOnly ihp As Integer() = {0, 3, 2}

        Public Function Decode(state As G722CodecState, outputBuffer As Short(), inputG722Data As Byte(), inputLength As Integer) As Integer
            Dim dlowt As Integer
            Dim rlow As Integer
            Dim ihigh As Integer
            Dim dhigh As Integer
            Dim rhigh As Integer
            Dim xout1 As Integer
            Dim xout2 As Integer
            Dim wd1 As Integer
            Dim wd2 As Integer
            Dim wd3 As Integer
            Dim code As Integer
            Dim outlen As Integer
            Dim i As Integer
            Dim j As Integer
            outlen = 0
            rhigh = 0
            j = 0
            While j < inputLength
                If state.Packed Then
                    If state.InBits < state.BitsPerSample Then
                        state.InBuffer = state.InBuffer Or CUInt(inputG722Data(System.Math.Max(System.Threading.Interlocked.Increment(j), j - 1)) << state.InBits)
                        state.InBits += 8
                    End If
                    code = CInt(state.InBuffer) And ((1 << state.BitsPerSample) - 1)
                    state.InBuffer >>= state.BitsPerSample
                    state.InBits -= state.BitsPerSample
                Else
                    code = inputG722Data(System.Math.Max(System.Threading.Interlocked.Increment(j), j - 1))
                End If
                Select Case state.BitsPerSample
                    Case 6
                        wd1 = code And &HF
                        ihigh = (code >> 4) And &H3
                        wd2 = qm4(wd1)
                        Exit Select
                    Case 7
                        wd1 = code And &H1F
                        ihigh = (code >> 5) And &H3
                        wd2 = qm5(wd1)
                        wd1 >>= 1
                        Exit Select
                    Case Else
                        wd1 = code And &H3F
                        ihigh = (code >> 6) And &H3
                        wd2 = qm6(wd1)
                        wd1 >>= 2
                        Exit Select
                End Select
                wd2 = (state.Band(0).det * wd2) >> 15
                rlow = state.Band(0).s + wd2
                If rlow > 16383 Then
                    rlow = 16383
                ElseIf rlow < -16384 Then
                    rlow = -16384
                End If
                wd2 = qm4(wd1)
                dlowt = (state.Band(0).det * wd2) >> 15
                wd2 = rl42(wd1)
                wd1 = (state.Band(0).nb * 127) >> 7
                wd1 += wl(wd2)
                If wd1 < 0 Then
                    wd1 = 0
                ElseIf wd1 > 18432 Then
                    wd1 = 18432
                End If
                state.Band(0).nb = wd1
                wd1 = (state.Band(0).nb >> 6) And 31
                wd2 = 8 - (state.Band(0).nb >> 11)
                wd3 = If((wd2 < 0), (ilb(wd1) << -wd2), (ilb(wd1) >> wd2))
                state.Band(0).det = wd3 << 2
                Block4(state, 0, dlowt)
                If Not state.EncodeFrom8000Hz Then
                    wd2 = qm2(ihigh)
                    dhigh = (state.Band(1).det * wd2) >> 15
                    rhigh = dhigh + state.Band(1).s
                    If rhigh > 16383 Then
                        rhigh = 16383
                    ElseIf rhigh < -16384 Then
                        rhigh = -16384
                    End If
                    wd2 = rh2(ihigh)
                    wd1 = (state.Band(1).nb * 127) >> 7
                    wd1 += wh(wd2)
                    If wd1 < 0 Then
                        wd1 = 0
                    ElseIf wd1 > 22528 Then
                        wd1 = 22528
                    End If
                    state.Band(1).nb = wd1
                    wd1 = (state.Band(1).nb >> 6) And 31
                    wd2 = 10 - (state.Band(1).nb >> 11)
                    wd3 = If((wd2 < 0), (ilb(wd1) << -wd2), (ilb(wd1) >> wd2))
                    state.Band(1).det = wd3 << 2
                    Block4(state, 1, dhigh)
                End If
                If state.ItuTestMode Then
                    outputBuffer(System.Math.Max(System.Threading.Interlocked.Increment(outlen), outlen - 1)) = CShort(rlow << 1)
                    outputBuffer(System.Math.Max(System.Threading.Interlocked.Increment(outlen), outlen - 1)) = CShort(rhigh << 1)
                Else
                    If state.EncodeFrom8000Hz Then
                        outputBuffer(System.Math.Max(System.Threading.Interlocked.Increment(outlen), outlen - 1)) = CShort(rlow << 1)
                    Else
                        For i = 0 To 21
                            state.QmfSignalHistory(i) = state.QmfSignalHistory(i + 2)
                        Next
                        state.QmfSignalHistory(22) = rlow + rhigh
                        state.QmfSignalHistory(23) = rlow - rhigh

                        xout1 = 0
                        xout2 = 0
                        For i = 0 To 11
                            xout2 += state.QmfSignalHistory(2 * i) * qmf_coeffs(i)
                            xout1 += state.QmfSignalHistory(2 * i + 1) * qmf_coeffs(11 - i)
                        Next
                        outputBuffer(System.Math.Max(System.Threading.Interlocked.Increment(outlen), outlen - 1)) = CShort(xout1 >> 11)
                        outputBuffer(System.Math.Max(System.Threading.Interlocked.Increment(outlen), outlen - 1)) = CShort(xout2 >> 11)
                    End If
                End If
            End While
            Return outlen
        End Function

        Public Function Encode(state As G722CodecState, outputBuffer As Byte(), inputBuffer As Short(), inputBufferCount As Integer) As Integer
            Dim dlow As Integer
            Dim dhigh As Integer
            Dim el As Integer
            Dim wd As Integer
            Dim wd1 As Integer
            Dim ril As Integer
            Dim wd2 As Integer
            Dim il4 As Integer
            Dim ih2 As Integer
            Dim wd3 As Integer
            Dim eh As Integer
            Dim mih As Integer
            Dim i As Integer
            Dim j As Integer
            Dim xlow As Integer
            Dim xhigh As Integer
            Dim g722_bytes As Integer
            Dim sumeven As Integer
            Dim sumodd As Integer
            Dim ihigh As Integer
            Dim ilow As Integer
            Dim code As Integer
            g722_bytes = 0
            xhigh = 0
            j = 0
            While j < inputBufferCount
                If state.ItuTestMode Then
                    xlow = InlineAssignHelper(xhigh, inputBuffer(System.Math.Max(System.Threading.Interlocked.Increment(j), j - 1)) >> 1)
                Else
                    If state.EncodeFrom8000Hz Then
                        xlow = inputBuffer(System.Math.Max(System.Threading.Interlocked.Increment(j), j - 1)) >> 1
                    Else
                        For i = 0 To 21
                            state.QmfSignalHistory(i) = state.QmfSignalHistory(i + 2)
                        Next
                        state.QmfSignalHistory(22) = inputBuffer(System.Math.Max(System.Threading.Interlocked.Increment(j), j - 1))
                        state.QmfSignalHistory(23) = inputBuffer(System.Math.Max(System.Threading.Interlocked.Increment(j), j - 1))
                        sumeven = 0
                        sumodd = 0
                        For i = 0 To 11
                            sumodd += state.QmfSignalHistory(2 * i) * qmf_coeffs(i)
                            sumeven += state.QmfSignalHistory(2 * i + 1) * qmf_coeffs(11 - i)
                        Next
                        xlow = (sumeven + sumodd) >> 14
                        xhigh = (sumeven - sumodd) >> 14
                    End If
                End If
                el = Saturate(xlow - state.Band(0).s)
                wd = If((el >= 0), el, -(el + 1))
                For i = 1 To 29
                    wd1 = (q6(i) * state.Band(0).det) >> 12
                    If wd < wd1 Then
                        Exit For
                    End If
                Next
                ilow = If((el < 0), iln(i), ilp(i))
                ril = ilow >> 2
                wd2 = qm4(ril)
                dlow = (state.Band(0).det * wd2) >> 15
                il4 = rl42(ril)
                wd = (state.Band(0).nb * 127) >> 7
                state.Band(0).nb = wd + wl(il4)
                If state.Band(0).nb < 0 Then
                    state.Band(0).nb = 0
                ElseIf state.Band(0).nb > 18432 Then
                    state.Band(0).nb = 18432
                End If
                wd1 = (state.Band(0).nb >> 6) And 31
                wd2 = 8 - (state.Band(0).nb >> 11)
                wd3 = If((wd2 < 0), (ilb(wd1) << -wd2), (ilb(wd1) >> wd2))
                state.Band(0).det = wd3 << 2
                Block4(state, 0, dlow)
                If state.EncodeFrom8000Hz Then
                    code = (&HC0 Or ilow) >> (8 - state.BitsPerSample)
                Else
                    eh = Saturate(xhigh - state.Band(1).s)
                    wd = If((eh >= 0), eh, -(eh + 1))
                    wd1 = (564 * state.Band(1).det) >> 12
                    mih = If((wd >= wd1), 2, 1)
                    ihigh = If((eh < 0), ihn(mih), ihp(mih))
                    wd2 = qm2(ihigh)
                    dhigh = (state.Band(1).det * wd2) >> 15
                    ih2 = rh2(ihigh)
                    wd = (state.Band(1).nb * 127) >> 7
                    state.Band(1).nb = wd + wh(ih2)
                    If state.Band(1).nb < 0 Then
                        state.Band(1).nb = 0
                    ElseIf state.Band(1).nb > 22528 Then
                        state.Band(1).nb = 22528
                    End If
                    wd1 = (state.Band(1).nb >> 6) And 31
                    wd2 = 10 - (state.Band(1).nb >> 11)
                    wd3 = If((wd2 < 0), (ilb(wd1) << -wd2), (ilb(wd1) >> wd2))
                    state.Band(1).det = wd3 << 2
                    Block4(state, 1, dhigh)
                    code = ((ihigh << 6) Or ilow) >> (8 - state.BitsPerSample)
                End If
                If state.Packed Then
                    state.OutBuffer = state.OutBuffer Or CUInt(code << state.OutBits)
                    state.OutBits += state.BitsPerSample
                    If state.OutBits >= 8 Then
                        outputBuffer(System.Math.Max(System.Threading.Interlocked.Increment(g722_bytes), g722_bytes - 1)) = CByte(state.OutBuffer And &HFF)
                        state.OutBits -= 8
                        state.OutBuffer >>= 8
                    End If
                Else
                    outputBuffer(System.Math.Max(System.Threading.Interlocked.Increment(g722_bytes), g722_bytes - 1)) = CByte(code)
                End If
            End While
            Return g722_bytes
        End Function

        Private Shared Function InlineAssignHelper(Of T)(ByRef target As T, value As T) As T
            target = value
            Return value
        End Function
    End Class

    Public Class G722CodecState
        Public Property ItuTestMode() As Boolean
            Get
                Return m_ItuTestMode
            End Get
            Set(value As Boolean)
                m_ItuTestMode = Value
            End Set
        End Property

        Private m_ItuTestMode As Boolean

        Public Property Packed() As Boolean
            Get
                Return m_Packed
            End Get
            Private Set(value As Boolean)
                m_Packed = value
            End Set
        End Property

        Private m_Packed As Boolean

        Public Property EncodeFrom8000Hz() As Boolean
            Get
                Return m_EncodeFrom8000Hz
            End Get
            Private Set(value As Boolean)
                m_EncodeFrom8000Hz = value
            End Set
        End Property

        Private m_EncodeFrom8000Hz As Boolean

        Public Property BitsPerSample() As Integer
            Get
                Return m_BitsPerSample
            End Get
            Private Set(value As Integer)
                m_BitsPerSample = value
            End Set
        End Property

        Private m_BitsPerSample As Integer

        Public Property QmfSignalHistory() As Integer()
            Get
                Return m_QmfSignalHistory
            End Get
            Private Set(value As Integer())
                m_QmfSignalHistory = value
            End Set
        End Property

        Private m_QmfSignalHistory As Integer()

        Public Property Band() As Band()
            Get
                Return m_Band
            End Get
            Private Set(value As Band())
                m_Band = value
            End Set
        End Property

        Private m_Band As Band()

        Public Property InBuffer() As UInteger
            Get
                Return m_InBuffer
            End Get
            Friend Set(value As UInteger)
                m_InBuffer = value
            End Set
        End Property

        Private m_InBuffer As UInteger

        Public Property InBits() As Integer
            Get
                Return m_InBits
            End Get
            Friend Set(value As Integer)
                m_InBits = value
            End Set
        End Property

        Private m_InBits As Integer

        Public Property OutBuffer() As UInteger
            Get
                Return m_OutBuffer
            End Get
            Friend Set(value As UInteger)
                m_OutBuffer = value
            End Set
        End Property

        Private m_OutBuffer As UInteger

        Public Property OutBits() As Integer
            Get
                Return m_OutBits
            End Get
            Friend Set(value As Integer)
                m_OutBits = value
            End Set
        End Property

        Private m_OutBits As Integer

        Public Sub New(rate As Integer, options As G722Flags)
            Me.Band = New Band(1) {New Band(), New Band()}
            Me.QmfSignalHistory = New Integer(23) {}
            Me.ItuTestMode = False
            If rate = 48000 Then
                Me.BitsPerSample = 6
            ElseIf rate = 56000 Then
                Me.BitsPerSample = 7
            ElseIf rate = 64000 Then
                Me.BitsPerSample = 8
            Else
                Throw New ArgumentException("Invalid rate, should be 48000, 56000 or 64000")
            End If
            If (options And G722Flags.SampleRate8000) = G722Flags.SampleRate8000 Then
                Me.EncodeFrom8000Hz = True
            End If
            If ((options And G722Flags.Packed) = G722Flags.Packed) AndAlso Me.BitsPerSample <> 8 Then
                Me.Packed = True
            Else
                Me.Packed = False
            End If
            Me.Band(0).det = 32
            Me.Band(1).det = 8
        End Sub
    End Class

    Public Class Band
        ''' <summary>s</summary>
        Public s As Integer
        ''' <summary>sp</summary>
        Public sp As Integer
        ''' <summary>sz</summary>
        Public sz As Integer
        ''' <summary>r</summary>
        Public r As Integer() = New Integer(2) {}
        ''' <summary>a</summary>
        Public a As Integer() = New Integer(2) {}
        ''' <summary>ap</summary>
        Public ap As Integer() = New Integer(2) {}
        ''' <summary>p</summary>
        Public p As Integer() = New Integer(2) {}
        ''' <summary>d</summary>
        Public d As Integer() = New Integer(6) {}
        ''' <summary>b</summary>
        Public b As Integer() = New Integer(6) {}
        ''' <summary>bp</summary>
        Public bp As Integer() = New Integer(6) {}
        ''' <summary>sg</summary>
        Public sg As Integer() = New Integer(6) {}
        ''' <summary>nb</summary>
        Public nb As Integer
        ''' <summary>det</summary>
        Public det As Integer
    End Class

    <Flags()> Public Enum G722Flags
        None = 0
        SampleRate8000 = &H1
        Packed = &H2
    End Enum
End Namespace