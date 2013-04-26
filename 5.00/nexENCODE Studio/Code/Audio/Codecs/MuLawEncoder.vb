'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Imports System.Collections.Generic
Imports System.Text
Imports NAudio.Wave

Namespace NAudio.Codecs
    Public NotInheritable Class MuLawEncoder
        Private Sub New()
        End Sub

        Private Const cBias As Integer = &H84
        Private Const cClip As Integer = 32635
        Private Shared ReadOnly MuLawCompressTable As Byte() = New Byte(255) {0, 0, 1, 1, 2, 2,
         2, 2, 3, 3, 3, 3, _
         3, 3, 3, 3, 4, 4, _
         4, 4, 4, 4, 4, 4, _
         4, 4, 4, 4, 4, 4, _
         4, 4, 5, 5, 5, 5, _
         5, 5, 5, 5, 5, 5, _
         5, 5, 5, 5, 5, 5, _
         5, 5, 5, 5, 5, 5, _
         5, 5, 5, 5, 5, 5, _
         5, 5, 5, 5, 6, 6, _
         6, 6, 6, 6, 6, 6, _
         6, 6, 6, 6, 6, 6, _
         6, 6, 6, 6, 6, 6, _
         6, 6, 6, 6, 6, 6, _
         6, 6, 6, 6, 6, 6, _
         6, 6, 6, 6, 6, 6, _
         6, 6, 6, 6, 6, 6, _
         6, 6, 6, 6, 6, 6, _
         6, 6, 6, 6, 6, 6, _
         6, 6, 6, 6, 6, 6, _
         6, 6, 7, 7, 7, 7, _
         7, 7, 7, 7, 7, 7, _
         7, 7, 7, 7, 7, 7, _
         7, 7, 7, 7, 7, 7, _
         7, 7, 7, 7, 7, 7, _
         7, 7, 7, 7, 7, 7, _
         7, 7, 7, 7, 7, 7, _
         7, 7, 7, 7, 7, 7, _
         7, 7, 7, 7, 7, 7, _
         7, 7, 7, 7, 7, 7, _
         7, 7, 7, 7, 7, 7, _
         7, 7, 7, 7, 7, 7, _
         7, 7, 7, 7, 7, 7, _
         7, 7, 7, 7, 7, 7, _
         7, 7, 7, 7, 7, 7, _
         7, 7, 7, 7, 7, 7, _
         7, 7, 7, 7, 7, 7, _
         7, 7, 7, 7, 7, 7, _
         7, 7, 7, 7, 7, 7, _
         7, 7, 7, 7, 7, 7, _
         7, 7, 7, 7, 7, 7, _
         7, 7, 7, 7}

        Public Shared Function LinearToMuLawSample(sample As Short) As Byte
            Dim sign As Integer = (sample >> 8) And &H80
            If sign <> 0 Then
                sample = CShort(-sample)
            End If
            If sample > cClip Then
                sample = cClip
            End If
            sample = CShort(sample + cBias)
            Dim exponent As Integer = CInt(MuLawCompressTable((sample >> 7) And &HFF))
            Dim mantissa As Integer = (sample >> (exponent + 3)) And &HF
            Dim compressedByte As Integer = Not (sign Or (exponent << 4) Or mantissa)
            Return CByte(compressedByte)
        End Function
    End Class
End Namespace