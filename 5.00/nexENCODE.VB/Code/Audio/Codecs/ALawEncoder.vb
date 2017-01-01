'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Imports System.Collections.Generic
Imports System.Text

Namespace NAudio.Codecs
    Public NotInheritable Class ALawEncoder
        Private Sub New()
        End Sub

        Private Const cBias As Integer = &H84
        Private Const cClip As Integer = 32635
        Private Shared ReadOnly ALawCompressTable As Byte() = New Byte(127) {1, 1, 2, 2, 3, 3, _
         3, 3, 4, 4, 4, 4, _
         4, 4, 4, 4, 5, 5, _
         5, 5, 5, 5, 5, 5, _
         5, 5, 5, 5, 5, 5, _
         5, 5, 6, 6, 6, 6, _
         6, 6, 6, 6, 6, 6, _
         6, 6, 6, 6, 6, 6, _
         6, 6, 6, 6, 6, 6, _
         6, 6, 6, 6, 6, 6, _
         6, 6, 6, 6, 7, 7, _
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
         7, 7}

        Public Shared Function LinearToALawSample(sample As Short) As Byte
            Dim sign As Integer
            Dim exponent As Integer
            Dim mantissa As Integer
            Dim compressedByte As Byte
            sign = ((Not sample) >> 8) And &H80
            If sign = 0 Then
                sample = CShort(-sample)
            End If
            If sample > cClip Then
                sample = cClip
            End If
            If sample >= 256 Then
                exponent = CInt(ALawCompressTable((sample >> 8) And &H7F))
                mantissa = (sample >> (exponent + 3)) And &HF
                compressedByte = CByte((exponent << 4) Or mantissa)
            Else
                compressedByte = CByte(sample >> 4)
            End If
            compressedByte = compressedByte Xor CByte(sign Xor &H55)
            Return compressedByte
        End Function
    End Class
End Namespace