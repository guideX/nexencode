'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Imports System.Collections.Generic
Imports System.Text
Imports System.Runtime.InteropServices

Namespace NAudio.CoreAudioApi.Interfaces
    <StructLayout(LayoutKind.Explicit)> Public Structure PropVariant
        <FieldOffset(0)> _
        Private vt As Short
        <FieldOffset(2)> _
        Private wReserved1 As Short
        <FieldOffset(4)> _
        Private wReserved2 As Short
        <FieldOffset(6)> _
        Private wReserved3 As Short
        <FieldOffset(8)> _
        Private cVal As SByte
        <FieldOffset(8)> _
        Private bVal As Byte
        <FieldOffset(8)> _
        Private iVal As Short
        <FieldOffset(8)> _
        Private uiVal As UShort
        <FieldOffset(8)> _
        Private lVal As Integer
        <FieldOffset(8)> _
        Private ulVal As UInteger
        <FieldOffset(8)> _
        Private intVal As Integer
        <FieldOffset(8)> _
        Private uintVal As UInteger
        <FieldOffset(8)> _
        Private hVal As Long
        <FieldOffset(8)> _
        Private uhVal As Long
        <FieldOffset(8)> _
        Private fltVal As Single
        <FieldOffset(8)> _
        Private dblVal As Double
        <FieldOffset(8)> _
        Private boolVal As Boolean
        <FieldOffset(8)> _
        Private scode As Integer
        'CY cyVal;
        <FieldOffset(8)> _
        Private [date] As DateTime
        <FieldOffset(8)> _
        Private filetime As System.Runtime.InteropServices.ComTypes.FILETIME
        <FieldOffset(8)> _
        Private blobVal As Blob
        <FieldOffset(8)> _
        Private pwszVal As IntPtr

        Private Function GetBlob() As Byte()
            Dim Result As Byte() = New Byte(blobVal.Length - 1) {}
            Marshal.Copy(blobVal.Data, Result, 0, Result.Length)
            Return Result
        End Function

        Public ReadOnly Property Value() As Object
            Get
                Dim ve As VarEnum = CType(vt, VarEnum)
                Select Case ve
                    Case VarEnum.VT_I1
                        Return bVal
                    Case VarEnum.VT_I2
                        Return iVal
                    Case VarEnum.VT_I4
                        Return lVal
                    Case VarEnum.VT_I8
                        Return hVal
                    Case VarEnum.VT_INT
                        Return iVal
                    Case VarEnum.VT_UI4
                        Return ulVal
                    Case VarEnum.VT_LPWSTR
                        Return Marshal.PtrToStringUni(pwszVal)
                    Case VarEnum.VT_BLOB
                        Return GetBlob()
                End Select
                Throw New NotImplementedException("PropVariant " & ve.ToString())
            End Get
        End Property
    End Structure
End Namespace