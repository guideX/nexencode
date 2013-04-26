'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Imports System.Runtime.InteropServices

Public Class clsAPI
    Public Event ProcessError(lError As String, lSub As String)
    Public lLastError As String

#Region "API"
#Region "DllImports"
    Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer) As Integer
    Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer) As Integer
    Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer, ByVal X3 As Integer, ByVal Y3 As Integer) As Integer
    Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Integer, ByVal hSrcRgn1 As Integer, ByVal hSrcRgn2 As Integer, ByVal nCombineMode As Integer) As Integer
    Private Declare Function AnimateWindow Lib "user32" (ByVal hwnd As Integer, ByVal dwTime As Integer, ByVal dwFlags As Integer) As Boolean
    <DllImport("user32.dll")> Private Shared Function SetWindowRgn(ByVal lhWnd As HandleRef, ByVal lRegion As Long, ByVal lRedraw As Boolean) As Long
    End Function
    <DllImport("user32.dll")> Private Shared Function GetWindowRect(lhWnd As HandleRef, ByRef lRECT As RECT) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Auto)> Private Shared Function GetClientRect(lhWnd As HandleRef, ByRef lRECT As RECT) As Integer
    End Function
#End Region
    <Flags()> Public Enum AnimateWindowFlags
        AW_HOR_POSITIVE = &H1
        AW_HOR_NEGATIVE = &H2
        AW_VER_POSITIVE = &H4
        AW_VER_NEGATIVE = &H8
        AW_CENTER = &H10
        AW_HIDE = &H10000
        AW_ACTIVATE = &H20000
        AW_SLIDE = &H40000
        AW_BLEND = &H80000
    End Enum

    Enum eRectTypes
        rRECT = 1
        rRectangle = 2
    End Enum

    Enum eCombineMode
        cRgn_None = 0
        cRgn_And = 1 'Creates the intersection of the two combined regions.
        cRgn_Or = 2 'Creates a copy of the region identified by hrgnSrc1.
        cRgn_XOr = 3 'Combines the parts of hrgnSrc1 that are not part of hrgnSrc2.
        cRgn_Diff = 4 'Creates the union of two combined regions.
        cRgn_Copy = 5 'Creates the union of two combined regions except for any overlapping areas.
    End Enum

    Enum eShapeTypes
        'sOther = 0
        sRectRgn = 1
        sEllipce = 2
        sRoundRectRgn = 3
    End Enum

    Enum eCombineRegionRet
        cNullRegion = 0
        cSimpleRegion = 1
        cComplexRegion = 2
        cError = 3
    End Enum

    <StructLayout(LayoutKind.Sequential)> Public Structure RECT
        Public Left As Integer ' x position of upper-left corner
        Public Top As Integer ' y position of upper-left corner
        Public Right As Integer ' x position of lower-right corner
        Public Bottom As Integer ' y position of lower-right corner
    End Structure

#Region "WINDOW_STYLE"
    Private Const GWL_EXSTYLE As Integer = (-20)
    Private Const WS_EX_TOPMOST As UInt32 = &H8
    Private Const WS_EX_CLIENTEDGE As UInt32 = &H200
#End Region
#End Region
    Private Structure gWindowRect
        Public wTypeSelected As eRectTypes
        Public wRECT As RECT
        Public wRectangle As Rectangle
    End Structure

    Public Structure gWindowSettings
        Public wTitleBarHeight As Integer
        Public wWindowBorder As Integer
    End Structure

    Private Function ReturnWindowRect(lForm As Form) As gWindowRect
        Try
            Dim lRECT As New gWindowRect
            lRECT.wRECT = New RECT
            lRECT.wRectangle = New Rectangle
            If Not GetWindowRect(New HandleRef(Me, lForm.Handle), lRECT.wRECT) Then
                lRECT.wRectangle.X = lRECT.wRECT.Left
                lRECT.wRectangle.Y = lRECT.wRECT.Top
                lRECT.wRectangle.Width = lRECT.wRECT.Right - lRECT.wRECT.Left + 1
                lRECT.wRectangle.Height = lRECT.wRECT.Bottom - lRECT.wRECT.Top + 1
            End If
            Return lRECT
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Private Function ReturnWindowRect(lForm As Form) As Rectangle")
            Return Nothing
        End Try
    End Function

    Private Function ReturnClientRect(lForm As Form) As gWindowRect
        Try
            Dim lRECT As New gWindowRect
            lRECT.wRECT = New RECT
            lRECT.wRectangle = New Rectangle
            If GetClientRect(New HandleRef(lForm, lForm.Handle), lRECT.wRECT) = 0 Then
                lRECT.wRectangle.X = lRECT.wRECT.Left
                lRECT.wRectangle.Y = lRECT.wRECT.Top
                lRECT.wRectangle.Width = lRECT.wRECT.Right - lRECT.wRECT.Left + 1
                lRECT.wRectangle.Height = lRECT.wRECT.Bottom - lRECT.wRECT.Top + 1
                Return lRECT
            Else
                Return Nothing
            End If
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Private Function ReturnClientRect(lForm As Form) As gWindowRect")
            Return Nothing
        End Try
    End Function

    Public ReadOnly Property WindowSettings(lForm As Form) As gWindowSettings
        Get
            Try
                Dim _WindowSettings As New gWindowSettings, lWindowPos As New gWindowRect, lClientPos As gWindowRect, lBorderWidth As Integer = CInt((lForm.Width - lForm.ClientSize.Width) / 2)
                lWindowPos = ReturnWindowRect(lForm)
                lClientPos = ReturnClientRect(lForm)
                _WindowSettings.wTitleBarHeight = lForm.Height - lForm.ClientSize.Height - 2 * lBorderWidth
                _WindowSettings.wWindowBorder = lBorderWidth
                Return _WindowSettings
            Catch ex As Exception
                RaiseEvent ProcessError(ex.Message, "Public Property WindowSettings(lForm As Form) As gWindowSettings")
                Return Nothing
            End Try
        End Get
    End Property

    Public Function SetWindowRegion(lForm As Form, lRegion As Long, lReDraw As Boolean) As Boolean
        Try
            Dim l As Long = SetWindowRgn(New HandleRef(Me, lForm.Handle), lRegion, lReDraw)
            Select Case l
                Case 0
                    RaiseEvent ProcessError("Failure: " & l.ToString, "SetWindowRegion")
                    Return False
                Case Else
                    Return True
            End Select
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Function SetWindowRegion(lForm As Form, lRegion As Long, lReDraw As Boolean) As Boolean")
            Return Nothing
        End Try
    End Function

    Public Function CombineRegion(lDestinationRegion As Integer, lSourceRegion1 As Integer, lSourceRegion2 As Integer, lCombineMode As eCombineMode) As eCombineRegionRet
        Try
            Dim n As Integer = CombineRgn(lDestinationRegion, lSourceRegion1, lSourceRegion2, CInt(lCombineMode))
            Select Case n
                Case 0 'NULLREGION
                    lLastError = "CombineRgn Fault: The region is empty."
                    Return Nothing
                Case 1 'SIMPLEREGION
                    Return eCombineRegionRet.cSimpleRegion
                Case 2 'COMPLEXREGION
                    Return eCombineRegionRet.cComplexRegion
                Case 3 'ERROR
                    lLastError = "CombineRgn Fault: No region is created."
                    Return Nothing
                Case Else
                    Return Nothing
            End Select
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Function CombineRegion(lDestinationRegion As Long, lSourceRegion1 As Long, lSourceRegion2 As Long, lCombineMode As eCombineMode) As Boolean")
            Return Nothing
        End Try
    End Function

    Public Function ReturnRegion(lType As eShapeTypes, x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer, Optional cx As Integer = 0, Optional cy As Integer = 0) As Integer
        Try
            Dim n As Integer
            Select Case lType
                Case eShapeTypes.sRectRgn
                    n = CreateRectRgn(x1, y1, x2, y2)
                Case eShapeTypes.sEllipce
                    n = CreateEllipticRgn(x1, y1, x2, y2)
                Case eShapeTypes.sRoundRectRgn
                    n = CreateRoundRectRgn(x1, y1, x2, y2, cx, cy)
                Case Else
                    Return Nothing
            End Select
            Select Case n
                Case 0
                    lLastError = "Failure"
                    Return Nothing
                Case Else
                    Return n
            End Select
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Function ReturnRegion(lType As eShapeTypes, x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer, Optional cx As Integer = 0, Optional cy As Integer = 0) As Long")
            Return Nothing
        End Try
    End Function

    Public Function AnimateWindowProc(lAnimationTime As Integer, lForm As Form, lFlags As AnimateWindowFlags) As Boolean
        Try
            AnimateWindow(CInt(lForm.Handle), lAnimationTime, lFlags)
            Return True
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "AnimateWindowProc")
            Return Nothing
        End Try
    End Function
End Class