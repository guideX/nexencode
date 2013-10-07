'nexENCODE Studio 5.0 Alpha 1.3
'October 6th, 2013
Option Explicit On
Option Strict On
Imports System.Runtime.InteropServices
Imports System.Drawing

Public Class clsSkin
    Public Event ProcessError(lError As String, lSub As String)
#Region "CLASSES"
    Private WithEvents lIniFiles As New clsIniFiles
    Private WithEvents lPrivateProfileString As New clsPrivateProfileString
    Private WithEvents lFiles As New clsFiles
    Private WithEvents lAPI As New clsAPI
#End Region
#Region "DECLARATIONS"
    Public Enum eWindowSize
        wLoading = 1
        wUnloading = 2
    End Enum

    Public Enum eLabelTypes
        lIdle = 0
        lStatus = 1
        lRipAndBurn = 2
        lEncodeAndDecode = 3
    End Enum

    Public Enum eButtonTypes
        oIdle = 0
        oRip = 1
        oRipCancel = 2
        oBurn = 3
        oEncode = 4
        oEncodeCancel = 5
        oDecode = 6
        oMinimize = 7
        oMaximize = 8
        oExit = 9
    End Enum

    Private Enum eObjectTypes
        oCustom = 0
        oImageButton = 1
        oStatusLabel = 2
    End Enum

    Public Structure gImageButtonTag
        Public iName As String
        Public iButtonType As eButtonTypes
        Public iFileName1 As String
        Public iFileName2 As String
        Public iFileName3 As String
    End Structure

    Private Structure gObject
        Public oName As String
        Public oButtonType As eButtonTypes
        Public oLabelType As eLabelTypes
        Public oObjectType As eObjectTypes
        Public oFilename As String
        Public oFilename2 As String
        Public oFilename3 As String
        Public oLeft As Integer
        Public oTop As Integer
        Public oWidth As Integer
        Public oHeight As Integer
        Public oVisible As Boolean
    End Structure

    Private Structure gShapes
        Public sShape() As gShape
        Public sCount As Integer
    End Structure

    Private Structure gRegions
        Public rRgn As Integer
        Public X1 As Integer
        Public X2 As Integer
        Public X3 As Integer
        Public Y1 As Integer
        Public Y2 As Integer
        Public Y3 As Integer
    End Structure

    Private Structure gShape
        Public sName As String
        Public sType As clsAPI.eShapeTypes
        Public sRgn As gRegions
        Public sCombineMode As clsAPI.eCombineMode
        Public sDestRgn As Integer
        Public sSrcRgn1 As Integer
        Public sSrcRgn2 As Integer
    End Structure

    Private Structure gSkin
        Public sMainWindow_Shape() As gShape
        Public sMainWindow_ShapeCount As Integer
        Public sMainWindow_ShapeFileName As String
        Public sMainWindow_ParentShapeRegion As Integer
        Public sMainWindow_BackgroundImage As String
        Public sMainWindow_Objects() As gObject
        Public sMainWindow_ObjectCount As Integer
        Public sMainWindow_ObjectFileName As String
        Public sMainWindow_SetShape As Boolean
        Public sMainWindow_CodeFile As String
        Public sFileName As String
        Public sName As String
        Public sWidth As Integer
        Public sHeight As Integer
        Public sCombine As Boolean
        Public sUseWindowMetrics As Boolean
        Public sIcon As String
    End Structure

    Private Structure gSkins
        Public sSkinIndex As Integer
        Public sSkin() As gSkin
        Public sCount As Integer
        Public sDefaultSkinIndex As Integer
    End Structure

    Private lSkins As New gSkins
#End Region
#Region "FUNCTIONS"
    Public Function ReturnSkinMainWindow_CodeFile(lSkinIndex As Integer) As String
        Try
            Return lSkins.sSkin(lSkinIndex).sMainWindow_CodeFile
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Function ReturnSkinMainWindow_CodeFile(lSkinIndex As Integer) As String")
            Return Nothing
        End Try
    End Function

    Public Function ReturnSkinIndex() As Integer
        Try
            Return lSkins.sSkinIndex
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Function ReturnSkinIndex() As Integer")
            Return Nothing
        End Try
    End Function

    Private Function LoadObjects(lForm As Form, lSkinIndex As Integer, lObjectHandler As clsObjectHandler) As Boolean
        Try
            Dim b As Boolean = True
            For i = 1 To lSkins.sSkin(lSkinIndex).sMainWindow_ObjectCount
                With lSkins.sSkin(lSkinIndex).sMainWindow_Objects(i)
                    Select Case .oObjectType
                        Case eObjectTypes.oImageButton
                            If Not lObjectHandler.CreateImageButton(.oButtonType, .oName, .oFilename, .oFilename2, .oFilename3, .oLeft, .oTop, .oWidth, .oHeight, .oVisible, lForm) Then
                                b = False
                                Exit For
                            End If
                        Case eObjectTypes.oStatusLabel
                            Select Case .oLabelType
                                Case eLabelTypes.lStatus
                                    If Not lObjectHandler.CreateStatusLabel(.oWidth, .oHeight, .oLeft, .oTop, lForm) Then
                                        b = False
                                        Exit For
                                    End If
                            End Select
                    End Select
                End With
            Next i
            Return b
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Private Function LoadObjects(lForm As Form, lSkinIndex As Integer) As Boolean")
            Return Nothing
        End Try
    End Function

    Private Function LoadShape(lForm As Form, lSkinIndex As Integer) As Boolean
        Try
            Dim lWindowSettings As clsAPI.gWindowSettings = lAPI.WindowSettings(lForm), X As Integer, Y As Integer, i As Integer, lCombineRegionRet As clsAPI.eCombineRegionRet
            X = lWindowSettings.wWindowBorder
            Y = lWindowSettings.wTitleBarHeight
            With lSkins.sSkin(lSkinIndex)
                If .sMainWindow_SetShape = False Then Return True
                If .sMainWindow_ShapeCount <> 0 Then
                    For i = 1 To .sMainWindow_ShapeCount
                        Select Case .sMainWindow_Shape(i).sType
                            Case clsAPI.eShapeTypes.sRoundRectRgn
                                If .sUseWindowMetrics = True Then
                                    .sMainWindow_Shape(i).sRgn.rRgn = lAPI.ReturnRegion(clsAPI.eShapeTypes.sRoundRectRgn, X + .sMainWindow_Shape(i).sRgn.X1, Y + .sMainWindow_Shape(i).sRgn.Y1, X + .sMainWindow_Shape(i).sRgn.X2, Y + .sMainWindow_Shape(i).sRgn.Y2, .sMainWindow_Shape(i).sRgn.X3, .sMainWindow_Shape(i).sRgn.Y3)
                                Else
                                    .sMainWindow_Shape(i).sRgn.rRgn = lAPI.ReturnRegion(clsAPI.eShapeTypes.sRoundRectRgn, .sMainWindow_Shape(i).sRgn.X1, .sMainWindow_Shape(i).sRgn.Y1, .sMainWindow_Shape(i).sRgn.X2, .sMainWindow_Shape(i).sRgn.Y2, .sMainWindow_Shape(i).sRgn.X3, .sMainWindow_Shape(i).sRgn.Y3)
                                End If
                            Case Else
                                If .sUseWindowMetrics = True Then
                                    .sMainWindow_Shape(i).sRgn.rRgn = lAPI.ReturnRegion(.sMainWindow_Shape(i).sType, X + .sMainWindow_Shape(i).sRgn.X1, Y + .sMainWindow_Shape(i).sRgn.Y1, X + .sMainWindow_Shape(i).sRgn.X2, Y + .sMainWindow_Shape(i).sRgn.Y2)
                                Else
                                    .sMainWindow_Shape(i).sRgn.rRgn = lAPI.ReturnRegion(.sMainWindow_Shape(i).sType, .sMainWindow_Shape(i).sRgn.X1, .sMainWindow_Shape(i).sRgn.Y1, .sMainWindow_Shape(i).sRgn.X2, .sMainWindow_Shape(i).sRgn.Y2)
                                End If
                        End Select
                    Next i
                    If .sCombine = True Then
                        For i = 1 To .sMainWindow_ShapeCount
                            If .sMainWindow_Shape(i).sCombineMode <> 0 And .sMainWindow_Shape(i).sDestRgn <> 0 And .sMainWindow_Shape(i).sSrcRgn1 <> 0 And .sMainWindow_Shape(i).sSrcRgn2 <> 0 Then
                                lCombineRegionRet = lAPI.CombineRegion(.sMainWindow_Shape(.sMainWindow_Shape(i).sDestRgn).sRgn.rRgn, .sMainWindow_Shape(.sMainWindow_Shape(i).sSrcRgn1).sRgn.rRgn, .sMainWindow_Shape(.sMainWindow_Shape(i).sSrcRgn2).sRgn.rRgn, .sMainWindow_Shape(i).sCombineMode)
                                If lCombineRegionRet <> clsAPI.eCombineRegionRet.cSimpleRegion And lCombineRegionRet <> clsAPI.eCombineRegionRet.cComplexRegion And lCombineRegionRet <> clsAPI.eCombineRegionRet.cNullRegion Then
                                    RaiseEvent ProcessError(lAPI.lLastError, "CombineRegion")
                                End If
                            End If
                        Next i
                    End If
                    Return lAPI.SetWindowRegion(lForm, .sMainWindow_Shape(.sMainWindow_ParentShapeRegion).sRgn.rRgn, True)
                Else
                    Return False
                End If
            End With
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Sub LoadShape(lForm As Form, lSkinIndex As Integer)")
            Return Nothing
        End Try
    End Function

    Public Function ReturnLastSkinIndex() As Integer
        Try
            Return lSkins.sSkinIndex
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Function ReturnLastSkinIndex() As Integer")
            Return Nothing
        End Try
    End Function

    Private Function ReplaceIndicators(lPath As String, Optional lSkinFile As String = "") As String
        Try
            Dim msg As String = lPath, msg2 As String
            msg = Replace(msg, "$apppath", Application.StartupPath)
            msg = Replace(msg, "$skinspath", Application.StartupPath & "\data\skins")
            If Len(lSkinFile) <> 0 Then
                msg2 = lFiles.ReturnDirectoryFromFilePath(lSkinFile)
                msg = Replace(msg, "$skinpath", msg2)
            End If
            Return msg
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Function ReplaceIndicators(lData As String) As String")
            Return Nothing
        End Try
    End Function

    Private Function DoSkinFilesExist(lSkinIndex As Integer) As Boolean
        Try
            If lSkinIndex <> 0 Then
                If lFiles.DoesFileExist(lSkins.sSkin(lSkinIndex).sFileName) = False Then
                    RaiseEvent ProcessError("Skin File Doesn't Exist!", "DoSkinFilesExist")
                    Return False
                End If
                If lFiles.DoesFileExist(lSkins.sSkin(lSkinIndex).sMainWindow_ShapeFileName) = False Then
                    RaiseEvent ProcessError("Main Window Shape File Doesn't Exist!", "DoSkinFilesExist")
                    Return False
                End If
                If lFiles.DoesFileExist(lSkins.sSkin(lSkinIndex).sMainWindow_ObjectFileName) = False Then
                    RaiseEvent ProcessError("Main Window Objects File Doesn't Exist!", "DoSkinFilesExist")
                    Return False
                End If
                Return True
            Else
                Return Nothing
            End If
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Function DoSkinFilesExist(lSkinIndex As Integer) As Boolean")
            Return Nothing
        End Try
    End Function

    Private Function FindSkinIndexByName(lName As String) As Integer
        Try
            Dim n As Integer
            For i As Integer = 1 To lSkins.sCount
                If LCase(Trim(lSkins.sSkin(i).sName)) = LCase(Trim(lName)) Then
                    n = i
                    Exit For
                End If
            Next i
            Return n
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Function FindSkinIndexByName(lName As String) As Integer")
            Return Nothing
        End Try
    End Function

    Public Sub LoadSkins()
        Try
            Dim i As Integer, n As Integer, x As Integer
            n = CInt(Trim(lPrivateProfileString.ReadINI(lIniFiles.SkinsINI, "Settings", "Count", "0")))
            lSkins.sDefaultSkinIndex = CInt(Trim(lPrivateProfileString.ReadINI(lIniFiles.SkinsINI, "Settings", "DefaultSkin", "0")))
            lSkins.sSkinIndex = CInt(Trim(lPrivateProfileString.ReadINI(lIniFiles.SkinsINI, "Settings", "SkinIndex", "0")))
            lSkins.sCount = n
            For i = 1 To n
                ReDim Preserve lSkins.sSkin(i)
                With lSkins.sSkin(i)
                    .sFileName = ReplaceIndicators(lPrivateProfileString.ReadINI(lIniFiles.SkinsINI, i.ToString, "File", ""))
                    .sMainWindow_ShapeFileName = ReplaceIndicators(lPrivateProfileString.ReadINI(.sFileName, "Settings", "MainWindow_ShapeFileName", ""), .sFileName)
                    .sMainWindow_ObjectFileName = ReplaceIndicators(lPrivateProfileString.ReadINI(.sFileName, "Settings", "MainWindow_ObjectFileName", ""), .sFileName)
                    .sMainWindow_BackgroundImage = ReplaceIndicators(lPrivateProfileString.ReadINI(.sFileName, "Settings", "MainWindow_BackgroundImage", ""), .sFileName)
                    .sMainWindow_CodeFile = ReplaceIndicators(lPrivateProfileString.ReadINI(.sFileName, "Settings", "MainWindow_CodeFile", ""), .sFileName)
                    If DoSkinFilesExist(i) = True Then
                        .sMainWindow_ShapeCount = CInt(Trim(lPrivateProfileString.ReadINI(.sMainWindow_ShapeFileName, "Settings", "Count", "0")))
                        .sMainWindow_ObjectCount = CInt(Trim(lPrivateProfileString.ReadINI(.sMainWindow_ObjectFileName, "Settings", "Count", "0")))
                        .sMainWindow_SetShape = CBool(Trim(lPrivateProfileString.ReadINI(.sFileName, "Settings", "MainWindow_SetShape", "False")))
                        .sName = lPrivateProfileString.ReadINI(.sMainWindow_ShapeFileName, "Settings", "Name", "")
                        .sIcon = ReplaceIndicators(lPrivateProfileString.ReadINI(.sFileName, "Settings", "Icon", ""), .sFileName)
                        .sWidth = CInt(Trim(lPrivateProfileString.ReadINI(.sFileName, "Settings", "Width", "0")))
                        .sHeight = CInt(Trim(lPrivateProfileString.ReadINI(.sFileName, "Settings", "Height", "0")))
                        .sMainWindow_ParentShapeRegion = CInt(Trim(lPrivateProfileString.ReadINI(.sMainWindow_ShapeFileName, "Settings", "ParentShapeRegion", "0")))
                        .sCombine = CBool(lPrivateProfileString.ReadINI(.sMainWindow_ShapeFileName, "Settings", "Combine", "True"))
                        .sUseWindowMetrics = CBool(lPrivateProfileString.ReadINI(.sMainWindow_ShapeFileName, "Settings", "UseWindowMetrics", "True"))
                        If .sMainWindow_ObjectCount <> 0 Then
                            For x = 1 To .sMainWindow_ObjectCount
                                ReDim Preserve .sMainWindow_Objects(x)
                                .sMainWindow_Objects(x).oName = lPrivateProfileString.ReadINI(.sMainWindow_ObjectFileName, x.ToString, "Name", "0")
                                If Len(.sMainWindow_Objects(x).oName) <> 0 Then
                                    .sMainWindow_Objects(x).oFilename = ReplaceIndicators(lPrivateProfileString.ReadINI(.sMainWindow_ObjectFileName, x.ToString, "Filename", ""), .sFileName)
                                    .sMainWindow_Objects(x).oFilename2 = ReplaceIndicators(lPrivateProfileString.ReadINI(.sMainWindow_ObjectFileName, x.ToString, "Filename2", ""), .sFileName)
                                    .sMainWindow_Objects(x).oFilename3 = ReplaceIndicators(lPrivateProfileString.ReadINI(.sMainWindow_ObjectFileName, x.ToString, "Filename3", ""), .sFileName)
                                    .sMainWindow_Objects(x).oHeight = CInt(Trim(lPrivateProfileString.ReadINI(.sMainWindow_ObjectFileName, x.ToString, "height", "0")))
                                    .sMainWindow_Objects(x).oLeft = CInt(Trim(lPrivateProfileString.ReadINI(.sMainWindow_ObjectFileName, x.ToString, "left", "0")))
                                    .sMainWindow_Objects(x).oWidth = CInt(Trim(lPrivateProfileString.ReadINI(.sMainWindow_ObjectFileName, x.ToString, "width", "0")))
                                    .sMainWindow_Objects(x).oTop = CInt(Trim(lPrivateProfileString.ReadINI(.sMainWindow_ObjectFileName, x.ToString, "top", "0")))
                                    .sMainWindow_Objects(x).oObjectType = CType(Trim(lPrivateProfileString.ReadINI(.sMainWindow_ObjectFileName, x.ToString, "objecttype", "0")), eObjectTypes)
                                    .sMainWindow_Objects(x).oLabelType = CType(Trim(lPrivateProfileString.ReadINI(.sMainWindow_ObjectFileName, x.ToString, "labeltype", "0")), eLabelTypes)
                                    .sMainWindow_Objects(x).oButtonType = CType(Trim(lPrivateProfileString.ReadINI(.sMainWindow_ObjectFileName, x.ToString, "buttontype", "0")), eButtonTypes)
                                    .sMainWindow_Objects(x).oVisible = CBool(Trim(lPrivateProfileString.ReadINI(.sMainWindow_ObjectFileName, x.ToString, "visible", "false")))
                                End If
                            Next x
                        End If
                        If .sMainWindow_ShapeCount <> 0 Then
                            For x = 1 To .sMainWindow_ShapeCount
                                ReDim Preserve .sMainWindow_Shape(x)
                                .sMainWindow_Shape(x).sName = lPrivateProfileString.ReadINI(.sMainWindow_ShapeFileName, x.ToString, "name", "")
                                .sMainWindow_Shape(x).sDestRgn = CInt(lPrivateProfileString.ReadINI(.sMainWindow_ShapeFileName, x.ToString, "destrgn", "0"))
                                .sMainWindow_Shape(x).sSrcRgn1 = CInt(lPrivateProfileString.ReadINI(.sMainWindow_ShapeFileName, x.ToString, "srcrgn1", "0"))
                                .sMainWindow_Shape(x).sSrcRgn2 = CInt(lPrivateProfileString.ReadINI(.sMainWindow_ShapeFileName, x.ToString, "srcrgn2", "0"))
                                .sMainWindow_Shape(x).sCombineMode = CType(lPrivateProfileString.ReadINI(.sMainWindow_ShapeFileName, x.ToString, "combinemode", "0"), clsAPI.eCombineMode)
                                .sMainWindow_Shape(x).sRgn.X1 = CInt(lPrivateProfileString.ReadINI(.sMainWindow_ShapeFileName, x.ToString, "x1", "0"))
                                .sMainWindow_Shape(x).sRgn.X2 = CInt(lPrivateProfileString.ReadINI(.sMainWindow_ShapeFileName, x.ToString, "x2", "0"))
                                .sMainWindow_Shape(x).sRgn.X3 = CInt(lPrivateProfileString.ReadINI(.sMainWindow_ShapeFileName, x.ToString, "x3", "0"))
                                .sMainWindow_Shape(x).sRgn.Y1 = CInt(lPrivateProfileString.ReadINI(.sMainWindow_ShapeFileName, x.ToString, "y1", "0"))
                                .sMainWindow_Shape(x).sRgn.Y2 = CInt(lPrivateProfileString.ReadINI(.sMainWindow_ShapeFileName, x.ToString, "y2", "0"))
                                .sMainWindow_Shape(x).sRgn.Y3 = CInt(lPrivateProfileString.ReadINI(.sMainWindow_ShapeFileName, x.ToString, "y3", "0"))
                                .sMainWindow_Shape(x).sType = CType(lPrivateProfileString.ReadINI(.sMainWindow_ShapeFileName, x.ToString, "type", "1"), clsAPI.eShapeTypes)
                            Next x
                        End If
                    End If
                End With
            Next i
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Sub LoadAllSkins()")
        End Try
    End Sub

    Public Sub WindowSize(lType As eWindowSize, lForm As Form)
        Try
            Dim msg As String = lForm.Name, lIni As String = lIniFiles.WindowPosINI
            If Len(msg) <> 0 Then
                If lType = eWindowSize.wLoading Then
                    lForm.Left = CInt(lPrivateProfileString.ReadINI(lIni, msg, "Left", lForm.Left.ToString))
                    lForm.Top = CInt(lPrivateProfileString.ReadINI(lIni, msg, "Top", lForm.Top.ToString))
                    lForm.Width = CInt(lPrivateProfileString.ReadINI(lIni, msg, "Width", lForm.Width.ToString))
                    lForm.Height = CInt(lPrivateProfileString.ReadINI(lIni, msg, "Height", lForm.Height.ToString))
                Else
                    lPrivateProfileString.WriteINI(lIni, msg, "Left", lForm.Left.ToString)
                    lPrivateProfileString.WriteINI(lIni, msg, "Top", lForm.Top.ToString)
                    lPrivateProfileString.WriteINI(lIni, msg, "Width", lForm.Width.ToString)
                    lPrivateProfileString.WriteINI(lIni, msg, "Height", lForm.Height.ToString)
                End If
            End If
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Sub WindowSize(lType As eWindowSize, lForm As Form)")
        End Try
    End Sub

    Public Sub ApplySkin(lForm As Form, lSkinIndex As Integer, lObjectHandler As clsObjectHandler)
        Try
            Dim n As Integer
            If lSkinIndex <> 0 Then
                n = lSkinIndex
            Else
                If lSkins.sDefaultSkinIndex <> 0 Then
                    n = lSkins.sDefaultSkinIndex
                End If
            End If
            If n <> 0 Then
                With lSkins.sSkin(n)
                    If Len(.sMainWindow_BackgroundImage) <> 0 Then lForm.BackgroundImage = System.Drawing.Image.FromFile(.sMainWindow_BackgroundImage)
                    lForm.Icon = New System.Drawing.Icon(.sIcon)
                    lForm.Width = .sWidth
                    lForm.Height = .sHeight
                    lSkins.sSkinIndex = n
                    lPrivateProfileString.WriteINI(lIniFiles.SkinsINI, "Settings", "SkinIndex", n.ToString)
                    If LoadShape(lForm, n) = False Then
                        RaiseEvent ProcessError("Failure", "ApplySkin - Shape")
                    End If
                    If LoadObjects(lForm, n, lObjectHandler) = False Then
                        RaiseEvent ProcessError("Failure", "ApplySkin - Objects")
                    End If
                End With
            Else
                RaiseEvent ProcessError("No skin was selected", "ApplySkin")
            End If
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Sub ApplySkin(lForm As Form, lSkinIndex As Integer)")
        End Try
    End Sub

    Public Function AnimateWindow(lAnimationTime As Integer, lForm As Form, lFlags As clsAPI.AnimateWindowFlags) As Boolean
        Try
            Return lAPI.AnimateWindowProc(lAnimationTime, lForm, lFlags)
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Function AnimateWindow(lForm As Form) As Boolean")
            Return Nothing
        End Try
    End Function
#End Region
#Region "ERRORHANDLING"
    Private Sub lAPI_ProcessError(lError As String, lSub As String) Handles lAPI.ProcessError
        Try
            RaiseEvent ProcessError(lError, lSub)
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Private Sub lAPI_ProcessError(lError As String, lSub As String) Handles lAPI.ProcessError")
        End Try
    End Sub

    Private Sub lFiles_ProcessError(lError As String, lSub As String) Handles lFiles.ProcessError
        Try
            RaiseEvent ProcessError(lError, lSub)
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Private Sub lFiles_ProcessError(lError As String, lSub As String) Handles lFiles.ProcessError")
        End Try
    End Sub

    Private Sub lIniFiles_ProcessError(lError As String, lSub As String) Handles lIniFiles.ProcessError
        Try
            RaiseEvent ProcessError(lError, lSub)
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Private Sub lIniFiles_ProcessError(lError As String, lSub As String) Handles lIniFiles.ProcessError")
        End Try
    End Sub

    Private Sub lPrivateProfileString_ProcessError(lError As String, lSub As String) Handles lPrivateProfileString.ProcessError
        Try
            RaiseEvent ProcessError(lError, lSub)
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Private Sub lPrivateProfileString_ProcessError(lError As String, lSub As String) Handles lPrivateProfileString.ProcessError")
        End Try
    End Sub
#End Region
End Class