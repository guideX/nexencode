Option Explicit On
Option Strict On
Imports nexENCODE.Enum.Skin
Imports nexENCODE.Models.Skin

Public Class clsSkin
    Public Event ProcessError(lError As String, lSub As String)
#Region "CLASSES"
    Private WithEvents lIniFiles As New clsIniFiles
    Private WithEvents lPrivateProfileString As New clsPrivateProfileString
    Private WithEvents lFiles As New clsFiles
    Private WithEvents lAPI As New clsAPI
#End Region
#Region "DECLARATIONS"
    Public lSkins As New SkinsModel
#End Region
#Region "FUNCTIONS"
    Private Function LoadObjects(lForm As Form, lSkinIndex As Integer, lObjectHandler As clsObjectHandler) As Boolean
        Try
            Dim b As Boolean = True
            For i = 1 To lSkins.Skin(lSkinIndex).MainWindow_ObjectCount
                With lSkins.Skin(lSkinIndex).MainWindow_Objects(i)
                    Select Case .ObjectType
                        Case ObjectTypes.ImageButton
                            If Not lObjectHandler.CreateImageButton(.ButtonType, .Name, .Filename, .Filename2, .Filename3, .Left, .Top, .Width, .Height, .Visible, lForm) Then
                                b = False
                                Exit For
                            End If
                        Case ObjectTypes.StatusLabel
                            Select Case .LabelType
                                Case LabelTypes.Status
                                    If Not lObjectHandler.CreateStatusLabel(.Width, .Height, .Left, .Top, lForm) Then
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
            With lSkins.Skin(lSkinIndex)
                If .MainWindow_SetShape = False Then Return True
                If .MainWindow_ShapeCount <> 0 Then
                    For i = 1 To .MainWindow_ShapeCount
                        Select Case .MainWindow_Shape(i).Type
                            Case ShapeTypes.RoundRectRgn
                                If .UseWindowMetrics = True Then
                                    .MainWindow_Shape(i).Rgn.Rgn = lAPI.ReturnRegion(ShapeTypes.RoundRectRgn, X + .MainWindow_Shape(i).Rgn.X1, Y + .MainWindow_Shape(i).Rgn.Y1, X + .MainWindow_Shape(i).Rgn.X2, Y + .MainWindow_Shape(i).Rgn.Y2, .MainWindow_Shape(i).Rgn.X3, .MainWindow_Shape(i).Rgn.Y3)
                                Else
                                    .MainWindow_Shape(i).Rgn.Rgn = lAPI.ReturnRegion(ShapeTypes.RoundRectRgn, .MainWindow_Shape(i).Rgn.X1, .MainWindow_Shape(i).Rgn.Y1, .MainWindow_Shape(i).Rgn.X2, .MainWindow_Shape(i).Rgn.Y2, .MainWindow_Shape(i).Rgn.X3, .MainWindow_Shape(i).Rgn.Y3)
                                End If
                            Case Else
                                If .UseWindowMetrics = True Then
                                    .MainWindow_Shape(i).Rgn.Rgn = lAPI.ReturnRegion(.MainWindow_Shape(i).Type, X + .MainWindow_Shape(i).Rgn.X1, Y + .MainWindow_Shape(i).Rgn.Y1, X + .MainWindow_Shape(i).Rgn.X2, Y + .MainWindow_Shape(i).Rgn.Y2)
                                Else
                                    .MainWindow_Shape(i).Rgn.Rgn = lAPI.ReturnRegion(.MainWindow_Shape(i).Type, .MainWindow_Shape(i).Rgn.X1, .MainWindow_Shape(i).Rgn.Y1, .MainWindow_Shape(i).Rgn.X2, .MainWindow_Shape(i).Rgn.Y2)
                                End If
                        End Select
                    Next i
                    If .Combine = True Then
                        For i = 1 To .MainWindow_ShapeCount
                            If .MainWindow_Shape(i).CombineMode <> 0 And .MainWindow_Shape(i).DestRgn <> 0 And .MainWindow_Shape(i).SrcRgn1 <> 0 And .MainWindow_Shape(i).SrcRgn2 <> 0 Then
                                lCombineRegionRet = lAPI.CombineRegion(.MainWindow_Shape(.MainWindow_Shape(i).DestRgn).Rgn.Rgn, .MainWindow_Shape(.MainWindow_Shape(i).SrcRgn1).Rgn.Rgn, .MainWindow_Shape(.MainWindow_Shape(i).SrcRgn2).Rgn.Rgn, .MainWindow_Shape(i).CombineMode)
                                If lCombineRegionRet <> clsAPI.eCombineRegionRet.cSimpleRegion And lCombineRegionRet <> clsAPI.eCombineRegionRet.cComplexRegion And lCombineRegionRet <> clsAPI.eCombineRegionRet.cNullRegion Then
                                    RaiseEvent ProcessError(lAPI.lLastError, "CombineRegion")
                                End If
                            End If
                        Next i
                    End If
                    Return lAPI.SetWindowRegion(lForm, .MainWindow_Shape(.MainWindow_ParentShapeRegion).Rgn.Rgn, True)
                Else
                    Return False
                End If
            End With
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Sub LoadShape(lForm As Form, lSkinIndex As Integer)")
            Return Nothing
        End Try
    End Function

    Public Sub LoadSkins()
        Try
            Dim i As Integer, n As Integer, x As Integer
            n = CInt(Trim(lPrivateProfileString.ReadINI(frmMain.lnexENCODE.GlobalController.Ini.Skins, "Settings", "Count", "0")))
            lSkins.DefaultSkinIndex = CInt(Trim(lPrivateProfileString.ReadINI(frmMain.lnexENCODE.GlobalController.Ini.Skins, "Settings", "DefaultSkin", "0")))
            lSkins.SkinIndex = CInt(Trim(lPrivateProfileString.ReadINI(frmMain.lnexENCODE.GlobalController.Ini.Skins, "Settings", "SkinIndex", "0")))
            lSkins.Count = n
            For i = 1 To n
                ReDim Preserve lSkins.Skin(i)
                With lSkins.Skin(i)
                    .FileName = frmMain.lnexENCODE.GlobalController.Skins.ReplaceIndicators(lPrivateProfileString.ReadINI(frmMain.lnexENCODE.GlobalController.Ini.Skins, i.ToString, "File", ""))
                    .MainWindow_ShapeFileName = frmMain.lnexENCODE.GlobalController.Skins.ReplaceIndicators(lPrivateProfileString.ReadINI(.FileName, "Settings", "MainWindow_ShapeFileName", ""), .FileName)
                    .MainWindow_ObjectFileName = frmMain.lnexENCODE.GlobalController.Skins.ReplaceIndicators(lPrivateProfileString.ReadINI(.FileName, "Settings", "MainWindow_ObjectFileName", ""), .FileName)
                    .MainWindow_BackgroundImage = frmMain.lnexENCODE.GlobalController.Skins.ReplaceIndicators(lPrivateProfileString.ReadINI(.FileName, "Settings", "MainWindow_BackgroundImage", ""), .FileName)
                    .MainWindow_CodeFile = frmMain.lnexENCODE.GlobalController.Skins.ReplaceIndicators(lPrivateProfileString.ReadINI(.FileName, "Settings", "MainWindow_CodeFile", ""), .FileName)
                    If frmMain.lnexENCODE.GlobalController.Skins.DoSkinFilesExist(i) = True Then
                        .MainWindow_ShapeCount = CInt(Trim(lPrivateProfileString.ReadINI(.MainWindow_ShapeFileName, "Settings", "Count", "0")))
                        .MainWindow_ObjectCount = CInt(Trim(lPrivateProfileString.ReadINI(.MainWindow_ObjectFileName, "Settings", "Count", "0")))
                        .MainWindow_SetShape = CBool(Trim(lPrivateProfileString.ReadINI(.FileName, "Settings", "MainWindow_SetShape", "False")))
                        .Name = lPrivateProfileString.ReadINI(.MainWindow_ShapeFileName, "Settings", "Name", "")
                        .Icon = frmMain.lnexENCODE.GlobalController.Skins.ReplaceIndicators(lPrivateProfileString.ReadINI(.FileName, "Settings", "Icon", ""), .FileName)
                        .Width = CInt(Trim(lPrivateProfileString.ReadINI(.FileName, "Settings", "Width", "0")))
                        .Height = CInt(Trim(lPrivateProfileString.ReadINI(.FileName, "Settings", "Height", "0")))
                        .MainWindow_ParentShapeRegion = CInt(Trim(lPrivateProfileString.ReadINI(.MainWindow_ShapeFileName, "Settings", "ParentShapeRegion", "0")))
                        .Combine = CBool(lPrivateProfileString.ReadINI(.MainWindow_ShapeFileName, "Settings", "Combine", "True"))
                        .UseWindowMetrics = CBool(lPrivateProfileString.ReadINI(.MainWindow_ShapeFileName, "Settings", "UseWindowMetrics", "True"))
                        If .MainWindow_ObjectCount <> 0 Then
                            For x = 1 To .MainWindow_ObjectCount
                                ReDim Preserve .MainWindow_Objects(x)
                                Dim obj = New ObjectModel
                                .MainWindow_Objects(x).Name = lPrivateProfileString.ReadINI(.MainWindow_ObjectFileName, x.ToString, "Name", "0")
                                If Len(.MainWindow_Objects(x).Name) <> 0 Then
                                    .MainWindow_Objects(x).Filename = frmMain.lnexENCODE.GlobalController.Skins.ReplaceIndicators(lPrivateProfileString.ReadINI(.MainWindow_ObjectFileName, x.ToString, "Filename", ""), .FileName)
                                    .MainWindow_Objects(x).Filename2 = frmMain.lnexENCODE.GlobalController.Skins.ReplaceIndicators(lPrivateProfileString.ReadINI(.MainWindow_ObjectFileName, x.ToString, "Filename2", ""), .FileName)
                                    .MainWindow_Objects(x).Filename3 = frmMain.lnexENCODE.GlobalController.Skins.ReplaceIndicators(lPrivateProfileString.ReadINI(.MainWindow_ObjectFileName, x.ToString, "Filename3", ""), .FileName)
                                    .MainWindow_Objects(x).Height = CInt(Trim(lPrivateProfileString.ReadINI(.MainWindow_ObjectFileName, x.ToString, "height", "0")))
                                    .MainWindow_Objects(x).Left = CInt(Trim(lPrivateProfileString.ReadINI(.MainWindow_ObjectFileName, x.ToString, "left", "0")))
                                    .MainWindow_Objects(x).Width = CInt(Trim(lPrivateProfileString.ReadINI(.MainWindow_ObjectFileName, x.ToString, "width", "0")))
                                    .MainWindow_Objects(x).Top = CInt(Trim(lPrivateProfileString.ReadINI(.MainWindow_ObjectFileName, x.ToString, "top", "0")))
                                    .MainWindow_Objects(x).ObjectType = CType(Trim(lPrivateProfileString.ReadINI(.MainWindow_ObjectFileName, x.ToString, "objecttype", "0")), ObjectTypes)
                                    .MainWindow_Objects(x).LabelType = CType(Trim(lPrivateProfileString.ReadINI(.MainWindow_ObjectFileName, x.ToString, "labeltype", "0")), LabelTypes)
                                    .MainWindow_Objects(x).ButtonType = CType(Trim(lPrivateProfileString.ReadINI(.MainWindow_ObjectFileName, x.ToString, "buttontype", "0")), ButtonTypes)
                                    .MainWindow_Objects(x).Visible = CBool(Trim(lPrivateProfileString.ReadINI(.MainWindow_ObjectFileName, x.ToString, "visible", "false")))
                                    .MainWindow_Objects(x).OnClick = Trim(lPrivateProfileString.ReadINI(.MainWindow_ObjectFileName, x.ToString, "onclick", ""))
                                End If
                            Next x
                        End If
                        If .MainWindow_ShapeCount <> 0 Then
                            For x = 1 To .MainWindow_ShapeCount
                                ReDim Preserve .MainWindow_Shape(x)
                                .MainWindow_Shape(x).Name = lPrivateProfileString.ReadINI(.MainWindow_ShapeFileName, x.ToString, "name", "")
                                .MainWindow_Shape(x).DestRgn = CInt(lPrivateProfileString.ReadINI(.MainWindow_ShapeFileName, x.ToString, "destrgn", "0"))
                                .MainWindow_Shape(x).SrcRgn1 = CInt(lPrivateProfileString.ReadINI(.MainWindow_ShapeFileName, x.ToString, "srcrgn1", "0"))
                                .MainWindow_Shape(x).SrcRgn2 = CInt(lPrivateProfileString.ReadINI(.MainWindow_ShapeFileName, x.ToString, "srcrgn2", "0"))
                                .MainWindow_Shape(x).CombineMode = CType(lPrivateProfileString.ReadINI(.MainWindow_ShapeFileName, x.ToString, "combinemode", "0"), CombineModes)
                                .MainWindow_Shape(x).Rgn.X1 = CInt(lPrivateProfileString.ReadINI(.MainWindow_ShapeFileName, x.ToString, "x1", "0"))
                                .MainWindow_Shape(x).Rgn.X2 = CInt(lPrivateProfileString.ReadINI(.MainWindow_ShapeFileName, x.ToString, "x2", "0"))
                                .MainWindow_Shape(x).Rgn.X3 = CInt(lPrivateProfileString.ReadINI(.MainWindow_ShapeFileName, x.ToString, "x3", "0"))
                                .MainWindow_Shape(x).Rgn.Y1 = CInt(lPrivateProfileString.ReadINI(.MainWindow_ShapeFileName, x.ToString, "y1", "0"))
                                .MainWindow_Shape(x).Rgn.Y2 = CInt(lPrivateProfileString.ReadINI(.MainWindow_ShapeFileName, x.ToString, "y2", "0"))
                                .MainWindow_Shape(x).Rgn.Y3 = CInt(lPrivateProfileString.ReadINI(.MainWindow_ShapeFileName, x.ToString, "y3", "0"))
                                .MainWindow_Shape(x).Type = CType(lPrivateProfileString.ReadINI(.MainWindow_ShapeFileName, x.ToString, "type", "1"), ShapeTypes)
                            Next x
                        End If
                    End If
                End With
            Next i
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Sub LoadAllSkins()")
        End Try
    End Sub

    Public Sub ApplySkin(lForm As Form, lSkinIndex As Integer, lObjectHandler As clsObjectHandler)
        Try
            Dim n As Integer
            If lSkinIndex <> 0 Then
                n = lSkinIndex
            Else
                If lSkins.DefaultSkinIndex <> 0 Then
                    n = lSkins.DefaultSkinIndex
                End If
            End If
            If n <> 0 Then
                With lSkins.Skin(n)
                    If Len(.MainWindow_BackgroundImage) <> 0 Then lForm.BackgroundImage = System.Drawing.Image.FromFile(.MainWindow_BackgroundImage)
                    lForm.Icon = New System.Drawing.Icon(.Icon)
                    lForm.Width = .Width
                    lForm.Height = .Height
                    lSkins.SkinIndex = n
                    lPrivateProfileString.WriteINI(frmMain.lnexENCODE.GlobalController.Ini.Skins, "Settings", "SkinIndex", n.ToString)
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