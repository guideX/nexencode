Attribute VB_Name = "mdlSkin"
Option Explicit

Public Sub FlashIN(lForm As Form)
'On Local Error Resume Next
pause 0.1
If lEvents.eSettings.iRememberWindowSizes = True Then WindowSize wLoading, lForm: DoEvents
lForm.Show
'If Len(lSkins.sSkin(lSkins.sSkinIndex).sWindowColor) <> 0 Then lForm.BackColor = lSkins.sSkin(lSkins.sSkinIndex).sWindowColor
If Err.Number <> 0 Then SetError "FlashIN()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub FlashOut(lForm As Form)
'On Local Error Resume Next
WindowSize wUnloading, lForm
lForm.Visible = False
If Err.Number <> 0 Then SetError "FlashOUT()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub LoadAllSkins()
'On Local Error Resume Next
Dim s As Integer, X As Integer, i As Integer, m As Integer, msg As String, msg2 As String
X = ReadINI(lIniFiles.iSettings, "MySkins", "Count", 0)
If X <> 0 Then
    For i = 1 To X
        msg = ReadINI(lIniFiles.iSettings, "MySkins", i, "")
        If Len(msg) <> 0 Then OpenSkin msg, False
        DoEvents
    Next i
End If
If Err.Number <> 0 Then SetError "LoadSkins()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Function OpenSkin(lFilename As String, lSelectContainer As Boolean) As Integer
'On Local Error Resume Next
Dim i As Integer, X As Integer, msg As String, f As Integer, msg2 As String, A As Integer

If Len(lFilename) = 0 Then
    lFilename = OpenDialog(frmMain, "NS4 Files (*.ns4)|*.ns4", "Select skin ...", CurDir)
End If
msg2 = lFilename
msg2 = GetFileTitle(msg2)
For A = 1 To lSkins.sCount
    If LCase(lSkins.sSkin(A).sFilename) = LCase(msg2) Then
        OpenSkin = A
        Exit Function
    End If
Next A
A = 0
If Len(lFilename) <> 0 Then
    With lSkins
        i = .sCount + 1
        .sCount = i
        .sSkin(i).sErrorGraphic = ReadINI(lFilename, "Settings", "ErrorGraphic", "")
        .sSkin(i).sSkinSettings.sWidth = ReadINI(lFilename, "Settings", "Width", 200)
        .sSkin(i).sSkinSettings.sHeight = ReadINI(lFilename, "Settings", "Height", 200)
        .sSkin(i).sGraphic = ReadINI(lFilename, "Settings", "Graphic", "")
        .sSkin(i).sName = ReadINI(lFilename, "Settings", "Name", "Default Skin")
        .sSkin(i).sShapeCount = ReadINI(lFilename, "Settings", "ShapeCount", 0)
        .sSkin(i).sObjectCount = ReadINI(lFilename, "Settings", "ObjectCount", 0)
        .sSkin(i).sAuthor = ReadINI(lFilename, "Settings", "Author", "")
        .sSkin(i).sFilename = msg2
        .sSkin(i).sFilepath = Left(lFilename, Len(lFilename) - Len(.sSkin(i).sFilename))
        .sSkin(i).sPlaylistGraphic = ReadINI(lFilename, "Settings", "Playlist", "")
        .sSkin(i).sBanner = ReadINI(lFilename, "Settings", "Banner", "")
        .sSkin(i).sBackground = ReadINI(lFilename, "Settings", "Background", "")
        If Len(.sSkin(i).sBanner) <> 0 And DoesFileExist(.sSkin(i).sFilepath & .sSkin(i).sBanner) = True Then frmGraphics.picNS4.Picture = LoadPicture(.sSkin(i).sFilepath & .sSkin(i).sBanner)
        If Len(.sSkin(i).sErrorGraphic) <> 0 And DoesFileExist(.sSkin(i).sFilepath & .sSkin(i).sErrorGraphic) = True Then frmMain.imgErrorBackground.Picture = LoadPicture(.sSkin(i).sFilepath & .sSkin(i).sErrorGraphic)
        If Len(.sSkin(i).sPlaylistGraphic) <> 0 And DoesFileExist(.sSkin(i).sFilepath & .sSkin(i).sPlaylistGraphic) = True Then frmGraphics.imgPlaylist.Picture = LoadPicture(.sSkin(i).sFilepath & .sSkin(i).sPlaylistGraphic)
        If Len(.sSkin(i).sName) <> 0 Then .sSkin(i).sEnabled = True
        If .sSkin(i).sShapeCount <> 0 Then
            For X = 1 To .sSkin(i).sShapeCount
                msg = "rgn" & X
                .sSkin(i).sShape(X).sEnabled = ReadINI(lFilename, msg, "enabled", "")
                .sSkin(i).sShape(X).sName = ReadINI(lFilename, msg, "name", "")
                .sSkin(i).sShape(X).sDestRgn = ReadINI(lFilename, msg, "destrgn", 0)
                .sSkin(i).sShape(X).sSrcRgn1 = ReadINI(lFilename, msg, "srcrgn1", 0)
                .sSkin(i).sShape(X).sSrcRgn2 = ReadINI(lFilename, msg, "srcrgn2", 0)
                .sSkin(i).sShape(X).sCombineMode = ReadINI(lFilename, msg, "combinemode", 0)
                .sSkin(i).sShape(X).sRgn.X1 = ReadINI(lFilename, msg, "x1", 0)
                .sSkin(i).sShape(X).sRgn.X2 = ReadINI(lFilename, msg, "x2", 0)
                .sSkin(i).sShape(X).sRgn.X3 = ReadINI(lFilename, msg, "x3", 0)
                .sSkin(i).sShape(X).sRgn.Y1 = ReadINI(lFilename, msg, "y1", 0)
                .sSkin(i).sShape(X).sRgn.Y2 = ReadINI(lFilename, msg, "y2", 0)
                .sSkin(i).sShape(X).sRgn.Y3 = ReadINI(lFilename, msg, "y3", 0)
                .sSkin(i).sShape(X).sType = ReadINI(lFilename, msg, "type", 1)
            Next X
        End If
        If .sSkin(i).sObjectCount <> 0 Then
            For X = 1 To .sSkin(i).sObjectCount
                msg = "object" & X
                .sSkin(i).sObject(X).oEnabled = ReadINI(lFilename, msg, "enabled", False)
                If .sSkin(i).sObject(X).oEnabled = True Then
                    .sSkin(i).sObject(X).oName = ReadINI(lFilename, msg, "name", "")
                    .sSkin(i).sObject(X).oFilename = ReadINI(lFilename, msg, "filename", "")
                    .sSkin(i).sObject(X).oFilename2 = ReadINI(lFilename, msg, "filename2", "")
                    .sSkin(i).sObject(X).oFilename3 = ReadINI(lFilename, msg, "filename3", "")
                    .sSkin(i).sObject(X).oPos.sHeight = ReadINI(lFilename, msg, "height", 0)
                    .sSkin(i).sObject(X).oPos.sWidth = ReadINI(lFilename, msg, "width", 0)
                    .sSkin(i).sObject(X).oPos.sLeft = ReadINI(lFilename, msg, "left", 0)
                    .sSkin(i).sObject(X).oPos.sTop = ReadINI(lFilename, msg, "top", 0)
                    .sSkin(i).sObject(X).oType = ReadINI(lFilename, msg, "type", 0)
                End If
            Next X
        End If
    End With
    A = frmMain.mnuSkinName.Count + 1
    Load frmMain.mnuSkinName(A)
    frmMain.mnuSkinName(A).Caption = UCase(lSkins.sSkin(i).sName)
    frmMain.mnuSkinName(A).Visible = True
    SaveSkins
    OpenSkin = i
End If
If Err.Number <> 0 Then SetError "OpenSkin()", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Sub InitSkins()
'On Local Error Resume Next
Dim k As Integer, msg As String

lSkins.sDefaultSkinLocation = App.Path & "\skins\inex\inex.ns4"
LoadAllSkins
lSkins.sLastSkin = ReadINI(lIniFiles.iSettings, "MySkins", "LastSkin", 0)
If lSkins.sLastSkin = 0 Then
    If DoesFileExist(lSkins.sDefaultSkinLocation) = True Then
        If k = lSkins.sLastSkin Then ApplySkin frmMain, OpenSkin(lSkins.sDefaultSkinLocation, False), True
        ApplySkin frmMain, OpenSkin(lSkins.sDefaultSkinLocation, False), True
        k = 0
    Else
        If DoesFileExist(lSkins.sDefaultSkinLocation) = True Then
            If k = lSkins.sLastSkin Then ApplySkin frmMain, OpenSkin(lSkins.sDefaultSkinLocation, False), True
            ApplySkin frmMain, OpenSkin(lSkins.sDefaultSkinLocation, False), True
            k = 0
        Else
            If lEvents.eSettings.iOverwritePrompts = True Then MsgBox "No skins found!", vbExclamation
            msg = OpenSkin("", False)
        End If
    End If
Else
    If Len(lSkins.sSkin(lSkins.sLastSkin).sName) <> 0 Then ApplySkin frmMain, lSkins.sLastSkin, True
End If
If Err.Number <> 0 Then SetError "InitSkins()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub ApplySkin(lForm As Form, lIndex As Integer, Optional lFadeInOnly As Boolean)
Dim i As Integer
'On Local Error Resume Next
'If lFadeInOnly = False Then FlashOut lForm
'ResetButtons
lForm.Width = lSkins.sSkin(lIndex).sSkinSettings.sWidth
lForm.Height = lSkins.sSkin(lIndex).sSkinSettings.sHeight
lSkins.sSkinIndex = lIndex
SetObjects lIndex
LoadShape lForm, lIndex
WriteINI lIniFiles.iSettings, "MySkins", "LastSkin", lIndex
If DoesFileExist(lSkins.sSkin(lIndex).sFilepath & lSkins.sSkin(lIndex).sGraphic) = True Then
    frmMain.Picture = LoadPicture(lSkins.sSkin(lIndex).sFilepath & lSkins.sSkin(lIndex).sGraphic)
    frmMain.imgBackground1.Picture = frmMain.Picture
End If
If DoesFileExist(lSkins.sSkin(lIndex).sFilepath & lSkins.sSkin(lIndex).sBackground) = True Then
    frmMain.imgBackground2.Picture = LoadPicture(lSkins.sSkin(lIndex).sFilepath & lSkins.sSkin(lIndex).sBackground)
End If
'If Len(lSkins.sSkin(lIndex).sToper) <> 0 Then frmGraphics.imgTopper.Picture = LoadPicture(lSkins.sSkin(lIndex).sFilepath & lSkins.sSkin(lIndex).sToper)
'If Len(lSkins.sSkin(lIndex).sSideGradient) <> 0 Then frmGraphics.imgSideGradient.Picture = LoadPicture(lSkins.sSkin(lIndex).sFilepath & lSkins.sSkin(lIndex).sSideGradient)
DoEvents

'MakeTransparent lForm.hWnd, 3
'lForm.Visible = True
'FlashIN lForm
For i = 1 To frmMain.mnuSkinName.Count
    If i - 1 = lIndex Then
        frmMain.mnuSkinName(i).Checked = True
    Else
        frmMain.mnuSkinName(i).Checked = False
    End If
Next i
If Err.Number <> 0 Then SetError "ApplySkin()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Function SaveSkins()
'On Local Error Resume Next
Dim A As Integer

For A = 1 To lSkins.sCount
    WriteINI lIniFiles.iSettings, "MySkins", A, lSkins.sSkin(A).sFilepath & lSkins.sSkin(A).sFilename
Next A
WriteINI lIniFiles.iSettings, "MySkins", "Count", lSkins.sCount
If Err.Number <> 0 Then SetError "SaveSkins()", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Function NewSkin() As String
'On Local Error Resume Next
Dim lFilename As String, lTitle As String

lTitle = InputBox("Enter title of new skin:")
If Len(lTitle) <> 0 Then
    lFilename = SaveDialog(frmSkinEditor, "Storage Containers (*.ns4)|*.ns4|All Files (*.*)|*.*", "Save as ...", App.Path & "\skins\")
    If Len(lFilename) <> 0 Then
        frmImagePreview.Visible = False
        frmShapeEdit.Visible = False
        frmShapePreview.Visible = False
        lFilename = Left(lFilename, Len(lFilename) - 1) & ".ns4"
        WriteINI lFilename, "Settings", "Author", SkinAuthor
        WriteINI lFilename, "Settings", "Name", lTitle
        WriteINI lFilename, "Settings", "Enabled", "True"
        WriteINI lFilename, "Settings", "ShapeCount", 1
        WriteINI lFilename, "rgn1", "enabled", "true"
        WriteINI lFilename, "rgn1", "x1", "10"
        WriteINI lFilename, "rgn1", "x2", "50"
        WriteINI lFilename, "rgn1", "y1", "10"
        WriteINI lFilename, "rgn1", "y2", "50"
        WriteINI lFilename, "rgn1", "type", "1"
        WriteINI lFilename, "rgn1", "name", "shape1"
        NewSkin = lFilename
    End If
End If
If Err.Number <> 0 Then SetError "NewSkin()", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Sub SetSkin(lIndex As Integer)
'On Local Error Resume Next

With lSkins.sSkin(lIndex)
    If .sName <> "" Then
        With frmSkinEditor
            .lstValues.Enabled = True
            .lstProporties.Enabled = True
            .cboType.Enabled = True
            .txtProporty.Enabled = True
            .lstProporties.Clear
            .lstValues.Clear
            .txtProporty.Text = ""
            lSkins.sSkinIndex = lIndex
            .Caption = "NexSkin [" & lSkins.sSkin(lIndex).sName & " - Author: " & lSkins.sSkin(lIndex).sAuthor & "]"
            .cboType.Text = "(Select)"
        End With
    End If
End With
If Err.Number <> 0 Then SetError "SetSkin()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub GetWindowSettings(lHandle As Long)
'On Local Error Resume Next
Dim lWindowPos As RECT, lClientPos As RECT, lBorderWidth As Long, lTopOffset As Long, i As Long

i = GetWindowRect(lHandle, lWindowPos)
i = GetClientRect(lHandle, lClientPos)
lMainWndSettings.wTitleBarHeight = lWindowPos.Bottom - lWindowPos.Top - lClientPos.Bottom - lBorderWidth
lMainWndSettings.wWindowBorder = lWindowPos.Right - lWindowPos.Left - lClientPos.Right - 2
If Err.Number <> 0 Then SetError "GetWindowSettings()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Function FindSkinIndexByFilename(lFilename As String) As Integer
'On Local Error Resume Next
Dim i As Integer

If Len(lFilename) <> 0 Then
    For i = 1 To lSkins.sCount
        If LCase(lFilename) = LCase(lSkins.sSkin(i).sFilename) Then FindSkinIndexByFilename = i
    Next i
End If
If Err.Number <> 0 Then SetError "FindSkinIndexByFilename()", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Function FindSkinIndex(lName As String) As Integer
'On Local Error Resume Next
Dim i As Integer

If Len(lName) <> 0 Then
    For i = 1 To lSkins.sCount
        If LCase(lName) = LCase(lSkins.sSkin(i).sName) Then FindSkinIndex = i
    Next i
End If
If Err.Number <> 0 Then SetError "FindSkinIndex()", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Sub SetImageBox(lImageBox As Image, lImageBox1 As Image, lImageBox2 As Image, lImg1 As String, lImg2 As String, lLeft As Long, lTop As Long, Optional lImg3 As String, Optional lImageBox3 As Image)
'On Local Error Resume Next
If Len(lImg1) <> 0 Then
    lImageBox1.Picture = LoadPicture(lImg1)
    With lImageBox
        .Picture = lImageBox1.Picture
        .Left = lLeft
        .Top = lTop
    End With
    'lImageBox.Visible = True
End If

If Len(lImg2) <> 0 Then lImageBox2.Picture = LoadPicture(lImg2)
If Len(lImg3) <> 0 Then lImageBox3.Picture = LoadPicture(lImg3)
If Err.Number <> 0 Then SetError "SetImageBox()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub SetPictureBox(lPictureBox As PictureBox, lPictureBox1 As PictureBox, lPictureBox2 As PictureBox, lImg1 As String, lImg2 As String, lLeft As Long, lTop As Long)
'On Local Error Resume Next

If Len(lImg1) <> 0 And Len(lImg1) <> 1 Then
    lPictureBox.Left = lLeft
    lPictureBox.Top = lTop
    lPictureBox.Picture = LoadPicture(lImg1)
    lPictureBox1.Picture = LoadPicture(lImg1)
    If Len(lImg2) <> 0 Then lPictureBox2.Picture = LoadPicture(lImg2)
End If
End Sub

Public Sub SetLabel(lLabel As Label, lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long)
'On Local Error Resume Next

lLabel.Left = lLeft
lLabel.Top = lTop
lLabel.Width = lWidth
lLabel.Height = lHeight
End Sub

Public Sub SetObjects(lSkinIndex As Integer)
'On Local Error Resume Next
Dim j As Integer, msg As String, lFile As String, lBorderColor As String, lBackColor As String, i As Integer

If lSkins.sSkin(lSkinIndex).sEnabled = True Then
    For i = 1 To lSkins.sSkin(lSkinIndex).sObjectCount
        With lSkins.sSkin(lSkinIndex).sObject(i)
            
            Select Case lSkins.sSkin(lSkinIndex).sObject(i).oType
            Case oWavFileLabel
                If .oEnabled = True Then
                    SetLabel frmMain.lblWavFile, .oPos.sLeft, .oPos.sTop, .oPos.sWidth, .oPos.sHeight
                End If
            Case oMp3FileLabel
                If .oEnabled = True Then
                    SetLabel frmMain.lblMp3File, .oPos.sLeft, .oPos.sTop, .oPos.sWidth, .oPos.sHeight
                End If
            Case oMinimize
                If .oEnabled = True And Len(.oFilename) <> 0 And Len(.oFilename2) <> 0 And DoesFileExist(lSkins.sSkin(lSkinIndex).sFilepath & .oFilename) Then
                    SetImageBox frmMain.imgMinimize, frmMain.imgMinimize1, frmMain.imgMinimize2, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename2, .oPos.sLeft, .oPos.sTop, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename3, frmMain.imgMinimize3
                End If
            Case oEnd
                If .oEnabled = True And Len(.oFilename) <> 0 And Len(.oFilename2) <> 0 And DoesFileExist(lSkins.sSkin(lSkinIndex).sFilepath & .oFilename) Then
                    SetImageBox frmMain.imgEnd, frmMain.imgEnd1, frmMain.imgEnd2, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename2, .oPos.sLeft, .oPos.sTop, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename3, frmMain.imgEnd3
                End If
            Case oRipperCrlEffect
                If .oEnabled = True And Len(.oFilename) <> 0 Then
                    For j = 1 To 10
                        lFile = lSkins.sSkin(lSkins.sSkinIndex).sFilepath & .oFilename
                        frmMain.shpRipper(j - 1).Left = Int(ReadINI(lFile, Str(j), "left", "0"))
                        frmMain.shpRipper(j - 1).Top = Int(ReadINI(lFile, Str(j), "top", ""))
                        frmMain.shpRipper(j - 1).Width = Int(ReadINI(lFile, Str(j), "width", "0"))
                        frmMain.shpRipper(j - 1).Height = Int(ReadINI(lFile, Str(j), "height", ""))
                        frmMain.shpRipper(j - 1).Shape = Int(ReadINI(lFile, Str(j), "shape", "2"))
                        lBackColor = ReadINI(lFile, Str(j), "backcolor", vbBlue)
                        lBorderColor = ReadINI(lFile, Str(j), "bordercolor", vbBlack)
                        'SetShapeColor frmMain.shpRipper(j - 1), lBackColor, lBorderColor
                    Next j
                End If
            Case oEncoderCrlEffect
                If .oEnabled = True And Len(.oFilename) <> 0 Then
                    For j = 1 To 10
                        lFile = lSkins.sSkin(lSkins.sSkinIndex).sFilepath & .oFilename
                        frmMain.shpEncoder(j - 1).Left = Int(ReadINI(lFile, Str(j), "left", "0"))
                        frmMain.shpEncoder(j - 1).Top = Int(ReadINI(lFile, Str(j), "top", ""))
                        frmMain.shpEncoder(j - 1).Width = Int(ReadINI(lFile, Str(j), "width", "0"))
                        frmMain.shpEncoder(j - 1).Height = Int(ReadINI(lFile, Str(j), "height", ""))
                        frmMain.shpEncoder(j - 1).Shape = Int(ReadINI(lFile, Str(j), "shape", "2"))
                        lBackColor = ReadINI(lFile, Str(j), "backcolor", vbBlue)
                        lBorderColor = ReadINI(lFile, Str(j), "bordercolor", vbBlack)
                        'SetShapeColor frmMain.shpEncoder(j - 1), lBackColor, lBorderColor
                    Next j
                End If
            Case oRip
                If .oEnabled = True And Len(.oFilename) <> 0 And Len(.oFilename2) <> 0 And DoesFileExist(lSkins.sSkin(lSkinIndex).sFilepath & .oFilename) Then
                    SetImageBox frmMain.imgRip, frmMain.imgRip1, frmMain.imgRip2, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename2, .oPos.sLeft, .oPos.sTop, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename3, frmMain.imgRip3
                End If
            Case oStopRipping
                If .oEnabled = True And Len(.oFilename) <> 0 And Len(.oFilename2) <> 0 And DoesFileExist(lSkins.sSkin(lSkinIndex).sFilepath & .oFilename) Then
                    SetImageBox frmMain.imgCancelRip, frmMain.imgStopRipping1, frmMain.imgStopRipping2, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename2, .oPos.sLeft, .oPos.sTop, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename3, frmMain.imgCancelRip3
                End If
            Case oPlayWav
                If .oEnabled = True And Len(.oFilename) <> 0 And Len(.oFilename2) <> 0 And DoesFileExist(lSkins.sSkin(lSkinIndex).sFilepath & .oFilename) Then
                    SetImageBox frmMain.imgPlayWav, frmMain.imgPlayWav1, frmMain.imgPlayWav2, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename2, .oPos.sLeft, .oPos.sTop, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename3, frmMain.imgPlayWAV3
                End If
            Case oEncode
                If .oEnabled = True And Len(.oFilename) <> 0 And Len(.oFilename2) <> 0 And DoesFileExist(lSkins.sSkin(lSkinIndex).sFilepath & .oFilename) Then
                    SetImageBox frmMain.imgEncode, frmMain.imgEncode1, frmMain.imgEncode2, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename2, .oPos.sLeft, .oPos.sTop, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename3, frmMain.imgEncode3
                End If
            Case oStopEncoding
                If .oEnabled = True And Len(.oFilename) <> 0 And Len(.oFilename2) <> 0 And DoesFileExist(lSkins.sSkin(lSkinIndex).sFilepath & .oFilename) Then
                    SetImageBox frmMain.imgStopEncoding, frmMain.imgStopEncoding1, frmMain.imgStopEncoding2, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename2, .oPos.sLeft, .oPos.sTop, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename3, frmMain.imgStopEncoding3
                End If
            Case oPlayMp3
                If .oEnabled = True And Len(.oFilename) <> 0 And Len(.oFilename2) <> 0 And DoesFileExist(lSkins.sSkin(lSkinIndex).sFilepath & .oFilename) Then
                    SetImageBox frmMain.imgPlayMp3, frmMain.imgPlayMp31, frmMain.imgPlayMp32, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename2, .oPos.sLeft, .oPos.sTop, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename3, frmMain.imgPlayMP33
                End If
            Case oStatusLabel
                If .oEnabled = True Then
                    SetLabel frmMain.lblInfo, .oPos.sLeft, .oPos.sTop, .oPos.sWidth, .oPos.sHeight
                End If
            Case oProgressFiller
                If .oEnabled = True And Len(.oFilename) <> 0 And DoesFileExist(lSkins.sSkin(lSkinIndex).sFilepath & .oFilename) Then
                    frmMain.imgPercent.Picture = LoadPicture(lSkins.sSkin(lSkinIndex).sFilepath & .oFilename)
                    frmMain.imgPercent.Left = .oPos.sLeft
                    frmMain.imgPercent.Top = .oPos.sTop
                    frmMain.imgPercent.Width = .oPos.sWidth
                    frmMain.imgPercent.Height = .oPos.sHeight
                End If
            Case oSkinEdit
                If .oEnabled = True And Len(.oFilename) <> 0 And DoesFileExist(lSkins.sSkin(lSkinIndex).sFilepath & .oFilename) Then
                    SetImageBox frmMain.imgSkinEdit, frmMain.imgSkinEdit1, frmMain.imgSkinEdit2, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename2, .oPos.sLeft, .oPos.sTop, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename3, frmMain.imgSkinEdit3
                End If
            Case oCDAudio
                If .oEnabled = True And Len(.oFilename) <> 0 And DoesFileExist(lSkins.sSkin(lSkinIndex).sFilepath & .oFilename) Then
                    SetImageBox frmMain.imgNexMedia, frmMain.imgNexMedia1, frmMain.imgNexMedia2, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename2, .oPos.sLeft, .oPos.sTop, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename3, frmMain.imgNexMEDIA3
                End If
            Case oTag
                If .oEnabled = True And Len(.oFilename) <> 0 And DoesFileExist(lSkins.sSkin(lSkinIndex).sFilepath & .oFilename) Then
                    SetImageBox frmMain.imgId3, frmMain.imgId31, frmMain.imgId32, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename2, .oPos.sLeft, .oPos.sTop, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename3, frmMain.imgID33
                End If
            Case oPlayButton
                If .oEnabled = True And Len(.oFilename) <> 0 And DoesFileExist(lSkins.sSkin(lSkinIndex).sFilepath & .oFilename) Then
                    SetImageBox frmMain.imgPlay, frmMain.imgPlay1, frmMain.imgPlay2, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename2, .oPos.sLeft, .oPos.sTop, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename3, frmMain.imgPlay3
                End If
            Case oBackwardButton
                If .oEnabled = True And Len(.oFilename) <> 0 And DoesFileExist(lSkins.sSkin(lSkinIndex).sFilepath & .oFilename) Then
                    SetImageBox frmMain.imgBackward, frmMain.imgBackward1, frmMain.imgBackward2, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename2, .oPos.sLeft, .oPos.sTop, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename3, frmMain.imgBackward3
                End If
            Case oForwardButton
                If .oEnabled = True And Len(.oFilename) <> 0 And DoesFileExist(lSkins.sSkin(lSkinIndex).sFilepath & .oFilename) Then
                    SetImageBox frmMain.imgForward, frmMain.imgForward1, frmMain.imgForward2, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename2, .oPos.sLeft, .oPos.sTop, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename3, frmMain.imgForward3
                End If
            Case oStopButton
                If .oEnabled = True And Len(.oFilename) <> 0 And DoesFileExist(lSkins.sSkin(lSkinIndex).sFilepath & .oFilename) Then
                    SetImageBox frmMain.imgStop, frmMain.imgStop1, frmMain.imgStop2, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename2, .oPos.sLeft, .oPos.sTop, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename3, frmMain.imgStop3
                End If
            Case oOptions
                If .oEnabled = True And Len(.oFilename) <> 0 And DoesFileExist(lSkins.sSkin(lSkinIndex).sFilepath & .oFilename) Then
                    SetImageBox frmMain.imgOptions, frmMain.imgOptions1, frmMain.imgOptions2, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename2, .oPos.sLeft, .oPos.sTop, lSkins.sSkin(lSkinIndex).sFilepath & .oFilename3, frmMain.imgOptions3
                End If
            End Select
        End With
        DoEvents
    Next i
Else
    SetError "SetObjects()", "Skins error", "Objects could not be set because no skin is enabled"
End If
If Err.Number <> 0 Then SetError "SetObjects()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub SetShapeColor(lShape As Shape, lBackColor As String, lBorderColor As String)
'On Local Error Resume Next

If Len(lBackColor) <> 0 And Len(lBorderColor) <> 0 Then
    With lShape
        .BackColor = lBackColor
        Select Case LCase(lBackColor)
            Case "vbblue"
                .BackColor = ColorConstants.vbBlue
            Case "vbblack"
                .BackColor = ColorConstants.vbBlack
            Case "vbcyan"
                .BackColor = ColorConstants.vbCyan
            Case "vbgreen"
                .BackColor = ColorConstants.vbGreen
            Case "vbmagenta"
                .BackColor = ColorConstants.vbMagenta
            Case "vbred"
                .BackColor = ColorConstants.vbRed
            Case "vbwhite"
                .BackColor = ColorConstants.vbWhite
            Case "vbyellow"
                .BackColor = ColorConstants.vbYellow
        End Select
        
        Select Case LCase(lBorderColor)
        Case "vbblue"
            .BorderColor = ColorConstants.vbBlue
        Case "vbblack"
            .BorderColor = ColorConstants.vbBlack
        Case "vbcyan"
            .BorderColor = ColorConstants.vbCyan
        Case "vbgreen"
            .BorderColor = ColorConstants.vbGreen
        Case "vbmagenta"
            .BorderColor = ColorConstants.vbMagenta
        Case "vbred"
            .BorderColor = ColorConstants.vbRed
        Case "vbwhite"
            .BorderColor = ColorConstants.vbWhite
        Case "vbyellow"
            .BorderColor = ColorConstants.vbYellow
        End Select
        
    End With
End If

If Err.Number <> 0 Then SetError "SetShapeColor()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub LoadShape(lForm As Form, lSkinIndex As Integer)
'On Local Error Resume Next
Dim i As Integer, X As Long, Y As Long, tmp As Long

GetWindowSettings lForm.hwnd
X = lMainWndSettings.wWindowBorder
Y = lMainWndSettings.wTitleBarHeight

With lSkins.sSkin(lSkins.sSkinIndex)
    For i = 1 To .sShapeCount
        If .sShape(i).sEnabled = True Then
            Select Case .sShape(i).sType
            Case 1
                .sShape(i).sRgn.rRgn = CreateRectRgn(X + .sShape(i).sRgn.X1, Y + .sShape(i).sRgn.Y1, X + .sShape(i).sRgn.X2, Y + .sShape(i).sRgn.Y2)
            Case 2
                .sShape(i).sRgn.rRgn = CreateEllipticRgn(X + .sShape(i).sRgn.X1, Y + .sShape(i).sRgn.Y1, X + .sShape(i).sRgn.X2, Y + .sShape(i).sRgn.Y2)
            Case 3
                .sShape(i).sRgn.rRgn = CreateRoundRectRgn(X + .sShape(i).sRgn.X1, Y + .sShape(i).sRgn.Y1, X + .sShape(i).sRgn.X2, Y + .sShape(i).sRgn.Y2, .sShape(i).sRgn.X3, .sShape(i).sRgn.Y3)
            End Select
        End If
    Next i
    For i = 1 To .sShapeCount
        If .sShape(i).sEnabled = True Then
            If .sShape(i).sCombineMode <> 0 And .sShape(i).sDestRgn <> 0 And .sShape(i).sSrcRgn1 <> 0 And .sShape(i).sSrcRgn2 <> 0 Then
                tmp = CombineRgn(.sShape(.sShape(i).sDestRgn).sRgn.rRgn, .sShape(.sShape(i).sSrcRgn1).sRgn.rRgn, .sShape(.sShape(i).sSrcRgn2).sRgn.rRgn, .sShape(i).sCombineMode)
            End If
        End If
    Next i
    SetWindowRgn lForm.hwnd, .sShape(1).sRgn.rRgn, True
End With
If Err.Number <> 0 Then SetError "LoadShape()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub PictureBoxMouseMove(lType As eObjectTypes, lButton As Integer, lPictureBox As Image, lPic1 As Image, lPic2 As Image, lX As Single, lY As Single, Optional lPic3 As Image, Optional lOver As Boolean)
'On Local Error Resume Next
If lButton = 0 Then
    If lOver = True Then
        CheckMainButtonsOver lType
        lPictureBox.Picture = lPic3.Picture
    End If
    If lEvents.eRipperBusy = False And lEvents.eEncoderBusy = False Then
        lEvents.ePercent = 0
        ConvertCaption lType
        Exit Sub
    End If
ElseIf lButton = 1 Then
    lX = frmMain.ScaleX(lX) * 1.8
    lY = frmMain.ScaleY(lY) * 1.8
    If lPictureBox.Picture = lPic2.Picture Then
        If lX > lPictureBox.Width Or lX < -1 Or lY > lPictureBox.Height Or lY < -1 Then lPictureBox.Picture = lPic1.Picture
    ElseIf lPictureBox.Picture = lPic1.Picture Then
        If lX < lPictureBox.Width And lX > -1 And lY < lPictureBox.Height And lY > -1 Then lPictureBox.Picture = lPic2.Picture
    End If
End If
End Sub
