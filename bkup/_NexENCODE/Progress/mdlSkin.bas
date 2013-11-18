Attribute VB_Name = "mdlSkin"
Option Explicit

Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Enum eCombineMode
    cRgn_None = 0
    cRgn_And = 1
    cRgn_Or = 2
    cRgn_XOr = 3
    cRgn_Diff = 4
    cRgn_Copy = 5
End Enum

Enum eObjectTypes
    oOther = 0
    oButton = 1
    oLabel = 2
    oProgress = 3
End Enum

Enum eShapeTypes
    sOther = 0
    sRectRgn = 1
    sEllipce = 2
    sRoundRectRgn = 3
End Enum

Private Type gRegions
    rRgn As Long
    X1 As Long
    X2 As Long
    X3 As Long
    Y1 As Long
    Y2 As Long
    Y3 As Long
End Type

Private Type gWindowPos
    wTitleBarHeight As Integer
    wWindowBorder As Integer
End Type

Private Type gSkinSettings
    sWidth As Long
    sHeight As Long
    sLeft As Long
    sTop As Long
End Type

Private Type gObject
    bName As String
    bType As eObjectTypes
    bScript As String
    
    bWidth As Long
    bHeight As Long
    bLeft As Long
    bTop As Long
    
    bFilename As String
    bFilename2 As String
    
    bEnabled As Boolean
End Type

Private Type gShape
    sName As String
    sType As eShapeTypes
    sRgn As gRegions
    sCombineMode As eCombineMode
    sDestRgn As Integer
    sSrcRgn1 As Integer
    sSrcRgn2 As Integer
    sEnabled As Boolean
End Type

Private Type gSkin
    sAuthor As String
    sEnabled As Boolean
    sName As String
    sShape(20) As gShape
    sObject(20) As gObject
    sSkinSettings As gSkinSettings
    sShapeCount As Integer
    sObjectCount As Integer
    sFilename As String
    sFilepath As String
    sGraphic As String
End Type

Private Type gSkins
    sSkinIndex As Integer
    sSkin(15) As gSkin
    sCount As Integer
End Type

Global lSkins As gSkins
Global lMainWndSettings As gWindowPos
Global SkinAuthor As String

Public Function OpenSkin(lFilename As String, lSelectContainer As Boolean) As Integer
''on local error resume next
Dim i As Integer, x As Integer, msg As String, f As Integer, msg2 As String
msg2 = lFilename
If Len(lFilename) <> 0 Then
    With lSkins
        i = lSkins.sCount + 1
        .sCount = i
        .sSkin(i).sSkinSettings.sWidth = ReadINI(lFilename, "Settings", "Width", 200)
        .sSkin(i).sSkinSettings.sHeight = ReadINI(lFilename, "Settings", "Height", 200)
        .sSkin(i).sSkinSettings.sLeft = ReadINI(lFilename, "Settings", "Left", 200)
        .sSkin(i).sSkinSettings.sTop = ReadINI(lFilename, "Settings", "Top", 200)
        .sSkin(i).sGraphic = ReadINI(lFilename, "Settings", "Graphic", "")
        .sSkin(i).sName = ReadINI(lFilename, "Settings", "Name", "Default Skin")
        frmSkinEditor.lstSkins.AddItem .sSkin(i).sName
        .sSkin(i).sShapeCount = ReadINI(lFilename, "Settings", "ShapeCount", 0)
        .sSkin(i).sAuthor = ReadINI(lFilename, "Settings", "Author", "")
        .sSkin(i).sFilename = GetFileTitle(msg2)
        .sSkin(i).sFilepath = Left(lFilename, Len(lFilename) - Len(.sSkin(i).sFilename))
        If .sSkin(i).sShapeCount <> 0 Then
            For x = 1 To .sSkin(i).sShapeCount
                msg = "rgn" & x
                .sSkin(i).sShape(x).sEnabled = ReadINI(lFilename, msg, "enabled", "")
                .sSkin(i).sShape(x).sName = ReadINI(lFilename, msg, "name", "")
                .sSkin(i).sShape(x).sDestRgn = ReadINI(lFilename, msg, "destrgn", 0)
                .sSkin(i).sShape(x).sSrcRgn1 = ReadINI(lFilename, msg, "srcrgn1", 0)
                .sSkin(i).sShape(x).sSrcRgn2 = ReadINI(lFilename, msg, "srcrgn2", 0)
                .sSkin(i).sShape(x).sCombineMode = ReadINI(lFilename, msg, "combinemode", 0)
                .sSkin(i).sShape(x).sRgn.X1 = ReadINI(lFilename, msg, "x1", 0)
                .sSkin(i).sShape(x).sRgn.X2 = ReadINI(lFilename, msg, "x2", 0)
                .sSkin(i).sShape(x).sRgn.X3 = ReadINI(lFilename, msg, "x3", 0)
                .sSkin(i).sShape(x).sRgn.Y1 = ReadINI(lFilename, msg, "y1", 0)
                .sSkin(i).sShape(x).sRgn.Y2 = ReadINI(lFilename, msg, "y2", 0)
                .sSkin(i).sShape(x).sRgn.Y3 = ReadINI(lFilename, msg, "y3", 0)
                .sSkin(i).sShape(x).sType = ReadINI(lFilename, msg, "type", 1)
'                MsgBox "Rgn info for: " & msg & vbCrLf & "Skin Index: " & i & vbCrLf & "Shape Index: " & i & vbCrLf & "Shape Type: " & .sSkin(i).sShape(x).sType
            Next x
        End If
    End With
    OpenSkin = i
End If
End Function

Public Function NewSkin() As String
Dim lFilename As String, lTitle As String
lTitle = InputBox("Enter title of new skin:")
If Len(lTitle) <> 0 Then
    lFilename = SaveDialog(frmSkinEditor, "Storage Containers (*.ns4)|*.ns4|All Files (*.*)|*.*", "Save as ...", App.Path)
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
End Function

Public Sub SetSkin(lIndex As Integer)
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
End Sub

Public Function DoesFileExist(lFilename As String) As Boolean
'on local error resume next

Dim msg As String
msg = Dir(lFilename)
If msg <> "" Then
    DoesFileExist = True
Else
    DoesFileExist = False
End If
End Function

Public Sub GetWindowSettings(lHandle As Long)
Dim lWindowPos As RECT, lClientPos As RECT
Dim lBorderWidth As Long, lTopOffset As Long
Dim i As Long
i = GetWindowRect(lHandle, lWindowPos)
i = GetClientRect(lHandle, lClientPos)
lMainWndSettings.wTitleBarHeight = lWindowPos.Bottom - lWindowPos.Top - lClientPos.Bottom - lBorderWidth
lMainWndSettings.wWindowBorder = lWindowPos.Right - lWindowPos.Left - lClientPos.Right - 2
End Sub

Public Function FindSkinIndex(lName As String) As Integer
Dim i As Integer

If Len(lName) <> 0 Then
    For i = 1 To lSkins.sCount
        If LCase(lName) = LCase(lSkins.sSkin(i).sName) Then FindSkinIndex = i
    Next i
End If
End Function

Public Sub LoadShape(lForm As Form, lSkinIndex As Integer)
'on local error resume next
Dim i As Integer, x As Long, Y As Long, tmp As Long

GetWindowSettings lForm.hwnd
x = lMainWndSettings.wWindowBorder
Y = lMainWndSettings.wTitleBarHeight

With lSkins.sSkin(lSkins.sSkinIndex)
    For i = 1 To .sShapeCount
        If .sShape(i).sEnabled = True Then
            Select Case .sShape(i).sType
            Case 1
                .sShape(i).sRgn.rRgn = CreateRectRgn(x + .sShape(i).sRgn.X1, Y + .sShape(i).sRgn.Y1, x + .sShape(i).sRgn.X2, Y + .sShape(i).sRgn.Y2)
            Case 2
                .sShape(i).sRgn.rRgn = CreateEllipticRgn(x + .sShape(i).sRgn.X1, Y + .sShape(i).sRgn.Y1, x + .sShape(i).sRgn.X2, Y + .sShape(i).sRgn.Y2)
            Case 3
                .sShape(i).sRgn.rRgn = CreateRoundRectRgn(x + .sShape(i).sRgn.X1, Y + .sShape(i).sRgn.Y1, x + .sShape(i).sRgn.X2, Y + .sShape(i).sRgn.Y2, .sShape(i).sRgn.X3, .sShape(i).sRgn.Y3)
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
End Sub

Public Sub SetShape()
Dim i As Integer

Dim rgn As Long, rgn1 As Long, rgn2 As Long, rgn3 As Long, rgn4 As Long, rgn5 As Long, rgn6 As Long, rgn7 As Long, tmp As Long
Dim x As Long, Y As Long

x = lMainWndSettings.wWindowBorder
Y = lMainWndSettings.wTitleBarHeight

rgn = CreateEllipticRgn(0, 0, frmMain.Width, frmMain.Height) ' whole image
rgn1 = CreateEllipticRgn(x + 147, Y + 90, x + 326, Y + 267) 'big crl in back (transparency)
rgn2 = CreateEllipticRgn(x + 104, Y + 46, x + 367, Y + 310) ' big crl in back
rgn3 = CreateEllipticRgn(x + 48, Y + 74, x + 257, Y + 287) 'left crl
rgn4 = CreateEllipticRgn(x + 65, Y + 92, x + 241, Y + 268) 'left crl (transperancy)
rgn5 = CreateEllipticRgn(x + 212, Y + 72, x + 421, Y + 285) 'right crl
rgn6 = CreateEllipticRgn(x + 230, Y + 91, x + 404, Y + 268) 'right crl (transparency)
rgn7 = CreateRoundRectRgn(x + 39, Y + 120, x + 429, Y + 237, 110, 110)  'pill

tmp = CombineRgn(rgn1, rgn2, rgn1, RGN_DIFF) ' back crl
tmp = CombineRgn(rgn3, rgn3, rgn4, RGN_DIFF) ' left crl
tmp = CombineRgn(rgn5, rgn5, rgn6, RGN_DIFF) 'right crl

tmp = CombineRgn(rgn3, rgn3, rgn5, RGN_OR)
tmp = CombineRgn(rgn1, rgn1, rgn7, RGN_OR)
tmp = CombineRgn(rgn, rgn1, rgn3, RGN_OR)
tmp = SetWindowRgn(frmMain.hwnd, rgn, True)
End Sub
