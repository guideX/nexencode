Attribute VB_Name = "mdlShapes"
Option Explicit

Enum eShapes
    RectRgn = 1
    EllipticRgn = 2
    RoundRect = 3
End Enum
Private Type gShape
    sType As eShapes
    sVisible As Boolean
    sColor As Integer
    sIndex As Integer
End Type
Private Type gShapes
    sShape(40) As gShape
    sCount As Integer
End Type
Global lShapes As gShapes

Public Sub InitShapes()
Dim i As Integer, x As Integer

For i = 1 To lShapes.sCount
    With lShapes.sShape(i)
        .sColor = 0
        .sIndex = 0
        .sType = 0
        .sVisible = False
    End With
Next i
lShapes.sCount = 0
With lSkins.sSkin(lSkins.sSkinIndex)
    MsgBox .sShapeCount
    For x = 1 To .sShapeCount
        lShapes.sCount = lShapes.sCount + 1
        lShapes.sShape(x).sIndex = x
        Load frmShapeEdit.shpDisplay(x)
        frmShapeEdit.shpDisplay(x).Visible = True
        frmShapeEdit.shpDisplay(x).Left = .sShape(x).sRgn.X1
        frmShapeEdit.shpDisplay(x).Top = .sShape(x).sRgn.Y1 + 2
        frmShapeEdit.shpDisplay(x).Width = .sShape(x).sRgn.X2 - .sShape(x).sRgn.X1
        frmShapeEdit.shpDisplay(x).Height = .sShape(x).sRgn.Y2 - .sShape(x).sRgn.Y1 + 1
        Select Case .sShape(x).sType
        Case 2
            frmShapeEdit.shpDisplay(x).Shape = 2
        Case 1
            frmShapeEdit.shpDisplay(x).Shape = 0
        Case 3
            frmShapeEdit.shpDisplay(x).Shape = 4
        End Select
    Next x
End With
End Sub

Public Sub RefreshAllShapes()
'on local error resume next
Dim i As Integer, x As Integer
For i = 1 To lShapes.sCount
    With lSkins.sSkin(lSkins.sSkinIndex).sShape(i)
        frmShapeEdit.shpDisplay(i).Left = .sRgn.X1
        frmShapeEdit.shpDisplay(i).Top = .sRgn.Y1 + 2
        frmShapeEdit.shpDisplay(i).Width = .sRgn.X2 - .sRgn.X1
        frmShapeEdit.shpDisplay(i).Height = .sRgn.Y2 - .sRgn.Y1 + 1
        Select Case .sType
        Case 2
            If frmShapeEdit.shpDisplay(i).Shape <> 2 Then frmShapeEdit.shpDisplay(i).Shape = 2
        Case 1
            If frmShapeEdit.shpDisplay(i).Shape <> 0 Then frmShapeEdit.shpDisplay(i).Shape = 0
        Case 3
            If frmShapeEdit.shpDisplay(i).Shape <> 4 Then frmShapeEdit.shpDisplay(i).Shape = 4
        End Select
    End With
Next i
End Sub


