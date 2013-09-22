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
On Local Error Resume Next
Dim i As Integer, X As Integer

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
    For X = 1 To .sShapeCount
        lShapes.sCount = lShapes.sCount + 1
        lShapes.sShape(X).sIndex = X
        Load frmShapeEdit.shpDisplay(X)
        frmShapeEdit.shpDisplay(X).Visible = True
        frmShapeEdit.shpDisplay(X).Left = .sShape(X).sRgn.X1
        frmShapeEdit.shpDisplay(X).Top = .sShape(X).sRgn.Y1 + 2
        frmShapeEdit.shpDisplay(X).Width = .sShape(X).sRgn.X2 - .sShape(X).sRgn.X1
        frmShapeEdit.shpDisplay(X).Height = .sShape(X).sRgn.Y2 - .sShape(X).sRgn.Y1 + 1
        Select Case .sShape(X).sType
        Case 2
            frmShapeEdit.shpDisplay(X).Shape = 2
        Case 1
            frmShapeEdit.shpDisplay(X).Shape = 0
        Case 3
            frmShapeEdit.shpDisplay(X).Shape = 4
        End Select
    Next X
End With
If Err.Number <> 0 Then SetError "InitShapes()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub RefreshAllShapes()
On Local Error Resume Next

Dim i As Integer, X As Integer
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
If Err.Number <> 0 Then SetError "RefreshAllShapes()", lEvents.eSettings.iErrDescription, Err.Description
End Sub
