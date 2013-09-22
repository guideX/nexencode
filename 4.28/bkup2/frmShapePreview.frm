VERSION 5.00
Begin VB.Form frmShapePreview 
   BackColor       =   &H00000000&
   Caption         =   "Shape Preview"
   ClientHeight    =   1725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2880
   LinkTopic       =   "Form1"
   ScaleHeight     =   115
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   192
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frmShapePreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
On Local Error Resume Next
Icon = frmSkinEditor.Icon
Me.Picture = LoadPicture(lSkins.sSkin(lSkins.sSkinIndex).sFilepath & lSkins.sSkin(lSkins.sSkinIndex).sGraphic)
If Err.Number <> 0 Then SetError "frmShapePreview_Load()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub
