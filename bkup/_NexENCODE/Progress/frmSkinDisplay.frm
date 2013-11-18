VERSION 5.00
Begin VB.Form frmShapeEdit 
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmSkinDisplay.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   Begin VB.Shape shpDisplay 
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   0
      Left            =   120
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmShapeEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim l As Integer, f As Integer
Icon = frmSkinEditor.Icon
MsgBox shpDisplay.Count
MsgBox lSkins.sSkin(lSkins.sSkinIndex).sShapeCount
InitShapes
End Sub
