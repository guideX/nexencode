VERSION 5.00
Begin VB.Form frmShapeEdit 
   BackColor       =   &H00000000&
   ClientHeight    =   1440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2115
   Icon            =   "frmSkinDisplay.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   96
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   141
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Image imgObject 
      Height          =   615
      Index           =   0
      Left            =   1080
      Top             =   120
      Width           =   735
   End
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
Option Explicit

Private Sub Form_Load()
On Local Error Resume Next
Dim l As Integer, f As Integer
Icon = frmSkinEditor.Icon
InitShapes
If Err.Number <> 0 Then SetError "frmShapeEdit_Load()", lEvents.eSettings.iErrDescription, Err.Description
End Sub
