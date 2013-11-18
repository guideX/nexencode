VERSION 5.00
Begin VB.Form frmPlayer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audica"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   Icon            =   "frmPlayer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   349
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   375
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SetPlayerShape()
Dim i As Integer

Dim rgn As Long, rgn1 As Long, rgn2 As Long, rgn3 As Long, rgn4 As Long, tmp As Long
Dim X As Long, Y As Long
X = lMainWndSettings.wWindowBorder
Y = lMainWndSettings.wTitleBarHeight
rgn1 = CreateEllipticRgn(X + 102, Y + 87, X + 102 + 157, Y + 87 + 163)
rgn2 = CreateEllipticRgn(X + 113, Y + 232, X + 112 + 159, Y + 391)
rgn3 = CreateEllipticRgn(X + 87, Y + 162, X + 87 + 160, Y + 315)
rgn4 = CreateEllipticRgn(X + 175, Y + 162, X + 155 + 168, Y + 313)
tmp = CombineRgn(rgn1, rgn1, rgn2, RGN_OR)
tmp = CombineRgn(rgn1, rgn1, rgn3, RGN_OR)
tmp = CombineRgn(rgn1, rgn1, rgn4, RGN_OR)
tmp = SetWindowRgn(Me.hWnd, rgn1, True)
End Sub

Private Sub Form_Load()

Icon = frmMain.Icon
SetPlayerShape
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SetPlayerShape
FormDrag Me
End Sub
