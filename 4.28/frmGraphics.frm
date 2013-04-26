VERSION 5.00
Begin VB.Form frmGraphics 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Graphics (Should be hidden)"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6000
   Icon            =   "frmGraphics.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Image picNS4 
      Height          =   225
      Left            =   120
      Top             =   1200
      Width           =   345
   End
   Begin VB.Image imgDisc 
      Height          =   480
      Left            =   120
      Picture         =   "frmGraphics.frx":08CA
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "frmGraphics.frx":0D0C
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image imgTopper 
      Height          =   255
      Left            =   120
      Top             =   480
      Width           =   360
   End
   Begin VB.Image imgPlaylist 
      Height          =   255
      Left            =   120
      Top             =   840
      Width           =   375
   End
   Begin VB.Image imgSideGradient 
      Height          =   225
      Left            =   120
      Top             =   120
      Width           =   330
   End
End
Attribute VB_Name = "frmGraphics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
