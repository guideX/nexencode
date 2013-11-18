VERSION 5.00
Begin VB.Form frmSkinBrowser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NexENCODE - Skin Browser"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSkinBrowser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   5985
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "&Modify"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.ListBox lstSkins 
      Appearance      =   0  'Flat
      Height          =   1980
      ItemData        =   "frmSkinBrowser.frx":08CA
      Left            =   1800
      List            =   "frmSkinBrowser.frx":08CC
      TabIndex        =   0
      Top             =   600
      Width           =   4095
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   6000
      X2              =   -120
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   1680
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   0
      Picture         =   "frmSkinBrowser.frx":08CE
      Top             =   0
      Width           =   6000
   End
   Begin VB.Image Image2 
      Height          =   3390
      Left            =   0
      Picture         =   "frmSkinBrowser.frx":31A8
      Top             =   -630
      Width           =   1920
   End
End
Attribute VB_Name = "frmSkinBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
Me.Visible = False
frmSkinEditor.Show

End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub

