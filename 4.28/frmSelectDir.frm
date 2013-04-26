VERSION 5.00
Object = "{EE128208-4F73-11D3-83BB-C47C02EE3D01}#1.0#0"; "ControlResizer.ocx"
Begin VB.Form frmSelectDir 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Select Directory"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3150
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelectDir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   224
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   240
      Pattern         =   "*.mp3"
      TabIndex        =   4
      Top             =   3840
      Width           =   2295
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   2880
      Width           =   1095
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   2895
   End
   Begin ControlResizer.AutoResizer AutoResizer1 
      Height          =   735
      Left            =   720
      Top             =   960
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1296
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   208
      Y1              =   185
      Y2              =   185
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   208
      Y1              =   184
      Y2              =   184
   End
End
Attribute VB_Name = "frmSelectDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
On Local Error Resume Next
Unload Me
End Sub

Private Sub cmdOK_Click()
On Local Error Resume Next
lEvents.eRetStr = Dir1.Path
Me.Visible = False
End Sub

Private Sub Dir1_Change()
On Local Error Resume Next
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Local Error Resume Next
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
On Local Error Resume Next
Drive1.Drive = Left(App.Path, 1) & ":"
Dir1.Path = App.Path
End Sub
