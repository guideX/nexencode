VERSION 5.00
Object = "{EE128208-4F73-11D3-83BB-C47C02EE3D01}#1.0#0"; "ControlResizer.ocx"
Begin VB.Form frmInsertDisc 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "NexENCODE - Insert Disc"
   ClientHeight    =   1185
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4185
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInsertDisc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRetry 
      Caption         =   "Retry"
      Default         =   -1  'True
      Height          =   315
      Left            =   3000
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin ControlResizer.AutoResizer AutoResizer1 
      Height          =   255
      Left            =   1560
      Top             =   960
      Visible         =   0   'False
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   450
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   4200
      Y1              =   735
      Y2              =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   0
      X2              =   4200
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      Caption         =   "No tracks or disc information was detected"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   240
      Width           =   3255
   End
   Begin VB.Image imgDisc 
      Height          =   480
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmInsertDisc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
'On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then SetError "cmdCancel_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdHelp_Click()
'On Local Error Resume Next
MsgBox "If you have a disc inserted but another program is accessing your cd drive, end that program, and click retry. If you do not have a cd rom drive click cancel. If you have a disc in the drive and no other programs are accessing your cd drive drive however you haven't yet updated your aspi drivers, run aspiupd.exe which can be found in <nexencode dir>\programs", vbInformation
If Err.Number <> 0 Then SetError "cmdHelp_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdRetry_Click()
'On Local Error Resume Next
Unload Me
LoadTrackGet True
End Sub

Private Sub Form_Load()
'On Local Error Resume Next
imgDisc.Picture = frmGraphics.imgDisc.Picture
If Err.Number <> 0 Then SetError "Form_Unload()", lEvents.eSettings.iErrDescription, Err.Description
End Sub
