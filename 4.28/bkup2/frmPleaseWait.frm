VERSION 5.00
Begin VB.Form frmPleaseWait 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "NexENCODE"
   ClientHeight    =   840
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3120
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPleaseWait.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   840
   ScaleWidth      =   3120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Caption         =   "Please wait while NexENCODE completes a task ..."
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   120
      Picture         =   "frmPleaseWait.frx":000C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmPleaseWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
On Local Error Resume Next
AlwaysOnTop Me, True
If Err.Number <> 0 Then SetError "Form_Load", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Local Error Resume Next
AlwaysOnTop Me, False
If Err.Number <> 0 Then SetError "Form_Unload", lEvents.eSettings.iErrDescription, Err.Description
End Sub
