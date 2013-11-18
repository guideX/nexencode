VERSION 5.00
Begin VB.Form frmImagePreview 
   Caption         =   "Image preview"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3375
   ScaleWidth      =   4320
   Begin VB.CommandButton cmdChange 
      Caption         =   "&Change"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   120
      ScaleHeight     =   3375
      ScaleWidth      =   4320
      TabIndex        =   0
      Top             =   120
      Width           =   4320
   End
End
Attribute VB_Name = "frmImagePreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdChange_Click()
'on local error resume next
Dim lFilename As String, msg As String, msg2 As String
lFilename = OpenDialog(frmImagePreview, "Gif Files (*.gif)|*.gif|Bitmaps (*.bmp)|*.bmp|Jpg Files (*.jpg)|*.jpg|All Files (*.*)|*.*", "Select graphic ...", App.Path)
If Len(lFilename) <> 0 Then
    msg2 = lFilename
    msg = GetFileTitle(msg2)
    If DoesFileExist(App.Path & "\" & msg2) = False Then FileCopy lFilename, lSkins.sSkin(lSkins.sSkinIndex).sFilepath & msg
    lSkins.sSkin(lSkins.sSkinIndex).sGraphic = msg
    WriteINI lSkins.sSkin(lSkins.sSkinIndex).sFilepath & lSkins.sSkin(lSkins.sSkinIndex).sFilename, "Settings", "Graphic", msg
    Picture1.Picture = LoadPicture(lSkins.sSkin(lSkins.sSkinIndex).sFilepath & lSkins.sSkin(lSkins.sSkinIndex).sGraphic)
End If
End Sub

Private Sub Form_Load()
With lSkins.sSkin(lSkins.sSkinIndex)
    
End With
End Sub

Private Sub Picture1_Resize()
Me.Width = Picture1.Width + 400
Me.Height = Picture1.Height + 600
End Sub
