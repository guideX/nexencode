VERSION 5.00
Object = "{5EE03E81-8B82-11D1-860C-0020AFE4DE54}#1.0#0"; "Mp3Info.ocx"
Begin VB.Form frmId3Editor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tag Editor"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   435
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   266
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstGenre 
      Height          =   1035
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtAlbum 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2760
      TabIndex        =   9
      Top             =   1440
      Width           =   3135
   End
   Begin VB.TextBox txtArtist 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2760
      TabIndex        =   7
      Top             =   1200
      Width           =   3135
   End
   Begin VB.ComboBox cboFile 
      Height          =   315
      Left            =   2040
      TabIndex        =   4
      Text            =   "cboFile"
      Top             =   600
      Width           =   3855
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   6015
      TabIndex        =   0
      Top             =   3400
      Width           =   6015
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   4680
         TabIndex        =   2
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   3360
         TabIndex        =   1
         Top             =   120
         Width           =   1215
      End
      Begin MP3INFOLib.Mp3Info Mp3Info1 
         Height          =   495
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   495
         _Version        =   65536
         _ExtentX        =   873
         _ExtentY        =   873
         _StockProps     =   0
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Genre:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Album:"
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   1485
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Artist:"
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   1245
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MP3 Filename:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1575
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   136
      X2              =   392
      Y1              =   73
      Y2              =   73
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   136
      X2              =   392
      Y1              =   72
      Y2              =   72
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   408
      X2              =   0
      Y1              =   224
      Y2              =   224
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   0
      Picture         =   "frmId3Editor.frx":0000
      Top             =   0
      Width           =   6000
   End
   Begin VB.Image Image2 
      Height          =   3390
      Left            =   0
      Picture         =   "frmId3Editor.frx":28DA
      Top             =   0
      Width           =   1920
   End
End
Attribute VB_Name = "frmId3Editor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboFile_Click()
Dim i As Integer, msg As String

i = FindMediaIndex(cboFile.Text)
msg = Playlist.pFiles(i).fPath & Playlist.pFiles(i).fFile
If DoesFileExist(msg) = True Then
    Mp3Info1.Open msg
    Caption = Mp3Info1.GetGenreString(Mp3Info1.Genre)
    If Mp3Info1.HasTag = True Then
        txtAlbum.Text = Mp3Info1.Album
        txtArtist.Text = Mp3Info1.Artist
    Else
        MsgBox "No tag present"
    End If
Else
    MsgBox "Doesn't exist"
End If
End Sub

Private Sub Form_Load()
'on local error Resume Next
Dim i As Integer, msg As String

Icon = frmMain.Icon
Mp3Info1.Authorize "Leon Aiossa", "698070606"
For i = 1 To Playlist.pFileCount
    If Len(Playlist.pFiles(i).fFile) <> 0 Then cboFile.AddItem Playlist.pFiles(i).fFile
Next i


'If Err.Number <> 0 Then SetError "Form_Load", lEvents.eSettings.iErrDescription, Err.Description
End Sub

