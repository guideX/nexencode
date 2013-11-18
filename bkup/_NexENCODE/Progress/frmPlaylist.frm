VERSION 5.00
Begin VB.Form frmPlaylist 
   Caption         =   "Playlist Browser"
   ClientHeight    =   6795
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6075
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
   Picture         =   "frmPlaylist.frx":0000
   ScaleHeight     =   453
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   405
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   120
      ScaleHeight     =   915
      ScaleWidth      =   5835
      TabIndex        =   1
      Top             =   6120
      Width           =   5895
      Begin VB.Image imgExit2 
         Height          =   180
         Left            =   1440
         Picture         =   "frmPlaylist.frx":48EC
         Top             =   240
         Width           =   615
      End
      Begin VB.Image imgExit1 
         Height          =   180
         Left            =   1440
         Picture         =   "frmPlaylist.frx":4E54
         Top             =   0
         Width           =   615
      End
      Begin VB.Image imgDel2 
         Height          =   180
         Left            =   720
         Picture         =   "frmPlaylist.frx":5036
         Top             =   240
         Width           =   615
      End
      Begin VB.Image imgDel1 
         Height          =   180
         Left            =   720
         Picture         =   "frmPlaylist.frx":552E
         Top             =   0
         Width           =   615
      End
      Begin VB.Image imgCopy2 
         Height          =   210
         Left            =   0
         Picture         =   "frmPlaylist.frx":56EA
         Top             =   240
         Width           =   615
      End
      Begin VB.Image imgCopy1 
         Height          =   210
         Left            =   0
         Picture         =   "frmPlaylist.frx":5C7B
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.ListBox lstTracks 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      IntegralHeight  =   0   'False
      Left            =   2475
      TabIndex        =   0
      Top             =   4020
      Width           =   2220
   End
   Begin VB.Image imgExit 
      Height          =   180
      Left            =   4845
      Picture         =   "frmPlaylist.frx":5F74
      Top             =   4560
      Width           =   615
   End
   Begin VB.Image imgDel 
      Height          =   180
      Left            =   4845
      Picture         =   "frmPlaylist.frx":6156
      Top             =   4200
      Width           =   615
   End
   Begin VB.Image imgCopy 
      Height          =   210
      Left            =   4845
      Picture         =   "frmPlaylist.frx":6312
      Top             =   4005
      Width           =   615
   End
   Begin VB.Shape Shape1 
      Height          =   1245
      Left            =   2325
      Top             =   3675
      Width           =   3315
   End
End
Attribute VB_Name = "frmPlaylist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SetPlaylistShape()
Dim i As Integer

Dim rgn As Long, rgn1 As Long, rgn2 As Long, rgn3 As Long, rgn4 As Long, rgn5 As Long, rgn6 As Long, rgn7 As Long, rgn8 As Long, tmp As Long
Dim x As Long, Y As Long
Dim rgn9 As Long

x = lMainWndSettings.wWindowBorder
Y = lMainWndSettings.wTitleBarHeight

rgn1 = CreateEllipticRgn(x + 96, Y + 93, x + 304, Y + 305)
rgn2 = CreateEllipticRgn(x + 154, Y + 149, x + 247, Y + 244)
rgn3 = CreateEllipticRgn(x + 114, Y + 109, x + 287, Y + 288)
rgn4 = CreateRectRgn(x + 90, Y + 141, x + 310, Y + 252)
rgn5 = CreateRectRgn(x + 155, Y + 245, x + 376, Y + 324)

tmp = CombineRgn(rgn1, rgn3, rgn1, RGN_OR)
tmp = CombineRgn(rgn1, rgn1, rgn3, RGN_DIFF)
tmp = CombineRgn(rgn1, rgn1, rgn4, RGN_DIFF)
tmp = CombineRgn(rgn1, rgn1, rgn2, RGN_OR)
tmp = CombineRgn(rgn1, rgn1, rgn5, RGN_OR)

tmp = SetWindowRgn(Me.hwnd, rgn1, True)
End Sub

Private Sub Form_DblClick()
SetPlaylistShape
End Sub

Private Sub Form_Load()
Dim i As Integer, z As Integer
Icon = frmMain.Icon
SetPlaylistShape

Top = GetSetting(App.Title, "frmTrack", "Top", Top)
Left = GetSetting(App.Title, "frmTrack", "Top", Left)
lstTracks.Clear
With frmMain
    .Ripper.Init: DoEvents
    .Ripper.OpenDriveByNumber 1
    z = .Ripper.TrackCount
    For i = 1 To .Ripper.TrackCount
        lstTracks.AddItem "Track " & i
    Next i
End With
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
FormDrag Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSetting App.Title, "frmTrack", "Top", Top
SaveSetting App.Title, "frmTrack", "Left", Left
End Sub

Private Sub imgCopy_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 Then
    imgCopy.Picture = imgCopy2.Picture
End If
End Sub

Private Sub imgCopy_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim i As Integer, msg As String
If Button = 1 Then
    imgCopy.Picture = imgCopy1.Picture
    For i = 1 To lstTracks.ListCount
        i = Int(Right(lstTracks.List(0), Len(lstTracks.List(0)) - 6))
        lstTracks.ListIndex = 0
        lstTracks.RemoveItem 0
        AddEvent Rip, "", "", App.Path, "track " & Str(i) & ".wav", i, ""
        
    Next i
    Unload Me
End If
End Sub

Private Sub imgDel_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 Then
    imgDel.Picture = imgDel2.Picture
End If
End Sub

Private Sub imgDel_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
'on local error resume next
Dim i As Integer
If Button = 1 Then
    imgDel.Picture = imgDel1.Picture
    If lstTracks.ListIndex <> -1 Then
        i = lstTracks.ListIndex
        lstTracks.RemoveItem i
        lstTracks.ListIndex = i
    End If
End If
End Sub

Private Sub imgExit_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 Then
    imgExit.Picture = imgExit2.Picture
End If
End Sub

Private Sub imgExit_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 Then
    imgExit.Picture = imgExit1.Picture
    Unload Me
End If
End Sub

Private Sub lstTracks_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If lstTracks.ListCount = 0 Then
    Form_Load
End If
End Sub
