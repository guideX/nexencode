VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSearchPlaylists 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "NexENCODE - Search playlists"
   ClientHeight    =   1020
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5250
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSearchPlaylists.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   68
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSearchHDD 
      Caption         =   "Search HDD"
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play/Close"
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdSaveAsPlaylist 
      Caption         =   "Save As"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txtSearchFor 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   3495
   End
   Begin MSComctlLib.ListView lvwResults 
      Height          =   1455
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   2566
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Line Line3 
      X1              =   256
      X2              =   256
      Y1              =   200
      Y2              =   224
   End
   Begin VB.Label Label3 
      Caption         =   "Results:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   0
      X2              =   408
      Y1              =   73
      Y2              =   73
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   0
      X2              =   408
      Y1              =   71
      Y2              =   71
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   120
      Picture         =   "frmSearchPlaylists.frx":000C
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Search For:"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmSearchPlaylists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdHelp_Click()
On Local Error Resume Next
MsgBox "Enter a search string, such as 'The Beatles'", vbInformation
End Sub

Private Sub cmdPlay_Click()
On Local Error Resume Next
Dim mItem As ListItem, i As Integer
For i = 1 To lvwResults.ListItems.Count
    Set mItem = lvwResults.ListItems(i)
    AddEvent Play, lvwResults.ListItems(i).SubItems(3), lvwResults.ListItems(i).Text, "", "", 0, ""
Next i
Unload Me
If Err.Number <> 0 Then SetError "cmdPlay_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdSaveAsPlaylist_Click()
On Local Error Resume Next
Dim msg As String, lFile As String, mItem As ListItem, i As Integer
For i = 1 To lvwResults.ListItems.Count
    Set mItem = lvwResults.ListItems(i)
    msg = mItem.SubItems(3) & mItem.Text & vbCrLf & msg
Next i
SaveAsPlaylist msg
If Err.Number <> 0 Then SetError "cmdSaveAsPlaylist_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdSearch_Click()
On Local Error Resume Next
Dim i As Integer, X As Integer, lListItem As ListItem, msg As String
If cmdSearch.Caption = "Search" Then
    If Len(txtSearchFor.Text) = 0 Then
        If lEvents.eSettings.iOverwritePrompts = True Then MsgBox "Enter a search string before pressing 'Search'", vbInformation
        txtSearchFor.SetFocus
        Beep
        Exit Sub
    End If
    For i = 1 To Playlist.pFileCount
        If InStr(LCase(Playlist.pFiles(i).fFile), LCase(txtSearchFor.Text)) And Len(Playlist.pFiles(i).fFile) <> 0 Then
            lTag.tFile = Playlist.pFiles(i).fPath & Playlist.pFiles(i).fFile
            GetTagInfo
            DoEvents
            Set lListItem = lvwResults.ListItems.Add(, , Playlist.pFiles(i).fFile)
            lListItem.SubItems(1) = lTag.tArtist
            lListItem.SubItems(2) = lTag.tTitle
            lListItem.SubItems(3) = Playlist.pFiles(i).fPath
        End If
    Next i
    Me.Height = 3825
    cmdSearch.Caption = "New Search"
ElseIf cmdSearch.Caption = "New Search" Then
    Unload Me
    frmSearchPlaylists.Show
End If
If Err.Number <> 0 Then SetError "cmdSearch_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdSearchHDD_Click()
On Local Error Resume Next
frmSearchForMedia.Show
Unload Me
End Sub

Private Sub Command2_Click()
On Local Error Resume Next
Unload Me
End Sub

Private Sub Form_Load()
On Local Error Resume Next
lvwResults.ColumnHeaders.Add , , "File", 160
lvwResults.ColumnHeaders.Add , , "Artist", 60
lvwResults.ColumnHeaders.Add , , "Title", 60
lvwResults.ColumnHeaders.Add , , "Path", 60
If Err.Number <> 0 Then SetError "Form_load()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub lvwResults_DblClick()
On Local Error Resume Next
Dim msg As String, lFile As String, mItem As ListItem, i As Integer
Set mItem = lvwResults.SelectedItem
msg = mItem.SubItems(3) & mItem.Text
i = FindMediaIndex(GetFileTitle(msg))
If Playlist.pFiles(i).fEnabled = True Then
    If lPlayer.pStatus = sPlaying Then StopMp3
    AddEvent Play, Playlist.pFiles(i).fPath, Playlist.pFiles(i).fFile, "", "", 0, ""
End If
If Err.Number <> 0 Then SetError "lvwResults_DBLClick()", lEvents.eSettings.iErrDescription, Err.Description
End Sub
