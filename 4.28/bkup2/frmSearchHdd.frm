VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{EE128208-4F73-11D3-83BB-C47C02EE3D01}#1.0#0"; "ControlResizer.ocx"
Begin VB.Form frmSearchForMedia 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "NexENCODE - Media Search"
   ClientHeight    =   3885
   ClientLeft      =   1860
   ClientTop       =   630
   ClientWidth     =   5970
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSearchHdd.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   360
      ScaleHeight     =   1395
      ScaleWidth      =   7155
      TabIndex        =   7
      Top             =   5520
      Width           =   7215
      Begin VB.ComboBox CmbType 
         Height          =   315
         ItemData        =   "frmSearchHdd.frx":000C
         Left            =   0
         List            =   "frmSearchHdd.frx":0016
         TabIndex        =   9
         Text            =   "Files in Directories"
         Top             =   270
         Width           =   5880
      End
      Begin VB.TextBox TxtFilters 
         Height          =   330
         Left            =   0
         TabIndex        =   8
         Top             =   900
         Visible         =   0   'False
         Width           =   5880
      End
      Begin VB.Label LblFilters 
         Caption         =   "File Filter"
         Height          =   240
         Left            =   0
         TabIndex        =   11
         Top             =   630
         Visible         =   0   'False
         Width           =   5760
      End
      Begin VB.Label LblType 
         Caption         =   "Search Type"
         Height          =   240
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   5880
      End
   End
   Begin VB.CheckBox ChkSubDirectorys 
      Caption         =   "&Include Sub-Directories"
      Height          =   330
      Left            =   3960
      TabIndex        =   6
      Top             =   3120
      Width           =   1995
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2085
      Left            =   45
      TabIndex        =   5
      Top             =   0
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   3678
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Path"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "DateTime"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Attr"
         Object.Width           =   882
      EndProperty
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "E&xit"
      Height          =   330
      Left            =   4680
      TabIndex        =   2
      Top             =   3540
      Width           =   1275
   End
   Begin VB.TextBox TxtPaths 
      Height          =   330
      Left            =   45
      TabIndex        =   0
      Top             =   2775
      Width           =   5880
   End
   Begin VB.CommandButton CmdSearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   330
      Left            =   3360
      TabIndex        =   1
      Top             =   3540
      Width           =   1275
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5880
      Y1              =   2400
      Y2              =   2400
   End
   Begin ControlResizer.AutoResizer AutoResizer1 
      Height          =   375
      Left            =   1800
      Top             =   4920
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
   End
   Begin VB.Label LblStatus 
      Alignment       =   1  'Right Justify
      Caption         =   "Total: 0 Results"
      Height          =   240
      Left            =   45
      TabIndex        =   4
      Top             =   2160
      Width           =   5880
   End
   Begin VB.Label LblPaths 
      Caption         =   "Search in Path"
      Height          =   240
      Left            =   45
      TabIndex        =   3
      Top             =   2505
      Width           =   5880
   End
End
Attribute VB_Name = "frmSearchForMedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
On Local Error Resume Next
CmbType.ListIndex = 1
TxtPaths.Text = lDrives.dHardDrives
ChkSubDirectorys.Value = 1
'TxtFilters.Text = "*.mp3;*.au;*.snd;*.aif;*.aifc;*.aiff;*.mid;*.midi;*.rmi;*.mpeg;*.mpg;*.mpe;*.m1v;*.mp1;*.mp2;*.mpa;*.avi;*.wm;*.wma;*.wmv;*.vob;*.wmx;*.wax;*.asf;*.asx;*.wmp"
TxtFilters.Text = "*.mp3"
If lAutoScanHdd = True Then cmdSearch_Click
If Err.Number <> 0 Then SetError "frmSearchForMedia_Load", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub CmbType_Click()
On Local Error Resume Next
TxtFilters.Visible = CmbType.ListIndex = 1
LblFilters.Visible = CmbType.ListIndex = 1
If Err.Number <> 0 Then SetError "cmbType_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdSearch_Click()
On Local Error Resume Next
Dim i As Integer, lTmp1 As Long, sStr1 As String, lItem As ListItem, cCol As tSearch
cmdSearch.Enabled = False
Me.MousePointer = vbHourglass
ListView1.ListItems.Clear
lblStatus.Alignment = vbLeftJustify
If CmbType.ListIndex = 0 Then
    If ChkSubDirectorys.Value Then
        lblStatus.Caption = "Please Wait, Searching Sub-Directories..."
        GetSubDirs TxtPaths.Text, vbDirectory, cCol
    Else
        lblStatus.Caption = "Please Wait, Searching Directories..."
        GetDirs TxtPaths.Text, vbDirectory, cCol
    End If
Else
    If ChkSubDirectorys.Value Then
        lblStatus.Caption = "Please Wait, Searching Sub-Files..."
        GetSubFiles TxtPaths.Text, TxtFilters.Text, vbDirectory, vbArchive, cCol
    Else
        lblStatus.Caption = "Please Wait, Searching Files..."
        GetFiles TxtPaths.Text, TxtFilters.Text, vbArchive, cCol
    End If
End If
i = AddPlaylist("Search Results", "search.m3u")
For lTmp1 = 1 To cCol.Count
    AddToPlaylist cCol.Path(lTmp1), i
    Set lItem = ListView1.ListItems.Add(, , cCol.Path(lTmp1))
    lItem.SubItems(1) = Format(cCol.Size(lTmp1), "###,###,##0")
    lItem.SubItems(2) = Format(cCol.DateTime(lTmp1), "DD-MM-YY HH:MM:SS")
    lItem.SubItems(3) = sAttr(cCol.Attr(lTmp1))
Next
frmPlaylist.SortByGenre i
frmPlaylist.cboPlaylists.ListIndex = 1
Unload frmPlaylist
lblStatus.Alignment = vbRightJustify
lblStatus.Caption = "Total: " & ListView1.ListItems.Count & " Results"
Me.MousePointer = vbDefault
cmdSearch.Enabled = False
If lAutoScanHdd = True Then
    lAutoScanHdd = False
    Unload Me
End If
If Err.Number <> 0 Then SetError "cmdSearch_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdExit_Click()
On Local Error Resume Next
Unload Me
End Sub

Function sAttr(Attr As VbFileAttribute) As String
On Local Error Resume Next
Dim sStr1 As String
sStr1 = ""
If Attr And vbReadOnly Then sStr1 = "r" Else sStr1 = "-"
If Attr And vbArchive Then sStr1 = sStr1 + "a" Else sStr1 = sStr1 + "-"
If Attr And vbHidden Then sStr1 = sStr1 + "h" Else sStr1 = sStr1 + "-"
If Attr And vbSystem Then sStr1 = sStr1 + "s" Else sStr1 = sStr1 + "-"
sAttr = sStr1
If Err.Number <> 0 Then SetError "sAttr", lEvents.eSettings.iErrDescription, Err.Description
End Function
