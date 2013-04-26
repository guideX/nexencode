VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{EE128208-4F73-11D3-83BB-C47C02EE3D01}#1.0#0"; "ControlResizer.ocx"
Begin VB.Form frmAddEvent 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "NexENCODE - Batch Events"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   330
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
   Icon            =   "frmAddEvent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   398
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDelLvw 
      Caption         =   "< Del"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "Click to remove event"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdAddlvw 
      Caption         =   "Add >"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      ToolTipText     =   "Click to add event"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.ComboBox cboCDAudioTrack 
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   12
      ToolTipText     =   "CD Audio Track"
      Top             =   1200
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select ..."
      Height          =   255
      Left            =   4800
      TabIndex        =   11
      ToolTipText     =   "Select output file"
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtOutputFile 
      Height          =   285
      Left            =   2040
      TabIndex        =   10
      ToolTipText     =   "File to write to"
      Top             =   840
      Width           =   2655
   End
   Begin VB.CommandButton cmdSelectInputFile 
      Caption         =   "Select ..."
      Height          =   255
      Left            =   4800
      TabIndex        =   7
      ToolTipText     =   "Select Input file"
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtInputFile 
      Height          =   285
      Left            =   2040
      TabIndex        =   6
      ToolTipText     =   "File to read from"
      Top             =   480
      Width           =   2655
   End
   Begin VB.ComboBox cboEventType 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmAddEvent.frx":000C
      Left            =   2040
      List            =   "frmAddEvent.frx":001F
      Style           =   2  'Dropdown List
      TabIndex        =   4
      ToolTipText     =   "Represents what kind of event you want to execute"
      Top             =   120
      Width           =   3855
   End
   Begin VB.PictureBox picBottom 
      BackColor       =   &H00E0E0E0&
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
      Top             =   2760
      Width           =   6015
      Begin VB.CommandButton cmdHelp 
         Caption         =   "Help"
         Height          =   315
         Left            =   120
         TabIndex        =   16
         ToolTipText     =   "Description of objects"
         Top             =   75
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   315
         Left            =   4920
         TabIndex        =   2
         Top             =   75
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Do Batch"
         Default         =   -1  'True
         Height          =   315
         Left            =   3840
         TabIndex        =   1
         ToolTipText     =   "Click to execute"
         Top             =   75
         Width           =   975
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   4
         X1              =   0
         X2              =   6000
         Y1              =   0
         Y2              =   0
      End
   End
   Begin MSComctlLib.ListView lvwEvents 
      Height          =   975
      Left            =   2040
      TabIndex        =   15
      ToolTipText     =   "Que of events"
      Top             =   1680
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1720
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
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Line Line1 
      X1              =   136
      X2              =   392
      Y1              =   104
      Y2              =   104
   End
   Begin VB.Label lblTrackNum 
      BackStyle       =   0  'Transparent
      Caption         =   "Track:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblOutputFile 
      BackStyle       =   0  'Transparent
      Caption         =   "Output File:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblInputFile 
      BackStyle       =   0  'Transparent
      Caption         =   "Input File:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblEventType 
      BackStyle       =   0  'Transparent
      Caption         =   "Event Type:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin ControlResizer.AutoResizer AutoResizer1 
      Height          =   375
      Left            =   5040
      Top             =   0
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
   End
End
Attribute VB_Name = "frmAddEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboEventType_Click()
On Local Error Resume Next
txtInputFile.Text = ""
txtOutputFile.Text = ""
Select Case cboEventType.ListIndex
Case 0
    lblTrackNum.Visible = False
    cboCDAudioTrack.Visible = False
    txtOutputFile.Visible = True
    txtOutputFile.Visible = True
    txtInputFile.Visible = True
    cmdSelectInputFile.Visible = True
Case 1
    cboCDAudioTrack.Visible = True
    lblTrackNum.Visible = True
    txtOutputFile.Visible = True
    txtOutputFile.Visible = True
    txtInputFile.Visible = False
    cmdSelectInputFile.Visible = False
Case 2
    cboCDAudioTrack.Visible = True
    lblTrackNum.Visible = True
    txtOutputFile.Visible = True
    txtOutputFile.Visible = True
    txtInputFile.Visible = False
    cmdSelectInputFile.Visible = False
Case 3
    cboCDAudioTrack.Visible = False
    lblTrackNum.Visible = False
    txtOutputFile.Visible = True
    txtOutputFile.Visible = True
    txtInputFile.Visible = True
    cmdSelectInputFile.Visible = True
Case 4
    cboCDAudioTrack.Visible = False
    lblTrackNum.Visible = False
    txtInputFile.Visible = True
    cmdSelectInputFile.Visible = True
    txtOutputFile.Visible = False
    txtOutputFile.Visible = False
End Select
If Err.Number <> 0 Then SetError "cboEventType_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdAdd_Click()
On Local Error Resume Next
Dim mItem As ListItem, msg As String
Dim i As Integer, f As Integer, lOutFile As String, lOutPath As String, InFile As String, lInPath As String

For i = 1 To lvwEvents.ListItems.Count
    Set mItem = lvwEvents.ListItems(i)
    Select Case LCase(mItem.Text)
    Case "play"
        InFile = mItem.SubItems(1)
        InFile = GetFileTitle(InFile)
        lInPath = Left(mItem.SubItems(1), Len(mItem.SubItems(1)) - Len(InFile))
        If Len(InFile) <> 0 And Len(lInPath) <> 0 Then AddEvent Play, lInPath, InFile, "", "", 0, ""
    Case "encode"
        InFile = mItem.SubItems(1)
        InFile = GetFileTitle(InFile)
        lInPath = Left(mItem.SubItems(1), Len(mItem.SubItems(1)) - Len(InFile))
        lOutFile = mItem.SubItems(2)
        lOutFile = GetFileTitle(lOutFile)
        lOutPath = Left(mItem.SubItems(2), Len(mItem.SubItems(2)) - Len(lOutFile))
        If Len(lOutFile) <> 0 And Len(lOutPath) <> 0 And Len(InFile) <> 0 And Len(lInPath) <> 0 Then AddEvent Encode, lInPath, InFile, lOutPath, lOutFile, 0, ""
    Case "rip"
        lOutFile = mItem.SubItems(2)
        lOutFile = GetFileTitle(lOutFile)
        lOutPath = Left(mItem.SubItems(2), Len(mItem.SubItems(2)) - Len(lOutFile))
        If Len(lOutFile) <> 0 And Len(lOutPath) <> 0 Then AddEvent Rip, "", "", lOutPath, lOutFile, mItem.SubItems(3), ""
    Case "decode"
        InFile = mItem.SubItems(1)
        InFile = GetFileTitle(InFile)
        lInPath = Left(mItem.SubItems(1), Len(mItem.SubItems(1)) - Len(InFile))
        lOutFile = mItem.SubItems(2)
        lOutFile = GetFileTitle(lOutFile)
        lOutPath = Left(mItem.SubItems(2), Len(mItem.SubItems(2)) - Len(lOutFile))
        AddEvent Decode, lInPath, InFile, lOutPath, lOutFile, 0, ""
    Case "rip2"
        lOutFile = mItem.SubItems(2)
        lOutFile = GetFileTitle(lOutFile)
        lOutPath = Left(mItem.SubItems(2), Len(mItem.SubItems(2)) - Len(lOutFile))
        If Len(InFile) <> 0 And Len(lInPath) <> 0 And Len(lOutFile) <> 0 And Len(lOutPath) <> 0 Then
            AddEvent Rip, "", "", lOutPath, Left(lOutFile, Len(lOutFile) - 3) & "wav", mItem.SubItems(3), ""
            AddEvent Encode, lOutPath, Left(lOutFile, Len(lOutFile) - 3) & "wav", lOutPath, lOutFile, 0, ""
        End If
    End Select
Next i
ProcessNextEvent
Unload Me
If Err.Number <> 0 Then SetError "cmdAdd_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdAddlvw_Click()
On Local Error Resume Next
Dim lListItem As ListItem, msg As String, lErr As Boolean
Select Case cboEventType.ListIndex
Case 0
    msg = "Encode"
    If Len(txtInputFile.Text) <> 0 And Len(txtOutputFile.Text) <> 0 Then
        Set lListItem = lvwEvents.ListItems.Add(, , msg)
        lListItem.SubItems(1) = txtInputFile.Text
        lListItem.SubItems(2) = txtOutputFile.Text
    Else
        lErr = True
    End If
Case 1
    msg = "Rip"
    If Len(txtOutputFile.Text) <> 0 And Len(cboCDAudioTrack.Text) <> 0 Then
        Set lListItem = lvwEvents.ListItems.Add(, , msg)
        lListItem.SubItems(2) = txtOutputFile.Text
        lListItem.SubItems(3) = cboCDAudioTrack.Text
    Else
        lErr = True
    End If
Case 2
    msg = "Rip2"
    If Len(txtOutputFile.Text) <> 0 And Len(cboCDAudioTrack.Text) <> 0 Then
        Set lListItem = lvwEvents.ListItems.Add(, , msg)
        lListItem.SubItems(2) = txtOutputFile.Text
        lListItem.SubItems(3) = cboCDAudioTrack.Text
    Else
        lErr = True
    End If
Case 3
    msg = "Decode"
    If Len(txtInputFile.Text) <> 0 And Len(txtOutputFile.Text) <> 0 Then
        Set lListItem = lvwEvents.ListItems.Add(, , msg)
        lListItem.SubItems(1) = txtInputFile.Text
        lListItem.SubItems(2) = txtOutputFile.Text
    Else
        lErr = True
    End If
Case 4
    If Len(txtInputFile.Text) <> 0 Then
        msg = "Play"
        If Right(LCase(txtInputFile.Text), 3) = "mp3" Then
            Set lListItem = lvwEvents.ListItems.Add(, , msg)
            lListItem.SubItems(1) = txtInputFile.Text
        ElseIf Right(LCase(txtInputFile.Text), 3) = "m3u" Then
            Dim CNumber, i As Integer
            CNumber = FreeFile
            Open txtInputFile.Text For Input As #CNumber
            Do While Not (EOF(CNumber))
                Line Input #CNumber, msg
                If Trim(msg) <> "" Then
                    Set lListItem = lvwEvents.ListItems.Add(, , "Play")
                    lListItem.SubItems(1) = msg
                End If
            Loop
            Close #CNumber
        End If
    Else
        lErr = True
    End If
End Select
If lErr = True And lEvents.eSettings.iOverwritePrompts = True Then MsgBox "Unable to add to batch", vbExclamation
If Err.Number <> 0 Then SetError "cmdAddlvw_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdCancel_Click()
On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then SetError "cmdCancel_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdDelLvw_Click()
On Local Error Resume Next
lvwEvents.ListItems.Remove lvwEvents.SelectedItem.Index
If Err.Number <> 0 Then SetError "cmdDelLVW()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdHelp_Click()
On Local Error Resume Next
MsgBox "Batch Event - Adds a job or a few jobs into the NexENCODE to do list and starts the process" & vbCrLf & "Event Type - What kind of event you want to do" & vbCrLf & "Input File - File to read from" & vbCrLf & "Output File - File to write to" & vbCrLf & "Track # - Which cd audio track to use" & vbCrLf & vbCrLf & "Input the needed information, then click add. When done editing your batch list, click 'Do Batch'", vbInformation
If Err.Number <> 0 Then SetError "cmdHelp_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdSelectInputFile_Click()
On Local Error Resume Next
Select Case cboEventType.ListIndex
Case 0
    txtInputFile.Text = OpenDialog(frmAddEvent, "Wav Files (*.wav)|*.wav", "Select Wave Audio File", CurDir)
Case 1
Case 2
Case 3
    txtInputFile.Text = OpenDialog(frmAddEvent, "Mp3 Files (*.mp3)|*.mp3", "Select Mpeg Audio", CurDir)
Case 4
    txtInputFile.Text = OpenDialog(frmAddEvent, "Mp3 Files (*.mp3)|*.mp3|M3u Playlists (*.m3u)|*.m3u", "Select Mpeg Audio", CurDir)
End Select
If Err.Number <> 0 Then SetError "cmdSelectInputFile_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Command1_Click()
On Local Error Resume Next
Dim msg As String
Select Case cboEventType.ListIndex
Case 0
    msg = Trim(SaveDialog(frmAddEvent, "Mp3 Files (*.mp3)|*.mp3", "Save as mp3 ...", CurDir))
    If Len(msg) <> 0 And LCase(Right(msg, 3)) <> "mp3" Then
        msg = Left(msg, Len(msg) - 1)
        msg = msg & ".mp3"
    End If
Case 1
    msg = Trim(SaveDialog(frmAddEvent, "Wav Files (*.wav)|*.wav", "Save as wav ...", CurDir))
    If Len(msg) <> 0 And LCase(Right(msg, 3)) <> "wav" Then
        msg = Left(msg, Len(msg) - 1)
        msg = msg & ".wav"
    End If

Case 2
    msg = Trim(SaveDialog(frmAddEvent, "Mp3 Files (*.mp3)|*.mp3", "Save as mp3 ...", CurDir))
    If Len(msg) <> 0 And LCase(Right(msg, 3)) <> "mp3" Then
        msg = Left(msg, Len(msg) - 1)
        msg = msg & ".mp3"
    End If
Case 3
    msg = Trim(SaveDialog(frmAddEvent, "Wav Files (*.wav)|*.wav", "Save as wav ...", CurDir))
    If Len(msg) <> 0 And LCase(Right(msg, 3)) <> "wav" Then
        msg = Left(msg, Len(msg) - 1)
        msg = msg & ".wav"
    End If
End Select
If Len(msg) <> 0 And Len(msg) <> 4 Then txtOutputFile.Text = msg
If Err.Number <> 0 Then SetError "cmdSelectOutputfile_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Load()
On Local Error Resume Next
Dim i As Integer
Icon = frmMain.Icon
lvwEvents.ColumnHeaders.Add , , "Type", 70
lvwEvents.ColumnHeaders.Add , , "Input", 70
lvwEvents.ColumnHeaders.Add , , "Output", 70
lvwEvents.ColumnHeaders.Add , , "Track", 43
cboCDAudioTrack.Clear
For i = 1 To lTracks.tCount
    cboCDAudioTrack.AddItem i
Next i
FlashIN Me
If Err.Number <> 0 Then SetError "cmdSelectInputFile_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Local Error Resume Next
FlashOut Me
If Err.Number <> 0 Then SetError "Form_Unload()", lEvents.eSettings.iErrDescription, Err.Description
End Sub
