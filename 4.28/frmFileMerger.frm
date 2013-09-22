VERSION 5.00
Object = "{EE128208-4F73-11D3-83BB-C47C02EE3D01}#1.0#0"; "ControlResizer.ocx"
Begin VB.Form frmFileMerger 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "NexENCODE - File Merger"
   ClientHeight    =   3240
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
   Icon            =   "frmFileMerger.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   216
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   398
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton cmdOpenMerged 
      Caption         =   "&Play Merged"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   240
      Width           =   1695
   End
   Begin VB.CheckBox chkExitWhenComplete 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      Caption         =   "Exit When Complete"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton cmdMerge 
      Caption         =   "&Merge"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdSaveAs 
      Caption         =   "Select ..."
      Height          =   255
      Left            =   4800
      TabIndex        =   7
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox txtSaveAs 
      Height          =   285
      Left            =   2760
      TabIndex        =   6
      Top             =   2040
      Width           =   3135
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Del"
      Height          =   255
      Left            =   5040
      TabIndex        =   4
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton cmdADD 
      Caption         =   "&Add"
      Height          =   255
      Left            =   4320
      TabIndex        =   3
      Top             =   1680
      Width           =   735
   End
   Begin VB.ListBox lstFiles 
      Height          =   1230
      Left            =   2040
      TabIndex        =   2
      Top             =   360
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
         TabIndex        =   14
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdOK 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   315
         Left            =   4800
         TabIndex        =   1
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1320
         TabIndex        =   15
         Top             =   120
         Width           =   3135
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         X1              =   0
         X2              =   6000
         Y1              =   0
         Y2              =   0
      End
   End
   Begin ControlResizer.AutoResizer AutoResizer1 
      Height          =   495
      Left            =   1440
      Top             =   120
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin VB.Label lblJustMerged 
      BackStyle       =   0  'Transparent
      Height          =   855
      Left            =   2640
      TabIndex        =   12
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Files to merge:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Save as:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   2040
      Width           =   1695
   End
End
Attribute VB_Name = "frmFileMerger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private TheFiles(250) As String
Private CancelFunc As Boolean
Option Explicit

Public Sub MergeFiles(Optional lSaveAs As String, Optional lExitOnComplete As Boolean)
'On Local Error Resume Next
Dim mbox As VbMsgBoxResult
If lExitOnComplete = True Then chkExitWhenComplete.Value = 1
If Len(txtSaveAs.Text) = 0 Then
Start:
    If Len(lSaveAs) <> 0 Then
        If Right(LCase(lSaveAs), 4) <> ".mp3" Then lSaveAs = lSaveAs & ".mp3"
        If DoesFileExist(lSaveAs) = True Then
            If lEvents.eSettings.iOverwritePrompts = True Then
                mbox = MsgBox("File """ & lSaveAs & """ exists, overwrite?", vbYesNo + vbQuestion, "Overwrite?")
                If mbox = vbYes Then
                    Kill lSaveAs
                ElseIf mbox = vbNo Then
                    Exit Sub
                End If
            Else
                Kill lSaveAs
            End If
        End If
        txtSaveAs.Text = lSaveAs
        cmdMerge_Click
    Else
        cmdSaveAs_Click
        DoEvents
        GoTo Start
    End If
Else
    cmdMerge_Click
End If
If Err.Number <> 0 Then SetError "MergeFiles", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub AddToMergeList(lFilename As String)
'On Local Error Resume Next
If Len(lFilename) <> 0 Then
    If DoesFileExist(lFilename) = True Then lstFiles.AddItem lFilename
End If
If Err.Number <> 0 Then SetError "AddToMergeList", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdAdd_Click()
'On Local Error Resume Next
Dim i As Integer, msg As String
msg = OpenDialog(Me, "Mpeg Layer 3 (*.mp3)|*.mp3|All Files (*.*)|*.*", "NexENCODE", CurDir)

If Len(msg) = 0 Then Exit Sub
For i = 0 To lstFiles.ListCount
    If lstFiles.List(i) = msg Then
        If lEvents.eSettings.iOverwritePrompts = True Then
            MsgBox "File exists in que", vbExclamation
            Exit Sub
        End If
    End If
Next i
If DoesFileExist(msg) = True Then lstFiles.AddItem msg
If Err.Number <> 0 Then SetError "cmdADD_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdCancel_Click()
'On Local Error Resume Next
CancelFunc = True
If Err.Number <> 0 Then SetError "cmdCancel_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdDel_Click()
'On Local Error Resume Next
If Len(lstFiles.Text) <> 0 Then lstFiles.RemoveItem lstFiles.ListIndex
If Err.Number <> 0 Then SetError "cmdDel_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdHelp_Click()
'On Local Error Resume Next
MsgBox "Click 'Add' to add mp3 files into the listbox. Once you have selected your files (must be more than 1) click 'select' to select a file to save, then click 'merge'. When it is done, click 'Play Merged'", vbInformation
End Sub

Private Sub cmdMerge_Click()
'On Local Error Resume Next
Dim mbox As VbMsgBoxResult, msg As String, i As Long, X As Long, SavedSpot As Long, theByte() As Byte, Length As Long, m As Integer
If lstFiles.ListCount = 0 Or lstFiles.ListCount = 1 Then
    If lEvents.eSettings.iOverwritePrompts = True Then MsgBox "Not enough files to merge! Aborting", vbExclamation
    Exit Sub
End If
If Len(txtSaveAs.Text) = 0 Then
    If lEvents.eSettings.iOverwritePrompts = True Then MsgBox "No file set! Aborting", vbExclamation
    Exit Sub
End If
For m = 0 To lstFiles.ListCount
    If DoesFileExist(TheFiles(m)) = True Then TheFiles(m) = lstFiles.List(m)
Next m
lstFiles.Enabled = False
cmdAdd.Enabled = False
cmdDel.Enabled = False
cmdSaveAs.Enabled = False
cmdMerge.Enabled = False
cmdCancel.Enabled = True
txtSaveAs.Enabled = False
SavedSpot = 1
For i = 0 To lstFiles.ListCount - 1
    If CancelFunc = True Then GoTo Done
    If Len(TheFiles(i)) <> 0 Then
        Length = FileLen(TheFiles(i))
        ReDim theByte(Length - 1)
        Open TheFiles(i) For Binary Access Read As #1
            Get #1, , theByte()
        Close #1
        Open txtSaveAs.Text For Binary As #1
            Put #1, SavedSpot, theByte()
        Close #1
        LblStatus.Caption = "File: " & i + 1 & " of " & lstFiles.ListCount & " [" & Int((100 / lstFiles.ListCount) * i + 1) & " %]"
        SavedSpot = SavedSpot + Length
        DoEvents
    End If
Next i
Done:
LblStatus.Caption = ""
lstFiles.Enabled = True
lstFiles.Clear
cmdAdd.Enabled = True
cmdDel.Enabled = True
cmdCancel.Enabled = False
txtSaveAs.Enabled = True
lblJustMerged.Caption = txtSaveAs.Text
txtSaveAs.Text = ""
cmdMerge.Enabled = True
cmdSaveAs.Enabled = True
If chkExitWhenComplete.Value = 1 Then
    ProcessNextEvent
    If lEvents.eSettings.iShowReports = True Then
        If lReports.rCount <> 0 Then
            frmReport.Show
            lEvents.eEventCount = 0
        End If
    End If
    Unload Me
End If
If Err.Number <> 0 Then SetError "cmdMerge_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdOK_Click()
'On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then SetError "cmdOK_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdOpenMerged_Click()
'On Local Error Resume Next
Dim msg As String, msg2 As String
msg = lblJustMerged.Caption
msg2 = msg
msg2 = GetFileTitle(msg2)
msg = Left(msg, Len(msg) - Len(msg2))
AddPlayEvent msg, msg2
If Err.Number <> 0 Then SetError "cmdOpenMerged_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdSaveAs_Click()
'On Local Error Resume Next
Dim msg As String, msg2 As VbMsgBoxResult
msg = SaveDialog(Me, "Mpeg Layer 3 (*.mp3)|*.mp3|All Files (*.*)|*.*", "Save as ..", CurDir)
If Len(msg) <> 0 Then
    msg = Left(msg, Len(msg) - 1)
    If Right(LCase(msg), 4) <> ".mp3" Then msg = msg & ".mp3"
    If Len(msg) <> 0 Then txtSaveAs.Text = msg
End If
If Err.Number <> 0 Then SetError "cmdSaveAs_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Load()
'On Local Error Resume Next
FlashIN Me
Icon = frmMain.Icon
If Err.Number <> 0 Then SetError "Form_Load()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Local Error Resume Next
FlashOut Me
If Err.Number <> 0 Then SetError "Form_Unload()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub lblJustMerged_Change()
'On Local Error Resume Next
If Len(lblJustMerged.Caption) <> 0 Then
    cmdOpenMerged.Enabled = True
Else
    cmdOpenMerged.Enabled = False
End If
If Err.Number <> 0 Then SetError "lblJustMerged_Change", lEvents.eSettings.iErrDescription, Err.Description
End Sub
