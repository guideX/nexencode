VERSION 5.00
Begin VB.Form frmDecode 
   Caption         =   "Decode"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   450
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   5970
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdSelectOutput 
      Caption         =   "Save as .."
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox txtOutputFile 
      Height          =   285
      Left            =   2040
      TabIndex        =   7
      Top             =   1560
      Width           =   3855
   End
   Begin VB.CommandButton cmdSelectMp3 
      Caption         =   "Open .."
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txtInputFilename 
      Height          =   285
      Left            =   2040
      TabIndex        =   4
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
      Top             =   3405
      Width           =   6015
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Decode"
         Default         =   -1  'True
         Height          =   375
         Left            =   3000
         TabIndex        =   2
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   4440
         TabIndex        =   1
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Decoding please wait ..."
      Height          =   255
      Left            =   2040
      TabIndex        =   9
      Top             =   3000
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Output File:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Input File:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1695
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      BorderWidth     =   4
      X1              =   6120
      X2              =   0
      Y1              =   3375
      Y2              =   3375
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   0
      Picture         =   "frmDecode.frx":0000
      Top             =   0
      Width           =   6000
   End
   Begin VB.Image Image2 
      Height          =   3390
      Left            =   0
      Picture         =   "frmDecode.frx":28DA
      Top             =   240
      Width           =   1920
   End
End
Attribute VB_Name = "frmDecode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
'on local error resume next
Unload Me
End Sub

Private Sub cmdOK_Click()
'on local error resume next
Dim lInput As String, lOutput As String

lInput = txtInputFilename.Text
lOutput = txtOutputFile.Text
If Right(LCase(lInput), 4) <> ".mp3" Then
    MsgBox "Input file must be an mp3 file"
    Exit Sub
End If
If Right(LCase(lOutput), 4) <> ".wav" Then
    MsgBox "Output file must be a wav file"
    Exit Sub
End If
If Len(lInput) = 0 Then
    SetError "cmdOK_Click()", "File system error", "No filename specified"
    Exit Sub
ElseIf DoesFileExist(lInput) = False Then
    SetError "cmdOK_Click()", "File system error", lInput & " could not be found"
    Exit Sub
End If
If Len(lOutput) = 0 Then
    SetError "cmdOK_Click()", "File system error", "No filename specified"
    Exit Sub
ElseIf DoesFileExist(lInput) = False Then
    SetError "cmdOK_Click()", "File system error", lOutput & " could not be found"
    Exit Sub
End If

txtInputFilename.Enabled = False
txtOutputFile.Enabled = False
DoEvents

Dim i As Integer
lOutput = Left(lOutput, Len(lOutput) - 4)
i = frmMain.Decoder.Open(lInput, lOutput): DoEvents
If i <> 0 Then
    MsgBox "An error occured"
    Exit Sub
End If
'pause 1
frmMain.Decoder.Play
Label3.Visible = True

If Err.Number <> 0 Then SetError "cmdOK_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdSelectMp3_Click()
'on local error resume next
Dim msg As String
msg = Trim(OpenDialog(frmDecode, "MP3 Files (*.mp3)|*.mp3|All Files (*.*)|*.*", "File to decode ...", CurDir))
If Len(msg) <> 0 And DoesFileExist(msg) = True Then
    If Right(LCase(msg), 4) <> ".mp3" Then
        txtInputFilename.Text = msg & ".mp3"
    Else
        txtInputFilename.Text = msg
    End If
End If
If Err.Number <> 0 Then SetError "cmdSelectMp3_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdSelectOutput_Click()
'on local error resume next
Dim msg As String, msg2 As String
msg = Trim(SaveDialog(frmDecode, "Wav Files (*.wav)|*.wav|All Files (*.*)|*.*", "Save as ...", CurDir))
msg = Left(msg, Len(msg) - 1)
If Right(LCase(msg), 4) <> ".wav" Then msg = msg & ".wav"

If Len(msg) <> 0 Then
    If DoesFileExist(msg) = True Then
        msg2 = MsgBox("File exists, overwrite?", vbYesNo + vbQuestion)
        If msg2 = vbYes Then
            Kill msg2
        Else
            Exit Sub
        End If
    End If
    If Right(LCase(msg), 4) <> ".wav" Then
        txtOutputFile.Text = msg & ".wav"
    Else
        txtOutputFile.Text = msg
    End If
End If

If Err.Number <> 0 Then SetError "cmdSelectOutput_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Decoder_ActFrame(ByVal ActFrame As Long)
Caption = "Decoding " & ActFrame
End Sub

Private Sub Decoder_Failure(ByVal ErrorCode As Long, ByVal ErrStr As String)
SetError "Decoder_Failure", "The decoder caused an error", ErrStr
End Sub

Private Sub Decoder_ThreadEnded()
Dim msg As VbMsgBoxResult

msg = MsgBox("Decode another?", vbYesNo + vbQuestion)
If msg = vbYes Then
    txtInputFilename.Enabled = True
    txtOutputFile.Enabled = True
    Caption = "Decode"
ElseIf msg = vbNo Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
'on local error resume next
Dim i As Integer, msg As String
Icon = frmMain.Icon
i = frmMain.Decoder.Authorize("Leon J Aiossa", "812144397")
'frmMain.Decoder.SetOutDevice 0
'frmMain.Decoder.GetOutDevice
FlashIN Me
If Err.Number <> 0 Then SetError "cmdOK_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
'on local error resume next
FlashOut Me
End Sub
