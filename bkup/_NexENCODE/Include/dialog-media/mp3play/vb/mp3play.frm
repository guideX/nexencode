VERSION 5.00
Object = "{3B00B10A-6EF0-11D1-A6AA-0020AFE4DE54}#1.0#0"; "Mp3play.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "MP3PLAY.OCX Test"
   ClientHeight    =   3015
   ClientLeft      =   3645
   ClientTop       =   3465
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   5370
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4440
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MPEGPLAYLib.Mp3Play Mp3Play1 
      Height          =   735
      Left            =   4440
      TabIndex        =   7
      Top             =   480
      Width           =   675
      _Version        =   65536
      _ExtentX        =   1191
      _ExtentY        =   1296
      _StockProps     =   0
   End
   Begin VB.Frame Frame10 
      Caption         =   "Seek"
      Height          =   705
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   3555
      Begin VB.HScrollBar HScroll1 
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   3250
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Acttime"
      Height          =   705
      Left            =   2280
      TabIndex        =   4
      Top             =   1560
      Width           =   1395
      Begin VB.TextBox Text3 
         Height          =   300
         Left            =   120
         TabIndex        =   5
         Text            =   "0"
         Top             =   270
         Width           =   1095
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Actframe"
      Height          =   675
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   2085
      Begin VB.TextBox Text2 
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Text            =   "0"
         Top             =   270
         Width           =   1785
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Command"
      Height          =   1395
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3555
      Begin VB.CommandButton Command4 
         Caption         =   "Pause"
         Height          =   400
         Left            =   2640
         TabIndex        =   11
         Top             =   360
         Width           =   750
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Stop"
         Height          =   400
         Left            =   1800
         TabIndex        =   10
         Top             =   360
         Width           =   750
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   3240
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Open"
         Height          =   400
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   750
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Play"
         Height          =   400
         Left            =   960
         TabIndex        =   1
         Top             =   360
         Width           =   750
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

CommonDialog1.CancelError = True
On Error GoTo ErrHandler
CommonDialog1.Flags = cdlOFNHideReadOnly
CommonDialog1.Filter = "MP3 Files" & "(*.mp3)|*.mp3"
CommonDialog1.FilterIndex = 2
CommonDialog1.ShowOpen

Text1.Text = "Open return: " & Mp3Play1.Open(CommonDialog1.FileName, "")
 
Exit Sub

ErrHandler:
Exit Sub


End Sub



Private Sub Command2_Click()

Text1.Text = "Play return: " & Mp3Play1.Play
HScroll1.Max = Mp3Play1.FrameCount

End Sub

Private Sub Command3_Click()
Text1.Text = "Stop return: " & Mp3Play1.Stop


End Sub

Private Sub Command4_Click()
Text1.Text = "Pause return: " & Mp3Play1.Pause


End Sub


Private Sub Form_Load()

Dim ret_val
ret_val = Mp3Play1.Authorize("xxxxxxxx", "yyyyyyyy")
Text1.Text = "MP3play.ocx Version: " & Mp3Play1.Version / 100

End Sub

Private Sub HScroll1_Change()
e = Mp3Play1.SetVolumeP(HScroll1.Value, HScroll1.Value)
End Sub


Private Sub Mp3Play1_ActFrame(ByVal ActFrame As Long)

Text2.Text = ActFrame & " of " & Mp3Play1.FrameCount
Text3 = (ActFrame * Mp3Play1.MsPerFrame) \ 1000 & " sec"
HScroll1.Value = ActFrame

End Sub

Private Sub Mp3Play1_ThreadEnded()

Text1.Text = "Thread ended"

End Sub

Private Sub VScroll1_Change()

End Sub

Private Sub Text6_Change()

End Sub
