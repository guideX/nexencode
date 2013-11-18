VERSION 5.00
Object = "{60819404-3CCE-11D2-A800-008048E89E3E}#1.0#0"; "EFFECT.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   9060
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame14 
      Caption         =   "Remove DC"
      Height          =   975
      Left            =   7200
      TabIndex        =   55
      Top             =   3360
      Width           =   1095
      Begin VB.CommandButton Command17 
         Caption         =   "Start"
         Height          =   375
         Left            =   240
         TabIndex        =   56
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame13 
      Caption         =   "Volume"
      Height          =   855
      Left            =   6600
      TabIndex        =   52
      Top             =   2400
      Width           =   1695
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   840
         TabIndex        =   54
         Text            =   "0,5"
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Set"
         Height          =   375
         Left            =   120
         TabIndex        =   53
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "Average"
      Height          =   975
      Left            =   6600
      TabIndex        =   49
      Top             =   1320
      Width           =   1695
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   720
         TabIndex        =   51
         Text            =   " "
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Get"
         Height          =   375
         Left            =   120
         TabIndex        =   50
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "Fading"
      Height          =   1215
      Left            =   6600
      TabIndex        =   44
      Top             =   0
      Width           =   1695
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   960
         TabIndex        =   48
         Text            =   "20"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   240
         TabIndex        =   47
         Text            =   "20"
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Out"
         Height          =   375
         Left            =   960
         TabIndex        =   46
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton Command13 
         Caption         =   "In"
         Height          =   375
         Left            =   240
         TabIndex        =   45
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "Volume"
      Height          =   975
      Left            =   0
      TabIndex        =   42
      Top             =   2280
      Width           =   1335
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   720
         TabIndex        =   57
         Text            =   "30"
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Set"
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   960
      TabIndex        =   41
      Text            =   "c:\test.wav"
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   960
      TabIndex        =   40
      Text            =   "c:\waffi.wav"
      Top             =   0
      Width           =   2295
   End
   Begin VB.Frame Frame9 
      Caption         =   "Nomalize"
      Height          =   975
      Left            =   2880
      TabIndex        =   38
      Top             =   3360
      Width           =   855
      Begin VB.CommandButton Command8 
         Caption         =   "Start"
         Height          =   375
         Left            =   120
         TabIndex        =   39
         ToolTipText     =   "start normalizing"
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Invert"
      Height          =   975
      Left            =   3840
      TabIndex        =   36
      Top             =   3360
      Width           =   855
      Begin VB.CommandButton Command10 
         Caption         =   "Start"
         Height          =   375
         Left            =   120
         TabIndex        =   37
         ToolTipText     =   "invert wave file"
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "DeClick"
      Height          =   975
      Left            =   6000
      TabIndex        =   34
      Top             =   3360
      Width           =   1095
      Begin VB.CommandButton Command11 
         Caption         =   "Start"
         Height          =   375
         Left            =   240
         TabIndex        =   35
         ToolTipText     =   "start declick"
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Stop Effect"
      Height          =   375
      Left            =   120
      TabIndex        =   33
      ToolTipText     =   "stop effect"
      Top             =   1680
      Width           =   1095
   End
   Begin EFFECTLib.Effect Effect1 
      Height          =   495
      Left            =   360
      TabIndex        =   32
      Top             =   960
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   873
      _StockProps     =   0
   End
   Begin VB.Frame Frame6 
      Caption         =   "Shifting"
      Height          =   975
      Left            =   1440
      TabIndex        =   28
      Top             =   2280
      Width           =   1815
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1080
         TabIndex        =   31
         Text            =   "2048"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1080
         TabIndex        =   30
         Text            =   "1"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Start"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         ToolTipText     =   "start shift effect with parameter shown right"
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Write"
      Height          =   375
      Left            =   0
      TabIndex        =   27
      ToolTipText     =   "write Wave file named right"
      Top             =   480
      Width           =   735
   End
   Begin VB.Frame Frame5 
      Caption         =   "Distortion"
      Height          =   975
      Left            =   4800
      TabIndex        =   25
      Top             =   3360
      Width           =   1095
      Begin VB.CommandButton Distortion 
         Caption         =   "Start"
         Height          =   375
         Left            =   240
         TabIndex        =   26
         ToolTipText     =   "start distortion effect whith fix parameters from source"
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Courus"
      Height          =   975
      Left            =   1920
      TabIndex        =   23
      Top             =   3360
      Width           =   855
      Begin VB.CommandButton Command7 
         Caption         =   "Start"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         ToolTipText     =   "start chourus with some parameter from source"
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Reverb"
      Height          =   975
      Left            =   1440
      TabIndex        =   19
      Top             =   1320
      Width           =   1815
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1080
         TabIndex        =   22
         Text            =   "40"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1080
         TabIndex        =   21
         Text            =   "80"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Start"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "start reverb with the parameter shown right"
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "EQ"
      Height          =   3255
      Left            =   3360
      TabIndex        =   4
      Top             =   0
      Width           =   3135
      Begin VB.CommandButton Start 
         Caption         =   "Start"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         ToolTipText     =   "start param. EQ"
         Top             =   2520
         Width           =   975
      End
      Begin VB.VScrollBar VScroll10 
         Height          =   2000
         Left            =   2640
         Max             =   0
         Min             =   100
         TabIndex        =   17
         Top             =   360
         Width           =   200
      End
      Begin VB.VScrollBar VScroll9 
         Height          =   2000
         Left            =   2160
         Max             =   -100
         Min             =   100
         TabIndex        =   16
         Top             =   360
         Width           =   200
      End
      Begin VB.VScrollBar VScroll8 
         Height          =   2000
         Left            =   1920
         Max             =   -100
         Min             =   100
         TabIndex        =   15
         Top             =   360
         Value           =   50
         Width           =   200
      End
      Begin VB.VScrollBar VScroll7 
         Height          =   2000
         Left            =   1680
         Max             =   -100
         Min             =   100
         TabIndex        =   14
         Top             =   360
         Value           =   100
         Width           =   200
      End
      Begin VB.VScrollBar VScroll6 
         Height          =   2000
         Left            =   1440
         Max             =   -100
         Min             =   100
         TabIndex        =   13
         Top             =   360
         Width           =   200
      End
      Begin VB.VScrollBar VScroll5 
         Height          =   2000
         Left            =   1200
         Max             =   -100
         Min             =   100
         TabIndex        =   12
         Top             =   360
         Width           =   200
      End
      Begin VB.VScrollBar VScroll4 
         Height          =   2000
         Left            =   960
         Max             =   -100
         Min             =   100
         TabIndex        =   11
         Top             =   360
         Width           =   200
      End
      Begin VB.VScrollBar VScroll3 
         Height          =   2000
         Left            =   720
         Max             =   -100
         Min             =   100
         TabIndex        =   10
         Top             =   360
         Width           =   200
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   2000
         Left            =   480
         Max             =   -100
         Min             =   100
         TabIndex        =   9
         Top             =   360
         Width           =   200
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   2000
         Left            =   240
         Max             =   -100
         Min             =   100
         TabIndex        =   8
         Top             =   360
         Width           =   200
      End
      Begin VB.Line Line1 
         X1              =   2520
         X2              =   2520
         Y1              =   480
         Y2              =   2160
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Echo"
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Top             =   3360
      Width           =   1815
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   960
         TabIndex        =   7
         Text            =   "40"
         ToolTipText     =   "% org wave"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   960
         TabIndex        =   6
         Text            =   "1000"
         ToolTipText     =   "Delaytime in ms"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Start"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "start echo effect with parameter shown right"
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stop"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      ToolTipText     =   "stop wave playback"
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00000000&
      Caption         =   "Play"
      Height          =   375
      Left            =   1560
      MaskColor       =   &H00000000&
      TabIndex        =   1
      ToolTipText     =   "play loaded wave file"
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Load 
      Caption         =   "Load"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "load Wave file named right"
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim err
Private Sub CachePreLoader1_Complete(ByVal lUrlID As Long)

End Sub

Private Sub Command1_Click()
err = Effect1.Shifting(Text5.Text, Text6.Text)
End Sub

Private Sub Command10_Click()
Effect1.Invert
End Sub

Private Sub Command11_Click()
Effect1.CFilter (5)
End Sub

Private Sub Command12_Click()
Effect1.StopEffect
End Sub

Private Sub Command13_Click()
Effect1.FadeIn (Text9.Text)
End Sub

Private Sub Command14_Click()
Effect1.FadeOut Text10.Text
End Sub

Private Sub Command15_Click()
Text11.Text = Effect1.Average
End Sub

Private Sub Command16_Click()
Effect1.Pegel (0.5)
End Sub

Private Sub Command17_Click()
Effect1.Amplitude
End Sub

Private Sub Command2_Click()
Effect1.Play
End Sub

Private Sub Command3_Click()
Effect1.Stop
End Sub

Private Sub Command4_Click()
err = Effect1.Echo(Text1.Text, Text2.Text)
End Sub


Private Sub Command5_Click()
err = Effect1.SetVolume(Text13.Text, Text13.Text)
End Sub

Private Sub Command6_Click()
err = Effect1.Reverb(Text3.Text, Text4.Text)

End Sub

Private Sub Command7_Click()
err = Effect1.Chorus(10, 60, 0.8, 1, 75, 80, 1, -1, 0) 'Uiiiii
End Sub

Private Sub Command8_Click()
Effect1.Pegel (-1) 'normalize
End Sub

Private Sub Command9_Click()
Effect1.InputFileSave Text8.Text
End Sub

Private Sub Distortion_Click()
err = Effect1.Distortion(46, 560, 2, 0, 0)
End Sub

Private Sub Effect1_EndOfAction(ByVal ActionType As Integer)
Dim alt

Form1.BackColor = &H80000007
DoEvents
Form1.Refresh
DoEvents
alt = Timer

While alt + 0.1 > Timer
DoEvents
Wend

Form1.BackColor = &H8000000F

End Sub

Private Sub Effect1_OnActionPosition(ByVal ActionPosition As Integer)
Form1.Caption = ActionPosition
End Sub

Private Sub Form_Load()
'return_value = Effect1.Authorize("Eff0000000DM", "1111111111")
End Sub


Private Sub Load_Click()
Effect1.InputFileOpen Text7.Text
End Sub

Private Sub Start_Click()
e = Effect1.Eq(VScroll1, VScroll2, VScroll3, VScroll4, VScroll5, VScroll6, VScroll7, VScroll8, VScroll9, VScroll10)
End Sub

