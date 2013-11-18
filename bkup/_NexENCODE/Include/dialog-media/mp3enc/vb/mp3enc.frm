VERSION 5.00
Object = "{3B00B10A-6EF0-11D1-A6AA-0020AFE4DE54}#1.0#0"; "MP3PLAY.OCX"
Object = "{FFBEC4C3-839E-11D1-85FE-0020AFE4DE54}#1.0#0"; "MP3ENC.OCX"
Begin VB.Form Form1 
   Caption         =   "MP3PLAY.OCX Test"
   ClientHeight    =   7860
   ClientLeft      =   2865
   ClientTop       =   2775
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   ScaleHeight     =   7860
   ScaleWidth      =   9465
   Begin MP3ENCLib.Mp3Enc Mp3Enc1 
      Height          =   435
      Left            =   390
      TabIndex        =   48
      Top             =   6270
      Width           =   465
      _Version        =   65536
      _ExtentX        =   820
      _ExtentY        =   767
      _StockProps     =   0
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   4020
      TabIndex        =   47
      Text            =   "false"
      Top             =   6990
      Width           =   1545
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   2100
      TabIndex        =   46
      Text            =   "false"
      Top             =   6930
      Width           =   1365
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Downmix"
      Height          =   615
      Left            =   3960
      TabIndex        =   45
      Top             =   6120
      Width           =   1665
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Downsample"
      Height          =   615
      Left            =   2070
      TabIndex        =   44
      Top             =   6150
      Width           =   1425
   End
   Begin MPEGPLAYLib.Mp3Play Mp3Play1 
      Height          =   735
      Left            =   1140
      TabIndex        =   42
      Top             =   6120
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   1296
      _StockProps     =   0
   End
   Begin VB.Frame Frame10 
      Caption         =   "MP3Play"
      Height          =   1095
      Left            =   6090
      TabIndex        =   40
      Top             =   3750
      Width           =   3195
      Begin VB.CommandButton Command5 
         Caption         =   "Play"
         Height          =   375
         Left            =   330
         TabIndex        =   43
         Top             =   450
         Width           =   975
      End
      Begin VB.CommandButton Command6 
         Caption         =   "stop"
         Height          =   405
         Left            =   1800
         TabIndex        =   41
         Top             =   420
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "kbit/s"
      Height          =   3105
      Left            =   6060
      TabIndex        =   30
      Top             =   330
      Width           =   3105
      Begin VB.OptionButton Option8 
         Caption         =   "Option8"
         Height          =   285
         Left            =   1680
         TabIndex        =   39
         Top             =   1950
         Width           =   1155
      End
      Begin VB.OptionButton Option7 
         Caption         =   "256"
         Height          =   345
         Left            =   1680
         TabIndex        =   38
         Top             =   1410
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton Option6 
         Caption         =   "128"
         Height          =   315
         Left            =   1710
         TabIndex        =   37
         Top             =   930
         Width           =   1245
      End
      Begin VB.OptionButton Option5 
         Caption         =   "112"
         Height          =   315
         Left            =   1710
         TabIndex        =   36
         Top             =   450
         Width           =   1245
      End
      Begin VB.OptionButton Option4 
         Caption         =   "96"
         Height          =   405
         Left            =   240
         TabIndex        =   35
         Top             =   1830
         Width           =   1035
      End
      Begin VB.OptionButton Option3 
         Caption         =   "64"
         Height          =   435
         Left            =   240
         TabIndex        =   34
         Top             =   1350
         Width           =   915
      End
      Begin VB.OptionButton Option2 
         Caption         =   "56"
         Height          =   345
         Left            =   210
         TabIndex        =   33
         Top             =   840
         Width           =   885
      End
      Begin VB.OptionButton Option1 
         Caption         =   "32"
         Height          =   405
         Left            =   180
         TabIndex        =   32
         Top             =   330
         Width           =   1035
      End
      Begin VB.CommandButton Command2 
         Caption         =   "set"
         Height          =   435
         Left            =   930
         TabIndex        =   31
         Top             =   2460
         Width           =   1275
      End
   End
   Begin VB.Frame Frame13 
      Caption         =   "MP3 Filename Output"
      Height          =   615
      Left            =   3120
      TabIndex        =   27
      Top             =   2370
      Width           =   2295
      Begin VB.TextBox Text27 
         Height          =   345
         Left            =   150
         TabIndex        =   28
         Text            =   "c:\song.wav"
         Top             =   210
         Width           =   2025
      End
   End
   Begin VB.CommandButton Command13 
      Caption         =   "About"
      Height          =   705
      Left            =   210
      TabIndex        =   26
      Top             =   420
      Width           =   705
   End
   Begin VB.Frame Frame14 
      Caption         =   "Password Key"
      Height          =   1185
      Left            =   2970
      TabIndex        =   21
      Top             =   120
      Width           =   2625
      Begin VB.TextBox Text25 
         Height          =   285
         Left            =   1500
         TabIndex        =   25
         Top             =   780
         Width           =   1005
      End
      Begin VB.CommandButton Command12 
         Caption         =   "set"
         Height          =   285
         Left            =   1590
         TabIndex        =   24
         Top             =   330
         Width           =   765
      End
      Begin VB.TextBox Text24 
         Height          =   375
         Left            =   180
         TabIndex        =   23
         Text            =   "key"
         Top             =   690
         Width           =   1095
      End
      Begin VB.TextBox Text23 
         Height          =   345
         Left            =   180
         TabIndex        =   22
         Text            =   "password"
         Top             =   270
         Width           =   1065
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "Wave Filename Input"
      Height          =   675
      Left            =   3090
      TabIndex        =   19
      Top             =   1530
      Width           =   2325
      Begin VB.TextBox Text22 
         Height          =   315
         Left            =   150
         TabIndex        =   20
         Text            =   "c:\song.mp3"
         Top             =   270
         Width           =   2085
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "Errortext"
      Height          =   975
      Left            =   1890
      TabIndex        =   16
      Top             =   4920
      Width           =   3675
      Begin VB.TextBox Text21 
         Height          =   345
         Left            =   1140
         TabIndex        =   18
         Top             =   450
         Width           =   2385
      End
      Begin VB.CommandButton Command10 
         Caption         =   "off"
         Height          =   525
         Left            =   150
         TabIndex        =   17
         Top             =   330
         Width           =   765
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Total %"
      Height          =   675
      Left            =   3120
      TabIndex        =   14
      Top             =   3930
      Width           =   2205
      Begin VB.TextBox Text19 
         Height          =   315
         Left            =   180
         TabIndex        =   15
         Text            =   "0"
         Top             =   210
         Width           =   1845
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Thread endet"
      Height          =   1395
      Left            =   300
      TabIndex        =   13
      Top             =   4410
      Width           =   1305
      Begin VB.TextBox Text5 
         Height          =   315
         Left            =   150
         TabIndex        =   29
         Top             =   930
         Width           =   1005
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0000C000&
         FillColor       =   &H000000C0&
         FillStyle       =   7  'Diagonalkreuz
         Height          =   495
         Left            =   120
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Actframe"
      Height          =   645
      Left            =   3090
      TabIndex        =   11
      Top             =   3120
      Width           =   2235
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   180
         TabIndex        =   12
         Text            =   "0"
         Top             =   240
         Width           =   1845
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Stop Result"
      Height          =   765
      Left            =   330
      TabIndex        =   8
      Top             =   3390
      Width           =   2355
      Begin VB.TextBox Text4 
         Height          =   345
         Left            =   1080
         TabIndex        =   10
         Top             =   270
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Stop"
         Height          =   435
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Enc Result"
      Height          =   765
      Left            =   300
      TabIndex        =   5
      Top             =   2520
      Width           =   2355
      Begin VB.TextBox Text3 
         Height          =   405
         Left            =   1050
         TabIndex        =   7
         Top             =   270
         Width           =   1065
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Encode"
         Height          =   435
         Left            =   120
         TabIndex        =   6
         Top             =   270
         Width           =   825
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Open Result"
      Height          =   825
      Left            =   300
      TabIndex        =   2
      Top             =   1560
      Width           =   2415
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   1050
         TabIndex        =   4
         Top             =   300
         Width           =   1065
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Open"
         Height          =   435
         Left            =   150
         TabIndex        =   3
         Top             =   270
         Width           =   675
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "MP3ENC OCX Ver:"
      Height          =   885
      Left            =   990
      TabIndex        =   0
      Top             =   270
      Width           =   1725
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   210
         TabIndex        =   1
         Top             =   300
         Width           =   1275
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim kbits As Variant

Dim start_time As Long

Private Sub Command1_Click()

On Error GoTo error_handle
Text5.Text = ""
start_time = Timer

Text3.Text = Mp3Enc1.Encode

Exit Sub
error_handle:
Text21.Text = Error
Resume Next

End Sub


Private Sub Command10_Click()

If Command10.Caption = "off" Then
    Mp3Enc1.SetErrorMode 1
    Command10.Caption = "on"
Else
    Mp3Enc1.SetErrorMode 0
    Command10.Caption = "off"
End If

End Sub



Private Sub Command12_Click()
On Error GoTo error_handle

Text25.Text = Mp3Enc1.Authorize(Text23.Text, Text24.Text)

Exit Sub
error_handle:
Text21.Text = Error
Resume Next

End Sub

Private Sub Command13_Click()
Mp3Enc1.AboutBox


End Sub

Private Sub Command14_Click()
On Error GoTo error_handle
Text26.Text = Mp3Enc1.Pause
Exit Sub
error_handle:
Text21.Text = Error
Resume Next

End Sub

Private Sub Command2_Click()
On Error GoTo error_handle

Mp3Enc1.BitRate = kbits
Exit Sub
error_handle:
Text21.Text = Error
Resume Next

End Sub

Private Sub Command3_Click()
On Error GoTo error_handle
Text2.Text = Mp3Enc1.Open(Text22.Text, Text27.Text)
Exit Sub
error_handle:
Text21.Text = Error
Resume Next

End Sub

Private Sub Command4_Click()
On Error GoTo error_handle
Text4.Text = Mp3Enc1.Stop
Exit Sub
error_handle:
Text21.Text = Error
Resume Next

End Sub

Private Sub Command5_Click()
e = Mp3Play1.Open(Text27.Text, "")
Mp3Play1.Play

End Sub

Private Sub Command6_Click()

Mp3Play1.Stop

End Sub

Private Sub Command7_Click()
If Mp3Enc1.AllowDownSample = True Then
Mp3Enc1.AllowDownSample = False
Text7.Text = "false"
Else
Mp3Enc1.AllowDownSample = True
Text7.Text = "true"
End If
End Sub

Private Sub Command8_Click()

If Mp3Enc1.DownMix = True Then
Mp3Enc1.DownMix = False
Text8.Text = "false"
Else
Mp3Enc1.DownMix = True
Text8.Text = "true"
End If

End Sub

Private Sub Command9_Click()
On Error GoTo error_handle

Text20.Text = Mp3Play1.Seek(Text6.Text + 400)
Exit Sub
error_handle:
Text21.Text = Error
Resume Next

End Sub

Private Sub Form_Load()

Text1.Text = Mp3Enc1.Version / 100
kbits = 256000 'default
End Sub

Private Sub Mp3Play1_ActFrame(ByVal ActFrame As Long)
Text6.Text = ActFrame
Text19 = (ActFrame * Mp3Play1.MsPerFrame) \ 1000
End Sub

Private Sub Mp3Play1_ThreadEnded()

If Shape1.FillColor = &HC000& Then
    Shape1.FillColor = &HC0&
Else
    Shape1.FillColor = &HC000&
End If


End Sub

Private Sub VScroll1_Change()
On Error GoTo error_handle

vol_left = VScroll1.Value
vol_right = VScroll1.Value
vol_left = vol_left * 640
vol_right = vol_right * 640
e = Mp3Play1.SetVolume(vol_right, vol_left)
Exit Sub
error_handle:
Text21.Text = Error
Resume Next
End Sub

Private Sub Mp3Enc1_ActFrame(ByVal ActFrame As Long)
Text6.Text = ActFrame
Text19.Text = (ActFrame * 100 / Mp3Enc1.GetFrameCount) \ 1 & " %"
Text19.Text = Text19.Text & " Faktor: " & (Timer - start_time) / (ActFrame * 0.026)

End Sub

Private Sub Mp3Enc1_ThreadEnded()

If Shape1.FillColor = &HC000& Then
    Shape1.FillColor = &HC0&
Else
    Shape1.FillColor = &HC000&
End If
Text5.Text = (((Timer - start_time) * 10) \ 1) / 10 & " sec"
End Sub

Private Sub Option1_Click()
kbits = 32000

End Sub

Private Sub Option2_Click()
kbits = 56000

End Sub

Private Sub Option3_Click()
kbits = 64000
End Sub

Private Sub Option4_Click()
kbits = 96000
End Sub

Private Sub Option5_Click()
kbits = 112000

End Sub

Private Sub Option6_Click()
kbits = 128000

End Sub

Private Sub Option7_Click()
kbits = 256000
End Sub
