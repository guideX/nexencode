VERSION 5.00
Object = "{9F5F61C6-83A0-11D2-A800-00A0CC20D781}#1.0#0"; "ACD.ocx"
Object = "{FFBEC4C3-839E-11D1-85FE-0020AFE4DE54}#1.0#0"; "Mp3Enc.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NexENCODE"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   510
   ClientWidth     =   7080
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":08CA
   ScaleHeight     =   472
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   472
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   6480
      ScaleHeight     =   5535
      ScaleWidth      =   615
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   615
      Begin VB.Timer tmrShowEncoderCircles 
         Enabled         =   0   'False
         Interval        =   40
         Left            =   0
         Top             =   2160
      End
      Begin VB.Timer tmrShowRipperCircles 
         Enabled         =   0   'False
         Interval        =   40
         Left            =   0
         Top             =   1800
      End
      Begin VB.Timer tmrFlash 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   0
         Top             =   1440
      End
      Begin VB.Timer tmrProcessEvent 
         Interval        =   1000
         Left            =   0
         Top             =   1080
      End
      Begin MP3ENCLib.Mp3Enc Encoder 
         Height          =   495
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   495
         _Version        =   65536
         _ExtentX        =   873
         _ExtentY        =   873
         _StockProps     =   0
      End
      Begin ACDLib.ACD Ripper 
         Height          =   495
         Left            =   0
         TabIndex        =   6
         Top             =   480
         Width           =   495
         _Version        =   65536
         _ExtentX        =   873
         _ExtentY        =   873
         _StockProps     =   0
      End
   End
   Begin VB.PictureBox picButtons 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   0
      ScaleHeight     =   1575
      ScaleWidth      =   7095
      TabIndex        =   0
      Top             =   5520
      Visible         =   0   'False
      Width           =   7095
      Begin VB.Image imgFlashEncode2 
         Height          =   735
         Left            =   0
         Picture         =   "frmMain.frx":84D8
         Top             =   480
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Image imgEnd1 
         Height          =   300
         Left            =   2760
         Picture         =   "frmMain.frx":894D
         Top             =   480
         Width           =   300
      End
      Begin VB.Image imgEnd2 
         Height          =   300
         Left            =   2760
         Picture         =   "frmMain.frx":8E3F
         Top             =   480
         Width           =   300
      End
      Begin VB.Image imgMinimize2 
         Height          =   300
         Left            =   2400
         Picture         =   "frmMain.frx":9331
         Top             =   480
         Width           =   300
      End
      Begin VB.Image imgMinimize1 
         Height          =   300
         Left            =   2400
         Picture         =   "frmMain.frx":9823
         Top             =   480
         Width           =   300
      End
      Begin VB.Image imgId32 
         Height          =   300
         Left            =   2040
         Picture         =   "frmMain.frx":9D15
         Top             =   480
         Width           =   300
      End
      Begin VB.Image imgId31 
         Height          =   300
         Left            =   2040
         Picture         =   "frmMain.frx":A1D5
         Top             =   480
         Width           =   300
      End
      Begin VB.Image imgRipper2 
         Height          =   735
         Left            =   960
         Picture         =   "frmMain.frx":A3B7
         Top             =   480
         Width           =   855
      End
      Begin VB.Image imgPlayMp32 
         Height          =   435
         Left            =   2520
         Picture         =   "frmMain.frx":AD20
         Top             =   0
         Width           =   435
      End
      Begin VB.Image imgPlayMp31 
         Height          =   435
         Left            =   2520
         Picture         =   "frmMain.frx":B30F
         Top             =   0
         Width           =   435
      End
      Begin VB.Image imgStopEncoding2 
         Height          =   435
         Left            =   2040
         Picture         =   "frmMain.frx":B762
         Top             =   0
         Width           =   435
      End
      Begin VB.Image imgStopEncoding1 
         Height          =   435
         Left            =   2040
         Picture         =   "frmMain.frx":BD1D
         Top             =   0
         Width           =   435
      End
      Begin VB.Image imgEncode2 
         Height          =   435
         Left            =   1560
         Picture         =   "frmMain.frx":C140
         Top             =   0
         Width           =   435
      End
      Begin VB.Image imgEncode1 
         Height          =   435
         Left            =   1560
         Picture         =   "frmMain.frx":C735
         Top             =   0
         Width           =   435
      End
      Begin VB.Image imgPlayWav2 
         Height          =   420
         Left            =   960
         Picture         =   "frmMain.frx":CB7E
         Top             =   0
         Width           =   435
      End
      Begin VB.Image imgPlayWav1 
         Height          =   420
         Left            =   960
         Picture         =   "frmMain.frx":D177
         Top             =   0
         Width           =   435
      End
      Begin VB.Image imgStopRipping2 
         Height          =   420
         Left            =   480
         Picture         =   "frmMain.frx":DB59
         Top             =   0
         Width           =   435
      End
      Begin VB.Image imgStopRipping1 
         Height          =   420
         Left            =   480
         Picture         =   "frmMain.frx":E0F7
         Top             =   0
         Width           =   435
      End
      Begin VB.Image imgRip2 
         Height          =   420
         Left            =   0
         Picture         =   "frmMain.frx":E514
         Top             =   0
         Width           =   435
      End
      Begin VB.Image imgRip1 
         Height          =   420
         Left            =   0
         Picture         =   "frmMain.frx":EAC8
         Top             =   0
         Width           =   435
      End
      Begin VB.Image imgFlashEncode1 
         Height          =   780
         Left            =   0
         Picture         =   "frmMain.frx":EEEF
         Top             =   480
         Width           =   900
      End
      Begin VB.Image imgRipper1 
         Height          =   735
         Left            =   960
         Picture         =   "frmMain.frx":113C1
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Shape shpEncoder 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   11
      Left            =   5280
      Shape           =   2  'Oval
      Top             =   3300
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpEncoder 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   10
      Left            =   5505
      Shape           =   2  'Oval
      Top             =   3330
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpEncoder 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   9
      Left            =   5745
      Shape           =   2  'Oval
      Top             =   3285
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpEncoder 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   8
      Left            =   5955
      Shape           =   2  'Oval
      Top             =   3180
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpEncoder 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   7
      Left            =   6105
      Shape           =   2  'Oval
      Top             =   3000
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpEncoder 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   6
      Left            =   6195
      Shape           =   2  'Oval
      Top             =   2760
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpEncoder 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   5
      Left            =   6195
      Shape           =   2  'Oval
      Top             =   2520
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpEncoder 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   4
      Left            =   6120
      Shape           =   2  'Oval
      Top             =   2280
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpEncoder 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   3
      Left            =   5955
      Shape           =   2  'Oval
      Top             =   2085
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpEncoder 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   2
      Left            =   5708
      Shape           =   2  'Oval
      Top             =   1950
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpEncoder 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   1
      Left            =   5460
      Shape           =   2  'Oval
      Top             =   1920
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpEncoder 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   0
      Left            =   5220
      Shape           =   2  'Oval
      Top             =   1965
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpRipper 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   11
      Left            =   1620
      Shape           =   2  'Oval
      Top             =   3285
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpRipper 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   10
      Left            =   1350
      Shape           =   2  'Oval
      Top             =   3330
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpRipper 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   9
      Left            =   1095
      Shape           =   2  'Oval
      Top             =   3270
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpRipper 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   8
      Left            =   900
      Shape           =   2  'Oval
      Top             =   3120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpRipper 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   7
      Left            =   750
      Shape           =   2  'Oval
      Top             =   2925
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpRipper 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   6
      Left            =   690
      Shape           =   2  'Oval
      Top             =   2685
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpRipper 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   5
      Left            =   705
      Shape           =   2  'Oval
      Top             =   2445
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpRipper 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   4
      Left            =   795
      Shape           =   2  'Oval
      Top             =   2235
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpRipper 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   3
      Left            =   945
      Shape           =   2  'Oval
      Top             =   2070
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpRipper 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   2
      Left            =   1125
      Shape           =   2  'Oval
      Top             =   1965
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpRipper 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   1
      Left            =   1335
      Shape           =   2  'Oval
      Top             =   1920
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpRipper 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   0
      Left            =   1560
      Shape           =   2  'Oval
      Top             =   1935
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image imgEnd 
      Height          =   300
      Left            =   5580
      Picture         =   "frmMain.frx":1167A
      Top             =   1515
      Width           =   300
   End
   Begin VB.Image imgMinimize 
      Height          =   300
      Left            =   5325
      Picture         =   "frmMain.frx":11B6C
      Top             =   1320
      Width           =   300
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "idle"
      BeginProperty Font 
         Name            =   "Digiface"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   3135
      Width           =   1815
   End
   Begin VB.Image imgId3 
      Height          =   300
      Left            =   5025
      Picture         =   "frmMain.frx":1205E
      Top             =   1185
      Width           =   300
   End
   Begin VB.Label lblMp3File 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C0FFFF&
      Height          =   615
      Left            =   5080
      TabIndex        =   2
      Top             =   2400
      Width           =   975
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgPlayMp3 
      Height          =   435
      Left            =   4680
      Picture         =   "frmMain.frx":12550
      Top             =   2940
      Width           =   435
   End
   Begin VB.Image imgStopEncoding 
      Height          =   435
      Left            =   4560
      Picture         =   "frmMain.frx":129A3
      Top             =   2490
      Width           =   435
   End
   Begin VB.Image imgEncode 
      Height          =   435
      Left            =   4680
      Picture         =   "frmMain.frx":12DC6
      Top             =   2050
      Width           =   435
   End
   Begin VB.Image imgPlayWav 
      Height          =   420
      Left            =   1920
      Picture         =   "frmMain.frx":1320F
      Top             =   2950
      Width           =   435
   End
   Begin VB.Image imgCancelRip 
      Height          =   420
      Left            =   2055
      Picture         =   "frmMain.frx":13648
      Top             =   2520
      Width           =   435
   End
   Begin VB.Image imgRip 
      Height          =   420
      Left            =   1920
      Picture         =   "frmMain.frx":13A65
      Top             =   2080
      Width           =   435
   End
   Begin VB.Label lblWavFile 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C0FFFF&
      Height          =   615
      Left            =   960
      TabIndex        =   1
      Top             =   2400
      Width           =   975
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgRipper 
      Height          =   735
      Left            =   1035
      Picture         =   "frmMain.frx":13E8C
      Top             =   2325
      Width           =   855
   End
   Begin VB.Image imgFlashEncode 
      Height          =   780
      Left            =   5115
      Picture         =   "frmMain.frx":14145
      Top             =   2310
      Width           =   900
   End
   Begin VB.Menu mnuHidden 
      Caption         =   "Hidden"
      Visible         =   0   'False
      Begin VB.Menu mnuNexENCODE 
         Caption         =   "NexENCODE"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ACD1_Failure(ByVal ErrorCode As Long, ByVal ErrorString As String)
MsgBox ErrorString
End Sub

Private Sub Encoder_ActFrame(ByVal ActFrame As Long)
Dim i As Long
lEvents.eEventBusy = True
i = ActFrame / Encoder.GetFrameCount * 100 / 1.5
lEvents.eDescription = "encode"
If tmrFlash.Enabled = False Then tmrFlash.Enabled = True
'EncodeCircleEffect Str(i)
'lEvents.ePercent = i
lblInfo.Caption = "Encoding " & i & "% Complete"
End Sub

Private Sub Encoder_Failure(ByVal ErrCode As Long, ByVal ErrStr As String)
MsgBox ErrCode
End Sub

Private Sub Encoder_ThreadEnded()
lblInfo.Caption = "Task completed successfully"
lEvents.eEventBusy = False
lblMp3File.Caption = ""
tmrFlash.Enabled = False
imgFlashEncode.Picture = imgFlashEncode1.Picture
End Sub

Private Sub Form_Load()
'on local error resume next

Ripper.Init
DoEvents
If Ripper.IsAspiLoaded = False Then
    Dim msg As String
    msg = MsgBox("Your WINASPI is out of date and needs to be updated. Do you wish to download it now?", vbYesNo + vbQuestion)
    If msg = vbYes Then
    
    ElseIf msg = vbNo Then
    
    End If
End If
GetWindowSettings hwnd
SetFiles
SetShape
RegisterComponents
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 Then
    FormDrag Me
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If lEvents.eEventBusy = False Then
    lEvents.eDescription = "idle"
    lEvents.ePercent = 0
    ConvertCaption
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Ripper.sTop
Encoder.sTop
End Sub

Private Sub imgCancelRip_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 Then
    imgCancelRip.Picture = imgStopRipping2.Picture
End If
End Sub

Private Sub imgCancelRip_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 And imgCancelRip.Picture = imgStopRipping2.Picture Then
    If x > 500 Or x < -1 Or Y > 500 Or Y < -1 Then imgCancelRip.Picture = imgStopRipping1.Picture
ElseIf Button = 1 And imgCancelRip.Picture = imgStopRipping1.Picture Then
    If x < 500 And x > -1 And Y < 500 And Y > -1 Then imgCancelRip.Picture = imgStopRipping2.Picture
ElseIf Button = 0 Then
    If lEvents.eEventBusy = False Then
        lEvents.eDescription = "stoprip"
        lEvents.ePercent = 0
        ConvertCaption
    End If
End If
End Sub

Private Sub imgCancelRip_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 And imgCancelRip.Picture = imgStopRipping2.Picture Then
    imgCancelRip.Picture = imgStopRipping1.Picture
End If
End Sub

Private Sub imgEncode_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 Then
    imgEncode.Picture = imgEncode2.Picture
End If
End Sub

Private Sub imgEncode_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 And imgEncode.Picture = imgEncode2.Picture Then
    If x > 500 Or x < -1 Or Y > 500 Or Y < -1 Then imgEncode.Picture = imgEncode1.Picture
ElseIf Button = 1 And imgEncode.Picture = imgEncode1.Picture Then
    If x < 500 And x > -1 And Y < 500 And Y > -1 Then imgEncode.Picture = imgEncode2.Picture
ElseIf Button = 0 Then
    If lEvents.eEventBusy = False Then
        lEvents.eDescription = "encode"
        lEvents.ePercent = 0
        ConvertCaption
    End If
End If
End Sub

Private Sub imgEncode_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 And imgEncode.Picture = imgEncode2.Picture Then
    imgEncode.Picture = imgEncode1.Picture
    Dim msg As String, msg2 As String
    msg = OpenDialog(Me, "Wav Audio Files (*.wav)|*.wav|All Files (*.*)|*.*", "Select file to compress", CurDir): DoEvents
    If Len(msg) <> 0 Then
        If InStr(LCase(msg), "wav") Then
            msg2 = Left(msg, Len(msg) - 4) & ".mp3"
            EncodeFile msg, msg2
        End If
    End If
End If
End Sub

Private Sub imgEnd_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 Then
    imgEnd.Picture = imgEnd2.Picture
End If
End Sub

Private Sub imgEnd_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 And imgEnd.Picture = imgEnd2.Picture Then
    If x > 250 Or x < -1 Or Y > 250 Or Y < -1 Then imgEnd.Picture = imgEnd1.Picture
ElseIf Button = 1 And imgEnd.Picture = imgEnd1.Picture Then
    If x < 250 And x > -1 And Y < 250 And Y > -1 Then imgEnd.Picture = imgEnd2.Picture
ElseIf Button = 0 Then
    If lEvents.eEventBusy = False Then
        lEvents.eDescription = "end"
        lEvents.ePercent = 0
        ConvertCaption
    End If
End If
End Sub

Private Sub imgEnd_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 And imgEnd.Picture = imgEnd2.Picture Then
    imgEnd.Picture = imgEnd1.Picture
    End
End If
End Sub

Private Sub imgFlashEncode_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
tmrFlash.Enabled = False
lEvents.eDescription = "idle"
ConvertCaption
End Sub

Private Sub imgId3_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 Then
    imgId3.Picture = imgId32.Picture
End If
End Sub

Private Sub imgId3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 And imgId3.Picture = imgId32.Picture Then
    If x > 250 Or x < -1 Or Y > 250 Or Y < -1 Then imgId3.Picture = imgId31.Picture
ElseIf Button = 1 And imgId3.Picture = imgId31.Picture Then
    If x < 250 And x > -1 And Y < 250 And Y > -1 Then imgId3.Picture = imgId32.Picture
End If
End Sub

Private Sub imgId3_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 And imgId3.Picture = imgId32.Picture Then
    imgId3.Picture = imgId31.Picture
End If
End Sub

Private Sub imgMinimize_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 Then
    imgMinimize.Picture = imgMinimize2.Picture
End If
End Sub

Private Sub imgMinimize_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 And imgMinimize.Picture = imgMinimize2.Picture Then
    If x > 250 Or x < -1 Or Y > 250 Or Y < -1 Then imgMinimize.Picture = imgMinimize1.Picture
ElseIf Button = 1 And imgMinimize.Picture = imgMinimize1.Picture Then
    If x < 250 And x > -1 And Y < 250 And Y > -1 Then imgMinimize.Picture = imgMinimize2.Picture
ElseIf Button = 0 Then
    If lEvents.eEventBusy = False Then
        lEvents.eDescription = "minimize"
        lEvents.ePercent = 0
        ConvertCaption
    End If
End If
End Sub

Private Sub imgMinimize_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 And imgMinimize.Picture = imgMinimize2.Picture Then
    imgMinimize.Picture = imgMinimize1.Picture
    WindowState = vbMinimized
End If
End Sub

Private Sub imgPlayMp3_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 Then
    imgPlayMp3.Picture = imgPlayMp32.Picture
End If
End Sub

Private Sub imgPlayMp3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 And imgPlayMp3.Picture = imgPlayMp32.Picture Then
    If x > 500 Or x < -1 Or Y > 500 Or Y < -1 Then imgPlayMp3.Picture = imgPlayMp31.Picture
ElseIf Button = 1 And imgPlayMp3.Picture = imgPlayMp31.Picture Then
    If x < 500 And x > -1 And Y < 500 And Y > -1 Then imgPlayMp3.Picture = imgPlayMp32.Picture
End If
End Sub

Private Sub imgPlayMp3_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 And imgPlayMp3.Picture = imgPlayMp32.Picture Then
    imgPlayMp3.Picture = imgPlayMp31.Picture
    'Shell App.Path & "\Audica.exe", vbNormalFocus
End If
End Sub

Private Sub imgPlayWav_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 Then
    imgPlayWav.Picture = imgPlayWav2.Picture
End If
End Sub

Private Sub imgPlayWav_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 And imgPlayWav.Picture = imgPlayWav2.Picture Then
    If x > 500 Or x < -1 Or Y > 500 Or Y < -1 Then imgPlayWav.Picture = imgPlayWav1.Picture
ElseIf Button = 1 And imgPlayWav.Picture = imgPlayWav1.Picture Then
    If x < 500 And x > -1 And Y < 500 And Y > -1 Then imgPlayWav.Picture = imgPlayWav2.Picture
ElseIf Button = 0 Then
    If lEvents.eEventBusy = False Then
        lEvents.eDescription = "playwav"
        lEvents.ePercent = 0
        ConvertCaption
    End If
End If
End Sub

Private Sub imgPlayWav_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 Then
    imgPlayWav.Picture = imgPlayWav1.Picture
End If
End Sub

Private Sub imgRip_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 Then
    imgRip.Picture = imgRip2.Picture
    tmrFlash.Enabled = False
    imgRipper.Picture = imgRipper1.Picture
End If
End Sub

Private Sub imgRip_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 And imgRip.Picture = imgRip2.Picture Then
    If x > 500 Or x < -1 Or Y > 500 Or Y < -1 Then imgRip.Picture = imgRip1.Picture
    
ElseIf Button = 1 And imgRip.Picture = imgRip1.Picture Then
    If x < 500 And x > -1 And Y < 500 And Y > -1 Then imgRip.Picture = imgRip2.Picture
ElseIf Button = 0 Then
    If lEvents.eEventBusy = False Then
        lEvents.eDescription = "rip"
        lEvents.ePercent = 0
        ConvertCaption
    End If
End If
End Sub

Private Sub imgRip_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim i As Integer, m As Integer
If Button = 1 And imgRip.Picture = imgRip2.Picture Then
    imgRip.Picture = imgRip1.Picture
    tmrFlash.Enabled = False
    imgRipper.Picture = imgRipper1.Picture
    lblInfo.Caption = "idle"
    frmPlaylist.Show 1
    
End If
End Sub

Private Sub imgRipper_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
lEvents.eDescription = "idle"
tmrFlash.Enabled = False
ConvertCaption
End Sub

Private Sub imgStopEncoding_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 Then
    imgStopEncoding.Picture = imgStopEncoding2.Picture
End If
End Sub

Private Sub imgStopEncoding_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 And imgStopEncoding.Picture = imgStopEncoding2.Picture Then
    If x > 500 Or x < -1 Or Y > 500 Or Y < -1 Then imgStopEncoding.Picture = imgStopEncoding1.Picture
ElseIf Button = 1 And imgStopEncoding.Picture = imgStopEncoding1.Picture Then
    If x < 500 And x > -1 And Y < 500 And Y > -1 Then imgStopEncoding.Picture = imgStopEncoding2.Picture
ElseIf Button = 0 Then
    If lEvents.eEventBusy = False Then
        lEvents.eDescription = "stopencode"
        lEvents.ePercent = 0
        ConvertCaption
        tmrFlash.Enabled = True
    End If
End If
End Sub

Private Sub imgStopEncoding_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 And imgStopEncoding.Picture = imgStopEncoding2.Picture Then
    imgStopEncoding.Picture = imgStopEncoding1.Picture
End If
End Sub

Private Sub Ripper_ActPosition(ByVal Position As Long)
Dim i As Long
lEvents.eEventBusy = True
i = Position / Ripper.GetTrackLength(lEvents.eEvent(lEvents.eEventCount).eTrack) * 100 / 1.5
lEvents.eDescription = "ripping"
If tmrFlash.Enabled = False Then tmrFlash.Enabled = True
RipCircleEffect Str(i)

lEvents.ePercent = i

lblInfo.Caption = "Ripping " & i & "% Complete"
End Sub

Private Sub Ripper_CopyStart()
'picPercent2.Visible = True
End Sub

Private Sub Ripper_CopyStop()
ResetRipperCircles
'lEvents.eEventBusy = False
lblWavFile.Caption = ""
tmrFlash.Enabled = False
imgRipper.Picture = imgRipper1.Picture
lblInfo.Caption = "idle"
EncodeFile lEvents.eCurrentFilename, Left(lEvents.eCurrentFilename, Len(lEvents.eCurrentFilename) - 3) & ".mp3"
End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub tmrFlash_Timer()
Dim msg As String
msg = lEvents.eDescription

'Stop
If msg = "rip" Or msg = "stoprip" Or msg = "playwav" Then
    imgFlashEncode.Picture = imgFlashEncode1.Picture
    If imgRipper.Picture = imgRipper1.Picture Then
        imgRipper.Picture = imgRipper2.Picture
    Else
        imgRipper.Picture = imgRipper1.Picture
    End If
ElseIf msg = "encode" Or msg = "stopencode" Then
    imgRipper.Picture = imgRipper1.Picture
    If imgFlashEncode.Picture = imgFlashEncode1.Picture Then
        imgFlashEncode.Picture = imgFlashEncode2.Picture
    Else
        imgFlashEncode.Picture = imgFlashEncode1.Picture
    End If
End If
End Sub

Private Sub tmrProcessEvent_Timer()
If lEvents.eEventBusy = False Then
    If lEvents.eEventCount = 0 Then
        tmrProcessEvent.Enabled = False
        Exit Sub
    End If
    ProcessNextEvent
End If
End Sub

Private Sub tmrShowEncoderCircles_Timer()
If lEvents.eCircleNum = 12 Then
    lEvents.eCircleNum = 0
    tmrShowEncoderCircles.Enabled = False
End If
shpEncoder(lEvents.eCircleNum).Visible = True
lEvents.eCircleNum = lEvents.eCircleNum + 1
End Sub

Private Sub tmrShowRipperCircles_Timer()
If lEvents.eCircleNum = 12 Then
    lEvents.eCircleNum = 0
    tmrShowRipperCircles.Enabled = False
End If
shpRipper(lEvents.eCircleNum).Visible = True
lEvents.eCircleNum = lEvents.eCircleNum + 1
End Sub
