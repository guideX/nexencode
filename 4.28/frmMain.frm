VERSION 5.00
Object = "{9F5F61C6-83A0-11D2-A800-00A0CC20D781}#1.0#0"; "ACD.OCX"
Object = "{FFBEC4C3-839E-11D1-85FE-0020AFE4DE54}#1.0#0"; "Mp3Enc.ocx"
Object = "{3B00B10A-6EF0-11D1-A6AA-0020AFE4DE54}#1.0#0"; "mp3play.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{60819404-3CCE-11D2-A800-008048E89E3E}#1.0#0"; "Effect.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loading NexENCODE"
   ClientHeight    =   1815
   ClientLeft      =   375
   ClientTop       =   765
   ClientWidth     =   6720
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
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   121
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   448
   Visible         =   0   'False
   Begin VB.PictureBox picButtons 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   9615
      Left            =   0
      ScaleHeight     =   641
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   0
      Top             =   -120
      Visible         =   0   'False
      Width           =   900
      Begin VB.Image imgNexMEDIA3 
         Height          =   300
         Left            =   120
         Top             =   8760
         Width           =   300
      End
      Begin VB.Image imgCancelRip3 
         Height          =   300
         Left            =   480
         Top             =   8760
         Width           =   300
      End
      Begin VB.Image imgForward3 
         Height          =   300
         Left            =   120
         Top             =   8400
         Width           =   300
      End
      Begin VB.Image imgBackward3 
         Height          =   300
         Left            =   480
         Top             =   8400
         Width           =   300
      End
      Begin VB.Image imgPlay3 
         Height          =   300
         Left            =   120
         Top             =   8040
         Width           =   300
      End
      Begin VB.Image imgStop3 
         Height          =   300
         Left            =   480
         Top             =   8040
         Width           =   300
      End
      Begin VB.Image imgEncode3 
         Height          =   300
         Left            =   120
         Top             =   7680
         Width           =   300
      End
      Begin VB.Image imgRip3 
         Height          =   300
         Left            =   480
         Top             =   7680
         Width           =   300
      End
      Begin VB.Image imgPlayMP33 
         Height          =   300
         Left            =   120
         Top             =   7320
         Width           =   300
      End
      Begin VB.Image imgPlayWAV3 
         Height          =   300
         Left            =   480
         Top             =   7320
         Width           =   300
      End
      Begin VB.Image imgStopEncoding3 
         Height          =   300
         Left            =   120
         Top             =   6960
         Width           =   300
      End
      Begin VB.Image imgEnd3 
         Height          =   300
         Left            =   480
         Top             =   6960
         Width           =   300
      End
      Begin VB.Image imgMinimize3 
         Height          =   300
         Left            =   120
         Top             =   6600
         Width           =   300
      End
      Begin VB.Image imgOptions3 
         Height          =   300
         Left            =   480
         Top             =   6600
         Width           =   300
      End
      Begin VB.Image imgID33 
         Height          =   300
         Left            =   120
         Top             =   6240
         Width           =   300
      End
      Begin VB.Image imgSkinEdit3 
         Height          =   300
         Left            =   480
         Top             =   6240
         Width           =   300
      End
      Begin VB.Image imgBackground1 
         Height          =   300
         Left            =   480
         Top             =   5880
         Width           =   300
      End
      Begin VB.Image imgBackground2 
         Height          =   300
         Left            =   120
         Top             =   5880
         Width           =   300
      End
      Begin VB.Image imgForward2 
         Height          =   300
         Left            =   480
         MousePointer    =   1  'Arrow
         Top             =   5160
         Width           =   300
      End
      Begin VB.Image imgForward1 
         Height          =   300
         Left            =   120
         MousePointer    =   1  'Arrow
         Top             =   5520
         Width           =   300
      End
      Begin VB.Image imgBackward1 
         Height          =   300
         Left            =   120
         MousePointer    =   1  'Arrow
         Top             =   3360
         Width           =   300
      End
      Begin VB.Image imgBackward2 
         Height          =   300
         Left            =   120
         MousePointer    =   1  'Arrow
         Top             =   2640
         Width           =   300
      End
      Begin VB.Image imgStop2 
         Height          =   300
         Left            =   480
         MousePointer    =   1  'Arrow
         Top             =   2640
         Width           =   300
      End
      Begin VB.Image imgStop1 
         Height          =   300
         Left            =   480
         MousePointer    =   1  'Arrow
         Top             =   3360
         Width           =   300
      End
      Begin VB.Image imgPlay1 
         Height          =   300
         Left            =   480
         MousePointer    =   1  'Arrow
         Top             =   1920
         Width           =   300
      End
      Begin VB.Image imgPlay2 
         Height          =   300
         Left            =   120
         MousePointer    =   1  'Arrow
         Top             =   3000
         Width           =   300
      End
      Begin VB.Image imgOptions2 
         Height          =   300
         Left            =   480
         MousePointer    =   1  'Arrow
         Top             =   2280
         Width           =   300
      End
      Begin VB.Image imgOptions1 
         Height          =   300
         Left            =   120
         MousePointer    =   1  'Arrow
         Top             =   2280
         Width           =   300
      End
      Begin VB.Image imgNexMedia2 
         Height          =   300
         Left            =   480
         MousePointer    =   1  'Arrow
         Top             =   1560
         Width           =   300
      End
      Begin VB.Image imgNexMedia1 
         Height          =   300
         Left            =   120
         MousePointer    =   1  'Arrow
         Top             =   1200
         Width           =   300
      End
      Begin VB.Image imgEncode2 
         Height          =   300
         Left            =   120
         MousePointer    =   1  'Arrow
         Top             =   840
         Width           =   300
      End
      Begin VB.Image imgSkinEdit2 
         Height          =   300
         Left            =   480
         MousePointer    =   1  'Arrow
         Top             =   840
         Width           =   300
      End
      Begin VB.Image imgSkinEdit1 
         Height          =   300
         Left            =   480
         MousePointer    =   1  'Arrow
         Top             =   4800
         Width           =   300
      End
      Begin VB.Image imgEnd1 
         Height          =   300
         Left            =   480
         MousePointer    =   1  'Arrow
         Top             =   3000
         Width           =   300
      End
      Begin VB.Image imgEnd2 
         Height          =   300
         Left            =   120
         MousePointer    =   1  'Arrow
         Top             =   1920
         Width           =   300
      End
      Begin VB.Image imgMinimize2 
         Height          =   300
         Left            =   480
         MousePointer    =   1  'Arrow
         Top             =   4080
         Width           =   300
      End
      Begin VB.Image imgMinimize1 
         Height          =   300
         Left            =   480
         MousePointer    =   1  'Arrow
         Top             =   3720
         Width           =   300
      End
      Begin VB.Image imgId32 
         Height          =   300
         Left            =   120
         MousePointer    =   1  'Arrow
         Top             =   3720
         Width           =   300
      End
      Begin VB.Image imgId31 
         Height          =   300
         Left            =   120
         MousePointer    =   1  'Arrow
         Top             =   1560
         Width           =   300
      End
      Begin VB.Image imgPlayMp32 
         Height          =   300
         Left            =   120
         MousePointer    =   1  'Arrow
         Top             =   4440
         Width           =   300
      End
      Begin VB.Image imgPlayMp31 
         Height          =   315
         Left            =   120
         MousePointer    =   1  'Arrow
         Top             =   4080
         Width           =   315
      End
      Begin VB.Image imgStopEncoding2 
         Height          =   300
         Left            =   480
         MousePointer    =   1  'Arrow
         Top             =   1200
         Width           =   300
      End
      Begin VB.Image imgStopEncoding1 
         Height          =   315
         Left            =   480
         MousePointer    =   1  'Arrow
         Top             =   4440
         Width           =   315
      End
      Begin VB.Image imgEncode1 
         Height          =   315
         Left            =   120
         MousePointer    =   1  'Arrow
         Top             =   4800
         Width           =   315
      End
      Begin VB.Image imgPlayWav2 
         Height          =   300
         Left            =   480
         MousePointer    =   1  'Arrow
         Top             =   5520
         Width           =   300
      End
      Begin VB.Image imgPlayWav1 
         Height          =   300
         Left            =   120
         MousePointer    =   1  'Arrow
         Top             =   5160
         Width           =   315
      End
      Begin VB.Image imgStopRipping2 
         Height          =   300
         Left            =   480
         MousePointer    =   1  'Arrow
         Top             =   480
         Width           =   300
      End
      Begin VB.Image imgStopRipping1 
         Height          =   300
         Left            =   480
         MousePointer    =   1  'Arrow
         Top             =   120
         Width           =   315
      End
      Begin VB.Image imgRip2 
         Height          =   300
         Left            =   120
         MousePointer    =   1  'Arrow
         Top             =   480
         Width           =   300
      End
      Begin VB.Image imgRip1 
         Height          =   300
         Left            =   120
         MousePointer    =   1  'Arrow
         Top             =   120
         Width           =   315
      End
   End
   Begin VB.PictureBox picActiveX 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   840
      ScaleHeight     =   975
      ScaleWidth      =   5895
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   5895
      Begin VB.Timer tmrPlayCommand 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   5160
         Top             =   0
      End
      Begin EFFECTLib.Effect ns4Effects 
         Height          =   495
         Left            =   3240
         TabIndex        =   9
         Top             =   0
         Width           =   495
         _Version        =   65536
         _ExtentX        =   873
         _ExtentY        =   873
         _StockProps     =   0
      End
      Begin VB.Timer tmrCheckActive 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   4800
         Top             =   0
      End
      Begin MSWinsockLib.Winsock wskUpdate 
         Left            =   2880
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wskFreeDB 
         Left            =   2520
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         LocalPort       =   4653
      End
      Begin VB.Timer tmrScrollStatus 
         Enabled         =   0   'False
         Interval        =   350
         Left            =   4440
         Top             =   0
      End
      Begin VB.Timer tmrShowEncoderCircles 
         Enabled         =   0   'False
         Interval        =   40
         Left            =   3720
         Top             =   0
      End
      Begin VB.Timer tmrShowRipperCircles 
         Enabled         =   0   'False
         Interval        =   40
         Left            =   4080
         Top             =   0
      End
      Begin MP3ENCLib.Mp3Enc Encoder 
         Height          =   495
         Left            =   2040
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
         Left            =   1560
         TabIndex        =   6
         Top             =   0
         Width           =   495
         _Version        =   65536
         _ExtentX        =   873
         _ExtentY        =   873
         _StockProps     =   0
      End
      Begin MPEGPLAYLib.Mp3Play SimpleMP3 
         Height          =   735
         Left            =   120
         TabIndex        =   7
         Top             =   0
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   0
      End
      Begin MPEGPLAYLib.Mp3Play Decoder 
         Height          =   735
         Left            =   840
         TabIndex        =   8
         Top             =   0
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   0
      End
   End
   Begin VB.Shape shpRipperColor 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Left            =   120
      Shape           =   2  'Oval
      Top             =   1440
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpEncoderColor 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Left            =   360
      Shape           =   2  'Oval
      Top             =   1440
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image imgErrorBackground 
      Height          =   300
      Left            =   2400
      MouseIcon       =   "frmMain.frx":000C
      MousePointer    =   99  'Custom
      Top             =   1080
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgStop 
      Height          =   300
      Left            =   1320
      MouseIcon       =   "frmMain.frx":015E
      MousePointer    =   99  'Custom
      ToolTipText     =   "Stop"
      Top             =   1080
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgBackward 
      Height          =   300
      Left            =   1680
      MouseIcon       =   "frmMain.frx":02B0
      MousePointer    =   99  'Custom
      ToolTipText     =   "Backward"
      Top             =   1080
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgForward 
      Height          =   300
      Left            =   2040
      MouseIcon       =   "frmMain.frx":0402
      MousePointer    =   99  'Custom
      ToolTipText     =   "Forward"
      Top             =   1080
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgPlay 
      Height          =   300
      Left            =   960
      MouseIcon       =   "frmMain.frx":0554
      MousePointer    =   99  'Custom
      ToolTipText     =   "Play MP3"
      Top             =   1080
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgOptions 
      Height          =   300
      Left            =   960
      MouseIcon       =   "frmMain.frx":06A6
      MousePointer    =   99  'Custom
      ToolTipText     =   "Options"
      Top             =   1440
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgNexMedia 
      Height          =   300
      Left            =   3480
      MouseIcon       =   "frmMain.frx":07F8
      MousePointer    =   99  'Custom
      ToolTipText     =   "Play CDAUDIO"
      Top             =   1080
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label lblMp3File 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   4560
      OLEDropMode     =   1  'Manual
      TabIndex        =   2
      Top             =   1440
      Width           =   300
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblWavFile 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   4200
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   1080
      Width           =   300
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgSkinEdit 
      Height          =   300
      Left            =   2760
      MouseIcon       =   "frmMain.frx":094A
      MousePointer    =   99  'Custom
      ToolTipText     =   "Edit Skin"
      Top             =   1080
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Shape shpEncoder 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   9
      Left            =   360
      Shape           =   2  'Oval
      Top             =   120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpEncoder 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   8
      Left            =   360
      Shape           =   2  'Oval
      Top             =   240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpEncoder 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   7
      Left            =   360
      Shape           =   2  'Oval
      Top             =   360
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpEncoder 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   6
      Left            =   360
      Shape           =   2  'Oval
      Top             =   480
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpEncoder 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   5
      Left            =   360
      Shape           =   2  'Oval
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpEncoder 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   4
      Left            =   360
      Shape           =   2  'Oval
      Top             =   720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpEncoder 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   3
      Left            =   360
      Shape           =   2  'Oval
      Top             =   840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpEncoder 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   2
      Left            =   360
      Shape           =   2  'Oval
      Top             =   960
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpEncoder 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   1
      Left            =   360
      Shape           =   2  'Oval
      Top             =   1080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpEncoder 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   0
      Left            =   360
      Shape           =   2  'Oval
      Top             =   1200
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpRipper 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   9
      Left            =   120
      Shape           =   2  'Oval
      Top             =   1200
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpRipper 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   8
      Left            =   120
      Shape           =   2  'Oval
      Top             =   840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpRipper 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   7
      Left            =   120
      Shape           =   2  'Oval
      Top             =   120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpRipper 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   6
      Left            =   120
      Shape           =   2  'Oval
      Top             =   1080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpRipper 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   5
      Left            =   120
      Shape           =   2  'Oval
      Top             =   360
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpRipper 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   4
      Left            =   120
      Shape           =   2  'Oval
      Top             =   960
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpRipper 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   3
      Left            =   120
      Shape           =   2  'Oval
      Top             =   480
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpRipper 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   2
      Left            =   120
      Shape           =   2  'Oval
      Top             =   240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpRipper 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   1
      Left            =   120
      Shape           =   2  'Oval
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpRipper 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   150
      Index           =   0
      Left            =   120
      Shape           =   2  'Oval
      Top             =   720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image imgEnd 
      Height          =   300
      Left            =   1680
      MouseIcon       =   "frmMain.frx":0A9C
      MousePointer    =   99  'Custom
      ToolTipText     =   "Power"
      Top             =   1440
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgMinimize 
      Height          =   300
      Left            =   2040
      MouseIcon       =   "frmMain.frx":0BEE
      MousePointer    =   99  'Custom
      ToolTipText     =   "Minimize"
      Top             =   1440
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgId3 
      Height          =   300
      Left            =   3120
      MouseIcon       =   "frmMain.frx":0D40
      MousePointer    =   99  'Custom
      ToolTipText     =   "Effects - Open .wav file and add effects to it"
      Top             =   1080
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgPlayMp3 
      Height          =   300
      Left            =   3840
      MouseIcon       =   "frmMain.frx":0E92
      MousePointer    =   99  'Custom
      ToolTipText     =   "Decode"
      Top             =   1440
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgStopEncoding 
      Height          =   300
      Left            =   2760
      MouseIcon       =   "frmMain.frx":0FE4
      MousePointer    =   99  'Custom
      ToolTipText     =   "Stop Encoding"
      Top             =   1440
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgEncode 
      Height          =   300
      Left            =   1320
      MouseIcon       =   "frmMain.frx":1136
      MousePointer    =   99  'Custom
      ToolTipText     =   "Encode"
      Top             =   1440
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgPlayWav 
      Height          =   300
      Left            =   3480
      MouseIcon       =   "frmMain.frx":1288
      MousePointer    =   99  'Custom
      ToolTipText     =   "Search for mp3 files"
      Top             =   1440
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgCancelRip 
      Height          =   300
      Left            =   3120
      MouseIcon       =   "frmMain.frx":13DA
      MousePointer    =   99  'Custom
      ToolTipText     =   "Cancel Rip"
      Top             =   1440
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgRip 
      Height          =   300
      Left            =   2400
      MouseIcon       =   "frmMain.frx":152C
      MousePointer    =   99  'Custom
      ToolTipText     =   "Rip CDAUDIO"
      Top             =   1440
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4200
      MouseIcon       =   "frmMain.frx":167E
      MousePointer    =   99  'Custom
      OLEDropMode     =   1  'Manual
      TabIndex        =   4
      Top             =   1440
      Width           =   285
   End
   Begin VB.Image imgPercent 
      Height          =   300
      Left            =   3840
      MouseIcon       =   "frmMain.frx":17D0
      MousePointer    =   99  'Custom
      Top             =   1080
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Menu mnuHidden 
      Caption         =   "Hidden"
      Visible         =   0   'False
      Begin VB.Menu mnuNexENCODE 
         Caption         =   "NexENCODE"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuBatch 
         Caption         =   "Batch"
      End
      Begin VB.Menu mnuPlaylist 
         Caption         =   "Playlist"
      End
      Begin VB.Menu mnuSep937927322343 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenFile 
         Caption         =   "Eject"
         Begin VB.Menu mnuWavAudio 
            Caption         =   "Wave Audio (*.wav)"
         End
         Begin VB.Menu mnuOpenMP3 
            Caption         =   "Mpeg Layer 3 (*.mp3)"
         End
         Begin VB.Menu mnuOpenPlaylist 
            Caption         =   "Playlist (*.m3u)"
         End
         Begin VB.Menu mnuSep397889273 
            Caption         =   "-"
         End
         Begin VB.Menu mnuHelp2 
            Caption         =   "Help"
         End
      End
      Begin VB.Menu mnuSkin 
         Caption         =   "Skins"
         Begin VB.Menu mnuLoadSkin 
            Caption         =   "Load"
         End
         Begin VB.Menu mnuSep937927 
            Caption         =   "-"
         End
         Begin VB.Menu mnuNexSkin 
            Caption         =   "Editor"
         End
         Begin VB.Menu mnuSkinBrowser 
            Caption         =   "Browser"
         End
         Begin VB.Menu mnuMoreSkinsOnline 
            Caption         =   "More Skins"
         End
         Begin VB.Menu mnuSep3978297392 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSkinName 
            Caption         =   "Nothing"
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu mnuSep9379237892468369 
            Caption         =   "-"
         End
         Begin VB.Menu mnuHelp4 
            Caption         =   "Help"
         End
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "Settings"
         Begin VB.Menu mnuEncoder 
            Caption         =   "Encoder"
         End
         Begin VB.Menu mnuRipperSettings 
            Caption         =   "Ripper"
         End
         Begin VB.Menu mnuPlayer_ 
            Caption         =   "Players"
         End
         Begin VB.Menu mnuGeneralSettings 
            Caption         =   "General"
         End
         Begin VB.Menu mnuCDDBOptions 
            Caption         =   "FreeDB"
         End
         Begin VB.Menu mnuSettingsASPI 
            Caption         =   "ASPI"
         End
         Begin VB.Menu mnuRegistered 
            Caption         =   "Register"
         End
         Begin VB.Menu mnuSep937293792 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSetupWizard 
            Caption         =   "Wizard"
         End
         Begin VB.Menu mnuDeleteSettings 
            Caption         =   "Delete Settings"
         End
         Begin VB.Menu mnuOnTop 
            Caption         =   "On Top"
         End
         Begin VB.Menu mnuSep938792733 
            Caption         =   "-"
         End
         Begin VB.Menu mnuHelp3 
            Caption         =   "Help"
         End
      End
      Begin VB.Menu mnuSep83298932 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEncode 
         Caption         =   "Encode"
         Begin VB.Menu mnuSingleFromDisk 
            Caption         =   "Encode (wav to mp3)"
         End
         Begin VB.Menu mnuMp3ToWav2 
            Caption         =   "Decode (mp3 to wav)"
         End
         Begin VB.Menu mnuSep797923 
            Caption         =   "-"
         End
         Begin VB.Menu mnuMergeMp3 
            Caption         =   "Merge"
         End
         Begin VB.Menu mnuSep34343 
            Caption         =   "-"
         End
         Begin VB.Menu mnuHelp 
            Caption         =   "Help"
         End
      End
      Begin VB.Menu mnuRip 
         Caption         =   "Rip"
         Begin VB.Menu mnuRipCDAToWav 
            Caption         =   "Rip (cda to wav)"
         End
         Begin VB.Menu mnuRipCDATOMP3 
            Caption         =   "Rip (cda to mp3)"
         End
         Begin VB.Menu mnuSep937927932 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEntireDiscToMp3 
            Caption         =   "Encode Disc"
         End
         Begin VB.Menu mnuEntireDisctoWav 
            Caption         =   "Rip Disc"
         End
         Begin VB.Menu mnuSep8937892689763 
            Caption         =   "-"
         End
         Begin VB.Menu mnuHelp5 
            Caption         =   "Help"
         End
      End
      Begin VB.Menu mnuPlayer 
         Caption         =   "Player"
         Begin VB.Menu mnuOpen 
            Caption         =   "Quick Play"
         End
         Begin VB.Menu mnuSep3979273 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAudica 
            Caption         =   "Player"
         End
         Begin VB.Menu mnuNexMedia 
            Caption         =   "CD Audio"
         End
         Begin VB.Menu mnuSep3979 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPlay 
            Caption         =   "Play"
         End
         Begin VB.Menu mnuPause 
            Caption         =   "Pause"
         End
         Begin VB.Menu mnuStopMp3 
            Caption         =   "Stop"
         End
         Begin VB.Menu mnuBackward 
            Caption         =   "Previous"
         End
         Begin VB.Menu mnuForward 
            Caption         =   "Next"
         End
         Begin VB.Menu mnuSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAdjustVolume 
            Caption         =   "Adjust Volume"
         End
         Begin VB.Menu mnuSpeed 
            Caption         =   "Speed"
            Begin VB.Menu mnu2xSlow 
               Caption         =   "2x Slow"
            End
            Begin VB.Menu mnu1xSlow 
               Caption         =   "1x Slow"
            End
            Begin VB.Menu mnuNormal 
               Caption         =   "Normal"
               Checked         =   -1  'True
            End
            Begin VB.Menu mnu1xFast 
               Caption         =   "1x Fast"
            End
            Begin VB.Menu mnu2xFast 
               Caption         =   "2x Fast"
            End
         End
         Begin VB.Menu mnuRandom 
            Caption         =   "Random"
         End
         Begin VB.Menu mnuContinuous 
            Caption         =   "Shuffle Mode"
         End
         Begin VB.Menu mnuSep93879273 
            Caption         =   "-"
         End
         Begin VB.Menu mnuHelp6 
            Caption         =   "Help"
         End
      End
      Begin VB.Menu mnuEffectsEditor2 
         Caption         =   "Effects"
         Begin VB.Menu mnuShowWavEffectsEditor 
            Caption         =   "Editor"
         End
         Begin VB.Menu mnuSep9739273978684 
            Caption         =   "-"
         End
         Begin VB.Menu mnuHelp9 
            Caption         =   "Help"
         End
      End
      Begin VB.Menu mnuSep938792739 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDiscInfo 
         Caption         =   "Disc"
         Begin VB.Menu mnuArtist 
            Caption         =   "Artist: <None>"
         End
         Begin VB.Menu mnuAlbum 
            Caption         =   "Album: <None>"
         End
         Begin VB.Menu mnuSep937927893 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEditTracks 
            Caption         =   "Edit"
         End
         Begin VB.Menu mnuDownloadTracks 
            Caption         =   "Download"
         End
      End
      Begin VB.Menu mnuMore 
         Caption         =   "More"
         Begin VB.Menu mnuMusicSearch 
            Caption         =   "Audio Galaxy Search"
         End
         Begin VB.Menu mnuSearchHardDrive 
            Caption         =   "Search Hard Drives"
         End
         Begin VB.Menu mnuSearchWithinPlaylists 
            Caption         =   "Search Playlists"
         End
         Begin VB.Menu mnuSep893727392 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCheckforUpdates 
            Caption         =   "Check for Updates"
         End
         Begin VB.Menu mnuSep3879273 
            Caption         =   "-"
         End
         Begin VB.Menu mnuReadMe 
            Caption         =   "Documentation"
         End
         Begin VB.Menu mnuSupport 
            Caption         =   "Support"
         End
         Begin VB.Menu mnuSep3972973 
            Caption         =   "-"
         End
         Begin VB.Menu mnushowTagEditor 
            Caption         =   "Tag editor"
         End
         Begin VB.Menu mnuUpdateASPI 
            Caption         =   "Update ASPI"
         End
      End
      Begin VB.Menu mnuWeb 
         Caption         =   "Web"
         Begin VB.Menu mnuTeamNexgen 
            Caption         =   "Team Nexgen"
         End
         Begin VB.Menu mnuSep93792793 
            Caption         =   "-"
         End
         Begin VB.Menu mnuMp3Website 
            Caption         =   "MP3.com"
         End
         Begin VB.Menu mnuKaza 
            Caption         =   "KaZzA"
         End
         Begin VB.Menu mnuAudioGalWeb 
            Caption         =   "AudioGalaxy"
         End
         Begin VB.Menu mnuRollingStonesWebsite 
            Caption         =   "Rolling Stones"
         End
         Begin VB.Menu mnuSep9387927392 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEmailDeveloper 
            Caption         =   "E-Mail Leon Aiossa"
         End
         Begin VB.Menu mnuEmailKnightFal 
            Caption         =   "E-Mail Colin Foss"
         End
      End
      Begin VB.Menu mnuSep093879273 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHide 
         Caption         =   "Minimize"
      End
      Begin VB.Menu mnuPower 
         Caption         =   "Power"
      End
      Begin VB.Menu mnuCancelSep 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuWaveMenu 
      Caption         =   "Wave Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuOpenWave 
         Caption         =   "Eject"
      End
      Begin VB.Menu mnuSep93789273 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEffects 
         Caption         =   "Effects"
         Begin VB.Menu mnuOpenCurrentFile 
            Caption         =   "Open"
         End
         Begin VB.Menu mnuSep89739723 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPlayEffectWav 
            Caption         =   "Play"
         End
         Begin VB.Menu mnuStopEffectWav 
            Caption         =   "Stop"
         End
         Begin VB.Menu mnuSaveWavAs 
            Caption         =   "Save As"
         End
         Begin VB.Menu mnuSep89379237 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEffectsEditor 
            Caption         =   "Editor"
         End
         Begin VB.Menu mnusep39873973 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAmplitude 
            Caption         =   "Amplitude"
         End
         Begin VB.Menu mnuChorus 
            Caption         =   "Chorus"
         End
         Begin VB.Menu mnuCFilter 
            Caption         =   "CFilter"
         End
         Begin VB.Menu mnuDistortion 
            Caption         =   "Distortion"
         End
         Begin VB.Menu mnuEcho 
            Caption         =   "Echo"
         End
         Begin VB.Menu mnuFadeIN 
            Caption         =   "Fade In"
         End
         Begin VB.Menu mnuFadeOut 
            Caption         =   "Fade Out"
         End
         Begin VB.Menu mnuInvert 
            Caption         =   "Invert"
         End
         Begin VB.Menu mnuShifting 
            Caption         =   "Shifting"
         End
         Begin VB.Menu mnuReverb 
            Caption         =   "Reverb"
         End
         Begin VB.Menu mnuSep927937 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCloseWavFile 
            Caption         =   "Close"
         End
      End
      Begin VB.Menu mnuSep3890927392 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConvertToMp3 
         Caption         =   "Convert to MP3"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuOpenInSoundRec 
         Caption         =   "Edit"
      End
      Begin VB.Menu mnuSep973927 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCloseWav 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuMpegMenu 
      Caption         =   "Mpeg Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuOpenMpeg 
         Caption         =   "Eject"
      End
      Begin VB.Menu mnSep39872987392 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlayMpeg 
         Caption         =   "Play"
      End
      Begin VB.Menu mnuPauseMpeg 
         Caption         =   "Pause"
      End
      Begin VB.Menu mnuStopMpeg 
         Caption         =   "Stop"
      End
      Begin VB.Menu mnusep6996789568587589 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBackwardMpeg 
         Caption         =   "Backward"
      End
      Begin VB.Menu mnuForwardMpeg 
         Caption         =   "Forward"
      End
      Begin VB.Menu mnuSep9379273 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConvertToWav 
         Caption         =   "Convert to WAV"
      End
      Begin VB.Menu mnuMoreInfo 
         Caption         =   "More Info"
      End
      Begin VB.Menu mnuRandom2 
         Caption         =   "Random"
      End
      Begin VB.Menu mnuSep8937298382964290 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lUpdateINI As String
Dim lSpindle As Integer

Public Sub InitMain()
On Local Error Resume Next
Dim msg As String, msg2 As String, i As Integer
Left = ReadINI(lIniFiles.iSettings, "MainWind", "Left", 0)
Top = ReadINI(lIniFiles.iSettings, "MainWind", "Top", 0)
LoadPlaylists
Ripper.Authorize "Leon Aiossa", "698070606"
Encoder.Authorize "Leon Aiossa", "680665552"
ns4Effects.Authorize "Leon J Aiossa", "1081841574"
Ripper.Init
If frmMain.Ripper.IsAspiLoaded = False Then
    lRipperSettings.eAspiEnabled = False
Else
    lRipperSettings.eAspiEnabled = True
End If
GetSimpleTracks
ResetPlayButtons
SetCaption Download
ConvertCaption oIdle
ToggleButtons oIdle
If Err.Number <> 0 Then SetError "InitMain()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Encoder_ActFrame(ByVal ActFrame As Long)
On Local Error Resume Next
Dim i As Long
If imgPercent.Visible = False Then
    imgPercent.Width = 1
    imgPercent.Visible = True
End If
If lEvents.eEncoderBusy = False Then lEvents.eEncoderBusy = True
i = (ActFrame * 100 / Encoder.GetFrameCount) \ 1
If i < 101 And i > -1 Then
    imgPercent.Width = i * 1.17
    EncodeCircleEffect Str(i)
    lEvents.ePercent = i
    lblInfo.Caption = "Encode: " & i & "%"
    SetCaption Encode, i, frmMain.lblMp3File.Caption
    If lEncWizard.eEnabled = True Then
        frmEncoderWizard.ProgressBar1.Value = i
        frmEncoderWizard.lblEncProgress.Caption = "Progress: " & i & "%"
    End If
Else
    lblInfo.Caption = "Encoding ..."
    lblWavFile.Caption = i
End If
If Err.Number <> 0 Then SetError "NEXENCODER_Actframe()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Encoder_Failure(ByVal ErrCode As Long, ByVal errStr As String)
On Local Error Resume Next
AddFinishedEvent Encode, Err.Description, ""
SetError "ENCODER_Failure", "The encoder ran into an error", errStr
If Err.Number <> 0 Then SetError "Encoder_Failure()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Encoder_ThreadEnded()
On Local Error Resume Next
If lEncWizard.eEnabled = True Then
    frmEncoderWizard.lblEncProgress = "Progress: Complete"
    If lEncWizard.eType = eSingleWav Then
        PlayWav App.Path & "\media\done.wav", SND_ASYNC
        frmEncoderWizard.NextEncWizardFrame
    ElseIf lEncWizard.eType = eMultiWav Then
        lEncWizard.eCount = lEncWizard.eCount - 1
        If lEncWizard.eCount = 0 Then
            PlayWav App.Path & "\media\done.wav", SND_ASYNC
            frmEncoderWizard.NextEncWizardFrame
        End If
    End If
End If
lblWavFile.Caption = ""
SetCaption Download
PlayWav App.Path & "\media\done.wav", SND_ASYNC
If Len(lEvents.eAutoDel) <> 0 Then
    If DoesFileExist(lEvents.eAutoDel) = True Then Kill lEvents.eAutoDel
    lEvents.eAutoDel = ""
End If
If lEncoderSettings.eAutoAddTags = True Then SaveTagInfo lTag.tFile
ToggleButtons oIdle
lEvents.eEncoderBusy = False
lblInfo.Caption = "Complete"
lblMp3File.Caption = ""
imgPercent.Width = 0
imgPercent.Visible = False
ResetEncoderCircles
AddFinishedEvent Encode, "Completed"
DoEvents
If lEvents.eSettings.iEnding = False Then ProcessNextEvent
If Err.Number <> 0 Then SetError "ENCODER_ThreadEnded()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Activate()
On Local Error Resume Next
If lEvents.eSettings.iCheckForActiveWindow = True Then IsActiveWindow
If Err.Number <> 0 Then SetError "Form_Activate()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Click()
On Local Error Resume Next
If lEvents.eSettings.iCheckForActiveWindow = True Then IsActiveWindow
If Err.Number <> 0 Then SetError "Form_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Deactivate()
On Local Error Resume Next
If lEvents.eSettings.iCheckForActiveWindow = True Then IsActiveWindow
If Err.Number <> 0 Then SetError "Form_Deactivate()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_GotFocus()
On Local Error Resume Next
If lEvents.eSettings.iCheckForActiveWindow = True Then IsActiveWindow
If Err.Number <> 0 Then SetError "Form_GotFocus()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Initialize()
On Local Error Resume Next
try.cbSize = Len(try)
try.hwnd = Me.hwnd
try.uId = vbNull
try.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
try.uCallBackMessage = WM_MOUSEMOVE
try.hIcon = Me.Icon
try.szTip = App.Title
Call Shell_NotifyIcon(NIM_ADD, try)
Call Shell_NotifyIcon(NIM_MODIFY, try)
If Err.Number <> 0 Then SetError "Form_Initialize", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Local Error Resume Next
If KeyAscii = 27 Then
    Me.Visible = False
    Me.WindowState = vbMinimized
    If lEvents.ePlaylistVisible = True Then Unload frmPlaylist
End If
If Err.Number <> 0 Then SetError "Form_KeyPress", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Load()
On Local Error Resume Next
Dim i As Integer, j As Integer, msg As String, msg2 As String
Me.Icon = frmGraphics.Icon
lEvents.eSettings.iLoading = True
lEffectsPresets.eSaved = True
CloseEffects
FindOpenDevice SimpleMP3
InitMain
DoEvents
If lEvents.eSettings.iUpdateCheck = True Then
    wskUpdate.Close
    wskUpdate.Connect "www.team-nexgen.com", 80
End If
If lEvents.eSettings.iShowAbout = True Then
    frmAbout.tmrDots.Enabled = False
    frmAbout.tmrUnload.Enabled = False
    Unload frmAbout
End If
InitSkins
DoEvents
If lEvents.eSettings.iAlwaysOnTop = True Then
    mnuOnTop.Checked = True
    AlwaysOnTop frmMain, True
End If
If lPlayer.pContinuous = True Then mnuContinuous.Checked = True
If lEvents.eRegistered = True Then frmMain.mnuRegistered.Visible = False
If lEvents.eSettings.iCheckForActiveWindow = True Then lEvents.eMainHWND = GetActiveWindow
If ReadINI(lIniFiles.iSettings, "PlaylistWind", "Visible", "True") = "True" Then frmPlaylist.Show
If Len(lEvents.eSettings.iCommand) <> 0 Then tmrPlayCommand.Enabled = True
lEvents.eSettings.iLoading = False
Playlist.pFileCount = Playlist.pFileCount + 1
If Err.Number <> 0 Then SetError "frmMain_Load", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_LostFocus()
On Local Error Resume Next
If lEvents.eSettings.iCheckForActiveWindow = True Then IsActiveWindow
If Err.Number <> 0 Then SetError "Form_LostFocus()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 Then
    FormDrag Me
Else
    PlayWav App.Path & "\media\click.wav", SND_ASYNC
    PopupMenu mnuHidden
End If
If Err.Number <> 0 Then SetError "frmMain_Mousedown()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
CheckMouseOver
Select Case X
Case 517
    frmMain.mnuCancel.Visible = True
    frmMain.mnuCancelSep.Visible = True
    frmMain.mnuNexENCODE.Checked = False
    PlayWav App.Path & "\media\click.wav", SND_ASYNC
    PopupMenu frmMain.mnuHidden
    frmMain.mnuNexENCODE.Checked = True
    frmMain.mnuCancel.Visible = False
    frmMain.mnuCancelSep.Visible = False
Case 515
    Dim b As Boolean
    Me.WindowState = vbNormal
    Me.Visible = True
    If lEvents.eSettings.iCheckForActiveWindow = True Then IsActiveWindow
    If lEvents.ePlaylistVisible = True Then frmPlaylist.Show
End Select
lEvents.ePercent = 0
ConvertCaption oIdle
If Err.Number <> 0 Then SetError "frmMain_Mousemove()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
DragDrop Data
If Err.Number <> 0 Then SetError "Form_OLEDragDrop()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Local Error Resume Next
If lEvents.eSettings.iEnding = False Then
    Cancel = 1
    UnloadMain
End If
If Err.Number <> 0 Then SetError "Form_Unload()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgBackward_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 Then
    imgBackward.Picture = imgBackward2.Picture
End If
If Err.Number <> 0 Then SetError "imgBackward_MouseDown()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgBackward_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
PictureBoxMouseMove oBackwardButton, Button, imgBackward, imgBackward1, imgBackward2, X, Y, imgBackward3, True
If Err.Number <> 0 Then SetError "imgBackward_MouseMove()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgBackward_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 And imgBackward.Picture = imgBackward2.Picture Then
    imgBackward.Picture = imgBackward1.Picture
    PlayWav App.Path & "\media\click.wav", SND_ASYNC
    PlaylistToHTMLFile False, 3
'    GoBackward
End If
If Err.Number <> 0 Then SetError "imgBackward_MouseUp()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgCancelRip_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 Then
    imgCancelRip.Picture = imgStopRipping2.Picture
End If
If Err.Number <> 0 Then SetError "imgCancelRip_MouseDown()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgCancelRip_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
PictureBoxMouseMove oStopRipping, Button, imgCancelRip, imgStopRipping1, imgStopRipping2, X, Y, imgCancelRip3, True
If Err.Number <> 0 Then SetError "imgCancelRip_MouseMove()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgCancelRip_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
Dim msg As String
If Button = 1 And imgCancelRip.Picture = imgStopRipping2.Picture Then
    ToggleButtons oIdle
    lEvents.eRipperBusy = False
    ResetRipperCircles
    tmrScrollStatus.Enabled = False
    Ripper.sTop
    If lEffectsPresets.eStatus <> ePlaying Then lblWavFile.Caption = ""
    If lEvents.eRipperBusy = True Then
        Ripper.sTop
    ElseIf lEffectsPresets.eStatus = eAddingEffect Then
        ns4Effects.StopEffect
        EnableEffects
    ElseIf lEffectsPresets.eStatus = ePlaying Then
        ns4Effects.sTop
        EnableEffects
    End If
    imgCancelRip.Picture = imgStopRipping1.Picture
End If
If Err.Number <> 0 Then SetError "imgCancelRip_MouseUp()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgEncode_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 Then
    imgEncode.Picture = imgEncode2.Picture
End If
If Err.Number <> 0 Then SetError "imgEncode_MouseDown()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgEncode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
PictureBoxMouseMove oEncode, Button, imgEncode, imgEncode1, imgEncode2, X, Y, imgEncode3, True
If Err.Number <> 0 Then SetError "imgEncode_MouseMove()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgEncode_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 And imgEncode.Picture = imgEncode2.Picture Then
    Dim msg As String, msg2 As String
    PlayWav App.Path & "\media\click2.wav", SND_ASYNC
    imgEncode.Picture = imgEncode1.Picture
    frmEncode.Show
End If
If Err.Number <> 0 Then SetError "imgEncode_mouseUp()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgEnd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 Then
    imgEnd.Picture = imgEnd2.Picture
End If
If Err.Number <> 0 Then SetError "imgEnd_mouseDown()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgEnd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
PictureBoxMouseMove oEnd, Button, imgEnd, imgEnd1, imgEnd2, X, Y, imgEnd3, True
If Err.Number <> 0 Then SetError "imgEnd_mouseMove()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgEnd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 And imgEnd.Picture = imgEnd2.Picture Then
    UnloadMain
End If
If Err.Number <> 0 Then SetError "imgEnd_MouseUp()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgForward_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 Then
    imgForward.Picture = imgForward2.Picture
End If
If Err.Number <> 0 Then SetError "imgForward_MouseDown()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgForward_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
PictureBoxMouseMove oForwardButton, Button, imgForward, imgForward1, imgForward2, X, Y, imgForward3, True
If Err.Number <> 0 Then SetError "imgForward_MouseMove()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgForward_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 And imgForward.Picture = imgForward2.Picture Then
    StopMp3
    imgForward.Picture = imgForward1.Picture
    PlayWav App.Path & "\media\click.wav", SND_ASYNC
    GoForward
End If
If Err.Number <> 0 Then SetError "imgforward_MouseUp()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgId3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 Then
    imgId3.Picture = imgId32.Picture
End If
If Err.Number <> 0 Then SetError "imgId3()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgId3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
PictureBoxMouseMove oTag, Button, imgId3, imgId31, imgId32, X, Y, imgID33, True
If Err.Number <> 0 Then SetError "imgId3_MouseMove()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgId3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
Dim msg As String, msg2 As String
If Button = 1 And imgId3.Picture = imgId32.Picture Then
    imgId3.Picture = imgId31.Picture
    If lEffectsPresets.eStatus = eOpen Then
        frmEffects.Show
        Exit Sub
    ElseIf lEffectsPresets.eStatus = eClosed And Len(lPlayer.pLabels.lWavFile) <> 0 Then
        OpenEffects lPlayer.pLabels.lWavPath & lPlayer.pLabels.lWavFile: DoEvents
        pause 1
        frmEffects.Show
        Exit Sub
    End If
    msg = OpenDialog(Me, "Wave Audio (*.wav)|*.wav", "Select Wave Audio", CurDir)
    If Len(msg) <> 0 Then
        If Right(LCase(msg), 4) <> ".wav" Then
            msg = msg & ".wav"
        End If
        msg2 = msg
        msg2 = GetFileTitle(msg2)
        OpenEffects msg: DoEvents
        SetFileLabels Left(msg, Len(msg) - Len(msg2)), msg2
        pause 1
        frmEffects.Show
    End If
    'LoadId3Editor
End If
If Err.Number <> 0 Then SetError "imgId3_MouseUp()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgMinimize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 Then
    imgMinimize.Picture = imgMinimize2.Picture
End If
If Err.Number <> 0 Then SetError "imgMinimize_MouseUp()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgMinimize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
PictureBoxMouseMove oMinimize, Button, imgMinimize, imgMinimize1, imgMinimize2, X, Y, imgMinimize3, True
If Err.Number <> 0 Then SetError "imgMinimize_MouseMove()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgMinimize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 And imgMinimize.Picture = imgMinimize2.Picture Then
    imgMinimize.Picture = imgMinimize1.Picture
    If mnuPlaylist.Checked = True Then
        If lEvents.ePlaylistVisible = True Then Unload frmPlaylist
        mnuPlaylist.Checked = False
    End If
    Me.Visible = False
    WindowState = vbMinimized
End If
If Err.Number <> 0 Then SetError "imgMinimize_MouseUp()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgNexMedia_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 Then
    imgNexMedia.Picture = imgNexMedia2.Picture
End If
If Err.Number <> 0 Then SetError "imgNexMedia_MouseDown()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgNexMedia_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
PictureBoxMouseMove oCDAudio, Button, imgNexMedia, imgNexMedia1, imgNexMedia2, X, Y, imgNexMEDIA3, True
If Err.Number <> 0 Then SetError "imgNexMedia_MouseMove()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgNexMedia_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 And imgNexMedia.Picture = imgNexMedia2.Picture Then
    GoCDPlayer
    imgNexMedia.Picture = imgNexMedia1.Picture
End If
If Err.Number <> 0 Then SetError "imgMinimize_MouseUp()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgOptions_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 Then
    imgOptions.Picture = imgOptions2.Picture
End If
If Err.Number <> 0 Then SetError "imgOptions_MouseDown()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
PictureBoxMouseMove oOptions, Button, imgOptions, imgOptions1, imgOptions2, X, Y, imgOptions3, True
If Err.Number <> 0 Then SetError "imgOptions_MouseMove()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgOptions_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 And imgOptions.Picture = imgOptions2.Picture Then
    imgOptions.Picture = imgOptions1.Picture
    PlayWav App.Path & "\media\click2.wav", SND_ASYNC
    frmSettings.ResetSettingsFrames eEncoder, True
End If
If Err.Number <> 0 Then SetError "imgOptions_MouseUp()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgPlay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 Then
    imgPlay.Picture = imgPlay2.Picture
End If
If Err.Number <> 0 Then SetError "imgPlay_MouseDown()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgPlay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
PictureBoxMouseMove oPlayButton, Button, imgPlay, imgPlay1, imgPlay2, X, Y, imgPlay3, True
If Err.Number <> 0 Then SetError "imgNexMedia_MouseMove()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgPlay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 And imgPlay.Picture = imgPlay2.Picture Then
    imgPlay.Picture = imgPlay1.Picture
    Select Case lPlayer.pStatus
    Case sPaused
        PauseMp3
    Case sPlaying
        PauseMp3
    Case sNotPlaying
        If Len(lblMp3File.Caption) <> 0 Then
            AddPlayEvent lPlayer.pLabels.lMp3Path, lPlayer.pLabels.lMp3File
        Else
            PlayWav App.Path & "\media\click.wav", SND_ASYNC
            PromptToPlay
        End If
    End Select
End If
If Err.Number <> 0 Then SetError "imgPlay_MouseUp()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgPlayMp3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 Then imgPlayMp3.Picture = imgPlayMp32.Picture
If Err.Number <> 0 Then SetError "imgPlayMp3_MouseDown()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgPlayMp3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
PictureBoxMouseMove oDecode, Button, imgPlayMp3, imgPlayMp31, imgPlayMp32, X, Y, imgPlayMP33, True
If Err.Number <> 0 Then SetError "imgPlayMp3_MouseMove()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgPlayMp3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
Dim msg As String, msg2 As String, lWavFile As String
If Button = 1 And imgPlayMp3.Picture = imgPlayMp32.Picture Then
    PlayWav App.Path & "\media\click2.wav", SND_ASYNC
    imgPlayMp3.Picture = imgPlayMp31.Picture
    If Len(lblMp3File.Caption) <> 0 Then
        lWavFile = Left(lPlayer.pLabels.lMp3File, Len(lPlayer.pLabels.lMp3File) - 4) & ".wav"
        AddEvent Decode, lPlayer.pLabels.lMp3Path, lPlayer.pLabels.lMp3File, lPlayer.pLabels.lMp3Path, lWavFile, 0, ""
    Else
        frmEncode.Show
        frmEncode.cboFormat.ListIndex = 1
    End If
End If
If Err.Number <> 0 Then SetError "imgPlayMp3_MouseUp()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgPlayWav_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 Then
    imgPlayWav.Picture = imgPlayWav2.Picture
End If
If Err.Number <> 0 Then SetError "imgPlayWav_MouseDown()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgPlayWav_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
PictureBoxMouseMove oPlayWav, Button, imgPlayWav, imgPlayWav1, imgPlayWav2, X, Y, imgPlayWAV3, True
If Err.Number <> 0 Then SetError "imgPlayWav_MouseMove()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgPlayWav_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If imgPlayWav.Picture = imgPlayWav2.Picture Then
    If Button = 1 Then
        imgPlayWav.Picture = imgPlayWav1.Picture
        frmSearch.Show
    ElseIf Button = 2 Then
        imgPlayWav.Picture = imgPlayWav1.Picture
        frmSearchForMedia.Show
    End If
End If
If Err.Number <> 0 Then SetError "imgPlayWav_MouseUp()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgRip_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 Then
    imgRip.Picture = imgRip2.Picture
End If
If Err.Number <> 0 Then SetError "imgRip_MouseDown()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgRip_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
PictureBoxMouseMove oRip, Button, imgRip, imgRip1, imgRip2, X, Y, imgRip3, True
If Err.Number <> 0 Then SetError "imgRip_MouseMove()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgRip_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
Dim i As Integer, m As Integer
If Button = 1 And imgRip.Picture = imgRip2.Picture Then
    If lRipperSettings.eAspiEnabled = False Then
        UpdateASPI
        imgRip.Picture = imgRip1.Picture
    Else
        PlayWav App.Path & "\media\click2.wav", SND_ASYNC
        ConvertCaption oIdle
        LoadTrackGet True
        imgRip.Picture = imgRip1.Picture
    End If
End If
If Err.Number <> 0 Then SetError "imgRip_MouseUp()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgSkinEdit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 Then
    imgSkinEdit.Picture = imgSkinEdit2.Picture
End If
If Err.Number <> 0 Then SetError "imgSkinEdit_MouseDown()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgSkinEdit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
PictureBoxMouseMove oSkinEdit, Button, imgSkinEdit, imgSkinEdit1, imgSkinEdit2, X, Y, imgSkinEdit3, True
If Err.Number <> 0 Then SetError "imgSkinEdit_MouseMove()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgSkinEdit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 And imgSkinEdit.Picture = imgSkinEdit2.Picture Then
    imgSkinEdit.Picture = imgSkinEdit1.Picture
    frmSkinBrowser.Show
End If
If Err.Number <> 0 Then SetError "imgSkinEdit_MouseUp()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgStop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 Then
    imgStop.Picture = imgStop2.Picture
End If
If Err.Number <> 0 Then SetError "imgStop_MouseDown()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgStop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
PictureBoxMouseMove oStopButton, Button, imgStop, imgStop1, imgStop2, X, Y, imgStop3, True
If Err.Number <> 0 Then SetError "imgStop_MouseMove()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgStop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 And imgStop.Picture = imgStop2.Picture Then
    imgStop.Picture = imgStop1.Picture
    PlayWav App.Path & "\media\click.wav", SND_ASYNC
    If lPlayer.pStatus = sPlaying Or sPaused Then
        PlayerDone
        StopMp3
        lPlayer.pStatus = sNotPlaying
    ElseIf lEffectsPresets.eStatus = eOpen Or lEffectsPresets.eStatus = eAddingEffect Or lEffectsPresets.eStatus = ePaused Or lEffectsPresets.eStatus = eOpening Then
        CloseEffects
    End If
End If
If Err.Number <> 0 Then SetError "imgStop_MouseUp()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgStopEncoding_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 Then
    imgStopEncoding.Picture = imgStopEncoding2.Picture
End If
If Err.Number <> 0 Then SetError "imgStopEncoding_MouseDown()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgStopEncoding_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
PictureBoxMouseMove oStopEncoding, Button, imgStopEncoding, imgStopEncoding1, imgStopEncoding2, X, Y, imgStopEncoding3, True
If Err.Number <> 0 Then SetError "imgStopEncoding_mouseMove()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub imgStopEncoding_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
Dim msg As Boolean
If Button = 1 And imgStopEncoding.Picture = imgStopEncoding2.Picture Then
    imgStopEncoding.Picture = imgStopEncoding1.Picture
    If lPlayer.pStatus = sPlaying Then
        StopPlayQue
        'StopAllEvents
        DoEvents
        StopMp3
        PlayerDone
        DoEvents
    Else
        Encoder.sTop
    End If
End If
If Err.Number <> 0 Then SetError "imgStopEncoding_MouseUp()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub lblInfo_DblClick()
On Local Error Resume Next
PlayWav App.Path & "\media\done.wav", SND_ASYNC
If lEvents.eTimeType = 1 Then
    lEvents.eTimeType = 2
ElseIf lEvents.eTimeType = 2 Then
    lEvents.eTimeType = 3
ElseIf lEvents.eTimeType = 3 Then
    lEvents.eTimeType = 1
End If
If Err.Number <> 0 Then SetError "lblInfo_DblClick()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub lblInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
Dim f As Long, c As Long
If Button = 1 Then
    If lPlayer.pStatus = sPlaying Then
        f = (SimpleMP3.FrameCount * X) / lblInfo.Width / 14.5
        If f < SimpleMP3.FrameCount Then SimpleMP3.Seek (f)
    Else
        FormDrag Me
    End If
ElseIf Button = 2 Then
    PlayWav App.Path & "\media\click.wav", SND_ASYNC
    PopupMenu mnuHidden
End If
If Err.Number <> 0 Then SetError "lblInfo_MouseDown()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub lblInfo_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
DragDrop Data
If Err.Number <> 0 Then SetError "lblInfo_OleDragDrop()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub lblMp3File_Change()
On Local Error Resume Next
lblMp3File.ToolTipText = lblMp3File.Caption
If Err.Number <> 0 Then SetError "lblInfo_Change()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub lblMp3File_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
Dim msg As String
If Button = 1 Then
    FormDrag Me
ElseIf Button = 2 Then
    PlayWav App.Path & "\media\click.wav", SND_ASYNC
    If Len(lblMp3File.Caption) <> 0 Then
        PopupMenu mnuMpegMenu
    Else
        PopupMenu mnuHidden
    End If
End If
If Err.Number <> 0 Then SetError "lblMp3File_MouseDown()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub lblMp3File_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
DragDrop Data
If Err.Number <> 0 Then SetError "lblMp3File_OleDragDrop()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub lblWavFile_Change()
On Local Error Resume Next
lblWavFile.ToolTipText = lblWavFile.Caption
If Err.Number <> 0 Then SetError "lblMp3File_Change()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub lblWavFile_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 Then
    FormDrag frmMain
ElseIf Button = 2 Then
    PlayWav App.Path & "\media\click.wav", SND_ASYNC
    If Len(lblWavFile.Caption) <> 0 Then
        If lEffectsPresets.eStatus = eClosed Then
            PopupMenu mnuHidden
            Exit Sub
        Else
            'If DoesFileExist(lPlayer.pLabels.lWavPath & lPlayer.pLabels.lWavFile) = True Then
            PopupMenu mnuWaveMenu
            'End If
        End If
    Else
        PopupMenu mnuHidden
    End If
End If
If Err.Number <> 0 Then SetError "lblWavFile_MouseDown()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub lblWavFile_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
DragDrop Data
If Err.Number <> 0 Then SetError "lblWavFile_OleDragDrop()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnu1xFast_Click()
On Local Error Resume Next
mnu1xFast.Checked = True
mnu2xFast.Checked = False
mnuNormal.Checked = False
mnu1xSlow.Checked = False
mnu2xSlow.Checked = False
SimpleMP3.SetSpeed 75
If Err.Number <> 0 Then SetError "mnu1xFast_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnu1xSlow_Click()
On Local Error Resume Next
mnu1xFast.Checked = False
mnu2xFast.Checked = False
mnuNormal.Checked = False
mnu1xSlow.Checked = True
mnu2xSlow.Checked = False
SimpleMP3.SetSpeed 125
If Err.Number <> 0 Then SetError "mnuNormal_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnu2xFast_Click()
On Local Error Resume Next
mnu1xFast.Checked = False
mnu2xFast.Checked = True
mnuNormal.Checked = False
mnu1xSlow.Checked = False
mnu2xSlow.Checked = False
SimpleMP3.SetSpeed 50
If Err.Number <> 0 Then SetError "mnuNormal_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnu2xSlow_Click()
On Local Error Resume Next
mnu1xFast.Checked = False
mnu2xFast.Checked = False
mnuNormal.Checked = False
mnu1xSlow.Checked = False
mnu2xSlow.Checked = True
SimpleMP3.SetSpeed 150
If Err.Number <> 0 Then SetError "mnu2xSlow_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuAdjustVolume_Click()
On Local Error Resume Next
frmVolume.Show
If Err.Number <> 0 Then SetError "mnuAdjustVolume_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuAlbum_Click()
On Local Error Resume Next
Surf "http://www.rollingstone.com/search/default.asp?st=music&ctgy=album&sf=" & lTracks.tTitle
If Err.Number <> 0 Then SetError "mnuAlbum_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuAmplitude_Click()
On Local Error Resume Next
AddAmplitude
If Err.Number <> 0 Then SetError "mnuAmplitude()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuArtist_Click()
On Local Error Resume Next
Surf "http://www.rollingstone.com/search/default.asp?st=music&ctgy=artists&sf=" & lTracks.tArtist
If Err.Number <> 0 Then SetError "mnuArtist_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuAudica_Click()
On Local Error Resume Next
If lEvents.eSettings.iPlayMp3sInNexENCODE = True Then
    lblInfo.Caption = "I am the MP3 Player"
Else
    GoMp3Player
End If
If Err.Number <> 0 Then SetError "mnuAudica_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuAudioGalWeb_Click()
On Local Error Resume Next
Surf "http://www.audiogalaxy.com"
If Err.Number <> 0 Then SetError "mnuAudioGalWeb_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuBackward_Click()
On Local Error Resume Next
GoBackward
If Err.Number <> 0 Then SetError "mnuBackWard_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuBackwardMpeg_Click()
On Local Error Resume Next
GoBackward
If Err.Number <> 0 Then SetError "mnuBackwardMpeg_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuBatch_Click()
On Local Error Resume Next
frmAddEvent.Show
If Err.Number <> 0 Then SetError "mnuBatch_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuCDDBOptions_Click()
On Local Error Resume Next
frmSettings.ResetSettingsFrames eFreeDB, True
If Err.Number <> 0 Then SetError "mnuEncoder_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuCFilter_Click()
On Local Error Resume Next
InitEffects
ns4Effects.CFilter 1
If Err.Number <> 0 Then SetError "mnuCFilter()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuCheckforUpdates_Click()
On Local Error Resume Next
frmLatestVersionCheck.Show
If Err.Number <> 0 Then SetError "mnuCheckForUpdates_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuChorus_Click()
On Local Error Resume Next
AddChorus 35, 25, 2, 1, 75, 75, 1, -1, 0
If Err.Number <> 0 Then SetError "mnuChorus()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuClose_Click()
On Local Error Resume Next
StopMp3
lblMp3File.Caption = ""
If Err.Number <> 0 Then SetError "mnuClose_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuCloseWav_Click()
On Local Error Resume Next
Dim msg As String
If lEffectsPresets.eStatus = eOpen Then
    CloseEffects
ElseIf lEffectsPresets.eStatus = eAddingEffect Then
    lblInfo.Caption = "Busy adding effect ..."
    pause 0.2
    Exit Sub
ElseIf lEffectsPresets.eStatus = ePlaying Then
    CloseEffects
ElseIf lEffectsPresets.eStatus = eOpening Then
    Exit Sub
ElseIf lEffectsPresets.eStatus = eStopped Then
    CloseEffects
End If
lblWavFile.Caption = ""
If Err.Number <> 0 Then SetError "mnuCloseWav_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuCloseWavFile_Click()
On Local Error Resume Next
CloseEffects
If Err.Number <> 0 Then SetError "mnuCloseEffects()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuContinuous_Click()
On Local Error Resume Next
If lPlayer.pContinuous = True Then
    lPlayer.pContinuous = False
    mnuContinuous.Checked = False
    WriteINI lIniFiles.iPlayers, "Settings", "Continuous", False
Else
    lPlayer.pContinuous = True
    mnuContinuous.Checked = True
    WriteINI lIniFiles.iPlayers, "Settings", "Continuous", True
End If
If Err.Number <> 0 Then SetError "mnuContinuous_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuConvertToMp3_Click()
On Local Error Resume Next
Dim msg As String, msg2 As String, lMp3 As String
msg = lPlayer.pLabels.lWavPath & lPlayer.pLabels.lWavFile
If DoesFileExist(msg) = True Then
    msg2 = msg
    msg2 = GetFileTitle(msg2)
    msg = Left(msg, Len(msg) - Len(msg2))
    lMp3 = Left(msg2, Len(msg2) - 3) & "mp3"
    AddEvent Encode, msg, msg2, msg, lMp3, 0, ""
End If
If Err.Number <> 0 Then SetError "mnuConvertToMp3_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuConvertToWav_Click()
On Local Error Resume Next
Dim msg As String, msg2 As String, lWav As String
msg = lPlayer.pLabels.lMp3Path & lPlayer.pLabels.lMp3File
If DoesFileExist(msg) = True Then
    msg2 = msg
    msg2 = GetFileTitle(msg2)
    msg = Left(msg, Len(msg) - Len(msg2))
    lWav = Left(msg2, Len(msg2) - 3) & "wav"
    AddEvent Decode, msg, msg2, msg, lWav, 0, ""
End If
If Err.Number <> 0 Then SetError "mnuConvertToWav_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuDelete_Click()
On Local Error Resume Next
Dim msg As String, mbox As VbMsgBoxResult

msg = lPlayer.pLabels.lWavPath & lPlayer.pLabels.lWavFile
If DoesFileExist(msg) = True Then
    If lEvents.eSettings.iOverwritePrompts = True Then
        mbox = MsgBox("Are you sure you want to delete the following file?" & vbCrLf & vbCrLf & msg, vbYesNo + vbQuestion, "Confirm Delete")
        If mbox = vbYes Then
            Kill msg
            lblWavFile.Caption = ""
            lPlayer.pLabels.lWavPath = ""
            lPlayer.pLabels.lWavFile = ""
        ElseIf mbox = vbNo Then
            Exit Sub
        End If
    Else
        Kill msg
        lblWavFile.Caption = ""
        lPlayer.pLabels.lWavPath = ""
        lPlayer.pLabels.lWavFile = ""
    End If
End If
End Sub

Private Sub mnuDeleteSettings_Click()
On Local Error Resume Next
Dim b As VbMsgBoxResult

If lEvents.eSettings.iOverwritePrompts = True Then
    b = MsgBox("Deleteing your settings will require restarting NexENCODE. Are you sure you wish to do this?", vbQuestion + vbYesNo)
    If b = vbYes Then
        Kill lIniFiles.iErrors
        Kill lIniFiles.iPlayers
        Kill lIniFiles.iPlaylists
        Kill lIniFiles.iSettings
        Kill lIniFiles.iWindowPos
        Shell App.Path & "\NexENCODE.exe", vbNormalFocus
        End
    ElseIf b = vbNo Then
        Exit Sub
    End If
Else
    Kill lIniFiles.iErrors
    Kill lIniFiles.iPlayers
    Kill lIniFiles.iPlaylists
    Kill lIniFiles.iSettings
    Kill lIniFiles.iWindowPos
    Shell App.Path & "\NexENCODE.exe", vbNormalFocus
    End
End If

If Err.Number <> 0 Then SetError "mnuDeleteSettings_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuDistortion_Click()
On Local Error Resume Next
AddDistortion 1005, 560, 3, 0, 0
If Err.Number <> 0 Then SetError "mnuAddDistortion()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuDownloadTracks_Click()
On Local Error Resume Next
GetSimpleTracks
If Err.Number <> 0 Then SetError "mnuDownloadTracks_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuEcho_Click()
On Local Error Resume Next
InitEffects
ns4Effects.Echo 900, 90
If Err.Number <> 0 Then SetError "mnuEcho_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuEditTracks_Click()
On Local Error Resume Next
frmEditTracks.Show
If Err.Number <> 0 Then SetError "mnuEditTracks_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuEffectsEditor_Click()
On Local Error Resume Next
frmEffects.Show
If Err.Number <> 0 Then SetError "mnuEffectsEditor_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuEmailDeveloper_Click()
On Local Error Resume Next
Surf "mailto:brendlefly3000@hotmail.com"
If Err.Number <> 0 Then SetError "mnuEmailDeveloper_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuEmailKnightFal_Click()
On Local Error Resume Next
Surf "mailto:knightfal@team-nexgen.com"
If Err.Number <> 0 Then SetError "mnuEmailKnightFal_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuEncoder_Click()
On Local Error Resume Next
frmSettings.ResetSettingsFrames eEncoder, True
If Err.Number <> 0 Then SetError "mnuEncoder_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuEntireDiscToMp3_Click()
On Local Error Resume Next
Dim i As Integer, msg As String
LoadTrackGet True
DoEvents
With frmTrackGet
    For i = 0 To .lstAvailableTracks.ListCount
        msg = .lstAvailableTracks.List(i)
        If Len(msg) <> 0 Then
            .lstQue.AddItem msg
            DoEvents
        End If
    Next i
    .lstAvailableTracks.Clear
End With
If Err.Number <> 0 Then SetError "mnuEntireDisctoMp3_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuEntireDisctoWav_Click()
On Local Error Resume Next
Dim i As Integer, msg As String
LoadTrackGet False
DoEvents
With frmTrackGet
    For i = 0 To .lstAvailableTracks.ListCount
        msg = .lstAvailableTracks.List(i)
        If Len(msg) <> 0 Then
            .lstQue.AddItem msg
            DoEvents
        End If
    Next i
    .lstAvailableTracks.Clear
End With
If Err.Number <> 0 Then SetError "mnuEntireDisctoWav_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuFadeIN_Click()
On Local Error Resume Next
InitEffects
ns4Effects.FadeIn 50
If Err.Number <> 0 Then SetError "mnuFadeIN_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuFadeOut_Click()
On Local Error Resume Next
InitEffects
ns4Effects.FadeOut 50
If Err.Number <> 0 Then SetError "mnuFadeOut_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuForward_Click()
On Local Error Resume Next
GoForward
If Err.Number <> 0 Then SetError "mnuForward_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuForwardMpeg_Click()
On Local Error Resume Next
GoForward
If Err.Number <> 0 Then SetError "mnuForwardMpeg_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuGeneralSettings_Click()
On Local Error Resume Next
frmSettings.ResetSettingsFrames eGeneral, True
If Err.Number <> 0 Then SetError "mnuGeneralSettings_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuHelp_Click()
On Local Error Resume Next
MsgBox "The Encoder menu is for converting wave audio files to mpeg layer 3 files. 'Encode (wav to mp3)' compresses wave audio into mp3, 'Decode' is for converting mp3 files back to wave audio files, "
End Sub

Private Sub mnuHelp2_Click()
On Local Error Resume Next
MsgBox "The eject menu is usefull if you want to set the file NexENCODE currently has open. 'Open Mpeg Layer 3' sets the mp3 file label (*.mp3) to any file you specify, 'Open Wave Audio' sets the wav (*.wav) file label to any file you specify, 'Open Playlists' sets the mp3 file label to any (*.m3u) file you specify. Once you have set this file, you can add effects, encode, decode, play, or delete the file currently set.", vbInformation, "Help"
End Sub

Private Sub mnuHelp3_Click()
On Local Error Resume Next
MsgBox "The Settings menu is useful if you wish to edit the way NexENCODE functions. 'Encoder' shows the encoder settings window which allows you to choose the encoder settings, what happens after you encode, and your current encoder profile. 'Ripper' shows the ripper settings interface, which allows you to select the CD drive, copy mode, ripper settings, and the directory from which files are stored. 'Player' shows the player settings interface, which stores a list of mp3 and cd players. 'General' shows the general settings interface. 'FreeDB' shows the freedb settings interface which allows you to connect to free db websites and download track names. 'ASPI' is for your aspi drivers and keeping them up to date. 'Register' (for unregistered users) shows how to register NexENCODE", vbInformation, "Help"
End Sub

Private Sub mnuHelp4_Click()
On Local Error Resume Next
MsgBox "The skins menu edits NexENCODE's interface and visual apeal. NexENCODE by default uses the inex skin, however you can set the skin to any .ns4 file. 'Load' shows the open dialog and lets you select a .ns4 file as the main interface. 'Editor' shows the skin editor which can create skins for NexENCODE. 'Browser' shows the skin browser, which lets you select the skin you wish to be using. 'More skins' launches you to the team nexgen website where the skins can be found", vbInformation, "Help"
End Sub

Private Sub mnuHelp5_Click()
On Local Error Resume Next
MsgBox "Ripping is copying tracks from your cd drive to your hard drive in wave audio format. Click 'Rip (cda to mp3)' to rip cd audio to mpeg layer 3 with the track get interface, Click 'Rip (cda to mp3)' to rip cd audio to wave audio with the track get interface, Click 'Encode Disc' to copy your entire cd to mpeg layer 3 audio", vbInformation, "Help"
End Sub

Private Sub mnuHelp6_Click()
MsgBox "This menu is for playing mpeg layer 3 audio", vbInformation, "Help"
End Sub

Private Sub mnuHelp7_Click()

End Sub

Private Sub mnuHelp8_Click()

End Sub

Private Sub mnuHelp9_Click()
On Local Error Resume Next
MsgBox "The effects editor adds different effects to your wave audio files, click 'Editor' to show the effects editor", vbInformation, "Help"
End Sub

Private Sub mnuHide_Click()
On Local Error Resume Next
Me.Visible = False
Me.WindowState = vbMinimized
If lEvents.ePlaylistVisible = True Then Unload frmPlaylist
If Err.Number <> 0 Then SetError "mnuHide_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuInvert_Click()
On Local Error Resume Next
AddInvert
If Err.Number <> 0 Then SetError "mnuAddInvert_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuKaza_Click()
On Local Error Resume Next
Surf "http://www.kaza.com/"
If Err.Number <> 0 Then SetError "mnuKaza_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuLoadSkin_Click()
On Local Error Resume Next
Dim i As Integer
i = OpenSkin(OpenDialog(frmMain, "Skin Files (*.ns4)|*.ns4|", "Load Skin", App.Path & "\Skins"), False)
If i <> 0 Then ApplySkin frmMain, i
If Err.Number <> 0 Then SetError "mnuLoadSkin_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuMergeMp3_Click()
On Local Error Resume Next
frmFileMerger.Show
If Err.Number <> 0 Then SetError "mnuMergeMP3_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuMoreInfo_Click()
On Local Error Resume Next
Dim i As Integer, msg As String
If Len(lPlayer.pLabels.lMp3Path) <> 0 Then
    If Right(lPlayer.pLabels.lMp3Path, 1) <> 0 Then lPlayer.pLabels.lMp3Path = lPlayer.pLabels.lMp3Path & "\"
    msg = lPlayer.pLabels.lMp3Path & lPlayer.pLabels.lMp3File
    If DoesFileExist(msg) = False Then
        If Len(msg) <> 0 Then
            i = FindMediaIndex(msg)
            If i <> 0 Then
                frmMP3Info.Show
                PromptGetTag Playlist.pFiles(i).fPath & Playlist.pFiles(i).fFile
                frmMP3Info.RefreshTagInfo
            End If
        End If
    Else
        frmMP3Info.Show
        PromptGetTag msg
        frmMP3Info.RefreshTagInfo
    End If
Else
    i = FindMediaIndex(lblMp3File.Caption)
    If i <> 0 Then
        frmMP3Info.Show
        PromptGetTag Playlist.pFiles(i).fPath & Playlist.pFiles(i).fFile
        frmMP3Info.RefreshTagInfo
    End If
End If
If Err.Number <> 0 Then SetError "mnuMoreInfo_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuMoreSkinsOnline_Click()
On Local Error Resume Next
Surf "http://www.team-nexgen.com/downloads/ns4skins/"
If Err.Number <> 0 Then SetError "mnuMoreSkinsOnline_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuMp3ToWav2_Click()
On Local Error Resume Next
frmEncode.Show
frmEncode.cboFormat.ListIndex = 1
If Err.Number <> 0 Then SetError "mnuMp3ToWav2_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuMp3Website_Click()
On Local Error Resume Next
Surf "http://www.mp3.com"
If Err.Number <> 0 Then SetError "mnuMp3Website_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuMusicSearch_Click()
On Local Error Resume Next
frmSearch.Show
If Err.Number <> 0 Then SetError "mnuMusicSearch_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuNexENCODE_Click()
On Local Error Resume Next
If mnuCancel.Visible = True Then
    frmMain.Visible = True
    frmMain.WindowState = vbNormal
    frmMain.mnuNexENCODE.Checked = True
Else
    frmAbout_NS.Show
End If
If Err.Number <> 0 Then SetError "mnuNexENCODE()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuNexMedia_Click()
On Local Error Resume Next
GoCDPlayer
If Err.Number <> 0 Then SetError "mnuNexMedia_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuNexSkin_Click()
On Local Error Resume Next
If mnuNexSkin.Checked = True Then
    frmSkinEditor.Hide
Else
    frmSkinEditor.Show
End If
If Err.Number <> 0 Then SetError "mnuNexSkin_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuNormal_Click()
On Local Error Resume Next
mnu1xFast.Checked = False
mnu2xFast.Checked = False
mnuNormal.Checked = True
mnu1xSlow.Checked = False
mnu2xSlow.Checked = False
SimpleMP3.SetSpeed 100
If Err.Number <> 0 Then SetError "mnuNormal_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuOnTop_Click()
On Local Error Resume Next
If mnuOnTop.Checked = False Then
    AlwaysOnTop frmMain, True
    WriteINI lIniFiles.iSettings, "Settings", "AlwaysOnTop", "True"
    mnuOnTop.Checked = True
Else
    AlwaysOnTop frmMain, False
    WriteINI lIniFiles.iSettings, "Settings", "AlwaysOnTop", "False"
    mnuOnTop.Checked = False
End If
If Err.Number <> 0 Then SetError "mnuOnTop_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuOpen_Click()
On Local Error Resume Next
PromptToPlay
If Err.Number <> 0 Then SetError "mnuOpen_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuOpenCurrentFile_Click()
On Local Error Resume Next
OpenEffects lPlayer.pLabels.lWavPath & lPlayer.pLabels.lWavFile
If Err.Number <> 0 Then SetError "mnuOpenCurrentFile()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuOpenInSoundRec_Click()
On Local Error Resume Next
Shell "sndrec32.exe " & lPlayer.pLabels.lWavPath & lPlayer.pLabels.lWavFile, vbNormalFocus
If Err.Number <> 0 Then SetError "mnuOpenInSoundRec_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuOpenMp3_Click()
On Local Error Resume Next
Dim msg As String, msg2 As String
msg = OpenDialog(Me, "Mpeg Layer-3 (*.mp3)|*.mp3", "Open Mpeg Layer-3", CurDir)
If Len(msg) <> 0 Then
    msg2 = msg
    msg2 = GetFileTitle(msg2)
    msg = Left(msg, Len(msg) - Len(msg2))
    SetFileLabels "", "", msg2, msg
End If
If Err.Number <> 0 Then SetError "mnuOpenMp3_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuOpenMpeg_Click()
On Local Error Resume Next
Dim msg As String, msg2 As String
msg = OpenDialog(Me, "Mpeg Layer-3 (*.mp3)|*.mp3", "Open Mpeg Layer-3", CurDir)
If Len(msg) <> 0 Then
    msg2 = msg
    msg2 = GetFileTitle(msg2)
    msg = Left(msg, Len(msg) - Len(msg2))
    SetFileLabels "", "", msg2, msg
End If

If Err.Number <> 0 Then SetError "mnuOpenMp3_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuOpenPlaylist_Click()
On Local Error Resume Next
Dim msg As String, msg2 As String
msg = OpenDialog(Me, "Playlist Files (*.m3u)|*.m3u", "Open Playlist File", CurDir)
If Len(msg) <> 0 Then
    msg2 = msg
    msg2 = GetFileTitle(msg2)
    msg = Left(msg, Len(msg) - Len(msg2))
    SetFileLabels "", "", msg2, msg
End If

If Err.Number <> 0 Then SetError "mnuOpenPlaylist_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuOpenWave_Click()
On Local Error Resume Next
Dim msg As String, msg2 As String
If lEffectsPresets.eStatus = eOpen Then
    CloseEffects
    DoEvents
End If
msg = OpenDialog(Me, "Wave Audio (*.wav)|*.wav", "Open Wave Audio", CurDir)
If Len(msg) <> 0 Then
    msg2 = msg
    msg2 = GetFileTitle(msg2)
    msg = Left(msg, Len(msg) - Len(msg2))
    SetFileLabels msg, msg2
    OpenEffects msg & msg2
End If
If Err.Number <> 0 Then SetError "mnuOpenWave_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuPause_Click()
On Local Error Resume Next
PauseMp3
If Err.Number <> 0 Then SetError "mnuPause_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuPauseMpeg_Click()
On Local Error Resume Next
PauseMp3
End Sub

Private Sub mnuPlay_Click()
On Local Error Resume Next
frmMain.SimpleMP3.Play
If Err.Number <> 0 Then SetError "mnuPlay_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuPlayEffectWav_Click()
On Local Error Resume Next
PlayEffect
If Err.Number <> 0 Then SetError "mnuPlayEffectWav_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuPlayer__Click()
On Local Error Resume Next
frmSettings.ResetSettingsFrames ePlayers, True
If Err.Number <> 0 Then SetError "mnuPlayerSettings_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuPlaylist_Click()
On Local Error Resume Next
If mnuPlaylist.Checked = True Then
    Unload frmPlaylist
Else
    frmPlaylist.Show
End If
If Err.Number <> 0 Then SetError "mnuPlaylist_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuPlayMpeg_Click()
On Local Error Resume Next
Dim msg As String, msg2 As String, i As Integer
msg = lPlayer.pLabels.lMp3Path & lPlayer.pLabels.lMp3File
msg2 = lPlayer.pLabels.lMp3File
Select Case LCase(Right(msg2, 3))
Case "mp3"
    i = FindMediaIndex(msg2)
    If i <> 0 Then
        lblMp3File.Caption = ""
        PlayMp3 i
    Else
        i = AddToPlaylist(msg, 1)
        If i <> 0 Then
            lblMp3File.Caption = ""
            PlayMp3 i
        End If
    End If
Case "m3u"
    Dim lPath As String, s As Integer
    lPath = Left(msg, Len(msg) - Len(msg2))
    s = FindPlaylistIndexByFile(msg2): DoEvents
    If s <> 0 Then
        PlayPlaylist "", "", s
    Else
        PlayPlaylist lPath, msg2
    End If
End Select
If Err.Number <> 0 Then SetError "mnuPlayMpeg_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuPower_Click()
On Local Error Resume Next
UnloadMain
If Err.Number <> 0 Then SetError "mnuPower_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuRandom_Click()
On Local Error Resume Next
LoadRandomMP3
If Err.Number <> 0 Then SetError "mnuRandom_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuRandom2_Click()
On Local Error Resume Next
LoadRandomMP3
If Err.Number <> 0 Then SetError "mnuRandom2_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuReadMe_Click()
On Local Error Resume Next
ShowText "Documentation", App.Path & "\documentation\nexencode.txt"
If Err.Number <> 0 Then SetError "mnuReadMe_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuRegistered_Click()
On Local Error Resume Next
frmSettings.ResetSettingsFrames eCDDB2, True
If Err.Number <> 0 Then SetError "mnuRegistered_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuReverb_Click()
On Local Error Resume Next
AddReverb 900, 90
If Err.Number <> 0 Then SetError "mnuReverb()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuRipCDATOMP3_Click()
On Local Error Resume Next
LoadTrackGet True
If Err.Number <> 0 Then SetError "mnuRipCDATOMp3()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuRipCDAToWav_Click()
On Local Error Resume Next
LoadTrackGet False
If Err.Number <> 0 Then SetError "mnuRipCDAToWav()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuRipperSettings_Click()
On Local Error Resume Next
frmSettings.ResetSettingsFrames eRipper, True
If Err.Number <> 0 Then SetError "mnuRipperSettings_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuRollingStonesWebsite_Click()
On Local Error Resume Next
Surf "http://www.rollingstones.com"
If Err.Number <> 0 Then SetError "mnuRollingStonesWebsite_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuSaveWavAs_Click()
On Local Error Resume Next
Dim msg As String
msg = SaveDialog(Me, "Wave Audio (*.wav)|*.wav", "Save wave audio as ...", CurDir)
If Len(msg) <> 0 Then
    If Right(LCase(msg), 4) <> ".wav" Then msg = msg & ".wav"
    InitEffects
    DoEvents
    ns4Effects.InputFileSave msg
End If
If Err.Number <> 0 Then SetError "mnuSaveWavAs_click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuSearchHardDrive_Click()
On Local Error Resume Next
frmSearchForMedia.Show
If Err.Number <> 0 Then SetError "mnuSearchHardDrive_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuSearchWithinPlaylists_Click()
On Local Error Resume Next
frmSearchPlaylists.Show
If Err.Number <> 0 Then SetError "mnuSearchWithinPlaylists_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuSettingsASPI_Click()
On Local Error Resume Next
frmSettings.ResetSettingsFrames eAspi, True
If Err.Number <> 0 Then SetError "mnuSettingsASPI_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuSetupWizard_Click()
On Local Error Resume Next
frmSetupWizard.Show
If Err.Number <> 0 Then SetError "mnuSetupWizard_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuShifting_Click()
On Local Error Resume Next
AddShifting 1, 2048
If Err.Number <> 0 Then SetError "mnuShifting_click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnushowTagEditor_Click()
On Local Error Resume Next
LoadId3Editor
If Err.Number <> 0 Then SetError "mnuShowTagEditor_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuShowWavEffectsEditor_Click()
On Local Error Resume Next
Dim msg As String, msg2 As String
If lEffectsPresets.eStatus = eOpen Then
    frmEffects.Show
    Exit Sub
ElseIf lEffectsPresets.eStatus = eClosed And Len(lPlayer.pLabels.lWavFile) <> 0 Then
    OpenEffects lPlayer.pLabels.lWavPath & lPlayer.pLabels.lWavFile: DoEvents
    pause 1
    frmEffects.Show
    Exit Sub
End If
msg = OpenDialog(Me, "Wave Audio (*.wav)|*.wav", "Select Wave Audio", CurDir)
If Len(msg) <> 0 Then
    If Right(LCase(msg), 4) <> ".wav" Then
        msg = msg & ".wav"
    End If
    msg2 = msg
    msg2 = GetFileTitle(msg2)
    OpenEffects msg: DoEvents
    SetFileLabels Left(msg, Len(msg) - Len(msg2)), msg2
    pause 1
    frmEffects.Show
End If
If Err.Number <> 0 Then SetError "mnuShowWavEffectsEditor_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuSingleFromDisk_Click()
On Local Error Resume Next
frmEncode.Show
If Err.Number <> 0 Then SetError "mnuSingleFromDisc()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuSkinBrowser_Click()
On Local Error Resume Next
frmSkinBrowser.Show
If Err.Number <> 0 Then SetError "mnuSkinBrowser()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuSkinName_Click(Index As Integer)
On Local Error Resume Next
ApplySkin frmMain, Index - 1
If Err.Number <> 0 Then SetError "mnuSkinName_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuStopEffectWav_Click()
On Local Error Resume Next
ns4Effects.sTop
If Err.Number <> 0 Then SetError "mnuStopEffectWav()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuStopMp3_Click()
On Local Error Resume Next
StopMp3
If Err.Number <> 0 Then SetError "mnuStopMp3_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuStopMpeg_Click()
On Local Error Resume Next
StopMp3
If Err.Number <> 0 Then SetError "mnuStopMpeg_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuSupport_Click()
On Local Error Resume Next
ShowText "Documentation", App.Path & "\documentation\support.txt"
If Err.Number <> 0 Then SetError "mnuSupport_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuTeamNexgen_Click()
On Local Error Resume Next
Surf "http://www.team-nexgen.com"
If Err.Number <> 0 Then SetError "mnuTeamNexgen_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuUpdateASPI_Click()
On Local Error Resume Next
Shell App.Path & "\programs\aspiupd.exe", vbNormalFocus
End
End Sub

Private Sub decoder_ActFrame(ByVal ActFrame As Long)
On Local Error Resume Next
Dim i As Integer
If imgPercent.Visible = False Then
    imgPercent.Width = 1
    imgPercent.Visible = True
End If
If lEvents.eEncoderBusy = False Then lEvents.eEncoderBusy = True
i = (ActFrame * 100 / Decoder.FrameCount) \ 1
If i < 101 Then
    imgPercent.Width = i * 1.17
    'imgPercent.Width = i
    EncodeCircleEffect Str(i)
    lEvents.ePercent = i
    lblInfo.Caption = "Decode: " & i & "%"
Else
    lblInfo.Caption = "Decoding ..."
End If
If Err.Number <> 0 Then SetError "decoder_Actframe", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub decoder_Failure(ByVal ErrorCode As Long, ByVal errStr As String)
On Local Error Resume Next
SetError "Decoder_Error()", "An error occured while decoding an mp3", errStr
End Sub

Private Sub decoder_ThreadEnded()
On Local Error Resume Next
PlayWav App.Path & "\media\done.wav", SND_ASYNC
lEvents.eRipperBusy = False
lEvents.eEncoderBusy = False
tmrScrollStatus.Enabled = False
imgPercent.Width = 0
ResetEncoderCircles
lblWavFile.Caption = ""
lblMp3File.Caption = ""
ToggleButtons oIdle
SetCaption Download
ConvertCaption oIdle
If lEvents.eSettings.iEnding = True Then Exit Sub
ProcessNextEvent
DoEvents
If Err.Number <> 0 Then SetError "Decoder_ThreadEnded()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub mnuWavAudio_Click()
On Local Error Resume Next
Dim msg As String, msg2 As String
msg = OpenDialog(Me, "Wave Audio (*.wav)|*.wav", "Open Wave Audio", CurDir)
If Len(msg) <> 0 Then
    msg2 = msg
    msg2 = GetFileTitle(msg2)
    msg = Left(msg, Len(msg) - Len(msg2))
    SetFileLabels msg, msg2
    OpenEffects msg & msg2
End If
If Err.Number <> 0 Then SetError "mnuWavAudio_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub ns4Effects_EndOfAction(ByVal ActionType As Integer)
On Local Error Resume Next
Dim f As Integer
EnableEffects
mnuStopEffectWav.Enabled = False
PlayWav App.Path & "\media\done.wav", SND_ASYNC
tmrScrollStatus.Enabled = False

lEvents.eRipperBusy = False
lEvents.eEncoderBusy = False
imgPercent.Width = 0
ResetRipperCircles
ToggleButtons oIdle
SetCaption Download
ConvertCaption oIdle
'lEffectsPresets.eStatus = eStopped
lEffectsPresets.eStatus = eOpen

If lEvents.eSettings.iEnding = True Then Exit Sub
ProcessNextEvent
DoEvents
End Sub

Private Sub ns4Effects_OnActionPosition(ByVal ActionPosition As Integer)
On Local Error Resume Next
Dim i As Integer
If imgPercent.Visible = False Then
    imgPercent.Width = 1
    imgPercent.Visible = True
End If
If lEvents.eRipperBusy = False Then lEvents.eRipperBusy = True
i = ActionPosition
If i < 101 Then
    imgPercent.Width = i * 1.17
    RipCircleEffect Str(i)
    lEvents.ePercent = i
    lblInfo.Caption = "Effects: " & i & "%"
Else
    lblInfo.Caption = "Effects ..."
End If
If Err.Number <> 0 Then SetError "ns4Effects_Actframe", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Ripper_ActPosition(ByVal Position As Long)
On Local Error Resume Next
Dim i As Long
With lEvents
    i = Position / Ripper.GetTrackLength(.eEvent(.eEventCount).eTrack) * 100 / 1.48
    RipCircleEffect Str(i)
    lEvents.ePercent = i
    imgPercent.Width = i * 1.17
    'imgPercent.Width = i
    lblInfo.Caption = "Rip " & i & "%"
    SetCaption Encode, i, frmMain.lblWavFile.Caption
End With
If Err.Number <> 0 Then SetError "NEXRIPPER_ActPosition()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Ripper_CopyStart()
On Local Error Resume Next
Dim i As Long, nag As Integer
With lEvents
    If .eRipperBusy = False Then .eRipperBusy = True
    If imgPercent.Visible = False Then
        imgPercent.Visible = True
        imgPercent.Width = 1
    End If
End With
If Err.Number <> 0 Then SetError "NEXRIPPER_CopyStart()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Ripper_CopyStop()
On Local Error Resume Next
Dim f As Integer
PlayWav App.Path & "\media\done.wav", SND_ASYNC
lEvents.eRipperBusy = False
lEvents.eEncoderBusy = False
imgPercent.Width = 0
ResetRipperCircles
lblWavFile.Caption = ""
ToggleButtons oIdle
SetCaption Download
AddFinishedEvent Rip, "Complete"
frmMain.Ripper.UnlockTray
ConvertCaption oIdle
If lEvents.eSettings.iEnding = True Then Exit Sub
ProcessNextEvent
DoEvents
If Err.Number <> 0 Then SetError "NEXRIPPER_CopyStop()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Ripper_Failure(ByVal ErrorCode As Long, ByVal ErrorString As String)
On Local Error Resume Next
If lEvents.eSettings.iEnding = False Then
    lEvents.ePercent = ErrorCode
    ConvertCaption oIdle
    SetError "RIPPER_Failure", "The ripper ran into an error", ErrorString
End If
If Err.Number <> 0 Then SetError "Ripper_Failure()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub SimpleMP3_ActFrame(ByVal ActFrame As Long)
On Local Error Resume Next
Dim i As Integer, h As Long

lblWavFile.Caption = Format((ActFrame * SimpleMP3.MsPerFrame) \ 1000, "00:00")
If imgPercent.Visible = False Then
    imgPercent.Width = 1
    imgPercent.Visible = True
End If
If lEvents.eEncoderBusy = False Then lEvents.eEncoderBusy = True
i = (ActFrame * 100 / SimpleMP3.FrameCount) \ 1
If i < 101 Then
    h = i
    imgPercent.Width = i * 1.17
    EncodeCircleEffect Str(i)
    lEvents.ePercent = i
    lblInfo.ToolTipText = "Play: " & i & "%"
    If Len(lblInfo.Caption) = 0 Then lblInfo.Caption = "Play: " & i & "%"
    SetCaption Play, h, frmMain.lblMp3File.Caption
Else
    lblInfo.Caption = "Playing ..."
End If
If Err.Number <> 0 Then SetError "SimpleMP3_Actframe", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub SimpleMP3_Failure(ByVal ErrorCode As Long, ByVal errStr As String)
On Local Error Resume Next
SetError "SimpleMP3", "A Mp3 Player Error occured", errStr
If Err.Number <> 0 Then SetError "SimpleMP3_Failure()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub SimpleMP3_ThreadEnded()
On Local Error Resume Next
Dim f As Integer
lEvents.eRipperBusy = False
lEvents.eEncoderBusy = False
imgPercent.Width = 0
ResetEncoderCircles
lblWavFile.Caption = ""
ToggleButtons oIdle
tmrScrollStatus.Enabled = False
lPlayer.pStatusString = ""
lblMp3File.Caption = ""
SetCaption Download
ConvertCaption oIdle
If lEvents.eSettings.iEnding = True Then Exit Sub
If lEvents.eEventCount < 201 Then
    If lEvents.eEventCount = 0 Or Len(lEvents.eEvent(lEvents.eEventCount).eInputFilename) = 0 Then
        If lPlayer.pPlayCanceled = False Then
            If lPlayer.pContinuous = True Then
                LoadRandomMP3
            End If
        Else
            lPlayer.pPlayCanceled = False
        End If
    Else
        ProcessNextEvent
    End If
Else
    lEvents.eEventCount = 200
End If
DoEvents
If Err.Number <> 0 Then SetError "Decoder_ThreadEnded()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub tmrCheckActive_Timer()
On Local Error Resume Next
If lEvents.eSettings.iCheckForActiveWindow = True Then IsActiveWindow
If Err.Number <> 0 Then SetError "tmrCheckActive_Timer()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub tmrPlayCommand_Timer()
On Local Error Resume Next
Dim lFile As String, lPath As String, i As Integer

If lEvents.eSettings.iLoading = True Then Exit Sub
lFile = Trim(lEvents.eSettings.iCommand)
If Len(lFile) <> 0 Then
    lPath = lFile
    lFile = GetFileTitle(lFile)
    lPath = Left(lPath, Len(lPath) - Len(lFile))
    If Right(LCase(lFile), 4) = ".mp3" Then
        AddPlayEvent lPath, lFile
    ElseIf Right(LCase(lFile), 4) = ".ns4" Then
        i = OpenSkin(lPath & lFile, False)
        If i <> 0 Then ApplySkin frmMain, i
    ElseIf Right(LCase(lFile), 4) = ".m3u" Then
        i = FindPlaylistIndexByFile(lFile)
        If i <> 0 Then
            PlayPlaylist "", "", i
        Else
            PlayPlaylist lPath, lFile
        End If
    End If
End If
tmrPlayCommand.Enabled = False
lEvents.eSettings.iCommand = ""
If Err.Number <> 0 Then SetError "tmrPlayCommand_Timer()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub tmrScrollStatus_Timer()
On Local Error Resume Next
Dim msg As String, lefty As String
If Len(lPlayer.pStatusString) <> 0 Then
    lefty = Left(lPlayer.pStatusString, 1)
    lPlayer.pStatusString = Right(lPlayer.pStatusString, Len(lPlayer.pStatusString) - 1) & lefty
    If Len(lPlayer.pStatusString) <> 0 Then
        lblInfo.Caption = Left(lPlayer.pStatusString, 22)
    End If
End If
If lEffectsPresets.eStatus = ePlaying Or lEffectsPresets.eStatus = eOpening Then
    Dim i As Integer
    For i = 0 To 9
        If shpRipper(i).BackColor <> shpRipperColor.BackColor Then shpRipper(i).BackColor = shpRipperColor.BackColor
    Next i
    If lSpindle = 9 Then
        shpRipper(0).BackColor = vbWhite
        lSpindle = 0
    Else
        lSpindle = lSpindle + 1
        shpRipper(lSpindle).BackColor = vbWhite
    End If
End If
If Err.Number <> 0 Then SetError "tmrScrollStatus_Timer()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub tmrShowEncoderCircles_Timer()
On Local Error Resume Next

If lEvents.eCircleNum = 10 Then
    lEvents.eCircleNum = 0
    tmrShowEncoderCircles.Enabled = False
End If
shpEncoder(lEvents.eCircleNum).Visible = True
lEvents.eCircleNum = lEvents.eCircleNum + 1
If Err.Number <> 0 Then SetError "tmrShowEncoderCircles_Timer()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub tmrShowRipperCircles_Timer()
On Local Error Resume Next

If lEvents.eCircleNum = 10 Then
    lEvents.eCircleNum = 0
    tmrShowRipperCircles.Enabled = False
End If
shpRipper(lEvents.eCircleNum).Visible = True
lEvents.eCircleNum = lEvents.eCircleNum + 1
If Err.Number <> 0 Then SetError "tmrShowRipperCircles_Timer()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub wskFreeDB_Close()
On Local Error Resume Next
wskFreeDB.Close
If lEvents.eSettings.iFreeDB.cShowDialog = True Then
    If frmWait.Height = 1460 Then Unload frmWait
End If
End Sub

Private Sub wskFreeDB_Connect()
On Local Error Resume Next

ShowWait "Downloading tracks", "Connecting"
wskFreeDB.SendData "cddb hello " & lEvents.eSettings.iFreeDB.cEmailAddress & " " & wskFreeDB.LocalHostName & " NexENCODE 4." & App.Minor & vbCrLf: DoEvents
SetWait "Downloading tracks", "Sending user info"

If Err.Number <> 0 Then SetError "wskFreeDB_Connect", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub wskFreeDB_DataArrival(ByVal bytesTotal As Long)
On Local Error Resume Next
Dim msg As String, msg2 As String, msg3 As String, lefty As String, msg4 As String, i As Integer, j As Integer, lGenre As String, msg5 As String

wskFreeDB.GetData msg, vbString

If lEvents.eSettings.iFreeDB.cShowDialog = True Then
    frmWait.txtIncoming.Text = frmWait.txtIncoming.Text & vbCrLf & msg
End If
If Left(msg, 3) = "200" Then
    If InStr(LCase(msg), "hello and welcome") Then
        wskFreeDB.SendData "cddb query " & Replace(ReturnFreeDBQueryString(lRipperSettings.eDiscID), "+", " ") & vbCrLf
        SetWait "Downloading tracks", "Sending disc info"
    Else
        Dim lCat As String, DiscID As String
        msg2 = Right(msg, Len(msg) - 4)
        lefty = Left(msg2, 1)
        lCat = lefty & ParseString(msg2, Left(msg2, 1), " ")
        msg2 = Right(msg2, Len(msg2) - Len(lCat) - 1)
        lefty = Left(msg2, 1)
        DiscID = lefty & ParseString(msg2, Left(msg2, 1), " ")
        wskFreeDB.SendData "cddb read " & lCat & " " & DiscID & vbCrLf
        lTracks.tArtist = Trim(ParseString(msg2, " ", "/"))
        msg4 = Right(msg2, Len(msg2) - Len(ParseString(msg2, Left(msg2, 1), "/")) - 3)
        lTracks.tTitle = Left(msg4, Len(msg4) - 2)
        frmMain.lblInfo.Caption = lTracks.tArtist & " / " & lTracks.tTitle
        SetWait "Downloading tracks", "Disc found " & lTracks.tArtist & " / " & lTracks.tTitle
        frmMain.mnuArtist.Caption = "Artist: " & lTracks.tArtist
        frmMain.mnuAlbum.Caption = "Album: " & lTracks.tTitle
        frmMain.mnuArtist.Enabled = True
        frmMain.mnuAlbum.Enabled = True
    End If
ElseIf Left(msg, 3) = "210" Then
    SetWait "Downloading tracks", "Recieving Track names"
    j = lTracks.tCount
    lGenre = ParseString(msg, "210 ", "CD ")
    lGenre = Right(lGenre, Len(lGenre) - 3)
    lefty = UCase(Left(lGenre, 1))
    lGenre = lefty & ParseString(lGenre, Left(lGenre, 1), " ")
    lTracks.tGenre = lGenre
    For i = 0 To j
        msg5 = "EXTD="
        msg3 = "TTITLE" & i & "="
        msg4 = "TTITLE" & i + 1 & "="
        If InStr(msg, msg3) And InStr(msg, msg3) Then
            msg2 = ParseString(msg, msg3, msg4)
            DoEvents
            If Len(msg2) <> 0 Then
                msg2 = Right(msg2, Len(msg2) - Len(msg3) + 1)
                msg2 = Left(msg2, Len(msg2) - 2)
                lTracks.tTrack(i + 1).tName = msg2
            End If
        Else
            lTracks.tTrack(i).tName = "Track " & i
        End If
        DoEvents
    Next i
    msg3 = "TTITLE" & j - 1 & "="
    msg2 = ParseString(msg, msg3, msg5)
    msg2 = Right(msg2, Len(msg2) - Len(msg3) + 1)
    lTracks.tTrack(j).tName = Left(msg2, Len(msg2) - 2)
    If lEvents.eSettings.iFreeDB.cSaveTracksToDisk = True Then SaveCDTracks lRipperSettings.eDiscID
    wskFreeDB.Close
    If lEvents.eSettings.iFreeDB.cShowDialog = True Then
        If frmWait.Height = 1460 Then Unload frmWait
    End If
ElseIf Left(msg, 3) = "211" Then
    frmMain.mnuArtist.Caption = "Artist: <None>"
    frmMain.mnuAlbum.Caption = "Album: <None>"
    frmMain.mnuArtist.Enabled = False
    frmMain.mnuAlbum.Enabled = False
    If lEvents.eSettings.iFreeDB.cShowDialog = True Then
        If frmWait.Height = 1460 Then Unload frmWait
    End If
    If lEvents.eSettings.iFreeDB.cAutoSubmit = True Then
        If lEvents.eSettings.iOverwritePrompts = True Then
            msg = MsgBox("No FreeDB match was found. Would you like to edit one yourself?", vbQuestion + vbYesNo, "No Match Found")
            If msg = vbYes Then ShowEditDisc
        Else
            ShowEditDisc
        End If
    End If
ElseIf Left(msg, 3) = "202" Then
    frmMain.mnuArtist.Caption = "Artist: <None>"
    frmMain.mnuAlbum.Caption = "Album: <None>"
    frmMain.mnuArtist.Enabled = False
    frmMain.mnuAlbum.Enabled = False
    If lEvents.eSettings.iFreeDB.cShowDialog = True Then
        If frmWait.Height = 1460 Then Unload frmWait
    End If
    If lEvents.eSettings.iFreeDB.cAutoSubmit = True Then
        If lEvents.eSettings.iOverwritePrompts = True Then
            msg = MsgBox("No FreeDB match was found. Would you like to edit one yourself?", vbQuestion + vbYesNo, "No Match Found")
            If msg = vbYes Then ShowEditDisc
        Else
            ShowEditDisc
        End If
    End If
ElseIf Left(msg, 3) = "431" Then
    frmMain.mnuArtist.Caption = "Artist: <None>"
    frmMain.mnuAlbum.Caption = "Album: <None>"
    frmMain.mnuArtist.Enabled = False
    frmMain.mnuAlbum.Enabled = False
    If lEvents.eSettings.iFreeDB.cShowDialog = True Then
        If frmWait.Height = 1460 Then Unload frmWait
    End If
    If lEvents.eSettings.iFreeDB.cAutoSubmit = True Then
        If lEvents.eSettings.iOverwritePrompts = True Then
            msg = MsgBox("No FreeDB match was found. Would you like to edit one yourself?", vbQuestion + vbYesNo, "No Match Found")
            If msg = vbYes Then ShowEditDisc
        Else
            ShowEditDisc
        End If
    End If
    wskFreeDB.Close
End If

If Err.Number <> 0 Then SetError "wskFreeDB", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub wskFreeDB_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Local Error Resume Next
SetError "wskFreeDB_Error", "a winsock error", Description
End Sub

Private Sub wskUpdate_Close()
On Local Error Resume Next
Dim msg As String, msg2 As String

If Len(lUpdateINI) <> 0 Then
    wskUpdate.Close: wskUpdate.Tag = "CLOSED"
    If DoesFileExist(lIniFiles.iUpdate) = True Then Kill lIniFiles.iUpdate: DoEvents
    SaveFile lIniFiles.iUpdate, lUpdateINI: DoEvents
    pause 0.2
    msg2 = ReadINI(lIniFiles.iUpdate, "Settings", "Version", "")
    If msg2 <> App.Major & "." & App.Minor Then
        lblInfo.Caption = "Update available"
        pause 0.5
        frmLatestVersionCheck.Show
    End If
Else
    wskUpdate.Close: wskUpdate.Tag = "CLOSED"
End If
If Err.Number <> 0 Then SetError "wskUpdate_Close()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub wskUpdate_Connect()
On Local Error Resume Next
Dim getString As String, ShortWebSite As String
wskUpdate.Tag = "OPEN"
ShortWebSite = "http://www.team-nexgen.com/ns4update.ini"
getString = "GET " + ShortWebSite + " HTTP/1.0" + vbCrLf
getString = getString + "Accept: */*" + vbCrLf
getString = getString + "Accept: text/html" + vbCrLf
getString = getString + vbCrLf
wskUpdate.SendData getString
If Err.Number <> 0 Then SetError "wskUpdate_Connect()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub wskUpdate_DataArrival(ByVal bytesTotal As Long)
On Local Error Resume Next
Dim Buffer As String
If wskUpdate.Tag = "OPEN" Then wskUpdate.GetData Buffer
lUpdateINI = lUpdateINI & Buffer
If Err.Number <> 0 Then SetError "wskUpdate_DataArrival()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub wskUpdate_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Local Error Resume Next
SetError "wskUpdate_Error", "A winsock error occured", Description
End Sub
