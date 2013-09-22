VERSION 5.00
Object = "{EE128208-4F73-11D3-83BB-C47C02EE3D01}#1.0#0"; "ControlResizer.ocx"
Begin VB.Form frmSettings 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "NexENCODE - Settings"
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   5955
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   237
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   397
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton cmdDefaults 
      Caption         =   "&Defaults"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   64
      Top             =   2640
      Width           =   1695
   End
   Begin VB.OptionButton cmdCDDB 
      Caption         =   "Register"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Width           =   1695
   End
   Begin VB.OptionButton cmdAspi 
      Caption         =   "Aspi"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   1920
      Width           =   1695
   End
   Begin VB.OptionButton cmdFreeDB 
      Caption         =   "FreeDB"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   1560
      Width           =   1695
   End
   Begin VB.OptionButton cmdGeneral 
      Caption         =   "General"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   1200
      Width           =   1695
   End
   Begin VB.OptionButton cmdPlayers 
      Caption         =   "Players"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   840
      Width           =   1695
   End
   Begin VB.OptionButton cmdRipper 
      Caption         =   "Ripper"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.OptionButton cmdEncoder 
      Caption         =   "Encoder"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   6015
      TabIndex        =   14
      Top             =   3120
      Width           =   6015
      Begin VB.ComboBox cboSave 
         Height          =   315
         ItemData        =   "frmSettings.frx":000C
         Left            =   120
         List            =   "frmSettings.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   90
         Width           =   3135
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   315
         Left            =   4680
         TabIndex        =   69
         Top             =   90
         Width           =   1215
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   315
         Left            =   3360
         TabIndex        =   68
         Top             =   90
         Width           =   1215
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         X1              =   6120
         X2              =   0
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.PictureBox TRASH 
      Height          =   975
      Left            =   1800
      ScaleHeight     =   915
      ScaleWidth      =   4155
      TabIndex        =   77
      Top             =   3960
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CommandButton cmdChangeDir 
         Caption         =   "..."
         Height          =   255
         Left            =   3000
         TabIndex        =   79
         Top             =   550
         Width           =   855
      End
      Begin VB.TextBox txtTempFiles 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   0
         TabIndex        =   78
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame fraCDDB2 
      BorderStyle     =   0  'None
      Caption         =   "Encoder Options"
      Height          =   3015
      Left            =   1920
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CommandButton Command1 
         Caption         =   "Register ..."
         Height          =   315
         Left            =   2760
         TabIndex        =   39
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Continue using NexENCODE for free for as long as you like. If you find NexENCODE to be of use, please register."
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   120
         TabIndex        =   40
         Top             =   720
         Width           =   3735
      End
      Begin VB.Label lblRegistered 
         BackStyle       =   0  'Transparent
         Caption         =   "Unregistered version"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   38
         Top             =   240
         Width           =   3015
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   120
         Picture         =   "frmSettings.frx":0049
         Stretch         =   -1  'True
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.Frame fraRipper 
      BorderStyle     =   0  'None
      Caption         =   "Ripper"
      Height          =   3015
      Left            =   1920
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   4095
      Begin VB.TextBox txtOutputDir 
         Height          =   285
         Left            =   120
         TabIndex        =   81
         Top             =   2640
         Width           =   3135
      End
      Begin VB.CommandButton cmdChangeOutDir 
         Caption         =   "..."
         Height          =   285
         Left            =   3360
         TabIndex        =   80
         Top             =   2640
         Width           =   615
      End
      Begin VB.ListBox lstCDDrive 
         Height          =   450
         Left            =   120
         TabIndex        =   58
         Top             =   1920
         Width           =   3855
      End
      Begin VB.CheckBox chkAutoDeleteWavs 
         Appearance      =   0  'Flat
         Caption         =   "Auto delete ripped wav files"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   120
         Width           =   2655
      End
      Begin VB.CheckBox chkLockTrayOnRip 
         Appearance      =   0  'Flat
         Caption         =   "Lock CD tray durring rip"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   360
         Width           =   2775
      End
      Begin VB.CheckBox chkAutoEject 
         Appearance      =   0  'Flat
         Caption         =   "Auto eject cd after rip"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   600
         Width           =   3855
      End
      Begin VB.ComboBox cboCopyMode 
         Height          =   315
         ItemData        =   "frmSettings.frx":0D13
         Left            =   120
         List            =   "frmSettings.frx":0D20
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1200
         Width           =   3855
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Output Directory:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   82
         Top             =   2400
         Width           =   3375
      End
      Begin VB.Label lblCdDrive 
         BackStyle       =   0  'Transparent
         Caption         =   "CD Drive:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Copy Mode:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   3495
      End
   End
   Begin VB.Frame fraEncoder 
      BorderStyle     =   0  'None
      Caption         =   "Encoder Options"
      Height          =   3015
      Left            =   1920
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   4095
      Begin VB.OptionButton optAfterEncode 
         Appearance      =   0  'Flat
         Caption         =   "Show process report"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   76
         Top             =   960
         Width           =   1815
      End
      Begin VB.OptionButton optAfterEncode 
         Appearance      =   0  'Flat
         Caption         =   "Do nothing"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   75
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CheckBox chkCreateAlbumFileOnEncode 
         Appearance      =   0  'Flat
         Caption         =   "Create Album"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2040
         TabIndex        =   71
         Top             =   600
         Width           =   1455
      End
      Begin VB.CheckBox chkCopyrighted 
         Appearance      =   0  'Flat
         Caption         =   "Copyrighted"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   840
         Width           =   1575
      End
      Begin VB.CheckBox chkDownsample 
         Appearance      =   0  'Flat
         Caption         =   "Downsample"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Width           =   1575
      End
      Begin VB.CheckBox chkOrigional 
         Appearance      =   0  'Flat
         Caption         =   "Origional work"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CheckBox chkDownmix 
         Appearance      =   0  'Flat
         Caption         =   "Downmix"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   1575
      End
      Begin VB.CheckBox chkAutoAddTags 
         Appearance      =   0  'Flat
         Caption         =   "Add Tags"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2040
         TabIndex        =   30
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cboProfile 
         Height          =   315
         ItemData        =   "frmSettings.frx":0D39
         Left            =   120
         List            =   "frmSettings.frx":0D49
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1680
         Width           =   3855
      End
      Begin VB.ListBox lstBitrate 
         Height          =   645
         ItemData        =   "frmSettings.frx":0D90
         Left            =   2160
         List            =   "frmSettings.frx":0DBE
         TabIndex        =   5
         Top             =   2280
         Width           =   1815
      End
      Begin VB.ListBox lstSampleRate 
         Height          =   645
         ItemData        =   "frmSettings.frx":0E01
         Left            =   120
         List            =   "frmSettings.frx":0E0E
         TabIndex        =   4
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Encoder Settings:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "After Encode:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2040
         TabIndex        =   73
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Current profile:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Sample Rate:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label lblBitrate 
         BackStyle       =   0  'Transparent
         Caption         =   "Bitrate:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2160
         TabIndex        =   7
         Top             =   2040
         Width           =   1575
      End
   End
   Begin VB.Frame fraFreeDB 
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   1920
      TabIndex        =   45
      Top             =   0
      Width           =   4095
      Begin VB.CheckBox chkShowDialog 
         Appearance      =   0  'Flat
         Caption         =   "Show Dialog"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   2040
         Width           =   3855
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   315
         Left            =   120
         TabIndex        =   56
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   315
         Left            =   120
         TabIndex        =   55
         Top             =   1200
         Width           =   975
      End
      Begin VB.ListBox lstCDDBServer 
         Height          =   840
         Left            =   1440
         TabIndex        =   54
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox txtFreeDBServer 
         Height          =   285
         Left            =   1440
         TabIndex        =   53
         Top             =   480
         Width           =   2535
      End
      Begin VB.CheckBox chkUseFirstMatchOnFuzzyMatch 
         Appearance      =   0  'Flat
         Caption         =   "Use first match on fuzzy match"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   1800
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.CheckBox chkAutoSubmitCDDB 
         Appearance      =   0  'Flat
         Caption         =   "Enter tracks manually when match not found"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   2280
         Width           =   3735
      End
      Begin VB.CheckBox chkSaveTracksToDisc 
         Appearance      =   0  'Flat
         Caption         =   "Save track names to disk"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   2520
         Width           =   3855
      End
      Begin VB.CheckBox chkCDDBEnabled 
         Appearance      =   0  'Flat
         Caption         =   "Enable CDDB"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   2760
         Width           =   3735
      End
      Begin VB.TextBox txtEmailAddress 
         Height          =   285
         Left            =   1440
         TabIndex        =   47
         Top             =   120
         Width           =   2535
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "FreeDB Server:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail address:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.Frame fraGeneral 
      BorderStyle     =   0  'None
      Caption         =   "Encoder Options"
      Height          =   3015
      Left            =   1920
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CheckBox chkRememberWindowSizes 
         Appearance      =   0  'Flat
         Caption         =   "Remember Window Sizes"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   85
         Top             =   1800
         Width           =   3735
      End
      Begin VB.CheckBox chkAlwaysOnTop 
         Appearance      =   0  'Flat
         Caption         =   "Always on Top"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Top             =   1560
         Width           =   3615
      End
      Begin VB.CheckBox chkCheckForActiveWindow 
         Appearance      =   0  'Flat
         Caption         =   "Check for active window"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   1320
         Width           =   3735
      End
      Begin VB.CheckBox chkUpdateCheck 
         Appearance      =   0  'Flat
         Caption         =   "Automatically Check for Updates"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   1080
         Width           =   3015
      End
      Begin VB.CheckBox chkShowAboutScreen 
         Appearance      =   0  'Flat
         Caption         =   "Show About"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   840
         Width           =   3735
      End
      Begin VB.CommandButton cmdWizard 
         Caption         =   "Wizard ..."
         Height          =   375
         Left            =   2880
         TabIndex        =   42
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CheckBox chkPlayWavs 
         Appearance      =   0  'Flat
         Caption         =   "Play system sounds"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   600
         Width           =   3735
      End
      Begin VB.CheckBox chkDisplayErrors 
         Appearance      =   0  'Flat
         Caption         =   "Display errors"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   3855
      End
      Begin VB.CheckBox chkOverwritePrompts 
         Appearance      =   0  'Flat
         Caption         =   "Prompts/Overwrite prompts"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   120
         Width           =   3855
      End
   End
   Begin VB.Frame fraASPI 
      BorderStyle     =   0  'None
      Caption         =   "ASPI"
      ForeColor       =   &H00000000&
      Height          =   3015
      Left            =   1920
      TabIndex        =   59
      Top             =   0
      Width           =   4095
      Begin VB.CommandButton cmdAspiChk 
         Caption         =   "ASPI INSTALLATION VERIFICATION"
         Height          =   855
         Left            =   120
         TabIndex        =   62
         Top             =   2040
         Width           =   3855
      End
      Begin VB.CommandButton cmdAspiUpd 
         Caption         =   "UPDATE V4.57 TO V4.60"
         Height          =   855
         Left            =   120
         TabIndex        =   61
         Top             =   1080
         Width           =   3855
      End
      Begin VB.CommandButton chkAspiInstall 
         Caption         =   "INSTALL V4.57 ASPI DRIVERS"
         Height          =   855
         Left            =   120
         TabIndex        =   60
         Top             =   120
         Width           =   3855
      End
   End
   Begin VB.Frame fraPlayers 
      BorderStyle     =   0  'None
      Caption         =   "Encoder Options"
      Height          =   3015
      Left            =   1920
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CheckBox chkContinuous 
         Appearance      =   0  'Flat
         Caption         =   "Continuous Mode"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   83
         Top             =   2520
         Width           =   3735
      End
      Begin VB.CommandButton cmdSearchMp3Players 
         Caption         =   "Search"
         Height          =   255
         Left            =   1920
         TabIndex        =   72
         Top             =   480
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdDelCdAudioPlayer 
         Caption         =   "Del"
         Height          =   255
         Left            =   3360
         TabIndex        =   23
         Top             =   1200
         Width           =   615
      End
      Begin VB.CommandButton cmdCDAudioPlayerAdd 
         Caption         =   "Add"
         Height          =   255
         Left            =   2760
         TabIndex        =   24
         Top             =   1200
         Width           =   615
      End
      Begin VB.CheckBox chkPlayMp3sInNexENCODE 
         Appearance      =   0  'Flat
         Caption         =   "Play Mpeg Layer 3 in NexENCODE"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   2760
         Width           =   3735
      End
      Begin VB.ComboBox cboCDAudioPlayer 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   840
         Width           =   2775
      End
      Begin VB.ComboBox cboMpegPlayer 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   120
         Width           =   2775
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "Del"
         Height          =   255
         Left            =   3360
         TabIndex        =   20
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton cmdAddMpegPlayer 
         Caption         =   "Add"
         Height          =   255
         Left            =   2760
         TabIndex        =   19
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "CD Player:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   840
         Width           =   3855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Mpeg Player:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   3855
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Temp Files:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4200
      TabIndex        =   86
      Top             =   4200
      Width           =   3615
   End
   Begin ControlResizer.AutoResizer AutoResizer1 
      Height          =   255
      Left            =   3720
      Top             =   120
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Enum eSettingsFrames
    eNone = 0
    eEncoder = 1
    eRipper = 2
    eCDDB2 = 3
    ePlayers = 4
    eGeneral = 5
    eFreeDB = 6
    eAspi = 7
End Enum

Public Function SaveSettings() As Boolean
'On Local Error Resume Next
If Len(txtOutputDir.Text) = 0 Then
    Beep
    ResetSettingsFrames eGeneral
    txtOutputDir.SetFocus
    Exit Function
End If
If optAfterEncode(0).Value = True Then
    lEvents.eSettings.iShowReports = True
    lEvents.eSettings.iAutoPlay = False
ElseIf optAfterEncode(2).Value = True Then
    lEvents.eSettings.iAutoPlay = False
    lEvents.eSettings.iShowReports = False
End If
lEncoderSettings.eBitrate = lstBitrate.Text
lEncoderSettings.eSampleRate = lstSampleRate.Text
lEncoderSettings.eProfile = cboProfile.ListIndex
lRipperSettings.eCopyMode = cboCopyMode.ListIndex + 1
lRipperSettings.eDriveLetter = lstCDDrive.Text
lEvents.eSettings.iRememberWindowSizes = GetCheckboxValue(chkRememberWindowSizes)
lEvents.eSettings.iAlwaysOnTop = GetCheckboxValue(chkAlwaysOnTop)
lEvents.eSettings.iRememberWindowSizes = GetCheckboxValue(chkRememberWindowSizes)
lEvents.eSettings.iUpdateCheck = GetCheckboxValue(chkUpdateCheck)
lEvents.eSettings.iPlayMp3sInNexENCODE = GetCheckboxValue(chkPlayMp3sInNexENCODE)
lEvents.eSettings.iOverwritePrompts = GetCheckboxValue(chkOverwritePrompts)
lEvents.eSettings.iShowErrors = GetCheckboxValue(chkDisplayErrors)
lEvents.eSettings.iPlayWavs = GetCheckboxValue(chkPlayWavs)
lEvents.eSettings.iShowAbout = GetCheckboxValue(chkShowAboutScreen)
lEvents.eSettings.iCheckForActiveWindow = GetCheckboxValue(chkCheckForActiveWindow)
lEvents.eSettings.iCreateAlbumFileOnEncode = GetCheckboxValue(chkCreateAlbumFileOnEncode)
With lEvents.eSettings.iFreeDB
    .cShowDialog = GetCheckboxValue(chkShowDialog)
    .cEnabled = GetCheckboxValue(chkCDDBEnabled)
    .cSaveTracksToDisk = GetCheckboxValue(chkSaveTracksToDisc)
    .cAutoSubmit = GetCheckboxValue(chkAutoSubmitCDDB)
    .cEmailAddress = txtEmailAddress.Text
    .cServer = txtFreeDBServer.Text
    .cUseFirstMatch = GetCheckboxValue(chkUseFirstMatchOnFuzzyMatch)
End With
lEncoderSettings.eOutputDir = txtOutputDir.Text
If Right(lEncoderSettings.eOutputDir, 1) <> "\" Then lEncoderSettings.eOutputDir = lEncoderSettings.eOutputDir & "\"
lEncoderSettings.eDownmix = GetCheckboxValue(chkDownmix)
lEncoderSettings.eAutoAddTags = GetCheckboxValue(chkAutoAddTags)
lEncoderSettings.eCopyrighted = GetCheckboxValue(chkCopyrighted)
lEncoderSettings.eDownsample = GetCheckboxValue(chkDownsample)
lEncoderSettings.eOrigionalWork = GetCheckboxValue(chkOrigional)
lRipperSettings.eAutoEject = GetCheckboxValue(chkAutoEject)
lRipperSettings.eAutoDeleteRipedFiles = GetCheckboxValue(chkAutoDeleteWavs)
lRipperSettings.eLockCDTrayDuringRip = GetCheckboxValue(chkLockTrayOnRip)
lPlayer.pContinuous = GetCheckboxValue(chkContinuous)
If Len(cboMpegPlayer.Text) <> 0 Then lPlayers.pMp3PlayerIndex = FindPlayerIndex(cboMpegPlayer.Text)
If Len(cboCDAudioPlayer.Text) <> 0 Then lPlayers.pCDPlayerIndex = FindPlayerIndex(cboCDAudioPlayer.Text)
If cboSave.ListIndex = 1 Then
    WriteINI lIniFiles.iSettings, "Settings", "RememberWindowSizes", lEvents.eSettings.iRememberWindowSizes
    WriteINI lIniFiles.iSettings, "Settings", "AlwaysOnTop", lEvents.eSettings.iAlwaysOnTop
    WriteINI lIniFiles.iSettings, "Settings", "CreateAlbumFileOnEncode", lEvents.eSettings.iCreateAlbumFileOnEncode
    WriteINI lIniFiles.iSettings, "Settings", "CheckForActiveWindow", lEvents.eSettings.iCheckForActiveWindow
    WriteINI lIniFiles.iSettings, "Settings", "ShowAbout", lEvents.eSettings.iShowAbout
    WriteINI lIniFiles.iSettings, "Settings", "DriveLetter", lRipperSettings.eDriveLetter
    WriteINI lIniFiles.iSettings, "Settings", "PlayMp3sInNexENCODE", lEvents.eSettings.iPlayMp3sInNexENCODE
    WriteINI lIniFiles.iSettings, "Settings", "OutputDir", lEncoderSettings.eOutputDir
    WriteINI lIniFiles.iSettings, "Settings", "PlayWavs", lEvents.eSettings.iPlayWavs
    WriteINI lIniFiles.iSettings, "Settings", "OverwritePrompts", lEvents.eSettings.iOverwritePrompts
    WriteINI lIniFiles.iSettings, "Settings", "ShowErrors", lEvents.eSettings.iShowErrors
    WriteINI lIniFiles.iSettings, "Settings", "ShowReports", lEvents.eSettings.iShowReports
    WriteINI lIniFiles.iSettings, "Settings", "AutoAddTags", lEncoderSettings.eAutoAddTags
    WriteINI lIniFiles.iSettings, "Settings", "UpdateCheck", lEvents.eSettings.iUpdateCheck
    WriteINI lIniFiles.iSettings, "Settings", "AutoPlay", lEvents.eSettings.iAutoPlay
    WriteINI lIniFiles.iPlayers, "Settings", "Mp3Player", lPlayers.pMp3PlayerIndex
    WriteINI lIniFiles.iPlayers, "Settings", "CDPlayer", lPlayers.pCDPlayerIndex
    WriteINI lIniFiles.iSettings, "Settings", "CopyMode", lRipperSettings.eCopyMode
    WriteINI lIniFiles.iSettings, "Settings", "AutoDeleteRipedFiles", lRipperSettings.eAutoDeleteRipedFiles
    WriteINI lIniFiles.iSettings, "Settings", "LockCDTrayDuringRip", lRipperSettings.eLockCDTrayDuringRip
    WriteINI lIniFiles.iSettings, "Settings", "Bitrate", lstBitrate.Text
    WriteINI lIniFiles.iSettings, "Settings", "SampleRate", lstSampleRate.Text
    WriteINI lIniFiles.iSettings, "Settings", "Profile", cboProfile.ListIndex
    WriteINI lIniFiles.iSettings, "Settings", "Downmix", lEncoderSettings.eDownmix
    WriteINI lIniFiles.iSettings, "Settings", "Downsample", lEncoderSettings.eDownsample
    WriteINI lIniFiles.iSettings, "Settings", "Origional", lEncoderSettings.eOrigionalWork
    WriteINI lIniFiles.iSettings, "Settings", "CopyRighted", lEncoderSettings.eCopyrighted
    WriteINI lIniFiles.iSettings, "CDDB", "UseFirstMatch", lEvents.eSettings.iFreeDB.cUseFirstMatch
    WriteINI lIniFiles.iSettings, "CDDB", "SaveTracksToDisk", lEvents.eSettings.iFreeDB.cSaveTracksToDisk
    WriteINI lIniFiles.iSettings, "CDDB", "UseFirstMatch", lEvents.eSettings.iFreeDB.cUseFirstMatch
    WriteINI lIniFiles.iSettings, "CDDB", "ShowDialog", lEvents.eSettings.iFreeDB.cShowDialog
    WriteINI lIniFiles.iSettings, "CDDB", "AutoSubmit", lEvents.eSettings.iFreeDB.cAutoSubmit
    WriteINI lIniFiles.iSettings, "CDDB", "EmailAddress", lEvents.eSettings.iFreeDB.cEmailAddress
    WriteINI lIniFiles.iSettings, "CDDB", "Server", lEvents.eSettings.iFreeDB.cServer
    WriteINI lIniFiles.iSettings, "CDDB", "Enabled", lEvents.eSettings.iFreeDB.cEnabled
    WriteINI lIniFiles.iPlayers, "Settings", "Continuous", lPlayer.pContinuous
End If
ResetPlayButtons
SaveSettings = True
If lEvents.eSettings.iCheckForActiveWindow = True Then
    frmMain.tmrCheckActive.Enabled = True
Else
    frmMain.tmrCheckActive.Enabled = False
    frmMain.Picture = frmMain.imgBackground1.Picture
End If
PlayWav App.Path & "\media\done.wav", SND_ASYNC
If Err.Number <> 0 Then SetError "SaveSettings()", lEvents.eSettings.iErrDescription, Err.Description
End Function

Public Sub SetSettingsObjects()
'On Local Error Resume Next
Dim i As Integer
Icon = frmMain.Icon
For i = 1 To lCDDBServ.cCount
    If Len(lCDDBServ.cServer(i).sLocation) <> 0 Then lstCDDBServer.AddItem lCDDBServ.cServer(i).sLocation
Next i
txtOutputDir.Text = lEncoderSettings.eOutputDir
If lEvents.eSettings.iFreeDB.cAutoSubmit = True Then chkAutoSubmitCDDB.Value = 1
txtEmailAddress.Text = lEvents.eSettings.iFreeDB.cEmailAddress
txtFreeDBServer.Text = lEvents.eSettings.iFreeDB.cServer
If lEvents.eSettings.iCreateAlbumFileOnEncode = True Then chkCreateAlbumFileOnEncode.Value = 1
If lEvents.eSettings.iFreeDB.cUseFirstMatch = True Then chkUseFirstMatchOnFuzzyMatch.Value = 1
If lEvents.eSettings.iFreeDB.cEnabled = True Then chkCDDBEnabled.Value = 1
If lEvents.eSettings.iFreeDB.cShowDialog = True Then chkShowDialog.Value = 1
If lEvents.eSettings.iFreeDB.cSaveTracksToDisk = True Then chkSaveTracksToDisc.Value = 1
If lEvents.eSettings.iShowAbout = True Then chkShowAboutScreen.Value = 1
If lEvents.eSettings.iPlayMp3sInNexENCODE = True Then chkPlayMp3sInNexENCODE.Value = 1
If lEvents.eSettings.iPlayWavs = True Then chkPlayWavs.Value = 1
If lEvents.eSettings.iUpdateCheck = True Then chkUpdateCheck.Value = 1
If lEvents.eSettings.iAlwaysOnTop = True Then chkAlwaysOnTop.Value = 1
If lEvents.eSettings.iAutoPlay = True Then
'    optAfterEncode(1).Value = True
ElseIf lEvents.eSettings.iShowReports = True Then
    optAfterEncode(0).Value = True
Else
    optAfterEncode(2).Value = True
End If
If lEvents.eSettings.iOverwritePrompts = True Then
    chkOverwritePrompts.Value = 1
Else
    chkOverwritePrompts.Value = 0
End If
If lEvents.eSettings.iCheckForActiveWindow = True Then
    chkCheckForActiveWindow.Value = 1
Else
    chkCheckForActiveWindow.Value = 0
End If
If lEvents.eSettings.iShowErrors = True Then
    chkDisplayErrors.Value = 1
Else
    chkDisplayErrors.Value = 0
End If
If lEncoderSettings.eAutoAddTags = True Then
    chkAutoAddTags.Value = 1
Else
    chkAutoAddTags.Value = 0
End If
If lRipperSettings.eAutoEject = True Then
    chkAutoEject.Value = 1
Else
    chkAutoEject.Value = 0
End If
If lEvents.eSettings.iRememberWindowSizes = True Then
    chkRememberWindowSizes.Value = 1
Else
    chkRememberWindowSizes.Value = 0
End If
cboSave.ListIndex = 1
If lPlayer.pContinuous = True Then chkContinuous.Value = 1
If lPlayers.pCount <> 0 Then
    For i = 1 To lPlayers.pCount
        If Len(lPlayers.pPlayer(i).pName) <> 0 Then
            If lPlayers.pPlayer(i).pType = pMp3Player Then
                cboMpegPlayer.AddItem LCase(lPlayers.pPlayer(i).pName)
'            ElseIf lPlayers.pPlayer(i).pType = pWavPlayer Then
'                cboWavPlayer.AddItem LCase(lPlayers.pPlayer(i).pName)
            ElseIf lPlayers.pPlayer(i).pType = pCDPlayer Then
                cboCDAudioPlayer.AddItem LCase(lPlayers.pPlayer(i).pName)
            End If
        End If
    Next i
    Dim u As Integer
    If lEvents.eSettings.iPlayMp3sInNexENCODE = False And lPlayers.pMp3PlayerIndex <> 0 And cboMpegPlayer.ListCount <> 0 Then
        u = FindComoboxIndex(cboMpegPlayer, lPlayers.pPlayer(lPlayers.pMp3PlayerIndex).pName)
        If u <> -1 Then cboMpegPlayer.ListIndex = u
    End If
    If lPlayers.pCDPlayerIndex <> 0 Then cboCDAudioPlayer.ListIndex = FindComoboxIndex(cboCDAudioPlayer, lPlayers.pPlayer(lPlayers.pCDPlayerIndex).pName)
    'If lPlayers.pWavPlayerIndex <> 0 Then cboWavPlayer.ListIndex = FindComoboxIndex(cboWavPlayer, lPlayers.pPlayer(lPlayers.pWavPlayerIndex).pName)
End If
'txtTempFiles.Text = lRipperSettings.eTempFiles
If lRipperSettings.eAutoDeleteRipedFiles = True Then chkAutoDeleteWavs.Value = 1
If lRipperSettings.eLockCDTrayDuringRip = True Then chkLockTrayOnRip.Value = 1
cboProfile.ListIndex = lEncoderSettings.eProfile
lstBitrate.Text = lEncoderSettings.eBitrate
lstSampleRate.Text = lEncoderSettings.eSampleRate
If lEncoderSettings.eCopyrighted = True Then chkCopyrighted.Value = 1
If lEncoderSettings.eDownsample = True Then chkDownsample.Value = 1
If lEncoderSettings.eOrigionalWork = True Then chkOrigional.Value = 1
If lEncoderSettings.eDownmix = True Then chkDownmix.Value = 1
cboCopyMode.ListIndex = lRipperSettings.eCopyMode - 1
frmMain.Ripper.Init
FillListboxWithDrives lstCDDrive
lstCDDrive.ListIndex = FindListboxIndex(lstCDDrive, lRipperSettings.eDriveLetter)
If Err.Number = 68 Then
    Err = 0
    If lEvents.eSettings.iOverwritePrompts = True Then MsgBox "NexENCODE Studio was unable to load your CD Audio device because it is in use, or no CD is in the drive.", vbExclamation, "Error"
End If
If Err.Number <> 0 Then SetError "SetSettingsObjects()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Public Sub ResetSettingsFrames(lFrame As eSettingsFrames, Optional lShow As Boolean)
'On Local Error Resume Next
fraRipper.Visible = False
fraPlayers.Visible = False
fraEncoder.Visible = False
fraCDDB2.Visible = False
fraFreeDB.Visible = False
fraGeneral.Visible = False
fraASPI.Visible = False
Select Case lFrame
Case eAspi
    fraASPI.Visible = True
    If cmdAspi.Value = False Then cmdAspi.Value = True
'    Caption = "Settings - Aspi"
Case eEncoder
    fraEncoder.Visible = True
    If cmdEncoder.Value = False Then cmdEncoder.Value = True
'    Caption = "Settings - MP3 Encoder"
Case eRipper
    fraRipper.Visible = True
    If cmdRipper.Value = False Then cmdRipper.Value = True
'    Caption = "Settings - CD Ripper"
Case eCDDB2
    fraCDDB2.Visible = True
    If cmdCDDB.Value = False Then cmdCDDB.Value = True
'    Caption = "Settings - Register"
Case ePlayers
    fraPlayers.Visible = True
    If cmdPlayers.Value = False Then cmdPlayers.Value = True
'    Caption = "Settings - Players"
Case eGeneral
    fraGeneral.Visible = True
    If cmdGeneral.Value = False Then cmdGeneral.Value = True
'    Caption = "Settings - General"
Case eFreeDB
    fraFreeDB.Visible = True
    If cmdFreeDB.Value = False Then cmdFreeDB.Value = True
'    Caption = "Settings - FreeDB"
End Select
If lShow = True Then Me.Show
If Err.Number <> 0 Then SetError "ResetSettingsFrames()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cboProfile_Click()
'On Local Error Resume Next
Select Case cboProfile.ListIndex
Case 0
    lstSampleRate.ListIndex = 0
    lstBitrate.ListIndex = 1
    chkDownsample.Value = 0
Case 1
    lstSampleRate.ListIndex = 1
    lstBitrate.ListIndex = 3
    chkDownsample.Value = 1
Case 2
    lstSampleRate.ListIndex = 1
    lstBitrate.ListIndex = 5
    chkDownsample.Value = 1
End Select
If Err.Number <> 0 Then SetError "cboProfile_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub chkAspiInstall_Click()
'On Local Error Resume Next
Shell App.Path & "\programs\aspiupd.exe", vbNormalFocus
End
End Sub

Private Sub chkPlayMp3sInNexENCODE_Click()
'On Local Error Resume Next
If chkPlayMp3sInNexENCODE.Value = 1 Then
    cboMpegPlayer.Enabled = False
    cmdAddMpegPlayer.Enabled = False
    cmdDel.Enabled = False
Else
    cboMpegPlayer.Enabled = True
    cmdAddMpegPlayer.Enabled = True
    cmdDel.Enabled = True
End If
If Err.Number <> 0 Then SetError "chkPlayMp3sInNexENCODE()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdAdd_Click()
'On Local Error Resume Next
Dim i As Integer
i = AddCDDBServer(InputBox("Enter Location (Example: st. paul, minnesota)", "Add CDDB Server", ""), InputBox("Enter Server Ip or address (Example: freedb.freedb.org)", "Add CDDB Server", ""))
If i <> 0 Then
    lstCDDBServer.AddItem lCDDBServ.cServer(i).sLocation
End If
If Err.Number <> 0 Then SetError "cmdAdd_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdAddMpegPlayer_Click()
'On Local Error Resume Next
Dim msg As String, msg2 As String, i As Integer, X As Integer
msg = InputBox("Enter name of player:", "", "Audica")
If Len(msg) = 0 Then Exit Sub
msg2 = OpenDialog(frmSettings, "Mp3 Players (*.exe)|*.exe|All Files (*.*)|*.*", "Select Player ...", App.Path)
If Len(msg2) = 0 Then Exit Sub
If Len(msg) <> 0 And Len(msg2) <> 0 Then
    X = FindPlayerIndex(msg)
    If X = 0 Then
        i = AddPlayer(msg, msg2, pMp3Player, msg & ".m3u")
        If i <> 0 Then cboMpegPlayer.AddItem lPlayers.pPlayer(i).pName
    Else
        cboMpegPlayer.AddItem lPlayers.pPlayer(X).pName
    End If
    cboMpegPlayer.ListIndex = FindComoboxIndex(cboMpegPlayer, lPlayers.pPlayer(X).pName)
End If
If Err.Number <> 0 Then SetError "cmdAddMpegPlayer_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdAspi_Click()
'On Local Error Resume Next
ResetSettingsFrames eAspi
PlayWav App.Path & "\media\click.wav", SND_ASYNC
If Err.Number <> 0 Then SetError "cmdAspi_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdAspiChk_Click()
'On Local Error Resume Next
Shell App.Path & "\programs\aspichk.exe", vbNormalFocus
End Sub

Private Sub cmdAspiUpd_Click()
'On Local Error Resume Next
Shell App.Path & "\programs\aspi32.exe", vbNormalFocus
End Sub

Private Sub cmdCancel_Click()
'On Local Error Resume Next
Unload Me
End Sub

Private Sub cmdCDAudioPlayerAdd_Click()
'On Local Error Resume Next
Dim msg As String, msg2 As String, i As Integer, X As Integer
msg = InputBox("Enter name of player:", "", "NexMedia")
If Len(msg) = 0 Then Exit Sub
msg2 = OpenDialog(frmSettings, "CD Players (*.exe)|*.exe|All Files (*.*)|*.*", "Select Player ...", App.Path)
If Len(msg2) = 0 Then Exit Sub
If Len(msg) <> 0 And Len(msg2) <> 0 Then
    X = FindPlayerIndex(msg)
    If X = 0 Then
        i = AddPlayer(msg, msg2, pCDPlayer, msg & ".m3u")
        If i <> 0 Then cboCDAudioPlayer.AddItem lPlayers.pPlayer(i).pName
    Else
        cboCDAudioPlayer.AddItem lPlayers.pPlayer(X).pName
    End If
End If
If Err.Number <> 0 Then SetError "cmdAddCDPlayer_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdCDDB_Click()
'On Local Error Resume Next
PlayWav App.Path & "\media\click.wav", SND_ASYNC
ResetSettingsFrames eCDDB2
If Err.Number <> 0 Then SetError "cmdCDDB_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdChangeDir_Click()
'On Local Error Resume Next
Dim msg As String, i As Integer
frmSelectDir.Show 1
If Len(lEvents.eRetStr) <> 0 And Len(lEvents.eRetStr) > 3 Then
    txtTempFiles.Text = lEvents.eRetStr
Else
    SetError "cmdChangeDir_Click", "Invalid specification", "Sorry, " & lEvents.eRetStr & " is not a valid path for temp files"
End If
If Err.Number <> 0 Then SetError "cmdChangeDir", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdChangeOutDir_Click()
'On Local Error Resume Next
Dim msg As String
frmSelectDir.Show 1
If Len(lEvents.eRetStr) <> 0 Then
    txtOutputDir.Text = lEvents.eRetStr & "\"
End If
If Err.Number <> 0 Then SetError "cmdChangeOutDir_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdDefaults_Click()
'On Local Error Resume Next
chkAutoAddTags.Value = 1
chkDownsample.Value = 1
chkDownmix.Value = 0
chkCopyrighted.Value = 0
chkOrigional.Value = 0
lstSampleRate.ListIndex = FindListboxIndex(lstSampleRate, "44100")
lstBitrate.ListIndex = FindListboxIndex(lstBitrate, "128")
txtEmailAddress.Text = "guide_X@live.com"
txtFreeDBServer.Text = "freedb.freedb.org"
chkAutoSubmitCDDB.Value = 1
chkSaveTracksToDisc.Value = 1
chkCDDBEnabled.Value = 1
chkAutoDeleteWavs.Value = 1
chkLockTrayOnRip.Value = 0
chkAutoEject.Value = 0
cboCopyMode.ListIndex = 1
lstCDDrive.ListIndex = FindListboxIndex(lstCDDrive, lDrives.dDrive(1).dLetter)
txtTempFiles.Text = App.Path & "\temp\"
chkOverwritePrompts.Value = 1
chkDisplayErrors.Value = 0
chkPlayWavs.Value = 1
chkShowAboutScreen.Value = 1
txtOutputDir.Text = App.Path & "\library\"
optAfterEncode(0).Value = True
If Err.Number <> 0 Then SetError "cmdDefaults_Click", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdDel_Click()
'On Local Error Resume Next
If Len(cboMpegPlayer.Text) <> 0 Then
    RemovePlayer cboMpegPlayer.Text
    cboMpegPlayer.RemoveItem cboMpegPlayer.ListIndex
End If
If Err.Number <> 0 Then SetError "cmdDel_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdDelCdAudioPlayer_Click()
'On Local Error Resume Next
RemovePlayer cboCDAudioPlayer.Text
If Len(cboCDAudioPlayer.Text) <> 0 Then cboCDAudioPlayer.RemoveItem cboCDAudioPlayer.ListIndex
If Err.Number <> 0 Then SetError "cmdDel_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdEncoder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Local Error Resume Next
ResetSettingsFrames eEncoder
PlayWav App.Path & "\media\click.wav", SND_ASYNC
If Err.Number <> 0 Then SetError "cmdEncoder_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdFreeDB_Click()
'On Local Error Resume Next
ResetSettingsFrames eFreeDB
PlayWav App.Path & "\media\click.wav", SND_ASYNC
If Err.Number <> 0 Then SetError "cmdFreeDB_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdGeneral_Click()
'On Local Error Resume Next
ResetSettingsFrames eGeneral
PlayWav App.Path & "\media\click.wav", SND_ASYNC
If Err.Number <> 0 Then SetError "cmdGeneral_MouseDown()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdOK_Click()
'On Local Error Resume Next
If SaveSettings = True Then Unload Me
If Err.Number <> 0 Then SetError "cmdOK_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdPlayers_Click()
'On Local Error Resume Next
ResetSettingsFrames ePlayers
PlayWav App.Path & "\media\click.wav", SND_ASYNC
If Err.Number <> 0 Then SetError "cmdPlayers_MouseDown()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdRemove_Click()
'On Local Error Resume Next
Dim i As Integer
If Len(lstCDDBServer.Text) <> 0 Then
    i = GetCDDBServerIndexByLocation(lstCDDBServer.Text)
    lCDDBServ.cServer(i).sIp = ""
    lCDDBServ.cServer(i).sLocation = ""
    lstCDDBServer.RemoveItem lstCDDBServer.ListIndex
    If i <> 0 Then
        WriteINI lIniFiles.iCDDBServers, Str(GetCDDBServerIndexByLocation(lstCDDBServer.Text)), "", ""
    End If
End If
If Err.Number <> 0 Then SetError "cmdRemove_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdRipper_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Local Error Resume Next
PlayWav App.Path & "\media\click.wav", SND_ASYNC
ResetSettingsFrames eRipper
If Err.Number <> 0 Then SetError "cmdRipper_MouseDown()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdWizard_Click()
'On Local Error Resume Next
frmSetupWizard.Show
Unload Me
If Err.Number <> 0 Then SetError "cmdWizard_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Command1_Click()
'On Local Error Resume Next
frmRegister.Show 1
If lEvents.eRegistered = True Then
    Label10.Caption = "Thanks for registering. Team Nexgen apretiates your help"
    cmdCDDB.Visible = False
    lblRegistered.Caption = "Registered Version"
    SetCaption 0
End If
If Err.Number <> 0 Then SetError "cmdCDDBWizard_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Load()
'On Local Error Resume Next
If lEvents.eRegistered = True Then
    lblRegistered.Caption = "Thanks for registering"
    cmdCDDB.Visible = False
End If
SetSettingsObjects
FlashIN frmSettings
If Err.Number <> 0 Then SetError "Form_Load()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Local Error Resume Next
FlashOut frmSettings
If Err.Number <> 0 Then SetError "Form_Unload()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub Image1_DblClick()
'On Local Error Resume Next
If Me.WindowState = vbMaximized Then
    Me.WindowState = vbNormal
Else
    Me.WindowState = vbMaximized
End If
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Local Error Resume Next
FormDrag Me
End Sub

Private Sub lstCDDBServer_Click()
'On Local Error Resume Next
Dim i As Integer
i = GetCDDBServerIndexByLocation(lstCDDBServer.Text)
If i <> 0 Then
    txtFreeDBServer.Text = lCDDBServ.cServer(i).sIp
End If
If Err.Number <> 0 Then SetError "lstCDDBServer_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub lstCDDrive_DblClick()
'On Local Error Resume Next
If lstCDDrive.Text = "<Select..>" Then
    lRipperSettings.eDriveLetter = ""
    SelectCDDrive
End If
If Err.Number <> 0 Then SetError "lstCDDrive()", lEvents.eSettings.iErrDescription, Err.Description
End Sub
