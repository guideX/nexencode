VERSION 5.00
Object = "{C0DD72E3-52F1-11D2-A800-0000E8545063}#1.0#0"; "ACM.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7170
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   10950
   StartUpPosition =   3  'Windows Default
   Begin ACMLib.ACM ACM1 
      Height          =   615
      Left            =   480
      TabIndex        =   11
      Top             =   2280
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   1085
      _StockProps     =   0
   End
   Begin VB.Frame Frame3 
      Caption         =   "Error"
      Height          =   735
      Left            =   9720
      TabIndex        =   9
      Top             =   120
      Width           =   1095
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "PCM -> CELP"
      Height          =   495
      Left            =   1080
      TabIndex        =   8
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Compressed"
      Height          =   735
      Left            =   6720
      TabIndex        =   6
      Top             =   120
      Width           =   2775
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Text            =   "c:\song.mp3"
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Source"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3975
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Text            =   "c:\song8.wav"
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Stop"
      Height          =   495
      Left            =   4680
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.ListBox List3 
      Height          =   5715
      Left            =   7200
      TabIndex        =   2
      Top             =   1200
      Width           =   3495
   End
   Begin VB.ListBox List2 
      Height          =   5715
      Left            =   4320
      TabIndex        =   1
      Top             =   1200
      Width           =   2775
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   120
      TabIndex        =   0
      Top             =   3960
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ACM_ConvertPosition(ByVal ConvertPosition As Long)
Form1.Caption = ConvertPosition
End Sub

Private Sub ACM1_ConvertPosition(ByVal ConvertPosition As Long)
Form1.Caption = ConvertPosition
End Sub

Private Sub ACM1_EndOfConversion()
text3text = ""
End Sub

Private Sub Command1_Click()
ACM1.Stop
End Sub

Private Sub Command2_Click()

Dim err

err = ACM1.SetInputFile(Text2.Text) 'PCM Format
err = ACM1.SetOutputFile("c:\temp.wav")


err = ACM1.SelectDriverByFileName("msadp32")
err = ACM1.SelectDriverByFileName("ms-pcm")

err = ACM1.SetFormatTags(1)
err = ACM1.SetFormat(3) ' from PCM to PCM 8k/16/m
err = ACM1.Convert

While ACM1.InConvert
    DoEvents
Wend

err = ACM1.SetInputFile("c:\temp.wav") ' PCM Format
err = ACM1.SetOutputFile(Text1.Text)

err = ACM1.SelectDriverByFileName("lhacm")

err = ACM1.SetFormatTags(1)
err = ACM1.SetFormat(1) ' from PCM to L&H Celp 4,8
err = ACM1.Convert

While ACM1.InConvert
    DoEvents
Wend

Kill "c:\temp.wav"

End Sub

Private Sub Command3_Click()


End Sub

Private Sub Form_Load()
Dim err

'err = ACM1.Authorize("xxxxxxxxxxxxxx", "xxxxxxxxxxxx")

For i = 1 To ACM1.GetNumACMs

   List1.AddItem (ACM1.GetNameACM(i))

Next i

End Sub

Private Sub List1_DblClick()
ACM1.SetInputFile (Text2.Text)
ACM1.SetOutputFile (Text1.Text)

ACM1.SelectACMDriver (List1.ListIndex + 1)

While List2.ListCount > 0

List2.RemoveItem (0)
DoEvents
List2.Refresh

Wend

For i = 1 To ACM1.GetFormatTagsCount
   List2.AddItem (ACM1.GetFormatTags(i))
Next i

While List3.ListCount > 0

List3.RemoveItem (0)
DoEvents
List3.Refresh

Wend

End Sub

Private Sub List2_DblClick()
ACM1.SetFormatTags (List2.ListIndex + 1)

While List3.ListCount > 0

List3.RemoveItem (0)
DoEvents
List3.Refresh

Wend

For i = 1 To ACM1.GetFormatsCount

   List3.AddItem (ACM1.GetFormat(i))

Next i

End Sub

Private Sub List3_DblClick()

Dim err

err = ACM1.SetFormat(List3.ListIndex + 1)
Text3.Text = "Error: " & err
err = ACM1.Convert

While ACM1.InConvert
    DoEvents
Wend

End Sub

