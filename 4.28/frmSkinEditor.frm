VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmSkinEditor 
   BackColor       =   &H8000000F&
   Caption         =   "NexSkin"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8520
   Icon            =   "frmSkinEditor.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picProporties 
      Align           =   4  'Align Right
      BackColor       =   &H00000000&
      Height          =   5565
      Left            =   7320
      ScaleHeight     =   5505
      ScaleWidth      =   1140
      TabIndex        =   6
      Top             =   315
      Width           =   1200
      Begin VB.TextBox txtProporty 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2805
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   2520
         Width           =   1215
      End
      Begin VB.ListBox lstProporties 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2460
         IntegralHeight  =   0   'False
         ItemData        =   "frmSkinEditor.frx":000C
         Left            =   0
         List            =   "frmSkinEditor.frx":0013
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   0
         Width           =   1215
      End
   End
   Begin MSComctlLib.Toolbar tlbButtons 
      Align           =   1  'Align Top
      Height          =   315
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   556
      ButtonWidth     =   609
      ButtonHeight    =   556
      ToolTips        =   0   'False
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "imgList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picToolbox 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   5565
      Left            =   0
      ScaleHeight     =   5535
      ScaleWidth      =   1095
      TabIndex        =   0
      Top             =   315
      Width           =   1125
      Begin VB.ListBox lstValues 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1900
         IntegralHeight  =   0   'False
         Left            =   0
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   320
         Width           =   1120
      End
      Begin MSComctlLib.ImageList imgList 
         Left            =   360
         Top             =   1200
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   15
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   9
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSkinEditor.frx":0026
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSkinEditor.frx":0568
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSkinEditor.frx":0AAA
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSkinEditor.frx":0FEC
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSkinEditor.frx":152E
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSkinEditor.frx":1A70
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSkinEditor.frx":1FB2
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSkinEditor.frx":24F4
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSkinEditor.frx":2A36
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         TabIndex        =   4
         Top             =   2760
         Width           =   1100
      End
      Begin VB.CommandButton cmdAdd 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         TabIndex        =   3
         Top             =   2400
         Width           =   1100
      End
      Begin VB.ComboBox cboType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Height          =   315
         ItemData        =   "frmSkinEditor.frx":2F78
         Left            =   0
         List            =   "frmSkinEditor.frx":2F85
         TabIndex        =   1
         Text            =   "(Select)"
         Top             =   0
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmSkinEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lstValueIndex As Integer
Public lstValueCount As Integer

Private Sub cboType_Click()
'On Local Error Resume Next
Dim i As Integer, msg As String

lstValues.Clear
lstProporties.Clear

If frmImagePreview.Visible = True Then frmImagePreview.Visible = False

With lSkins.sSkin(lSkins.sSkinIndex)
Select Case LCase(cboType.Text)
Case "objects"
    Unload frmShapeEdit
    For i = 1 To .sObjectCount
        lstValues.AddItem "object" & i
    Next i
Case "shapes"
    msg = .sFilepath & .sGraphic
    frmShapeEdit.Show
    frmSkinEditor.lstValues.AddItem "(Preview)"
    For i = 1 To lSkins.sSkin(lSkins.sSkinIndex).sShapeCount
        frmSkinEditor.lstValues.AddItem "rgn" & i
    Next i
    If DoesFileExist(msg) = False Then Exit Sub
    If Len(.sGraphic) <> 0 Then frmShapeEdit.Picture = LoadPicture(msg)
Case "settings"
    Unload frmShapeEdit
    lstValues.AddItem "Position"
    lstValues.AddItem "Graphics"
End Select
End With
If Err.Number <> 0 Then SetError "cboType_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdAdd_Click()
'On Local Error Resume Next
Dim i As Integer, msg As String, msg2 As String, X As Integer

Select Case LCase(cboType.Text)
Case "shapes"
    X = lSkins.sSkin(lSkins.sSkinIndex).sShapeCount + 1
    With lSkins.sSkin(lSkins.sSkinIndex)
        msg = "rgn" & X
        For i = 0 To lstValues.ListCount
            If msg = lstValues.List(i) Then Exit Sub
        Next i
        msg2 = InputBox("Enter description of shape:")
        If Len(msg2) <> 0 Then
            lstValues.AddItem msg
            WriteINI .sFilepath & .sFilename, msg, "enabled", "true"
            WriteINI .sFilepath & .sFilename, msg, "name", msg2
            WriteINI .sFilepath & .sFilename, msg, "x1", "20"
            WriteINI .sFilepath & .sFilename, msg, "y1", "20"
            WriteINI .sFilepath & .sFilename, msg, "x2", "50"
            WriteINI .sFilepath & .sFilename, msg, "y2", "50"
            WriteINI .sFilepath & .sFilename, "Settings", "ShapeCount", X
            WriteINI .sFilepath & .sFilename, "Settings", "type", "3"
            .sShapeCount = X
            .sShape(X).sType = 1
            .sShape(X).sName = msg2
            .sShape(X).sEnabled = True
            .sShape(X).sRgn.X1 = 20
            .sShape(X).sRgn.Y1 = 20
            .sShape(X).sRgn.X2 = 50
            .sShape(X).sRgn.Y2 = 50
            Unload frmShapeEdit
            InitShapes
            frmShapeEdit.Picture = LoadPicture(lSkins.sSkin(lSkins.sSkinIndex).sGraphic)
        End If
    End With
Case "objects"
    X = lSkins.sSkin(lSkins.sSkinIndex).sObjectCount + 1
    msg = "object" & X
    msg2 = InputBox("Enter description of object")
        If Len(msg2) <> 0 Then
            With lSkins.sSkin(lSkins.sSkinIndex)
                lstValues.AddItem msg
                WriteINI .sFilepath & .sFilename, msg, "enabled", "true"
                WriteINI .sFilepath & .sFilename, msg, "name", msg2
                WriteINI .sFilepath & .sFilename, msg, "type", "0"
                WriteINI .sFilepath & .sFilename, msg, "left", "20"
                WriteINI .sFilepath & .sFilename, msg, "filename", ""
                WriteINI .sFilepath & .sFilename, msg, "filename2", ""
                WriteINI .sFilepath & .sFilename, "Settings", "ObjectCount", X
            End With
        End If
Case "settings"
End Select
If Err.Number <> 0 Then SetError "cmdAdd_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub cmdDelete_Click()
'On Local Error Resume Next
Dim i As Integer, msg As String

Select Case LCase(cboType.Text)
Case "shapes"
    msg = lstValues.Text
    i = Int(Right(lstValues.Text, Len(lstValues.Text) - 3))
    With lSkins.sSkin(lSkins.sSkinIndex)
        WriteINI .sFilepath & .sFilename, msg, "enabled", ""
        WriteINI .sFilepath & .sFilename, msg, "name", ""
        WriteINI .sFilepath & .sFilename, msg, "x1", ""
        WriteINI .sFilepath & .sFilename, msg, "y1", ""
        WriteINI .sFilepath & .sFilename, msg, "x2", ""
        WriteINI .sFilepath & .sFilename, msg, "y2", ""
        .sShapeCount = .sShapeCount - 1
        WriteINI .sFilepath & .sFilename, "Settings", "ShapeCount", .sShapeCount
        .sShape(i).sEnabled = False
    End With
    Unload frmShapeEdit
    InitShapes
    frmShapeEdit.Picture = LoadPicture(lSkins.sSkin(lSkins.sSkinIndex).sGraphic)
    lstValues.Text = msg
    If lstValues.Text = msg Then lstValues.RemoveItem lstValues.ListIndex
End Select

If Err.Number <> 0 Then SetError "cmdDelete_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub lstProporties_Click()
'On Local Error Resume Next
Dim lIndex As Integer

lIndex = lSkins.sSkinIndex
txtProporty.Text = ""

Select Case LCase(cboType.Text)
Case "shapes"
    With lSkins.sSkin(lIndex).sShape(lstValueIndex)
        Select Case LCase(lstProporties.Text)
        Case "name"
            txtProporty.Text = .sName
        Case "x1"
            txtProporty.Text = .sRgn.X1
        Case "x2"
            txtProporty.Text = .sRgn.X2
        Case "x3"
            txtProporty.Text = .sRgn.X3
        Case "y1"
            txtProporty.Text = .sRgn.Y1
        Case "y2"
            txtProporty.Text = .sRgn.Y2
        Case "y3"
            txtProporty.Text = .sRgn.Y3
        Case "type"
            txtProporty.Text = .sType
        Case "combine"
            txtProporty.Text = .sCombineMode
        Case "srcrgn1"
            txtProporty.Text = .sSrcRgn1
        Case "srcrgn2"
            txtProporty.Text = .sSrcRgn2
        Case "destrgn"
            txtProporty.Text = .sDestRgn
        End Select
    End With
Case "objects"
    With lSkins.sSkin(lSkins.sSkinIndex).sObject(lstValueIndex)
        Select Case LCase(lstProporties.Text)
        Case "left"
            txtProporty.Text = .oPos.sLeft
        Case "top"
            txtProporty.Text = .oPos.sTop
        Case "width"
            txtProporty.Text = .oPos.sWidth
        Case "height"
            txtProporty.Text = .oPos.sHeight
        Case "type"
            txtProporty.Text = .oType
        Case "enabled"
            txtProporty.Text = .oEnabled
        Case "filename"
            If Len(.oFilename) <> 0 Then
                txtProporty.Text = .oFilename
                frmImagePreview.Show
                frmImagePreview.Picture1.Picture = LoadPicture(lSkins.sSkin(lSkins.sSkinIndex).sFilepath & lSkins.sSkin(lSkins.sSkinIndex).sObject(lstValueIndex).oFilename)
            End If
        Case "filename2"
            txtProporty.Text = .oFilename2
            If Len(.oFilename2) <> 0 Then
                txtProporty.Text = .oFilename2
                frmImagePreview.Show
                frmImagePreview.Picture1.Picture = LoadPicture(lSkins.sSkin(lSkins.sSkinIndex).sFilepath & lSkins.sSkin(lSkins.sSkinIndex).sObject(lstValueIndex).oFilename2)
            End If
        Case "filename3"
            txtProporty.Text = .oFilename3
            If Len(.oFilename3) <> 0 Then
                txtProporty.Text = .oFilename3
                frmImagePreview.Show
                frmImagePreview.Picture1.Picture = LoadPicture(lSkins.sSkin(lSkins.sSkinIndex).sFilepath & lSkins.sSkin(lSkins.sSkinIndex).sObject(lstValueIndex).oFilename3)
            End If
        
        End Select
    End With
Case "settings"
    If lstValues.Text = "Position" Then
        Select Case LCase(lstProporties.Text)
        Case "left"
            txtProporty.Text = lSkins.sSkin(lSkins.sSkinIndex).sSkinSettings.sLeft
        Case "top"
            txtProporty.Text = lSkins.sSkin(lSkins.sSkinIndex).sSkinSettings.sTop
        Case "width"
            txtProporty.Text = lSkins.sSkin(lSkins.sSkinIndex).sSkinSettings.sWidth
        Case "height"
            txtProporty.Text = lSkins.sSkin(lSkins.sSkinIndex).sSkinSettings.sHeight
        End Select
    ElseIf lstValues.Text = "Graphics" Then
        If lstProporties.Text = "Main" Then
            txtProporty.Text = lSkins.sSkin(lSkins.sSkinIndex).sGraphic
            frmImagePreview.Show
            frmImagePreview.Picture1.Picture = LoadPicture(lSkins.sSkin(lSkins.sSkinIndex).sFilepath & lSkins.sSkin(lSkins.sSkinIndex).sGraphic)
        ElseIf lstProporties.Text = "Main2" Then
            txtProporty.Text = lSkins.sSkin(lSkins.sSkinIndex).sBackground
            frmImagePreview.Show
            frmImagePreview.Picture1.Picture = LoadPicture(lSkins.sSkin(lSkins.sSkinIndex).sFilepath & lSkins.sSkin(lSkins.sSkinIndex).sBackground)
        ElseIf lstProporties.Text = "Main3" Then
            txtProporty.Text = lSkins.sSkin(lSkins.sSkinIndex).sErrorGraphic
            frmImagePreview.Show
            frmImagePreview.Picture1.Picture = LoadPicture(lSkins.sSkin(lSkins.sSkinIndex).sFilepath & lSkins.sSkin(lSkins.sSkinIndex).sErrorGraphic)
        'ElseIf lstProporties.Text = "Side" Then
            'txtProporty.Text = lSkins.sSkin(lSkins.sSkinIndex).sSideGradient
        '    frmImagePreview.Show
        '    frmImagePreview.Picture1.Picture = LoadPicture(lSkins.sSkin(lSkins.sSkinIndex).sFilepath & lSkins.sSkin(lSkins.sSkinIndex).sSideGradient)
        'ElseIf lstProporties.Text = "Top" Then
        '    txtProporty.Text = lSkins.sSkin(lSkins.sSkinIndex).sToper
        '    frmImagePreview.Show
        '    frmImagePreview.Picture1.Picture = LoadPicture(lSkins.sSkin(lSkins.sSkinIndex).sFilepath & lSkins.sSkin(lSkins.sSkinIndex).sToper)
        ElseIf lstProporties.Text = "Playlist" Then
            txtProporty.Text = lSkins.sSkin(lSkins.sSkinIndex).sPlaylistGraphic
            frmImagePreview.Show
            frmImagePreview.Picture1.Picture = LoadPicture(lSkins.sSkin(lSkins.sSkinIndex).sFilepath & lSkins.sSkin(lSkins.sSkinIndex).sPlaylistGraphic)
        End If
    End If
End Select
If Err.Number <> 0 Then SetError "lstProporties_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub lstValues_Click()
'On Local Error Resume Next
Dim i As Integer, X As Integer

txtProporty.Text = ""
lstProporties.Clear
If cboType.Text = "Shapes" Then
    If frmImagePreview.Visible = True Then Unload frmImagePreview
    If lstValues.Text = "(Preview)" Then
        LoadShape frmShapePreview, lSkins.sSkinIndex
        frmShapePreview.Show
        Exit Sub
    End If
    If Left(LCase(lstValues.Text), 3) = "rgn" Then
        lstValueCount = lstValues.ListCount
        i = Int(Right(lstValues.Text, Len(lstValues.Text) - 3))
        lstValueIndex = i
    Else
        lstValueIndex = 0
    End If
    With lSkins.sSkin(lSkins.sSkinIndex)
    Select Case .sShape(lstValueIndex).sType
        Case 2
            lstProporties.AddItem "Name"
            lstProporties.AddItem "x1"
            lstProporties.AddItem "x2"
            lstProporties.AddItem "y1"
            lstProporties.AddItem "y2"
            lstProporties.AddItem "Type"
            lstProporties.AddItem "Combine"
            lstProporties.AddItem "DestRgn"
            lstProporties.AddItem "SrcRgn1"
            lstProporties.AddItem "SrcRgn2"
        Case 1
            lstProporties.AddItem "Name"
            lstProporties.AddItem "x1"
            lstProporties.AddItem "x2"
            lstProporties.AddItem "y1"
            lstProporties.AddItem "y2"
            lstProporties.AddItem "Type"
            lstProporties.AddItem "Combine"
            lstProporties.AddItem "DestRgn"
            lstProporties.AddItem "SrcRgn1"
            lstProporties.AddItem "SrcRgn2"
        Case 3
            lstProporties.AddItem "Name"
            lstProporties.AddItem "x1"
            lstProporties.AddItem "x2"
            lstProporties.AddItem "x3"
            lstProporties.AddItem "y1"
            lstProporties.AddItem "y2"
            lstProporties.AddItem "y3"
            lstProporties.AddItem "Type"
            lstProporties.AddItem "Combine"
            lstProporties.AddItem "DestRgn"
            lstProporties.AddItem "SrcRgn1"
            lstProporties.AddItem "SrcRgn2"
        End Select
        For X = 1 To frmShapeEdit.shpDisplay.Count - 1
            frmShapeEdit.shpDisplay(X).BorderWidth = 1
            frmShapeEdit.shpDisplay(X).BorderColor = vbWhite
        Next X
        frmShapeEdit.shpDisplay(lstValueIndex).BorderWidth = 3
        frmShapeEdit.shpDisplay(lstValueIndex).BorderColor = &HC0C0&
        frmShapeEdit.Caption = "Name: " & lSkins.sSkin(lSkins.sSkinIndex).sShape(lstValueIndex).sName & " , Shape #:" & lstValueIndex
        RefreshAllShapes
        DoEvents
    End With
End If
If cboType.Text = "Settings" Then
    If lstValues.Text = "Graphics" Then
        lstProporties.AddItem "Main"
        lstProporties.AddItem "Main2"
        lstProporties.AddItem "Main3"
        lstProporties.AddItem "Playlist"
        'lstProporties.AddItem "Top"
        'lstProporties.AddItem "Side"
        
    ElseIf lstValues.Text = "Position" Then
        lstProporties.AddItem "Width"
        lstProporties.AddItem "Height"
        lstProporties.AddItem "Top"
        lstProporties.AddItem "Left"
    End If
End If
If cboType.Text = "Objects" Then
    If frmImagePreview.Visible = True Then Unload frmImagePreview
    If Left(LCase(lstValues.Text), 6) = "object" Then
        lstValueCount = lstValues.ListCount
        i = Int(Right(lstValues.Text, Len(lstValues.Text) - 6))
        lstValueIndex = i
        lstProporties.Clear
        lstProporties.AddItem "Enabled"
        lstProporties.AddItem "Filename"
        lstProporties.AddItem "Filename2"
        lstProporties.AddItem "Filename3"
        lstProporties.AddItem "Name"
        lstProporties.AddItem "Width"
        lstProporties.AddItem "Height"
        lstProporties.AddItem "Left"
        lstProporties.AddItem "Top"
        lstProporties.AddItem "Type"
    Else
        lstValueIndex = 0
    End If
End If
If Err.Number <> 0 Then SetError "lstValues_Click()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub MDIForm_Load()
'On Local Error Resume Next
Dim msg As String
Icon = frmGraphics.Icon
frmMain.mnuNexSkin.Checked = True
SkinAuthor = ReadINI(lIniFiles.iSettings, "Settings", "SkinAuthor", "|GuideX|")
If Len(SkinAuthor) = 0 Then
    msg = InputBox("Enter name of author:")
    WriteINI lIniFiles.iSettings, "Settings", "SkinAuthor", msg
End If
lstProporties.Clear
lstValues.Clear
FlashIN frmSkinEditor
MDIForm_Resize
If Err.Number <> 0 Then SetError "frmSkinEditor_Load()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub MDIForm_Resize()
'On Local Error Resume Next

lstValues.Height = ScaleHeight - 1100
cmdAdd.Top = lstValues.Height + 400
cmdDelete.Top = lstValues.Height + cmdAdd.Height + 400

lstProporties.Height = ScaleHeight / 2
txtProporty.Top = lstProporties.Height
txtProporty.Height = ScaleHeight / 2 + 40
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
'On Local Error Resume Next

frmMain.mnuNexSkin.Checked = False
FlashOut frmSkinEditor
Me.Visible = False
Unload frmImagePreview
Unload frmShapePreview
Unload frmShapeEdit
End Sub

Private Sub tlbButtons_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Local Error Resume Next
Dim msg As String, i As Integer, msg2 As String

Select Case Button.Index
Case 1
    msg = NewSkin
    SetSkin OpenSkin(msg, False)
Case 2
    Unload frmShapeEdit
    msg = OpenDialog(frmSkinEditor, "Storage Containers (*.ns4)|*.ns4|All Files (*.*)|*.*", "Select Skin Container", App.Path)
    If Len(msg) <> 0 Then
        WriteINI msg, "Settings", "Enabled", "True"
        i = OpenSkin(msg, False)
        SetSkin i
        cboType.ListIndex = 1
    End If
End Select
If Err.Number <> 0 Then SetError "tlbButtons()", lEvents.eSettings.iErrDescription, Err.Description
End Sub

Private Sub txtProporty_Change()
'On Local Error Resume Next
Dim i As Integer, X As Integer, msg As String, msg2 As String

If Len(txtProporty.Text) = 0 Then Exit Sub
msg2 = lSkins.sSkin(lSkins.sSkinIndex).sFilepath & lSkins.sSkin(lSkins.sSkinIndex).sFilename
Select Case LCase(cboType.Text)
Case "objects"
    With lSkins.sSkin(lSkins.sSkinIndex).sObject(lstValueIndex)
    If lstValueIndex <> 0 Then
        Select Case LCase(lstProporties.Text)
        Case "filename"
            If .oFilename <> txtProporty.Text Then
                .oFilename = txtProporty.Text
                WriteINI msg2, lstValues.Text, "filename", .oFilename
            End If
        Case "filename2"
            If .oFilename2 <> txtProporty.Text Then
                .oFilename2 = txtProporty.Text
                WriteINI msg2, lstValues.Text, "filename2", .oFilename2
            End If
        Case "filename3"
            If .oFilename3 <> txtProporty.Text Then
                .oFilename3 = txtProporty.Text
                WriteINI msg2, lstValues.Text, "filename3", .oFilename3
            End If
        Case "enabled"
            If .oEnabled <> txtProporty.Text Then
                .oEnabled = txtProporty.Text
                WriteINI msg2, lstValues.Text, "enabled", .oEnabled
            End If
        Case "name"
            If .oName <> txtProporty.Text Then
                .oName = txtProporty.Text
                WriteINI msg2, lstValues.Text, "name", .oName
            End If
        Case "type"
            If .oType <> txtProporty.Text Then
                .oType = txtProporty.Text
                WriteINI msg2, lstValues.Text, "type", .oType
            End If
        Case "left"
            If .oPos.sLeft <> txtProporty.Text Then
                .oPos.sLeft = txtProporty.Text
                WriteINI msg2, lstValues.Text, "left", .oPos.sLeft
            End If
        Case "top"
            If .oPos.sTop <> txtProporty.Text Then
                .oPos.sTop = txtProporty.Text
                WriteINI msg2, lstValues.Text, "top", .oPos.sTop
            End If
        Case "width"
            If .oPos.sWidth <> txtProporty.Text Then
                .oPos.sWidth = txtProporty.Text
                WriteINI msg2, lstValues.Text, "width", .oPos.sWidth
            End If
        Case "height"
            If .oPos.sHeight <> txtProporty.Text Then
                .oPos.sHeight = txtProporty.Text
                WriteINI msg2, lstValues.Text, "height", .oPos.sHeight
            End If
        End Select
    End If
    End With
Case "shapes"
    With lSkins.sSkin(lSkins.sSkinIndex).sShape(lstValueIndex)
    If lstValueIndex <> 0 Then
        Select Case LCase(lstProporties.Text)
        Case "x1"
            If .sRgn.X1 <> txtProporty.Text Then
                .sRgn.X1 = txtProporty.Text
                WriteINI msg2, lstValues.Text, "x1", .sRgn.X1
            End If
        Case "x2"
            If .sRgn.X2 <> txtProporty.Text Then
                .sRgn.X2 = txtProporty.Text
                WriteINI msg2, lstValues.Text, "x2", .sRgn.X2
            End If
        Case "x3"
            If .sRgn.X3 <> txtProporty.Text Then
                .sRgn.X3 = txtProporty.Text
                WriteINI msg2, lstValues.Text, "x3", .sRgn.X3
            End If
        Case "y1"
            If .sRgn.Y1 <> txtProporty.Text Then
                .sRgn.Y1 = txtProporty.Text
                WriteINI msg2, lstValues.Text, "y1", .sRgn.Y1
            End If
        Case "y2"
            If .sRgn.Y2 <> txtProporty.Text Then
                .sRgn.Y2 = txtProporty.Text
                WriteINI msg2, lstValues.Text, "y2", .sRgn.Y2
            End If
        Case "y3"
            If .sRgn.Y3 <> txtProporty.Text Then
                .sRgn.Y3 = txtProporty.Text
                WriteINI msg2, lstValues.Text, "y3", .sRgn.Y3
            End If
        Case "destrgn"
            If .sDestRgn <> txtProporty.Text Then
                .sDestRgn = txtProporty.Text
                WriteINI msg2, lstValues.Text, "destrgn", .sDestRgn
            End If
        Case "srcrgn1"
            If .sSrcRgn1 <> txtProporty.Text Then
                .sSrcRgn1 = txtProporty.Text
                WriteINI msg2, lstValues.Text, "srcrgn1", .sSrcRgn1
            End If
        Case "srcrgn2"
            If .sSrcRgn2 <> txtProporty.Text Then
                .sSrcRgn2 = txtProporty.Text
                WriteINI msg2, lstValues.Text, "srcrgn2", .sSrcRgn2
            End If
        Case "combine"
            .sCombineMode = txtProporty.Text
            WriteINI msg2, lstValues.Text, "combinemode", .sCombineMode
        Case "type"
            .sType = txtProporty.Text
            WriteINI msg2, lstValues.Text, "type", .sType
        End Select
    End If
    RefreshAllShapes
    End With
Case "settings"
    Select Case LCase(lstProporties.Text)
    Case "left"
        lSkins.sSkin(lSkins.sSkinIndex).sSkinSettings.sLeft = txtProporty.Text
        WriteINI msg2, "Settings", "Left", lSkins.sSkin(lSkins.sSkinIndex).sSkinSettings.sLeft
    Case "width"
        lSkins.sSkin(lSkins.sSkinIndex).sSkinSettings.sWidth = txtProporty.Text
        WriteINI msg2, "Settings", "Width", lSkins.sSkin(lSkins.sSkinIndex).sSkinSettings.sWidth
    Case "height"
        lSkins.sSkin(lSkins.sSkinIndex).sSkinSettings.sHeight = txtProporty.Text
        WriteINI msg2, "Settings", "Left", lSkins.sSkin(lSkins.sSkinIndex).sSkinSettings.sHeight
    Case "playlist"
        lSkins.sSkin(lSkins.sSkinIndex).sPlaylistGraphic = txtProporty.Text
        WriteINI msg2, "Settings", "Playlist", lSkins.sSkin(lSkins.sSkinIndex).sPlaylistGraphic
    Case "main"
        lSkins.sSkin(lSkins.sSkinIndex).sGraphic = txtProporty.Text
        WriteINI msg2, "Settings", "Graphic", lSkins.sSkin(lSkins.sSkinIndex).sGraphic
    Case "main2"
        lSkins.sSkin(lSkins.sSkinIndex).sBackground = txtProporty.Text
        WriteINI msg2, "Settings", "Background", lSkins.sSkin(lSkins.sSkinIndex).sBackground
    Case "main3"
        lSkins.sSkin(lSkins.sSkinIndex).sErrorGraphic = txtProporty.Text
        WriteINI msg2, "Settings", "ErrorGraphic", lSkins.sSkin(lSkins.sSkinIndex).sErrorGraphic
    'Case "side"
    '    lSkins.sSkin(lSkins.sSkinIndex).sSideGradient = txtProporty.Text
    '    WriteINI msg2, "Settings", "Side", lSkins.sSkin(lSkins.sSkinIndex).sSideGradient
    'Case "top"
    '    If lstValues.Text = "Graphics" Then
    '        lSkins.sSkin(lSkins.sSkinIndex).sToper = txtProporty.Text
    '        WriteINI msg2, "Settings", "Topper", lSkins.sSkin(lSkins.sSkinIndex).sToper
    '    Else
    '        lSkins.sSkin(lSkins.sSkinIndex).sSkinSettings.sTop = txtProporty.Text
    '        WriteINI msg2, "Settings", "Top", lSkins.sSkin(lSkins.sSkinIndex).sSkinSettings.sTop
    '    End If
    End Select
End Select
If Err.Number <> 0 Then SetError "txtProporty_Change()", lEvents.eSettings.iErrDescription, Err.Description
End Sub
