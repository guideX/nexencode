VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmSkinEditor 
   BackColor       =   &H00800000&
   Caption         =   "NexSkin"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8520
   Icon            =   "frmSkinEditor.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin VB.PictureBox picProporties 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   5565
      Left            =   7320
      ScaleHeight     =   5565
      ScaleWidth      =   1200
      TabIndex        =   6
      Top             =   315
      Width           =   1200
      Begin VB.TextBox txtProporty 
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
         Height          =   2805
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   2640
         Width           =   1215
      End
      Begin VB.ListBox lstProporties 
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
         Height          =   2460
         IntegralHeight  =   0   'False
         ItemData        =   "frmSkinEditor.frx":08CA
         Left            =   0
         List            =   "frmSkinEditor.frx":08D1
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
      BorderStyle     =   0  'None
      Height          =   5565
      Left            =   0
      ScaleHeight     =   5565
      ScaleWidth      =   1125
      TabIndex        =   0
      Top             =   315
      Width           =   1125
      Begin VB.ListBox lstValues 
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
         Height          =   1935
         IntegralHeight  =   0   'False
         Left            =   0
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.ListBox lstSkins 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1980
         IntegralHeight  =   0   'False
         ItemData        =   "frmSkinEditor.frx":08E4
         Left            =   0
         List            =   "frmSkinEditor.frx":08E6
         TabIndex        =   9
         Top             =   3240
         Width           =   1095
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
               Picture         =   "frmSkinEditor.frx":08E8
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSkinEditor.frx":0E2A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSkinEditor.frx":136C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSkinEditor.frx":18AE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSkinEditor.frx":1DF0
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSkinEditor.frx":2332
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSkinEditor.frx":2874
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSkinEditor.frx":2DB6
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSkinEditor.frx":32F8
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdDelete 
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
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
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
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   2400
         Width           =   1095
      End
      Begin VB.ComboBox cboType 
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
         ItemData        =   "frmSkinEditor.frx":383A
         Left            =   0
         List            =   "frmSkinEditor.frx":3847
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
''on local error resume next
Dim i As Integer, msg As String

lstValues.Clear
lstProporties.Clear

If frmImagePreview.Visible = True Then frmImagePreview.Visible = False

With lSkins.sSkin(lSkins.sSkinIndex)
Select Case LCase(cboType.Text)
Case "objects"
    Unload frmShapeEdit
Case "shapes"
    msg = .sFilepath & .sGraphic
    frmShapeEdit.Show
    frmSkinEditor.lstValues.AddItem "(Preview)"
    For i = 1 To lSkins.sSkin(lSkins.sSkinIndex).sShapeCount
        frmSkinEditor.lstValues.AddItem "rgn" & i
    Next i
    If DoesFileExist(msg) = False Then
        MsgBox "Unable to find the graphic " & GetFileTitle(msg), vbExclamation
        Exit Sub
    End If
    If Len(.sGraphic) <> 0 Then frmShapeEdit.Picture = LoadPicture(msg)
Case "settings"
    Unload frmShapeEdit
    lstValues.AddItem "Position"
    lstValues.AddItem "Graphics"
End Select
End With
End Sub

Private Sub cmdAdd_Click()
'on local error resume next
Dim i As Integer, msg As String, msg2 As String, x As Integer

Select Case LCase(cboType.Text)
Case "shapes"
    x = lSkins.sSkin(lSkins.sSkinIndex).sShapeCount + 1
    With lSkins.sSkin(lSkins.sSkinIndex)
        msg = "rgn" & x
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
            WriteINI .sFilepath & .sFilename, "Settings", "ShapeCount", x
            WriteINI .sFilepath & .sFilename, "Settings", "type", "3"
            .sShapeCount = x
            .sShape(x).sType = 1
            .sShape(x).sName = msg2
            .sShape(x).sEnabled = True
            .sShape(x).sRgn.X1 = 20
            .sShape(x).sRgn.Y1 = 20
            .sShape(x).sRgn.X2 = 50
            .sShape(x).sRgn.Y2 = 50
            Unload frmShapeEdit
            InitShapes
        End If
    End With
Case "objects"
Case "settings"
End Select
End Sub

Private Sub lstProporties_Click()
'on local error resume next
Dim lIndex As Integer
lIndex = lSkins.sSkinIndex
txtProporty.Text = ""

If frmImagePreview.Visible = True Then frmImagePreview.Visible = False
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
            frmImagePreview.Picture1.Picture = LoadPicture(lSkins.sSkin(lSkins.sSkinIndex).sGraphic)
        End If
    End If
    
End Select
End Sub

Private Sub lstSkins_DblClick()
Dim i As Integer

i = FindSkinIndex(lstSkins.Text)
If i <> lSkins.sSkinIndex And i <> 0 Then
    SetSkin i
    frmShapePreview.Visible = False
    frmShapeEdit.Visible = False
    frmImagePreview.Visible = False
End If
End Sub

Private Sub lstValues_Click()
''on local error resume next
Dim i As Integer, x As Integer

txtProporty.Text = ""
lstProporties.Clear
If cboType.Text = "Shapes" Then
    If lstValues.Text = "(Preview)" Then
        LoadShape frmShapePreview, lSkins.sSkinIndex
        frmShapePreview.Show
        Exit Sub
    End If
    If Left(LCase(lstValues.Text), 3) = "rgn" Then
        lstValueCount = lstValues.ListCount
        i = Int(Right(lstValues.Text, 1))
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
        For x = 1 To frmShapeEdit.shpDisplay.Count - 1
            frmShapeEdit.shpDisplay(x).BorderWidth = 1
            frmShapeEdit.shpDisplay(x).BorderColor = vbBlack
        Next x
        'frmShapeEdit.Refresh
        
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
    ElseIf lstValues.Text = "Position" Then
        lstProporties.AddItem "Width"
        lstProporties.AddItem "Height"
        lstProporties.AddItem "Top"
        lstProporties.AddItem "Left"
    End If
End If

End Sub

Private Sub MDIForm_Load()
Dim msg As String
SkinAuthor = GetSetting(App.Title, "Settings", "SkinAuthor", "")
If Len(SkinAuthor) = 0 Then
    msg = InputBox("Enter name of author:")
    SaveSetting App.Title, "Settings", "SkinAuthor", msg
End If
Me.Width = GetSetting(App.Title, "Settings", "Width", Width)
Me.Height = GetSetting(App.Title, "Settings", "Height", Height)
Me.Left = GetSetting(App.Title, "Settings", "Left", Left)
Me.Top = GetSetting(App.Title, "Settings", "Top", Top)

lstProporties.Clear
lstValues.Clear
lstSkins.Clear

SetSkin OpenSkin(GetSetting(App.Title, "Settings", "LastProject", ""), False)

End Sub

Private Sub MDIForm_Resize()
'on local error resume next
lstValues.Height = ScaleHeight / 2 - 600
cmdAdd.Top = lstValues.Height + 400
cmdDelete.Top = lstValues.Height + cmdAdd.Height + 400
'cmdApply.Top = ScaleHeight - 360
lstSkins.Top = lstValues.Height + cmdAdd.Height + cmdDelete.Height + 500
lstSkins.Height = lstValues.Height + cmdAdd.Height + cmdDelete.Height - 760

lstProporties.Height = ScaleHeight / 2
txtProporty.Top = lstProporties.Height
txtProporty.Height = ScaleHeight / 2 + 40
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Unload frmImagePreview
Unload frmShapePreview
Unload frmShapeEdit
SaveSetting App.Title, "Settings", "Width", Width
SaveSetting App.Title, "Settings", "Height", Height
SaveSetting App.Title, "Settings", "Left", Left
SaveSetting App.Title, "Settings", "Top", Top
SaveSetting App.Title, "Settings", "LastProject", lSkins.sSkin(lSkins.sSkinIndex).sFilepath & lSkins.sSkin(lSkins.sSkinIndex).sFilename
End Sub

Private Sub tlbButtons_ButtonClick(ByVal Button As MSComctlLib.Button)
'on local error resume next
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
    End If
    
End Select
End Sub

Private Sub txtProporty_Change()
'on local error resume next
Dim i As Integer, x As Integer, msg As String, msg2 As String

If Len(txtProporty.Text) = 0 Then Exit Sub
msg2 = lSkins.sSkin(lSkins.sSkinIndex).sFilepath & lSkins.sSkin(lSkins.sSkinIndex).sFilename
Select Case LCase(cboType.Text)
Case "objects"
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
    Case "top"
        lSkins.sSkin(lSkins.sSkinIndex).sSkinSettings.sTop = txtProporty.Text
        WriteINI msg2, "Settings", "Top", lSkins.sSkin(lSkins.sSkinIndex).sSkinSettings.sTop
    Case "width"
        lSkins.sSkin(lSkins.sSkinIndex).sSkinSettings.sWidth = txtProporty.Text
        WriteINI msg2, "Settings", "Width", lSkins.sSkin(lSkins.sSkinIndex).sSkinSettings.sWidth
    Case "height"
        lSkins.sSkin(lSkins.sSkinIndex).sSkinSettings.sHeight = txtProporty.Text
        WriteINI msg2, "Settings", "Left", lSkins.sSkin(lSkins.sSkinIndex).sSkinSettings.sHeight
    End Select
End Select

End Sub
