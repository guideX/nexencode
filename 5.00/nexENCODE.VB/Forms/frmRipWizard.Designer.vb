<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRipWizard
    Inherits Telerik.WinControls.UI.RadForm

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim ListViewDetailColumn1 As Telerik.WinControls.UI.ListViewDetailColumn = New Telerik.WinControls.UI.ListViewDetailColumn("Column 0", "Track Number")
        Dim ListViewDetailColumn2 As Telerik.WinControls.UI.ListViewDetailColumn = New Telerik.WinControls.UI.ListViewDetailColumn("Column 1", "Title")
        Dim ListViewDetailColumn3 As Telerik.WinControls.UI.ListViewDetailColumn = New Telerik.WinControls.UI.ListViewDetailColumn("Column 2", "Duration")
        Me.pnlCurrentStep = New Telerik.WinControls.UI.RadPanel()
        Me.lblStep0 = New Telerik.WinControls.UI.RadLabel()
        Me.lblStep5 = New Telerik.WinControls.UI.RadLabel()
        Me.lblStep4 = New Telerik.WinControls.UI.RadLabel()
        Me.lblStep3 = New Telerik.WinControls.UI.RadLabel()
        Me.lblStep2 = New Telerik.WinControls.UI.RadLabel()
        Me.lblStep1 = New Telerik.WinControls.UI.RadLabel()
        Me.RadPanel1 = New Telerik.WinControls.UI.RadPanel()
        Me.cmdFinish = New Telerik.WinControls.UI.RadButton()
        Me.cmdCancel = New Telerik.WinControls.UI.RadButton()
        Me.cmdPrevious = New Telerik.WinControls.UI.RadButton()
        Me.cmdNext = New Telerik.WinControls.UI.RadButton()
        Me.RadPanel6 = New Telerik.WinControls.UI.RadPanel()
        Me.RadSeparator6 = New Telerik.WinControls.UI.RadSeparator()
        Me.pnlStep3 = New Telerik.WinControls.UI.RadPanel()
        Me.RadSeparator3 = New Telerik.WinControls.UI.RadSeparator()
        Me.RadPanel5 = New Telerik.WinControls.UI.RadPanel()
        Me.RadSeparator5 = New Telerik.WinControls.UI.RadSeparator()
        Me.pnlStep4 = New Telerik.WinControls.UI.RadPanel()
        Me.RadSeparator4 = New Telerik.WinControls.UI.RadSeparator()
        Me.pnlStep2 = New Telerik.WinControls.UI.RadPanel()
        Me.RadSeparator2 = New Telerik.WinControls.UI.RadSeparator()
        Me.RadPanel8 = New Telerik.WinControls.UI.RadPanel()
        Me.RadSeparator8 = New Telerik.WinControls.UI.RadSeparator()
        Me.pnlStep5 = New Telerik.WinControls.UI.RadPanel()
        Me.RadSeparator7 = New Telerik.WinControls.UI.RadSeparator()
        Me.pnlStep1 = New Telerik.WinControls.UI.RadPanel()
        Me.rlvSelectTracks = New Telerik.WinControls.UI.RadListView()
        Me.RadSeparator9 = New Telerik.WinControls.UI.RadSeparator()
        Me.pnlStep6 = New Telerik.WinControls.UI.RadPanel()
        CType(Me.pnlCurrentStep, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblStep0, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblStep5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblStep4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblStep3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblStep2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblStep1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadPanel1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmdFinish, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmdCancel, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmdPrevious, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmdNext, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadPanel6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadSeparator6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pnlStep3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadSeparator3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadPanel5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadSeparator5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pnlStep4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadSeparator4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pnlStep2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadSeparator2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadPanel8, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadSeparator8, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pnlStep5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadSeparator7, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pnlStep1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rlvSelectTracks, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadSeparator9, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pnlStep6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pnlCurrentStep
        '
        Me.pnlCurrentStep.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.pnlCurrentStep.Controls.Add(Me.lblStep0)
        Me.pnlCurrentStep.Controls.Add(Me.lblStep5)
        Me.pnlCurrentStep.Controls.Add(Me.lblStep4)
        Me.pnlCurrentStep.Controls.Add(Me.lblStep3)
        Me.pnlCurrentStep.Controls.Add(Me.lblStep2)
        Me.pnlCurrentStep.Controls.Add(Me.lblStep1)
        Me.pnlCurrentStep.Location = New System.Drawing.Point(0, 1)
        Me.pnlCurrentStep.Name = "pnlCurrentStep"
        '
        '
        '
        Me.pnlCurrentStep.RootElement.AccessibleDescription = Nothing
        Me.pnlCurrentStep.RootElement.AccessibleName = Nothing
        Me.pnlCurrentStep.RootElement.ControlBounds = New System.Drawing.Rectangle(0, 0, 200, 100)
        Me.pnlCurrentStep.Size = New System.Drawing.Size(137, 310)
        Me.pnlCurrentStep.TabIndex = 1
        '
        'lblStep0
        '
        Me.lblStep0.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.lblStep0.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStep0.Location = New System.Drawing.Point(12, 11)
        Me.lblStep0.Name = "lblStep0"
        '
        '
        '
        Me.lblStep0.RootElement.AccessibleDescription = Nothing
        Me.lblStep0.RootElement.AccessibleName = Nothing
        Me.lblStep0.Size = New System.Drawing.Size(83, 18)
        Me.lblStep0.TabIndex = 5
        Me.lblStep0.Text = "Rip CD Wizard"
        '
        'lblStep5
        '
        Me.lblStep5.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.lblStep5.ForeColor = System.Drawing.Color.Teal
        Me.lblStep5.Location = New System.Drawing.Point(24, 131)
        Me.lblStep5.Name = "lblStep5"
        '
        '
        '
        Me.lblStep5.RootElement.AccessibleDescription = Nothing
        Me.lblStep5.RootElement.AccessibleName = Nothing
        Me.lblStep5.RootElement.ForeColor = System.Drawing.Color.Teal
        Me.lblStep5.Size = New System.Drawing.Size(50, 18)
        Me.lblStep5.TabIndex = 4
        Me.lblStep5.Text = "Playback"
        '
        'lblStep4
        '
        Me.lblStep4.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.lblStep4.ForeColor = System.Drawing.Color.Teal
        Me.lblStep4.Location = New System.Drawing.Point(24, 107)
        Me.lblStep4.Name = "lblStep4"
        '
        '
        '
        Me.lblStep4.RootElement.AccessibleDescription = Nothing
        Me.lblStep4.RootElement.AccessibleName = Nothing
        Me.lblStep4.RootElement.ForeColor = System.Drawing.Color.Teal
        Me.lblStep4.Size = New System.Drawing.Size(49, 18)
        Me.lblStep4.TabIndex = 3
        Me.lblStep4.Text = "Progress"
        '
        'lblStep3
        '
        Me.lblStep3.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.lblStep3.ForeColor = System.Drawing.Color.Teal
        Me.lblStep3.Location = New System.Drawing.Point(24, 83)
        Me.lblStep3.Name = "lblStep3"
        '
        '
        '
        Me.lblStep3.RootElement.AccessibleDescription = Nothing
        Me.lblStep3.RootElement.AccessibleName = Nothing
        Me.lblStep3.RootElement.ForeColor = System.Drawing.Color.Teal
        Me.lblStep3.Size = New System.Drawing.Size(49, 18)
        Me.lblStep3.TabIndex = 2
        Me.lblStep3.Text = "Location"
        '
        'lblStep2
        '
        Me.lblStep2.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.lblStep2.ForeColor = System.Drawing.Color.Teal
        Me.lblStep2.Location = New System.Drawing.Point(24, 59)
        Me.lblStep2.Name = "lblStep2"
        '
        '
        '
        Me.lblStep2.RootElement.AccessibleDescription = Nothing
        Me.lblStep2.RootElement.AccessibleName = Nothing
        Me.lblStep2.RootElement.ForeColor = System.Drawing.Color.Teal
        Me.lblStep2.Size = New System.Drawing.Size(42, 18)
        Me.lblStep2.TabIndex = 1
        Me.lblStep2.Text = "Format"
        '
        'lblStep1
        '
        Me.lblStep1.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.lblStep1.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStep1.ForeColor = System.Drawing.Color.DarkSlateGray
        Me.lblStep1.Location = New System.Drawing.Point(24, 35)
        Me.lblStep1.Name = "lblStep1"
        '
        '
        '
        Me.lblStep1.RootElement.AccessibleDescription = Nothing
        Me.lblStep1.RootElement.AccessibleName = Nothing
        Me.lblStep1.RootElement.ForeColor = System.Drawing.Color.DarkSlateGray
        Me.lblStep1.Size = New System.Drawing.Size(37, 18)
        Me.lblStep1.TabIndex = 0
        Me.lblStep1.Text = "Tracks"
        '
        'RadPanel1
        '
        Me.RadPanel1.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.RadPanel1.Controls.Add(Me.cmdFinish)
        Me.RadPanel1.Controls.Add(Me.cmdCancel)
        Me.RadPanel1.Controls.Add(Me.cmdPrevious)
        Me.RadPanel1.Controls.Add(Me.cmdNext)
        Me.RadPanel1.Location = New System.Drawing.Point(140, 262)
        Me.RadPanel1.Name = "RadPanel1"
        '
        '
        '
        Me.RadPanel1.RootElement.AccessibleDescription = Nothing
        Me.RadPanel1.RootElement.AccessibleName = Nothing
        Me.RadPanel1.RootElement.ControlBounds = New System.Drawing.Rectangle(0, 0, 200, 100)
        Me.RadPanel1.Size = New System.Drawing.Size(429, 49)
        Me.RadPanel1.TabIndex = 3
        '
        'cmdFinish
        '
        Me.cmdFinish.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.cmdFinish.Location = New System.Drawing.Point(332, 13)
        Me.cmdFinish.Name = "cmdFinish"
        '
        '
        '
        Me.cmdFinish.RootElement.AccessibleDescription = Nothing
        Me.cmdFinish.RootElement.AccessibleName = Nothing
        Me.cmdFinish.RootElement.ControlBounds = New System.Drawing.Rectangle(0, 0, 130, 24)
        Me.cmdFinish.Size = New System.Drawing.Size(85, 24)
        Me.cmdFinish.TabIndex = 3
        Me.cmdFinish.Text = "Finish"
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.cmdCancel.Location = New System.Drawing.Point(13, 13)
        Me.cmdCancel.Name = "cmdCancel"
        '
        '
        '
        Me.cmdCancel.RootElement.AccessibleDescription = Nothing
        Me.cmdCancel.RootElement.AccessibleName = Nothing
        Me.cmdCancel.RootElement.ControlBounds = New System.Drawing.Rectangle(0, 0, 130, 24)
        Me.cmdCancel.Size = New System.Drawing.Size(85, 24)
        Me.cmdCancel.TabIndex = 2
        Me.cmdCancel.Text = "Cancel"
        '
        'cmdPrevious
        '
        Me.cmdPrevious.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.cmdPrevious.Location = New System.Drawing.Point(126, 13)
        Me.cmdPrevious.Name = "cmdPrevious"
        '
        '
        '
        Me.cmdPrevious.RootElement.AccessibleDescription = Nothing
        Me.cmdPrevious.RootElement.AccessibleName = Nothing
        Me.cmdPrevious.RootElement.ControlBounds = New System.Drawing.Rectangle(0, 0, 130, 24)
        Me.cmdPrevious.Size = New System.Drawing.Size(85, 24)
        Me.cmdPrevious.TabIndex = 1
        Me.cmdPrevious.Text = "<< Previous"
        '
        'cmdNext
        '
        Me.cmdNext.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.cmdNext.Location = New System.Drawing.Point(217, 13)
        Me.cmdNext.Name = "cmdNext"
        '
        '
        '
        Me.cmdNext.RootElement.AccessibleDescription = Nothing
        Me.cmdNext.RootElement.AccessibleName = Nothing
        Me.cmdNext.RootElement.ControlBounds = New System.Drawing.Rectangle(0, 0, 130, 24)
        Me.cmdNext.Size = New System.Drawing.Size(85, 24)
        Me.cmdNext.TabIndex = 0
        Me.cmdNext.Text = "Next >>"
        '
        'RadPanel6
        '
        Me.RadPanel6.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.RadPanel6.Controls.Add(Me.RadSeparator6)
        Me.RadPanel6.Location = New System.Drawing.Point(1036, 172)
        Me.RadPanel6.Name = "RadPanel6"
        '
        '
        '
        Me.RadPanel6.RootElement.AccessibleDescription = Nothing
        Me.RadPanel6.RootElement.AccessibleName = Nothing
        Me.RadPanel6.RootElement.ControlBounds = New System.Drawing.Rectangle(0, 0, 200, 100)
        Me.RadPanel6.Size = New System.Drawing.Size(429, 258)
        Me.RadPanel6.TabIndex = 13
        Me.RadPanel6.Visible = False
        '
        'RadSeparator6
        '
        Me.RadSeparator6.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.RadSeparator6.Location = New System.Drawing.Point(13, 137)
        Me.RadSeparator6.Name = "RadSeparator6"
        '
        '
        '
        Me.RadSeparator6.RootElement.AccessibleDescription = Nothing
        Me.RadSeparator6.RootElement.AccessibleName = Nothing
        Me.RadSeparator6.RootElement.ControlBounds = New System.Drawing.Rectangle(0, 0, 200, 4)
        Me.RadSeparator6.ShadowOffset = New System.Drawing.Point(0, 0)
        Me.RadSeparator6.Size = New System.Drawing.Size(404, 12)
        Me.RadSeparator6.TabIndex = 0
        Me.RadSeparator6.Text = "RadSeparator6"
        '
        'pnlStep3
        '
        Me.pnlStep3.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.pnlStep3.Controls.Add(Me.RadSeparator3)
        Me.pnlStep3.Location = New System.Drawing.Point(1052, 72)
        Me.pnlStep3.Name = "pnlStep3"
        '
        '
        '
        Me.pnlStep3.RootElement.AccessibleDescription = Nothing
        Me.pnlStep3.RootElement.AccessibleName = Nothing
        Me.pnlStep3.RootElement.ControlBounds = New System.Drawing.Rectangle(0, 0, 200, 100)
        Me.pnlStep3.Size = New System.Drawing.Size(429, 258)
        Me.pnlStep3.TabIndex = 14
        Me.pnlStep3.Visible = False
        '
        'RadSeparator3
        '
        Me.RadSeparator3.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.RadSeparator3.Location = New System.Drawing.Point(13, 137)
        Me.RadSeparator3.Name = "RadSeparator3"
        '
        '
        '
        Me.RadSeparator3.RootElement.AccessibleDescription = Nothing
        Me.RadSeparator3.RootElement.AccessibleName = Nothing
        Me.RadSeparator3.RootElement.ControlBounds = New System.Drawing.Rectangle(0, 0, 200, 4)
        Me.RadSeparator3.ShadowOffset = New System.Drawing.Point(0, 0)
        Me.RadSeparator3.Size = New System.Drawing.Size(404, 12)
        Me.RadSeparator3.TabIndex = 0
        Me.RadSeparator3.Text = "RadSeparator3"
        '
        'RadPanel5
        '
        Me.RadPanel5.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.RadPanel5.Controls.Add(Me.RadSeparator5)
        Me.RadPanel5.Location = New System.Drawing.Point(1100, 23)
        Me.RadPanel5.Name = "RadPanel5"
        '
        '
        '
        Me.RadPanel5.RootElement.AccessibleDescription = Nothing
        Me.RadPanel5.RootElement.AccessibleName = Nothing
        Me.RadPanel5.RootElement.ControlBounds = New System.Drawing.Rectangle(0, 0, 200, 100)
        Me.RadPanel5.Size = New System.Drawing.Size(429, 258)
        Me.RadPanel5.TabIndex = 15
        Me.RadPanel5.Visible = False
        '
        'RadSeparator5
        '
        Me.RadSeparator5.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.RadSeparator5.Location = New System.Drawing.Point(13, 137)
        Me.RadSeparator5.Name = "RadSeparator5"
        '
        '
        '
        Me.RadSeparator5.RootElement.AccessibleDescription = Nothing
        Me.RadSeparator5.RootElement.AccessibleName = Nothing
        Me.RadSeparator5.RootElement.ControlBounds = New System.Drawing.Rectangle(0, 0, 200, 4)
        Me.RadSeparator5.ShadowOffset = New System.Drawing.Point(0, 0)
        Me.RadSeparator5.Size = New System.Drawing.Size(404, 12)
        Me.RadSeparator5.TabIndex = 0
        Me.RadSeparator5.Text = "RadSeparator5"
        '
        'pnlStep4
        '
        Me.pnlStep4.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.pnlStep4.Controls.Add(Me.RadSeparator4)
        Me.pnlStep4.Location = New System.Drawing.Point(1108, 31)
        Me.pnlStep4.Name = "pnlStep4"
        '
        '
        '
        Me.pnlStep4.RootElement.AccessibleDescription = Nothing
        Me.pnlStep4.RootElement.AccessibleName = Nothing
        Me.pnlStep4.RootElement.ControlBounds = New System.Drawing.Rectangle(0, 0, 200, 100)
        Me.pnlStep4.Size = New System.Drawing.Size(429, 258)
        Me.pnlStep4.TabIndex = 16
        Me.pnlStep4.Visible = False
        '
        'RadSeparator4
        '
        Me.RadSeparator4.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.RadSeparator4.Location = New System.Drawing.Point(13, 137)
        Me.RadSeparator4.Name = "RadSeparator4"
        '
        '
        '
        Me.RadSeparator4.RootElement.AccessibleDescription = Nothing
        Me.RadSeparator4.RootElement.AccessibleName = Nothing
        Me.RadSeparator4.RootElement.ControlBounds = New System.Drawing.Rectangle(0, 0, 200, 4)
        Me.RadSeparator4.ShadowOffset = New System.Drawing.Point(0, 0)
        Me.RadSeparator4.Size = New System.Drawing.Size(404, 12)
        Me.RadSeparator4.TabIndex = 0
        Me.RadSeparator4.Text = "RadSeparator4"
        '
        'pnlStep2
        '
        Me.pnlStep2.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.pnlStep2.Controls.Add(Me.RadSeparator2)
        Me.pnlStep2.Location = New System.Drawing.Point(1116, 39)
        Me.pnlStep2.Name = "pnlStep2"
        '
        '
        '
        Me.pnlStep2.RootElement.AccessibleDescription = Nothing
        Me.pnlStep2.RootElement.AccessibleName = Nothing
        Me.pnlStep2.RootElement.ControlBounds = New System.Drawing.Rectangle(0, 0, 200, 100)
        Me.pnlStep2.Size = New System.Drawing.Size(429, 258)
        Me.pnlStep2.TabIndex = 17
        Me.pnlStep2.Visible = False
        '
        'RadSeparator2
        '
        Me.RadSeparator2.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.RadSeparator2.Location = New System.Drawing.Point(13, 137)
        Me.RadSeparator2.Name = "RadSeparator2"
        '
        '
        '
        Me.RadSeparator2.RootElement.AccessibleDescription = Nothing
        Me.RadSeparator2.RootElement.AccessibleName = Nothing
        Me.RadSeparator2.RootElement.ControlBounds = New System.Drawing.Rectangle(0, 0, 200, 4)
        Me.RadSeparator2.ShadowOffset = New System.Drawing.Point(0, 0)
        Me.RadSeparator2.Size = New System.Drawing.Size(404, 12)
        Me.RadSeparator2.TabIndex = 0
        Me.RadSeparator2.Text = "RadSeparator2"
        '
        'RadPanel8
        '
        Me.RadPanel8.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.RadPanel8.Controls.Add(Me.RadSeparator8)
        Me.RadPanel8.Location = New System.Drawing.Point(1124, 47)
        Me.RadPanel8.Name = "RadPanel8"
        '
        '
        '
        Me.RadPanel8.RootElement.AccessibleDescription = Nothing
        Me.RadPanel8.RootElement.AccessibleName = Nothing
        Me.RadPanel8.RootElement.ControlBounds = New System.Drawing.Rectangle(0, 0, 200, 100)
        Me.RadPanel8.Size = New System.Drawing.Size(429, 258)
        Me.RadPanel8.TabIndex = 18
        Me.RadPanel8.Visible = False
        '
        'RadSeparator8
        '
        Me.RadSeparator8.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.RadSeparator8.Location = New System.Drawing.Point(13, 137)
        Me.RadSeparator8.Name = "RadSeparator8"
        '
        '
        '
        Me.RadSeparator8.RootElement.AccessibleDescription = Nothing
        Me.RadSeparator8.RootElement.AccessibleName = Nothing
        Me.RadSeparator8.RootElement.ControlBounds = New System.Drawing.Rectangle(0, 0, 200, 4)
        Me.RadSeparator8.ShadowOffset = New System.Drawing.Point(0, 0)
        Me.RadSeparator8.Size = New System.Drawing.Size(404, 12)
        Me.RadSeparator8.TabIndex = 0
        Me.RadSeparator8.Text = "RadSeparator8"
        '
        'pnlStep5
        '
        Me.pnlStep5.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.pnlStep5.Controls.Add(Me.RadSeparator7)
        Me.pnlStep5.Location = New System.Drawing.Point(67, 327)
        Me.pnlStep5.Name = "pnlStep5"
        '
        '
        '
        Me.pnlStep5.RootElement.AccessibleDescription = Nothing
        Me.pnlStep5.RootElement.AccessibleName = Nothing
        Me.pnlStep5.RootElement.ControlBounds = New System.Drawing.Rectangle(0, 0, 200, 100)
        Me.pnlStep5.Size = New System.Drawing.Size(429, 258)
        Me.pnlStep5.TabIndex = 19
        Me.pnlStep5.Visible = False
        '
        'RadSeparator7
        '
        Me.RadSeparator7.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.RadSeparator7.Location = New System.Drawing.Point(13, 137)
        Me.RadSeparator7.Name = "RadSeparator7"
        '
        '
        '
        Me.RadSeparator7.RootElement.AccessibleDescription = Nothing
        Me.RadSeparator7.RootElement.AccessibleName = Nothing
        Me.RadSeparator7.RootElement.ControlBounds = New System.Drawing.Rectangle(0, 0, 200, 4)
        Me.RadSeparator7.ShadowOffset = New System.Drawing.Point(0, 0)
        Me.RadSeparator7.Size = New System.Drawing.Size(404, 12)
        Me.RadSeparator7.TabIndex = 0
        Me.RadSeparator7.Text = "RadSeparator7"
        '
        'pnlStep1
        '
        Me.pnlStep1.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.pnlStep1.Controls.Add(Me.rlvSelectTracks)
        Me.pnlStep1.Location = New System.Drawing.Point(140, 1)
        Me.pnlStep1.Name = "pnlStep1"
        '
        '
        '
        Me.pnlStep1.RootElement.AccessibleDescription = Nothing
        Me.pnlStep1.RootElement.AccessibleName = Nothing
        Me.pnlStep1.RootElement.ControlBounds = New System.Drawing.Rectangle(0, 0, 200, 100)
        Me.pnlStep1.Size = New System.Drawing.Size(429, 258)
        Me.pnlStep1.TabIndex = 20
        '
        'rlvSelectTracks
        '
        Me.rlvSelectTracks.BackColor = System.Drawing.SystemColors.ControlLightLight
        ListViewDetailColumn1.HeaderText = "Track Number"
        ListViewDetailColumn1.Width = 100.0!
        ListViewDetailColumn2.HeaderText = "Title"
        ListViewDetailColumn3.HeaderText = "Duration"
        ListViewDetailColumn3.Width = 100.0!
        Me.rlvSelectTracks.Columns.AddRange(New Telerik.WinControls.UI.ListViewDetailColumn() {ListViewDetailColumn1, ListViewDetailColumn2, ListViewDetailColumn3})
        Me.rlvSelectTracks.GroupItemSize = New System.Drawing.Size(200, 20)
        Me.rlvSelectTracks.ItemSize = New System.Drawing.Size(200, 20)
        Me.rlvSelectTracks.ItemSpacing = -1
        Me.rlvSelectTracks.Location = New System.Drawing.Point(3, 3)
        Me.rlvSelectTracks.Name = "rlvSelectTracks"
        '
        '
        '
        Me.rlvSelectTracks.RootElement.AccessibleDescription = Nothing
        Me.rlvSelectTracks.RootElement.AccessibleName = Nothing
        Me.rlvSelectTracks.RootElement.ControlBounds = New System.Drawing.Rectangle(0, 0, 120, 95)
        Me.rlvSelectTracks.Size = New System.Drawing.Size(423, 252)
        Me.rlvSelectTracks.TabIndex = 1
        Me.rlvSelectTracks.ViewType = Telerik.WinControls.UI.ListViewType.DetailsView
        '
        'RadSeparator9
        '
        Me.RadSeparator9.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.RadSeparator9.Location = New System.Drawing.Point(13, 137)
        Me.RadSeparator9.Name = "RadSeparator9"
        '
        '
        '
        Me.RadSeparator9.RootElement.AccessibleDescription = Nothing
        Me.RadSeparator9.RootElement.AccessibleName = Nothing
        Me.RadSeparator9.RootElement.ControlBounds = New System.Drawing.Rectangle(0, 0, 200, 4)
        Me.RadSeparator9.ShadowOffset = New System.Drawing.Point(0, 0)
        Me.RadSeparator9.Size = New System.Drawing.Size(404, 12)
        Me.RadSeparator9.TabIndex = 0
        Me.RadSeparator9.Text = "RadSeparator9"
        '
        'pnlStep6
        '
        Me.pnlStep6.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.pnlStep6.Controls.Add(Me.RadSeparator9)
        Me.pnlStep6.Location = New System.Drawing.Point(140, 1)
        Me.pnlStep6.Name = "pnlStep6"
        '
        '
        '
        Me.pnlStep6.RootElement.AccessibleDescription = Nothing
        Me.pnlStep6.RootElement.AccessibleName = Nothing
        Me.pnlStep6.RootElement.ControlBounds = New System.Drawing.Rectangle(0, 0, 200, 100)
        Me.pnlStep6.Size = New System.Drawing.Size(429, 258)
        Me.pnlStep6.TabIndex = 12
        Me.pnlStep6.Visible = False
        '
        'frmRipWizard
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(219, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(887, 506)
        Me.Controls.Add(Me.pnlStep4)
        Me.Controls.Add(Me.pnlStep5)
        Me.Controls.Add(Me.RadPanel8)
        Me.Controls.Add(Me.pnlStep2)
        Me.Controls.Add(Me.RadPanel5)
        Me.Controls.Add(Me.pnlStep3)
        Me.Controls.Add(Me.RadPanel6)
        Me.Controls.Add(Me.RadPanel1)
        Me.Controls.Add(Me.pnlCurrentStep)
        Me.Controls.Add(Me.pnlStep1)
        Me.Controls.Add(Me.pnlStep6)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "frmRipWizard"
        '
        '
        '
        Me.RootElement.ApplyShapeToControl = True
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "nexENCODE Studio - Rip CD Wizard"
        CType(Me.pnlCurrentStep, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblStep0, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblStep5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblStep4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblStep3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblStep2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblStep1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadPanel1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmdFinish, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmdCancel, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmdPrevious, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmdNext, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadPanel6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadSeparator6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pnlStep3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadSeparator3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadPanel5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadSeparator5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pnlStep4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadSeparator4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pnlStep2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadSeparator2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadPanel8, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadSeparator8, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pnlStep5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadSeparator7, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pnlStep1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rlvSelectTracks, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadSeparator9, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pnlStep6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Private WithEvents pnlCurrentStep As Telerik.WinControls.UI.RadPanel
    Private WithEvents RadPanel1 As Telerik.WinControls.UI.RadPanel
    Private WithEvents cmdNext As Telerik.WinControls.UI.RadButton
    Private WithEvents cmdFinish As Telerik.WinControls.UI.RadButton
    Private WithEvents cmdCancel As Telerik.WinControls.UI.RadButton
    Private WithEvents cmdPrevious As Telerik.WinControls.UI.RadButton
    Private WithEvents lblStep5 As Telerik.WinControls.UI.RadLabel
    Private WithEvents lblStep4 As Telerik.WinControls.UI.RadLabel
    Private WithEvents lblStep3 As Telerik.WinControls.UI.RadLabel
    Private WithEvents lblStep2 As Telerik.WinControls.UI.RadLabel
    Private WithEvents lblStep1 As Telerik.WinControls.UI.RadLabel
    Private WithEvents lblStep0 As Telerik.WinControls.UI.RadLabel
    Private WithEvents RadPanel6 As Telerik.WinControls.UI.RadPanel
    Private WithEvents RadSeparator6 As Telerik.WinControls.UI.RadSeparator
    Private WithEvents pnlStep3 As Telerik.WinControls.UI.RadPanel
    Private WithEvents RadSeparator3 As Telerik.WinControls.UI.RadSeparator
    Private WithEvents RadPanel5 As Telerik.WinControls.UI.RadPanel
    Private WithEvents RadSeparator5 As Telerik.WinControls.UI.RadSeparator
    Private WithEvents pnlStep4 As Telerik.WinControls.UI.RadPanel
    Private WithEvents RadSeparator4 As Telerik.WinControls.UI.RadSeparator
    Private WithEvents pnlStep2 As Telerik.WinControls.UI.RadPanel
    Private WithEvents RadSeparator2 As Telerik.WinControls.UI.RadSeparator
    Private WithEvents RadPanel8 As Telerik.WinControls.UI.RadPanel
    Private WithEvents RadSeparator8 As Telerik.WinControls.UI.RadSeparator
    Private WithEvents pnlStep5 As Telerik.WinControls.UI.RadPanel
    Private WithEvents RadSeparator7 As Telerik.WinControls.UI.RadSeparator
    Private WithEvents pnlStep1 As Telerik.WinControls.UI.RadPanel
    Private WithEvents RadSeparator9 As Telerik.WinControls.UI.RadSeparator
    Private WithEvents pnlStep6 As Telerik.WinControls.UI.RadPanel
    Private WithEvents rlvSelectTracks As Telerik.WinControls.UI.RadListView
End Class

