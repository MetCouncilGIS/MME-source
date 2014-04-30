<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SyncManagerForm
    Inherits System.Windows.Forms.Form

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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SyncManagerForm))
        Me.tcEPASync = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.clbSynchronizers = New System.Windows.Forms.CheckedListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.cbRetainPublishedDocID = New System.Windows.Forms.CheckBox()
        Me.cbWipeEsriTagsOnSync = New System.Windows.Forms.CheckBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cbProcSteps = New System.Windows.Forms.CheckBox()
        Me.cbExtent = New System.Windows.Forms.CheckBox()
        Me.cbCSI = New System.Windows.Forms.CheckBox()
        Me.cbAttrInfo = New System.Windows.Forms.CheckBox()
        Me.cbOnlink = New System.Windows.Forms.CheckBox()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.Help_Synchronizer_____help = New System.Windows.Forms.LinkLabel()
        Me.tcEPASync.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.SuspendLayout()
        '
        'tcEPASync
        '
        Me.tcEPASync.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tcEPASync.Controls.Add(Me.TabPage1)
        Me.tcEPASync.Controls.Add(Me.TabPage2)
        Me.tcEPASync.Location = New System.Drawing.Point(8, 12)
        Me.tcEPASync.Name = "tcEPASync"
        Me.tcEPASync.SelectedIndex = 0
        Me.tcEPASync.Size = New System.Drawing.Size(290, 182)
        Me.tcEPASync.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.clbSynchronizers)
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(282, 156)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Select Synchronizers"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'clbSynchronizers
        '
        Me.clbSynchronizers.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.clbSynchronizers.CheckOnClick = True
        Me.clbSynchronizers.FormattingEnabled = True
        Me.clbSynchronizers.Items.AddRange(New Object() {"MGMG Synchronizer", "FGDC Synchronizer"})
        Me.clbSynchronizers.Location = New System.Drawing.Point(8, 40)
        Me.clbSynchronizers.Name = "clbSynchronizers"
        Me.clbSynchronizers.Size = New System.Drawing.Size(268, 79)
        Me.clbSynchronizers.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(7, 7)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(268, 26)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Select which synchronizers to apply to your metadata records"
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.cbRetainPublishedDocID)
        Me.TabPage2.Controls.Add(Me.cbWipeEsriTagsOnSync)
        Me.TabPage2.Controls.Add(Me.Label2)
        Me.TabPage2.Controls.Add(Me.cbProcSteps)
        Me.TabPage2.Controls.Add(Me.cbExtent)
        Me.TabPage2.Controls.Add(Me.cbCSI)
        Me.TabPage2.Controls.Add(Me.cbAttrInfo)
        Me.TabPage2.Controls.Add(Me.cbOnlink)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(282, 156)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "MGMG Sync Settings"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'cbRetainPublishedDocID
        '
        Me.cbRetainPublishedDocID.AutoSize = True
        Me.cbRetainPublishedDocID.Location = New System.Drawing.Point(24, 119)
        Me.cbRetainPublishedDocID.Name = "cbRetainPublishedDocID"
        Me.cbRetainPublishedDocID.Size = New System.Drawing.Size(137, 17)
        Me.cbRetainPublishedDocID.TabIndex = 7
        Me.cbRetainPublishedDocID.Text = "Retain PublishedDocID"
        Me.cbRetainPublishedDocID.UseVisualStyleBackColor = True
        '
        'cbWipeEsriTagsOnSync
        '
        Me.cbWipeEsriTagsOnSync.AutoSize = True
        Me.cbWipeEsriTagsOnSync.Location = New System.Drawing.Point(9, 96)
        Me.cbWipeEsriTagsOnSync.Name = "cbWipeEsriTagsOnSync"
        Me.cbWipeEsriTagsOnSync.Size = New System.Drawing.Size(117, 17)
        Me.cbWipeEsriTagsOnSync.TabIndex = 6
        Me.cbWipeEsriTagsOnSync.Text = "Remove ESRI tags"
        Me.cbWipeEsriTagsOnSync.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(6, 7)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(270, 30)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Select what metadata information to update automatically using the MGMG synchroni" & _
            "zer"
        '
        'cbProcSteps
        '
        Me.cbProcSteps.AutoSize = True
        Me.cbProcSteps.Location = New System.Drawing.Point(132, 96)
        Me.cbProcSteps.Name = "cbProcSteps"
        Me.cbProcSteps.Size = New System.Drawing.Size(108, 17)
        Me.cbProcSteps.TabIndex = 4
        Me.cbProcSteps.Text = "Processing Steps"
        Me.cbProcSteps.UseVisualStyleBackColor = True
        Me.cbProcSteps.Visible = False
        '
        'cbExtent
        '
        Me.cbExtent.AutoSize = True
        Me.cbExtent.Location = New System.Drawing.Point(132, 73)
        Me.cbExtent.Name = "cbExtent"
        Me.cbExtent.Size = New System.Drawing.Size(91, 17)
        Me.cbExtent.TabIndex = 3
        Me.cbExtent.Text = "Spatial Extent"
        Me.cbExtent.UseVisualStyleBackColor = True
        '
        'cbCSI
        '
        Me.cbCSI.AutoSize = True
        Me.cbCSI.Location = New System.Drawing.Point(132, 50)
        Me.cbCSI.Name = "cbCSI"
        Me.cbCSI.Size = New System.Drawing.Size(114, 17)
        Me.cbCSI.TabIndex = 2
        Me.cbCSI.Text = "Coordinate System"
        Me.cbCSI.UseVisualStyleBackColor = True
        '
        'cbAttrInfo
        '
        Me.cbAttrInfo.AutoSize = True
        Me.cbAttrInfo.Location = New System.Drawing.Point(9, 50)
        Me.cbAttrInfo.Name = "cbAttrInfo"
        Me.cbAttrInfo.Size = New System.Drawing.Size(70, 17)
        Me.cbAttrInfo.TabIndex = 0
        Me.cbAttrInfo.Text = "Attributes"
        Me.cbAttrInfo.UseVisualStyleBackColor = True
        '
        'cbOnlink
        '
        Me.cbOnlink.AutoSize = True
        Me.cbOnlink.Location = New System.Drawing.Point(9, 73)
        Me.cbOnlink.Name = "cbOnlink"
        Me.cbOnlink.Size = New System.Drawing.Size(97, 17)
        Me.cbOnlink.TabIndex = 1
        Me.cbOnlink.Text = "Online Linkage"
        Me.cbOnlink.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSave.Location = New System.Drawing.Point(171, 200)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(56, 25)
        Me.btnSave.TabIndex = 2
        Me.btnSave.Text = "OK"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.Location = New System.Drawing.Point(244, 200)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(50, 25)
        Me.btnExit.TabIndex = 3
        Me.btnExit.Text = "Cancel"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'Help_Synchronizer_____help
        '
        Me.Help_Synchronizer_____help.AutoSize = True
        Me.Help_Synchronizer_____help.Location = New System.Drawing.Point(12, 207)
        Me.Help_Synchronizer_____help.Name = "Help_Synchronizer_____help"
        Me.Help_Synchronizer_____help.Size = New System.Drawing.Size(29, 13)
        Me.Help_Synchronizer_____help.TabIndex = 4
        Me.Help_Synchronizer_____help.TabStop = True
        Me.Help_Synchronizer_____help.Text = "Help"
        '
        'SyncManagerForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(310, 232)
        Me.Controls.Add(Me.Help_Synchronizer_____help)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.tcEPASync)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "SyncManagerForm"
        Me.Text = "Manage Metadata Synchronization"
        Me.tcEPASync.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents tcEPASync As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents cbProcSteps As System.Windows.Forms.CheckBox
    Friend WithEvents cbExtent As System.Windows.Forms.CheckBox
    Friend WithEvents cbCSI As System.Windows.Forms.CheckBox
    Friend WithEvents cbOnlink As System.Windows.Forms.CheckBox
    Friend WithEvents cbAttrInfo As System.Windows.Forms.CheckBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents clbSynchronizers As System.Windows.Forms.CheckedListBox
    Friend WithEvents cbWipeEsriTagsOnSync As System.Windows.Forms.CheckBox
    Friend WithEvents Help_Synchronizer_____help As System.Windows.Forms.LinkLabel
    Friend WithEvents cbRetainPublishedDocID As System.Windows.Forms.CheckBox
End Class
