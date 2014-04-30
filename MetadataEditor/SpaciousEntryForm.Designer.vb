<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SpaciousEntryForm
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
        Me.txtText = New System.Windows.Forms.TextBox()
        Me.btnCloseDiscard = New System.Windows.Forms.Button()
        Me.btnCloseSave = New System.Windows.Forms.Button()
        Me.panelMain = New System.Windows.Forms.Panel()
        Me.llHelp = New System.Windows.Forms.LinkLabel()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.panelMain.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtText
        '
        Me.txtText.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtText.Location = New System.Drawing.Point(12, 26)
        Me.txtText.Multiline = True
        Me.txtText.Name = "txtText"
        Me.txtText.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtText.Size = New System.Drawing.Size(568, 175)
        Me.txtText.TabIndex = 0
        '
        'btnCloseDiscard
        '
        Me.btnCloseDiscard.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCloseDiscard.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCloseDiscard.Location = New System.Drawing.Point(502, 231)
        Me.btnCloseDiscard.Name = "btnCloseDiscard"
        Me.btnCloseDiscard.Size = New System.Drawing.Size(78, 25)
        Me.btnCloseDiscard.TabIndex = 25
        Me.btnCloseDiscard.Text = "Cancel"
        Me.btnCloseDiscard.UseVisualStyleBackColor = True
        '
        'btnCloseSave
        '
        Me.btnCloseSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCloseSave.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCloseSave.Location = New System.Drawing.Point(404, 231)
        Me.btnCloseSave.Name = "btnCloseSave"
        Me.btnCloseSave.Size = New System.Drawing.Size(78, 25)
        Me.btnCloseSave.TabIndex = 24
        Me.btnCloseSave.Text = "Save && Close"
        Me.btnCloseSave.UseVisualStyleBackColor = True
        '
        'panelMain
        '
        Me.panelMain.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.panelMain.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.panelMain.Controls.Add(Me.llHelp)
        Me.panelMain.Controls.Add(Me.txtText)
        Me.panelMain.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.panelMain.Location = New System.Drawing.Point(0, 0)
        Me.panelMain.Name = "panelMain"
        Me.panelMain.Size = New System.Drawing.Size(593, 215)
        Me.panelMain.TabIndex = 26
        '
        'llHelp
        '
        Me.llHelp.AutoSize = True
        Me.llHelp.Location = New System.Drawing.Point(10, 7)
        Me.llHelp.Name = "llHelp"
        Me.llHelp.Size = New System.Drawing.Size(28, 13)
        Me.llHelp.TabIndex = 27
        Me.llHelp.TabStop = True
        Me.llHelp.Text = "Help"
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSave.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSave.Location = New System.Drawing.Point(302, 231)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(78, 23)
        Me.btnSave.TabIndex = 23
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'SpaciousEntryForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(592, 268)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnCloseDiscard)
        Me.Controls.Add(Me.btnCloseSave)
        Me.Controls.Add(Me.panelMain)
        Me.MinimumSize = New System.Drawing.Size(300, 200)
        Me.Name = "SpaciousEntryForm"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "SpaciousEntryForm"
        Me.panelMain.ResumeLayout(False)
        Me.panelMain.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents txtText As System.Windows.Forms.TextBox
    Friend WithEvents btnCloseDiscard As System.Windows.Forms.Button
    Friend WithEvents btnCloseSave As System.Windows.Forms.Button
    Friend WithEvents panelMain As System.Windows.Forms.Panel
    Friend WithEvents llHelp As System.Windows.Forms.LinkLabel
    Friend WithEvents btnSave As System.Windows.Forms.Button
End Class
