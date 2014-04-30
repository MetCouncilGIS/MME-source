<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FindReplaceForm
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FindReplaceForm))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnReplaceAll = New System.Windows.Forms.Button()
        Me.tbReplace = New System.Windows.Forms.TextBox()
        Me.tbFind = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(19, 23)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Find what:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(19, 49)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(72, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Replace with:"
        '
        'btnReplaceAll
        '
        Me.btnReplaceAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnReplaceAll.Location = New System.Drawing.Point(237, 84)
        Me.btnReplaceAll.Name = "btnReplaceAll"
        Me.btnReplaceAll.Size = New System.Drawing.Size(75, 23)
        Me.btnReplaceAll.TabIndex = 4
        Me.btnReplaceAll.Text = "Replace All"
        Me.btnReplaceAll.UseVisualStyleBackColor = True
        '
        'tbReplace
        '
        Me.tbReplace.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbReplace.DataBindings.Add(New System.Windows.Forms.Binding("Text", Global.MetadataEditor.My.MySettings.Default, "ReplaceString", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.tbReplace.Location = New System.Drawing.Point(97, 46)
        Me.tbReplace.Name = "tbReplace"
        Me.tbReplace.Size = New System.Drawing.Size(215, 20)
        Me.tbReplace.TabIndex = 2
        Me.tbReplace.Text = Global.MetadataEditor.My.MySettings.Default.ReplaceString
        '
        'tbFind
        '
        Me.tbFind.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbFind.DataBindings.Add(New System.Windows.Forms.Binding("Text", Global.MetadataEditor.My.MySettings.Default, "FindString", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.tbFind.Location = New System.Drawing.Point(97, 20)
        Me.tbFind.Name = "tbFind"
        Me.tbFind.Size = New System.Drawing.Size(215, 20)
        Me.tbFind.TabIndex = 0
        Me.tbFind.Text = Global.MetadataEditor.My.MySettings.Default.FindString
        '
        'FindReplaceForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(341, 126)
        Me.Controls.Add(Me.btnReplaceAll)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.tbReplace)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.tbFind)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(1020, 154)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(349, 154)
        Me.Name = "FindReplaceForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Find and Replace"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents tbFind As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents tbReplace As System.Windows.Forms.TextBox
    Friend WithEvents btnReplaceAll As System.Windows.Forms.Button
End Class
