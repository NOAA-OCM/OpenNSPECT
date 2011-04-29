<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class BaseDialogForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> Protected Overrides Sub Dispose(ByVal disposing As Boolean)
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
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(BaseDialogForm))
        Me.FooterButtonTableLayoutPanel = New System.Windows.Forms.TableLayoutPanel()
        Me.OK_Button = New System.Windows.Forms.Button()
        Me.Cancel_Button = New System.Windows.Forms.Button()
        Me.FooterButtonPanel = New System.Windows.Forms.Panel()
        Me.FooterButtonTableLayoutPanel.SuspendLayout()
        Me.FooterButtonPanel.SuspendLayout()
        Me.SuspendLayout()
        '
        'FooterButtonTableLayoutPanel
        '
        Me.FooterButtonTableLayoutPanel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.FooterButtonTableLayoutPanel.ColumnCount = 2
        Me.FooterButtonTableLayoutPanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.FooterButtonTableLayoutPanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.FooterButtonTableLayoutPanel.Controls.Add(Me.OK_Button, 0, 0)
        Me.FooterButtonTableLayoutPanel.Controls.Add(Me.Cancel_Button, 1, 0)
        Me.FooterButtonTableLayoutPanel.Location = New System.Drawing.Point(286, 3)
        Me.FooterButtonTableLayoutPanel.Name = "FooterButtonTableLayoutPanel"
        Me.FooterButtonTableLayoutPanel.RowCount = 1
        Me.FooterButtonTableLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.FooterButtonTableLayoutPanel.Size = New System.Drawing.Size(146, 29)
        Me.FooterButtonTableLayoutPanel.TabIndex = 0
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.OK_Button.Location = New System.Drawing.Point(3, 3)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(67, 23)
        Me.OK_Button.TabIndex = 0
        Me.OK_Button.Text = "OK"
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.Location = New System.Drawing.Point(76, 3)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(67, 23)
        Me.Cancel_Button.TabIndex = 1
        Me.Cancel_Button.Text = "Cancel"
        '
        'FooterButtonPanel
        '
        Me.FooterButtonPanel.Controls.Add(Me.FooterButtonTableLayoutPanel)
        Me.FooterButtonPanel.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.FooterButtonPanel.Location = New System.Drawing.Point(0, 280)
        Me.FooterButtonPanel.Name = "FooterButtonPanel"
        Me.FooterButtonPanel.Size = New System.Drawing.Size(435, 35)
        Me.FooterButtonPanel.TabIndex = 1
        '
        'BaseDialogForm
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(435, 315)
        Me.Controls.Add(Me.FooterButtonPanel)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "BaseDialogForm"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "OpenNSPECT"
        Me.FooterButtonTableLayoutPanel.ResumeLayout(False)
        Me.FooterButtonPanel.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents FooterButtonTableLayoutPanel As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents FooterButtonPanel As System.Windows.Forms.Panel

End Class
