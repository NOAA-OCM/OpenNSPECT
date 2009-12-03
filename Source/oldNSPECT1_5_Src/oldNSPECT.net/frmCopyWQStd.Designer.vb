<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmCopyWQStd
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents txtStdName As System.Windows.Forms.TextBox
	Public WithEvents cboStdName As System.Windows.Forms.ComboBox
	Public WithEvents _Label1_5 As System.Windows.Forms.Label
	Public WithEvents _Label1_0 As System.Windows.Forms.Label
	Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmCopyWQStd))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cmdOK = New System.Windows.Forms.Button
		Me.txtStdName = New System.Windows.Forms.TextBox
		Me.cboStdName = New System.Windows.Forms.ComboBox
		Me._Label1_5 = New System.Windows.Forms.Label
		Me._Label1_0 = New System.Windows.Forms.Label
		Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "Copy Water Quality Standard"
		Me.ClientSize = New System.Drawing.Size(367, 115)
		Me.Location = New System.Drawing.Point(473, 457)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmCopyWQStd"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(65, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(263, 81)
		Me.cmdCancel.TabIndex = 3
		Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCancel.CausesValidation = True
		Me.cmdCancel.Enabled = True
		Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.TabStop = True
		Me.cmdCancel.Name = "cmdCancel"
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOK.Text = "OK"
		Me.cmdOK.Size = New System.Drawing.Size(65, 25)
		Me.cmdOK.Location = New System.Drawing.Point(192, 81)
		Me.cmdOK.TabIndex = 2
		Me.cmdOK.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOK.CausesValidation = True
		Me.cmdOK.Enabled = True
		Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOK.TabStop = True
		Me.cmdOK.Name = "cmdOK"
		Me.txtStdName.AutoSize = False
		Me.txtStdName.Size = New System.Drawing.Size(158, 19)
		Me.txtStdName.Location = New System.Drawing.Point(175, 49)
		Me.txtStdName.TabIndex = 1
		Me.txtStdName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtStdName.AcceptsReturn = True
		Me.txtStdName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtStdName.BackColor = System.Drawing.SystemColors.Window
		Me.txtStdName.CausesValidation = True
		Me.txtStdName.Enabled = True
		Me.txtStdName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtStdName.HideSelection = True
		Me.txtStdName.ReadOnly = False
		Me.txtStdName.Maxlength = 0
		Me.txtStdName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtStdName.MultiLine = False
		Me.txtStdName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtStdName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtStdName.TabStop = True
		Me.txtStdName.Visible = True
		Me.txtStdName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtStdName.Name = "txtStdName"
		Me.cboStdName.Size = New System.Drawing.Size(158, 21)
		Me.cboStdName.Location = New System.Drawing.Point(175, 15)
		Me.cboStdName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cboStdName.TabIndex = 0
		Me.cboStdName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboStdName.BackColor = System.Drawing.SystemColors.Window
		Me.cboStdName.CausesValidation = True
		Me.cboStdName.Enabled = True
		Me.cboStdName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboStdName.IntegralHeight = True
		Me.cboStdName.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboStdName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboStdName.Sorted = False
		Me.cboStdName.TabStop = True
		Me.cboStdName.Visible = True
		Me.cboStdName.Name = "cboStdName"
		Me._Label1_5.Text = "New Standard Name:"
		Me._Label1_5.Size = New System.Drawing.Size(110, 17)
		Me._Label1_5.Location = New System.Drawing.Point(30, 51)
		Me._Label1_5.TabIndex = 5
		Me._Label1_5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_5.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_5.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_5.Enabled = True
		Me._Label1_5.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_5.UseMnemonic = True
		Me._Label1_5.Visible = True
		Me._Label1_5.AutoSize = False
		Me._Label1_5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_5.Name = "_Label1_5"
		Me._Label1_0.Text = "Copy from Standard Name:"
		Me._Label1_0.Size = New System.Drawing.Size(152, 17)
		Me._Label1_0.Location = New System.Drawing.Point(30, 17)
		Me._Label1_0.TabIndex = 4
		Me._Label1_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_0.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_0.Enabled = True
		Me._Label1_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_0.UseMnemonic = True
		Me._Label1_0.Visible = True
		Me._Label1_0.AutoSize = False
		Me._Label1_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_0.Name = "_Label1_0"
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(txtStdName)
		Me.Controls.Add(cboStdName)
		Me.Controls.Add(_Label1_5)
		Me.Controls.Add(_Label1_0)
		Me.Label1.SetIndex(_Label1_5, CType(5, Short))
		Me.Label1.SetIndex(_Label1_0, CType(0, Short))
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class