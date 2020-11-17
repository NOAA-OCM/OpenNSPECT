<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmImportWQStd
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
	Public WithEvents cmdBrowse As System.Windows.Forms.Button
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents txtImpFile As System.Windows.Forms.TextBox
	Public WithEvents txtStdName As System.Windows.Forms.TextBox
	Public dlgCMD1Open As System.Windows.Forms.OpenFileDialog
	Public WithEvents _Label1_0 As System.Windows.Forms.Label
	Public WithEvents _Label1_5 As System.Windows.Forms.Label
	Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmImportWQStd))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdBrowse = New System.Windows.Forms.Button
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cmdOK = New System.Windows.Forms.Button
		Me.txtImpFile = New System.Windows.Forms.TextBox
		Me.txtStdName = New System.Windows.Forms.TextBox
		Me.dlgCMD1Open = New System.Windows.Forms.OpenFileDialog
		Me._Label1_0 = New System.Windows.Forms.Label
		Me._Label1_5 = New System.Windows.Forms.Label
		Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "Import Water Quality Standard"
		Me.ClientSize = New System.Drawing.Size(367, 115)
		Me.Location = New System.Drawing.Point(3, 21)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
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
		Me.Name = "frmImportWQStd"
		Me.cmdBrowse.TextAlign = System.Drawing.ContentAlignment.BottomCenter
		Me.cmdBrowse.Size = New System.Drawing.Size(25, 21)
		Me.cmdBrowse.Location = New System.Drawing.Point(326, 46)
		Me.cmdBrowse.Image = CType(resources.GetObject("cmdBrowse.Image"), System.Drawing.Image)
		Me.cmdBrowse.TabIndex = 2
		Me.cmdBrowse.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdBrowse.BackColor = System.Drawing.SystemColors.Control
		Me.cmdBrowse.CausesValidation = True
		Me.cmdBrowse.Enabled = True
		Me.cmdBrowse.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdBrowse.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdBrowse.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdBrowse.TabStop = True
		Me.cmdBrowse.Name = "cmdBrowse"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdCancel
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(65, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(259, 81)
		Me.cmdCancel.TabIndex = 4
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
		Me.cmdOK.Enabled = False
		Me.cmdOK.Size = New System.Drawing.Size(65, 25)
		Me.cmdOK.Location = New System.Drawing.Point(189, 81)
		Me.cmdOK.TabIndex = 3
		Me.cmdOK.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOK.CausesValidation = True
		Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOK.TabStop = True
		Me.cmdOK.Name = "cmdOK"
		Me.txtImpFile.AutoSize = False
		Me.txtImpFile.Size = New System.Drawing.Size(200, 19)
		Me.txtImpFile.Location = New System.Drawing.Point(123, 47)
		Me.txtImpFile.TabIndex = 1
		Me.txtImpFile.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtImpFile.AcceptsReturn = True
		Me.txtImpFile.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtImpFile.BackColor = System.Drawing.SystemColors.Window
		Me.txtImpFile.CausesValidation = True
		Me.txtImpFile.Enabled = True
		Me.txtImpFile.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtImpFile.HideSelection = True
		Me.txtImpFile.ReadOnly = False
		Me.txtImpFile.Maxlength = 0
		Me.txtImpFile.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtImpFile.MultiLine = False
		Me.txtImpFile.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtImpFile.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtImpFile.TabStop = True
		Me.txtImpFile.Visible = True
		Me.txtImpFile.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtImpFile.Name = "txtImpFile"
		Me.txtStdName.AutoSize = False
		Me.txtStdName.Size = New System.Drawing.Size(162, 19)
		Me.txtStdName.Location = New System.Drawing.Point(123, 12)
		Me.txtStdName.TabIndex = 0
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
		Me._Label1_0.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_0.Text = "Import File:"
		Me._Label1_0.Size = New System.Drawing.Size(100, 17)
		Me._Label1_0.Location = New System.Drawing.Point(17, 50)
		Me._Label1_0.TabIndex = 6
		Me._Label1_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
		Me._Label1_5.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_5.Text = "New Standard Name:"
		Me._Label1_5.Size = New System.Drawing.Size(111, 17)
		Me._Label1_5.Location = New System.Drawing.Point(6, 16)
		Me._Label1_5.TabIndex = 5
		Me._Label1_5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
		Me.Controls.Add(cmdBrowse)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(txtImpFile)
		Me.Controls.Add(txtStdName)
		Me.Controls.Add(_Label1_0)
		Me.Controls.Add(_Label1_5)
		Me.Label1.SetIndex(_Label1_0, CType(0, Short))
		Me.Label1.SetIndex(_Label1_5, CType(5, Short))
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class