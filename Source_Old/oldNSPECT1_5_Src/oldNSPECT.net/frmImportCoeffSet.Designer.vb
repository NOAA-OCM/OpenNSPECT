<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmImportCoeffSet
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
	Public WithEvents cboLCType As System.Windows.Forms.ComboBox
	Public WithEvents cmdBrowse As System.Windows.Forms.Button
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents txtImpFile As System.Windows.Forms.TextBox
	Public WithEvents txtCoeffSetName As System.Windows.Forms.TextBox
	Public dlgCMD1Open As System.Windows.Forms.OpenFileDialog
	Public WithEvents _Label1_5 As System.Windows.Forms.Label
	Public WithEvents _Label1_0 As System.Windows.Forms.Label
	Public WithEvents _Label1_7 As System.Windows.Forms.Label
	Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmImportCoeffSet))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cboLCType = New System.Windows.Forms.ComboBox
		Me.cmdBrowse = New System.Windows.Forms.Button
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cmdOK = New System.Windows.Forms.Button
		Me.txtImpFile = New System.Windows.Forms.TextBox
		Me.txtCoeffSetName = New System.Windows.Forms.TextBox
		Me.dlgCMD1Open = New System.Windows.Forms.OpenFileDialog
		Me._Label1_5 = New System.Windows.Forms.Label
		Me._Label1_0 = New System.Windows.Forms.Label
		Me._Label1_7 = New System.Windows.Forms.Label
		Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "Import Coefficient Set"
		Me.ClientSize = New System.Drawing.Size(436, 132)
		Me.Location = New System.Drawing.Point(3, 21)
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
		Me.Name = "frmImportCoeffSet"
		Me.cboLCType.Size = New System.Drawing.Size(147, 21)
		Me.cboLCType.Location = New System.Drawing.Point(162, 39)
		Me.cboLCType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cboLCType.TabIndex = 1
		Me.cboLCType.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboLCType.BackColor = System.Drawing.SystemColors.Window
		Me.cboLCType.CausesValidation = True
		Me.cboLCType.Enabled = True
		Me.cboLCType.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboLCType.IntegralHeight = True
		Me.cboLCType.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboLCType.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboLCType.Sorted = False
		Me.cboLCType.TabStop = True
		Me.cboLCType.Visible = True
		Me.cboLCType.Name = "cboLCType"
		Me.cmdBrowse.TextAlign = System.Drawing.ContentAlignment.BottomCenter
		Me.cmdBrowse.Size = New System.Drawing.Size(25, 21)
		Me.cmdBrowse.Location = New System.Drawing.Point(388, 68)
		Me.cmdBrowse.Image = CType(resources.GetObject("cmdBrowse.Image"), System.Drawing.Image)
		Me.cmdBrowse.TabIndex = 3
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
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(65, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(342, 100)
		Me.cmdCancel.TabIndex = 5
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
		Me.cmdOK.Location = New System.Drawing.Point(273, 100)
		Me.cmdOK.TabIndex = 4
		Me.cmdOK.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOK.CausesValidation = True
		Me.cmdOK.Enabled = True
		Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOK.TabStop = True
		Me.cmdOK.Name = "cmdOK"
		Me.txtImpFile.AutoSize = False
		Me.txtImpFile.Size = New System.Drawing.Size(224, 19)
		Me.txtImpFile.Location = New System.Drawing.Point(161, 69)
		Me.txtImpFile.TabIndex = 2
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
		Me.txtCoeffSetName.AutoSize = False
		Me.txtCoeffSetName.Size = New System.Drawing.Size(143, 19)
		Me.txtCoeffSetName.Location = New System.Drawing.Point(162, 12)
		Me.txtCoeffSetName.TabIndex = 0
		Me.txtCoeffSetName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtCoeffSetName.AcceptsReturn = True
		Me.txtCoeffSetName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtCoeffSetName.BackColor = System.Drawing.SystemColors.Window
		Me.txtCoeffSetName.CausesValidation = True
		Me.txtCoeffSetName.Enabled = True
		Me.txtCoeffSetName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtCoeffSetName.HideSelection = True
		Me.txtCoeffSetName.ReadOnly = False
		Me.txtCoeffSetName.Maxlength = 0
		Me.txtCoeffSetName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtCoeffSetName.MultiLine = False
		Me.txtCoeffSetName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtCoeffSetName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtCoeffSetName.TabStop = True
		Me.txtCoeffSetName.Visible = True
		Me.txtCoeffSetName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtCoeffSetName.Name = "txtCoeffSetName"
		Me._Label1_5.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_5.Text = "New Coefficient Set Name:"
		Me._Label1_5.Size = New System.Drawing.Size(189, 17)
		Me._Label1_5.Location = New System.Drawing.Point(-48, 15)
		Me._Label1_5.TabIndex = 8
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
		Me._Label1_0.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_0.Text = "Import File:"
		Me._Label1_0.Size = New System.Drawing.Size(133, 17)
		Me._Label1_0.Location = New System.Drawing.Point(7, 69)
		Me._Label1_0.TabIndex = 7
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
		Me._Label1_7.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_7.Text = "Land Cover Type:"
		Me._Label1_7.Size = New System.Drawing.Size(127, 17)
		Me._Label1_7.Location = New System.Drawing.Point(14, 42)
		Me._Label1_7.TabIndex = 6
		Me._Label1_7.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_7.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_7.Enabled = True
		Me._Label1_7.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_7.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_7.UseMnemonic = True
		Me._Label1_7.Visible = True
		Me._Label1_7.AutoSize = False
		Me._Label1_7.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_7.Name = "_Label1_7"
		Me.Controls.Add(cboLCType)
		Me.Controls.Add(cmdBrowse)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(txtImpFile)
		Me.Controls.Add(txtCoeffSetName)
		Me.Controls.Add(_Label1_5)
		Me.Controls.Add(_Label1_0)
		Me.Controls.Add(_Label1_7)
		Me.Label1.SetIndex(_Label1_5, CType(5, Short))
		Me.Label1.SetIndex(_Label1_0, CType(0, Short))
		Me.Label1.SetIndex(_Label1_7, CType(7, Short))
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class