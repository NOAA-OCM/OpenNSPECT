<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmCopyCoeffSet
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
	Public WithEvents cboCoeffSet As System.Windows.Forms.ComboBox
	Public WithEvents txtCoeffSetName As System.Windows.Forms.TextBox
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents _Label1_0 As System.Windows.Forms.Label
	Public WithEvents _Label1_5 As System.Windows.Forms.Label
	Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmCopyCoeffSet))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cboCoeffSet = New System.Windows.Forms.ComboBox
		Me.txtCoeffSetName = New System.Windows.Forms.TextBox
		Me.cmdOK = New System.Windows.Forms.Button
		Me.cmdCancel = New System.Windows.Forms.Button
		Me._Label1_0 = New System.Windows.Forms.Label
		Me._Label1_5 = New System.Windows.Forms.Label
		Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "Copy Coefficient Set"
		Me.ClientSize = New System.Drawing.Size(367, 115)
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
		Me.Name = "frmCopyCoeffSet"
		Me.cboCoeffSet.Size = New System.Drawing.Size(158, 21)
		Me.cboCoeffSet.Location = New System.Drawing.Point(176, 13)
		Me.cboCoeffSet.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cboCoeffSet.TabIndex = 0
		Me.cboCoeffSet.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboCoeffSet.BackColor = System.Drawing.SystemColors.Window
		Me.cboCoeffSet.CausesValidation = True
		Me.cboCoeffSet.Enabled = True
		Me.cboCoeffSet.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboCoeffSet.IntegralHeight = True
		Me.cboCoeffSet.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboCoeffSet.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboCoeffSet.Sorted = False
		Me.cboCoeffSet.TabStop = True
		Me.cboCoeffSet.Visible = True
		Me.cboCoeffSet.Name = "cboCoeffSet"
		Me.txtCoeffSetName.AutoSize = False
		Me.txtCoeffSetName.Size = New System.Drawing.Size(156, 20)
		Me.txtCoeffSetName.Location = New System.Drawing.Point(176, 46)
		Me.txtCoeffSetName.TabIndex = 1
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
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOK.Text = "OK"
		Me.cmdOK.Size = New System.Drawing.Size(65, 25)
		Me.cmdOK.Location = New System.Drawing.Point(199, 81)
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
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(65, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(269, 81)
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
		Me._Label1_0.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_0.Text = "Copy from Coefficient Set:"
		Me._Label1_0.Size = New System.Drawing.Size(160, 17)
		Me._Label1_0.Location = New System.Drawing.Point(0, 18)
		Me._Label1_0.TabIndex = 5
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
		Me._Label1_5.Text = "New Coefficient Set Name:"
		Me._Label1_5.Size = New System.Drawing.Size(140, 17)
		Me._Label1_5.Location = New System.Drawing.Point(24, 50)
		Me._Label1_5.TabIndex = 4
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
		Me.Controls.Add(cboCoeffSet)
		Me.Controls.Add(txtCoeffSetName)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(cmdCancel)
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