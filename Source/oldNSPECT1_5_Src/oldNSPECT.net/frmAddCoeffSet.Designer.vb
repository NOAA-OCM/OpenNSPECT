<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmAddCoeffSet
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
	Public WithEvents txtCoeffSetName As System.Windows.Forms.TextBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents _Label1_5 As System.Windows.Forms.Label
	Public WithEvents _Label1_7 As System.Windows.Forms.Label
	Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmAddCoeffSet))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cboLCType = New System.Windows.Forms.ComboBox
		Me.txtCoeffSetName = New System.Windows.Forms.TextBox
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cmdOK = New System.Windows.Forms.Button
		Me._Label1_5 = New System.Windows.Forms.Label
		Me._Label1_7 = New System.Windows.Forms.Label
		Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "Add Coefficient Set"
		Me.ClientSize = New System.Drawing.Size(372, 115)
		Me.Location = New System.Drawing.Point(519, 231)
		Me.Icon = CType(resources.GetObject("frmAddCoeffSet.Icon"), System.Drawing.Icon)
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
		Me.Name = "frmAddCoeffSet"
		Me.cboLCType.Size = New System.Drawing.Size(173, 21)
		Me.cboLCType.Location = New System.Drawing.Point(154, 46)
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
		Me.txtCoeffSetName.AutoSize = False
		Me.txtCoeffSetName.Size = New System.Drawing.Size(175, 21)
		Me.txtCoeffSetName.Location = New System.Drawing.Point(154, 15)
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
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(65, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(265, 84)
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
		Me.cmdOK.Enabled = False
		Me.cmdOK.Size = New System.Drawing.Size(65, 25)
		Me.cmdOK.Location = New System.Drawing.Point(194, 84)
		Me.cmdOK.TabIndex = 2
		Me.cmdOK.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOK.CausesValidation = True
		Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOK.TabStop = True
		Me.cmdOK.Name = "cmdOK"
		Me._Label1_5.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_5.Text = "Coefficient Set Name:"
		Me._Label1_5.Size = New System.Drawing.Size(150, 17)
		Me._Label1_5.Location = New System.Drawing.Point(-14, 18)
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
		Me._Label1_7.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_7.Text = "Land Cover Type:"
		Me._Label1_7.Size = New System.Drawing.Size(102, 17)
		Me._Label1_7.Location = New System.Drawing.Point(33, 49)
		Me._Label1_7.TabIndex = 4
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
		Me.Controls.Add(txtCoeffSetName)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(_Label1_5)
		Me.Controls.Add(_Label1_7)
		Me.Label1.SetIndex(_Label1_5, CType(5, Short))
		Me.Label1.SetIndex(_Label1_7, CType(7, Short))
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class