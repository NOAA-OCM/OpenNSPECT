<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmSoils
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
	Public WithEvents mnuNew As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuDelete As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuOptions As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuSoilsHelp As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuHelp As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
	Public WithEvents txtSoilsKGrid As System.Windows.Forms.TextBox
	Public WithEvents txtSoilsGrid As System.Windows.Forms.TextBox
	Public WithEvents cboSoils As System.Windows.Forms.ComboBox
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents cmdSave As System.Windows.Forms.Button
	Public WithEvents cmdQuit As System.Windows.Forms.Button
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmSoils))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.MainMenu1 = New System.Windows.Forms.MenuStrip
		Me.mnuOptions = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuNew = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuDelete = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuHelp = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuSoilsHelp = New System.Windows.Forms.ToolStripMenuItem
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me.txtSoilsKGrid = New System.Windows.Forms.TextBox
		Me.txtSoilsGrid = New System.Windows.Forms.TextBox
		Me.cboSoils = New System.Windows.Forms.ComboBox
		Me.Label5 = New System.Windows.Forms.Label
		Me.Label4 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.cmdSave = New System.Windows.Forms.Button
		Me.cmdQuit = New System.Windows.Forms.Button
		Me.MainMenu1.SuspendLayout()
		Me.Frame1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "Soils"
		Me.ClientSize = New System.Drawing.Size(312, 184)
		Me.Location = New System.Drawing.Point(11, 30)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MaximizeBox = True
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmSoils"
		Me.mnuOptions.Name = "mnuOptions"
		Me.mnuOptions.Text = "Options"
		Me.mnuOptions.Checked = False
		Me.mnuOptions.Enabled = True
		Me.mnuOptions.Visible = True
		Me.mnuNew.Name = "mnuNew"
		Me.mnuNew.Text = "New..."
		Me.mnuNew.Checked = False
		Me.mnuNew.Enabled = True
		Me.mnuNew.Visible = True
		Me.mnuDelete.Name = "mnuDelete"
		Me.mnuDelete.Text = "Delete..."
		Me.mnuDelete.Checked = False
		Me.mnuDelete.Enabled = True
		Me.mnuDelete.Visible = True
		Me.mnuHelp.Name = "mnuHelp"
		Me.mnuHelp.Text = "&Help"
		Me.mnuHelp.Checked = False
		Me.mnuHelp.Enabled = True
		Me.mnuHelp.Visible = True
		Me.mnuSoilsHelp.Name = "mnuSoilsHelp"
		Me.mnuSoilsHelp.Text = "Soils..."
		Me.mnuSoilsHelp.ShortcutKeys = CType(System.Windows.Forms.Keys.Shift or System.Windows.Forms.Keys.F1, System.Windows.Forms.Keys)
		Me.mnuSoilsHelp.Checked = False
		Me.mnuSoilsHelp.Enabled = True
		Me.mnuSoilsHelp.Visible = True
		Me.Frame1.Text = "Soils Configuration  "
		Me.Frame1.Size = New System.Drawing.Size(285, 118)
		Me.Frame1.Location = New System.Drawing.Point(13, 25)
		Me.Frame1.TabIndex = 2
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.BackColor = System.Drawing.SystemColors.Control
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame1.Name = "Frame1"
		Me.txtSoilsKGrid.AutoSize = False
		Me.txtSoilsKGrid.Size = New System.Drawing.Size(194, 20)
		Me.txtSoilsKGrid.Location = New System.Drawing.Point(75, 86)
		Me.txtSoilsKGrid.TabIndex = 5
		Me.txtSoilsKGrid.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtSoilsKGrid.AcceptsReturn = True
		Me.txtSoilsKGrid.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtSoilsKGrid.BackColor = System.Drawing.SystemColors.Window
		Me.txtSoilsKGrid.CausesValidation = True
		Me.txtSoilsKGrid.Enabled = True
		Me.txtSoilsKGrid.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtSoilsKGrid.HideSelection = True
		Me.txtSoilsKGrid.ReadOnly = False
		Me.txtSoilsKGrid.Maxlength = 0
		Me.txtSoilsKGrid.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtSoilsKGrid.MultiLine = False
		Me.txtSoilsKGrid.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtSoilsKGrid.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtSoilsKGrid.TabStop = True
		Me.txtSoilsKGrid.Visible = True
		Me.txtSoilsKGrid.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtSoilsKGrid.Name = "txtSoilsKGrid"
		Me.txtSoilsGrid.AutoSize = False
		Me.txtSoilsGrid.Size = New System.Drawing.Size(192, 20)
		Me.txtSoilsGrid.Location = New System.Drawing.Point(75, 53)
		Me.txtSoilsGrid.TabIndex = 4
		Me.txtSoilsGrid.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtSoilsGrid.AcceptsReturn = True
		Me.txtSoilsGrid.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtSoilsGrid.BackColor = System.Drawing.SystemColors.Window
		Me.txtSoilsGrid.CausesValidation = True
		Me.txtSoilsGrid.Enabled = True
		Me.txtSoilsGrid.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtSoilsGrid.HideSelection = True
		Me.txtSoilsGrid.ReadOnly = False
		Me.txtSoilsGrid.Maxlength = 0
		Me.txtSoilsGrid.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtSoilsGrid.MultiLine = False
		Me.txtSoilsGrid.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtSoilsGrid.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtSoilsGrid.TabStop = True
		Me.txtSoilsGrid.Visible = True
		Me.txtSoilsGrid.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtSoilsGrid.Name = "txtSoilsGrid"
		Me.cboSoils.Size = New System.Drawing.Size(143, 21)
		Me.cboSoils.Location = New System.Drawing.Point(76, 21)
		Me.cboSoils.TabIndex = 3
		Me.cboSoils.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboSoils.BackColor = System.Drawing.SystemColors.Window
		Me.cboSoils.CausesValidation = True
		Me.cboSoils.Enabled = True
		Me.cboSoils.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboSoils.IntegralHeight = True
		Me.cboSoils.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboSoils.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboSoils.Sorted = False
		Me.cboSoils.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboSoils.TabStop = True
		Me.cboSoils.Visible = True
		Me.cboSoils.Name = "cboSoils"
		Me.Label5.Text = "Soils K Grid:"
		Me.Label5.Size = New System.Drawing.Size(77, 19)
		Me.Label5.Location = New System.Drawing.Point(7, 88)
		Me.Label5.TabIndex = 8
		Me.Label5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label5.BackColor = System.Drawing.SystemColors.Control
		Me.Label5.Enabled = True
		Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label5.UseMnemonic = True
		Me.Label5.Visible = True
		Me.Label5.AutoSize = False
		Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label5.Name = "Label5"
		Me.Label4.Text = "Soils GRID:"
		Me.Label4.Size = New System.Drawing.Size(59, 18)
		Me.Label4.Location = New System.Drawing.Point(7, 55)
		Me.Label4.TabIndex = 7
		Me.Label4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label4.BackColor = System.Drawing.SystemColors.Control
		Me.Label4.Enabled = True
		Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label4.UseMnemonic = True
		Me.Label4.Visible = True
		Me.Label4.AutoSize = False
		Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label4.Name = "Label4"
		Me.Label1.Text = "Name:"
		Me.Label1.Size = New System.Drawing.Size(53, 18)
		Me.Label1.Location = New System.Drawing.Point(8, 23)
		Me.Label1.TabIndex = 6
		Me.Label1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.BackColor = System.Drawing.SystemColors.Control
		Me.Label1.Enabled = True
		Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.cmdSave.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdSave.Text = "OK"
		Me.cmdSave.Size = New System.Drawing.Size(65, 25)
		Me.cmdSave.Location = New System.Drawing.Point(153, 152)
		Me.cmdSave.TabIndex = 1
		Me.cmdSave.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdSave.BackColor = System.Drawing.SystemColors.Control
		Me.cmdSave.CausesValidation = True
		Me.cmdSave.Enabled = True
		Me.cmdSave.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdSave.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdSave.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdSave.TabStop = True
		Me.cmdSave.Name = "cmdSave"
		Me.cmdQuit.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdQuit.Text = "Cancel"
		Me.cmdQuit.Size = New System.Drawing.Size(65, 25)
		Me.cmdQuit.Location = New System.Drawing.Point(224, 152)
		Me.cmdQuit.TabIndex = 0
		Me.cmdQuit.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdQuit.BackColor = System.Drawing.SystemColors.Control
		Me.cmdQuit.CausesValidation = True
		Me.cmdQuit.Enabled = True
		Me.cmdQuit.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdQuit.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdQuit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdQuit.TabStop = True
		Me.cmdQuit.Name = "cmdQuit"
		Me.Controls.Add(Frame1)
		Me.Controls.Add(cmdSave)
		Me.Controls.Add(cmdQuit)
		Me.Frame1.Controls.Add(txtSoilsKGrid)
		Me.Frame1.Controls.Add(txtSoilsGrid)
		Me.Frame1.Controls.Add(cboSoils)
		Me.Frame1.Controls.Add(Label5)
		Me.Frame1.Controls.Add(Label4)
		Me.Frame1.Controls.Add(Label1)
		MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuOptions, Me.mnuHelp})
		mnuOptions.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuNew, Me.mnuDelete})
		mnuHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuSoilsHelp})
		Me.Controls.Add(MainMenu1)
		Me.MainMenu1.ResumeLayout(False)
		Me.Frame1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class