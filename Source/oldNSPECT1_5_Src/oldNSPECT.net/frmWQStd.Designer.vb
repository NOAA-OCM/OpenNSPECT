<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmWQStd
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
	Public WithEvents mnuNewWQStd As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuDelWQStd As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuCopyWQStd As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuImpWQStd As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuExpWQStd As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuWQStd As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuAddRow As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuInsertRow As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuDeleteRow As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuEditCell As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuWQHelp As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuHelp As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
	Public WithEvents txtActiveCell As System.Windows.Forms.TextBox
	Public WithEvents grdWQStd As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
	Public WithEvents cmdQuit As System.Windows.Forms.Button
	Public WithEvents cmdSave As System.Windows.Forms.Button
	Public WithEvents cboWQStdName As System.Windows.Forms.ComboBox
	Public WithEvents txtWQStdDesc As System.Windows.Forms.TextBox
	Public dlgCMD1Open As System.Windows.Forms.OpenFileDialog
	Public dlgCMD1Save As System.Windows.Forms.SaveFileDialog
	Public WithEvents _Label1_1 As System.Windows.Forms.Label
	Public WithEvents _Label1_0 As System.Windows.Forms.Label
	Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmWQStd))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.MainMenu1 = New System.Windows.Forms.MenuStrip
		Me.mnuWQStd = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuNewWQStd = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuDelWQStd = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuCopyWQStd = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuImpWQStd = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuExpWQStd = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuEditCell = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuAddRow = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuInsertRow = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuDeleteRow = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuHelp = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuWQHelp = New System.Windows.Forms.ToolStripMenuItem
		Me.txtActiveCell = New System.Windows.Forms.TextBox
		Me.grdWQStd = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
		Me.cmdQuit = New System.Windows.Forms.Button
		Me.cmdSave = New System.Windows.Forms.Button
		Me.cboWQStdName = New System.Windows.Forms.ComboBox
		Me.txtWQStdDesc = New System.Windows.Forms.TextBox
		Me.dlgCMD1Open = New System.Windows.Forms.OpenFileDialog
		Me.dlgCMD1Save = New System.Windows.Forms.SaveFileDialog
		Me._Label1_1 = New System.Windows.Forms.Label
		Me._Label1_0 = New System.Windows.Forms.Label
		Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.MainMenu1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.grdWQStd, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Text = "Water Quality Standards"
		Me.ClientSize = New System.Drawing.Size(416, 302)
		Me.Location = New System.Drawing.Point(9, 27)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmWQStd"
		Me.mnuWQStd.Name = "mnuWQStd"
		Me.mnuWQStd.Text = "&Options"
		Me.mnuWQStd.Checked = False
		Me.mnuWQStd.Enabled = True
		Me.mnuWQStd.Visible = True
		Me.mnuNewWQStd.Name = "mnuNewWQStd"
		Me.mnuNewWQStd.Text = "New..."
		Me.mnuNewWQStd.Checked = False
		Me.mnuNewWQStd.Enabled = True
		Me.mnuNewWQStd.Visible = True
		Me.mnuDelWQStd.Name = "mnuDelWQStd"
		Me.mnuDelWQStd.Text = "Delete..."
		Me.mnuDelWQStd.Checked = False
		Me.mnuDelWQStd.Enabled = True
		Me.mnuDelWQStd.Visible = True
		Me.mnuCopyWQStd.Name = "mnuCopyWQStd"
		Me.mnuCopyWQStd.Text = "Copy..."
		Me.mnuCopyWQStd.Checked = False
		Me.mnuCopyWQStd.Enabled = True
		Me.mnuCopyWQStd.Visible = True
		Me.mnuImpWQStd.Name = "mnuImpWQStd"
		Me.mnuImpWQStd.Text = "Import..."
		Me.mnuImpWQStd.Checked = False
		Me.mnuImpWQStd.Enabled = True
		Me.mnuImpWQStd.Visible = True
		Me.mnuExpWQStd.Name = "mnuExpWQStd"
		Me.mnuExpWQStd.Text = "Export..."
		Me.mnuExpWQStd.Checked = False
		Me.mnuExpWQStd.Enabled = True
		Me.mnuExpWQStd.Visible = True
		Me.mnuEditCell.Name = "mnuEditCell"
		Me.mnuEditCell.Text = "Edit"
		Me.mnuEditCell.Visible = False
		Me.mnuEditCell.Checked = False
		Me.mnuEditCell.Enabled = True
		Me.mnuAddRow.Name = "mnuAddRow"
		Me.mnuAddRow.Text = "Add Row"
		Me.mnuAddRow.Checked = False
		Me.mnuAddRow.Enabled = True
		Me.mnuAddRow.Visible = True
		Me.mnuInsertRow.Name = "mnuInsertRow"
		Me.mnuInsertRow.Text = "Insert Row"
		Me.mnuInsertRow.Checked = False
		Me.mnuInsertRow.Enabled = True
		Me.mnuInsertRow.Visible = True
		Me.mnuDeleteRow.Name = "mnuDeleteRow"
		Me.mnuDeleteRow.Text = "Delete Row"
		Me.mnuDeleteRow.Checked = False
		Me.mnuDeleteRow.Enabled = True
		Me.mnuDeleteRow.Visible = True
		Me.mnuHelp.Name = "mnuHelp"
		Me.mnuHelp.Text = "&Help"
		Me.mnuHelp.Checked = False
		Me.mnuHelp.Enabled = True
		Me.mnuHelp.Visible = True
		Me.mnuWQHelp.Name = "mnuWQHelp"
		Me.mnuWQHelp.Text = "Water Quality Standards..."
		Me.mnuWQHelp.ShortcutKeys = CType(System.Windows.Forms.Keys.Shift or System.Windows.Forms.Keys.F1, System.Windows.Forms.Keys)
		Me.mnuWQHelp.Checked = False
		Me.mnuWQHelp.Enabled = True
		Me.mnuWQHelp.Visible = True
		Me.txtActiveCell.AutoSize = False
		Me.txtActiveCell.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
		Me.txtActiveCell.BackColor = System.Drawing.Color.White
		Me.txtActiveCell.Size = New System.Drawing.Size(66, 19)
		Me.txtActiveCell.Location = New System.Drawing.Point(131, 263)
		Me.txtActiveCell.TabIndex = 7
		Me.txtActiveCell.Text = "Text1"
		Me.txtActiveCell.Visible = False
		Me.txtActiveCell.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtActiveCell.AcceptsReturn = True
		Me.txtActiveCell.CausesValidation = True
		Me.txtActiveCell.Enabled = True
		Me.txtActiveCell.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtActiveCell.HideSelection = True
		Me.txtActiveCell.ReadOnly = False
		Me.txtActiveCell.Maxlength = 0
		Me.txtActiveCell.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtActiveCell.MultiLine = False
		Me.txtActiveCell.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtActiveCell.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtActiveCell.TabStop = True
		Me.txtActiveCell.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.txtActiveCell.Name = "txtActiveCell"
		grdWQStd.OcxState = CType(resources.GetObject("grdWQStd.OcxState"), System.Windows.Forms.AxHost.State)
		Me.grdWQStd.Size = New System.Drawing.Size(364, 161)
		Me.grdWQStd.Location = New System.Drawing.Point(23, 89)
		Me.grdWQStd.TabIndex = 6
		Me.grdWQStd.Name = "grdWQStd"
		Me.cmdQuit.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdQuit.Text = "Cancel"
		Me.cmdQuit.Size = New System.Drawing.Size(65, 25)
		Me.cmdQuit.Location = New System.Drawing.Point(316, 262)
		Me.cmdQuit.TabIndex = 3
		Me.cmdQuit.TabStop = False
		Me.cmdQuit.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdQuit.BackColor = System.Drawing.SystemColors.Control
		Me.cmdQuit.CausesValidation = True
		Me.cmdQuit.Enabled = True
		Me.cmdQuit.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdQuit.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdQuit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdQuit.Name = "cmdQuit"
		Me.cmdSave.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdSave.Text = "OK"
		Me.cmdSave.Enabled = False
		Me.cmdSave.Size = New System.Drawing.Size(65, 25)
		Me.cmdSave.Location = New System.Drawing.Point(247, 262)
		Me.cmdSave.TabIndex = 2
		Me.cmdSave.TabStop = False
		Me.cmdSave.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdSave.BackColor = System.Drawing.SystemColors.Control
		Me.cmdSave.CausesValidation = True
		Me.cmdSave.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdSave.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdSave.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdSave.Name = "cmdSave"
		Me.cboWQStdName.Size = New System.Drawing.Size(160, 21)
		Me.cboWQStdName.Location = New System.Drawing.Point(97, 30)
		Me.cboWQStdName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cboWQStdName.TabIndex = 0
		Me.cboWQStdName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboWQStdName.BackColor = System.Drawing.SystemColors.Window
		Me.cboWQStdName.CausesValidation = True
		Me.cboWQStdName.Enabled = True
		Me.cboWQStdName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboWQStdName.IntegralHeight = True
		Me.cboWQStdName.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboWQStdName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboWQStdName.Sorted = False
		Me.cboWQStdName.TabStop = True
		Me.cboWQStdName.Visible = True
		Me.cboWQStdName.Name = "cboWQStdName"
		Me.txtWQStdDesc.AutoSize = False
		Me.txtWQStdDesc.Size = New System.Drawing.Size(286, 20)
		Me.txtWQStdDesc.Location = New System.Drawing.Point(97, 58)
		Me.txtWQStdDesc.Maxlength = 100
		Me.txtWQStdDesc.TabIndex = 1
		Me.txtWQStdDesc.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtWQStdDesc.AcceptsReturn = True
		Me.txtWQStdDesc.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtWQStdDesc.BackColor = System.Drawing.SystemColors.Window
		Me.txtWQStdDesc.CausesValidation = True
		Me.txtWQStdDesc.Enabled = True
		Me.txtWQStdDesc.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtWQStdDesc.HideSelection = True
		Me.txtWQStdDesc.ReadOnly = False
		Me.txtWQStdDesc.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtWQStdDesc.MultiLine = False
		Me.txtWQStdDesc.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtWQStdDesc.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtWQStdDesc.TabStop = True
		Me.txtWQStdDesc.Visible = True
		Me.txtWQStdDesc.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtWQStdDesc.Name = "txtWQStdDesc"
		Me._Label1_1.Text = "Description:"
		Me._Label1_1.Size = New System.Drawing.Size(67, 19)
		Me._Label1_1.Location = New System.Drawing.Point(15, 60)
		Me._Label1_1.TabIndex = 5
		Me._Label1_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_1.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_1.Enabled = True
		Me._Label1_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_1.UseMnemonic = True
		Me._Label1_1.Visible = True
		Me._Label1_1.AutoSize = False
		Me._Label1_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_1.Name = "_Label1_1"
		Me._Label1_0.Text = "Standard Name:"
		Me._Label1_0.Size = New System.Drawing.Size(95, 19)
		Me._Label1_0.Location = New System.Drawing.Point(14, 32)
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
		Me.Controls.Add(txtActiveCell)
		Me.Controls.Add(grdWQStd)
		Me.Controls.Add(cmdQuit)
		Me.Controls.Add(cmdSave)
		Me.Controls.Add(cboWQStdName)
		Me.Controls.Add(txtWQStdDesc)
		Me.Controls.Add(_Label1_1)
		Me.Controls.Add(_Label1_0)
		Me.Label1.SetIndex(_Label1_1, CType(1, Short))
		Me.Label1.SetIndex(_Label1_0, CType(0, Short))
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.grdWQStd, System.ComponentModel.ISupportInitialize).EndInit()
		MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuWQStd, Me.mnuEditCell, Me.mnuHelp})
		mnuWQStd.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuNewWQStd, Me.mnuDelWQStd, Me.mnuCopyWQStd, Me.mnuImpWQStd, Me.mnuExpWQStd})
		mnuEditCell.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuAddRow, Me.mnuInsertRow, Me.mnuDeleteRow})
		mnuHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuWQHelp})
		Me.Controls.Add(MainMenu1)
		Me.MainMenu1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class