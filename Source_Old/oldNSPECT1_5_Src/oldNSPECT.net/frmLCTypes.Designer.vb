<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmLCTypes
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
	Public WithEvents mnuNewLCType As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuDelLCType As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuImpLCType As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuExpLCType As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuLCTypes As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuAppend As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuInsertRow As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuDeleteRow As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopUp As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuLCHelp As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuHelp As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
	Public WithEvents Timer1 As System.Windows.Forms.Timer
	Public WithEvents _chkWWL_0 As System.Windows.Forms.CheckBox
	Public WithEvents cboWWL As System.Windows.Forms.ComboBox
	Public WithEvents txtActiveCell As System.Windows.Forms.TextBox
	Public WithEvents grdLCClasses As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
	Public dlgCMD1Open As System.Windows.Forms.OpenFileDialog
	Public dlgCMD1Save As System.Windows.Forms.SaveFileDialog
	Public WithEvents cmdRestore As System.Windows.Forms.Button
	Public WithEvents cmdQuit As System.Windows.Forms.Button
	Public WithEvents cmdSave As System.Windows.Forms.Button
	Public WithEvents txtLCTypeDesc As System.Windows.Forms.TextBox
	Public WithEvents cboLCType As System.Windows.Forms.ComboBox
	Public WithEvents _Label2_2 As System.Windows.Forms.Label
	Public WithEvents _Label2_1 As System.Windows.Forms.Label
	Public WithEvents _Label2_0 As System.Windows.Forms.Label
	Public WithEvents _Label1_6 As System.Windows.Forms.Label
	Public WithEvents _Label1_7 As System.Windows.Forms.Label
	Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents Label2 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents chkWWL As Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmLCTypes))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.MainMenu1 = New System.Windows.Forms.MenuStrip
		Me.mnuLCTypes = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuNewLCType = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuDelLCType = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuImpLCType = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuExpLCType = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopUp = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuAppend = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuInsertRow = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuDeleteRow = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuHelp = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuLCHelp = New System.Windows.Forms.ToolStripMenuItem
		Me.Timer1 = New System.Windows.Forms.Timer(components)
		Me._chkWWL_0 = New System.Windows.Forms.CheckBox
		Me.cboWWL = New System.Windows.Forms.ComboBox
		Me.txtActiveCell = New System.Windows.Forms.TextBox
		Me.grdLCClasses = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
		Me.dlgCMD1Open = New System.Windows.Forms.OpenFileDialog
		Me.dlgCMD1Save = New System.Windows.Forms.SaveFileDialog
		Me.cmdRestore = New System.Windows.Forms.Button
		Me.cmdQuit = New System.Windows.Forms.Button
		Me.cmdSave = New System.Windows.Forms.Button
		Me.txtLCTypeDesc = New System.Windows.Forms.TextBox
		Me.cboLCType = New System.Windows.Forms.ComboBox
		Me._Label2_2 = New System.Windows.Forms.Label
		Me._Label2_1 = New System.Windows.Forms.Label
		Me._Label2_0 = New System.Windows.Forms.Label
		Me._Label1_6 = New System.Windows.Forms.Label
		Me._Label1_7 = New System.Windows.Forms.Label
		Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.Label2 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.chkWWL = New Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray(components)
		Me.MainMenu1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.grdLCClasses, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Label2, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.chkWWL, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "Land Cover Types"
		Me.ClientSize = New System.Drawing.Size(618, 590)
		Me.Location = New System.Drawing.Point(268, 71)
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
		Me.Name = "frmLCTypes"
		Me.mnuLCTypes.Name = "mnuLCTypes"
		Me.mnuLCTypes.Text = "&Options"
		Me.mnuLCTypes.Checked = False
		Me.mnuLCTypes.Enabled = True
		Me.mnuLCTypes.Visible = True
		Me.mnuNewLCType.Name = "mnuNewLCType"
		Me.mnuNewLCType.Text = "&New..."
		Me.mnuNewLCType.Checked = False
		Me.mnuNewLCType.Enabled = True
		Me.mnuNewLCType.Visible = True
		Me.mnuDelLCType.Name = "mnuDelLCType"
		Me.mnuDelLCType.Text = "&Delete..."
		Me.mnuDelLCType.Checked = False
		Me.mnuDelLCType.Enabled = True
		Me.mnuDelLCType.Visible = True
		Me.mnuImpLCType.Name = "mnuImpLCType"
		Me.mnuImpLCType.Text = "&Import..."
		Me.mnuImpLCType.Checked = False
		Me.mnuImpLCType.Enabled = True
		Me.mnuImpLCType.Visible = True
		Me.mnuExpLCType.Name = "mnuExpLCType"
		Me.mnuExpLCType.Text = "&Export..."
		Me.mnuExpLCType.Checked = False
		Me.mnuExpLCType.Enabled = True
		Me.mnuExpLCType.Visible = True
		Me.mnuPopUp.Name = "mnuPopUp"
		Me.mnuPopUp.Text = "Edit"
		Me.mnuPopUp.Visible = False
		Me.mnuPopUp.Checked = False
		Me.mnuPopUp.Enabled = True
		Me.mnuAppend.Name = "mnuAppend"
		Me.mnuAppend.Text = "Add Row"
		Me.mnuAppend.Checked = False
		Me.mnuAppend.Enabled = True
		Me.mnuAppend.Visible = True
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
		Me.mnuLCHelp.Name = "mnuLCHelp"
		Me.mnuLCHelp.Text = "Land Cover Types..."
		Me.mnuLCHelp.ShortcutKeys = CType(System.Windows.Forms.Keys.Shift or System.Windows.Forms.Keys.F1, System.Windows.Forms.Keys)
		Me.mnuLCHelp.Checked = False
		Me.mnuLCHelp.Enabled = True
		Me.mnuLCHelp.Visible = True
		Me.Timer1.Enabled = False
		Me.Timer1.Interval = 1
		Me._chkWWL_0.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
		Me._chkWWL_0.FlatStyle = System.Windows.Forms.FlatStyle.Flat
		Me._chkWWL_0.BackColor = System.Drawing.SystemColors.Window
		Me._chkWWL_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._chkWWL_0.Size = New System.Drawing.Size(13, 13)
		Me._chkWWL_0.Location = New System.Drawing.Point(207, 553)
		Me._chkWWL_0.TabIndex = 13
		Me.ToolTip1.SetToolTip(Me._chkWWL_0, "Check if landuse is water or wetland")
		Me._chkWWL_0.Visible = False
		Me._chkWWL_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkWWL_0.Text = ""
		Me._chkWWL_0.CausesValidation = True
		Me._chkWWL_0.Enabled = True
		Me._chkWWL_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkWWL_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkWWL_0.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkWWL_0.TabStop = True
		Me._chkWWL_0.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._chkWWL_0.Name = "_chkWWL_0"
		Me.cboWWL.Size = New System.Drawing.Size(45, 21)
		Me.cboWWL.Location = New System.Drawing.Point(370, 559)
		Me.cboWWL.Items.AddRange(New Object(){"Y", "N"})
		Me.cboWWL.TabIndex = 12
		Me.cboWWL.Visible = False
		Me.cboWWL.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboWWL.BackColor = System.Drawing.SystemColors.Window
		Me.cboWWL.CausesValidation = True
		Me.cboWWL.Enabled = True
		Me.cboWWL.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboWWL.IntegralHeight = True
		Me.cboWWL.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboWWL.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboWWL.Sorted = False
		Me.cboWWL.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboWWL.TabStop = True
		Me.cboWWL.Name = "cboWWL"
		Me.txtActiveCell.AutoSize = False
		Me.txtActiveCell.BackColor = System.Drawing.Color.White
		Me.txtActiveCell.Size = New System.Drawing.Size(70, 16)
		Me.txtActiveCell.HideSelection = False
		Me.txtActiveCell.Location = New System.Drawing.Point(268, 561)
		Me.txtActiveCell.TabIndex = 11
		Me.txtActiveCell.Text = "Text1"
		Me.txtActiveCell.Visible = False
		Me.txtActiveCell.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtActiveCell.AcceptsReturn = True
		Me.txtActiveCell.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtActiveCell.CausesValidation = True
		Me.txtActiveCell.Enabled = True
		Me.txtActiveCell.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtActiveCell.ReadOnly = False
		Me.txtActiveCell.Maxlength = 0
		Me.txtActiveCell.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtActiveCell.MultiLine = False
		Me.txtActiveCell.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtActiveCell.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtActiveCell.TabStop = True
		Me.txtActiveCell.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.txtActiveCell.Name = "txtActiveCell"
		grdLCClasses.OcxState = CType(resources.GetObject("grdLCClasses.OcxState"), System.Windows.Forms.AxHost.State)
		Me.grdLCClasses.Size = New System.Drawing.Size(590, 445)
		Me.grdLCClasses.Location = New System.Drawing.Point(7, 100)
		Me.grdLCClasses.TabIndex = 10
		Me.grdLCClasses.Name = "grdLCClasses"
		Me.cmdRestore.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdRestore.Text = "Restore Defaults"
		Me.cmdRestore.Enabled = False
		Me.cmdRestore.Size = New System.Drawing.Size(103, 25)
		Me.cmdRestore.Location = New System.Drawing.Point(25, 551)
		Me.cmdRestore.TabIndex = 6
		Me.cmdRestore.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdRestore.BackColor = System.Drawing.SystemColors.Control
		Me.cmdRestore.CausesValidation = True
		Me.cmdRestore.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdRestore.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdRestore.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdRestore.TabStop = True
		Me.cmdRestore.Name = "cmdRestore"
		Me.cmdQuit.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdQuit.Text = "Cancel"
		Me.cmdQuit.Size = New System.Drawing.Size(65, 25)
		Me.cmdQuit.Location = New System.Drawing.Point(512, 550)
		Me.cmdQuit.TabIndex = 5
		Me.cmdQuit.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdQuit.BackColor = System.Drawing.SystemColors.Control
		Me.cmdQuit.CausesValidation = True
		Me.cmdQuit.Enabled = True
		Me.cmdQuit.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdQuit.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdQuit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdQuit.TabStop = True
		Me.cmdQuit.Name = "cmdQuit"
		Me.cmdSave.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdSave.Text = "Save"
		Me.cmdSave.Enabled = False
		Me.cmdSave.Size = New System.Drawing.Size(65, 25)
		Me.cmdSave.Location = New System.Drawing.Point(441, 550)
		Me.cmdSave.TabIndex = 4
		Me.cmdSave.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdSave.BackColor = System.Drawing.SystemColors.Control
		Me.cmdSave.CausesValidation = True
		Me.cmdSave.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdSave.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdSave.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdSave.TabStop = True
		Me.cmdSave.Name = "cmdSave"
		Me.txtLCTypeDesc.AutoSize = False
		Me.txtLCTypeDesc.Size = New System.Drawing.Size(362, 19)
		Me.txtLCTypeDesc.Location = New System.Drawing.Point(124, 52)
		Me.txtLCTypeDesc.TabIndex = 2
		Me.txtLCTypeDesc.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtLCTypeDesc.AcceptsReturn = True
		Me.txtLCTypeDesc.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtLCTypeDesc.BackColor = System.Drawing.SystemColors.Window
		Me.txtLCTypeDesc.CausesValidation = True
		Me.txtLCTypeDesc.Enabled = True
		Me.txtLCTypeDesc.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtLCTypeDesc.HideSelection = True
		Me.txtLCTypeDesc.ReadOnly = False
		Me.txtLCTypeDesc.Maxlength = 0
		Me.txtLCTypeDesc.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtLCTypeDesc.MultiLine = False
		Me.txtLCTypeDesc.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtLCTypeDesc.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtLCTypeDesc.TabStop = True
		Me.txtLCTypeDesc.Visible = True
		Me.txtLCTypeDesc.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtLCTypeDesc.Name = "txtLCTypeDesc"
		Me.cboLCType.Size = New System.Drawing.Size(149, 21)
		Me.cboLCType.Location = New System.Drawing.Point(124, 26)
		Me.cboLCType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cboLCType.TabIndex = 0
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
		Me._Label2_2.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._Label2_2.Text = "RUSLE"
		Me._Label2_2.Size = New System.Drawing.Size(160, 19)
		Me._Label2_2.Location = New System.Drawing.Point(451, 78)
		Me._Label2_2.TabIndex = 9
		Me._Label2_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label2_2.BackColor = System.Drawing.SystemColors.Control
		Me._Label2_2.Enabled = True
		Me._Label2_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label2_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label2_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label2_2.UseMnemonic = True
		Me._Label2_2.Visible = True
		Me._Label2_2.AutoSize = False
		Me._Label2_2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._Label2_2.Name = "_Label2_2"
		Me._Label2_1.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._Label2_1.Text = "Classification"
		Me._Label2_1.Size = New System.Drawing.Size(227, 19)
		Me._Label2_1.Location = New System.Drawing.Point(7, 78)
		Me._Label2_1.TabIndex = 8
		Me._Label2_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label2_1.BackColor = System.Drawing.SystemColors.Control
		Me._Label2_1.Enabled = True
		Me._Label2_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label2_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label2_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label2_1.UseMnemonic = True
		Me._Label2_1.Visible = True
		Me._Label2_1.AutoSize = False
		Me._Label2_1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._Label2_1.Name = "_Label2_1"
		Me._Label2_0.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._Label2_0.Text = "SCS Curve Numbers"
		Me._Label2_0.Size = New System.Drawing.Size(213, 19)
		Me._Label2_0.Location = New System.Drawing.Point(236, 78)
		Me._Label2_0.TabIndex = 7
		Me._Label2_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label2_0.BackColor = System.Drawing.SystemColors.Control
		Me._Label2_0.Enabled = True
		Me._Label2_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label2_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label2_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label2_0.UseMnemonic = True
		Me._Label2_0.Visible = True
		Me._Label2_0.AutoSize = False
		Me._Label2_0.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._Label2_0.Name = "_Label2_0"
		Me._Label1_6.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_6.Text = "Description:"
		Me._Label1_6.Size = New System.Drawing.Size(86, 17)
		Me._Label1_6.Location = New System.Drawing.Point(21, 51)
		Me._Label1_6.TabIndex = 3
		Me._Label1_6.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_6.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_6.Enabled = True
		Me._Label1_6.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_6.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_6.UseMnemonic = True
		Me._Label1_6.Visible = True
		Me._Label1_6.AutoSize = False
		Me._Label1_6.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_6.Name = "_Label1_6"
		Me._Label1_7.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_7.Text = "Land Cover Type:"
		Me._Label1_7.Size = New System.Drawing.Size(91, 17)
		Me._Label1_7.Location = New System.Drawing.Point(16, 29)
		Me._Label1_7.TabIndex = 1
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
		Me.Controls.Add(_chkWWL_0)
		Me.Controls.Add(cboWWL)
		Me.Controls.Add(txtActiveCell)
		Me.Controls.Add(grdLCClasses)
		Me.Controls.Add(cmdRestore)
		Me.Controls.Add(cmdQuit)
		Me.Controls.Add(cmdSave)
		Me.Controls.Add(txtLCTypeDesc)
		Me.Controls.Add(cboLCType)
		Me.Controls.Add(_Label2_2)
		Me.Controls.Add(_Label2_1)
		Me.Controls.Add(_Label2_0)
		Me.Controls.Add(_Label1_6)
		Me.Controls.Add(_Label1_7)
		Me.Label1.SetIndex(_Label1_6, CType(6, Short))
		Me.Label1.SetIndex(_Label1_7, CType(7, Short))
		Me.Label2.SetIndex(_Label2_2, CType(2, Short))
		Me.Label2.SetIndex(_Label2_1, CType(1, Short))
		Me.Label2.SetIndex(_Label2_0, CType(0, Short))
		Me.chkWWL.SetIndex(_chkWWL_0, CType(0, Short))
		CType(Me.chkWWL, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Label2, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.grdLCClasses, System.ComponentModel.ISupportInitialize).EndInit()
		MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuLCTypes, Me.mnuPopUp, Me.mnuHelp})
		mnuLCTypes.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuNewLCType, Me.mnuDelLCType, Me.mnuImpLCType, Me.mnuExpLCType})
		mnuPopUp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuAppend, Me.mnuInsertRow, Me.mnuDeleteRow})
		mnuHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuLCHelp})
		Me.Controls.Add(MainMenu1)
		Me.MainMenu1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class