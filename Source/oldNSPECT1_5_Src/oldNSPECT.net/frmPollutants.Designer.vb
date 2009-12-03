<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmPollutants
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
	Public WithEvents mnuAddPoll As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuDeletePoll As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuPoll_1 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuCoeffNewSet As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuCoeffCopySet As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuCoeffDeleteSet As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuCoeffImportSet As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuCoeffExportSet As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuCoeff As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPollHelp As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuCoeffHelp As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuHelp As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
	Public dlgColorOpen As System.Windows.Forms.OpenFileDialog
	Public dlgColorSave As System.Windows.Forms.SaveFileDialog
	Public dlgColorFont As System.Windows.Forms.FontDialog
	Public dlgColorColor As System.Windows.Forms.ColorDialog
	Public dlgColorPrint As System.Windows.Forms.PrintDialog
	Public WithEvents txtActiveCellWQStd As System.Windows.Forms.TextBox
	Public WithEvents txtActiveCell As System.Windows.Forms.TextBox
	Public WithEvents _Label1_5 As System.Windows.Forms.Label
	Public WithEvents _Label1_6 As System.Windows.Forms.Label
	Public WithEvents _Label1_7 As System.Windows.Forms.Label
	Public WithEvents _Label1_1 As System.Windows.Forms.Label
	Public WithEvents _Label1_2 As System.Windows.Forms.Label
	Public WithEvents _Label1_3 As System.Windows.Forms.Label
	Public WithEvents grdPollDef As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
	Public WithEvents cboCoeffSet As System.Windows.Forms.ComboBox
	Public WithEvents txtCoeffSetDesc As System.Windows.Forms.TextBox
	Public WithEvents txtLCType As System.Windows.Forms.TextBox
	Public WithEvents _SSTab1_TabPage0 As System.Windows.Forms.TabPage
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents grdWQStd As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
	Public WithEvents _SSTab1_TabPage1 As System.Windows.Forms.TabPage
	Public WithEvents SSTab1 As System.Windows.Forms.TabControl
	Public WithEvents cmdSave As System.Windows.Forms.Button
	Public WithEvents cmdQuit As System.Windows.Forms.Button
	Public WithEvents cboPollName As System.Windows.Forms.ComboBox
	Public dlgCMD1Open As System.Windows.Forms.OpenFileDialog
	Public dlgCMD1Save As System.Windows.Forms.SaveFileDialog
	Public WithEvents _Label1_0 As System.Windows.Forms.Label
	Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents mnuPoll As Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPollutants))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.MainMenu1 = New System.Windows.Forms.MenuStrip
		Me._mnuPoll_1 = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuAddPoll = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuDeletePoll = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuCoeff = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuCoeffNewSet = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuCoeffCopySet = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuCoeffDeleteSet = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuCoeffImportSet = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuCoeffExportSet = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuHelp = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPollHelp = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuCoeffHelp = New System.Windows.Forms.ToolStripMenuItem
		Me.dlgColorOpen = New System.Windows.Forms.OpenFileDialog
		Me.dlgColorSave = New System.Windows.Forms.SaveFileDialog
		Me.dlgColorFont = New System.Windows.Forms.FontDialog
		Me.dlgColorColor = New System.Windows.Forms.ColorDialog
		Me.dlgColorPrint = New System.Windows.Forms.PrintDialog
		Me.txtActiveCellWQStd = New System.Windows.Forms.TextBox
		Me.txtActiveCell = New System.Windows.Forms.TextBox
		Me.SSTab1 = New System.Windows.Forms.TabControl
		Me._SSTab1_TabPage0 = New System.Windows.Forms.TabPage
		Me._Label1_5 = New System.Windows.Forms.Label
		Me._Label1_6 = New System.Windows.Forms.Label
		Me._Label1_7 = New System.Windows.Forms.Label
		Me._Label1_1 = New System.Windows.Forms.Label
		Me._Label1_2 = New System.Windows.Forms.Label
		Me._Label1_3 = New System.Windows.Forms.Label
		Me.grdPollDef = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
		Me.cboCoeffSet = New System.Windows.Forms.ComboBox
		Me.txtCoeffSetDesc = New System.Windows.Forms.TextBox
		Me.txtLCType = New System.Windows.Forms.TextBox
		Me._SSTab1_TabPage1 = New System.Windows.Forms.TabPage
		Me.Label2 = New System.Windows.Forms.Label
		Me.grdWQStd = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
		Me.cmdSave = New System.Windows.Forms.Button
		Me.cmdQuit = New System.Windows.Forms.Button
		Me.cboPollName = New System.Windows.Forms.ComboBox
		Me.dlgCMD1Open = New System.Windows.Forms.OpenFileDialog
		Me.dlgCMD1Save = New System.Windows.Forms.SaveFileDialog
		Me._Label1_0 = New System.Windows.Forms.Label
		Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.mnuPoll = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(components)
		Me.MainMenu1.SuspendLayout()
		Me.SSTab1.SuspendLayout()
		Me._SSTab1_TabPage0.SuspendLayout()
		Me._SSTab1_TabPage1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.grdPollDef, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.grdWQStd, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.mnuPoll, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "Pollutants"
		Me.ClientSize = New System.Drawing.Size(561, 590)
		Me.Location = New System.Drawing.Point(113, 127)
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
		Me.Name = "frmPollutants"
		Me._mnuPoll_1.Name = "_mnuPoll_1"
		Me._mnuPoll_1.Text = "&Pollutants"
		Me._mnuPoll_1.Checked = False
		Me._mnuPoll_1.Enabled = True
		Me._mnuPoll_1.Visible = True
		Me.mnuAddPoll.Name = "mnuAddPoll"
		Me.mnuAddPoll.Text = "&Add..."
		Me.mnuAddPoll.Checked = False
		Me.mnuAddPoll.Enabled = True
		Me.mnuAddPoll.Visible = True
		Me.mnuDeletePoll.Name = "mnuDeletePoll"
		Me.mnuDeletePoll.Text = "&Delete..."
		Me.mnuDeletePoll.Checked = False
		Me.mnuDeletePoll.Enabled = True
		Me.mnuDeletePoll.Visible = True
		Me.mnuCoeff.Name = "mnuCoeff"
		Me.mnuCoeff.Text = "&Coefficients"
		Me.mnuCoeff.Checked = False
		Me.mnuCoeff.Enabled = True
		Me.mnuCoeff.Visible = True
		Me.mnuCoeffNewSet.Name = "mnuCoeffNewSet"
		Me.mnuCoeffNewSet.Text = "&New Set..."
		Me.mnuCoeffNewSet.Checked = False
		Me.mnuCoeffNewSet.Enabled = True
		Me.mnuCoeffNewSet.Visible = True
		Me.mnuCoeffCopySet.Name = "mnuCoeffCopySet"
		Me.mnuCoeffCopySet.Text = "&Copy Set..."
		Me.mnuCoeffCopySet.Checked = False
		Me.mnuCoeffCopySet.Enabled = True
		Me.mnuCoeffCopySet.Visible = True
		Me.mnuCoeffDeleteSet.Name = "mnuCoeffDeleteSet"
		Me.mnuCoeffDeleteSet.Text = "&Delete Set..."
		Me.mnuCoeffDeleteSet.Checked = False
		Me.mnuCoeffDeleteSet.Enabled = True
		Me.mnuCoeffDeleteSet.Visible = True
		Me.mnuCoeffImportSet.Name = "mnuCoeffImportSet"
		Me.mnuCoeffImportSet.Text = "&Import Set..."
		Me.mnuCoeffImportSet.Checked = False
		Me.mnuCoeffImportSet.Enabled = True
		Me.mnuCoeffImportSet.Visible = True
		Me.mnuCoeffExportSet.Name = "mnuCoeffExportSet"
		Me.mnuCoeffExportSet.Text = "&Export Set..."
		Me.mnuCoeffExportSet.Checked = False
		Me.mnuCoeffExportSet.Enabled = True
		Me.mnuCoeffExportSet.Visible = True
		Me.mnuHelp.Name = "mnuHelp"
		Me.mnuHelp.Text = "&Help"
		Me.mnuHelp.Checked = False
		Me.mnuHelp.Enabled = True
		Me.mnuHelp.Visible = True
		Me.mnuPollHelp.Name = "mnuPollHelp"
		Me.mnuPollHelp.Text = "Pollutants..."
		Me.mnuPollHelp.ShortcutKeys = CType(System.Windows.Forms.Keys.Shift or System.Windows.Forms.Keys.F1, System.Windows.Forms.Keys)
		Me.mnuPollHelp.Checked = False
		Me.mnuPollHelp.Enabled = True
		Me.mnuPollHelp.Visible = True
		Me.mnuCoeffHelp.Name = "mnuCoeffHelp"
		Me.mnuCoeffHelp.Text = "Coefficients..."
		Me.mnuCoeffHelp.ShortcutKeys = CType(System.Windows.Forms.Keys.Shift or System.Windows.Forms.Keys.F2, System.Windows.Forms.Keys)
		Me.mnuCoeffHelp.Checked = False
		Me.mnuCoeffHelp.Enabled = True
		Me.mnuCoeffHelp.Visible = True
		Me.txtActiveCellWQStd.AutoSize = False
		Me.txtActiveCellWQStd.BackColor = System.Drawing.Color.White
		Me.txtActiveCellWQStd.Size = New System.Drawing.Size(70, 16)
		Me.txtActiveCellWQStd.Location = New System.Drawing.Point(209, 555)
		Me.txtActiveCellWQStd.TabIndex = 15
		Me.txtActiveCellWQStd.Text = "Text1"
		Me.txtActiveCellWQStd.Visible = False
		Me.txtActiveCellWQStd.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtActiveCellWQStd.AcceptsReturn = True
		Me.txtActiveCellWQStd.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtActiveCellWQStd.CausesValidation = True
		Me.txtActiveCellWQStd.Enabled = True
		Me.txtActiveCellWQStd.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtActiveCellWQStd.HideSelection = True
		Me.txtActiveCellWQStd.ReadOnly = False
		Me.txtActiveCellWQStd.Maxlength = 0
		Me.txtActiveCellWQStd.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtActiveCellWQStd.MultiLine = False
		Me.txtActiveCellWQStd.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtActiveCellWQStd.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtActiveCellWQStd.TabStop = True
		Me.txtActiveCellWQStd.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.txtActiveCellWQStd.Name = "txtActiveCellWQStd"
		Me.txtActiveCell.AutoSize = False
		Me.txtActiveCell.BackColor = System.Drawing.Color.White
		Me.txtActiveCell.Size = New System.Drawing.Size(70, 16)
		Me.txtActiveCell.Location = New System.Drawing.Point(120, 555)
		Me.txtActiveCell.TabIndex = 14
		Me.txtActiveCell.Text = "Text1"
		Me.txtActiveCell.Visible = False
		Me.txtActiveCell.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtActiveCell.AcceptsReturn = True
		Me.txtActiveCell.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
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
		Me.SSTab1.Size = New System.Drawing.Size(533, 483)
		Me.SSTab1.Location = New System.Drawing.Point(13, 62)
		Me.SSTab1.TabIndex = 4
		Me.SSTab1.Appearance = System.Windows.Forms.TabAppearance.FlatButtons
		Me.SSTab1.ItemSize = New System.Drawing.Size(42, 18)
		Me.SSTab1.HotTrack = False
		Me.SSTab1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.SSTab1.Name = "SSTab1"
		Me._SSTab1_TabPage0.Text = "Coefficients"
		Me._Label1_5.Text = "Coefficient Set:"
		Me._Label1_5.Size = New System.Drawing.Size(97, 17)
		Me._Label1_5.Location = New System.Drawing.Point(17, 34)
		Me._Label1_5.TabIndex = 8
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
		Me._Label1_6.Text = "Description:"
		Me._Label1_6.Size = New System.Drawing.Size(74, 17)
		Me._Label1_6.Location = New System.Drawing.Point(18, 59)
		Me._Label1_6.TabIndex = 9
		Me._Label1_6.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_6.TextAlign = System.Drawing.ContentAlignment.TopLeft
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
		Me._Label1_7.Text = "Land Cover Type:"
		Me._Label1_7.Size = New System.Drawing.Size(97, 17)
		Me._Label1_7.Location = New System.Drawing.Point(271, 34)
		Me._Label1_7.TabIndex = 10
		Me._Label1_7.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_7.TextAlign = System.Drawing.ContentAlignment.TopLeft
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
		Me._Label1_1.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._Label1_1.Text = "Class"
		Me._Label1_1.ForeColor = System.Drawing.Color.Black
		Me._Label1_1.Size = New System.Drawing.Size(230, 17)
		Me._Label1_1.Location = New System.Drawing.Point(44, 86)
		Me._Label1_1.TabIndex = 11
		Me._Label1_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_1.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_1.Enabled = True
		Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_1.UseMnemonic = True
		Me._Label1_1.Visible = True
		Me._Label1_1.AutoSize = False
		Me._Label1_1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._Label1_1.Name = "_Label1_1"
		Me._Label1_2.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._Label1_2.Text = "Coefficients (mg/L)"
		Me._Label1_2.ForeColor = System.Drawing.Color.Black
		Me._Label1_2.Size = New System.Drawing.Size(224, 17)
		Me._Label1_2.Location = New System.Drawing.Point(276, 86)
		Me._Label1_2.TabIndex = 12
		Me._Label1_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_2.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_2.Enabled = True
		Me._Label1_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_2.UseMnemonic = True
		Me._Label1_2.Visible = True
		Me._Label1_2.AutoSize = False
		Me._Label1_2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._Label1_2.Name = "_Label1_2"
		Me._Label1_3.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._Label1_3.Size = New System.Drawing.Size(27, 17)
		Me._Label1_3.Location = New System.Drawing.Point(14, 86)
		Me._Label1_3.TabIndex = 13
		Me._Label1_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_3.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_3.Enabled = True
		Me._Label1_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_3.UseMnemonic = True
		Me._Label1_3.Visible = True
		Me._Label1_3.AutoSize = False
		Me._Label1_3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._Label1_3.Name = "_Label1_3"
		grdPollDef.OcxState = CType(resources.GetObject("grdPollDef.OcxState"), System.Windows.Forms.AxHost.State)
		Me.grdPollDef.Size = New System.Drawing.Size(509, 369)
		Me.grdPollDef.Location = New System.Drawing.Point(13, 104)
		Me.grdPollDef.TabIndex = 17
		Me.grdPollDef.Name = "grdPollDef"
		Me.cboCoeffSet.Size = New System.Drawing.Size(147, 21)
		Me.cboCoeffSet.Location = New System.Drawing.Point(100, 32)
		Me.cboCoeffSet.Sorted = True
		Me.cboCoeffSet.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cboCoeffSet.TabIndex = 5
		Me.cboCoeffSet.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboCoeffSet.BackColor = System.Drawing.SystemColors.Window
		Me.cboCoeffSet.CausesValidation = True
		Me.cboCoeffSet.Enabled = True
		Me.cboCoeffSet.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboCoeffSet.IntegralHeight = True
		Me.cboCoeffSet.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboCoeffSet.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboCoeffSet.TabStop = True
		Me.cboCoeffSet.Visible = True
		Me.cboCoeffSet.Name = "cboCoeffSet"
		Me.txtCoeffSetDesc.AutoSize = False
		Me.txtCoeffSetDesc.Size = New System.Drawing.Size(375, 19)
		Me.txtCoeffSetDesc.Location = New System.Drawing.Point(100, 60)
		Me.txtCoeffSetDesc.TabIndex = 6
		Me.txtCoeffSetDesc.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtCoeffSetDesc.AcceptsReturn = True
		Me.txtCoeffSetDesc.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtCoeffSetDesc.BackColor = System.Drawing.SystemColors.Window
		Me.txtCoeffSetDesc.CausesValidation = True
		Me.txtCoeffSetDesc.Enabled = True
		Me.txtCoeffSetDesc.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtCoeffSetDesc.HideSelection = True
		Me.txtCoeffSetDesc.ReadOnly = False
		Me.txtCoeffSetDesc.Maxlength = 0
		Me.txtCoeffSetDesc.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtCoeffSetDesc.MultiLine = False
		Me.txtCoeffSetDesc.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtCoeffSetDesc.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtCoeffSetDesc.TabStop = True
		Me.txtCoeffSetDesc.Visible = True
		Me.txtCoeffSetDesc.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtCoeffSetDesc.Name = "txtCoeffSetDesc"
		Me.txtLCType.AutoSize = False
		Me.txtLCType.Enabled = False
		Me.txtLCType.Size = New System.Drawing.Size(111, 19)
		Me.txtLCType.Location = New System.Drawing.Point(365, 32)
		Me.txtLCType.TabIndex = 7
		Me.txtLCType.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtLCType.AcceptsReturn = True
		Me.txtLCType.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtLCType.BackColor = System.Drawing.SystemColors.Window
		Me.txtLCType.CausesValidation = True
		Me.txtLCType.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtLCType.HideSelection = True
		Me.txtLCType.ReadOnly = False
		Me.txtLCType.Maxlength = 0
		Me.txtLCType.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtLCType.MultiLine = False
		Me.txtLCType.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtLCType.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtLCType.TabStop = True
		Me.txtLCType.Visible = True
		Me.txtLCType.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtLCType.Name = "txtLCType"
		Me._SSTab1_TabPage1.Text = "Water Quality Standards"
		Me.Label2.Text = "Threshold Units: ug/L"
		Me.Label2.Size = New System.Drawing.Size(137, 17)
		Me.Label2.Location = New System.Drawing.Point(32, 344)
		Me.Label2.TabIndex = 18
		Me.Label2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label2.BackColor = System.Drawing.SystemColors.Control
		Me.Label2.Enabled = True
		Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label2.UseMnemonic = True
		Me.Label2.Visible = True
		Me.Label2.AutoSize = False
		Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label2.Name = "Label2"
		grdWQStd.OcxState = CType(resources.GetObject("grdWQStd.OcxState"), System.Windows.Forms.AxHost.State)
		Me.grdWQStd.Size = New System.Drawing.Size(447, 288)
		Me.grdWQStd.Location = New System.Drawing.Point(24, 48)
		Me.grdWQStd.TabIndex = 16
		Me.grdWQStd.Name = "grdWQStd"
		Me.cmdSave.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdSave.Text = "OK"
		Me.cmdSave.Enabled = False
		Me.cmdSave.Size = New System.Drawing.Size(65, 25)
		Me.cmdSave.Location = New System.Drawing.Point(371, 553)
		Me.cmdSave.TabIndex = 2
		Me.cmdSave.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdSave.BackColor = System.Drawing.SystemColors.Control
		Me.cmdSave.CausesValidation = True
		Me.cmdSave.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdSave.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdSave.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdSave.TabStop = True
		Me.cmdSave.Name = "cmdSave"
		Me.cmdQuit.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdQuit.Text = "Cancel"
		Me.cmdQuit.Size = New System.Drawing.Size(65, 25)
		Me.cmdQuit.Location = New System.Drawing.Point(442, 553)
		Me.cmdQuit.TabIndex = 3
		Me.cmdQuit.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdQuit.BackColor = System.Drawing.SystemColors.Control
		Me.cmdQuit.CausesValidation = True
		Me.cmdQuit.Enabled = True
		Me.cmdQuit.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdQuit.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdQuit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdQuit.TabStop = True
		Me.cmdQuit.Name = "cmdQuit"
		Me.cboPollName.Size = New System.Drawing.Size(150, 21)
		Me.cboPollName.Location = New System.Drawing.Point(106, 34)
		Me.cboPollName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cboPollName.TabIndex = 1
		Me.cboPollName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboPollName.BackColor = System.Drawing.SystemColors.Window
		Me.cboPollName.CausesValidation = True
		Me.cboPollName.Enabled = True
		Me.cboPollName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboPollName.IntegralHeight = True
		Me.cboPollName.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboPollName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboPollName.Sorted = False
		Me.cboPollName.TabStop = True
		Me.cboPollName.Visible = True
		Me.cboPollName.Name = "cboPollName"
		Me._Label1_0.Text = "Pollutant Name:"
		Me._Label1_0.Size = New System.Drawing.Size(89, 17)
		Me._Label1_0.Location = New System.Drawing.Point(21, 35)
		Me._Label1_0.TabIndex = 0
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
		Me.Controls.Add(txtActiveCellWQStd)
		Me.Controls.Add(txtActiveCell)
		Me.Controls.Add(SSTab1)
		Me.Controls.Add(cmdSave)
		Me.Controls.Add(cmdQuit)
		Me.Controls.Add(cboPollName)
		Me.Controls.Add(_Label1_0)
		Me.SSTab1.Controls.Add(_SSTab1_TabPage0)
		Me.SSTab1.Controls.Add(_SSTab1_TabPage1)
		Me._SSTab1_TabPage0.Controls.Add(_Label1_5)
		Me._SSTab1_TabPage0.Controls.Add(_Label1_6)
		Me._SSTab1_TabPage0.Controls.Add(_Label1_7)
		Me._SSTab1_TabPage0.Controls.Add(_Label1_1)
		Me._SSTab1_TabPage0.Controls.Add(_Label1_2)
		Me._SSTab1_TabPage0.Controls.Add(_Label1_3)
		Me._SSTab1_TabPage0.Controls.Add(grdPollDef)
		Me._SSTab1_TabPage0.Controls.Add(cboCoeffSet)
		Me._SSTab1_TabPage0.Controls.Add(txtCoeffSetDesc)
		Me._SSTab1_TabPage0.Controls.Add(txtLCType)
		Me._SSTab1_TabPage1.Controls.Add(Label2)
		Me._SSTab1_TabPage1.Controls.Add(grdWQStd)
		Me.Label1.SetIndex(_Label1_3, CType(3, Short))
		Me.Label1.SetIndex(_Label1_2, CType(2, Short))
		Me.Label1.SetIndex(_Label1_1, CType(1, Short))
		Me.Label1.SetIndex(_Label1_7, CType(7, Short))
		Me.Label1.SetIndex(_Label1_6, CType(6, Short))
		Me.Label1.SetIndex(_Label1_5, CType(5, Short))
		Me.Label1.SetIndex(_Label1_0, CType(0, Short))
		Me.mnuPoll.SetIndex(_mnuPoll_1, CType(1, Short))
		CType(Me.mnuPoll, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.grdWQStd, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.grdPollDef, System.ComponentModel.ISupportInitialize).EndInit()
		MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem(){Me._mnuPoll_1, Me.mnuCoeff, Me.mnuHelp})
		_mnuPoll_1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuAddPoll, Me.mnuDeletePoll})
		mnuCoeff.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuCoeffNewSet, Me.mnuCoeffCopySet, Me.mnuCoeffDeleteSet, Me.mnuCoeffImportSet, Me.mnuCoeffExportSet})
		mnuHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuPollHelp, Me.mnuCoeffHelp})
		Me.Controls.Add(MainMenu1)
		Me.MainMenu1.ResumeLayout(False)
		Me.SSTab1.ResumeLayout(False)
		Me._SSTab1_TabPage0.ResumeLayout(False)
		Me._SSTab1_TabPage1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class