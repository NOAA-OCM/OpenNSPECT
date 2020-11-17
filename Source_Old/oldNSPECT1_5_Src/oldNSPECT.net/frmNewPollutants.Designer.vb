<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmNewPollutants
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
	Public WithEvents mnuCoeffNewSet As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuCoeffCopySet As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuCoeff As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
	Public WithEvents txtPollutant As System.Windows.Forms.TextBox
	Public WithEvents txtActiveCellWQStd As System.Windows.Forms.TextBox
	Public WithEvents txtActiveCell As System.Windows.Forms.TextBox
	Public WithEvents _Label1_6 As System.Windows.Forms.Label
	Public WithEvents _Label1_1 As System.Windows.Forms.Label
	Public WithEvents _Label1_2 As System.Windows.Forms.Label
	Public WithEvents _Label1_3 As System.Windows.Forms.Label
	Public WithEvents _Label1_7 As System.Windows.Forms.Label
	Public WithEvents _Label1_5 As System.Windows.Forms.Label
	Public WithEvents grdPollDef As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
	Public WithEvents cboLCType As System.Windows.Forms.ComboBox
	Public WithEvents txtCoeffSetDesc As System.Windows.Forms.TextBox
	Public WithEvents txtCoeffSet As System.Windows.Forms.TextBox
	Public WithEvents _SSTab1_TabPage0 As System.Windows.Forms.TabPage
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents grdWQStd As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
	Public WithEvents _SSTab1_TabPage1 As System.Windows.Forms.TabPage
	Public WithEvents SSTab1 As System.Windows.Forms.TabControl
	Public WithEvents cmdSave As System.Windows.Forms.Button
	Public WithEvents cmdQuit As System.Windows.Forms.Button
	Public dlgCMD1Open As System.Windows.Forms.OpenFileDialog
	Public dlgCMD1Save As System.Windows.Forms.SaveFileDialog
	Public dlgCMD1Font As System.Windows.Forms.FontDialog
	Public dlgCMD1Color As System.Windows.Forms.ColorDialog
	Public dlgCMD1Print As System.Windows.Forms.PrintDialog
	Public WithEvents _Label1_0 As System.Windows.Forms.Label
	Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmNewPollutants))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.MainMenu1 = New System.Windows.Forms.MenuStrip
		Me.mnuCoeff = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuCoeffNewSet = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuCoeffCopySet = New System.Windows.Forms.ToolStripMenuItem
		Me.txtPollutant = New System.Windows.Forms.TextBox
		Me.txtActiveCellWQStd = New System.Windows.Forms.TextBox
		Me.txtActiveCell = New System.Windows.Forms.TextBox
		Me.SSTab1 = New System.Windows.Forms.TabControl
		Me._SSTab1_TabPage0 = New System.Windows.Forms.TabPage
		Me._Label1_6 = New System.Windows.Forms.Label
		Me._Label1_1 = New System.Windows.Forms.Label
		Me._Label1_2 = New System.Windows.Forms.Label
		Me._Label1_3 = New System.Windows.Forms.Label
		Me._Label1_7 = New System.Windows.Forms.Label
		Me._Label1_5 = New System.Windows.Forms.Label
		Me.grdPollDef = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
		Me.cboLCType = New System.Windows.Forms.ComboBox
		Me.txtCoeffSetDesc = New System.Windows.Forms.TextBox
		Me.txtCoeffSet = New System.Windows.Forms.TextBox
		Me._SSTab1_TabPage1 = New System.Windows.Forms.TabPage
		Me.Label2 = New System.Windows.Forms.Label
		Me.grdWQStd = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
		Me.cmdSave = New System.Windows.Forms.Button
		Me.cmdQuit = New System.Windows.Forms.Button
		Me.dlgCMD1Open = New System.Windows.Forms.OpenFileDialog
		Me.dlgCMD1Save = New System.Windows.Forms.SaveFileDialog
		Me.dlgCMD1Font = New System.Windows.Forms.FontDialog
		Me.dlgCMD1Color = New System.Windows.Forms.ColorDialog
		Me.dlgCMD1Print = New System.Windows.Forms.PrintDialog
		Me._Label1_0 = New System.Windows.Forms.Label
		Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.MainMenu1.SuspendLayout()
		Me.SSTab1.SuspendLayout()
		Me._SSTab1_TabPage0.SuspendLayout()
		Me._SSTab1_TabPage1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.grdPollDef, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.grdWQStd, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "Add Pollutant"
		Me.ClientSize = New System.Drawing.Size(561, 590)
		Me.Location = New System.Drawing.Point(268, 127)
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
		Me.Name = "frmNewPollutants"
		Me.mnuCoeff.Name = "mnuCoeff"
		Me.mnuCoeff.Text = "&Coefficients"
		Me.mnuCoeff.Checked = False
		Me.mnuCoeff.Enabled = True
		Me.mnuCoeff.Visible = True
		Me.mnuCoeffNewSet.Name = "mnuCoeffNewSet"
		Me.mnuCoeffNewSet.Text = "&Add Coefficient Set..."
		Me.mnuCoeffNewSet.Checked = False
		Me.mnuCoeffNewSet.Enabled = True
		Me.mnuCoeffNewSet.Visible = True
		Me.mnuCoeffCopySet.Name = "mnuCoeffCopySet"
		Me.mnuCoeffCopySet.Text = "&Copy Coefficient Set..."
		Me.mnuCoeffCopySet.Checked = False
		Me.mnuCoeffCopySet.Enabled = True
		Me.mnuCoeffCopySet.Visible = True
		Me.txtPollutant.AutoSize = False
		Me.txtPollutant.Size = New System.Drawing.Size(134, 20)
		Me.txtPollutant.Location = New System.Drawing.Point(106, 26)
		Me.txtPollutant.TabIndex = 0
		Me.txtPollutant.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtPollutant.AcceptsReturn = True
		Me.txtPollutant.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtPollutant.BackColor = System.Drawing.SystemColors.Window
		Me.txtPollutant.CausesValidation = True
		Me.txtPollutant.Enabled = True
		Me.txtPollutant.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtPollutant.HideSelection = True
		Me.txtPollutant.ReadOnly = False
		Me.txtPollutant.Maxlength = 0
		Me.txtPollutant.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtPollutant.MultiLine = False
		Me.txtPollutant.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtPollutant.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtPollutant.TabStop = True
		Me.txtPollutant.Visible = True
		Me.txtPollutant.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtPollutant.Name = "txtPollutant"
		Me.txtActiveCellWQStd.AutoSize = False
		Me.txtActiveCellWQStd.BackColor = System.Drawing.Color.White
		Me.txtActiveCellWQStd.Size = New System.Drawing.Size(70, 16)
		Me.txtActiveCellWQStd.Location = New System.Drawing.Point(25, 565)
		Me.txtActiveCellWQStd.TabIndex = 14
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
		Me.txtActiveCell.Location = New System.Drawing.Point(26, 543)
		Me.txtActiveCell.TabIndex = 13
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
		Me.SSTab1.Size = New System.Drawing.Size(521, 482)
		Me.SSTab1.Location = New System.Drawing.Point(16, 53)
		Me.SSTab1.TabIndex = 8
		Me.SSTab1.Appearance = System.Windows.Forms.TabAppearance.FlatButtons
		Me.SSTab1.SelectedIndex = 1
		Me.SSTab1.ItemSize = New System.Drawing.Size(42, 18)
		Me.SSTab1.HotTrack = False
		Me.SSTab1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.SSTab1.Name = "SSTab1"
		Me._SSTab1_TabPage0.Text = "Coefficients"
		Me._Label1_6.Text = "Description:"
		Me._Label1_6.Size = New System.Drawing.Size(74, 17)
		Me._Label1_6.Location = New System.Drawing.Point(18, 63)
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
		Me._Label1_1.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._Label1_1.Text = "Class"
		Me._Label1_1.ForeColor = System.Drawing.Color.Black
		Me._Label1_1.Size = New System.Drawing.Size(230, 17)
		Me._Label1_1.Location = New System.Drawing.Point(45, 88)
		Me._Label1_1.TabIndex = 10
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
		Me._Label1_2.Size = New System.Drawing.Size(226, 17)
		Me._Label1_2.Location = New System.Drawing.Point(278, 88)
		Me._Label1_2.TabIndex = 11
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
		Me._Label1_3.Location = New System.Drawing.Point(15, 88)
		Me._Label1_3.TabIndex = 12
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
		Me._Label1_7.BackColor = System.Drawing.Color.Transparent
		Me._Label1_7.Text = "Land Cover Type:"
		Me._Label1_7.Size = New System.Drawing.Size(97, 17)
		Me._Label1_7.Location = New System.Drawing.Point(257, 35)
		Me._Label1_7.TabIndex = 16
		Me._Label1_7.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_7.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_7.Enabled = True
		Me._Label1_7.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_7.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_7.UseMnemonic = True
		Me._Label1_7.Visible = True
		Me._Label1_7.AutoSize = False
		Me._Label1_7.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_7.Name = "_Label1_7"
		Me._Label1_5.BackColor = System.Drawing.Color.Transparent
		Me._Label1_5.Text = "Coefficient Set:"
		Me._Label1_5.Size = New System.Drawing.Size(78, 17)
		Me._Label1_5.Location = New System.Drawing.Point(18, 37)
		Me._Label1_5.TabIndex = 17
		Me._Label1_5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_5.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_5.Enabled = True
		Me._Label1_5.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_5.UseMnemonic = True
		Me._Label1_5.Visible = True
		Me._Label1_5.AutoSize = False
		Me._Label1_5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_5.Name = "_Label1_5"
		grdPollDef.OcxState = CType(resources.GetObject("grdPollDef.OcxState"), System.Windows.Forms.AxHost.State)
		Me.grdPollDef.Size = New System.Drawing.Size(491, 356)
		Me.grdPollDef.Location = New System.Drawing.Point(13, 109)
		Me.grdPollDef.TabIndex = 4
		Me.grdPollDef.Name = "grdPollDef"
		Me.cboLCType.Size = New System.Drawing.Size(147, 21)
		Me.cboLCType.Location = New System.Drawing.Point(350, 32)
		Me.cboLCType.Sorted = True
		Me.cboLCType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cboLCType.TabIndex = 2
		Me.cboLCType.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboLCType.BackColor = System.Drawing.SystemColors.Window
		Me.cboLCType.CausesValidation = True
		Me.cboLCType.Enabled = True
		Me.cboLCType.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboLCType.IntegralHeight = True
		Me.cboLCType.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboLCType.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboLCType.TabStop = True
		Me.cboLCType.Visible = True
		Me.cboLCType.Name = "cboLCType"
		Me.txtCoeffSetDesc.AutoSize = False
		Me.txtCoeffSetDesc.Size = New System.Drawing.Size(399, 19)
		Me.txtCoeffSetDesc.Location = New System.Drawing.Point(100, 62)
		Me.txtCoeffSetDesc.TabIndex = 3
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
		Me.txtCoeffSet.AutoSize = False
		Me.txtCoeffSet.Size = New System.Drawing.Size(134, 19)
		Me.txtCoeffSet.Location = New System.Drawing.Point(100, 35)
		Me.txtCoeffSet.TabIndex = 1
		Me.txtCoeffSet.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtCoeffSet.AcceptsReturn = True
		Me.txtCoeffSet.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtCoeffSet.BackColor = System.Drawing.SystemColors.Window
		Me.txtCoeffSet.CausesValidation = True
		Me.txtCoeffSet.Enabled = True
		Me.txtCoeffSet.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtCoeffSet.HideSelection = True
		Me.txtCoeffSet.ReadOnly = False
		Me.txtCoeffSet.Maxlength = 0
		Me.txtCoeffSet.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtCoeffSet.MultiLine = False
		Me.txtCoeffSet.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtCoeffSet.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtCoeffSet.TabStop = True
		Me.txtCoeffSet.Visible = True
		Me.txtCoeffSet.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtCoeffSet.Name = "txtCoeffSet"
		Me._SSTab1_TabPage1.Text = "Water Quality Standards"
		Me.Label2.Text = "Threshold Units: ug/L"
		Me.Label2.Size = New System.Drawing.Size(145, 17)
		Me.Label2.Location = New System.Drawing.Point(24, 336)
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
		Me.grdWQStd.Size = New System.Drawing.Size(465, 285)
		Me.grdWQStd.Location = New System.Drawing.Point(23, 42)
		Me.grdWQStd.TabIndex = 15
		Me.grdWQStd.Name = "grdWQStd"
		Me.cmdSave.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdSave.Text = "OK"
		Me.cmdSave.Enabled = False
		Me.cmdSave.Size = New System.Drawing.Size(65, 25)
		Me.cmdSave.Location = New System.Drawing.Point(378, 553)
		Me.cmdSave.TabIndex = 6
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
		Me.cmdQuit.Location = New System.Drawing.Point(448, 553)
		Me.cmdQuit.TabIndex = 7
		Me.cmdQuit.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdQuit.BackColor = System.Drawing.SystemColors.Control
		Me.cmdQuit.CausesValidation = True
		Me.cmdQuit.Enabled = True
		Me.cmdQuit.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdQuit.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdQuit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdQuit.TabStop = True
		Me.cmdQuit.Name = "cmdQuit"
		Me._Label1_0.Text = "Pollutant Name:"
		Me._Label1_0.Size = New System.Drawing.Size(86, 17)
		Me._Label1_0.Location = New System.Drawing.Point(21, 30)
		Me._Label1_0.TabIndex = 5
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
		Me.Controls.Add(txtPollutant)
		Me.Controls.Add(txtActiveCellWQStd)
		Me.Controls.Add(txtActiveCell)
		Me.Controls.Add(SSTab1)
		Me.Controls.Add(cmdSave)
		Me.Controls.Add(cmdQuit)
		Me.Controls.Add(_Label1_0)
		Me.SSTab1.Controls.Add(_SSTab1_TabPage0)
		Me.SSTab1.Controls.Add(_SSTab1_TabPage1)
		Me._SSTab1_TabPage0.Controls.Add(_Label1_6)
		Me._SSTab1_TabPage0.Controls.Add(_Label1_1)
		Me._SSTab1_TabPage0.Controls.Add(_Label1_2)
		Me._SSTab1_TabPage0.Controls.Add(_Label1_3)
		Me._SSTab1_TabPage0.Controls.Add(_Label1_7)
		Me._SSTab1_TabPage0.Controls.Add(_Label1_5)
		Me._SSTab1_TabPage0.Controls.Add(grdPollDef)
		Me._SSTab1_TabPage0.Controls.Add(cboLCType)
		Me._SSTab1_TabPage0.Controls.Add(txtCoeffSetDesc)
		Me._SSTab1_TabPage0.Controls.Add(txtCoeffSet)
		Me._SSTab1_TabPage1.Controls.Add(Label2)
		Me._SSTab1_TabPage1.Controls.Add(grdWQStd)
		Me.Label1.SetIndex(_Label1_5, CType(5, Short))
		Me.Label1.SetIndex(_Label1_7, CType(7, Short))
		Me.Label1.SetIndex(_Label1_3, CType(3, Short))
		Me.Label1.SetIndex(_Label1_2, CType(2, Short))
		Me.Label1.SetIndex(_Label1_1, CType(1, Short))
		Me.Label1.SetIndex(_Label1_6, CType(6, Short))
		Me.Label1.SetIndex(_Label1_0, CType(0, Short))
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.grdWQStd, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.grdPollDef, System.ComponentModel.ISupportInitialize).EndInit()
		MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuCoeff})
		mnuCoeff.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuCoeffNewSet, Me.mnuCoeffCopySet})
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