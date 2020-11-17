<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmWSDelin
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
	Public WithEvents mnuNewWSDelin As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuNewExist As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuDelWSDelin As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuDefWSDelin As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuWSDelin As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuHelp As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
	Public WithEvents txtLSGrid As System.Windows.Forms.TextBox
	Public WithEvents cboDEMUnits As System.Windows.Forms.ComboBox
	Public WithEvents cboWSSize As System.Windows.Forms.ComboBox
	Public WithEvents chkHydroCorr As System.Windows.Forms.CheckBox
	Public WithEvents txtFlowAccumGrid As System.Windows.Forms.TextBox
	Public WithEvents txtWSFile As System.Windows.Forms.TextBox
	Public WithEvents txtStream As System.Windows.Forms.TextBox
	Public WithEvents cboWSDelin As System.Windows.Forms.ComboBox
	Public WithEvents txtDEMFile As System.Windows.Forms.TextBox
	Public WithEvents _Label1_6 As System.Windows.Forms.Label
	Public WithEvents _Label1_12 As System.Windows.Forms.Label
	Public WithEvents _Label1_2 As System.Windows.Forms.Label
	Public WithEvents _Label1_0 As System.Windows.Forms.Label
	Public WithEvents _Label1_3 As System.Windows.Forms.Label
	Public WithEvents _Label1_4 As System.Windows.Forms.Label
	Public WithEvents _Label1_5 As System.Windows.Forms.Label
	Public WithEvents _Label1_1 As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents cmdQuit As System.Windows.Forms.Button
	Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmWSDelin))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.MainMenu1 = New System.Windows.Forms.MenuStrip
		Me.mnuDefWSDelin = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuNewWSDelin = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuNewExist = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuDelWSDelin = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuHelp = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuWSDelin = New System.Windows.Forms.ToolStripMenuItem
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me.txtLSGrid = New System.Windows.Forms.TextBox
		Me.cboDEMUnits = New System.Windows.Forms.ComboBox
		Me.cboWSSize = New System.Windows.Forms.ComboBox
		Me.chkHydroCorr = New System.Windows.Forms.CheckBox
		Me.txtFlowAccumGrid = New System.Windows.Forms.TextBox
		Me.txtWSFile = New System.Windows.Forms.TextBox
		Me.txtStream = New System.Windows.Forms.TextBox
		Me.cboWSDelin = New System.Windows.Forms.ComboBox
		Me.txtDEMFile = New System.Windows.Forms.TextBox
		Me._Label1_6 = New System.Windows.Forms.Label
		Me._Label1_12 = New System.Windows.Forms.Label
		Me._Label1_2 = New System.Windows.Forms.Label
		Me._Label1_0 = New System.Windows.Forms.Label
		Me._Label1_3 = New System.Windows.Forms.Label
		Me._Label1_4 = New System.Windows.Forms.Label
		Me._Label1_5 = New System.Windows.Forms.Label
		Me._Label1_1 = New System.Windows.Forms.Label
		Me.cmdQuit = New System.Windows.Forms.Button
		Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.MainMenu1.SuspendLayout()
		Me.Frame1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "Watershed Delineations"
		Me.ClientSize = New System.Drawing.Size(504, 302)
		Me.Location = New System.Drawing.Point(3, 39)
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
		Me.Name = "frmWSDelin"
		Me.mnuDefWSDelin.Name = "mnuDefWSDelin"
		Me.mnuDefWSDelin.Text = "&Options"
		Me.mnuDefWSDelin.Checked = False
		Me.mnuDefWSDelin.Enabled = True
		Me.mnuDefWSDelin.Visible = True
		Me.mnuNewWSDelin.Name = "mnuNewWSDelin"
		Me.mnuNewWSDelin.Text = "&New..."
		Me.mnuNewWSDelin.Checked = False
		Me.mnuNewWSDelin.Enabled = True
		Me.mnuNewWSDelin.Visible = True
		Me.mnuNewExist.Name = "mnuNewExist"
		Me.mnuNewExist.Text = "New from existing data..."
		Me.mnuNewExist.Checked = False
		Me.mnuNewExist.Enabled = True
		Me.mnuNewExist.Visible = True
		Me.mnuDelWSDelin.Name = "mnuDelWSDelin"
		Me.mnuDelWSDelin.Text = "&Delete..."
		Me.mnuDelWSDelin.Checked = False
		Me.mnuDelWSDelin.Enabled = True
		Me.mnuDelWSDelin.Visible = True
		Me.mnuHelp.Name = "mnuHelp"
		Me.mnuHelp.Text = "&Help"
		Me.mnuHelp.Checked = False
		Me.mnuHelp.Enabled = True
		Me.mnuHelp.Visible = True
		Me.mnuWSDelin.Name = "mnuWSDelin"
		Me.mnuWSDelin.Text = "Watershed Delineations..."
		Me.mnuWSDelin.ShortcutKeys = CType(System.Windows.Forms.Keys.Shift or System.Windows.Forms.Keys.F1, System.Windows.Forms.Keys)
		Me.mnuWSDelin.Checked = False
		Me.mnuWSDelin.Enabled = True
		Me.mnuWSDelin.Visible = True
		Me.Frame1.Text = "Browse Watershed Delineations  "
		Me.Frame1.Size = New System.Drawing.Size(484, 230)
		Me.Frame1.Location = New System.Drawing.Point(10, 28)
		Me.Frame1.TabIndex = 1
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.BackColor = System.Drawing.SystemColors.Control
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame1.Name = "Frame1"
		Me.txtLSGrid.AutoSize = False
		Me.txtLSGrid.Enabled = False
		Me.txtLSGrid.Size = New System.Drawing.Size(293, 19)
		Me.txtLSGrid.Location = New System.Drawing.Point(178, 193)
		Me.txtLSGrid.TabIndex = 17
		Me.txtLSGrid.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtLSGrid.AcceptsReturn = True
		Me.txtLSGrid.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtLSGrid.BackColor = System.Drawing.SystemColors.Window
		Me.txtLSGrid.CausesValidation = True
		Me.txtLSGrid.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtLSGrid.HideSelection = True
		Me.txtLSGrid.ReadOnly = False
		Me.txtLSGrid.Maxlength = 0
		Me.txtLSGrid.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtLSGrid.MultiLine = False
		Me.txtLSGrid.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtLSGrid.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtLSGrid.TabStop = True
		Me.txtLSGrid.Visible = True
		Me.txtLSGrid.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtLSGrid.Name = "txtLSGrid"
		Me.cboDEMUnits.Enabled = False
		Me.cboDEMUnits.Size = New System.Drawing.Size(170, 21)
		Me.cboDEMUnits.Location = New System.Drawing.Point(178, 67)
		Me.cboDEMUnits.Items.AddRange(New Object(){"meters", "feet"})
		Me.cboDEMUnits.TabIndex = 16
		Me.cboDEMUnits.Text = "Combo1"
		Me.cboDEMUnits.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboDEMUnits.BackColor = System.Drawing.SystemColors.Window
		Me.cboDEMUnits.CausesValidation = True
		Me.cboDEMUnits.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboDEMUnits.IntegralHeight = True
		Me.cboDEMUnits.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboDEMUnits.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboDEMUnits.Sorted = False
		Me.cboDEMUnits.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboDEMUnits.TabStop = True
		Me.cboDEMUnits.Visible = True
		Me.cboDEMUnits.Name = "cboDEMUnits"
		Me.cboWSSize.CausesValidation = False
		Me.cboWSSize.Enabled = False
		Me.cboWSSize.Size = New System.Drawing.Size(120, 21)
		Me.cboWSSize.Location = New System.Drawing.Point(178, 114)
		Me.cboWSSize.Items.AddRange(New Object(){"small", "medium", "large"})
		Me.cboWSSize.TabIndex = 15
		Me.cboWSSize.Text = "cboWSSize"
		Me.cboWSSize.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboWSSize.BackColor = System.Drawing.SystemColors.Window
		Me.cboWSSize.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboWSSize.IntegralHeight = True
		Me.cboWSSize.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboWSSize.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboWSSize.Sorted = False
		Me.cboWSSize.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboWSSize.TabStop = True
		Me.cboWSSize.Visible = True
		Me.cboWSSize.Name = "cboWSSize"
		Me.chkHydroCorr.Text = "Hydrologically Corrected DEM"
		Me.chkHydroCorr.Enabled = False
		Me.chkHydroCorr.Size = New System.Drawing.Size(173, 19)
		Me.chkHydroCorr.Location = New System.Drawing.Point(179, 93)
		Me.chkHydroCorr.TabIndex = 7
		Me.chkHydroCorr.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkHydroCorr.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkHydroCorr.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkHydroCorr.BackColor = System.Drawing.SystemColors.Control
		Me.chkHydroCorr.CausesValidation = True
		Me.chkHydroCorr.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkHydroCorr.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkHydroCorr.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkHydroCorr.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkHydroCorr.TabStop = True
		Me.chkHydroCorr.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkHydroCorr.Visible = True
		Me.chkHydroCorr.Name = "chkHydroCorr"
		Me.txtFlowAccumGrid.AutoSize = False
		Me.txtFlowAccumGrid.Enabled = False
		Me.txtFlowAccumGrid.Size = New System.Drawing.Size(293, 19)
		Me.txtFlowAccumGrid.Location = New System.Drawing.Point(178, 167)
		Me.txtFlowAccumGrid.TabIndex = 6
		Me.txtFlowAccumGrid.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtFlowAccumGrid.AcceptsReturn = True
		Me.txtFlowAccumGrid.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtFlowAccumGrid.BackColor = System.Drawing.SystemColors.Window
		Me.txtFlowAccumGrid.CausesValidation = True
		Me.txtFlowAccumGrid.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtFlowAccumGrid.HideSelection = True
		Me.txtFlowAccumGrid.ReadOnly = False
		Me.txtFlowAccumGrid.Maxlength = 0
		Me.txtFlowAccumGrid.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtFlowAccumGrid.MultiLine = False
		Me.txtFlowAccumGrid.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtFlowAccumGrid.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtFlowAccumGrid.TabStop = True
		Me.txtFlowAccumGrid.Visible = True
		Me.txtFlowAccumGrid.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtFlowAccumGrid.Name = "txtFlowAccumGrid"
		Me.txtWSFile.AutoSize = False
		Me.txtWSFile.Enabled = False
		Me.txtWSFile.Size = New System.Drawing.Size(293, 19)
		Me.txtWSFile.Location = New System.Drawing.Point(178, 140)
		Me.txtWSFile.TabIndex = 5
		Me.txtWSFile.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtWSFile.AcceptsReturn = True
		Me.txtWSFile.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtWSFile.BackColor = System.Drawing.SystemColors.Window
		Me.txtWSFile.CausesValidation = True
		Me.txtWSFile.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtWSFile.HideSelection = True
		Me.txtWSFile.ReadOnly = False
		Me.txtWSFile.Maxlength = 0
		Me.txtWSFile.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtWSFile.MultiLine = False
		Me.txtWSFile.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtWSFile.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtWSFile.TabStop = True
		Me.txtWSFile.Visible = True
		Me.txtWSFile.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtWSFile.Name = "txtWSFile"
		Me.txtStream.AutoSize = False
		Me.txtStream.Enabled = False
		Me.txtStream.Size = New System.Drawing.Size(202, 19)
		Me.txtStream.Location = New System.Drawing.Point(178, 115)
		Me.txtStream.TabIndex = 4
		Me.txtStream.Visible = False
		Me.txtStream.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtStream.AcceptsReturn = True
		Me.txtStream.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtStream.BackColor = System.Drawing.SystemColors.Window
		Me.txtStream.CausesValidation = True
		Me.txtStream.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtStream.HideSelection = True
		Me.txtStream.ReadOnly = False
		Me.txtStream.Maxlength = 0
		Me.txtStream.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtStream.MultiLine = False
		Me.txtStream.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtStream.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtStream.TabStop = True
		Me.txtStream.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtStream.Name = "txtStream"
		Me.cboWSDelin.Size = New System.Drawing.Size(113, 21)
		Me.cboWSDelin.Location = New System.Drawing.Point(178, 16)
		Me.cboWSDelin.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cboWSDelin.TabIndex = 3
		Me.cboWSDelin.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboWSDelin.BackColor = System.Drawing.SystemColors.Window
		Me.cboWSDelin.CausesValidation = True
		Me.cboWSDelin.Enabled = True
		Me.cboWSDelin.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboWSDelin.IntegralHeight = True
		Me.cboWSDelin.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboWSDelin.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboWSDelin.Sorted = False
		Me.cboWSDelin.TabStop = True
		Me.cboWSDelin.Visible = True
		Me.cboWSDelin.Name = "cboWSDelin"
		Me.txtDEMFile.AutoSize = False
		Me.txtDEMFile.Enabled = False
		Me.txtDEMFile.Size = New System.Drawing.Size(290, 19)
		Me.txtDEMFile.Location = New System.Drawing.Point(178, 43)
		Me.txtDEMFile.TabIndex = 2
		Me.txtDEMFile.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtDEMFile.AcceptsReturn = True
		Me.txtDEMFile.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtDEMFile.BackColor = System.Drawing.SystemColors.Window
		Me.txtDEMFile.CausesValidation = True
		Me.txtDEMFile.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtDEMFile.HideSelection = True
		Me.txtDEMFile.ReadOnly = False
		Me.txtDEMFile.Maxlength = 0
		Me.txtDEMFile.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtDEMFile.MultiLine = False
		Me.txtDEMFile.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtDEMFile.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtDEMFile.TabStop = True
		Me.txtDEMFile.Visible = True
		Me.txtDEMFile.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtDEMFile.Name = "txtDEMFile"
		Me._Label1_6.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_6.Text = "LS Grid:"
		Me._Label1_6.Size = New System.Drawing.Size(38, 13)
		Me._Label1_6.Location = New System.Drawing.Point(124, 194)
		Me._Label1_6.TabIndex = 18
		Me._Label1_6.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_6.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_6.Enabled = True
		Me._Label1_6.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_6.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_6.UseMnemonic = True
		Me._Label1_6.Visible = True
		Me._Label1_6.AutoSize = True
		Me._Label1_6.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_6.Name = "_Label1_6"
		Me._Label1_12.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_12.Text = "DEM Grid:"
		Me._Label1_12.Size = New System.Drawing.Size(49, 13)
		Me._Label1_12.Location = New System.Drawing.Point(113, 42)
		Me._Label1_12.TabIndex = 14
		Me._Label1_12.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_12.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_12.Enabled = True
		Me._Label1_12.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_12.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_12.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_12.UseMnemonic = True
		Me._Label1_12.Visible = True
		Me._Label1_12.AutoSize = True
		Me._Label1_12.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_12.Name = "_Label1_12"
		Me._Label1_2.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_2.Text = "Units:"
		Me._Label1_2.Size = New System.Drawing.Size(27, 13)
		Me._Label1_2.Location = New System.Drawing.Point(135, 67)
		Me._Label1_2.TabIndex = 13
		Me._Label1_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_2.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_2.Enabled = True
		Me._Label1_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_2.UseMnemonic = True
		Me._Label1_2.Visible = True
		Me._Label1_2.AutoSize = True
		Me._Label1_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_2.Name = "_Label1_2"
		Me._Label1_0.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_0.Text = "Stream Agreement Layer:"
		Me._Label1_0.Size = New System.Drawing.Size(119, 13)
		Me._Label1_0.Location = New System.Drawing.Point(43, 115)
		Me._Label1_0.TabIndex = 12
		Me._Label1_0.Visible = False
		Me._Label1_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_0.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_0.Enabled = True
		Me._Label1_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_0.UseMnemonic = True
		Me._Label1_0.AutoSize = True
		Me._Label1_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_0.Name = "_Label1_0"
		Me._Label1_3.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_3.Text = "Watershed Delineation Name:"
		Me._Label1_3.Size = New System.Drawing.Size(142, 13)
		Me._Label1_3.Location = New System.Drawing.Point(25, 19)
		Me._Label1_3.TabIndex = 11
		Me._Label1_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_3.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_3.Enabled = True
		Me._Label1_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_3.UseMnemonic = True
		Me._Label1_3.Visible = True
		Me._Label1_3.AutoSize = True
		Me._Label1_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_3.Name = "_Label1_3"
		Me._Label1_4.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_4.Text = "Watershed:"
		Me._Label1_4.Size = New System.Drawing.Size(55, 13)
		Me._Label1_4.Location = New System.Drawing.Point(107, 140)
		Me._Label1_4.TabIndex = 10
		Me._Label1_4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_4.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_4.Enabled = True
		Me._Label1_4.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_4.UseMnemonic = True
		Me._Label1_4.Visible = True
		Me._Label1_4.AutoSize = True
		Me._Label1_4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_4.Name = "_Label1_4"
		Me._Label1_5.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_5.Text = "Flow Accumulation Grid:"
		Me._Label1_5.Size = New System.Drawing.Size(114, 13)
		Me._Label1_5.Location = New System.Drawing.Point(48, 167)
		Me._Label1_5.TabIndex = 9
		Me._Label1_5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_5.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_5.Enabled = True
		Me._Label1_5.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_5.UseMnemonic = True
		Me._Label1_5.Visible = True
		Me._Label1_5.AutoSize = True
		Me._Label1_5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_5.Name = "_Label1_5"
		Me._Label1_1.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_1.Text = "Subwatershed Size:"
		Me._Label1_1.Size = New System.Drawing.Size(94, 13)
		Me._Label1_1.Location = New System.Drawing.Point(68, 115)
		Me._Label1_1.TabIndex = 8
		Me._Label1_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_1.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_1.Enabled = True
		Me._Label1_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_1.UseMnemonic = True
		Me._Label1_1.Visible = True
		Me._Label1_1.AutoSize = True
		Me._Label1_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_1.Name = "_Label1_1"
		Me.cmdQuit.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdQuit.Text = "Cancel"
		Me.cmdQuit.Size = New System.Drawing.Size(65, 25)
		Me.cmdQuit.Location = New System.Drawing.Point(416, 267)
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
		Me.Controls.Add(cmdQuit)
		Me.Frame1.Controls.Add(txtLSGrid)
		Me.Frame1.Controls.Add(cboDEMUnits)
		Me.Frame1.Controls.Add(cboWSSize)
		Me.Frame1.Controls.Add(chkHydroCorr)
		Me.Frame1.Controls.Add(txtFlowAccumGrid)
		Me.Frame1.Controls.Add(txtWSFile)
		Me.Frame1.Controls.Add(txtStream)
		Me.Frame1.Controls.Add(cboWSDelin)
		Me.Frame1.Controls.Add(txtDEMFile)
		Me.Frame1.Controls.Add(_Label1_6)
		Me.Frame1.Controls.Add(_Label1_12)
		Me.Frame1.Controls.Add(_Label1_2)
		Me.Frame1.Controls.Add(_Label1_0)
		Me.Frame1.Controls.Add(_Label1_3)
		Me.Frame1.Controls.Add(_Label1_4)
		Me.Frame1.Controls.Add(_Label1_5)
		Me.Frame1.Controls.Add(_Label1_1)
		Me.Label1.SetIndex(_Label1_6, CType(6, Short))
		Me.Label1.SetIndex(_Label1_12, CType(12, Short))
		Me.Label1.SetIndex(_Label1_2, CType(2, Short))
		Me.Label1.SetIndex(_Label1_0, CType(0, Short))
		Me.Label1.SetIndex(_Label1_3, CType(3, Short))
		Me.Label1.SetIndex(_Label1_4, CType(4, Short))
		Me.Label1.SetIndex(_Label1_5, CType(5, Short))
		Me.Label1.SetIndex(_Label1_1, CType(1, Short))
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
		MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuDefWSDelin, Me.mnuHelp})
		mnuDefWSDelin.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuNewWSDelin, Me.mnuNewExist, Me.mnuDelWSDelin})
		mnuHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuWSDelin})
		Me.Controls.Add(MainMenu1)
		Me.MainMenu1.ResumeLayout(False)
		Me.Frame1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class