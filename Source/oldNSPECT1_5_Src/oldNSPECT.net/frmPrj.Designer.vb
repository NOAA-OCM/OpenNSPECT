<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmPrj
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
	Public WithEvents mnuOpen As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuSpace As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuSave As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuSaveAs As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuSpace1 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuExit As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuFile As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuGeneralHelp As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuBigHelp As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuLUAdd As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuLUEdit As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuLUDelete As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuOptions As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuManAppen As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuManInsert As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuManDelete As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuManagement As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
	Public WithEvents cboWQStd As System.Windows.Forms.ComboBox
	Public WithEvents _Label1_6 As System.Windows.Forms.Label
	Public WithEvents Frame6 As System.Windows.Forms.GroupBox
	Public WithEvents cboWSDelin As System.Windows.Forms.ComboBox
	Public WithEvents _Label1_3 As System.Windows.Forms.Label
	Public WithEvents Frame5 As System.Windows.Forms.GroupBox
	Public WithEvents cboPrecipScen As System.Windows.Forms.ComboBox
	Public WithEvents _Label1_7 As System.Windows.Forms.Label
	Public WithEvents Frame4 As System.Windows.Forms.GroupBox
	Public WithEvents cboSelectPoly As System.Windows.Forms.ComboBox
	Public WithEvents chkSelectedPolys As System.Windows.Forms.CheckBox
	Public WithEvents chkLocalEffects As System.Windows.Forms.CheckBox
	Public WithEvents lblLayer As System.Windows.Forms.Label
	Public WithEvents frm_raintype As System.Windows.Forms.GroupBox
	Public WithEvents cmdOpenWS As System.Windows.Forms.Button
	Public WithEvents txtOutputWS As System.Windows.Forms.TextBox
	Public WithEvents txtProjectName As System.Windows.Forms.TextBox
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Frame3 As System.Windows.Forms.GroupBox
	Public WithEvents cboSoilsLayer As System.Windows.Forms.ComboBox
	Public WithEvents Label6 As System.Windows.Forms.Label
	Public WithEvents lblSoilsHyd As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents cboLCUnits As System.Windows.Forms.ComboBox
	Public WithEvents cboLCLayer As System.Windows.Forms.ComboBox
	Public WithEvents cboLCType As System.Windows.Forms.ComboBox
	Public WithEvents _Label1_5 As System.Windows.Forms.Label
	Public WithEvents _Label1_0 As System.Windows.Forms.Label
	Public WithEvents _Label1_2 As System.Windows.Forms.Label
	Public WithEvents fraLC As System.Windows.Forms.GroupBox
	Public WithEvents cboAreaLayer As System.Windows.Forms.ComboBox
	Public WithEvents cboClass As System.Windows.Forms.ComboBox
	Public WithEvents Timer1 As System.Windows.Forms.Timer
	Public WithEvents cboCoeffSet As System.Windows.Forms.ComboBox
	Public WithEvents cboCoeff As System.Windows.Forms.ComboBox
	Public WithEvents grdCoeffs As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
	Public WithEvents _SSTab1_TabPage0 As System.Windows.Forms.TabPage
	Public WithEvents lblErodFactor As System.Windows.Forms.Label
	Public WithEvents lblKFactor As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label7 As System.Windows.Forms.Label
	Public WithEvents chkCalcErosion As System.Windows.Forms.CheckBox
	Public WithEvents txtRainValue As System.Windows.Forms.TextBox
	Public WithEvents cboRainGrid As System.Windows.Forms.ComboBox
	Public WithEvents optUseValue As System.Windows.Forms.RadioButton
	Public WithEvents optUseGRID As System.Windows.Forms.RadioButton
	Public WithEvents frameRainFall As System.Windows.Forms.GroupBox
	Public WithEvents cboErodFactor As System.Windows.Forms.ComboBox
	Public WithEvents cboSoilAttribute As System.Windows.Forms.ComboBox
	Public WithEvents cmdOpenSDR As System.Windows.Forms.Button
	Public WithEvents txtSDRGRID As System.Windows.Forms.TextBox
	Public WithEvents chkSDR As System.Windows.Forms.CheckBox
	Public WithEvents frmSDR As System.Windows.Forms.GroupBox
	Public WithEvents _SSTab1_TabPage1 As System.Windows.Forms.TabPage
	Public WithEvents grdLU As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
	Public WithEvents _SSTab1_TabPage2 As System.Windows.Forms.TabPage
	Public WithEvents grdLCChanges As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
	Public WithEvents _SSTab1_TabPage3 As System.Windows.Forms.TabPage
	Public WithEvents SSTab1 As System.Windows.Forms.TabControl
	Public WithEvents cmdRun As System.Windows.Forms.Button
	Public WithEvents cmdQuit As System.Windows.Forms.Button
	Public dlgXMLOpen As System.Windows.Forms.OpenFileDialog
	Public dlgXMLSave As System.Windows.Forms.SaveFileDialog
	Public WithEvents _chkIgnore_0 As System.Windows.Forms.CheckBox
	Public WithEvents _chkIgnoreMgmt_0 As System.Windows.Forms.CheckBox
	Public WithEvents _chkIgnoreLU_0 As System.Windows.Forms.CheckBox
	Public WithEvents Picture2 As System.Windows.Forms.PictureBox
	Public WithEvents txtThemeName As System.Windows.Forms.TextBox
	Public WithEvents cmdOutputBrowse As System.Windows.Forms.Button
	Public WithEvents txtOutputFile As System.Windows.Forms.TextBox
	Public WithEvents _Label1_11 As System.Windows.Forms.Label
	Public WithEvents _Label1_12 As System.Windows.Forms.Label
	Public WithEvents Frame2 As System.Windows.Forms.GroupBox
	Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents chkIgnore As Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray
	Public WithEvents chkIgnoreLU As Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray
	Public WithEvents chkIgnoreMgmt As Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPrj))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.MainMenu1 = New System.Windows.Forms.MenuStrip
		Me.mnuFile = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuNew = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuOpen = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuSpace = New System.Windows.Forms.ToolStripSeparator
		Me.mnuSave = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuSaveAs = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuSpace1 = New System.Windows.Forms.ToolStripSeparator
		Me.mnuExit = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuBigHelp = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuGeneralHelp = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuOptions = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuLUAdd = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuLUEdit = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuLUDelete = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuManagement = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuManAppen = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuManInsert = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuManDelete = New System.Windows.Forms.ToolStripMenuItem
		Me.Frame6 = New System.Windows.Forms.GroupBox
		Me.cboWQStd = New System.Windows.Forms.ComboBox
		Me._Label1_6 = New System.Windows.Forms.Label
		Me.Frame5 = New System.Windows.Forms.GroupBox
		Me.cboWSDelin = New System.Windows.Forms.ComboBox
		Me._Label1_3 = New System.Windows.Forms.Label
		Me.Frame4 = New System.Windows.Forms.GroupBox
		Me.cboPrecipScen = New System.Windows.Forms.ComboBox
		Me._Label1_7 = New System.Windows.Forms.Label
		Me.frm_raintype = New System.Windows.Forms.GroupBox
		Me.cboSelectPoly = New System.Windows.Forms.ComboBox
		Me.chkSelectedPolys = New System.Windows.Forms.CheckBox
		Me.chkLocalEffects = New System.Windows.Forms.CheckBox
		Me.lblLayer = New System.Windows.Forms.Label
		Me.Frame3 = New System.Windows.Forms.GroupBox
		Me.cmdOpenWS = New System.Windows.Forms.Button
		Me.txtOutputWS = New System.Windows.Forms.TextBox
		Me.txtProjectName = New System.Windows.Forms.TextBox
		Me.Label5 = New System.Windows.Forms.Label
		Me.Label4 = New System.Windows.Forms.Label
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me.cboSoilsLayer = New System.Windows.Forms.ComboBox
		Me.Label6 = New System.Windows.Forms.Label
		Me.lblSoilsHyd = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.fraLC = New System.Windows.Forms.GroupBox
		Me.cboLCUnits = New System.Windows.Forms.ComboBox
		Me.cboLCLayer = New System.Windows.Forms.ComboBox
		Me.cboLCType = New System.Windows.Forms.ComboBox
		Me._Label1_5 = New System.Windows.Forms.Label
		Me._Label1_0 = New System.Windows.Forms.Label
		Me._Label1_2 = New System.Windows.Forms.Label
		Me.cboAreaLayer = New System.Windows.Forms.ComboBox
		Me.cboClass = New System.Windows.Forms.ComboBox
		Me.Timer1 = New System.Windows.Forms.Timer(components)
		Me.cboCoeffSet = New System.Windows.Forms.ComboBox
		Me.cboCoeff = New System.Windows.Forms.ComboBox
		Me.SSTab1 = New System.Windows.Forms.TabControl
		Me._SSTab1_TabPage0 = New System.Windows.Forms.TabPage
		Me.grdCoeffs = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
		Me._SSTab1_TabPage1 = New System.Windows.Forms.TabPage
		Me.lblErodFactor = New System.Windows.Forms.Label
		Me.lblKFactor = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label7 = New System.Windows.Forms.Label
		Me.chkCalcErosion = New System.Windows.Forms.CheckBox
		Me.frameRainFall = New System.Windows.Forms.GroupBox
		Me.txtRainValue = New System.Windows.Forms.TextBox
		Me.cboRainGrid = New System.Windows.Forms.ComboBox
		Me.optUseValue = New System.Windows.Forms.RadioButton
		Me.optUseGRID = New System.Windows.Forms.RadioButton
		Me.cboErodFactor = New System.Windows.Forms.ComboBox
		Me.cboSoilAttribute = New System.Windows.Forms.ComboBox
		Me.frmSDR = New System.Windows.Forms.GroupBox
		Me.cmdOpenSDR = New System.Windows.Forms.Button
		Me.txtSDRGRID = New System.Windows.Forms.TextBox
		Me.chkSDR = New System.Windows.Forms.CheckBox
		Me._SSTab1_TabPage2 = New System.Windows.Forms.TabPage
		Me.grdLU = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
		Me._SSTab1_TabPage3 = New System.Windows.Forms.TabPage
		Me.grdLCChanges = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
		Me.cmdRun = New System.Windows.Forms.Button
		Me.cmdQuit = New System.Windows.Forms.Button
		Me.dlgXMLOpen = New System.Windows.Forms.OpenFileDialog
		Me.dlgXMLSave = New System.Windows.Forms.SaveFileDialog
		Me._chkIgnore_0 = New System.Windows.Forms.CheckBox
		Me._chkIgnoreMgmt_0 = New System.Windows.Forms.CheckBox
		Me._chkIgnoreLU_0 = New System.Windows.Forms.CheckBox
		Me.Picture2 = New System.Windows.Forms.PictureBox
		Me.Frame2 = New System.Windows.Forms.GroupBox
		Me.txtThemeName = New System.Windows.Forms.TextBox
		Me.cmdOutputBrowse = New System.Windows.Forms.Button
		Me.txtOutputFile = New System.Windows.Forms.TextBox
		Me._Label1_11 = New System.Windows.Forms.Label
		Me._Label1_12 = New System.Windows.Forms.Label
		Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.chkIgnore = New Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray(components)
		Me.chkIgnoreLU = New Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray(components)
		Me.chkIgnoreMgmt = New Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray(components)
		Me.MainMenu1.SuspendLayout()
		Me.Frame6.SuspendLayout()
		Me.Frame5.SuspendLayout()
		Me.Frame4.SuspendLayout()
		Me.frm_raintype.SuspendLayout()
		Me.Frame3.SuspendLayout()
		Me.Frame1.SuspendLayout()
		Me.fraLC.SuspendLayout()
		Me.SSTab1.SuspendLayout()
		Me._SSTab1_TabPage0.SuspendLayout()
		Me._SSTab1_TabPage1.SuspendLayout()
		Me.frameRainFall.SuspendLayout()
		Me.frmSDR.SuspendLayout()
		Me._SSTab1_TabPage2.SuspendLayout()
		Me._SSTab1_TabPage3.SuspendLayout()
		Me.Frame2.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.grdCoeffs, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.grdLU, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.grdLCChanges, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.chkIgnore, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.chkIgnoreLU, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.chkIgnoreMgmt, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.ClientSize = New System.Drawing.Size(642, 502)
		Me.Location = New System.Drawing.Point(157, 105)
		Me.Icon = CType(resources.GetObject("frmPrj.Icon"), System.Drawing.Icon)
		Me.KeyPreview = True
		Me.MaximizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmPrj"
		Me.mnuFile.Name = "mnuFile"
		Me.mnuFile.Text = "&File"
		Me.mnuFile.Checked = False
		Me.mnuFile.Enabled = True
		Me.mnuFile.Visible = True
		Me.mnuNew.Name = "mnuNew"
		Me.mnuNew.Text = "New Project"
		Me.mnuNew.ShortcutKeys = CType(System.Windows.Forms.Keys.Control or System.Windows.Forms.Keys.N, System.Windows.Forms.Keys)
		Me.mnuNew.Checked = False
		Me.mnuNew.Enabled = True
		Me.mnuNew.Visible = True
		Me.mnuOpen.Name = "mnuOpen"
		Me.mnuOpen.Text = "Open Project..."
		Me.mnuOpen.ShortcutKeys = CType(System.Windows.Forms.Keys.Control or System.Windows.Forms.Keys.O, System.Windows.Forms.Keys)
		Me.mnuOpen.Checked = False
		Me.mnuOpen.Enabled = True
		Me.mnuOpen.Visible = True
		Me.mnuSpace.Enabled = True
		Me.mnuSpace.Visible = True
		Me.mnuSpace.Name = "mnuSpace"
		Me.mnuSave.Name = "mnuSave"
		Me.mnuSave.Text = "Save"
		Me.mnuSave.ShortcutKeys = CType(System.Windows.Forms.Keys.Control or System.Windows.Forms.Keys.S, System.Windows.Forms.Keys)
		Me.mnuSave.Checked = False
		Me.mnuSave.Enabled = True
		Me.mnuSave.Visible = True
		Me.mnuSaveAs.Name = "mnuSaveAs"
		Me.mnuSaveAs.Text = "Save As..."
		Me.mnuSaveAs.Checked = False
		Me.mnuSaveAs.Enabled = True
		Me.mnuSaveAs.Visible = True
		Me.mnuSpace1.Enabled = True
		Me.mnuSpace1.Visible = True
		Me.mnuSpace1.Name = "mnuSpace1"
		Me.mnuExit.Name = "mnuExit"
		Me.mnuExit.Text = "E&xit"
		Me.mnuExit.ShortcutKeys = CType(System.Windows.Forms.Keys.Control or System.Windows.Forms.Keys.X, System.Windows.Forms.Keys)
		Me.mnuExit.Checked = False
		Me.mnuExit.Enabled = True
		Me.mnuExit.Visible = True
		Me.mnuBigHelp.Name = "mnuBigHelp"
		Me.mnuBigHelp.Text = "&Help"
		Me.mnuBigHelp.Checked = False
		Me.mnuBigHelp.Enabled = True
		Me.mnuBigHelp.Visible = True
		Me.mnuGeneralHelp.Name = "mnuGeneralHelp"
		Me.mnuGeneralHelp.Text = "N-SPECT Help..."
		Me.mnuGeneralHelp.ShortcutKeys = CType(System.Windows.Forms.Keys.Shift or System.Windows.Forms.Keys.F1, System.Windows.Forms.Keys)
		Me.mnuGeneralHelp.Checked = False
		Me.mnuGeneralHelp.Enabled = True
		Me.mnuGeneralHelp.Visible = True
		Me.mnuOptions.Name = "mnuOptions"
		Me.mnuOptions.Text = ""
		Me.mnuOptions.Visible = False
		Me.mnuOptions.Checked = False
		Me.mnuOptions.Enabled = True
		Me.mnuLUAdd.Name = "mnuLUAdd"
		Me.mnuLUAdd.Text = "Add Scenario..."
		Me.mnuLUAdd.Checked = False
		Me.mnuLUAdd.Enabled = True
		Me.mnuLUAdd.Visible = True
		Me.mnuLUEdit.Name = "mnuLUEdit"
		Me.mnuLUEdit.Text = "Edit Scenario..."
		Me.mnuLUEdit.Checked = False
		Me.mnuLUEdit.Enabled = True
		Me.mnuLUEdit.Visible = True
		Me.mnuLUDelete.Name = "mnuLUDelete"
		Me.mnuLUDelete.Text = "Delete Scenario..."
		Me.mnuLUDelete.Checked = False
		Me.mnuLUDelete.Enabled = True
		Me.mnuLUDelete.Visible = True
		Me.mnuManagement.Name = "mnuManagement"
		Me.mnuManagement.Text = ""
		Me.mnuManagement.Visible = False
		Me.mnuManagement.Checked = False
		Me.mnuManagement.Enabled = True
		Me.mnuManAppen.Name = "mnuManAppen"
		Me.mnuManAppen.Text = "Append Row"
		Me.mnuManAppen.Checked = False
		Me.mnuManAppen.Enabled = True
		Me.mnuManAppen.Visible = True
		Me.mnuManInsert.Name = "mnuManInsert"
		Me.mnuManInsert.Text = "Insert Row"
		Me.mnuManInsert.Checked = False
		Me.mnuManInsert.Enabled = True
		Me.mnuManInsert.Visible = True
		Me.mnuManDelete.Name = "mnuManDelete"
		Me.mnuManDelete.Text = "Delete Current Row"
		Me.mnuManDelete.Checked = False
		Me.mnuManDelete.Enabled = True
		Me.mnuManDelete.Visible = True
		Me.Frame6.Text = "Water Quality Standard "
		Me.Frame6.Size = New System.Drawing.Size(209, 58)
		Me.Frame6.Location = New System.Drawing.Point(426, 192)
		Me.Frame6.TabIndex = 58
		Me.Frame6.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame6.BackColor = System.Drawing.SystemColors.Control
		Me.Frame6.Enabled = True
		Me.Frame6.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame6.Visible = True
		Me.Frame6.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame6.Name = "Frame6"
		Me.cboWQStd.Size = New System.Drawing.Size(141, 21)
		Me.cboWQStd.Location = New System.Drawing.Point(49, 22)
		Me.cboWQStd.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cboWQStd.TabIndex = 59
		Me.cboWQStd.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboWQStd.BackColor = System.Drawing.SystemColors.Window
		Me.cboWQStd.CausesValidation = True
		Me.cboWQStd.Enabled = True
		Me.cboWQStd.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboWQStd.IntegralHeight = True
		Me.cboWQStd.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboWQStd.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboWQStd.Sorted = False
		Me.cboWQStd.TabStop = True
		Me.cboWQStd.Visible = True
		Me.cboWQStd.Name = "cboWQStd"
		Me._Label1_6.Text = "Name: "
		Me._Label1_6.Size = New System.Drawing.Size(44, 16)
		Me._Label1_6.Location = New System.Drawing.Point(9, 25)
		Me._Label1_6.TabIndex = 60
		Me._Label1_6.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_6.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_6.BackColor = System.Drawing.Color.Transparent
		Me._Label1_6.Enabled = True
		Me._Label1_6.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_6.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_6.UseMnemonic = True
		Me._Label1_6.Visible = True
		Me._Label1_6.AutoSize = False
		Me._Label1_6.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_6.Name = "_Label1_6"
		Me.Frame5.Text = "Watershed Delineation "
		Me.Frame5.Size = New System.Drawing.Size(202, 58)
		Me.Frame5.Location = New System.Drawing.Point(217, 192)
		Me.Frame5.TabIndex = 55
		Me.Frame5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame5.BackColor = System.Drawing.SystemColors.Control
		Me.Frame5.Enabled = True
		Me.Frame5.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame5.Visible = True
		Me.Frame5.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame5.Name = "Frame5"
		Me.cboWSDelin.Size = New System.Drawing.Size(141, 21)
		Me.cboWSDelin.Location = New System.Drawing.Point(46, 22)
		Me.cboWSDelin.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cboWSDelin.TabIndex = 56
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
		Me._Label1_3.Text = "Name:"
		Me._Label1_3.Size = New System.Drawing.Size(39, 17)
		Me._Label1_3.Location = New System.Drawing.Point(7, 25)
		Me._Label1_3.TabIndex = 57
		Me._Label1_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_3.BackColor = System.Drawing.Color.Transparent
		Me._Label1_3.Enabled = True
		Me._Label1_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_3.UseMnemonic = True
		Me._Label1_3.Visible = True
		Me._Label1_3.AutoSize = False
		Me._Label1_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_3.Name = "_Label1_3"
		Me.Frame4.Text = "Precipitation Scenario  "
		Me.Frame4.Size = New System.Drawing.Size(202, 58)
		Me.Frame4.Location = New System.Drawing.Point(10, 192)
		Me.Frame4.TabIndex = 52
		Me.Frame4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame4.BackColor = System.Drawing.SystemColors.Control
		Me.Frame4.Enabled = True
		Me.Frame4.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame4.Visible = True
		Me.Frame4.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame4.Name = "Frame4"
		Me.cboPrecipScen.Size = New System.Drawing.Size(141, 21)
		Me.cboPrecipScen.Location = New System.Drawing.Point(47, 22)
		Me.cboPrecipScen.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cboPrecipScen.TabIndex = 53
		Me.cboPrecipScen.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboPrecipScen.BackColor = System.Drawing.SystemColors.Window
		Me.cboPrecipScen.CausesValidation = True
		Me.cboPrecipScen.Enabled = True
		Me.cboPrecipScen.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboPrecipScen.IntegralHeight = True
		Me.cboPrecipScen.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboPrecipScen.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboPrecipScen.Sorted = False
		Me.cboPrecipScen.TabStop = True
		Me.cboPrecipScen.Visible = True
		Me.cboPrecipScen.Name = "cboPrecipScen"
		Me._Label1_7.Text = "Name: "
		Me._Label1_7.Size = New System.Drawing.Size(49, 18)
		Me._Label1_7.Location = New System.Drawing.Point(9, 25)
		Me._Label1_7.TabIndex = 54
		Me._Label1_7.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_7.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_7.BackColor = System.Drawing.Color.Transparent
		Me._Label1_7.Enabled = True
		Me._Label1_7.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_7.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_7.UseMnemonic = True
		Me._Label1_7.Visible = True
		Me._Label1_7.AutoSize = False
		Me._Label1_7.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_7.Name = "_Label1_7"
		Me.frm_raintype.Text = "Miscellaneous "
		Me.frm_raintype.Size = New System.Drawing.Size(175, 113)
		Me.frm_raintype.Location = New System.Drawing.Point(461, 73)
		Me.frm_raintype.TabIndex = 50
		Me.frm_raintype.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frm_raintype.BackColor = System.Drawing.SystemColors.Control
		Me.frm_raintype.Enabled = True
		Me.frm_raintype.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frm_raintype.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frm_raintype.Visible = True
		Me.frm_raintype.Padding = New System.Windows.Forms.Padding(0)
		Me.frm_raintype.Name = "frm_raintype"
		Me.cboSelectPoly.Enabled = False
		Me.cboSelectPoly.Size = New System.Drawing.Size(112, 21)
		Me.cboSelectPoly.Location = New System.Drawing.Point(45, 53)
		Me.cboSelectPoly.TabIndex = 62
		Me.cboSelectPoly.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboSelectPoly.BackColor = System.Drawing.SystemColors.Window
		Me.cboSelectPoly.CausesValidation = True
		Me.cboSelectPoly.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboSelectPoly.IntegralHeight = True
		Me.cboSelectPoly.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboSelectPoly.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboSelectPoly.Sorted = False
		Me.cboSelectPoly.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboSelectPoly.TabStop = True
		Me.cboSelectPoly.Visible = True
		Me.cboSelectPoly.Name = "cboSelectPoly"
		Me.chkSelectedPolys.Text = "Selected Polygons Only"
		Me.chkSelectedPolys.Size = New System.Drawing.Size(147, 22)
		Me.chkSelectedPolys.Location = New System.Drawing.Point(7, 22)
		Me.chkSelectedPolys.TabIndex = 61
		Me.ToolTip1.SetToolTip(Me.chkSelectedPolys, "Select to limit analysis to selected polygons from a map layer")
		Me.chkSelectedPolys.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkSelectedPolys.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkSelectedPolys.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkSelectedPolys.BackColor = System.Drawing.SystemColors.Control
		Me.chkSelectedPolys.CausesValidation = True
		Me.chkSelectedPolys.Enabled = True
		Me.chkSelectedPolys.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkSelectedPolys.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkSelectedPolys.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkSelectedPolys.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkSelectedPolys.TabStop = True
		Me.chkSelectedPolys.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkSelectedPolys.Visible = True
		Me.chkSelectedPolys.Name = "chkSelectedPolys"
		Me.chkLocalEffects.Text = "Local Effects Only"
		Me.chkLocalEffects.Size = New System.Drawing.Size(131, 17)
		Me.chkLocalEffects.Location = New System.Drawing.Point(8, 86)
		Me.chkLocalEffects.TabIndex = 51
		Me.ToolTip1.SetToolTip(Me.chkLocalEffects, "Select for analysis of local effects only")
		Me.chkLocalEffects.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkLocalEffects.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkLocalEffects.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkLocalEffects.BackColor = System.Drawing.SystemColors.Control
		Me.chkLocalEffects.CausesValidation = True
		Me.chkLocalEffects.Enabled = True
		Me.chkLocalEffects.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkLocalEffects.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkLocalEffects.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkLocalEffects.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkLocalEffects.TabStop = True
		Me.chkLocalEffects.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkLocalEffects.Visible = True
		Me.chkLocalEffects.Name = "chkLocalEffects"
		Me.lblLayer.Text = "Layer: "
		Me.lblLayer.Enabled = False
		Me.lblLayer.Size = New System.Drawing.Size(44, 15)
		Me.lblLayer.Location = New System.Drawing.Point(6, 56)
		Me.lblLayer.TabIndex = 63
		Me.lblLayer.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblLayer.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblLayer.BackColor = System.Drawing.SystemColors.Control
		Me.lblLayer.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblLayer.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblLayer.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblLayer.UseMnemonic = True
		Me.lblLayer.Visible = True
		Me.lblLayer.AutoSize = False
		Me.lblLayer.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblLayer.Name = "lblLayer"
		Me.Frame3.Text = "Project Information "
		Me.Frame3.Size = New System.Drawing.Size(625, 41)
		Me.Frame3.Location = New System.Drawing.Point(9, 27)
		Me.Frame3.TabIndex = 37
		Me.Frame3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame3.BackColor = System.Drawing.SystemColors.Control
		Me.Frame3.Enabled = True
		Me.Frame3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame3.Visible = True
		Me.Frame3.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame3.Name = "Frame3"
		Me.cmdOpenWS.TextAlign = System.Drawing.ContentAlignment.BottomCenter
		Me.cmdOpenWS.Size = New System.Drawing.Size(23, 20)
		Me.cmdOpenWS.Location = New System.Drawing.Point(570, 14)
		Me.cmdOpenWS.Image = CType(resources.GetObject("cmdOpenWS.Image"), System.Drawing.Image)
		Me.cmdOpenWS.TabIndex = 40
		Me.cmdOpenWS.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOpenWS.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOpenWS.CausesValidation = True
		Me.cmdOpenWS.Enabled = True
		Me.cmdOpenWS.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOpenWS.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOpenWS.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOpenWS.TabStop = True
		Me.cmdOpenWS.Name = "cmdOpenWS"
		Me.txtOutputWS.AutoSize = False
		Me.txtOutputWS.Size = New System.Drawing.Size(192, 20)
		Me.txtOutputWS.Location = New System.Drawing.Point(373, 15)
		Me.txtOutputWS.TabIndex = 39
		Me.txtOutputWS.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtOutputWS.AcceptsReturn = True
		Me.txtOutputWS.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtOutputWS.BackColor = System.Drawing.SystemColors.Window
		Me.txtOutputWS.CausesValidation = True
		Me.txtOutputWS.Enabled = True
		Me.txtOutputWS.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtOutputWS.HideSelection = True
		Me.txtOutputWS.ReadOnly = False
		Me.txtOutputWS.Maxlength = 0
		Me.txtOutputWS.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtOutputWS.MultiLine = False
		Me.txtOutputWS.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtOutputWS.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtOutputWS.TabStop = True
		Me.txtOutputWS.Visible = True
		Me.txtOutputWS.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtOutputWS.Name = "txtOutputWS"
		Me.txtProjectName.AutoSize = False
		Me.txtProjectName.Size = New System.Drawing.Size(161, 20)
		Me.txtProjectName.Location = New System.Drawing.Point(63, 15)
		Me.txtProjectName.TabIndex = 38
		Me.txtProjectName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtProjectName.AcceptsReturn = True
		Me.txtProjectName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtProjectName.BackColor = System.Drawing.SystemColors.Window
		Me.txtProjectName.CausesValidation = True
		Me.txtProjectName.Enabled = True
		Me.txtProjectName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtProjectName.HideSelection = True
		Me.txtProjectName.ReadOnly = False
		Me.txtProjectName.Maxlength = 0
		Me.txtProjectName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtProjectName.MultiLine = False
		Me.txtProjectName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtProjectName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtProjectName.TabStop = True
		Me.txtProjectName.Visible = True
		Me.txtProjectName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtProjectName.Name = "txtProjectName"
		Me.Label5.Text = "Working Directory:"
		Me.Label5.Size = New System.Drawing.Size(99, 18)
		Me.Label5.Location = New System.Drawing.Point(270, 16)
		Me.Label5.TabIndex = 42
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
		Me.Label4.Text = "Name:"
		Me.Label4.Size = New System.Drawing.Size(54, 16)
		Me.Label4.Location = New System.Drawing.Point(11, 17)
		Me.Label4.TabIndex = 41
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
		Me.Frame1.Text = "Soils"
		Me.Frame1.Size = New System.Drawing.Size(214, 113)
		Me.Frame1.Location = New System.Drawing.Point(240, 73)
		Me.Frame1.TabIndex = 9
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.BackColor = System.Drawing.SystemColors.Control
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame1.Name = "Frame1"
		Me.cboSoilsLayer.Size = New System.Drawing.Size(141, 21)
		Me.cboSoilsLayer.Location = New System.Drawing.Point(7, 37)
		Me.cboSoilsLayer.TabIndex = 11
		Me.cboSoilsLayer.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboSoilsLayer.BackColor = System.Drawing.SystemColors.Window
		Me.cboSoilsLayer.CausesValidation = True
		Me.cboSoilsLayer.Enabled = True
		Me.cboSoilsLayer.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboSoilsLayer.IntegralHeight = True
		Me.cboSoilsLayer.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboSoilsLayer.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboSoilsLayer.Sorted = False
		Me.cboSoilsLayer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboSoilsLayer.TabStop = True
		Me.cboSoilsLayer.Visible = True
		Me.cboSoilsLayer.Name = "cboSoilsLayer"
		Me.Label6.Text = "Hydrologic Soils Data Set:"
		Me.Label6.Size = New System.Drawing.Size(157, 16)
		Me.Label6.Location = New System.Drawing.Point(7, 63)
		Me.Label6.TabIndex = 47
		Me.Label6.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label6.BackColor = System.Drawing.SystemColors.Control
		Me.Label6.Enabled = True
		Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label6.UseMnemonic = True
		Me.Label6.Visible = True
		Me.Label6.AutoSize = False
		Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label6.Name = "Label6"
		Me.lblSoilsHyd.Size = New System.Drawing.Size(198, 21)
		Me.lblSoilsHyd.Location = New System.Drawing.Point(7, 79)
		Me.lblSoilsHyd.TabIndex = 46
		Me.lblSoilsHyd.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblSoilsHyd.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblSoilsHyd.BackColor = System.Drawing.SystemColors.Control
		Me.lblSoilsHyd.Enabled = True
		Me.lblSoilsHyd.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblSoilsHyd.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblSoilsHyd.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblSoilsHyd.UseMnemonic = True
		Me.lblSoilsHyd.Visible = True
		Me.lblSoilsHyd.AutoSize = False
		Me.lblSoilsHyd.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblSoilsHyd.Name = "lblSoilsHyd"
		Me.Label2.Text = "Soils Definition:"
		Me.Label2.Size = New System.Drawing.Size(136, 17)
		Me.Label2.Location = New System.Drawing.Point(7, 19)
		Me.Label2.TabIndex = 10
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
		Me.fraLC.Text = "Land Cover"
		Me.fraLC.Size = New System.Drawing.Size(226, 113)
		Me.fraLC.Location = New System.Drawing.Point(9, 73)
		Me.fraLC.TabIndex = 5
		Me.fraLC.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fraLC.BackColor = System.Drawing.SystemColors.Control
		Me.fraLC.Enabled = True
		Me.fraLC.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraLC.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraLC.Visible = True
		Me.fraLC.Padding = New System.Windows.Forms.Padding(0)
		Me.fraLC.Name = "fraLC"
		Me.cboLCUnits.Enabled = False
		Me.cboLCUnits.Size = New System.Drawing.Size(141, 21)
		Me.cboLCUnits.Location = New System.Drawing.Point(69, 51)
		Me.cboLCUnits.Items.AddRange(New Object(){"meters", "feet"})
		Me.cboLCUnits.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cboLCUnits.TabIndex = 1
		Me.cboLCUnits.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboLCUnits.BackColor = System.Drawing.SystemColors.Window
		Me.cboLCUnits.CausesValidation = True
		Me.cboLCUnits.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboLCUnits.IntegralHeight = True
		Me.cboLCUnits.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboLCUnits.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboLCUnits.Sorted = False
		Me.cboLCUnits.TabStop = True
		Me.cboLCUnits.Visible = True
		Me.cboLCUnits.Name = "cboLCUnits"
		Me.cboLCLayer.Size = New System.Drawing.Size(141, 21)
		Me.cboLCLayer.Location = New System.Drawing.Point(69, 23)
		Me.cboLCLayer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cboLCLayer.TabIndex = 0
		Me.ToolTip1.SetToolTip(Me.cboLCLayer, "Choose a Land Cover GRID from current layers in map view")
		Me.cboLCLayer.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboLCLayer.BackColor = System.Drawing.SystemColors.Window
		Me.cboLCLayer.CausesValidation = True
		Me.cboLCLayer.Enabled = True
		Me.cboLCLayer.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboLCLayer.IntegralHeight = True
		Me.cboLCLayer.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboLCLayer.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboLCLayer.Sorted = False
		Me.cboLCLayer.TabStop = True
		Me.cboLCLayer.Visible = True
		Me.cboLCLayer.Name = "cboLCLayer"
		Me.cboLCType.Size = New System.Drawing.Size(141, 21)
		Me.cboLCType.Location = New System.Drawing.Point(69, 80)
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
		Me.cboLCType.Sorted = False
		Me.cboLCType.TabStop = True
		Me.cboLCType.Visible = True
		Me.cboLCType.Name = "cboLCType"
		Me._Label1_5.Text = "Grid Units:"
		Me._Label1_5.Size = New System.Drawing.Size(52, 19)
		Me._Label1_5.Location = New System.Drawing.Point(10, 56)
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
		Me._Label1_0.Text = "Grid:"
		Me._Label1_0.Size = New System.Drawing.Size(32, 19)
		Me._Label1_0.Location = New System.Drawing.Point(10, 28)
		Me._Label1_0.TabIndex = 7
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
		Me._Label1_2.Text = "Type:"
		Me._Label1_2.Size = New System.Drawing.Size(36, 19)
		Me._Label1_2.Location = New System.Drawing.Point(10, 81)
		Me._Label1_2.TabIndex = 6
		Me._Label1_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_2.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_2.Enabled = True
		Me._Label1_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_2.UseMnemonic = True
		Me._Label1_2.Visible = True
		Me._Label1_2.AutoSize = False
		Me._Label1_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_2.Name = "_Label1_2"
		Me.cboAreaLayer.Size = New System.Drawing.Size(97, 21)
		Me.cboAreaLayer.Location = New System.Drawing.Point(231, 480)
		Me.cboAreaLayer.TabIndex = 33
		Me.cboAreaLayer.Visible = False
		Me.cboAreaLayer.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboAreaLayer.BackColor = System.Drawing.SystemColors.Window
		Me.cboAreaLayer.CausesValidation = True
		Me.cboAreaLayer.Enabled = True
		Me.cboAreaLayer.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboAreaLayer.IntegralHeight = True
		Me.cboAreaLayer.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboAreaLayer.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboAreaLayer.Sorted = False
		Me.cboAreaLayer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboAreaLayer.TabStop = True
		Me.cboAreaLayer.Name = "cboAreaLayer"
		Me.cboClass.Size = New System.Drawing.Size(81, 21)
		Me.cboClass.Location = New System.Drawing.Point(110, 475)
		Me.cboClass.TabIndex = 32
		Me.cboClass.Text = "cboClass"
		Me.cboClass.Visible = False
		Me.cboClass.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboClass.BackColor = System.Drawing.SystemColors.Window
		Me.cboClass.CausesValidation = True
		Me.cboClass.Enabled = True
		Me.cboClass.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboClass.IntegralHeight = True
		Me.cboClass.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboClass.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboClass.Sorted = False
		Me.cboClass.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboClass.TabStop = True
		Me.cboClass.Name = "cboClass"
		Me.Timer1.Enabled = False
		Me.Timer1.Interval = 1
		Me.cboCoeffSet.Size = New System.Drawing.Size(97, 21)
		Me.cboCoeffSet.Location = New System.Drawing.Point(100, 474)
		Me.cboCoeffSet.TabIndex = 31
		Me.cboCoeffSet.Text = "cboCoeffSet"
		Me.cboCoeffSet.Visible = False
		Me.cboCoeffSet.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboCoeffSet.BackColor = System.Drawing.SystemColors.Window
		Me.cboCoeffSet.CausesValidation = True
		Me.cboCoeffSet.Enabled = True
		Me.cboCoeffSet.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboCoeffSet.IntegralHeight = True
		Me.cboCoeffSet.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboCoeffSet.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboCoeffSet.Sorted = False
		Me.cboCoeffSet.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboCoeffSet.TabStop = True
		Me.cboCoeffSet.Name = "cboCoeffSet"
		Me.cboCoeff.Size = New System.Drawing.Size(113, 21)
		Me.cboCoeff.Location = New System.Drawing.Point(94, 476)
		Me.cboCoeff.Items.AddRange(New Object(){"Type 1", "Type 2", "Type 3", "Type 4"})
		Me.cboCoeff.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cboCoeff.TabIndex = 14
		Me.cboCoeff.Visible = False
		Me.cboCoeff.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboCoeff.BackColor = System.Drawing.SystemColors.Window
		Me.cboCoeff.CausesValidation = True
		Me.cboCoeff.Enabled = True
		Me.cboCoeff.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboCoeff.IntegralHeight = True
		Me.cboCoeff.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboCoeff.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboCoeff.Sorted = False
		Me.cboCoeff.TabStop = True
		Me.cboCoeff.Name = "cboCoeff"
		Me.SSTab1.Size = New System.Drawing.Size(618, 203)
		Me.SSTab1.Location = New System.Drawing.Point(16, 256)
		Me.SSTab1.TabIndex = 12
		Me.SSTab1.Appearance = System.Windows.Forms.TabAppearance.FlatButtons
		Me.SSTab1.SelectedIndex = 3
		Me.SSTab1.ItemSize = New System.Drawing.Size(42, 18)
		Me.SSTab1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.SSTab1.Name = "SSTab1"
		Me._SSTab1_TabPage0.Text = "Pollutants"
		grdCoeffs.OcxState = CType(resources.GetObject("grdCoeffs.OcxState"), System.Windows.Forms.AxHost.State)
		Me.grdCoeffs.Size = New System.Drawing.Size(569, 153)
		Me.grdCoeffs.Location = New System.Drawing.Point(12, 32)
		Me.grdCoeffs.TabIndex = 27
		Me.grdCoeffs.Name = "grdCoeffs"
		Me._SSTab1_TabPage1.Text = "Erosion"
		Me.lblErodFactor.Text = "Erodibility Factor Attribute: "
		Me.lblErodFactor.Size = New System.Drawing.Size(129, 15)
		Me.lblErodFactor.Location = New System.Drawing.Point(440, 184)
		Me.lblErodFactor.TabIndex = 28
		Me.lblErodFactor.Visible = False
		Me.lblErodFactor.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblErodFactor.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblErodFactor.BackColor = System.Drawing.SystemColors.Control
		Me.lblErodFactor.Enabled = True
		Me.lblErodFactor.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblErodFactor.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblErodFactor.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblErodFactor.UseMnemonic = True
		Me.lblErodFactor.AutoSize = False
		Me.lblErodFactor.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblErodFactor.Name = "lblErodFactor"
		Me.lblKFactor.Size = New System.Drawing.Size(254, 21)
		Me.lblKFactor.Location = New System.Drawing.Point(105, 61)
		Me.lblKFactor.TabIndex = 43
		Me.lblKFactor.Visible = False
		Me.lblKFactor.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblKFactor.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblKFactor.BackColor = System.Drawing.SystemColors.Control
		Me.lblKFactor.Enabled = True
		Me.lblKFactor.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblKFactor.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblKFactor.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblKFactor.UseMnemonic = True
		Me.lblKFactor.AutoSize = False
		Me.lblKFactor.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblKFactor.Name = "lblKFactor"
		Me.Label3.Text = "Hydrologic Soil Group Attribute:"
		Me.Label3.Size = New System.Drawing.Size(180, 18)
		Me.Label3.Location = New System.Drawing.Point(16, 176)
		Me.Label3.TabIndex = 45
		Me.Label3.Visible = False
		Me.Label3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label3.BackColor = System.Drawing.SystemColors.Control
		Me.Label3.Enabled = True
		Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label3.UseMnemonic = True
		Me.Label3.AutoSize = False
		Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label3.Name = "Label3"
		Me.Label7.Text = "K Factor Dataset:"
		Me.Label7.Size = New System.Drawing.Size(84, 21)
		Me.Label7.Location = New System.Drawing.Point(16, 61)
		Me.Label7.TabIndex = 48
		Me.Label7.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label7.BackColor = System.Drawing.SystemColors.Control
		Me.Label7.Enabled = True
		Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label7.UseMnemonic = True
		Me.Label7.Visible = True
		Me.Label7.AutoSize = False
		Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label7.Name = "Label7"
		Me.chkCalcErosion.Text = "Calculate Erosion for Annual Type Precipitation Scenario"
		Me.chkCalcErosion.Size = New System.Drawing.Size(331, 19)
		Me.chkCalcErosion.Location = New System.Drawing.Point(14, 37)
		Me.chkCalcErosion.TabIndex = 15
		Me.chkCalcErosion.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkCalcErosion.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkCalcErosion.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkCalcErosion.BackColor = System.Drawing.SystemColors.Control
		Me.chkCalcErosion.CausesValidation = True
		Me.chkCalcErosion.Enabled = True
		Me.chkCalcErosion.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkCalcErosion.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkCalcErosion.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkCalcErosion.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkCalcErosion.TabStop = True
		Me.chkCalcErosion.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkCalcErosion.Visible = True
		Me.chkCalcErosion.Name = "chkCalcErosion"
		Me.frameRainFall.Text = "Rainfall Factor "
		Me.frameRainFall.Size = New System.Drawing.Size(268, 89)
		Me.frameRainFall.Location = New System.Drawing.Point(13, 86)
		Me.frameRainFall.TabIndex = 22
		Me.frameRainFall.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frameRainFall.BackColor = System.Drawing.SystemColors.Control
		Me.frameRainFall.Enabled = True
		Me.frameRainFall.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frameRainFall.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frameRainFall.Visible = True
		Me.frameRainFall.Padding = New System.Windows.Forms.Padding(0)
		Me.frameRainFall.Name = "frameRainFall"
		Me.txtRainValue.AutoSize = False
		Me.txtRainValue.Size = New System.Drawing.Size(84, 19)
		Me.txtRainValue.Location = New System.Drawing.Point(139, 57)
		Me.txtRainValue.TabIndex = 26
		Me.ToolTip1.SetToolTip(Me.txtRainValue, "Functionality to be implemented in Alpha2")
		Me.txtRainValue.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtRainValue.AcceptsReturn = True
		Me.txtRainValue.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtRainValue.BackColor = System.Drawing.SystemColors.Window
		Me.txtRainValue.CausesValidation = True
		Me.txtRainValue.Enabled = True
		Me.txtRainValue.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtRainValue.HideSelection = True
		Me.txtRainValue.ReadOnly = False
		Me.txtRainValue.Maxlength = 0
		Me.txtRainValue.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtRainValue.MultiLine = False
		Me.txtRainValue.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtRainValue.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtRainValue.TabStop = True
		Me.txtRainValue.Visible = True
		Me.txtRainValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtRainValue.Name = "txtRainValue"
		Me.cboRainGrid.Enabled = False
		Me.cboRainGrid.Size = New System.Drawing.Size(122, 21)
		Me.cboRainGrid.Location = New System.Drawing.Point(100, 25)
		Me.cboRainGrid.TabIndex = 25
		Me.cboRainGrid.Text = "cboRainGRID"
		Me.cboRainGrid.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboRainGrid.BackColor = System.Drawing.SystemColors.Window
		Me.cboRainGrid.CausesValidation = True
		Me.cboRainGrid.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboRainGrid.IntegralHeight = True
		Me.cboRainGrid.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboRainGrid.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboRainGrid.Sorted = False
		Me.cboRainGrid.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboRainGrid.TabStop = True
		Me.cboRainGrid.Visible = True
		Me.cboRainGrid.Name = "cboRainGrid"
		Me.optUseValue.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optUseValue.Text = "Use Constant Value: "
		Me.optUseValue.Size = New System.Drawing.Size(126, 21)
		Me.optUseValue.Location = New System.Drawing.Point(19, 57)
		Me.optUseValue.TabIndex = 24
		Me.ToolTip1.SetToolTip(Me.optUseValue, "Functionality to be implemented in Alpha2")
		Me.optUseValue.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optUseValue.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optUseValue.BackColor = System.Drawing.SystemColors.Control
		Me.optUseValue.CausesValidation = True
		Me.optUseValue.Enabled = True
		Me.optUseValue.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optUseValue.Cursor = System.Windows.Forms.Cursors.Default
		Me.optUseValue.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optUseValue.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optUseValue.TabStop = True
		Me.optUseValue.Checked = False
		Me.optUseValue.Visible = True
		Me.optUseValue.Name = "optUseValue"
		Me.optUseGRID.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optUseGRID.Text = "Use GRID: "
		Me.optUseGRID.Enabled = False
		Me.optUseGRID.Size = New System.Drawing.Size(91, 13)
		Me.optUseGRID.Location = New System.Drawing.Point(19, 30)
		Me.optUseGRID.TabIndex = 23
		Me.optUseGRID.Checked = True
		Me.optUseGRID.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optUseGRID.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optUseGRID.BackColor = System.Drawing.SystemColors.Control
		Me.optUseGRID.CausesValidation = True
		Me.optUseGRID.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optUseGRID.Cursor = System.Windows.Forms.Cursors.Default
		Me.optUseGRID.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optUseGRID.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optUseGRID.TabStop = True
		Me.optUseGRID.Visible = True
		Me.optUseGRID.Name = "optUseGRID"
		Me.cboErodFactor.Size = New System.Drawing.Size(136, 21)
		Me.cboErodFactor.Location = New System.Drawing.Point(304, 176)
		Me.cboErodFactor.TabIndex = 29
		Me.cboErodFactor.Visible = False
		Me.cboErodFactor.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboErodFactor.BackColor = System.Drawing.SystemColors.Window
		Me.cboErodFactor.CausesValidation = True
		Me.cboErodFactor.Enabled = True
		Me.cboErodFactor.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboErodFactor.IntegralHeight = True
		Me.cboErodFactor.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboErodFactor.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboErodFactor.Sorted = False
		Me.cboErodFactor.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboErodFactor.TabStop = True
		Me.cboErodFactor.Name = "cboErodFactor"
		Me.cboSoilAttribute.Size = New System.Drawing.Size(141, 21)
		Me.cboSoilAttribute.Location = New System.Drawing.Point(168, 176)
		Me.cboSoilAttribute.TabIndex = 44
		Me.cboSoilAttribute.Visible = False
		Me.cboSoilAttribute.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboSoilAttribute.BackColor = System.Drawing.SystemColors.Window
		Me.cboSoilAttribute.CausesValidation = True
		Me.cboSoilAttribute.Enabled = True
		Me.cboSoilAttribute.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboSoilAttribute.IntegralHeight = True
		Me.cboSoilAttribute.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboSoilAttribute.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboSoilAttribute.Sorted = False
		Me.cboSoilAttribute.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboSoilAttribute.TabStop = True
		Me.cboSoilAttribute.Name = "cboSoilAttribute"
		Me.frmSDR.Text = "Frame7"
		Me.frmSDR.Size = New System.Drawing.Size(297, 57)
		Me.frmSDR.Location = New System.Drawing.Point(296, 88)
		Me.frmSDR.TabIndex = 64
		Me.frmSDR.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frmSDR.BackColor = System.Drawing.SystemColors.Control
		Me.frmSDR.Enabled = True
		Me.frmSDR.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frmSDR.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmSDR.Visible = True
		Me.frmSDR.Padding = New System.Windows.Forms.Padding(0)
		Me.frmSDR.Name = "frmSDR"
		Me.cmdOpenSDR.TextAlign = System.Drawing.ContentAlignment.BottomCenter
		Me.cmdOpenSDR.Enabled = False
		Me.cmdOpenSDR.Size = New System.Drawing.Size(23, 20)
		Me.cmdOpenSDR.Location = New System.Drawing.Point(256, 24)
		Me.cmdOpenSDR.Image = CType(resources.GetObject("cmdOpenSDR.Image"), System.Drawing.Image)
		Me.cmdOpenSDR.TabIndex = 67
		Me.cmdOpenSDR.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOpenSDR.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOpenSDR.CausesValidation = True
		Me.cmdOpenSDR.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOpenSDR.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOpenSDR.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOpenSDR.TabStop = True
		Me.cmdOpenSDR.Name = "cmdOpenSDR"
		Me.txtSDRGRID.AutoSize = False
		Me.txtSDRGRID.Enabled = False
		Me.txtSDRGRID.Size = New System.Drawing.Size(233, 19)
		Me.txtSDRGRID.Location = New System.Drawing.Point(16, 24)
		Me.txtSDRGRID.TabIndex = 66
		Me.txtSDRGRID.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtSDRGRID.AcceptsReturn = True
		Me.txtSDRGRID.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtSDRGRID.BackColor = System.Drawing.SystemColors.Window
		Me.txtSDRGRID.CausesValidation = True
		Me.txtSDRGRID.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtSDRGRID.HideSelection = True
		Me.txtSDRGRID.ReadOnly = False
		Me.txtSDRGRID.Maxlength = 0
		Me.txtSDRGRID.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtSDRGRID.MultiLine = False
		Me.txtSDRGRID.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtSDRGRID.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtSDRGRID.TabStop = True
		Me.txtSDRGRID.Visible = True
		Me.txtSDRGRID.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtSDRGRID.Name = "txtSDRGRID"
		Me.chkSDR.Text = "Sediment Delivery Ratio GRID (optional)"
		Me.chkSDR.Size = New System.Drawing.Size(217, 17)
		Me.chkSDR.Location = New System.Drawing.Point(8, 0)
		Me.chkSDR.TabIndex = 65
		Me.chkSDR.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkSDR.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkSDR.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkSDR.BackColor = System.Drawing.SystemColors.Control
		Me.chkSDR.CausesValidation = True
		Me.chkSDR.Enabled = True
		Me.chkSDR.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkSDR.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkSDR.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkSDR.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkSDR.TabStop = True
		Me.chkSDR.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkSDR.Visible = True
		Me.chkSDR.Name = "chkSDR"
		Me._SSTab1_TabPage2.Text = "Land Uses"
		grdLU.OcxState = CType(resources.GetObject("grdLU.OcxState"), System.Windows.Forms.AxHost.State)
		Me.grdLU.Size = New System.Drawing.Size(403, 153)
		Me.grdLU.Location = New System.Drawing.Point(12, 32)
		Me.grdLU.TabIndex = 13
		Me.grdLU.Name = "grdLU"
		Me._SSTab1_TabPage3.Text = "Management Scenarios"
		grdLCChanges.OcxState = CType(resources.GetObject("grdLCChanges.OcxState"), System.Windows.Forms.AxHost.State)
		Me.grdLCChanges.Size = New System.Drawing.Size(524, 153)
		Me.grdLCChanges.Location = New System.Drawing.Point(12, 32)
		Me.grdLCChanges.TabIndex = 34
		Me.grdLCChanges.Name = "grdLCChanges"
		Me.cmdRun.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdRun.Text = "Run"
		Me.cmdRun.CausesValidation = False
		Me.cmdRun.Size = New System.Drawing.Size(72, 25)
		Me.cmdRun.Location = New System.Drawing.Point(459, 471)
		Me.cmdRun.TabIndex = 3
		Me.cmdRun.TabStop = False
		Me.cmdRun.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdRun.BackColor = System.Drawing.SystemColors.Control
		Me.cmdRun.Enabled = True
		Me.cmdRun.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdRun.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdRun.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdRun.Name = "cmdRun"
		Me.cmdQuit.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdQuit.Text = "Cancel"
		Me.cmdQuit.Size = New System.Drawing.Size(72, 25)
		Me.cmdQuit.Location = New System.Drawing.Point(537, 471)
		Me.cmdQuit.TabIndex = 4
		Me.cmdQuit.TabStop = False
		Me.cmdQuit.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdQuit.BackColor = System.Drawing.SystemColors.Control
		Me.cmdQuit.CausesValidation = True
		Me.cmdQuit.Enabled = True
		Me.cmdQuit.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdQuit.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdQuit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdQuit.Name = "cmdQuit"
		Me._chkIgnore_0.FlatStyle = System.Windows.Forms.FlatStyle.Flat
		Me._chkIgnore_0.BackColor = System.Drawing.SystemColors.Window
		Me._chkIgnore_0.Text = "Check2"
		Me._chkIgnore_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._chkIgnore_0.Size = New System.Drawing.Size(15, 13)
		Me._chkIgnore_0.Location = New System.Drawing.Point(77, 383)
		Me._chkIgnore_0.TabIndex = 30
		Me._chkIgnore_0.Visible = False
		Me._chkIgnore_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkIgnore_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkIgnore_0.CausesValidation = True
		Me._chkIgnore_0.Enabled = True
		Me._chkIgnore_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkIgnore_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkIgnore_0.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkIgnore_0.TabStop = True
		Me._chkIgnore_0.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._chkIgnore_0.Name = "_chkIgnore_0"
		Me._chkIgnoreMgmt_0.FlatStyle = System.Windows.Forms.FlatStyle.Flat
		Me._chkIgnoreMgmt_0.BackColor = System.Drawing.SystemColors.Window
		Me._chkIgnoreMgmt_0.Text = "Check2"
		Me._chkIgnoreMgmt_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._chkIgnoreMgmt_0.Size = New System.Drawing.Size(15, 13)
		Me._chkIgnoreMgmt_0.Location = New System.Drawing.Point(535, 372)
		Me._chkIgnoreMgmt_0.TabIndex = 35
		Me._chkIgnoreMgmt_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkIgnoreMgmt_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkIgnoreMgmt_0.CausesValidation = True
		Me._chkIgnoreMgmt_0.Enabled = True
		Me._chkIgnoreMgmt_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkIgnoreMgmt_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkIgnoreMgmt_0.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkIgnoreMgmt_0.TabStop = True
		Me._chkIgnoreMgmt_0.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._chkIgnoreMgmt_0.Visible = True
		Me._chkIgnoreMgmt_0.Name = "_chkIgnoreMgmt_0"
		Me._chkIgnoreLU_0.FlatStyle = System.Windows.Forms.FlatStyle.Flat
		Me._chkIgnoreLU_0.BackColor = System.Drawing.SystemColors.Window
		Me._chkIgnoreLU_0.Text = "Check2"
		Me._chkIgnoreLU_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._chkIgnoreLU_0.Size = New System.Drawing.Size(15, 13)
		Me._chkIgnoreLU_0.Location = New System.Drawing.Point(467, 384)
		Me._chkIgnoreLU_0.TabIndex = 36
		Me._chkIgnoreLU_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkIgnoreLU_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkIgnoreLU_0.CausesValidation = True
		Me._chkIgnoreLU_0.Enabled = True
		Me._chkIgnoreLU_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkIgnoreLU_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkIgnoreLU_0.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkIgnoreLU_0.TabStop = True
		Me._chkIgnoreLU_0.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._chkIgnoreLU_0.Visible = True
		Me._chkIgnoreLU_0.Name = "_chkIgnoreLU_0"
		Me.Picture2.Size = New System.Drawing.Size(305, 215)
		Me.Picture2.Location = New System.Drawing.Point(336, 72)
		Me.Picture2.TabIndex = 49
		Me.Picture2.Visible = False
		Me.Picture2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Picture2.Dock = System.Windows.Forms.DockStyle.None
		Me.Picture2.BackColor = System.Drawing.SystemColors.Control
		Me.Picture2.CausesValidation = True
		Me.Picture2.Enabled = True
		Me.Picture2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Picture2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Picture2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Picture2.TabStop = True
		Me.Picture2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.Picture2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Picture2.Name = "Picture2"
		Me.Frame2.Text = "Results Output "
		Me.Frame2.Size = New System.Drawing.Size(614, 52)
		Me.Frame2.Location = New System.Drawing.Point(12, 407)
		Me.Frame2.TabIndex = 16
		Me.Frame2.Visible = False
		Me.Frame2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame2.BackColor = System.Drawing.SystemColors.Control
		Me.Frame2.Enabled = True
		Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame2.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame2.Name = "Frame2"
		Me.txtThemeName.AutoSize = False
		Me.txtThemeName.Size = New System.Drawing.Size(154, 19)
		Me.txtThemeName.Location = New System.Drawing.Point(422, 22)
		Me.txtThemeName.Maxlength = 30
		Me.txtThemeName.TabIndex = 19
		Me.txtThemeName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtThemeName.AcceptsReturn = True
		Me.txtThemeName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtThemeName.BackColor = System.Drawing.SystemColors.Window
		Me.txtThemeName.CausesValidation = True
		Me.txtThemeName.Enabled = True
		Me.txtThemeName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtThemeName.HideSelection = True
		Me.txtThemeName.ReadOnly = False
		Me.txtThemeName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtThemeName.MultiLine = False
		Me.txtThemeName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtThemeName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtThemeName.TabStop = True
		Me.txtThemeName.Visible = True
		Me.txtThemeName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtThemeName.Name = "txtThemeName"
		Me.cmdOutputBrowse.TextAlign = System.Drawing.ContentAlignment.BottomCenter
		Me.cmdOutputBrowse.Enabled = False
		Me.cmdOutputBrowse.Size = New System.Drawing.Size(23, 20)
		Me.cmdOutputBrowse.Location = New System.Drawing.Point(298, 21)
		Me.cmdOutputBrowse.Image = CType(resources.GetObject("cmdOutputBrowse.Image"), System.Drawing.Image)
		Me.cmdOutputBrowse.TabIndex = 18
		Me.cmdOutputBrowse.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOutputBrowse.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOutputBrowse.CausesValidation = True
		Me.cmdOutputBrowse.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOutputBrowse.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOutputBrowse.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOutputBrowse.TabStop = True
		Me.cmdOutputBrowse.Name = "cmdOutputBrowse"
		Me.txtOutputFile.AutoSize = False
		Me.txtOutputFile.Enabled = False
		Me.txtOutputFile.Size = New System.Drawing.Size(196, 19)
		Me.txtOutputFile.Location = New System.Drawing.Point(100, 22)
		Me.txtOutputFile.TabIndex = 17
		Me.txtOutputFile.Text = "C:\Temp\test.shp"
		Me.ToolTip1.SetToolTip(Me.txtOutputFile, "Functionality to be finalized in Alpha2")
		Me.txtOutputFile.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtOutputFile.AcceptsReturn = True
		Me.txtOutputFile.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtOutputFile.BackColor = System.Drawing.SystemColors.Window
		Me.txtOutputFile.CausesValidation = True
		Me.txtOutputFile.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtOutputFile.HideSelection = True
		Me.txtOutputFile.ReadOnly = False
		Me.txtOutputFile.Maxlength = 0
		Me.txtOutputFile.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtOutputFile.MultiLine = False
		Me.txtOutputFile.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtOutputFile.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtOutputFile.TabStop = True
		Me.txtOutputFile.Visible = True
		Me.txtOutputFile.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtOutputFile.Name = "txtOutputFile"
		Me._Label1_11.Text = "Layer Name:"
		Me._Label1_11.Size = New System.Drawing.Size(73, 14)
		Me._Label1_11.Location = New System.Drawing.Point(349, 22)
		Me._Label1_11.TabIndex = 21
		Me._Label1_11.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_11.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_11.BackColor = System.Drawing.Color.Transparent
		Me._Label1_11.Enabled = True
		Me._Label1_11.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_11.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_11.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_11.UseMnemonic = True
		Me._Label1_11.Visible = True
		Me._Label1_11.AutoSize = False
		Me._Label1_11.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_11.Name = "_Label1_11"
		Me._Label1_12.Text = "Output Shapefile:"
		Me._Label1_12.Size = New System.Drawing.Size(100, 19)
		Me._Label1_12.Location = New System.Drawing.Point(6, 22)
		Me._Label1_12.TabIndex = 20
		Me.ToolTip1.SetToolTip(Me._Label1_12, "Functionality to be finalized in Alpha2")
		Me._Label1_12.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_12.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_12.BackColor = System.Drawing.Color.Transparent
		Me._Label1_12.Enabled = True
		Me._Label1_12.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_12.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_12.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_12.UseMnemonic = True
		Me._Label1_12.Visible = True
		Me._Label1_12.AutoSize = False
		Me._Label1_12.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_12.Name = "_Label1_12"
		Me.Controls.Add(Frame6)
		Me.Controls.Add(Frame5)
		Me.Controls.Add(Frame4)
		Me.Controls.Add(frm_raintype)
		Me.Controls.Add(Frame3)
		Me.Controls.Add(Frame1)
		Me.Controls.Add(fraLC)
		Me.Controls.Add(cboAreaLayer)
		Me.Controls.Add(cboClass)
		Me.Controls.Add(cboCoeffSet)
		Me.Controls.Add(cboCoeff)
		Me.Controls.Add(SSTab1)
		Me.Controls.Add(cmdRun)
		Me.Controls.Add(cmdQuit)
		Me.Controls.Add(_chkIgnore_0)
		Me.Controls.Add(_chkIgnoreMgmt_0)
		Me.Controls.Add(_chkIgnoreLU_0)
		Me.Controls.Add(Picture2)
		Me.Controls.Add(Frame2)
		Me.Frame6.Controls.Add(cboWQStd)
		Me.Frame6.Controls.Add(_Label1_6)
		Me.Frame5.Controls.Add(cboWSDelin)
		Me.Frame5.Controls.Add(_Label1_3)
		Me.Frame4.Controls.Add(cboPrecipScen)
		Me.Frame4.Controls.Add(_Label1_7)
		Me.frm_raintype.Controls.Add(cboSelectPoly)
		Me.frm_raintype.Controls.Add(chkSelectedPolys)
		Me.frm_raintype.Controls.Add(chkLocalEffects)
		Me.frm_raintype.Controls.Add(lblLayer)
		Me.Frame3.Controls.Add(cmdOpenWS)
		Me.Frame3.Controls.Add(txtOutputWS)
		Me.Frame3.Controls.Add(txtProjectName)
		Me.Frame3.Controls.Add(Label5)
		Me.Frame3.Controls.Add(Label4)
		Me.Frame1.Controls.Add(cboSoilsLayer)
		Me.Frame1.Controls.Add(Label6)
		Me.Frame1.Controls.Add(lblSoilsHyd)
		Me.Frame1.Controls.Add(Label2)
		Me.fraLC.Controls.Add(cboLCUnits)
		Me.fraLC.Controls.Add(cboLCLayer)
		Me.fraLC.Controls.Add(cboLCType)
		Me.fraLC.Controls.Add(_Label1_5)
		Me.fraLC.Controls.Add(_Label1_0)
		Me.fraLC.Controls.Add(_Label1_2)
		Me.SSTab1.Controls.Add(_SSTab1_TabPage0)
		Me.SSTab1.Controls.Add(_SSTab1_TabPage1)
		Me.SSTab1.Controls.Add(_SSTab1_TabPage2)
		Me.SSTab1.Controls.Add(_SSTab1_TabPage3)
		Me._SSTab1_TabPage0.Controls.Add(grdCoeffs)
		Me._SSTab1_TabPage1.Controls.Add(lblErodFactor)
		Me._SSTab1_TabPage1.Controls.Add(lblKFactor)
		Me._SSTab1_TabPage1.Controls.Add(Label3)
		Me._SSTab1_TabPage1.Controls.Add(Label7)
		Me._SSTab1_TabPage1.Controls.Add(chkCalcErosion)
		Me._SSTab1_TabPage1.Controls.Add(frameRainFall)
		Me._SSTab1_TabPage1.Controls.Add(cboErodFactor)
		Me._SSTab1_TabPage1.Controls.Add(cboSoilAttribute)
		Me._SSTab1_TabPage1.Controls.Add(frmSDR)
		Me.frameRainFall.Controls.Add(txtRainValue)
		Me.frameRainFall.Controls.Add(cboRainGrid)
		Me.frameRainFall.Controls.Add(optUseValue)
		Me.frameRainFall.Controls.Add(optUseGRID)
		Me.frmSDR.Controls.Add(cmdOpenSDR)
		Me.frmSDR.Controls.Add(txtSDRGRID)
		Me.frmSDR.Controls.Add(chkSDR)
		Me._SSTab1_TabPage2.Controls.Add(grdLU)
		Me._SSTab1_TabPage3.Controls.Add(grdLCChanges)
		Me.Frame2.Controls.Add(txtThemeName)
		Me.Frame2.Controls.Add(cmdOutputBrowse)
		Me.Frame2.Controls.Add(txtOutputFile)
		Me.Frame2.Controls.Add(_Label1_11)
		Me.Frame2.Controls.Add(_Label1_12)
		Me.Label1.SetIndex(_Label1_6, CType(6, Short))
		Me.Label1.SetIndex(_Label1_3, CType(3, Short))
		Me.Label1.SetIndex(_Label1_7, CType(7, Short))
		Me.Label1.SetIndex(_Label1_5, CType(5, Short))
		Me.Label1.SetIndex(_Label1_0, CType(0, Short))
		Me.Label1.SetIndex(_Label1_2, CType(2, Short))
		Me.Label1.SetIndex(_Label1_11, CType(11, Short))
		Me.Label1.SetIndex(_Label1_12, CType(12, Short))
		Me.chkIgnore.SetIndex(_chkIgnore_0, CType(0, Short))
		Me.chkIgnoreLU.SetIndex(_chkIgnoreLU_0, CType(0, Short))
		Me.chkIgnoreMgmt.SetIndex(_chkIgnoreMgmt_0, CType(0, Short))
		CType(Me.chkIgnoreMgmt, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.chkIgnoreLU, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.chkIgnore, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.grdLCChanges, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.grdLU, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.grdCoeffs, System.ComponentModel.ISupportInitialize).EndInit()
		MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuFile, Me.mnuBigHelp, Me.mnuOptions, Me.mnuManagement})
		mnuFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuNew, Me.mnuOpen, Me.mnuSpace, Me.mnuSave, Me.mnuSaveAs, Me.mnuSpace1, Me.mnuExit})
		mnuBigHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuGeneralHelp})
		mnuOptions.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuLUAdd, Me.mnuLUEdit, Me.mnuLUDelete})
		mnuManagement.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuManAppen, Me.mnuManInsert, Me.mnuManDelete})
		Me.Controls.Add(MainMenu1)
		Me.MainMenu1.ResumeLayout(False)
		Me.Frame6.ResumeLayout(False)
		Me.Frame5.ResumeLayout(False)
		Me.Frame4.ResumeLayout(False)
		Me.frm_raintype.ResumeLayout(False)
		Me.Frame3.ResumeLayout(False)
		Me.Frame1.ResumeLayout(False)
		Me.fraLC.ResumeLayout(False)
		Me.SSTab1.ResumeLayout(False)
		Me._SSTab1_TabPage0.ResumeLayout(False)
		Me._SSTab1_TabPage1.ResumeLayout(False)
		Me.frameRainFall.ResumeLayout(False)
		Me.frmSDR.ResumeLayout(False)
		Me._SSTab1_TabPage2.ResumeLayout(False)
		Me._SSTab1_TabPage3.ResumeLayout(False)
		Me.Frame2.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class