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
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPrj))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.chkSelectedPolys = New System.Windows.Forms.CheckBox
        Me.chkLocalEffects = New System.Windows.Forms.CheckBox
        Me.cboLCLayer = New System.Windows.Forms.ComboBox
        Me.txtRainValue = New System.Windows.Forms.TextBox
        Me.optUseValue = New System.Windows.Forms.RadioButton
        Me.txtOutputFile = New System.Windows.Forms.TextBox
        Me._Label1_12 = New System.Windows.Forms.Label
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
        Me.cboLCType = New System.Windows.Forms.ComboBox
        Me._Label1_5 = New System.Windows.Forms.Label
        Me._Label1_0 = New System.Windows.Forms.Label
        Me._Label1_2 = New System.Windows.Forms.Label
        Me.cboAreaLayer = New System.Windows.Forms.ComboBox
        Me.cboClass = New System.Windows.Forms.ComboBox
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
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
        Me.cboRainGrid = New System.Windows.Forms.ComboBox
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
        Me._Label1_11 = New System.Windows.Forms.Label
        Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.chkIgnore = New Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray(Me.components)
        Me.chkIgnoreLU = New Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray(Me.components)
        Me.chkIgnoreMgmt = New Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray(Me.components)
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
        CType(Me.grdCoeffs, System.ComponentModel.ISupportInitialize).BeginInit()
        Me._SSTab1_TabPage1.SuspendLayout()
        Me.frameRainFall.SuspendLayout()
        Me.frmSDR.SuspendLayout()
        Me._SSTab1_TabPage2.SuspendLayout()
        CType(Me.grdLU, System.ComponentModel.ISupportInitialize).BeginInit()
        Me._SSTab1_TabPage3.SuspendLayout()
        CType(Me.grdLCChanges, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Picture2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Frame2.SuspendLayout()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkIgnore, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkIgnoreLU, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkIgnoreMgmt, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'chkSelectedPolys
        '
        Me.chkSelectedPolys.BackColor = System.Drawing.SystemColors.Control
        Me.chkSelectedPolys.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkSelectedPolys.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkSelectedPolys.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkSelectedPolys.Location = New System.Drawing.Point(7, 22)
        Me.chkSelectedPolys.Name = "chkSelectedPolys"
        Me.chkSelectedPolys.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkSelectedPolys.Size = New System.Drawing.Size(147, 22)
        Me.chkSelectedPolys.TabIndex = 61
        Me.chkSelectedPolys.Text = "Selected Polygons Only"
        Me.ToolTip1.SetToolTip(Me.chkSelectedPolys, "Select to limit analysis to selected polygons from a map layer")
        Me.chkSelectedPolys.UseVisualStyleBackColor = False
        '
        'chkLocalEffects
        '
        Me.chkLocalEffects.BackColor = System.Drawing.SystemColors.Control
        Me.chkLocalEffects.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkLocalEffects.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkLocalEffects.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkLocalEffects.Location = New System.Drawing.Point(8, 86)
        Me.chkLocalEffects.Name = "chkLocalEffects"
        Me.chkLocalEffects.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkLocalEffects.Size = New System.Drawing.Size(131, 17)
        Me.chkLocalEffects.TabIndex = 51
        Me.chkLocalEffects.Text = "Local Effects Only"
        Me.ToolTip1.SetToolTip(Me.chkLocalEffects, "Select for analysis of local effects only")
        Me.chkLocalEffects.UseVisualStyleBackColor = False
        '
        'cboLCLayer
        '
        Me.cboLCLayer.BackColor = System.Drawing.SystemColors.Window
        Me.cboLCLayer.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboLCLayer.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboLCLayer.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboLCLayer.Location = New System.Drawing.Point(69, 23)
        Me.cboLCLayer.Name = "cboLCLayer"
        Me.cboLCLayer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboLCLayer.Size = New System.Drawing.Size(141, 22)
        Me.cboLCLayer.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.cboLCLayer, "Choose a Land Cover GRID from current layers in map view")
        '
        'txtRainValue
        '
        Me.txtRainValue.AcceptsReturn = True
        Me.txtRainValue.BackColor = System.Drawing.SystemColors.Window
        Me.txtRainValue.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRainValue.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRainValue.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRainValue.Location = New System.Drawing.Point(139, 57)
        Me.txtRainValue.MaxLength = 0
        Me.txtRainValue.Name = "txtRainValue"
        Me.txtRainValue.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRainValue.Size = New System.Drawing.Size(84, 19)
        Me.txtRainValue.TabIndex = 26
        Me.ToolTip1.SetToolTip(Me.txtRainValue, "Functionality to be implemented in Alpha2")
        '
        'optUseValue
        '
        Me.optUseValue.BackColor = System.Drawing.SystemColors.Control
        Me.optUseValue.Cursor = System.Windows.Forms.Cursors.Default
        Me.optUseValue.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optUseValue.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optUseValue.Location = New System.Drawing.Point(19, 57)
        Me.optUseValue.Name = "optUseValue"
        Me.optUseValue.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optUseValue.Size = New System.Drawing.Size(126, 21)
        Me.optUseValue.TabIndex = 24
        Me.optUseValue.TabStop = True
        Me.optUseValue.Text = "Use Constant Value: "
        Me.ToolTip1.SetToolTip(Me.optUseValue, "Functionality to be implemented in Alpha2")
        Me.optUseValue.UseVisualStyleBackColor = False
        '
        'txtOutputFile
        '
        Me.txtOutputFile.AcceptsReturn = True
        Me.txtOutputFile.BackColor = System.Drawing.SystemColors.Window
        Me.txtOutputFile.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtOutputFile.Enabled = False
        Me.txtOutputFile.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOutputFile.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtOutputFile.Location = New System.Drawing.Point(100, 22)
        Me.txtOutputFile.MaxLength = 0
        Me.txtOutputFile.Name = "txtOutputFile"
        Me.txtOutputFile.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtOutputFile.Size = New System.Drawing.Size(196, 19)
        Me.txtOutputFile.TabIndex = 17
        Me.txtOutputFile.Text = "C:\Temp\test.shp"
        Me.ToolTip1.SetToolTip(Me.txtOutputFile, "Functionality to be finalized in Alpha2")
        '
        '_Label1_12
        '
        Me._Label1_12.BackColor = System.Drawing.Color.Transparent
        Me._Label1_12.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_12.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_12.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_12, CType(12, Short))
        Me._Label1_12.Location = New System.Drawing.Point(6, 22)
        Me._Label1_12.Name = "_Label1_12"
        Me._Label1_12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_12.Size = New System.Drawing.Size(100, 19)
        Me._Label1_12.TabIndex = 20
        Me._Label1_12.Text = "Output Shapefile:"
        Me.ToolTip1.SetToolTip(Me._Label1_12, "Functionality to be finalized in Alpha2")
        '
        'MainMenu1
        '
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFile, Me.mnuBigHelp, Me.mnuOptions, Me.mnuManagement})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(642, 24)
        Me.MainMenu1.TabIndex = 59
        '
        'mnuFile
        '
        Me.mnuFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuNew, Me.mnuOpen, Me.mnuSpace, Me.mnuSave, Me.mnuSaveAs, Me.mnuSpace1, Me.mnuExit})
        Me.mnuFile.Name = "mnuFile"
        Me.mnuFile.Size = New System.Drawing.Size(37, 20)
        Me.mnuFile.Text = "&File"
        '
        'mnuNew
        '
        Me.mnuNew.Name = "mnuNew"
        Me.mnuNew.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.N), System.Windows.Forms.Keys)
        Me.mnuNew.Size = New System.Drawing.Size(195, 22)
        Me.mnuNew.Text = "New Project"
        '
        'mnuOpen
        '
        Me.mnuOpen.Name = "mnuOpen"
        Me.mnuOpen.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.O), System.Windows.Forms.Keys)
        Me.mnuOpen.Size = New System.Drawing.Size(195, 22)
        Me.mnuOpen.Text = "Open Project..."
        '
        'mnuSpace
        '
        Me.mnuSpace.Name = "mnuSpace"
        Me.mnuSpace.Size = New System.Drawing.Size(192, 6)
        '
        'mnuSave
        '
        Me.mnuSave.Name = "mnuSave"
        Me.mnuSave.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.S), System.Windows.Forms.Keys)
        Me.mnuSave.Size = New System.Drawing.Size(195, 22)
        Me.mnuSave.Text = "Save"
        '
        'mnuSaveAs
        '
        Me.mnuSaveAs.Name = "mnuSaveAs"
        Me.mnuSaveAs.Size = New System.Drawing.Size(195, 22)
        Me.mnuSaveAs.Text = "Save As..."
        '
        'mnuSpace1
        '
        Me.mnuSpace1.Name = "mnuSpace1"
        Me.mnuSpace1.Size = New System.Drawing.Size(192, 6)
        '
        'mnuExit
        '
        Me.mnuExit.Name = "mnuExit"
        Me.mnuExit.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.X), System.Windows.Forms.Keys)
        Me.mnuExit.Size = New System.Drawing.Size(195, 22)
        Me.mnuExit.Text = "E&xit"
        '
        'mnuBigHelp
        '
        Me.mnuBigHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuGeneralHelp})
        Me.mnuBigHelp.Name = "mnuBigHelp"
        Me.mnuBigHelp.Size = New System.Drawing.Size(44, 20)
        Me.mnuBigHelp.Text = "&Help"
        '
        'mnuGeneralHelp
        '
        Me.mnuGeneralHelp.Name = "mnuGeneralHelp"
        Me.mnuGeneralHelp.ShortcutKeys = CType((System.Windows.Forms.Keys.Shift Or System.Windows.Forms.Keys.F1), System.Windows.Forms.Keys)
        Me.mnuGeneralHelp.Size = New System.Drawing.Size(210, 22)
        Me.mnuGeneralHelp.Text = "N-SPECT Help..."
        '
        'mnuOptions
        '
        Me.mnuOptions.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuLUAdd, Me.mnuLUEdit, Me.mnuLUDelete})
        Me.mnuOptions.Name = "mnuOptions"
        Me.mnuOptions.Size = New System.Drawing.Size(12, 20)
        Me.mnuOptions.Visible = False
        '
        'mnuLUAdd
        '
        Me.mnuLUAdd.Name = "mnuLUAdd"
        Me.mnuLUAdd.Size = New System.Drawing.Size(164, 22)
        Me.mnuLUAdd.Text = "Add Scenario..."
        '
        'mnuLUEdit
        '
        Me.mnuLUEdit.Name = "mnuLUEdit"
        Me.mnuLUEdit.Size = New System.Drawing.Size(164, 22)
        Me.mnuLUEdit.Text = "Edit Scenario..."
        '
        'mnuLUDelete
        '
        Me.mnuLUDelete.Name = "mnuLUDelete"
        Me.mnuLUDelete.Size = New System.Drawing.Size(164, 22)
        Me.mnuLUDelete.Text = "Delete Scenario..."
        '
        'mnuManagement
        '
        Me.mnuManagement.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuManAppen, Me.mnuManInsert, Me.mnuManDelete})
        Me.mnuManagement.Name = "mnuManagement"
        Me.mnuManagement.Size = New System.Drawing.Size(12, 20)
        Me.mnuManagement.Visible = False
        '
        'mnuManAppen
        '
        Me.mnuManAppen.Name = "mnuManAppen"
        Me.mnuManAppen.Size = New System.Drawing.Size(176, 22)
        Me.mnuManAppen.Text = "Append Row"
        '
        'mnuManInsert
        '
        Me.mnuManInsert.Name = "mnuManInsert"
        Me.mnuManInsert.Size = New System.Drawing.Size(176, 22)
        Me.mnuManInsert.Text = "Insert Row"
        '
        'mnuManDelete
        '
        Me.mnuManDelete.Name = "mnuManDelete"
        Me.mnuManDelete.Size = New System.Drawing.Size(176, 22)
        Me.mnuManDelete.Text = "Delete Current Row"
        '
        'Frame6
        '
        Me.Frame6.BackColor = System.Drawing.SystemColors.Control
        Me.Frame6.Controls.Add(Me.cboWQStd)
        Me.Frame6.Controls.Add(Me._Label1_6)
        Me.Frame6.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame6.Location = New System.Drawing.Point(426, 192)
        Me.Frame6.Name = "Frame6"
        Me.Frame6.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame6.Size = New System.Drawing.Size(209, 58)
        Me.Frame6.TabIndex = 58
        Me.Frame6.TabStop = False
        Me.Frame6.Text = "Water Quality Standard "
        '
        'cboWQStd
        '
        Me.cboWQStd.BackColor = System.Drawing.SystemColors.Window
        Me.cboWQStd.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboWQStd.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboWQStd.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboWQStd.Location = New System.Drawing.Point(49, 22)
        Me.cboWQStd.Name = "cboWQStd"
        Me.cboWQStd.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboWQStd.Size = New System.Drawing.Size(141, 22)
        Me.cboWQStd.TabIndex = 59
        '
        '_Label1_6
        '
        Me._Label1_6.BackColor = System.Drawing.Color.Transparent
        Me._Label1_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_6.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_6, CType(6, Short))
        Me._Label1_6.Location = New System.Drawing.Point(9, 25)
        Me._Label1_6.Name = "_Label1_6"
        Me._Label1_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_6.Size = New System.Drawing.Size(44, 16)
        Me._Label1_6.TabIndex = 60
        Me._Label1_6.Text = "Name: "
        '
        'Frame5
        '
        Me.Frame5.BackColor = System.Drawing.SystemColors.Control
        Me.Frame5.Controls.Add(Me.cboWSDelin)
        Me.Frame5.Controls.Add(Me._Label1_3)
        Me.Frame5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame5.Location = New System.Drawing.Point(217, 192)
        Me.Frame5.Name = "Frame5"
        Me.Frame5.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame5.Size = New System.Drawing.Size(202, 58)
        Me.Frame5.TabIndex = 55
        Me.Frame5.TabStop = False
        Me.Frame5.Text = "Watershed Delineation "
        '
        'cboWSDelin
        '
        Me.cboWSDelin.BackColor = System.Drawing.SystemColors.Window
        Me.cboWSDelin.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboWSDelin.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboWSDelin.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboWSDelin.Location = New System.Drawing.Point(46, 22)
        Me.cboWSDelin.Name = "cboWSDelin"
        Me.cboWSDelin.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboWSDelin.Size = New System.Drawing.Size(141, 22)
        Me.cboWSDelin.TabIndex = 56
        '
        '_Label1_3
        '
        Me._Label1_3.BackColor = System.Drawing.Color.Transparent
        Me._Label1_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_3, CType(3, Short))
        Me._Label1_3.Location = New System.Drawing.Point(7, 25)
        Me._Label1_3.Name = "_Label1_3"
        Me._Label1_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_3.Size = New System.Drawing.Size(39, 17)
        Me._Label1_3.TabIndex = 57
        Me._Label1_3.Text = "Name:"
        '
        'Frame4
        '
        Me.Frame4.BackColor = System.Drawing.SystemColors.Control
        Me.Frame4.Controls.Add(Me.cboPrecipScen)
        Me.Frame4.Controls.Add(Me._Label1_7)
        Me.Frame4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame4.Location = New System.Drawing.Point(10, 192)
        Me.Frame4.Name = "Frame4"
        Me.Frame4.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame4.Size = New System.Drawing.Size(202, 58)
        Me.Frame4.TabIndex = 52
        Me.Frame4.TabStop = False
        Me.Frame4.Text = "Precipitation Scenario  "
        '
        'cboPrecipScen
        '
        Me.cboPrecipScen.BackColor = System.Drawing.SystemColors.Window
        Me.cboPrecipScen.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboPrecipScen.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboPrecipScen.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboPrecipScen.Location = New System.Drawing.Point(47, 22)
        Me.cboPrecipScen.Name = "cboPrecipScen"
        Me.cboPrecipScen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboPrecipScen.Size = New System.Drawing.Size(141, 22)
        Me.cboPrecipScen.TabIndex = 53
        '
        '_Label1_7
        '
        Me._Label1_7.BackColor = System.Drawing.Color.Transparent
        Me._Label1_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_7.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_7, CType(7, Short))
        Me._Label1_7.Location = New System.Drawing.Point(9, 25)
        Me._Label1_7.Name = "_Label1_7"
        Me._Label1_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_7.Size = New System.Drawing.Size(49, 18)
        Me._Label1_7.TabIndex = 54
        Me._Label1_7.Text = "Name: "
        '
        'frm_raintype
        '
        Me.frm_raintype.BackColor = System.Drawing.SystemColors.Control
        Me.frm_raintype.Controls.Add(Me.cboSelectPoly)
        Me.frm_raintype.Controls.Add(Me.chkSelectedPolys)
        Me.frm_raintype.Controls.Add(Me.chkLocalEffects)
        Me.frm_raintype.Controls.Add(Me.lblLayer)
        Me.frm_raintype.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.frm_raintype.ForeColor = System.Drawing.SystemColors.ControlText
        Me.frm_raintype.Location = New System.Drawing.Point(461, 73)
        Me.frm_raintype.Name = "frm_raintype"
        Me.frm_raintype.Padding = New System.Windows.Forms.Padding(0)
        Me.frm_raintype.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.frm_raintype.Size = New System.Drawing.Size(175, 113)
        Me.frm_raintype.TabIndex = 50
        Me.frm_raintype.TabStop = False
        Me.frm_raintype.Text = "Miscellaneous "
        '
        'cboSelectPoly
        '
        Me.cboSelectPoly.BackColor = System.Drawing.SystemColors.Window
        Me.cboSelectPoly.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboSelectPoly.Enabled = False
        Me.cboSelectPoly.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboSelectPoly.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboSelectPoly.Location = New System.Drawing.Point(45, 53)
        Me.cboSelectPoly.Name = "cboSelectPoly"
        Me.cboSelectPoly.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboSelectPoly.Size = New System.Drawing.Size(112, 22)
        Me.cboSelectPoly.TabIndex = 62
        '
        'lblLayer
        '
        Me.lblLayer.BackColor = System.Drawing.SystemColors.Control
        Me.lblLayer.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblLayer.Enabled = False
        Me.lblLayer.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLayer.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLayer.Location = New System.Drawing.Point(6, 56)
        Me.lblLayer.Name = "lblLayer"
        Me.lblLayer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblLayer.Size = New System.Drawing.Size(44, 15)
        Me.lblLayer.TabIndex = 63
        Me.lblLayer.Text = "Layer: "
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me.cmdOpenWS)
        Me.Frame3.Controls.Add(Me.txtOutputWS)
        Me.Frame3.Controls.Add(Me.txtProjectName)
        Me.Frame3.Controls.Add(Me.Label5)
        Me.Frame3.Controls.Add(Me.Label4)
        Me.Frame3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame3.Location = New System.Drawing.Point(9, 27)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(625, 41)
        Me.Frame3.TabIndex = 37
        Me.Frame3.TabStop = False
        Me.Frame3.Text = "Project Information "
        '
        'cmdOpenWS
        '
        Me.cmdOpenWS.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOpenWS.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOpenWS.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOpenWS.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOpenWS.Image = CType(resources.GetObject("cmdOpenWS.Image"), System.Drawing.Image)
        Me.cmdOpenWS.Location = New System.Drawing.Point(570, 14)
        Me.cmdOpenWS.Name = "cmdOpenWS"
        Me.cmdOpenWS.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOpenWS.Size = New System.Drawing.Size(23, 20)
        Me.cmdOpenWS.TabIndex = 40
        Me.cmdOpenWS.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdOpenWS.UseVisualStyleBackColor = False
        '
        'txtOutputWS
        '
        Me.txtOutputWS.AcceptsReturn = True
        Me.txtOutputWS.BackColor = System.Drawing.SystemColors.Window
        Me.txtOutputWS.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtOutputWS.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOutputWS.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtOutputWS.Location = New System.Drawing.Point(373, 15)
        Me.txtOutputWS.MaxLength = 0
        Me.txtOutputWS.Name = "txtOutputWS"
        Me.txtOutputWS.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtOutputWS.Size = New System.Drawing.Size(192, 20)
        Me.txtOutputWS.TabIndex = 39
        '
        'txtProjectName
        '
        Me.txtProjectName.AcceptsReturn = True
        Me.txtProjectName.BackColor = System.Drawing.SystemColors.Window
        Me.txtProjectName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtProjectName.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtProjectName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtProjectName.Location = New System.Drawing.Point(63, 15)
        Me.txtProjectName.MaxLength = 0
        Me.txtProjectName.Name = "txtProjectName"
        Me.txtProjectName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtProjectName.Size = New System.Drawing.Size(161, 20)
        Me.txtProjectName.TabIndex = 38
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(270, 16)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(99, 18)
        Me.Label5.TabIndex = 42
        Me.Label5.Text = "Working Directory:"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(11, 17)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(54, 16)
        Me.Label4.TabIndex = 41
        Me.Label4.Text = "Name:"
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.cboSoilsLayer)
        Me.Frame1.Controls.Add(Me.Label6)
        Me.Frame1.Controls.Add(Me.lblSoilsHyd)
        Me.Frame1.Controls.Add(Me.Label2)
        Me.Frame1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(240, 73)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(214, 113)
        Me.Frame1.TabIndex = 9
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Soils"
        '
        'cboSoilsLayer
        '
        Me.cboSoilsLayer.BackColor = System.Drawing.SystemColors.Window
        Me.cboSoilsLayer.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboSoilsLayer.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboSoilsLayer.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboSoilsLayer.Location = New System.Drawing.Point(7, 37)
        Me.cboSoilsLayer.Name = "cboSoilsLayer"
        Me.cboSoilsLayer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboSoilsLayer.Size = New System.Drawing.Size(141, 22)
        Me.cboSoilsLayer.TabIndex = 11
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(7, 63)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(157, 16)
        Me.Label6.TabIndex = 47
        Me.Label6.Text = "Hydrologic Soils Data Set:"
        '
        'lblSoilsHyd
        '
        Me.lblSoilsHyd.BackColor = System.Drawing.SystemColors.Control
        Me.lblSoilsHyd.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSoilsHyd.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSoilsHyd.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblSoilsHyd.Location = New System.Drawing.Point(7, 79)
        Me.lblSoilsHyd.Name = "lblSoilsHyd"
        Me.lblSoilsHyd.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSoilsHyd.Size = New System.Drawing.Size(198, 21)
        Me.lblSoilsHyd.TabIndex = 46
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(7, 19)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(136, 17)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "Soils Definition:"
        '
        'fraLC
        '
        Me.fraLC.BackColor = System.Drawing.SystemColors.Control
        Me.fraLC.Controls.Add(Me.cboLCUnits)
        Me.fraLC.Controls.Add(Me.cboLCLayer)
        Me.fraLC.Controls.Add(Me.cboLCType)
        Me.fraLC.Controls.Add(Me._Label1_5)
        Me.fraLC.Controls.Add(Me._Label1_0)
        Me.fraLC.Controls.Add(Me._Label1_2)
        Me.fraLC.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraLC.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraLC.Location = New System.Drawing.Point(9, 73)
        Me.fraLC.Name = "fraLC"
        Me.fraLC.Padding = New System.Windows.Forms.Padding(0)
        Me.fraLC.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraLC.Size = New System.Drawing.Size(226, 113)
        Me.fraLC.TabIndex = 5
        Me.fraLC.TabStop = False
        Me.fraLC.Text = "Land Cover"
        '
        'cboLCUnits
        '
        Me.cboLCUnits.BackColor = System.Drawing.SystemColors.Window
        Me.cboLCUnits.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboLCUnits.Enabled = False
        Me.cboLCUnits.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboLCUnits.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboLCUnits.Items.AddRange(New Object() {"meters", "feet"})
        Me.cboLCUnits.Location = New System.Drawing.Point(69, 51)
        Me.cboLCUnits.Name = "cboLCUnits"
        Me.cboLCUnits.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboLCUnits.Size = New System.Drawing.Size(141, 22)
        Me.cboLCUnits.TabIndex = 1
        '
        'cboLCType
        '
        Me.cboLCType.BackColor = System.Drawing.SystemColors.Window
        Me.cboLCType.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboLCType.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboLCType.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboLCType.Location = New System.Drawing.Point(69, 80)
        Me.cboLCType.Name = "cboLCType"
        Me.cboLCType.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboLCType.Size = New System.Drawing.Size(141, 22)
        Me.cboLCType.TabIndex = 2
        '
        '_Label1_5
        '
        Me._Label1_5.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_5, CType(5, Short))
        Me._Label1_5.Location = New System.Drawing.Point(10, 56)
        Me._Label1_5.Name = "_Label1_5"
        Me._Label1_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_5.Size = New System.Drawing.Size(52, 19)
        Me._Label1_5.TabIndex = 8
        Me._Label1_5.Text = "Grid Units:"
        '
        '_Label1_0
        '
        Me._Label1_0.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_0, CType(0, Short))
        Me._Label1_0.Location = New System.Drawing.Point(10, 28)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_0.Size = New System.Drawing.Size(32, 19)
        Me._Label1_0.TabIndex = 7
        Me._Label1_0.Text = "Grid:"
        '
        '_Label1_2
        '
        Me._Label1_2.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_2, CType(2, Short))
        Me._Label1_2.Location = New System.Drawing.Point(10, 81)
        Me._Label1_2.Name = "_Label1_2"
        Me._Label1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_2.Size = New System.Drawing.Size(36, 19)
        Me._Label1_2.TabIndex = 6
        Me._Label1_2.Text = "Type:"
        '
        'cboAreaLayer
        '
        Me.cboAreaLayer.BackColor = System.Drawing.SystemColors.Window
        Me.cboAreaLayer.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboAreaLayer.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboAreaLayer.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboAreaLayer.Location = New System.Drawing.Point(231, 480)
        Me.cboAreaLayer.Name = "cboAreaLayer"
        Me.cboAreaLayer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboAreaLayer.Size = New System.Drawing.Size(97, 22)
        Me.cboAreaLayer.TabIndex = 33
        Me.cboAreaLayer.Visible = False
        '
        'cboClass
        '
        Me.cboClass.BackColor = System.Drawing.SystemColors.Window
        Me.cboClass.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboClass.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboClass.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboClass.Location = New System.Drawing.Point(110, 475)
        Me.cboClass.Name = "cboClass"
        Me.cboClass.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboClass.Size = New System.Drawing.Size(81, 22)
        Me.cboClass.TabIndex = 32
        Me.cboClass.Text = "cboClass"
        Me.cboClass.Visible = False
        '
        'Timer1
        '
        Me.Timer1.Interval = 1
        '
        'cboCoeffSet
        '
        Me.cboCoeffSet.BackColor = System.Drawing.SystemColors.Window
        Me.cboCoeffSet.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboCoeffSet.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboCoeffSet.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboCoeffSet.Location = New System.Drawing.Point(100, 474)
        Me.cboCoeffSet.Name = "cboCoeffSet"
        Me.cboCoeffSet.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboCoeffSet.Size = New System.Drawing.Size(97, 22)
        Me.cboCoeffSet.TabIndex = 31
        Me.cboCoeffSet.Text = "cboCoeffSet"
        Me.cboCoeffSet.Visible = False
        '
        'cboCoeff
        '
        Me.cboCoeff.BackColor = System.Drawing.SystemColors.Window
        Me.cboCoeff.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboCoeff.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCoeff.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboCoeff.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboCoeff.Items.AddRange(New Object() {"Type 1", "Type 2", "Type 3", "Type 4"})
        Me.cboCoeff.Location = New System.Drawing.Point(94, 476)
        Me.cboCoeff.Name = "cboCoeff"
        Me.cboCoeff.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboCoeff.Size = New System.Drawing.Size(113, 22)
        Me.cboCoeff.TabIndex = 14
        Me.cboCoeff.Visible = False
        '
        'SSTab1
        '
        Me.SSTab1.Appearance = System.Windows.Forms.TabAppearance.FlatButtons
        Me.SSTab1.Controls.Add(Me._SSTab1_TabPage0)
        Me.SSTab1.Controls.Add(Me._SSTab1_TabPage1)
        Me.SSTab1.Controls.Add(Me._SSTab1_TabPage2)
        Me.SSTab1.Controls.Add(Me._SSTab1_TabPage3)
        Me.SSTab1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SSTab1.ItemSize = New System.Drawing.Size(42, 18)
        Me.SSTab1.Location = New System.Drawing.Point(16, 256)
        Me.SSTab1.Name = "SSTab1"
        Me.SSTab1.SelectedIndex = 3
        Me.SSTab1.Size = New System.Drawing.Size(618, 203)
        Me.SSTab1.TabIndex = 12
        '
        '_SSTab1_TabPage0
        '
        Me._SSTab1_TabPage0.Controls.Add(Me.grdCoeffs)
        Me._SSTab1_TabPage0.Location = New System.Drawing.Point(4, 22)
        Me._SSTab1_TabPage0.Name = "_SSTab1_TabPage0"
        Me._SSTab1_TabPage0.Size = New System.Drawing.Size(610, 177)
        Me._SSTab1_TabPage0.TabIndex = 0
        Me._SSTab1_TabPage0.Text = "Pollutants"
        '
        'grdCoeffs
        '
        Me.grdCoeffs.DataSource = Nothing
        Me.grdCoeffs.Location = New System.Drawing.Point(12, 32)
        Me.grdCoeffs.Name = "grdCoeffs"
        Me.grdCoeffs.OcxState = CType(resources.GetObject("grdCoeffs.OcxState"), System.Windows.Forms.AxHost.State)
        Me.grdCoeffs.Size = New System.Drawing.Size(569, 153)
        Me.grdCoeffs.TabIndex = 27
        '
        '_SSTab1_TabPage1
        '
        Me._SSTab1_TabPage1.Controls.Add(Me.lblErodFactor)
        Me._SSTab1_TabPage1.Controls.Add(Me.lblKFactor)
        Me._SSTab1_TabPage1.Controls.Add(Me.Label3)
        Me._SSTab1_TabPage1.Controls.Add(Me.Label7)
        Me._SSTab1_TabPage1.Controls.Add(Me.chkCalcErosion)
        Me._SSTab1_TabPage1.Controls.Add(Me.frameRainFall)
        Me._SSTab1_TabPage1.Controls.Add(Me.cboErodFactor)
        Me._SSTab1_TabPage1.Controls.Add(Me.cboSoilAttribute)
        Me._SSTab1_TabPage1.Controls.Add(Me.frmSDR)
        Me._SSTab1_TabPage1.Location = New System.Drawing.Point(4, 22)
        Me._SSTab1_TabPage1.Name = "_SSTab1_TabPage1"
        Me._SSTab1_TabPage1.Size = New System.Drawing.Size(610, 177)
        Me._SSTab1_TabPage1.TabIndex = 1
        Me._SSTab1_TabPage1.Text = "Erosion"
        '
        'lblErodFactor
        '
        Me.lblErodFactor.BackColor = System.Drawing.SystemColors.Control
        Me.lblErodFactor.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblErodFactor.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblErodFactor.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblErodFactor.Location = New System.Drawing.Point(440, 184)
        Me.lblErodFactor.Name = "lblErodFactor"
        Me.lblErodFactor.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblErodFactor.Size = New System.Drawing.Size(129, 15)
        Me.lblErodFactor.TabIndex = 28
        Me.lblErodFactor.Text = "Erodibility Factor Attribute: "
        Me.lblErodFactor.Visible = False
        '
        'lblKFactor
        '
        Me.lblKFactor.BackColor = System.Drawing.SystemColors.Control
        Me.lblKFactor.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblKFactor.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblKFactor.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblKFactor.Location = New System.Drawing.Point(105, 61)
        Me.lblKFactor.Name = "lblKFactor"
        Me.lblKFactor.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblKFactor.Size = New System.Drawing.Size(254, 21)
        Me.lblKFactor.TabIndex = 43
        Me.lblKFactor.Visible = False
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(16, 176)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(180, 18)
        Me.Label3.TabIndex = 45
        Me.Label3.Text = "Hydrologic Soil Group Attribute:"
        Me.Label3.Visible = False
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(16, 61)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(84, 21)
        Me.Label7.TabIndex = 48
        Me.Label7.Text = "K Factor Dataset:"
        '
        'chkCalcErosion
        '
        Me.chkCalcErosion.BackColor = System.Drawing.SystemColors.Control
        Me.chkCalcErosion.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkCalcErosion.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkCalcErosion.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkCalcErosion.Location = New System.Drawing.Point(14, 37)
        Me.chkCalcErosion.Name = "chkCalcErosion"
        Me.chkCalcErosion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkCalcErosion.Size = New System.Drawing.Size(331, 19)
        Me.chkCalcErosion.TabIndex = 15
        Me.chkCalcErosion.Text = "Calculate Erosion for Annual Type Precipitation Scenario"
        Me.chkCalcErosion.UseVisualStyleBackColor = False
        '
        'frameRainFall
        '
        Me.frameRainFall.BackColor = System.Drawing.SystemColors.Control
        Me.frameRainFall.Controls.Add(Me.txtRainValue)
        Me.frameRainFall.Controls.Add(Me.cboRainGrid)
        Me.frameRainFall.Controls.Add(Me.optUseValue)
        Me.frameRainFall.Controls.Add(Me.optUseGRID)
        Me.frameRainFall.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.frameRainFall.ForeColor = System.Drawing.SystemColors.ControlText
        Me.frameRainFall.Location = New System.Drawing.Point(13, 86)
        Me.frameRainFall.Name = "frameRainFall"
        Me.frameRainFall.Padding = New System.Windows.Forms.Padding(0)
        Me.frameRainFall.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.frameRainFall.Size = New System.Drawing.Size(268, 89)
        Me.frameRainFall.TabIndex = 22
        Me.frameRainFall.TabStop = False
        Me.frameRainFall.Text = "Rainfall Factor "
        '
        'cboRainGrid
        '
        Me.cboRainGrid.BackColor = System.Drawing.SystemColors.Window
        Me.cboRainGrid.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboRainGrid.Enabled = False
        Me.cboRainGrid.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboRainGrid.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboRainGrid.Location = New System.Drawing.Point(100, 25)
        Me.cboRainGrid.Name = "cboRainGrid"
        Me.cboRainGrid.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboRainGrid.Size = New System.Drawing.Size(122, 22)
        Me.cboRainGrid.TabIndex = 25
        Me.cboRainGrid.Text = "cboRainGRID"
        '
        'optUseGRID
        '
        Me.optUseGRID.BackColor = System.Drawing.SystemColors.Control
        Me.optUseGRID.Checked = True
        Me.optUseGRID.Cursor = System.Windows.Forms.Cursors.Default
        Me.optUseGRID.Enabled = False
        Me.optUseGRID.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optUseGRID.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optUseGRID.Location = New System.Drawing.Point(19, 30)
        Me.optUseGRID.Name = "optUseGRID"
        Me.optUseGRID.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optUseGRID.Size = New System.Drawing.Size(91, 13)
        Me.optUseGRID.TabIndex = 23
        Me.optUseGRID.TabStop = True
        Me.optUseGRID.Text = "Use GRID: "
        Me.optUseGRID.UseVisualStyleBackColor = False
        '
        'cboErodFactor
        '
        Me.cboErodFactor.BackColor = System.Drawing.SystemColors.Window
        Me.cboErodFactor.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboErodFactor.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboErodFactor.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboErodFactor.Location = New System.Drawing.Point(304, 176)
        Me.cboErodFactor.Name = "cboErodFactor"
        Me.cboErodFactor.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboErodFactor.Size = New System.Drawing.Size(136, 22)
        Me.cboErodFactor.TabIndex = 29
        Me.cboErodFactor.Visible = False
        '
        'cboSoilAttribute
        '
        Me.cboSoilAttribute.BackColor = System.Drawing.SystemColors.Window
        Me.cboSoilAttribute.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboSoilAttribute.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboSoilAttribute.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboSoilAttribute.Location = New System.Drawing.Point(168, 176)
        Me.cboSoilAttribute.Name = "cboSoilAttribute"
        Me.cboSoilAttribute.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboSoilAttribute.Size = New System.Drawing.Size(141, 22)
        Me.cboSoilAttribute.TabIndex = 44
        Me.cboSoilAttribute.Visible = False
        '
        'frmSDR
        '
        Me.frmSDR.BackColor = System.Drawing.SystemColors.Control
        Me.frmSDR.Controls.Add(Me.cmdOpenSDR)
        Me.frmSDR.Controls.Add(Me.txtSDRGRID)
        Me.frmSDR.Controls.Add(Me.chkSDR)
        Me.frmSDR.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.frmSDR.ForeColor = System.Drawing.SystemColors.ControlText
        Me.frmSDR.Location = New System.Drawing.Point(296, 88)
        Me.frmSDR.Name = "frmSDR"
        Me.frmSDR.Padding = New System.Windows.Forms.Padding(0)
        Me.frmSDR.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.frmSDR.Size = New System.Drawing.Size(297, 57)
        Me.frmSDR.TabIndex = 64
        Me.frmSDR.TabStop = False
        Me.frmSDR.Text = "Frame7"
        '
        'cmdOpenSDR
        '
        Me.cmdOpenSDR.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOpenSDR.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOpenSDR.Enabled = False
        Me.cmdOpenSDR.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOpenSDR.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOpenSDR.Image = CType(resources.GetObject("cmdOpenSDR.Image"), System.Drawing.Image)
        Me.cmdOpenSDR.Location = New System.Drawing.Point(256, 24)
        Me.cmdOpenSDR.Name = "cmdOpenSDR"
        Me.cmdOpenSDR.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOpenSDR.Size = New System.Drawing.Size(23, 20)
        Me.cmdOpenSDR.TabIndex = 67
        Me.cmdOpenSDR.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdOpenSDR.UseVisualStyleBackColor = False
        '
        'txtSDRGRID
        '
        Me.txtSDRGRID.AcceptsReturn = True
        Me.txtSDRGRID.BackColor = System.Drawing.SystemColors.Window
        Me.txtSDRGRID.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSDRGRID.Enabled = False
        Me.txtSDRGRID.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSDRGRID.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSDRGRID.Location = New System.Drawing.Point(16, 24)
        Me.txtSDRGRID.MaxLength = 0
        Me.txtSDRGRID.Name = "txtSDRGRID"
        Me.txtSDRGRID.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSDRGRID.Size = New System.Drawing.Size(233, 19)
        Me.txtSDRGRID.TabIndex = 66
        '
        'chkSDR
        '
        Me.chkSDR.BackColor = System.Drawing.SystemColors.Control
        Me.chkSDR.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkSDR.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkSDR.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkSDR.Location = New System.Drawing.Point(8, 0)
        Me.chkSDR.Name = "chkSDR"
        Me.chkSDR.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkSDR.Size = New System.Drawing.Size(217, 17)
        Me.chkSDR.TabIndex = 65
        Me.chkSDR.Text = "Sediment Delivery Ratio GRID (optional)"
        Me.chkSDR.UseVisualStyleBackColor = False
        '
        '_SSTab1_TabPage2
        '
        Me._SSTab1_TabPage2.Controls.Add(Me.grdLU)
        Me._SSTab1_TabPage2.Location = New System.Drawing.Point(4, 22)
        Me._SSTab1_TabPage2.Name = "_SSTab1_TabPage2"
        Me._SSTab1_TabPage2.Size = New System.Drawing.Size(610, 177)
        Me._SSTab1_TabPage2.TabIndex = 2
        Me._SSTab1_TabPage2.Text = "Land Uses"
        '
        'grdLU
        '
        Me.grdLU.DataSource = Nothing
        Me.grdLU.Location = New System.Drawing.Point(12, 32)
        Me.grdLU.Name = "grdLU"
        Me.grdLU.OcxState = CType(resources.GetObject("grdLU.OcxState"), System.Windows.Forms.AxHost.State)
        Me.grdLU.Size = New System.Drawing.Size(403, 153)
        Me.grdLU.TabIndex = 13
        '
        '_SSTab1_TabPage3
        '
        Me._SSTab1_TabPage3.Controls.Add(Me.grdLCChanges)
        Me._SSTab1_TabPage3.Location = New System.Drawing.Point(4, 22)
        Me._SSTab1_TabPage3.Name = "_SSTab1_TabPage3"
        Me._SSTab1_TabPage3.Size = New System.Drawing.Size(610, 177)
        Me._SSTab1_TabPage3.TabIndex = 3
        Me._SSTab1_TabPage3.Text = "Management Scenarios"
        '
        'grdLCChanges
        '
        Me.grdLCChanges.DataSource = Nothing
        Me.grdLCChanges.Location = New System.Drawing.Point(12, 32)
        Me.grdLCChanges.Name = "grdLCChanges"
        Me.grdLCChanges.OcxState = CType(resources.GetObject("grdLCChanges.OcxState"), System.Windows.Forms.AxHost.State)
        Me.grdLCChanges.Size = New System.Drawing.Size(524, 153)
        Me.grdLCChanges.TabIndex = 34
        '
        'cmdRun
        '
        Me.cmdRun.BackColor = System.Drawing.SystemColors.Control
        Me.cmdRun.CausesValidation = False
        Me.cmdRun.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdRun.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdRun.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdRun.Location = New System.Drawing.Point(459, 471)
        Me.cmdRun.Name = "cmdRun"
        Me.cmdRun.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdRun.Size = New System.Drawing.Size(72, 25)
        Me.cmdRun.TabIndex = 3
        Me.cmdRun.TabStop = False
        Me.cmdRun.Text = "Run"
        Me.cmdRun.UseVisualStyleBackColor = False
        '
        'cmdQuit
        '
        Me.cmdQuit.BackColor = System.Drawing.SystemColors.Control
        Me.cmdQuit.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdQuit.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdQuit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdQuit.Location = New System.Drawing.Point(537, 471)
        Me.cmdQuit.Name = "cmdQuit"
        Me.cmdQuit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdQuit.Size = New System.Drawing.Size(72, 25)
        Me.cmdQuit.TabIndex = 4
        Me.cmdQuit.TabStop = False
        Me.cmdQuit.Text = "Cancel"
        Me.cmdQuit.UseVisualStyleBackColor = False
        '
        '_chkIgnore_0
        '
        Me._chkIgnore_0.BackColor = System.Drawing.SystemColors.Window
        Me._chkIgnore_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._chkIgnore_0.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me._chkIgnore_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._chkIgnore_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.chkIgnore.SetIndex(Me._chkIgnore_0, CType(0, Short))
        Me._chkIgnore_0.Location = New System.Drawing.Point(77, 383)
        Me._chkIgnore_0.Name = "_chkIgnore_0"
        Me._chkIgnore_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._chkIgnore_0.Size = New System.Drawing.Size(15, 13)
        Me._chkIgnore_0.TabIndex = 30
        Me._chkIgnore_0.Text = "Check2"
        Me._chkIgnore_0.UseVisualStyleBackColor = False
        Me._chkIgnore_0.Visible = False
        '
        '_chkIgnoreMgmt_0
        '
        Me._chkIgnoreMgmt_0.BackColor = System.Drawing.SystemColors.Window
        Me._chkIgnoreMgmt_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._chkIgnoreMgmt_0.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me._chkIgnoreMgmt_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._chkIgnoreMgmt_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.chkIgnoreMgmt.SetIndex(Me._chkIgnoreMgmt_0, CType(0, Short))
        Me._chkIgnoreMgmt_0.Location = New System.Drawing.Point(535, 372)
        Me._chkIgnoreMgmt_0.Name = "_chkIgnoreMgmt_0"
        Me._chkIgnoreMgmt_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._chkIgnoreMgmt_0.Size = New System.Drawing.Size(15, 13)
        Me._chkIgnoreMgmt_0.TabIndex = 35
        Me._chkIgnoreMgmt_0.Text = "Check2"
        Me._chkIgnoreMgmt_0.UseVisualStyleBackColor = False
        '
        '_chkIgnoreLU_0
        '
        Me._chkIgnoreLU_0.BackColor = System.Drawing.SystemColors.Window
        Me._chkIgnoreLU_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._chkIgnoreLU_0.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me._chkIgnoreLU_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._chkIgnoreLU_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.chkIgnoreLU.SetIndex(Me._chkIgnoreLU_0, CType(0, Short))
        Me._chkIgnoreLU_0.Location = New System.Drawing.Point(467, 384)
        Me._chkIgnoreLU_0.Name = "_chkIgnoreLU_0"
        Me._chkIgnoreLU_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._chkIgnoreLU_0.Size = New System.Drawing.Size(15, 13)
        Me._chkIgnoreLU_0.TabIndex = 36
        Me._chkIgnoreLU_0.Text = "Check2"
        Me._chkIgnoreLU_0.UseVisualStyleBackColor = False
        '
        'Picture2
        '
        Me.Picture2.BackColor = System.Drawing.SystemColors.Control
        Me.Picture2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Picture2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Picture2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Picture2.Location = New System.Drawing.Point(336, 72)
        Me.Picture2.Name = "Picture2"
        Me.Picture2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Picture2.Size = New System.Drawing.Size(305, 215)
        Me.Picture2.TabIndex = 49
        Me.Picture2.Visible = False
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.txtThemeName)
        Me.Frame2.Controls.Add(Me.cmdOutputBrowse)
        Me.Frame2.Controls.Add(Me.txtOutputFile)
        Me.Frame2.Controls.Add(Me._Label1_11)
        Me.Frame2.Controls.Add(Me._Label1_12)
        Me.Frame2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame2.Location = New System.Drawing.Point(12, 407)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(614, 52)
        Me.Frame2.TabIndex = 16
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Results Output "
        Me.Frame2.Visible = False
        '
        'txtThemeName
        '
        Me.txtThemeName.AcceptsReturn = True
        Me.txtThemeName.BackColor = System.Drawing.SystemColors.Window
        Me.txtThemeName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtThemeName.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtThemeName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtThemeName.Location = New System.Drawing.Point(422, 22)
        Me.txtThemeName.MaxLength = 30
        Me.txtThemeName.Name = "txtThemeName"
        Me.txtThemeName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtThemeName.Size = New System.Drawing.Size(154, 19)
        Me.txtThemeName.TabIndex = 19
        '
        'cmdOutputBrowse
        '
        Me.cmdOutputBrowse.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOutputBrowse.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOutputBrowse.Enabled = False
        Me.cmdOutputBrowse.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOutputBrowse.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOutputBrowse.Image = CType(resources.GetObject("cmdOutputBrowse.Image"), System.Drawing.Image)
        Me.cmdOutputBrowse.Location = New System.Drawing.Point(298, 21)
        Me.cmdOutputBrowse.Name = "cmdOutputBrowse"
        Me.cmdOutputBrowse.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOutputBrowse.Size = New System.Drawing.Size(23, 20)
        Me.cmdOutputBrowse.TabIndex = 18
        Me.cmdOutputBrowse.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdOutputBrowse.UseVisualStyleBackColor = False
        '
        '_Label1_11
        '
        Me._Label1_11.BackColor = System.Drawing.Color.Transparent
        Me._Label1_11.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_11.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_11.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_11, CType(11, Short))
        Me._Label1_11.Location = New System.Drawing.Point(349, 22)
        Me._Label1_11.Name = "_Label1_11"
        Me._Label1_11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_11.Size = New System.Drawing.Size(73, 14)
        Me._Label1_11.TabIndex = 21
        Me._Label1_11.Text = "Layer Name:"
        '
        'chkIgnore
        '
        '
        'chkIgnoreLU
        '
        '
        'chkIgnoreMgmt
        '
        '
        'frmPrj
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(642, 502)
        Me.Controls.Add(Me.Frame6)
        Me.Controls.Add(Me.Frame5)
        Me.Controls.Add(Me.Frame4)
        Me.Controls.Add(Me.frm_raintype)
        Me.Controls.Add(Me.Frame3)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.fraLC)
        Me.Controls.Add(Me.cboAreaLayer)
        Me.Controls.Add(Me.cboClass)
        Me.Controls.Add(Me.cboCoeffSet)
        Me.Controls.Add(Me.cboCoeff)
        Me.Controls.Add(Me.SSTab1)
        Me.Controls.Add(Me.cmdRun)
        Me.Controls.Add(Me.cmdQuit)
        Me.Controls.Add(Me._chkIgnore_0)
        Me.Controls.Add(Me._chkIgnoreMgmt_0)
        Me.Controls.Add(Me._chkIgnoreLU_0)
        Me.Controls.Add(Me.Picture2)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.MainMenu1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(157, 105)
        Me.MaximizeBox = False
        Me.Name = "frmPrj"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        Me.Frame6.ResumeLayout(False)
        Me.Frame5.ResumeLayout(False)
        Me.Frame4.ResumeLayout(False)
        Me.frm_raintype.ResumeLayout(False)
        Me.Frame3.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.fraLC.ResumeLayout(False)
        Me.SSTab1.ResumeLayout(False)
        Me._SSTab1_TabPage0.ResumeLayout(False)
        CType(Me.grdCoeffs, System.ComponentModel.ISupportInitialize).EndInit()
        Me._SSTab1_TabPage1.ResumeLayout(False)
        Me.frameRainFall.ResumeLayout(False)
        Me.frmSDR.ResumeLayout(False)
        Me._SSTab1_TabPage2.ResumeLayout(False)
        CType(Me.grdLU, System.ComponentModel.ISupportInitialize).EndInit()
        Me._SSTab1_TabPage3.ResumeLayout(False)
        CType(Me.grdLCChanges, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Picture2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Frame2.ResumeLayout(False)
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkIgnore, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkIgnoreLU, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkIgnoreMgmt, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region 
End Class