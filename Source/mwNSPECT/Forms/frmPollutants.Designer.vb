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
    Public WithEvents cboCoeffSet As System.Windows.Forms.ComboBox
	Public WithEvents txtCoeffSetDesc As System.Windows.Forms.TextBox
	Public WithEvents txtLCType As System.Windows.Forms.TextBox
	Public WithEvents _SSTab1_TabPage0 As System.Windows.Forms.TabPage
	Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents _SSTab1_TabPage1 As System.Windows.Forms.TabPage
	Public WithEvents SSTab1 As System.Windows.Forms.TabControl
	Public WithEvents cmdSave As System.Windows.Forms.Button
	Public WithEvents cmdQuit As System.Windows.Forms.Button
	Public WithEvents cboPollName As System.Windows.Forms.ComboBox
	Public dlgCMD1Open As System.Windows.Forms.OpenFileDialog
	Public dlgCMD1Save As System.Windows.Forms.SaveFileDialog
	Public WithEvents _Label1_0 As System.Windows.Forms.Label
    'Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    'Public WithEvents mnuPoll As Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
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
        Me.cboCoeffSet = New System.Windows.Forms.ComboBox
        Me.txtCoeffSetDesc = New System.Windows.Forms.TextBox
        Me.txtLCType = New System.Windows.Forms.TextBox
        Me._SSTab1_TabPage1 = New System.Windows.Forms.TabPage
        Me.Label2 = New System.Windows.Forms.Label
        Me.cmdSave = New System.Windows.Forms.Button
        Me.cmdQuit = New System.Windows.Forms.Button
        Me.cboPollName = New System.Windows.Forms.ComboBox
        Me.dlgCMD1Open = New System.Windows.Forms.OpenFileDialog
        Me.dlgCMD1Save = New System.Windows.Forms.SaveFileDialog
        Me._Label1_0 = New System.Windows.Forms.Label
        Me.MainMenu1.SuspendLayout()
        Me.SSTab1.SuspendLayout()
        Me._SSTab1_TabPage0.SuspendLayout()
        Me._SSTab1_TabPage1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuPoll_1, Me.mnuCoeff, Me.mnuHelp})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(561, 24)
        Me.MainMenu1.TabIndex = 16
        '
        '_mnuPoll_1
        '
        Me._mnuPoll_1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuAddPoll, Me.mnuDeletePoll})
        Me._mnuPoll_1.Name = "_mnuPoll_1"
        Me._mnuPoll_1.Size = New System.Drawing.Size(72, 20)
        Me._mnuPoll_1.Text = "&Pollutants"
        '
        'mnuAddPoll
        '
        Me.mnuAddPoll.Name = "mnuAddPoll"
        Me.mnuAddPoll.Size = New System.Drawing.Size(116, 22)
        Me.mnuAddPoll.Text = "&Add..."
        '
        'mnuDeletePoll
        '
        Me.mnuDeletePoll.Name = "mnuDeletePoll"
        Me.mnuDeletePoll.Size = New System.Drawing.Size(116, 22)
        Me.mnuDeletePoll.Text = "&Delete..."
        '
        'mnuCoeff
        '
        Me.mnuCoeff.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuCoeffNewSet, Me.mnuCoeffCopySet, Me.mnuCoeffDeleteSet, Me.mnuCoeffImportSet, Me.mnuCoeffExportSet})
        Me.mnuCoeff.Name = "mnuCoeff"
        Me.mnuCoeff.Size = New System.Drawing.Size(82, 20)
        Me.mnuCoeff.Text = "&Coefficients"
        '
        'mnuCoeffNewSet
        '
        Me.mnuCoeffNewSet.Name = "mnuCoeffNewSet"
        Me.mnuCoeffNewSet.Size = New System.Drawing.Size(138, 22)
        Me.mnuCoeffNewSet.Text = "&New Set..."
        '
        'mnuCoeffCopySet
        '
        Me.mnuCoeffCopySet.Name = "mnuCoeffCopySet"
        Me.mnuCoeffCopySet.Size = New System.Drawing.Size(138, 22)
        Me.mnuCoeffCopySet.Text = "&Copy Set..."
        '
        'mnuCoeffDeleteSet
        '
        Me.mnuCoeffDeleteSet.Name = "mnuCoeffDeleteSet"
        Me.mnuCoeffDeleteSet.Size = New System.Drawing.Size(138, 22)
        Me.mnuCoeffDeleteSet.Text = "&Delete Set..."
        '
        'mnuCoeffImportSet
        '
        Me.mnuCoeffImportSet.Name = "mnuCoeffImportSet"
        Me.mnuCoeffImportSet.Size = New System.Drawing.Size(138, 22)
        Me.mnuCoeffImportSet.Text = "&Import Set..."
        '
        'mnuCoeffExportSet
        '
        Me.mnuCoeffExportSet.Name = "mnuCoeffExportSet"
        Me.mnuCoeffExportSet.Size = New System.Drawing.Size(138, 22)
        Me.mnuCoeffExportSet.Text = "&Export Set..."
        '
        'mnuHelp
        '
        Me.mnuHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuPollHelp, Me.mnuCoeffHelp})
        Me.mnuHelp.Name = "mnuHelp"
        Me.mnuHelp.Size = New System.Drawing.Size(44, 20)
        Me.mnuHelp.Text = "&Help"
        '
        'mnuPollHelp
        '
        Me.mnuPollHelp.Name = "mnuPollHelp"
        Me.mnuPollHelp.ShortcutKeys = CType((System.Windows.Forms.Keys.Shift Or System.Windows.Forms.Keys.F1), System.Windows.Forms.Keys)
        Me.mnuPollHelp.Size = New System.Drawing.Size(197, 22)
        Me.mnuPollHelp.Text = "Pollutants..."
        '
        'mnuCoeffHelp
        '
        Me.mnuCoeffHelp.Name = "mnuCoeffHelp"
        Me.mnuCoeffHelp.ShortcutKeys = CType((System.Windows.Forms.Keys.Shift Or System.Windows.Forms.Keys.F2), System.Windows.Forms.Keys)
        Me.mnuCoeffHelp.Size = New System.Drawing.Size(197, 22)
        Me.mnuCoeffHelp.Text = "Coefficients..."
        '
        'txtActiveCellWQStd
        '
        Me.txtActiveCellWQStd.AcceptsReturn = True
        Me.txtActiveCellWQStd.BackColor = System.Drawing.Color.White
        Me.txtActiveCellWQStd.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtActiveCellWQStd.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtActiveCellWQStd.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtActiveCellWQStd.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtActiveCellWQStd.Location = New System.Drawing.Point(209, 555)
        Me.txtActiveCellWQStd.MaxLength = 0
        Me.txtActiveCellWQStd.Name = "txtActiveCellWQStd"
        Me.txtActiveCellWQStd.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtActiveCellWQStd.Size = New System.Drawing.Size(70, 16)
        Me.txtActiveCellWQStd.TabIndex = 15
        Me.txtActiveCellWQStd.Text = "Text1"
        Me.txtActiveCellWQStd.Visible = False
        '
        'txtActiveCell
        '
        Me.txtActiveCell.AcceptsReturn = True
        Me.txtActiveCell.BackColor = System.Drawing.Color.White
        Me.txtActiveCell.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtActiveCell.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtActiveCell.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtActiveCell.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtActiveCell.Location = New System.Drawing.Point(120, 555)
        Me.txtActiveCell.MaxLength = 0
        Me.txtActiveCell.Name = "txtActiveCell"
        Me.txtActiveCell.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtActiveCell.Size = New System.Drawing.Size(70, 16)
        Me.txtActiveCell.TabIndex = 14
        Me.txtActiveCell.Text = "Text1"
        Me.txtActiveCell.Visible = False
        '
        'SSTab1
        '
        Me.SSTab1.Appearance = System.Windows.Forms.TabAppearance.FlatButtons
        Me.SSTab1.Controls.Add(Me._SSTab1_TabPage0)
        Me.SSTab1.Controls.Add(Me._SSTab1_TabPage1)
        Me.SSTab1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SSTab1.ItemSize = New System.Drawing.Size(42, 18)
        Me.SSTab1.Location = New System.Drawing.Point(13, 62)
        Me.SSTab1.Name = "SSTab1"
        Me.SSTab1.SelectedIndex = 0
        Me.SSTab1.Size = New System.Drawing.Size(533, 483)
        Me.SSTab1.TabIndex = 4
        '
        '_SSTab1_TabPage0
        '
        Me._SSTab1_TabPage0.Controls.Add(Me._Label1_5)
        Me._SSTab1_TabPage0.Controls.Add(Me._Label1_6)
        Me._SSTab1_TabPage0.Controls.Add(Me._Label1_7)
        Me._SSTab1_TabPage0.Controls.Add(Me._Label1_1)
        Me._SSTab1_TabPage0.Controls.Add(Me._Label1_2)
        Me._SSTab1_TabPage0.Controls.Add(Me._Label1_3)
        Me._SSTab1_TabPage0.Controls.Add(Me.cboCoeffSet)
        Me._SSTab1_TabPage0.Controls.Add(Me.txtCoeffSetDesc)
        Me._SSTab1_TabPage0.Controls.Add(Me.txtLCType)
        Me._SSTab1_TabPage0.Location = New System.Drawing.Point(4, 22)
        Me._SSTab1_TabPage0.Name = "_SSTab1_TabPage0"
        Me._SSTab1_TabPage0.Size = New System.Drawing.Size(525, 457)
        Me._SSTab1_TabPage0.TabIndex = 0
        Me._SSTab1_TabPage0.Text = "Coefficients"
        '
        '_Label1_5
        '
        Me._Label1_5.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_5.Location = New System.Drawing.Point(17, 34)
        Me._Label1_5.Name = "_Label1_5"
        Me._Label1_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_5.Size = New System.Drawing.Size(97, 17)
        Me._Label1_5.TabIndex = 8
        Me._Label1_5.Text = "Coefficient Set:"
        '
        '_Label1_6
        '
        Me._Label1_6.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_6.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_6.Location = New System.Drawing.Point(18, 59)
        Me._Label1_6.Name = "_Label1_6"
        Me._Label1_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_6.Size = New System.Drawing.Size(74, 17)
        Me._Label1_6.TabIndex = 9
        Me._Label1_6.Text = "Description:"
        '
        '_Label1_7
        '
        Me._Label1_7.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_7.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_7.Location = New System.Drawing.Point(271, 34)
        Me._Label1_7.Name = "_Label1_7"
        Me._Label1_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_7.Size = New System.Drawing.Size(97, 17)
        Me._Label1_7.TabIndex = 10
        Me._Label1_7.Text = "Land Cover Type:"
        '
        '_Label1_1
        '
        Me._Label1_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_1.ForeColor = System.Drawing.Color.Black
        Me._Label1_1.Location = New System.Drawing.Point(44, 86)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_1.Size = New System.Drawing.Size(230, 17)
        Me._Label1_1.TabIndex = 11
        Me._Label1_1.Text = "Class"
        Me._Label1_1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_Label1_2
        '
        Me._Label1_2.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me._Label1_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_2.ForeColor = System.Drawing.Color.Black
        Me._Label1_2.Location = New System.Drawing.Point(276, 86)
        Me._Label1_2.Name = "_Label1_2"
        Me._Label1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_2.Size = New System.Drawing.Size(224, 17)
        Me._Label1_2.TabIndex = 12
        Me._Label1_2.Text = "Coefficients (mg/L)"
        Me._Label1_2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_Label1_3
        '
        Me._Label1_3.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me._Label1_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_3.Location = New System.Drawing.Point(14, 86)
        Me._Label1_3.Name = "_Label1_3"
        Me._Label1_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_3.Size = New System.Drawing.Size(27, 17)
        Me._Label1_3.TabIndex = 13
        Me._Label1_3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'cboCoeffSet
        '
        Me.cboCoeffSet.BackColor = System.Drawing.SystemColors.Window
        Me.cboCoeffSet.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboCoeffSet.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCoeffSet.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboCoeffSet.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboCoeffSet.Location = New System.Drawing.Point(100, 32)
        Me.cboCoeffSet.Name = "cboCoeffSet"
        Me.cboCoeffSet.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboCoeffSet.Size = New System.Drawing.Size(147, 22)
        Me.cboCoeffSet.Sorted = True
        Me.cboCoeffSet.TabIndex = 5
        '
        'txtCoeffSetDesc
        '
        Me.txtCoeffSetDesc.AcceptsReturn = True
        Me.txtCoeffSetDesc.BackColor = System.Drawing.SystemColors.Window
        Me.txtCoeffSetDesc.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCoeffSetDesc.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCoeffSetDesc.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCoeffSetDesc.Location = New System.Drawing.Point(100, 60)
        Me.txtCoeffSetDesc.MaxLength = 0
        Me.txtCoeffSetDesc.Name = "txtCoeffSetDesc"
        Me.txtCoeffSetDesc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCoeffSetDesc.Size = New System.Drawing.Size(375, 19)
        Me.txtCoeffSetDesc.TabIndex = 6
        '
        'txtLCType
        '
        Me.txtLCType.AcceptsReturn = True
        Me.txtLCType.BackColor = System.Drawing.SystemColors.Window
        Me.txtLCType.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtLCType.Enabled = False
        Me.txtLCType.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLCType.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtLCType.Location = New System.Drawing.Point(365, 32)
        Me.txtLCType.MaxLength = 0
        Me.txtLCType.Name = "txtLCType"
        Me.txtLCType.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtLCType.Size = New System.Drawing.Size(111, 19)
        Me.txtLCType.TabIndex = 7
        '
        '_SSTab1_TabPage1
        '
        Me._SSTab1_TabPage1.Controls.Add(Me.Label2)
        Me._SSTab1_TabPage1.Location = New System.Drawing.Point(4, 22)
        Me._SSTab1_TabPage1.Name = "_SSTab1_TabPage1"
        Me._SSTab1_TabPage1.Size = New System.Drawing.Size(525, 457)
        Me._SSTab1_TabPage1.TabIndex = 1
        Me._SSTab1_TabPage1.Text = "Water Quality Standards"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(32, 344)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(137, 17)
        Me.Label2.TabIndex = 18
        Me.Label2.Text = "Threshold Units: ug/L"
        '
        'cmdSave
        '
        Me.cmdSave.BackColor = System.Drawing.SystemColors.Control
        Me.cmdSave.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdSave.Enabled = False
        Me.cmdSave.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSave.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSave.Location = New System.Drawing.Point(371, 553)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdSave.Size = New System.Drawing.Size(65, 25)
        Me.cmdSave.TabIndex = 2
        Me.cmdSave.Text = "OK"
        Me.cmdSave.UseVisualStyleBackColor = False
        '
        'cmdQuit
        '
        Me.cmdQuit.BackColor = System.Drawing.SystemColors.Control
        Me.cmdQuit.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdQuit.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdQuit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdQuit.Location = New System.Drawing.Point(442, 553)
        Me.cmdQuit.Name = "cmdQuit"
        Me.cmdQuit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdQuit.Size = New System.Drawing.Size(65, 25)
        Me.cmdQuit.TabIndex = 3
        Me.cmdQuit.Text = "Cancel"
        Me.cmdQuit.UseVisualStyleBackColor = False
        '
        'cboPollName
        '
        Me.cboPollName.BackColor = System.Drawing.SystemColors.Window
        Me.cboPollName.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboPollName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPollName.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboPollName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboPollName.Location = New System.Drawing.Point(106, 34)
        Me.cboPollName.Name = "cboPollName"
        Me.cboPollName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboPollName.Size = New System.Drawing.Size(150, 22)
        Me.cboPollName.TabIndex = 1
        '
        '_Label1_0
        '
        Me._Label1_0.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_0.Location = New System.Drawing.Point(21, 35)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_0.Size = New System.Drawing.Size(89, 17)
        Me._Label1_0.TabIndex = 0
        Me._Label1_0.Text = "Pollutant Name:"
        '
        'frmPollutants
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(561, 590)
        Me.Controls.Add(Me.txtActiveCellWQStd)
        Me.Controls.Add(Me.txtActiveCell)
        Me.Controls.Add(Me.SSTab1)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.cmdQuit)
        Me.Controls.Add(Me.cboPollName)
        Me.Controls.Add(Me._Label1_0)
        Me.Controls.Add(Me.MainMenu1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(113, 127)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmPollutants"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Pollutants"
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        Me.SSTab1.ResumeLayout(False)
        Me._SSTab1_TabPage0.ResumeLayout(False)
        Me._SSTab1_TabPage1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region 
End Class