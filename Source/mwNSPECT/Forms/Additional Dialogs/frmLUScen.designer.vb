<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmLUScen
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
	Public WithEvents mnuAppendP As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuInsertP As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuDeleteP As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopLU As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
	Public WithEvents chkSelectedPolys As System.Windows.Forms.CheckBox
    Public WithEvents chkWatWetlands As System.Windows.Forms.CheckBox
	Public WithEvents txtLUName As System.Windows.Forms.TextBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents _txtLUCN_4 As System.Windows.Forms.TextBox
	Public WithEvents _txtLUCN_3 As System.Windows.Forms.TextBox
	Public WithEvents _txtLUCN_2 As System.Windows.Forms.TextBox
	Public WithEvents _txtLUCN_1 As System.Windows.Forms.TextBox
	Public WithEvents _txtLUCN_0 As System.Windows.Forms.TextBox
	Public WithEvents cboLULayer As System.Windows.Forms.ComboBox
    Public WithEvents _Label1_5 As System.Windows.Forms.Label
	Public WithEvents _Label1_3 As System.Windows.Forms.Label
	Public WithEvents _Label1_2 As System.Windows.Forms.Label
	Public WithEvents _Label1_0 As System.Windows.Forms.Label
	Public WithEvents _Label1_15 As System.Windows.Forms.Label
	Public WithEvents _Label1_4 As System.Windows.Forms.Label
	Public WithEvents _Label1_1 As System.Windows.Forms.Label
	Public WithEvents _Label1_16 As System.Windows.Forms.Label
	Public WithEvents _Label1_17 As System.Windows.Forms.Label
	Public WithEvents _Label1_18 As System.Windows.Forms.Label
	Public WithEvents _Label1_19 As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip
        Me.mnuPopLU = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuAppendP = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuInsertP = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuDeleteP = New System.Windows.Forms.ToolStripMenuItem
        Me.chkSelectedPolys = New System.Windows.Forms.CheckBox
        Me.chkWatWetlands = New System.Windows.Forms.CheckBox
        Me.txtLUName = New System.Windows.Forms.TextBox
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmdOK = New System.Windows.Forms.Button
        Me._txtLUCN_4 = New System.Windows.Forms.TextBox
        Me._txtLUCN_3 = New System.Windows.Forms.TextBox
        Me._txtLUCN_2 = New System.Windows.Forms.TextBox
        Me._txtLUCN_1 = New System.Windows.Forms.TextBox
        Me._txtLUCN_0 = New System.Windows.Forms.TextBox
        Me.cboLULayer = New System.Windows.Forms.ComboBox
        Me._Label1_5 = New System.Windows.Forms.Label
        Me._Label1_3 = New System.Windows.Forms.Label
        Me._Label1_2 = New System.Windows.Forms.Label
        Me._Label1_0 = New System.Windows.Forms.Label
        Me._Label1_15 = New System.Windows.Forms.Label
        Me._Label1_4 = New System.Windows.Forms.Label
        Me._Label1_1 = New System.Windows.Forms.Label
        Me._Label1_16 = New System.Windows.Forms.Label
        Me._Label1_17 = New System.Windows.Forms.Label
        Me._Label1_18 = New System.Windows.Forms.Label
        Me._Label1_19 = New System.Windows.Forms.Label
        Me.dgvCoef = New System.Windows.Forms.DataGridView
        Me.Pollutant = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Type1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Type2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Type3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Type4 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.MainMenu1.SuspendLayout()
        CType(Me.dgvCoef, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuPopLU})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(410, 24)
        Me.MainMenu1.TabIndex = 25
        '
        'mnuPopLU
        '
        Me.mnuPopLU.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuAppendP, Me.mnuInsertP, Me.mnuDeleteP})
        Me.mnuPopLU.Name = "mnuPopLU"
        Me.mnuPopLU.Size = New System.Drawing.Size(39, 20)
        Me.mnuPopLU.Text = "Edit"
        Me.mnuPopLU.Visible = False
        '
        'mnuAppendP
        '
        Me.mnuAppendP.Name = "mnuAppendP"
        Me.mnuAppendP.Size = New System.Drawing.Size(142, 22)
        Me.mnuAppendP.Text = "Append Row"
        '
        'mnuInsertP
        '
        Me.mnuInsertP.Name = "mnuInsertP"
        Me.mnuInsertP.Size = New System.Drawing.Size(142, 22)
        Me.mnuInsertP.Text = "Insert Row"
        '
        'mnuDeleteP
        '
        Me.mnuDeleteP.Name = "mnuDeleteP"
        Me.mnuDeleteP.Size = New System.Drawing.Size(142, 22)
        Me.mnuDeleteP.Text = "Delete Row"
        '
        'chkSelectedPolys
        '
        Me.chkSelectedPolys.BackColor = System.Drawing.SystemColors.Control
        Me.chkSelectedPolys.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkSelectedPolys.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkSelectedPolys.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkSelectedPolys.Location = New System.Drawing.Point(137, 86)
        Me.chkSelectedPolys.Name = "chkSelectedPolys"
        Me.chkSelectedPolys.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkSelectedPolys.Size = New System.Drawing.Size(231, 16)
        Me.chkSelectedPolys.TabIndex = 24
        Me.chkSelectedPolys.Text = "Use Selected Polygons Only"
        Me.chkSelectedPolys.UseVisualStyleBackColor = False
        '
        'chkWatWetlands
        '
        Me.chkWatWetlands.BackColor = System.Drawing.SystemColors.Control
        Me.chkWatWetlands.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkWatWetlands.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkWatWetlands.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkWatWetlands.Location = New System.Drawing.Point(231, 161)
        Me.chkWatWetlands.Name = "chkWatWetlands"
        Me.chkWatWetlands.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkWatWetlands.Size = New System.Drawing.Size(109, 19)
        Me.chkWatWetlands.TabIndex = 7
        Me.chkWatWetlands.Text = "Water/Wetlands"
        Me.chkWatWetlands.UseVisualStyleBackColor = False
        '
        'txtLUName
        '
        Me.txtLUName.AcceptsReturn = True
        Me.txtLUName.BackColor = System.Drawing.SystemColors.Window
        Me.txtLUName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtLUName.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLUName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtLUName.Location = New System.Drawing.Point(137, 31)
        Me.txtLUName.MaxLength = 30
        Me.txtLUName.Name = "txtLUName"
        Me.txtLUName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtLUName.Size = New System.Drawing.Size(234, 20)
        Me.txtLUName.TabIndex = 0
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.Location = New System.Drawing.Point(312, 458)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancel.Size = New System.Drawing.Size(65, 25)
        Me.cmdCancel.TabIndex = 9
        Me.cmdCancel.TabStop = False
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'cmdOK
        '
        Me.cmdOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOK.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOK.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOK.Location = New System.Drawing.Point(240, 458)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOK.Size = New System.Drawing.Size(65, 25)
        Me.cmdOK.TabIndex = 8
        Me.cmdOK.TabStop = False
        Me.cmdOK.Text = "OK"
        Me.cmdOK.UseVisualStyleBackColor = False
        '
        '_txtLUCN_4
        '
        Me._txtLUCN_4.AcceptsReturn = True
        Me._txtLUCN_4.BackColor = System.Drawing.SystemColors.Window
        Me._txtLUCN_4.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtLUCN_4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtLUCN_4.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtLUCN_4.Location = New System.Drawing.Point(99, 162)
        Me._txtLUCN_4.MaxLength = 30
        Me._txtLUCN_4.Name = "_txtLUCN_4"
        Me._txtLUCN_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtLUCN_4.Size = New System.Drawing.Size(60, 20)
        Me._txtLUCN_4.TabIndex = 6
        Me._txtLUCN_4.Text = "0"
        '
        '_txtLUCN_3
        '
        Me._txtLUCN_3.AcceptsReturn = True
        Me._txtLUCN_3.BackColor = System.Drawing.SystemColors.Window
        Me._txtLUCN_3.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtLUCN_3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtLUCN_3.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtLUCN_3.Location = New System.Drawing.Point(316, 127)
        Me._txtLUCN_3.MaxLength = 30
        Me._txtLUCN_3.Name = "_txtLUCN_3"
        Me._txtLUCN_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtLUCN_3.Size = New System.Drawing.Size(60, 20)
        Me._txtLUCN_3.TabIndex = 5
        Me._txtLUCN_3.Text = "0"
        '
        '_txtLUCN_2
        '
        Me._txtLUCN_2.AcceptsReturn = True
        Me._txtLUCN_2.BackColor = System.Drawing.SystemColors.Window
        Me._txtLUCN_2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtLUCN_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtLUCN_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtLUCN_2.Location = New System.Drawing.Point(256, 127)
        Me._txtLUCN_2.MaxLength = 30
        Me._txtLUCN_2.Name = "_txtLUCN_2"
        Me._txtLUCN_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtLUCN_2.Size = New System.Drawing.Size(60, 20)
        Me._txtLUCN_2.TabIndex = 4
        Me._txtLUCN_2.Text = "0"
        '
        '_txtLUCN_1
        '
        Me._txtLUCN_1.AcceptsReturn = True
        Me._txtLUCN_1.BackColor = System.Drawing.SystemColors.Window
        Me._txtLUCN_1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtLUCN_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtLUCN_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtLUCN_1.Location = New System.Drawing.Point(197, 127)
        Me._txtLUCN_1.MaxLength = 30
        Me._txtLUCN_1.Name = "_txtLUCN_1"
        Me._txtLUCN_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtLUCN_1.Size = New System.Drawing.Size(60, 20)
        Me._txtLUCN_1.TabIndex = 3
        Me._txtLUCN_1.Text = "0"
        '
        '_txtLUCN_0
        '
        Me._txtLUCN_0.AcceptsReturn = True
        Me._txtLUCN_0.BackColor = System.Drawing.SystemColors.Window
        Me._txtLUCN_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtLUCN_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtLUCN_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtLUCN_0.Location = New System.Drawing.Point(137, 127)
        Me._txtLUCN_0.MaxLength = 30
        Me._txtLUCN_0.Name = "_txtLUCN_0"
        Me._txtLUCN_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtLUCN_0.Size = New System.Drawing.Size(60, 20)
        Me._txtLUCN_0.TabIndex = 2
        Me._txtLUCN_0.Text = "0"
        '
        'cboLULayer
        '
        Me.cboLULayer.BackColor = System.Drawing.SystemColors.Window
        Me.cboLULayer.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboLULayer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboLULayer.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboLULayer.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboLULayer.Location = New System.Drawing.Point(136, 59)
        Me.cboLULayer.Name = "cboLULayer"
        Me.cboLULayer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboLULayer.Size = New System.Drawing.Size(237, 22)
        Me.cboLULayer.TabIndex = 1
        '
        '_Label1_5
        '
        Me._Label1_5.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me._Label1_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_5.Location = New System.Drawing.Point(316, 111)
        Me._Label1_5.Name = "_Label1_5"
        Me._Label1_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_5.Size = New System.Drawing.Size(60, 17)
        Me._Label1_5.TabIndex = 20
        Me._Label1_5.Text = "D"
        Me._Label1_5.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_Label1_3
        '
        Me._Label1_3.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me._Label1_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_3.Location = New System.Drawing.Point(257, 111)
        Me._Label1_3.Name = "_Label1_3"
        Me._Label1_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_3.Size = New System.Drawing.Size(60, 17)
        Me._Label1_3.TabIndex = 19
        Me._Label1_3.Text = "C"
        Me._Label1_3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_Label1_2
        '
        Me._Label1_2.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me._Label1_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_2.Location = New System.Drawing.Point(197, 111)
        Me._Label1_2.Name = "_Label1_2"
        Me._Label1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_2.Size = New System.Drawing.Size(60, 17)
        Me._Label1_2.TabIndex = 18
        Me._Label1_2.Text = "B"
        Me._Label1_2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_Label1_0
        '
        Me._Label1_0.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_0.Location = New System.Drawing.Point(20, 32)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_0.Size = New System.Drawing.Size(104, 17)
        Me._Label1_0.TabIndex = 17
        Me._Label1_0.Text = "Scenario Name:"
        '
        '_Label1_15
        '
        Me._Label1_15.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_15.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_15.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_15.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_15.Location = New System.Drawing.Point(24, 164)
        Me._Label1_15.Name = "_Label1_15"
        Me._Label1_15.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_15.Size = New System.Drawing.Size(76, 19)
        Me._Label1_15.TabIndex = 16
        Me._Label1_15.Text = "Cover Factor:"
        '
        '_Label1_4
        '
        Me._Label1_4.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me._Label1_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_4.Location = New System.Drawing.Point(137, 111)
        Me._Label1_4.Name = "_Label1_4"
        Me._Label1_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_4.Size = New System.Drawing.Size(60, 17)
        Me._Label1_4.TabIndex = 15
        Me._Label1_4.Text = "A"
        Me._Label1_4.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_Label1_1
        '
        Me._Label1_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_1.Location = New System.Drawing.Point(22, 111)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_1.Size = New System.Drawing.Size(104, 17)
        Me._Label1_1.TabIndex = 14
        Me._Label1_1.Text = "SCS Curve Numbers:"
        '
        '_Label1_16
        '
        Me._Label1_16.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_16.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me._Label1_16.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_16.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_16.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_16.Location = New System.Drawing.Point(38, 194)
        Me._Label1_16.Name = "_Label1_16"
        Me._Label1_16.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_16.Size = New System.Drawing.Size(133, 17)
        Me._Label1_16.TabIndex = 13
        Me._Label1_16.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_17
        '
        Me._Label1_17.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_17.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me._Label1_17.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_17.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_17.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_17.Location = New System.Drawing.Point(171, 194)
        Me._Label1_17.Name = "_Label1_17"
        Me._Label1_17.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_17.Size = New System.Drawing.Size(217, 17)
        Me._Label1_17.TabIndex = 12
        Me._Label1_17.Text = "Coefficients"
        Me._Label1_17.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_Label1_18
        '
        Me._Label1_18.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_18.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me._Label1_18.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_18.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_18.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_18.Location = New System.Drawing.Point(11, 194)
        Me._Label1_18.Name = "_Label1_18"
        Me._Label1_18.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_18.Size = New System.Drawing.Size(27, 17)
        Me._Label1_18.TabIndex = 11
        Me._Label1_18.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_19
        '
        Me._Label1_19.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_19.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_19.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_19.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_19.Location = New System.Drawing.Point(21, 58)
        Me._Label1_19.Name = "_Label1_19"
        Me._Label1_19.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_19.Size = New System.Drawing.Size(62, 19)
        Me._Label1_19.TabIndex = 10
        Me._Label1_19.Text = "Layer:"
        '
        'dgvCoef
        '
        Me.dgvCoef.AllowUserToAddRows = False
        Me.dgvCoef.AllowUserToDeleteRows = False
        Me.dgvCoef.AllowUserToResizeColumns = False
        Me.dgvCoef.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvCoef.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvCoef.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Pollutant, Me.Type1, Me.Type2, Me.Type3, Me.Type4})
        Me.dgvCoef.Location = New System.Drawing.Point(11, 214)
        Me.dgvCoef.Name = "dgvCoef"
        Me.dgvCoef.Size = New System.Drawing.Size(377, 236)
        Me.dgvCoef.TabIndex = 26
        '
        'Pollutant
        '
        Me.Pollutant.HeaderText = "Pollutant"
        Me.Pollutant.Name = "Pollutant"
        Me.Pollutant.ReadOnly = True
        Me.Pollutant.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.Pollutant.Width = 120
        '
        'Type1
        '
        Me.Type1.HeaderText = "Type 1"
        Me.Type1.Name = "Type1"
        Me.Type1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.Type1.Width = 53
        '
        'Type2
        '
        Me.Type2.HeaderText = "Type 2"
        Me.Type2.Name = "Type2"
        Me.Type2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.Type2.Width = 53
        '
        'Type3
        '
        Me.Type3.HeaderText = "Type 3"
        Me.Type3.Name = "Type3"
        Me.Type3.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.Type3.Width = 53
        '
        'Type4
        '
        Me.Type4.HeaderText = "Type 4"
        Me.Type4.Name = "Type4"
        Me.Type4.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.Type4.Width = 53
        '
        'frmLUScen
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(410, 494)
        Me.Controls.Add(Me.dgvCoef)
        Me.Controls.Add(Me.chkSelectedPolys)
        Me.Controls.Add(Me.chkWatWetlands)
        Me.Controls.Add(Me.txtLUName)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me._txtLUCN_4)
        Me.Controls.Add(Me._txtLUCN_3)
        Me.Controls.Add(Me._txtLUCN_2)
        Me.Controls.Add(Me._txtLUCN_1)
        Me.Controls.Add(Me._txtLUCN_0)
        Me.Controls.Add(Me.cboLULayer)
        Me.Controls.Add(Me._Label1_5)
        Me.Controls.Add(Me._Label1_3)
        Me.Controls.Add(Me._Label1_2)
        Me.Controls.Add(Me._Label1_0)
        Me.Controls.Add(Me._Label1_15)
        Me.Controls.Add(Me._Label1_4)
        Me.Controls.Add(Me._Label1_1)
        Me.Controls.Add(Me._Label1_16)
        Me.Controls.Add(Me._Label1_17)
        Me.Controls.Add(Me._Label1_18)
        Me.Controls.Add(Me._Label1_19)
        Me.Controls.Add(Me.MainMenu1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(4, 23)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmLUScen"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Edit Land Use Scenario"
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        CType(Me.dgvCoef, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvCoef As System.Windows.Forms.DataGridView
    Friend WithEvents Pollutant As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Type1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Type2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Type3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Type4 As System.Windows.Forms.DataGridViewTextBoxColumn
#End Region
End Class