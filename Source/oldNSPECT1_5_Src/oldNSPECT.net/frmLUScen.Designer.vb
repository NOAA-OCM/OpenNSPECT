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
	Public WithEvents txtTypes As System.Windows.Forms.TextBox
	Public WithEvents cboPollName As System.Windows.Forms.ComboBox
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
	Public WithEvents grdPoll As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
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
	Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents txtLUCN As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmLUScen))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.MainMenu1 = New System.Windows.Forms.MenuStrip
		Me.mnuPopLU = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuAppendP = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuInsertP = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuDeleteP = New System.Windows.Forms.ToolStripMenuItem
		Me.chkSelectedPolys = New System.Windows.Forms.CheckBox
		Me.txtTypes = New System.Windows.Forms.TextBox
		Me.cboPollName = New System.Windows.Forms.ComboBox
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
		Me.grdPoll = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
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
		Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.txtLUCN = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(components)
		Me.MainMenu1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.grdPoll, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.txtLUCN, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.Text = "Edit Land Use Scenario"
		Me.ClientSize = New System.Drawing.Size(410, 389)
		Me.Location = New System.Drawing.Point(4, 23)
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
		Me.Name = "frmLUScen"
		Me.mnuPopLU.Name = "mnuPopLU"
		Me.mnuPopLU.Text = "Edit"
		Me.mnuPopLU.Visible = False
		Me.mnuPopLU.Checked = False
		Me.mnuPopLU.Enabled = True
		Me.mnuAppendP.Name = "mnuAppendP"
		Me.mnuAppendP.Text = "Append Row"
		Me.mnuAppendP.Checked = False
		Me.mnuAppendP.Enabled = True
		Me.mnuAppendP.Visible = True
		Me.mnuInsertP.Name = "mnuInsertP"
		Me.mnuInsertP.Text = "Insert Row"
		Me.mnuInsertP.Checked = False
		Me.mnuInsertP.Enabled = True
		Me.mnuInsertP.Visible = True
		Me.mnuDeleteP.Name = "mnuDeleteP"
		Me.mnuDeleteP.Text = "Delete Row"
		Me.mnuDeleteP.Checked = False
		Me.mnuDeleteP.Enabled = True
		Me.mnuDeleteP.Visible = True
		Me.chkSelectedPolys.Text = "Use Selected Polygons Only"
		Me.chkSelectedPolys.Size = New System.Drawing.Size(231, 16)
		Me.chkSelectedPolys.Location = New System.Drawing.Point(137, 86)
		Me.chkSelectedPolys.TabIndex = 24
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
		Me.txtTypes.AutoSize = False
		Me.txtTypes.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
		Me.txtTypes.Size = New System.Drawing.Size(61, 19)
		Me.txtTypes.Location = New System.Drawing.Point(151, 352)
		Me.txtTypes.TabIndex = 22
		Me.txtTypes.Text = "Text1"
		Me.txtTypes.Visible = False
		Me.txtTypes.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtTypes.AcceptsReturn = True
		Me.txtTypes.BackColor = System.Drawing.SystemColors.Window
		Me.txtTypes.CausesValidation = True
		Me.txtTypes.Enabled = True
		Me.txtTypes.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtTypes.HideSelection = True
		Me.txtTypes.ReadOnly = False
		Me.txtTypes.Maxlength = 0
		Me.txtTypes.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtTypes.MultiLine = False
		Me.txtTypes.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtTypes.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtTypes.TabStop = True
		Me.txtTypes.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.txtTypes.Name = "txtTypes"
		Me.cboPollName.Size = New System.Drawing.Size(112, 21)
		Me.cboPollName.Location = New System.Drawing.Point(17, 351)
		Me.cboPollName.TabIndex = 21
		Me.cboPollName.Text = "Combo1"
		Me.cboPollName.Visible = False
		Me.cboPollName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboPollName.BackColor = System.Drawing.SystemColors.Window
		Me.cboPollName.CausesValidation = True
		Me.cboPollName.Enabled = True
		Me.cboPollName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboPollName.IntegralHeight = True
		Me.cboPollName.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboPollName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboPollName.Sorted = False
		Me.cboPollName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboPollName.TabStop = True
		Me.cboPollName.Name = "cboPollName"
		Me.chkWatWetlands.Text = "Water/Wetlands"
		Me.chkWatWetlands.Size = New System.Drawing.Size(109, 19)
		Me.chkWatWetlands.Location = New System.Drawing.Point(231, 161)
		Me.chkWatWetlands.TabIndex = 7
		Me.chkWatWetlands.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkWatWetlands.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkWatWetlands.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkWatWetlands.BackColor = System.Drawing.SystemColors.Control
		Me.chkWatWetlands.CausesValidation = True
		Me.chkWatWetlands.Enabled = True
		Me.chkWatWetlands.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkWatWetlands.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkWatWetlands.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkWatWetlands.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkWatWetlands.TabStop = True
		Me.chkWatWetlands.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkWatWetlands.Visible = True
		Me.chkWatWetlands.Name = "chkWatWetlands"
		Me.txtLUName.AutoSize = False
		Me.txtLUName.Size = New System.Drawing.Size(234, 19)
		Me.txtLUName.Location = New System.Drawing.Point(137, 31)
		Me.txtLUName.Maxlength = 30
		Me.txtLUName.TabIndex = 0
		Me.txtLUName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtLUName.AcceptsReturn = True
		Me.txtLUName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtLUName.BackColor = System.Drawing.SystemColors.Window
		Me.txtLUName.CausesValidation = True
		Me.txtLUName.Enabled = True
		Me.txtLUName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtLUName.HideSelection = True
		Me.txtLUName.ReadOnly = False
		Me.txtLUName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtLUName.MultiLine = False
		Me.txtLUName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtLUName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtLUName.TabStop = True
		Me.txtLUName.Visible = True
		Me.txtLUName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtLUName.Name = "txtLUName"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(65, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(312, 353)
		Me.cmdCancel.TabIndex = 9
		Me.cmdCancel.TabStop = False
		Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCancel.CausesValidation = True
		Me.cmdCancel.Enabled = True
		Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.Name = "cmdCancel"
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOK.Text = "OK"
		Me.cmdOK.Size = New System.Drawing.Size(65, 25)
		Me.cmdOK.Location = New System.Drawing.Point(240, 353)
		Me.cmdOK.TabIndex = 8
		Me.cmdOK.TabStop = False
		Me.cmdOK.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOK.CausesValidation = True
		Me.cmdOK.Enabled = True
		Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOK.Name = "cmdOK"
		Me._txtLUCN_4.AutoSize = False
		Me._txtLUCN_4.Size = New System.Drawing.Size(60, 19)
		Me._txtLUCN_4.Location = New System.Drawing.Point(99, 162)
		Me._txtLUCN_4.Maxlength = 30
		Me._txtLUCN_4.TabIndex = 6
		Me._txtLUCN_4.Text = "0"
		Me._txtLUCN_4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txtLUCN_4.AcceptsReturn = True
		Me._txtLUCN_4.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtLUCN_4.BackColor = System.Drawing.SystemColors.Window
		Me._txtLUCN_4.CausesValidation = True
		Me._txtLUCN_4.Enabled = True
		Me._txtLUCN_4.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtLUCN_4.HideSelection = True
		Me._txtLUCN_4.ReadOnly = False
		Me._txtLUCN_4.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtLUCN_4.MultiLine = False
		Me._txtLUCN_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtLUCN_4.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtLUCN_4.TabStop = True
		Me._txtLUCN_4.Visible = True
		Me._txtLUCN_4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtLUCN_4.Name = "_txtLUCN_4"
		Me._txtLUCN_3.AutoSize = False
		Me._txtLUCN_3.Size = New System.Drawing.Size(60, 19)
		Me._txtLUCN_3.Location = New System.Drawing.Point(316, 127)
		Me._txtLUCN_3.Maxlength = 30
		Me._txtLUCN_3.TabIndex = 5
		Me._txtLUCN_3.Text = "0"
		Me._txtLUCN_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txtLUCN_3.AcceptsReturn = True
		Me._txtLUCN_3.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtLUCN_3.BackColor = System.Drawing.SystemColors.Window
		Me._txtLUCN_3.CausesValidation = True
		Me._txtLUCN_3.Enabled = True
		Me._txtLUCN_3.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtLUCN_3.HideSelection = True
		Me._txtLUCN_3.ReadOnly = False
		Me._txtLUCN_3.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtLUCN_3.MultiLine = False
		Me._txtLUCN_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtLUCN_3.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtLUCN_3.TabStop = True
		Me._txtLUCN_3.Visible = True
		Me._txtLUCN_3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtLUCN_3.Name = "_txtLUCN_3"
		Me._txtLUCN_2.AutoSize = False
		Me._txtLUCN_2.Size = New System.Drawing.Size(60, 19)
		Me._txtLUCN_2.Location = New System.Drawing.Point(256, 127)
		Me._txtLUCN_2.Maxlength = 30
		Me._txtLUCN_2.TabIndex = 4
		Me._txtLUCN_2.Text = "0"
		Me._txtLUCN_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txtLUCN_2.AcceptsReturn = True
		Me._txtLUCN_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtLUCN_2.BackColor = System.Drawing.SystemColors.Window
		Me._txtLUCN_2.CausesValidation = True
		Me._txtLUCN_2.Enabled = True
		Me._txtLUCN_2.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtLUCN_2.HideSelection = True
		Me._txtLUCN_2.ReadOnly = False
		Me._txtLUCN_2.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtLUCN_2.MultiLine = False
		Me._txtLUCN_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtLUCN_2.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtLUCN_2.TabStop = True
		Me._txtLUCN_2.Visible = True
		Me._txtLUCN_2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtLUCN_2.Name = "_txtLUCN_2"
		Me._txtLUCN_1.AutoSize = False
		Me._txtLUCN_1.Size = New System.Drawing.Size(60, 19)
		Me._txtLUCN_1.Location = New System.Drawing.Point(197, 127)
		Me._txtLUCN_1.Maxlength = 30
		Me._txtLUCN_1.TabIndex = 3
		Me._txtLUCN_1.Text = "0"
		Me._txtLUCN_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txtLUCN_1.AcceptsReturn = True
		Me._txtLUCN_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtLUCN_1.BackColor = System.Drawing.SystemColors.Window
		Me._txtLUCN_1.CausesValidation = True
		Me._txtLUCN_1.Enabled = True
		Me._txtLUCN_1.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtLUCN_1.HideSelection = True
		Me._txtLUCN_1.ReadOnly = False
		Me._txtLUCN_1.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtLUCN_1.MultiLine = False
		Me._txtLUCN_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtLUCN_1.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtLUCN_1.TabStop = True
		Me._txtLUCN_1.Visible = True
		Me._txtLUCN_1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtLUCN_1.Name = "_txtLUCN_1"
		Me._txtLUCN_0.AutoSize = False
		Me._txtLUCN_0.Size = New System.Drawing.Size(60, 19)
		Me._txtLUCN_0.Location = New System.Drawing.Point(137, 127)
		Me._txtLUCN_0.Maxlength = 30
		Me._txtLUCN_0.TabIndex = 2
		Me._txtLUCN_0.Text = "0"
		Me._txtLUCN_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txtLUCN_0.AcceptsReturn = True
		Me._txtLUCN_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtLUCN_0.BackColor = System.Drawing.SystemColors.Window
		Me._txtLUCN_0.CausesValidation = True
		Me._txtLUCN_0.Enabled = True
		Me._txtLUCN_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtLUCN_0.HideSelection = True
		Me._txtLUCN_0.ReadOnly = False
		Me._txtLUCN_0.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtLUCN_0.MultiLine = False
		Me._txtLUCN_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtLUCN_0.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtLUCN_0.TabStop = True
		Me._txtLUCN_0.Visible = True
		Me._txtLUCN_0.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtLUCN_0.Name = "_txtLUCN_0"
		Me.cboLULayer.Size = New System.Drawing.Size(237, 21)
		Me.cboLULayer.Location = New System.Drawing.Point(136, 59)
		Me.cboLULayer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cboLULayer.TabIndex = 1
		Me.cboLULayer.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboLULayer.BackColor = System.Drawing.SystemColors.Window
		Me.cboLULayer.CausesValidation = True
		Me.cboLULayer.Enabled = True
		Me.cboLULayer.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboLULayer.IntegralHeight = True
		Me.cboLULayer.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboLULayer.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboLULayer.Sorted = False
		Me.cboLULayer.TabStop = True
		Me.cboLULayer.Visible = True
		Me.cboLULayer.Name = "cboLULayer"
		grdPoll.OcxState = CType(resources.GetObject("grdPoll.OcxState"), System.Windows.Forms.AxHost.State)
		Me.grdPoll.Size = New System.Drawing.Size(378, 124)
		Me.grdPoll.Location = New System.Drawing.Point(11, 216)
		Me.grdPoll.TabIndex = 23
		Me.grdPoll.Name = "grdPoll"
		Me._Label1_5.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._Label1_5.Text = "D"
		Me._Label1_5.Size = New System.Drawing.Size(60, 17)
		Me._Label1_5.Location = New System.Drawing.Point(316, 111)
		Me._Label1_5.TabIndex = 20
		Me._Label1_5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_5.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_5.Enabled = True
		Me._Label1_5.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_5.UseMnemonic = True
		Me._Label1_5.Visible = True
		Me._Label1_5.AutoSize = False
		Me._Label1_5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._Label1_5.Name = "_Label1_5"
		Me._Label1_3.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._Label1_3.Text = "C"
		Me._Label1_3.Size = New System.Drawing.Size(60, 17)
		Me._Label1_3.Location = New System.Drawing.Point(257, 111)
		Me._Label1_3.TabIndex = 19
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
		Me._Label1_2.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._Label1_2.Text = "B"
		Me._Label1_2.Size = New System.Drawing.Size(60, 17)
		Me._Label1_2.Location = New System.Drawing.Point(197, 111)
		Me._Label1_2.TabIndex = 18
		Me._Label1_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_2.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_2.Enabled = True
		Me._Label1_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_2.UseMnemonic = True
		Me._Label1_2.Visible = True
		Me._Label1_2.AutoSize = False
		Me._Label1_2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._Label1_2.Name = "_Label1_2"
		Me._Label1_0.Text = "Scenario Name:"
		Me._Label1_0.Size = New System.Drawing.Size(104, 17)
		Me._Label1_0.Location = New System.Drawing.Point(20, 32)
		Me._Label1_0.TabIndex = 17
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
		Me._Label1_15.Text = "Cover Factor:"
		Me._Label1_15.Size = New System.Drawing.Size(76, 19)
		Me._Label1_15.Location = New System.Drawing.Point(24, 164)
		Me._Label1_15.TabIndex = 16
		Me._Label1_15.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_15.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_15.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_15.Enabled = True
		Me._Label1_15.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_15.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_15.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_15.UseMnemonic = True
		Me._Label1_15.Visible = True
		Me._Label1_15.AutoSize = False
		Me._Label1_15.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_15.Name = "_Label1_15"
		Me._Label1_4.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._Label1_4.Text = "A"
		Me._Label1_4.Size = New System.Drawing.Size(60, 17)
		Me._Label1_4.Location = New System.Drawing.Point(137, 111)
		Me._Label1_4.TabIndex = 15
		Me._Label1_4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_4.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_4.Enabled = True
		Me._Label1_4.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_4.UseMnemonic = True
		Me._Label1_4.Visible = True
		Me._Label1_4.AutoSize = False
		Me._Label1_4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._Label1_4.Name = "_Label1_4"
		Me._Label1_1.Text = "SCS Curve Numbers:"
		Me._Label1_1.Size = New System.Drawing.Size(104, 17)
		Me._Label1_1.Location = New System.Drawing.Point(22, 111)
		Me._Label1_1.TabIndex = 14
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
		Me._Label1_16.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_16.Size = New System.Drawing.Size(133, 17)
		Me._Label1_16.Location = New System.Drawing.Point(38, 194)
		Me._Label1_16.TabIndex = 13
		Me._Label1_16.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_16.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_16.Enabled = True
		Me._Label1_16.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_16.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_16.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_16.UseMnemonic = True
		Me._Label1_16.Visible = True
		Me._Label1_16.AutoSize = False
		Me._Label1_16.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._Label1_16.Name = "_Label1_16"
		Me._Label1_17.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._Label1_17.Text = "Coefficients"
		Me._Label1_17.Size = New System.Drawing.Size(217, 17)
		Me._Label1_17.Location = New System.Drawing.Point(171, 194)
		Me._Label1_17.TabIndex = 12
		Me._Label1_17.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_17.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_17.Enabled = True
		Me._Label1_17.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_17.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_17.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_17.UseMnemonic = True
		Me._Label1_17.Visible = True
		Me._Label1_17.AutoSize = False
		Me._Label1_17.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._Label1_17.Name = "_Label1_17"
		Me._Label1_18.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_18.Size = New System.Drawing.Size(27, 17)
		Me._Label1_18.Location = New System.Drawing.Point(11, 194)
		Me._Label1_18.TabIndex = 11
		Me._Label1_18.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_18.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_18.Enabled = True
		Me._Label1_18.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_18.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_18.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_18.UseMnemonic = True
		Me._Label1_18.Visible = True
		Me._Label1_18.AutoSize = False
		Me._Label1_18.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._Label1_18.Name = "_Label1_18"
		Me._Label1_19.Text = "Layer:"
		Me._Label1_19.Size = New System.Drawing.Size(33, 19)
		Me._Label1_19.Location = New System.Drawing.Point(21, 58)
		Me._Label1_19.TabIndex = 10
		Me._Label1_19.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_19.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_19.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_19.Enabled = True
		Me._Label1_19.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_19.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_19.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_19.UseMnemonic = True
		Me._Label1_19.Visible = True
		Me._Label1_19.AutoSize = False
		Me._Label1_19.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_19.Name = "_Label1_19"
		Me.Controls.Add(chkSelectedPolys)
		Me.Controls.Add(txtTypes)
		Me.Controls.Add(cboPollName)
		Me.Controls.Add(chkWatWetlands)
		Me.Controls.Add(txtLUName)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(_txtLUCN_4)
		Me.Controls.Add(_txtLUCN_3)
		Me.Controls.Add(_txtLUCN_2)
		Me.Controls.Add(_txtLUCN_1)
		Me.Controls.Add(_txtLUCN_0)
		Me.Controls.Add(cboLULayer)
		Me.Controls.Add(grdPoll)
		Me.Controls.Add(_Label1_5)
		Me.Controls.Add(_Label1_3)
		Me.Controls.Add(_Label1_2)
		Me.Controls.Add(_Label1_0)
		Me.Controls.Add(_Label1_15)
		Me.Controls.Add(_Label1_4)
		Me.Controls.Add(_Label1_1)
		Me.Controls.Add(_Label1_16)
		Me.Controls.Add(_Label1_17)
		Me.Controls.Add(_Label1_18)
		Me.Controls.Add(_Label1_19)
		Me.Label1.SetIndex(_Label1_5, CType(5, Short))
		Me.Label1.SetIndex(_Label1_3, CType(3, Short))
		Me.Label1.SetIndex(_Label1_2, CType(2, Short))
		Me.Label1.SetIndex(_Label1_0, CType(0, Short))
		Me.Label1.SetIndex(_Label1_15, CType(15, Short))
		Me.Label1.SetIndex(_Label1_4, CType(4, Short))
		Me.Label1.SetIndex(_Label1_1, CType(1, Short))
		Me.Label1.SetIndex(_Label1_16, CType(16, Short))
		Me.Label1.SetIndex(_Label1_17, CType(17, Short))
		Me.Label1.SetIndex(_Label1_18, CType(18, Short))
		Me.Label1.SetIndex(_Label1_19, CType(19, Short))
		Me.txtLUCN.SetIndex(_txtLUCN_4, CType(4, Short))
		Me.txtLUCN.SetIndex(_txtLUCN_3, CType(3, Short))
		Me.txtLUCN.SetIndex(_txtLUCN_2, CType(2, Short))
		Me.txtLUCN.SetIndex(_txtLUCN_1, CType(1, Short))
		Me.txtLUCN.SetIndex(_txtLUCN_0, CType(0, Short))
		CType(Me.txtLUCN, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.grdPoll, System.ComponentModel.ISupportInitialize).EndInit()
		MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuPopLU})
		mnuPopLU.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuAppendP, Me.mnuInsertP, Me.mnuDeleteP})
		Me.Controls.Add(MainMenu1)
		Me.MainMenu1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class