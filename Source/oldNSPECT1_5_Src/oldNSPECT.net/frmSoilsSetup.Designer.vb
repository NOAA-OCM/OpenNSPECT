<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmSoilsSetup
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
	Public WithEvents txtMUSLEVal As System.Windows.Forms.TextBox
	Public WithEvents txtMUSLEExp As System.Windows.Forms.TextBox
	Public WithEvents Label12 As System.Windows.Forms.Label
	Public WithEvents Label7 As System.Windows.Forms.Label
	Public WithEvents Label6 As System.Windows.Forms.Label
	Public WithEvents Label11 As System.Windows.Forms.Label
	Public WithEvents Label10 As System.Windows.Forms.Label
	Public WithEvents Label9 As System.Windows.Forms.Label
	Public WithEvents Label8 As System.Windows.Forms.Label
	Public WithEvents Frame2 As System.Windows.Forms.GroupBox
	Public WithEvents txtSoilsName As System.Windows.Forms.TextBox
	Public WithEvents cmdDEMBrowse As System.Windows.Forms.Button
	Public WithEvents txtDEMFile As System.Windows.Forms.TextBox
	Public WithEvents cmdQuit As System.Windows.Forms.Button
	Public WithEvents cmdSave As System.Windows.Forms.Button
	Public WithEvents cboSoilFieldsK As System.Windows.Forms.ComboBox
	Public WithEvents cboSoilFields As System.Windows.Forms.ComboBox
	Public WithEvents cmdBrowseFile As System.Windows.Forms.Button
	Public WithEvents txtSoilsDS As System.Windows.Forms.TextBox
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public dlgOpenOpen As System.Windows.Forms.OpenFileDialog
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label5 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmSoilsSetup))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Frame2 = New System.Windows.Forms.GroupBox
		Me.txtMUSLEVal = New System.Windows.Forms.TextBox
		Me.txtMUSLEExp = New System.Windows.Forms.TextBox
		Me.Label12 = New System.Windows.Forms.Label
		Me.Label7 = New System.Windows.Forms.Label
		Me.Label6 = New System.Windows.Forms.Label
		Me.Label11 = New System.Windows.Forms.Label
		Me.Label10 = New System.Windows.Forms.Label
		Me.Label9 = New System.Windows.Forms.Label
		Me.Label8 = New System.Windows.Forms.Label
		Me.txtSoilsName = New System.Windows.Forms.TextBox
		Me.cmdDEMBrowse = New System.Windows.Forms.Button
		Me.txtDEMFile = New System.Windows.Forms.TextBox
		Me.cmdQuit = New System.Windows.Forms.Button
		Me.cmdSave = New System.Windows.Forms.Button
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me.cboSoilFieldsK = New System.Windows.Forms.ComboBox
		Me.cboSoilFields = New System.Windows.Forms.ComboBox
		Me.cmdBrowseFile = New System.Windows.Forms.Button
		Me.txtSoilsDS = New System.Windows.Forms.TextBox
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.dlgOpenOpen = New System.Windows.Forms.OpenFileDialog
		Me.Label4 = New System.Windows.Forms.Label
		Me.Label5 = New System.Windows.Forms.Label
		Me.Frame2.SuspendLayout()
		Me.Frame1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Text = "Soils Setup"
		Me.ClientSize = New System.Drawing.Size(360, 392)
		Me.Location = New System.Drawing.Point(3, 22)
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
		Me.Name = "frmSoilsSetup"
		Me.Frame2.Text = "Advanced MUSLE Specific Coefficients"
		Me.Frame2.Size = New System.Drawing.Size(329, 157)
		Me.Frame2.Location = New System.Drawing.Point(16, 192)
		Me.Frame2.TabIndex = 15
		Me.Frame2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame2.BackColor = System.Drawing.SystemColors.Control
		Me.Frame2.Enabled = True
		Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame2.Visible = True
		Me.Frame2.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame2.Name = "Frame2"
		Me.txtMUSLEVal.AutoSize = False
		Me.txtMUSLEVal.Size = New System.Drawing.Size(43, 22)
		Me.txtMUSLEVal.Location = New System.Drawing.Point(56, 56)
		Me.txtMUSLEVal.TabIndex = 17
		Me.txtMUSLEVal.Text = "95"
		Me.txtMUSLEVal.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtMUSLEVal.AcceptsReturn = True
		Me.txtMUSLEVal.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtMUSLEVal.BackColor = System.Drawing.SystemColors.Window
		Me.txtMUSLEVal.CausesValidation = True
		Me.txtMUSLEVal.Enabled = True
		Me.txtMUSLEVal.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtMUSLEVal.HideSelection = True
		Me.txtMUSLEVal.ReadOnly = False
		Me.txtMUSLEVal.Maxlength = 0
		Me.txtMUSLEVal.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtMUSLEVal.MultiLine = False
		Me.txtMUSLEVal.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtMUSLEVal.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtMUSLEVal.TabStop = True
		Me.txtMUSLEVal.Visible = True
		Me.txtMUSLEVal.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtMUSLEVal.Name = "txtMUSLEVal"
		Me.txtMUSLEExp.AutoSize = False
		Me.txtMUSLEExp.Size = New System.Drawing.Size(35, 22)
		Me.txtMUSLEExp.Location = New System.Drawing.Point(136, 56)
		Me.txtMUSLEExp.TabIndex = 16
		Me.txtMUSLEExp.Text = "0.56"
		Me.txtMUSLEExp.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtMUSLEExp.AcceptsReturn = True
		Me.txtMUSLEExp.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtMUSLEExp.BackColor = System.Drawing.SystemColors.Window
		Me.txtMUSLEExp.CausesValidation = True
		Me.txtMUSLEExp.Enabled = True
		Me.txtMUSLEExp.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtMUSLEExp.HideSelection = True
		Me.txtMUSLEExp.ReadOnly = False
		Me.txtMUSLEExp.Maxlength = 0
		Me.txtMUSLEExp.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtMUSLEExp.MultiLine = False
		Me.txtMUSLEExp.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtMUSLEExp.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtMUSLEExp.TabStop = True
		Me.txtMUSLEExp.Visible = True
		Me.txtMUSLEExp.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtMUSLEExp.Name = "txtMUSLEExp"
		Me.Label12.Text = "Warning: Q and qp are calculated in English units (acre-feet and cubic feet per second respectively). ""a"" and ""b"" must be derived accordingly."
		Me.Label12.Size = New System.Drawing.Size(305, 57)
		Me.Label12.Location = New System.Drawing.Point(16, 104)
		Me.Label12.TabIndex = 24
		Me.Label12.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label12.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label12.BackColor = System.Drawing.Color.Transparent
		Me.Label12.Enabled = True
		Me.Label12.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label12.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label12.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label12.UseMnemonic = True
		Me.Label12.Visible = True
		Me.Label12.AutoSize = False
		Me.Label12.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label12.Name = "Label12"
		Me.Label7.Text = "b="
		Me.Label7.Size = New System.Drawing.Size(17, 17)
		Me.Label7.Location = New System.Drawing.Point(112, 64)
		Me.Label7.TabIndex = 23
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
		Me.Label6.Text = "a="
		Me.Label6.Size = New System.Drawing.Size(17, 17)
		Me.Label6.Location = New System.Drawing.Point(32, 64)
		Me.Label6.TabIndex = 22
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
		Me.Label11.BackColor = System.Drawing.Color.Transparent
		Me.Label11.Text = "    b"
		Me.Label11.Size = New System.Drawing.Size(25, 13)
		Me.Label11.Location = New System.Drawing.Point(72, 32)
		Me.Label11.TabIndex = 21
		Me.Label11.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label11.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label11.Enabled = True
		Me.Label11.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label11.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label11.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label11.UseMnemonic = True
		Me.Label11.Visible = True
		Me.Label11.AutoSize = False
		Me.Label11.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label11.Name = "Label11"
		Me.Label10.Text = "MUSLE Equation for sediment yield:"
		Me.Label10.Size = New System.Drawing.Size(252, 16)
		Me.Label10.Location = New System.Drawing.Point(18, 19)
		Me.Label10.TabIndex = 20
		Me.Label10.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label10.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label10.BackColor = System.Drawing.SystemColors.Control
		Me.Label10.Enabled = True
		Me.Label10.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label10.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label10.UseMnemonic = True
		Me.Label10.Visible = True
		Me.Label10.AutoSize = False
		Me.Label10.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label10.Name = "Label10"
		Me.Label9.Text = "a * (Q * qp)   * K * C * P * LS"
		Me.Label9.Size = New System.Drawing.Size(212, 20)
		Me.Label9.Location = New System.Drawing.Point(32, 38)
		Me.Label9.TabIndex = 19
		Me.Label9.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label9.BackColor = System.Drawing.SystemColors.Control
		Me.Label9.Enabled = True
		Me.Label9.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label9.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label9.UseMnemonic = True
		Me.Label9.Visible = True
		Me.Label9.AutoSize = False
		Me.Label9.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label9.Name = "Label9"
		Me.Label8.Text = "Locally calibrated MUSLE coefficients can be entered above."
		Me.Label8.Size = New System.Drawing.Size(303, 16)
		Me.Label8.Location = New System.Drawing.Point(16, 88)
		Me.Label8.TabIndex = 18
		Me.Label8.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label8.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label8.BackColor = System.Drawing.SystemColors.Control
		Me.Label8.Enabled = True
		Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label8.UseMnemonic = True
		Me.Label8.Visible = True
		Me.Label8.AutoSize = False
		Me.Label8.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label8.Name = "Label8"
		Me.txtSoilsName.AutoSize = False
		Me.txtSoilsName.Size = New System.Drawing.Size(212, 21)
		Me.txtSoilsName.Location = New System.Drawing.Point(89, 6)
		Me.txtSoilsName.TabIndex = 0
		Me.ToolTip1.SetToolTip(Me.txtSoilsName, "Choose the filled DEM you are using.  This will provide Spatial Analyst with the proper analysis environment.")
		Me.txtSoilsName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtSoilsName.AcceptsReturn = True
		Me.txtSoilsName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtSoilsName.BackColor = System.Drawing.SystemColors.Window
		Me.txtSoilsName.CausesValidation = True
		Me.txtSoilsName.Enabled = True
		Me.txtSoilsName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtSoilsName.HideSelection = True
		Me.txtSoilsName.ReadOnly = False
		Me.txtSoilsName.Maxlength = 0
		Me.txtSoilsName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtSoilsName.MultiLine = False
		Me.txtSoilsName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtSoilsName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtSoilsName.TabStop = True
		Me.txtSoilsName.Visible = True
		Me.txtSoilsName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtSoilsName.Name = "txtSoilsName"
		Me.cmdDEMBrowse.TextAlign = System.Drawing.ContentAlignment.BottomCenter
		Me.cmdDEMBrowse.Size = New System.Drawing.Size(25, 21)
		Me.cmdDEMBrowse.Location = New System.Drawing.Point(307, 39)
		Me.cmdDEMBrowse.Image = CType(resources.GetObject("cmdDEMBrowse.Image"), System.Drawing.Image)
		Me.cmdDEMBrowse.TabIndex = 2
		Me.cmdDEMBrowse.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdDEMBrowse.BackColor = System.Drawing.SystemColors.Control
		Me.cmdDEMBrowse.CausesValidation = True
		Me.cmdDEMBrowse.Enabled = True
		Me.cmdDEMBrowse.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdDEMBrowse.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdDEMBrowse.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdDEMBrowse.TabStop = True
		Me.cmdDEMBrowse.Name = "cmdDEMBrowse"
		Me.txtDEMFile.AutoSize = False
		Me.txtDEMFile.Size = New System.Drawing.Size(212, 21)
		Me.txtDEMFile.Location = New System.Drawing.Point(89, 38)
		Me.txtDEMFile.TabIndex = 1
		Me.ToolTip1.SetToolTip(Me.txtDEMFile, "Choose the filled DEM you are using.  This will provide Spatial Analyst with the proper analysis environment.")
		Me.txtDEMFile.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtDEMFile.AcceptsReturn = True
		Me.txtDEMFile.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtDEMFile.BackColor = System.Drawing.SystemColors.Window
		Me.txtDEMFile.CausesValidation = True
		Me.txtDEMFile.Enabled = True
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
		Me.cmdQuit.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdQuit.Text = "Cancel"
		Me.cmdQuit.Size = New System.Drawing.Size(65, 25)
		Me.cmdQuit.Location = New System.Drawing.Point(274, 358)
		Me.cmdQuit.TabIndex = 8
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
		Me.cmdSave.Text = "OK"
		Me.cmdSave.Size = New System.Drawing.Size(65, 25)
		Me.cmdSave.Location = New System.Drawing.Point(204, 358)
		Me.cmdSave.TabIndex = 7
		Me.cmdSave.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdSave.BackColor = System.Drawing.SystemColors.Control
		Me.cmdSave.CausesValidation = True
		Me.cmdSave.Enabled = True
		Me.cmdSave.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdSave.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdSave.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdSave.TabStop = True
		Me.cmdSave.Name = "cmdSave"
		Me.Frame1.Text = " Soils  "
		Me.Frame1.Size = New System.Drawing.Size(328, 120)
		Me.Frame1.Location = New System.Drawing.Point(16, 69)
		Me.Frame1.TabIndex = 9
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.BackColor = System.Drawing.SystemColors.Control
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame1.Name = "Frame1"
		Me.cboSoilFieldsK.Size = New System.Drawing.Size(143, 21)
		Me.cboSoilFieldsK.Location = New System.Drawing.Point(173, 83)
		Me.cboSoilFieldsK.TabIndex = 6
		Me.cboSoilFieldsK.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboSoilFieldsK.BackColor = System.Drawing.SystemColors.Window
		Me.cboSoilFieldsK.CausesValidation = True
		Me.cboSoilFieldsK.Enabled = True
		Me.cboSoilFieldsK.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboSoilFieldsK.IntegralHeight = True
		Me.cboSoilFieldsK.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboSoilFieldsK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboSoilFieldsK.Sorted = False
		Me.cboSoilFieldsK.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboSoilFieldsK.TabStop = True
		Me.cboSoilFieldsK.Visible = True
		Me.cboSoilFieldsK.Name = "cboSoilFieldsK"
		Me.cboSoilFields.Size = New System.Drawing.Size(143, 21)
		Me.cboSoilFields.Location = New System.Drawing.Point(173, 49)
		Me.cboSoilFields.TabIndex = 5
		Me.cboSoilFields.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboSoilFields.BackColor = System.Drawing.SystemColors.Window
		Me.cboSoilFields.CausesValidation = True
		Me.cboSoilFields.Enabled = True
		Me.cboSoilFields.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboSoilFields.IntegralHeight = True
		Me.cboSoilFields.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboSoilFields.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboSoilFields.Sorted = False
		Me.cboSoilFields.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboSoilFields.TabStop = True
		Me.cboSoilFields.Visible = True
		Me.cboSoilFields.Name = "cboSoilFields"
		Me.cmdBrowseFile.TextAlign = System.Drawing.ContentAlignment.BottomCenter
		Me.cmdBrowseFile.Size = New System.Drawing.Size(25, 21)
		Me.cmdBrowseFile.Location = New System.Drawing.Point(291, 18)
		Me.cmdBrowseFile.Image = CType(resources.GetObject("cmdBrowseFile.Image"), System.Drawing.Image)
		Me.cmdBrowseFile.TabIndex = 4
		Me.cmdBrowseFile.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdBrowseFile.BackColor = System.Drawing.SystemColors.Control
		Me.cmdBrowseFile.CausesValidation = True
		Me.cmdBrowseFile.Enabled = True
		Me.cmdBrowseFile.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdBrowseFile.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdBrowseFile.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdBrowseFile.TabStop = True
		Me.cmdBrowseFile.Name = "cmdBrowseFile"
		Me.txtSoilsDS.AutoSize = False
		Me.txtSoilsDS.Size = New System.Drawing.Size(193, 20)
		Me.txtSoilsDS.Location = New System.Drawing.Point(93, 19)
		Me.txtSoilsDS.TabIndex = 3
		Me.txtSoilsDS.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtSoilsDS.AcceptsReturn = True
		Me.txtSoilsDS.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtSoilsDS.BackColor = System.Drawing.SystemColors.Window
		Me.txtSoilsDS.CausesValidation = True
		Me.txtSoilsDS.Enabled = True
		Me.txtSoilsDS.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtSoilsDS.HideSelection = True
		Me.txtSoilsDS.ReadOnly = False
		Me.txtSoilsDS.Maxlength = 0
		Me.txtSoilsDS.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtSoilsDS.MultiLine = False
		Me.txtSoilsDS.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtSoilsDS.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtSoilsDS.TabStop = True
		Me.txtSoilsDS.Visible = True
		Me.txtSoilsDS.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtSoilsDS.Name = "txtSoilsDS"
		Me.Label2.Text = "K Factor Attribute:"
		Me.Label2.Size = New System.Drawing.Size(122, 19)
		Me.Label2.Location = New System.Drawing.Point(13, 85)
		Me.Label2.TabIndex = 13
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
		Me.Label3.Text = "Hydrologic Soil Group Attribute:"
		Me.Label3.Size = New System.Drawing.Size(159, 30)
		Me.Label3.Location = New System.Drawing.Point(11, 52)
		Me.Label3.TabIndex = 11
		Me.Label3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label3.BackColor = System.Drawing.SystemColors.Control
		Me.Label3.Enabled = True
		Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label3.UseMnemonic = True
		Me.Label3.Visible = True
		Me.Label3.AutoSize = False
		Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label3.Name = "Label3"
		Me.Label1.Text = "Soils Data Set:"
		Me.Label1.Size = New System.Drawing.Size(82, 18)
		Me.Label1.Location = New System.Drawing.Point(12, 20)
		Me.Label1.TabIndex = 10
		Me.Label1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.BackColor = System.Drawing.SystemColors.Control
		Me.Label1.Enabled = True
		Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Label4.Text = "Name:"
		Me.Label4.Size = New System.Drawing.Size(59, 18)
		Me.Label4.Location = New System.Drawing.Point(18, 9)
		Me.Label4.TabIndex = 14
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
		Me.Label5.Text = "DEM GRID:"
		Me.Label5.Size = New System.Drawing.Size(77, 19)
		Me.Label5.Location = New System.Drawing.Point(17, 41)
		Me.Label5.TabIndex = 12
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
		Me.Controls.Add(Frame2)
		Me.Controls.Add(txtSoilsName)
		Me.Controls.Add(cmdDEMBrowse)
		Me.Controls.Add(txtDEMFile)
		Me.Controls.Add(cmdQuit)
		Me.Controls.Add(cmdSave)
		Me.Controls.Add(Frame1)
		Me.Controls.Add(Label4)
		Me.Controls.Add(Label5)
		Me.Frame2.Controls.Add(txtMUSLEVal)
		Me.Frame2.Controls.Add(txtMUSLEExp)
		Me.Frame2.Controls.Add(Label12)
		Me.Frame2.Controls.Add(Label7)
		Me.Frame2.Controls.Add(Label6)
		Me.Frame2.Controls.Add(Label11)
		Me.Frame2.Controls.Add(Label10)
		Me.Frame2.Controls.Add(Label9)
		Me.Frame2.Controls.Add(Label8)
		Me.Frame1.Controls.Add(cboSoilFieldsK)
		Me.Frame1.Controls.Add(cboSoilFields)
		Me.Frame1.Controls.Add(cmdBrowseFile)
		Me.Frame1.Controls.Add(txtSoilsDS)
		Me.Frame1.Controls.Add(Label2)
		Me.Frame1.Controls.Add(Label3)
		Me.Frame1.Controls.Add(Label1)
		Me.Frame2.ResumeLayout(False)
		Me.Frame1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class