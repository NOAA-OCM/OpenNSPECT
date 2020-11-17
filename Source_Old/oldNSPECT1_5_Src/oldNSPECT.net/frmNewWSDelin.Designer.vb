<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmNewWSDelin
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
	Public WithEvents chkHydroCorr As System.Windows.Forms.CheckBox
	Public WithEvents cmdBrowseDEMFile As System.Windows.Forms.Button
	Public WithEvents txtWSDelinName As System.Windows.Forms.TextBox
	Public WithEvents cboDEMUnits As System.Windows.Forms.ComboBox
	Public WithEvents txtDEMFile As System.Windows.Forms.TextBox
	Public WithEvents cboSubWSSize As System.Windows.Forms.ComboBox
	Public WithEvents cmdOptions As System.Windows.Forms.Button
	Public WithEvents cboStreamLayer As System.Windows.Forms.ComboBox
	Public WithEvents chkStreamAgree As System.Windows.Forms.CheckBox
	Public WithEvents _lblStream_0 As System.Windows.Forms.Label
	Public WithEvents frmAdvanced As System.Windows.Forms.GroupBox
	Public WithEvents _Label1_3 As System.Windows.Forms.Label
	Public WithEvents _Label1_2 As System.Windows.Forms.Label
	Public WithEvents _Label1_12 As System.Windows.Forms.Label
	Public WithEvents _Label1_1 As System.Windows.Forms.Label
	Public WithEvents frmMain As System.Windows.Forms.GroupBox
	Public WithEvents cmdCreate As System.Windows.Forms.Button
	Public WithEvents cmdQuit As System.Windows.Forms.Button
	Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblStream As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmNewWSDelin))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.frmMain = New System.Windows.Forms.GroupBox
		Me.chkHydroCorr = New System.Windows.Forms.CheckBox
		Me.cmdBrowseDEMFile = New System.Windows.Forms.Button
		Me.txtWSDelinName = New System.Windows.Forms.TextBox
		Me.cboDEMUnits = New System.Windows.Forms.ComboBox
		Me.txtDEMFile = New System.Windows.Forms.TextBox
		Me.cboSubWSSize = New System.Windows.Forms.ComboBox
		Me.frmAdvanced = New System.Windows.Forms.GroupBox
		Me.cmdOptions = New System.Windows.Forms.Button
		Me.cboStreamLayer = New System.Windows.Forms.ComboBox
		Me.chkStreamAgree = New System.Windows.Forms.CheckBox
		Me._lblStream_0 = New System.Windows.Forms.Label
		Me._Label1_3 = New System.Windows.Forms.Label
		Me._Label1_2 = New System.Windows.Forms.Label
		Me._Label1_12 = New System.Windows.Forms.Label
		Me._Label1_1 = New System.Windows.Forms.Label
		Me.cmdCreate = New System.Windows.Forms.Button
		Me.cmdQuit = New System.Windows.Forms.Button
		Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.lblStream = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.frmMain.SuspendLayout()
		Me.frmAdvanced.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.lblStream, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "New Watershed Delineation"
		Me.ClientSize = New System.Drawing.Size(469, 209)
		Me.Location = New System.Drawing.Point(213, 196)
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
		Me.Name = "frmNewWSDelin"
		Me.frmMain.Text = "Create a new watershed delineation  "
		Me.frmMain.Size = New System.Drawing.Size(437, 160)
		Me.frmMain.Location = New System.Drawing.Point(13, 9)
		Me.frmMain.TabIndex = 7
		Me.frmMain.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frmMain.BackColor = System.Drawing.SystemColors.Control
		Me.frmMain.Enabled = True
		Me.frmMain.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frmMain.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmMain.Visible = True
		Me.frmMain.Padding = New System.Windows.Forms.Padding(0)
		Me.frmMain.Name = "frmMain"
		Me.chkHydroCorr.Text = "DEM is hyrdologically correct (filled)"
		Me.chkHydroCorr.Size = New System.Drawing.Size(241, 15)
		Me.chkHydroCorr.Location = New System.Drawing.Point(119, 76)
		Me.chkHydroCorr.TabIndex = 17
		Me.chkHydroCorr.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkHydroCorr.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkHydroCorr.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkHydroCorr.BackColor = System.Drawing.SystemColors.Control
		Me.chkHydroCorr.CausesValidation = True
		Me.chkHydroCorr.Enabled = True
		Me.chkHydroCorr.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkHydroCorr.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkHydroCorr.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkHydroCorr.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkHydroCorr.TabStop = True
		Me.chkHydroCorr.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkHydroCorr.Visible = True
		Me.chkHydroCorr.Name = "chkHydroCorr"
		Me.cmdBrowseDEMFile.TextAlign = System.Drawing.ContentAlignment.BottomCenter
		Me.cmdBrowseDEMFile.Size = New System.Drawing.Size(25, 21)
		Me.cmdBrowseDEMFile.Location = New System.Drawing.Point(391, 46)
		Me.cmdBrowseDEMFile.Image = CType(resources.GetObject("cmdBrowseDEMFile.Image"), System.Drawing.Image)
		Me.cmdBrowseDEMFile.TabIndex = 1
		Me.cmdBrowseDEMFile.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdBrowseDEMFile.BackColor = System.Drawing.SystemColors.Control
		Me.cmdBrowseDEMFile.CausesValidation = True
		Me.cmdBrowseDEMFile.Enabled = True
		Me.cmdBrowseDEMFile.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdBrowseDEMFile.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdBrowseDEMFile.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdBrowseDEMFile.TabStop = True
		Me.cmdBrowseDEMFile.Name = "cmdBrowseDEMFile"
		Me.txtWSDelinName.AutoSize = False
		Me.txtWSDelinName.Size = New System.Drawing.Size(134, 19)
		Me.txtWSDelinName.Location = New System.Drawing.Point(118, 20)
		Me.txtWSDelinName.TabIndex = 0
		Me.txtWSDelinName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtWSDelinName.AcceptsReturn = True
		Me.txtWSDelinName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtWSDelinName.BackColor = System.Drawing.SystemColors.Window
		Me.txtWSDelinName.CausesValidation = True
		Me.txtWSDelinName.Enabled = True
		Me.txtWSDelinName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtWSDelinName.HideSelection = True
		Me.txtWSDelinName.ReadOnly = False
		Me.txtWSDelinName.Maxlength = 0
		Me.txtWSDelinName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtWSDelinName.MultiLine = False
		Me.txtWSDelinName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtWSDelinName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtWSDelinName.TabStop = True
		Me.txtWSDelinName.Visible = True
		Me.txtWSDelinName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtWSDelinName.Name = "txtWSDelinName"
		Me.cboDEMUnits.Size = New System.Drawing.Size(152, 21)
		Me.cboDEMUnits.Location = New System.Drawing.Point(119, 99)
		Me.cboDEMUnits.Items.AddRange(New Object(){"meters", "feet"})
		Me.cboDEMUnits.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cboDEMUnits.TabIndex = 2
		Me.cboDEMUnits.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboDEMUnits.BackColor = System.Drawing.SystemColors.Window
		Me.cboDEMUnits.CausesValidation = True
		Me.cboDEMUnits.Enabled = True
		Me.cboDEMUnits.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboDEMUnits.IntegralHeight = True
		Me.cboDEMUnits.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboDEMUnits.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboDEMUnits.Sorted = False
		Me.cboDEMUnits.TabStop = True
		Me.cboDEMUnits.Visible = True
		Me.cboDEMUnits.Name = "cboDEMUnits"
		Me.txtDEMFile.AutoSize = False
		Me.txtDEMFile.Size = New System.Drawing.Size(271, 19)
		Me.txtDEMFile.Location = New System.Drawing.Point(118, 48)
		Me.txtDEMFile.TabIndex = 4
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
		Me.cboSubWSSize.Size = New System.Drawing.Size(151, 21)
		Me.cboSubWSSize.Location = New System.Drawing.Point(119, 129)
		Me.cboSubWSSize.Items.AddRange(New Object(){"small", "medium", "large"})
		Me.cboSubWSSize.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cboSubWSSize.TabIndex = 3
		Me.cboSubWSSize.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboSubWSSize.BackColor = System.Drawing.SystemColors.Window
		Me.cboSubWSSize.CausesValidation = True
		Me.cboSubWSSize.Enabled = True
		Me.cboSubWSSize.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboSubWSSize.IntegralHeight = True
		Me.cboSubWSSize.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboSubWSSize.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboSubWSSize.Sorted = False
		Me.cboSubWSSize.TabStop = True
		Me.cboSubWSSize.Visible = True
		Me.cboSubWSSize.Name = "cboSubWSSize"
		Me.frmAdvanced.Text = "Advanced Parameters (optional) "
		Me.frmAdvanced.Size = New System.Drawing.Size(406, 110)
		Me.frmAdvanced.Location = New System.Drawing.Point(19, 154)
		Me.frmAdvanced.TabIndex = 12
		Me.frmAdvanced.Visible = False
		Me.frmAdvanced.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frmAdvanced.BackColor = System.Drawing.SystemColors.Control
		Me.frmAdvanced.Enabled = True
		Me.frmAdvanced.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frmAdvanced.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmAdvanced.Padding = New System.Windows.Forms.Padding(0)
		Me.frmAdvanced.Name = "frmAdvanced"
		Me.cmdOptions.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOptions.Text = "Options..."
		Me.cmdOptions.Enabled = False
		Me.cmdOptions.Size = New System.Drawing.Size(59, 25)
		Me.cmdOptions.Location = New System.Drawing.Point(281, 71)
		Me.cmdOptions.TabIndex = 16
		Me.cmdOptions.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOptions.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOptions.CausesValidation = True
		Me.cmdOptions.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOptions.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOptions.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOptions.TabStop = True
		Me.cmdOptions.Name = "cmdOptions"
		Me.cboStreamLayer.Enabled = False
		Me.cboStreamLayer.Size = New System.Drawing.Size(155, 21)
		Me.cboStreamLayer.Location = New System.Drawing.Point(122, 73)
		Me.cboStreamLayer.Items.AddRange(New Object(){"Stream1"})
		Me.cboStreamLayer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cboStreamLayer.TabIndex = 14
		Me.cboStreamLayer.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboStreamLayer.BackColor = System.Drawing.SystemColors.Window
		Me.cboStreamLayer.CausesValidation = True
		Me.cboStreamLayer.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboStreamLayer.IntegralHeight = True
		Me.cboStreamLayer.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboStreamLayer.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboStreamLayer.Sorted = False
		Me.cboStreamLayer.TabStop = True
		Me.cboStreamLayer.Visible = True
		Me.cboStreamLayer.Name = "cboStreamLayer"
		Me.chkStreamAgree.Text = "Force Stream Agreement"
		Me.chkStreamAgree.Enabled = False
		Me.chkStreamAgree.Size = New System.Drawing.Size(149, 19)
		Me.chkStreamAgree.Location = New System.Drawing.Point(48, 49)
		Me.chkStreamAgree.TabIndex = 13
		Me.chkStreamAgree.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkStreamAgree.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkStreamAgree.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkStreamAgree.BackColor = System.Drawing.SystemColors.Control
		Me.chkStreamAgree.CausesValidation = True
		Me.chkStreamAgree.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkStreamAgree.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkStreamAgree.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkStreamAgree.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkStreamAgree.TabStop = True
		Me.chkStreamAgree.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkStreamAgree.Visible = True
		Me.chkStreamAgree.Name = "chkStreamAgree"
		Me._lblStream_0.Text = "Stream Layer:"
		Me._lblStream_0.Enabled = False
		Me._lblStream_0.Size = New System.Drawing.Size(65, 13)
		Me._lblStream_0.Location = New System.Drawing.Point(49, 76)
		Me._lblStream_0.TabIndex = 15
		Me._lblStream_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblStream_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblStream_0.BackColor = System.Drawing.SystemColors.Control
		Me._lblStream_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblStream_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblStream_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblStream_0.UseMnemonic = True
		Me._lblStream_0.Visible = True
		Me._lblStream_0.AutoSize = True
		Me._lblStream_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblStream_0.Name = "_lblStream_0"
		Me._Label1_3.Text = "Delineation Name:"
		Me._Label1_3.Size = New System.Drawing.Size(87, 13)
		Me._Label1_3.Location = New System.Drawing.Point(17, 24)
		Me._Label1_3.TabIndex = 11
		Me._Label1_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_3.TextAlign = System.Drawing.ContentAlignment.TopLeft
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
		Me._Label1_2.Text = "DEM Units:"
		Me._Label1_2.Size = New System.Drawing.Size(54, 13)
		Me._Label1_2.Location = New System.Drawing.Point(18, 102)
		Me._Label1_2.TabIndex = 10
		Me._Label1_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
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
		Me._Label1_12.Text = "DEM Grid:"
		Me._Label1_12.Size = New System.Drawing.Size(49, 13)
		Me._Label1_12.Location = New System.Drawing.Point(17, 52)
		Me._Label1_12.TabIndex = 9
		Me._Label1_12.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_12.TextAlign = System.Drawing.ContentAlignment.TopLeft
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
		Me._Label1_1.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_1.Text = "Subwatershed Size:"
		Me._Label1_1.Size = New System.Drawing.Size(94, 13)
		Me._Label1_1.Location = New System.Drawing.Point(18, 129)
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
		Me.cmdCreate.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdCreate.Text = "OK"
		Me.cmdCreate.Enabled = False
		Me.cmdCreate.Size = New System.Drawing.Size(65, 25)
		Me.cmdCreate.Location = New System.Drawing.Point(282, 178)
		Me.cmdCreate.TabIndex = 5
		Me.cmdCreate.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCreate.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCreate.CausesValidation = True
		Me.cmdCreate.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCreate.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCreate.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCreate.TabStop = True
		Me.cmdCreate.Name = "cmdCreate"
		Me.cmdQuit.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdQuit.Text = "Cancel"
		Me.cmdQuit.Size = New System.Drawing.Size(65, 25)
		Me.cmdQuit.Location = New System.Drawing.Point(357, 178)
		Me.cmdQuit.TabIndex = 6
		Me.cmdQuit.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdQuit.BackColor = System.Drawing.SystemColors.Control
		Me.cmdQuit.CausesValidation = True
		Me.cmdQuit.Enabled = True
		Me.cmdQuit.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdQuit.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdQuit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdQuit.TabStop = True
		Me.cmdQuit.Name = "cmdQuit"
		Me.Controls.Add(frmMain)
		Me.Controls.Add(cmdCreate)
		Me.Controls.Add(cmdQuit)
		Me.frmMain.Controls.Add(chkHydroCorr)
		Me.frmMain.Controls.Add(cmdBrowseDEMFile)
		Me.frmMain.Controls.Add(txtWSDelinName)
		Me.frmMain.Controls.Add(cboDEMUnits)
		Me.frmMain.Controls.Add(txtDEMFile)
		Me.frmMain.Controls.Add(cboSubWSSize)
		Me.frmMain.Controls.Add(frmAdvanced)
		Me.frmMain.Controls.Add(_Label1_3)
		Me.frmMain.Controls.Add(_Label1_2)
		Me.frmMain.Controls.Add(_Label1_12)
		Me.frmMain.Controls.Add(_Label1_1)
		Me.frmAdvanced.Controls.Add(cmdOptions)
		Me.frmAdvanced.Controls.Add(cboStreamLayer)
		Me.frmAdvanced.Controls.Add(chkStreamAgree)
		Me.frmAdvanced.Controls.Add(_lblStream_0)
		Me.Label1.SetIndex(_Label1_3, CType(3, Short))
		Me.Label1.SetIndex(_Label1_2, CType(2, Short))
		Me.Label1.SetIndex(_Label1_12, CType(12, Short))
		Me.Label1.SetIndex(_Label1_1, CType(1, Short))
		Me.lblStream.SetIndex(_lblStream_0, CType(0, Short))
		CType(Me.lblStream, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.frmMain.ResumeLayout(False)
		Me.frmAdvanced.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class