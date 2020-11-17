<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmUserWShed
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
	Public WithEvents cmdQuit As System.Windows.Forms.Button
	Public WithEvents cmdCreate As System.Windows.Forms.Button
	Public WithEvents cmdBrowseLS As System.Windows.Forms.Button
	Public WithEvents cboDEMUnits As System.Windows.Forms.ComboBox
	Public WithEvents txtFlowDir As System.Windows.Forms.TextBox
	Public WithEvents cmdBrowseFlowDir As System.Windows.Forms.Button
	Public WithEvents cmdBrowseWS As System.Windows.Forms.Button
	Public WithEvents txtWaterSheds As System.Windows.Forms.TextBox
	Public WithEvents cmdBrowseFlowAcc As System.Windows.Forms.Button
	Public WithEvents txtFlowAcc As System.Windows.Forms.TextBox
	Public WithEvents cmdBrowseDEMFile As System.Windows.Forms.Button
	Public WithEvents txtWSDelinName As System.Windows.Forms.TextBox
	Public WithEvents txtDEMFile As System.Windows.Forms.TextBox
	Public WithEvents txtLS As System.Windows.Forms.TextBox
	Public WithEvents _Label1_2 As System.Windows.Forms.Label
	Public WithEvents _Label1_1 As System.Windows.Forms.Label
	Public WithEvents _Label1_5 As System.Windows.Forms.Label
	Public WithEvents _Label1_4 As System.Windows.Forms.Label
	Public WithEvents _Label1_3 As System.Windows.Forms.Label
	Public WithEvents _Label1_0 As System.Windows.Forms.Label
	Public WithEvents _Label1_12 As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmUserWShed))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdQuit = New System.Windows.Forms.Button
		Me.cmdCreate = New System.Windows.Forms.Button
		Me.cmdBrowseLS = New System.Windows.Forms.Button
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me.cboDEMUnits = New System.Windows.Forms.ComboBox
		Me.txtFlowDir = New System.Windows.Forms.TextBox
		Me.cmdBrowseFlowDir = New System.Windows.Forms.Button
		Me.cmdBrowseWS = New System.Windows.Forms.Button
		Me.txtWaterSheds = New System.Windows.Forms.TextBox
		Me.cmdBrowseFlowAcc = New System.Windows.Forms.Button
		Me.txtFlowAcc = New System.Windows.Forms.TextBox
		Me.cmdBrowseDEMFile = New System.Windows.Forms.Button
		Me.txtWSDelinName = New System.Windows.Forms.TextBox
		Me.txtDEMFile = New System.Windows.Forms.TextBox
		Me.txtLS = New System.Windows.Forms.TextBox
		Me._Label1_2 = New System.Windows.Forms.Label
		Me._Label1_1 = New System.Windows.Forms.Label
		Me._Label1_5 = New System.Windows.Forms.Label
		Me._Label1_4 = New System.Windows.Forms.Label
		Me._Label1_3 = New System.Windows.Forms.Label
		Me._Label1_0 = New System.Windows.Forms.Label
		Me._Label1_12 = New System.Windows.Forms.Label
		Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.Frame1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.Text = "New Watershed Delineation"
		Me.ClientSize = New System.Drawing.Size(486, 273)
		Me.Location = New System.Drawing.Point(4, 30)
		Me.ShowInTaskbar = False
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
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmUserWShed"
		Me.cmdQuit.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdQuit.Text = "Cancel"
		Me.cmdQuit.Size = New System.Drawing.Size(65, 25)
		Me.cmdQuit.Location = New System.Drawing.Point(403, 240)
		Me.cmdQuit.TabIndex = 14
		Me.cmdQuit.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdQuit.BackColor = System.Drawing.SystemColors.Control
		Me.cmdQuit.CausesValidation = True
		Me.cmdQuit.Enabled = True
		Me.cmdQuit.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdQuit.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdQuit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdQuit.TabStop = True
		Me.cmdQuit.Name = "cmdQuit"
		Me.cmdCreate.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdCreate.Text = "OK"
		Me.cmdCreate.Size = New System.Drawing.Size(65, 25)
		Me.cmdCreate.Location = New System.Drawing.Point(328, 240)
		Me.cmdCreate.TabIndex = 13
		Me.cmdCreate.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCreate.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCreate.CausesValidation = True
		Me.cmdCreate.Enabled = True
		Me.cmdCreate.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCreate.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCreate.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCreate.TabStop = True
		Me.cmdCreate.Name = "cmdCreate"
		Me.cmdBrowseLS.TextAlign = System.Drawing.ContentAlignment.BottomCenter
		Me.cmdBrowseLS.Size = New System.Drawing.Size(25, 21)
		Me.cmdBrowseLS.Location = New System.Drawing.Point(432, 168)
		Me.cmdBrowseLS.Image = CType(resources.GetObject("cmdBrowseLS.Image"), System.Drawing.Image)
		Me.cmdBrowseLS.TabIndex = 12
		Me.cmdBrowseLS.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdBrowseLS.BackColor = System.Drawing.SystemColors.Control
		Me.cmdBrowseLS.CausesValidation = True
		Me.cmdBrowseLS.Enabled = True
		Me.cmdBrowseLS.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdBrowseLS.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdBrowseLS.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdBrowseLS.TabStop = True
		Me.cmdBrowseLS.Name = "cmdBrowseLS"
		Me.Frame1.Text = "Define Watershed Delineation"
		Me.Frame1.Size = New System.Drawing.Size(460, 214)
		Me.Frame1.Location = New System.Drawing.Point(16, 8)
		Me.Frame1.TabIndex = 3
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.BackColor = System.Drawing.SystemColors.Control
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame1.Name = "Frame1"
		Me.cboDEMUnits.Size = New System.Drawing.Size(153, 21)
		Me.cboDEMUnits.Location = New System.Drawing.Point(170, 77)
		Me.cboDEMUnits.Items.AddRange(New Object(){"meters", "feet"})
		Me.cboDEMUnits.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cboDEMUnits.TabIndex = 20
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
		Me.txtFlowDir.AutoSize = False
		Me.txtFlowDir.Size = New System.Drawing.Size(245, 19)
		Me.txtFlowDir.Location = New System.Drawing.Point(170, 102)
		Me.txtFlowDir.TabIndex = 19
		Me.txtFlowDir.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtFlowDir.AcceptsReturn = True
		Me.txtFlowDir.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtFlowDir.BackColor = System.Drawing.SystemColors.Window
		Me.txtFlowDir.CausesValidation = True
		Me.txtFlowDir.Enabled = True
		Me.txtFlowDir.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtFlowDir.HideSelection = True
		Me.txtFlowDir.ReadOnly = False
		Me.txtFlowDir.Maxlength = 0
		Me.txtFlowDir.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtFlowDir.MultiLine = False
		Me.txtFlowDir.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtFlowDir.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtFlowDir.TabStop = True
		Me.txtFlowDir.Visible = True
		Me.txtFlowDir.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtFlowDir.Name = "txtFlowDir"
		Me.cmdBrowseFlowDir.TextAlign = System.Drawing.ContentAlignment.BottomCenter
		Me.cmdBrowseFlowDir.Size = New System.Drawing.Size(25, 21)
		Me.cmdBrowseFlowDir.Location = New System.Drawing.Point(416, 102)
		Me.cmdBrowseFlowDir.Image = CType(resources.GetObject("cmdBrowseFlowDir.Image"), System.Drawing.Image)
		Me.cmdBrowseFlowDir.TabIndex = 18
		Me.cmdBrowseFlowDir.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdBrowseFlowDir.BackColor = System.Drawing.SystemColors.Control
		Me.cmdBrowseFlowDir.CausesValidation = True
		Me.cmdBrowseFlowDir.Enabled = True
		Me.cmdBrowseFlowDir.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdBrowseFlowDir.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdBrowseFlowDir.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdBrowseFlowDir.TabStop = True
		Me.cmdBrowseFlowDir.Name = "cmdBrowseFlowDir"
		Me.cmdBrowseWS.TextAlign = System.Drawing.ContentAlignment.BottomCenter
		Me.cmdBrowseWS.Size = New System.Drawing.Size(25, 21)
		Me.cmdBrowseWS.Location = New System.Drawing.Point(416, 184)
		Me.cmdBrowseWS.Image = CType(resources.GetObject("cmdBrowseWS.Image"), System.Drawing.Image)
		Me.cmdBrowseWS.TabIndex = 17
		Me.cmdBrowseWS.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdBrowseWS.BackColor = System.Drawing.SystemColors.Control
		Me.cmdBrowseWS.CausesValidation = True
		Me.cmdBrowseWS.Enabled = True
		Me.cmdBrowseWS.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdBrowseWS.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdBrowseWS.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdBrowseWS.TabStop = True
		Me.cmdBrowseWS.Name = "cmdBrowseWS"
		Me.txtWaterSheds.AutoSize = False
		Me.txtWaterSheds.Size = New System.Drawing.Size(245, 19)
		Me.txtWaterSheds.Location = New System.Drawing.Point(170, 184)
		Me.txtWaterSheds.TabIndex = 15
		Me.txtWaterSheds.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtWaterSheds.AcceptsReturn = True
		Me.txtWaterSheds.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtWaterSheds.BackColor = System.Drawing.SystemColors.Window
		Me.txtWaterSheds.CausesValidation = True
		Me.txtWaterSheds.Enabled = True
		Me.txtWaterSheds.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtWaterSheds.HideSelection = True
		Me.txtWaterSheds.ReadOnly = False
		Me.txtWaterSheds.Maxlength = 0
		Me.txtWaterSheds.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtWaterSheds.MultiLine = False
		Me.txtWaterSheds.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtWaterSheds.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtWaterSheds.TabStop = True
		Me.txtWaterSheds.Visible = True
		Me.txtWaterSheds.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtWaterSheds.Name = "txtWaterSheds"
		Me.cmdBrowseFlowAcc.TextAlign = System.Drawing.ContentAlignment.BottomCenter
		Me.cmdBrowseFlowAcc.Size = New System.Drawing.Size(25, 21)
		Me.cmdBrowseFlowAcc.Location = New System.Drawing.Point(416, 130)
		Me.cmdBrowseFlowAcc.Image = CType(resources.GetObject("cmdBrowseFlowAcc.Image"), System.Drawing.Image)
		Me.cmdBrowseFlowAcc.TabIndex = 11
		Me.cmdBrowseFlowAcc.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdBrowseFlowAcc.BackColor = System.Drawing.SystemColors.Control
		Me.cmdBrowseFlowAcc.CausesValidation = True
		Me.cmdBrowseFlowAcc.Enabled = True
		Me.cmdBrowseFlowAcc.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdBrowseFlowAcc.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdBrowseFlowAcc.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdBrowseFlowAcc.TabStop = True
		Me.cmdBrowseFlowAcc.Name = "cmdBrowseFlowAcc"
		Me.txtFlowAcc.AutoSize = False
		Me.txtFlowAcc.Size = New System.Drawing.Size(245, 19)
		Me.txtFlowAcc.Location = New System.Drawing.Point(170, 130)
		Me.txtFlowAcc.TabIndex = 10
		Me.txtFlowAcc.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtFlowAcc.AcceptsReturn = True
		Me.txtFlowAcc.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtFlowAcc.BackColor = System.Drawing.SystemColors.Window
		Me.txtFlowAcc.CausesValidation = True
		Me.txtFlowAcc.Enabled = True
		Me.txtFlowAcc.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtFlowAcc.HideSelection = True
		Me.txtFlowAcc.ReadOnly = False
		Me.txtFlowAcc.Maxlength = 0
		Me.txtFlowAcc.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtFlowAcc.MultiLine = False
		Me.txtFlowAcc.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtFlowAcc.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtFlowAcc.TabStop = True
		Me.txtFlowAcc.Visible = True
		Me.txtFlowAcc.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtFlowAcc.Name = "txtFlowAcc"
		Me.cmdBrowseDEMFile.TextAlign = System.Drawing.ContentAlignment.BottomCenter
		Me.cmdBrowseDEMFile.Size = New System.Drawing.Size(25, 21)
		Me.cmdBrowseDEMFile.Location = New System.Drawing.Point(416, 49)
		Me.cmdBrowseDEMFile.Image = CType(resources.GetObject("cmdBrowseDEMFile.Image"), System.Drawing.Image)
		Me.cmdBrowseDEMFile.TabIndex = 2
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
		Me.txtWSDelinName.Size = New System.Drawing.Size(245, 19)
		Me.txtWSDelinName.Location = New System.Drawing.Point(170, 24)
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
		Me.txtDEMFile.AutoSize = False
		Me.txtDEMFile.Size = New System.Drawing.Size(245, 19)
		Me.txtDEMFile.Location = New System.Drawing.Point(170, 52)
		Me.txtDEMFile.TabIndex = 1
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
		Me.txtLS.AutoSize = False
		Me.txtLS.Size = New System.Drawing.Size(245, 19)
		Me.txtLS.Location = New System.Drawing.Point(170, 157)
		Me.txtLS.TabIndex = 4
		Me.txtLS.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtLS.AcceptsReturn = True
		Me.txtLS.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtLS.BackColor = System.Drawing.SystemColors.Window
		Me.txtLS.CausesValidation = True
		Me.txtLS.Enabled = True
		Me.txtLS.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtLS.HideSelection = True
		Me.txtLS.ReadOnly = False
		Me.txtLS.Maxlength = 0
		Me.txtLS.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtLS.MultiLine = False
		Me.txtLS.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtLS.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtLS.TabStop = True
		Me.txtLS.Visible = True
		Me.txtLS.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtLS.Name = "txtLS"
		Me._Label1_2.Text = "DEM Units:"
		Me._Label1_2.Size = New System.Drawing.Size(54, 13)
		Me._Label1_2.Location = New System.Drawing.Point(96, 80)
		Me._Label1_2.TabIndex = 21
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
		Me._Label1_1.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_1.Text = "Watersheds:"
		Me._Label1_1.Size = New System.Drawing.Size(60, 13)
		Me._Label1_1.Location = New System.Drawing.Point(94, 184)
		Me._Label1_1.TabIndex = 16
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
		Me._Label1_5.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_5.Text = "Length-slope Grid:"
		Me._Label1_5.Size = New System.Drawing.Size(86, 13)
		Me._Label1_5.Location = New System.Drawing.Point(68, 159)
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
		Me._Label1_4.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_4.Text = "Flow Accumulation Grid:"
		Me._Label1_4.Size = New System.Drawing.Size(114, 13)
		Me._Label1_4.Location = New System.Drawing.Point(40, 132)
		Me._Label1_4.TabIndex = 8
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
		Me._Label1_3.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_3.Text = "Watershed Delineation Name:"
		Me._Label1_3.Size = New System.Drawing.Size(142, 13)
		Me._Label1_3.Location = New System.Drawing.Point(17, 27)
		Me._Label1_3.TabIndex = 7
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
		Me._Label1_0.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_0.Text = "Flow Direction Grid:"
		Me._Label1_0.Size = New System.Drawing.Size(92, 13)
		Me._Label1_0.Location = New System.Drawing.Point(62, 107)
		Me._Label1_0.TabIndex = 6
		Me._Label1_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_0.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_0.Enabled = True
		Me._Label1_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_0.UseMnemonic = True
		Me._Label1_0.Visible = True
		Me._Label1_0.AutoSize = True
		Me._Label1_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_0.Name = "_Label1_0"
		Me._Label1_12.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_12.Text = "DEM Grid:"
		Me._Label1_12.Size = New System.Drawing.Size(49, 17)
		Me._Label1_12.Location = New System.Drawing.Point(104, 56)
		Me._Label1_12.TabIndex = 5
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
		Me.Controls.Add(cmdQuit)
		Me.Controls.Add(cmdCreate)
		Me.Controls.Add(cmdBrowseLS)
		Me.Controls.Add(Frame1)
		Me.Frame1.Controls.Add(cboDEMUnits)
		Me.Frame1.Controls.Add(txtFlowDir)
		Me.Frame1.Controls.Add(cmdBrowseFlowDir)
		Me.Frame1.Controls.Add(cmdBrowseWS)
		Me.Frame1.Controls.Add(txtWaterSheds)
		Me.Frame1.Controls.Add(cmdBrowseFlowAcc)
		Me.Frame1.Controls.Add(txtFlowAcc)
		Me.Frame1.Controls.Add(cmdBrowseDEMFile)
		Me.Frame1.Controls.Add(txtWSDelinName)
		Me.Frame1.Controls.Add(txtDEMFile)
		Me.Frame1.Controls.Add(txtLS)
		Me.Frame1.Controls.Add(_Label1_2)
		Me.Frame1.Controls.Add(_Label1_1)
		Me.Frame1.Controls.Add(_Label1_5)
		Me.Frame1.Controls.Add(_Label1_4)
		Me.Frame1.Controls.Add(_Label1_3)
		Me.Frame1.Controls.Add(_Label1_0)
		Me.Frame1.Controls.Add(_Label1_12)
		Me.Label1.SetIndex(_Label1_2, CType(2, Short))
		Me.Label1.SetIndex(_Label1_1, CType(1, Short))
		Me.Label1.SetIndex(_Label1_5, CType(5, Short))
		Me.Label1.SetIndex(_Label1_4, CType(4, Short))
		Me.Label1.SetIndex(_Label1_3, CType(3, Short))
		Me.Label1.SetIndex(_Label1_0, CType(0, Short))
		Me.Label1.SetIndex(_Label1_12, CType(12, Short))
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.Frame1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class