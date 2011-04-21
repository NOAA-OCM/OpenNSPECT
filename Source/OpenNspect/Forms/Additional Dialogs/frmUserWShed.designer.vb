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
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmUserWShed))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdQuit = New System.Windows.Forms.Button()
        Me.cmdCreate = New System.Windows.Forms.Button()
        Me.cmdBrowseLS = New System.Windows.Forms.Button()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.cboDEMUnits = New System.Windows.Forms.ComboBox()
        Me.txtFlowDir = New System.Windows.Forms.TextBox()
        Me.cmdBrowseFlowDir = New System.Windows.Forms.Button()
        Me.cmdBrowseWS = New System.Windows.Forms.Button()
        Me.txtWaterSheds = New System.Windows.Forms.TextBox()
        Me.cmdBrowseFlowAcc = New System.Windows.Forms.Button()
        Me.txtFlowAcc = New System.Windows.Forms.TextBox()
        Me.cmdBrowseDEMFile = New System.Windows.Forms.Button()
        Me.txtWSDelinName = New System.Windows.Forms.TextBox()
        Me.txtDEMFile = New System.Windows.Forms.TextBox()
        Me.txtLS = New System.Windows.Forms.TextBox()
        Me._Label1_2 = New System.Windows.Forms.Label()
        Me._Label1_1 = New System.Windows.Forms.Label()
        Me._Label1_5 = New System.Windows.Forms.Label()
        Me._Label1_4 = New System.Windows.Forms.Label()
        Me._Label1_3 = New System.Windows.Forms.Label()
        Me._Label1_0 = New System.Windows.Forms.Label()
        Me._Label1_12 = New System.Windows.Forms.Label()
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdQuit
        '
        Me.cmdQuit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdQuit.BackColor = System.Drawing.SystemColors.Control
        Me.cmdQuit.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdQuit.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdQuit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdQuit.Location = New System.Drawing.Point(501, 329)
        Me.cmdQuit.Name = "cmdQuit"
        Me.cmdQuit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdQuit.Size = New System.Drawing.Size(65, 25)
        Me.cmdQuit.TabIndex = 14
        Me.cmdQuit.Text = "Cancel"
        Me.cmdQuit.UseVisualStyleBackColor = False
        '
        'cmdCreate
        '
        Me.cmdCreate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCreate.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCreate.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCreate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCreate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCreate.Location = New System.Drawing.Point(426, 329)
        Me.cmdCreate.Name = "cmdCreate"
        Me.cmdCreate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCreate.Size = New System.Drawing.Size(65, 25)
        Me.cmdCreate.TabIndex = 13
        Me.cmdCreate.Text = "OK"
        Me.cmdCreate.UseVisualStyleBackColor = False
        '
        'cmdBrowseLS
        '
        Me.cmdBrowseLS.BackColor = System.Drawing.SystemColors.Control
        Me.cmdBrowseLS.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdBrowseLS.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowseLS.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdBrowseLS.Image = CType(resources.GetObject("cmdBrowseLS.Image"), System.Drawing.Image)
        Me.cmdBrowseLS.Location = New System.Drawing.Point(522, 157)
        Me.cmdBrowseLS.Name = "cmdBrowseLS"
        Me.cmdBrowseLS.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdBrowseLS.Size = New System.Drawing.Size(25, 21)
        Me.cmdBrowseLS.TabIndex = 12
        Me.cmdBrowseLS.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdBrowseLS.UseVisualStyleBackColor = False
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.cboDEMUnits)
        Me.Frame1.Controls.Add(Me.txtFlowDir)
        Me.Frame1.Controls.Add(Me.cmdBrowseLS)
        Me.Frame1.Controls.Add(Me.cmdBrowseFlowDir)
        Me.Frame1.Controls.Add(Me.cmdBrowseWS)
        Me.Frame1.Controls.Add(Me.txtWaterSheds)
        Me.Frame1.Controls.Add(Me.cmdBrowseFlowAcc)
        Me.Frame1.Controls.Add(Me.txtFlowAcc)
        Me.Frame1.Controls.Add(Me.cmdBrowseDEMFile)
        Me.Frame1.Controls.Add(Me.txtWSDelinName)
        Me.Frame1.Controls.Add(Me.txtDEMFile)
        Me.Frame1.Controls.Add(Me.txtLS)
        Me.Frame1.Controls.Add(Me._Label1_2)
        Me.Frame1.Controls.Add(Me._Label1_1)
        Me.Frame1.Controls.Add(Me._Label1_5)
        Me.Frame1.Controls.Add(Me._Label1_4)
        Me.Frame1.Controls.Add(Me._Label1_3)
        Me.Frame1.Controls.Add(Me._Label1_0)
        Me.Frame1.Controls.Add(Me._Label1_12)
        Me.Frame1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(16, 8)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(550, 300)
        Me.Frame1.TabIndex = 3
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Define Watershed Delineation"
        '
        'cboDEMUnits
        '
        Me.cboDEMUnits.BackColor = System.Drawing.SystemColors.Window
        Me.cboDEMUnits.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboDEMUnits.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDEMUnits.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboDEMUnits.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboDEMUnits.Items.AddRange(New Object() {"meters", "feet"})
        Me.cboDEMUnits.Location = New System.Drawing.Point(170, 77)
        Me.cboDEMUnits.Name = "cboDEMUnits"
        Me.cboDEMUnits.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboDEMUnits.Size = New System.Drawing.Size(153, 22)
        Me.cboDEMUnits.TabIndex = 20
        '
        'txtFlowDir
        '
        Me.txtFlowDir.AcceptsReturn = True
        Me.txtFlowDir.BackColor = System.Drawing.SystemColors.Window
        Me.txtFlowDir.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFlowDir.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFlowDir.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFlowDir.Location = New System.Drawing.Point(170, 102)
        Me.txtFlowDir.MaxLength = 0
        Me.txtFlowDir.Name = "txtFlowDir"
        Me.txtFlowDir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFlowDir.Size = New System.Drawing.Size(346, 20)
        Me.txtFlowDir.TabIndex = 19
        '
        'cmdBrowseFlowDir
        '
        Me.cmdBrowseFlowDir.BackColor = System.Drawing.SystemColors.Control
        Me.cmdBrowseFlowDir.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdBrowseFlowDir.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowseFlowDir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdBrowseFlowDir.Image = CType(resources.GetObject("cmdBrowseFlowDir.Image"), System.Drawing.Image)
        Me.cmdBrowseFlowDir.Location = New System.Drawing.Point(522, 100)
        Me.cmdBrowseFlowDir.Name = "cmdBrowseFlowDir"
        Me.cmdBrowseFlowDir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdBrowseFlowDir.Size = New System.Drawing.Size(25, 21)
        Me.cmdBrowseFlowDir.TabIndex = 18
        Me.cmdBrowseFlowDir.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdBrowseFlowDir.UseVisualStyleBackColor = False
        '
        'cmdBrowseWS
        '
        Me.cmdBrowseWS.BackColor = System.Drawing.SystemColors.Control
        Me.cmdBrowseWS.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdBrowseWS.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowseWS.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdBrowseWS.Image = CType(resources.GetObject("cmdBrowseWS.Image"), System.Drawing.Image)
        Me.cmdBrowseWS.Location = New System.Drawing.Point(522, 183)
        Me.cmdBrowseWS.Name = "cmdBrowseWS"
        Me.cmdBrowseWS.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdBrowseWS.Size = New System.Drawing.Size(25, 21)
        Me.cmdBrowseWS.TabIndex = 17
        Me.cmdBrowseWS.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdBrowseWS.UseVisualStyleBackColor = False
        '
        'txtWaterSheds
        '
        Me.txtWaterSheds.AcceptsReturn = True
        Me.txtWaterSheds.BackColor = System.Drawing.SystemColors.Window
        Me.txtWaterSheds.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtWaterSheds.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtWaterSheds.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtWaterSheds.Location = New System.Drawing.Point(170, 184)
        Me.txtWaterSheds.MaxLength = 0
        Me.txtWaterSheds.Name = "txtWaterSheds"
        Me.txtWaterSheds.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtWaterSheds.Size = New System.Drawing.Size(346, 20)
        Me.txtWaterSheds.TabIndex = 15
        '
        'cmdBrowseFlowAcc
        '
        Me.cmdBrowseFlowAcc.BackColor = System.Drawing.SystemColors.Control
        Me.cmdBrowseFlowAcc.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdBrowseFlowAcc.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowseFlowAcc.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdBrowseFlowAcc.Image = CType(resources.GetObject("cmdBrowseFlowAcc.Image"), System.Drawing.Image)
        Me.cmdBrowseFlowAcc.Location = New System.Drawing.Point(522, 132)
        Me.cmdBrowseFlowAcc.Name = "cmdBrowseFlowAcc"
        Me.cmdBrowseFlowAcc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdBrowseFlowAcc.Size = New System.Drawing.Size(25, 21)
        Me.cmdBrowseFlowAcc.TabIndex = 11
        Me.cmdBrowseFlowAcc.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdBrowseFlowAcc.UseVisualStyleBackColor = False
        '
        'txtFlowAcc
        '
        Me.txtFlowAcc.AcceptsReturn = True
        Me.txtFlowAcc.BackColor = System.Drawing.SystemColors.Window
        Me.txtFlowAcc.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFlowAcc.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFlowAcc.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFlowAcc.Location = New System.Drawing.Point(170, 130)
        Me.txtFlowAcc.MaxLength = 0
        Me.txtFlowAcc.Name = "txtFlowAcc"
        Me.txtFlowAcc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFlowAcc.Size = New System.Drawing.Size(346, 20)
        Me.txtFlowAcc.TabIndex = 10
        '
        'cmdBrowseDEMFile
        '
        Me.cmdBrowseDEMFile.BackColor = System.Drawing.SystemColors.Control
        Me.cmdBrowseDEMFile.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdBrowseDEMFile.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowseDEMFile.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdBrowseDEMFile.Image = CType(resources.GetObject("cmdBrowseDEMFile.Image"), System.Drawing.Image)
        Me.cmdBrowseDEMFile.Location = New System.Drawing.Point(522, 51)
        Me.cmdBrowseDEMFile.Name = "cmdBrowseDEMFile"
        Me.cmdBrowseDEMFile.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdBrowseDEMFile.Size = New System.Drawing.Size(25, 21)
        Me.cmdBrowseDEMFile.TabIndex = 2
        Me.cmdBrowseDEMFile.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdBrowseDEMFile.UseVisualStyleBackColor = False
        '
        'txtWSDelinName
        '
        Me.txtWSDelinName.AcceptsReturn = True
        Me.txtWSDelinName.BackColor = System.Drawing.SystemColors.Window
        Me.txtWSDelinName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtWSDelinName.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtWSDelinName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtWSDelinName.Location = New System.Drawing.Point(170, 24)
        Me.txtWSDelinName.MaxLength = 0
        Me.txtWSDelinName.Name = "txtWSDelinName"
        Me.txtWSDelinName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtWSDelinName.Size = New System.Drawing.Size(346, 20)
        Me.txtWSDelinName.TabIndex = 0
        '
        'txtDEMFile
        '
        Me.txtDEMFile.AcceptsReturn = True
        Me.txtDEMFile.BackColor = System.Drawing.SystemColors.Window
        Me.txtDEMFile.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDEMFile.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDEMFile.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDEMFile.Location = New System.Drawing.Point(170, 52)
        Me.txtDEMFile.MaxLength = 0
        Me.txtDEMFile.Name = "txtDEMFile"
        Me.txtDEMFile.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDEMFile.Size = New System.Drawing.Size(346, 20)
        Me.txtDEMFile.TabIndex = 1
        '
        'txtLS
        '
        Me.txtLS.AcceptsReturn = True
        Me.txtLS.BackColor = System.Drawing.SystemColors.Window
        Me.txtLS.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtLS.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLS.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtLS.Location = New System.Drawing.Point(170, 157)
        Me.txtLS.MaxLength = 0
        Me.txtLS.Name = "txtLS"
        Me.txtLS.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtLS.Size = New System.Drawing.Size(346, 20)
        Me.txtLS.TabIndex = 4
        '
        '_Label1_2
        '
        Me._Label1_2.AutoSize = True
        Me._Label1_2.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_2.Location = New System.Drawing.Point(96, 80)
        Me._Label1_2.Name = "_Label1_2"
        Me._Label1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_2.Size = New System.Drawing.Size(58, 14)
        Me._Label1_2.TabIndex = 21
        Me._Label1_2.Text = "DEM Units:"
        '
        '_Label1_1
        '
        Me._Label1_1.AutoSize = True
        Me._Label1_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_1.Location = New System.Drawing.Point(94, 184)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_1.Size = New System.Drawing.Size(69, 14)
        Me._Label1_1.TabIndex = 16
        Me._Label1_1.Text = "Watersheds:"
        Me._Label1_1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_5
        '
        Me._Label1_5.AutoSize = True
        Me._Label1_5.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_5.Location = New System.Drawing.Point(68, 159)
        Me._Label1_5.Name = "_Label1_5"
        Me._Label1_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_5.Size = New System.Drawing.Size(96, 14)
        Me._Label1_5.TabIndex = 9
        Me._Label1_5.Text = "Length-slope Grid:"
        Me._Label1_5.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_4
        '
        Me._Label1_4.AutoSize = True
        Me._Label1_4.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_4.Location = New System.Drawing.Point(40, 132)
        Me._Label1_4.Name = "_Label1_4"
        Me._Label1_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_4.Size = New System.Drawing.Size(124, 14)
        Me._Label1_4.TabIndex = 8
        Me._Label1_4.Text = "Flow Accumulation Grid:"
        Me._Label1_4.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_3
        '
        Me._Label1_3.AutoSize = True
        Me._Label1_3.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_3.Location = New System.Drawing.Point(17, 27)
        Me._Label1_3.Name = "_Label1_3"
        Me._Label1_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_3.Size = New System.Drawing.Size(148, 14)
        Me._Label1_3.TabIndex = 7
        Me._Label1_3.Text = "Watershed Delineation Name:"
        Me._Label1_3.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_0
        '
        Me._Label1_0.AutoSize = True
        Me._Label1_0.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_0.Location = New System.Drawing.Point(62, 107)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_0.Size = New System.Drawing.Size(102, 14)
        Me._Label1_0.TabIndex = 6
        Me._Label1_0.Text = "Flow Direction Grid:"
        Me._Label1_0.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_12
        '
        Me._Label1_12.AutoSize = True
        Me._Label1_12.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_12.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_12.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_12.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_12.Location = New System.Drawing.Point(104, 56)
        Me._Label1_12.Name = "_Label1_12"
        Me._Label1_12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_12.Size = New System.Drawing.Size(54, 14)
        Me._Label1_12.TabIndex = 5
        Me._Label1_12.Text = "DEM Grid:"
        Me._Label1_12.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'frmUserWShed
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(584, 362)
        Me.Controls.Add(Me.cmdQuit)
        Me.Controls.Add(Me.cmdCreate)
        Me.Controls.Add(Me.Frame1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(4, 30)
        Me.Name = "frmUserWShed"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "New Watershed Delineation"
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
#End Region 
End Class