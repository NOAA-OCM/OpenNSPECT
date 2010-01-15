<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmNewPrecip
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
	Public WithEvents txtRainingDays As System.Windows.Forms.TextBox
	Public WithEvents cboTimePeriod As System.Windows.Forms.ComboBox
	Public WithEvents cboPrecipType As System.Windows.Forms.ComboBox
	Public WithEvents txtPrecipName As System.Windows.Forms.TextBox
	Public WithEvents txtDesc As System.Windows.Forms.TextBox
	Public WithEvents txtPrecipFile As System.Windows.Forms.TextBox
	Public WithEvents cmdBrowseFile As System.Windows.Forms.Button
	Public WithEvents cboPrecipUnits As System.Windows.Forms.ComboBox
	Public WithEvents cboGridUnits As System.Windows.Forms.ComboBox
	Public WithEvents lblRainingDays As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents _Label1_7 As System.Windows.Forms.Label
	Public WithEvents _Label1_1 As System.Windows.Forms.Label
	Public WithEvents _Label1_0 As System.Windows.Forms.Label
	Public WithEvents _Label1_2 As System.Windows.Forms.Label
	Public WithEvents _Label1_6 As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmNewPrecip))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Frame1 = New System.Windows.Forms.GroupBox
        Me.txtRainingDays = New System.Windows.Forms.TextBox
        Me.cboTimePeriod = New System.Windows.Forms.ComboBox
        Me.cboPrecipType = New System.Windows.Forms.ComboBox
        Me.txtPrecipName = New System.Windows.Forms.TextBox
        Me.txtDesc = New System.Windows.Forms.TextBox
        Me.txtPrecipFile = New System.Windows.Forms.TextBox
        Me.cmdBrowseFile = New System.Windows.Forms.Button
        Me.cboPrecipUnits = New System.Windows.Forms.ComboBox
        Me.cboGridUnits = New System.Windows.Forms.ComboBox
        Me.lblRainingDays = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me._Label1_7 = New System.Windows.Forms.Label
        Me._Label1_1 = New System.Windows.Forms.Label
        Me._Label1_0 = New System.Windows.Forms.Label
        Me._Label1_2 = New System.Windows.Forms.Label
        Me._Label1_6 = New System.Windows.Forms.Label
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmdOK = New System.Windows.Forms.Button
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.txtRainingDays)
        Me.Frame1.Controls.Add(Me.cboTimePeriod)
        Me.Frame1.Controls.Add(Me.cboPrecipType)
        Me.Frame1.Controls.Add(Me.txtPrecipName)
        Me.Frame1.Controls.Add(Me.txtDesc)
        Me.Frame1.Controls.Add(Me.txtPrecipFile)
        Me.Frame1.Controls.Add(Me.cmdBrowseFile)
        Me.Frame1.Controls.Add(Me.cboPrecipUnits)
        Me.Frame1.Controls.Add(Me.cboGridUnits)
        Me.Frame1.Controls.Add(Me.lblRainingDays)
        Me.Frame1.Controls.Add(Me.Label2)
        Me.Frame1.Controls.Add(Me.Label3)
        Me.Frame1.Controls.Add(Me._Label1_7)
        Me.Frame1.Controls.Add(Me._Label1_1)
        Me.Frame1.Controls.Add(Me._Label1_0)
        Me.Frame1.Controls.Add(Me._Label1_2)
        Me.Frame1.Controls.Add(Me._Label1_6)
        Me.Frame1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(8, 7)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(461, 240)
        Me.Frame1.TabIndex = 8
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Enter new scenario information  "
        '
        'txtRainingDays
        '
        Me.txtRainingDays.AcceptsReturn = True
        Me.txtRainingDays.BackColor = System.Drawing.SystemColors.Window
        Me.txtRainingDays.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRainingDays.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRainingDays.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRainingDays.Location = New System.Drawing.Point(337, 178)
        Me.txtRainingDays.MaxLength = 0
        Me.txtRainingDays.Name = "txtRainingDays"
        Me.txtRainingDays.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRainingDays.Size = New System.Drawing.Size(46, 20)
        Me.txtRainingDays.TabIndex = 19
        Me.txtRainingDays.Visible = False
        '
        'cboTimePeriod
        '
        Me.cboTimePeriod.BackColor = System.Drawing.SystemColors.Window
        Me.cboTimePeriod.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboTimePeriod.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTimePeriod.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboTimePeriod.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboTimePeriod.Items.AddRange(New Object() {"Annual", "Event"})
        Me.cboTimePeriod.Location = New System.Drawing.Point(112, 177)
        Me.cboTimePeriod.Name = "cboTimePeriod"
        Me.cboTimePeriod.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboTimePeriod.Size = New System.Drawing.Size(143, 22)
        Me.cboTimePeriod.TabIndex = 16
        '
        'cboPrecipType
        '
        Me.cboPrecipType.BackColor = System.Drawing.SystemColors.Window
        Me.cboPrecipType.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboPrecipType.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboPrecipType.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboPrecipType.Items.AddRange(New Object() {"Type I", "Type IA", "Type II", "Type III"})
        Me.cboPrecipType.Location = New System.Drawing.Point(112, 209)
        Me.cboPrecipType.Name = "cboPrecipType"
        Me.cboPrecipType.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboPrecipType.Size = New System.Drawing.Size(143, 22)
        Me.cboPrecipType.TabIndex = 14
        '
        'txtPrecipName
        '
        Me.txtPrecipName.AcceptsReturn = True
        Me.txtPrecipName.BackColor = System.Drawing.SystemColors.Window
        Me.txtPrecipName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPrecipName.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPrecipName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPrecipName.Location = New System.Drawing.Point(112, 25)
        Me.txtPrecipName.MaxLength = 0
        Me.txtPrecipName.Name = "txtPrecipName"
        Me.txtPrecipName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPrecipName.Size = New System.Drawing.Size(109, 21)
        Me.txtPrecipName.TabIndex = 0
        '
        'txtDesc
        '
        Me.txtDesc.AcceptsReturn = True
        Me.txtDesc.BackColor = System.Drawing.SystemColors.Window
        Me.txtDesc.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDesc.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDesc.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDesc.Location = New System.Drawing.Point(112, 54)
        Me.txtDesc.MaxLength = 0
        Me.txtDesc.Name = "txtDesc"
        Me.txtDesc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDesc.Size = New System.Drawing.Size(269, 21)
        Me.txtDesc.TabIndex = 1
        '
        'txtPrecipFile
        '
        Me.txtPrecipFile.AcceptsReturn = True
        Me.txtPrecipFile.BackColor = System.Drawing.SystemColors.Window
        Me.txtPrecipFile.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPrecipFile.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPrecipFile.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPrecipFile.Location = New System.Drawing.Point(112, 82)
        Me.txtPrecipFile.MaxLength = 0
        Me.txtPrecipFile.Name = "txtPrecipFile"
        Me.txtPrecipFile.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPrecipFile.Size = New System.Drawing.Size(269, 21)
        Me.txtPrecipFile.TabIndex = 3
        '
        'cmdBrowseFile
        '
        Me.cmdBrowseFile.BackColor = System.Drawing.SystemColors.Control
        Me.cmdBrowseFile.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdBrowseFile.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowseFile.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdBrowseFile.Image = CType(resources.GetObject("cmdBrowseFile.Image"), System.Drawing.Image)
        Me.cmdBrowseFile.Location = New System.Drawing.Point(385, 81)
        Me.cmdBrowseFile.Name = "cmdBrowseFile"
        Me.cmdBrowseFile.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdBrowseFile.Size = New System.Drawing.Size(25, 21)
        Me.cmdBrowseFile.TabIndex = 2
        Me.cmdBrowseFile.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdBrowseFile.UseVisualStyleBackColor = False
        '
        'cboPrecipUnits
        '
        Me.cboPrecipUnits.BackColor = System.Drawing.SystemColors.Window
        Me.cboPrecipUnits.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboPrecipUnits.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPrecipUnits.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboPrecipUnits.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboPrecipUnits.Items.AddRange(New Object() {"centimeters", "inches"})
        Me.cboPrecipUnits.Location = New System.Drawing.Point(112, 144)
        Me.cboPrecipUnits.Name = "cboPrecipUnits"
        Me.cboPrecipUnits.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboPrecipUnits.Size = New System.Drawing.Size(143, 22)
        Me.cboPrecipUnits.TabIndex = 5
        '
        'cboGridUnits
        '
        Me.cboGridUnits.BackColor = System.Drawing.SystemColors.Window
        Me.cboGridUnits.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboGridUnits.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboGridUnits.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboGridUnits.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboGridUnits.Items.AddRange(New Object() {"meters", "feet"})
        Me.cboGridUnits.Location = New System.Drawing.Point(112, 112)
        Me.cboGridUnits.Name = "cboGridUnits"
        Me.cboGridUnits.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboGridUnits.Size = New System.Drawing.Size(143, 22)
        Me.cboGridUnits.TabIndex = 4
        '
        'lblRainingDays
        '
        Me.lblRainingDays.BackColor = System.Drawing.SystemColors.Control
        Me.lblRainingDays.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblRainingDays.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRainingDays.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRainingDays.Location = New System.Drawing.Point(265, 180)
        Me.lblRainingDays.Name = "lblRainingDays"
        Me.lblRainingDays.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblRainingDays.Size = New System.Drawing.Size(72, 21)
        Me.lblRainingDays.TabIndex = 18
        Me.lblRainingDays.Text = "Raining Days: "
        Me.lblRainingDays.Visible = False
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(21, 179)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(81, 17)
        Me.Label2.TabIndex = 17
        Me.Label2.Text = "Time Period:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(21, 210)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(81, 17)
        Me.Label3.TabIndex = 15
        Me.Label3.Text = "Type:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_7
        '
        Me._Label1_7.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_7.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_7.Location = New System.Drawing.Point(23, 25)
        Me._Label1_7.Name = "_Label1_7"
        Me._Label1_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_7.Size = New System.Drawing.Size(79, 17)
        Me._Label1_7.TabIndex = 13
        Me._Label1_7.Text = "Scenario Name:"
        Me._Label1_7.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_1
        '
        Me._Label1_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_1.Location = New System.Drawing.Point(5, 54)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_1.Size = New System.Drawing.Size(97, 19)
        Me._Label1_1.TabIndex = 12
        Me._Label1_1.Text = "Description:"
        Me._Label1_1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_0
        '
        Me._Label1_0.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_0.Location = New System.Drawing.Point(5, 83)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_0.Size = New System.Drawing.Size(97, 19)
        Me._Label1_0.TabIndex = 11
        Me._Label1_0.Text = "Precipitation Grid:"
        Me._Label1_0.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_2
        '
        Me._Label1_2.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_2.Location = New System.Drawing.Point(5, 143)
        Me._Label1_2.Name = "_Label1_2"
        Me._Label1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_2.Size = New System.Drawing.Size(97, 19)
        Me._Label1_2.TabIndex = 10
        Me._Label1_2.Text = "Precipitation Units:"
        Me._Label1_2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_6
        '
        Me._Label1_6.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_6.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_6.Location = New System.Drawing.Point(5, 112)
        Me._Label1_6.Name = "_Label1_6"
        Me._Label1_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_6.Size = New System.Drawing.Size(97, 19)
        Me._Label1_6.TabIndex = 9
        Me._Label1_6.Text = "Grid Units:"
        Me._Label1_6.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.Location = New System.Drawing.Point(396, 257)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancel.Size = New System.Drawing.Size(65, 25)
        Me.cmdCancel.TabIndex = 7
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'cmdOK
        '
        Me.cmdOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOK.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOK.Location = New System.Drawing.Point(326, 257)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOK.Size = New System.Drawing.Size(65, 25)
        Me.cmdOK.TabIndex = 6
        Me.cmdOK.Text = "OK"
        Me.cmdOK.UseVisualStyleBackColor = False
        '
        'frmNewPrecip
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(477, 290)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdOK)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(3, 21)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmNewPrecip"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "New Precipitation Scenario"
        Me.Frame1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
#End Region 
End Class