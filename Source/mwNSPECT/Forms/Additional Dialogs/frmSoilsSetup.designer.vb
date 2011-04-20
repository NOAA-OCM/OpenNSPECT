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
    Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label5 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSoilsSetup))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtSoilsName = New System.Windows.Forms.TextBox()
        Me.txtDEMFile = New System.Windows.Forms.TextBox()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.txtMUSLEVal = New System.Windows.Forms.TextBox()
        Me.txtMUSLEExp = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.cmdDEMBrowse = New System.Windows.Forms.Button()
        Me.cmdQuit = New System.Windows.Forms.Button()
        Me.cmdSave = New System.Windows.Forms.Button()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.cboSoilFieldsK = New System.Windows.Forms.ComboBox()
        Me.cboSoilFields = New System.Windows.Forms.ComboBox()
        Me.cmdBrowseFile = New System.Windows.Forms.Button()
        Me.txtSoilsDS = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Frame2.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtSoilsName
        '
        Me.txtSoilsName.AcceptsReturn = True
        Me.txtSoilsName.BackColor = System.Drawing.SystemColors.Window
        Me.txtSoilsName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSoilsName.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSoilsName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSoilsName.Location = New System.Drawing.Point(89, 6)
        Me.txtSoilsName.MaxLength = 0
        Me.txtSoilsName.Name = "txtSoilsName"
        Me.txtSoilsName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSoilsName.Size = New System.Drawing.Size(212, 20)
        Me.txtSoilsName.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.txtSoilsName, "Choose the filled DEM you are using.  This will provide Spatial Analyst with the " & _
        "proper analysis environment.")
        '
        'txtDEMFile
        '
        Me.txtDEMFile.AcceptsReturn = True
        Me.txtDEMFile.BackColor = System.Drawing.SystemColors.Window
        Me.txtDEMFile.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDEMFile.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDEMFile.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDEMFile.Location = New System.Drawing.Point(89, 38)
        Me.txtDEMFile.MaxLength = 0
        Me.txtDEMFile.Name = "txtDEMFile"
        Me.txtDEMFile.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDEMFile.Size = New System.Drawing.Size(212, 20)
        Me.txtDEMFile.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.txtDEMFile, "Choose the filled DEM you are using.  This will provide Spatial Analyst with the " & _
        "proper analysis environment.")
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.txtMUSLEVal)
        Me.Frame2.Controls.Add(Me.txtMUSLEExp)
        Me.Frame2.Controls.Add(Me.Label12)
        Me.Frame2.Controls.Add(Me.Label7)
        Me.Frame2.Controls.Add(Me.Label6)
        Me.Frame2.Controls.Add(Me.Label11)
        Me.Frame2.Controls.Add(Me.Label10)
        Me.Frame2.Controls.Add(Me.Label9)
        Me.Frame2.Controls.Add(Me.Label8)
        Me.Frame2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame2.Location = New System.Drawing.Point(16, 192)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(328, 157)
        Me.Frame2.TabIndex = 15
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Advanced MUSLE Specific Coefficients"
        '
        'txtMUSLEVal
        '
        Me.txtMUSLEVal.AcceptsReturn = True
        Me.txtMUSLEVal.BackColor = System.Drawing.SystemColors.Window
        Me.txtMUSLEVal.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMUSLEVal.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMUSLEVal.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMUSLEVal.Location = New System.Drawing.Point(56, 56)
        Me.txtMUSLEVal.MaxLength = 0
        Me.txtMUSLEVal.Name = "txtMUSLEVal"
        Me.txtMUSLEVal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMUSLEVal.Size = New System.Drawing.Size(43, 20)
        Me.txtMUSLEVal.TabIndex = 17
        Me.txtMUSLEVal.Text = "95"
        '
        'txtMUSLEExp
        '
        Me.txtMUSLEExp.AcceptsReturn = True
        Me.txtMUSLEExp.BackColor = System.Drawing.SystemColors.Window
        Me.txtMUSLEExp.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMUSLEExp.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMUSLEExp.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMUSLEExp.Location = New System.Drawing.Point(136, 56)
        Me.txtMUSLEExp.MaxLength = 0
        Me.txtMUSLEExp.Name = "txtMUSLEExp"
        Me.txtMUSLEExp.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMUSLEExp.Size = New System.Drawing.Size(35, 20)
        Me.txtMUSLEExp.TabIndex = 16
        Me.txtMUSLEExp.Text = "0.56"
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.Color.Transparent
        Me.Label12.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label12.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label12.Location = New System.Drawing.Point(16, 104)
        Me.Label12.Name = "Label12"
        Me.Label12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label12.Size = New System.Drawing.Size(305, 57)
        Me.Label12.TabIndex = 24
        Me.Label12.Text = "Warning: Q and qp are calculated in English units (acre-feet and cubic feet per s" & _
    "econd respectively). ""a"" and ""b"" must be derived accordingly."
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(112, 64)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(17, 17)
        Me.Label7.TabIndex = 23
        Me.Label7.Text = "b="
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(32, 64)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(17, 17)
        Me.Label6.TabIndex = 22
        Me.Label6.Text = "a="
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.Color.Transparent
        Me.Label11.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label11.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label11.Location = New System.Drawing.Point(72, 32)
        Me.Label11.Name = "Label11"
        Me.Label11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label11.Size = New System.Drawing.Size(25, 13)
        Me.Label11.TabIndex = 21
        Me.Label11.Text = "    b"
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.SystemColors.Control
        Me.Label10.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label10.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label10.Location = New System.Drawing.Point(18, 19)
        Me.Label10.Name = "Label10"
        Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label10.Size = New System.Drawing.Size(252, 16)
        Me.Label10.TabIndex = 20
        Me.Label10.Text = "MUSLE Equation for sediment yield:"
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.SystemColors.Control
        Me.Label9.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label9.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label9.Location = New System.Drawing.Point(32, 38)
        Me.Label9.Name = "Label9"
        Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label9.Size = New System.Drawing.Size(212, 20)
        Me.Label9.TabIndex = 19
        Me.Label9.Text = "a * (Q * qp)   * K * C * P * LS"
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.SystemColors.Control
        Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label8.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label8.Location = New System.Drawing.Point(16, 88)
        Me.Label8.Name = "Label8"
        Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label8.Size = New System.Drawing.Size(303, 16)
        Me.Label8.TabIndex = 18
        Me.Label8.Text = "Locally calibrated MUSLE coefficients can be entered above."
        '
        'cmdDEMBrowse
        '
        Me.cmdDEMBrowse.BackColor = System.Drawing.SystemColors.Control
        Me.cmdDEMBrowse.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdDEMBrowse.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDEMBrowse.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDEMBrowse.Image = CType(resources.GetObject("cmdDEMBrowse.Image"), System.Drawing.Image)
        Me.cmdDEMBrowse.Location = New System.Drawing.Point(307, 39)
        Me.cmdDEMBrowse.Name = "cmdDEMBrowse"
        Me.cmdDEMBrowse.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdDEMBrowse.Size = New System.Drawing.Size(25, 21)
        Me.cmdDEMBrowse.TabIndex = 2
        Me.cmdDEMBrowse.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdDEMBrowse.UseVisualStyleBackColor = False
        '
        'cmdQuit
        '
        Me.cmdQuit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdQuit.BackColor = System.Drawing.SystemColors.Control
        Me.cmdQuit.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdQuit.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdQuit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdQuit.Location = New System.Drawing.Point(508, 338)
        Me.cmdQuit.Name = "cmdQuit"
        Me.cmdQuit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdQuit.Size = New System.Drawing.Size(65, 25)
        Me.cmdQuit.TabIndex = 8
        Me.cmdQuit.Text = "Cancel"
        Me.cmdQuit.UseVisualStyleBackColor = False
        '
        'cmdSave
        '
        Me.cmdSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdSave.BackColor = System.Drawing.SystemColors.Control
        Me.cmdSave.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdSave.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSave.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSave.Location = New System.Drawing.Point(438, 338)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdSave.Size = New System.Drawing.Size(65, 25)
        Me.cmdSave.TabIndex = 7
        Me.cmdSave.Text = "OK"
        Me.cmdSave.UseVisualStyleBackColor = False
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.cboSoilFieldsK)
        Me.Frame1.Controls.Add(Me.cboSoilFields)
        Me.Frame1.Controls.Add(Me.cmdBrowseFile)
        Me.Frame1.Controls.Add(Me.txtSoilsDS)
        Me.Frame1.Controls.Add(Me.Label2)
        Me.Frame1.Controls.Add(Me.Label3)
        Me.Frame1.Controls.Add(Me.Label1)
        Me.Frame1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(16, 69)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(328, 120)
        Me.Frame1.TabIndex = 9
        Me.Frame1.TabStop = False
        Me.Frame1.Text = " Soils  "
        '
        'cboSoilFieldsK
        '
        Me.cboSoilFieldsK.BackColor = System.Drawing.SystemColors.Window
        Me.cboSoilFieldsK.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboSoilFieldsK.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboSoilFieldsK.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboSoilFieldsK.Location = New System.Drawing.Point(173, 83)
        Me.cboSoilFieldsK.Name = "cboSoilFieldsK"
        Me.cboSoilFieldsK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboSoilFieldsK.Size = New System.Drawing.Size(143, 22)
        Me.cboSoilFieldsK.TabIndex = 6
        '
        'cboSoilFields
        '
        Me.cboSoilFields.BackColor = System.Drawing.SystemColors.Window
        Me.cboSoilFields.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboSoilFields.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboSoilFields.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboSoilFields.Location = New System.Drawing.Point(173, 49)
        Me.cboSoilFields.Name = "cboSoilFields"
        Me.cboSoilFields.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboSoilFields.Size = New System.Drawing.Size(143, 22)
        Me.cboSoilFields.TabIndex = 5
        '
        'cmdBrowseFile
        '
        Me.cmdBrowseFile.BackColor = System.Drawing.SystemColors.Control
        Me.cmdBrowseFile.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdBrowseFile.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowseFile.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdBrowseFile.Image = CType(resources.GetObject("cmdBrowseFile.Image"), System.Drawing.Image)
        Me.cmdBrowseFile.Location = New System.Drawing.Point(291, 18)
        Me.cmdBrowseFile.Name = "cmdBrowseFile"
        Me.cmdBrowseFile.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdBrowseFile.Size = New System.Drawing.Size(25, 21)
        Me.cmdBrowseFile.TabIndex = 4
        Me.cmdBrowseFile.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdBrowseFile.UseVisualStyleBackColor = False
        '
        'txtSoilsDS
        '
        Me.txtSoilsDS.AcceptsReturn = True
        Me.txtSoilsDS.BackColor = System.Drawing.SystemColors.Window
        Me.txtSoilsDS.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSoilsDS.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSoilsDS.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSoilsDS.Location = New System.Drawing.Point(93, 19)
        Me.txtSoilsDS.MaxLength = 0
        Me.txtSoilsDS.Name = "txtSoilsDS"
        Me.txtSoilsDS.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSoilsDS.Size = New System.Drawing.Size(193, 20)
        Me.txtSoilsDS.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(13, 85)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(122, 19)
        Me.Label2.TabIndex = 13
        Me.Label2.Text = "K Factor Attribute:"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(11, 52)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(159, 30)
        Me.Label3.TabIndex = 11
        Me.Label3.Text = "Hydrologic Soil Group Attribute:"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(12, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(82, 18)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Soils Data Set:"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(18, 9)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(59, 18)
        Me.Label4.TabIndex = 14
        Me.Label4.Text = "Name:"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(17, 41)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(77, 19)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "DEM GRID:"
        '
        'frmSoilsSetup
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(594, 372)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.txtSoilsName)
        Me.Controls.Add(Me.cmdDEMBrowse)
        Me.Controls.Add(Me.txtDEMFile)
        Me.Controls.Add(Me.cmdQuit)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label5)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(3, 22)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmSoilsSetup"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Soils Setup"
        Me.Frame2.ResumeLayout(False)
        Me.Frame2.PerformLayout()
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region 
End Class