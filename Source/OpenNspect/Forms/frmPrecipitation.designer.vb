<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmPrecipitation
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
    Public WithEvents mnuNewPrecip As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuDelPrecip As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuPrecipOpts As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuPrecipHelp As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuHelp As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
    Public WithEvents cboScenName As System.Windows.Forms.ComboBox
    Public WithEvents txtRainingDays As System.Windows.Forms.TextBox
    Public WithEvents cboTimePeriod As System.Windows.Forms.ComboBox
    Public WithEvents cboPrecipType As System.Windows.Forms.ComboBox
    Public WithEvents cmdBrowseFile As System.Windows.Forms.Button
    Public WithEvents txtDesc As System.Windows.Forms.TextBox
    Public WithEvents txtPrecipFile As System.Windows.Forms.TextBox
    Public WithEvents cboGridUnits As System.Windows.Forms.ComboBox
    Public WithEvents cboPrecipUnits As System.Windows.Forms.ComboBox
    Public WithEvents _Label1_0 As System.Windows.Forms.Label
    Public WithEvents lblRainingDays As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents _Label1_2 As System.Windows.Forms.Label
    Public WithEvents _Label1_1 As System.Windows.Forms.Label
    Public WithEvents _Label1_6 As System.Windows.Forms.Label
    Public WithEvents _Label1_7 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents cmdSave As System.Windows.Forms.Button
    Public WithEvents cmdQuit As System.Windows.Forms.Button
    'Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPrecipitation))
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip
        Me.mnuPrecipOpts = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuNewPrecip = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuDelPrecip = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuHelp = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuPrecipHelp = New System.Windows.Forms.ToolStripMenuItem
        Me.Frame1 = New System.Windows.Forms.GroupBox
        Me.cboScenName = New System.Windows.Forms.ComboBox
        Me.txtRainingDays = New System.Windows.Forms.TextBox
        Me.cboTimePeriod = New System.Windows.Forms.ComboBox
        Me.cboPrecipType = New System.Windows.Forms.ComboBox
        Me.cmdBrowseFile = New System.Windows.Forms.Button
        Me.txtDesc = New System.Windows.Forms.TextBox
        Me.txtPrecipFile = New System.Windows.Forms.TextBox
        Me.cboGridUnits = New System.Windows.Forms.ComboBox
        Me.cboPrecipUnits = New System.Windows.Forms.ComboBox
        Me._Label1_0 = New System.Windows.Forms.Label
        Me.lblRainingDays = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me._Label1_2 = New System.Windows.Forms.Label
        Me._Label1_1 = New System.Windows.Forms.Label
        Me._Label1_6 = New System.Windows.Forms.Label
        Me._Label1_7 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.cmdSave = New System.Windows.Forms.Button
        Me.cmdQuit = New System.Windows.Forms.Button
        'Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
        Me.MainMenu1.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        Me.ToolTip1.Active = True
        'CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Text = "Precipitation Scenarios"
        Me.ClientSize = New System.Drawing.Size(477, 302)
        Me.Location = New System.Drawing.Point(192, 246)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ControlBox = True
        Me.Enabled = True
        Me.KeyPreview = False
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.HelpButton = False
        Me.WindowState = System.Windows.Forms.FormWindowState.Normal
        Me.Name = "frmPrecip"
        Me.mnuPrecipOpts.Name = "mnuPrecipOpts"
        Me.mnuPrecipOpts.Text = "Options"
        Me.mnuPrecipOpts.Checked = False
        Me.mnuPrecipOpts.Enabled = True
        Me.mnuPrecipOpts.Visible = True
        Me.mnuNewPrecip.Name = "mnuNewPrecip"
        Me.mnuNewPrecip.Text = "New..."
        Me.mnuNewPrecip.Checked = False
        Me.mnuNewPrecip.Enabled = True
        Me.mnuNewPrecip.Visible = True
        Me.mnuDelPrecip.Name = "mnuDelPrecip"
        Me.mnuDelPrecip.Text = "Delete..."
        Me.mnuDelPrecip.Checked = False
        Me.mnuDelPrecip.Enabled = True
        Me.mnuDelPrecip.Visible = True
        Me.mnuHelp.Name = "mnuHelp"
        Me.mnuHelp.Text = "&Help"
        Me.mnuHelp.Checked = False
        Me.mnuHelp.Enabled = True
        Me.mnuHelp.Visible = True
        Me.mnuPrecipHelp.Name = "mnuPrecipHelp"
        Me.mnuPrecipHelp.Text = "Precipitation Scenarios..."
        Me.mnuPrecipHelp.ShortcutKeys = CType(System.Windows.Forms.Keys.Shift Or System.Windows.Forms.Keys.F1, System.Windows.Forms.Keys)
        Me.mnuPrecipHelp.Checked = False
        Me.mnuPrecipHelp.Enabled = True
        Me.mnuPrecipHelp.Visible = True
        Me.Frame1.Text = "Choose a precipitation scenario to view or edit  "
        Me.Frame1.Size = New System.Drawing.Size(455, 236)
        Me.Frame1.Location = New System.Drawing.Point(12, 26)
        Me.Frame1.TabIndex = 6
        Me.Frame1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Enabled = True
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Visible = True
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.Name = "Frame1"
        Me.cboScenName.Size = New System.Drawing.Size(143, 21)
        Me.cboScenName.Location = New System.Drawing.Point(139, 24)
        Me.cboScenName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboScenName.TabIndex = 18
        Me.cboScenName.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboScenName.BackColor = System.Drawing.SystemColors.Window
        Me.cboScenName.CausesValidation = True
        Me.cboScenName.Enabled = True
        Me.cboScenName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboScenName.IntegralHeight = True
        Me.cboScenName.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboScenName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboScenName.Sorted = False
        Me.cboScenName.TabStop = True
        Me.cboScenName.Visible = True
        Me.cboScenName.Name = "cboScenName"
        Me.txtRainingDays.AutoSize = False
        Me.txtRainingDays.Size = New System.Drawing.Size(46, 20)
        Me.txtRainingDays.Location = New System.Drawing.Point(366, 173)
        Me.txtRainingDays.TabIndex = 17
        Me.txtRainingDays.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRainingDays.AcceptsReturn = True
        Me.txtRainingDays.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
        Me.txtRainingDays.BackColor = System.Drawing.SystemColors.Window
        Me.txtRainingDays.CausesValidation = True
        Me.txtRainingDays.Enabled = True
        Me.txtRainingDays.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRainingDays.HideSelection = True
        Me.txtRainingDays.ReadOnly = False
        Me.txtRainingDays.Maxlength = 0
        Me.txtRainingDays.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRainingDays.MultiLine = False
        Me.txtRainingDays.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRainingDays.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.txtRainingDays.TabStop = True
        Me.txtRainingDays.Visible = True
        Me.txtRainingDays.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtRainingDays.Name = "txtRainingDays"
        Me.cboTimePeriod.Size = New System.Drawing.Size(143, 21)
        Me.cboTimePeriod.Location = New System.Drawing.Point(139, 173)
        Me.cboTimePeriod.Items.AddRange(New Object() {"Annual", "Event"})
        Me.cboTimePeriod.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTimePeriod.TabIndex = 15
        Me.cboTimePeriod.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboTimePeriod.BackColor = System.Drawing.SystemColors.Window
        Me.cboTimePeriod.CausesValidation = True
        Me.cboTimePeriod.Enabled = True
        Me.cboTimePeriod.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboTimePeriod.IntegralHeight = True
        Me.cboTimePeriod.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboTimePeriod.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboTimePeriod.Sorted = False
        Me.cboTimePeriod.TabStop = True
        Me.cboTimePeriod.Visible = True
        Me.cboTimePeriod.Name = "cboTimePeriod"
        Me.cboPrecipType.Size = New System.Drawing.Size(143, 21)
        Me.cboPrecipType.Location = New System.Drawing.Point(139, 202)
        Me.cboPrecipType.Items.AddRange(New Object() {"Type I", "Type IA", "Type II", "Type III"})
        Me.cboPrecipType.TabIndex = 8
        Me.cboPrecipType.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboPrecipType.BackColor = System.Drawing.SystemColors.Window
        Me.cboPrecipType.CausesValidation = True
        Me.cboPrecipType.Enabled = True
        Me.cboPrecipType.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboPrecipType.IntegralHeight = True
        Me.cboPrecipType.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboPrecipType.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboPrecipType.Sorted = False
        Me.cboPrecipType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
        Me.cboPrecipType.TabStop = True
        Me.cboPrecipType.Visible = True
        Me.cboPrecipType.Name = "cboPrecipType"
        Me.cmdBrowseFile.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdBrowseFile.Size = New System.Drawing.Size(25, 21)
        Me.cmdBrowseFile.Location = New System.Drawing.Point(398, 82)
        Me.cmdBrowseFile.Image = CType(resources.GetObject("cmdBrowseFile.Image"), System.Drawing.Image)
        Me.cmdBrowseFile.TabIndex = 1
        Me.cmdBrowseFile.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowseFile.BackColor = System.Drawing.SystemColors.Control
        Me.cmdBrowseFile.CausesValidation = True
        Me.cmdBrowseFile.Enabled = True
        Me.cmdBrowseFile.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdBrowseFile.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdBrowseFile.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdBrowseFile.TabStop = True
        Me.cmdBrowseFile.Name = "cmdBrowseFile"
        Me.txtDesc.AutoSize = False
        Me.txtDesc.Size = New System.Drawing.Size(255, 21)
        Me.txtDesc.Location = New System.Drawing.Point(139, 54)
        Me.txtDesc.TabIndex = 0
        Me.txtDesc.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDesc.AcceptsReturn = True
        Me.txtDesc.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
        Me.txtDesc.BackColor = System.Drawing.SystemColors.Window
        Me.txtDesc.CausesValidation = True
        Me.txtDesc.Enabled = True
        Me.txtDesc.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDesc.HideSelection = True
        Me.txtDesc.ReadOnly = False
        Me.txtDesc.Maxlength = 0
        Me.txtDesc.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDesc.MultiLine = False
        Me.txtDesc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDesc.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.txtDesc.TabStop = True
        Me.txtDesc.Visible = True
        Me.txtDesc.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtDesc.Name = "txtDesc"
        Me.txtPrecipFile.AutoSize = False
        Me.txtPrecipFile.Size = New System.Drawing.Size(255, 21)
        Me.txtPrecipFile.Location = New System.Drawing.Point(139, 84)
        Me.txtPrecipFile.TabIndex = 2
        Me.txtPrecipFile.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPrecipFile.AcceptsReturn = True
        Me.txtPrecipFile.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
        Me.txtPrecipFile.BackColor = System.Drawing.SystemColors.Window
        Me.txtPrecipFile.CausesValidation = True
        Me.txtPrecipFile.Enabled = True
        Me.txtPrecipFile.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPrecipFile.HideSelection = True
        Me.txtPrecipFile.ReadOnly = False
        Me.txtPrecipFile.Maxlength = 0
        Me.txtPrecipFile.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPrecipFile.MultiLine = False
        Me.txtPrecipFile.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPrecipFile.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.txtPrecipFile.TabStop = True
        Me.txtPrecipFile.Visible = True
        Me.txtPrecipFile.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtPrecipFile.Name = "txtPrecipFile"
        Me.cboGridUnits.Size = New System.Drawing.Size(142, 21)
        Me.cboGridUnits.Location = New System.Drawing.Point(139, 113)
        Me.cboGridUnits.Items.AddRange(New Object() {"meters", "feet"})
        Me.cboGridUnits.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboGridUnits.TabIndex = 3
        Me.cboGridUnits.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboGridUnits.BackColor = System.Drawing.SystemColors.Window
        Me.cboGridUnits.CausesValidation = True
        Me.cboGridUnits.Enabled = True
        Me.cboGridUnits.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboGridUnits.IntegralHeight = True
        Me.cboGridUnits.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboGridUnits.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboGridUnits.Sorted = False
        Me.cboGridUnits.TabStop = True
        Me.cboGridUnits.Visible = True
        Me.cboGridUnits.Name = "cboGridUnits"
        Me.cboPrecipUnits.Size = New System.Drawing.Size(143, 21)
        Me.cboPrecipUnits.Location = New System.Drawing.Point(139, 143)
        Me.cboPrecipUnits.Items.AddRange(New Object() {"centimeters", "inches"})
        Me.cboPrecipUnits.Sorted = True
        Me.cboPrecipUnits.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPrecipUnits.TabIndex = 4
        Me.cboPrecipUnits.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboPrecipUnits.BackColor = System.Drawing.SystemColors.Window
        Me.cboPrecipUnits.CausesValidation = True
        Me.cboPrecipUnits.Enabled = True
        Me.cboPrecipUnits.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboPrecipUnits.IntegralHeight = True
        Me.cboPrecipUnits.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboPrecipUnits.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboPrecipUnits.TabStop = True
        Me.cboPrecipUnits.Visible = True
        Me.cboPrecipUnits.Name = "cboPrecipUnits"
        Me._Label1_0.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me._Label1_0.Text = "Scenario Name:"
        Me._Label1_0.Size = New System.Drawing.Size(102, 19)
        Me._Label1_0.Location = New System.Drawing.Point(26, 25)
        Me._Label1_0.TabIndex = 19
        Me._Label1_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
        Me.lblRainingDays.Text = "Raining Days: "
        Me.lblRainingDays.Size = New System.Drawing.Size(72, 21)
        Me.lblRainingDays.Location = New System.Drawing.Point(294, 175)
        Me.lblRainingDays.TabIndex = 16
        Me.lblRainingDays.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRainingDays.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.lblRainingDays.BackColor = System.Drawing.SystemColors.Control
        Me.lblRainingDays.Enabled = True
        Me.lblRainingDays.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRainingDays.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblRainingDays.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblRainingDays.UseMnemonic = True
        Me.lblRainingDays.Visible = True
        Me.lblRainingDays.AutoSize = False
        Me.lblRainingDays.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lblRainingDays.Name = "lblRainingDays"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.Label3.Text = "Type:"
        Me.Label3.Size = New System.Drawing.Size(81, 17)
        Me.Label3.Location = New System.Drawing.Point(47, 204)
        Me.Label3.TabIndex = 14
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
        Me._Label1_2.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me._Label1_2.Text = "Precipitation Units:"
        Me._Label1_2.Size = New System.Drawing.Size(122, 19)
        Me._Label1_2.Location = New System.Drawing.Point(6, 144)
        Me._Label1_2.TabIndex = 13
        Me._Label1_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_2.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_2.Enabled = True
        Me._Label1_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_2.UseMnemonic = True
        Me._Label1_2.Visible = True
        Me._Label1_2.AutoSize = False
        Me._Label1_2.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me._Label1_2.Name = "_Label1_2"
        Me._Label1_1.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me._Label1_1.Text = "Description:"
        Me._Label1_1.Size = New System.Drawing.Size(89, 19)
        Me._Label1_1.Location = New System.Drawing.Point(39, 55)
        Me._Label1_1.TabIndex = 12
        Me._Label1_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
        Me._Label1_6.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me._Label1_6.Text = "Grid Units:"
        Me._Label1_6.Size = New System.Drawing.Size(74, 19)
        Me._Label1_6.Location = New System.Drawing.Point(54, 114)
        Me._Label1_6.TabIndex = 11
        Me._Label1_6.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_6.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_6.Enabled = True
        Me._Label1_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_6.UseMnemonic = True
        Me._Label1_6.Visible = True
        Me._Label1_6.AutoSize = False
        Me._Label1_6.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me._Label1_6.Name = "_Label1_6"
        Me._Label1_7.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me._Label1_7.Text = "Precipitation Grid:"
        Me._Label1_7.Size = New System.Drawing.Size(97, 19)
        Me._Label1_7.Location = New System.Drawing.Point(31, 85)
        Me._Label1_7.TabIndex = 10
        Me._Label1_7.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_7.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_7.Enabled = True
        Me._Label1_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_7.UseMnemonic = True
        Me._Label1_7.Visible = True
        Me._Label1_7.AutoSize = False
        Me._Label1_7.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me._Label1_7.Name = "_Label1_7"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.Label2.Text = "Time Period:"
        Me.Label2.Size = New System.Drawing.Size(81, 17)
        Me.Label2.Location = New System.Drawing.Point(47, 175)
        Me.Label2.TabIndex = 9
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
        Me.cmdSave.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.cmdSave.Text = "OK"
        Me.cmdSave.Enabled = False
        Me.cmdSave.Size = New System.Drawing.Size(65, 25)
        Me.cmdSave.Location = New System.Drawing.Point(314, 269)
        Me.cmdSave.TabIndex = 5
        Me.cmdSave.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSave.BackColor = System.Drawing.SystemColors.Control
        Me.cmdSave.CausesValidation = True
        Me.cmdSave.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSave.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdSave.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdSave.TabStop = True
        Me.cmdSave.Name = "cmdSave"
        Me.cmdQuit.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.cmdQuit.Text = "Cancel"
        Me.cmdQuit.Size = New System.Drawing.Size(65, 25)
        Me.cmdQuit.Location = New System.Drawing.Point(385, 269)
        Me.cmdQuit.TabIndex = 7
        Me.cmdQuit.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdQuit.BackColor = System.Drawing.SystemColors.Control
        Me.cmdQuit.CausesValidation = True
        Me.cmdQuit.Enabled = True
        Me.cmdQuit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdQuit.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdQuit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdQuit.TabStop = True
        Me.cmdQuit.Name = "cmdQuit"
        Me.Controls.Add(Frame1)
        Me.Controls.Add(cmdSave)
        Me.Controls.Add(cmdQuit)
        Me.Frame1.Controls.Add(cboScenName)
        Me.Frame1.Controls.Add(txtRainingDays)
        Me.Frame1.Controls.Add(cboTimePeriod)
        Me.Frame1.Controls.Add(cboPrecipType)
        Me.Frame1.Controls.Add(cmdBrowseFile)
        Me.Frame1.Controls.Add(txtDesc)
        Me.Frame1.Controls.Add(txtPrecipFile)
        Me.Frame1.Controls.Add(cboGridUnits)
        Me.Frame1.Controls.Add(cboPrecipUnits)
        Me.Frame1.Controls.Add(_Label1_0)
        Me.Frame1.Controls.Add(lblRainingDays)
        Me.Frame1.Controls.Add(Label3)
        Me.Frame1.Controls.Add(_Label1_2)
        Me.Frame1.Controls.Add(_Label1_1)
        Me.Frame1.Controls.Add(_Label1_6)
        Me.Frame1.Controls.Add(_Label1_7)
        Me.Frame1.Controls.Add(Label2)
        'Me.Label1.SetIndex(_Label1_0, CType(0, Short))
        'Me.Label1.SetIndex(_Label1_2, CType(2, Short))
        'Me.Label1.SetIndex(_Label1_1, CType(1, Short))
        'Me.Label1.SetIndex(_Label1_6, CType(6, Short))
        'Me.Label1.SetIndex(_Label1_7, CType(7, Short))
        'CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
        MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuPrecipOpts, Me.mnuHelp})
        mnuPrecipOpts.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuNewPrecip, Me.mnuDelPrecip})
        mnuHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuPrecipHelp})
        Me.Controls.Add(MainMenu1)
        Me.MainMenu1.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub
#End Region
End Class