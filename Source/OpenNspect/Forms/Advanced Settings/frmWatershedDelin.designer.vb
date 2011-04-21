<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmWatershedDelin
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
    Public WithEvents mnuNewWSDelin As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuNewExist As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuDelWSDelin As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuDefWSDelin As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuWSDelin As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuHelp As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
    Public WithEvents txtLSGrid As System.Windows.Forms.TextBox
    Public WithEvents cboDEMUnits As System.Windows.Forms.ComboBox
    Public WithEvents cboWSSize As System.Windows.Forms.ComboBox
    Public WithEvents chkHydroCorr As System.Windows.Forms.CheckBox
    Public WithEvents txtFlowAccumGrid As System.Windows.Forms.TextBox
    Public WithEvents txtWSFile As System.Windows.Forms.TextBox
    Public WithEvents txtStream As System.Windows.Forms.TextBox
    Public WithEvents cboWSDelin As System.Windows.Forms.ComboBox
    Public WithEvents txtDEMFile As System.Windows.Forms.TextBox
    Public WithEvents _Label1_6 As System.Windows.Forms.Label
    Public WithEvents _Label1_12 As System.Windows.Forms.Label
    Public WithEvents _Label1_2 As System.Windows.Forms.Label
    Public WithEvents _Label1_0 As System.Windows.Forms.Label
    Public WithEvents _Label1_3 As System.Windows.Forms.Label
    Public WithEvents _Label1_4 As System.Windows.Forms.Label
    Public WithEvents _Label1_5 As System.Windows.Forms.Label
    Public WithEvents _Label1_1 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents cmdQuit As System.Windows.Forms.Button
    'Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip()
        Me.mnuDefWSDelin = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuNewWSDelin = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuNewExist = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuDelWSDelin = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuHelp = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuWSDelin = New System.Windows.Forms.ToolStripMenuItem()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.txtLSGrid = New System.Windows.Forms.TextBox()
        Me.cboDEMUnits = New System.Windows.Forms.ComboBox()
        Me.cboWSSize = New System.Windows.Forms.ComboBox()
        Me.chkHydroCorr = New System.Windows.Forms.CheckBox()
        Me.txtFlowAccumGrid = New System.Windows.Forms.TextBox()
        Me.txtWSFile = New System.Windows.Forms.TextBox()
        Me.txtStream = New System.Windows.Forms.TextBox()
        Me.cboWSDelin = New System.Windows.Forms.ComboBox()
        Me.txtDEMFile = New System.Windows.Forms.TextBox()
        Me._Label1_6 = New System.Windows.Forms.Label()
        Me._Label1_12 = New System.Windows.Forms.Label()
        Me._Label1_2 = New System.Windows.Forms.Label()
        Me._Label1_0 = New System.Windows.Forms.Label()
        Me._Label1_3 = New System.Windows.Forms.Label()
        Me._Label1_4 = New System.Windows.Forms.Label()
        Me._Label1_5 = New System.Windows.Forms.Label()
        Me._Label1_1 = New System.Windows.Forms.Label()
        Me.cmdQuit = New System.Windows.Forms.Button()
        Me.MainMenu1.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuDefWSDelin, Me.mnuHelp})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(594, 24)
        Me.MainMenu1.TabIndex = 2
        '
        'mnuDefWSDelin
        '
        Me.mnuDefWSDelin.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuNewWSDelin, Me.mnuNewExist, Me.mnuDelWSDelin})
        Me.mnuDefWSDelin.Name = "mnuDefWSDelin"
        Me.mnuDefWSDelin.Size = New System.Drawing.Size(61, 20)
        Me.mnuDefWSDelin.Text = "&Options"
        '
        'mnuNewWSDelin
        '
        Me.mnuNewWSDelin.Name = "mnuNewWSDelin"
        Me.mnuNewWSDelin.Size = New System.Drawing.Size(205, 22)
        Me.mnuNewWSDelin.Text = "&New..."
        '
        'mnuNewExist
        '
        Me.mnuNewExist.Name = "mnuNewExist"
        Me.mnuNewExist.Size = New System.Drawing.Size(205, 22)
        Me.mnuNewExist.Text = "New from existing data..."
        '
        'mnuDelWSDelin
        '
        Me.mnuDelWSDelin.Name = "mnuDelWSDelin"
        Me.mnuDelWSDelin.Size = New System.Drawing.Size(205, 22)
        Me.mnuDelWSDelin.Text = "&Delete..."
        '
        'mnuHelp
        '
        Me.mnuHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuWSDelin})
        Me.mnuHelp.Name = "mnuHelp"
        Me.mnuHelp.Size = New System.Drawing.Size(44, 20)
        Me.mnuHelp.Text = "&Help"
        '
        'mnuWSDelin
        '
        Me.mnuWSDelin.Name = "mnuWSDelin"
        Me.mnuWSDelin.ShortcutKeys = CType((System.Windows.Forms.Keys.Shift Or System.Windows.Forms.Keys.F1), System.Windows.Forms.Keys)
        Me.mnuWSDelin.Size = New System.Drawing.Size(258, 22)
        Me.mnuWSDelin.Text = "Watershed Delineations..."
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.txtLSGrid)
        Me.Frame1.Controls.Add(Me.cboDEMUnits)
        Me.Frame1.Controls.Add(Me.cboWSSize)
        Me.Frame1.Controls.Add(Me.chkHydroCorr)
        Me.Frame1.Controls.Add(Me.txtFlowAccumGrid)
        Me.Frame1.Controls.Add(Me.txtWSFile)
        Me.Frame1.Controls.Add(Me.txtStream)
        Me.Frame1.Controls.Add(Me.cboWSDelin)
        Me.Frame1.Controls.Add(Me.txtDEMFile)
        Me.Frame1.Controls.Add(Me._Label1_6)
        Me.Frame1.Controls.Add(Me._Label1_12)
        Me.Frame1.Controls.Add(Me._Label1_2)
        Me.Frame1.Controls.Add(Me._Label1_0)
        Me.Frame1.Controls.Add(Me._Label1_3)
        Me.Frame1.Controls.Add(Me._Label1_4)
        Me.Frame1.Controls.Add(Me._Label1_5)
        Me.Frame1.Controls.Add(Me._Label1_1)
        Me.Frame1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(10, 28)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(572, 287)
        Me.Frame1.TabIndex = 1
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Browse Watershed Delineations  "
        '
        'txtLSGrid
        '
        Me.txtLSGrid.AcceptsReturn = True
        Me.txtLSGrid.BackColor = System.Drawing.SystemColors.Window
        Me.txtLSGrid.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtLSGrid.Enabled = False
        Me.txtLSGrid.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLSGrid.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtLSGrid.Location = New System.Drawing.Point(178, 193)
        Me.txtLSGrid.MaxLength = 0
        Me.txtLSGrid.Name = "txtLSGrid"
        Me.txtLSGrid.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtLSGrid.Size = New System.Drawing.Size(391, 20)
        Me.txtLSGrid.TabIndex = 17
        '
        'cboDEMUnits
        '
        Me.cboDEMUnits.BackColor = System.Drawing.SystemColors.Window
        Me.cboDEMUnits.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboDEMUnits.Enabled = False
        Me.cboDEMUnits.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboDEMUnits.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboDEMUnits.Items.AddRange(New Object() {"meters", "feet"})
        Me.cboDEMUnits.Location = New System.Drawing.Point(178, 67)
        Me.cboDEMUnits.Name = "cboDEMUnits"
        Me.cboDEMUnits.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboDEMUnits.Size = New System.Drawing.Size(202, 22)
        Me.cboDEMUnits.TabIndex = 16
        Me.cboDEMUnits.Text = "Combo1"
        '
        'cboWSSize
        '
        Me.cboWSSize.BackColor = System.Drawing.SystemColors.Window
        Me.cboWSSize.CausesValidation = False
        Me.cboWSSize.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboWSSize.Enabled = False
        Me.cboWSSize.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboWSSize.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboWSSize.Items.AddRange(New Object() {"small", "medium", "large"})
        Me.cboWSSize.Location = New System.Drawing.Point(178, 114)
        Me.cboWSSize.Name = "cboWSSize"
        Me.cboWSSize.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboWSSize.Size = New System.Drawing.Size(120, 22)
        Me.cboWSSize.TabIndex = 15
        Me.cboWSSize.Text = "cboWSSize"
        '
        'chkHydroCorr
        '
        Me.chkHydroCorr.BackColor = System.Drawing.SystemColors.Control
        Me.chkHydroCorr.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkHydroCorr.Enabled = False
        Me.chkHydroCorr.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkHydroCorr.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkHydroCorr.Location = New System.Drawing.Point(179, 93)
        Me.chkHydroCorr.Name = "chkHydroCorr"
        Me.chkHydroCorr.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkHydroCorr.Size = New System.Drawing.Size(173, 19)
        Me.chkHydroCorr.TabIndex = 7
        Me.chkHydroCorr.Text = "Hydrologically Corrected DEM"
        Me.chkHydroCorr.UseVisualStyleBackColor = False
        '
        'txtFlowAccumGrid
        '
        Me.txtFlowAccumGrid.AcceptsReturn = True
        Me.txtFlowAccumGrid.BackColor = System.Drawing.SystemColors.Window
        Me.txtFlowAccumGrid.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFlowAccumGrid.Enabled = False
        Me.txtFlowAccumGrid.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFlowAccumGrid.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFlowAccumGrid.Location = New System.Drawing.Point(178, 167)
        Me.txtFlowAccumGrid.MaxLength = 0
        Me.txtFlowAccumGrid.Name = "txtFlowAccumGrid"
        Me.txtFlowAccumGrid.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFlowAccumGrid.Size = New System.Drawing.Size(391, 20)
        Me.txtFlowAccumGrid.TabIndex = 6
        '
        'txtWSFile
        '
        Me.txtWSFile.AcceptsReturn = True
        Me.txtWSFile.BackColor = System.Drawing.SystemColors.Window
        Me.txtWSFile.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtWSFile.Enabled = False
        Me.txtWSFile.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtWSFile.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtWSFile.Location = New System.Drawing.Point(178, 140)
        Me.txtWSFile.MaxLength = 0
        Me.txtWSFile.Name = "txtWSFile"
        Me.txtWSFile.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtWSFile.Size = New System.Drawing.Size(391, 20)
        Me.txtWSFile.TabIndex = 5
        '
        'txtStream
        '
        Me.txtStream.AcceptsReturn = True
        Me.txtStream.BackColor = System.Drawing.SystemColors.Window
        Me.txtStream.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtStream.Enabled = False
        Me.txtStream.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStream.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtStream.Location = New System.Drawing.Point(178, 115)
        Me.txtStream.MaxLength = 0
        Me.txtStream.Name = "txtStream"
        Me.txtStream.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtStream.Size = New System.Drawing.Size(202, 20)
        Me.txtStream.TabIndex = 4
        Me.txtStream.Visible = False
        '
        'cboWSDelin
        '
        Me.cboWSDelin.BackColor = System.Drawing.SystemColors.Window
        Me.cboWSDelin.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboWSDelin.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboWSDelin.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboWSDelin.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboWSDelin.Location = New System.Drawing.Point(178, 16)
        Me.cboWSDelin.Name = "cboWSDelin"
        Me.cboWSDelin.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboWSDelin.Size = New System.Drawing.Size(202, 22)
        Me.cboWSDelin.TabIndex = 3
        '
        'txtDEMFile
        '
        Me.txtDEMFile.AcceptsReturn = True
        Me.txtDEMFile.BackColor = System.Drawing.SystemColors.Window
        Me.txtDEMFile.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDEMFile.Enabled = False
        Me.txtDEMFile.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDEMFile.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDEMFile.Location = New System.Drawing.Point(178, 43)
        Me.txtDEMFile.MaxLength = 0
        Me.txtDEMFile.Name = "txtDEMFile"
        Me.txtDEMFile.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDEMFile.Size = New System.Drawing.Size(388, 20)
        Me.txtDEMFile.TabIndex = 2
        '
        '_Label1_6
        '
        Me._Label1_6.AutoSize = True
        Me._Label1_6.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_6.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_6.Location = New System.Drawing.Point(124, 194)
        Me._Label1_6.Name = "_Label1_6"
        Me._Label1_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_6.Size = New System.Drawing.Size(46, 14)
        Me._Label1_6.TabIndex = 18
        Me._Label1_6.Text = "LS Grid:"
        Me._Label1_6.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_12
        '
        Me._Label1_12.AutoSize = True
        Me._Label1_12.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_12.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_12.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_12.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_12.Location = New System.Drawing.Point(113, 42)
        Me._Label1_12.Name = "_Label1_12"
        Me._Label1_12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_12.Size = New System.Drawing.Size(54, 14)
        Me._Label1_12.TabIndex = 14
        Me._Label1_12.Text = "DEM Grid:"
        Me._Label1_12.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_2
        '
        Me._Label1_2.AutoSize = True
        Me._Label1_2.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_2.Location = New System.Drawing.Point(135, 67)
        Me._Label1_2.Name = "_Label1_2"
        Me._Label1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_2.Size = New System.Drawing.Size(34, 14)
        Me._Label1_2.TabIndex = 13
        Me._Label1_2.Text = "Units:"
        Me._Label1_2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_0
        '
        Me._Label1_0.AutoSize = True
        Me._Label1_0.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_0.Location = New System.Drawing.Point(43, 115)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_0.Size = New System.Drawing.Size(130, 14)
        Me._Label1_0.TabIndex = 12
        Me._Label1_0.Text = "Stream Agreement Layer:"
        Me._Label1_0.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me._Label1_0.Visible = False
        '
        '_Label1_3
        '
        Me._Label1_3.AutoSize = True
        Me._Label1_3.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_3.Location = New System.Drawing.Point(25, 19)
        Me._Label1_3.Name = "_Label1_3"
        Me._Label1_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_3.Size = New System.Drawing.Size(148, 14)
        Me._Label1_3.TabIndex = 11
        Me._Label1_3.Text = "Watershed Delineation Name:"
        Me._Label1_3.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_4
        '
        Me._Label1_4.AutoSize = True
        Me._Label1_4.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_4.Location = New System.Drawing.Point(107, 140)
        Me._Label1_4.Name = "_Label1_4"
        Me._Label1_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_4.Size = New System.Drawing.Size(63, 14)
        Me._Label1_4.TabIndex = 10
        Me._Label1_4.Text = "Watershed:"
        Me._Label1_4.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_5
        '
        Me._Label1_5.AutoSize = True
        Me._Label1_5.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_5.Location = New System.Drawing.Point(48, 167)
        Me._Label1_5.Name = "_Label1_5"
        Me._Label1_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_5.Size = New System.Drawing.Size(124, 14)
        Me._Label1_5.TabIndex = 9
        Me._Label1_5.Text = "Flow Accumulation Grid:"
        Me._Label1_5.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_1
        '
        Me._Label1_1.AutoSize = True
        Me._Label1_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_1.Location = New System.Drawing.Point(68, 115)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_1.Size = New System.Drawing.Size(106, 14)
        Me._Label1_1.TabIndex = 8
        Me._Label1_1.Text = "Subwatershed Size:"
        Me._Label1_1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'cmdQuit
        '
        Me.cmdQuit.BackColor = System.Drawing.SystemColors.Control
        Me.cmdQuit.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdQuit.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdQuit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdQuit.Location = New System.Drawing.Point(517, 335)
        Me.cmdQuit.Name = "cmdQuit"
        Me.cmdQuit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdQuit.Size = New System.Drawing.Size(65, 25)
        Me.cmdQuit.TabIndex = 0
        Me.cmdQuit.Text = "Cancel"
        Me.cmdQuit.UseVisualStyleBackColor = False
        '
        'frmWatershedDelin
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(594, 372)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.cmdQuit)
        Me.Controls.Add(Me.MainMenu1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(3, 39)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmWatershedDelin"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Watershed Delineations"
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
End Class