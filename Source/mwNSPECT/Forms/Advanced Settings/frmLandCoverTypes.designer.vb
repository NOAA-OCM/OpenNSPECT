<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmLandCoverTypes
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
    Public WithEvents mnuNewLCType As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuDelLCType As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuImpLCType As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuExpLCType As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuLCTypes As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuAppend As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuInsertRow As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuDeleteRow As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuEdit As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuLCHelp As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuHelp As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
    Public dlgCMD1Open As System.Windows.Forms.OpenFileDialog
    Public dlgCMD1Save As System.Windows.Forms.SaveFileDialog
    Public WithEvents btnRestoreDefaults As System.Windows.Forms.Button
    Public WithEvents btnCancel As System.Windows.Forms.Button
    Public WithEvents btnSave As System.Windows.Forms.Button
    Public WithEvents txtLCTypeDesc As System.Windows.Forms.TextBox
    Public WithEvents cmbxLCType As System.Windows.Forms.ComboBox
    Public WithEvents _Label2_2 As System.Windows.Forms.Label
    Public WithEvents _Label2_1 As System.Windows.Forms.Label
    Public WithEvents _Label2_0 As System.Windows.Forms.Label
    Public WithEvents _Label1_6 As System.Windows.Forms.Label
    Public WithEvents _Label1_7 As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip
        Me.mnuLCTypes = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuNewLCType = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuDelLCType = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuImpLCType = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuExpLCType = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuEdit = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuAppend = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuInsertRow = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuDeleteRow = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuHelp = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuLCHelp = New System.Windows.Forms.ToolStripMenuItem
        Me.dlgCMD1Open = New System.Windows.Forms.OpenFileDialog
        Me.dlgCMD1Save = New System.Windows.Forms.SaveFileDialog
        Me.btnRestoreDefaults = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.txtLCTypeDesc = New System.Windows.Forms.TextBox
        Me.cmbxLCType = New System.Windows.Forms.ComboBox
        Me._Label2_2 = New System.Windows.Forms.Label
        Me._Label2_1 = New System.Windows.Forms.Label
        Me._Label2_0 = New System.Windows.Forms.Label
        Me._Label1_6 = New System.Windows.Forms.Label
        Me._Label1_7 = New System.Windows.Forms.Label
        Me.dgvLCTypes = New System.Windows.Forms.DataGridView
        Me.colValue = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colNameCol = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colCNA = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colCNB = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colCNC = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colCND = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colCoverFactor = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colWetCheck = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.colLCTYPEID = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colLCClassID = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.cntxmnuGrid = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.AddRowToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.InsertRowToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.DeleteRowToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.MainMenu1.SuspendLayout()
        CType(Me.dgvLCTypes, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.cntxmnuGrid.SuspendLayout()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuLCTypes, Me.mnuEdit, Me.mnuHelp})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(618, 24)
        Me.MainMenu1.TabIndex = 14
        '
        'mnuLCTypes
        '
        Me.mnuLCTypes.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuNewLCType, Me.mnuDelLCType, Me.mnuImpLCType, Me.mnuExpLCType})
        Me.mnuLCTypes.Name = "mnuLCTypes"
        Me.mnuLCTypes.Size = New System.Drawing.Size(61, 20)
        Me.mnuLCTypes.Text = "&Options"
        '
        'mnuNewLCType
        '
        Me.mnuNewLCType.Name = "mnuNewLCType"
        Me.mnuNewLCType.Size = New System.Drawing.Size(119, 22)
        Me.mnuNewLCType.Text = "&New..."
        '
        'mnuDelLCType
        '
        Me.mnuDelLCType.Name = "mnuDelLCType"
        Me.mnuDelLCType.Size = New System.Drawing.Size(119, 22)
        Me.mnuDelLCType.Text = "&Delete..."
        '
        'mnuImpLCType
        '
        Me.mnuImpLCType.Name = "mnuImpLCType"
        Me.mnuImpLCType.Size = New System.Drawing.Size(119, 22)
        Me.mnuImpLCType.Text = "&Import..."
        '
        'mnuExpLCType
        '
        Me.mnuExpLCType.Name = "mnuExpLCType"
        Me.mnuExpLCType.Size = New System.Drawing.Size(119, 22)
        Me.mnuExpLCType.Text = "&Export..."
        '
        'mnuEdit
        '
        Me.mnuEdit.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuAppend, Me.mnuInsertRow, Me.mnuDeleteRow})
        Me.mnuEdit.Name = "mnuEdit"
        Me.mnuEdit.Size = New System.Drawing.Size(39, 20)
        Me.mnuEdit.Text = "Edit"
        '
        'mnuAppend
        '
        Me.mnuAppend.Name = "mnuAppend"
        Me.mnuAppend.Size = New System.Drawing.Size(133, 22)
        Me.mnuAppend.Text = "Add Row"
        '
        'mnuInsertRow
        '
        Me.mnuInsertRow.Name = "mnuInsertRow"
        Me.mnuInsertRow.Size = New System.Drawing.Size(133, 22)
        Me.mnuInsertRow.Text = "Insert Row"
        '
        'mnuDeleteRow
        '
        Me.mnuDeleteRow.Name = "mnuDeleteRow"
        Me.mnuDeleteRow.Size = New System.Drawing.Size(133, 22)
        Me.mnuDeleteRow.Text = "Delete Row"
        '
        'mnuHelp
        '
        Me.mnuHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuLCHelp})
        Me.mnuHelp.Name = "mnuHelp"
        Me.mnuHelp.Size = New System.Drawing.Size(44, 20)
        Me.mnuHelp.Text = "&Help"
        '
        'mnuLCHelp
        '
        Me.mnuLCHelp.Name = "mnuLCHelp"
        Me.mnuLCHelp.ShortcutKeys = CType((System.Windows.Forms.Keys.Shift Or System.Windows.Forms.Keys.F1), System.Windows.Forms.Keys)
        Me.mnuLCHelp.Size = New System.Drawing.Size(228, 22)
        Me.mnuLCHelp.Text = "Land Cover Types..."
        '
        'btnRestoreDefaults
        '
        Me.btnRestoreDefaults.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnRestoreDefaults.BackColor = System.Drawing.SystemColors.Control
        Me.btnRestoreDefaults.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnRestoreDefaults.Enabled = False
        Me.btnRestoreDefaults.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRestoreDefaults.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnRestoreDefaults.Location = New System.Drawing.Point(25, 551)
        Me.btnRestoreDefaults.Name = "btnRestoreDefaults"
        Me.btnRestoreDefaults.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnRestoreDefaults.Size = New System.Drawing.Size(103, 25)
        Me.btnRestoreDefaults.TabIndex = 6
        Me.btnRestoreDefaults.Text = "Restore Defaults"
        Me.btnRestoreDefaults.UseVisualStyleBackColor = False
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.BackColor = System.Drawing.SystemColors.Control
        Me.btnCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnCancel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnCancel.Location = New System.Drawing.Point(512, 550)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnCancel.Size = New System.Drawing.Size(65, 25)
        Me.btnCancel.TabIndex = 5
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = False
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSave.BackColor = System.Drawing.SystemColors.Control
        Me.btnSave.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnSave.Enabled = False
        Me.btnSave.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSave.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnSave.Location = New System.Drawing.Point(441, 550)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnSave.Size = New System.Drawing.Size(65, 25)
        Me.btnSave.TabIndex = 4
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = False
        '
        'txtLCTypeDesc
        '
        Me.txtLCTypeDesc.AcceptsReturn = True
        Me.txtLCTypeDesc.BackColor = System.Drawing.SystemColors.Window
        Me.txtLCTypeDesc.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtLCTypeDesc.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLCTypeDesc.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtLCTypeDesc.Location = New System.Drawing.Point(124, 52)
        Me.txtLCTypeDesc.MaxLength = 0
        Me.txtLCTypeDesc.Name = "txtLCTypeDesc"
        Me.txtLCTypeDesc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtLCTypeDesc.Size = New System.Drawing.Size(362, 20)
        Me.txtLCTypeDesc.TabIndex = 2
        '
        'cmbxLCType
        '
        Me.cmbxLCType.BackColor = System.Drawing.SystemColors.Window
        Me.cmbxLCType.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbxLCType.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbxLCType.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbxLCType.Location = New System.Drawing.Point(124, 26)
        Me.cmbxLCType.Name = "cmbxLCType"
        Me.cmbxLCType.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbxLCType.Size = New System.Drawing.Size(149, 22)
        Me.cmbxLCType.TabIndex = 0
        '
        '_Label2_2
        '
        Me._Label2_2.BackColor = System.Drawing.SystemColors.Control
        Me._Label2_2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me._Label2_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label2_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label2_2.Location = New System.Drawing.Point(451, 78)
        Me._Label2_2.Name = "_Label2_2"
        Me._Label2_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_2.Size = New System.Drawing.Size(160, 19)
        Me._Label2_2.TabIndex = 9
        Me._Label2_2.Text = "RUSLE"
        Me._Label2_2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_Label2_1
        '
        Me._Label2_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label2_1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me._Label2_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label2_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label2_1.Location = New System.Drawing.Point(7, 78)
        Me._Label2_1.Name = "_Label2_1"
        Me._Label2_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_1.Size = New System.Drawing.Size(227, 19)
        Me._Label2_1.TabIndex = 8
        Me._Label2_1.Text = "Classification"
        Me._Label2_1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_Label2_0
        '
        Me._Label2_0.BackColor = System.Drawing.SystemColors.Control
        Me._Label2_0.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me._Label2_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label2_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label2_0.Location = New System.Drawing.Point(236, 78)
        Me._Label2_0.Name = "_Label2_0"
        Me._Label2_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_0.Size = New System.Drawing.Size(213, 19)
        Me._Label2_0.TabIndex = 7
        Me._Label2_0.Text = "SCS Curve Numbers"
        Me._Label2_0.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_Label1_6
        '
        Me._Label1_6.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_6.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_6.Location = New System.Drawing.Point(21, 51)
        Me._Label1_6.Name = "_Label1_6"
        Me._Label1_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_6.Size = New System.Drawing.Size(86, 17)
        Me._Label1_6.TabIndex = 3
        Me._Label1_6.Text = "Description:"
        Me._Label1_6.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_7
        '
        Me._Label1_7.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_7.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_7.Location = New System.Drawing.Point(16, 29)
        Me._Label1_7.Name = "_Label1_7"
        Me._Label1_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_7.Size = New System.Drawing.Size(91, 17)
        Me._Label1_7.TabIndex = 1
        Me._Label1_7.Text = "Land Cover Type:"
        Me._Label1_7.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'dgvLCTypes
        '
        Me.dgvLCTypes.AllowUserToAddRows = False
        Me.dgvLCTypes.AllowUserToDeleteRows = False
        Me.dgvLCTypes.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvLCTypes.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvLCTypes.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colValue, Me.colNameCol, Me.colCNA, Me.colCNB, Me.colCNC, Me.colCND, Me.colCoverFactor, Me.colWetCheck, Me.colLCTYPEID, Me.colLCClassID})
        Me.dgvLCTypes.Location = New System.Drawing.Point(7, 100)
        Me.dgvLCTypes.MultiSelect = False
        Me.dgvLCTypes.Name = "dgvLCTypes"
        Me.dgvLCTypes.ShowCellToolTips = False
        Me.dgvLCTypes.Size = New System.Drawing.Size(604, 444)
        Me.dgvLCTypes.TabIndex = 15
        '
        'colValue
        '
        Me.colValue.DataPropertyName = "Value"
        Me.colValue.HeaderText = "Value"
        Me.colValue.Name = "colValue"
        Me.colValue.Width = 42
        '
        'colNameCol
        '
        Me.colNameCol.DataPropertyName = "Name"
        Me.colNameCol.HeaderText = "Name"
        Me.colNameCol.Name = "colNameCol"
        Me.colNameCol.Width = 145
        '
        'colCNA
        '
        Me.colCNA.DataPropertyName = "CN-A"
        Me.colCNA.HeaderText = "CN-A"
        Me.colCNA.Name = "colCNA"
        Me.colCNA.Width = 53
        '
        'colCNB
        '
        Me.colCNB.DataPropertyName = "CN-B"
        Me.colCNB.HeaderText = "CN-B"
        Me.colCNB.Name = "colCNB"
        Me.colCNB.Width = 53
        '
        'colCNC
        '
        Me.colCNC.DataPropertyName = "CN-C"
        Me.colCNC.HeaderText = "CN-C"
        Me.colCNC.Name = "colCNC"
        Me.colCNC.Width = 53
        '
        'colCND
        '
        Me.colCND.DataPropertyName = "CN-D"
        Me.colCND.HeaderText = "CN-D"
        Me.colCND.Name = "colCND"
        Me.colCND.Width = 55
        '
        'colCoverFactor
        '
        Me.colCoverFactor.DataPropertyName = "CoverFactor"
        Me.colCoverFactor.HeaderText = "Cover-Factor"
        Me.colCoverFactor.Name = "colCoverFactor"
        Me.colCoverFactor.Width = 95
        '
        'colWetCheck
        '
        Me.colWetCheck.DataPropertyName = "W_WL"
        Me.colWetCheck.FalseValue = "0"
        Me.colWetCheck.HeaderText = "Wet"
        Me.colWetCheck.Name = "colWetCheck"
        Me.colWetCheck.TrueValue = "1"
        Me.colWetCheck.Width = 53
        '
        'colLCTYPEID
        '
        Me.colLCTYPEID.DataPropertyName = "LCTypeID"
        Me.colLCTYPEID.HeaderText = "LCTYPEID"
        Me.colLCTYPEID.Name = "colLCTYPEID"
        Me.colLCTYPEID.Visible = False
        '
        'colLCClassID
        '
        Me.colLCClassID.DataPropertyName = "LCClassID"
        Me.colLCClassID.HeaderText = "LCClassID"
        Me.colLCClassID.Name = "colLCClassID"
        Me.colLCClassID.Visible = False
        '
        'cntxmnuGrid
        '
        Me.cntxmnuGrid.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AddRowToolStripMenuItem, Me.InsertRowToolStripMenuItem, Me.DeleteRowToolStripMenuItem})
        Me.cntxmnuGrid.Name = "ContextMenuStrip1"
        Me.cntxmnuGrid.Size = New System.Drawing.Size(134, 70)
        '
        'AddRowToolStripMenuItem
        '
        Me.AddRowToolStripMenuItem.Name = "AddRowToolStripMenuItem"
        Me.AddRowToolStripMenuItem.Size = New System.Drawing.Size(133, 22)
        Me.AddRowToolStripMenuItem.Text = "Add Row"
        '
        'InsertRowToolStripMenuItem
        '
        Me.InsertRowToolStripMenuItem.Name = "InsertRowToolStripMenuItem"
        Me.InsertRowToolStripMenuItem.Size = New System.Drawing.Size(133, 22)
        Me.InsertRowToolStripMenuItem.Text = "Insert Row"
        '
        'DeleteRowToolStripMenuItem
        '
        Me.DeleteRowToolStripMenuItem.Name = "DeleteRowToolStripMenuItem"
        Me.DeleteRowToolStripMenuItem.Size = New System.Drawing.Size(133, 22)
        Me.DeleteRowToolStripMenuItem.Text = "Delete Row"
        '
        'frmLandCoverTypes
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(618, 590)
        Me.Controls.Add(Me.dgvLCTypes)
        Me.Controls.Add(Me.btnRestoreDefaults)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.txtLCTypeDesc)
        Me.Controls.Add(Me.cmbxLCType)
        Me.Controls.Add(Me._Label2_2)
        Me.Controls.Add(Me._Label2_1)
        Me.Controls.Add(Me._Label2_0)
        Me.Controls.Add(Me._Label1_6)
        Me.Controls.Add(Me._Label1_7)
        Me.Controls.Add(Me.MainMenu1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(268, 71)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmLandCoverTypes"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Land Cover Types"
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        CType(Me.dgvLCTypes, System.ComponentModel.ISupportInitialize).EndInit()
        Me.cntxmnuGrid.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvLCTypes As System.Windows.Forms.DataGridView
    Friend WithEvents colValue As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colNameCol As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colCNA As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colCNB As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colCNC As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colCND As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colCoverFactor As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colWetCheck As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents colLCTYPEID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colLCClassID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents cntxmnuGrid As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents AddRowToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents InsertRowToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DeleteRowToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
#End Region
End Class