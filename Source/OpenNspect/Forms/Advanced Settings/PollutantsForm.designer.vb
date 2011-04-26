<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class PollutantsForm
    Inherits OpenNspect.BaseDialogForm
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
    Public WithEvents mnuAddPoll As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuDeletePoll As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuPoll_1 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuCoeffNewSet As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuCoeffCopySet As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuCoeffDeleteSet As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuCoeffImportSet As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuCoeffExportSet As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuCoeff As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuPollHelp As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuCoeffHelp As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuHelp As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
    Public WithEvents _Label1_5 As System.Windows.Forms.Label
    Public WithEvents _Label1_6 As System.Windows.Forms.Label
    Public WithEvents _Label1_7 As System.Windows.Forms.Label
    Public WithEvents _Label1_1 As System.Windows.Forms.Label
    Public WithEvents _Label1_2 As System.Windows.Forms.Label
    Public WithEvents _Label1_3 As System.Windows.Forms.Label
    Public WithEvents cboCoeffSet As System.Windows.Forms.ComboBox
    Public WithEvents txtCoeffSetDesc As System.Windows.Forms.TextBox
    Public WithEvents txtLCType As System.Windows.Forms.TextBox
    Public WithEvents _SSTab1_TabPage0 As System.Windows.Forms.TabPage
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents _SSTab1_TabPage1 As System.Windows.Forms.TabPage
    Public WithEvents SSTab1 As System.Windows.Forms.TabControl
    Public WithEvents cboPollName As System.Windows.Forms.ComboBox
    Public WithEvents _Label1_0 As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PollutantsForm))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip()
        Me._mnuPoll_1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuAddPoll = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuDeletePoll = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuCoeff = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuCoeffNewSet = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuCoeffCopySet = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuCoeffDeleteSet = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuCoeffImportSet = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuCoeffExportSet = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuHelp = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuPollHelp = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuCoeffHelp = New System.Windows.Forms.ToolStripMenuItem()
        Me.SSTab1 = New System.Windows.Forms.TabControl()
        Me._SSTab1_TabPage0 = New System.Windows.Forms.TabPage()
        Me.dgvCoef = New System.Windows.Forms.DataGridView()
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn6 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn7 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn8 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me._Label1_6 = New System.Windows.Forms.Label()
        Me._Label1_7 = New System.Windows.Forms.Label()
        Me._Label1_1 = New System.Windows.Forms.Label()
        Me._Label1_2 = New System.Windows.Forms.Label()
        Me._Label1_3 = New System.Windows.Forms.Label()
        Me.cboCoeffSet = New System.Windows.Forms.ComboBox()
        Me.txtCoeffSetDesc = New System.Windows.Forms.TextBox()
        Me.txtLCType = New System.Windows.Forms.TextBox()
        Me._Label1_5 = New System.Windows.Forms.Label()
        Me._SSTab1_TabPage1 = New System.Windows.Forms.TabPage()
        Me.dgvWaterQuality = New System.Windows.Forms.DataGridView()
        Me.ColName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Description = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Threshold = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PollID = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cboPollName = New System.Windows.Forms.ComboBox()
        Me._Label1_0 = New System.Windows.Forms.Label()
        Me.MainMenu1.SuspendLayout()
        Me.SSTab1.SuspendLayout()
        Me._SSTab1_TabPage0.SuspendLayout()
        CType(Me.dgvCoef, System.ComponentModel.ISupportInitialize).BeginInit()
        Me._SSTab1_TabPage1.SuspendLayout()
        CType(Me.dgvWaterQuality, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuPoll_1, Me.mnuCoeff, Me.mnuHelp})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(594, 24)
        Me.MainMenu1.TabIndex = 16
        '
        '_mnuPoll_1
        '
        Me._mnuPoll_1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuAddPoll, Me.mnuDeletePoll})
        Me._mnuPoll_1.Name = "_mnuPoll_1"
        Me._mnuPoll_1.Size = New System.Drawing.Size(72, 20)
        Me._mnuPoll_1.Text = "&Pollutants"
        '
        'mnuAddPoll
        '
        Me.mnuAddPoll.Name = "mnuAddPoll"
        Me.mnuAddPoll.Size = New System.Drawing.Size(116, 22)
        Me.mnuAddPoll.Text = "&Add..."
        '
        'mnuDeletePoll
        '
        Me.mnuDeletePoll.Name = "mnuDeletePoll"
        Me.mnuDeletePoll.Size = New System.Drawing.Size(116, 22)
        Me.mnuDeletePoll.Text = "&Delete..."
        '
        'mnuCoeff
        '
        Me.mnuCoeff.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuCoeffNewSet, Me.mnuCoeffCopySet, Me.mnuCoeffDeleteSet, Me.mnuCoeffImportSet, Me.mnuCoeffExportSet})
        Me.mnuCoeff.Name = "mnuCoeff"
        Me.mnuCoeff.Size = New System.Drawing.Size(82, 20)
        Me.mnuCoeff.Text = "&Coefficients"
        '
        'mnuCoeffNewSet
        '
        Me.mnuCoeffNewSet.Name = "mnuCoeffNewSet"
        Me.mnuCoeffNewSet.Size = New System.Drawing.Size(138, 22)
        Me.mnuCoeffNewSet.Text = "&New Set..."
        '
        'mnuCoeffCopySet
        '
        Me.mnuCoeffCopySet.Name = "mnuCoeffCopySet"
        Me.mnuCoeffCopySet.Size = New System.Drawing.Size(138, 22)
        Me.mnuCoeffCopySet.Text = "&Copy Set..."
        '
        'mnuCoeffDeleteSet
        '
        Me.mnuCoeffDeleteSet.Name = "mnuCoeffDeleteSet"
        Me.mnuCoeffDeleteSet.Size = New System.Drawing.Size(138, 22)
        Me.mnuCoeffDeleteSet.Text = "&Delete Set..."
        '
        'mnuCoeffImportSet
        '
        Me.mnuCoeffImportSet.Name = "mnuCoeffImportSet"
        Me.mnuCoeffImportSet.Size = New System.Drawing.Size(138, 22)
        Me.mnuCoeffImportSet.Text = "&Import Set..."
        '
        'mnuCoeffExportSet
        '
        Me.mnuCoeffExportSet.Name = "mnuCoeffExportSet"
        Me.mnuCoeffExportSet.Size = New System.Drawing.Size(138, 22)
        Me.mnuCoeffExportSet.Text = "&Export Set..."
        '
        'mnuHelp
        '
        Me.mnuHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuPollHelp, Me.mnuCoeffHelp})
        Me.mnuHelp.Name = "mnuHelp"
        Me.mnuHelp.Size = New System.Drawing.Size(44, 20)
        Me.mnuHelp.Text = "&Help"
        '
        'mnuPollHelp
        '
        Me.mnuPollHelp.Name = "mnuPollHelp"
        Me.mnuPollHelp.ShortcutKeys = CType((System.Windows.Forms.Keys.Shift Or System.Windows.Forms.Keys.F1), System.Windows.Forms.Keys)
        Me.mnuPollHelp.Size = New System.Drawing.Size(197, 22)
        Me.mnuPollHelp.Text = "Pollutants..."
        '
        'mnuCoeffHelp
        '
        Me.mnuCoeffHelp.Name = "mnuCoeffHelp"
        Me.mnuCoeffHelp.ShortcutKeys = CType((System.Windows.Forms.Keys.Shift Or System.Windows.Forms.Keys.F2), System.Windows.Forms.Keys)
        Me.mnuCoeffHelp.Size = New System.Drawing.Size(197, 22)
        Me.mnuCoeffHelp.Text = "Coefficients..."
        '
        'SSTab1
        '
        Me.SSTab1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SSTab1.Controls.Add(Me._SSTab1_TabPage0)
        Me.SSTab1.Controls.Add(Me._SSTab1_TabPage1)
        Me.SSTab1.ItemSize = New System.Drawing.Size(42, 18)
        Me.SSTab1.Location = New System.Drawing.Point(13, 58)
        Me.SSTab1.Name = "SSTab1"
        Me.SSTab1.SelectedIndex = 0
        Me.SSTab1.Size = New System.Drawing.Size(566, 617)
        Me.SSTab1.TabIndex = 4
        '
        '_SSTab1_TabPage0
        '
        Me._SSTab1_TabPage0.Controls.Add(Me.dgvCoef)
        Me._SSTab1_TabPage0.Controls.Add(Me._Label1_6)
        Me._SSTab1_TabPage0.Controls.Add(Me._Label1_7)
        Me._SSTab1_TabPage0.Controls.Add(Me._Label1_1)
        Me._SSTab1_TabPage0.Controls.Add(Me._Label1_2)
        Me._SSTab1_TabPage0.Controls.Add(Me._Label1_3)
        Me._SSTab1_TabPage0.Controls.Add(Me.cboCoeffSet)
        Me._SSTab1_TabPage0.Controls.Add(Me.txtCoeffSetDesc)
        Me._SSTab1_TabPage0.Controls.Add(Me.txtLCType)
        Me._SSTab1_TabPage0.Controls.Add(Me._Label1_5)
        Me._SSTab1_TabPage0.Location = New System.Drawing.Point(4, 22)
        Me._SSTab1_TabPage0.Name = "_SSTab1_TabPage0"
        Me._SSTab1_TabPage0.Size = New System.Drawing.Size(558, 591)
        Me._SSTab1_TabPage0.TabIndex = 0
        Me._SSTab1_TabPage0.Text = "Coefficients"
        '
        'dgvCoef
        '
        Me.dgvCoef.AllowUserToAddRows = False
        Me.dgvCoef.AllowUserToDeleteRows = False
        Me.dgvCoef.AllowUserToResizeColumns = False
        Me.dgvCoef.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvCoef.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewTextBoxColumn1, Me.DataGridViewTextBoxColumn2, Me.DataGridViewTextBoxColumn3, Me.DataGridViewTextBoxColumn4, Me.DataGridViewTextBoxColumn5, Me.DataGridViewTextBoxColumn6, Me.DataGridViewTextBoxColumn7, Me.DataGridViewTextBoxColumn8})
        Me.dgvCoef.Location = New System.Drawing.Point(14, 98)
        Me.dgvCoef.Name = "dgvCoef"
        Me.dgvCoef.Size = New System.Drawing.Size(541, 492)
        Me.dgvCoef.TabIndex = 20
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.DataPropertyName = "Value"
        Me.DataGridViewTextBoxColumn1.HeaderText = "Value"
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        Me.DataGridViewTextBoxColumn1.ReadOnly = True
        Me.DataGridViewTextBoxColumn1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.DataGridViewTextBoxColumn1.Width = 53
        '
        'DataGridViewTextBoxColumn2
        '
        Me.DataGridViewTextBoxColumn2.DataPropertyName = "Name"
        Me.DataGridViewTextBoxColumn2.HeaderText = "Name"
        Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
        Me.DataGridViewTextBoxColumn2.ReadOnly = True
        Me.DataGridViewTextBoxColumn2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.DataGridViewTextBoxColumn2.Width = 165
        '
        'DataGridViewTextBoxColumn3
        '
        Me.DataGridViewTextBoxColumn3.DataPropertyName = "Type1"
        Me.DataGridViewTextBoxColumn3.HeaderText = "Type 1"
        Me.DataGridViewTextBoxColumn3.Name = "DataGridViewTextBoxColumn3"
        Me.DataGridViewTextBoxColumn3.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.DataGridViewTextBoxColumn3.Width = 53
        '
        'DataGridViewTextBoxColumn4
        '
        Me.DataGridViewTextBoxColumn4.DataPropertyName = "Type2"
        Me.DataGridViewTextBoxColumn4.HeaderText = "Type 2"
        Me.DataGridViewTextBoxColumn4.Name = "DataGridViewTextBoxColumn4"
        Me.DataGridViewTextBoxColumn4.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.DataGridViewTextBoxColumn4.Width = 53
        '
        'DataGridViewTextBoxColumn5
        '
        Me.DataGridViewTextBoxColumn5.DataPropertyName = "Type3"
        Me.DataGridViewTextBoxColumn5.HeaderText = "Type 3"
        Me.DataGridViewTextBoxColumn5.Name = "DataGridViewTextBoxColumn5"
        Me.DataGridViewTextBoxColumn5.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.DataGridViewTextBoxColumn5.Width = 53
        '
        'DataGridViewTextBoxColumn6
        '
        Me.DataGridViewTextBoxColumn6.DataPropertyName = "Type4"
        Me.DataGridViewTextBoxColumn6.HeaderText = "Type 4"
        Me.DataGridViewTextBoxColumn6.Name = "DataGridViewTextBoxColumn6"
        Me.DataGridViewTextBoxColumn6.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.DataGridViewTextBoxColumn6.Width = 53
        '
        'DataGridViewTextBoxColumn7
        '
        Me.DataGridViewTextBoxColumn7.DataPropertyName = "CoeffID"
        Me.DataGridViewTextBoxColumn7.HeaderText = "SetID"
        Me.DataGridViewTextBoxColumn7.Name = "DataGridViewTextBoxColumn7"
        Me.DataGridViewTextBoxColumn7.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.DataGridViewTextBoxColumn7.Visible = False
        '
        'DataGridViewTextBoxColumn8
        '
        Me.DataGridViewTextBoxColumn8.DataPropertyName = "LCCLASSID"
        Me.DataGridViewTextBoxColumn8.HeaderText = "LCClassID"
        Me.DataGridViewTextBoxColumn8.Name = "DataGridViewTextBoxColumn8"
        Me.DataGridViewTextBoxColumn8.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.DataGridViewTextBoxColumn8.Visible = False
        '
        '_Label1_6
        '
        Me._Label1_6.Location = New System.Drawing.Point(18, 55)
        Me._Label1_6.Name = "_Label1_6"
        Me._Label1_6.Size = New System.Drawing.Size(74, 16)
        Me._Label1_6.TabIndex = 9
        Me._Label1_6.Text = "Description:"
        '
        '_Label1_7
        '
        Me._Label1_7.Location = New System.Drawing.Point(271, 32)
        Me._Label1_7.Name = "_Label1_7"
        Me._Label1_7.Size = New System.Drawing.Size(97, 16)
        Me._Label1_7.TabIndex = 10
        Me._Label1_7.Text = "Land Cover Type:"
        '
        '_Label1_1
        '
        Me._Label1_1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me._Label1_1.Location = New System.Drawing.Point(56, 80)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.Size = New System.Drawing.Size(214, 16)
        Me._Label1_1.TabIndex = 11
        Me._Label1_1.Text = "Class"
        Me._Label1_1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_Label1_2
        '
        Me._Label1_2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me._Label1_2.Location = New System.Drawing.Point(276, 80)
        Me._Label1_2.Name = "_Label1_2"
        Me._Label1_2.Size = New System.Drawing.Size(224, 16)
        Me._Label1_2.TabIndex = 12
        Me._Label1_2.Text = "Coefficients (mg/L)"
        Me._Label1_2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_Label1_3
        '
        Me._Label1_3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me._Label1_3.Location = New System.Drawing.Point(14, 80)
        Me._Label1_3.Name = "_Label1_3"
        Me._Label1_3.Size = New System.Drawing.Size(36, 16)
        Me._Label1_3.TabIndex = 13
        Me._Label1_3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'cboCoeffSet
        '
        Me.cboCoeffSet.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCoeffSet.Location = New System.Drawing.Point(100, 30)
        Me.cboCoeffSet.Name = "cboCoeffSet"
        Me.cboCoeffSet.Size = New System.Drawing.Size(147, 21)
        Me.cboCoeffSet.Sorted = True
        Me.cboCoeffSet.TabIndex = 5
        '
        'txtCoeffSetDesc
        '
        Me.txtCoeffSetDesc.AcceptsReturn = True
        Me.txtCoeffSetDesc.Location = New System.Drawing.Point(100, 56)
        Me.txtCoeffSetDesc.MaxLength = 0
        Me.txtCoeffSetDesc.Name = "txtCoeffSetDesc"
        Me.txtCoeffSetDesc.Size = New System.Drawing.Size(375, 20)
        Me.txtCoeffSetDesc.TabIndex = 6
        '
        'txtLCType
        '
        Me.txtLCType.AcceptsReturn = True
        Me.txtLCType.Enabled = False
        Me.txtLCType.Location = New System.Drawing.Point(374, 30)
        Me.txtLCType.MaxLength = 0
        Me.txtLCType.Name = "txtLCType"
        Me.txtLCType.Size = New System.Drawing.Size(102, 20)
        Me.txtLCType.TabIndex = 7
        '
        '_Label1_5
        '
        Me._Label1_5.Location = New System.Drawing.Point(17, 32)
        Me._Label1_5.Name = "_Label1_5"
        Me._Label1_5.Size = New System.Drawing.Size(97, 16)
        Me._Label1_5.TabIndex = 8
        Me._Label1_5.Text = "Coefficient Set:"
        '
        '_SSTab1_TabPage1
        '
        Me._SSTab1_TabPage1.Controls.Add(Me.dgvWaterQuality)
        Me._SSTab1_TabPage1.Controls.Add(Me.Label2)
        Me._SSTab1_TabPage1.Location = New System.Drawing.Point(4, 22)
        Me._SSTab1_TabPage1.Name = "_SSTab1_TabPage1"
        Me._SSTab1_TabPage1.Size = New System.Drawing.Size(558, 591)
        Me._SSTab1_TabPage1.TabIndex = 1
        Me._SSTab1_TabPage1.Text = "Water Quality Standards"
        '
        'dgvWaterQuality
        '
        Me.dgvWaterQuality.AllowUserToAddRows = False
        Me.dgvWaterQuality.AllowUserToDeleteRows = False
        Me.dgvWaterQuality.AllowUserToResizeColumns = False
        Me.dgvWaterQuality.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvWaterQuality.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColName, Me.Description, Me.Threshold, Me.PollID})
        Me.dgvWaterQuality.Location = New System.Drawing.Point(22, 27)
        Me.dgvWaterQuality.Name = "dgvWaterQuality"
        Me.dgvWaterQuality.Size = New System.Drawing.Size(479, 290)
        Me.dgvWaterQuality.TabIndex = 19
        '
        'ColName
        '
        Me.ColName.DataPropertyName = "Name"
        Me.ColName.HeaderText = "Name"
        Me.ColName.Name = "ColName"
        Me.ColName.ReadOnly = True
        Me.ColName.Width = 133
        '
        'Description
        '
        Me.Description.DataPropertyName = "Description"
        Me.Description.HeaderText = "Description"
        Me.Description.Name = "Description"
        Me.Description.ReadOnly = True
        Me.Description.Width = 205
        '
        'Threshold
        '
        Me.Threshold.DataPropertyName = "Threshold"
        Me.Threshold.HeaderText = "Threshold"
        Me.Threshold.Name = "Threshold"
        Me.Threshold.Width = 95
        '
        'PollID
        '
        Me.PollID.DataPropertyName = "POLL_WQCRITID"
        Me.PollID.HeaderText = "PollID"
        Me.PollID.Name = "PollID"
        Me.PollID.Visible = False
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(19, 561)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(137, 16)
        Me.Label2.TabIndex = 18
        Me.Label2.Text = "Threshold Units: ug/L"
        '
        'cboPollName
        '
        Me.cboPollName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPollName.Location = New System.Drawing.Point(106, 32)
        Me.cboPollName.Name = "cboPollName"
        Me.cboPollName.Size = New System.Drawing.Size(150, 21)
        Me.cboPollName.TabIndex = 1
        '
        '_Label1_0
        '
        Me._Label1_0.Location = New System.Drawing.Point(21, 32)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.Size = New System.Drawing.Size(89, 16)
        Me._Label1_0.TabIndex = 0
        Me._Label1_0.Text = "Pollutant Name:"
        '
        'PollutantsForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(594, 717)
        Me.Controls.Add(Me.SSTab1)
        Me.Controls.Add(Me.cboPollName)
        Me.Controls.Add(Me._Label1_0)
        Me.Controls.Add(Me.MainMenu1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(113, 127)
        Me.Name = "PollutantsForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Pollutants"
        Me.Controls.SetChildIndex(Me.MainMenu1, 0)
        Me.Controls.SetChildIndex(Me._Label1_0, 0)
        Me.Controls.SetChildIndex(Me.cboPollName, 0)
        Me.Controls.SetChildIndex(Me.SSTab1, 0)
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        Me.SSTab1.ResumeLayout(False)
        Me._SSTab1_TabPage0.ResumeLayout(False)
        Me._SSTab1_TabPage0.PerformLayout()
        CType(Me.dgvCoef, System.ComponentModel.ISupportInitialize).EndInit()
        Me._SSTab1_TabPage1.ResumeLayout(False)
        CType(Me.dgvWaterQuality, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvWaterQuality As System.Windows.Forms.DataGridView
    Friend WithEvents dgvCoef As System.Windows.Forms.DataGridView
    Friend WithEvents ColName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Description As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Threshold As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PollID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn6 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn7 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn8 As System.Windows.Forms.DataGridViewTextBoxColumn
#End Region
End Class