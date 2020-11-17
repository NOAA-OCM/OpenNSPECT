<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class WaterQualityStandardsForm
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
    Public WithEvents mnuNewWQStd As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuDelWQStd As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuCopyWQStd As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuImpWQStd As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuExpWQStd As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuWQStd As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuAddRow As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuInsertRow As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuDeleteRow As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuEditCell As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuWQHelp As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuHelp As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
    Public WithEvents cboWQStdName As System.Windows.Forms.ComboBox
    Public WithEvents txtWQStdDesc As System.Windows.Forms.TextBox
    Public WithEvents _Label1_1 As System.Windows.Forms.Label
    Public WithEvents _Label1_0 As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip()
        Me.mnuWQStd = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuNewWQStd = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuDelWQStd = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuCopyWQStd = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuImpWQStd = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuExpWQStd = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuEditCell = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuAddRow = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuInsertRow = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuDeleteRow = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuHelp = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuWQHelp = New System.Windows.Forms.ToolStripMenuItem()
        Me.cboWQStdName = New System.Windows.Forms.ComboBox()
        Me.txtWQStdDesc = New System.Windows.Forms.TextBox()
        Me._Label1_1 = New System.Windows.Forms.Label()
        Me._Label1_0 = New System.Windows.Forms.Label()
        Me.dgvWaterQuality = New System.Windows.Forms.DataGridView()
        Me.colPollutant = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colThreshold = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Poll_WQCritID = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.MainMenu1.SuspendLayout()
        CType(Me.dgvWaterQuality, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuWQStd, Me.mnuEditCell, Me.mnuHelp})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(594, 24)
        Me.MainMenu1.TabIndex = 8
        '
        'mnuWQStd
        '
        Me.mnuWQStd.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuNewWQStd, Me.mnuDelWQStd, Me.mnuCopyWQStd, Me.mnuImpWQStd, Me.mnuExpWQStd})
        Me.mnuWQStd.Name = "mnuWQStd"
        Me.mnuWQStd.Size = New System.Drawing.Size(61, 20)
        Me.mnuWQStd.Text = "&Options"
        '
        'mnuNewWQStd
        '
        Me.mnuNewWQStd.Name = "mnuNewWQStd"
        Me.mnuNewWQStd.Size = New System.Drawing.Size(119, 22)
        Me.mnuNewWQStd.Text = "New..."
        '
        'mnuDelWQStd
        '
        Me.mnuDelWQStd.Name = "mnuDelWQStd"
        Me.mnuDelWQStd.Size = New System.Drawing.Size(119, 22)
        Me.mnuDelWQStd.Text = "Delete..."
        '
        'mnuCopyWQStd
        '
        Me.mnuCopyWQStd.Name = "mnuCopyWQStd"
        Me.mnuCopyWQStd.Size = New System.Drawing.Size(119, 22)
        Me.mnuCopyWQStd.Text = "Copy..."
        '
        'mnuImpWQStd
        '
        Me.mnuImpWQStd.Name = "mnuImpWQStd"
        Me.mnuImpWQStd.Size = New System.Drawing.Size(119, 22)
        Me.mnuImpWQStd.Text = "Import..."
        '
        'mnuExpWQStd
        '
        Me.mnuExpWQStd.Name = "mnuExpWQStd"
        Me.mnuExpWQStd.Size = New System.Drawing.Size(119, 22)
        Me.mnuExpWQStd.Text = "Export..."
        '
        'mnuEditCell
        '
        Me.mnuEditCell.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuAddRow, Me.mnuInsertRow, Me.mnuDeleteRow})
        Me.mnuEditCell.Name = "mnuEditCell"
        Me.mnuEditCell.Size = New System.Drawing.Size(39, 20)
        Me.mnuEditCell.Text = "Edit"
        Me.mnuEditCell.Visible = False
        '
        'mnuAddRow
        '
        Me.mnuAddRow.Name = "mnuAddRow"
        Me.mnuAddRow.Size = New System.Drawing.Size(133, 22)
        Me.mnuAddRow.Text = "Add Row"
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
        Me.mnuHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuWQHelp})
        Me.mnuHelp.Name = "mnuHelp"
        Me.mnuHelp.Size = New System.Drawing.Size(44, 20)
        Me.mnuHelp.Text = "&Help"
        '
        'mnuWQHelp
        '
        Me.mnuWQHelp.Name = "mnuWQHelp"
        Me.mnuWQHelp.ShortcutKeys = CType((System.Windows.Forms.Keys.Shift Or System.Windows.Forms.Keys.F1), System.Windows.Forms.Keys)
        Me.mnuWQHelp.Size = New System.Drawing.Size(261, 22)
        Me.mnuWQHelp.Text = "Water Quality Standards..."
        '
        'cboWQStdName
        '
        Me.cboWQStdName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboWQStdName.Location = New System.Drawing.Point(97, 28)
        Me.cboWQStdName.Name = "cboWQStdName"
        Me.cboWQStdName.Size = New System.Drawing.Size(464, 21)
        Me.cboWQStdName.TabIndex = 0
        '
        'txtWQStdDesc
        '
        Me.txtWQStdDesc.AcceptsReturn = True
        Me.txtWQStdDesc.Location = New System.Drawing.Point(97, 54)
        Me.txtWQStdDesc.MaxLength = 100
        Me.txtWQStdDesc.Name = "txtWQStdDesc"
        Me.txtWQStdDesc.Size = New System.Drawing.Size(464, 20)
        Me.txtWQStdDesc.TabIndex = 1
        '
        '_Label1_1
        '
        Me._Label1_1.Location = New System.Drawing.Point(15, 56)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.Size = New System.Drawing.Size(67, 18)
        Me._Label1_1.TabIndex = 5
        Me._Label1_1.Text = "Description:"
        '
        '_Label1_0
        '
        Me._Label1_0.Location = New System.Drawing.Point(14, 30)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.Size = New System.Drawing.Size(95, 18)
        Me._Label1_0.TabIndex = 4
        Me._Label1_0.Text = "Standard Name:"
        '
        'dgvWaterQuality
        '
        Me.dgvWaterQuality.AllowUserToAddRows = False
        Me.dgvWaterQuality.AllowUserToDeleteRows = False
        Me.dgvWaterQuality.AllowUserToResizeColumns = False
        Me.dgvWaterQuality.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvWaterQuality.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvWaterQuality.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colPollutant, Me.colThreshold, Me.Poll_WQCritID})
        Me.dgvWaterQuality.Location = New System.Drawing.Point(18, 78)
        Me.dgvWaterQuality.Name = "dgvWaterQuality"
        Me.dgvWaterQuality.Size = New System.Drawing.Size(543, 225)
        Me.dgvWaterQuality.TabIndex = 9
        '
        'colPollutant
        '
        Me.colPollutant.DataPropertyName = "Name"
        Me.colPollutant.HeaderText = "Pollutant"
        Me.colPollutant.Name = "colPollutant"
        Me.colPollutant.ReadOnly = True
        Me.colPollutant.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.colPollutant.Width = 185
        '
        'colThreshold
        '
        Me.colThreshold.DataPropertyName = "Threshold"
        Me.colThreshold.HeaderText = "Threshold"
        Me.colThreshold.Name = "colThreshold"
        Me.colThreshold.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.colThreshold.Width = 130
        '
        'Poll_WQCritID
        '
        Me.Poll_WQCritID.DataPropertyName = "POLL_WQCRITID"
        Me.Poll_WQCritID.HeaderText = "Poll_WQCritID"
        Me.Poll_WQCritID.Name = "Poll_WQCritID"
        Me.Poll_WQCritID.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.Poll_WQCritID.Visible = False
        '
        'WaterQualityStandardsForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(594, 345)
        Me.Controls.Add(Me.dgvWaterQuality)
        Me.Controls.Add(Me.cboWQStdName)
        Me.Controls.Add(Me.txtWQStdDesc)
        Me.Controls.Add(Me._Label1_1)
        Me.Controls.Add(Me._Label1_0)
        Me.Controls.Add(Me.MainMenu1)
        Me.Location = New System.Drawing.Point(9, 27)
        Me.Name = "WaterQualityStandardsForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Water Quality Standards"
        Me.Controls.SetChildIndex(Me.MainMenu1, 0)
        Me.Controls.SetChildIndex(Me._Label1_0, 0)
        Me.Controls.SetChildIndex(Me._Label1_1, 0)
        Me.Controls.SetChildIndex(Me.txtWQStdDesc, 0)
        Me.Controls.SetChildIndex(Me.cboWQStdName, 0)
        Me.Controls.SetChildIndex(Me.dgvWaterQuality, 0)
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        CType(Me.dgvWaterQuality, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvWaterQuality As System.Windows.Forms.DataGridView
    Friend WithEvents colPollutant As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colThreshold As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Poll_WQCritID As System.Windows.Forms.DataGridViewTextBoxColumn
#End Region
End Class