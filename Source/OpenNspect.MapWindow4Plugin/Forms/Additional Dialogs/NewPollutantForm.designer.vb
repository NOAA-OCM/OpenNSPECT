<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class NewPollutantForm
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
    Public WithEvents mnuCoeffNewSet As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuCoeffCopySet As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuCoeff As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
    Public WithEvents txtPollutant As System.Windows.Forms.TextBox
    Public WithEvents _Label1_6 As System.Windows.Forms.Label
    Public WithEvents _Label1_1 As System.Windows.Forms.Label
    Public WithEvents _Label1_2 As System.Windows.Forms.Label
    Public WithEvents _Label1_3 As System.Windows.Forms.Label
    Public WithEvents _Label1_7 As System.Windows.Forms.Label
    Public WithEvents _Label1_5 As System.Windows.Forms.Label
    Public WithEvents cboLCType As System.Windows.Forms.ComboBox
    Public WithEvents txtCoeffSetDesc As System.Windows.Forms.TextBox
    Public WithEvents txtCoeffSet As System.Windows.Forms.TextBox
    Public WithEvents _SSTab1_TabPage0 As System.Windows.Forms.TabPage
    Public WithEvents SSTab1 As System.Windows.Forms.TabControl
    Public dlgCMD1Open As System.Windows.Forms.OpenFileDialog
    Public dlgCMD1Save As System.Windows.Forms.SaveFileDialog
    Public dlgCMD1Font As System.Windows.Forms.FontDialog
    Public dlgCMD1Color As System.Windows.Forms.ColorDialog
    Public dlgCMD1Print As System.Windows.Forms.PrintDialog
    Public WithEvents _Label1_0 As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip()
        Me.mnuCoeff = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuCoeffNewSet = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuCoeffCopySet = New System.Windows.Forms.ToolStripMenuItem()
        Me.txtPollutant = New System.Windows.Forms.TextBox()
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
        Me._Label1_1 = New System.Windows.Forms.Label()
        Me._Label1_2 = New System.Windows.Forms.Label()
        Me._Label1_3 = New System.Windows.Forms.Label()
        Me._Label1_5 = New System.Windows.Forms.Label()
        Me.cboLCType = New System.Windows.Forms.ComboBox()
        Me.txtCoeffSetDesc = New System.Windows.Forms.TextBox()
        Me.txtCoeffSet = New System.Windows.Forms.TextBox()
        Me._Label1_7 = New System.Windows.Forms.Label()
        Me.dlgCMD1Open = New System.Windows.Forms.OpenFileDialog()
        Me.dlgCMD1Save = New System.Windows.Forms.SaveFileDialog()
        Me.dlgCMD1Font = New System.Windows.Forms.FontDialog()
        Me.dlgCMD1Color = New System.Windows.Forms.ColorDialog()
        Me.dlgCMD1Print = New System.Windows.Forms.PrintDialog()
        Me._Label1_0 = New System.Windows.Forms.Label()
        Me.MainMenu1.SuspendLayout()
        Me.SSTab1.SuspendLayout()
        Me._SSTab1_TabPage0.SuspendLayout()
        CType(Me.dgvCoef, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuCoeff})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(594, 24)
        Me.MainMenu1.TabIndex = 15
        '
        'mnuCoeff
        '
        Me.mnuCoeff.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuCoeffNewSet, Me.mnuCoeffCopySet})
        Me.mnuCoeff.Name = "mnuCoeff"
        Me.mnuCoeff.Size = New System.Drawing.Size(82, 20)
        Me.mnuCoeff.Text = "&Coefficients"
        '
        'mnuCoeffNewSet
        '
        Me.mnuCoeffNewSet.Name = "mnuCoeffNewSet"
        Me.mnuCoeffNewSet.Size = New System.Drawing.Size(191, 22)
        Me.mnuCoeffNewSet.Text = "&Add Coefficient Set..."
        '
        'mnuCoeffCopySet
        '
        Me.mnuCoeffCopySet.Name = "mnuCoeffCopySet"
        Me.mnuCoeffCopySet.Size = New System.Drawing.Size(191, 22)
        Me.mnuCoeffCopySet.Text = "&Copy Coefficient Set..."
        '
        'txtPollutant
        '
        Me.txtPollutant.AcceptsReturn = True
        Me.txtPollutant.Location = New System.Drawing.Point(106, 24)
        Me.txtPollutant.MaxLength = 0
        Me.txtPollutant.Name = "txtPollutant"
        Me.txtPollutant.Size = New System.Drawing.Size(460, 20)
        Me.txtPollutant.TabIndex = 0
        '
        'SSTab1
        '
        Me.SSTab1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SSTab1.Controls.Add(Me._SSTab1_TabPage0)
        Me.SSTab1.ItemSize = New System.Drawing.Size(42, 18)
        Me.SSTab1.Location = New System.Drawing.Point(16, 49)
        Me.SSTab1.Name = "SSTab1"
        Me.SSTab1.SelectedIndex = 1
        Me.SSTab1.Size = New System.Drawing.Size(554, 459)
        Me.SSTab1.TabIndex = 8
        '
        '_SSTab1_TabPage0
        '
        Me._SSTab1_TabPage0.Controls.Add(Me.dgvCoef)
        Me._SSTab1_TabPage0.Controls.Add(Me._Label1_6)
        Me._SSTab1_TabPage0.Controls.Add(Me._Label1_1)
        Me._SSTab1_TabPage0.Controls.Add(Me._Label1_2)
        Me._SSTab1_TabPage0.Controls.Add(Me._Label1_3)
        Me._SSTab1_TabPage0.Controls.Add(Me._Label1_5)
        Me._SSTab1_TabPage0.Controls.Add(Me.cboLCType)
        Me._SSTab1_TabPage0.Controls.Add(Me.txtCoeffSetDesc)
        Me._SSTab1_TabPage0.Controls.Add(Me.txtCoeffSet)
        Me._SSTab1_TabPage0.Controls.Add(Me._Label1_7)
        Me._SSTab1_TabPage0.Location = New System.Drawing.Point(4, 22)
        Me._SSTab1_TabPage0.Name = "_SSTab1_TabPage0"
        Me._SSTab1_TabPage0.Size = New System.Drawing.Size(546, 433)
        Me._SSTab1_TabPage0.TabIndex = 0
        Me._SSTab1_TabPage0.Text = "Coefficients"
        '
        'dgvCoef
        '
        Me.dgvCoef.AllowUserToAddRows = False
        Me.dgvCoef.AllowUserToDeleteRows = False
        Me.dgvCoef.AllowUserToResizeColumns = False
        Me.dgvCoef.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.dgvCoef.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvCoef.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewTextBoxColumn1, Me.DataGridViewTextBoxColumn2, Me.DataGridViewTextBoxColumn3, Me.DataGridViewTextBoxColumn4, Me.DataGridViewTextBoxColumn5, Me.DataGridViewTextBoxColumn6, Me.DataGridViewTextBoxColumn7, Me.DataGridViewTextBoxColumn8})
        Me.dgvCoef.Location = New System.Drawing.Point(15, 100)
        Me.dgvCoef.Name = "dgvCoef"
        Me.dgvCoef.Size = New System.Drawing.Size(489, 320)
        Me.dgvCoef.TabIndex = 21
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.HeaderText = "Value"
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        Me.DataGridViewTextBoxColumn1.ReadOnly = True
        Me.DataGridViewTextBoxColumn1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.DataGridViewTextBoxColumn1.Width = 53
        '
        'DataGridViewTextBoxColumn2
        '
        Me.DataGridViewTextBoxColumn2.HeaderText = "Name"
        Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
        Me.DataGridViewTextBoxColumn2.ReadOnly = True
        Me.DataGridViewTextBoxColumn2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.DataGridViewTextBoxColumn2.Width = 165
        '
        'DataGridViewTextBoxColumn3
        '
        Me.DataGridViewTextBoxColumn3.HeaderText = "Type 1"
        Me.DataGridViewTextBoxColumn3.Name = "DataGridViewTextBoxColumn3"
        Me.DataGridViewTextBoxColumn3.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.DataGridViewTextBoxColumn3.Width = 53
        '
        'DataGridViewTextBoxColumn4
        '
        Me.DataGridViewTextBoxColumn4.HeaderText = "Type 2"
        Me.DataGridViewTextBoxColumn4.Name = "DataGridViewTextBoxColumn4"
        Me.DataGridViewTextBoxColumn4.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.DataGridViewTextBoxColumn4.Width = 53
        '
        'DataGridViewTextBoxColumn5
        '
        Me.DataGridViewTextBoxColumn5.HeaderText = "Type 3"
        Me.DataGridViewTextBoxColumn5.Name = "DataGridViewTextBoxColumn5"
        Me.DataGridViewTextBoxColumn5.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.DataGridViewTextBoxColumn5.Width = 53
        '
        'DataGridViewTextBoxColumn6
        '
        Me.DataGridViewTextBoxColumn6.HeaderText = "Type 4"
        Me.DataGridViewTextBoxColumn6.Name = "DataGridViewTextBoxColumn6"
        Me.DataGridViewTextBoxColumn6.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.DataGridViewTextBoxColumn6.Width = 53
        '
        'DataGridViewTextBoxColumn7
        '
        Me.DataGridViewTextBoxColumn7.HeaderText = "SetID"
        Me.DataGridViewTextBoxColumn7.Name = "DataGridViewTextBoxColumn7"
        Me.DataGridViewTextBoxColumn7.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.DataGridViewTextBoxColumn7.Visible = False
        '
        'DataGridViewTextBoxColumn8
        '
        Me.DataGridViewTextBoxColumn8.HeaderText = "LCClassID"
        Me.DataGridViewTextBoxColumn8.Name = "DataGridViewTextBoxColumn8"
        Me.DataGridViewTextBoxColumn8.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.DataGridViewTextBoxColumn8.Visible = False
        '
        '_Label1_6
        '
        Me._Label1_6.Location = New System.Drawing.Point(18, 58)
        Me._Label1_6.Name = "_Label1_6"
        Me._Label1_6.Size = New System.Drawing.Size(74, 16)
        Me._Label1_6.TabIndex = 9
        Me._Label1_6.Text = "Description:"
        '
        '_Label1_1
        '
        Me._Label1_1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me._Label1_1.Location = New System.Drawing.Point(57, 82)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.Size = New System.Drawing.Size(215, 16)
        Me._Label1_1.TabIndex = 10
        Me._Label1_1.Text = "Class"
        Me._Label1_1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_Label1_2
        '
        Me._Label1_2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me._Label1_2.Location = New System.Drawing.Point(278, 82)
        Me._Label1_2.Name = "_Label1_2"
        Me._Label1_2.Size = New System.Drawing.Size(226, 16)
        Me._Label1_2.TabIndex = 11
        Me._Label1_2.Text = "Coefficients (mg/L)"
        Me._Label1_2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_Label1_3
        '
        Me._Label1_3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me._Label1_3.Location = New System.Drawing.Point(15, 82)
        Me._Label1_3.Name = "_Label1_3"
        Me._Label1_3.Size = New System.Drawing.Size(36, 16)
        Me._Label1_3.TabIndex = 12
        Me._Label1_3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_Label1_5
        '
        Me._Label1_5.Location = New System.Drawing.Point(18, 34)
        Me._Label1_5.Name = "_Label1_5"
        Me._Label1_5.Size = New System.Drawing.Size(78, 16)
        Me._Label1_5.TabIndex = 17
        Me._Label1_5.Text = "Coefficient Set:"
        '
        'cboLCType
        '
        Me.cboLCType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboLCType.Location = New System.Drawing.Point(350, 30)
        Me.cboLCType.Name = "cboLCType"
        Me.cboLCType.Size = New System.Drawing.Size(147, 21)
        Me.cboLCType.Sorted = True
        Me.cboLCType.TabIndex = 2
        '
        'txtCoeffSetDesc
        '
        Me.txtCoeffSetDesc.AcceptsReturn = True
        Me.txtCoeffSetDesc.Location = New System.Drawing.Point(100, 58)
        Me.txtCoeffSetDesc.MaxLength = 0
        Me.txtCoeffSetDesc.Name = "txtCoeffSetDesc"
        Me.txtCoeffSetDesc.Size = New System.Drawing.Size(399, 20)
        Me.txtCoeffSetDesc.TabIndex = 3
        '
        'txtCoeffSet
        '
        Me.txtCoeffSet.AcceptsReturn = True
        Me.txtCoeffSet.Location = New System.Drawing.Point(100, 32)
        Me.txtCoeffSet.MaxLength = 0
        Me.txtCoeffSet.Name = "txtCoeffSet"
        Me.txtCoeffSet.Size = New System.Drawing.Size(134, 20)
        Me.txtCoeffSet.TabIndex = 1
        '
        '_Label1_7
        '
        Me._Label1_7.Location = New System.Drawing.Point(257, 32)
        Me._Label1_7.Name = "_Label1_7"
        Me._Label1_7.Size = New System.Drawing.Size(97, 16)
        Me._Label1_7.TabIndex = 16
        Me._Label1_7.Text = "Land Cover Type:"
        '
        '_Label1_0
        '
        Me._Label1_0.Location = New System.Drawing.Point(21, 28)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.Size = New System.Drawing.Size(86, 16)
        Me._Label1_0.TabIndex = 5
        Me._Label1_0.Text = "Pollutant Name:"
        '
        'NewPollutantForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(594, 559)
        Me.Controls.Add(Me.txtPollutant)
        Me.Controls.Add(Me.SSTab1)
        Me.Controls.Add(Me._Label1_0)
        Me.Controls.Add(Me.MainMenu1)
        Me.Location = New System.Drawing.Point(268, 127)
        Me.Name = "NewPollutantForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Add Pollutant"
        Me.Controls.SetChildIndex(Me.MainMenu1, 0)
        Me.Controls.SetChildIndex(Me._Label1_0, 0)
        Me.Controls.SetChildIndex(Me.SSTab1, 0)
        Me.Controls.SetChildIndex(Me.txtPollutant, 0)
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        Me.SSTab1.ResumeLayout(False)
        Me._SSTab1_TabPage0.ResumeLayout(False)
        Me._SSTab1_TabPage0.PerformLayout()
        CType(Me.dgvCoef, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvCoef As System.Windows.Forms.DataGridView
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