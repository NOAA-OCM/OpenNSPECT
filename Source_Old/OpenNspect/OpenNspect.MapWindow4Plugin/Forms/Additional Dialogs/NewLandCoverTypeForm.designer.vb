<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class NewLandCoverTypeForm
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
    Public WithEvents txtLCTypeDesc As System.Windows.Forms.TextBox
    Public WithEvents txtLCType As System.Windows.Forms.TextBox
    Public WithEvents _Label1_6 As System.Windows.Forms.Label
    Public WithEvents _Label2_0 As System.Windows.Forms.Label
    Public WithEvents _Label2_1 As System.Windows.Forms.Label
    Public WithEvents _Label2_2 As System.Windows.Forms.Label
    Public WithEvents _Label1_7 As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtLCTypeDesc = New System.Windows.Forms.TextBox()
        Me.txtLCType = New System.Windows.Forms.TextBox()
        Me._Label1_6 = New System.Windows.Forms.Label()
        Me._Label2_0 = New System.Windows.Forms.Label()
        Me._Label2_1 = New System.Windows.Forms.Label()
        Me._Label2_2 = New System.Windows.Forms.Label()
        Me._Label1_7 = New System.Windows.Forms.Label()
        Me.dgvLCTypes = New System.Windows.Forms.DataGridView()
        Me.cntxmnuGrid = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.AddRowToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.InsertRowToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DeleteRowToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Value = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.NameCol = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CNA = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CNB = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CNC = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CND = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CoverFactor = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.WetCheck = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.LCTYPEID = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.LCClassID = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.dgvLCTypes, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.cntxmnuGrid.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtLCTypeDesc
        '
        Me.txtLCTypeDesc.AcceptsReturn = True
        Me.txtLCTypeDesc.Location = New System.Drawing.Point(120, 31)
        Me.txtLCTypeDesc.MaxLength = 0
        Me.txtLCTypeDesc.Name = "txtLCTypeDesc"
        Me.txtLCTypeDesc.Size = New System.Drawing.Size(374, 20)
        Me.txtLCTypeDesc.TabIndex = 1
        '
        'txtLCType
        '
        Me.txtLCType.AcceptsReturn = True
        Me.txtLCType.Location = New System.Drawing.Point(121, 6)
        Me.txtLCType.MaxLength = 0
        Me.txtLCType.Name = "txtLCType"
        Me.txtLCType.Size = New System.Drawing.Size(373, 20)
        Me.txtLCType.TabIndex = 0
        '
        '_Label1_6
        '
        Me._Label1_6.Location = New System.Drawing.Point(25, 31)
        Me._Label1_6.Name = "_Label1_6"
        Me._Label1_6.Size = New System.Drawing.Size(82, 16)
        Me._Label1_6.TabIndex = 8
        Me._Label1_6.Text = "Description:"
        Me._Label1_6.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label2_0
        '
        Me._Label2_0.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me._Label2_0.Location = New System.Drawing.Point(240, 59)
        Me._Label2_0.Name = "_Label2_0"
        Me._Label2_0.Size = New System.Drawing.Size(213, 18)
        Me._Label2_0.TabIndex = 7
        Me._Label2_0.Text = "SCS Curve Numbers"
        Me._Label2_0.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_Label2_1
        '
        Me._Label2_1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me._Label2_1.Location = New System.Drawing.Point(11, 58)
        Me._Label2_1.Name = "_Label2_1"
        Me._Label2_1.Size = New System.Drawing.Size(227, 18)
        Me._Label2_1.TabIndex = 6
        Me._Label2_1.Text = "Classification"
        Me._Label2_1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_Label2_2
        '
        Me._Label2_2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me._Label2_2.Location = New System.Drawing.Point(454, 59)
        Me._Label2_2.Name = "_Label2_2"
        Me._Label2_2.Size = New System.Drawing.Size(160, 18)
        Me._Label2_2.TabIndex = 5
        Me._Label2_2.Text = "RUSLE"
        Me._Label2_2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_Label1_7
        '
        Me._Label1_7.Location = New System.Drawing.Point(13, 7)
        Me._Label1_7.Name = "_Label1_7"
        Me._Label1_7.Size = New System.Drawing.Size(94, 16)
        Me._Label1_7.TabIndex = 4
        Me._Label1_7.Text = "Land Cover Type:"
        Me._Label1_7.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'dgvLCTypes
        '
        Me.dgvLCTypes.AllowUserToAddRows = False
        Me.dgvLCTypes.AllowUserToDeleteRows = False
        Me.dgvLCTypes.AllowUserToResizeColumns = False
        Me.dgvLCTypes.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvLCTypes.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Value, Me.NameCol, Me.CNA, Me.CNB, Me.CNC, Me.CND, Me.CoverFactor, Me.WetCheck, Me.LCTYPEID, Me.LCClassID})
        Me.dgvLCTypes.ContextMenuStrip = Me.cntxmnuGrid
        Me.dgvLCTypes.Location = New System.Drawing.Point(10, 79)
        Me.dgvLCTypes.Name = "dgvLCTypes"
        Me.dgvLCTypes.Size = New System.Drawing.Size(582, 215)
        Me.dgvLCTypes.TabIndex = 16
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
        'Value
        '
        Me.Value.DataPropertyName = "Value"
        Me.Value.HeaderText = "Value"
        Me.Value.Name = "Value"
        Me.Value.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.Value.Width = 42
        '
        'NameCol
        '
        Me.NameCol.DataPropertyName = "Name"
        Me.NameCol.HeaderText = "Name"
        Me.NameCol.Name = "NameCol"
        Me.NameCol.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.NameCol.Width = 145
        '
        'CNA
        '
        Me.CNA.DataPropertyName = "CN-A"
        Me.CNA.HeaderText = "CN-A"
        Me.CNA.Name = "CNA"
        Me.CNA.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.CNA.Width = 53
        '
        'CNB
        '
        Me.CNB.DataPropertyName = "CN-B"
        Me.CNB.HeaderText = "CN-B"
        Me.CNB.Name = "CNB"
        Me.CNB.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.CNB.Width = 53
        '
        'CNC
        '
        Me.CNC.DataPropertyName = "CN-C"
        Me.CNC.HeaderText = "CN-C"
        Me.CNC.Name = "CNC"
        Me.CNC.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.CNC.Width = 53
        '
        'CND
        '
        Me.CND.DataPropertyName = "CN-D"
        Me.CND.HeaderText = "CN-D"
        Me.CND.Name = "CND"
        Me.CND.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.CND.Width = 55
        '
        'CoverFactor
        '
        Me.CoverFactor.DataPropertyName = "CoverFactor"
        Me.CoverFactor.HeaderText = "Cover-Factor"
        Me.CoverFactor.Name = "CoverFactor"
        Me.CoverFactor.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.CoverFactor.Width = 95
        '
        'WetCheck
        '
        Me.WetCheck.DataPropertyName = "W_WL"
        Me.WetCheck.FalseValue = "0"
        Me.WetCheck.HeaderText = "Wet"
        Me.WetCheck.Name = "WetCheck"
        Me.WetCheck.TrueValue = "1"
        Me.WetCheck.Width = 35
        '
        'LCTYPEID
        '
        Me.LCTYPEID.DataPropertyName = "LCTypeID"
        Me.LCTYPEID.HeaderText = "LCTYPEID"
        Me.LCTYPEID.Name = "LCTYPEID"
        Me.LCTYPEID.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.LCTYPEID.Visible = False
        '
        'LCClassID
        '
        Me.LCClassID.DataPropertyName = "LCClassID"
        Me.LCClassID.HeaderText = "LCClassID"
        Me.LCClassID.Name = "LCClassID"
        Me.LCClassID.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.LCClassID.Visible = False
        '
        'NewLandCoverTypeForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(594, 345)
        Me.Controls.Add(Me.dgvLCTypes)
        Me.Controls.Add(Me.txtLCTypeDesc)
        Me.Controls.Add(Me.txtLCType)
        Me.Controls.Add(Me._Label1_6)
        Me.Controls.Add(Me._Label2_0)
        Me.Controls.Add(Me._Label2_1)
        Me.Controls.Add(Me._Label2_2)
        Me.Controls.Add(Me._Label1_7)
        Me.Location = New System.Drawing.Point(3, 21)
        Me.Name = "NewLandCoverTypeForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "New Land Cover Type"
        Me.Controls.SetChildIndex(Me._Label1_7, 0)
        Me.Controls.SetChildIndex(Me._Label2_2, 0)
        Me.Controls.SetChildIndex(Me._Label2_1, 0)
        Me.Controls.SetChildIndex(Me._Label2_0, 0)
        Me.Controls.SetChildIndex(Me._Label1_6, 0)
        Me.Controls.SetChildIndex(Me.txtLCType, 0)
        Me.Controls.SetChildIndex(Me.txtLCTypeDesc, 0)
        Me.Controls.SetChildIndex(Me.dgvLCTypes, 0)
        CType(Me.dgvLCTypes, System.ComponentModel.ISupportInitialize).EndInit()
        Me.cntxmnuGrid.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvLCTypes As System.Windows.Forms.DataGridView
    Friend WithEvents cntxmnuGrid As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents AddRowToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents InsertRowToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DeleteRowToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Value As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NameCol As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CNA As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CNB As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CNC As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CND As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CoverFactor As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents WetCheck As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents LCTYPEID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LCClassID As System.Windows.Forms.DataGridViewTextBoxColumn
#End Region
End Class