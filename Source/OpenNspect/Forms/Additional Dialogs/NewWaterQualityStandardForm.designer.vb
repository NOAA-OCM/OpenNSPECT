<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class NewWaterQualityStandardForm
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
    Public WithEvents txtWQStdDesc As System.Windows.Forms.TextBox
    Public WithEvents cmdCancel As System.Windows.Forms.Button
    Public WithEvents cmdSave As System.Windows.Forms.Button
    Public WithEvents txtWQStdName As System.Windows.Forms.TextBox
    Public WithEvents _Label1_1 As System.Windows.Forms.Label
    Public WithEvents _Label1_7 As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(NewWaterQualityStandardForm))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtWQStdDesc = New System.Windows.Forms.TextBox()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdSave = New System.Windows.Forms.Button()
        Me.txtWQStdName = New System.Windows.Forms.TextBox()
        Me._Label1_1 = New System.Windows.Forms.Label()
        Me._Label1_7 = New System.Windows.Forms.Label()
        Me.dgvWaterQuality = New System.Windows.Forms.DataGridView()
        Me.Pollutant = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Threshold = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.dgvWaterQuality, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtWQStdDesc
        '
        Me.txtWQStdDesc.AcceptsReturn = True
        Me.txtWQStdDesc.BackColor = System.Drawing.SystemColors.Window
        Me.txtWQStdDesc.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtWQStdDesc.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtWQStdDesc.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtWQStdDesc.Location = New System.Drawing.Point(103, 40)
        Me.txtWQStdDesc.MaxLength = 100
        Me.txtWQStdDesc.Name = "txtWQStdDesc"
        Me.txtWQStdDesc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtWQStdDesc.Size = New System.Drawing.Size(450, 20)
        Me.txtWQStdDesc.TabIndex = 1
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.Location = New System.Drawing.Point(488, 336)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
        Me.cmdCancel.TabIndex = 3
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdSave
        '
        Me.cmdSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdSave.BackColor = System.Drawing.SystemColors.Control
        Me.cmdSave.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdSave.Enabled = False
        Me.cmdSave.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSave.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSave.Location = New System.Drawing.Point(407, 336)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdSave.Size = New System.Drawing.Size(75, 23)
        Me.cmdSave.TabIndex = 2
        Me.cmdSave.Text = "Save"
        Me.cmdSave.UseVisualStyleBackColor = True
        '
        'txtWQStdName
        '
        Me.txtWQStdName.AcceptsReturn = True
        Me.txtWQStdName.BackColor = System.Drawing.SystemColors.Window
        Me.txtWQStdName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtWQStdName.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtWQStdName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtWQStdName.Location = New System.Drawing.Point(103, 13)
        Me.txtWQStdName.MaxLength = 0
        Me.txtWQStdName.Name = "txtWQStdName"
        Me.txtWQStdName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtWQStdName.Size = New System.Drawing.Size(450, 20)
        Me.txtWQStdName.TabIndex = 0
        '
        '_Label1_1
        '
        Me._Label1_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_1.Location = New System.Drawing.Point(19, 41)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_1.Size = New System.Drawing.Size(71, 19)
        Me._Label1_1.TabIndex = 5
        Me._Label1_1.Text = "Description"
        '
        '_Label1_7
        '
        Me._Label1_7.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_7.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_7.Location = New System.Drawing.Point(19, 17)
        Me._Label1_7.Name = "_Label1_7"
        Me._Label1_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_7.Size = New System.Drawing.Size(83, 17)
        Me._Label1_7.TabIndex = 4
        Me._Label1_7.Text = "Standard Name"
        '
        'dgvWaterQuality
        '
        Me.dgvWaterQuality.AllowUserToAddRows = False
        Me.dgvWaterQuality.AllowUserToDeleteRows = False
        Me.dgvWaterQuality.AllowUserToResizeColumns = False
        Me.dgvWaterQuality.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvWaterQuality.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Pollutant, Me.Threshold})
        Me.dgvWaterQuality.Location = New System.Drawing.Point(22, 66)
        Me.dgvWaterQuality.Name = "dgvWaterQuality"
        Me.dgvWaterQuality.Size = New System.Drawing.Size(531, 253)
        Me.dgvWaterQuality.TabIndex = 10
        '
        'Pollutant
        '
        Me.Pollutant.HeaderText = "Pollutant"
        Me.Pollutant.Name = "Pollutant"
        Me.Pollutant.ReadOnly = True
        Me.Pollutant.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.Pollutant.Width = 185
        '
        'Threshold
        '
        Me.Threshold.HeaderText = "Threshold"
        Me.Threshold.Name = "Threshold"
        Me.Threshold.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.Threshold.Width = 130
        '
        'NewWaterQualityStandardForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(594, 372)
        Me.Controls.Add(Me.dgvWaterQuality)
        Me.Controls.Add(Me.txtWQStdDesc)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.txtWQStdName)
        Me.Controls.Add(Me._Label1_1)
        Me.Controls.Add(Me._Label1_7)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(3, 21)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "NewWaterQualityStandardForm"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Add Water Quality Standard"
        CType(Me.dgvWaterQuality, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvWaterQuality As System.Windows.Forms.DataGridView
    Friend WithEvents Pollutant As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Threshold As System.Windows.Forms.DataGridViewTextBoxColumn
#End Region
End Class