<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class NewWaterQualityStandardForm
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
    Public WithEvents txtWQStdDesc As System.Windows.Forms.TextBox
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
        Me.txtWQStdDesc.Location = New System.Drawing.Point(103, 37)
        Me.txtWQStdDesc.MaxLength = 100
        Me.txtWQStdDesc.Name = "txtWQStdDesc"
        Me.txtWQStdDesc.Size = New System.Drawing.Size(450, 20)
        Me.txtWQStdDesc.TabIndex = 1
        '
        'txtWQStdName
        '
        Me.txtWQStdName.AcceptsReturn = True
        Me.txtWQStdName.Location = New System.Drawing.Point(103, 12)
        Me.txtWQStdName.MaxLength = 0
        Me.txtWQStdName.Name = "txtWQStdName"
        Me.txtWQStdName.Size = New System.Drawing.Size(450, 20)
        Me.txtWQStdName.TabIndex = 0
        '
        '_Label1_1
        '
        Me._Label1_1.Location = New System.Drawing.Point(19, 38)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.Size = New System.Drawing.Size(71, 18)
        Me._Label1_1.TabIndex = 5
        Me._Label1_1.Text = "Description"
        '
        '_Label1_7
        '
        Me._Label1_7.Location = New System.Drawing.Point(19, 16)
        Me._Label1_7.Name = "_Label1_7"
        Me._Label1_7.Size = New System.Drawing.Size(83, 16)
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
        Me.dgvWaterQuality.Location = New System.Drawing.Point(22, 61)
        Me.dgvWaterQuality.Name = "dgvWaterQuality"
        Me.dgvWaterQuality.Size = New System.Drawing.Size(531, 235)
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
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(594, 345)
        Me.Controls.Add(Me.dgvWaterQuality)
        Me.Controls.Add(Me.txtWQStdDesc)
        Me.Controls.Add(Me.txtWQStdName)
        Me.Controls.Add(Me._Label1_1)
        Me.Controls.Add(Me._Label1_7)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(3, 21)
        Me.Name = "NewWaterQualityStandardForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Add Water Quality Standard"
        Me.Controls.SetChildIndex(Me._Label1_7, 0)
        Me.Controls.SetChildIndex(Me._Label1_1, 0)
        Me.Controls.SetChildIndex(Me.txtWQStdName, 0)
        Me.Controls.SetChildIndex(Me.txtWQStdDesc, 0)
        Me.Controls.SetChildIndex(Me.dgvWaterQuality, 0)
        CType(Me.dgvWaterQuality, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvWaterQuality As System.Windows.Forms.DataGridView
    Friend WithEvents Pollutant As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Threshold As System.Windows.Forms.DataGridViewTextBoxColumn
#End Region
End Class