<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class CopyWaterQualityStandardForm
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
    Public WithEvents txtStdName As System.Windows.Forms.TextBox
    Public WithEvents cboStdName As System.Windows.Forms.ComboBox
    Public WithEvents _Label1_5 As System.Windows.Forms.Label
    Public WithEvents _Label1_0 As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(CopyWaterQualityStandardForm))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtStdName = New System.Windows.Forms.TextBox()
        Me.cboStdName = New System.Windows.Forms.ComboBox()
        Me._Label1_5 = New System.Windows.Forms.Label()
        Me._Label1_0 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'txtStdName
        '
        Me.txtStdName.AcceptsReturn = True
        Me.txtStdName.Location = New System.Drawing.Point(175, 45)
        Me.txtStdName.MaxLength = 0
        Me.txtStdName.Name = "txtStdName"
        Me.txtStdName.Size = New System.Drawing.Size(158, 20)
        Me.txtStdName.TabIndex = 1
        '
        'cboStdName
        '
        Me.cboStdName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboStdName.Location = New System.Drawing.Point(175, 14)
        Me.cboStdName.Name = "cboStdName"
        Me.cboStdName.Size = New System.Drawing.Size(158, 21)
        Me.cboStdName.TabIndex = 0
        '
        '_Label1_5
        '
        Me._Label1_5.Location = New System.Drawing.Point(30, 47)
        Me._Label1_5.Name = "_Label1_5"
        Me._Label1_5.Size = New System.Drawing.Size(110, 16)
        Me._Label1_5.TabIndex = 5
        Me._Label1_5.Text = "New Standard Name:"
        '
        '_Label1_0
        '
        Me._Label1_0.Location = New System.Drawing.Point(30, 16)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.Size = New System.Drawing.Size(152, 16)
        Me._Label1_0.TabIndex = 4
        Me._Label1_0.Text = "Copy from Standard Name:"
        '
        'CopyWaterQualityStandardForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(367, 107)
        Me.Controls.Add(Me.txtStdName)
        Me.Controls.Add(Me.cboStdName)
        Me.Controls.Add(Me._Label1_5)
        Me.Controls.Add(Me._Label1_0)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(473, 457)
        Me.Name = "CopyWaterQualityStandardForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Copy Water Quality Standard"
        Me.Controls.SetChildIndex(Me._Label1_0, 0)
        Me.Controls.SetChildIndex(Me._Label1_5, 0)
        Me.Controls.SetChildIndex(Me.cboStdName, 0)
        Me.Controls.SetChildIndex(Me.txtStdName, 0)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
End Class