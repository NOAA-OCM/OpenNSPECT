<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CopyCoefficientSetForm
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
    Public WithEvents cboCoeffSet As System.Windows.Forms.ComboBox
    Public WithEvents txtCoeffSetName As System.Windows.Forms.TextBox
    Public WithEvents _Label1_0 As System.Windows.Forms.Label
    Public WithEvents _Label1_5 As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cboCoeffSet = New System.Windows.Forms.ComboBox()
        Me.txtCoeffSetName = New System.Windows.Forms.TextBox()
        Me._Label1_0 = New System.Windows.Forms.Label()
        Me._Label1_5 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'cboCoeffSet
        '
        Me.cboCoeffSet.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCoeffSet.Location = New System.Drawing.Point(176, 12)
        Me.cboCoeffSet.Name = "cboCoeffSet"
        Me.cboCoeffSet.Size = New System.Drawing.Size(158, 21)
        Me.cboCoeffSet.TabIndex = 0
        '
        'txtCoeffSetName
        '
        Me.txtCoeffSetName.AcceptsReturn = True
        Me.txtCoeffSetName.Location = New System.Drawing.Point(176, 43)
        Me.txtCoeffSetName.MaxLength = 0
        Me.txtCoeffSetName.Name = "txtCoeffSetName"
        Me.txtCoeffSetName.Size = New System.Drawing.Size(156, 20)
        Me.txtCoeffSetName.TabIndex = 1
        '
        '_Label1_0
        '
        Me._Label1_0.Location = New System.Drawing.Point(0, 17)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.Size = New System.Drawing.Size(160, 16)
        Me._Label1_0.TabIndex = 5
        Me._Label1_0.Text = "Copy from Coefficient Set:"
        Me._Label1_0.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_5
        '
        Me._Label1_5.Location = New System.Drawing.Point(24, 46)
        Me._Label1_5.Name = "_Label1_5"
        Me._Label1_5.Size = New System.Drawing.Size(140, 16)
        Me._Label1_5.TabIndex = 4
        Me._Label1_5.Text = "New Coefficient Set Name:"
        Me._Label1_5.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'CopyCoefficientSetForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(367, 107)
        Me.Controls.Add(Me.cboCoeffSet)
        Me.Controls.Add(Me.txtCoeffSetName)
        Me.Controls.Add(Me._Label1_0)
        Me.Controls.Add(Me._Label1_5)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(3, 21)
        Me.Name = "CopyCoefficientSetForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Copy Coefficient Set"
        Me.Controls.SetChildIndex(Me._Label1_5, 0)
        Me.Controls.SetChildIndex(Me._Label1_0, 0)
        Me.Controls.SetChildIndex(Me.txtCoeffSetName, 0)
        Me.Controls.SetChildIndex(Me.cboCoeffSet, 0)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
End Class