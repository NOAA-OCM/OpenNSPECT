<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class ImportWaterQualityStandardForm
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
    Public WithEvents cmdBrowse As System.Windows.Forms.Button
    Public WithEvents txtImpFile As System.Windows.Forms.TextBox
    Public WithEvents txtStdName As System.Windows.Forms.TextBox
    Public WithEvents _Label1_0 As System.Windows.Forms.Label
    Public WithEvents _Label1_5 As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ImportWaterQualityStandardForm))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdBrowse = New System.Windows.Forms.Button()
        Me.txtImpFile = New System.Windows.Forms.TextBox()
        Me.txtStdName = New System.Windows.Forms.TextBox()
        Me._Label1_0 = New System.Windows.Forms.Label()
        Me._Label1_5 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'cmdBrowse
        '
        Me.cmdBrowse.Image = CType(resources.GetObject("cmdBrowse.Image"), System.Drawing.Image)
        Me.cmdBrowse.Location = New System.Drawing.Point(326, 43)
        Me.cmdBrowse.Name = "cmdBrowse"
        Me.cmdBrowse.Size = New System.Drawing.Size(25, 19)
        Me.cmdBrowse.TabIndex = 2
        Me.cmdBrowse.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdBrowse.UseVisualStyleBackColor = True
        '
        'txtImpFile
        '
        Me.txtImpFile.AcceptsReturn = True
        Me.txtImpFile.Location = New System.Drawing.Point(123, 44)
        Me.txtImpFile.MaxLength = 0
        Me.txtImpFile.Name = "txtImpFile"
        Me.txtImpFile.Size = New System.Drawing.Size(200, 20)
        Me.txtImpFile.TabIndex = 1
        '
        'txtStdName
        '
        Me.txtStdName.AcceptsReturn = True
        Me.txtStdName.Location = New System.Drawing.Point(123, 11)
        Me.txtStdName.MaxLength = 0
        Me.txtStdName.Name = "txtStdName"
        Me.txtStdName.Size = New System.Drawing.Size(162, 20)
        Me.txtStdName.TabIndex = 0
        '
        '_Label1_0
        '
        Me._Label1_0.Location = New System.Drawing.Point(17, 46)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.Size = New System.Drawing.Size(100, 16)
        Me._Label1_0.TabIndex = 6
        Me._Label1_0.Text = "Import File:"
        Me._Label1_0.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_5
        '
        Me._Label1_5.Location = New System.Drawing.Point(6, 15)
        Me._Label1_5.Name = "_Label1_5"
        Me._Label1_5.Size = New System.Drawing.Size(111, 16)
        Me._Label1_5.TabIndex = 5
        Me._Label1_5.Text = "New Standard Name:"
        Me._Label1_5.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'ImportWaterQualityStandardForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Nothing
        Me.ClientSize = New System.Drawing.Size(367, 107)
        Me.Controls.Add(Me.cmdBrowse)
        Me.Controls.Add(Me.txtImpFile)
        Me.Controls.Add(Me.txtStdName)
        Me.Controls.Add(Me._Label1_0)
        Me.Controls.Add(Me._Label1_5)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(3, 21)
        Me.Name = "ImportWaterQualityStandardForm"
        Me.Text = "Import Water Quality Standard"
        Me.Controls.SetChildIndex(Me._Label1_5, 0)
        Me.Controls.SetChildIndex(Me._Label1_0, 0)
        Me.Controls.SetChildIndex(Me.txtStdName, 0)
        Me.Controls.SetChildIndex(Me.txtImpFile, 0)
        Me.Controls.SetChildIndex(Me.cmdBrowse, 0)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
End Class