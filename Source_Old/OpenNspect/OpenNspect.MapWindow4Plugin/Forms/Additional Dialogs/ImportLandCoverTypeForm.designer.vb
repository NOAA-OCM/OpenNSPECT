<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class ImportLandCoverTypeForm
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
    Public WithEvents txtLCType As System.Windows.Forms.TextBox
    Public WithEvents txtImpFile As System.Windows.Forms.TextBox
    Public WithEvents cmdBrowse As System.Windows.Forms.Button
    Public WithEvents _Label1_5 As System.Windows.Forms.Label
    Public WithEvents _Label1_0 As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ImportLandCoverTypeForm))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtLCType = New System.Windows.Forms.TextBox()
        Me.txtImpFile = New System.Windows.Forms.TextBox()
        Me.cmdBrowse = New System.Windows.Forms.Button()
        Me._Label1_5 = New System.Windows.Forms.Label()
        Me._Label1_0 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'txtLCType
        '
        Me.txtLCType.AcceptsReturn = True
        Me.txtLCType.Location = New System.Drawing.Point(164, 16)
        Me.txtLCType.MaxLength = 0
        Me.txtLCType.Name = "txtLCType"
        Me.txtLCType.Size = New System.Drawing.Size(163, 20)
        Me.txtLCType.TabIndex = 1
        '
        'txtImpFile
        '
        Me.txtImpFile.AcceptsReturn = True
        Me.txtImpFile.Location = New System.Drawing.Point(91, 45)
        Me.txtImpFile.MaxLength = 0
        Me.txtImpFile.Name = "txtImpFile"
        Me.txtImpFile.Size = New System.Drawing.Size(236, 20)
        Me.txtImpFile.TabIndex = 2
        '
        'cmdBrowse
        '
        Me.cmdBrowse.Image = CType(resources.GetObject("cmdBrowse.Image"), System.Drawing.Image)
        Me.cmdBrowse.Location = New System.Drawing.Point(331, 45)
        Me.cmdBrowse.Name = "cmdBrowse"
        Me.cmdBrowse.Size = New System.Drawing.Size(25, 21)
        Me.cmdBrowse.TabIndex = 0
        Me.cmdBrowse.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdBrowse.UseVisualStyleBackColor = True
        '
        '_Label1_5
        '
        Me._Label1_5.Location = New System.Drawing.Point(13, 19)
        Me._Label1_5.Name = "_Label1_5"
        Me._Label1_5.Size = New System.Drawing.Size(126, 16)
        Me._Label1_5.TabIndex = 6
        Me._Label1_5.Text = "Land Cover Type Name:"
        '
        '_Label1_0
        '
        Me._Label1_0.Location = New System.Drawing.Point(15, 46)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.Size = New System.Drawing.Size(89, 16)
        Me._Label1_0.TabIndex = 5
        Me._Label1_0.Text = "Import File:"
        '
        'ImportLandCoverTypeForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(367, 109)
        Me.Controls.Add(Me.txtLCType)
        Me.Controls.Add(Me.txtImpFile)
        Me.Controls.Add(Me.cmdBrowse)
        Me.Controls.Add(Me._Label1_5)
        Me.Controls.Add(Me._Label1_0)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(3, 21)
        Me.Name = "ImportLandCoverTypeForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Import Land Cover Type"
        Me.Controls.SetChildIndex(Me._Label1_0, 0)
        Me.Controls.SetChildIndex(Me._Label1_5, 0)
        Me.Controls.SetChildIndex(Me.cmdBrowse, 0)
        Me.Controls.SetChildIndex(Me.txtImpFile, 0)
        Me.Controls.SetChildIndex(Me.txtLCType, 0)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
End Class