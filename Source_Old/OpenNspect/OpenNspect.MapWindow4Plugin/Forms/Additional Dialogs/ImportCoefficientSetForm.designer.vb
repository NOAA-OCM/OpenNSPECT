<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class ImportCoefficientSetForm
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
    Public WithEvents cboLCType As System.Windows.Forms.ComboBox
    Public WithEvents cmdBrowse As System.Windows.Forms.Button
    Public WithEvents txtImpFile As System.Windows.Forms.TextBox
    Public WithEvents txtCoeffSetName As System.Windows.Forms.TextBox
    Public WithEvents _Label1_5 As System.Windows.Forms.Label
    Public WithEvents _Label1_0 As System.Windows.Forms.Label
    Public WithEvents _Label1_7 As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ImportCoefficientSetForm))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cboLCType = New System.Windows.Forms.ComboBox()
        Me.cmdBrowse = New System.Windows.Forms.Button()
        Me.txtImpFile = New System.Windows.Forms.TextBox()
        Me.txtCoeffSetName = New System.Windows.Forms.TextBox()
        Me._Label1_5 = New System.Windows.Forms.Label()
        Me._Label1_0 = New System.Windows.Forms.Label()
        Me._Label1_7 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'cboLCType
        '
        Me.cboLCType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboLCType.Location = New System.Drawing.Point(162, 36)
        Me.cboLCType.Name = "cboLCType"
        Me.cboLCType.Size = New System.Drawing.Size(223, 21)
        Me.cboLCType.TabIndex = 1
        '
        'cmdBrowse
        '
        Me.cmdBrowse.Image = CType(resources.GetObject("cmdBrowse.Image"), System.Drawing.Image)
        Me.cmdBrowse.Location = New System.Drawing.Point(388, 63)
        Me.cmdBrowse.Name = "cmdBrowse"
        Me.cmdBrowse.Size = New System.Drawing.Size(25, 21)
        Me.cmdBrowse.TabIndex = 3
        Me.cmdBrowse.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdBrowse.UseVisualStyleBackColor = True
        '
        'txtImpFile
        '
        Me.txtImpFile.AcceptsReturn = True
        Me.txtImpFile.Location = New System.Drawing.Point(161, 64)
        Me.txtImpFile.MaxLength = 0
        Me.txtImpFile.Name = "txtImpFile"
        Me.txtImpFile.Size = New System.Drawing.Size(224, 20)
        Me.txtImpFile.TabIndex = 2
        '
        'txtCoeffSetName
        '
        Me.txtCoeffSetName.AcceptsReturn = True
        Me.txtCoeffSetName.Location = New System.Drawing.Point(162, 11)
        Me.txtCoeffSetName.MaxLength = 0
        Me.txtCoeffSetName.Name = "txtCoeffSetName"
        Me.txtCoeffSetName.Size = New System.Drawing.Size(222, 20)
        Me.txtCoeffSetName.TabIndex = 0
        '
        '_Label1_5
        '
        Me._Label1_5.Location = New System.Drawing.Point(-48, 14)
        Me._Label1_5.Name = "_Label1_5"
        Me._Label1_5.Size = New System.Drawing.Size(189, 16)
        Me._Label1_5.TabIndex = 8
        Me._Label1_5.Text = "New Coefficient Set Name:"
        Me._Label1_5.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_0
        '
        Me._Label1_0.Location = New System.Drawing.Point(7, 64)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.Size = New System.Drawing.Size(133, 16)
        Me._Label1_0.TabIndex = 7
        Me._Label1_0.Text = "Import File:"
        Me._Label1_0.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_7
        '
        Me._Label1_7.Location = New System.Drawing.Point(14, 39)
        Me._Label1_7.Name = "_Label1_7"
        Me._Label1_7.Size = New System.Drawing.Size(127, 16)
        Me._Label1_7.TabIndex = 6
        Me._Label1_7.Text = "Land Cover Type:"
        Me._Label1_7.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'ImportCoefficientSetForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(436, 123)
        Me.Controls.Add(Me.cboLCType)
        Me.Controls.Add(Me.cmdBrowse)
        Me.Controls.Add(Me.txtImpFile)
        Me.Controls.Add(Me.txtCoeffSetName)
        Me.Controls.Add(Me._Label1_5)
        Me.Controls.Add(Me._Label1_0)
        Me.Controls.Add(Me._Label1_7)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(3, 21)
        Me.Name = "ImportCoefficientSetForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Import Coefficient Set"
        Me.Controls.SetChildIndex(Me._Label1_7, 0)
        Me.Controls.SetChildIndex(Me._Label1_0, 0)
        Me.Controls.SetChildIndex(Me._Label1_5, 0)
        Me.Controls.SetChildIndex(Me.txtCoeffSetName, 0)
        Me.Controls.SetChildIndex(Me.txtImpFile, 0)
        Me.Controls.SetChildIndex(Me.cmdBrowse, 0)
        Me.Controls.SetChildIndex(Me.cboLCType, 0)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
End Class