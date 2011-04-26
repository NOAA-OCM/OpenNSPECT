<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class NewCoefficientSetForm
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
    Public WithEvents txtCoeffSetName As System.Windows.Forms.TextBox
    Public WithEvents _Label1_5 As System.Windows.Forms.Label
    Public WithEvents _Label1_7 As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cboLCType = New System.Windows.Forms.ComboBox()
        Me.txtCoeffSetName = New System.Windows.Forms.TextBox()
        Me._Label1_5 = New System.Windows.Forms.Label()
        Me._Label1_7 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'cboLCType
        '
        Me.cboLCType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboLCType.Location = New System.Drawing.Point(154, 43)
        Me.cboLCType.Name = "cboLCType"
        Me.cboLCType.Size = New System.Drawing.Size(173, 21)
        Me.cboLCType.TabIndex = 1
        '
        'txtCoeffSetName
        '
        Me.txtCoeffSetName.AcceptsReturn = True
        Me.txtCoeffSetName.Location = New System.Drawing.Point(154, 14)
        Me.txtCoeffSetName.MaxLength = 0
        Me.txtCoeffSetName.Name = "txtCoeffSetName"
        Me.txtCoeffSetName.Size = New System.Drawing.Size(175, 20)
        Me.txtCoeffSetName.TabIndex = 0
        '
        '_Label1_5
        '
        Me._Label1_5.Location = New System.Drawing.Point(-14, 17)
        Me._Label1_5.Name = "_Label1_5"
        Me._Label1_5.Size = New System.Drawing.Size(150, 16)
        Me._Label1_5.TabIndex = 5
        Me._Label1_5.Text = "Coefficient Set Name:"
        Me._Label1_5.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_7
        '
        Me._Label1_7.Location = New System.Drawing.Point(33, 45)
        Me._Label1_7.Name = "_Label1_7"
        Me._Label1_7.Size = New System.Drawing.Size(102, 16)
        Me._Label1_7.TabIndex = 4
        Me._Label1_7.Text = "Land Cover Type:"
        Me._Label1_7.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'NewCoefficientSetForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(372, 107)
        Me.Controls.Add(Me.cboLCType)
        Me.Controls.Add(Me.txtCoeffSetName)
        Me.Controls.Add(Me._Label1_5)
        Me.Controls.Add(Me._Label1_7)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(519, 231)
        Me.Name = "NewCoefficientSetForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Add Coefficient Set"
        Me.Controls.SetChildIndex(Me._Label1_7, 0)
        Me.Controls.SetChildIndex(Me._Label1_5, 0)
        Me.Controls.SetChildIndex(Me.txtCoeffSetName, 0)
        Me.Controls.SetChildIndex(Me.cboLCType, 0)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
End Class