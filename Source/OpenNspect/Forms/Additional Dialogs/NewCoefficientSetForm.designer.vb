<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class NewCoefficientSetForm
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
    Public WithEvents cmdCancel As System.Windows.Forms.Button
    Public WithEvents cmdOK As System.Windows.Forms.Button
    Public WithEvents _Label1_5 As System.Windows.Forms.Label
    Public WithEvents _Label1_7 As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(NewCoefficientSetForm))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cboLCType = New System.Windows.Forms.ComboBox()
        Me.txtCoeffSetName = New System.Windows.Forms.TextBox()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me._Label1_5 = New System.Windows.Forms.Label()
        Me._Label1_7 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'cboLCType
        '
        Me.cboLCType.BackColor = System.Drawing.SystemColors.Window
        Me.cboLCType.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboLCType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboLCType.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboLCType.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboLCType.Location = New System.Drawing.Point(154, 46)
        Me.cboLCType.Name = "cboLCType"
        Me.cboLCType.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboLCType.Size = New System.Drawing.Size(173, 22)
        Me.cboLCType.TabIndex = 1
        '
        'txtCoeffSetName
        '
        Me.txtCoeffSetName.AcceptsReturn = True
        Me.txtCoeffSetName.BackColor = System.Drawing.SystemColors.Window
        Me.txtCoeffSetName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCoeffSetName.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCoeffSetName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCoeffSetName.Location = New System.Drawing.Point(154, 15)
        Me.txtCoeffSetName.MaxLength = 0
        Me.txtCoeffSetName.Name = "txtCoeffSetName"
        Me.txtCoeffSetName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCoeffSetName.Size = New System.Drawing.Size(175, 20)
        Me.txtCoeffSetName.TabIndex = 0
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.Location = New System.Drawing.Point(275, 84)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
        Me.cmdCancel.TabIndex = 3
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdOK
        '
        Me.cmdOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOK.Enabled = False
        Me.cmdOK.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOK.Location = New System.Drawing.Point(194, 84)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOK.Size = New System.Drawing.Size(75, 23)
        Me.cmdOK.TabIndex = 2
        Me.cmdOK.Text = "OK"
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        '_Label1_5
        '
        Me._Label1_5.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_5.Location = New System.Drawing.Point(-14, 18)
        Me._Label1_5.Name = "_Label1_5"
        Me._Label1_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_5.Size = New System.Drawing.Size(150, 17)
        Me._Label1_5.TabIndex = 5
        Me._Label1_5.Text = "Coefficient Set Name:"
        Me._Label1_5.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_7
        '
        Me._Label1_7.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_7.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_7.Location = New System.Drawing.Point(33, 49)
        Me._Label1_7.Name = "_Label1_7"
        Me._Label1_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_7.Size = New System.Drawing.Size(102, 17)
        Me._Label1_7.TabIndex = 4
        Me._Label1_7.Text = "Land Cover Type:"
        Me._Label1_7.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'NewCoefficientSetForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(372, 115)
        Me.Controls.Add(Me.cboLCType)
        Me.Controls.Add(Me.txtCoeffSetName)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me._Label1_5)
        Me.Controls.Add(Me._Label1_7)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(519, 231)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "NewCoefficientSetForm"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Add Coefficient Set"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
End Class