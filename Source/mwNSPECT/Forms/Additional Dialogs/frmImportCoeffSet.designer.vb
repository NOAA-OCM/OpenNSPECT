<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmImportCoeffSet
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
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents txtImpFile As System.Windows.Forms.TextBox
	Public WithEvents txtCoeffSetName As System.Windows.Forms.TextBox
	Public dlgCMD1Open As System.Windows.Forms.OpenFileDialog
	Public WithEvents _Label1_5 As System.Windows.Forms.Label
	Public WithEvents _Label1_0 As System.Windows.Forms.Label
	Public WithEvents _Label1_7 As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmImportCoeffSet))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cboLCType = New System.Windows.Forms.ComboBox
        Me.cmdBrowse = New System.Windows.Forms.Button
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmdOK = New System.Windows.Forms.Button
        Me.txtImpFile = New System.Windows.Forms.TextBox
        Me.txtCoeffSetName = New System.Windows.Forms.TextBox
        Me.dlgCMD1Open = New System.Windows.Forms.OpenFileDialog
        Me._Label1_5 = New System.Windows.Forms.Label
        Me._Label1_0 = New System.Windows.Forms.Label
        Me._Label1_7 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'cboLCType
        '
        Me.cboLCType.BackColor = System.Drawing.SystemColors.Window
        Me.cboLCType.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboLCType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboLCType.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboLCType.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboLCType.Location = New System.Drawing.Point(162, 39)
        Me.cboLCType.Name = "cboLCType"
        Me.cboLCType.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboLCType.Size = New System.Drawing.Size(147, 22)
        Me.cboLCType.TabIndex = 1
        '
        'cmdBrowse
        '
        Me.cmdBrowse.BackColor = System.Drawing.SystemColors.Control
        Me.cmdBrowse.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdBrowse.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowse.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdBrowse.Image = CType(resources.GetObject("cmdBrowse.Image"), System.Drawing.Image)
        Me.cmdBrowse.Location = New System.Drawing.Point(388, 68)
        Me.cmdBrowse.Name = "cmdBrowse"
        Me.cmdBrowse.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdBrowse.Size = New System.Drawing.Size(25, 21)
        Me.cmdBrowse.TabIndex = 3
        Me.cmdBrowse.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdBrowse.UseVisualStyleBackColor = False
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.Location = New System.Drawing.Point(342, 100)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancel.Size = New System.Drawing.Size(65, 25)
        Me.cmdCancel.TabIndex = 5
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'cmdOK
        '
        Me.cmdOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOK.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOK.Location = New System.Drawing.Point(273, 100)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOK.Size = New System.Drawing.Size(65, 25)
        Me.cmdOK.TabIndex = 4
        Me.cmdOK.Text = "OK"
        Me.cmdOK.UseVisualStyleBackColor = False
        '
        'txtImpFile
        '
        Me.txtImpFile.AcceptsReturn = True
        Me.txtImpFile.BackColor = System.Drawing.SystemColors.Window
        Me.txtImpFile.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtImpFile.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtImpFile.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtImpFile.Location = New System.Drawing.Point(161, 69)
        Me.txtImpFile.MaxLength = 0
        Me.txtImpFile.Name = "txtImpFile"
        Me.txtImpFile.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtImpFile.Size = New System.Drawing.Size(224, 20)
        Me.txtImpFile.TabIndex = 2
        '
        'txtCoeffSetName
        '
        Me.txtCoeffSetName.AcceptsReturn = True
        Me.txtCoeffSetName.BackColor = System.Drawing.SystemColors.Window
        Me.txtCoeffSetName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCoeffSetName.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCoeffSetName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCoeffSetName.Location = New System.Drawing.Point(162, 12)
        Me.txtCoeffSetName.MaxLength = 0
        Me.txtCoeffSetName.Name = "txtCoeffSetName"
        Me.txtCoeffSetName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCoeffSetName.Size = New System.Drawing.Size(143, 20)
        Me.txtCoeffSetName.TabIndex = 0
        '
        '_Label1_5
        '
        Me._Label1_5.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_5.Location = New System.Drawing.Point(-48, 15)
        Me._Label1_5.Name = "_Label1_5"
        Me._Label1_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_5.Size = New System.Drawing.Size(189, 17)
        Me._Label1_5.TabIndex = 8
        Me._Label1_5.Text = "New Coefficient Set Name:"
        Me._Label1_5.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_0
        '
        Me._Label1_0.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_0.Location = New System.Drawing.Point(7, 69)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_0.Size = New System.Drawing.Size(133, 17)
        Me._Label1_0.TabIndex = 7
        Me._Label1_0.Text = "Import File:"
        Me._Label1_0.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_7
        '
        Me._Label1_7.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_7.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_7.Location = New System.Drawing.Point(14, 42)
        Me._Label1_7.Name = "_Label1_7"
        Me._Label1_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_7.Size = New System.Drawing.Size(127, 17)
        Me._Label1_7.TabIndex = 6
        Me._Label1_7.Text = "Land Cover Type:"
        Me._Label1_7.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'frmImportCoeffSet
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(436, 132)
        Me.Controls.Add(Me.cboLCType)
        Me.Controls.Add(Me.cmdBrowse)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.txtImpFile)
        Me.Controls.Add(Me.txtCoeffSetName)
        Me.Controls.Add(Me._Label1_5)
        Me.Controls.Add(Me._Label1_0)
        Me.Controls.Add(Me._Label1_7)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(3, 21)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmImportCoeffSet"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Import Coefficient Set"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region 
End Class