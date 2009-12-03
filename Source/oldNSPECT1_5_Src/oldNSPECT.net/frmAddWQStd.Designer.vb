<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmAddWQStd
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
	Public WithEvents txtActiveCell As System.Windows.Forms.TextBox
	Public WithEvents txtWQStdDesc As System.Windows.Forms.TextBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdSave As System.Windows.Forms.Button
	Public WithEvents txtWQStdName As System.Windows.Forms.TextBox
	Public WithEvents grdWQStd As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
	Public WithEvents _Label1_1 As System.Windows.Forms.Label
	Public WithEvents _Label1_7 As System.Windows.Forms.Label
	Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmAddWQStd))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.txtActiveCell = New System.Windows.Forms.TextBox
		Me.txtWQStdDesc = New System.Windows.Forms.TextBox
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cmdSave = New System.Windows.Forms.Button
		Me.txtWQStdName = New System.Windows.Forms.TextBox
		Me.grdWQStd = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
		Me._Label1_1 = New System.Windows.Forms.Label
		Me._Label1_7 = New System.Windows.Forms.Label
		Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.grdWQStd, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "Add Water Quality Standard"
		Me.ClientSize = New System.Drawing.Size(416, 297)
		Me.Location = New System.Drawing.Point(3, 21)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmAddWQStd"
		Me.txtActiveCell.AutoSize = False
		Me.txtActiveCell.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
		Me.txtActiveCell.BackColor = System.Drawing.Color.White
		Me.txtActiveCell.ForeColor = System.Drawing.Color.Black
		Me.txtActiveCell.Size = New System.Drawing.Size(95, 19)
		Me.txtActiveCell.Location = New System.Drawing.Point(81, 265)
		Me.txtActiveCell.TabIndex = 7
		Me.txtActiveCell.Visible = False
		Me.txtActiveCell.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtActiveCell.AcceptsReturn = True
		Me.txtActiveCell.CausesValidation = True
		Me.txtActiveCell.Enabled = True
		Me.txtActiveCell.HideSelection = True
		Me.txtActiveCell.ReadOnly = False
		Me.txtActiveCell.Maxlength = 0
		Me.txtActiveCell.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtActiveCell.MultiLine = False
		Me.txtActiveCell.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtActiveCell.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtActiveCell.TabStop = True
		Me.txtActiveCell.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.txtActiveCell.Name = "txtActiveCell"
		Me.txtWQStdDesc.AutoSize = False
		Me.txtWQStdDesc.Size = New System.Drawing.Size(286, 20)
		Me.txtWQStdDesc.Location = New System.Drawing.Point(103, 40)
		Me.txtWQStdDesc.Maxlength = 100
		Me.txtWQStdDesc.TabIndex = 1
		Me.txtWQStdDesc.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtWQStdDesc.AcceptsReturn = True
		Me.txtWQStdDesc.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtWQStdDesc.BackColor = System.Drawing.SystemColors.Window
		Me.txtWQStdDesc.CausesValidation = True
		Me.txtWQStdDesc.Enabled = True
		Me.txtWQStdDesc.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtWQStdDesc.HideSelection = True
		Me.txtWQStdDesc.ReadOnly = False
		Me.txtWQStdDesc.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtWQStdDesc.MultiLine = False
		Me.txtWQStdDesc.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtWQStdDesc.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtWQStdDesc.TabStop = True
		Me.txtWQStdDesc.Visible = True
		Me.txtWQStdDesc.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtWQStdDesc.Name = "txtWQStdDesc"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(65, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(310, 261)
		Me.cmdCancel.TabIndex = 3
		Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCancel.CausesValidation = True
		Me.cmdCancel.Enabled = True
		Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.TabStop = True
		Me.cmdCancel.Name = "cmdCancel"
		Me.cmdSave.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdSave.Text = "Save"
		Me.cmdSave.Enabled = False
		Me.cmdSave.Size = New System.Drawing.Size(65, 25)
		Me.cmdSave.Location = New System.Drawing.Point(242, 261)
		Me.cmdSave.TabIndex = 2
		Me.cmdSave.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdSave.BackColor = System.Drawing.SystemColors.Control
		Me.cmdSave.CausesValidation = True
		Me.cmdSave.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdSave.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdSave.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdSave.TabStop = True
		Me.cmdSave.Name = "cmdSave"
		Me.txtWQStdName.AutoSize = False
		Me.txtWQStdName.Size = New System.Drawing.Size(134, 20)
		Me.txtWQStdName.Location = New System.Drawing.Point(103, 13)
		Me.txtWQStdName.TabIndex = 0
		Me.txtWQStdName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtWQStdName.AcceptsReturn = True
		Me.txtWQStdName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtWQStdName.BackColor = System.Drawing.SystemColors.Window
		Me.txtWQStdName.CausesValidation = True
		Me.txtWQStdName.Enabled = True
		Me.txtWQStdName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtWQStdName.HideSelection = True
		Me.txtWQStdName.ReadOnly = False
		Me.txtWQStdName.Maxlength = 0
		Me.txtWQStdName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtWQStdName.MultiLine = False
		Me.txtWQStdName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtWQStdName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtWQStdName.TabStop = True
		Me.txtWQStdName.Visible = True
		Me.txtWQStdName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtWQStdName.Name = "txtWQStdName"
		grdWQStd.OcxState = CType(resources.GetObject("grdWQStd.OcxState"), System.Windows.Forms.AxHost.State)
		Me.grdWQStd.Size = New System.Drawing.Size(364, 174)
		Me.grdWQStd.Location = New System.Drawing.Point(21, 76)
		Me.grdWQStd.TabIndex = 6
		Me.grdWQStd.Name = "grdWQStd"
		Me._Label1_1.Text = "Description"
		Me._Label1_1.Size = New System.Drawing.Size(71, 19)
		Me._Label1_1.Location = New System.Drawing.Point(19, 41)
		Me._Label1_1.TabIndex = 5
		Me._Label1_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_1.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_1.Enabled = True
		Me._Label1_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_1.UseMnemonic = True
		Me._Label1_1.Visible = True
		Me._Label1_1.AutoSize = False
		Me._Label1_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_1.Name = "_Label1_1"
		Me._Label1_7.Text = "Standard Name"
		Me._Label1_7.Size = New System.Drawing.Size(83, 17)
		Me._Label1_7.Location = New System.Drawing.Point(19, 17)
		Me._Label1_7.TabIndex = 4
		Me._Label1_7.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_7.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_7.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_7.Enabled = True
		Me._Label1_7.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_7.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_7.UseMnemonic = True
		Me._Label1_7.Visible = True
		Me._Label1_7.AutoSize = False
		Me._Label1_7.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_7.Name = "_Label1_7"
		Me.Controls.Add(txtActiveCell)
		Me.Controls.Add(txtWQStdDesc)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdSave)
		Me.Controls.Add(txtWQStdName)
		Me.Controls.Add(grdWQStd)
		Me.Controls.Add(_Label1_1)
		Me.Controls.Add(_Label1_7)
		Me.Label1.SetIndex(_Label1_1, CType(1, Short))
		Me.Label1.SetIndex(_Label1_7, CType(7, Short))
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.grdWQStd, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class