<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmNewLCType
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
	Public WithEvents mnuAppendRow As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuInsertRow As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuDeleteRow As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopUp As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
	Public WithEvents _chkWWL_0 As System.Windows.Forms.CheckBox
	Public WithEvents txtActiveCell As System.Windows.Forms.TextBox
	Public WithEvents grdLcClasses As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
	Public WithEvents txtLCTypeDesc As System.Windows.Forms.TextBox
	Public WithEvents txtLCType As System.Windows.Forms.TextBox
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents _Label1_6 As System.Windows.Forms.Label
	Public WithEvents _Label2_0 As System.Windows.Forms.Label
	Public WithEvents _Label2_1 As System.Windows.Forms.Label
	Public WithEvents _Label2_2 As System.Windows.Forms.Label
	Public WithEvents _Label1_7 As System.Windows.Forms.Label
	Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents Label2 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents chkWWL As Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmNewLCType))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.MainMenu1 = New System.Windows.Forms.MenuStrip
		Me.mnuPopUp = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuAppendRow = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuInsertRow = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuDeleteRow = New System.Windows.Forms.ToolStripMenuItem
		Me._chkWWL_0 = New System.Windows.Forms.CheckBox
		Me.txtActiveCell = New System.Windows.Forms.TextBox
		Me.grdLcClasses = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
		Me.txtLCTypeDesc = New System.Windows.Forms.TextBox
		Me.txtLCType = New System.Windows.Forms.TextBox
		Me.cmdOK = New System.Windows.Forms.Button
		Me.cmdCancel = New System.Windows.Forms.Button
		Me._Label1_6 = New System.Windows.Forms.Label
		Me._Label2_0 = New System.Windows.Forms.Label
		Me._Label2_1 = New System.Windows.Forms.Label
		Me._Label2_2 = New System.Windows.Forms.Label
		Me._Label1_7 = New System.Windows.Forms.Label
		Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.Label2 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.chkWWL = New Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray(components)
		Me.MainMenu1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.grdLcClasses, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Label2, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.chkWWL, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "New Land Cover Type"
		Me.ClientSize = New System.Drawing.Size(623, 397)
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
		Me.Name = "frmNewLCType"
		Me.mnuPopUp.Name = "mnuPopUp"
		Me.mnuPopUp.Text = "Edit"
		Me.mnuPopUp.Visible = False
		Me.mnuPopUp.Checked = False
		Me.mnuPopUp.Enabled = True
		Me.mnuAppendRow.Name = "mnuAppendRow"
		Me.mnuAppendRow.Text = "Add Row"
		Me.mnuAppendRow.Checked = False
		Me.mnuAppendRow.Enabled = True
		Me.mnuAppendRow.Visible = True
		Me.mnuInsertRow.Name = "mnuInsertRow"
		Me.mnuInsertRow.Text = "Insert Row"
		Me.mnuInsertRow.Checked = False
		Me.mnuInsertRow.Enabled = True
		Me.mnuInsertRow.Visible = True
		Me.mnuDeleteRow.Name = "mnuDeleteRow"
		Me.mnuDeleteRow.Text = "Delete Row"
		Me.mnuDeleteRow.Checked = False
		Me.mnuDeleteRow.Enabled = True
		Me.mnuDeleteRow.Visible = True
		Me._chkWWL_0.FlatStyle = System.Windows.Forms.FlatStyle.Flat
		Me._chkWWL_0.BackColor = System.Drawing.Color.Black
		Me._chkWWL_0.Text = "Check1"
		Me._chkWWL_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._chkWWL_0.Size = New System.Drawing.Size(13, 13)
		Me._chkWWL_0.Location = New System.Drawing.Point(306, 363)
		Me._chkWWL_0.TabIndex = 11
		Me._chkWWL_0.Visible = False
		Me._chkWWL_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkWWL_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkWWL_0.CausesValidation = True
		Me._chkWWL_0.Enabled = True
		Me._chkWWL_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkWWL_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkWWL_0.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkWWL_0.TabStop = True
		Me._chkWWL_0.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._chkWWL_0.Name = "_chkWWL_0"
		Me.txtActiveCell.AutoSize = False
		Me.txtActiveCell.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
		Me.txtActiveCell.BackColor = System.Drawing.Color.White
		Me.txtActiveCell.Size = New System.Drawing.Size(90, 17)
		Me.txtActiveCell.Location = New System.Drawing.Point(153, 356)
		Me.txtActiveCell.TabIndex = 10
		Me.txtActiveCell.Text = "Text1"
		Me.txtActiveCell.Visible = False
		Me.txtActiveCell.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtActiveCell.AcceptsReturn = True
		Me.txtActiveCell.CausesValidation = True
		Me.txtActiveCell.Enabled = True
		Me.txtActiveCell.ForeColor = System.Drawing.SystemColors.WindowText
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
		grdLcClasses.OcxState = CType(resources.GetObject("grdLcClasses.OcxState"), System.Windows.Forms.AxHost.State)
		Me.grdLcClasses.Size = New System.Drawing.Size(581, 231)
		Me.grdLcClasses.Location = New System.Drawing.Point(12, 117)
		Me.grdLcClasses.TabIndex = 9
		Me.grdLcClasses.Name = "grdLcClasses"
		Me.txtLCTypeDesc.AutoSize = False
		Me.txtLCTypeDesc.Size = New System.Drawing.Size(374, 19)
		Me.txtLCTypeDesc.Location = New System.Drawing.Point(120, 66)
		Me.txtLCTypeDesc.TabIndex = 1
		Me.txtLCTypeDesc.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtLCTypeDesc.AcceptsReturn = True
		Me.txtLCTypeDesc.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtLCTypeDesc.BackColor = System.Drawing.SystemColors.Window
		Me.txtLCTypeDesc.CausesValidation = True
		Me.txtLCTypeDesc.Enabled = True
		Me.txtLCTypeDesc.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtLCTypeDesc.HideSelection = True
		Me.txtLCTypeDesc.ReadOnly = False
		Me.txtLCTypeDesc.Maxlength = 0
		Me.txtLCTypeDesc.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtLCTypeDesc.MultiLine = False
		Me.txtLCTypeDesc.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtLCTypeDesc.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtLCTypeDesc.TabStop = True
		Me.txtLCTypeDesc.Visible = True
		Me.txtLCTypeDesc.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtLCTypeDesc.Name = "txtLCTypeDesc"
		Me.txtLCType.AutoSize = False
		Me.txtLCType.Size = New System.Drawing.Size(134, 19)
		Me.txtLCType.Location = New System.Drawing.Point(121, 39)
		Me.txtLCType.TabIndex = 0
		Me.txtLCType.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtLCType.AcceptsReturn = True
		Me.txtLCType.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtLCType.BackColor = System.Drawing.SystemColors.Window
		Me.txtLCType.CausesValidation = True
		Me.txtLCType.Enabled = True
		Me.txtLCType.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtLCType.HideSelection = True
		Me.txtLCType.ReadOnly = False
		Me.txtLCType.Maxlength = 0
		Me.txtLCType.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtLCType.MultiLine = False
		Me.txtLCType.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtLCType.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtLCType.TabStop = True
		Me.txtLCType.Visible = True
		Me.txtLCType.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtLCType.Name = "txtLCType"
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOK.Text = "OK"
		Me.cmdOK.Size = New System.Drawing.Size(65, 25)
		Me.cmdOK.Location = New System.Drawing.Point(434, 360)
		Me.cmdOK.TabIndex = 2
		Me.cmdOK.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOK.CausesValidation = True
		Me.cmdOK.Enabled = True
		Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOK.TabStop = True
		Me.cmdOK.Name = "cmdOK"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(65, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(509, 360)
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
		Me._Label1_6.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_6.Text = "Description:"
		Me._Label1_6.Size = New System.Drawing.Size(82, 17)
		Me._Label1_6.Location = New System.Drawing.Point(25, 66)
		Me._Label1_6.TabIndex = 8
		Me._Label1_6.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_6.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_6.Enabled = True
		Me._Label1_6.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_6.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_6.UseMnemonic = True
		Me._Label1_6.Visible = True
		Me._Label1_6.AutoSize = False
		Me._Label1_6.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_6.Name = "_Label1_6"
		Me._Label2_0.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._Label2_0.Text = "SCS Curve Numbers"
		Me._Label2_0.Size = New System.Drawing.Size(213, 19)
		Me._Label2_0.Location = New System.Drawing.Point(240, 96)
		Me._Label2_0.TabIndex = 7
		Me._Label2_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label2_0.BackColor = System.Drawing.SystemColors.Control
		Me._Label2_0.Enabled = True
		Me._Label2_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label2_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label2_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label2_0.UseMnemonic = True
		Me._Label2_0.Visible = True
		Me._Label2_0.AutoSize = False
		Me._Label2_0.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._Label2_0.Name = "_Label2_0"
		Me._Label2_1.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._Label2_1.Text = "Classification"
		Me._Label2_1.Size = New System.Drawing.Size(227, 19)
		Me._Label2_1.Location = New System.Drawing.Point(11, 96)
		Me._Label2_1.TabIndex = 6
		Me._Label2_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label2_1.BackColor = System.Drawing.SystemColors.Control
		Me._Label2_1.Enabled = True
		Me._Label2_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label2_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label2_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label2_1.UseMnemonic = True
		Me._Label2_1.Visible = True
		Me._Label2_1.AutoSize = False
		Me._Label2_1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._Label2_1.Name = "_Label2_1"
		Me._Label2_2.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._Label2_2.Text = "RUSLE"
		Me._Label2_2.Size = New System.Drawing.Size(160, 19)
		Me._Label2_2.Location = New System.Drawing.Point(454, 96)
		Me._Label2_2.TabIndex = 5
		Me._Label2_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label2_2.BackColor = System.Drawing.SystemColors.Control
		Me._Label2_2.Enabled = True
		Me._Label2_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label2_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label2_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label2_2.UseMnemonic = True
		Me._Label2_2.Visible = True
		Me._Label2_2.AutoSize = False
		Me._Label2_2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._Label2_2.Name = "_Label2_2"
		Me._Label1_7.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label1_7.Text = "Land Cover Type:"
		Me._Label1_7.Size = New System.Drawing.Size(94, 17)
		Me._Label1_7.Location = New System.Drawing.Point(13, 41)
		Me._Label1_7.TabIndex = 4
		Me._Label1_7.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
		Me.Controls.Add(_chkWWL_0)
		Me.Controls.Add(txtActiveCell)
		Me.Controls.Add(grdLcClasses)
		Me.Controls.Add(txtLCTypeDesc)
		Me.Controls.Add(txtLCType)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(_Label1_6)
		Me.Controls.Add(_Label2_0)
		Me.Controls.Add(_Label2_1)
		Me.Controls.Add(_Label2_2)
		Me.Controls.Add(_Label1_7)
		Me.Label1.SetIndex(_Label1_6, CType(6, Short))
		Me.Label1.SetIndex(_Label1_7, CType(7, Short))
		Me.Label2.SetIndex(_Label2_0, CType(0, Short))
		Me.Label2.SetIndex(_Label2_1, CType(1, Short))
		Me.Label2.SetIndex(_Label2_2, CType(2, Short))
		Me.chkWWL.SetIndex(_chkWWL_0, CType(0, Short))
		CType(Me.chkWWL, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Label2, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.grdLcClasses, System.ComponentModel.ISupportInitialize).EndInit()
		MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuPopUp})
		mnuPopUp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuAppendRow, Me.mnuInsertRow, Me.mnuDeleteRow})
		Me.Controls.Add(MainMenu1)
		Me.MainMenu1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class