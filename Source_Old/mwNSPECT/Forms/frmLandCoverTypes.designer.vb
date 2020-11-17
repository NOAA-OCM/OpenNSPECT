<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmLandCoverTypes
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
    Public WithEvents mnuNewLCType As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuDelLCType As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuImpLCType As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuExpLCType As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuLCTypes As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuAppend As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuInsertRow As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuDeleteRow As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuPopUp As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuLCHelp As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuHelp As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
    Public WithEvents Timer1 As System.Windows.Forms.Timer
    Public WithEvents _chkWWL_0 As System.Windows.Forms.CheckBox
    Public WithEvents cboWWL As System.Windows.Forms.ComboBox
    Public WithEvents txtActiveCell As System.Windows.Forms.TextBox
    Public dlgCMD1Open As System.Windows.Forms.OpenFileDialog
    Public dlgCMD1Save As System.Windows.Forms.SaveFileDialog
    Public WithEvents cmdRestore As System.Windows.Forms.Button
    Public WithEvents cmdQuit As System.Windows.Forms.Button
    Public WithEvents cmdSave As System.Windows.Forms.Button
    Public WithEvents txtLCTypeDesc As System.Windows.Forms.TextBox
    Public WithEvents cboLCType As System.Windows.Forms.ComboBox
    Public WithEvents _Label2_2 As System.Windows.Forms.Label
    Public WithEvents _Label2_1 As System.Windows.Forms.Label
    Public WithEvents _Label2_0 As System.Windows.Forms.Label
    Public WithEvents _Label1_6 As System.Windows.Forms.Label
    Public WithEvents _Label1_7 As System.Windows.Forms.Label
    'Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    'Public WithEvents Label2 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    'Public WithEvents chkWWL As Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me._chkWWL_0 = New System.Windows.Forms.CheckBox
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip
        Me.mnuLCTypes = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuNewLCType = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuDelLCType = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuImpLCType = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuExpLCType = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuPopUp = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuAppend = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuInsertRow = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuDeleteRow = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuHelp = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuLCHelp = New System.Windows.Forms.ToolStripMenuItem
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.cboWWL = New System.Windows.Forms.ComboBox
        Me.txtActiveCell = New System.Windows.Forms.TextBox
        Me.dlgCMD1Open = New System.Windows.Forms.OpenFileDialog
        Me.dlgCMD1Save = New System.Windows.Forms.SaveFileDialog
        Me.cmdRestore = New System.Windows.Forms.Button
        Me.cmdQuit = New System.Windows.Forms.Button
        Me.cmdSave = New System.Windows.Forms.Button
        Me.txtLCTypeDesc = New System.Windows.Forms.TextBox
        Me.cboLCType = New System.Windows.Forms.ComboBox
        Me._Label2_2 = New System.Windows.Forms.Label
        Me._Label2_1 = New System.Windows.Forms.Label
        Me._Label2_0 = New System.Windows.Forms.Label
        Me._Label1_6 = New System.Windows.Forms.Label
        Me._Label1_7 = New System.Windows.Forms.Label
        Me.MainMenu1.SuspendLayout()
        Me.SuspendLayout()
        '
        '_chkWWL_0
        '
        Me._chkWWL_0.BackColor = System.Drawing.SystemColors.Window
        Me._chkWWL_0.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me._chkWWL_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._chkWWL_0.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me._chkWWL_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._chkWWL_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me._chkWWL_0.Location = New System.Drawing.Point(207, 553)
        Me._chkWWL_0.Name = "_chkWWL_0"
        Me._chkWWL_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._chkWWL_0.Size = New System.Drawing.Size(13, 13)
        Me._chkWWL_0.TabIndex = 13
        Me.ToolTip1.SetToolTip(Me._chkWWL_0, "Check if landuse is water or wetland")
        Me._chkWWL_0.UseVisualStyleBackColor = False
        Me._chkWWL_0.Visible = False
        '
        'MainMenu1
        '
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuLCTypes, Me.mnuPopUp, Me.mnuHelp})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(618, 24)
        Me.MainMenu1.TabIndex = 14
        '
        'mnuLCTypes
        '
        Me.mnuLCTypes.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuNewLCType, Me.mnuDelLCType, Me.mnuImpLCType, Me.mnuExpLCType})
        Me.mnuLCTypes.Name = "mnuLCTypes"
        Me.mnuLCTypes.Size = New System.Drawing.Size(61, 20)
        Me.mnuLCTypes.Text = "&Options"
        '
        'mnuNewLCType
        '
        Me.mnuNewLCType.Name = "mnuNewLCType"
        Me.mnuNewLCType.Size = New System.Drawing.Size(119, 22)
        Me.mnuNewLCType.Text = "&New..."
        '
        'mnuDelLCType
        '
        Me.mnuDelLCType.Name = "mnuDelLCType"
        Me.mnuDelLCType.Size = New System.Drawing.Size(119, 22)
        Me.mnuDelLCType.Text = "&Delete..."
        '
        'mnuImpLCType
        '
        Me.mnuImpLCType.Name = "mnuImpLCType"
        Me.mnuImpLCType.Size = New System.Drawing.Size(119, 22)
        Me.mnuImpLCType.Text = "&Import..."
        '
        'mnuExpLCType
        '
        Me.mnuExpLCType.Name = "mnuExpLCType"
        Me.mnuExpLCType.Size = New System.Drawing.Size(119, 22)
        Me.mnuExpLCType.Text = "&Export..."
        '
        'mnuPopUp
        '
        Me.mnuPopUp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuAppend, Me.mnuInsertRow, Me.mnuDeleteRow})
        Me.mnuPopUp.Name = "mnuPopUp"
        Me.mnuPopUp.Size = New System.Drawing.Size(39, 20)
        Me.mnuPopUp.Text = "Edit"
        Me.mnuPopUp.Visible = False
        '
        'mnuAppend
        '
        Me.mnuAppend.Name = "mnuAppend"
        Me.mnuAppend.Size = New System.Drawing.Size(133, 22)
        Me.mnuAppend.Text = "Add Row"
        '
        'mnuInsertRow
        '
        Me.mnuInsertRow.Name = "mnuInsertRow"
        Me.mnuInsertRow.Size = New System.Drawing.Size(133, 22)
        Me.mnuInsertRow.Text = "Insert Row"
        '
        'mnuDeleteRow
        '
        Me.mnuDeleteRow.Name = "mnuDeleteRow"
        Me.mnuDeleteRow.Size = New System.Drawing.Size(133, 22)
        Me.mnuDeleteRow.Text = "Delete Row"
        '
        'mnuHelp
        '
        Me.mnuHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuLCHelp})
        Me.mnuHelp.Name = "mnuHelp"
        Me.mnuHelp.Size = New System.Drawing.Size(44, 20)
        Me.mnuHelp.Text = "&Help"
        '
        'mnuLCHelp
        '
        Me.mnuLCHelp.Name = "mnuLCHelp"
        Me.mnuLCHelp.ShortcutKeys = CType((System.Windows.Forms.Keys.Shift Or System.Windows.Forms.Keys.F1), System.Windows.Forms.Keys)
        Me.mnuLCHelp.Size = New System.Drawing.Size(228, 22)
        Me.mnuLCHelp.Text = "Land Cover Types..."
        '
        'Timer1
        '
        Me.Timer1.Interval = 1
        '
        'cboWWL
        '
        Me.cboWWL.BackColor = System.Drawing.SystemColors.Window
        Me.cboWWL.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboWWL.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboWWL.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboWWL.Items.AddRange(New Object() {"Y", "N"})
        Me.cboWWL.Location = New System.Drawing.Point(370, 559)
        Me.cboWWL.Name = "cboWWL"
        Me.cboWWL.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboWWL.Size = New System.Drawing.Size(45, 22)
        Me.cboWWL.TabIndex = 12
        Me.cboWWL.Visible = False
        '
        'txtActiveCell
        '
        Me.txtActiveCell.AcceptsReturn = True
        Me.txtActiveCell.BackColor = System.Drawing.Color.White
        Me.txtActiveCell.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtActiveCell.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtActiveCell.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtActiveCell.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtActiveCell.HideSelection = False
        Me.txtActiveCell.Location = New System.Drawing.Point(268, 561)
        Me.txtActiveCell.MaxLength = 0
        Me.txtActiveCell.Name = "txtActiveCell"
        Me.txtActiveCell.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtActiveCell.Size = New System.Drawing.Size(70, 16)
        Me.txtActiveCell.TabIndex = 11
        Me.txtActiveCell.Text = "Text1"
        Me.txtActiveCell.Visible = False
        '
        'cmdRestore
        '
        Me.cmdRestore.BackColor = System.Drawing.SystemColors.Control
        Me.cmdRestore.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdRestore.Enabled = False
        Me.cmdRestore.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdRestore.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdRestore.Location = New System.Drawing.Point(25, 551)
        Me.cmdRestore.Name = "cmdRestore"
        Me.cmdRestore.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdRestore.Size = New System.Drawing.Size(103, 25)
        Me.cmdRestore.TabIndex = 6
        Me.cmdRestore.Text = "Restore Defaults"
        Me.cmdRestore.UseVisualStyleBackColor = False
        '
        'cmdQuit
        '
        Me.cmdQuit.BackColor = System.Drawing.SystemColors.Control
        Me.cmdQuit.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdQuit.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdQuit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdQuit.Location = New System.Drawing.Point(512, 550)
        Me.cmdQuit.Name = "cmdQuit"
        Me.cmdQuit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdQuit.Size = New System.Drawing.Size(65, 25)
        Me.cmdQuit.TabIndex = 5
        Me.cmdQuit.Text = "Cancel"
        Me.cmdQuit.UseVisualStyleBackColor = False
        '
        'cmdSave
        '
        Me.cmdSave.BackColor = System.Drawing.SystemColors.Control
        Me.cmdSave.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdSave.Enabled = False
        Me.cmdSave.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSave.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSave.Location = New System.Drawing.Point(441, 550)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdSave.Size = New System.Drawing.Size(65, 25)
        Me.cmdSave.TabIndex = 4
        Me.cmdSave.Text = "Save"
        Me.cmdSave.UseVisualStyleBackColor = False
        '
        'txtLCTypeDesc
        '
        Me.txtLCTypeDesc.AcceptsReturn = True
        Me.txtLCTypeDesc.BackColor = System.Drawing.SystemColors.Window
        Me.txtLCTypeDesc.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtLCTypeDesc.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLCTypeDesc.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtLCTypeDesc.Location = New System.Drawing.Point(124, 52)
        Me.txtLCTypeDesc.MaxLength = 0
        Me.txtLCTypeDesc.Name = "txtLCTypeDesc"
        Me.txtLCTypeDesc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtLCTypeDesc.Size = New System.Drawing.Size(362, 19)
        Me.txtLCTypeDesc.TabIndex = 2
        '
        'cboLCType
        '
        Me.cboLCType.BackColor = System.Drawing.SystemColors.Window
        Me.cboLCType.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboLCType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboLCType.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboLCType.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboLCType.Location = New System.Drawing.Point(124, 26)
        Me.cboLCType.Name = "cboLCType"
        Me.cboLCType.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboLCType.Size = New System.Drawing.Size(149, 22)
        Me.cboLCType.TabIndex = 0
        '
        '_Label2_2
        '
        Me._Label2_2.BackColor = System.Drawing.SystemColors.Control
        Me._Label2_2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me._Label2_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label2_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label2_2.Location = New System.Drawing.Point(451, 78)
        Me._Label2_2.Name = "_Label2_2"
        Me._Label2_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_2.Size = New System.Drawing.Size(160, 19)
        Me._Label2_2.TabIndex = 9
        Me._Label2_2.Text = "RUSLE"
        Me._Label2_2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_Label2_1
        '
        Me._Label2_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label2_1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me._Label2_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label2_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label2_1.Location = New System.Drawing.Point(7, 78)
        Me._Label2_1.Name = "_Label2_1"
        Me._Label2_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_1.Size = New System.Drawing.Size(227, 19)
        Me._Label2_1.TabIndex = 8
        Me._Label2_1.Text = "Classification"
        Me._Label2_1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_Label2_0
        '
        Me._Label2_0.BackColor = System.Drawing.SystemColors.Control
        Me._Label2_0.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me._Label2_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label2_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label2_0.Location = New System.Drawing.Point(236, 78)
        Me._Label2_0.Name = "_Label2_0"
        Me._Label2_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_0.Size = New System.Drawing.Size(213, 19)
        Me._Label2_0.TabIndex = 7
        Me._Label2_0.Text = "SCS Curve Numbers"
        Me._Label2_0.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_Label1_6
        '
        Me._Label1_6.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_6.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_6.Location = New System.Drawing.Point(21, 51)
        Me._Label1_6.Name = "_Label1_6"
        Me._Label1_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_6.Size = New System.Drawing.Size(86, 17)
        Me._Label1_6.TabIndex = 3
        Me._Label1_6.Text = "Description:"
        Me._Label1_6.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_7
        '
        Me._Label1_7.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_7.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_7.Location = New System.Drawing.Point(16, 29)
        Me._Label1_7.Name = "_Label1_7"
        Me._Label1_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_7.Size = New System.Drawing.Size(91, 17)
        Me._Label1_7.TabIndex = 1
        Me._Label1_7.Text = "Land Cover Type:"
        Me._Label1_7.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'frmLCTypes
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(618, 590)
        Me.Controls.Add(Me._chkWWL_0)
        Me.Controls.Add(Me.cboWWL)
        Me.Controls.Add(Me.txtActiveCell)
        Me.Controls.Add(Me.cmdRestore)
        Me.Controls.Add(Me.cmdQuit)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.txtLCTypeDesc)
        Me.Controls.Add(Me.cboLCType)
        Me.Controls.Add(Me._Label2_2)
        Me.Controls.Add(Me._Label2_1)
        Me.Controls.Add(Me._Label2_0)
        Me.Controls.Add(Me._Label1_6)
        Me.Controls.Add(Me._Label1_7)
        Me.Controls.Add(Me.MainMenu1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(268, 71)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmLCTypes"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Land Cover Types"
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
End Class