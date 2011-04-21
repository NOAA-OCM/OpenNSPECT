<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmNewWSDelin
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
	Public WithEvents chkHydroCorr As System.Windows.Forms.CheckBox
	Public WithEvents cmdBrowseDEMFile As System.Windows.Forms.Button
	Public WithEvents txtWSDelinName As System.Windows.Forms.TextBox
	Public WithEvents cboDEMUnits As System.Windows.Forms.ComboBox
	Public WithEvents txtDEMFile As System.Windows.Forms.TextBox
	Public WithEvents cboSubWSSize As System.Windows.Forms.ComboBox
	Public WithEvents cmdOptions As System.Windows.Forms.Button
	Public WithEvents cboStreamLayer As System.Windows.Forms.ComboBox
	Public WithEvents chkStreamAgree As System.Windows.Forms.CheckBox
	Public WithEvents _lblStream_0 As System.Windows.Forms.Label
	Public WithEvents frmAdvanced As System.Windows.Forms.GroupBox
	Public WithEvents _Label1_3 As System.Windows.Forms.Label
	Public WithEvents _Label1_2 As System.Windows.Forms.Label
	Public WithEvents _Label1_12 As System.Windows.Forms.Label
	Public WithEvents _Label1_1 As System.Windows.Forms.Label
	Public WithEvents frmMain As System.Windows.Forms.GroupBox
	Public WithEvents cmdCreate As System.Windows.Forms.Button
	Public WithEvents cmdQuit As System.Windows.Forms.Button
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmNewWSDelin))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.frmMain = New System.Windows.Forms.GroupBox()
        Me.cboDEMUnits = New System.Windows.Forms.ComboBox()
        Me.chkHydroCorr = New System.Windows.Forms.CheckBox()
        Me.cmdBrowseDEMFile = New System.Windows.Forms.Button()
        Me.txtWSDelinName = New System.Windows.Forms.TextBox()
        Me.txtDEMFile = New System.Windows.Forms.TextBox()
        Me.cboSubWSSize = New System.Windows.Forms.ComboBox()
        Me._Label1_3 = New System.Windows.Forms.Label()
        Me._Label1_2 = New System.Windows.Forms.Label()
        Me._Label1_12 = New System.Windows.Forms.Label()
        Me._Label1_1 = New System.Windows.Forms.Label()
        Me.frmAdvanced = New System.Windows.Forms.GroupBox()
        Me.cmdOptions = New System.Windows.Forms.Button()
        Me.cboStreamLayer = New System.Windows.Forms.ComboBox()
        Me.chkStreamAgree = New System.Windows.Forms.CheckBox()
        Me._lblStream_0 = New System.Windows.Forms.Label()
        Me.cmdCreate = New System.Windows.Forms.Button()
        Me.cmdQuit = New System.Windows.Forms.Button()
        Me.frmMain.SuspendLayout()
        Me.frmAdvanced.SuspendLayout()
        Me.SuspendLayout()
        '
        'frmMain
        '
        Me.frmMain.BackColor = System.Drawing.SystemColors.Control
        Me.frmMain.Controls.Add(Me.cboDEMUnits)
        Me.frmMain.Controls.Add(Me.chkHydroCorr)
        Me.frmMain.Controls.Add(Me.cmdBrowseDEMFile)
        Me.frmMain.Controls.Add(Me.txtWSDelinName)
        Me.frmMain.Controls.Add(Me.txtDEMFile)
        Me.frmMain.Controls.Add(Me.cboSubWSSize)
        Me.frmMain.Controls.Add(Me._Label1_3)
        Me.frmMain.Controls.Add(Me._Label1_2)
        Me.frmMain.Controls.Add(Me._Label1_12)
        Me.frmMain.Controls.Add(Me._Label1_1)
        Me.frmMain.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.frmMain.ForeColor = System.Drawing.SystemColors.ControlText
        Me.frmMain.Location = New System.Drawing.Point(13, 9)
        Me.frmMain.Name = "frmMain"
        Me.frmMain.Padding = New System.Windows.Forms.Padding(0)
        Me.frmMain.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.frmMain.Size = New System.Drawing.Size(437, 160)
        Me.frmMain.TabIndex = 7
        Me.frmMain.TabStop = False
        Me.frmMain.Text = "Create a new watershed delineation  "
        '
        'cboDEMUnits
        '
        Me.cboDEMUnits.BackColor = System.Drawing.SystemColors.Window
        Me.cboDEMUnits.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboDEMUnits.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDEMUnits.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboDEMUnits.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboDEMUnits.Items.AddRange(New Object() {"meters", "feet"})
        Me.cboDEMUnits.Location = New System.Drawing.Point(119, 99)
        Me.cboDEMUnits.Name = "cboDEMUnits"
        Me.cboDEMUnits.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboDEMUnits.Size = New System.Drawing.Size(152, 22)
        Me.cboDEMUnits.TabIndex = 2
        '
        'chkHydroCorr
        '
        Me.chkHydroCorr.BackColor = System.Drawing.SystemColors.Control
        Me.chkHydroCorr.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkHydroCorr.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkHydroCorr.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkHydroCorr.Location = New System.Drawing.Point(119, 73)
        Me.chkHydroCorr.Name = "chkHydroCorr"
        Me.chkHydroCorr.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkHydroCorr.Size = New System.Drawing.Size(241, 24)
        Me.chkHydroCorr.TabIndex = 17
        Me.chkHydroCorr.Text = "DEM is hyrdologically correct (filled)"
        Me.chkHydroCorr.UseVisualStyleBackColor = False
        '
        'cmdBrowseDEMFile
        '
        Me.cmdBrowseDEMFile.BackColor = System.Drawing.SystemColors.Control
        Me.cmdBrowseDEMFile.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdBrowseDEMFile.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowseDEMFile.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdBrowseDEMFile.Image = CType(resources.GetObject("cmdBrowseDEMFile.Image"), System.Drawing.Image)
        Me.cmdBrowseDEMFile.Location = New System.Drawing.Point(391, 46)
        Me.cmdBrowseDEMFile.Name = "cmdBrowseDEMFile"
        Me.cmdBrowseDEMFile.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdBrowseDEMFile.Size = New System.Drawing.Size(25, 21)
        Me.cmdBrowseDEMFile.TabIndex = 1
        Me.cmdBrowseDEMFile.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdBrowseDEMFile.UseVisualStyleBackColor = False
        '
        'txtWSDelinName
        '
        Me.txtWSDelinName.AcceptsReturn = True
        Me.txtWSDelinName.BackColor = System.Drawing.SystemColors.Window
        Me.txtWSDelinName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtWSDelinName.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtWSDelinName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtWSDelinName.Location = New System.Drawing.Point(118, 20)
        Me.txtWSDelinName.MaxLength = 0
        Me.txtWSDelinName.Name = "txtWSDelinName"
        Me.txtWSDelinName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtWSDelinName.Size = New System.Drawing.Size(134, 20)
        Me.txtWSDelinName.TabIndex = 0
        '
        'txtDEMFile
        '
        Me.txtDEMFile.AcceptsReturn = True
        Me.txtDEMFile.BackColor = System.Drawing.SystemColors.Window
        Me.txtDEMFile.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDEMFile.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDEMFile.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDEMFile.Location = New System.Drawing.Point(118, 48)
        Me.txtDEMFile.MaxLength = 0
        Me.txtDEMFile.Name = "txtDEMFile"
        Me.txtDEMFile.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDEMFile.Size = New System.Drawing.Size(271, 20)
        Me.txtDEMFile.TabIndex = 4
        '
        'cboSubWSSize
        '
        Me.cboSubWSSize.BackColor = System.Drawing.SystemColors.Window
        Me.cboSubWSSize.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboSubWSSize.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSubWSSize.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboSubWSSize.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboSubWSSize.Items.AddRange(New Object() {"small", "medium", "large"})
        Me.cboSubWSSize.Location = New System.Drawing.Point(119, 129)
        Me.cboSubWSSize.Name = "cboSubWSSize"
        Me.cboSubWSSize.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboSubWSSize.Size = New System.Drawing.Size(151, 22)
        Me.cboSubWSSize.TabIndex = 3
        '
        '_Label1_3
        '
        Me._Label1_3.AutoSize = True
        Me._Label1_3.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_3.Location = New System.Drawing.Point(13, 24)
        Me._Label1_3.Name = "_Label1_3"
        Me._Label1_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_3.Size = New System.Drawing.Size(92, 14)
        Me._Label1_3.TabIndex = 11
        Me._Label1_3.Text = "Delineation Name:"
        '
        '_Label1_2
        '
        Me._Label1_2.AutoSize = True
        Me._Label1_2.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_2.Location = New System.Drawing.Point(14, 102)
        Me._Label1_2.Name = "_Label1_2"
        Me._Label1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_2.Size = New System.Drawing.Size(58, 14)
        Me._Label1_2.TabIndex = 10
        Me._Label1_2.Text = "DEM Units:"
        '
        '_Label1_12
        '
        Me._Label1_12.AutoSize = True
        Me._Label1_12.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_12.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_12.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_12.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_12.Location = New System.Drawing.Point(13, 52)
        Me._Label1_12.Name = "_Label1_12"
        Me._Label1_12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_12.Size = New System.Drawing.Size(54, 14)
        Me._Label1_12.TabIndex = 9
        Me._Label1_12.Text = "DEM Grid:"
        '
        '_Label1_1
        '
        Me._Label1_1.AutoSize = True
        Me._Label1_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_1.Location = New System.Drawing.Point(14, 132)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_1.Size = New System.Drawing.Size(106, 14)
        Me._Label1_1.TabIndex = 8
        Me._Label1_1.Text = "Subwatershed Size:"
        Me._Label1_1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'frmAdvanced
        '
        Me.frmAdvanced.BackColor = System.Drawing.SystemColors.Control
        Me.frmAdvanced.Controls.Add(Me.cmdOptions)
        Me.frmAdvanced.Controls.Add(Me.cboStreamLayer)
        Me.frmAdvanced.Controls.Add(Me.chkStreamAgree)
        Me.frmAdvanced.Controls.Add(Me._lblStream_0)
        Me.frmAdvanced.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.frmAdvanced.ForeColor = System.Drawing.SystemColors.ControlText
        Me.frmAdvanced.Location = New System.Drawing.Point(13, 200)
        Me.frmAdvanced.Name = "frmAdvanced"
        Me.frmAdvanced.Padding = New System.Windows.Forms.Padding(0)
        Me.frmAdvanced.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.frmAdvanced.Size = New System.Drawing.Size(406, 110)
        Me.frmAdvanced.TabIndex = 12
        Me.frmAdvanced.TabStop = False
        Me.frmAdvanced.Text = "Advanced Parameters (optional) "
        Me.frmAdvanced.Visible = False
        '
        'cmdOptions
        '
        Me.cmdOptions.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOptions.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOptions.Enabled = False
        Me.cmdOptions.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOptions.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOptions.Location = New System.Drawing.Point(281, 71)
        Me.cmdOptions.Name = "cmdOptions"
        Me.cmdOptions.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOptions.Size = New System.Drawing.Size(59, 25)
        Me.cmdOptions.TabIndex = 16
        Me.cmdOptions.Text = "Options..."
        Me.cmdOptions.UseVisualStyleBackColor = False
        '
        'cboStreamLayer
        '
        Me.cboStreamLayer.BackColor = System.Drawing.SystemColors.Window
        Me.cboStreamLayer.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboStreamLayer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboStreamLayer.Enabled = False
        Me.cboStreamLayer.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboStreamLayer.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboStreamLayer.Items.AddRange(New Object() {"Stream1"})
        Me.cboStreamLayer.Location = New System.Drawing.Point(122, 73)
        Me.cboStreamLayer.Name = "cboStreamLayer"
        Me.cboStreamLayer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboStreamLayer.Size = New System.Drawing.Size(155, 22)
        Me.cboStreamLayer.TabIndex = 14
        '
        'chkStreamAgree
        '
        Me.chkStreamAgree.BackColor = System.Drawing.SystemColors.Control
        Me.chkStreamAgree.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkStreamAgree.Enabled = False
        Me.chkStreamAgree.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkStreamAgree.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkStreamAgree.Location = New System.Drawing.Point(48, 49)
        Me.chkStreamAgree.Name = "chkStreamAgree"
        Me.chkStreamAgree.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkStreamAgree.Size = New System.Drawing.Size(149, 19)
        Me.chkStreamAgree.TabIndex = 13
        Me.chkStreamAgree.Text = "Force Stream Agreement"
        Me.chkStreamAgree.UseVisualStyleBackColor = False
        '
        '_lblStream_0
        '
        Me._lblStream_0.AutoSize = True
        Me._lblStream_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblStream_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblStream_0.Enabled = False
        Me._lblStream_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblStream_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblStream_0.Location = New System.Drawing.Point(49, 76)
        Me._lblStream_0.Name = "_lblStream_0"
        Me._lblStream_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblStream_0.Size = New System.Drawing.Size(75, 14)
        Me._lblStream_0.TabIndex = 15
        Me._lblStream_0.Text = "Stream Layer:"
        '
        'cmdCreate
        '
        Me.cmdCreate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCreate.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCreate.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCreate.Enabled = False
        Me.cmdCreate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCreate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCreate.Location = New System.Drawing.Point(407, 341)
        Me.cmdCreate.Name = "cmdCreate"
        Me.cmdCreate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCreate.Size = New System.Drawing.Size(65, 25)
        Me.cmdCreate.TabIndex = 5
        Me.cmdCreate.Text = "OK"
        Me.cmdCreate.UseVisualStyleBackColor = False
        '
        'cmdQuit
        '
        Me.cmdQuit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdQuit.BackColor = System.Drawing.SystemColors.Control
        Me.cmdQuit.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdQuit.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdQuit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdQuit.Location = New System.Drawing.Point(482, 341)
        Me.cmdQuit.Name = "cmdQuit"
        Me.cmdQuit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdQuit.Size = New System.Drawing.Size(65, 25)
        Me.cmdQuit.TabIndex = 6
        Me.cmdQuit.Text = "Cancel"
        Me.cmdQuit.UseVisualStyleBackColor = False
        '
        'frmNewWSDelin
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(594, 372)
        Me.Controls.Add(Me.frmMain)
        Me.Controls.Add(Me.cmdCreate)
        Me.Controls.Add(Me.cmdQuit)
        Me.Controls.Add(Me.frmAdvanced)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(213, 196)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmNewWSDelin"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "New Watershed Delineation"
        Me.frmMain.ResumeLayout(False)
        Me.frmMain.PerformLayout()
        Me.frmAdvanced.ResumeLayout(False)
        Me.frmAdvanced.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
#End Region 
End Class