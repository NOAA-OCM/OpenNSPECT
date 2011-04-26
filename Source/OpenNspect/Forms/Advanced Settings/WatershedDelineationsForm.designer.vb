<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class WatershedDelineationsForm
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
    Public WithEvents mnuNewWSDelin As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuNewExist As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuDelWSDelin As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuDefWSDelin As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuWSDelin As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuHelp As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
    Public WithEvents txtLSGrid As System.Windows.Forms.TextBox
    Public WithEvents cboDEMUnits As System.Windows.Forms.ComboBox
    Public WithEvents cboWSSize As System.Windows.Forms.ComboBox
    Public WithEvents chkHydroCorr As System.Windows.Forms.CheckBox
    Public WithEvents txtFlowAccumGrid As System.Windows.Forms.TextBox
    Public WithEvents txtWSFile As System.Windows.Forms.TextBox
    Public WithEvents txtStream As System.Windows.Forms.TextBox
    Public WithEvents cboWSDelin As System.Windows.Forms.ComboBox
    Public WithEvents txtDEMFile As System.Windows.Forms.TextBox
    Public WithEvents _Label1_6 As System.Windows.Forms.Label
    Public WithEvents _Label1_12 As System.Windows.Forms.Label
    Public WithEvents _Label1_2 As System.Windows.Forms.Label
    Public WithEvents _Label1_0 As System.Windows.Forms.Label
    Public WithEvents _Label1_3 As System.Windows.Forms.Label
    Public WithEvents _Label1_4 As System.Windows.Forms.Label
    Public WithEvents _Label1_5 As System.Windows.Forms.Label
    Public WithEvents _Label1_1 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    'Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip()
        Me.mnuDefWSDelin = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuNewWSDelin = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuNewExist = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuDelWSDelin = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuHelp = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuWSDelin = New System.Windows.Forms.ToolStripMenuItem()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.txtLSGrid = New System.Windows.Forms.TextBox()
        Me.cboDEMUnits = New System.Windows.Forms.ComboBox()
        Me.cboWSSize = New System.Windows.Forms.ComboBox()
        Me.chkHydroCorr = New System.Windows.Forms.CheckBox()
        Me.txtFlowAccumGrid = New System.Windows.Forms.TextBox()
        Me.txtWSFile = New System.Windows.Forms.TextBox()
        Me.txtStream = New System.Windows.Forms.TextBox()
        Me.cboWSDelin = New System.Windows.Forms.ComboBox()
        Me.txtDEMFile = New System.Windows.Forms.TextBox()
        Me._Label1_6 = New System.Windows.Forms.Label()
        Me._Label1_12 = New System.Windows.Forms.Label()
        Me._Label1_2 = New System.Windows.Forms.Label()
        Me._Label1_0 = New System.Windows.Forms.Label()
        Me._Label1_3 = New System.Windows.Forms.Label()
        Me._Label1_4 = New System.Windows.Forms.Label()
        Me._Label1_5 = New System.Windows.Forms.Label()
        Me._Label1_1 = New System.Windows.Forms.Label()
        Me.MainMenu1.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuDefWSDelin, Me.mnuHelp})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(594, 24)
        Me.MainMenu1.TabIndex = 2
        '
        'mnuDefWSDelin
        '
        Me.mnuDefWSDelin.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuNewWSDelin, Me.mnuNewExist, Me.mnuDelWSDelin})
        Me.mnuDefWSDelin.Name = "mnuDefWSDelin"
        Me.mnuDefWSDelin.Size = New System.Drawing.Size(61, 20)
        Me.mnuDefWSDelin.Text = "&Options"
        '
        'mnuNewWSDelin
        '
        Me.mnuNewWSDelin.Name = "mnuNewWSDelin"
        Me.mnuNewWSDelin.Size = New System.Drawing.Size(205, 22)
        Me.mnuNewWSDelin.Text = "&New..."
        '
        'mnuNewExist
        '
        Me.mnuNewExist.Name = "mnuNewExist"
        Me.mnuNewExist.Size = New System.Drawing.Size(205, 22)
        Me.mnuNewExist.Text = "New from existing data..."
        '
        'mnuDelWSDelin
        '
        Me.mnuDelWSDelin.Name = "mnuDelWSDelin"
        Me.mnuDelWSDelin.Size = New System.Drawing.Size(205, 22)
        Me.mnuDelWSDelin.Text = "&Delete..."
        '
        'mnuHelp
        '
        Me.mnuHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuWSDelin})
        Me.mnuHelp.Name = "mnuHelp"
        Me.mnuHelp.Size = New System.Drawing.Size(44, 20)
        Me.mnuHelp.Text = "&Help"
        '
        'mnuWSDelin
        '
        Me.mnuWSDelin.Name = "mnuWSDelin"
        Me.mnuWSDelin.ShortcutKeys = CType((System.Windows.Forms.Keys.Shift Or System.Windows.Forms.Keys.F1), System.Windows.Forms.Keys)
        Me.mnuWSDelin.Size = New System.Drawing.Size(258, 22)
        Me.mnuWSDelin.Text = "Watershed Delineations..."
        '
        'Frame1
        '
        Me.Frame1.Controls.Add(Me.txtLSGrid)
        Me.Frame1.Controls.Add(Me.cboDEMUnits)
        Me.Frame1.Controls.Add(Me.cboWSSize)
        Me.Frame1.Controls.Add(Me.chkHydroCorr)
        Me.Frame1.Controls.Add(Me.txtFlowAccumGrid)
        Me.Frame1.Controls.Add(Me.txtWSFile)
        Me.Frame1.Controls.Add(Me.txtStream)
        Me.Frame1.Controls.Add(Me.cboWSDelin)
        Me.Frame1.Controls.Add(Me.txtDEMFile)
        Me.Frame1.Controls.Add(Me._Label1_6)
        Me.Frame1.Controls.Add(Me._Label1_12)
        Me.Frame1.Controls.Add(Me._Label1_2)
        Me.Frame1.Controls.Add(Me._Label1_0)
        Me.Frame1.Controls.Add(Me._Label1_3)
        Me.Frame1.Controls.Add(Me._Label1_4)
        Me.Frame1.Controls.Add(Me._Label1_5)
        Me.Frame1.Controls.Add(Me._Label1_1)
        Me.Frame1.Location = New System.Drawing.Point(10, 26)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.Size = New System.Drawing.Size(572, 266)
        Me.Frame1.TabIndex = 1
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Browse Watershed Delineations  "
        '
        'txtLSGrid
        '
        Me.txtLSGrid.AcceptsReturn = True
        Me.txtLSGrid.Enabled = False
        Me.txtLSGrid.Location = New System.Drawing.Point(178, 179)
        Me.txtLSGrid.MaxLength = 0
        Me.txtLSGrid.Name = "txtLSGrid"
        Me.txtLSGrid.Size = New System.Drawing.Size(391, 20)
        Me.txtLSGrid.TabIndex = 17
        '
        'cboDEMUnits
        '
        Me.cboDEMUnits.Enabled = False
        Me.cboDEMUnits.Items.AddRange(New Object() {"meters", "feet"})
        Me.cboDEMUnits.Location = New System.Drawing.Point(178, 62)
        Me.cboDEMUnits.Name = "cboDEMUnits"
        Me.cboDEMUnits.Size = New System.Drawing.Size(202, 21)
        Me.cboDEMUnits.TabIndex = 16
        Me.cboDEMUnits.Text = "Combo1"
        '
        'cboWSSize
        '
        Me.cboWSSize.CausesValidation = False
        Me.cboWSSize.Enabled = False
        Me.cboWSSize.Items.AddRange(New Object() {"small", "medium", "large"})
        Me.cboWSSize.Location = New System.Drawing.Point(178, 106)
        Me.cboWSSize.Name = "cboWSSize"
        Me.cboWSSize.Size = New System.Drawing.Size(120, 21)
        Me.cboWSSize.TabIndex = 15
        Me.cboWSSize.Text = "cboWSSize"
        '
        'chkHydroCorr
        '
        Me.chkHydroCorr.Enabled = False
        Me.chkHydroCorr.Location = New System.Drawing.Point(179, 86)
        Me.chkHydroCorr.Name = "chkHydroCorr"
        Me.chkHydroCorr.Size = New System.Drawing.Size(173, 18)
        Me.chkHydroCorr.TabIndex = 7
        Me.chkHydroCorr.Text = "Hydrologically Corrected DEM"
        Me.chkHydroCorr.UseVisualStyleBackColor = True
        '
        'txtFlowAccumGrid
        '
        Me.txtFlowAccumGrid.AcceptsReturn = True
        Me.txtFlowAccumGrid.Enabled = False
        Me.txtFlowAccumGrid.Location = New System.Drawing.Point(178, 155)
        Me.txtFlowAccumGrid.MaxLength = 0
        Me.txtFlowAccumGrid.Name = "txtFlowAccumGrid"
        Me.txtFlowAccumGrid.Size = New System.Drawing.Size(391, 20)
        Me.txtFlowAccumGrid.TabIndex = 6
        '
        'txtWSFile
        '
        Me.txtWSFile.AcceptsReturn = True
        Me.txtWSFile.Enabled = False
        Me.txtWSFile.Location = New System.Drawing.Point(178, 130)
        Me.txtWSFile.MaxLength = 0
        Me.txtWSFile.Name = "txtWSFile"
        Me.txtWSFile.Size = New System.Drawing.Size(391, 20)
        Me.txtWSFile.TabIndex = 5
        '
        'txtStream
        '
        Me.txtStream.AcceptsReturn = True
        Me.txtStream.Enabled = False
        Me.txtStream.Location = New System.Drawing.Point(178, 107)
        Me.txtStream.MaxLength = 0
        Me.txtStream.Name = "txtStream"
        Me.txtStream.Size = New System.Drawing.Size(202, 20)
        Me.txtStream.TabIndex = 4
        Me.txtStream.Visible = False
        '
        'cboWSDelin
        '
        Me.cboWSDelin.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboWSDelin.Location = New System.Drawing.Point(178, 15)
        Me.cboWSDelin.Name = "cboWSDelin"
        Me.cboWSDelin.Size = New System.Drawing.Size(202, 21)
        Me.cboWSDelin.TabIndex = 3
        '
        'txtDEMFile
        '
        Me.txtDEMFile.AcceptsReturn = True
        Me.txtDEMFile.Enabled = False
        Me.txtDEMFile.Location = New System.Drawing.Point(178, 40)
        Me.txtDEMFile.MaxLength = 0
        Me.txtDEMFile.Name = "txtDEMFile"
        Me.txtDEMFile.Size = New System.Drawing.Size(388, 20)
        Me.txtDEMFile.TabIndex = 2
        '
        '_Label1_6
        '
        Me._Label1_6.AutoSize = True
        Me._Label1_6.Location = New System.Drawing.Point(124, 180)
        Me._Label1_6.Name = "_Label1_6"
        Me._Label1_6.Size = New System.Drawing.Size(45, 13)
        Me._Label1_6.TabIndex = 18
        Me._Label1_6.Text = "LS Grid:"
        Me._Label1_6.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_12
        '
        Me._Label1_12.AutoSize = True
        Me._Label1_12.Location = New System.Drawing.Point(113, 39)
        Me._Label1_12.Name = "_Label1_12"
        Me._Label1_12.Size = New System.Drawing.Size(56, 13)
        Me._Label1_12.TabIndex = 14
        Me._Label1_12.Text = "DEM Grid:"
        Me._Label1_12.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_2
        '
        Me._Label1_2.AutoSize = True
        Me._Label1_2.Location = New System.Drawing.Point(135, 62)
        Me._Label1_2.Name = "_Label1_2"
        Me._Label1_2.Size = New System.Drawing.Size(34, 13)
        Me._Label1_2.TabIndex = 13
        Me._Label1_2.Text = "Units:"
        Me._Label1_2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_0
        '
        Me._Label1_0.AutoSize = True
        Me._Label1_0.Location = New System.Drawing.Point(43, 107)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.Size = New System.Drawing.Size(126, 13)
        Me._Label1_0.TabIndex = 12
        Me._Label1_0.Text = "Stream Agreement Layer:"
        Me._Label1_0.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me._Label1_0.Visible = False
        '
        '_Label1_3
        '
        Me._Label1_3.AutoSize = True
        Me._Label1_3.Location = New System.Drawing.Point(25, 18)
        Me._Label1_3.Name = "_Label1_3"
        Me._Label1_3.Size = New System.Drawing.Size(149, 13)
        Me._Label1_3.TabIndex = 11
        Me._Label1_3.Text = "Watershed Delineation Name:"
        Me._Label1_3.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_4
        '
        Me._Label1_4.AutoSize = True
        Me._Label1_4.Location = New System.Drawing.Point(107, 130)
        Me._Label1_4.Name = "_Label1_4"
        Me._Label1_4.Size = New System.Drawing.Size(62, 13)
        Me._Label1_4.TabIndex = 10
        Me._Label1_4.Text = "Watershed:"
        Me._Label1_4.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_5
        '
        Me._Label1_5.AutoSize = True
        Me._Label1_5.Location = New System.Drawing.Point(48, 155)
        Me._Label1_5.Name = "_Label1_5"
        Me._Label1_5.Size = New System.Drawing.Size(121, 13)
        Me._Label1_5.TabIndex = 9
        Me._Label1_5.Text = "Flow Accumulation Grid:"
        Me._Label1_5.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_1
        '
        Me._Label1_1.AutoSize = True
        Me._Label1_1.Location = New System.Drawing.Point(68, 107)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.Size = New System.Drawing.Size(101, 13)
        Me._Label1_1.TabIndex = 8
        Me._Label1_1.Text = "Subwatershed Size:"
        Me._Label1_1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'WatershedDelineationsForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(594, 345)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.MainMenu1)
        Me.Location = New System.Drawing.Point(3, 39)
        Me.Name = "WatershedDelineationsForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Watershed Delineations"
        Me.Controls.SetChildIndex(Me.MainMenu1, 0)
        Me.Controls.SetChildIndex(Me.Frame1, 0)
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
End Class