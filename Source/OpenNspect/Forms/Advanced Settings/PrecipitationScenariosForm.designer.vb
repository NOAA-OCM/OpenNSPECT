<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class PrecipitationScenariosForm
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
    Public WithEvents mnuNewPrecip As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuDelPrecip As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuPrecipOpts As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuPrecipHelp As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuHelp As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
    Public WithEvents cboScenName As System.Windows.Forms.ComboBox
    Public WithEvents txtRainingDays As System.Windows.Forms.TextBox
    Public WithEvents cboTimePeriod As System.Windows.Forms.ComboBox
    Public WithEvents cboPrecipType As System.Windows.Forms.ComboBox
    Public WithEvents cmdBrowseFile As System.Windows.Forms.Button
    Public WithEvents txtDesc As System.Windows.Forms.TextBox
    Public WithEvents txtPrecipFile As System.Windows.Forms.TextBox
    Public WithEvents cboGridUnits As System.Windows.Forms.ComboBox
    Public WithEvents cboPrecipUnits As System.Windows.Forms.ComboBox
    Public WithEvents _Label1_0 As System.Windows.Forms.Label
    Public WithEvents lblRainingDays As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents _Label1_2 As System.Windows.Forms.Label
    Public WithEvents _Label1_1 As System.Windows.Forms.Label
    Public WithEvents _Label1_6 As System.Windows.Forms.Label
    Public WithEvents _Label1_7 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    'Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PrecipitationScenariosForm))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip()
        Me.mnuPrecipOpts = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuNewPrecip = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuDelPrecip = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuHelp = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuPrecipHelp = New System.Windows.Forms.ToolStripMenuItem()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.cboScenName = New System.Windows.Forms.ComboBox()
        Me.txtRainingDays = New System.Windows.Forms.TextBox()
        Me.cboTimePeriod = New System.Windows.Forms.ComboBox()
        Me.cboPrecipType = New System.Windows.Forms.ComboBox()
        Me.cmdBrowseFile = New System.Windows.Forms.Button()
        Me.txtDesc = New System.Windows.Forms.TextBox()
        Me.txtPrecipFile = New System.Windows.Forms.TextBox()
        Me.cboGridUnits = New System.Windows.Forms.ComboBox()
        Me.cboPrecipUnits = New System.Windows.Forms.ComboBox()
        Me._Label1_0 = New System.Windows.Forms.Label()
        Me.lblRainingDays = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me._Label1_2 = New System.Windows.Forms.Label()
        Me._Label1_1 = New System.Windows.Forms.Label()
        Me._Label1_6 = New System.Windows.Forms.Label()
        Me._Label1_7 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.MainMenu1.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuPrecipOpts, Me.mnuHelp})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(594, 24)
        Me.MainMenu1.TabIndex = 8
        '
        'mnuPrecipOpts
        '
        Me.mnuPrecipOpts.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuNewPrecip, Me.mnuDelPrecip})
        Me.mnuPrecipOpts.Name = "mnuPrecipOpts"
        Me.mnuPrecipOpts.Size = New System.Drawing.Size(61, 20)
        Me.mnuPrecipOpts.Text = "Options"
        '
        'mnuNewPrecip
        '
        Me.mnuNewPrecip.Name = "mnuNewPrecip"
        Me.mnuNewPrecip.Size = New System.Drawing.Size(116, 22)
        Me.mnuNewPrecip.Text = "New..."
        '
        'mnuDelPrecip
        '
        Me.mnuDelPrecip.Name = "mnuDelPrecip"
        Me.mnuDelPrecip.Size = New System.Drawing.Size(116, 22)
        Me.mnuDelPrecip.Text = "Delete..."
        '
        'mnuHelp
        '
        Me.mnuHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuPrecipHelp})
        Me.mnuHelp.Name = "mnuHelp"
        Me.mnuHelp.Size = New System.Drawing.Size(44, 20)
        Me.mnuHelp.Text = "&Help"
        '
        'mnuPrecipHelp
        '
        Me.mnuPrecipHelp.Name = "mnuPrecipHelp"
        Me.mnuPrecipHelp.ShortcutKeys = CType((System.Windows.Forms.Keys.Shift Or System.Windows.Forms.Keys.F1), System.Windows.Forms.Keys)
        Me.mnuPrecipHelp.Size = New System.Drawing.Size(254, 22)
        Me.mnuPrecipHelp.Text = "Precipitation Scenarios..."
        '
        'Frame1
        '
        Me.Frame1.Controls.Add(Me.cboScenName)
        Me.Frame1.Controls.Add(Me.txtRainingDays)
        Me.Frame1.Controls.Add(Me.cboTimePeriod)
        Me.Frame1.Controls.Add(Me.cboPrecipType)
        Me.Frame1.Controls.Add(Me.cmdBrowseFile)
        Me.Frame1.Controls.Add(Me.txtDesc)
        Me.Frame1.Controls.Add(Me.txtPrecipFile)
        Me.Frame1.Controls.Add(Me.cboGridUnits)
        Me.Frame1.Controls.Add(Me.cboPrecipUnits)
        Me.Frame1.Controls.Add(Me._Label1_0)
        Me.Frame1.Controls.Add(Me.lblRainingDays)
        Me.Frame1.Controls.Add(Me.Label3)
        Me.Frame1.Controls.Add(Me._Label1_2)
        Me.Frame1.Controls.Add(Me._Label1_1)
        Me.Frame1.Controls.Add(Me._Label1_6)
        Me.Frame1.Controls.Add(Me._Label1_7)
        Me.Frame1.Controls.Add(Me.Label2)
        Me.Frame1.Location = New System.Drawing.Point(12, 24)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.Size = New System.Drawing.Size(555, 270)
        Me.Frame1.TabIndex = 6
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Choose a precipitation scenario to view or edit  "
        '
        'cboScenName
        '
        Me.cboScenName.Location = New System.Drawing.Point(139, 22)
        Me.cboScenName.Name = "cboScenName"
        Me.cboScenName.Size = New System.Drawing.Size(143, 21)
        Me.cboScenName.TabIndex = 18
        '
        'txtRainingDays
        '
        Me.txtRainingDays.AcceptsReturn = True
        Me.txtRainingDays.Location = New System.Drawing.Point(366, 161)
        Me.txtRainingDays.MaxLength = 0
        Me.txtRainingDays.Name = "txtRainingDays"
        Me.txtRainingDays.Size = New System.Drawing.Size(46, 20)
        Me.txtRainingDays.TabIndex = 17
        '
        'cboTimePeriod
        '
        Me.cboTimePeriod.Items.AddRange(New Object() {"Annual", "Event"})
        Me.cboTimePeriod.Location = New System.Drawing.Point(139, 161)
        Me.cboTimePeriod.Name = "cboTimePeriod"
        Me.cboTimePeriod.Size = New System.Drawing.Size(143, 21)
        Me.cboTimePeriod.TabIndex = 15
        '
        'cboPrecipType
        '
        Me.cboPrecipType.Items.AddRange(New Object() {"Type I", "Type IA", "Type II", "Type III"})
        Me.cboPrecipType.Location = New System.Drawing.Point(139, 188)
        Me.cboPrecipType.Name = "cboPrecipType"
        Me.cboPrecipType.Size = New System.Drawing.Size(143, 21)
        Me.cboPrecipType.TabIndex = 8
        '
        'cmdBrowseFile
        '
        Me.cmdBrowseFile.Image = CType(resources.GetObject("cmdBrowseFile.Image"), System.Drawing.Image)
        Me.cmdBrowseFile.Location = New System.Drawing.Point(398, 76)
        Me.cmdBrowseFile.Name = "cmdBrowseFile"
        Me.cmdBrowseFile.Size = New System.Drawing.Size(25, 19)
        Me.cmdBrowseFile.TabIndex = 1
        Me.cmdBrowseFile.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdBrowseFile.UseVisualStyleBackColor = True
        '
        'txtDesc
        '
        Me.txtDesc.AcceptsReturn = True
        Me.txtDesc.Location = New System.Drawing.Point(139, 50)
        Me.txtDesc.MaxLength = 0
        Me.txtDesc.Name = "txtDesc"
        Me.txtDesc.Size = New System.Drawing.Size(255, 20)
        Me.txtDesc.TabIndex = 0
        '
        'txtPrecipFile
        '
        Me.txtPrecipFile.AcceptsReturn = True
        Me.txtPrecipFile.Location = New System.Drawing.Point(139, 78)
        Me.txtPrecipFile.MaxLength = 0
        Me.txtPrecipFile.Name = "txtPrecipFile"
        Me.txtPrecipFile.Size = New System.Drawing.Size(255, 20)
        Me.txtPrecipFile.TabIndex = 2
        '
        'cboGridUnits
        '
        Me.cboGridUnits.Items.AddRange(New Object() {"meters", "feet"})
        Me.cboGridUnits.Location = New System.Drawing.Point(139, 105)
        Me.cboGridUnits.Name = "cboGridUnits"
        Me.cboGridUnits.Size = New System.Drawing.Size(142, 21)
        Me.cboGridUnits.TabIndex = 3
        '
        'cboPrecipUnits
        '
        Me.cboPrecipUnits.Items.AddRange(New Object() {"centimeters", "inches"})
        Me.cboPrecipUnits.Location = New System.Drawing.Point(139, 133)
        Me.cboPrecipUnits.Name = "cboPrecipUnits"
        Me.cboPrecipUnits.Size = New System.Drawing.Size(143, 21)
        Me.cboPrecipUnits.Sorted = True
        Me.cboPrecipUnits.TabIndex = 4
        '
        '_Label1_0
        '
        Me._Label1_0.Location = New System.Drawing.Point(26, 23)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.Size = New System.Drawing.Size(102, 18)
        Me._Label1_0.TabIndex = 19
        Me._Label1_0.Text = "Scenario Name:"
        Me._Label1_0.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblRainingDays
        '
        Me.lblRainingDays.Location = New System.Drawing.Point(294, 162)
        Me.lblRainingDays.Name = "lblRainingDays"
        Me.lblRainingDays.Size = New System.Drawing.Size(72, 19)
        Me.lblRainingDays.TabIndex = 16
        Me.lblRainingDays.Text = "Raining Days: "
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(47, 189)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(81, 16)
        Me.Label3.TabIndex = 14
        Me.Label3.Text = "Type:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_2
        '
        Me._Label1_2.Location = New System.Drawing.Point(6, 134)
        Me._Label1_2.Name = "_Label1_2"
        Me._Label1_2.Size = New System.Drawing.Size(122, 18)
        Me._Label1_2.TabIndex = 13
        Me._Label1_2.Text = "Precipitation Units:"
        Me._Label1_2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_1
        '
        Me._Label1_1.Location = New System.Drawing.Point(39, 51)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.Size = New System.Drawing.Size(89, 18)
        Me._Label1_1.TabIndex = 12
        Me._Label1_1.Text = "Description:"
        Me._Label1_1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_6
        '
        Me._Label1_6.Location = New System.Drawing.Point(54, 106)
        Me._Label1_6.Name = "_Label1_6"
        Me._Label1_6.Size = New System.Drawing.Size(74, 18)
        Me._Label1_6.TabIndex = 11
        Me._Label1_6.Text = "Grid Units:"
        Me._Label1_6.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_7
        '
        Me._Label1_7.Location = New System.Drawing.Point(31, 79)
        Me._Label1_7.Name = "_Label1_7"
        Me._Label1_7.Size = New System.Drawing.Size(97, 18)
        Me._Label1_7.TabIndex = 10
        Me._Label1_7.Text = "Precipitation Grid:"
        Me._Label1_7.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(47, 162)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(81, 16)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "Time Period:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'PrecipitationScenariosForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(594, 345)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.MainMenu1)
        Me.Location = New System.Drawing.Point(192, 246)
        Me.Name = "PrecipitationScenariosForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Precipitation Scenarios"
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