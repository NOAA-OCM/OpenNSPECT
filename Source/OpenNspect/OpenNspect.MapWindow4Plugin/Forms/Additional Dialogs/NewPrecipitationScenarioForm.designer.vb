<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class NewPrecipitationScenarioForm
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
    Public WithEvents txtRainingDays As System.Windows.Forms.TextBox
    Public WithEvents cboTimePeriod As System.Windows.Forms.ComboBox
    Public WithEvents cboPrecipType As System.Windows.Forms.ComboBox
    Public WithEvents txtPrecipName As System.Windows.Forms.TextBox
    Public WithEvents txtDesc As System.Windows.Forms.TextBox
    Public WithEvents txtPrecipFile As System.Windows.Forms.TextBox
    Public WithEvents cmdBrowseFile As System.Windows.Forms.Button
    Public WithEvents cboPrecipUnits As System.Windows.Forms.ComboBox
    Public WithEvents cboGridUnits As System.Windows.Forms.ComboBox
    Public WithEvents lblRainingDays As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents _Label1_7 As System.Windows.Forms.Label
    Public WithEvents _Label1_1 As System.Windows.Forms.Label
    Public WithEvents _Label1_0 As System.Windows.Forms.Label
    Public WithEvents _Label1_2 As System.Windows.Forms.Label
    Public WithEvents _Label1_6 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(NewPrecipitationScenarioForm))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.txtRainingDays = New System.Windows.Forms.TextBox()
        Me.cboTimePeriod = New System.Windows.Forms.ComboBox()
        Me.cboPrecipType = New System.Windows.Forms.ComboBox()
        Me.txtPrecipName = New System.Windows.Forms.TextBox()
        Me.txtDesc = New System.Windows.Forms.TextBox()
        Me.txtPrecipFile = New System.Windows.Forms.TextBox()
        Me.cmdBrowseFile = New System.Windows.Forms.Button()
        Me.cboPrecipUnits = New System.Windows.Forms.ComboBox()
        Me.cboGridUnits = New System.Windows.Forms.ComboBox()
        Me.lblRainingDays = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me._Label1_7 = New System.Windows.Forms.Label()
        Me._Label1_1 = New System.Windows.Forms.Label()
        Me._Label1_0 = New System.Windows.Forms.Label()
        Me._Label1_2 = New System.Windows.Forms.Label()
        Me._Label1_6 = New System.Windows.Forms.Label()
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Frame1
        '
        Me.Frame1.Controls.Add(Me.txtRainingDays)
        Me.Frame1.Controls.Add(Me.cboTimePeriod)
        Me.Frame1.Controls.Add(Me.cboPrecipType)
        Me.Frame1.Controls.Add(Me.txtPrecipName)
        Me.Frame1.Controls.Add(Me.txtDesc)
        Me.Frame1.Controls.Add(Me.txtPrecipFile)
        Me.Frame1.Controls.Add(Me.cmdBrowseFile)
        Me.Frame1.Controls.Add(Me.cboPrecipUnits)
        Me.Frame1.Controls.Add(Me.cboGridUnits)
        Me.Frame1.Controls.Add(Me.lblRainingDays)
        Me.Frame1.Controls.Add(Me.Label2)
        Me.Frame1.Controls.Add(Me.Label3)
        Me.Frame1.Controls.Add(Me._Label1_7)
        Me.Frame1.Controls.Add(Me._Label1_1)
        Me.Frame1.Controls.Add(Me._Label1_0)
        Me.Frame1.Controls.Add(Me._Label1_2)
        Me.Frame1.Controls.Add(Me._Label1_6)
        Me.Frame1.Location = New System.Drawing.Point(8, 6)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.Size = New System.Drawing.Size(570, 294)
        Me.Frame1.TabIndex = 8
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Enter new scenario information  "
        '
        'txtRainingDays
        '
        Me.txtRainingDays.AcceptsReturn = True
        Me.txtRainingDays.Location = New System.Drawing.Point(337, 165)
        Me.txtRainingDays.MaxLength = 0
        Me.txtRainingDays.Name = "txtRainingDays"
        Me.txtRainingDays.Size = New System.Drawing.Size(46, 20)
        Me.txtRainingDays.TabIndex = 19
        Me.txtRainingDays.Visible = False
        '
        'cboTimePeriod
        '
        Me.cboTimePeriod.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTimePeriod.Items.AddRange(New Object() {"Annual", "Event"})
        Me.cboTimePeriod.Location = New System.Drawing.Point(112, 164)
        Me.cboTimePeriod.Name = "cboTimePeriod"
        Me.cboTimePeriod.Size = New System.Drawing.Size(143, 21)
        Me.cboTimePeriod.TabIndex = 16
        '
        'cboPrecipType
        '
        Me.cboPrecipType.Items.AddRange(New Object() {"Type I", "Type IA", "Type II", "Type III"})
        Me.cboPrecipType.Location = New System.Drawing.Point(112, 194)
        Me.cboPrecipType.Name = "cboPrecipType"
        Me.cboPrecipType.Size = New System.Drawing.Size(143, 21)
        Me.cboPrecipType.TabIndex = 14
        '
        'txtPrecipName
        '
        Me.txtPrecipName.AcceptsReturn = True
        Me.txtPrecipName.Location = New System.Drawing.Point(112, 23)
        Me.txtPrecipName.MaxLength = 0
        Me.txtPrecipName.Name = "txtPrecipName"
        Me.txtPrecipName.Size = New System.Drawing.Size(109, 20)
        Me.txtPrecipName.TabIndex = 0
        '
        'txtDesc
        '
        Me.txtDesc.AcceptsReturn = True
        Me.txtDesc.Location = New System.Drawing.Point(112, 50)
        Me.txtDesc.MaxLength = 0
        Me.txtDesc.Name = "txtDesc"
        Me.txtDesc.Size = New System.Drawing.Size(269, 20)
        Me.txtDesc.TabIndex = 1
        '
        'txtPrecipFile
        '
        Me.txtPrecipFile.AcceptsReturn = True
        Me.txtPrecipFile.Location = New System.Drawing.Point(112, 76)
        Me.txtPrecipFile.MaxLength = 0
        Me.txtPrecipFile.Name = "txtPrecipFile"
        Me.txtPrecipFile.Size = New System.Drawing.Size(269, 20)
        Me.txtPrecipFile.TabIndex = 3
        '
        'cmdBrowseFile
        '
        Me.cmdBrowseFile.Image = CType(resources.GetObject("cmdBrowseFile.Image"), System.Drawing.Image)
        Me.cmdBrowseFile.Location = New System.Drawing.Point(385, 75)
        Me.cmdBrowseFile.Name = "cmdBrowseFile"
        Me.cmdBrowseFile.Size = New System.Drawing.Size(25, 21)
        Me.cmdBrowseFile.TabIndex = 2
        Me.cmdBrowseFile.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdBrowseFile.UseVisualStyleBackColor = True
        '
        'cboPrecipUnits
        '
        Me.cboPrecipUnits.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPrecipUnits.Items.AddRange(New Object() {"centimeters", "inches"})
        Me.cboPrecipUnits.Location = New System.Drawing.Point(112, 134)
        Me.cboPrecipUnits.Name = "cboPrecipUnits"
        Me.cboPrecipUnits.Size = New System.Drawing.Size(143, 21)
        Me.cboPrecipUnits.TabIndex = 5
        '
        'cboGridUnits
        '
        Me.cboGridUnits.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboGridUnits.Items.AddRange(New Object() {"meters", "feet"})
        Me.cboGridUnits.Location = New System.Drawing.Point(112, 104)
        Me.cboGridUnits.Name = "cboGridUnits"
        Me.cboGridUnits.Size = New System.Drawing.Size(143, 21)
        Me.cboGridUnits.TabIndex = 4
        '
        'lblRainingDays
        '
        Me.lblRainingDays.Location = New System.Drawing.Point(265, 167)
        Me.lblRainingDays.Name = "lblRainingDays"
        Me.lblRainingDays.Size = New System.Drawing.Size(72, 19)
        Me.lblRainingDays.TabIndex = 18
        Me.lblRainingDays.Text = "Raining Days: "
        Me.lblRainingDays.Visible = False
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(21, 166)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(81, 16)
        Me.Label2.TabIndex = 17
        Me.Label2.Text = "Time Period:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(21, 195)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(81, 16)
        Me.Label3.TabIndex = 15
        Me.Label3.Text = "Type:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_7
        '
        Me._Label1_7.Location = New System.Drawing.Point(23, 23)
        Me._Label1_7.Name = "_Label1_7"
        Me._Label1_7.Size = New System.Drawing.Size(79, 16)
        Me._Label1_7.TabIndex = 13
        Me._Label1_7.Text = "Scenario Name:"
        Me._Label1_7.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_1
        '
        Me._Label1_1.Location = New System.Drawing.Point(5, 50)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.Size = New System.Drawing.Size(97, 18)
        Me._Label1_1.TabIndex = 12
        Me._Label1_1.Text = "Description:"
        Me._Label1_1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_0
        '
        Me._Label1_0.Location = New System.Drawing.Point(5, 77)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.Size = New System.Drawing.Size(97, 18)
        Me._Label1_0.TabIndex = 11
        Me._Label1_0.Text = "Precipitation Grid:"
        Me._Label1_0.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_2
        '
        Me._Label1_2.Location = New System.Drawing.Point(5, 133)
        Me._Label1_2.Name = "_Label1_2"
        Me._Label1_2.Size = New System.Drawing.Size(97, 18)
        Me._Label1_2.TabIndex = 10
        Me._Label1_2.Text = "Precipitation Units:"
        Me._Label1_2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_6
        '
        Me._Label1_6.Location = New System.Drawing.Point(5, 104)
        Me._Label1_6.Name = "_Label1_6"
        Me._Label1_6.Size = New System.Drawing.Size(97, 18)
        Me._Label1_6.TabIndex = 9
        Me._Label1_6.Text = "Grid Units:"
        Me._Label1_6.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'NewPrecipitationScenarioForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(594, 345)
        Me.Controls.Add(Me.Frame1)
        Me.Location = New System.Drawing.Point(3, 21)
        Me.Name = "NewPrecipitationScenarioForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "New Precipitation Scenario"
        Me.Controls.SetChildIndex(Me.Frame1, 0)
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
#End Region
End Class