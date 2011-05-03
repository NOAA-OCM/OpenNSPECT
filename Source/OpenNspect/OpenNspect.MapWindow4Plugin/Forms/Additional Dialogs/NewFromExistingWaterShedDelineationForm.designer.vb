<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class NewFromExistingWaterShedDelineationForm
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
    Public WithEvents cmdBrowseLS As System.Windows.Forms.Button
    Public WithEvents cboDEMUnits As System.Windows.Forms.ComboBox
    Public WithEvents txtFlowDir As System.Windows.Forms.TextBox
    Public WithEvents cmdBrowseFlowDir As System.Windows.Forms.Button
    Public WithEvents cmdBrowseWS As System.Windows.Forms.Button
    Public WithEvents txtWaterSheds As System.Windows.Forms.TextBox
    Public WithEvents cmdBrowseFlowAcc As System.Windows.Forms.Button
    Public WithEvents txtFlowAcc As System.Windows.Forms.TextBox
    Public WithEvents cmdBrowseDEMFile As System.Windows.Forms.Button
    Public WithEvents txtWSDelinName As System.Windows.Forms.TextBox
    Public WithEvents txtDEMFile As System.Windows.Forms.TextBox
    Public WithEvents txtLS As System.Windows.Forms.TextBox
    Public WithEvents _Label1_2 As System.Windows.Forms.Label
    Public WithEvents _Label1_1 As System.Windows.Forms.Label
    Public WithEvents _Label1_5 As System.Windows.Forms.Label
    Public WithEvents _Label1_4 As System.Windows.Forms.Label
    Public WithEvents _Label1_3 As System.Windows.Forms.Label
    Public WithEvents _Label1_0 As System.Windows.Forms.Label
    Public WithEvents _Label1_12 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(NewFromExistingWaterShedDelineationForm))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdBrowseLS = New System.Windows.Forms.Button()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.cboDEMUnits = New System.Windows.Forms.ComboBox()
        Me.txtFlowDir = New System.Windows.Forms.TextBox()
        Me.cmdBrowseFlowDir = New System.Windows.Forms.Button()
        Me.cmdBrowseWS = New System.Windows.Forms.Button()
        Me.txtWaterSheds = New System.Windows.Forms.TextBox()
        Me.cmdBrowseFlowAcc = New System.Windows.Forms.Button()
        Me.txtFlowAcc = New System.Windows.Forms.TextBox()
        Me.cmdBrowseDEMFile = New System.Windows.Forms.Button()
        Me.txtWSDelinName = New System.Windows.Forms.TextBox()
        Me.txtDEMFile = New System.Windows.Forms.TextBox()
        Me.txtLS = New System.Windows.Forms.TextBox()
        Me._Label1_2 = New System.Windows.Forms.Label()
        Me._Label1_1 = New System.Windows.Forms.Label()
        Me._Label1_5 = New System.Windows.Forms.Label()
        Me._Label1_4 = New System.Windows.Forms.Label()
        Me._Label1_3 = New System.Windows.Forms.Label()
        Me._Label1_0 = New System.Windows.Forms.Label()
        Me._Label1_12 = New System.Windows.Forms.Label()
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdBrowseLS
        '
        Me.cmdBrowseLS.Image = CType(resources.GetObject("cmdBrowseLS.Image"), System.Drawing.Image)
        Me.cmdBrowseLS.Location = New System.Drawing.Point(522, 146)
        Me.cmdBrowseLS.Name = "cmdBrowseLS"
        Me.cmdBrowseLS.Size = New System.Drawing.Size(25, 19)
        Me.cmdBrowseLS.TabIndex = 12
        Me.cmdBrowseLS.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdBrowseLS.UseVisualStyleBackColor = True
        '
        'Frame1
        '
        Me.Frame1.Controls.Add(Me.cboDEMUnits)
        Me.Frame1.Controls.Add(Me.txtFlowDir)
        Me.Frame1.Controls.Add(Me.cmdBrowseLS)
        Me.Frame1.Controls.Add(Me.cmdBrowseFlowDir)
        Me.Frame1.Controls.Add(Me.cmdBrowseWS)
        Me.Frame1.Controls.Add(Me.txtWaterSheds)
        Me.Frame1.Controls.Add(Me.cmdBrowseFlowAcc)
        Me.Frame1.Controls.Add(Me.txtFlowAcc)
        Me.Frame1.Controls.Add(Me.cmdBrowseDEMFile)
        Me.Frame1.Controls.Add(Me.txtWSDelinName)
        Me.Frame1.Controls.Add(Me.txtDEMFile)
        Me.Frame1.Controls.Add(Me.txtLS)
        Me.Frame1.Controls.Add(Me._Label1_2)
        Me.Frame1.Controls.Add(Me._Label1_1)
        Me.Frame1.Controls.Add(Me._Label1_5)
        Me.Frame1.Controls.Add(Me._Label1_4)
        Me.Frame1.Controls.Add(Me._Label1_3)
        Me.Frame1.Controls.Add(Me._Label1_0)
        Me.Frame1.Controls.Add(Me._Label1_12)
        Me.Frame1.Location = New System.Drawing.Point(16, 7)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.Size = New System.Drawing.Size(550, 279)
        Me.Frame1.TabIndex = 3
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Define Watershed Delineation"
        '
        'cboDEMUnits
        '
        Me.cboDEMUnits.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDEMUnits.Items.AddRange(New Object() {"meters", "feet"})
        Me.cboDEMUnits.Location = New System.Drawing.Point(170, 71)
        Me.cboDEMUnits.Name = "cboDEMUnits"
        Me.cboDEMUnits.Size = New System.Drawing.Size(153, 21)
        Me.cboDEMUnits.TabIndex = 20
        '
        'txtFlowDir
        '
        Me.txtFlowDir.AcceptsReturn = True
        Me.txtFlowDir.Location = New System.Drawing.Point(170, 95)
        Me.txtFlowDir.MaxLength = 0
        Me.txtFlowDir.Name = "txtFlowDir"
        Me.txtFlowDir.Size = New System.Drawing.Size(346, 20)
        Me.txtFlowDir.TabIndex = 19
        '
        'cmdBrowseFlowDir
        '
        Me.cmdBrowseFlowDir.Image = CType(resources.GetObject("cmdBrowseFlowDir.Image"), System.Drawing.Image)
        Me.cmdBrowseFlowDir.Location = New System.Drawing.Point(522, 93)
        Me.cmdBrowseFlowDir.Name = "cmdBrowseFlowDir"
        Me.cmdBrowseFlowDir.Size = New System.Drawing.Size(25, 19)
        Me.cmdBrowseFlowDir.TabIndex = 18
        Me.cmdBrowseFlowDir.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdBrowseFlowDir.UseVisualStyleBackColor = True
        '
        'cmdBrowseWS
        '
        Me.cmdBrowseWS.Image = CType(resources.GetObject("cmdBrowseWS.Image"), System.Drawing.Image)
        Me.cmdBrowseWS.Location = New System.Drawing.Point(522, 170)
        Me.cmdBrowseWS.Name = "cmdBrowseWS"
        Me.cmdBrowseWS.Size = New System.Drawing.Size(25, 19)
        Me.cmdBrowseWS.TabIndex = 17
        Me.cmdBrowseWS.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdBrowseWS.UseVisualStyleBackColor = True
        '
        'txtWaterSheds
        '
        Me.txtWaterSheds.AcceptsReturn = True
        Me.txtWaterSheds.Location = New System.Drawing.Point(170, 171)
        Me.txtWaterSheds.MaxLength = 0
        Me.txtWaterSheds.Name = "txtWaterSheds"
        Me.txtWaterSheds.Size = New System.Drawing.Size(346, 20)
        Me.txtWaterSheds.TabIndex = 15
        '
        'cmdBrowseFlowAcc
        '
        Me.cmdBrowseFlowAcc.Image = CType(resources.GetObject("cmdBrowseFlowAcc.Image"), System.Drawing.Image)
        Me.cmdBrowseFlowAcc.Location = New System.Drawing.Point(522, 123)
        Me.cmdBrowseFlowAcc.Name = "cmdBrowseFlowAcc"
        Me.cmdBrowseFlowAcc.Size = New System.Drawing.Size(25, 19)
        Me.cmdBrowseFlowAcc.TabIndex = 11
        Me.cmdBrowseFlowAcc.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdBrowseFlowAcc.UseVisualStyleBackColor = True
        '
        'txtFlowAcc
        '
        Me.txtFlowAcc.AcceptsReturn = True
        Me.txtFlowAcc.Location = New System.Drawing.Point(170, 121)
        Me.txtFlowAcc.MaxLength = 0
        Me.txtFlowAcc.Name = "txtFlowAcc"
        Me.txtFlowAcc.Size = New System.Drawing.Size(346, 20)
        Me.txtFlowAcc.TabIndex = 10
        '
        'cmdBrowseDEMFile
        '
        Me.cmdBrowseDEMFile.Image = CType(resources.GetObject("cmdBrowseDEMFile.Image"), System.Drawing.Image)
        Me.cmdBrowseDEMFile.Location = New System.Drawing.Point(522, 47)
        Me.cmdBrowseDEMFile.Name = "cmdBrowseDEMFile"
        Me.cmdBrowseDEMFile.Size = New System.Drawing.Size(25, 19)
        Me.cmdBrowseDEMFile.TabIndex = 2
        Me.cmdBrowseDEMFile.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdBrowseDEMFile.UseVisualStyleBackColor = True
        '
        'txtWSDelinName
        '
        Me.txtWSDelinName.AcceptsReturn = True
        Me.txtWSDelinName.Location = New System.Drawing.Point(170, 22)
        Me.txtWSDelinName.MaxLength = 0
        Me.txtWSDelinName.Name = "txtWSDelinName"
        Me.txtWSDelinName.Size = New System.Drawing.Size(346, 20)
        Me.txtWSDelinName.TabIndex = 0
        '
        'txtDEMFile
        '
        Me.txtDEMFile.AcceptsReturn = True
        Me.txtDEMFile.Location = New System.Drawing.Point(170, 48)
        Me.txtDEMFile.MaxLength = 0
        Me.txtDEMFile.Name = "txtDEMFile"
        Me.txtDEMFile.Size = New System.Drawing.Size(346, 20)
        Me.txtDEMFile.TabIndex = 1
        '
        'txtLS
        '
        Me.txtLS.AcceptsReturn = True
        Me.txtLS.Location = New System.Drawing.Point(170, 146)
        Me.txtLS.MaxLength = 0
        Me.txtLS.Name = "txtLS"
        Me.txtLS.Size = New System.Drawing.Size(346, 20)
        Me.txtLS.TabIndex = 4
        '
        '_Label1_2
        '
        Me._Label1_2.AutoSize = True
        Me._Label1_2.Location = New System.Drawing.Point(96, 74)
        Me._Label1_2.Name = "_Label1_2"
        Me._Label1_2.Size = New System.Drawing.Size(61, 13)
        Me._Label1_2.TabIndex = 21
        Me._Label1_2.Text = "DEM Units:"
        '
        '_Label1_1
        '
        Me._Label1_1.AutoSize = True
        Me._Label1_1.Location = New System.Drawing.Point(94, 171)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.Size = New System.Drawing.Size(67, 13)
        Me._Label1_1.TabIndex = 16
        Me._Label1_1.Text = "Watersheds:"
        Me._Label1_1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_5
        '
        Me._Label1_5.AutoSize = True
        Me._Label1_5.Location = New System.Drawing.Point(68, 148)
        Me._Label1_5.Name = "_Label1_5"
        Me._Label1_5.Size = New System.Drawing.Size(93, 13)
        Me._Label1_5.TabIndex = 9
        Me._Label1_5.Text = "Length-slope Grid:"
        Me._Label1_5.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_4
        '
        Me._Label1_4.AutoSize = True
        Me._Label1_4.Location = New System.Drawing.Point(40, 123)
        Me._Label1_4.Name = "_Label1_4"
        Me._Label1_4.Size = New System.Drawing.Size(121, 13)
        Me._Label1_4.TabIndex = 8
        Me._Label1_4.Text = "Flow Accumulation Grid:"
        Me._Label1_4.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_3
        '
        Me._Label1_3.AutoSize = True
        Me._Label1_3.Location = New System.Drawing.Point(17, 25)
        Me._Label1_3.Name = "_Label1_3"
        Me._Label1_3.Size = New System.Drawing.Size(149, 13)
        Me._Label1_3.TabIndex = 7
        Me._Label1_3.Text = "Watershed Delineation Name:"
        Me._Label1_3.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_0
        '
        Me._Label1_0.AutoSize = True
        Me._Label1_0.Location = New System.Drawing.Point(62, 99)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.Size = New System.Drawing.Size(99, 13)
        Me._Label1_0.TabIndex = 6
        Me._Label1_0.Text = "Flow Direction Grid:"
        Me._Label1_0.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_Label1_12
        '
        Me._Label1_12.AutoSize = True
        Me._Label1_12.Location = New System.Drawing.Point(104, 52)
        Me._Label1_12.Name = "_Label1_12"
        Me._Label1_12.Size = New System.Drawing.Size(56, 13)
        Me._Label1_12.TabIndex = 5
        Me._Label1_12.Text = "DEM Grid:"
        Me._Label1_12.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'UserWaterShedDelineationForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(584, 336)
        Me.Controls.Add(Me.Frame1)
        Me.Location = New System.Drawing.Point(4, 30)
        Me.Name = "UserWaterShedDelineationForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "New Watershed Delineation"
        Me.Controls.SetChildIndex(Me.Frame1, 0)
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
#End Region
End Class