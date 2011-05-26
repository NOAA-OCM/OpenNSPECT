<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class CreateNewWatershedDelineationForm
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
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(CreateNewWatershedDelineationForm))
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
        Me.frmMain.SuspendLayout()
        Me.frmAdvanced.SuspendLayout()
        Me.SuspendLayout()
        '
        'frmMain
        '
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
        Me.frmMain.Location = New System.Drawing.Point(13, 7)
        Me.frmMain.Name = "frmMain"
        Me.frmMain.Padding = New System.Windows.Forms.Padding(0)
        Me.frmMain.Size = New System.Drawing.Size(569, 149)
        Me.frmMain.TabIndex = 7
        Me.frmMain.TabStop = False
        Me.frmMain.Text = "Create a new watershed delineation  "
        '
        'cboDEMUnits
        '
        Me.cboDEMUnits.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDEMUnits.Items.AddRange(New Object() {"meters", "feet"})
        Me.cboDEMUnits.Location = New System.Drawing.Point(119, 92)
        Me.cboDEMUnits.Name = "cboDEMUnits"
        Me.cboDEMUnits.Size = New System.Drawing.Size(270, 21)
        Me.cboDEMUnits.TabIndex = 2
        '
        'chkHydroCorr
        '
        Me.chkHydroCorr.Location = New System.Drawing.Point(119, 68)
        Me.chkHydroCorr.Name = "chkHydroCorr"
        Me.chkHydroCorr.Size = New System.Drawing.Size(241, 22)
        Me.chkHydroCorr.TabIndex = 17
        Me.chkHydroCorr.Text = "DEM is hydrologically correct (filled)"
        Me.chkHydroCorr.UseVisualStyleBackColor = True
        '
        'cmdBrowseDEMFile
        '
        Me.cmdBrowseDEMFile.Image = CType(resources.GetObject("cmdBrowseDEMFile.Image"), System.Drawing.Image)
        Me.cmdBrowseDEMFile.Location = New System.Drawing.Point(391, 43)
        Me.cmdBrowseDEMFile.Name = "cmdBrowseDEMFile"
        Me.cmdBrowseDEMFile.Size = New System.Drawing.Size(25, 21)
        Me.cmdBrowseDEMFile.TabIndex = 1
        Me.cmdBrowseDEMFile.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdBrowseDEMFile.UseVisualStyleBackColor = True
        '
        'txtWSDelinName
        '
        Me.txtWSDelinName.AcceptsReturn = True
        Me.txtWSDelinName.Location = New System.Drawing.Point(118, 19)
        Me.txtWSDelinName.MaxLength = 0
        Me.txtWSDelinName.Name = "txtWSDelinName"
        Me.txtWSDelinName.Size = New System.Drawing.Size(271, 20)
        Me.txtWSDelinName.TabIndex = 0
        '
        'txtDEMFile
        '
        Me.txtDEMFile.AcceptsReturn = True
        Me.txtDEMFile.Location = New System.Drawing.Point(118, 45)
        Me.txtDEMFile.MaxLength = 0
        Me.txtDEMFile.Name = "txtDEMFile"
        Me.txtDEMFile.Size = New System.Drawing.Size(271, 20)
        Me.txtDEMFile.TabIndex = 4
        '
        'cboSubWSSize
        '
        Me.cboSubWSSize.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSubWSSize.Items.AddRange(New Object() {"small", "medium", "large"})
        Me.cboSubWSSize.Location = New System.Drawing.Point(119, 120)
        Me.cboSubWSSize.Name = "cboSubWSSize"
        Me.cboSubWSSize.Size = New System.Drawing.Size(270, 21)
        Me.cboSubWSSize.TabIndex = 3
        '
        '_Label1_3
        '
        Me._Label1_3.AutoSize = True
        Me._Label1_3.Location = New System.Drawing.Point(13, 22)
        Me._Label1_3.Name = "_Label1_3"
        Me._Label1_3.Size = New System.Drawing.Size(94, 13)
        Me._Label1_3.TabIndex = 11
        Me._Label1_3.Text = "Delineation Name:"
        '
        '_Label1_2
        '
        Me._Label1_2.AutoSize = True
        Me._Label1_2.Location = New System.Drawing.Point(14, 95)
        Me._Label1_2.Name = "_Label1_2"
        Me._Label1_2.Size = New System.Drawing.Size(61, 13)
        Me._Label1_2.TabIndex = 10
        Me._Label1_2.Text = "DEM Units:"
        '
        '_Label1_12
        '
        Me._Label1_12.AutoSize = True
        Me._Label1_12.Location = New System.Drawing.Point(13, 48)
        Me._Label1_12.Name = "_Label1_12"
        Me._Label1_12.Size = New System.Drawing.Size(56, 13)
        Me._Label1_12.TabIndex = 9
        Me._Label1_12.Text = "DEM Grid:"
        '
        '_Label1_1
        '
        Me._Label1_1.AutoSize = True
        Me._Label1_1.Location = New System.Drawing.Point(14, 123)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.Size = New System.Drawing.Size(101, 13)
        Me._Label1_1.TabIndex = 8
        Me._Label1_1.Text = "Subwatershed Size:"
        Me._Label1_1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'frmAdvanced
        '
        Me.frmAdvanced.Controls.Add(Me.cmdOptions)
        Me.frmAdvanced.Controls.Add(Me.cboStreamLayer)
        Me.frmAdvanced.Controls.Add(Me.chkStreamAgree)
        Me.frmAdvanced.Controls.Add(Me._lblStream_0)
        Me.frmAdvanced.Location = New System.Drawing.Point(12, 162)
        Me.frmAdvanced.Name = "frmAdvanced"
        Me.frmAdvanced.Padding = New System.Windows.Forms.Padding(0)
        Me.frmAdvanced.Size = New System.Drawing.Size(570, 102)
        Me.frmAdvanced.TabIndex = 12
        Me.frmAdvanced.TabStop = False
        Me.frmAdvanced.Text = "Advanced Parameters (optional) "
        Me.frmAdvanced.Visible = False
        '
        'cmdOptions
        '
        Me.cmdOptions.Enabled = False
        Me.cmdOptions.Location = New System.Drawing.Point(281, 66)
        Me.cmdOptions.Name = "cmdOptions"
        Me.cmdOptions.Size = New System.Drawing.Size(67, 23)
        Me.cmdOptions.TabIndex = 16
        Me.cmdOptions.Text = "Options..."
        Me.cmdOptions.UseVisualStyleBackColor = True
        '
        'cboStreamLayer
        '
        Me.cboStreamLayer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboStreamLayer.Enabled = False
        Me.cboStreamLayer.Items.AddRange(New Object() {"Stream1"})
        Me.cboStreamLayer.Location = New System.Drawing.Point(122, 68)
        Me.cboStreamLayer.Name = "cboStreamLayer"
        Me.cboStreamLayer.Size = New System.Drawing.Size(155, 21)
        Me.cboStreamLayer.TabIndex = 14
        '
        'chkStreamAgree
        '
        Me.chkStreamAgree.Enabled = False
        Me.chkStreamAgree.Location = New System.Drawing.Point(48, 45)
        Me.chkStreamAgree.Name = "chkStreamAgree"
        Me.chkStreamAgree.Size = New System.Drawing.Size(149, 18)
        Me.chkStreamAgree.TabIndex = 13
        Me.chkStreamAgree.Text = "Force Stream Agreement"
        Me.chkStreamAgree.UseVisualStyleBackColor = True
        '
        '_lblStream_0
        '
        Me._lblStream_0.AutoSize = True
        Me._lblStream_0.Enabled = False
        Me._lblStream_0.Location = New System.Drawing.Point(49, 71)
        Me._lblStream_0.Name = "_lblStream_0"
        Me._lblStream_0.Size = New System.Drawing.Size(72, 13)
        Me._lblStream_0.TabIndex = 15
        Me._lblStream_0.Text = "Stream Layer:"
        '
        'CreateNewWatershedDelineationForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(594, 345)
        Me.Controls.Add(Me.frmMain)
        Me.Controls.Add(Me.frmAdvanced)
        Me.Location = New System.Drawing.Point(213, 196)
        Me.Name = "CreateNewWatershedDelineationForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Create New Watershed Delineation"
        Me.Controls.SetChildIndex(Me.frmAdvanced, 0)
        Me.Controls.SetChildIndex(Me.frmMain, 0)
        Me.frmMain.ResumeLayout(False)
        Me.frmMain.PerformLayout()
        Me.frmAdvanced.ResumeLayout(False)
        Me.frmAdvanced.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
#End Region
End Class