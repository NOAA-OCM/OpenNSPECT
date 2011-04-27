<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class SoilsSetupForm
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
    Public WithEvents txtMUSLEVal As System.Windows.Forms.TextBox
    Public WithEvents txtMUSLEExp As System.Windows.Forms.TextBox
    Public WithEvents Label12 As System.Windows.Forms.Label
    Public WithEvents Label7 As System.Windows.Forms.Label
    Public WithEvents Label6 As System.Windows.Forms.Label
    Public WithEvents Label11 As System.Windows.Forms.Label
    Public WithEvents Label10 As System.Windows.Forms.Label
    Public WithEvents Label9 As System.Windows.Forms.Label
    Public WithEvents Label8 As System.Windows.Forms.Label
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents txtSoilsName As System.Windows.Forms.TextBox
    Public WithEvents cmdDEMBrowse As System.Windows.Forms.Button
    Public WithEvents txtDEMFile As System.Windows.Forms.TextBox
    Public WithEvents cboSoilFieldsK As System.Windows.Forms.ComboBox
    Public WithEvents cboSoilFields As System.Windows.Forms.ComboBox
    Public WithEvents cmdBrowseFile As System.Windows.Forms.Button
    Public WithEvents txtSoilsDS As System.Windows.Forms.TextBox
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents Label5 As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SoilsSetupForm))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtSoilsName = New System.Windows.Forms.TextBox()
        Me.txtDEMFile = New System.Windows.Forms.TextBox()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.txtMUSLEVal = New System.Windows.Forms.TextBox()
        Me.txtMUSLEExp = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.cmdDEMBrowse = New System.Windows.Forms.Button()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.cboSoilFieldsK = New System.Windows.Forms.ComboBox()
        Me.cboSoilFields = New System.Windows.Forms.ComboBox()
        Me.cmdBrowseFile = New System.Windows.Forms.Button()
        Me.txtSoilsDS = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Frame2.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtSoilsName
        '
        Me.txtSoilsName.AcceptsReturn = True
        Me.txtSoilsName.Location = New System.Drawing.Point(89, 6)
        Me.txtSoilsName.MaxLength = 0
        Me.txtSoilsName.Name = "txtSoilsName"
        Me.txtSoilsName.Size = New System.Drawing.Size(212, 20)
        Me.txtSoilsName.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.txtSoilsName, "Choose the filled DEM you are using.  This will provide Spatial Analyst with the " & _
        "proper analysis environment.")
        '
        'txtDEMFile
        '
        Me.txtDEMFile.AcceptsReturn = True
        Me.txtDEMFile.Location = New System.Drawing.Point(89, 35)
        Me.txtDEMFile.MaxLength = 0
        Me.txtDEMFile.Name = "txtDEMFile"
        Me.txtDEMFile.Size = New System.Drawing.Size(212, 20)
        Me.txtDEMFile.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.txtDEMFile, "Choose the filled DEM you are using.  This will provide Spatial Analyst with the " & _
        "proper analysis environment.")
        '
        'Frame2
        '
        Me.Frame2.Controls.Add(Me.txtMUSLEVal)
        Me.Frame2.Controls.Add(Me.txtMUSLEExp)
        Me.Frame2.Controls.Add(Me.Label12)
        Me.Frame2.Controls.Add(Me.Label7)
        Me.Frame2.Controls.Add(Me.Label6)
        Me.Frame2.Controls.Add(Me.Label11)
        Me.Frame2.Controls.Add(Me.Label10)
        Me.Frame2.Controls.Add(Me.Label9)
        Me.Frame2.Controls.Add(Me.Label8)
        Me.Frame2.Location = New System.Drawing.Point(16, 178)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame2.Size = New System.Drawing.Size(328, 157)
        Me.Frame2.TabIndex = 15
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Advanced MUSLE Specific Coefficients"
        '
        'txtMUSLEVal
        '
        Me.txtMUSLEVal.AcceptsReturn = True
        Me.txtMUSLEVal.Location = New System.Drawing.Point(56, 52)
        Me.txtMUSLEVal.MaxLength = 0
        Me.txtMUSLEVal.Name = "txtMUSLEVal"
        Me.txtMUSLEVal.Size = New System.Drawing.Size(43, 20)
        Me.txtMUSLEVal.TabIndex = 17
        Me.txtMUSLEVal.Text = "95"
        '
        'txtMUSLEExp
        '
        Me.txtMUSLEExp.AcceptsReturn = True
        Me.txtMUSLEExp.Location = New System.Drawing.Point(136, 52)
        Me.txtMUSLEExp.MaxLength = 0
        Me.txtMUSLEExp.Name = "txtMUSLEExp"
        Me.txtMUSLEExp.Size = New System.Drawing.Size(35, 20)
        Me.txtMUSLEExp.TabIndex = 16
        Me.txtMUSLEExp.Text = "0.56"
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(16, 97)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(305, 49)
        Me.Label12.TabIndex = 24
        Me.Label12.Text = "Warning: Q and qp are calculated in English units (acre-feet and cubic feet per s" & _
    "econd respectively). ""a"" and ""b"" must be derived accordingly."
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(112, 59)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(17, 16)
        Me.Label7.TabIndex = 23
        Me.Label7.Text = "b="
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(32, 59)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(17, 16)
        Me.Label6.TabIndex = 22
        Me.Label6.Text = "a="
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(72, 30)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(25, 12)
        Me.Label11.TabIndex = 21
        Me.Label11.Text = "    b"
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(18, 18)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(252, 15)
        Me.Label10.TabIndex = 20
        Me.Label10.Text = "MUSLE Equation for sediment yield:"
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(32, 35)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(212, 19)
        Me.Label9.TabIndex = 19
        Me.Label9.Text = "a * (Q * qp)   * K * C * P * LS"
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(16, 82)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(303, 15)
        Me.Label8.TabIndex = 18
        Me.Label8.Text = "Locally calibrated MUSLE coefficients can be entered above."
        '
        'cmdDEMBrowse
        '
        Me.cmdDEMBrowse.Image = CType(resources.GetObject("cmdDEMBrowse.Image"), System.Drawing.Image)
        Me.cmdDEMBrowse.Location = New System.Drawing.Point(307, 35)
        Me.cmdDEMBrowse.Name = "cmdDEMBrowse"
        Me.cmdDEMBrowse.Size = New System.Drawing.Size(25, 23)
        Me.cmdDEMBrowse.TabIndex = 2
        Me.cmdDEMBrowse.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdDEMBrowse.UseVisualStyleBackColor = True
        '
        'Frame1
        '
        Me.Frame1.Controls.Add(Me.cboSoilFieldsK)
        Me.Frame1.Controls.Add(Me.cboSoilFields)
        Me.Frame1.Controls.Add(Me.cmdBrowseFile)
        Me.Frame1.Controls.Add(Me.txtSoilsDS)
        Me.Frame1.Controls.Add(Me.Label2)
        Me.Frame1.Controls.Add(Me.Label3)
        Me.Frame1.Controls.Add(Me.Label1)
        Me.Frame1.Location = New System.Drawing.Point(16, 64)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.Size = New System.Drawing.Size(328, 111)
        Me.Frame1.TabIndex = 9
        Me.Frame1.TabStop = False
        Me.Frame1.Text = " Soils  "
        '
        'cboSoilFieldsK
        '
        Me.cboSoilFieldsK.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSoilFieldsK.Location = New System.Drawing.Point(173, 77)
        Me.cboSoilFieldsK.Name = "cboSoilFieldsK"
        Me.cboSoilFieldsK.Size = New System.Drawing.Size(143, 21)
        Me.cboSoilFieldsK.TabIndex = 6
        '
        'cboSoilFields
        '
        Me.cboSoilFields.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSoilFields.Location = New System.Drawing.Point(173, 45)
        Me.cboSoilFields.Name = "cboSoilFields"
        Me.cboSoilFields.Size = New System.Drawing.Size(143, 21)
        Me.cboSoilFields.TabIndex = 5
        '
        'cmdBrowseFile
        '
        Me.cmdBrowseFile.Image = CType(resources.GetObject("cmdBrowseFile.Image"), System.Drawing.Image)
        Me.cmdBrowseFile.Location = New System.Drawing.Point(291, 17)
        Me.cmdBrowseFile.Name = "cmdBrowseFile"
        Me.cmdBrowseFile.Size = New System.Drawing.Size(25, 23)
        Me.cmdBrowseFile.TabIndex = 4
        Me.cmdBrowseFile.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdBrowseFile.UseVisualStyleBackColor = True
        '
        'txtSoilsDS
        '
        Me.txtSoilsDS.AcceptsReturn = True
        Me.txtSoilsDS.Location = New System.Drawing.Point(93, 18)
        Me.txtSoilsDS.MaxLength = 0
        Me.txtSoilsDS.Name = "txtSoilsDS"
        Me.txtSoilsDS.Size = New System.Drawing.Size(193, 20)
        Me.txtSoilsDS.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(13, 79)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(122, 18)
        Me.Label2.TabIndex = 13
        Me.Label2.Text = "K Factor Attribute:"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(11, 48)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(159, 28)
        Me.Label3.TabIndex = 11
        Me.Label3.Text = "Hydrologic Soil Group Attribute:"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(12, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(82, 17)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Soils Data Set:"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(18, 8)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(59, 17)
        Me.Label4.TabIndex = 14
        Me.Label4.Text = "Name:"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(17, 38)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(77, 18)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "DEM GRID:"
        '
        'SoilsSetupForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(594, 376)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.txtSoilsName)
        Me.Controls.Add(Me.txtDEMFile)
        Me.Controls.Add(Me.cmdDEMBrowse)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label5)
        Me.Location = New System.Drawing.Point(3, 22)
        Me.Name = "SoilsSetupForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Soils Setup"
        Me.Controls.SetChildIndex(Me.Label5, 0)
        Me.Controls.SetChildIndex(Me.Label4, 0)
        Me.Controls.SetChildIndex(Me.Frame1, 0)
        Me.Controls.SetChildIndex(Me.cmdDEMBrowse, 0)
        Me.Controls.SetChildIndex(Me.txtDEMFile, 0)
        Me.Controls.SetChildIndex(Me.txtSoilsName, 0)
        Me.Controls.SetChildIndex(Me.Frame2, 0)
        Me.Frame2.ResumeLayout(False)
        Me.Frame2.PerformLayout()
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
End Class