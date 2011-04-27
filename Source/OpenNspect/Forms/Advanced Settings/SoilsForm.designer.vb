<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class SoilsForm
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
    Public WithEvents mnuNew As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuDelete As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuOptions As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuSoilsHelp As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuHelp As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
    Public WithEvents txtSoilsKGrid As System.Windows.Forms.TextBox
    Public WithEvents txtSoilsGrid As System.Windows.Forms.TextBox
    Public WithEvents cboSoils As System.Windows.Forms.ComboBox
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip()
        Me.mnuOptions = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuNew = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuDelete = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuHelp = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuSoilsHelp = New System.Windows.Forms.ToolStripMenuItem()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.txtSoilsKGrid = New System.Windows.Forms.TextBox()
        Me.txtSoilsGrid = New System.Windows.Forms.TextBox()
        Me.cboSoils = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.MainMenu1.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuOptions, Me.mnuHelp})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(594, 24)
        Me.MainMenu1.TabIndex = 3
        '
        'mnuOptions
        '
        Me.mnuOptions.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuNew, Me.mnuDelete})
        Me.mnuOptions.Name = "mnuOptions"
        Me.mnuOptions.Size = New System.Drawing.Size(61, 20)
        Me.mnuOptions.Text = "Options"
        '
        'mnuNew
        '
        Me.mnuNew.Name = "mnuNew"
        Me.mnuNew.Size = New System.Drawing.Size(116, 22)
        Me.mnuNew.Text = "New..."
        '
        'mnuDelete
        '
        Me.mnuDelete.Name = "mnuDelete"
        Me.mnuDelete.Size = New System.Drawing.Size(116, 22)
        Me.mnuDelete.Text = "Delete..."
        '
        'mnuHelp
        '
        Me.mnuHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuSoilsHelp})
        Me.mnuHelp.Name = "mnuHelp"
        Me.mnuHelp.Size = New System.Drawing.Size(44, 20)
        Me.mnuHelp.Text = "&Help"
        '
        'mnuSoilsHelp
        '
        Me.mnuSoilsHelp.Name = "mnuSoilsHelp"
        Me.mnuSoilsHelp.ShortcutKeys = CType((System.Windows.Forms.Keys.Shift Or System.Windows.Forms.Keys.F1), System.Windows.Forms.Keys)
        Me.mnuSoilsHelp.Size = New System.Drawing.Size(158, 22)
        Me.mnuSoilsHelp.Text = "Soils..."
        '
        'Frame1
        '
        Me.Frame1.Controls.Add(Me.txtSoilsKGrid)
        Me.Frame1.Controls.Add(Me.txtSoilsGrid)
        Me.Frame1.Controls.Add(Me.cboSoils)
        Me.Frame1.Controls.Add(Me.Label5)
        Me.Frame1.Controls.Add(Me.Label4)
        Me.Frame1.Controls.Add(Me.Label1)
        Me.Frame1.Location = New System.Drawing.Point(13, 23)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.Size = New System.Drawing.Size(569, 282)
        Me.Frame1.TabIndex = 2
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Soils Configuration  "
        '
        'txtSoilsKGrid
        '
        Me.txtSoilsKGrid.AcceptsReturn = True
        Me.txtSoilsKGrid.Location = New System.Drawing.Point(75, 80)
        Me.txtSoilsKGrid.MaxLength = 0
        Me.txtSoilsKGrid.Name = "txtSoilsKGrid"
        Me.txtSoilsKGrid.Size = New System.Drawing.Size(421, 20)
        Me.txtSoilsKGrid.TabIndex = 5
        '
        'txtSoilsGrid
        '
        Me.txtSoilsGrid.AcceptsReturn = True
        Me.txtSoilsGrid.Location = New System.Drawing.Point(75, 49)
        Me.txtSoilsGrid.MaxLength = 0
        Me.txtSoilsGrid.Name = "txtSoilsGrid"
        Me.txtSoilsGrid.Size = New System.Drawing.Size(421, 20)
        Me.txtSoilsGrid.TabIndex = 4
        '
        'cboSoils
        '
        Me.cboSoils.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSoils.Location = New System.Drawing.Point(76, 19)
        Me.cboSoils.Name = "cboSoils"
        Me.cboSoils.Size = New System.Drawing.Size(420, 21)
        Me.cboSoils.TabIndex = 3
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(7, 82)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(77, 18)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "Soils K Grid:"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(7, 51)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(59, 17)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Soils GRID:"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 21)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(53, 17)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Name:"
        '
        'SoilsForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(594, 345)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.MainMenu1)
        Me.Location = New System.Drawing.Point(11, 30)
        Me.Name = "SoilsForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Soils"
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