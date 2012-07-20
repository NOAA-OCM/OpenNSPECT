<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class CompareOutputsForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(CompareOutputsForm))
        Me.pnlTop = New System.Windows.Forms.Panel()
        Me.spltconBase = New System.Windows.Forms.SplitContainer()
        Me.lstbxLeft = New System.Windows.Forms.ListBox()
        Me.chkbxLeftUseLegend = New System.Windows.Forms.CheckBox()
        Me.lstbxRight = New System.Windows.Forms.ListBox()
        Me.chkbxRightUseLegend = New System.Windows.Forms.CheckBox()
        Me.cmdRun = New System.Windows.Forms.Button()
        Me.cmdQuit = New System.Windows.Forms.Button()
        Me.mnustrpMain = New System.Windows.Forms.MenuStrip()
        Me.mnuFile = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuitmAddToLegend = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuitmExit = New System.Windows.Forms.ToolStripMenuItem()
        Me.chkSelectedPolys = New System.Windows.Forms.CheckBox()
        Me.btnSelect = New System.Windows.Forms.Button()
        Me.lblSelected = New System.Windows.Forms.Label()
        Me.lblOriginal = New System.Windows.Forms.Label()
        Me.lblModified = New System.Windows.Forms.Label()
        Me.lblCompareInstruction = New System.Windows.Forms.Label()
        Me.pnlTop.SuspendLayout()
        Me.spltconBase.Panel1.SuspendLayout()
        Me.spltconBase.Panel2.SuspendLayout()
        Me.spltconBase.SuspendLayout()
        Me.mnustrpMain.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlTop
        '
        Me.pnlTop.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlTop.Controls.Add(Me.spltconBase)
        Me.pnlTop.Location = New System.Drawing.Point(0, 27)
        Me.pnlTop.Name = "pnlTop"
        Me.pnlTop.Size = New System.Drawing.Size(584, 299)
        Me.pnlTop.TabIndex = 0
        '
        'spltconBase
        '
        Me.spltconBase.Dock = System.Windows.Forms.DockStyle.Fill
        Me.spltconBase.Location = New System.Drawing.Point(0, 0)
        Me.spltconBase.Name = "spltconBase"
        '
        'spltconBase.Panel1
        '
        Me.spltconBase.Panel1.Controls.Add(Me.lblModified)
        Me.spltconBase.Panel1.Controls.Add(Me.lstbxLeft)
        Me.spltconBase.Panel1.Controls.Add(Me.chkbxLeftUseLegend)
        Me.spltconBase.Panel1MinSize = 100
        '
        'spltconBase.Panel2
        '
        Me.spltconBase.Panel2.Controls.Add(Me.lblOriginal)
        Me.spltconBase.Panel2.Controls.Add(Me.lstbxRight)
        Me.spltconBase.Panel2.Controls.Add(Me.chkbxRightUseLegend)
        Me.spltconBase.Panel2MinSize = 100
        Me.spltconBase.Size = New System.Drawing.Size(584, 299)
        Me.spltconBase.SplitterDistance = 288
        Me.spltconBase.TabIndex = 0
        '
        'lstbxLeft
        '
        Me.lstbxLeft.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstbxLeft.FormattingEnabled = True
        Me.lstbxLeft.Location = New System.Drawing.Point(12, 33)
        Me.lstbxLeft.Name = "lstbxLeft"
        Me.lstbxLeft.Size = New System.Drawing.Size(273, 238)
        Me.lstbxLeft.TabIndex = 1
        '
        'chkbxLeftUseLegend
        '
        Me.chkbxLeftUseLegend.AutoSize = True
        Me.chkbxLeftUseLegend.Location = New System.Drawing.Point(12, 277)
        Me.chkbxLeftUseLegend.Name = "chkbxLeftUseLegend"
        Me.chkbxLeftUseLegend.Size = New System.Drawing.Size(206, 17)
        Me.chkbxLeftUseLegend.TabIndex = 0
        Me.chkbxLeftUseLegend.Text = "Show Only Output Group from Legend"
        Me.chkbxLeftUseLegend.UseVisualStyleBackColor = True
        '
        'lstbxRight
        '
        Me.lstbxRight.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstbxRight.FormattingEnabled = True
        Me.lstbxRight.Location = New System.Drawing.Point(4, 33)
        Me.lstbxRight.Name = "lstbxRight"
        Me.lstbxRight.Size = New System.Drawing.Size(278, 238)
        Me.lstbxRight.TabIndex = 2
        '
        'chkbxRightUseLegend
        '
        Me.chkbxRightUseLegend.AutoSize = True
        Me.chkbxRightUseLegend.Location = New System.Drawing.Point(4, 277)
        Me.chkbxRightUseLegend.Name = "chkbxRightUseLegend"
        Me.chkbxRightUseLegend.Size = New System.Drawing.Size(206, 17)
        Me.chkbxRightUseLegend.TabIndex = 1
        Me.chkbxRightUseLegend.Text = "Show Only Output Group from Legend"
        Me.chkbxRightUseLegend.UseVisualStyleBackColor = True
        '
        'cmdRun
        '
        Me.cmdRun.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdRun.CausesValidation = False
        Me.cmdRun.Location = New System.Drawing.Point(414, 332)
        Me.cmdRun.Name = "cmdRun"
        Me.cmdRun.Size = New System.Drawing.Size(75, 23)
        Me.cmdRun.TabIndex = 5
        Me.cmdRun.TabStop = False
        Me.cmdRun.Text = "Run"
        Me.cmdRun.UseVisualStyleBackColor = True
        '
        'cmdQuit
        '
        Me.cmdQuit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdQuit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdQuit.Location = New System.Drawing.Point(502, 332)
        Me.cmdQuit.Name = "cmdQuit"
        Me.cmdQuit.Size = New System.Drawing.Size(75, 23)
        Me.cmdQuit.TabIndex = 6
        Me.cmdQuit.TabStop = False
        Me.cmdQuit.Text = "Cancel"
        Me.cmdQuit.UseVisualStyleBackColor = True
        '
        'mnustrpMain
        '
        Me.mnustrpMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFile})
        Me.mnustrpMain.Location = New System.Drawing.Point(0, 0)
        Me.mnustrpMain.Name = "mnustrpMain"
        Me.mnustrpMain.Size = New System.Drawing.Size(584, 24)
        Me.mnustrpMain.TabIndex = 7
        Me.mnustrpMain.Text = "MenuStrip1"
        '
        'mnuFile
        '
        Me.mnuFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuitmAddToLegend, Me.mnuitmExit})
        Me.mnuFile.Name = "mnuFile"
        Me.mnuFile.Size = New System.Drawing.Size(37, 20)
        Me.mnuFile.Text = "File"
        '
        'mnuitmAddToLegend
        '
        Me.mnuitmAddToLegend.Name = "mnuitmAddToLegend"
        Me.mnuitmAddToLegend.Size = New System.Drawing.Size(247, 22)
        Me.mnuitmAddToLegend.Text = "Add To Legend From Project File"
        '
        'mnuitmExit
        '
        Me.mnuitmExit.Name = "mnuitmExit"
        Me.mnuitmExit.Size = New System.Drawing.Size(247, 22)
        Me.mnuitmExit.Text = "Exit"
        '
        'chkSelectedPolys
        '
        Me.chkSelectedPolys.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.chkSelectedPolys.Location = New System.Drawing.Point(12, 335)
        Me.chkSelectedPolys.Name = "chkSelectedPolys"
        Me.chkSelectedPolys.Size = New System.Drawing.Size(147, 22)
        Me.chkSelectedPolys.TabIndex = 63
        Me.chkSelectedPolys.Text = "Selected Polygons Only"
        Me.chkSelectedPolys.UseVisualStyleBackColor = True
        Me.chkSelectedPolys.Visible = False
        '
        'btnSelect
        '
        Me.btnSelect.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnSelect.CausesValidation = False
        Me.btnSelect.Location = New System.Drawing.Point(149, 332)
        Me.btnSelect.Name = "btnSelect"
        Me.btnSelect.Size = New System.Drawing.Size(75, 23)
        Me.btnSelect.TabIndex = 65
        Me.btnSelect.TabStop = False
        Me.btnSelect.Text = "Select"
        Me.btnSelect.UseVisualStyleBackColor = True
        Me.btnSelect.Visible = False
        '
        'lblSelected
        '
        Me.lblSelected.Location = New System.Drawing.Point(254, 338)
        Me.lblSelected.Name = "lblSelected"
        Me.lblSelected.Size = New System.Drawing.Size(100, 23)
        Me.lblSelected.TabIndex = 66
        Me.lblSelected.Text = "0 selected"
        Me.lblSelected.Visible = False
        '
        'lblOriginal
        '
        Me.lblOriginal.AutoSize = True
        Me.lblOriginal.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOriginal.Location = New System.Drawing.Point(3, 12)
        Me.lblOriginal.Name = "lblOriginal"
        Me.lblOriginal.Size = New System.Drawing.Size(148, 15)
        Me.lblOriginal.TabIndex = 67
        Me.lblOriginal.Text = "Select Original Output"
        '
        'lblModified
        '
        Me.lblModified.AutoSize = True
        Me.lblModified.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblModified.Location = New System.Drawing.Point(12, 12)
        Me.lblModified.Name = "lblModified"
        Me.lblModified.Size = New System.Drawing.Size(153, 15)
        Me.lblModified.TabIndex = 68
        Me.lblModified.Text = "Select Modified Output"
        '
        'lblCompareInstruction
        '
        Me.lblCompareInstruction.AutoSize = True
        Me.lblCompareInstruction.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompareInstruction.Location = New System.Drawing.Point(146, 6)
        Me.lblCompareInstruction.Name = "lblCompareInstruction"
        Me.lblCompareInstruction.Size = New System.Drawing.Size(280, 18)
        Me.lblCompareInstruction.TabIndex = 69
        Me.lblCompareInstruction.Text = "Compare Modified to Orignal Output"
        '
        'CompareOutputsForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.cmdQuit
        Me.ClientSize = New System.Drawing.Size(584, 362)
        Me.Controls.Add(Me.lblCompareInstruction)
        Me.Controls.Add(Me.lblSelected)
        Me.Controls.Add(Me.btnSelect)
        Me.Controls.Add(Me.chkSelectedPolys)
        Me.Controls.Add(Me.cmdRun)
        Me.Controls.Add(Me.cmdQuit)
        Me.Controls.Add(Me.pnlTop)
        Me.Controls.Add(Me.mnustrpMain)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.mnustrpMain
        Me.MinimumSize = New System.Drawing.Size(600, 400)
        Me.Name = "CompareOutputsForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Compare OpenNSPECT Outputs"
        Me.pnlTop.ResumeLayout(False)
        Me.spltconBase.Panel1.ResumeLayout(False)
        Me.spltconBase.Panel1.PerformLayout()
        Me.spltconBase.Panel2.ResumeLayout(False)
        Me.spltconBase.Panel2.PerformLayout()
        Me.spltconBase.ResumeLayout(False)
        Me.mnustrpMain.ResumeLayout(False)
        Me.mnustrpMain.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents pnlTop As System.Windows.Forms.Panel
    Private WithEvents cmdRun As System.Windows.Forms.Button
    Private WithEvents cmdQuit As System.Windows.Forms.Button
    Friend WithEvents spltconBase As System.Windows.Forms.SplitContainer
    Friend WithEvents chkbxLeftUseLegend As System.Windows.Forms.CheckBox
    Friend WithEvents lstbxLeft As System.Windows.Forms.ListBox
    Friend WithEvents chkbxRightUseLegend As System.Windows.Forms.CheckBox
    Friend WithEvents lstbxRight As System.Windows.Forms.ListBox
    Friend WithEvents mnustrpMain As System.Windows.Forms.MenuStrip
    Friend WithEvents mnuFile As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuitmAddToLegend As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuitmExit As System.Windows.Forms.ToolStripMenuItem
    Private WithEvents chkSelectedPolys As System.Windows.Forms.CheckBox
    Private WithEvents btnSelect As System.Windows.Forms.Button
    Friend WithEvents lblSelected As System.Windows.Forms.Label
    Friend WithEvents lblOriginal As System.Windows.Forms.Label
    Friend WithEvents lblModified As System.Windows.Forms.Label
    Friend WithEvents lblCompareInstruction As System.Windows.Forms.Label
End Class
