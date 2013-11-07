<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DataPrepForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DataPrepForm))
        Me.txtAOI = New System.Windows.Forms.TextBox()
        Me.lbl = New System.Windows.Forms.Label()
        Me.lblDPInstructions = New System.Windows.Forms.Label()
        Me.btnAOI = New System.Windows.Forms.Button()
        Me.txtProjName = New System.Windows.Forms.TextBox()
        Me.lblAOI_Proj = New System.Windows.Forms.Label()
        Me.diaOpenPrep = New System.Windows.Forms.OpenFileDialog()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtProjParams = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.TableLayoutPanel5 = New System.Windows.Forms.TableLayoutPanel()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.txtPrecipParams = New System.Windows.Forms.TextBox()
        Me.TableLayoutPanel9 = New System.Windows.Forms.TableLayoutPanel()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.txtPrecipProj = New System.Windows.Forms.TextBox()
        Me.txtPrecipName = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.btnOpenPrecip = New System.Windows.Forms.Button()
        Me.TableLayoutPanel6 = New System.Windows.Forms.TableLayoutPanel()
        Me.txtSizeUnitsPrecip = New System.Windows.Forms.TextBox()
        Me.txtCellSizePrecip = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.TableLayoutPanel3 = New System.Windows.Forms.TableLayoutPanel()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.txtLCParams = New System.Windows.Forms.TextBox()
        Me.TableLayoutPanel7 = New System.Windows.Forms.TableLayoutPanel()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.txtLCProj = New System.Windows.Forms.TextBox()
        Me.txtLCName = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.btnOpenLC = New System.Windows.Forms.Button()
        Me.TableLayoutPanel4 = New System.Windows.Forms.TableLayoutPanel()
        Me.txtSizeUnitsLC = New System.Windows.Forms.TextBox()
        Me.txtCellSizeLC = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.TableLayoutPanel8 = New System.Windows.Forms.TableLayoutPanel()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtDEMProj = New System.Windows.Forms.TextBox()
        Me.txtDEMParams = New System.Windows.Forms.TextBox()
        Me.txtDEMName = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnOpenDEM = New System.Windows.Forms.Button()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.txtSizeUnitsDEM = New System.Windows.Forms.TextBox()
        Me.txtCellSizeDEM = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.HelpProvider1 = New System.Windows.Forms.HelpProvider()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.TableLayoutPanel11 = New System.Windows.Forms.TableLayoutPanel()
        Me.txtUserBuffer = New System.Windows.Forms.TextBox()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.TableLayoutPanel10 = New System.Windows.Forms.TableLayoutPanel()
        Me.txtFinalCell = New System.Windows.Forms.TextBox()
        Me.lblTarCell = New System.Windows.Forms.Label()
        Me.txtFinalCellUnits = New System.Windows.Forms.TextBox()
        Me.cbKeep = New System.Windows.Forms.CheckBox()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnRun = New System.Windows.Forms.Button()
        Me.cbLoadFinal = New System.Windows.Forms.CheckBox()
        Me.GroupBox1.SuspendLayout()
        Me.TableLayoutPanel5.SuspendLayout()
        Me.TableLayoutPanel9.SuspendLayout()
        Me.TableLayoutPanel6.SuspendLayout()
        Me.TableLayoutPanel3.SuspendLayout()
        Me.TableLayoutPanel7.SuspendLayout()
        Me.TableLayoutPanel4.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.TableLayoutPanel8.SuspendLayout()
        Me.TableLayoutPanel2.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.TableLayoutPanel11.SuspendLayout()
        Me.TableLayoutPanel10.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtAOI
        '
        Me.txtAOI.Location = New System.Drawing.Point(111, 19)
        Me.txtAOI.Name = "txtAOI"
        Me.txtAOI.Size = New System.Drawing.Size(510, 20)
        Me.txtAOI.TabIndex = 0
        '
        'lbl
        '
        Me.lbl.AutoSize = True
        Me.lbl.Location = New System.Drawing.Point(33, 22)
        Me.lbl.Name = "lbl"
        Me.lbl.Size = New System.Drawing.Size(72, 13)
        Me.lbl.TabIndex = 1
        Me.lbl.Text = "AOI Shapefile"
        '
        'lblDPInstructions
        '
        Me.lblDPInstructions.AutoSize = True
        Me.lblDPInstructions.Location = New System.Drawing.Point(26, 9)
        Me.lblDPInstructions.Name = "lblDPInstructions"
        Me.lblDPInstructions.Size = New System.Drawing.Size(677, 78)
        Me.lblDPInstructions.TabIndex = 2
        Me.lblDPInstructions.Text = resources.GetString("lblDPInstructions.Text")
        '
        'btnAOI
        '
        Me.btnAOI.Location = New System.Drawing.Point(627, 16)
        Me.btnAOI.Name = "btnAOI"
        Me.btnAOI.Size = New System.Drawing.Size(75, 23)
        Me.btnAOI.TabIndex = 3
        Me.btnAOI.Text = "Browse"
        Me.btnAOI.UseVisualStyleBackColor = True
        '
        'txtProjName
        '
        Me.txtProjName.ForeColor = System.Drawing.SystemColors.GrayText
        Me.txtProjName.Location = New System.Drawing.Point(111, 52)
        Me.txtProjName.Name = "txtProjName"
        Me.txtProjName.ReadOnly = True
        Me.txtProjName.Size = New System.Drawing.Size(591, 20)
        Me.txtProjName.TabIndex = 4
        '
        'lblAOI_Proj
        '
        Me.lblAOI_Proj.AutoSize = True
        Me.lblAOI_Proj.Location = New System.Drawing.Point(20, 55)
        Me.lblAOI_Proj.Name = "lblAOI_Proj"
        Me.lblAOI_Proj.Size = New System.Drawing.Size(85, 13)
        Me.lblAOI_Proj.TabIndex = 5
        Me.lblAOI_Proj.Text = "Projection Name"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(1, 83)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(104, 13)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Projection Prameters"
        '
        'txtProjParams
        '
        Me.txtProjParams.ForeColor = System.Drawing.SystemColors.WindowFrame
        Me.txtProjParams.Location = New System.Drawing.Point(111, 80)
        Me.txtProjParams.Name = "txtProjParams"
        Me.txtProjParams.ReadOnly = True
        Me.txtProjParams.Size = New System.Drawing.Size(591, 20)
        Me.txtProjParams.TabIndex = 6
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.TableLayoutPanel5)
        Me.GroupBox1.Controls.Add(Me.TableLayoutPanel3)
        Me.GroupBox1.Controls.Add(Me.TableLayoutPanel1)
        Me.GroupBox1.Location = New System.Drawing.Point(16, 276)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(708, 389)
        Me.GroupBox1.TabIndex = 10
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Initial Data"
        '
        'TableLayoutPanel5
        '
        Me.TableLayoutPanel5.ColumnCount = 4
        Me.TableLayoutPanel5.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20.36424!))
        Me.TableLayoutPanel5.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 79.63576!))
        Me.TableLayoutPanel5.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 149.0!))
        Me.TableLayoutPanel5.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 94.0!))
        Me.TableLayoutPanel5.Controls.Add(Me.Label13, 0, 2)
        Me.TableLayoutPanel5.Controls.Add(Me.txtPrecipParams, 1, 2)
        Me.TableLayoutPanel5.Controls.Add(Me.TableLayoutPanel9, 0, 1)
        Me.TableLayoutPanel5.Controls.Add(Me.txtPrecipName, 1, 0)
        Me.TableLayoutPanel5.Controls.Add(Me.Label7, 0, 0)
        Me.TableLayoutPanel5.Controls.Add(Me.btnOpenPrecip, 3, 0)
        Me.TableLayoutPanel5.Controls.Add(Me.TableLayoutPanel6, 2, 1)
        Me.TableLayoutPanel5.Location = New System.Drawing.Point(4, 267)
        Me.TableLayoutPanel5.Name = "TableLayoutPanel5"
        Me.TableLayoutPanel5.RowCount = 3
        Me.TableLayoutPanel5.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel5.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel5.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel5.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel5.Size = New System.Drawing.Size(698, 109)
        Me.TableLayoutPanel5.TabIndex = 11
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(3, 69)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(57, 26)
        Me.Label13.TabIndex = 17
        Me.Label13.Text = "Projection Prameters"
        '
        'txtPrecipParams
        '
        Me.TableLayoutPanel5.SetColumnSpan(Me.txtPrecipParams, 3)
        Me.txtPrecipParams.ForeColor = System.Drawing.SystemColors.WindowFrame
        Me.txtPrecipParams.Location = New System.Drawing.Point(95, 72)
        Me.txtPrecipParams.Name = "txtPrecipParams"
        Me.txtPrecipParams.ReadOnly = True
        Me.txtPrecipParams.Size = New System.Drawing.Size(591, 20)
        Me.txtPrecipParams.TabIndex = 16
        '
        'TableLayoutPanel9
        '
        Me.TableLayoutPanel9.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel9.ColumnCount = 2
        Me.TableLayoutPanel5.SetColumnSpan(Me.TableLayoutPanel9, 2)
        Me.TableLayoutPanel9.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 26.78186!))
        Me.TableLayoutPanel9.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 73.21814!))
        Me.TableLayoutPanel9.Controls.Add(Me.Label10, 0, 0)
        Me.TableLayoutPanel9.Controls.Add(Me.txtPrecipProj, 1, 0)
        Me.TableLayoutPanel9.Location = New System.Drawing.Point(3, 32)
        Me.TableLayoutPanel9.Name = "TableLayoutPanel9"
        Me.TableLayoutPanel9.RowCount = 1
        Me.TableLayoutPanel9.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel9.Size = New System.Drawing.Size(448, 34)
        Me.TableLayoutPanel9.TabIndex = 14
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(3, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(85, 13)
        Me.Label10.TabIndex = 12
        Me.Label10.Text = "Projection Name"
        '
        'txtPrecipProj
        '
        Me.txtPrecipProj.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtPrecipProj.ForeColor = System.Drawing.SystemColors.GrayText
        Me.txtPrecipProj.Location = New System.Drawing.Point(122, 3)
        Me.txtPrecipProj.Name = "txtPrecipProj"
        Me.txtPrecipProj.ReadOnly = True
        Me.txtPrecipProj.Size = New System.Drawing.Size(323, 20)
        Me.txtPrecipProj.TabIndex = 11
        '
        'txtPrecipName
        '
        Me.TableLayoutPanel5.SetColumnSpan(Me.txtPrecipName, 2)
        Me.txtPrecipName.Location = New System.Drawing.Point(95, 3)
        Me.txtPrecipName.Name = "txtPrecipName"
        Me.txtPrecipName.Size = New System.Drawing.Size(484, 20)
        Me.txtPrecipName.TabIndex = 9
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(3, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(68, 26)
        Me.Label7.TabIndex = 2
        Me.Label7.Text = "Precipitation Raster"
        '
        'btnOpenPrecip
        '
        Me.btnOpenPrecip.Location = New System.Drawing.Point(606, 3)
        Me.btnOpenPrecip.Name = "btnOpenPrecip"
        Me.btnOpenPrecip.Size = New System.Drawing.Size(73, 23)
        Me.btnOpenPrecip.TabIndex = 12
        Me.btnOpenPrecip.Text = "Browse"
        Me.btnOpenPrecip.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel6
        '
        Me.TableLayoutPanel6.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel6.ColumnCount = 3
        Me.TableLayoutPanel5.SetColumnSpan(Me.TableLayoutPanel6, 2)
        Me.TableLayoutPanel6.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel6.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel6.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel6.Controls.Add(Me.txtSizeUnitsPrecip, 0, 0)
        Me.TableLayoutPanel6.Controls.Add(Me.txtCellSizePrecip, 0, 0)
        Me.TableLayoutPanel6.Controls.Add(Me.Label6, 0, 0)
        Me.TableLayoutPanel6.Location = New System.Drawing.Point(473, 32)
        Me.TableLayoutPanel6.Name = "TableLayoutPanel6"
        Me.TableLayoutPanel6.RowCount = 1
        Me.TableLayoutPanel6.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel6.Size = New System.Drawing.Size(222, 28)
        Me.TableLayoutPanel6.TabIndex = 10
        '
        'txtSizeUnitsPrecip
        '
        Me.txtSizeUnitsPrecip.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtSizeUnitsPrecip.Location = New System.Drawing.Point(150, 3)
        Me.txtSizeUnitsPrecip.Name = "txtSizeUnitsPrecip"
        Me.txtSizeUnitsPrecip.ReadOnly = True
        Me.txtSizeUnitsPrecip.Size = New System.Drawing.Size(62, 20)
        Me.txtSizeUnitsPrecip.TabIndex = 15
        '
        'txtCellSizePrecip
        '
        Me.txtCellSizePrecip.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCellSizePrecip.Location = New System.Drawing.Point(56, 3)
        Me.txtCellSizePrecip.Name = "txtCellSizePrecip"
        Me.txtCellSizePrecip.ReadOnly = True
        Me.txtCellSizePrecip.Size = New System.Drawing.Size(88, 20)
        Me.txtCellSizePrecip.TabIndex = 14
        '
        'Label6
        '
        Me.Label6.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(3, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(47, 28)
        Me.Label6.TabIndex = 10
        Me.Label6.Text = "Cell Size"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TableLayoutPanel3
        '
        Me.TableLayoutPanel3.ColumnCount = 4
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20.36424!))
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 79.63576!))
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 147.0!))
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 96.0!))
        Me.TableLayoutPanel3.Controls.Add(Me.Label12, 0, 2)
        Me.TableLayoutPanel3.Controls.Add(Me.txtLCParams, 1, 2)
        Me.TableLayoutPanel3.Controls.Add(Me.TableLayoutPanel7, 0, 1)
        Me.TableLayoutPanel3.Controls.Add(Me.txtLCName, 1, 0)
        Me.TableLayoutPanel3.Controls.Add(Me.Label5, 0, 0)
        Me.TableLayoutPanel3.Controls.Add(Me.btnOpenLC, 3, 0)
        Me.TableLayoutPanel3.Controls.Add(Me.TableLayoutPanel4, 2, 1)
        Me.TableLayoutPanel3.Location = New System.Drawing.Point(4, 143)
        Me.TableLayoutPanel3.Name = "TableLayoutPanel3"
        Me.TableLayoutPanel3.RowCount = 3
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel3.Size = New System.Drawing.Size(698, 104)
        Me.TableLayoutPanel3.TabIndex = 10
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(3, 69)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(57, 26)
        Me.Label12.TabIndex = 15
        Me.Label12.Text = "Projection Prameters"
        '
        'txtLCParams
        '
        Me.TableLayoutPanel3.SetColumnSpan(Me.txtLCParams, 3)
        Me.txtLCParams.ForeColor = System.Drawing.SystemColors.WindowFrame
        Me.txtLCParams.Location = New System.Drawing.Point(95, 72)
        Me.txtLCParams.Name = "txtLCParams"
        Me.txtLCParams.ReadOnly = True
        Me.txtLCParams.Size = New System.Drawing.Size(591, 20)
        Me.txtLCParams.TabIndex = 14
        '
        'TableLayoutPanel7
        '
        Me.TableLayoutPanel7.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel7.ColumnCount = 2
        Me.TableLayoutPanel3.SetColumnSpan(Me.TableLayoutPanel7, 2)
        Me.TableLayoutPanel7.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 26.78186!))
        Me.TableLayoutPanel7.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 73.21814!))
        Me.TableLayoutPanel7.Controls.Add(Me.Label9, 0, 0)
        Me.TableLayoutPanel7.Controls.Add(Me.txtLCProj, 1, 0)
        Me.TableLayoutPanel7.Location = New System.Drawing.Point(3, 32)
        Me.TableLayoutPanel7.Name = "TableLayoutPanel7"
        Me.TableLayoutPanel7.RowCount = 2
        Me.TableLayoutPanel7.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel7.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel7.Size = New System.Drawing.Size(448, 34)
        Me.TableLayoutPanel7.TabIndex = 14
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(3, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(85, 13)
        Me.Label9.TabIndex = 12
        Me.Label9.Text = "Projection Name"
        '
        'txtLCProj
        '
        Me.txtLCProj.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtLCProj.ForeColor = System.Drawing.SystemColors.GrayText
        Me.txtLCProj.Location = New System.Drawing.Point(122, 3)
        Me.txtLCProj.Name = "txtLCProj"
        Me.txtLCProj.ReadOnly = True
        Me.txtLCProj.Size = New System.Drawing.Size(323, 20)
        Me.txtLCProj.TabIndex = 11
        '
        'txtLCName
        '
        Me.TableLayoutPanel3.SetColumnSpan(Me.txtLCName, 2)
        Me.txtLCName.Location = New System.Drawing.Point(95, 3)
        Me.txtLCName.Name = "txtLCName"
        Me.txtLCName.Size = New System.Drawing.Size(484, 20)
        Me.txtLCName.TabIndex = 9
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(3, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(61, 26)
        Me.Label5.TabIndex = 2
        Me.Label5.Text = "Landcover Raster"
        '
        'btnOpenLC
        '
        Me.btnOpenLC.Location = New System.Drawing.Point(604, 3)
        Me.btnOpenLC.Name = "btnOpenLC"
        Me.btnOpenLC.Size = New System.Drawing.Size(75, 23)
        Me.btnOpenLC.TabIndex = 12
        Me.btnOpenLC.Text = "Browse"
        Me.btnOpenLC.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel4
        '
        Me.TableLayoutPanel4.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel4.ColumnCount = 3
        Me.TableLayoutPanel3.SetColumnSpan(Me.TableLayoutPanel4, 2)
        Me.TableLayoutPanel4.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel4.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel4.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel4.Controls.Add(Me.txtSizeUnitsLC, 0, 0)
        Me.TableLayoutPanel4.Controls.Add(Me.txtCellSizeLC, 0, 0)
        Me.TableLayoutPanel4.Controls.Add(Me.Label4, 0, 0)
        Me.TableLayoutPanel4.Location = New System.Drawing.Point(473, 32)
        Me.TableLayoutPanel4.Name = "TableLayoutPanel4"
        Me.TableLayoutPanel4.RowCount = 1
        Me.TableLayoutPanel4.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel4.Size = New System.Drawing.Size(222, 28)
        Me.TableLayoutPanel4.TabIndex = 10
        '
        'txtSizeUnitsLC
        '
        Me.txtSizeUnitsLC.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtSizeUnitsLC.Location = New System.Drawing.Point(150, 3)
        Me.txtSizeUnitsLC.Name = "txtSizeUnitsLC"
        Me.txtSizeUnitsLC.ReadOnly = True
        Me.txtSizeUnitsLC.Size = New System.Drawing.Size(62, 20)
        Me.txtSizeUnitsLC.TabIndex = 15
        '
        'txtCellSizeLC
        '
        Me.txtCellSizeLC.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCellSizeLC.Location = New System.Drawing.Point(56, 3)
        Me.txtCellSizeLC.Name = "txtCellSizeLC"
        Me.txtCellSizeLC.ReadOnly = True
        Me.txtCellSizeLC.Size = New System.Drawing.Size(88, 20)
        Me.txtCellSizeLC.TabIndex = 14
        '
        'Label4
        '
        Me.Label4.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(3, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(47, 28)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Cell Size"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 4
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20.36424!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 79.63576!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 147.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 96.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.Label11, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.TableLayoutPanel8, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.txtDEMParams, 1, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.txtDEMName, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Label2, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.btnOpenDEM, 3, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.TableLayoutPanel2, 2, 1)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(4, 23)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 3
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(698, 104)
        Me.TableLayoutPanel1.TabIndex = 9
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(3, 69)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(57, 26)
        Me.Label11.TabIndex = 13
        Me.Label11.Text = "Projection Prameters"
        '
        'TableLayoutPanel8
        '
        Me.TableLayoutPanel8.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel8.ColumnCount = 2
        Me.TableLayoutPanel1.SetColumnSpan(Me.TableLayoutPanel8, 2)
        Me.TableLayoutPanel8.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 26.78186!))
        Me.TableLayoutPanel8.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 73.21814!))
        Me.TableLayoutPanel8.Controls.Add(Me.Label8, 0, 0)
        Me.TableLayoutPanel8.Controls.Add(Me.txtDEMProj, 1, 0)
        Me.TableLayoutPanel8.Location = New System.Drawing.Point(3, 32)
        Me.TableLayoutPanel8.Name = "TableLayoutPanel8"
        Me.TableLayoutPanel8.RowCount = 2
        Me.TableLayoutPanel8.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel8.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel8.Size = New System.Drawing.Size(448, 34)
        Me.TableLayoutPanel8.TabIndex = 13
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(3, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(85, 13)
        Me.Label8.TabIndex = 12
        Me.Label8.Text = "Projection Name"
        '
        'txtDEMProj
        '
        Me.txtDEMProj.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtDEMProj.ForeColor = System.Drawing.SystemColors.GrayText
        Me.txtDEMProj.Location = New System.Drawing.Point(122, 3)
        Me.txtDEMProj.Name = "txtDEMProj"
        Me.txtDEMProj.ReadOnly = True
        Me.txtDEMProj.Size = New System.Drawing.Size(323, 20)
        Me.txtDEMProj.TabIndex = 11
        '
        'txtDEMParams
        '
        Me.TableLayoutPanel1.SetColumnSpan(Me.txtDEMParams, 3)
        Me.txtDEMParams.ForeColor = System.Drawing.SystemColors.WindowFrame
        Me.txtDEMParams.Location = New System.Drawing.Point(95, 72)
        Me.txtDEMParams.Name = "txtDEMParams"
        Me.txtDEMParams.ReadOnly = True
        Me.txtDEMParams.Size = New System.Drawing.Size(591, 20)
        Me.txtDEMParams.TabIndex = 12
        '
        'txtDEMName
        '
        Me.TableLayoutPanel1.SetColumnSpan(Me.txtDEMName, 2)
        Me.txtDEMName.Location = New System.Drawing.Point(95, 3)
        Me.txtDEMName.Name = "txtDEMName"
        Me.txtDEMName.Size = New System.Drawing.Size(484, 20)
        Me.txtDEMName.TabIndex = 9
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(3, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(85, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Elevation Raster"
        '
        'btnOpenDEM
        '
        Me.btnOpenDEM.Location = New System.Drawing.Point(604, 3)
        Me.btnOpenDEM.Name = "btnOpenDEM"
        Me.btnOpenDEM.Size = New System.Drawing.Size(75, 23)
        Me.btnOpenDEM.TabIndex = 12
        Me.btnOpenDEM.Text = "Browse"
        Me.btnOpenDEM.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel2.ColumnCount = 3
        Me.TableLayoutPanel1.SetColumnSpan(Me.TableLayoutPanel2, 2)
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel2.Controls.Add(Me.txtSizeUnitsDEM, 0, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.txtCellSizeDEM, 0, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.Label3, 0, 0)
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(473, 32)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 1
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28.0!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(222, 28)
        Me.TableLayoutPanel2.TabIndex = 10
        '
        'txtSizeUnitsDEM
        '
        Me.txtSizeUnitsDEM.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtSizeUnitsDEM.Location = New System.Drawing.Point(150, 3)
        Me.txtSizeUnitsDEM.Name = "txtSizeUnitsDEM"
        Me.txtSizeUnitsDEM.ReadOnly = True
        Me.txtSizeUnitsDEM.Size = New System.Drawing.Size(62, 20)
        Me.txtSizeUnitsDEM.TabIndex = 15
        '
        'txtCellSizeDEM
        '
        Me.txtCellSizeDEM.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCellSizeDEM.Location = New System.Drawing.Point(56, 3)
        Me.txtCellSizeDEM.Name = "txtCellSizeDEM"
        Me.txtCellSizeDEM.ReadOnly = True
        Me.txtCellSizeDEM.Size = New System.Drawing.Size(88, 20)
        Me.txtCellSizeDEM.TabIndex = 14
        '
        'Label3
        '
        Me.Label3.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(3, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(47, 28)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "Cell Size"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label14
        '
        Me.Label14.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label14.AutoSize = True
        Me.HelpProvider1.SetHelpString(Me.Label14, "Increase if the final grids do not align precisely.")
        Me.Label14.Location = New System.Drawing.Point(3, 0)
        Me.Label14.Name = "Label14"
        Me.HelpProvider1.SetShowHelp(Me.Label14, True)
        Me.Label14.Size = New System.Drawing.Size(126, 28)
        Me.Label14.TabIndex = 10
        Me.Label14.Text = "Intermediate Buffer Size:*"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Label15)
        Me.GroupBox2.Controls.Add(Me.TableLayoutPanel11)
        Me.GroupBox2.Controls.Add(Me.TableLayoutPanel10)
        Me.GroupBox2.Controls.Add(Me.txtAOI)
        Me.GroupBox2.Controls.Add(Me.lbl)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.btnAOI)
        Me.GroupBox2.Controls.Add(Me.txtProjParams)
        Me.GroupBox2.Controls.Add(Me.txtProjName)
        Me.GroupBox2.Controls.Add(Me.lblAOI_Proj)
        Me.GroupBox2.Location = New System.Drawing.Point(16, 96)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(708, 170)
        Me.GroupBox2.TabIndex = 11
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Final Area of Interest and Projection"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(2, 144)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(684, 13)
        Me.Label15.TabIndex = 13
        Me.Label15.Text = "* Depends on the projections of the raw data. Generally a value of 20-50 grid cel" & _
    "ls works well. Only increase if the final grids do not align exactly."
        '
        'TableLayoutPanel11
        '
        Me.TableLayoutPanel11.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel11.ColumnCount = 3
        Me.TableLayoutPanel11.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel11.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel11.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel11.Controls.Add(Me.txtUserBuffer, 0, 0)
        Me.TableLayoutPanel11.Controls.Add(Me.Label14, 0, 0)
        Me.TableLayoutPanel11.Controls.Add(Me.TextBox2, 2, 0)
        Me.TableLayoutPanel11.Location = New System.Drawing.Point(420, 106)
        Me.TableLayoutPanel11.Name = "TableLayoutPanel11"
        Me.TableLayoutPanel11.RowCount = 1
        Me.TableLayoutPanel11.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel11.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28.0!))
        Me.TableLayoutPanel11.Size = New System.Drawing.Size(271, 28)
        Me.TableLayoutPanel11.TabIndex = 12
        '
        'txtUserBuffer
        '
        Me.txtUserBuffer.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtUserBuffer.Location = New System.Drawing.Point(135, 3)
        Me.txtUserBuffer.Name = "txtUserBuffer"
        Me.txtUserBuffer.Size = New System.Drawing.Size(88, 20)
        Me.txtUserBuffer.TabIndex = 14
        Me.txtUserBuffer.Text = "50"
        '
        'TextBox2
        '
        Me.TextBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.TextBox2.Location = New System.Drawing.Point(229, 3)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.ReadOnly = True
        Me.TextBox2.Size = New System.Drawing.Size(40, 20)
        Me.TextBox2.TabIndex = 15
        Me.TextBox2.Text = "Cells"
        '
        'TableLayoutPanel10
        '
        Me.TableLayoutPanel10.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel10.ColumnCount = 3
        Me.TableLayoutPanel10.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel10.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel10.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel10.Controls.Add(Me.txtFinalCell, 0, 0)
        Me.TableLayoutPanel10.Controls.Add(Me.lblTarCell, 0, 0)
        Me.TableLayoutPanel10.Controls.Add(Me.txtFinalCellUnits, 2, 0)
        Me.TableLayoutPanel10.Location = New System.Drawing.Point(23, 106)
        Me.TableLayoutPanel10.Name = "TableLayoutPanel10"
        Me.TableLayoutPanel10.RowCount = 1
        Me.TableLayoutPanel10.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel10.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28.0!))
        Me.TableLayoutPanel10.Size = New System.Drawing.Size(282, 28)
        Me.TableLayoutPanel10.TabIndex = 11
        '
        'txtFinalCell
        '
        Me.txtFinalCell.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtFinalCell.Location = New System.Drawing.Point(81, 3)
        Me.txtFinalCell.Name = "txtFinalCell"
        Me.txtFinalCell.Size = New System.Drawing.Size(88, 20)
        Me.txtFinalCell.TabIndex = 14
        '
        'lblTarCell
        '
        Me.lblTarCell.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblTarCell.AutoSize = True
        Me.lblTarCell.Location = New System.Drawing.Point(3, 0)
        Me.lblTarCell.Name = "lblTarCell"
        Me.lblTarCell.Size = New System.Drawing.Size(72, 28)
        Me.lblTarCell.TabIndex = 10
        Me.lblTarCell.Text = "Final Cell Size"
        Me.lblTarCell.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtFinalCellUnits
        '
        Me.txtFinalCellUnits.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtFinalCellUnits.Location = New System.Drawing.Point(175, 3)
        Me.txtFinalCellUnits.Name = "txtFinalCellUnits"
        Me.txtFinalCellUnits.ReadOnly = True
        Me.txtFinalCellUnits.Size = New System.Drawing.Size(62, 20)
        Me.txtFinalCellUnits.TabIndex = 15
        '
        'cbKeep
        '
        Me.cbKeep.AutoSize = True
        Me.cbKeep.Location = New System.Drawing.Point(142, 679)
        Me.cbKeep.Name = "cbKeep"
        Me.cbKeep.Size = New System.Drawing.Size(141, 17)
        Me.cbKeep.TabIndex = 12
        Me.cbKeep.Text = "Keep intermediate data?"
        Me.cbKeep.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(549, 675)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 13
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnRun
        '
        Me.btnRun.Location = New System.Drawing.Point(639, 675)
        Me.btnRun.Name = "btnRun"
        Me.btnRun.Size = New System.Drawing.Size(75, 23)
        Me.btnRun.TabIndex = 14
        Me.btnRun.Text = "Run"
        Me.btnRun.UseVisualStyleBackColor = True
        '
        'cbLoadFinal
        '
        Me.cbLoadFinal.AutoSize = True
        Me.cbLoadFinal.Checked = True
        Me.cbLoadFinal.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbLoadFinal.Location = New System.Drawing.Point(350, 679)
        Me.cbLoadFinal.Name = "cbLoadFinal"
        Me.cbLoadFinal.Size = New System.Drawing.Size(164, 17)
        Me.cbLoadFinal.TabIndex = 15
        Me.cbLoadFinal.Text = "Display AOI and final rasters?"
        Me.cbLoadFinal.UseVisualStyleBackColor = True
        '
        'DataPrepForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(742, 714)
        Me.Controls.Add(Me.cbLoadFinal)
        Me.Controls.Add(Me.btnRun)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.cbKeep)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.lblDPInstructions)
        Me.HelpButton = True
        Me.Name = "DataPrepForm"
        Me.Text = "Preprocess New Data Sets"
        Me.GroupBox1.ResumeLayout(False)
        Me.TableLayoutPanel5.ResumeLayout(False)
        Me.TableLayoutPanel5.PerformLayout()
        Me.TableLayoutPanel9.ResumeLayout(False)
        Me.TableLayoutPanel9.PerformLayout()
        Me.TableLayoutPanel6.ResumeLayout(False)
        Me.TableLayoutPanel6.PerformLayout()
        Me.TableLayoutPanel3.ResumeLayout(False)
        Me.TableLayoutPanel3.PerformLayout()
        Me.TableLayoutPanel7.ResumeLayout(False)
        Me.TableLayoutPanel7.PerformLayout()
        Me.TableLayoutPanel4.ResumeLayout(False)
        Me.TableLayoutPanel4.PerformLayout()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.TableLayoutPanel8.ResumeLayout(False)
        Me.TableLayoutPanel8.PerformLayout()
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.TableLayoutPanel2.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.TableLayoutPanel11.ResumeLayout(False)
        Me.TableLayoutPanel11.PerformLayout()
        Me.TableLayoutPanel10.ResumeLayout(False)
        Me.TableLayoutPanel10.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtAOI As System.Windows.Forms.TextBox
    Friend WithEvents lbl As System.Windows.Forms.Label
    Friend WithEvents lblDPInstructions As System.Windows.Forms.Label
    Friend WithEvents btnAOI As System.Windows.Forms.Button
    Friend WithEvents txtProjName As System.Windows.Forms.TextBox
    Friend WithEvents lblAOI_Proj As System.Windows.Forms.Label
    Friend WithEvents diaOpenPrep As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtProjParams As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents TableLayoutPanel5 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents TableLayoutPanel6 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents txtSizeUnitsPrecip As System.Windows.Forms.TextBox
    Friend WithEvents txtCellSizePrecip As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnOpenPrecip As System.Windows.Forms.Button
    Friend WithEvents txtPrecipName As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TableLayoutPanel3 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents TableLayoutPanel4 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents txtSizeUnitsLC As System.Windows.Forms.TextBox
    Friend WithEvents txtCellSizeLC As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents btnOpenLC As System.Windows.Forms.Button
    Friend WithEvents txtLCName As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents btnOpenDEM As System.Windows.Forms.Button
    Friend WithEvents txtDEMName As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TableLayoutPanel8 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents txtDEMProj As System.Windows.Forms.TextBox
    Friend WithEvents TableLayoutPanel9 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtPrecipProj As System.Windows.Forms.TextBox
    Friend WithEvents TableLayoutPanel7 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtLCProj As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TableLayoutPanel2 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents txtSizeUnitsDEM As System.Windows.Forms.TextBox
    Friend WithEvents txtCellSizeDEM As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents HelpProvider1 As System.Windows.Forms.HelpProvider
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents TableLayoutPanel10 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents txtFinalCell As System.Windows.Forms.TextBox
    Friend WithEvents lblTarCell As System.Windows.Forms.Label
    Friend WithEvents txtFinalCellUnits As System.Windows.Forms.TextBox
    Friend WithEvents cbKeep As System.Windows.Forms.CheckBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnRun As System.Windows.Forms.Button
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents txtPrecipParams As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents txtLCParams As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtDEMParams As System.Windows.Forms.TextBox
    Friend WithEvents cbLoadFinal As System.Windows.Forms.CheckBox
    Friend WithEvents TableLayoutPanel11 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents txtUserBuffer As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
End Class
