Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing
Imports System.Windows.Forms

Friend Class frmProjectSetup
    Inherits System.Windows.Forms.Form

#Region "Class Vars"
    Private _strFileName As String 'Name of Open doc
    Private _strWShed As String 'String

    Private _XMLPrjParams As clsXMLPrjFile 'xml doc that holds inputs

    Private _bolFirstLoad As Boolean 'Is initial Load event
    Private _booNew As Boolean 'New
    Private _booExists As Boolean 'Has file been saved
    Private _booAnnualPrecip As Boolean 'Is the precip scenario annual, if so = TRUE

    Private _strPrecipFile As String
    Private _strOpenFileName As String

    Private _intMgmtCount As Short 'Count for management scenarios
    Private _intLUCount As Short 'Count for Land Use grid
    Private _intPollCount As Short 'count for pollutant grid

    'Font DPI API
    Private Declare Function GetDC Lib "user32" (ByVal hwnd As Integer) As Integer
    Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Integer, ByVal hDC As Integer) As Integer
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Integer, ByVal nIndex As Integer) As Integer

    ' Win 32 Constant Declarations
    Private Const LOGPIXELSX As Short = 88 'Logical pixels/inch in X


    Const c_sModuleFileName As String = "frmProjectSetup.vb"
    Private m_ParentHWND As Integer 'Set this to get correct parenting of Error handler forms

    Private arrAreaList As New System.Collections.ArrayList
    Private arrClassList As New System.Collections.ArrayList
#End Region

#Region "Events"

    Private Sub frmProjectSetup_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            SSTab1.SelectedIndex = 0
            'Check the DPI setting: This was put in place because the alignment of the checkboxes will be
            'messed up if DPI settings are anything other than Normal
            Dim lngMapDC As Integer
            Dim lngDPI As Integer
            lngMapDC = GetDC(Me.Handle.ToInt32)
            lngDPI = GetDeviceCaps(lngMapDC, LOGPIXELSX)
            ReleaseDC(Me.Handle.ToInt32, lngMapDC)

            If lngDPI <> 96 Then
                MsgBox("Warning: N-SPECT requires your font size to be 96 DPI." & vbNewLine & "Some controls may appear out of alignment on this form.", MsgBoxStyle.Critical, "Warning!")
            End If

            _bolFirstLoad = True 'It's the first load
            _booExists = False

            'ComboBox::LandCover Type
            modUtil.InitComboBox(cboLCType, "LCType")

            'ComboBox::Precipitation Scenarios
            modUtil.InitComboBox(cboPrecipScen, "PrecipScenario")
            cboPrecipScen.Items.Insert(cboPrecipScen.Items.Count, "New precipitation scenario...")

            'ComboBox::WaterShed Delineations
            modUtil.InitComboBox(cboWSDelin, "WSDelineation")
            cboWSDelin.Items.Insert(cboWSDelin.Items.Count, "New watershed delineation...")

            'ComboBox::WaterQuality Criteria
            modUtil.InitComboBox(cboWQStd, "WQCriteria")
            cboWQStd.Items.Insert(cboWQStd.Items.Count, "New water quality standard...")

            cboLCLayer.Items.Clear()
            cboRainGrid.Items.Clear()
            cboSelectPoly.Items.Clear()
            arrAreaList.Clear()
            Dim currLyr As MapWindow.Interfaces.Layer
            For i As Integer = 0 To g_MapWin.Layers.NumLayers - 1
                currLyr = g_MapWin.Layers(i)
                If currLyr.LayerType = MapWindow.Interfaces.eLayerType.Grid Then
                    cboLCLayer.Items.Add(currLyr.Name)
                    cboRainGrid.Items.Add(currLyr.Name)
                ElseIf currLyr.LayerType = MapWindow.Interfaces.eLayerType.PolygonShapefile Then
                    cboSelectPoly.Items.Add(currLyr.Name)
                    arrAreaList.Add(currLyr.Name)
                End If
            Next


            'Soils, now a 'scenario', not just a datalayer
            modUtil.InitComboBox(cboSoilsLayer, "Soils")

            'Fill LandClass
            FillCboLCCLass()

            _intMgmtCount = 0
            _intLUCount = 0

            'Initialize parameter file
            _XMLPrjParams = New clsXMLPrjFile

            Me.Text = "Untitled"

            chkSelectedPolys.Enabled = EnableChkWaterShed()

            'Add one blank management row
            dgvManagementScen.Rows.Clear()
            dgvManagementScen.Rows.Add()
            PopulateManagement(0)

            'Test workspace persistence
            If Len(g_strWorkspace) > 0 Then
                txtOutputWS.Text = g_strWorkspace
            End If

        Catch ex As Exception
            HandleError(True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Sub


    Private Sub txtProjectName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtProjectName.TextChanged
        Me.Text = txtProjectName.Text
    End Sub

    Private Sub txtOutputWS_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOutputWS.TextChanged

    End Sub

    Private Sub cmdOpenWS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOpenWS.Click
        Try
            Dim initFolder As String = modUtil.g_nspectDocPath & "\workspace"

            If Not System.IO.Directory.Exists(initFolder) Then
                MkDir(initFolder)
            End If
            Dim dlgBrowser As New System.Windows.Forms.FolderBrowserDialog

            dlgBrowser.Description = "Choose a directory for analysis output: "
            dlgBrowser.SelectedPath = initFolder
            If dlgBrowser.ShowDialog = Windows.Forms.DialogResult.OK Then
                txtOutputWS.Text = dlgBrowser.SelectedPath
                g_strWorkspace = txtOutputWS.Text
            End If
        Catch ex As Exception
            HandleError(True, "cmdOpenWS_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Sub

    Private Sub cboLCLayer_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboLCLayer.SelectedIndexChanged
        Try
            cboLCUnits.SelectedIndex = modUtil.GetRasterDistanceUnits(cboLCLayer.Text)
        Catch ex As Exception
            HandleError(True, "cboLCLayer_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Sub

    Private Sub cboLCType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboLCType.SelectedIndexChanged
        Try
            FillCboLCCLass()
        Catch ex As Exception
            HandleError(True, "cboLCType_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Sub

    Private Sub cboSoilsLayer_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSoilsLayer.SelectedIndexChanged
        Try

            Dim strSelect As String = "SELECT * FROM Soils WHERE NAME LIKE '" & cboSoilsLayer.Text & "'"
            Dim soilCmd As New OleDbCommand(strSelect, modUtil.g_DBConn)
            Dim soilData As OleDbDataReader = soilCmd.ExecuteReader()
            soilData.Read()

            lblKFactor.Text = soilData.Item("SoilsKFileName")
            lblSoilsHyd.Text = soilData.Item("SoilsFileName")

            soilData.Close()
        Catch ex As Exception
            HandleError(True, "cboSoilsLayer_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Sub

    Private Sub chkSelectedPolys_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSelectedPolys.CheckedChanged
        If chkSelectedPolys.CheckState = 1 Then
            cboSelectPoly.Enabled = True
            lblLayer.Enabled = True
        Else
            cboSelectPoly.Enabled = False
            lblLayer.Enabled = False
        End If
    End Sub

    Private Sub cboPrecipScen_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPrecipScen.SelectedIndexChanged
        Try
            'Have to change Erosion tab based on Annual/Event driven rain event
            Dim strEvent As String

            'If define, then open new window for new definition, else select from database
            If cboPrecipScen.Text = "New precipitation scenario..." Then
                Dim newPre As New frmNewPrecip
                newPre.Init(Me, Nothing)
                newPre.ShowDialog()
            Else
                strEvent = "SELECT * FROM PRECIPSCENARIO WHERE NAME LIKE '" & cboPrecipScen.Text & "'"
                Dim eventCmd As New OleDbCommand(strEvent, modUtil.g_DBConn)
                Dim eventData As OleDbDataReader = eventCmd.ExecuteReader()
                eventData.Read()


                Select Case eventData.Item("Type").ToString
                    Case "0" 'Annual
                        frmSDR.Visible = True
                        frameRainFall.Visible = True
                        chkCalcErosion.Text = "Calculate Erosion for Annual Type Precipitation Scenario"
                        _booAnnualPrecip = True 'Set flag
                    Case "1" 'Event
                        frmSDR.Visible = False
                        frameRainFall.Visible = False
                        chkCalcErosion.Text = "Calculate Erosion for Event Type Precipitation Scenario"
                        _booAnnualPrecip = False 'Set flag
                End Select

                _strPrecipFile = eventData.Item("PrecipFileName")
                eventData.Close()
            End If
        Catch ex As Exception
            HandleError(True, "cboPrecipScen_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Sub

    Private Sub cboWSDelin_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboWSDelin.SelectedIndexChanged
        Try
            If cboWSDelin.Text = "New watershed delineation..." Then

                g_boolNewWShed = True

                Dim newWS As New frmNewWSDelin
                newWS.ShowDialog()
            End If

        Catch ex As Exception
            HandleError(True, "cboWSDelin_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Sub

    Private Sub cboWQStd_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboWQStd.SelectedIndexChanged
        Try
            If cboWQStd.Text = "New water quality standard..." Then
                Dim fNewWQ As New frmAddWQStd
                fNewWQ.ShowDialog()
            Else
                PopulatePollutants()
            End If
        Catch ex As Exception
            HandleError(True, "cboWQStd_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Sub


    Private Sub mnuNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuNew.Click
        Try
            Dim intvbYesNo As Short

            intvbYesNo = MsgBox("Do you want to save changes you made to " & Me.Text & "?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Exclamation, "N-SPECT")

            If intvbYesNo = MsgBoxResult.Yes Then
                If SaveXMLFile() Then
                    ClearForm()
                    frmProjectSetup_Load(Me, New System.EventArgs())
                Else
                    Exit Sub
                End If
            ElseIf intvbYesNo = MsgBoxResult.No Then
                ClearForm()
                frmProjectSetup_Load(Me, New System.EventArgs())
            ElseIf intvbYesNo = MsgBoxResult.Cancel Then
                Exit Sub
            End If

        Catch ex As Exception
            HandleError(True, "mnuNew_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Sub

    Private Sub mnuOpen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuOpen.Click
        LoadXMLFile()
    End Sub

    Private Sub mnuSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSave.Click
        Try
            SaveXMLFile()
        Catch ex As Exception
            HandleError(True, "mnuSave_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Sub

    Private Sub mnuSaveAs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSaveAs.Click
        Try
            _booExists = False
            SaveXMLFile()
        Catch ex As Exception
            HandleError(True, "mnuSaveAs_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Sub

    Private Sub mnuExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExit.Click
        Try
            Dim intvbYesNo As Short

            intvbYesNo = MsgBox("Do you want to save changes you made to " & Me.Text & "?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Exclamation, "N-SPECT")

            If intvbYesNo = MsgBoxResult.Yes Then
                If SaveXMLFile() Then
                    Me.Close()
                End If
            ElseIf intvbYesNo = MsgBoxResult.No Then
                Me.Close()
            ElseIf intvbYesNo = MsgBoxResult.Cancel Then
                Exit Sub
            End If

        Catch ex As Exception
            HandleError(True, "mnuExit_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Sub

    Private Sub mnuGeneralHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuGeneralHelp.Click
        System.Windows.Forms.Help.ShowHelp(Me, modUtil.g_nspectPath & "\Help\nspect.chm", "project_setup.htm")
    End Sub


    Private Sub cboLCLayer_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboWSDelin.KeyDown, cboWQStd.KeyDown, cboPrecipScen.KeyDown, cboLCUnits.KeyDown, cboLCType.KeyDown, cboLCLayer.KeyDown
        e.SuppressKeyPress = True
    End Sub

    Private Sub optUseGRID_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optUseGRID.CheckedChanged
        If sender.Checked Then
            Try
                cboRainGrid.Enabled = optUseGRID.Checked
                txtRainValue.Enabled = optUseValue.Checked
            Catch ex As Exception
                HandleError(True, "optUseGRID_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
            End Try
        End If
    End Sub

    Private Sub optUseValue_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optUseValue.CheckedChanged
        If sender.Checked Then
            Try
                txtRainValue.Enabled = optUseValue.Checked
                cboRainGrid.Enabled = optUseGRID.Checked
            Catch ex As Exception
                HandleError(True, "optUseValue_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
            End Try
        End If
    End Sub

    Private Sub chkSDR_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkSDR.CheckStateChanged
        If chkSDR.CheckState = 1 Then
            txtSDRGRID.Enabled = True
            cmdOpenSDR.Enabled = True
        Else
            txtSDRGRID.Enabled = False
            cmdOpenSDR.Enabled = False
        End If
    End Sub

    Private Sub cmdOpenSDR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOpenSDR.Click
        Dim g As New MapWinGIS.Grid
        Dim dlgOpen As New Windows.Forms.OpenFileDialog
        dlgOpen.Title = "Choose SDR GRID"
        dlgOpen.Filter = g.CdlgFilter
        If dlgOpen.ShowDialog = Windows.Forms.DialogResult.OK Then
            txtSDRGRID.Text = dlgOpen.FileName
        End If
    End Sub

    Private Sub chkCalcErosion_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkCalcErosion.CheckStateChanged
        Try
            frameRainFall.Enabled = chkCalcErosion.CheckState
            cboErodFactor.Enabled = chkCalcErosion.CheckState
            lblErodFactor.Enabled = chkCalcErosion.CheckState
            optUseGRID.Enabled = chkCalcErosion.CheckState
            optUseValue.Enabled = True 'chkCalcErosion.Value
            lblKFactor.Visible = chkCalcErosion.CheckState
            Label7.Visible = chkCalcErosion.CheckState
        Catch ex As Exception
            HandleError(True, "chkCalcErosion_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Sub


    Private Sub dgvPollutants_DataError(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvPollutants.DataError
        MsgBox("Please enter a valid value in row " + (e.RowIndex + 1).ToString + " and column " + (e.ColumnIndex + 1).ToString + ".")
    End Sub

    Private Sub dgvLandUse_DataError(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvLandUse.DataError
        MsgBox("Please enter a valid value in row " + (e.RowIndex + 1).ToString + " and column " + (e.ColumnIndex + 1).ToString + ".")
    End Sub

    Private Sub dgvManagementScen_DataError(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvManagementScen.DataError
        MsgBox("Please enter a valid value in row " + (e.RowIndex + 1).ToString + " and column " + (e.ColumnIndex + 1).ToString + ".")
    End Sub


    Private Sub dgvLandUse_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgvLandUse.MouseClick
        If e.Button = Windows.Forms.MouseButtons.Right Then
            cnxtmnuLandUse.Show(dgvLandUse, New Drawing.Point(e.X, e.Y))
            If Not dgvLandUse.CurrentRow Is Nothing Then
                If dgvLandUse.CurrentRow.Cells("LUApply").FormattedValue Or dgvLandUse.CurrentRow.Cells("LUScenario").Value <> "" Then
                    EditScenarioToolStripMenuItem.Enabled = True
                Else
                    EditScenarioToolStripMenuItem.Enabled = False
                End If
            End If
            If dgvLandUse.Rows.Count > 0 Then
                DeleteScenarioToolStripMenuItem.Enabled = True
            Else
                EditScenarioToolStripMenuItem.Enabled = False
                DeleteScenarioToolStripMenuItem.Enabled = False
            End If
        End If
    End Sub

    Private Sub dgvManagementScen_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgvManagementScen.MouseClick
        If e.Button = Windows.Forms.MouseButtons.Right Then
            cnxtmnuManagement.Show(dgvManagementScen, New Drawing.Point(e.X, e.Y))
        End If
    End Sub

    'TODO
    Private Sub AddScenarioToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddScenarioToolStripMenuItem.Click
        Try

            Dim intRow As Short = dgvLandUse.Rows.Add()

            g_intManScenRow = intRow.ToString

            g_strLUScenFileName = ""

            Dim newscen As New frmLUScen
            With newscen
                .init(cboWQStd.Text, Me)
                .Text = "Add Land Use Scenario"
                If .ShowDialog() = Windows.Forms.DialogResult.Cancel Then
                    dgvLandUse.Rows.RemoveAt(g_intManScenRow)
                End If
            End With

        Catch ex As Exception
            HandleError(False, "MnuLUAdd_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Sub

    Private Sub EditScenarioToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditScenarioToolStripMenuItem.Click
        Try
            If Not dgvLandUse.CurrentRow Is Nothing Then
                g_intManScenRow = dgvLandUse.CurrentRow.Index.ToString
                g_strLUScenFileName = dgvLandUse.CurrentRow.Cells("LUScenarioXML").Value

                Dim newscen As New frmLUScen
                With newscen
                    .init(cboWQStd.Text, Me)
                    .Text = "Edit Land Use Scenario"
                    .ShowDialog()
                End With
            End If

        Catch ex As Exception
            HandleError(False, "MnuLUEdit_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try

    End Sub

    Private Sub DeleteScenarioToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteScenarioToolStripMenuItem.Click
        Try
            With dgvLandUse
                If Not .CurrentRow Is Nothing Then
                    If .CurrentRow.Cells("LUApply").FormattedValue Or .CurrentRow.Cells("LUScenario").Value <> "" Then
                        If MsgBox("There is data in Row " & (.CurrentRow.Index + 1).ToString & ". Would you still like to delete?", MsgBoxStyle.YesNo, "Delete Row") = MsgBoxResult.Yes Then
                            .Rows.Remove(.CurrentRow)
                        End If
                    Else
                        .Rows.Remove(.CurrentRow)
                    End If
                End If
            End With
        Catch ex As Exception
            HandleError(True, "mnuLUDelete_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try

    End Sub


    Private Sub AppendRowToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AppendRowToolStripMenuItem.Click
        Dim idx As Integer = dgvManagementScen.Rows.Add()
        PopulateManagement(idx)
    End Sub

    Private Sub InsertRowToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InsertRowToolStripMenuItem.Click
        If Not dgvManagementScen.CurrentRow Is Nothing Then
            Dim idx As Integer = dgvManagementScen.CurrentRow.Index
            dgvManagementScen.Rows.Insert(idx, 1)
            PopulateManagement(idx)
        End If
    End Sub

    Private Sub DeleteCurrentRowToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteCurrentRowToolStripMenuItem.Click
        With dgvManagementScen
            If Not .CurrentRow Is Nothing Then
                Dim idx As Integer = .CurrentRow.Index
                If .Rows(idx).Cells("ManageApply").Value Or .Rows(idx).Cells("ChangeAreaLayer").Value <> "" Or .Rows(idx).Cells("ChangeToClass").Value <> "" Then
                    If MsgBox("There is data in Row " & (idx + 1).ToString & ". Would you still like to delete?", MsgBoxStyle.YesNo, "Delete Row") = MsgBoxResult.Yes Then
                        .Rows.Remove(.CurrentRow)
                    End If
                Else
                    .Rows.Remove(.CurrentRow)
                End If
            End If
            If .Rows.Count = 0 Then
                .Rows.Add()
                PopulateManagement(0)
            End If
        End With
    End Sub


    Private Sub cmdQuit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdQuit.Click
        Try
            Dim intvbYesNo As Short

            intvbYesNo = MsgBox("Do you want to save changes you made to " & Me.Text & "?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Exclamation, "N-SPECT")

            If intvbYesNo = MsgBoxResult.Yes Then
                If SaveXMLFile() Then
                    Me.Close()
                End If
            ElseIf intvbYesNo = MsgBoxResult.No Then
                Me.Close()
            ElseIf intvbYesNo = MsgBoxResult.Cancel Then
                Exit Sub
            End If


        Catch ex As Exception
            HandleError(True, "cmdQuit_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Sub


    'TODO
    Private Sub cmdRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRun.Click

        Dim strWaterShed As String 'Connection string
        Dim strPrecip As String 'Connection String
        Dim lngWShedLayerIndex As Integer 'Watershed layer index

        Dim booLUItems As Boolean 'Are there Landuse Scenarios???
        Dim dictPollutants As New Generic.Dictionary(Of String, String) 'Dict to hold all pollutants
        Dim i As Integer
        Dim strProjectInfo As String 'String that will hold contents of prj file for inclusion in metatdata

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        'STEP 1: Save file, populate xml params: -------------------------------------------------------------------------
        If Not SaveXMLFile() Then
            System.Windows.Forms.Cursor.Current = Cursors.Default
            Exit Sub
        End If

        'Init your global dictionary to hold the metadata records as well as the global xml prj file
        g_dicMetadata = New Generic.Dictionary(Of String, String)
        g_clsXMLPrjFile = _XMLPrjParams
        'END STEP 1: -----------------------------------------------------------------------------------------------------


        'STEP 2: Identify if local effects are being used : --------------------------------------------------------------
        'Local Effects Global
        If _XMLPrjParams.intLocalEffects = 1 Then
            g_booLocalEffects = True
        Else
            g_booLocalEffects = False
        End If
        'END STEP 2: -----------------------------------------------------------------------------------------------------

        'STEP 3: Find out if user is making use of only the selected Sheds -----------------------------------------------
        'Selected Sheds only
        Dim lyrSelectedPolyLayer As MapWindow.Interfaces.Layer = Nothing
        If _XMLPrjParams.intSelectedPolys = 1 Then
            g_booSelectedPolys = True
            lngWShedLayerIndex = modUtil.GetLayerIndex(cboSelectPoly.Text)
            lyrSelectedPolyLayer = g_MapWin.Layers(lngWShedLayerIndex)
        Else
            g_booSelectedPolys = False
        End If
        'END STEP 3: ---------------------------------------------------------------------------------------------------------

        'STEP 4: Get the Management Scenarios: ------------------------------------------------------------------------------------
        'If they're using, we send them over to modMgmtScen to implement
        If _XMLPrjParams.clsMgmtScenHolder.Count > 0 Then
            modMgmtScen.MgmtScenSetup(_XMLPrjParams.clsMgmtScenHolder, _XMLPrjParams.strLCGridType, _XMLPrjParams.strLCGridFileName, _XMLPrjParams.strProjectWorkspace)
        End If
        'END STEP 4: ---------------------------------------------------------------------------------------------------------

        'STEP 5: Pollutant Dictionary creation, needed for Landuse -----------------------------------------------------------
        'Go through and find the pollutants, if they're used and what the CoeffSet is
        'We're creating a dictionary that will hold Pollutant, Coefficient Set for use in the Landuse Scenarios
        For i = 0 To _XMLPrjParams.clsPollItems.Count - 1
            If _XMLPrjParams.clsPollItems.Item(i).intApply = 1 Then
                dictPollutants.Add(_XMLPrjParams.clsPollItems.Item(i).strPollName, _XMLPrjParams.clsPollItems.Item(i).strCoeffSet)
            End If
        Next i
        'END STEP 5: ---------------------------------------------------------------------------------------------------------

        'STEP 6: Landuses sent off to modLanduse for processing -----------------------------------------------------
        For i = 0 To _XMLPrjParams.clsLUItems.Count - 1
            If _XMLPrjParams.clsLUItems.Item(i).intApply = 1 Then
                booLUItems = True
                modLanduse.Begin(_XMLPrjParams.strLCGridType, _XMLPrjParams.clsLUItems, dictPollutants, _XMLPrjParams.strLCGridFileName, _XMLPrjParams.strProjectWorkspace)
                Exit For
            Else
                booLUItems = False
            End If
        Next i
        'END STEP 6: ---------------------------------------------------------------------------------------------------------

        'STEP 7: ---------------------------------------------------------------------------------------------------------
        'Obtain Watershed values

        strWaterShed = "Select * from WSDelineation Where Name like '" & _XMLPrjParams.strWaterShedDelin & "'"
        Dim cmdWS As New OleDbCommand(strWaterShed, g_DBConn)

        'END STEP 7: -----------------------------------------------------------------------------------------------------

        'STEP 8: ---------------------------------------------------------------------------------------------------------
        'Set the Analysis Environment and globals for output workspace

        modMainRun.SetGlobalEnvironment(cmdWS, lyrSelectedPolyLayer)

        'END STEP 8: -----------------------------------------------------------------------------------------------------

        'STEP 8a: --------------------------------------------------------------------------------------------------------
        'Added 1/08/2007 to account for non-adjacent polygons
        If _XMLPrjParams.intSelectedPolys = 1 Then
            If modMainRun.CheckMultiPartPolygon(g_pSelectedPolyClip) Then
                MsgBox("Warning: Your selected polygons are not adjacent.  Please select only polygons that are adjacent.", MsgBoxStyle.Critical, "Non-adjacent Polygons Detected")
                System.Windows.Forms.Cursor.Current = Cursors.Default
                Exit Sub
            End If
        End If

        'STEP 9: ---------------------------------------------------------------------------------------------------------
        'Create the runoff GRID
        'Get the precip scenario stuff
        strPrecip = "Select * from PrecipScenario where name like '" & _XMLPrjParams.strPrecipScenario & "'"
        Dim cmdPrecip As New OleDbCommand(strPrecip, g_DBConn)
        Dim dataPrecip As OleDbDataReader = cmdPrecip.ExecuteReader()
        dataPrecip.Read()
        'Added 6/04 to account for different PrecipTypes
        modMainRun.g_intPrecipType = dataPrecip.Item("PrecipType")
        dataPrecip.Close()

        'If there has been a land use added, then a new LCType has been created, hence we get it from g_strLCTypename
        Dim strLCType As String
        If booLUItems Then
            strLCType = modLanduse.g_strLCTypeName
        Else
            strLCType = _XMLPrjParams.strLCGridType
        End If

        If Not modRunoff.CreateRunoffGrid(_XMLPrjParams.strLCGridFileName, strLCType, cmdPrecip, _XMLPrjParams.strSoilsHydFileName) Then
            Exit Sub
        End If
        'END STEP 9: -----------------------------------------------------------------------------------------------------

        'STEP 10: ---------------------------------------------------------------------------------------------------------
        'Process pollutants
        For i = 0 To _XMLPrjParams.clsPollItems.Count - 1
            If _XMLPrjParams.clsPollItems.Item(i).intApply = 1 Then
                'If user is NOT ignoring the pollutant then send the whole item over along with LCType
                If Not modPollutantCalcs.PollutantConcentrationSetup(_XMLPrjParams.clsPollItems.Item(i), _XMLPrjParams.strLCGridType, _XMLPrjParams.strWaterQuality) Then
                    Exit Sub
                End If
            End If
        Next i
        'END STEP 10: -----------------------------------------------------------------------------------------------------

        'Step 11: Erosion -------------------------------------------------------------------------------------------------
        'Check that they have chosen Erosion
        Dim dataWS As OleDbDataReader = cmdWS.ExecuteReader
        dataWS.Read()
        If _XMLPrjParams.intCalcErosion = 1 Then
            If _booAnnualPrecip Then 'If Annual (0) then TRUE, ergo RUSLE
                If _XMLPrjParams.intRainGridBool Then
                    If Not modRusle.RUSLESetup(dataWS.Item("NibbleFileName"), dataWS.Item("dem2bfilename"), _XMLPrjParams.strRainGridFileName, _XMLPrjParams.strSoilsKFileName, _XMLPrjParams.strSDRGridFileName, _XMLPrjParams.strLCGridType) Then
                        Exit Sub
                    End If
                ElseIf _XMLPrjParams.intRainConstBool Then
                    If Not modRusle.RUSLESetup(dataWS.Item("NibbleFileName"), dataWS.Item("dem2bfilename"), _XMLPrjParams.strRainGridFileName, _XMLPrjParams.strSoilsKFileName, _XMLPrjParams.strSDRGridFileName, _XMLPrjParams.strLCGridType, _XMLPrjParams.dblRainConstValue) Then
                        Exit Sub
                    End If
                End If

            Else 'If event (1) then False, ergo MUSLE
                If Not modMUSLE.MUSLESetup(_XMLPrjParams.strSoilsDefName, _XMLPrjParams.strSoilsKFileName, _XMLPrjParams.strLCGridType) Then
                    Exit Sub
                End If
            End If
        End If
        dataWS.Close()
        'STEP 11: ----------------------------------------------------------------------------------------------------------

        'STEP 12 : Cleanup any temp critters -------------------------------------------------------------------------------
        'g_DictTempNames holds the names of all temporary landuses and/or coefficient sets created during the Landuse scenario
        'portion of our program, for example CCAP1, or NitSet1.  We now must eliminate them from the database if they exist.
        If g_DictTempNames.Count > 0 Then
            If booLUItems Then
                modLanduse.Cleanup(g_DictTempNames, (_XMLPrjParams.clsPollItems), (_XMLPrjParams.strLCGridType))
            End If
        End If
        'END STEP 12: -------------------------------------------------------------------------------------------------------

        'TODO
        'STEP 13: -----------------------------------------------------------------------------------------------------------
        'g_pGroupLayer has been created earlier and has been taken on GRIDs since.  Now lets add it
        'Add the group layer.
        'With g_pGroupLayer
        '    .Expanded = True 'Are going to 'expand' it
        '    .Name = _XMLPrjParams.strProjectName 'The name equals whatever the user entered
        'End With

        ''UPGRADE_WARNING: Couldn't resolve default property of object g_pGroupLayer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'm_pMap.AddLayer(g_pGroupLayer)
        'END STEP 13: -------------------------------------------------------------------------------------------------------

        'STEP 14 save out group layer ---------------------------------------------------------------------------------------
        'TODO figure alternative for layer
        'modUtil.ExportLayerToPath(g_pGroupLayer, _XMLPrjParams.strProjectWorkspace & "\" & _XMLPrjParams.strProjectName & ".lyr")
        'END STEP 14: -------------------------------------------------------------------------------------------------------

        'STEP 15: create string describing project parameters ---------------------------------------------------------------
        'TODO metadata stuff
        'strProjectInfo = modUtil.ParseProjectforMetadata(_XMLPrjParams, _strFileName)
        'END STEP 15: -------------------------------------------------------------------------------------------------------

        'STEP 16: Apply the metadata to each of the rasters in the group layer ----------------------------------------------
        'TODO
        'm_App.StatusBar.Message(0) = "Creating metadata for the N-SPECT group layer..."
        'modUtil.CreateMetadata(g_pGroupLayer, strProjectInfo)
        'END STEP 16: -------------------------------------------------------------------------------------------------------

        'Cleanup ------------------------------------------------------------------------------------------------------------
        'TODO
        'm_App.StatusBar.Message(0) = "Deleting temporary files..."

        'Go into workspace and rid it of all rasters
        modUtil.CleanGlobals()
        modUtil.CleanupRasterFolder((_XMLPrjParams.strProjectWorkspace))

        System.Windows.Forms.Cursor.Current = Cursors.Default

        'TODO
        'm_App.StatusBar.Message(0) = "N-SPECT processing complete!"

        Me.Close()

        Exit Sub

UserCancel:
        modProgDialog.KillDialog()
        System.Windows.Forms.Cursor.Current = Cursors.Default
        MsgBox("Processing has been stopped.", MsgBoxStyle.Information, "Analysis Stopped")

ErrorHandler:
        HandleError(True, "cmdRun_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        System.Windows.Forms.Cursor.Current = Cursors.Default

    End Sub

#End Region

#Region "Helper Functions"
    Private Sub ClearForm()
        Try
            'Gotta clean up before new, clean form

            'LandClass stuff
            cboLCLayer.Items.Clear()
            cboLCType.Items.Clear()

            'DBase scens
            cboPrecipScen.Items.Clear()
            cboWSDelin.Items.Clear()
            cboWQStd.Items.Clear()
            cboSoilsLayer.Items.Clear()

            'Text
            txtProjectName.Text = ""
            txtOutputWS.Text = ""

            'Checkboxes
            chkSelectedPolys.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkLocalEffects.CheckState = System.Windows.Forms.CheckState.Unchecked

            'Erosion
            cboRainGrid.Items.Clear()
            optUseGRID.Checked = True
            chkCalcErosion.CheckState = System.Windows.Forms.CheckState.Unchecked
            txtRainValue.Text = ""

            'txtOutputFile.Text = ""
            txtThemeName.Text = ""

            'clear the grids
            dgvPollutants.Rows.Clear()
            dgvLandUse.Rows.Clear()
            dgvManagementScen.Rows.Clear()

            frmProjectSetup_Load(Nothing, Nothing)
        Catch ex As Exception
            HandleError(False, "ClearForm " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Sub

    Private Sub LoadXMLFile()

        'browse...get output filename
        Dim strFolder As String = modUtil.g_nspectDocPath & "\projects"
        If Not IO.Directory.Exists(strFolder) Then
            MkDir(strFolder)
        End If
        Dim dlgXMLOpen As New Windows.Forms.OpenFileDialog
        With dlgXMLOpen
            .Filter = MSG8
            .InitialDirectory = strFolder
            .Title = "Open N-SPECT Project File"
            .FilterIndex = 1
            .ShowDialog()
        End With

        If Len(dlgXMLOpen.FileName) > 0 Then
            _strFileName = Trim(dlgXMLOpen.FileName)
            _XMLPrjParams.XML = _strFileName
            FillForm()
        Else
            Exit Sub
        End If

        'Pop this string with the incoming name, if they change, we'll prompt to 'save as'.
        _strOpenFileName = txtProjectName.Text


        'Exit Sub
        'ErrorHandler:
        '  HandleError False, "LoadXMLFile " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
    End Sub

    Private Function EnableChkWaterShed() As Boolean

        Dim strWShed As String

        Try
            strWShed = "SELECT WSFILENAME FROM WSDELINEATION WHERE NAME LIKE '" & cboWSDelin.Text & "'"
            Dim wsCmd As New OleDbCommand(strWShed, modUtil.g_DBConn)
            Dim ws As OleDbDataReader = wsCmd.ExecuteReader()
            ws.Read()

            _strWShed = modUtil.SplitFileName(ws.Item("wsfilename"))

            If g_MapWin.View.SelectedShapes.NumSelected > 0 Then
                EnableChkWaterShed = True
            Else
                EnableChkWaterShed = False
            End If
            ws.Close()
        Catch ex As Exception
            EnableChkWaterShed = False
        End Try
    End Function

    Private Sub FillCboLCCLass()
        Try
            Dim strLCChanges As String
            Dim i As Short

            strLCChanges = "SELECT LCCLASS.Name as Name2, LCTYPE.LCTYPEID FROM LCTYPE INNER JOIN LCCLASS " & "ON LCTYPE.LCTYPEID = LCCLASS.LCTYPEID WHERE LCTYPE.NAME LIKE '" & cboLCType.Text & "'"
            Dim lcChangesCmd As New OleDbCommand(strLCChanges, modUtil.g_DBConn)
            Dim lcChanges As OleDbDataReader = lcChangesCmd.ExecuteReader()

            arrClassList.Clear()

            Do While lcChanges.Read()
                arrClassList.Insert(i, lcChanges.Item("Name2"))
            Loop

            lcChanges.Close()

        Catch ex As Exception
            HandleError(False, "FillCboLCCLass " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Sub

    Private Sub PopulatePollutants()
        Try
            Dim strSQLWQStd As String

            Dim strSQLWQStdPoll As String

            'Selection based on combo box
            strSQLWQStd = "SELECT * FROM WQCRITERIA WHERE NAME LIKE '" & cboWQStd.Text & "'"
            Dim WQCritCmd As OleDbCommand = New OleDbCommand(strSQLWQStd, g_DBConn)
            Dim dataWQCrit As OleDbDataReader = WQCritCmd.ExecuteReader()


            If dataWQCrit.HasRows Then
                dataWQCrit.Read()

                strSQLWQStdPoll = "SELECT POLLUTANT.NAME, POLL_WQCRITERIA.THRESHOLD " & "FROM POLL_WQCRITERIA INNER JOIN POLLUTANT " & "ON POLL_WQCRITERIA.POLLID = POLLUTANT.POLLID Where POLL_WQCRITERIA.WQCRITID = " & dataWQCrit.Item("WQCRITID")

                Dim WQStdCmd As OleDbCommand = New OleDbCommand(strSQLWQStdPoll, g_DBConn)

                Dim WQStdAdapter As OleDbDataAdapter = New OleDbDataAdapter(WQStdCmd)
                Dim PollutantsTable As New DataTable

                WQStdAdapter.Fill(PollutantsTable)

                'Don't actually datasource since there's no two-way communication. Just fill grid from the table
                Dim strPollName As String
                Dim idx As Integer
                dgvPollutants.Rows.Clear()
                For Each row As DataRow In PollutantsTable.Rows
                    idx = dgvPollutants.Rows.Add()
                    strPollName = row.Item(0)

                    dgvPollutants.Rows(idx).Cells("PollutantName").Value = strPollName
                    dgvPollutants.Rows(idx).Cells("Threshold").Value = row.Item(1)

                    PopulateCoefType(strPollName, idx)
                Next
            Else
                MsgBox("Warning: There are no water quality standards remaining.  Please add a new one.", MsgBoxStyle.Critical, "Recordset Empty")
            End If

        Catch ex As Exception
            HandleError(False, "PopPollutants " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Sub

    Private Sub PopulateCoefType(ByVal strPollutantName As String, ByVal rowidx As Integer)
        Dim strSelectCoeff As String = "SELECT POLLUTANT.POLLID, POLLUTANT.NAME, COEFFICIENTSET.NAME AS NAME2 FROM POLLUTANT INNER JOIN COEFFICIENTSET " & "ON POLLUTANT.POLLID = COEFFICIENTSET.POLLID Where POLLUTANT.NAME LIKE '" & strPollutantName & "'"
        Dim coefCmd As New OleDbCommand(strSelectCoeff, modUtil.g_DBConn)
        Dim coefData As OleDbDataReader = coefCmd.ExecuteReader()
        Dim cell As DataGridViewComboBoxCell = dgvPollutants.Rows(rowidx).Cells("CoefSet")
        While coefData.Read
            cell.Items.Add(coefData.Item("Name2"))
        End While
        coefData.Close()
    End Sub

    Private Sub PopulateManagement(ByVal rowidx As Integer)

        Dim areacell As DataGridViewComboBoxCell = dgvManagementScen.Rows(rowidx).Cells("ChangeAreaLayer")
        areacell.Items.Clear()
        For i As Integer = 0 To arrAreaList.Count - 1
            areacell.Items.Add(arrAreaList(i))
        Next

        Dim classcell As DataGridViewComboBoxCell = dgvManagementScen.Rows(rowidx).Cells("ChangeToClass")
        classcell.Items.Clear()
        For i As Integer = 0 To arrClassList.Count - 1
            classcell.Items.Add(arrClassList(i))
        Next
    End Sub

    Private Sub FillForm()
        Try
            Dim i As Integer
            Dim strCurrWShed As String
            Dim lngCurrWshedPolyIndex As Integer
            Dim intYesNo As Short

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            txtProjectName.Text = _XMLPrjParams.strProjectName
            txtOutputWS.Text = _XMLPrjParams.strProjectWorkspace

            'Step 1:  LandCoverGrid
            'Check to see if the LC cover is in the map, if so, set the combobox
            If modUtil.LayerInMap((_XMLPrjParams.strLCGridName)) Then
                cboLCLayer.SelectedIndex = modUtil.GetCboIndex((_XMLPrjParams.strLCGridName), cboLCLayer)
                cboLCLayer.Refresh()
            Else
                If modUtil.AddRasterLayerToMapFromFileName(_XMLPrjParams.strLCGridFileName) Then

                    With cboLCLayer
                        .Items.Add(_XMLPrjParams.strLCGridName)
                        .Refresh()
                        .SelectedIndex = modUtil.GetCboIndex(_XMLPrjParams.strLCGridName, cboLCLayer)
                    End With

                Else
                    intYesNo = MsgBox("Could not find the Land Cover dataset: " & _XMLPrjParams.strLCGridFileName & ".  Would you like " & "to browse for it?", MsgBoxStyle.YesNo, "Cannot Locate Dataset")
                    If intYesNo = MsgBoxResult.Yes Then
                        _XMLPrjParams.strLCGridFileName = modUtil.AddInputFromGxBrowser("Raster")
                        If _XMLPrjParams.strLCGridFileName <> "" Then
                            If modUtil.AddRasterLayerToMapFromFileName(_XMLPrjParams.strLCGridFileName) Then
                                cboLCLayer.SelectedIndex = 0
                            End If
                        Else
                            Exit Sub
                        End If
                    Else
                        Exit Sub
                    End If
                End If

            End If

            cboLCUnits.SelectedIndex = modUtil.GetCboIndex((_XMLPrjParams.strLCGridUnits), cboLCUnits)
            cboLCUnits.Refresh()

            cboLCType.SelectedIndex = modUtil.GetCboIndex((_XMLPrjParams.strLCGridType), cboLCType)
            cboLCType.Refresh()

            'Step 2: Soils - same process, if in doc and map, OK, else send em looking
            If System.IO.File.Exists(_XMLPrjParams.strSoilsHydFileName) Or System.IO.File.Exists(_XMLPrjParams.strSoilsHydFileName + "\sta.adf") Then
                cboSoilsLayer.SelectedIndex = modUtil.GetCboIndex((_XMLPrjParams.strSoilsDefName), cboSoilsLayer)
                cboSoilsLayer.Refresh()
            Else
                MsgBox("Could not find soils dataset.  Please correct the soils definition in the Advanced Settings.", MsgBoxStyle.Critical, "Dataset Missing")
                Exit Sub
            End If

            'Step5: Precip Scenario
            cboPrecipScen.SelectedIndex = modUtil.GetCboIndex((_XMLPrjParams.strPrecipScenario), cboPrecipScen)

            'Step6: Watershed Delineation
            cboWSDelin.SelectedIndex = modUtil.GetCboIndex((_XMLPrjParams.strWaterShedDelin), cboWSDelin)

            'Add the basinpoly to the map
            strCurrWShed = "Select * from WSDelineation where Name Like '" & _XMLPrjParams.strWaterShedDelin & "'"
            Dim WSCmd As New OleDbCommand(strCurrWShed, modUtil.g_DBConn)
            Dim wsData As OleDbDataReader = WSCmd.ExecuteReader
            wsData.Read()
            Dim strBasin As String = wsData.Item("wsfilename")
            If IO.Path.GetExtension(strBasin) <> ".shp" Then
                strBasin = strBasin + ".shp"
            End If

            If Not modUtil.LayerInMapByFileName(strBasin) Then
                If modUtil.AddFeatureLayerToMapFromFileName(strBasin, wsData.Item("Name") & " " & "Drainage Basins") Then
                    lngCurrWshedPolyIndex = modUtil.GetLayerIndex(wsData.Item("Name") & " " & "Drainage Basins")

                    cboSelectPoly.Items.Add(wsData.Item("Name") & " " & "Drainage Basins")
                    arrAreaList.Add(wsData.Item("Name") & " " & "Drainage Basins")
                Else
                    MsgBox("Could not find watershed layer: " & wsData.Item("wsfilename") & " .  Please add the watershed layer to the map.", MsgBoxStyle.Critical, "File Not Found")
                End If
            End If
            wsData.Close()

            'Step7: Water Quality
            cboWQStd.SelectedIndex = modUtil.GetCboIndex((_XMLPrjParams.strWaterQuality), cboWQStd)

            'Step8: LocalEffects/Selected Watersheds
            chkLocalEffects.CheckState = _XMLPrjParams.intLocalEffects
            chkSelectedPolys.CheckState = _XMLPrjParams.intSelectedPolys

            If chkSelectedPolys.CheckState = 1 Then
                '1st see if it's in the map
                Dim strSelected As String = _XMLPrjParams.strSelectedPolyFileName
                If IO.Path.GetExtension(strSelected) <> ".shp" Then
                    strSelected = strSelected + ".shp"
                End If

                If modUtil.LayerInMapByFileName(strSelected) Then
                    cboSelectPoly.SelectedIndex = modUtil.GetCboIndex(strSelected, cboSelectPoly)
                Else
                    'Not there then add it
                    If modUtil.AddFeatureLayerToMapFromFileName(strSelected, _XMLPrjParams.strSelectedPolyLyrName) Then
                        cboSelectPoly.Items.Add(_XMLPrjParams.strSelectedPolyLyrName)
                        arrAreaList.Add(_XMLPrjParams.strSelectedPolyLyrName)

                        cboSelectPoly.SelectedIndex = modUtil.GetCboIndex(_XMLPrjParams.strSelectedPolyLyrName, cboSelectPoly)
                    Else
                        'Can't find it, then send em searching
                        intYesNo = MsgBox("Could not find the Selected Polygons file used to limit extent: " & strSelected & ".  Would you like to browse for it? ", MsgBoxStyle.YesNo, "Cannot Locate Dataset")
                        If intYesNo = MsgBoxResult.Yes Then
                            'if they want to look for it then give em the browser
                            _XMLPrjParams.strSelectedPolyFileName = modUtil.AddInputFromGxBrowser("Feature")
                            'if they actually find something, throw it in the map
                            If Len(_XMLPrjParams.strSelectedPolyFileName) > 0 Then
                                If modUtil.AddFeatureLayerToMapFromFileName(_XMLPrjParams.strSelectedPolyFileName) Then
                                    cboSelectPoly.Items.Add(modUtil.SplitFileName(_XMLPrjParams.strSelectedPolyFileName))
                                    arrAreaList.Add(modUtil.SplitFileName(_XMLPrjParams.strSelectedPolyFileName))

                                    cboSelectPoly.SelectedIndex = modUtil.GetCboIndex(modUtil.SplitFileName(_XMLPrjParams.strSelectedPolyFileName), cboSelectPoly)
                                End If
                            End If
                        Else
                            _XMLPrjParams.intSelectedPolys = 0
                            chkSelectedPolys.CheckState = System.Windows.Forms.CheckState.Unchecked
                        End If
                    End If
                End If
            End If

            'Step: Erosion Tab - Calc Erosion, Erosion Attribute
            If _XMLPrjParams.intCalcErosion = 1 Then
                chkCalcErosion.CheckState = System.Windows.Forms.CheckState.Checked
            Else
                chkCalcErosion.CheckState = System.Windows.Forms.CheckState.Unchecked
            End If

            'Step: Erosion Tab - Precip
            'Either they use the GRID
            optUseGRID.Checked = _XMLPrjParams.intRainGridBool
            If optUseGRID.Checked Then
                If modUtil.LayerInMap(_XMLPrjParams.strRainGridName) Then
                    cboRainGrid.SelectedIndex = modUtil.GetCboIndex(_XMLPrjParams.strRainGridName, cboRainGrid)
                Else
                    If modUtil.AddRasterLayerToMapFromFileName(_XMLPrjParams.strRainGridFileName) Then
                        With cboRainGrid
                            .Items.Add(_XMLPrjParams.strRainGridName)
                            .Refresh()
                            .SelectedIndex = modUtil.GetCboIndex(_XMLPrjParams.strRainGridName, cboRainGrid)
                        End With
                    Else
                        intYesNo = MsgBox("Could not find Rainfall GRID: " & _XMLPrjParams.strRainGridName & ".  Would you like " & "to browse for it?", MsgBoxStyle.YesNo, "Cannot Locate Dataset")
                        If intYesNo = MsgBoxResult.Yes Then
                            _XMLPrjParams.strRainGridFileName = modUtil.AddInputFromGxBrowser("Raster")
                            If modUtil.AddRasterLayerToMapFromFileName(_XMLPrjParams.strRainGridFileName) Then
                                With cboRainGrid
                                    .Items.Add(_XMLPrjParams.strRainGridName)
                                    .Refresh()
                                    .SelectedIndex = modUtil.GetCboIndex(_XMLPrjParams.strRainGridName, cboRainGrid)
                                End With
                            End If
                        Else
                            Exit Sub
                        End If
                    End If
                End If
            End If

            'Or they use a constant value
            optUseValue.Checked = _XMLPrjParams.intRainConstBool

            If optUseValue.Checked Then
                txtRainValue.Text = CStr(_XMLPrjParams.dblRainConstValue)
            End If

            'SDR GRID
            'If Not _XMLPrjParams.intUseOwnSDR Is Nothing Then
            'On Error GoTo VersionProblem
            If _XMLPrjParams.intUseOwnSDR = 1 Then
                chkSDR.CheckState = System.Windows.Forms.CheckState.Checked
                txtSDRGRID.Text = _XMLPrjParams.strLCGridFileName
            Else
                chkSDR.CheckState = System.Windows.Forms.CheckState.Unchecked
                txtSDRGRID.Text = _XMLPrjParams.strSDRGridFileName
            End If
            'End If

            'Step Pollutants
            _intPollCount = _XMLPrjParams.clsPollItems.Count
            Dim idx As Integer
            Dim strPollName As String
            If _intPollCount > 0 Then
                dgvPollutants.Rows.Clear()

                For i = 0 To _intPollCount - 1
                    With dgvPollutants
                        idx = .Rows.Add()

                        .Rows(idx).Cells("PollApply").Value = _XMLPrjParams.clsPollItems.Item(i).intApply
                        .Rows(idx).Cells("PollutantName").Value = _XMLPrjParams.clsPollItems.Item(i).strPollName
                        strPollName = _XMLPrjParams.clsPollItems.Item(i).strPollName
                        PopulateCoefType(strPollName, idx)
                        .Rows(idx).Cells("CoefSet").Value = _XMLPrjParams.clsPollItems.Item(i).strCoeffSet

                        .Rows(idx).Cells("WhichCoeff").Value = _XMLPrjParams.clsPollItems.Item(i).strCoeff
                        .Rows(idx).Cells("Threshold").Value = CStr(_XMLPrjParams.clsPollItems.Item(i).intThreshold)

                        If Len(_XMLPrjParams.clsPollItems.Item(i).strTypeDefXMLFile) > 0 Then
                            .Rows(idx).Cells("TypeDef").Value = CStr(_XMLPrjParams.clsPollItems.Item(i).strTypeDefXMLFile)
                        End If
                    End With
                Next i

            End If

            'Step - Land Uses
            _intLUCount = _XMLPrjParams.clsLUItems.Count

            If _intLUCount > 0 Then
                dgvLandUse.Rows.Clear()
                For i = 0 To _intLUCount - 1
                    With dgvLandUse
                        idx = .Rows.Add()
                        .Rows(idx).Cells("LUApply").Value = _XMLPrjParams.clsLUItems.Item(i).intApply
                        .Rows(idx).Cells("LUScenario").Value = _XMLPrjParams.clsLUItems.Item(i).strLUScenName
                        .Rows(idx).Cells("LUScenarioXML").Value = _XMLPrjParams.clsLUItems.Item(i).strLUScenXMLFile
                    End With
                Next i
            End If

            'Step Management Scenarios
            _intMgmtCount = _XMLPrjParams.clsMgmtScenHolder.Count

            If _intMgmtCount > 0 Then
                dgvManagementScen.Rows.Clear()
                For i = 0 To _intMgmtCount - 1
                    With dgvManagementScen
                        idx = .Rows.Add()

                        If modUtil.LayerInMap(_XMLPrjParams.clsMgmtScenHolder.Item(i).strAreaName) Then
                            PopulateManagement(idx)
                            .Rows(idx).Cells("ManageApply").Value = _XMLPrjParams.clsMgmtScenHolder.Item(i).intApply
                            .Rows(idx).Cells("ChangeAreaLayer").Value = _XMLPrjParams.clsMgmtScenHolder.Item(i).strAreaName
                            .Rows(idx).Cells("ChangeToClass").Value = _XMLPrjParams.clsMgmtScenHolder.Item(i).strChangeToClass
                        Else
                            If modUtil.AddFeatureLayerToMapFromFileName(_XMLPrjParams.clsMgmtScenHolder.Item(i).strAreaFileName) Then
                                cboSelectPoly.Items.Add(_XMLPrjParams.clsMgmtScenHolder.Item(i).strAreaName)
                                arrAreaList.Add(_XMLPrjParams.clsMgmtScenHolder.Item(i).strAreaName)

                                PopulateManagement(idx)
                                .Rows(idx).Cells("ManageApply").Value = _XMLPrjParams.clsMgmtScenHolder.Item(i).intApply
                                .Rows(idx).Cells("ChangeAreaLayer").Value = _XMLPrjParams.clsMgmtScenHolder.Item(i).strAreaName
                                .Rows(idx).Cells("ChangeToClass").Value = _XMLPrjParams.clsMgmtScenHolder.Item(i).strChangeToClass
                            Else
                                intYesNo = MsgBox("Could not find Management Sceario Area Layer: " & _XMLPrjParams.clsMgmtScenHolder.Item(i).strAreaFileName & ".  Would you like " & "to browse for it?", MsgBoxStyle.YesNo, "Cannot Locate Dataset:" & _XMLPrjParams.clsMgmtScenHolder.Item(i).strAreaFileName)
                                If intYesNo = MsgBoxResult.Yes Then
                                    _XMLPrjParams.clsMgmtScenHolder.Item(i).strAreaFileName = modUtil.AddInputFromGxBrowser("Feature")
                                    If modUtil.AddFeatureLayerToMapFromFileName(_XMLPrjParams.clsMgmtScenHolder.Item(i).strAreaFileName) Then
                                        cboSelectPoly.Items.Add(_XMLPrjParams.clsMgmtScenHolder.Item(i).strAreaName)
                                        arrAreaList.Add(_XMLPrjParams.clsMgmtScenHolder.Item(i).strAreaName)

                                        PopulateManagement(idx)
                                        .Rows(idx).Cells("ManageApply").Value = _XMLPrjParams.clsMgmtScenHolder.Item(i).intApply
                                        .Rows(idx).Cells("ChangeAreaLayer").Value = _XMLPrjParams.clsMgmtScenHolder.Item(i).strAreaName
                                        .Rows(idx).Cells("ChangeToClass").Value = _XMLPrjParams.clsMgmtScenHolder.Item(i).strChangeToClass
                                    End If
                                Else
                                    Exit Sub
                                End If
                            End If
                        End If
                    End With
                Next i
            Else
                'Clear and add new row to catch new area files and such
                dgvManagementScen.Rows.Clear()
                dgvManagementScen.Rows.Add()
                PopulateManagement(0)
            End If

            'Reset to first tab
            SSTab1.SelectedIndex = 0
            _booExists = True
            System.Windows.Forms.Cursor.Current = Cursors.Default

        Catch ex As Exception
            HandleError(False, "FillForm " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Sub


    Private Function SaveXMLFile() As Boolean
        Try
            Dim strFolder As String
            Dim intvbYesNo As Short
            Dim dlgXML As New SaveFileDialog


            strFolder = modUtil.g_nspectDocPath & "\projects"
            If Not IO.Directory.Exists(strFolder) Then
                MkDir(strFolder)
            End If

            If Not ValidateData() Then 'check form inputs
                SaveXMLFile = False
                Exit Function
            End If

            'If it does not already exist, open Save As... dialog
            If Not _booExists Then
                With dlgXML
                    .Filter = MSG8
                    .Title = "Save Project File As..."
                    .InitialDirectory = strFolder
                    .FileName = txtProjectName.Text
                    .ShowDialog()
                End With

                'check to make sure filename length is greater than zeros
                If Len(dlgXML.FileName) > 0 Then
                    _strFileName = Trim(dlgXML.FileName)
                    _booExists = True
                    _XMLPrjParams.SaveFile(_strFileName)
                    SaveXMLFile = True
                Else
                    SaveXMLFile = False
                    Exit Function
                End If

            Else
                'Now check to see if the name changed
                If _strOpenFileName <> txtProjectName.Text Then
                    intvbYesNo = MsgBox("You have changed the name of this project.  Would you like to save your settings as a new file?" & vbNewLine & vbTab & "Yes" & vbTab & " -    Save as new N-SPECT project file" & vbNewLine & vbTab & "No" & vbTab & " -    Save changes to current N-SPECT project file" & vbNewLine & vbTab & "Cancel" & vbTab & " -    Return to the project window", MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, "N-SPECT")

                    If intvbYesNo = MsgBoxResult.Yes Then
                        With dlgXML
                            .Filter = MSG8
                            .Title = "Save Project File As..."
                            .InitialDirectory = strFolder
                            .FilterIndex = 1
                            .FileName = txtProjectName.Text
                            .ShowDialog()
                        End With

                        'check to make sure filename length is greater than zeros
                        If Len(dlgXML.FileName) > 0 Then
                            _strFileName = Trim(dlgXML.FileName)
                            _booExists = True
                            _XMLPrjParams.SaveFile(_strFileName)
                            SaveXMLFile = True
                        Else
                            SaveXMLFile = False
                            Exit Function
                        End If
                    ElseIf intvbYesNo = MsgBoxResult.No Then
                        _XMLPrjParams.SaveFile(_strFileName)
                        _booExists = True
                        SaveXMLFile = True
                    Else
                        SaveXMLFile = False
                        Exit Function
                    End If
                Else
                    _XMLPrjParams.SaveFile(_strFileName)
                    _booExists = True
                    SaveXMLFile = True
                End If
            End If

        Catch ex As Exception
            If Err.Number = 32755 Then
                SaveXMLFile = False
                Exit Function
            Else
                MsgBox(Err.Number & " " & Err.Description)
                SaveXMLFile = False
            End If
        End Try

    End Function

    Private Function ValidateData() As Boolean
        'Time to rifle through the form ensuring kosher data across the board.
        Try
            Dim strUpdatePrecip As String

            Dim clsParamsPrj As New clsXMLPrjFile 'Just a holder for the xml

            SSTab1.SelectedIndex = 0

            'First check Selected Watersheds
            If chkSelectedPolys.Enabled = True And chkSelectedPolys.CheckState = 1 Then
                If Len(cboSelectPoly.Text) > 0 Then
                    If g_MapWin.View.SelectedShapes.NumSelected > 0 Then
                        ValidateData = True
                    End If
                Else
                    MsgBox("You have chosen 'Selected watersheds only'.  Please select watersheds.", MsgBoxStyle.Critical, "No Selected Features Found")
                    ValidateData = False
                    Exit Function
                End If
            End If

            'Project Name
            If Len(txtProjectName.Text) > 0 Then
                clsParamsPrj.strProjectName = Trim(txtProjectName.Text)
            Else
                MsgBox("Please enter a name for this project.", MsgBoxStyle.Information, "Enter Name")
                txtProjectName.Focus()
                ValidateData = False
                Exit Function
            End If

            'Working Directory
            If (Len(txtOutputWS.Text) > 0) And IO.Directory.Exists(txtOutputWS.Text) Then
                clsParamsPrj.strProjectWorkspace = Trim(txtOutputWS.Text)
            Else
                MsgBox("Please choose a valid output working directory.", MsgBoxStyle.Information, "Choose Workspace")
                txtOutputWS.Focus()
                ValidateData = False
                Exit Function
            End If

            'LandCover
            If cboLCLayer.Text = "" Then
                MsgBox("Please select a Land Cover layer before continuing.", MsgBoxStyle.Information, "Select Land Cover Layer")
                cboLCLayer.Focus()
                ValidateData = False
                Exit Function
            Else
                If modUtil.LayerInMap(cboLCLayer.Text) Then
                    clsParamsPrj.strLCGridName = cboLCLayer.Text
                    clsParamsPrj.strLCGridFileName = modUtil.GetLayerFilename(cboLCLayer.Text)
                    clsParamsPrj.strLCGridUnits = CStr(cboLCUnits.SelectedIndex)
                Else
                    MsgBox("The Land Cover layer you have choosen is not in the current map frame.", MsgBoxStyle.Information, "Layer Not Found")
                    ValidateData = False
                    Exit Function
                End If
            End If

            'LC Type
            If cboLCType.Text = "" Then
                MsgBox("Please select a Land Class Type before continuing.", MsgBoxStyle.Information, "Select Land Class Type")
                cboLCType.Focus()
                ValidateData = False
                Exit Function
            Else
                clsParamsPrj.strLCGridType = cboLCType.Text
            End If

            'Soils - use definition to find datasets, if there use, if not tell the user
            If cboSoilsLayer.Text = "" Then
                MsgBox("Please select a Soils definition before continuing.", MsgBoxStyle.Information, "Select Soils Layer")
                cboSoilsLayer.Focus()
                ValidateData = False
                Exit Function
            Else
                If modUtil.RasterExists((lblSoilsHyd.Text)) Then
                    clsParamsPrj.strSoilsDefName = cboSoilsLayer.Text
                    clsParamsPrj.strSoilsHydFileName = lblSoilsHyd.Text
                Else
                    MsgBox("The hydrologic soils layer " & lblSoilsHyd.Text & " you have selected is missing.  Please check you soils definition.", MsgBoxStyle.Information, "Soils Layer Not Found")
                    ValidateData = False
                    Exit Function
                End If
            End If

            'PrecipScenario
            'If the layer is in the map, get out, all is well- _strPrecipFile is established on the
            'PrecipCbo Click event
            If modUtil.LayerInMap(modUtil.SplitFileName(_strPrecipFile)) Then
                clsParamsPrj.strPrecipScenario = cboPrecipScen.Text
            Else
                'Check if you can add it, if so, all is well
                If modUtil.RasterExists(_strPrecipFile) Then
                    clsParamsPrj.strPrecipScenario = cboPrecipScen.Text
                Else
                    'Can't find it...well, then send user to Browse
                    MsgBox("Unable to find precip dataset: " & _strPrecipFile & ".  Please Correct", MsgBoxStyle.Information, "Cannot Find Dataset")
                    _strPrecipFile = modUtil.BrowseForFileName("Raster", Me, "Browse for Precipitation Dataset...")
                    'If new one found, then we must update DataBase
                    If Len(_strPrecipFile) > 0 Then
                        strUpdatePrecip = "UPDATE PrecipScenario SET precipScenario.PrecipFileName = '" & _strPrecipFile & "'" & "WHERE NAME = '" & cboPrecipScen.Text & "'"
                        Dim PreUpdCmd As OleDbCommand = New OleDbCommand(strUpdatePrecip, g_DBConn)
                        PreUpdCmd.ExecuteNonQuery()

                        'Now we can set the xmlParams
                        clsParamsPrj.strPrecipScenario = cboPrecipScen.Text
                        'modUtil.AddRasterLayerToMapFromFileName _strPrecipFile, m_pMap
                    Else
                        MsgBox("Invalid File.", MsgBoxStyle.Information, "Invalid File")
                        cboPrecipScen.Focus()
                        ValidateData = False
                    End If
                End If
            End If

            'Go out to a separate function for this one...WaterShed
            If ValidateWaterShed() Then
                clsParamsPrj.strWaterShedDelin = cboWSDelin.Text
            Else
                MsgBox("There is a problems with the selected Watershed Delineation.", MsgBoxStyle.Information, "Watershed Delineation")
                ValidateData = False
                Exit Function
            End If

            'Water Quality
            If Len(cboWQStd.Text) > 0 Then
                clsParamsPrj.strWaterQuality = cboWQStd.Text
            Else
                MsgBox("Please select a water quality standard.", MsgBoxStyle.Information, "Water Quality Standard Missing")
                ValidateData = False
                Exit Function
            End If

            'Checkboxes, straight up values
            clsParamsPrj.intLocalEffects = chkLocalEffects.CheckState

            'Theoreretically, user could open file that had selected sheds.
            If chkSelectedPolys.Enabled = True Then
                clsParamsPrj.intSelectedPolys = chkSelectedPolys.CheckState
                clsParamsPrj.strSelectedPolyFileName = modUtil.GetLayerFilename(cboSelectPoly.Text)
                clsParamsPrj.strSelectedPolyLyrName = cboSelectPoly.Text
            Else
                clsParamsPrj.intSelectedPolys = 0
            End If

            SSTab1.SelectedIndex = 1

            'Erosion Tab
            'Calc Erosion checkbox
            clsParamsPrj.intCalcErosion = chkCalcErosion.CheckState

            If chkCalcErosion.CheckState Then
                If modUtil.RasterExists((lblKFactor.Text)) Then
                    clsParamsPrj.strSoilsKFileName = lblKFactor.Text
                Else
                    MsgBox("The K Factor soils dataset " & lblSoilsHyd.Text & " you have selected is missing.  Please check your soils definition.", MsgBoxStyle.Information, "Soils K Factor Not Found")
                    ValidateData = False
                    Exit Function
                End If

                'Check the Rainfall Factor grid objects.
                If frameRainFall.Visible = True Then

                    If optUseGRID.Checked Then

                        If Len(cboRainGrid.Text) > 0 And (InStr(1, cboRainGrid.Text, cboLCLayer.Text, 1) = 0) Then
                            clsParamsPrj.intRainGridBool = 1
                            clsParamsPrj.intRainConstBool = 0
                            clsParamsPrj.strRainGridName = cboRainGrid.Text
                            clsParamsPrj.strRainGridFileName = modUtil.GetLayerFilename(cboRainGrid.Text)
                        Else
                            MsgBox("Please choose a rainfall Grid.", MsgBoxStyle.Information, "Select Rainfall GRID")
                            SSTab1.SelectedIndex = 1
                            ValidateData = False
                            Exit Function

                        End If

                    ElseIf optUseValue.Checked Then

                        If Not IsNumeric(txtRainValue.Text) Then
                            MsgBox("Numbers Only for Rain Values.", MsgBoxStyle.Information, "Numbers Only Please")
                            txtRainValue.Focus()
                        Else
                            If CDbl(txtRainValue.Text) < 0 Then
                                MsgBox("Positive values only please for rainfall values.", MsgBoxStyle.Information, "Postive Values Only")
                                txtRainValue.Focus()
                            Else
                                clsParamsPrj.intRainConstBool = 1
                                clsParamsPrj.dblRainConstValue = CDbl(txtRainValue.Text)
                                clsParamsPrj.strRainGridFileName = ""
                            End If
                        End If

                    Else
                        MsgBox("You must choose a rainfall factor.", MsgBoxStyle.Information, "Rainfall Factor Missing")
                        ValidateData = False
                        Exit Function
                    End If
                End If

                'Soil Delivery Ratio
                'Added 12/03/07 to account for soil delivery ratio GRID, user can now provide.
                If chkSDR.CheckState = 1 Then
                    If Len(txtSDRGRID.Text) > 0 Then
                        If modUtil.RasterExists((txtSDRGRID.Text)) Then
                            clsParamsPrj.intUseOwnSDR = 1
                            clsParamsPrj.strSDRGridFileName = txtSDRGRID.Text
                        Else
                            MsgBox("SDR GRID " & txtSDRGRID.Text & " not found.", MsgBoxStyle.Information, "SDR GRID Not Found")
                            ValidateData = False
                            Exit Function
                        End If
                    Else
                        MsgBox("Please select an SDR GRID.", MsgBoxStyle.Information, "SDR GRID Not Selected")
                        ValidateData = False
                        Exit Function
                    End If
                Else
                    clsParamsPrj.intUseOwnSDR = 0
                    clsParamsPrj.strSDRGridFileName = txtSDRGRID.Text
                End If

            End If

            SSTab1.SelectedIndex = 3
            'Managment Scenarios
            If ValidateMgmtScenario() Then
                For Each row As DataGridViewRow In dgvManagementScen.Rows
                    If Len(row.Cells("ChangeAreaLayer").Value) > 0 Then
                        clsParamsPrj.clsMgmtScenItem = New clsXMLMgmtScenItem
                        clsParamsPrj.clsMgmtScenItem.intID = row.Index
                        If row.Cells("ManageApply").FormattedValue Then
                            clsParamsPrj.clsMgmtScenItem.intApply = 1
                        Else
                            clsParamsPrj.clsMgmtScenItem.intApply = 0
                        End If
                        clsParamsPrj.clsMgmtScenItem.strAreaName = row.Cells("ChangeAreaLayer").Value
                        clsParamsPrj.clsMgmtScenItem.strAreaFileName = modUtil.GetLayerFilename(row.Cells("ChangeAreaLayer").Value)
                        clsParamsPrj.clsMgmtScenItem.strChangeToClass = row.Cells("ChangeToClass").Value
                        clsParamsPrj.clsMgmtScenHolder.Add(clsParamsPrj.clsMgmtScenItem)
                    End If
                Next
            Else
                ValidateData = False
                dgvManagementScen.Focus()
                Exit Function
            End If

            SSTab1.SelectedIndex = 0
            'Pollutants
            If ValidatePollutants() Then
                For Each row As DataGridViewRow In dgvPollutants.Rows
                    'Adding a New Pollutantant Item to the Project file
                    clsParamsPrj.clsPollItem = New clsXMLPollutantItem
                    clsParamsPrj.clsPollItem.intID = row.Index
                    If row.Cells("PollApply").FormattedValue Then
                        clsParamsPrj.clsPollItem.intApply = 1
                    Else
                        clsParamsPrj.clsPollItem.intApply = 0
                    End If
                    clsParamsPrj.clsPollItem.strPollName = row.Cells("PollutantName").Value
                    clsParamsPrj.clsPollItem.strCoeffSet = row.Cells("CoefSet").Value
                    clsParamsPrj.clsPollItem.strCoeff = row.Cells("WhichCoeff").Value
                    clsParamsPrj.clsPollItem.intThreshold = CShort(row.Cells("Threshold").Value)
                    If row.Cells("TypeDef").Value <> "" Then
                        clsParamsPrj.clsPollItem.strTypeDefXMLFile = row.Cells("TypeDef").Value
                    End If
                    clsParamsPrj.clsPollItems.Add(clsParamsPrj.clsPollItem)
                Next
            Else
                ValidateData = False
                dgvPollutants.Focus()
                Exit Function
            End If

            SSTab1.SelectedIndex = 2
            'Land Uses
            For Each row As DataGridViewRow In dgvLandUse.Rows
                If Len(row.Cells("LUScenario").Value) > 0 Then
                    clsParamsPrj.clsLUItem = New clsXMLLandUseItem
                    clsParamsPrj.clsLUItem.intID = row.Index
                    clsParamsPrj.clsLUItem.intApply = CShort(row.Cells("LUApply").FormattedValue)
                    clsParamsPrj.clsLUItem.strLUScenName = row.Cells("LUScenario").Value
                    clsParamsPrj.clsLUItem.strLUScenXMLFile = row.Cells("LUScenarioXML").Value
                    clsParamsPrj.clsLUItems.Add(clsParamsPrj.clsLUItem)
                End If
            Next
            'If it gets to here, all is well
            ValidateData = True

            _XMLPrjParams = clsParamsPrj

        Catch ex As Exception
            HandleError(False, "ValidateData " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        Finally
            SSTab1.SelectedIndex = 0
        End Try
    End Function

    Private Function ValidatePollutants() As Boolean
        'Function to validate pollutants
        Try
            For Each row As DataGridViewRow In dgvPollutants.Rows
                If row.Cells("PollApply").FormattedValue = True Then
                    If Len(row.Cells("CoefSet").Value) = 0 Then
                        MsgBox("Please select a coefficient set for pollutant: " & row.Cells("PollutantName").Value.ToString, MsgBoxStyle.Critical, "Coefficient Set Missing")
                        ValidatePollutants = False
                        Exit Function
                    Else
                        If Len(row.Cells("WhichCoeff").Value) = 0 Then
                            MsgBox("Please select a coefficient for pollutant: " & row.Cells("PollutantName").Value.ToString, MsgBoxStyle.Critical, "Coefficient Missing")
                            ValidatePollutants = False
                            Exit Function
                        Else
                            ValidatePollutants = True
                        End If
                    End If
                Else
                    ValidatePollutants = True
                End If
            Next
        Catch ex As Exception
            HandleError(False, "ValidatePollutants " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Function

    Private Function ValidateWaterShed() As Boolean
        Try

            'Validate the Watershed
            Dim strWShed As String
            Dim booUpdate As Boolean

            Dim strDEM As String
            Dim strFlowDirFileName As String
            Dim strFlowAccumFileName As String
            Dim strFilledDEMFileName As String

            booUpdate = False

            'Select record from current cbo Selection
            strWShed = "SELECT * FROM WSDELINEATION WHERE NAME LIKE '" & cboWSDelin.Text & "'"
            Dim WSCmd As OleDbCommand = New OleDbCommand(strWShed, g_DBConn)
            Dim WSData As OleDbDataReader = WSCmd.ExecuteReader()
            WSData.Read()

            'Check to make sure all datasets exist, if not
            'DEM
            If Not modUtil.RasterExists(WSData.Item("DEMFileName")) Then
                MsgBox("Unable to locate DEM dataset: " & WSData.Item("DEMFileName") & ".", MsgBoxStyle.Critical, "Missing Dataset")
                strDEM = modUtil.BrowseForFileName("Raster", Me, "Browse for DEM...")
                If Len(strDEM) > 0 Then
                    'TODO write demfilename
                    'rsWShed.Fields("DEMFileName").Value = strDEM
                    booUpdate = True
                Else
                    ValidateWaterShed = False
                    Exit Function
                End If
                'WaterShed Delineation
            ElseIf Not modUtil.FeatureExists(WSData.Item("wsfilename")) Then
                MsgBox("Unable to locate Watershed dataset: " & WSData.Item("wsfilename") & ".", MsgBoxStyle.Critical, "Missing Dataset")
                strWShed = modUtil.BrowseForFileName("Feature", Me, "Browse for Watershed Dataset...")
                If Len(strWShed) > 0 Then
                    'TODO write wsfilename
                    'rsWShed.Fields("wsfilename").Value = strWShed
                    booUpdate = True
                Else
                    ValidateWaterShed = False
                    Exit Function
                End If
                'Flow Direction
            ElseIf Not modUtil.RasterExists(WSData.Item("FlowDirFileName")) Then
                MsgBox("Unable to locate Flow Direction GRID: " & WSData.Item("FlowDirFileName") & ".", MsgBoxStyle.Critical, "Missing Dataset")
                strFlowDirFileName = modUtil.BrowseForFileName("Raster", Me, "Browse for Flow Direction GRID...")
                If Len(strFlowDirFileName) > 0 Then
                    'TODO write flow dir
                    'rsWShed.Fields("FlowDirFileName").Value = strFlowDirFileName
                    booUpdate = True
                Else
                    ValidateWaterShed = False
                    Exit Function
                End If
                'Flow Accumulation
            ElseIf Not modUtil.RasterExists(WSData.Item("FlowAccumFileName")) Then
                MsgBox("Unable to locate Flow Accumulation GRID: " & WSData.Item("FlowAccumFileName") & ".", MsgBoxStyle.Critical, "Missing Dataset")
                strFlowAccumFileName = modUtil.BrowseForFileName("Raster", Me, "Browse for Flow Accumulation GRID...")
                If Len(strFlowAccumFileName) > 0 Then
                    'TODO Write flow accum
                    'rsWShed.Fields("FlowAccumFileName").Value = strFlowAccumFileName
                    booUpdate = True
                Else
                    ValidateWaterShed = False
                    Exit Function
                End If
                'Check for non-hydro correct GRIDS
            ElseIf WSData.Item("HydroCorrected") = 0 Then
                If Not modUtil.RasterExists(WSData.Item("FilledDEMFileName")) Then
                    MsgBox("Unable to locate the Filled DEM: " & WSData.Item("FilledDEMFileName") & ".", MsgBoxStyle.Critical, "Missing Dataset")
                    strFilledDEMFileName = modUtil.BrowseForFileName("Raster", Me, "Browse for Filled DEM...")
                    If Len(strFilledDEMFileName) > 0 Then
                        'TODO write filled dem
                        'rsWShed.Fields("FilledDEMFileName").Value = strFilledDEMFileName
                        booUpdate = True
                    Else
                        ValidateWaterShed = False
                        Exit Function
                    End If
                End If
            End If

            If booUpdate Then
                'TODO Figure out how to update
                'rsWShed.Update()
            End If

            ValidateWaterShed = True

            WSData.Close()
        Catch ex As Exception
            HandleError(False, "ValidateWaterShed " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Function

    Private Function ValidateMgmtScenario() As Boolean
        Try
            For Each row As DataGridViewRow In dgvManagementScen.Rows
                If row.Cells("ManageApply").FormattedValue <> False Then
                    If Len(row.Cells("ChangeAreaLayer").Value) > 0 Then
                        If Not modUtil.LayerInMap(row.Cells("ChangeAreaLayer").Value) Then
                            ValidateMgmtScenario = False
                            Exit Function
                        End If
                    End If

                    If row.Cells("ChangeToClass").Value = "" Then
                        MsgBox("Please select a land class in row " + (row.Index + 1).ToString, MsgBoxStyle.Critical, "Missing Value")
                        ValidateMgmtScenario = False
                        Exit Function
                    End If

                End If
            Next
           

            ValidateMgmtScenario = True

        Catch ex As Exception
            HandleError(False, "ValidateMgmtScenario " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Function

    Public Sub SetLURow(ByVal row As Integer, ByVal name As String, ByVal strXML As String)
        dgvLandUse.Rows(row).Cells("LUApply").Value = True
        dgvLandUse.Rows(row).Cells("LUScenario").Value = name
        dgvLandUse.Rows(row).Cells("LUScenarioXML").Value = strXML
    End Sub

    Public Sub UpdatePrecip(ByVal strPrecName As String)
        cboPrecipScen.Items.Clear()
        modUtil.InitComboBox(cboPrecipScen, "PrecipScenario")
        cboPrecipScen.Items.Insert(cboPrecipScen.Items.Count, "New precipitation scenario...")
        cboPrecipScen.SelectedIndex = modUtil.GetCboIndex(strPrecName, cboPrecipScen)
    End Sub

    Public Sub UpdateWQ(ByVal strWQName As String)
        cboWQStd.Items.Clear()
        modUtil.InitComboBox(cboWQStd, "WQCRITERIA")
        cboWQStd.Items.Insert(cboWQStd.Items.Count, "Define a new water quality standard...")
        cboWQStd.SelectedIndex = modUtil.GetCboIndex(strWQName, cboWQStd)
    End Sub
#End Region


End Class