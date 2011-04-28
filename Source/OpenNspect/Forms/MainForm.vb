'********************************************************************************************************
'File Name: frmProjectSetup.vb
'Description: Main  input and model activation form
'********************************************************************************************************
'The contents of this file are subject to the Mozilla Public License Version 1.1 (the "License"); 
'you may not use this file except in compliance with the License. You may obtain a copy of the License at 
'http://www.mozilla.org/MPL/ 
'Software distributed under the License is distributed on an "AS IS" basis, WITHOUT WARRANTY OF 
'ANY KIND, either express or implied. See the License for the specificlanguage governing rights and 
'limitations under the License. 
'
'Contributor(s): (Open source contributors should list themselves and their modifications here). 
'Oct 20, 2010:  Allen Anselmo allen.anselmo@gmail.com - 
'               Added licensing and comments to code
Imports System.Windows.Forms
Imports System.Collections.Generic
Imports System.IO
Imports System.Data
Imports MapWindow.Interfaces
Imports System.Data.OleDb
Imports MapWinGIS
Imports OpenNspect.Xml
Imports Point = System.Drawing.Point

Friend Class MainForm
    Inherits Form

#Region "Class Vars"

    Private _XmlPrjParams As ProjectFile
    ' xml doc that holds inputs

    Private _isFileSaved As Boolean
    Private _IsAnnualPrecipScenario As Boolean
    Private _currentFileName As String
    Private _strPrecipFile As String
    Private _strOpenFileName As String


    Private arrAreaList As New ArrayList
    Private arrClassList As New ArrayList

    Private _SelectLyrPath As String
    Private _SelectedShapes As List(Of Integer)
    Private _CurrentProjectPath As String
#End Region

#Region "Events"

    Private Sub GetMapWindowLayers()
        Dim currLyr As Layer
        For i As Integer = 0 To MapWindowPlugin.MapWindowInstance.Layers.NumLayers - 1
            currLyr = MapWindowPlugin.MapWindowInstance.Layers(i)
            If currLyr.LayerType = eLayerType.Grid Then
                cboLCLayer.Items.Add(currLyr.Name)
            ElseIf currLyr.LayerType = eLayerType.PolygonShapefile Then
                arrAreaList.Add(currLyr.Name)
            End If
        Next
    End Sub
    ''' <summary>
    ''' Load form that initializes globals and various form elements
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub frmProjectSetup_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        Try
            g_cbMainForm = Me
            TabsForGrids.SelectedIndex = 0

            InitComboBox(cboLandCoverType, "LCType")

            InitComboBox(cboPrecipitationScenarios, "PrecipScenario")
            cboPrecipitationScenarios.Items.Add("New precipitation scenario...")

            InitComboBox(cboWaterShedDelineations, "WSDelineation")
            cboWaterShedDelineations.Items.Add("New watershed delineation...")

            InitComboBox(cboWaterQualityCriteriaStd, "WQCriteria")
            cboWaterQualityCriteriaStd.Items.Add("New water quality standard...")

            GetMapWindowLayers()

            'Soils, now a 'scenario', not just a datalayer
            InitComboBox(cboSoilsLayer, "Soils")

            'Fill LandClass
            FillCboLCCLass()

            'Initialize parameter file
            _XmlPrjParams = New ProjectFile

            UpdateFormTitle("*")

            chkCalcErosion_CheckStateChanged(Me, Nothing)

            'Add one blank management row
            dgvManagementScen.Rows.Add()
            PopulateManagement(0)

            'Test workspace persistence
            If g_strWorkspace IsNot Nothing Then
                txtOutputWS.Text = g_strWorkspace
            End If

            txtProjectName.Focus()
        Catch ex As Exception
            HandleError(ex)
        End Try

    End Sub

    ''' <summary>
    ''' Loads the previous Xml settings file when reshown
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub frmProjectSetup_Shown(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Shown
        LoadPreviousXmlFile()
    End Sub

    Private Sub UpdateFormTitle(ByVal projectName As String)
        Me.Text = String.Format("OpenNSPECT - {0}", projectName)
    End Sub

    ''' <summary>
    ''' Changes the form title with the project name
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub txtProjectName_TextChanged(ByVal sender As Object, ByVal e As EventArgs) _
        Handles txtProjectName.TextChanged
        UpdateFormTitle(txtProjectName.Text)
    End Sub

    ''' <summary>
    ''' Button press opening a working directory
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub cmdOpenWS_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdOpenWS.Click
        Try
            'Starts by default in the workspace subdirectory of whatever the NSPECT default path is
            Dim initFolder As String = g_nspectDocPath & "\workspace"

            'Makes sure the workspace exists
            If Not Directory.Exists(initFolder) Then
                MkDir(initFolder)
            End If
            Using dlgBrowser As New FolderBrowserDialog() With
                {.Description = "Choose a directory for analysis output: ", .SelectedPath = initFolder}
                If dlgBrowser.ShowDialog = System.Windows.Forms.DialogResult.OK Then
                    txtOutputWS.Text = dlgBrowser.SelectedPath
                    g_strWorkspace = txtOutputWS.Text
                End If
            End Using
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    ''' <summary>
    ''' Changes the LC units if LC layer changes
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub cboLCLayer_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) _
        Handles cboLCLayer.SelectedIndexChanged
        Try
            cboLCUnits.SelectedIndex = GetRasterDistanceUnits(cboLCLayer.Text)
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    ''' <summary>
    ''' Fills the LC combo whena type is selected
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub cboLCType_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) _
        Handles cboLandCoverType.SelectedIndexChanged
        Try
            FillCboLCCLass()
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    ''' <summary>
    ''' When Soils layer changes, it sets the appropriate soils filenames from the DB
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub cboSoilsLayer_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) _
        Handles cboSoilsLayer.SelectedIndexChanged
        Try
            'Execute a selection query
            Dim strSelect As String = String.Format("SELECT * FROM Soils WHERE NAME LIKE '{0}'", cboSoilsLayer.Text)
            Using soilCmd As New OleDbCommand(strSelect, g_DBConn)
                Dim soilData As OleDbDataReader = soilCmd.ExecuteReader()
                soilData.Read()
                lblKFactor.Text = soilData.Item("SoilsKFileName")
                lblSoilsHyd.Text = soilData.Item("SoilsFileName")
                soilData.Close()
            End Using
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    ''' <summary>
    ''' Controlls what happens when the precipitation scenario is changed
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub cboPrecipScen_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) _
        Handles cboPrecipitationScenarios.SelectedIndexChanged
        Try
            'Have to change Erosion tab based on Annual/Event driven rain event
            Dim strEvent As String

            'If define, then open new window for new definition, else select from database
            If cboPrecipitationScenarios.Text = "New precipitation scenario..." Then
                Using newPre As New NewPrecipitationScenarioForm()
                    newPre.Init(Me, Nothing)
                    newPre.ShowDialog()
                End Using
            Else
                strEvent = String.Format("SELECT * FROM PRECIPSCENARIO WHERE NAME LIKE '{0}'", cboPrecipitationScenarios.Text)
                Using eventCmd As New OleDbCommand(strEvent, g_DBConn)
                    Dim eventData As OleDbDataReader = eventCmd.ExecuteReader()
                    eventData.Read()

                    Select Case eventData.Item("Type").ToString
                        Case "0"
                            'Annual
                            frmSDR.Visible = True
                            frameRainFall.Visible = True
                            chkCalcErosion.Text = "Calculate Erosion for Annual Type Precipitation Scenario"
                            _IsAnnualPrecipScenario = True
                            'Set flag
                        Case "1"
                            'Event
                            frmSDR.Visible = False
                            frameRainFall.Visible = False
                            chkCalcErosion.Text = "Calculate Erosion for Event Type Precipitation Scenario"
                            _IsAnnualPrecipScenario = False
                            'Set flag
                    End Select

                    _strPrecipFile = eventData.Item("PrecipFileName")
                    eventData.Close()
                End Using
            End If
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    ''' <summary>
    ''' Controls what happens when the WS delin combo is changed
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub cboWSDelin_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) _
        Handles cboWaterShedDelineations.SelectedIndexChanged
        Try
            If cboWaterShedDelineations.Text = "New watershed delineation..." Then

                g_boolNewWShed = True

                Using newWS As New NewWatershedDelineationForm()
                    newWS.Init(Nothing, Me)
                    newWS.ShowDialog()
                End Using
            End If

        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    ''' <summary>
    ''' Controls what happens when the WQ Std combo is changed
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub cboWQStd_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) _
        Handles cboWaterQualityCriteriaStd.SelectedIndexChanged
        Try
            If cboWaterQualityCriteriaStd.Text = "New water quality standard..." Then
                Using fNewWQ As New NewWaterQualityStandardForm()
                    fNewWQ.Init(Nothing, Me)
                    fNewWQ.ShowDialog()
                End Using
            Else
                PopulatePollutants()
            End If
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    ''' <summary>
    ''' New project menu click
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub mnuNew_Click(ByVal sender As Object, ByVal e As EventArgs) Handles mnuNew.Click
        Try
            Dim intvbYesNo As Short

            intvbYesNo = _
                MsgBox(String.Format("Do you want to save changes you made to {0}?", Me.Text), _
                        MsgBoxStyle.YesNoCancel + MsgBoxStyle.Exclamation, "OpenNSPECT")
            'Make sure they save current before it's lost forever if they want to
            If intvbYesNo = MsgBoxResult.Yes Then
                If SaveXmlFile() Then
                    ClearForm()
                    'Calling the load directly just reinitializes everything after it's been cleared out.
                    frmProjectSetup_Load(Me, New EventArgs())
                Else
                    Return
                End If
            ElseIf intvbYesNo = MsgBoxResult.No Then
                ClearForm()
                frmProjectSetup_Load(Me, New EventArgs())
            ElseIf intvbYesNo = MsgBoxResult.Cancel Then
                Return
            End If

        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    ''' <summary>
    ''' Opens a new xml setting file
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub mnuOpen_Click(ByVal sender As Object, ByVal e As EventArgs) Handles mnuOpen.Click
        LoadXmlFile()
    End Sub

    ''' <summary>
    ''' Saves the current Xml settings file
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub mnuSave_Click(ByVal sender As Object, ByVal e As EventArgs) Handles mnuSave.Click
        Try
            SaveXmlFile()
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    ''' <summary>
    ''' Save new xml file as a new file
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub mnuSaveAs_Click(ByVal sender As Object, ByVal e As EventArgs) Handles mnuSaveAs.Click
        Try
            _isFileSaved = False
            SaveXmlFile()
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    ''' <summary>
    ''' Exit the form
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub mnuExit_Click(ByVal sender As Object, ByVal e As EventArgs) Handles mnuExit.Click
        Try
            Dim intvbYesNo As Short

            intvbYesNo = _
                MsgBox(String.Format("Do you want to save changes you made to {0}?", Me.Text), _
                        MsgBoxStyle.YesNoCancel + MsgBoxStyle.Exclamation, "OpenNSPECT")

            If intvbYesNo = MsgBoxResult.Yes Then
                If SaveXmlFile() Then
                    Close()
                End If
            ElseIf intvbYesNo = MsgBoxResult.No Then
                Close()
            ElseIf intvbYesNo = MsgBoxResult.Cancel Then
                Return
            End If

        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    ''' <summary>
    ''' Call up the generic help window
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub mnuGeneralHelp_Click(ByVal sender As Object, ByVal e As EventArgs) _
        Handles mnuGeneralHelp.Click
        Help.ShowHelp(Me, g_nspectPath & "\Help\nspect.chm", "project_setup.htm")
    End Sub

    ''' <summary>
    ''' Handles showing the appropriate elements when the use grid checkbox is used
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub optUseGRID_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) _
        Handles optUseGRID.CheckedChanged
        otUseValue(sender)
    End Sub

    Private Sub otUseValue(ByVal sender As Object)
        If sender.Checked Then
            txtRainValue.Enabled = optUseValue.Checked
            txtbxRainGrid.Enabled = optUseGRID.Checked
        End If
    End Sub
    ''' <summary>
    ''' Handles showing the appropriate elements when the use value checkbox is used
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub optUseValue_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) _
        Handles optUseValue.CheckedChanged
        otUseValue(sender)
    End Sub

    ''' <summary>
    ''' Handles showing the appropriate elements when the SDR checkbox is used
    ''' </summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    Private Sub chkSDR_CheckStateChanged(ByVal eventSender As Object, ByVal eventArgs As EventArgs) _
        Handles chkSDR.CheckStateChanged
        Dim state As Boolean = chkSDR.CheckState = 1
        txtSDRGRID.Enabled = state
        cmdOpenSDR.Enabled = state
    End Sub

    ''' <summary>
    ''' Button to open SDR grid
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub cmdOpenSDR_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdOpenSDR.Click
        Dim g As New Grid
        Using dlgOpen As New OpenFileDialog()
            dlgOpen.Title = "Choose SDR GRID"
            dlgOpen.Filter = g.CdlgFilter
            If dlgOpen.ShowDialog = System.Windows.Forms.DialogResult.OK Then
                txtSDRGRID.Text = dlgOpen.FileName
            End If
        End Using
    End Sub

    ''' <summary>
    ''' Handles showing the appropriate elements when the calculate erosion checkbox is used
    ''' </summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    Private Sub chkCalcErosion_CheckStateChanged(ByVal eventSender As Object, _
                                                  ByVal eventArgs As EventArgs) _
        Handles chkCalcErosion.CheckStateChanged
        frameRainFall.Enabled = chkCalcErosion.CheckState
        cboErodFactor.Enabled = chkCalcErosion.CheckState
        lblErodFactor.Enabled = chkCalcErosion.CheckState
        optUseGRID.Enabled = chkCalcErosion.CheckState
        optUseValue.Enabled = True
        lblKFactor.Visible = chkCalcErosion.CheckState
        Label7.Visible = chkCalcErosion.CheckState
    End Sub

    ''' <summary>
    ''' Button used to open a rainfall grid
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnOpenRainfallFactorGrid_Click(ByVal sender As Object, ByVal e As EventArgs) _
        Handles btnOpenRainfallFactorGrid.Click
        Dim g As New Grid
        Using dlgOpen As New OpenFileDialog()
            dlgOpen.Title = "Choose Rainfall Factor GRID"
            dlgOpen.Filter = g.CdlgFilter
            If dlgOpen.ShowDialog = System.Windows.Forms.DialogResult.OK Then
                txtbxRainGrid.Text = dlgOpen.FileName
            End If
        End Using
    End Sub

    ''' <summary>
    ''' Handles errors if a value is typed into the table that doesn't match constraints
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub dgvPollutants_DataError(ByVal sender As Object, _
                                         ByVal e As DataGridViewDataErrorEventArgs) _
        Handles dgvPollutants.DataError
        MsgBox( _
                String.Format("Please enter a valid value in row {0} and column {1}.", e.RowIndex + 1, _
                               e.ColumnIndex + 1))
    End Sub

    ''' <summary>
    ''' Handles errors if a value is typed into the table that doesn't match constraints
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub dgvLandUse_DataError(ByVal sender As Object, _
                                      ByVal e As DataGridViewDataErrorEventArgs) _
        Handles dgvLandUse.DataError
        MsgBox( _
                String.Format("Please enter a valid value in row {0} and column {1}.", e.RowIndex + 1, _
                               e.ColumnIndex + 1))
    End Sub

    ''' <summary>
    ''' Handles errors if a value is typed into the table that doesn't match constraints
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub dgvManagementScen_DataError(ByVal sender As Object, _
                                             ByVal e As DataGridViewDataErrorEventArgs) _
        Handles dgvManagementScen.DataError
        MsgBox( _
                String.Format("Please enter a valid value in row {0} and column {1}.", e.RowIndex + 1, _
                               e.ColumnIndex + 1))
    End Sub

    ''' <summary>
    ''' Controls what happens when the land use table is right-clicked
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub dgvLandUse_MouseClick(ByVal sender As Object, ByVal e As MouseEventArgs) _
        Handles dgvLandUse.MouseClick
        'limit to right click
        If e.Button = System.Windows.Forms.MouseButtons.Right Then
            'show the context menu
            cnxtmnuLandUse.Show(dgvLandUse, New Point(e.X, e.Y))
            If Not dgvLandUse.CurrentRow Is Nothing Then
                'Enable or disable the edit scenario menu items, though they aren't currently used
                If _
                    dgvLandUse.CurrentRow.Cells("LUApply").FormattedValue Or _
                    dgvLandUse.CurrentRow.Cells("LUScenario").Value <> "" Then
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

    ''' <summary>
    ''' Handle right click of the Management Scenarios table
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub dgvManagementScen_MouseClick(ByVal sender As Object, _
                                              ByVal e As MouseEventArgs) _
        Handles dgvManagementScen.MouseClick
        If e.Button = System.Windows.Forms.MouseButtons.Right Then
            cnxtmnuManagement.Show(dgvManagementScen, New Point(e.X, e.Y))
        End If
    End Sub

    ''' <summary>
    ''' Used to add a scenario to the table
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub AddScenarioToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) _
        Handles AddScenarioToolStripMenuItem.Click
        Try

            Dim intRow As Short = dgvLandUse.Rows.Add()

            g_intManScenRow = intRow.ToString

            g_strLUScenFileName = ""

            'Generate the scenario form
            Using newscen As New EditLandUseScenario()
                With newscen
                    .init(cboWaterQualityCriteriaStd.Text, Me)
                    .Text = "Add Land Use Scenario"
                    'If they cancel, then remove the added
                    If .ShowDialog() = System.Windows.Forms.DialogResult.Cancel Then
                        dgvLandUse.Rows.RemoveAt(g_intManScenRow)
                    End If
                End With
            End Using

        Catch ex As Exception
            HandleError(ex)
            'False, "MnuLUAdd_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Sub

    ''' <summary>
    ''' Handles editing a scenario in the table
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub EditScenarioToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) _
        Handles EditScenarioToolStripMenuItem.Click
        Try
            If Not dgvLandUse.CurrentRow Is Nothing Then
                g_intManScenRow = dgvLandUse.CurrentRow.Index.ToString
                g_strLUScenFileName = dgvLandUse.CurrentRow.Cells("LUScenarioXml").Value

                Using newscen As New EditLandUseScenario()
                    With newscen
                        .init(cboWaterQualityCriteriaStd.Text, Me)
                        .Text = "Edit Land Use Scenario"
                        .ShowDialog()
                    End With
                End Using
            End If

        Catch ex As Exception
            HandleError(ex)
            'False, "MnuLUEdit_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try

    End Sub

    ''' <summary>
    ''' Handles removing a scenario from the table
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub DeleteScenarioToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) _
        Handles DeleteScenarioToolStripMenuItem.Click
        Try
            With dgvLandUse
                If Not .CurrentRow Is Nothing Then
                    If .CurrentRow.Cells("LUApply").FormattedValue Or .CurrentRow.Cells("LUScenario").Value <> "" Then
                        If _
                            MsgBox( _
                                    String.Format("There is data in Row {0}. Would you still like to delete it?", _
                                                   .CurrentRow.Index + 1), MsgBoxStyle.YesNo, "Delete Row") = _
                            MsgBoxResult.Yes Then
                            .Rows.Remove(.CurrentRow)
                        End If
                    Else
                        .Rows.Remove(.CurrentRow)
                    End If
                End If
            End With
        Catch ex As Exception
            HandleError(ex)
        End Try

    End Sub

    ''' <summary>
    ''' Handles appending a row to the management table
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub AppendRowToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) _
        Handles AppendRowToolStripMenuItem.Click
        Dim idx As Integer = dgvManagementScen.Rows.Add()
        PopulateManagement(idx)
    End Sub

    ''' <summary>
    ''' Handles inserting a row into the management table
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub InsertRowToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) _
        Handles InsertRowToolStripMenuItem.Click
        If Not dgvManagementScen.CurrentRow Is Nothing Then
            Dim idx As Integer = dgvManagementScen.CurrentRow.Index
            dgvManagementScen.Rows.Insert(idx, 1)
            PopulateManagement(idx)
        End If
    End Sub

    ''' <summary>
    ''' Handles deleting the currently selected row in the management table
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub DeleteCurrentRowToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) _
        Handles DeleteCurrentRowToolStripMenuItem.Click
        With dgvManagementScen
            If Not .CurrentRow Is Nothing Then
                Dim idx As Integer = .CurrentRow.Index
                If _
                    .Rows(idx).Cells("ManageApply").Value Or .Rows(idx).Cells("ChangeAreaLayer").Value <> "" Or _
                    .Rows(idx).Cells("ChangeToClass").Value <> "" Then
                    If _
                        MsgBox(String.Format("There is data in Row {0}. Would you still like to delete it?", idx + 1), _
                                MsgBoxStyle.YesNo, "Delete Row") = MsgBoxResult.Yes Then
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

    ''' <summary>
    ''' Handles opening the shape selection form
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnSelect_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSelect.Click
        Dim selectfrm As New SelectionModeForm()
        selectfrm.InitializeAndShow()

    End Sub

    ''' <summary>
    ''' Handles the close button click
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub cmdQuit_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdQuit.Click
        Try
            Dim intvbYesNo As Short

            intvbYesNo = _
                MsgBox(String.Format("Do you want to save changes you made to {0}?", Me.Text), _
                        MsgBoxStyle.YesNoCancel + MsgBoxStyle.Exclamation, "OpenNSPECT")

            'Make sure to let them save before it's lost if they choose to
            If intvbYesNo = MsgBoxResult.Yes Then
                If SaveXmlFile() Then
                    Close()
                End If
            ElseIf intvbYesNo = MsgBoxResult.No Then
                Close()
            ElseIf intvbYesNo = MsgBoxResult.Cancel Then
                Return
            End If

        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    ''' <summary>
    ''' The workhorse of NSPECT. Automates the entire process of the nspect processing
    ''' </summary>
    Private Sub RunAnalysis()
        Try
            Dim strWaterShed As String
            'Connection string
            Dim strPrecip As String
            'Connection String

            Dim AreThereLandUseItems As Boolean
            'Are there Landuse Scenarios???
            Dim dictPollutants As New Dictionary(Of String, String)
            'Dict to hold all pollutants
            Dim i As Integer

            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'STEP 1: Save file, populate xml params: -------------------------------------------------------------------------
            If Not SaveXmlFile() Then
                Return
            End If

            'Handles whether to overwrite existing groups of the same name or to generate a new group for outputs
            If g_pGroupLayer <> -1 Then
                If MapWindowPlugin.MapWindowInstance.Layers.Groups.ItemByHandle(g_pGroupLayer).Text = _XmlPrjParams.strProjectName Then
                    Dim _
                        res As MsgBoxResult = _
                            MsgBox( _
                                    String.Format("Would you like to overwrite the last results group named {0}?", _
                                                   _XmlPrjParams.strProjectName), MsgBoxStyle.YesNoCancel, _
                                    "Replace Results?")
                    If res = MsgBoxResult.Yes Then
                        MapWindowPlugin.MapWindowInstance.Layers.Groups.Remove(g_pGroupLayer)
                    ElseIf res = MsgBoxResult.Cancel Then
                        Return
                    End If
                End If
            End If

            g_pGroupLayer = MapWindowPlugin.MapWindowInstance.Layers.Groups.Add(_XmlPrjParams.strProjectName)

            'Init your global dictionary to hold the metadata records as well as the global xml prj file
            g_dicMetadata = New Dictionary(Of String, String)
            g_XmlPrjFile = _XmlPrjParams
            g_strWorkspace = g_XmlPrjFile.strProjectWorkspace
            'END STEP 1: -----------------------------------------------------------------------------------------------------

            'STEP 2: Identify if local effects are being used : --------------------------------------------------------------
            'Local Effects Global
            If _XmlPrjParams.intLocalEffects = 1 Then
                g_booLocalEffects = True
            Else
                g_booLocalEffects = False
            End If
            'END STEP 2: -----------------------------------------------------------------------------------------------------

            'STEP 3: Find out if user is making use of only the selected Sheds -----------------------------------------------
            'Selected Sheds only
            If _XmlPrjParams.intSelectedPolys = 1 Then
                g_booSelectedPolys = True
            Else
                g_booSelectedPolys = False
            End If
            'END STEP 3: ---------------------------------------------------------------------------------------------------------

            'STEP 4: Get the Management Scenarios: ------------------------------------------------------------------------------------
            'If they're using, we send them over to modMgmtScen to implement
            If _XmlPrjParams.MgmtScenHolder.Count > 0 Then
                MgmtScenSetup(_XmlPrjParams.MgmtScenHolder, _XmlPrjParams.strLCGridType, _
                               _XmlPrjParams.strLCGridFileName, _XmlPrjParams.strProjectWorkspace)
            End If
            'END STEP 4: ---------------------------------------------------------------------------------------------------------

            'STEP 5: Pollutant Dictionary creation, needed for Landuse -----------------------------------------------------------
            'Go through and find the pollutants, if they're used and what the CoeffSet is
            'We're creating a dictionary that will hold Pollutant, Coefficient Set for use in the Landuse Scenarios
            For i = 0 To _XmlPrjParams.PollItems.Count - 1
                If _XmlPrjParams.PollItems.Item(i).intApply = 1 Then
                    dictPollutants.Add(_XmlPrjParams.PollItems.Item(i).strPollName, _
                                        _XmlPrjParams.PollItems.Item(i).strCoeffSet)
                End If
            Next i
            'END STEP 5: ---------------------------------------------------------------------------------------------------------

            'STEP 6: Landuses sent off to modLanduse for processing -----------------------------------------------------
            For i = 0 To _XmlPrjParams.LUItems.Count - 1
                If _XmlPrjParams.LUItems.Item(i).intApply = 1 Then
                    AreThereLandUseItems = True
                    Begin(_XmlPrjParams.strLCGridType, _XmlPrjParams.LUItems, dictPollutants, _
                           _XmlPrjParams.strLCGridFileName)
                    Exit For
                Else
                    AreThereLandUseItems = False
                End If
            Next i
            'END STEP 6: ---------------------------------------------------------------------------------------------------------

            'STEP 7: ---------------------------------------------------------------------------------------------------------
            'Obtain Watershed values

            strWaterShed = String.Format("Select * from WSDelineation Where Name like '{0}'", _XmlPrjParams.strWaterShedDelin)
            Dim cmdWS As New DataHelper(strWaterShed)

            'END STEP 7: -----------------------------------------------------------------------------------------------------

            'STEP 8: ---------------------------------------------------------------------------------------------------------
            'Set the Analysis Environment and globals for output workspace

            SetGlobalEnvironment(cmdWS.GetCommand(), _SelectLyrPath, _SelectedShapes)

            'END STEP 8: -----------------------------------------------------------------------------------------------------

            'STEP 8a: --------------------------------------------------------------------------------------------------------
            'Added 1/08/2007 to account for non-adjacent polygons
            If _XmlPrjParams.intSelectedPolys = 1 Then
                If CheckMultiPartPolygon(g_pSelectedPolyClip) Then
                    MsgBox( _
                            "Warning: Your selected polygons are not adjacent.  Please select only polygons that are adjacent.", _
                            MsgBoxStyle.Critical, "Non-adjacent Polygons Detected")
                    Return
                End If
            End If

            'STEP 9: ---------------------------------------------------------------------------------------------------------
            'Create the runoff GRID
            'Get the precip scenario stuff
            strPrecip = _
                String.Format("Select * from PrecipScenario where name like '{0}'", _XmlPrjParams.strPrecipScenario)
            Using cmdPrecip As New DataHelper(strPrecip)
                Using dataPrecip As OleDbDataReader = cmdPrecip.ExecuteReader()
                    dataPrecip.Read()
                    'Added 6/04 to account for different PrecipTypes
                    g_intPrecipType = dataPrecip.Item("PrecipType")
                    dataPrecip.Close()
                End Using
                'If there has been a land use added, then a new LCType has been created, hence we get it from g_strLCTypename
                Dim strLCType As String
                If AreThereLandUseItems Then
                    strLCType = g_strLCTypeName
                Else
                    strLCType = _XmlPrjParams.strLCGridType
                End If
                If _
                    Not _
                    CreateRunoffGrid(_XmlPrjParams.strLCGridFileName, strLCType, cmdPrecip.GetCommand(), _
                                      _XmlPrjParams.strSoilsHydFileName, _XmlPrjParams.OutputItems) Then
                    Return
                End If
            End Using
            'END STEP 9: -----------------------------------------------------------------------------------------------------

            'STEP 10: ---------------------------------------------------------------------------------------------------------
            'Process pollutants
            For i = 0 To _XmlPrjParams.PollItems.Count - 1
                If _XmlPrjParams.PollItems.Item(i).intApply = 1 Then
                    'If user is NOT ignoring the pollutant then send the whole item over along with LCType
                    Dim pollitem As PollutantItem = _XmlPrjParams.PollItems.Item(i)
                    If _
                        Not _
                        PollutantConcentrationSetup(pollitem, _XmlPrjParams.strLCGridType, _
                                                     _XmlPrjParams.strWaterQuality, _
                                                     _XmlPrjParams.OutputItems) Then
                        Return
                    End If
                End If
            Next i
            'END STEP 10: -----------------------------------------------------------------------------------------------------

            'Step 11: Erosion -------------------------------------------------------------------------------------------------
            'Check that they have chosen Erosion
            Dim dataWS As OleDbDataReader = cmdWS.ExecuteReader
            dataWS.Read()
            If _XmlPrjParams.intCalcErosion = 1 Then
                If _IsAnnualPrecipScenario Then 'If Annual (0) then TRUE, ergo RUSLE
                    If _XmlPrjParams.intRainGridBool Then
                        If _
                            Not _
                            RUSLESetup(dataWS.Item("NibbleFileName"), dataWS.Item("dem2bfilename"), _
                                        _XmlPrjParams.strRainGridFileName, _XmlPrjParams.strSoilsKFileName, _
                                        _XmlPrjParams.strSDRGridFileName, _XmlPrjParams.strLCGridType, _
                                        _XmlPrjParams.OutputItems) Then
                            Return
                        End If
                    ElseIf _XmlPrjParams.intRainConstBool Then
                        If _
                            Not _
                            RUSLESetup(dataWS.Item("NibbleFileName"), dataWS.Item("dem2bfilename"), _
                                        _XmlPrjParams.strRainGridFileName, _XmlPrjParams.strSoilsKFileName, _
                                        _XmlPrjParams.strSDRGridFileName, _XmlPrjParams.strLCGridType, _
                                        _XmlPrjParams.OutputItems, _XmlPrjParams.dblRainConstValue) Then
                            Return
                        End If
                    End If
                Else 'If event (1) then False, ergo MUSLE
                    If _
                        Not _
                        MUSLESetup(_XmlPrjParams.strSoilsDefName, _XmlPrjParams.strSoilsKFileName, _
                                    _XmlPrjParams.strLCGridType, _XmlPrjParams.OutputItems) Then
                        Return
                    End If
                End If
            End If
            dataWS.Close()
            'STEP 11: ----------------------------------------------------------------------------------------------------------

            'STEP 12 : Cleanup any temp critters -------------------------------------------------------------------------------
            'g_DictTempNames holds the names of all temporary landuses and/or coefficient sets created during the Landuse scenario
            'portion of our program, for example CCAP1, or NitSet1.  We now must eliminate them from the database if they exist.
            If g_DictTempNames.Count > 0 Then
                If AreThereLandUseItems Then
                    Cleanup(g_DictTempNames, (_XmlPrjParams.PollItems), (_XmlPrjParams.strLCGridType))
                End If
            End If
            'END STEP 12: -------------------------------------------------------------------------------------------------------

            'STEP 15: create string describing project parameters ---------------------------------------------------------------
            'TODO metadata stuff
            'strProjectInfo = modUtil.ParseProjectforMetadata(_XmlPrjParams, _strFileName)
            'END STEP 15: -------------------------------------------------------------------------------------------------------

            'STEP 16: Apply the metadata to each of the rasters in the group layer ----------------------------------------------
            'TODO
            'm_App.StatusBar.Message(0) = "Creating metadata for the OpenNSPECT group layer..."
            'modUtil.CreateMetadata(g_pGroupLayer, strProjectInfo)
            'END STEP 16: -------------------------------------------------------------------------------------------------------

            'Cleanup ------------------------------------------------------------------------------------------------------------
            'Go into workspace and rid it of all rasters
            CleanGlobals()

            MapWindowPlugin.MapWindowInstance.StatusBar.ProgressBarValue = 0

            'Save xml to ensure outputs are saved
            _XmlPrjParams.SaveFile(_currentFileName)

            Close()

            Return
        Catch ex As Exception
            HandleError(ex)
        Finally
            CloseProgressDialog()
            System.Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub cmdRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdRun.Click
        RunAnalysis()
    End Sub

#End Region

#Region "Helper Functions"

    ''' <summary>
    ''' Used by the selection form to set the selected shape
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub SetSelectedShape()
        'Uses the current layer and cycles the select shapes, populating a list of shape index values
        If MapWindowPlugin.MapWindowInstance.Layers.CurrentLayer <> -1 And MapWindowPlugin.MapWindowInstance.View.SelectedShapes.NumSelected > 0 Then
            chkSelectedPolys.Checked = True
            _SelectLyrPath = MapWindowPlugin.MapWindowInstance.Layers(MapWindowPlugin.MapWindowInstance.Layers.CurrentLayer).FileName
            _SelectedShapes = New List(Of Integer)
            For i As Integer = 0 To MapWindowPlugin.MapWindowInstance.View.SelectedShapes.NumSelected - 1
                _SelectedShapes.Add(MapWindowPlugin.MapWindowInstance.View.SelectedShapes(i).ShapeIndex)
            Next
            lblSelected.Text = MapWindowPlugin.MapWindowInstance.View.SelectedShapes.NumSelected.ToString + " selected"
        End If
    End Sub

    ''' <summary>
    ''' Sets the form to a default state
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ClearForm()
        Try
            'Gotta clean up before new, clean form

            'LandClass stuff
            cboLCLayer.Items.Clear()
            cboLandCoverType.Items.Clear()

            'DBase scens
            cboPrecipitationScenarios.Items.Clear()
            cboWaterShedDelineations.Items.Clear()
            cboWaterQualityCriteriaStd.Items.Clear()
            cboSoilsLayer.Items.Clear()

            'Text
            txtProjectName.Text = ""
            txtOutputWS.Text = ""

            'Checkboxes
            chkSelectedPolys.CheckState = CheckState.Unchecked
            chkLocalEffects.CheckState = CheckState.Unchecked

            'Erosion
            optUseGRID.Checked = True
            chkCalcErosion.CheckState = CheckState.Unchecked
            txtRainValue.Text = ""

            'clear the grids
            dgvPollutants.Rows.Clear()
            dgvLandUse.Rows.Clear()
            dgvManagementScen.Rows.Clear()

            frmProjectSetup_Load(Nothing, Nothing)
        Catch ex As Exception
            HandleError(ex)
            'False, "ClearForm " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Sub

    ''' <summary>
    ''' Prompts for an Xml project file to load
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub LoadXmlFile()
        Try

            'browse...get output filename
            Dim initialDirectory As String = g_nspectDocPath & "\projects"
            If Not Directory.Exists(initialDirectory) Then
                MkDir(initialDirectory)
            End If
            Using dlgXmlOpen As New OpenFileDialog()
                With dlgXmlOpen
                    .Filter = MSG8XmlFile
                    .InitialDirectory = initialDirectory
                    .Title = "Open OpenNSPECT Project File"
                    .FilterIndex = 1
                    .ShowDialog()
                End With

                If Len(dlgXmlOpen.FileName) > 0 Then
                    _currentFileName = Trim(dlgXmlOpen.FileName)
                    _CurrentProjectPath = _currentFileName
                    'Xml Class autopopulates when passed a file
                    _XmlPrjParams.Xml = _currentFileName
                    'Populate from the local Xml params
                    FillForm()
                Else
                    Return
                End If
            End Using

            'Pop this string with the incoming name, if they change, we'll prompt to 'save as'.
            _strOpenFileName = txtProjectName.Text

        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    ''' <summary>
    ''' Loads the previously set Xml file and populates the form from it
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub LoadPreviousXmlFile()
        Try
            If Len(_CurrentProjectPath) > 0 Then
                _currentFileName = Trim(_CurrentProjectPath)
                _XmlPrjParams.Xml = _currentFileName
                FillForm()
            Else
                Return
            End If

            'Pop this string with the incoming name, if they change, we'll prompt to 'save as'.
            _strOpenFileName = txtProjectName.Text

        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    ''' <summary>
    ''' Fill the LC class combo from the DB
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub FillCboLCCLass()
        Try
            Dim strLCChanges As String
            Dim i As Short

            strLCChanges = _
                String.Format( _
                               "SELECT LCCLASS.Name as Name2, LCTYPE.LCTYPEID FROM LCTYPE INNER JOIN LCCLASS ON LCTYPE.LCTYPEID = LCCLASS.LCTYPEID WHERE LCTYPE.NAME LIKE '{0}'", _
                               cboLandCoverType.Text)
            Using lcChangesCmd As New OleDbCommand(strLCChanges, g_DBConn)
                Using lcChanges As OleDbDataReader = lcChangesCmd.ExecuteReader()
                    arrClassList.Clear()
                    Do While lcChanges.Read()
                        arrClassList.Insert(i, lcChanges.Item("Name2"))
                    Loop
                    lcChanges.Close()
                End Using
            End Using

        Catch ex As Exception
            HandleError(ex)
            'False, "FillCboLCCLass " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Sub

    ''' <summary>
    ''' Populates the pollutants table
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub PopulatePollutants()
        Try
            Dim strSQLWQStd As String

            Dim strSQLWQStdPoll As String

            'Selection based on combo box
            strSQLWQStd = String.Format("SELECT * FROM WQCRITERIA WHERE NAME LIKE '{0}'", cboWaterQualityCriteriaStd.Text)
            Using WQCritCmd As OleDbCommand = New OleDbCommand(strSQLWQStd, g_DBConn)
                Dim dataWQCrit As OleDbDataReader = WQCritCmd.ExecuteReader()

                If dataWQCrit.HasRows Then
                    dataWQCrit.Read()
                    strSQLWQStdPoll = _
                        String.Format( _
                                       "SELECT POLLUTANT.NAME, POLL_WQCRITERIA.THRESHOLD FROM POLL_WQCRITERIA INNER JOIN POLLUTANT ON POLL_WQCRITERIA.POLLID = POLLUTANT.POLLID Where POLL_WQCRITERIA.WQCRITID = {0}", _
                                       dataWQCrit.Item("WQCRITID"))
                    Using WQStdCmd As OleDbCommand = New OleDbCommand(strSQLWQStdPoll, g_DBConn)
                        Using WQStdAdapter As OleDbDataAdapter = New OleDbDataAdapter(WQStdCmd)
                            Using PollutantsTable As New DataTable()
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
                            End Using
                        End Using
                    End Using
                Else
                    MsgBox("Warning: There are no water quality standards remaining.  Please add a new one.", _
                            MsgBoxStyle.Critical, "Recordset Empty")
                End If
            End Using

        Catch ex As Exception
            HandleError(ex)
            'False, "PopPollutants " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Sub

    ''' <summary>
    ''' Populate the coefficients table
    ''' </summary>
    ''' <param name="strPollutantName"></param>
    ''' <param name="rowidx"></param>
    ''' <remarks></remarks>
    Private Sub PopulateCoefType(ByVal strPollutantName As String, ByVal rowidx As Integer)
        Try
            Dim _
                strSelectCoeff As String = _
                    String.Format( _
                                   "SELECT POLLUTANT.POLLID, POLLUTANT.NAME, COEFFICIENTSET.NAME AS NAME2 FROM POLLUTANT INNER JOIN COEFFICIENTSET ON POLLUTANT.POLLID = COEFFICIENTSET.POLLID Where POLLUTANT.NAME LIKE '{0}'", _
                                   strPollutantName)
            Using coefCmd As New OleDbCommand(strSelectCoeff, g_DBConn)
                Dim coefData As OleDbDataReader = coefCmd.ExecuteReader()
                Dim cell As DataGridViewComboBoxCell = dgvPollutants.Rows(rowidx).Cells("CoefSet")
                While coefData.Read
                    cell.Items.Add(coefData.Item("Name2"))
                End While
                coefData.Close()
            End Using
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    ''' <summary>
    ''' Populate the Management table
    ''' </summary>
    ''' <param name="rowidx"></param>
    ''' <remarks></remarks>
    Private Sub PopulateManagement(ByVal rowidx As Integer)
        Try
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
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub AddBasinpolyToMap()
        Dim _
            strCurrWShed As String = _
                String.Format("Select * from WSDelineation where Name Like '{0}'", _XmlPrjParams.strWaterShedDelin)
        ' TODO: remove lngCurrWshedPolyIndex as it is not used.
        Dim lngCurrWshedPolyIndex As Integer
        Using WSCmd As OleDbCommand = New OleDbCommand(strCurrWShed, g_DBConn)
            Using wsData As OleDbDataReader = WSCmd.ExecuteReader
                wsData.Read()
                Dim strBasin As String = ""
                If wsData.HasRows() Then
                    strBasin = wsData.Item("wsfilename")
                    If Path.GetExtension(strBasin) <> ".shp" Then
                        strBasin = strBasin + ".shp"
                    End If
                    If Not LayerInMapByFileName(strBasin) Then
                        If _
                            AddFeatureLayerToMapFromFileName(strBasin, _
                                                              String.Format("{0} Drainage Basins", _
                                                                             wsData.Item("Name"))) Then
                            lngCurrWshedPolyIndex = _
                                GetLayerIndex(String.Format("{0} Drainage Basins", wsData.Item("Name")))
                            arrAreaList.Add(String.Format("{0} Drainage Basins", wsData.Item("Name")))
                        Else
                            MsgBox( _
                                    String.Format( _
                                                   "Could not find watershed layer: {0} .  Please add the watershed layer to the map.", _
                                                   wsData.Item("wsfilename")), MsgBoxStyle.Critical, "File Not Found")
                        End If
                    End If
                Else
                    _XmlPrjParams.strWaterShedDelin = cboWaterShedDelineations.Text
                    strCurrWShed = _
                        String.Format("Select * from WSDelineation where Name Like '{0}'", _
                                       _XmlPrjParams.strWaterShedDelin)
                    Using WSCmd2 = New OleDbCommand(strCurrWShed, g_DBConn)
                        Dim wsData2 = WSCmd2.ExecuteReader()
                        wsData2.Read()
                        If wsData2.HasRows() Then
                            strBasin = wsData2.Item("wsfilename")
                            If Path.GetExtension(strBasin) <> ".shp" Then
                                strBasin = strBasin + ".shp"
                            End If
                            If Not LayerInMapByFileName(strBasin) Then
                                Dim basinName As String = String.Format("{0} Drainage Basins", wsData2.Item("Name"))
                                If AddFeatureLayerToMapFromFileName(strBasin, basinName) Then
                                    lngCurrWshedPolyIndex = GetLayerIndex(basinName)
                                    arrAreaList.Add(String.Format("{0} Drainage Basins", wsData2.Item("Name")))
                                Else
                                    MsgBox( _
                                            String.Format( _
                                                           "Could not find watershed layer: {0} .  Please add the watershed layer to the map.", _
                                                           wsData.Item("wsfilename")), MsgBoxStyle.Critical, _
                                            "File Not Found")
                                End If
                            End If
                        End If
                    End Using
                End If
            End Using
        End Using
    End Sub

    Private Function GetPollutantsIdx() As Integer
        Dim _intPollCount As Short
        'count for pollutant grid
        _intPollCount = _XmlPrjParams.PollItems.Count
        Dim idx As Integer
        Dim strPollName As String
        If _intPollCount > 0 Then
            dgvPollutants.Rows.Clear()

            For i = 0 To _intPollCount - 1
                With dgvPollutants
                    idx = .Rows.Add()

                    .Rows(idx).Cells("PollApply").Value = _XmlPrjParams.PollItems.Item(i).intApply
                    .Rows(idx).Cells("PollutantName").Value = _XmlPrjParams.PollItems.Item(i).strPollName
                    strPollName = _XmlPrjParams.PollItems.Item(i).strPollName
                    PopulateCoefType(strPollName, idx)
                    .Rows(idx).Cells("CoefSet").Value = _XmlPrjParams.PollItems.Item(i).strCoeffSet

                    .Rows(idx).Cells("WhichCoeff").Value = _XmlPrjParams.PollItems.Item(i).strCoeff
                    .Rows(idx).Cells("Threshold").Value = CStr(_XmlPrjParams.PollItems.Item(i).intThreshold)

                    If Len(_XmlPrjParams.PollItems.Item(i).strTypeDefXmlFile) > 0 Then
                        .Rows(idx).Cells("TypeDef").Value = _
                            CStr(_XmlPrjParams.PollItems.Item(i).strTypeDefXmlFile)
                    End If
                End With
            Next i

        End If
        Return idx
    End Function

    Private Function LoadLandCoverGrid() As Boolean

        'Check to see if the LC cover is in the map, if so, set the combobox
        If LayerLoadedInMap(_XmlPrjParams.strLCGridName) Then
            cboLCLayer.SelectedIndex = GetCboIndex((_XmlPrjParams.strLCGridName), cboLCLayer)
            cboLCLayer.Refresh()
        Else
            If AddRasterLayerToMapFromFileName(_XmlPrjParams.strLCGridFileName) Then

                With cboLCLayer
                    .Items.Add(_XmlPrjParams.strLCGridName)
                    .Refresh()
                    .SelectedIndex = GetCboIndex(_XmlPrjParams.strLCGridName, cboLCLayer)
                End With

            Else
                Dim browseForDataSet As Short = MsgBox(String.Format( _
                                                               "Could not find the Land Cover dataset: {0}.  Would you like to browse for it?", _
                                                               _XmlPrjParams.strLCGridFileName), MsgBoxStyle.YesNo, _
                                                "Cannot Locate Dataset")
                If browseForDataSet = MsgBoxResult.Yes Then
                    _XmlPrjParams.strLCGridFileName = AddInputFromGxBrowser("Raster")
                    If _XmlPrjParams.strLCGridFileName <> "" Then
                        If AddRasterLayerToMapFromFileName(_XmlPrjParams.strLCGridFileName) Then
                            cboLCLayer.SelectedIndex = 0
                        End If
                    Else
                        Return False
                    End If
                Else
                    Return False
                End If
            End If

        End If

        cboLCUnits.SelectedIndex = GetCboIndex((_XmlPrjParams.strLCGridUnits), cboLCUnits)
        cboLCUnits.Refresh()

        cboLandCoverType.SelectedIndex = GetCboIndex((_XmlPrjParams.strLCGridType), cboLandCoverType)
        cboLandCoverType.Refresh()

        Return True
    End Function
    Private Function LoadSoils() As Boolean
        If RasterExists(_XmlPrjParams.strSoilsHydFileName) Then
            cboSoilsLayer.SelectedIndex = GetCboIndex((_XmlPrjParams.strSoilsDefName), cboSoilsLayer)
            cboSoilsLayer.Refresh()
            Return True
        Else
            MsgBox("Could not find soils dataset.  Please correct the soils definition in the Advanced Settings.", _
                    MsgBoxStyle.Critical, "Dataset Missing")
            Return False
        End If
    End Function
    Private Sub LoadWatersheds()
        chkLocalEffects.CheckState = _XmlPrjParams.intLocalEffects
        chkSelectedPolys.CheckState = _XmlPrjParams.intSelectedPolys

        If chkSelectedPolys.CheckState = 1 Then
            '1st see if it's in the map
            Dim strSelected As String = _XmlPrjParams.strSelectedPolyFileName
            If Path.GetExtension(strSelected) <> ".shp" Then
                strSelected = strSelected + ".shp"
            End If
            _SelectLyrPath = strSelected
            _SelectedShapes = _XmlPrjParams.intSelectedPolyList
            lblSelected.Text = _SelectedShapes.Count.ToString + " selected"

            If Not LayerInMapByFileName(strSelected) Then
                'Not there then add it
                If AddFeatureLayerToMapFromFileName(strSelected, _XmlPrjParams.strSelectedPolyLyrName) Then
                    arrAreaList.Add(_XmlPrjParams.strSelectedPolyLyrName)
                Else
                    'Can't find it, then send em searching
                    Dim browseForData = _
                        MsgBox( _
                                String.Format( _
                                               "Could not find the Selected Polygons file used to limit extent: {0}.  Would you like to browse for it? ", _
                                               strSelected), MsgBoxStyle.YesNo, "Cannot Locate Dataset")
                    If browseForData = MsgBoxResult.Yes Then
                        'if they want to look for it then give em the browser
                        _XmlPrjParams.strSelectedPolyFileName = AddInputFromGxBrowser("Feature")
                        'if they actually find something, throw it in the map
                        If Len(_XmlPrjParams.strSelectedPolyFileName) > 0 Then
                            If AddFeatureLayerToMapFromFileName(_XmlPrjParams.strSelectedPolyFileName) Then
                                arrAreaList.Add(SplitFileName(_XmlPrjParams.strSelectedPolyFileName))

                            End If
                        End If
                    Else
                        _XmlPrjParams.intSelectedPolys = 0
                        chkSelectedPolys.CheckState = CheckState.Unchecked
                    End If
                End If
            End If
        End If
    End Sub
    Private Function LoadErosion() As Boolean
        If _XmlPrjParams.intCalcErosion = 1 Then
            chkCalcErosion.CheckState = CheckState.Checked
        Else
            chkCalcErosion.CheckState = CheckState.Unchecked
        End If

        'Step: Erosion Tab - Precip
        'Either they use the GRID
        optUseGRID.Checked = _XmlPrjParams.intRainGridBool
        If optUseGRID.Checked Then
            If RasterExists(_XmlPrjParams.strRainGridFileName) Then
                Dim browseForData = _
                      MsgBox( _
                              String.Format("Could not find Rainfall GRID: {0}.  Would you like to browse for it?", _
                                             _XmlPrjParams.strRainGridFileName), MsgBoxStyle.YesNo, _
                              "Cannot Locate Dataset")
                If browseForData = MsgBoxResult.Yes Then
                    Dim g As New Grid
                    Using dlgOpen As New OpenFileDialog()
                        dlgOpen.Title = "Choose Rainfall Factor GRID"
                        dlgOpen.Filter = g.CdlgFilter
                        If dlgOpen.ShowDialog = System.Windows.Forms.DialogResult.OK Then
                            txtbxRainGrid.Text = dlgOpen.FileName
                            _XmlPrjParams.strRainGridFileName = dlgOpen.FileName
                        End If
                    End Using
                Else
                    Return False
                End If
            Else
                txtbxRainGrid.Text = _XmlPrjParams.strRainGridFileName
            End If
        End If

        'Or they use a constant value
        optUseValue.Checked = _XmlPrjParams.intRainConstBool

        If optUseValue.Checked Then
            txtRainValue.Text = CStr(_XmlPrjParams.dblRainConstValue)
        End If

        'SDR GRID
        'If Not _XmlPrjParams.intUseOwnSDR Is Nothing Then
        If _XmlPrjParams.intUseOwnSDR = 1 Then
            chkSDR.CheckState = CheckState.Checked
            txtSDRGRID.Text = _XmlPrjParams.strLCGridFileName
        Else
            chkSDR.CheckState = CheckState.Unchecked
            txtSDRGRID.Text = _XmlPrjParams.strSDRGridFileName
        End If
        'End If

        Return True
    End Function
    Private Sub LoadLandUses(ByRef idx As Integer)
        Dim landUsesCount As Short = _XmlPrjParams.LUItems.Count

        If landUsesCount > 0 Then
            dgvLandUse.Rows.Clear()
            For i = 0 To landUsesCount - 1
                With dgvLandUse
                    idx = .Rows.Add()
                    .Rows(idx).Cells("LUApply").Value = _XmlPrjParams.LUItems.Item(i).intApply
                    .Rows(idx).Cells("LUScenario").Value = _XmlPrjParams.LUItems.Item(i).strLUScenName
                    .Rows(idx).Cells("LUScenarioXml").Value = _XmlPrjParams.LUItems.Item(i).strLUScenXmlFile
                End With
            Next i
        End If
    End Sub
    Private Function LoadManagementScenarios(ByVal idx As Integer) As Boolean
        Dim scenarioCount As Short = _XmlPrjParams.MgmtScenHolder.Count

        If scenarioCount > 0 Then
            dgvManagementScen.Rows.Clear()
            For i = 0 To scenarioCount - 1
                With dgvManagementScen
                    idx = .Rows.Add()

                    If LayerLoadedInMap(_XmlPrjParams.MgmtScenHolder.Item(i).strAreaName) Then
                        PopulateManagement(idx)
                        .Rows(idx).Cells("ManageApply").Value = _XmlPrjParams.MgmtScenHolder.Item(i).intApply
                        .Rows(idx).Cells("ChangeAreaLayer").Value = _
                            _XmlPrjParams.MgmtScenHolder.Item(i).strAreaName
                        .Rows(idx).Cells("ChangeToClass").Value = _
                            _XmlPrjParams.MgmtScenHolder.Item(i).strChangeToClass
                    Else
                        If _
                            AddFeatureLayerToMapFromFileName( _
                                                              _XmlPrjParams.MgmtScenHolder.Item(i). _
                                                                 strAreaFileName) Then
                            arrAreaList.Add(_XmlPrjParams.MgmtScenHolder.Item(i).strAreaName)

                            PopulateManagement(idx)
                            .Rows(idx).Cells("ManageApply").Value = _
                                _XmlPrjParams.MgmtScenHolder.Item(i).intApply
                            .Rows(idx).Cells("ChangeAreaLayer").Value = _
                                _XmlPrjParams.MgmtScenHolder.Item(i).strAreaName
                            .Rows(idx).Cells("ChangeToClass").Value = _
                                _XmlPrjParams.MgmtScenHolder.Item(i).strChangeToClass
                        Else
                            Dim browseForData = MsgBox( _
                                        String.Format( _
                                                       "Could not find Management Sceario Area Layer: {0}.  Would you like to browse for it?", _
                                                       _XmlPrjParams.MgmtScenHolder.Item(i).strAreaFileName), _
                                        MsgBoxStyle.YesNo, _
                                        "Cannot Locate Dataset:" & _
                                        _XmlPrjParams.MgmtScenHolder.Item(i).strAreaFileName)
                            If browseForData = MsgBoxResult.Yes Then
                                _XmlPrjParams.MgmtScenHolder.Item(i).strAreaFileName = _
                                    AddInputFromGxBrowser("Feature")
                                If _
                                    AddFeatureLayerToMapFromFileName( _
                                                                      _XmlPrjParams.MgmtScenHolder.Item( _
                                                                                                            i) _
                                                                         .strAreaFileName) Then
                                    arrAreaList.Add(_XmlPrjParams.MgmtScenHolder.Item(i).strAreaName)

                                    PopulateManagement(idx)
                                    .Rows(idx).Cells("ManageApply").Value = _
                                        _XmlPrjParams.MgmtScenHolder.Item(i).intApply
                                    .Rows(idx).Cells("ChangeAreaLayer").Value = _
                                        _XmlPrjParams.MgmtScenHolder.Item(i).strAreaName
                                    .Rows(idx).Cells("ChangeToClass").Value = _
                                        _XmlPrjParams.MgmtScenHolder.Item(i).strChangeToClass
                                End If
                            Else
                                Return False
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

        Return True
    End Function
    ''' <summary>
    ''' Populates the form from the currently loaded xml project
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub FillForm()
        Try
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor

            txtProjectName.Text = _XmlPrjParams.strProjectName
            txtOutputWS.Text = _XmlPrjParams.strProjectWorkspace

            If Not LoadLandCoverGrid() Then Return
            If Not LoadSoils() Then Return
            cboPrecipitationScenarios.SelectedIndex = GetCboIndex((_XmlPrjParams.strPrecipScenario), cboPrecipitationScenarios)
            cboWaterShedDelineations.SelectedIndex = GetCboIndex((_XmlPrjParams.strWaterShedDelin), cboWaterShedDelineations)
            AddBasinpolyToMap()
            cboWaterQualityCriteriaStd.SelectedIndex = GetCboIndex((_XmlPrjParams.strWaterQuality), cboWaterQualityCriteriaStd)
            LoadWatersheds()
            If Not LoadErosion() Then Return
            Dim idx As Integer = GetPollutantsIdx()
            LoadLandUses(idx)
            LoadManagementScenarios(idx)

            'Reset to first tab
            TabsForGrids.SelectedIndex = 0
            _isFileSaved = True
            System.Windows.Forms.Cursor.Current = Cursors.Default

        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    ''' <summary>
    ''' Validates the present form settings and saves them to an Xml file
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function SaveXmlFile() As Boolean
        Try
            Dim strFolder As String
            Dim intvbYesNo As Short
            Using dlgXml As New SaveFileDialog()
                strFolder = g_nspectDocPath & "\projects"
                If Not Directory.Exists(strFolder) Then
                    MkDir(strFolder)
                End If
                If Not ValidateData() Then 'check form inputs
                    SaveXmlFile = False
                    Exit Function
                End If
                'If it does not already exist, open Save As... dialog
                If Not _isFileSaved Then
                    With dlgXml
                        .Filter = MSG8XmlFile
                        .Title = "Save Project File As..."
                        .InitialDirectory = strFolder
                        .FileName = txtProjectName.Text
                        .ShowDialog()
                    End With
                    'check to make sure filename length is greater than zeros
                    If Len(dlgXml.FileName) > 0 Then
                        _currentFileName = Trim(dlgXml.FileName)
                        _isFileSaved = True
                        _XmlPrjParams.SaveFile(_currentFileName)
                        SaveXmlFile = True
                    Else
                        SaveXmlFile = False
                        Exit Function
                    End If
                Else
                    'Now check to see if the name changed
                    If _strOpenFileName <> txtProjectName.Text Then
                        intvbYesNo = _
                            MsgBox( _
                                    String.Format( _
                                                   "You have changed the name of this project.  Would you like to save your settings as a new file?{0}{1}Yes{1} -    Save as new OpenNSPECT project file{0}{1}No{1} -    Save changes to current OpenNSPECT project file{0}{1}Cancel{1} -    Return to the project window", _
                                                   vbNewLine, vbTab), MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, _
                                    "OpenNSPECT")
                        If intvbYesNo = MsgBoxResult.Yes Then
                            With dlgXml
                                .Filter = MSG8XmlFile
                                .Title = "Save Project File As..."
                                .InitialDirectory = strFolder
                                .FilterIndex = 1
                                .FileName = txtProjectName.Text
                                .ShowDialog()
                            End With
                            'check to make sure filename length is greater than zeros
                            If Len(dlgXml.FileName) > 0 Then
                                _currentFileName = Trim(dlgXml.FileName)
                                _isFileSaved = True
                                _XmlPrjParams.SaveFile(_currentFileName)
                                SaveXmlFile = True
                            Else
                                SaveXmlFile = False
                                Exit Function
                            End If
                        ElseIf intvbYesNo = MsgBoxResult.No Then
                            _XmlPrjParams.SaveFile(_currentFileName)
                            _isFileSaved = True
                            SaveXmlFile = True
                        Else
                            SaveXmlFile = False
                            Exit Function
                        End If
                    Else
                        _XmlPrjParams.SaveFile(_currentFileName)
                        _isFileSaved = True
                        SaveXmlFile = True
                    End If
                End If
            End Using

        Catch ex As Exception
            If Err.Number = 32755 Then
                SaveXmlFile = False
                Exit Function
            Else
                HandleError(ex)

                SaveXmlFile = False
            End If
        End Try

    End Function

    ''' <summary>
    ''' Validates the form data
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ValidateData() As Boolean
        'Time to rifle through the form ensuring kosher data across the board.
        Try
            Dim strUpdatePrecip As String

            Dim ParamsPrj As New ProjectFile
            'Just a holder for the xml

            TabsForGrids.SelectedIndex = 0

            'First check Selected Watersheds
            If chkSelectedPolys.Enabled = True And chkSelectedPolys.CheckState = 1 Then
                If Len(_SelectLyrPath) > 0 Then
                    If _SelectedShapes.Count > 0 Then
                        ValidateData = True
                    End If
                Else
                    MsgBox("You have chosen 'Selected watersheds only'.  Please select watersheds.", _
                            MsgBoxStyle.Critical, "No Selected Features Found")
                    ValidateData = False
                    Exit Function
                End If
            End If

            'Project Name
            If Len(txtProjectName.Text) > 0 Then
                ParamsPrj.strProjectName = Trim(txtProjectName.Text)
            Else
                MsgBox("Please enter a name for this project.", MsgBoxStyle.Information, "Enter Name")
                txtProjectName.Focus()
                ValidateData = False
                Exit Function
            End If

            'Working Directory
            If (Len(txtOutputWS.Text) > 0) And Directory.Exists(txtOutputWS.Text) Then
                ParamsPrj.strProjectWorkspace = Trim(txtOutputWS.Text)
            Else
                MsgBox("Please choose a valid output working directory.", MsgBoxStyle.Information, "Choose Workspace")
                txtOutputWS.Focus()
                ValidateData = False
                Exit Function
            End If

            'LandCover
            If cboLCLayer.Text = "" Then
                MsgBox("Please select a Land Cover layer before continuing.", MsgBoxStyle.Information, _
                        "Select Land Cover Layer")
                cboLCLayer.Focus()
                ValidateData = False
                Exit Function
            Else
                If LayerLoadedInMap(cboLCLayer.Text) Then
                    ParamsPrj.strLCGridName = cboLCLayer.Text
                    ParamsPrj.strLCGridFileName = GetLayerFilename(cboLCLayer.Text)
                    ParamsPrj.strLCGridUnits = CStr(cboLCUnits.SelectedIndex)
                Else
                    MsgBox("The Land Cover layer you have choosen is not in the current map frame.", _
                            MsgBoxStyle.Information, "Layer Not Found")
                    ValidateData = False
                    Exit Function
                End If
            End If

            'LC Type
            If cboLandCoverType.Text = "" Then
                MsgBox("Please select a Land Class Type before continuing.", MsgBoxStyle.Information, _
                        "Select Land Class Type")
                cboLandCoverType.Focus()
                ValidateData = False
                Exit Function
            Else
                ParamsPrj.strLCGridType = cboLandCoverType.Text
            End If

            'Soils - use definition to find datasets, if there use, if not tell the user
            If cboSoilsLayer.Text = "" Then
                MsgBox("Please select a Soils definition before continuing.", MsgBoxStyle.Information, _
                        "Select Soils Layer")
                cboSoilsLayer.Focus()
                ValidateData = False
                Exit Function
            Else
                If RasterExists((lblSoilsHyd.Text)) Then
                    ParamsPrj.strSoilsDefName = cboSoilsLayer.Text
                    ParamsPrj.strSoilsHydFileName = lblSoilsHyd.Text
                Else
                    MsgBox( _
                            String.Format( _
                                           "The hydrologic soils layer {0} you have selected is missing.  Please check you soils definition.", _
                                           lblSoilsHyd.Text), MsgBoxStyle.Information, "Soils Layer Not Found")
                    ValidateData = False
                    Exit Function
                End If
            End If

            'PrecipScenario
            'If the layer is in the map, get out, all is well- _strPrecipFile is established on the
            'PrecipCbo Click event
            If LayerLoadedInMap(SplitFileName(_strPrecipFile)) Then
                ParamsPrj.strPrecipScenario = cboPrecipitationScenarios.Text
            Else
                'Check if you can add it, if so, all is well
                If RasterExists(_strPrecipFile) Then
                    ParamsPrj.strPrecipScenario = cboPrecipitationScenarios.Text
                Else
                    'Can't find it...well, then send user to Browse
                    MsgBox(String.Format("Unable to find precip dataset: {0}.  Please Correct", _strPrecipFile), _
                            MsgBoxStyle.Information, "Cannot Find Dataset")
                    _strPrecipFile = BrowseForFileName("Raster")
                    'If new one found, then we must update DataBase
                    If Len(_strPrecipFile) > 0 Then
                        strUpdatePrecip = _
                            String.Format( _
                                           "UPDATE PrecipScenario SET precipScenario.PrecipFileName = '{0}'WHERE NAME = '{1}'", _
                                           _strPrecipFile, cboPrecipitationScenarios.Text)
                        Using PreUpdCmd As OleDbCommand = New OleDbCommand(strUpdatePrecip, g_DBConn)
                            PreUpdCmd.ExecuteNonQuery()
                        End Using

                        'Now we can set the xmlParams
                        ParamsPrj.strPrecipScenario = cboPrecipitationScenarios.Text
                        'modUtil.AddRasterLayerToMapFromFileName _strPrecipFile, m_pMap
                    Else
                        MsgBox("Invalid File.", MsgBoxStyle.Information, "Invalid File")
                        cboPrecipitationScenarios.Focus()
                        ValidateData = False
                    End If
                End If
            End If

            'Go out to a separate function for this one...WaterShed
            If ValidateWaterShed() Then
                ParamsPrj.strWaterShedDelin = cboWaterShedDelineations.Text
            Else
                MsgBox("There is a problems with the selected Watershed Delineation.", MsgBoxStyle.Information, _
                        "Watershed Delineation")
                ValidateData = False
                Exit Function
            End If

            'Water Quality
            If Len(cboWaterQualityCriteriaStd.Text) > 0 Then
                ParamsPrj.strWaterQuality = cboWaterQualityCriteriaStd.Text
            Else
                MsgBox("Please select a water quality standard.", MsgBoxStyle.Information, _
                        "Water Quality Standard Missing")
                ValidateData = False
                Exit Function
            End If

            'Checkboxes, straight up values
            ParamsPrj.intLocalEffects = chkLocalEffects.CheckState

            'Theoreretically, user could open file that had selected sheds.
            If chkSelectedPolys.Checked = True Then
                ParamsPrj.intSelectedPolys = chkSelectedPolys.CheckState
                ParamsPrj.strSelectedPolyFileName = _SelectLyrPath
                Dim tmpidx As Integer = GetLayerIndexByFilename(_SelectLyrPath)
                If tmpidx <> -1 Then
                    ParamsPrj.strSelectedPolyLyrName = MapWindowPlugin.MapWindowInstance.Layers(tmpidx).Name
                Else
                    ParamsPrj.strSelectedPolyLyrName = ""
                End If
                ParamsPrj.intSelectedPolyList = _SelectedShapes
            Else
                ParamsPrj.intSelectedPolys = 0
            End If

            TabsForGrids.SelectedIndex = 1

            'Erosion Tab
            'Calc Erosion checkbox
            ParamsPrj.intCalcErosion = chkCalcErosion.CheckState

            If chkCalcErosion.CheckState Then
                If RasterExists((lblKFactor.Text)) Then
                    ParamsPrj.strSoilsKFileName = lblKFactor.Text
                Else
                    MsgBox( _
                            String.Format( _
                                           "The K Factor soils dataset {0} you have selected is missing.  Please check your soils definition.", _
                                           lblSoilsHyd.Text), MsgBoxStyle.Information, "Soils K Factor Not Found")
                    ValidateData = False
                    Exit Function
                End If

                'Check the Rainfall Factor grid objects.
                If frameRainFall.Visible = True Then

                    If optUseGRID.Checked Then

                        If Len(txtbxRainGrid.Text) > 0 And (InStr(1, txtbxRainGrid.Text, cboLCLayer.Text, 1) = 0) Then
                            ParamsPrj.intRainGridBool = 1
                            ParamsPrj.intRainConstBool = 0
                            ParamsPrj.strRainGridName = txtbxRainGrid.Text
                            ParamsPrj.strRainGridFileName = txtbxRainGrid.Text
                        Else
                            MsgBox("Please choose a rainfall Grid.", MsgBoxStyle.Information, "Select Rainfall GRID")
                            TabsForGrids.SelectedIndex = 1
                            ValidateData = False
                            Exit Function

                        End If

                    ElseIf optUseValue.Checked Then

                        If Not IsNumeric(txtRainValue.Text) Then
                            MsgBox("Numbers Only for Rain Values.", MsgBoxStyle.Information, "Numbers Only Please")
                            txtRainValue.Focus()
                        Else
                            If CDbl(txtRainValue.Text) < 0 Then
                                MsgBox("Positive values only please for rainfall values.", MsgBoxStyle.Information, _
                                        "Postive Values Only")
                                txtRainValue.Focus()
                            Else
                                ParamsPrj.intRainConstBool = 1
                                ParamsPrj.dblRainConstValue = CDbl(txtRainValue.Text)
                                ParamsPrj.strRainGridFileName = ""
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
                        If RasterExists((txtSDRGRID.Text)) Then
                            ParamsPrj.intUseOwnSDR = 1
                            ParamsPrj.strSDRGridFileName = txtSDRGRID.Text
                        Else
                            MsgBox(String.Format("SDR GRID {0} not found.", txtSDRGRID.Text), MsgBoxStyle.Information, _
                                    "SDR GRID Not Found")
                            ValidateData = False
                            Exit Function
                        End If
                    Else
                        MsgBox("Please select an SDR GRID.", MsgBoxStyle.Information, "SDR GRID Not Selected")
                        ValidateData = False
                        Exit Function
                    End If
                Else
                    ParamsPrj.intUseOwnSDR = 0
                    ParamsPrj.strSDRGridFileName = txtSDRGRID.Text
                End If

            End If

            TabsForGrids.SelectedIndex = 3
            'Managment Scenarios
            If ValidateMgmtScenario() Then
                Dim mgmtitem As ManagementScenarioItem
                For Each row As DataGridViewRow In dgvManagementScen.Rows
                    If Len(row.Cells("ChangeAreaLayer").Value) > 0 Then
                        mgmtitem = New ManagementScenarioItem
                        mgmtitem.intID = row.Index + 1
                        If row.Cells("ManageApply").FormattedValue Then
                            mgmtitem.intApply = 1
                        Else
                            mgmtitem.intApply = 0
                        End If
                        mgmtitem.strAreaName = row.Cells("ChangeAreaLayer").Value
                        mgmtitem.strAreaFileName = GetLayerFilename(row.Cells("ChangeAreaLayer").Value)
                        mgmtitem.strChangeToClass = row.Cells("ChangeToClass").Value
                        ParamsPrj.MgmtScenHolder.Add(mgmtitem)
                    End If
                Next
            Else
                ValidateData = False
                dgvManagementScen.Focus()
                Exit Function
            End If

            TabsForGrids.SelectedIndex = 0
            'Pollutants
            If ValidatePollutants() Then
                Dim pollitem As PollutantItem
                For Each row As DataGridViewRow In dgvPollutants.Rows
                    'Adding a New Pollutantant Item to the Project file
                    pollitem = New PollutantItem
                    pollitem.intID = row.Index + 1
                    If row.Cells("PollApply").FormattedValue Then
                        pollitem.intApply = 1
                    Else
                        pollitem.intApply = 0
                    End If
                    pollitem.strPollName = row.Cells("PollutantName").Value
                    pollitem.strCoeffSet = row.Cells("CoefSet").Value
                    pollitem.strCoeff = row.Cells("WhichCoeff").Value
                    pollitem.intThreshold = CShort(row.Cells("Threshold").Value)
                    If row.Cells("TypeDef").Value <> "" Then
                        pollitem.strTypeDefXmlFile = row.Cells("TypeDef").Value
                    End If
                    ParamsPrj.PollItems.Add(pollitem)
                Next
            Else
                ValidateData = False
                dgvPollutants.Focus()
                Exit Function
            End If

            TabsForGrids.SelectedIndex = 2
            'Land Uses
            For Each row As DataGridViewRow In dgvLandUse.Rows
                Dim luitem As LandUseItem
                If Len(row.Cells("LUScenario").Value) > 0 Then
                    luitem = New LandUseItem
                    luitem.intID = row.Index + 1
                    luitem.intApply = CShort(row.Cells("LUApply").FormattedValue)
                    luitem.strLUScenName = row.Cells("LUScenario").Value
                    luitem.strLUScenXmlFile = row.Cells("LUScenarioXml").Value
                    ParamsPrj.LUItems.Add(luitem)
                End If
            Next
            'If it gets to here, all is well
            ValidateData = True

            _XmlPrjParams = ParamsPrj

        Catch ex As Exception
            HandleError(ex)
            'False, "ValidateData " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        Finally
            TabsForGrids.SelectedIndex = 0
        End Try
    End Function

    ''' <summary>
    ''' Validate the pollutant table values
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ValidatePollutants() As Boolean
        'Function to validate pollutants
        Try
            For Each row As DataGridViewRow In dgvPollutants.Rows
                If row.Cells("PollApply").FormattedValue = True Then
                    If Len(row.Cells("CoefSet").Value) = 0 Then
                        MsgBox( _
                                "Please select a coefficient set for pollutant: " & _
                                row.Cells("PollutantName").Value.ToString, MsgBoxStyle.Critical, _
                                "Coefficient Set Missing")
                        ValidatePollutants = False
                        Exit Function
                    Else
                        If Len(row.Cells("WhichCoeff").Value) = 0 Then
                            MsgBox( _
                                    "Please select a coefficient for pollutant: " & _
                                    row.Cells("PollutantName").Value.ToString, MsgBoxStyle.Critical, _
                                    "Coefficient Missing")
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
            HandleError(ex)
            'False, "ValidatePollutants " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Function

    ''' <summary>
    ''' Validate the Watershed data
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ValidateWaterShed() As Boolean
        Try
            Dim strWShed As String
            Dim booUpdate As Boolean

            Dim strDEM As String
            Dim strFlowDirFileName As String
            Dim strFlowAccumFileName As String
            Dim strFilledDEMFileName As String

            booUpdate = False

            'Select record from current cbo Selection
            strWShed = String.Format("SELECT * FROM WSDELINEATION WHERE NAME LIKE '{0}'", cboWaterShedDelineations.Text)
            Using WSCmd As OleDbCommand = New OleDbCommand(strWShed, g_DBConn)
                Using adWS As New OleDbDataAdapter(WSCmd)
                    Using buWS As New OleDbCommandBuilder(adWS)
                        buWS.QuotePrefix = "["
                        buWS.QuoteSuffix = "]"
                    End Using
                    Using dtWS As New DataTable()
                        adWS.Fill(dtWS)
                        'Check to make sure all datasets exist, if not
                        'DEM
                        If Not RasterExists(dtWS.Rows(0)("DEMFileName")) Then
                            MsgBox(String.Format("Unable to locate DEM dataset: {0}.", dtWS.Rows(0)("DEMFileName")), _
                                    MsgBoxStyle.Critical, "Missing Dataset")
                            strDEM = BrowseForFileName("Raster")
                            If Len(strDEM) > 0 Then
                                dtWS.Rows(0)("DEMFileName") = strDEM
                                'rsWShed.Fields("DEMFileName").Value = strDEM
                                booUpdate = True
                            Else
                                ValidateWaterShed = False
                                Exit Function
                            End If
                            'WaterShed Delineation
                        ElseIf Not FeatureExists(dtWS.Rows(0)("wsfilename")) Then
                            MsgBox( _
                                    String.Format("Unable to locate Watershed dataset: {0}.", _
                                                   dtWS.Rows(0)("wsfilename")), MsgBoxStyle.Critical, _
                                    "Missing Dataset")
                            strWShed = BrowseForFileName("Feature")
                            If Len(strWShed) > 0 Then
                                dtWS.Rows(0)("wsfilename") = strWShed
                                booUpdate = True
                            Else
                                ValidateWaterShed = False
                                Exit Function
                            End If
                            'Flow Direction
                        ElseIf Not RasterExists(dtWS.Rows(0)("FlowDirFileName")) Then
                            MsgBox( _
                                    String.Format("Unable to locate Flow Direction GRID: {0}.", _
                                                   dtWS.Rows(0)("FlowDirFileName")), MsgBoxStyle.Critical, _
                                    "Missing Dataset")
                            strFlowDirFileName = _
                                BrowseForFileName("Raster")
                            If Len(strFlowDirFileName) > 0 Then
                                dtWS.Rows(0)("FlowDirFileName") = strFlowDirFileName
                                booUpdate = True
                            Else
                                ValidateWaterShed = False
                                Exit Function
                            End If
                            'Flow Accumulation
                        ElseIf Not RasterExists(dtWS.Rows(0)("FlowAccumFileName")) Then
                            MsgBox( _
                                    String.Format("Unable to locate Flow Accumulation GRID: {0}.", _
                                                   dtWS.Rows(0)("FlowAccumFileName")), MsgBoxStyle.Critical, _
                                    "Missing Dataset")
                            strFlowAccumFileName = _
                                BrowseForFileName("Raster")
                            If Len(strFlowAccumFileName) > 0 Then
                                dtWS.Rows(0)("FlowAccumFileName") = strFlowAccumFileName
                                booUpdate = True
                            Else
                                ValidateWaterShed = False
                                Exit Function
                            End If
                            'Check for non-hydro correct GRIDS
                        ElseIf dtWS.Rows(0)("HydroCorrected") = 0 Then
                            If Not RasterExists(dtWS.Rows(0)("FilledDEMFileName")) Then
                                MsgBox( _
                                        String.Format("Unable to locate the Filled DEM: {0}.", _
                                                       dtWS.Rows(0)("FilledDEMFileName")), MsgBoxStyle.Critical, _
                                        "Missing Dataset")
                                strFilledDEMFileName = _
                                    BrowseForFileName("Raster")
                                If Len(strFilledDEMFileName) > 0 Then
                                    dtWS.Rows(0)("FilledDEMFileName") = strFilledDEMFileName
                                    booUpdate = True
                                Else
                                    ValidateWaterShed = False
                                    Exit Function
                                End If
                            End If
                        End If
                        If booUpdate Then
                            adWS.Update(dtWS)
                        End If
                    End Using
                End Using
            End Using

            ValidateWaterShed = True
        Catch ex As Exception
            HandleError(ex)
            'False, "ValidateWaterShed " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Function

    ''' <summary>
    ''' Validate the management scenario table
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ValidateMgmtScenario() As Boolean
        Try
            For Each row As DataGridViewRow In dgvManagementScen.Rows
                If row.Cells("ManageApply").FormattedValue <> False Then
                    If Len(row.Cells("ChangeAreaLayer").Value) > 0 Then
                        If Not LayerLoadedInMap(row.Cells("ChangeAreaLayer").Value) Then
                            ValidateMgmtScenario = False
                            Exit Function
                        End If
                    End If

                    If row.Cells("ChangeToClass").Value = "" Then
                        MsgBox("Please select a land class in row " + (row.Index + 1).ToString, MsgBoxStyle.Critical, _
                                "Missing Value")
                        ValidateMgmtScenario = False
                        Exit Function
                    End If

                End If
            Next

            ValidateMgmtScenario = True

        Catch ex As Exception
            HandleError(ex)
            'False, "ValidateMgmtScenario " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Function

    ''' <summary>
    ''' used to set the Land use row values
    ''' </summary>
    ''' <param name="row"></param>
    ''' <param name="name"></param>
    ''' <param name="strXml"></param>
    ''' <remarks></remarks>
    Public Sub SetLURow(ByVal row As Integer, ByVal name As String, ByVal strXml As String)
        Try
            dgvLandUse.Rows(row).Cells("LUApply").Value = True
            dgvLandUse.Rows(row).Cells("LUScenario").Value = name
            dgvLandUse.Rows(row).Cells("LUScenarioXml").Value = strXml
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    ''' <summary>
    ''' updates the preciptiation combo after changes from outside forms
    ''' </summary>
    ''' <param name="strPrecName"></param>
    ''' <remarks></remarks>
    Public Sub UpdatePrecip(ByVal strPrecName As String)
        Try
            cboPrecipitationScenarios.Items.Clear()
            InitComboBox(cboPrecipitationScenarios, "PrecipScenario")
            cboPrecipitationScenarios.Items.Insert(cboPrecipitationScenarios.Items.Count, "New precipitation scenario...")
            cboPrecipitationScenarios.SelectedIndex = GetCboIndex(strPrecName, cboPrecipitationScenarios)
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    ''' <summary>
    ''' Updates the WQ combo after changes from outside forms
    ''' </summary>
    ''' <param name="strWQName"></param>
    ''' <remarks></remarks>
    Public Sub UpdateWQ(ByVal strWQName As String)
        Try
            cboWaterQualityCriteriaStd.Items.Clear()
            InitComboBox(cboWaterQualityCriteriaStd, "WQCRITERIA")
            cboWaterQualityCriteriaStd.Items.Insert(cboWaterQualityCriteriaStd.Items.Count, "Define a new water quality standard...")
            cboWaterQualityCriteriaStd.SelectedIndex = GetCboIndex(strWQName, cboWaterQualityCriteriaStd)
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

#End Region
End Class