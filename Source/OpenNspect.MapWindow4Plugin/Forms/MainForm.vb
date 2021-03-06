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
Imports MapWinGeoProc

Friend Class MainForm
    Inherits Form

#Region "Class Vars"

    Private _isFileOnDisk As Boolean
    Private _IsAnnualPrecipScenario As Boolean
    Private _currentFileName As String
    Private _soilsFileName As Object
    Private _strPrecipFile As String
    Private _strOpenFileName As String


    Private arrAreaList As New ArrayList
    Private arrClassList As New ArrayList

    Private _SelectLyrPath As String
    Private _SelectedShapes As List(Of Integer)
    Private _CurrentProjectPath As String
#End Region

#Region "Events"

    Private Sub ReloadTargetLayers()
        cboTargetLayer.Items.Clear()
        Dim layer As Layer
        For i As Integer = 0 To MapWindowPlugin.MapWindowInstance.Layers.NumLayers - 1
            layer = MapWindowPlugin.MapWindowInstance.Layers(i)
            If layer IsNot Nothing Then
                If layer.LayerType = eLayerType.PolygonShapefile Then
                    cboTargetLayer.Items.Add(layer.Name)
                End If
            End If

        Next
        cboTargetLayer.SelectedIndex = GetIndexOfEntry(MapWindowPlugin.MapWindowInstance.Layers(MapWindowPlugin.MapWindowInstance.Layers.CurrentLayer).Name, cboTargetLayer)
    End Sub
    Private Sub GetMapWindowLayers()
        Dim layer As Layer
        For i As Integer = 0 To MapWindowPlugin.MapWindowInstance.Layers.NumLayers - 1
            layer = MapWindowPlugin.MapWindowInstance.Layers(i)
            If layer IsNot Nothing Then
                If layer.LayerType = eLayerType.Grid Then
                    cboLCLayer.Items.Add(layer.Name)
                ElseIf layer.LayerType = eLayerType.PolygonShapefile Then
                    arrAreaList.Add(layer.Name)
                    cboTargetLayer.Items.Add(layer.Name)
                End If
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
            g_Project = New ProjectFile

            UpdateFormTitle("*")

            chkCalcErosion_CheckStateChanged(Me, Nothing)

            'Add one blank management row
            dgvManagementScen.Rows.Add()
            PopulateManagement(0)

            PopulatePollutants()

            'Test workspace persistence
            If g_Project IsNot Nothing AndAlso g_Project.ProjectWorkspace IsNot Nothing Then
                txtOutputWS.Text = g_Project.ProjectWorkspace
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
    Private Sub txtProjectName_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles txtProjectName.TextChanged
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
                    g_Project.ProjectWorkspace = txtOutputWS.Text
                End If
            End Using
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
    Private Sub cboLCType_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles cboLandCoverType.SelectedIndexChanged
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
    Private Sub cboSoilsLayer_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles cboSoilsLayer.SelectedIndexChanged
        Try
            'Execute a selection query
            Dim strSelect As String = String.Format("SELECT * FROM Soils WHERE NAME LIKE '{0}'", cboSoilsLayer.Text)
            Using soilCmd As New OleDbCommand(strSelect, g_DBConn)
                Dim soilData As OleDbDataReader = soilCmd.ExecuteReader()
                soilData.Read()
                lblKFactor.Text = soilData.Item("SoilsKFileName")
                _soilsFileName = soilData.Item("SoilsFileName")
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
    Private Sub cboPrecipScen_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles cboPrecipitationScenarios.SelectedIndexChanged
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
    Private Sub cboWSDelin_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles cboWaterShedDelineations.SelectedIndexChanged
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
    Private Sub cboWQStd_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles cboWaterQualityCriteriaStd.SelectedIndexChanged
        Try
            If cboWaterQualityCriteriaStd.Text = "New water quality standard..." Then
                Using fNewWQ As New NewWaterQualityStandardForm()
                    fNewWQ.Init(Nothing, Me)
                    fNewWQ.ShowDialog()
                End Using
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

            intvbYesNo = MsgBox(String.Format("Do you want to save changes you made to {0}?", Me.Text), MsgBoxStyle.YesNoCancel + MsgBoxStyle.Exclamation, "OpenNSPECT")
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
            _isFileOnDisk = False
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

            intvbYesNo = MsgBox(String.Format("Do you want to save changes you made to {0}?", Me.Text), MsgBoxStyle.YesNoCancel + MsgBoxStyle.Exclamation, "OpenNSPECT")

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
    Private Sub mnuGeneralHelp_Click(ByVal sender As Object, ByVal e As EventArgs) Handles mnuGeneralHelp.Click
        Help.ShowHelp(Me, g_nspectPath & "\Help\nspect.chm", "project_setup.htm")
    End Sub

    ''' <summary>
    ''' Handles showing the appropriate elements when the use grid checkbox is used
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub optUseGRID_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles optUseGRID.CheckedChanged
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
    Private Sub optUseValue_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles optUseValue.CheckedChanged
        otUseValue(sender)
    End Sub

    ''' <summary>
    ''' Handles showing the appropriate elements when the SDR checkbox is used
    ''' </summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    Private Sub chkSDR_CheckStateChanged(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles chkSDR.CheckStateChanged
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
    Private Sub chkCalcErosion_CheckStateChanged(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles chkCalcErosion.CheckStateChanged
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
    Private Sub btnOpenRainfallFactorGrid_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnOpenRainfallFactorGrid.Click
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
    Private Sub dgvLandUse_DataError(ByVal sender As Object, ByVal e As DataGridViewDataErrorEventArgs) Handles dgvLandUse.DataError
        MsgBox(String.Format("Please enter a valid value in row {0} and column {1}.", e.RowIndex + 1, e.ColumnIndex + 1))
    End Sub

    ''' <summary>
    ''' Handles errors if a value is typed into the table that doesn't match constraints
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub dgvManagementScen_DataError(ByVal sender As Object, ByVal e As DataGridViewDataErrorEventArgs) Handles dgvManagementScen.DataError
        MsgBox(String.Format("Please enter a valid value in row {0} and column {1}.", e.RowIndex + 1, e.ColumnIndex + 1))
    End Sub

    ''' <summary>
    ''' Controls what happens when the land use table is right-clicked
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub dgvLandUse_MouseClick(ByVal sender As Object, ByVal e As MouseEventArgs) Handles dgvLandUse.MouseClick
        'limit to right click
        If e.Button = System.Windows.Forms.MouseButtons.Right Then
            If Not dgvLandUse.CurrentRow Is Nothing Then
                'Enable or disable the edit scenario menu items, though they aren't currently used
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


    ''' <summary>
    ''' Used to add a scenario to the table
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub AddScenarioToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles AddScenarioToolStripMenuItem.Click
        Try
            Dim intRow As Short = dgvLandUse.Rows.Add()
            g_ManagementScenarioRowNumber = intRow.ToString
            g_ManagementScenarioLUScenFileName = ""

            'Generate the scenario form
            Using newscen As New EditLandUseScenario()
                With newscen
                    .init(cboWaterQualityCriteriaStd.Text, Me)
                    .Text = "Add Land Use Scenario"
                    'If they cancel, then remove the added
                    If .ShowDialog() = System.Windows.Forms.DialogResult.Cancel Then
                        dgvLandUse.Rows.RemoveAt(g_ManagementScenarioRowNumber)
                    End If
                End With
            End Using

        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    ''' <summary>
    ''' Handles editing a scenario in the table
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub EditScenarioToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles EditScenarioToolStripMenuItem.Click
        Try
            If Not dgvLandUse.CurrentRow Is Nothing Then
                g_ManagementScenarioRowNumber = dgvLandUse.CurrentRow.Index.ToString
                g_ManagementScenarioLUScenFileName = dgvLandUse.CurrentRow.Cells("LUScenarioXml").Value

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
        End Try

    End Sub

    ''' <summary>
    ''' Handles removing a scenario from the table
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub DeleteScenarioToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles DeleteScenarioToolStripMenuItem.Click
        Try
            With dgvLandUse
                If Not .CurrentRow Is Nothing Then
                    If .CurrentRow.Cells("LUApply").FormattedValue Or .CurrentRow.Cells("LUScenario").Value <> "" Then
                        If MsgBox(String.Format("There is data in Row {0}. Would you still like to delete it?", .CurrentRow.Index + 1), MsgBoxStyle.YesNo, "Delete Row") = MsgBoxResult.Yes Then
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
    Private Sub AppendRowToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles AppendRowToolStripMenuItem.Click
        Dim idx As Integer = dgvManagementScen.Rows.Add()
        PopulateManagement(idx)
    End Sub

    ''' <summary>
    ''' Handles inserting a row into the management table
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub InsertRowToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles InsertRowToolStripMenuItem.Click
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
    Private Sub DeleteCurrentRowToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles DeleteCurrentRowToolStripMenuItem.Click
        With dgvManagementScen
            If Not .CurrentRow Is Nothing Then
                Dim idx As Integer = .CurrentRow.Index
                If .Rows(idx).Cells("ManageApply").Value Or .Rows(idx).Cells("ChangeAreaLayer").Value <> "" Or .Rows(idx).Cells("ChangeToClass").Value <> "" Then
                    If MsgBox(String.Format("There is data in Row {0}. Would you still like to delete it?", idx + 1), MsgBoxStyle.YesNo, "Delete Row") = MsgBoxResult.Yes Then
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

            intvbYesNo = MsgBox(String.Format("Do you want to save changes you made to {0}?", Me.Text), MsgBoxStyle.YesNoCancel + MsgBoxStyle.Exclamation, "OpenNSPECT")

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
#Region "Pollutants"

    ''' <summary>
    ''' Gets a list of the pollutents to apply (name, coeff)
    ''' Go through and find the pollutants, if they're used and what the CoeffSet is
    ''' We're creating a dictionary that will hold Pollutant, Coefficient Set for use in the Landuse Scenarios
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetPollutants() As Dictionary(Of String, String)
        Dim dictPollutants As New Dictionary(Of String, String)
        Dim i As Integer
        For i = 0 To g_Project.PollItems.Count - 1
            If g_Project.PollItems.Item(i).intApply = 1 Then
                dictPollutants.Add(g_Project.PollItems.Item(i).strPollName, g_Project.PollItems.Item(i).strCoeffSet)
            End If
        Next i
        Return dictPollutants
    End Function

    'Private Sub dgvPollutants_MouseClick(sender As Object, e As MouseEventArgs) Handles dgvPollutants.MouseClick
    '    Dim sfName As String = ""
    '    Dim fieldName As String = ""
    '    With dgvPollutants
    '        If (.CurrentCellAddress.X = 3) Then
    '            If .CurrentCell.Value = "Use shapefile..." Then
    '                Dim row As Integer = .CurrentCellAddress.Y
    '                Dim pollName As String = .Rows(row).Cells(1).Value.ToString
    '                Debug.WriteLine("Pollutant is  " & pollName)
    '                Dim getShapeinfo As New PollCoeffSelectionShapefileForm(pollName, sfName, fieldName)
    '                'getShapeinfo = New PollCoeffSelectionShapefileForm(pollName)
    '                'Using getShapeinfo
    '                getShapeinfo.ShowDialog()
    '                If (getShapeinfo.DialogResult = Windows.Forms.DialogResult.OK) Then
    '                    Dim sf As String = getShapeinfo.indexShapefileName
    '                    ' .Rows(row).Cells(4).Value = sf
    '                    Dim sfld As String = getShapeinfo.indexField
    '                    '.Rows(row).Cells(5).Value = sfld
    '                    MsgBox("On MainForm, " & pollName & " shapefile is " & getShapeinfo.indexShapefileName & " and index field is " & getShapeinfo.indexField)
    '                End If

    '                'End Using

    '            End If
    '        End If

    '    End With

    'End Sub
    '   Private Sub dgvPollutants_SelectionChanged(sender As Object, e As EventArgs) Handles dgvPollutants.SelectionChanged
    'Dim sfName As String = ""
    'Dim fieldName As String = ""
    'With dgvPollutants
    '    If (.CurrentCellAddress.X = 3) Then
    '        If .CurrentCell.Value = "Use shapefile..." Then
    '            Dim row As Integer = .CurrentCellAddress.Y
    '            Dim pollName As String = .Rows(row).Cells(1).Value.ToString
    '            Debug.WriteLine("Pollutant is  " & pollName)
    '            Dim getShapeinfo As New PollCoeffSelectionShapefileForm(pollName, sfName, fieldName)
    '            'getShapeinfo = New PollCoeffSelectionShapefileForm(pollName)
    '            'Using getShapeinfo
    '            getShapeinfo.ShowDialog()
    '            If (getShapeinfo.DialogResult = Windows.Forms.DialogResult.OK) Then
    '                Dim sf As String = getShapeinfo.indexShapefileName
    '                ' .Rows(row).Cells(4).Value = sf
    '                Dim sfld As String = getShapeinfo.indexField
    '                '.Rows(row).Cells(5).Value = sfld
    '                MsgBox("On MainForm, " & pollName & " shapefile is " & getShapeinfo.indexShapefileName & " and index field is " & getShapeinfo.indexField)
    '            End If

    '            'End Using

    '        End If
    '    End If

    'End With

    '  End Sub

    'Private Sub dgvPollutants_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvPollutants.CellValueChanged
    '    Dim sfName As String = ""
    '    Dim fieldName As String = ""
    '    With dgvPollutants
    '        If (.CurrentCellAddress.X = 3) Then
    '            If .CurrentCell.Value = "Use shapefile..." Then
    '                Dim row As Integer = e.RowIndex
    '                Dim pollName As String = .Rows(e.RowIndex).Cells(1).Value.ToString
    '                Debug.WriteLine("Pollutant is  " & pollName)
    '                Dim getShapeinfo As New PollCoeffSelectionShapefileForm(pollName, sfName, fieldName)
    '                'getShapeinfo = New PollCoeffSelectionShapefileForm(pollName)
    '                'Using getShapeinfo
    '                getShapeinfo.ShowDialog()
    '                If (getShapeinfo.DialogResult = Windows.Forms.DialogResult.OK) Then
    '                    Dim sf As String = getShapeinfo.indexShapefileName
    '                    ' .Rows(row).Cells(4).Value = sf
    '                    Dim sfld As String = getShapeinfo.indexField
    '                    '.Rows(row).Cells(5).Value = sfld
    '                    MsgBox("On MainForm, " & pollName & " shapefile is " & getShapeinfo.indexShapefileName & " and index field is " & getShapeinfo.indexField)
    '                End If

    '                'End Using

    '            End If
    '        End If

    '    End With

    'End Sub

    ''' <summary>
    ''' Handles errors if a value is typed into the table that doesn't match constraints
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub dgvPollutants_DataError(ByVal sender As Object, ByVal e As DataGridViewDataErrorEventArgs) Handles dgvPollutants.DataError
        MsgBox(String.Format("Please enter a valid value in row {0} and column {1}.", e.RowIndex + 1, e.ColumnIndex + 1))
    End Sub



#End Region

    Private Function ProcessLandUseScenarioItems() As Boolean
        Dim dictPollutants As Dictionary(Of String, String) = GetPollutants()

        Dim i As Integer
        For i = 0 To g_Project.LUItems.Count - 1
            If g_Project.LUItems.Item(i).Enabled Then
                Begin(g_Project.LandCoverGridType, g_Project.LUItems, dictPollutants, g_Project.LandCoverGridDirectory)
                Return True
            End If
        Next i
        Return False
    End Function
    ''' <summary>
    ''' The workhorse of NSPECT. Automates the entire process of the nspect processing
    ''' </summary>
    Private Sub RunAnalysis()
        Try
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'STEP 1: Save file, populate xml params: -------------------------------------------------------------------------
            If Not SaveXmlFile() Then Return

            'Handles whether to overwrite existing groups of the same name or to generate a new group for outputs
            ' the global group layer might have stored a value, but the user may have reset the project or removed the group.
            If g_pGroupLayer <> -1 AndAlso MapWindowPlugin.MapWindowInstance.Layers.Groups.ItemByHandle(g_pGroupLayer) IsNot Nothing Then
                If MapWindowPlugin.MapWindowInstance.Layers.Groups.ItemByHandle(g_pGroupLayer).Text = g_Project.ProjectName Then
                    Dim res As MsgBoxResult = MsgBox(String.Format("Would you like to overwrite the last results group named {0}?", g_Project.ProjectName), MsgBoxStyle.YesNoCancel, "Replace Results?")
                    If res = MsgBoxResult.Yes Then
                        MapWindowPlugin.MapWindowInstance.Layers.Groups.Remove(g_pGroupLayer)
                    ElseIf res = MsgBoxResult.Cancel Then
                        Return
                    End If
                End If
            End If

            g_pGroupLayer = MapWindowPlugin.MapWindowInstance.Layers.Groups.Add(g_Project.ProjectName)

            'Init your global dictionary to hold the metadata records as well as the global xml prj file
            g_dicMetadata = New Dictionary(Of String, String)
            'END STEP 1: -----------------------------------------------------------------------------------------------------

            'STEP 4: Get the Management Scenarios: ------------------------------------------------------------------------------------
            'If they're using, we send them over to modMgmtScen to implement
            If g_Project.MgmtScenHolder.Count > 0 Then  'Return here!
                MgmtScenSetup(g_Project.MgmtScenHolder, g_Project.LandCoverGridType, g_Project.LandCoverGridDirectory)
            End If

            'STEP 6: Landuses sent off to modLanduse for processing -----------------------------------------------------
            Dim AreThereLandUseScenarioItems As Boolean = ProcessLandUseScenarioItems()

            'STEP 7: ---------------------------------------------------------------------------------------------------------
            'Obtain Watershed values
            Dim strWaterShed As String = String.Format("Select * from WSDelineation Where Name like '{0}'", g_Project.WaterShedDelin)
            Using cmdWS As New DataHelper(strWaterShed)
                'Set the Analysis Environment and globals for output workspace
                SetGlobalEnvironment(cmdWS.GetCommand(), _SelectLyrPath, _SelectedShapes)
            End Using

            '' g_pSelectedPolyClip is set by SetGlobalEnvironment and must be called afterward.
            'account for non-adjacent polygons
            If g_Project.UseSelectedPolygons Then
                If CheckMultiPartPolygon(g_pSelectedPolyClip) Then
                    'MsgBox("Warning: Your selected polygons are not adjacent.  Please select only polygons that are adjacent.", MsgBoxStyle.Critical, "Non-adjacent Polygons Detected")
                    'Return DLE 4/3/2014 THIS SHOULD BE TURNED BACK ON OR TESTED  XXXX
                    '4/3/2014: Tested and seems to work just fine with multi-part and non-adjacent polygons.
                End If
            End If

            'STEP 9: ---------------------------------------------------------------------------------------------------------
            'Create the runoff GRID
            'Get the precip scenario stuff
            Dim strPrecip As String = String.Format("Select * from PrecipScenario where name like '{0}'", g_Project.PrecipScenario)
            Using cmdPrecip As New DataHelper(strPrecip)
                Using dataPrecip As OleDbDataReader = cmdPrecip.ExecuteReader()
                    dataPrecip.Read()
                    'Added 6/04 to account for different PrecipTypes
                    g_intPrecipType = dataPrecip.Item("PrecipType")
                End Using
                'If there has been a land use added, then a new LCType has been created, hence we get it from g_strLCTypename
                Dim strLCType As String
                If AreThereLandUseScenarioItems Then
                    strLCType = g_LandUse_LCTypeName
                Else
                    strLCType = g_Project.LandCoverGridType
                End If
                If Not CreateRunoffGrid(g_Project.LandCoverGridDirectory, strLCType, cmdPrecip.GetCommand(), g_Project.SoilsHydDirectory, g_Project.OutputItems) Then
                    Return
                End If
            End Using

            'STEP 10: ---------------------------------------------------------------------------------------------------------
            'Process pollutants
            Dim i As Integer
            For i = 0 To g_Project.PollItems.Count - 1
                If g_Project.PollItems.Item(i).intApply = 1 Then
                    'If user is NOT ignoring the pollutant then send the whole item over along with LCType
                    Dim pollitem As PollutantItem = g_Project.PollItems.Item(i)
                    If Not PollutantConcentrationSetup(pollitem, g_Project.WaterQuality, g_Project.OutputItems) Then
                        Return
                    End If
                End If
            Next i

            'Step 11: Erosion -------------------------------------------------------------------------------------------------
            'Check that they have chosen Erosion
            If g_Project.IntCalcErosion = 1 Then
                If _IsAnnualPrecipScenario Then 'If Annual (0) then TRUE, ergo RUSLE
                    If g_Project.IntRainGridBool Then
                        If Not RUSLESetup(g_Project.StrRainGridFileName, g_Project.SoilsKFileName, g_Project.StrSDRGridFileName, g_Project.LandCoverGridType, g_Project.OutputItems) Then
                            Return
                        End If
                    ElseIf g_Project.IntRainConstBool Then
                        If Not RUSLESetup(g_Project.StrRainGridFileName, g_Project.SoilsKFileName, g_Project.StrSDRGridFileName, g_Project.LandCoverGridType, g_Project.OutputItems, g_Project.DblRainConstValue) Then
                            Return
                        End If
                    End If
                Else 'If event (1) then False, ergo MUSLE
                    If Not MUSLESetup(g_Project.SoilsDefName, g_Project.SoilsKFileName, g_Project.LandCoverGridType, g_Project.OutputItems) Then
                        Return
                    End If
                End If
            End If

            'STEP 12 : Cleanup any temp critters -------------------------------------------------------------------------------
            'g_DictTempNames holds the names of all temporary landuses and/or coefficient sets created during the Landuse scenario
            'portion of our program, for example CCAP1, or NitSet1.  We now must eliminate them from the database if they exist.
            If g_LandUse_DictTempNames.Count > 0 Then
                If AreThereLandUseScenarioItems Then
                    Cleanup(g_LandUse_DictTempNames, (g_Project.PollItems), (g_Project.LandCoverGridType))
                End If
            End If

            'STEP 15: create string describing project parameters ---------------------------------------------------------------
            'TODO metadata stuff
            'strProjectInfo = modUtil.ParseProjectforMetadata(_XmlPrjParams, _strFileName)

            'STEP 16: Apply the metadata to each of the rasters in the group layer ----------------------------------------------
            'TODO metadata stuff
            'modUtil.CreateMetadata(g_pGroupLayer, strProjectInfo)

            'Go into workspace and rid it of all rasters
            CleanGlobals()

            'Save xml to ensure outputs are saved
            g_Project.SaveFile(_currentFileName)

            Close()

            Return
        Catch ex As Exception
            HandleError(ex)
        Finally
            System.Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub CleanGlobals()
        Try
            'Sub to rid the world of stragling GRIDS, i.e. the ones established for global usse

            If Not g_pSCS100Raster Is Nothing Then
                g_pSCS100Raster.Close()
                g_pSCS100Raster = Nothing
            End If

            If Not g_pMetRunoffRaster Is Nothing Then
                g_pMetRunoffRaster.Close()
                g_pMetRunoffRaster = Nothing
            End If

            If Not g_pRunoffRaster Is Nothing Then
                g_pRunoffRaster.Close()
                g_pRunoffRaster = Nothing
            End If

            If Not g_pDEMRaster Is Nothing Then
                g_pDEMRaster.Close()
                g_pDEMRaster = Nothing
            End If

            If Not g_pFlowAccRaster Is Nothing Then
                g_pFlowAccRaster.Close()
                g_pFlowAccRaster = Nothing
            End If

            If Not g_pFlowDirRaster Is Nothing Then
                g_pFlowDirRaster.Close()
                g_pFlowDirRaster = Nothing
            End If

            If Not g_pLSRaster Is Nothing Then
                g_pLSRaster.Close()
                g_pLSRaster = Nothing
            End If

            If Not g_KFactorRaster Is Nothing Then
                g_KFactorRaster.Close()
                g_KFactorRaster = Nothing
            End If

            If Not g_pPrecipRaster Is Nothing Then
                g_pPrecipRaster.Close()
                g_pPrecipRaster = Nothing
            End If

            If Not g_LandCoverRaster Is Nothing Then
                g_LandCoverRaster.Close()
                g_LandCoverRaster = Nothing
            End If

            If Not g_pSelectedPolyClip Is Nothing Then
                g_pSelectedPolyClip = Nothing
            End If

            For Each file As String In g_TempFilesToDel
                ' Try to delete the file specified as well as any related files.
                DataManagement.DeleteGrid(file)
                DataManagement.DeleteShapefile(file)
                DataManagement.TryDelete(file)
            Next

        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub
    Private Sub cmdRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdRun.Click
        RunAnalysis()
    End Sub

#End Region

#Region "Helper Functions"

    Private Sub RefreshPolygonCount()
        chkSelectedPolys.Text = String.Format("{0} Selected Polygons Only", MapWindowPlugin.MapWindowInstance.View.SelectedShapes.NumSelected)
    End Sub

    ''' <summary>
    ''' Used by the selection form to set the selected shape
    ''' </summary>
    ''' <remarks></remarks>
    ''' <param name="layerHandle"></param>
    Public Sub SetSelectedShape(ByVal layerHandle As Integer)
        'Uses the current layer and cycles the select shapes, populating a list of shape index values
        If layerHandle <> -1 And MapWindowPlugin.MapWindowInstance.View.SelectedShapes.NumSelected > 0 Then
            chkSelectedPolys.Checked = True
            _SelectLyrPath = MapWindowPlugin.MapWindowInstance.Layers(layerHandle).FileName
            _SelectedShapes = New List(Of Integer)

            Dim axMap As AxMapWinGIS.AxMap = CType(MapWindowPlugin.MapWindowInstance.GetOCX(), AxMapWinGIS.AxMap)
            Dim sf As MapWinGIS.Shapefile = CType(axMap.get_GetObject(layerHandle), MapWinGIS.Shapefile)
            If Not sf Is Nothing Then   ' layer object can be an image as well, nothing wil be returned in this case
                For i As Integer = 0 To sf.NumShapes - 1
                    If sf.ShapeSelected(i) Then
                        _SelectedShapes.Add(i)
                    End If
                Next i
            End If

            RefreshPolygonCount()
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
            cboTargetLayer.Items.Clear()

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
                    g_Project.Xml = _currentFileName
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
                g_Project.Xml = _currentFileName
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

            strLCChanges = String.Format("SELECT LCCLASS.Name as Name2, LCTYPE.LCTYPEID FROM LCTYPE INNER JOIN LCCLASS ON LCTYPE.LCTYPEID = LCCLASS.LCTYPEID WHERE LCTYPE.NAME LIKE '{0}'", cboLandCoverType.Text)
            Using lcChangesCmd As New OleDbCommand(strLCChanges, g_DBConn)
                Using lcChanges As OleDbDataReader = lcChangesCmd.ExecuteReader()
                    arrClassList.Clear()
                    Do While lcChanges.Read()
                        arrClassList.Insert(i, lcChanges.Item("Name2"))
                    Loop
                End Using
            End Using

        Catch ex As Exception
            HandleError(ex)
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
                    strSQLWQStdPoll = String.Format("SELECT POLLUTANT.NAME, POLL_WQCRITERIA.THRESHOLD FROM POLL_WQCRITERIA INNER JOIN POLLUTANT ON POLL_WQCRITERIA.POLLID = POLLUTANT.POLLID Where POLL_WQCRITERIA.WQCRITID = {0}", dataWQCrit.Item("WQCRITID"))
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
                    MsgBox("Warning: There are no water quality standards remaining.  Please add a new one.", MsgBoxStyle.Critical, "Recordset Empty")
                End If
            End Using

        Catch ex As Exception
            HandleError(ex)
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
            Dim strSelectCoeff As String = String.Format("SELECT POLLUTANT.POLLID, POLLUTANT.NAME, COEFFICIENTSET.NAME AS NAME2 FROM POLLUTANT INNER JOIN COEFFICIENTSET ON POLLUTANT.POLLID = COEFFICIENTSET.POLLID Where POLLUTANT.NAME LIKE '{0}'", strPollutantName)
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
        Dim strCurrWShed As String = String.Format("Select * from WSDelineation where Name Like '{0}'", g_Project.WaterShedDelin)

        Using WSCmd As OleDbCommand = New OleDbCommand(strCurrWShed, g_DBConn)
            Using wsData As OleDbDataReader = WSCmd.ExecuteReader
                wsData.Read()
                If wsData.HasRows() Then
                    Dim strBasin As String = wsData.Item("wsfilename")
                    If Path.GetExtension(strBasin) <> ".shp" Then
                        strBasin = strBasin + ".shp"
                    End If
                    If Not ExistsLayerInMapByFileName(strBasin) Then
                        If AddFeatureLayerToMapFromFileName(strBasin, String.Format("{0} Drainage Basins", wsData.Item("Name"))) Then

                            arrAreaList.Add(String.Format("{0} Drainage Basins", wsData.Item("Name")))
                        Else
                            MsgBox(String.Format("Could not find watershed layer: {0} .  Please add the watershed layer to the map.", wsData.Item("wsfilename")), MsgBoxStyle.Critical, "File Not Found")
                        End If
                    End If
                Else
                    MsgBox(String.Format("Could not find watershed data: {0}. Perhaps we should have tried with {1}.", g_Project.WaterShedDelin, cboWaterShedDelineations.Text), MsgBoxStyle.Critical, "File Not Found")
                End If
            End Using
        End Using
    End Sub

    Private Function GetPollutantsIdx() As Integer
        Dim _intPollCount As Short = g_Project.PollItems.Count
        Dim idx As Integer
        If _intPollCount > 0 Then
            dgvPollutants.Rows.Clear()

            For i = 0 To _intPollCount - 1
                With dgvPollutants
                    idx = .Rows.Add()

                    .Rows(idx).Cells("PollApply").Value = g_Project.PollItems.Item(i).intApply
                    .Rows(idx).Cells("PollutantName").Value = g_Project.PollItems.Item(i).strPollName
                    Dim strPollName = g_Project.PollItems.Item(i).strPollName
                    PopulateCoefType(strPollName, idx)
                    .Rows(idx).Cells("CoefSet").Value = g_Project.PollItems.Item(i).strCoeffSet

                    .Rows(idx).Cells("WhichCoeff").Value = g_Project.PollItems.Item(i).strCoeff
                    .Rows(idx).Cells("Threshold").Value = CStr(g_Project.PollItems.Item(i).intThreshold)

                    If Len(g_Project.PollItems.Item(i).strTypeDefXmlFile) > 0 Then
                        .Rows(idx).Cells("TypeDef").Value = CStr(g_Project.PollItems.Item(i).strTypeDefXmlFile)
                    End If
                End With
            Next i

        End If
        Return idx
    End Function

    Private Function LoadLandCoverGrid() As Boolean

        'Check to see if the LC cover is in the map, if so, set the combobox
        If LayerLoadedInMap(g_Project.LandCoverGridName) Then
            cboLCLayer.SelectedIndex = GetIndexOfEntry((g_Project.LandCoverGridName), cboLCLayer)
        Else
            If AddRasterLayerToMapFromFileName(g_Project.LandCoverGridDirectory) Then

                With cboLCLayer
                    .Items.Add(g_Project.LandCoverGridName)
                    .SelectedIndex = GetIndexOfEntry(g_Project.LandCoverGridName, cboLCLayer)
                End With

            Else
                Dim browseForDataSet As Short = MsgBox(String.Format("Could not find the Land Cover dataset: {0}.  Would you like to browse for it?", g_Project.LandCoverGridDirectory), MsgBoxStyle.YesNo, "Cannot Locate Dataset")
                If browseForDataSet = MsgBoxResult.Yes Then
                    g_Project.LandCoverGridDirectory = AddInputFromGxBrowser("Raster")
                    If g_Project.LandCoverGridDirectory <> "" Then
                        If AddRasterLayerToMapFromFileName(g_Project.LandCoverGridDirectory) Then
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

        cboLandCoverType.SelectedIndex = GetIndexOfEntry((g_Project.LandCoverGridType), cboLandCoverType)

        Return True
    End Function
    Private Function LoadSoils() As Boolean
        If RasterExists(g_Project.SoilsHydDirectory) Then
            cboSoilsLayer.SelectedIndex = GetIndexOfEntry(g_Project.SoilsDefName, cboSoilsLayer)
            Return True
        Else
            MsgBox("Could not find soils dataset.  Please correct the soils definition in the Advanced Settings.", MsgBoxStyle.Critical, "Dataset Missing")
            Return False
        End If
    End Function
    Private Sub LoadWatersheds()
        chkLocalEffects.CheckState = g_Project.IntLocalEffects
        chkSelectedPolys.CheckState = g_Project.IntSelectedPolys

        If chkSelectedPolys.CheckState = 1 Then
            '1st see if it's in the map
            Dim strSelected As String = g_Project.StrSelectedPolyFileName
            If Path.GetExtension(strSelected) <> ".shp" Then
                strSelected = strSelected + ".shp"
            End If
            _SelectLyrPath = strSelected
            _SelectedShapes = g_Project.intSelectedPolyList
            RefreshPolygonCount()
            'chkSelectedPolys.Text = String.Format("{0} Selected Polygons Only", _SelectedShapes.Count)

            If Not ExistsLayerInMapByFileName(strSelected) Then
                'Not there then add it
                If AddFeatureLayerToMapFromFileName(strSelected, g_Project.StrSelectedPolyLyrName) Then
                    arrAreaList.Add(g_Project.StrSelectedPolyLyrName)
                Else
                    'Can't find it, then send em searching
                    Dim browseForData = MsgBox(String.Format("Could not find the Selected Polygons file used to limit extent: {0}.  Would you like to browse for it? ", strSelected), MsgBoxStyle.YesNo, "Cannot Locate Dataset")
                    If browseForData = MsgBoxResult.Yes Then
                        'if they want to look for it then give em the browser
                        g_Project.StrSelectedPolyFileName = AddInputFromGxBrowser("Feature")
                        'if they actually find something, throw it in the map
                        If Len(g_Project.StrSelectedPolyFileName) > 0 Then
                            If AddFeatureLayerToMapFromFileName(g_Project.StrSelectedPolyFileName) Then
                                arrAreaList.Add(SplitFileName(g_Project.StrSelectedPolyFileName))

                            End If
                        End If
                    Else
                        g_Project.IntSelectedPolys = 0
                        chkSelectedPolys.CheckState = CheckState.Unchecked
                    End If
                End If
            End If
        End If
    End Sub
    Private Function LoadErosion() As Boolean
        If g_Project.IntCalcErosion = 1 Then
            chkCalcErosion.CheckState = CheckState.Checked
        Else
            chkCalcErosion.CheckState = CheckState.Unchecked
        End If

        'Step: Erosion Tab - Precip
        'Either they use the GRID
        optUseGRID.Checked = g_Project.IntRainGridBool
        If optUseGRID.Checked Then
            If Not RasterExists(g_Project.StrRainGridFileName) Then
                Dim browseForData = MsgBox(String.Format("Could not find Rainfall GRID: {0}.  Would you like to browse for it?", g_Project.StrRainGridFileName), MsgBoxStyle.YesNo, "Cannot Locate Dataset")
                If browseForData = MsgBoxResult.Yes Then
                    Using dlgOpen As New OpenFileDialog()
                        dlgOpen.Title = "Choose Rainfall Factor GRID"
                        dlgOpen.Filter = (New Grid).CdlgFilter
                        If dlgOpen.ShowDialog = System.Windows.Forms.DialogResult.OK Then
                            txtbxRainGrid.Text = dlgOpen.FileName
                            g_Project.StrRainGridFileName = dlgOpen.FileName
                        End If
                    End Using
                Else
                    Return False
                End If
            Else
                txtbxRainGrid.Text = g_Project.StrRainGridFileName
            End If
        End If

        'Or they use a constant value
        optUseValue.Checked = g_Project.IntRainConstBool

        If optUseValue.Checked Then
            txtRainValue.Text = CStr(g_Project.DblRainConstValue)
        End If

        'SDR GRID
        'If Not _XmlPrjParams.intUseOwnSDR Is Nothing Then
        If g_Project.IntUseOwnSDR = 1 Then
            chkSDR.CheckState = CheckState.Checked
            txtSDRGRID.Text = g_Project.LandCoverGridDirectory
        Else
            chkSDR.CheckState = CheckState.Unchecked
            txtSDRGRID.Text = g_Project.StrSDRGridFileName
        End If
        'End If

        Return True
    End Function
    Private Sub LoadLandUses(ByRef idx As Integer)
        Dim landUsesCount As Short = g_Project.LUItems.Count

        If landUsesCount > 0 Then
            dgvLandUse.Rows.Clear()
            For i = 0 To landUsesCount - 1
                With dgvLandUse
                    idx = .Rows.Add()
                    .Rows(idx).Cells("LUApply").Value = g_Project.LUItems.Item(i).Enabled
                    .Rows(idx).Cells("LUScenario").Value = g_Project.LUItems.Item(i).strLUScenName
                    .Rows(idx).Cells("LUScenarioXml").Value = g_Project.LUItems.Item(i).strLUScenXmlFile
                End With
            Next i
        End If
    End Sub
    Private Function LoadManagementScenarios(ByVal idx As Integer) As Boolean
        Dim scenarioCount As Short = g_Project.MgmtScenHolder.Count

        If scenarioCount > 0 Then
            dgvManagementScen.Rows.Clear()
            For i = 0 To scenarioCount - 1
                With dgvManagementScen
                    idx = .Rows.Add()

                    If LayerLoadedInMap(g_Project.MgmtScenHolder.Item(i).strAreaName) Then
                        PopulateManagement(idx)
                        .Rows(idx).Cells("ManageApply").Value = g_Project.MgmtScenHolder.Item(i).intApply
                        .Rows(idx).Cells("ChangeAreaLayer").Value = g_Project.MgmtScenHolder.Item(i).strAreaName
                        .Rows(idx).Cells("ChangeToClass").Value = g_Project.MgmtScenHolder.Item(i).strChangeToClass
                    Else
                        If AddFeatureLayerToMapFromFileName(g_Project.MgmtScenHolder.Item(i).strAreaFileName) Then
                            arrAreaList.Add(g_Project.MgmtScenHolder.Item(i).strAreaName)

                            PopulateManagement(idx)
                            .Rows(idx).Cells("ManageApply").Value = g_Project.MgmtScenHolder.Item(i).intApply
                            .Rows(idx).Cells("ChangeAreaLayer").Value = g_Project.MgmtScenHolder.Item(i).strAreaName
                            .Rows(idx).Cells("ChangeToClass").Value = g_Project.MgmtScenHolder.Item(i).strChangeToClass
                        Else
                            Dim browseForData = MsgBox(String.Format("Could not find Management Sceario Area Layer: {0}.  Would you like to browse for it?", g_Project.MgmtScenHolder.Item(i).strAreaFileName), MsgBoxStyle.YesNo, "Cannot Locate Dataset:" & g_Project.MgmtScenHolder.Item(i).strAreaFileName)
                            If browseForData = MsgBoxResult.Yes Then
                                g_Project.MgmtScenHolder.Item(i).strAreaFileName = AddInputFromGxBrowser("Feature")
                                If AddFeatureLayerToMapFromFileName(g_Project.MgmtScenHolder.Item(i).strAreaFileName) Then
                                    arrAreaList.Add(g_Project.MgmtScenHolder.Item(i).strAreaName)

                                    PopulateManagement(idx)
                                    .Rows(idx).Cells("ManageApply").Value = g_Project.MgmtScenHolder.Item(i).intApply
                                    .Rows(idx).Cells("ChangeAreaLayer").Value = g_Project.MgmtScenHolder.Item(i).strAreaName
                                    .Rows(idx).Cells("ChangeToClass").Value = g_Project.MgmtScenHolder.Item(i).strChangeToClass
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

            txtProjectName.Text = g_Project.ProjectName
            txtOutputWS.Text = g_Project.ProjectWorkspace

            If Not LoadLandCoverGrid() Then Return
            If Not LoadSoils() Then Return
            cboPrecipitationScenarios.SelectedIndex = GetIndexOfEntry(g_Project.PrecipScenario, cboPrecipitationScenarios)
            cboWaterShedDelineations.SelectedIndex = GetIndexOfEntry(g_Project.WaterShedDelin, cboWaterShedDelineations)
            AddBasinpolyToMap()
            cboWaterQualityCriteriaStd.SelectedIndex = GetIndexOfEntry(g_Project.WaterQuality, cboWaterQualityCriteriaStd)
            LoadWatersheds()
            If Not LoadErosion() Then Return
            Dim idx As Integer = GetPollutantsIdx()
            LoadLandUses(idx)
            LoadManagementScenarios(idx)

            ReloadTargetLayers()

            'Reset to first tab
            TabsForGrids.SelectedIndex = 0
            _isFileOnDisk = True
        Catch ex As Exception
            HandleError(ex)
        Finally
            System.Windows.Forms.Cursor.Current = Cursors.Default
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
                If Not _isFileOnDisk Then
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
                        _isFileOnDisk = True
                        g_Project.SaveFile(_currentFileName)
                        SaveXmlFile = True
                    Else
                        SaveXmlFile = False
                        Exit Function
                    End If
                Else
                    'Now check to see if the name changed
                    If _strOpenFileName <> txtProjectName.Text Then
                        intvbYesNo = MsgBox(String.Format("You have changed the name of this project.  Would you like to save your settings as a new file?{0}{1}Yes{1} -    Save as new OpenNSPECT project file{0}{1}No{1} -    Save changes to current OpenNSPECT project file{0}{1}Cancel{1} -    Return to the project window", vbNewLine, vbTab), MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, "OpenNSPECT")
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
                                _isFileOnDisk = True
                                g_Project.SaveFile(_currentFileName)
                                SaveXmlFile = True
                            Else
                                SaveXmlFile = False
                                Exit Function
                            End If
                        ElseIf intvbYesNo = MsgBoxResult.No Then
                            g_Project.SaveFile(_currentFileName)
                            _isFileOnDisk = True
                            SaveXmlFile = True
                        Else
                            SaveXmlFile = False
                            Exit Function
                        End If
                    Else
                        g_Project.SaveFile(_currentFileName)
                        _isFileOnDisk = True
                        SaveXmlFile = True
                    End If
                End If
            End Using

        Catch ex As Exception
            HandleError(ex)
            Return False
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
                    MsgBox("You have chosen 'Selected watersheds only'.  Please select watersheds.", MsgBoxStyle.Critical, "No Selected Features Found")
                    Return False
                End If
            End If

            'Project Name
            If Len(txtProjectName.Text) > 0 Then
                ParamsPrj.ProjectName = Trim(txtProjectName.Text)
            Else
                MsgBox("Please enter a name for this project.", MsgBoxStyle.Information, "Enter Name")
                txtProjectName.Focus()
                Return False
            End If

            'Working Directory
            If (Len(txtOutputWS.Text) > 0) And Directory.Exists(txtOutputWS.Text) Then
                ParamsPrj.ProjectWorkspace = Trim(txtOutputWS.Text)
            Else
                MsgBox("Please choose a valid output working directory.", MsgBoxStyle.Information, "Choose Workspace")
                txtOutputWS.Focus()
                Return False
            End If

            'LandCover
            If cboLCLayer.Text = "" Then
                MsgBox("Please select a Land Cover layer before continuing.", MsgBoxStyle.Information, "Select Land Cover Layer")
                cboLCLayer.Focus()
                Return False
            Else
                If LayerLoadedInMap(cboLCLayer.Text) Then
                    ParamsPrj.LandCoverGridName = cboLCLayer.Text
                    ParamsPrj.LandCoverGridDirectory = GetLayerFilename(cboLCLayer.Text)
                    ParamsPrj.LandCoverGridUnits = GetRasterDistanceUnits(cboLCLayer.Text)
                Else
                    MsgBox("The Land Cover layer you have choosen is not in the current map frame.", MsgBoxStyle.Information, "Layer Not Found")
                    Return False
                End If
            End If

            'LC Type
            If cboLandCoverType.Text = "" Then
                MsgBox("Please select a Land Class Type before continuing.", MsgBoxStyle.Information, "Select Land Class Type")
                cboLandCoverType.Focus()
                Return False
            Else
                ParamsPrj.LandCoverGridType = cboLandCoverType.Text
            End If

            'Soils - use definition to find datasets, if there use, if not tell the user
            If cboSoilsLayer.Text = "" Then
                MsgBox("Please select a Soils definition before continuing.", MsgBoxStyle.Information, "Select Soils Layer")
                cboSoilsLayer.Focus()
                Return False
            Else
                If RasterExists(_soilsFileName) Then
                    ParamsPrj.SoilsDefName = cboSoilsLayer.Text
                    ParamsPrj.SoilsHydDirectory = _soilsFileName
                Else
                    MsgBox(String.Format("The hydrologic soils layer {0} you have selected is missing.  Please check you soils definition.", _soilsFileName), MsgBoxStyle.Information, "Soils Layer Not Found")
                    Return False
                End If
            End If

            'PrecipScenario
            'If the layer is in the map, get out, all is well- _strPrecipFile is established on the
            'PrecipCbo Click event
            If LayerLoadedInMap(SplitFileName(_strPrecipFile)) Then
                ParamsPrj.PrecipScenario = cboPrecipitationScenarios.Text
            Else
                'Check if you can add it, if so, all is well
                If RasterExists(_strPrecipFile) Then
                    ParamsPrj.PrecipScenario = cboPrecipitationScenarios.Text
                Else
                    'Can't find it...well, then send user to Browse
                    MsgBox(String.Format("Unable to find precip dataset: {0}.  Please Correct", _strPrecipFile), MsgBoxStyle.Information, "Cannot Find Dataset")
                    _strPrecipFile = BrowseForFileName("Raster")
                    'If new one found, then we must update Database
                    If Len(_strPrecipFile) > 0 Then
                        strUpdatePrecip = String.Format("UPDATE PrecipScenario SET precipScenario.PrecipFileName = '{0}' WHERE NAME = '{1}'", _strPrecipFile, cboPrecipitationScenarios.Text)
                        Using PreUpdCmd As OleDbCommand = New OleDbCommand(strUpdatePrecip, g_DBConn)
                            PreUpdCmd.ExecuteNonQuery()
                        End Using

                        'Now we can set the xmlParams
                        ParamsPrj.PrecipScenario = cboPrecipitationScenarios.Text
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
                ParamsPrj.WaterShedDelin = cboWaterShedDelineations.Text
            Else
                MsgBox("There is a problems with the selected Watershed Delineation.", MsgBoxStyle.Information, "Watershed Delineation")
                Return False
            End If

            'Water Quality
            If Len(cboWaterQualityCriteriaStd.Text) > 0 Then
                ParamsPrj.WaterQuality = cboWaterQualityCriteriaStd.Text
            Else
                MsgBox("Please select a water quality standard.", MsgBoxStyle.Information, "Water Quality Standard Missing")
                Return False
            End If

            'Checkboxes, straight up values
            ParamsPrj.IntLocalEffects = chkLocalEffects.CheckState

            'Theoreretically, user could open file that had selected sheds.
            If chkSelectedPolys.Checked = True Then
                ParamsPrj.IntSelectedPolys = chkSelectedPolys.CheckState
                ParamsPrj.StrSelectedPolyFileName = _SelectLyrPath
                Dim tmpidx As Integer = GetLayerIndexByFilename(_SelectLyrPath)
                If tmpidx <> -1 Then
                    ParamsPrj.StrSelectedPolyLyrName = MapWindowPlugin.MapWindowInstance.Layers(tmpidx).Name
                Else
                    ParamsPrj.StrSelectedPolyLyrName = ""
                End If
                ParamsPrj.intSelectedPolyList = _SelectedShapes
            Else
                ParamsPrj.IntSelectedPolys = 0
            End If

            TabsForGrids.SelectedIndex = 1

            'Erosion Tab
            'Calc Erosion checkbox
            ParamsPrj.IntCalcErosion = chkCalcErosion.CheckState

            If chkCalcErosion.CheckState Then
                If RasterExists((lblKFactor.Text)) Then
                    ParamsPrj.SoilsKFileName = lblKFactor.Text
                Else
                    MsgBox(String.Format("The K Factor soils dataset {0} you have selected is missing.  Please check your soils definition.", lblKFactor.Text), MsgBoxStyle.Information, "Soils K Factor Not Found")
                    Return False
                End If

                'Check the Rainfall Factor grid objects.
                If frameRainFall.Visible = True Then

                    If optUseGRID.Checked Then

                        If Len(txtbxRainGrid.Text) > 0 And (InStr(1, txtbxRainGrid.Text, cboLCLayer.Text, 1) = 0) Then
                            ParamsPrj.IntRainGridBool = 1
                            ParamsPrj.IntRainConstBool = 0
                            ParamsPrj.StrRainGridName = txtbxRainGrid.Text
                            ParamsPrj.StrRainGridFileName = txtbxRainGrid.Text
                        Else
                            MsgBox("Please choose a rainfall Grid.", MsgBoxStyle.Information, "Select Rainfall GRID")
                            TabsForGrids.SelectedIndex = 1
                            Return False

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
                                ParamsPrj.IntRainConstBool = 1
                                ParamsPrj.DblRainConstValue = CDbl(txtRainValue.Text)
                                ParamsPrj.StrRainGridFileName = ""
                            End If
                        End If

                    Else
                        MsgBox("You must choose a rainfall factor.", MsgBoxStyle.Information, "Rainfall Factor Missing")
                        Return False
                    End If
                End If

                'Soil Delivery Ratio
                'Added 12/03/07 to account for soil delivery ratio GRID, user can now provide.
                If chkSDR.CheckState = 1 Then
                    If Len(txtSDRGRID.Text) > 0 Then
                        If RasterExists((txtSDRGRID.Text)) Then
                            ParamsPrj.IntUseOwnSDR = 1
                            ParamsPrj.StrSDRGridFileName = txtSDRGRID.Text
                        Else
                            MsgBox(String.Format("SDR GRID {0} not found.", txtSDRGRID.Text), MsgBoxStyle.Information, "SDR GRID Not Found")
                            Return False
                        End If
                    Else
                        MsgBox("Please select an SDR GRID.", MsgBoxStyle.Information, "SDR GRID Not Selected")
                        Return False
                    End If
                Else
                    ParamsPrj.IntUseOwnSDR = 0
                    ParamsPrj.StrSDRGridFileName = txtSDRGRID.Text
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
                dgvManagementScen.Focus()
                Return False
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
                    If (pollitem.strCoeff = "Use shapefile...") Then
                        pollitem.idxShapefile = row.Cells("IndexShapefile").Value
                        pollitem.idxField = row.Cells("IndexFieldName").Value
                    End If
                    pollitem.intThreshold = CShort(row.Cells("Threshold").Value)
                    If row.Cells("TypeDef").Value <> "" Then

                        pollitem.strTypeDefXmlFile = row.Cells("TypeDef").Value
                    End If
                    ParamsPrj.PollItems.Add(pollitem)
                Next
            Else
                dgvPollutants.Focus()
                Return False
            End If

            TabsForGrids.SelectedIndex = 2
            'Land Uses
            For Each row As DataGridViewRow In dgvLandUse.Rows
                Dim luitem As LandUseItem
                If Len(row.Cells("LUScenario").Value) > 0 Then
                    luitem = New LandUseItem
                    luitem.intID = row.Index + 1
                    luitem.Enabled = row.Cells("LUApply").FormattedValue
                    luitem.strLUScenName = row.Cells("LUScenario").Value
                    luitem.strLUScenXmlFile = row.Cells("LUScenarioXml").Value
                    ParamsPrj.LUItems.Add(luitem)  ' DLE 10/26/2012  Test if this is extra LU param addition... NOPE!  Needed
                End If
            Next
            'If it gets to here, all is well
            ValidateData = True

            g_Project = ParamsPrj

        Catch ex As Exception
            HandleError(ex)
        Finally
            TabsForGrids.SelectedIndex = 0
        End Try
    End Function

    ''' <summary>
    ''' Find and store index shapefile and field names for pollutant
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetPickShapeInfo(ByVal row As Integer) As Boolean
         Dim pollName As String = Me.dgvPollutants.Rows(row).Cells(1).Value.ToString
        Debug.WriteLine("Pollutant is  " & pollName)
        Dim getShapeinfo As New PollCoeffSelectionShapefileForm(pollName)
        getShapeinfo.ShowDialog()
        If (getShapeinfo.DialogResult = Windows.Forms.DialogResult.OK) Then
            Dim sf As String = getShapeinfo.indexShapefileName
            Me.dgvPollutants.Rows(row).Cells(4).Value = sf
            Dim sfld As String = getShapeinfo.indexField
            Me.dgvPollutants.Rows(row).Cells(5).Value = sfld
            MsgBox("On MainForm, " & pollName & " shapefile is " & getShapeinfo.indexShapefileName & " and index field is " & getShapeinfo.indexField)
            GetPickShapeInfo = True
        End If

    End Function

    ''' <summary>
    ''' Validate the pollutant table values
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ValidatePollutants() As Boolean
        'Function to validate pollutants
        Try
            'For Each row As DataGridViewRow In dgvPollutants.Rows
            For irow As Integer = 0 To dgvPollutants.Rows.Count - 1
                Dim row As DataGridViewRow = dgvPollutants.Rows(irow)
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
                            'Add code to get Picks shapefile name and field name here
                            If row.Cells("WhichCoeff").Value = "Use shapefile..." Then
                                If (GetPickShapeInfo(irow)) Then
                                    ValidatePollutants = True
                                Else
                                    ValidatePollutants = False
                                End If
                            Else
                                ValidatePollutants = True
                            End If
                        End If

                    End If
                Else
                    ValidatePollutants = True
                End If
            Next
        Catch ex As Exception
            HandleError(ex)
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
                            MsgBox(String.Format("Unable to locate DEM dataset: {0}.", dtWS.Rows(0)("DEMFileName")), MsgBoxStyle.Critical, "Missing Dataset")
                            strDEM = BrowseForFileName("Raster")
                            If Len(strDEM) > 0 Then
                                dtWS.Rows(0)("DEMFileName") = strDEM
                                'rsWShed.Fields("DEMFileName").Value = strDEM
                                booUpdate = True
                            Else
                                Return False
                            End If
                            'WaterShed Delineation
                        ElseIf Not FeatureExists(dtWS.Rows(0)("wsfilename")) Then
                            MsgBox(String.Format("Unable to locate Watershed dataset: {0}.", dtWS.Rows(0)("wsfilename")), MsgBoxStyle.Critical, "Missing Dataset")
                            Dim strWShed1 = BrowseForFileName("Feature")
                            If Len(strWShed1) > 0 Then
                                dtWS.Rows(0)("wsfilename") = strWShed1
                                booUpdate = True
                            Else
                                Return False
                            End If
                            'Flow Direction
                        ElseIf Not RasterExists(dtWS.Rows(0)("FlowDirFileName")) Then
                            MsgBox(String.Format("Unable to locate Flow Direction GRID: {0}.", dtWS.Rows(0)("FlowDirFileName")), MsgBoxStyle.Critical, "Missing Dataset")
                            strFlowDirFileName = BrowseForFileName("Raster")
                            If Len(strFlowDirFileName) > 0 Then
                                dtWS.Rows(0)("FlowDirFileName") = strFlowDirFileName
                                booUpdate = True
                            Else
                                Return False
                            End If
                            'Flow Accumulation
                        ElseIf Not RasterExists(dtWS.Rows(0)("FlowAccumFileName")) Then
                            MsgBox(String.Format("Unable to locate Flow Accumulation GRID: {0}.", dtWS.Rows(0)("FlowAccumFileName")), MsgBoxStyle.Critical, "Missing Dataset")
                            strFlowAccumFileName = BrowseForFileName("Raster")
                            If Len(strFlowAccumFileName) > 0 Then
                                dtWS.Rows(0)("FlowAccumFileName") = strFlowAccumFileName
                                booUpdate = True
                            Else
                                Return False
                            End If
                            'Check for non-hydro correct GRIDS
                        ElseIf dtWS.Rows(0)("HydroCorrected") = 0 Then
                            If Not RasterExists(dtWS.Rows(0)("FilledDEMFileName")) Then
                                MsgBox(String.Format("Unable to locate the Filled DEM: {0}.", dtWS.Rows(0)("FilledDEMFileName")), MsgBoxStyle.Critical, "Missing Dataset")
                                strFilledDEMFileName = BrowseForFileName("Raster")
                                If Len(strFilledDEMFileName) > 0 Then
                                    dtWS.Rows(0)("FilledDEMFileName") = strFilledDEMFileName
                                    booUpdate = True
                                Else
                                    Return False
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
                        MsgBox("Please select a land class in row " + (row.Index + 1).ToString, MsgBoxStyle.Critical, "Missing Value")
                        ValidateMgmtScenario = False
                        Exit Function
                    End If

                End If
            Next

            ValidateMgmtScenario = True

        Catch ex As Exception
            HandleError(ex)
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
            cboPrecipitationScenarios.SelectedIndex = GetIndexOfEntry(strPrecName, cboPrecipitationScenarios)
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
            cboWaterQualityCriteriaStd.SelectedIndex = GetIndexOfEntry(strWQName, cboWaterQualityCriteriaStd)
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

#End Region

    Private Sub cboTargetLayer_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cboTargetLayer.SelectedIndexChanged
        ' get mapwindow layer index
        For i As Integer = 0 To MapWindowPlugin.MapWindowInstance.Layers.NumLayers - 1
            Dim layer As Layer = MapWindowPlugin.MapWindowInstance.Layers(i)
            If layer IsNot Nothing Then
                If layer.Name = cboTargetLayer.SelectedItem.ToString() Then
                    SetSelectedShape(i)
                    Return
                End If
            End If
        Next

    End Sub



End Class