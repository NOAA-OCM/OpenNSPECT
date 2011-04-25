'********************************************************************************************************
'File Name: frmNewWSDelin.vb
'Description: Form for new watershed delineation
'********************************************************************************************************
'The contents of this file are subject to the Mozilla Public License Version 1.1 (the "License"); 
'you may not use this file except in compliance with the License. You may obtain a copy of the License at 
'http://www.mozilla.org/MPL/ 
'Software distributed under the License is distributed on an "AS IS" basis, WITHOUT WARRANTY OF 
'ANY KIND, either express or implied. See the License for the specificlanguage governing rights and 
'limitations under the License. 
'
'Note: This code was converted from the vb6 NSPECT ArcGIS extension and so bears many of the old comments
'in the files where it was possible to leave them.
'
'Contributor(s): (Open source contributors should list themselves and their modifications here). 
'Oct 20, 2010:  Allen Anselmo allen.anselmo@gmail.com - 
'               Added licensing and comments to code

Imports System.Data.OleDb

Friend Class NewWatershedDelineationForm
    Inherits System.Windows.Forms.Form

    Private Const c_sModuleFileName As String = "frmNewWSDelin.vb"

    Private boolChange(3) As Boolean 'Array set to track changes in controls: On Change, cmdCreate is enabled
    Private _booProject As Boolean
    Private _frmWS As WatershedDelineationsForm
    Private _frmPrj As MainForm

    Private _intSize As Short 'Index for Size Combo
    Private _intCellSize As Short 'Cell Size of DEM Grid, used in Length Slope Calculation
    Private _InputDEMPath As String

    Private _strDirFileName As String
    Private _strAccumFileName As String
    Private _strFilledDEMFileName As String
    Private _strStreamLayer As String
    Private _strWShedFileName As String
    Private _strLSFileName As String
    Private _strNibbleName As String
    Private _strDEM2BName As String

    Private Const _dblSmall As Double = 0.03 '0.001    '
    Private Const _dblMedium As Double = 0.06 '0.01    '-Subwatershed sizes (3%, 6%, 10%)
    Private Const _dblLarge As Double = 0.1 '


#Region "Events"


    Private Sub frmNewWSDelin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'Init bool variables
            g_boolAgree = False
            g_boolHydCorr = False
            g_boolParams = False

            Dim i As Short

            For i = 0 To UBound(boolChange)
                boolChange(i) = False
            Next i

            cboStreamLayer.Items.Clear()
            For i = 0 To g_MapWin.Layers.NumLayers - 1
                If g_MapWin.Layers(g_MapWin.Layers.GetHandle(i)).LayerType = MapWindow.Interfaces.eLayerType.LineShapefile Then
                    cboStreamLayer.Items.Add(g_MapWin.Layers(g_MapWin.Layers.GetHandle(i)).Name)
                End If
            Next
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub txtWSDelinName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtWSDelinName.TextChanged
        Try
            boolChange(0) = True
            CheckEnabled()
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub txtDEMFile_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDEMFile.TextChanged
        Try
            boolChange(1) = True
            CheckEnabled()

        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub cmdBrowseDEMFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowseDEMFile.Click
        Try
            Dim pDEMRasterDataset As MapWinGIS.Grid

            pDEMRasterDataset = AddInputFromGxBrowserText(txtDEMFile, "Choose DEM GRID", Me, 0)
            If Not pDEMRasterDataset Is Nothing Then
                'Get the spatial reference
                Dim strProj As String = CheckSpatialReference(pDEMRasterDataset)
                If strProj = "" Then
                    MsgBox("The GRID you have choosen has no spatial reference information.  Please define a projection before continuing.", MsgBoxStyle.Exclamation, "No Project Information Detected")
                    Exit Sub
                Else
                    If strProj.ToLower.Contains("units=m") Then
                        cboDEMUnits.SelectedIndex = 0
                    Else
                        cboDEMUnits.SelectedIndex = 1
                    End If

                    cboDEMUnits.Refresh()

                End If

                _InputDEMPath = pDEMRasterDataset.Filename
                txtDEMFile.Text = _InputDEMPath

                pDEMRasterDataset.Close()

                CheckEnabled()

                ''Get the name
                '_strDemArray = Split(pDEMRasterDataset.CompleteName, "\")
                '_strDemName = m_strDemArray(UBound(m_strDemArray))
            End If
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub chkHydroCorr_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkHydroCorr.CheckedChanged
        Try
            Select Case chkHydroCorr.CheckState
                Case Windows.Forms.CheckState.Checked
                    chkStreamAgree.Enabled = True
                    cboStreamLayer.Enabled = True
                    cmdOptions.Enabled = True
                    g_boolHydCorr = True

                Case Windows.Forms.CheckState.Unchecked

                    chkStreamAgree.Enabled = False
                    cboStreamLayer.Enabled = False
                    cmdOptions.Enabled = False
                    g_boolHydCorr = False

            End Select

            CheckEnabled()

        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub cboDEMUnits_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDEMUnits.SelectedIndexChanged
        Try
            boolChange(2) = True
            CheckEnabled()
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub cboSubWSSize_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSubWSSize.SelectedIndexChanged
        Try
            boolChange(3) = True
            CheckEnabled()
            _intSize = cboSubWSSize.SelectedIndex
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub cmdQuit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdQuit.Click
        Try
            Dim intvbYesNo As Short

            intvbYesNo = MsgBox("Are you sure you want to exit?  Your changes will not be saved.", MsgBoxStyle.YesNo, "Exit")

            If intvbYesNo = MsgBoxResult.Yes Then
                If Not _frmPrj Is Nothing Then
                    _frmPrj.cboPrecipScen.SelectedIndex = 0
                End If
                Close()
            Else
                Exit Sub
            End If

            _booProject = False


        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try

    End Sub


    Private Sub cmdCreate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCreate.Click
        Try
            If _InputDEMPath = "" Then
                If txtDEMFile.Text <> "" Then
                    If IO.Directory.Exists(txtDEMFile.Text) Then
                        If IO.File.Exists(txtDEMFile.Text + "sta.adf") Then
                            _InputDEMPath = txtDEMFile.Text + "sta.adf"
                        Else
                            If IO.File.Exists(txtDEMFile.Text + IO.Path.DirectorySeparatorChar + "sta.adf") Then
                                _InputDEMPath = txtDEMFile.Text + IO.Path.DirectorySeparatorChar + "sta.adf"
                            End If
                        End If
                    Else
                        If IO.File.Exists(txtDEMFile.Text) Then
                            _InputDEMPath = txtDEMFile.Text
                        End If
                    End If
                Else
                End If
            End If

            If _InputDEMPath = "" Then
                MsgBox("The File you have choosen does not exist.", MsgBoxStyle.Critical, "File Not Found")
                txtDEMFile.Focus()
                Exit Sub
            End If

            Dim pRasterDataset As New MapWinGIS.Grid
            If Not pRasterDataset.Open(_InputDEMPath) Then
                MsgBox("The File you have choosen is not a raster.", MsgBoxStyle.Critical, "File Not Raster")
                txtDEMFile.Focus()
                Exit Sub
            End If

            Dim outpath As String
            If Not IO.Directory.Exists(g_nspectPath + "\wsdelin\" + txtWSDelinName.Text) Then
                outpath = g_nspectPath + "\wsdelin\" + txtWSDelinName.Text + "\"
                IO.Directory.CreateDirectory(outpath)
            Else
                MsgBox("Name in use.  Please select another.", MsgBoxStyle.Critical, "Choose New Name")
                txtWSDelinName.Focus()
                Exit Sub
            End If

            'Give the call; if successful insert new record
            If DelineateWatershed(pRasterDataset, outpath) Then
                'SQL Insert
                'DataBase Update
                'Compose the INSERT statement.
                Dim strCmdInsert As String = "INSERT INTO WSDelineation " & "(Name, DEMFileName, DEMGridUnits, FlowDirFileName, FlowAccumFileName," & "FilledDEMFileName, HydroCorrected, StreamFileName, SubWSSize, WSFileName, LSFileName, NibbleFileName, DEM2bFileName) " & " VALUES (" & "'" & CStr(txtWSDelinName.Text) & "', " & "'" & CStr(_InputDEMPath) & "', " & "'" & cboDEMUnits.SelectedIndex & "', " & "'" & _strDirFileName & "', " & "'" & _strAccumFileName & "', " & "'" & _strFilledDEMFileName & "', " & "'" & chkHydroCorr.CheckState & "', " & "'" & _strStreamLayer & "', " & "'" & cboSubWSSize.SelectedIndex & "', " & "'" & _strWShedFileName & "', " & "'" & _strLSFileName & "', " & "'" & _strNibbleName & "', " & "'" & _strDEM2BName & "')"

                'Execute the statement.
                Dim insCmd As New DataHelper(strCmdInsert)
                insCmd.ExecuteNonQuery()

                'Confirm
                MsgBox(txtWSDelinName.Text & " successfully added.", MsgBoxStyle.OkOnly, "Record Added")

                If g_boolNewWShed Then
                    'frmPrj.Show
                    _frmPrj.Visible = True
                    _frmPrj.cboWSDelin.Items.Clear()
                    modUtil.InitComboBox((_frmPrj.cboWSDelin), "WSDelineation")
                    _frmPrj.cboWSDelin.SelectedIndex = modUtil.GetCboIndex((txtWSDelinName.Text), (_frmPrj.cboWSDelin))
                    Close()
                Else
                    Close()
                    _frmWS.Close()
                End If
            Else
                Exit Sub
            End If

            'Reset project boolean
            _booProject = False
        Catch ex As Exception
            MsgBox(Err.Number & Err.Description & " :New Watershed Delineation")
        End Try

    End Sub

#End Region

#Region "Helper Functions"
    Public Sub Init(ByRef frmWS As WatershedDelineationsForm, ByRef frmPrj As MainForm)
        _frmWS = frmWS
        _frmPrj = frmPrj
    End Sub


    Public Sub CheckEnabled()
        Try

            If chkHydroCorr.CheckState And chkStreamAgree.CheckState Then
                If boolChange(0) And boolChange(1) And boolChange(2) And boolChange(3) And g_boolParams Then
                    cmdCreate.Enabled = True
                Else
                    cmdCreate.Enabled = False
                End If
            ElseIf chkHydroCorr.CheckState And Not chkStreamAgree.CheckState Then
                If boolChange(0) And boolChange(1) And boolChange(2) And boolChange(3) Then
                    cmdCreate.Enabled = True
                Else
                    cmdCreate.Enabled = False
                End If
            ElseIf Not chkHydroCorr.CheckState Then
                If boolChange(0) And boolChange(1) And boolChange(2) And boolChange(3) Then
                    cmdCreate.Enabled = True
                Else
                    cmdCreate.Enabled = False
                End If
            End If

        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Function DelineateWatershed(ByRef pSurfaceDatasetIn As MapWinGIS.Grid, ByVal OutPath As String) As Boolean
        'Declare the raster objects
        Dim pFlowDirRaster As New MapWinGIS.Grid 'Flow Direction
        Dim pAccumRaster As New MapWinGIS.Grid 'Flow Accumulation
        Dim pFillRaster As New MapWinGIS.Grid 'Fill GDS
        Dim pWShedRaster As New MapWinGIS.Grid 'WaterShed GDS
        Dim pBasinRaster As New MapWinGIS.Grid 'Basin GDS
        Dim pOutputFeatClass As New MapWinGIS.Shapefile 'output basins
        Dim strahlordout As String = ""
        Dim longestupslopeout As String = ""
        Dim totalupslopeout As String = ""
        Dim streamgridout As String = ""
        Dim streamordout As String = ""
        Dim treedatout As String = ""
        Dim coorddatout As String = ""
        Dim strWSGridOut As String = ""
        Dim strWSSFOut As String = ""

        'Featureclass objects
        Dim pBasinFeatClass As New MapWinGIS.Shapefile 'Basin Featureclass

        Try
            Dim strProgTitle As String = "Watershed Delineation Processing..."

            'Have to hide the existing forms, to be able to show the progress
            If g_boolNewWShed Then
                'Me.Hide()
                '_frmPrj.Hide()
            ElseIf Not g_boolNewWShed Then
                'Me.Hide()
                '_frmWS.Hide()
            End If


            'Get cell size from DEM; needed later in the Length Slope Calculation
            _intCellSize = pSurfaceDatasetIn.Header.dX

            Dim ret As Integer

            'STEP 1:  Fill the Surface
            'if hydrocorrect, then skip the Fill, just use the incoming DEM
            modProgDialog.ShowProgress("Filling DEM...", strProgTitle, 0, 10, 1, Me)
            If chkHydroCorr.CheckState = 1 Then
                pFillRaster = pSurfaceDatasetIn
                _strFilledDEMFileName = pSurfaceDatasetIn.Filename
            Else
                'Call to ProgDialog to use throughout process: keep user informed.

                If modProgDialog.g_KeepRunning Then
                    _strFilledDEMFileName = OutPath + "demfill" + g_OutputGridExt
                    MapWinGeoProc.Hydrology.Fill(pSurfaceDatasetIn.Filename, _strFilledDEMFileName, False)
                    pFillRaster.Open(_strFilledDEMFileName)
                Else
                    Return False
                End If
            End If 'End if Fill


            'STEP 2: Flow Direction
            modProgDialog.ShowProgress("Computing Flow Direction...", strProgTitle, 0, 10, 2, Me)

            Dim mwDirFileName As String = OutPath + "mwflowdir" + g_OutputGridExt
            _strDirFileName = OutPath + "flowdir" + g_OutputGridExt
            Dim strSlpFileName As String = OutPath + "slope" + g_OutputGridExt
            If modProgDialog.g_KeepRunning Then
                ret = MapWinGeoProc.Hydrology.D8(pFillRaster.Filename, mwDirFileName, strSlpFileName, Environment.ProcessorCount, Nothing)
                If ret <> 0 Then Return False
                pFlowDirRaster.Open(mwDirFileName)

                Dim pESRID8Flow As New MapWinGIS.Grid
                Dim tmphead As New MapWinGIS.GridHeader
                tmphead.CopyFrom(pFlowDirRaster.Header)
                pESRID8Flow.CreateNew(_strDirFileName, tmphead, MapWinGIS.GridDataType.FloatDataType, -1)
                Dim tauD8ToESRIcalc As New RasterMathCellCalcNulls(AddressOf tauD8ToESRICellCalc)
                RasterMath(pFlowDirRaster, Nothing, Nothing, Nothing, Nothing, pESRID8Flow, Nothing, False, tauD8ToESRIcalc)
                pESRID8Flow.Header.NodataValue = -1
                pESRID8Flow.Save(_strDirFileName)
                pESRID8Flow.Close()
            Else
                Return False
            End If

            'STEP 3: Flow Accumulation
            modProgDialog.ShowProgress("Computing Flow Accumulation...", strProgTitle, 0, 10, 3, Me)
            _strAccumFileName = OutPath + "flowacc" + g_OutputGridExt
            If modProgDialog.g_KeepRunning Then
                ret = MapWinGeoProc.Hydrology.AreaD8(pFlowDirRaster.Filename, "", _strAccumFileName, False, False, Environment.ProcessorCount, Nothing)
                If ret <> 0 Then Return False
                pAccumRaster.Open(_strAccumFileName)
            Else
                Return False
            End If

            '        'STEP 4: Stream Layer
            '        'How it's done:
            '        '1. Get the Stat.Max of the Accum GRID
            '        '2. Multiply the Max * percentage as deemed by user: small = 3% med = 6% large = 10%

            Dim dblMax As Double = pAccumRaster.Maximum
            Dim dblSubShedSize As Double

            Select Case _intSize
                Case 0 'small
                    dblSubShedSize = dblMax * _dblSmall
                Case 1 'medium
                    dblSubShedSize = dblMax * _dblMedium
                Case 2 'large
                    dblSubShedSize = dblMax * _dblLarge
            End Select


            strahlordout = OutPath + "strahlord" + g_OutputGridExt
            longestupslopeout = OutPath + "longestupslope" + g_OutputGridExt
            totalupslopeout = OutPath + "totalupslope" + g_OutputGridExt
            streamgridout = OutPath + "streamgrid" + g_OutputGridExt
            streamordout = OutPath + "streamord" + g_OutputGridExt
            treedatout = OutPath + "tree.dat"
            coorddatout = OutPath + "coord.dat"
            strWSGridOut = OutPath + "wsgrid" + g_OutputGridExt
            strWSSFOut = OutPath + "ws.shp"

            '        'Step 5: Using Hydrology Op to create stream network
            _strStreamLayer = OutPath + "stream.shp"
            modProgDialog.ShowProgress("Creating Stream Network...", strProgTitle, 0, 10, 4, Me)
            If modProgDialog.g_KeepRunning Then
                ret = MapWinGeoProc.Hydrology.DelinStreamGrids(
       pSurfaceDatasetIn.Filename, _
             pFillRaster.Filename, _
             pFlowDirRaster.Filename, _
             strSlpFileName, _
             pAccumRaster.Filename, _
             "", _
             strahlordout, _
             longestupslopeout, _
             totalupslopeout, _
             streamgridout, _
             streamordout, _
             treedatout, _
             coorddatout, _
              _strStreamLayer, _
             strWSGridOut, _
             dblSubShedSize, _
             False, _
             False, _
             2, _
             Nothing)
                If ret <> 0 Then Return False
            Else
                Return False
            End If

            'The 14th and 15th parameters of DelinStreamGrids are the 4th and 5th parameters of the old function DelinStreamsAndSubBasins, and you do need them.
            '        'Step 6: Do WaterShed Op got moved into above step.
            '_strWShedFileName = OutPath + "basinpoly.shp"
            'modProgDialog.ProgDialog("Creating Watershed GRID...", strProgTitle, 0, 10, 5, Me)
            'If modProgDialog.g_boolCancel Then

            '    'HACK ret = MapWinGeoProc.Hydrology.DelinStreamsAndSubBasins(pFlowDirRaster.Filename, treedatout, coorddatout, _strStreamLayer, strWSGridOut, Nothing)
            '    If ret <> 0 Then Return False
            'Else
            '    Return False
            'End If

            modProgDialog.ShowProgress("Creating Watershed Shape...", strProgTitle, 0, 10, 7, Me)
            If modProgDialog.g_KeepRunning Then
                ret = MapWinGeoProc.Hydrology.SubbasinsToShape(pFlowDirRaster.Filename, strWSGridOut, strWSSFOut, Nothing)
                If ret <> 0 Then Return False
            Else
                Return False
            End If

            modProgDialog.ShowProgress("Removing Small Polygons...", strProgTitle, 0, 10, 9, Me)
            If modProgDialog.g_KeepRunning Then
                pBasinFeatClass.Open(strWSSFOut)

                pOutputFeatClass = RemoveSmallPolys(pBasinFeatClass, pFillRaster)
            Else
                Return False
            End If

            'Save final output tif versions (since Taudem needed bgds before this)
            Dim proj As String = pOutputFeatClass.Projection
            Dim tmpfile As String
            tmpfile = _strFilledDEMFileName
            _strFilledDEMFileName = OutPath + "demfill" + g_FinalOutputGridExt
            pFillRaster.Close()
            pFillRaster = New MapWinGIS.Grid
            pFillRaster.Open(tmpfile)
            pFillRaster.Header.Projection = proj
            pFillRaster.Save(_strFilledDEMFileName)
            pFillRaster.Close()
            IO.File.Delete(tmpfile)

            tmpfile = strSlpFileName
            strSlpFileName = OutPath + "slope" + g_FinalOutputGridExt
            Dim pslope As New MapWinGIS.Grid
            pslope.Open(tmpfile)
            pslope.Header.Projection = proj
            pslope.Save(strSlpFileName)
            pslope.Close()
            IO.File.Delete(tmpfile)

            tmpfile = _strDirFileName
            _strDirFileName = OutPath + "flowdir" + g_FinalOutputGridExt
            pFlowDirRaster.Close()
            pFlowDirRaster = New MapWinGIS.Grid
            pFlowDirRaster.Open(tmpfile)
            pFlowDirRaster.Header.Projection = proj
            pFlowDirRaster.Save(_strDirFileName)
            pFlowDirRaster.Close()
            IO.File.Delete(tmpfile)

            tmpfile = _strAccumFileName
            _strAccumFileName = OutPath + "flowacc" + g_FinalOutputGridExt
            pAccumRaster.Close()
            pAccumRaster = New MapWinGIS.Grid
            pAccumRaster.Open(tmpfile)
            pAccumRaster.Header.Projection = proj
            pAccumRaster.Save(_strAccumFileName)
            pAccumRaster.Close()
            IO.File.Delete(tmpfile)

            'With all of that done, now go get the name of the LS Grid while actually computing said LS Grid
            '_strLSFileName = CalcLengthSlope(pFillRaster, pFlowDirRaster, pAccumRaster, pEnv, "0", pWorkspace)
            _strLSFileName = OutPath + "lsgrid" + g_FinalOutputGridExt
            Dim g As New MapWinGIS.Grid
            g.Open(longestupslopeout)
            g.Save(_strLSFileName)
            g.Close()
            _strNibbleName = _strDirFileName
            _strDEM2BName = pSurfaceDatasetIn.Filename
            'TODO: create these if really needed

            DelineateWatershed = True
        Catch ex As Exception
            If Err.Number = -2147217297 Then 'User cancelled operation
                modProgDialog.g_KeepRunning = False
                DelineateWatershed = False
            Else
                HandleError(c_sModuleFileName, ex)
                DelineateWatershed = False
            End If
        Finally
            modProgDialog.CloseDialog()
            pFlowDirRaster.Close()
            pAccumRaster.Close()
            pFillRaster.Close()
            pWShedRaster.Close()
            pBasinRaster.Close()
            pBasinFeatClass.Close()
            pOutputFeatClass.Close()
            MapWinGeoProc.DataManagement.DeleteGrid(strahlordout)
            MapWinGeoProc.DataManagement.DeleteGrid(longestupslopeout)
            MapWinGeoProc.DataManagement.DeleteGrid(totalupslopeout)
            MapWinGeoProc.DataManagement.DeleteGrid(streamgridout)
            MapWinGeoProc.DataManagement.DeleteGrid(streamordout)
            IO.File.Delete(treedatout)
            IO.File.Delete(coorddatout)
            MapWinGeoProc.DataManagement.DeleteGrid(strWSGridOut)
            MapWinGeoProc.DataManagement.DeleteShapefile(strWSSFOut)
        End Try
    End Function


    Private Function RemoveSmallPolys(ByRef pFeatureClass As MapWinGIS.Shapefile, ByRef pDEMRaster As MapWinGIS.Grid) As MapWinGIS.Shapefile
        Dim pMaskCalcRaster As New MapWinGIS.Grid
        Dim rastersf As New MapWinGIS.Shapefile
        Try
            'TODO: Find way in mapwindow to generate a decent outline of the grid
            '#1 First step, get a 1, 0 representation of the DEM, so we can get the outline
            'Dim maskcalc As New RasterMathCellCalc(AddressOf maskCellCalc)
            'RasterMath(pDEMRaster, Nothing, Nothing, Nothing, Nothing, pMaskCalcRaster, maskcalc)
            'strExpression = "Con([DEMo] >= 0, 1, 0)"


            '#2 init the rasterconverter and create the poly
            'Dim outMaskPath As String = IO.Path.GetTempFileName() + g_OutputGridExt
            'pMaskCalcRaster.Save(outMaskPath)
            'Dim outShapePath As String = IO.Path.GetTempFileName() + ".shp"


            'Dim outMaskPath2 As String = IO.Path.GetTempFileName() + g_OutputGridExt
            'Dim g As New MapWinGIS.Grid
            'Dim ghead As New MapWinGIS.GridHeader
            'ghead.CopyFrom(pMaskCalcRaster.Header)
            'g.CreateNew(outMaskPath2, ghead, MapWinGIS.GridDataType.ShortDataType, ghead.NodataValue)
            'Dim nr As Integer = ghead.NumberRows - 1
            'Dim nc As Integer = ghead.NumberCols - 1
            'Dim oldnull As Double = pMaskCalcRaster.Header.NodataValue
            'Dim rowsvals() As Single
            'ReDim rowsvals(nc)
            'For row As Integer = 0 To nr
            '    pMaskCalcRaster.GetRow(row, rowsvals(0))
            '    g.PutRow(row, rowsvals(0))
            'Next
            'pMaskCalcRaster.Close()
            'g.Save(outMaskPath)
            'Dim outdirPath As String = IO.Path.GetTempFileName() + "dir" + g_OutputGridExt
            'g.Save(outdirPath)
            'pMaskCalcRaster.Open(outMaskPath)

            'Dim u As New MapWinGIS.Utils
            'rastersf = u.GridToShapefile(pMaskCalcRaster, g)
            'rastersf.SaveAs(outShapePath)


            'MapWinGeoProc.FlowArea.SimpleRasterToPolygon(pMaskCalcRaster.Filename, outShapePath, Nothing)
            'rastersf.Open(outShapePath)
            'Dim u As New MapWinGIS.Utils
            'rastersf = u.GridToShapefile(pMaskCalcRaster)
            'rastersf.SaveAs(outShapePath)
            'MapWinGeoProc.Utils.GridToShapefile(pMaskCalcRaster.Filename, outShapePath)
            'rastersf.Open(outShapePath)

            '#3 determine size of 'small' watersheds, this is the area
            'Number of cells in the DEM area that are not null * a number dave came up with * CellSize Squared
            Dim dblArea As Double = ((pDEMRaster.Header.NumberCols * pDEMRaster.Header.NumberRows) * 0.004) * _intCellSize * _intCellSize

            '#4 Now with the Area of small sheds determined we can remove polygons that are too small.  To do this
            pFeatureClass.StartEditingShapes()
            For i As Integer = pFeatureClass.NumShapes - 1 To 0
                If pFeatureClass.Shape(i).Area() < dblArea Then
                    pFeatureClass.EditDeleteShape(i)
                End If
            Next
            pFeatureClass.StopEditingShapes()

            'TODO: Once we have an outline, union it, but for now just output the base with smalls removed
            '#5  Now, time to union the outline of the of the DEM with the newly paired down basin poly
            'Dim outputSf As MapWinGIS.Shapefile = rastersf.GetIntersection(False, pFeatureClass, False, pFeatureClass.ShapefileType, Nothing)
            'outputSf.SaveAs(_strWShedFileName)
            pFeatureClass.SaveAs(_strWShedFileName)
            pFeatureClass.Close()
            Dim outputSf As New MapWinGIS.Shapefile
            outputSf.Open(_strWShedFileName)
            Return outputSf

            'Dim currshape As MapWinGIS.Shape = rastersf.Shape(0)
            'For i As Integer = 0 To pFeatureClass.NumShapes() - 1
            '    currshape = MapWinGeoProc.SpatialOperations.Union(currshape, pFeatureClass.Shape(0))
            'Next
            'outputSf.CreateNew(_strWShedFileName, MapWinGIS.ShpfileType.SHP_POLYGON)
            'outputSf.StartEditingShapes()
            'outputSf.EditInsertShape(currshape, 0)
            'outputSf.StopEditingTable()
            'Return outputSf
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
            Return Nothing
        Finally
            pMaskCalcRaster.Close()
            pFeatureClass.Close()
            rastersf.Close()
        End Try
    End Function


    'Private Function CalcLengthSlope(ByRef pDEMRaster As MapWinGIS.Grid, ByRef pFlowDirRaster As MapWinGIS.Grid, ByRef pAccumRaster As MapWinGIS.Grid, ByRef strUnits As String) As String
    '    Try
    '        '        'From the Delineate watershed, we are garnering the DEM, Flow Direction, flowaccum, environment and units
    '        '        'to be used in this function
    '        '        'Returns: Name of the LS File for use in the database table
    '        '        On Error GoTo ErrHandler

    '        '        'The rasters                                'Associated Steps
    '        '        Dim pDEMOneCell As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 1: 1 Cell Buffer DEM
    '        '        Dim pDEMTwoCell As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 2: 2 Cell Buffer DEM
    '        '        Dim pFlowDir As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 3: Flow Direction
    '        '        Dim pFlowDirBV As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 4: Pre Nibble Null
    '        '        Dim pMask As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 4: Mask
    '        '        Dim pNibble As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 4: Nibble
    '        '        Dim pDownSlope As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 5: Down Slope
    '        '        Dim pDownAngle As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 5a: Tweak down slope
    '        '        Dim pRelativeSlope As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 6: Relative Slope
    '        '        Dim pRelSlopeThreshold As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 7: Relative slope threshold
    '        '        Dim pSlopeBreak As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 7a: Slope Break
    '        '        Dim pFlowDirBreak As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 8: Flow Direction Break
    '        '        Dim pWeight As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 9: Weight GRID
    '        '        Dim pFlowLength As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 10: Flow Length
    '        '        Dim pFlowLengthFt As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 11: Flow Length to Feet
    '        '        Dim pSlopeExp As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 12: Slope Exponent
    '        '        Dim pRusleLFactor As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 13: Rusle L Factor
    '        '        Dim pRusleSFactor As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 14: Rusle S Factor
    '        '        Dim pLSFactor As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 15: LS Factor
    '        '        Dim pFinalLS As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 15a: clippage
    '        '        Dim strProgTitle As String

    '        '        'Analysis Environment
    '        '        Dim pDEMGeoDS As ESRI.ArcGIS.Geodatabase.IGeoDataset 'GeoDataset to get spat. ref
    '        '        Dim pSpatRef As ESRI.ArcGIS.Geometry.ISpatialReference 'Spatial Reference

    '        '        'Create Map Algebra Operator
    '        '        Dim pMapAlgebraOp As ESRI.ArcGIS.SpatialAnalyst.IMapAlgebraOp 'Workhorse
    '        '        'String to hold calculations
    '        '        Dim strExpression As String

    '        '        pDEMGeoDS = pDEMRaster
    '        '        pSpatRef = pDEMGeoDS.SpatialReference

    '        '        If Not modUtil.CheckSpatialAnalystLicense Then
    '        '            MsgBox("No Spatial Analyst License Available.", MsgBoxStyle.Critical, "Pay Your Licensing Fee")
    '        '            CalcLengthSlope = "ERROR"
    '        '            Exit Function
    '        '        End If

    '        '        If pDEMRaster Is Nothing Or pFlowDirRaster Is Nothing Then
    '        '            MsgBox("CaclLengthSlope Error.")
    '        '            CalcLengthSlope = "ERROR"
    '        '        End If

    '        '        'Initialize the Map AlgebraOp, same thing as the Map Calculator
    '        '        'All of the following steps use the same methodology.  They take a Raster, bind it
    '        '        'to a symbol, in this case a string that represents them in some map calculation.
    '        '        'You then simply use the MapAlgebraOp to execute the expression.

    '        '        'Get a hold of the Spatial Reference of the DEM, and init MapAlgebra
    '        '        pMapAlgebraOp = New ESRI.ArcGIS.SpatialAnalyst.RasterMapAlgebraOp

    '        '        'Set the environment
    '        '        pEnv = pMapAlgebraOp

    '        '        strProgTitle = "Processing the LS GRID..."

    '        '        'STEP 1: ----------------------------------------------------------------------
    '        '        'Buffer the DEM by one cell
    '        '        modProgDialog.ProgDialog("Creating one cell buffer...", strProgTitle, 0, 15, 1, (m_App.hwnd))

    '        '        'UPGRADE_NOTE: Object pDEMOneCell may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '        '        pDEMOneCell = Nothing

    '        '        If modProgDialog.g_boolCancel Then

    '        '            'UPGRADE_WARNING: Couldn't resolve default property of object pDEMRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        '            pMapAlgebraOp.BindRaster(pDEMRaster, "aml_fdem")
    '        '            strExpression = "Con(isnull([aml_fdem]), focalmin([aml_fdem]), [aml_fdem])"

    '        '            pDEMOneCell = pMapAlgebraOp.Execute(strExpression)

    '        '            pMapAlgebraOp.UnbindRaster("aml_fdem")

    '        '        Else
    '        '            GoTo ProgCancel
    '        '        End If

    '        '        'END STEP 1: ------------------------------------------------------------------

    '        '        'STEP 2: ----------------------------------------------------------------------
    '        '        'Buffer the DEM buffer by one more cell, that's 2
    '        '        modProgDialog.ProgDialog("Creating two cell buffer...", strProgTitle, 0, 15, 2, (m_App.hwnd))

    '        '        If modProgDialog.g_boolCancel Then

    '        '            'UPGRADE_WARNING: Couldn't resolve default property of object pDEMOneCell. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        '            pMapAlgebraOp.BindRaster(pDEMOneCell, "dem_b")
    '        '            strExpression = "Con(isnull([dem_b]), focalmin([dem_b]), [dem_b])"

    '        '            pDEMTwoCell = pMapAlgebraOp.Execute(strExpression)

    '        '            m_strDEM2BName = modUtil.MakePerminentGrid(pDEMTwoCell, (pWS.PathName), "dem2b")

    '        '            pMapAlgebraOp.UnbindRaster("dem_b")

    '        '        Else
    '        '            GoTo ProgCancel
    '        '        End If
    '        '        'END STEP 2: ------------------------------------------------------------------

    '        '        'STEP 3: ----------------------------------------------------------------------
    '        '        'Flow Direction

    '        '        pFlowDir = pFlowDirRaster

    '        '        'END STEP 3: ------------------------------------------------------------------


    '        '        'STEP 3a: ---------------------------------------------------------------------
    '        '        modProgDialog.ProgDialog("Creating mask...", strProgTitle, 0, 15, 3, (m_App.hwnd))

    '        '        If modProgDialog.g_boolCancel Then

    '        '            With pMapAlgebraOp
    '        '                'UPGRADE_WARNING: Couldn't resolve default property of object pDEMOneCell. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        '                .BindRaster(pDEMOneCell, "mask")
    '        '            End With

    '        '            'strExpression = "con(isnull([mask]),0,1)"
    '        '            strExpression = "con([mask] >= 0, 1, 0)"


    '        '            pMask = pMapAlgebraOp.Execute(strExpression)

    '        '            With pEnv
    '        '                .Mask = pMask
    '        '                'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        '                .OutWorkspace = pWS
    '        '            End With

    '        '            pEnv = pMapAlgebraOp
    '        '        Else
    '        '            GoTo ProgCancel
    '        '        End If

    '        '        'STEP 4: ----------------------------------------------------------------------
    '        '        'Buffering Flow Direction and do the nibble to fill it in
    '        '        'Needed in case there is outflow from the DEM grid.
    '        '        'The following algorithms need to access the downslope DEM grid cell.
    '        '        'We find this by nibbling the original flow direction grid instead of recalculating
    '        '        'flow direction from the nibbled DEM because we want the elevation that is assumed
    '        '        'to be downstream of the edge cell.  Using flow direction on the buffered DEM
    '        '        'may not give that same result.
    '        '        modProgDialog.ProgDialog("Buffering slope direction...", strProgTitle, 0, 15, 4, (m_App.hwnd))

    '        '        If modProgDialog.g_boolCancel Then

    '        '            With pMapAlgebraOp
    '        '                'UPGRADE_WARNING: Couldn't resolve default property of object pFlowDir. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        '                .BindRaster(pFlowDir, "fdr_b")
    '        '            End With
    '        '            strExpression = "con(isnull([fdr_b]),0,[fdr_b])"

    '        '            pFlowDirBV = pMapAlgebraOp.Execute(strExpression)

    '        '            pMapAlgebraOp.UnbindRaster("fdr_b")

    '        '            'Nibble
    '        '            'UPGRADE_WARNING: Couldn't resolve default property of object pFlowDirBV. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        '            pMapAlgebraOp.BindRaster(pFlowDirBV, "fdr_bv")
    '        '            'UPGRADE_WARNING: Couldn't resolve default property of object pMask. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        '            pMapAlgebraOp.BindRaster(pMask, "waia_reg")
    '        '            strExpression = "nibble([fdr_bv],[waia_reg], dataonly)"

    '        '            pNibble = pMapAlgebraOp.Execute(strExpression)

    '        '            'Get nibble's path for use in the database
    '        '            m_strNibbleName = modUtil.MakePerminentGrid(pNibble, (pWS.PathName), "nibble")

    '        '            With pMapAlgebraOp
    '        '                .UnbindRaster("fdr_bv")
    '        '                .UnbindRaster("waia_reg")
    '        '            End With
    '        '        Else
    '        '            GoTo ProgCancel
    '        '        End If
    '        '        'END STEP 4: ------------------------------------------------------------------

    '        '        'STEP 5: ----------------------------------------------------------------------
    '        '        'Calculate Slope
    '        '        'Actually this is calculating the SLOPE, in degrees, not the slope change.
    '        '        'Note that ESRI's SLOPE command should not be used here.  That command fits
    '        '        'a plane to the 3x3 grid surrounding the central point and assigns the central
    '        '        'point the slope of that plane.  The algorithm used here calculates only the slope
    '        '        'between the central point and it's immediate downstream neighbor.
    '        '        'That is what is needed by RUSLE.

    '        '        modProgDialog.ProgDialog("Calculating Slope change...", strProgTitle, 0, 15, 5, (m_App.hwnd))

    '        '        If modProgDialog.g_boolCancel Then

    '        '            With pMapAlgebraOp
    '        '                'UPGRADE_WARNING: Couldn't resolve default property of object pNibble. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        '                .BindRaster(pNibble, "fdrnib")
    '        '                'UPGRADE_WARNING: Couldn't resolve default property of object pDEMTwoCell. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        '                .BindRaster(pDEMTwoCell, "dem_2b")
    '        '            End With
    '        '            strExpression = "Con(([fdrnib] ge 0.5 and [fdrnib] lt 1.5), deg * atan(([dem_2b] - [dem_2b](1,0)) / (" & m_intCellSize & "))," & "Con(([fdrnib] ge 1.5 and [fdrnib] lt 3.0), deg * atan(([dem_2b] - [dem_2b](1,1)) / (" & m_intCellSize & " * 1.4142))," & "Con(([fdrnib] ge 3.0 and [fdrnib] lt 6.0), deg * atan(([dem_2b] - [dem_2b](0,1)) / (" & m_intCellSize & "))," & "Con(([fdrnib] ge 6.0 and [fdrnib] lt 12.0), deg * atan(([dem_2b] - [dem_2b](-1,1)) / (" & m_intCellSize & " * 1.4142))," & "Con(([fdrnib] ge 12.0 and [fdrnib] lt 24.0), deg * atan(([dem_2b] - [dem_2b](-1,0)) / (" & m_intCellSize & "))," & "Con(([fdrnib] ge 24.0 and [fdrnib] lt 48.0), deg * atan(([dem_2b] - [dem_2b](-1,-1)) / (" & m_intCellSize & " * 1.4142))," & "Con(([fdrnib] ge 48.0 and [fdrnib] lt 96.0), deg * atan(([dem_2b] - [dem_2b](0,-1)) / (" & m_intCellSize & "))," & "Con(([fdrnib] ge 96.0 and [fdrnib] lt 192.0), deg * atan(([dem_2b] - [dem_2b](1,-1)) / (" & m_intCellSize & " * 1.4142))," & "Con(([fdrnib] ge 192.0 and [fdrnib] le 255.0), deg * atan(([dem_2b] - [dem_2b](1,0)) / (" & m_intCellSize & "))," & "0.1 )))))))))"

    '        '            pDownSlope = pMapAlgebraOp.Execute(strExpression)
    '        '            'Cleanup
    '        '            With pMapAlgebraOp
    '        '                .UnbindRaster("fdrnib")
    '        '                .UnbindRaster("dem_2b")
    '        '            End With
    '        '        Else
    '        '            GoTo ProgCancel
    '        '        End If
    '        '        'END STEP 5 ---------------------------------------------------------------------

    '        '        'STEP 5a: -----------------------------------------------------------------------
    '        '        'Tweak slope where it equals 0 to 0.1
    '        '        'UPGRADE_WARNING: Couldn't resolve default property of object pDownSlope. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        '        pMapAlgebraOp.BindRaster(pDownSlope, "dwnsltmp")
    '        '        strExpression = "Con([dwnsltmp] le 0, 0.1,[dwnsltmp])"

    '        '        pDownAngle = pMapAlgebraOp.Execute(strExpression)

    '        '        pMapAlgebraOp.UnbindRaster("dwnsltmp")

    '        '        'END STEP 5a: -------------------------------------------------------------------

    '        '        'STEP 6: ------------------------------------------------------------------------
    '        '        'Relative Slope Change
    '        '        modProgDialog.ProgDialog("Calculating Relative Slope Change...", strProgTitle, 0, 15, 6, (m_App.hwnd))

    '        '        If modProgDialog.g_boolCancel Then
    '        '            With pMapAlgebraOp
    '        '                'UPGRADE_WARNING: Couldn't resolve default property of object pDownAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        '                .BindRaster(pDownAngle, "dwnslangle")
    '        '                'UPGRADE_WARNING: Couldn't resolve default property of object pNibble. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        '                .BindRaster(pNibble, "fdrnib")
    '        '            End With
    '        '            strExpression = "Con([fdrnib] == 1, ([dwnslangle] - [dwnslangle](1,0)) / [dwnslangle]," & "Con([fdrnib] == 2, ([dwnslangle] - [dwnslangle](1,1)) / [dwnslangle]," & "Con([fdrnib] == 4, ([dwnslangle] - [dwnslangle](0,1)) / [dwnslangle]," & "Con([fdrnib] == 8, ([dwnslangle] - [dwnslangle](-1,1)) / [dwnslangle]," & "Con([fdrnib] == 16, ([dwnslangle] - [dwnslangle](-1,0)) / [dwnslangle]," & "Con([fdrnib] == 32, ([dwnslangle] - [dwnslangle](-1,-1)) / [dwnslangle]," & "Con([fdrnib] == 64, ([dwnslangle] - [dwnslangle](0,-1)) / [dwnslangle]," & "Con([fdrnib] == 128, ([dwnslangle] - [dwnslangle](1,-1)) / [dwnslangle]," & "Con([fdrnib] == 255, ([dwnslangle] - [dwnslangle](1,0)) / [dwnslangle]," & "0.1 )))))))))"

    '        '            pRelativeSlope = pMapAlgebraOp.Execute(strExpression)

    '        '            With pMapAlgebraOp
    '        '                .UnbindRaster("dwnslangle")
    '        '                .UnbindRaster("fdrnib")
    '        '            End With
    '        '        Else
    '        '            GoTo ProgCancel
    '        '        End If
    '        '        'END STEP 6 -----------------------------------------------------------------------

    '        '        'STEP 7: --------------------------------------------------------------------------
    '        '        'Identify breakpoints: relative difference where slope angle exceeds threshold values
    '        '        modProgDialog.ProgDialog("Identifying breakpoints...", strProgTitle, 0, 15, 7, (m_App.hwnd))

    '        '        If modProgDialog.g_boolCancel Then

    '        '            'UPGRADE_WARNING: Couldn't resolve default property of object pDownAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        '            pMapAlgebraOp.BindRaster(pDownAngle, "dwnslangle")
    '        '            strExpression = "Con(([dwnslangle] gt 2.86240), 0.5, Con(([dwnslangle] le 2.86240), 0.7, 0.0 ))"

    '        '            pRelSlopeThreshold = pMapAlgebraOp.Execute(strExpression)

    '        '            pMapAlgebraOp.UnbindRaster("dwnslangle")
    '        '            'END STEP 7 -----------------------------------------------------------------------

    '        '            'STEP 7a: -------------------------------------------------------------------------
    '        '            'Slope break
    '        '            With pMapAlgebraOp
    '        '                'UPGRADE_WARNING: Couldn't resolve default property of object pRelSlopeThreshold. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        '                .BindRaster(pRelSlopeThreshold, "threshold")
    '        '                'UPGRADE_WARNING: Couldn't resolve default property of object pRelativeSlope. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        '                .BindRaster(pRelativeSlope, "delslprel")
    '        '            End With
    '        '            strExpression = "Con(([delslprel] gt [threshold]), 1, 0 )"

    '        '            pSlopeBreak = pMapAlgebraOp.Execute(strExpression)

    '        '            With pMapAlgebraOp
    '        '                .UnbindRaster("threshold")
    '        '                .UnbindRaster("delslprel")
    '        '            End With
    '        '        Else
    '        '            GoTo ProgCancel
    '        '        End If
    '        '        'END STEP 7a -------------------------------------------------------------------------

    '        '        'STEP 8 -------------------------------------------------------------------------------
    '        '        'Create Modified Flow Direction GRID
    '        '        modProgDialog.ProgDialog("Creating modified flow direction GRID...", strProgTitle, 0, 15, 8, (m_App.hwnd))

    '        '        If modProgDialog.g_boolCancel Then

    '        '            With pMapAlgebraOp
    '        '                'UPGRADE_WARNING: Couldn't resolve default property of object pSlopeBreak. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        '                .BindRaster(pSlopeBreak, "slopebreak")
    '        '                'UPGRADE_WARNING: Couldn't resolve default property of object pFlowDir. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        '                .BindRaster(pFlowDir, "fdr")
    '        '            End With
    '        '            strExpression = "Con([slopebreak] eq 0, [fdr], 0)"

    '        '            pFlowDirBreak = pMapAlgebraOp.Execute(strExpression)

    '        '            With pMapAlgebraOp
    '        '                .UnbindRaster("slopebreak")
    '        '                .UnbindRaster("fdr")
    '        '            End With
    '        '        Else
    '        '            GoTo ProgCancel
    '        '        End If
    '        '        'END STEP 8: ----------------------------------------------------------------------

    '        '        'STEP 9: --------------------------------------------------------------------------
    '        '        'Create weight grid
    '        '        modProgDialog.ProgDialog("Creating weight GRID...", strProgTitle, 0, 15, 9, (m_App.hwnd))

    '        '        If modProgDialog.g_boolCancel Then


    '        '            'Dave's comments:
    '        '            'This is an error.  In the original AML code, it is needed to correctly account for
    '        '            'diagonal flow in the flow length calculation of step 10.  However, ArcMap already
    '        '            'makes this correction.  There is another weighting function needed, however, to be
    '        '            'consistent with the procedure used in the original AML code. That is what should replace this con.

    '        '            'Removed 12/19/07
    '        '            'pMapAlgebraOp.BindRaster pFlowDir, "fdr"
    '        '            'strExpression = "Con([fdr] eq 2, 1.41421," & _
    '        '            ''    "Con([fdr] eq 8, 1.41421," & _
    '        '            ''    "Con([fdr] eq 32, 1.41421," & _
    '        '            ''    "Con([fdr] eq 128, 1.41421," & _
    '        '            ''    "1.0))))"
    '        '            'End Remove

    '        '            'Added 12/19/07
    '        '            'pAccumRaster is passed over from delin watershed
    '        '            'UPGRADE_WARNING: Couldn't resolve default property of object pAccumRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        '            pMapAlgebraOp.BindRaster(pAccumRaster, "flowacc")
    '        '            strExpression = "Con([flowacc] eq 0, 0.5,1.0)"

    '        '            pWeight = pMapAlgebraOp.Execute(strExpression)

    '        '            pMapAlgebraOp.UnbindRaster("flowacc")
    '        '            'End Added
    '        '        Else
    '        '            GoTo ProgCancel
    '        '        End If
    '        '        'END STEP 9 -------------------------------------------------------------------------

    '        '        'STEP 10: ---------------------------------------------------------------------------
    '        '        'Flow Length GRID
    '        '        modProgDialog.ProgDialog("Creating flow length GRID...", strProgTitle, 0, 15, 10, (m_App.hwnd))

    '        '        If modProgDialog.g_boolCancel Then

    '        '            With pMapAlgebraOp
    '        '                'UPGRADE_WARNING: Couldn't resolve default property of object pFlowDirBreak. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        '                .BindRaster(pFlowDirBreak, "fdrbrk")
    '        '                'UPGRADE_WARNING: Couldn't resolve default property of object pWeight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        '                .BindRaster(pWeight, "weight")
    '        '            End With
    '        '            strExpression = "FlowLength([fdrbrk], [weight], UPSTREAM)"

    '        '            pFlowLength = pMapAlgebraOp.Execute(strExpression)

    '        '            With pMapAlgebraOp
    '        '                .UnbindRaster("fdrbrk")
    '        '                .UnbindRaster("weight")
    '        '            End With
    '        '        Else
    '        '            GoTo ProgCancel
    '        '        End If
    '        '        'END STEP 10: -----------------------------------------------------------------------

    '        '        'STEP 11: ---------------------------------------------------------------------------
    '        '        'Convert Meters To Feet
    '        '        'TODO: Check measure units, won't have to do if already in Feet
    '        '        modProgDialog.ProgDialog("Checking measurement units...", strProgTitle, 0, 15, 11, (m_App.hwnd))

    '        '        If modProgDialog.g_boolCancel Then
    '        '            'UPGRADE_WARNING: Couldn't resolve default property of object pFlowLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        '            pMapAlgebraOp.BindRaster(pFlowLength, "flowlen")
    '        '            strExpression = "[flowlen] / 0.3048"

    '        '            pFlowLengthFt = pMapAlgebraOp.Execute(strExpression)

    '        '            pMapAlgebraOp.UnbindRaster("flowlen")
    '        '        Else
    '        '            GoTo ProgCancel
    '        '        End If
    '        '        'END STEP 11: -----------------------------------------------------------------------

    '        '        'STEP 12: ---------------------------------------------------------------------------
    '        '        'Calculate the slope length exponent value 'M'
    '        '        modProgDialog.ProgDialog("Calculating slope length exponent value 'M'...", strProgTitle, 0, 15, 12, (m_App.hwnd))

    '        '        If modProgDialog.g_boolCancel Then

    '        '            'UPGRADE_WARNING: Couldn't resolve default property of object pDownAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        '            pMapAlgebraOp.BindRaster(pDownAngle, "dwnslangle")
    '        '            strExpression = "Con(([dwnslangle] le 0.1), 0.01," & "Con(([dwnslangle] gt 0.1 and [dwnslangle] lt 0.2), 0.02," & "Con(([dwnslangle] ge 0.2 and [dwnslangle] lt 0.4), 0.04," & "Con(([dwnslangle] ge 0.4 and [dwnslangle] lt 0.85), 0.08," & "Con(([dwnslangle] ge 0.85 and [dwnslangle] lt 1.4), 0.14," & "Con(([dwnslangle] ge 1.4 and [dwnslangle] lt 2.0), 0.18," & "Con(([dwnslangle] ge 2.0 and [dwnslangle] lt 2.6), 0.22," & "Con(([dwnslangle] ge 2.6 and [dwnslangle] lt 3.1), 0.25," & "Con(([dwnslangle] ge 3.1 and [dwnslangle] lt 3.7), 0.28," & "Con(([dwnslangle] ge 3.7 and [dwnslangle] lt 5.2), 0.32," & "Con(([dwnslangle] ge 5.2 and [dwnslangle] lt 6.3), 0.35," & "Con(([dwnslangle] ge 6.3 and [dwnslangle] lt 7.4), 0.37," & "Con(([dwnslangle] ge 7.4 and [dwnslangle] lt 8.6), 0.40," & "Con(([dwnslangle] ge 8.6 and [dwnslangle] lt 10.3), 0.41," & "Con(([dwnslangle] ge 10.3 and [dwnslangle] lt 12.9), 0.44," & "Con(([dwnslangle] ge 12.9 and [dwnslangle] lt 15.7), 0.47," & "Con(([dwnslangle] ge 15.7 and [dwnslangle] lt 20.0), 0.49," & "Con(([dwnslangle] ge 20.0 and [dwnslangle] lt 25.8), 0.52," & "Con(([dwnslangle] ge 25.8 and [dwnslangle] lt 31.5), 0.54," & "Con(([dwnslangle] ge 31.5 and [dwnslangle] lt 37.2), 0.55," & "0.56))))))))))))))))))))"

    '        '            pSlopeExp = pMapAlgebraOp.Execute(strExpression)

    '        '            pMapAlgebraOp.UnbindRaster("dwnslangle")
    '        '        Else
    '        '            GoTo ProgCancel
    '        '        End If
    '        '        'END STEP 12: ------------------------------------------------------------------------

    '        '        'STEP 13: ----------------------------------------------------------------------------
    '        '        'Calculate the L-Factor
    '        '        modProgDialog.ProgDialog("Calculating the L factor...", strProgTitle, 0, 15, 13, (m_App.hwnd))

    '        '        If modProgDialog.g_boolCancel Then

    '        '            With pMapAlgebraOp
    '        '                'UPGRADE_WARNING: Couldn't resolve default property of object pFlowLengthFt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        '                .BindRaster(pFlowLengthFt, "flowlenft")
    '        '                'UPGRADE_WARNING: Couldn't resolve default property of object pSlopeExp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        '                .BindRaster(pSlopeExp, "new_slpexp")
    '        '            End With
    '        '            'This non-dimensionalizes flowlenft and, after raising the result to the power in new_slpexp, we have the L factor.
    '        '            strExpression = "Pow(([flowlenft] / 72.6), [new_slpexp])"

    '        '            pRusleLFactor = pMapAlgebraOp.Execute(strExpression)
    '        '            'AddRasterLayer Application, pRusleLFactor, "RussleL"

    '        '            With pMapAlgebraOp
    '        '                .UnbindRaster("flowlenft")
    '        '                .UnbindRaster("new_slpexp")
    '        '            End With
    '        '        Else
    '        '            GoTo ProgCancel
    '        '        End If
    '        '        'END STEP 13: ------------------------------------------------------------------------

    '        '        'STEP 14: ----------------------------------------------------------------------------
    '        '        'Calculate the S-Factor
    '        '        modProgDialog.ProgDialog("Creating flow length GRID...", strProgTitle, 0, 15, 14, (m_App.hwnd))

    '        '        If modProgDialog.g_boolCancel Then

    '        '            'Here is the calculation for S, which is not, actually the slope,
    '        '            'but IS a function of the slope between a cell and its immediate downstream neighbor.
    '        '            'UPGRADE_WARNING: Couldn't resolve default property of object pDownAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        '            pMapAlgebraOp.BindRaster(pDownAngle, "dwnslangle")
    '        '            strExpression = "Con([dwnslangle] ge 5.1428, 16.8 * (sin(([dwnslangle] - 0.5) div deg))," & "10.8 * (sin(([dwnslangle] + 0.03) div deg)))"

    '        '            pRusleSFactor = pMapAlgebraOp.Execute(strExpression)

    '        '            pMapAlgebraOp.UnbindRaster("dwnslangle")
    '        '        Else
    '        '            GoTo ProgCancel
    '        '        End If
    '        '        'END STEP 14: -------------------------------------------------------------------------

    '        '        'STEP 15: ----------------------------------------------------------------------------
    '        '        'Calculate the LS Factor
    '        '        modProgDialog.ProgDialog("Calculating the LS Factor...", strProgTitle, 0, 15, 15, (m_App.hwnd))

    '        '        If modProgDialog.g_boolCancel Then

    '        '            'quick math to clip this bugger
    '        '            With pMapAlgebraOp
    '        '                'UPGRADE_WARNING: Couldn't resolve default property of object pRusleSFactor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        '                .BindRaster(pRusleSFactor, "Sfactor")
    '        '                'UPGRADE_WARNING: Couldn't resolve default property of object pRusleLFactor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        '                .BindRaster(pRusleLFactor, "Lfactor")
    '        '            End With

    '        '            strExpression = "[Sfactor] * [Lfactor]"

    '        '            pLSFactor = pMapAlgebraOp.Execute(strExpression)

    '        '            With pMapAlgebraOp
    '        '                .UnbindRaster("Sfactor")
    '        '                .UnbindRaster("Lfactor")
    '        '            End With
    '        '            'END STEP 15: -------------------------------------------------------------------------

    '        '            'STEP 15a: ----------------------------------------------------------------------------
    '        '            With pMapAlgebraOp
    '        '                'UPGRADE_WARNING: Couldn't resolve default property of object pLSFactor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        '                .BindRaster(pLSFactor, "LSFactor")
    '        '                'UPGRADE_WARNING: Couldn't resolve default property of object pMask. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        '                .BindRaster(pMask, "Mask")
    '        '            End With

    '        '            strExpression = "[LSFactor] * [Mask]"

    '        '            pFinalLS = pMapAlgebraOp.Execute(strExpression)

    '        '            With pMapAlgebraOp
    '        '                .UnbindRaster("LSFactor")
    '        '                .UnbindRaster("Mask")
    '        '            End With
    '        '        Else
    '        '            GoTo ProgCancel
    '        '        End If

    '        '        CalcLengthSlope = modUtil.MakePerminentGrid(pLSFactor, (pWS.PathName), "LSGrid")


    '    Catch ex As Exception
    '        '        MsgBox(Err.Number & ": " & Err.Description)
    '    Finally
    '        modProgDialog.KillDialog()
    '    End Try
    'End Function
#End Region

#Region "Raster Calc"


    Private Function maskCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        Try
            If Input1 > 0 Then
                Return 1
            Else
                Return OutNull
            End If
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Function


#End Region

End Class