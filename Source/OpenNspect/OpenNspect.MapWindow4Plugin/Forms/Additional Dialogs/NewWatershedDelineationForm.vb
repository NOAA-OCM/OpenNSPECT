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
Imports System.IO
Imports MapWindow.Interfaces
Imports MapWinGeoProc
Imports MapWinGIS
Imports System.Windows.Forms
Imports System.Runtime.InteropServices

Friend Class CreateNewWatershedDelineationForm
    Private boolChange(3) As Boolean
    'Array set to track changes in controls: On Change,OK_Button is enabled
    Private _frmWS As WatershedDelineationsForm
    Private _frmPrj As MainForm

    Private _intSize As Short
    'Index for Size Combo
    Private _intCellSize As Short
    'Cell Size of DEM Grid, used in Length Slope Calculation
    Private _InputDEMPath As String

    Private _strDirFileName As String
    Private _strAccumFileName As String
    Private _strFilledDEMFileName As String
    Private _strStreamLayer As String
    Private _strWShedFileName As String
    Private _strLSFileName As String

    Private Const _dblSmall As Double = 0.03
    '0.001    '
    Private Const _dblMedium As Double = 0.06
    '0.01    '-Subwatershed sizes (3%, 6%, 10%)
    Private Const _dblLarge As Double = 0.1
    '

    'Agree DEM Stuff
    Private AgreeParams As Boolean
    'Flag to indicate Agree params have been entered

#Region "Events"

    Private Sub frmNewWSDelin_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        Try
            'Init bool variables
            AgreeParams = False

            Dim i As Short

            For i = 0 To UBound(boolChange)
                boolChange(i) = False
            Next i

            cboStreamLayer.Items.Clear()
            For i = 0 To MapWindowPlugin.MapWindowInstance.Layers.NumLayers - 1
                If MapWindowPlugin.MapWindowInstance.Layers(MapWindowPlugin.MapWindowInstance.Layers.GetHandle(i)).LayerType = eLayerType.LineShapefile Then
                    cboStreamLayer.Items.Add(MapWindowPlugin.MapWindowInstance.Layers(MapWindowPlugin.MapWindowInstance.Layers.GetHandle(i)).Name)
                End If
            Next
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub txtWSDelinName_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles txtWSDelinName.TextChanged
        Try
            boolChange(0) = True
            CheckEnabled()
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub txtDEMFile_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles txtDEMFile.TextChanged
        Try
            boolChange(1) = True
            CheckEnabled()

        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub cmdBrowseDEMFile_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdBrowseDEMFile.Click
        Try
            Dim pDEMRasterDataset As Grid

            pDEMRasterDataset = AddInputFromGxBrowserText(txtDEMFile, "Choose DEM GRID")
            If Not pDEMRasterDataset Is Nothing Then
                'Get the spatial reference
                Dim strProj As String = CheckSpatialReference(pDEMRasterDataset)
                If strProj = "" Then
                    MsgBox("The GRID you have choosen has no spatial reference information.  Please define a projection before continuing.", MsgBoxStyle.Exclamation, "No Project Information Detected")
                    Return
                Else
                    If strProj.ToLower.Contains("units=m") Then
                        cboDEMUnits.SelectedIndex = 0
                    Else
                        cboDEMUnits.SelectedIndex = 1
                    End If
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
            HandleError(ex)
        End Try
    End Sub

    Private Sub chkHydroCorr_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles chkHydroCorr.CheckedChanged
        Try
            Select Case chkHydroCorr.Checked

                Case True
                    chkStreamAgree.Enabled = True
                    cboStreamLayer.Enabled = True
                    cmdOptions.Enabled = True
                Case False

                    chkStreamAgree.Enabled = False
                    cboStreamLayer.Enabled = False
                    cmdOptions.Enabled = False
            End Select

            CheckEnabled()

        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub cboDEMUnits_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles cboDEMUnits.SelectedIndexChanged
        Try
            boolChange(2) = True
            CheckEnabled()
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub cboSubWSSize_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles cboSubWSSize.SelectedIndexChanged
        Try
            boolChange(3) = True
            CheckEnabled()
            _intSize = cboSubWSSize.SelectedIndex
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub


    Protected Overrides Sub OK_Button_Click(ByVal sender As Object, ByVal e As EventArgs)
        Try
            If _InputDEMPath = "" Then
                If txtDEMFile.Text <> "" Then
                    Dim path = GetRasterFullPath(txtDEMFile.Text)
                    If File.Exists(path) Then
                        _InputDEMPath = path
                    End If
                End If

                MsgBox("The File you have choosen does not exist.", MsgBoxStyle.Critical, "File Not Found")
                txtDEMFile.Focus()
                Return
            End If

            Dim pRasterDataset As New Grid
            If Not pRasterDataset.Open(_InputDEMPath) Then
                MsgBox("The File you have choosen is not a raster.", MsgBoxStyle.Critical, "File Not Raster")
                txtDEMFile.Focus()
                Return
            End If

            Dim outpath As String
            If Not Directory.Exists(String.Format("{0}\wsdelin\{1}", g_nspectPath, txtWSDelinName.Text)) Then
                outpath = String.Format("{0}\wsdelin\{1}\", g_nspectPath, txtWSDelinName.Text)
                Directory.CreateDirectory(outpath)
            Else
                MsgBox("Name in use.  Please select another.", MsgBoxStyle.Critical, "Choose New Name")
                txtWSDelinName.Focus()
                Return
            End If

            'Give the call; if successful insert new record
            If DelineateWatershed(pRasterDataset, outpath) Then

                InsertWaterShedDelineation()

                MyBase.OK_Button_Click(sender, e)
            Else
                Return
            End If

        Catch ex As Exception
            HandleError(ex)
        End Try

    End Sub
    Private Sub InsertWaterShedDelineation()
        'TODO: refactor this duplicate method.
        Dim strCmdInsert As String = String.Format("INSERT INTO WSDelineation (Name, DEMFileName, DEMGridUnits, FlowDirFileName, FlowAccumFileName,FilledDEMFileName, HydroCorrected, StreamFileName, SubWSSize, WSFileName, LSFileName)  VALUES ('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}', '{9}', '{10}')", CStr(txtWSDelinName.Text), CStr(_InputDEMPath), cboDEMUnits.SelectedIndex, _strDirFileName, _strAccumFileName, _strFilledDEMFileName, IIf(chkHydroCorr.Checked, 1, 0), _strStreamLayer, cboSubWSSize.SelectedIndex, _strWShedFileName, _strLSFileName)

        'Execute the statement.
        Using insCmd As New DataHelper(strCmdInsert)
            insCmd.ExecuteNonQuery()
        End Using

        'Confirm
        MsgBox(txtWSDelinName.Text & " successfully added.", MsgBoxStyle.OkOnly, "Record Added")

        If g_boolNewWShed Then
            _frmPrj.Visible = True
            _frmPrj.cboWaterShedDelineations.Items.Clear()
            InitComboBox((_frmPrj.cboWaterShedDelineations), "WSDelineation")
            _frmPrj.cboWaterShedDelineations.SelectedIndex = GetIndexOfEntry((txtWSDelinName.Text), (_frmPrj.cboWaterShedDelineations))
        Else
            _frmWS.Close()
        End If
    End Sub

    Protected Overrides Sub Cancel_Button_Click(ByVal sender As Object, ByVal e As EventArgs)
        If Not _frmPrj Is Nothing Then
            _frmPrj.cboPrecipitationScenarios.SelectedIndex = 0
        End If
        MyBase.Cancel_Button_Click(sender, e)
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
                If boolChange(0) And boolChange(1) And boolChange(2) And boolChange(3) And AgreeParams Then
                    OK_Button.Enabled = True
                Else
                    OK_Button.Enabled = False
                End If
            ElseIf chkHydroCorr.CheckState And Not chkStreamAgree.CheckState Then
                If boolChange(0) And boolChange(1) And boolChange(2) And boolChange(3) Then
                    OK_Button.Enabled = True
                Else
                    OK_Button.Enabled = False
                End If
            ElseIf Not chkHydroCorr.CheckState Then
                If boolChange(0) And boolChange(1) And boolChange(2) And boolChange(3) Then
                    OK_Button.Enabled = True
                Else
                    OK_Button.Enabled = False
                End If
            End If

        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub FatalError(ByVal ex As SEHException)
        Trace.TraceError(ex.Message)
        Dim ret = MessageBox.Show("An error occured in the Hydrology class. This leaves everything in an unstable state, so MapWindow needs to be shut down. Close MapWindow now?", "Close MapWindow?", MessageBoxButtons.YesNo)
        If ret = Windows.Forms.DialogResult.Yes Then
            Application.ExitThread()
        Else
            Close()
        End If

    End Sub
    Private Function DelineateWatershed(ByRef pSurfaceDatasetIn As Grid, ByVal OutPath As String) As Boolean
        'Declare the raster objects
        Dim pFlowDirRaster As New Grid
        'Flow Direction
        Dim pAccumRaster As New Grid
        'Flow Accumulation
        Dim pFillRaster As New Grid
        'Fill GDS
        Dim pWShedRaster As New Grid
        'WaterShed GDS
        Dim pBasinRaster As New Grid
        'Basin GDS
        Dim pOutputFeatClass As New Shapefile
        'output basins
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
        Dim pBasinFeatClass As New Shapefile
        'Basin Featureclass

        Dim progress = New SynchronousProgressDialog("Filling DEM...", "Watershed Delineation Processing...", 10, Me)
        Try
            'Get cell size from DEM; needed later in the Length Slope Calculation
            _intCellSize = pSurfaceDatasetIn.Header.dX

            Dim ret As Integer

            'STEP 1:  Fill the Surface
            'if hydrocorrect, then skip the Fill, just use the incoming DEM

            If chkHydroCorr.CheckState = 1 Then
                pFillRaster = pSurfaceDatasetIn
                _strFilledDEMFileName = pSurfaceDatasetIn.Filename
            Else
                'Call to ProgDialog to use throughout process: keep user informed.

                If Not SynchronousProgressDialog.KeepRunning Then
                    Return False
                End If

                _strFilledDEMFileName = OutPath + "demfill" + OutputGridExt
                Hydrology.Fill(pSurfaceDatasetIn.Filename, _strFilledDEMFileName, False)

                pFillRaster.Open(_strFilledDEMFileName)

            End If
            'End if Fill

            'STEP 2: Flow Direction
            progress.Increment("Computing Flow Direction...")

            Dim mwDirFileName As String = OutPath + "mwflowdir" + OutputGridExt
            _strDirFileName = OutPath + "flowdir" + OutputGridExt
            Dim strSlpFileName As String = OutPath + "slope" + OutputGridExt
            If SynchronousProgressDialog.KeepRunning Then
                Try
                    ret = Hydrology.D8(pFillRaster.Filename, mwDirFileName, strSlpFileName, Environment.ProcessorCount, False, Nothing)
                Catch ex As SEHException
                    FatalError(ex)
                    Return False
                End Try

                If ret <> 0 Then Return False

                pFlowDirRaster.Open(mwDirFileName)

                Dim pESRID8Flow As New Grid
                Dim tmphead As New GridHeader
                tmphead.CopyFrom(pFlowDirRaster.Header)
                pESRID8Flow.CreateNew(_strDirFileName, tmphead, GridDataType.FloatDataType, -1)
                RasterMath(pFlowDirRaster, Nothing, Nothing, Nothing, Nothing, pESRID8Flow, Nothing, False, GetConverterToEsriFromTauDem())
                pESRID8Flow.Header.NodataValue = -1
                pESRID8Flow.Save()  'Saving because it seemed necessary.
                pESRID8Flow.Save(_strDirFileName)
                pESRID8Flow.Close()
            Else
                Return False
            End If

            'STEP 3: Flow Accumulation
            progress.Increment("Computing Flow Accumulation...")
            _strAccumFileName = OutPath + "flowacc" + OutputGridExt
            If SynchronousProgressDialog.KeepRunning Then
                Try
                    ret = Hydrology.AreaD8(pFlowDirRaster.Filename, "", _strAccumFileName, False, False, Environment.ProcessorCount, False, Nothing)
                Catch ex As SEHException
                    FatalError(ex)
                    Return False
                End Try
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

            strahlordout = OutPath + "strahlord" + OutputGridExt
            longestupslopeout = OutPath + "longestupslope" + OutputGridExt
            totalupslopeout = OutPath + "totalupslope" + OutputGridExt
            streamgridout = OutPath + "streamgrid" + OutputGridExt
            streamordout = OutPath + "streamord" + OutputGridExt
            treedatout = OutPath + "tree.dat"
            coorddatout = OutPath + "coord.dat"
            strWSGridOut = OutPath + "wsgrid" + OutputGridExt
            strWSSFOut = OutPath + "ws.shp"

            '        'Step 5: Using Hydrology Op to create stream network
            _strStreamLayer = OutPath + "stream.shp"
            progress.Increment("Creating Stream Network...")
            If SynchronousProgressDialog.KeepRunning Then
                Try
                    ret = Hydrology.DelinStreamGrids(pSurfaceDatasetIn.Filename, pFillRaster.Filename, pFlowDirRaster.Filename, strSlpFileName, pAccumRaster.Filename, "", strahlordout, longestupslopeout, totalupslopeout, streamgridout, streamordout, treedatout, coorddatout, _strStreamLayer, strWSGridOut, dblSubShedSize, False, False, 2, False, Nothing)
                Catch ex As SEHException
                    FatalError(ex)
                    Return False
                End Try
                If ret <> 0 Then Return False
            Else
                Return False
            End If

            ' Copy prj file for "basinpoly.shp" http://nspect.codeplex.com/workitem/20849
            File.Copy(OutPath + "stream.prj", OutPath + "basinpoly.prj")

            FillInMissingSpaces(strWSGridOut, pFlowDirRaster)

            'Step 6: Do WaterShed Op got moved into above step.
            _strWShedFileName = OutPath + "basinpoly.shp"

            progress.Increment("Creating Watershed Shape...")
            If Not SynchronousProgressDialog.KeepRunning Then
                Return False
            Else
                pFlowDirRaster.Save()
                Dim file = pFlowDirRaster.Filename
                pFlowDirRaster.Close()
                Try
                    ret = Hydrology.SubbasinsToShape(file, strWSGridOut, strWSSFOut, Nothing)
                Catch ex As SEHException
                    FatalError(ex)
                    Return False
                End Try
                If ret <> 0 Then Return False
            End If

            progress.Increment("Removing Small Polygons...")
            If Not SynchronousProgressDialog.KeepRunning Then
                Return False
            Else
                pBasinFeatClass.Open(strWSSFOut)

                pOutputFeatClass = RemoveSmallPolys(pBasinFeatClass, pFillRaster)
            End If

            _strLSFileName = OutPath + "lsgrid" + OutputGridExt
            Dim g As New Grid
            g.Open(longestupslopeout)
            g.Save(_strLSFileName)
            g.Close()

            Return True
        Catch ex As Exception
            HandleError(ex)
            Return False
        Finally
            progress.Dispose()
            pFlowDirRaster.Close()
            pAccumRaster.Close()
            pFillRaster.Close()
            pWShedRaster.Close()
            pBasinRaster.Close()
            pBasinFeatClass.Close()
            pOutputFeatClass.Close()
            DataManagement.DeleteGrid(strahlordout)
            DataManagement.DeleteGrid(longestupslopeout)
            DataManagement.DeleteGrid(totalupslopeout)
            DataManagement.DeleteGrid(streamgridout)
            DataManagement.DeleteGrid(streamordout)
            DataManagement.TryDelete(treedatout)
            DataManagement.TryDelete(coorddatout)
            DataManagement.DeleteGrid(strWSGridOut)
            DataManagement.DeleteShapefile(strWSSFOut)
        End Try
    End Function

    Private Function RemoveSmallPolys(ByRef pFeatureClass As Shapefile, ByRef pDEMRaster As Grid) As Shapefile
        Try
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

            'Once we have an outline, union it, but for now just output the base with smalls removed

            '#5  Now, time to union the outline of the of the DEM with the newly paired down basin poly
            'Dim outputSf As MapWinGIS.Shapefile = rastersf.GetIntersection(False, pFeatureClass, False, pFeatureClass.ShapefileType, Nothing)
            'outputSf.SaveAs(_strWShedFileName)
            pFeatureClass.SaveAs(_strWShedFileName)
            pFeatureClass.Close()
            Dim outputSf As New Shapefile
            outputSf.Open(_strWShedFileName)


            'Dim currshape As MapWinGIS.Shape = rastersf.Shape(0)
            'For i As Integer = 0 To pFeatureClass.NumShapes() - 1
            '    currshape = MapWinGeoProc.SpatialOperations.Union(currshape, pFeatureClass.Shape(0))
            'Next
            'outputSf.CreateNew(_strWShedFileName, MapWinGIS.ShpfileType.SHP_POLYGON)
            'outputSf.StartEditingShapes()
            'outputSf.EditInsertShape(currshape, 0)
            'outputSf.StopEditingTable()
            Return outputSf
        Catch ex As Exception
            HandleError(ex)
            Return Nothing
        Finally
            pFeatureClass.Close()

        End Try
    End Function
#End Region

    Private Shared Sub FillInMissingSpaces(ByVal gridPath As String, ByVal dem As Grid)

        Dim g As New Grid
        g.Open(gridPath)

        Dim header = g.Header
        Dim nr As Integer = header.NumberRows - 1
        Dim nc As Integer = header.NumberCols - 1
        Dim gridNoData = header.NodataValue
        Dim demNoData = dem.Header.NodataValue
        Dim newValue = g.Maximum + 1

        For row As Integer = 0 To nr
            For col = 0 To nc
                Dim demValue = dem.Value(col, row)
                If demValue <> demNoData Then
                    If g.Value(col, row) = gridNoData Then
                        g.Value(col, row) = newValue
                    End If
                End If
            Next
        Next
        g.Save()

    End Sub
    ''' <summary>
    ''' Flattens a grid to simple (nodata or 1) values
    ''' </summary>
    ''' <param name="grid"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function Flatten(ByVal grid As Grid) As Grid
        Dim tempPath = GetTempFileNameOutputGridExt()
        grid.Save()
        grid.Save(tempPath)

        Dim g As New Grid
        g.Open(tempPath)
        Dim header = g.Header
        Dim nr As Integer = header.NumberRows - 1
        Dim nc As Integer = header.NumberCols - 1
        Dim gridNoData As Double = header.NodataValue

        ' Convert grid to a simple 0, 1 grid
        For row As Integer = 0 To nr
            For col = 0 To nc
                Dim value As Double = grid.Value(col, row)
                If value = gridNoData Then
                    Debug.Assert(g.Value(col, row) = gridNoData)
                Else
                    g.Value(col, row) = 1
                End If
            Next
        Next
        g.Save()

        Return g

    End Function

    ''' <summary>
    ''' Removes the first and last pixel on each row in the grid. Assumes no holes.
    ''' </summary>
    ''' <param name="grid"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function RightLeftTrim(ByVal grid As Grid) As Grid
        Dim tempPath = GetTempFileNameOutputGridExt()
        grid.Save()
        grid.Save(tempPath)

        Dim g As New Grid
        g.Open(tempPath)
        Dim header = g.Header
        Dim nr As Integer = header.NumberRows - 1
        Dim nc As Integer = header.NumberCols - 1
        Dim gridNoData As Double = header.NodataValue

        For row As Integer = 0 To nr
            Dim overwroteStart = False
            Dim overwroteFinish = False

            For col = 0 To nc
                Dim value As Double = grid.Value(col, row)
                If value = 1 Then
                    If overwroteStart Then
                        ' look for the finish
                        Dim nextCol = Math.Min(nc, col + 1)
                        If grid.Value(nextCol, row) = gridNoData Then
                            'we're in the last cell with data
                            g.Value(col, row) = gridNoData
                            overwroteFinish = True
                        End If
                    Else
                        g.Value(col, row) = gridNoData
                        overwroteStart = True
                    End If
                End If

                If overwroteFinish Then
                    ' skip to the next row.
                    Exit For
                End If
            Next
        Next
        g.Save()

        Return g

    End Function

    ''' <summary>
    ''' Trims the specified grid.
    ''' </summary>
    ''' <param name="gridPath">The grid.</param>
    ''' <param name="clippingArea">The clipping area.</param><returns></returns>
    Private Function Trim(ByVal gridPath As String, ByVal clippingArea As Grid) As Grid

        Dim g As New Grid
        g.Open(gridPath)
        Dim header = g.Header
        Dim nr As Integer = header.NumberRows - 1
        Dim nc As Integer = header.NumberCols - 1
        Dim gridNoData As Single = header.NodataValue
        Dim clippingNoData As Single = clippingArea.Header.NodataValue

        For row As Integer = 0 To nr
            For col = 0 To nc
                ' if the grid has data
                If g.Value(col, row) <> gridNoData Then
                    ' and the clipping area does not
                    If clippingArea.Value(col, row) = clippingNoData Then
                        ' remove the data from the grid
                        g.Value(col, row) = gridNoData
                    End If
                End If
            Next
        Next
        g.Save()

        Return g

    End Function
End Class
