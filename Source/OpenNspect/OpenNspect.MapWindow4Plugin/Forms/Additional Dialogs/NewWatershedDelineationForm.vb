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
Imports System.Windows.Forms
Imports System.IO
Imports MapWindow.Interfaces
Imports MapWinGeoProc
Imports MapWinGIS

Friend Class NewWatershedDelineationForm
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
    Private _strNibbleName As String
    Private _strDEM2BName As String

    Private Const _dblSmall As Double = 0.03
    '0.001    '
    Private Const _dblMedium As Double = 0.06
    '0.01    '-Subwatershed sizes (3%, 6%, 10%)
    Private Const _dblLarge As Double = 0.1
    '

#Region "Events"

    Private Sub frmNewWSDelin_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
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
            Select Case chkHydroCorr.CheckState
                Case CheckState.Checked
                    chkStreamAgree.Enabled = True
                    cboStreamLayer.Enabled = True
                    cmdOptions.Enabled = True
                    g_boolHydCorr = True

                Case CheckState.Unchecked

                    chkStreamAgree.Enabled = False
                    cboStreamLayer.Enabled = False
                    cmdOptions.Enabled = False
                    g_boolHydCorr = False

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

    Protected Overrides Sub OK_Button_Click(sender As Object, e As EventArgs)
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
                'SQL Insert
                'DataBase Update
                'Compose the INSERT statement.
                Dim strCmdInsert As String = "INSERT INTO WSDelineation " & "(Name, DEMFileName, DEMGridUnits, FlowDirFileName, FlowAccumFileName," & "FilledDEMFileName, HydroCorrected, StreamFileName, SubWSSize, WSFileName, LSFileName, NibbleFileName, DEM2bFileName) " & " VALUES (" & "'" & CStr(txtWSDelinName.Text) & "', " & "'" & CStr(_InputDEMPath) & "', " & "'" & cboDEMUnits.SelectedIndex & "', " & "'" & _strDirFileName & "', " & "'" & _strAccumFileName & "', " & "'" & _strFilledDEMFileName & "', " & "'" & chkHydroCorr.CheckState & "', " & "'" & _strStreamLayer & "', " & "'" & cboSubWSSize.SelectedIndex & "', " & "'" & _strWShedFileName & "', " & "'" & _strLSFileName & "', " & "'" & _strNibbleName & "', " & "'" & _strDEM2BName & "')"

                'Execute the statement.
                Using insCmd As New DataHelper(strCmdInsert)
                    insCmd.ExecuteNonQuery()
                End Using

                'Confirm
                MsgBox(txtWSDelinName.Text & " successfully added.", MsgBoxStyle.OkOnly, "Record Added")

                If g_boolNewWShed Then
                    'frmPrj.Show
                    _frmPrj.Visible = True
                    _frmPrj.cboWaterShedDelineations.Items.Clear()
                    InitComboBox((_frmPrj.cboWaterShedDelineations), "WSDelineation")
                    _frmPrj.cboWaterShedDelineations.SelectedIndex = GetIndexOfEntry((txtWSDelinName.Text), (_frmPrj.cboWaterShedDelineations))
                    MyBase.OK_Button_Click(sender, e)
                Else
                    MyBase.OK_Button_Click(sender, e)
                    _frmWS.Close()
                End If
            Else
                Return
            End If

        Catch ex As Exception
            HandleError(ex)
        End Try

    End Sub

    Protected Overrides Sub Cancel_Button_Click(sender As Object, e As EventArgs)
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
                If boolChange(0) And boolChange(1) And boolChange(2) And boolChange(3) And g_boolParams Then
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

                If SynchronousProgressDialog.KeepRunning Then
                    _strFilledDEMFileName = OutPath + "demfill" + OutputGridExt
                    Hydrology.Fill(pSurfaceDatasetIn.Filename, _strFilledDEMFileName, False)
                    pFillRaster.Open(_strFilledDEMFileName)
                Else
                    Return False
                End If
            End If
            'End if Fill

            'STEP 2: Flow Direction
            progress.Increment("Computing Flow Direction...")

            Dim mwDirFileName As String = OutPath + "mwflowdir" + OutputGridExt
            _strDirFileName = OutPath + "flowdir" + OutputGridExt
            Dim strSlpFileName As String = OutPath + "slope" + OutputGridExt
            If SynchronousProgressDialog.KeepRunning Then
                ret = Hydrology.D8(pFillRaster.Filename, mwDirFileName, strSlpFileName, Environment.ProcessorCount, Nothing)
                If ret <> 0 Then Return False
                pFlowDirRaster.Open(mwDirFileName)

                Dim pESRID8Flow As New Grid
                Dim tmphead As New GridHeader
                tmphead.CopyFrom(pFlowDirRaster.Header)
                pESRID8Flow.CreateNew(_strDirFileName, tmphead, GridDataType.FloatDataType, -1)
                RasterMath(pFlowDirRaster, Nothing, Nothing, Nothing, Nothing, pESRID8Flow, Nothing, False, GetConverterToEsriFromTauDem())
                pESRID8Flow.Header.NodataValue = -1
                pESRID8Flow.Save(_strDirFileName)
                pESRID8Flow.Close()
            Else
                Return False
            End If

            'STEP 3: Flow Accumulation
            progress.Increment("Computing Flow Accumulation...")
            _strAccumFileName = OutPath + "flowacc" + OutputGridExt
            If SynchronousProgressDialog.KeepRunning Then
                ret = Hydrology.AreaD8(pFlowDirRaster.Filename, "", _strAccumFileName, False, False, Environment.ProcessorCount, Nothing)
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
                ret = Hydrology.DelinStreamGrids(pSurfaceDatasetIn.Filename, pFillRaster.Filename, pFlowDirRaster.Filename, strSlpFileName, pAccumRaster.Filename, "", strahlordout, longestupslopeout, totalupslopeout, streamgridout, streamordout, treedatout, coorddatout, _strStreamLayer, strWSGridOut, dblSubShedSize, False, False, 2, Nothing)
                If ret <> 0 Then Return False
            Else
                Return False
            End If

            'Step 6: Do WaterShed Op got moved into above step.
            _strWShedFileName = OutPath + "basinpoly.shp"

            progress.Increment("Creating Watershed Shape...")
            If SynchronousProgressDialog.KeepRunning Then
                Dim file = pFlowDirRaster.Filename
                pFlowDirRaster.Close()
                ret = Hydrology.SubbasinsToShape(file, strWSGridOut, strWSSFOut, Nothing)
                If ret <> 0 Then Return False
            Else
                Return False
            End If

            progress.Increment("Removing Small Polygons...")
            If SynchronousProgressDialog.KeepRunning Then
                pBasinFeatClass.Open(strWSSFOut)

                pOutputFeatClass = RemoveSmallPolys(pBasinFeatClass, pFillRaster)
            Else
                Return False
            End If

            'Save final output tif versions (since Taudem needed bgds before this)
            Dim proj As String = pOutputFeatClass.Projection
            Dim tmpfile As String
            tmpfile = _strFilledDEMFileName
            _strFilledDEMFileName = OutPath + "demfill" + FinalOutputGridExt
            pFillRaster.Close()
            pFillRaster = New Grid
            pFillRaster.Open(tmpfile)
            pFillRaster.Header.Projection = proj
            pFillRaster.Save(_strFilledDEMFileName)
            pFillRaster.Close()
            File.Delete(tmpfile)

            tmpfile = strSlpFileName
            strSlpFileName = OutPath + "slope" + FinalOutputGridExt
            Dim pslope As New Grid
            pslope.Open(tmpfile)
            pslope.Header.Projection = proj
            pslope.Save(strSlpFileName)
            pslope.Close()
            File.Delete(tmpfile)

            tmpfile = _strDirFileName
            _strDirFileName = OutPath + "flowdir" + FinalOutputGridExt

            pFlowDirRaster = New Grid
            pFlowDirRaster.Open(tmpfile)
            pFlowDirRaster.Header.Projection = proj
            pFlowDirRaster.Save(_strDirFileName)
            pFlowDirRaster.Close()
            File.Delete(tmpfile)

            tmpfile = _strAccumFileName
            _strAccumFileName = OutPath + "flowacc" + FinalOutputGridExt
            pAccumRaster.Close()
            pAccumRaster = New Grid
            pAccumRaster.Open(tmpfile)
            pAccumRaster.Header.Projection = proj
            pAccumRaster.Save(_strAccumFileName)
            pAccumRaster.Close()
            File.Delete(tmpfile)

            'With all of that done, now go get the name of the LS Grid while actually computing said LS Grid
            '_strLSFileName = CalcLengthSlope(pFillRaster, pFlowDirRaster, pAccumRaster, pEnv, "0", pWorkspace)
            _strLSFileName = OutPath + "lsgrid" + FinalOutputGridExt
            Dim g As New Grid
            g.Open(longestupslopeout)
            g.Save(_strLSFileName)
            g.Close()
            _strNibbleName = _strDirFileName
            _strDEM2BName = pSurfaceDatasetIn.Filename
            'TODO: create these if really needed

            DelineateWatershed = True
        Catch ex As Exception
            HandleError(ex)
            DelineateWatershed = False
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
            File.Delete(treedatout) 'TODO: trying to delete this file will not work if the user cancels.
            File.Delete(coorddatout)
            DataManagement.DeleteGrid(strWSGridOut)
            DataManagement.DeleteShapefile(strWSSFOut)
        End Try
    End Function

    Private Function RemoveSmallPolys(ByRef pFeatureClass As Shapefile, ByRef pDEMRaster As Grid) As Shapefile
        Dim pMaskCalcRaster As New Grid
        Dim rastersf As New Shapefile
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
            Dim outputSf As New Shapefile
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
            HandleError(ex)
            Return Nothing
        Finally
            pMaskCalcRaster.Close()
            pFeatureClass.Close()
            rastersf.Close()
        End Try
    End Function
#End Region

End Class