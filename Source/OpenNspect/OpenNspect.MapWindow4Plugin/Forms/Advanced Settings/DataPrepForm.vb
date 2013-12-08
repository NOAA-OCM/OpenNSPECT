'********************************************************************************************************
'File Name: DataPrepForm.vb
'Description: Form for Preprocessing data from various sources to a consistent projection 
'             and clipping to a target Area of Interest
'********************************************************************************************************
'The contents of this file are subject to the Mozilla Public License Version 1.1 (the "License"); 
'you may not use this file except in compliance with the License. You may obtain a copy of the License at 
'http://www.mozilla.org/MPL/ 
'Software distributed under the License is distributed on an "AS IS" basis, WITHOUT WARRANTY OF 
'ANY KIND, either express or implied. See the License for the specificlanguage governing rights and 
'limitations under the License. 
'
'Contributor(s): (Open source contributors should list themselves and their modifications here). 
'Sept. 26, 2013:  Dave Eslinger dave.eslinger@noaa.gov 
'      Initial commit to repository, not actaully working.
'Oct. 31, 2013: Working well, still some cleanup and usability tweaks needed.
'Nov. 11, 2013: Working, errors caught (I think) and should be pretty solid.

Imports System.IO
Imports MapWindow.Interfaces
Imports MapWinGIS
Imports MapWinGeoProc
Imports MapWindow
Imports MapWinUtility
Imports MapWindow.Controls.GisToolbox
Imports MapWindow.Controls.Data

Public Class DataPrepForm
    Inherits BaseDialogForm
    Public Property Filter As String
    Public aoiFName As String
    Public demFName As String
    Public lcFName As String
    Public precipFName As String
    Public refXll As Double
    Public refYll As Double
    Public refdX As Double
    Public refdY As Double
    Public tarAOI As New Shapefile
    Private aoiBuff20FName As String
    Private demBuffFname As String
    Private lcBuffFname As String
    Private precipBuffFName As String
    Private dirDPRoot As String
    Private dirTarProj As String
    Private dirDataPrep As String
    Private dirOtherProj As String
    Private dirFinal As String
    Private finalCellSize As Double
    Private bufferSize As Double

    Private Sub btnAOI_Click(sender As Object, e As EventArgs) Handles btnAOI.Click
        Dim AOI As New MapWinGIS.Shapefile
        Dim tmpAOI As New MapWinGIS.Shapefile
        Dim tmpaoiFName As String
        diaOpenPrep.Reset()
        diaOpenPrep.Title = "Open AOI Shapefile"
        diaOpenPrep.Filter = "Shapefiles|*.shp"
        If diaOpenPrep.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txtAOI.Text = diaOpenPrep.FileName
            tmpaoiFName = txtAOI.Text
            'MsgBox("Shapefile name is " & aoiFName)
            AOI.Open(tmpaoiFName)
            If (AOI.NumShapes > 1) Then
                Dim msgResponse As New MsgBoxResult
                msgResponse = MsgBox("Your AOI shapefile contains mulitple polygons, ALL of which will be merged into one polygon for use and saved as a new shapefile.  " & _
                        "If this is not desired, select 'No' to choose another shapefile, or select 'Cancel' to close the entire Data Prep dialog and create a different AOI shapefile. " & vbCrLf & vbCrLf & _
                        "Note: If you want to use only some polygons from a shapefile, you must manually select those polygons and use the 'Export Selected' tool on the toolbar " & _
                        "to create a new shapefile of just the desired area.  Use 'Cancel' in this case, to close the Data Prep window.", MsgBoxStyle.YesNoCancel, _
                        "Merge all polygons?")
                If (msgResponse = MsgBoxResult.Cancel) Then
                    Close()
                ElseIf (msgResponse = MsgBoxResult.Yes) Then
                    AOI.SelectAll()
                    tmpAOI = AOI.AggregateShapes(True)
                    AOI.Save()
                    AOI.Close()
                    AOI = tmpAOI.Dissolve(0, False)
                    aoiFName = Path.GetDirectoryName(tmpaoiFName) & "\" _
                    & Path.GetFileNameWithoutExtension(tmpaoiFName) _
                    & "_DP.shp"
                    AOI.SaveAs(aoiFName)
                    AOI.Close()
                    AOI.Open(aoiFName)
                    txtAOI.Text = aoiFName
                    MsgBox("New Merged AOIFilename is " & aoiFName)
                    If (Not AddFeatureLayerToMapFromFileName(aoiFName, "Merged AOI")) Then
                        MsgBox("ERROR: Merged AOI Shapfile not found: " & vbLf & txtAOI.Text.ToString)
                    End If
                Else
                    Exit Sub
                End If
            Else
                aoiFName = tmpaoiFName
            End If
            txtProjParams.Text = AOI.Projection.ToString
            txtProjName.Text = AOI.GeoProjection.ProjectionName
            txtFinalCellUnits.Text = ProjectionUnits(txtProjParams.Text, AOI.GeoProjection)
            If (AOI.GeoProjection.ProjectionName = "") Then
                If (AOI.Projection.Substring(6, 7) = "longlat") Then
                    txtProjName.Text = "Geographic (unprojected), see parameters below:"
                    MsgBox("Hmmm... your AOI shapefile is in geographic coordinates.  " & _
                           "That is almost always a bad idea for your target shapefile.  " & _
                           "You should probably check the shapefile name and try again.")
                End If
            End If

        End If
        diaOpenPrep.Filter = ""
    End Sub

    Private Sub btnOpenDEM_Click(sender As Object, e As EventArgs) Handles btnOpenDEM.Click
        Dim g As New MapWinGIS.Grid
        diaOpenPrep.Reset()
        diaOpenPrep.Title = "Open Elevation Raster"
        diaOpenPrep.Filter = g.CdlgFilter
        If diaOpenPrep.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txtDEMName.Text = diaOpenPrep.FileName.ToString
            demFName = txtDEMName.Text
            'MsgBox("DEM file is " & txtDEMName.Text)
            g.Open(demFName)
            txtCellSizeDEM.Text = g.Header.dX.ToString
            txtSizeUnitsDEM.Text = ProjectionUnits(g.Header.Projection.ToString, g.Header.GeoProjection)
            txtDEMProj.Text = g.Header.GeoProjection.ProjectionName
            txtDEMParams.Text = g.Header.Projection.ToString
            If (g.Header.GeoProjection.ProjectionName = "") Then
                If (g.Header.Projection.Substring(6, 7) = "longlat") Then
                    txtDEMProj.Text = "Geographic (unprojected), see parameters below:"
                End If
            End If
            g.Close()
        End If
    End Sub

    Private Sub btnOpenLC_Click(sender As Object, e As EventArgs) Handles btnOpenLC.Click
        Dim g As New MapWinGIS.Grid
        diaOpenPrep.Reset()
        diaOpenPrep.Title = "Open Land Cover/Land Use Raster"
        diaOpenPrep.Filter = g.CdlgFilter
        If diaOpenPrep.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txtLCName.Text = diaOpenPrep.FileName.ToString
            lcFName = txtLCName.Text
            'MsgBox("LC file is " & txtLCName.Text)
            g.Open(lcFName)
            txtCellSizeLC.Text = g.Header.dX.ToString
            txtSizeUnitsLC.Text = ProjectionUnits(g.Header.Projection.ToString, g.Header.GeoProjection)
            txtLCProj.Text = g.Header.GeoProjection.ProjectionName
            txtLCParams.Text = g.Header.Projection.ToString
            If (g.Header.GeoProjection.ProjectionName = "") Then
                If (g.Header.Projection.Substring(6, 7) = "longlat") Then
                    txtLCProj.Text = "Geographic (unprojected), see parameters below:"
                End If
            End If
            g.Close()
        End If

    End Sub

    Private Sub btnOpenPrecip_Click(sender As Object, e As EventArgs) Handles btnOpenPrecip.Click
        Dim g As New MapWinGIS.Grid
        diaOpenPrep.Reset()
        diaOpenPrep.Title = "Open Precipitation Raster"
        diaOpenPrep.Filter = g.CdlgFilter
        If diaOpenPrep.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txtPrecipName.Text = diaOpenPrep.FileName.ToString
            precipFName = txtPrecipName.Text
            'MsgBox("Precip file is " & txtPrecipName.Text)
            g.Open(precipFName)
            txtCellSizePrecip.Text = g.Header.dX.ToString
            txtSizeUnitsPrecip.Text = ProjectionUnits(g.Header.Projection.ToString, g.Header.GeoProjection)
            txtPrecipProj.Text = g.Header.GeoProjection.ProjectionName
            txtPrecipParams.Text = g.Header.Projection.ToString
            If (g.Header.GeoProjection.ProjectionName = "") Then
                If (g.Header.Projection.Substring(6, 7) = "longlat") Then
                    txtPrecipProj.Text = "Geographic (unprojected), see parameters below:"
                End If
            End If
            g.Close()
        End If
    End Sub

    'Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
    Protected Overrides Sub OK_Button_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim boolCell As Boolean = False
        Try
            'Do Until boolCell
            ' Check for valid data and set up main variables
            Try
                finalCellSize = Convert.ToDouble(txtFinalCell.Text)
                bufferSize = Convert.ToDouble(txtUserBuffer.Text)
                boolCell = True
                'MsgBox("Cell size is " & finalCellSize)
                ' If it gets to here, the cell size is good, so check the rest of the input and run the 
                ' data preparation functions if it all checks out. If it isn't working, prompt user to correct and retry.
                If (aoiFName = "" Or demFName = "" Or lcFName = "" Or precipFName = "") Then
                    If aoiFName = "" Then
                        txtAOI.Focus()
                    ElseIf demFName = "" Then
                        txtDEMName.Focus()
                    ElseIf lcFName = "" Then
                        txtLCName.Focus()
                    ElseIf precipFName = "" Then
                        txtPrecipName.Focus()
                    End If

                    Dim msgResult As MsgBoxResult = MsgBox("Please specify all input items", MsgBoxStyle.RetryCancel)
                    If (msgResult = MsgBoxResult.Cancel) Then
                        Close()
                    Else
                        boolCell = False
                    End If
                Else
                    '  MsgBox("txtFinalCell.Text = '" & txtFinalCell.Text & "'")
                    Try
                        If (PrepRawData()) Then
                            boolCell = True
                        Else
                            boolCell = False
                        End If
                    Catch ex As Exception
                        boolCell = False
                        MsgBox("Error in PrepRawData:" & ex.ToString, MsgBoxStyle.Critical)
                    End Try
                    If (cbLoadFinal.Checked) Then
                        ' Load AOI into MapWindow
                        If (Not AddFeatureLayerToMapFromFileName(aoiFName, "Area of Interest")) Then
                            MsgBox("ERROR: AOI Shapfile not found: " & vbLf & txtAOI.Text.ToString)
                        End If
                    End If
                End If
            Catch ex As Exception
                boolCell = False
                MsgBox("Final cell or buffer size is invalid.  Please correct." & vbLf & ex.ToString, MsgBoxStyle.Critical)
            End Try

        Catch ex As Exception
            MsgBox("Error: " & ex.ToString, MsgBoxStyle.Critical)
        Finally
            If boolCell Then
                Close()
            End If
        End Try
    End Sub

    Private Function PrepRawData() As Boolean
        '  MsgBox("More processing")
        ' Create intermediate subdirectories
        Dim fileInfoDPRoot As New FileInfo(aoiFName)
        Dim refGridName As String
        Dim removeOld As New Boolean
        removeOld = False

        dirDPRoot = fileInfoDPRoot.DirectoryName.ToString
        dirDataPrep = dirDPRoot & "\ON_DataPrep"
        dirTarProj = dirDataPrep & "\TargetProj\"
        dirOtherProj = dirDataPrep & "\OtherProj\"
        If cbKeep.Checked = True Then
            If (Directory.Exists(dirDataPrep)) Then
                Dim response As MsgBoxResult = MsgBox("Root data directory exists.  Do you want to proceed, deleting any duplicate files?", _
                                                      MsgBoxStyle.OkCancel)
                If response = MsgBoxResult.Cancel Then
                    MsgBox("Canceling.  To rerun, put AOI Shapefile in a new directory.")
                    Return False
                ElseIf response = MsgBoxResult.Ok Then
                    removeOld = True
                End If
            End If
        End If

        Directory.CreateDirectory(dirDataPrep)
        Directory.CreateDirectory(dirTarProj)
        Directory.CreateDirectory(dirOtherProj)

        ' Buffer AOI Shapefile by 20cells -> AOIB20
        aoiBuff20FName = dirTarProj & Path.GetFileNameWithoutExtension(aoiFName) & "B20.shp"
        'MsgBox("aoiBuffName is " & aoiBuffFName)
        If (File.Exists(aoiBuff20FName) And removeOld) Then
            ' DelShapefile(aoiBuffFName)
            MapWinGeoProc.DataManagement.DeleteShapefile(aoiBuff20FName)
        End If
        'Dim aoi As New Shapefile
        Dim aoiB20 As New Shapefile
        Dim B20Dist As New Double
        B20Dist = bufferSize * finalCellSize
        tarAOI.Open(aoiFName)
        aoiB20.Open(aoiBuff20FName)
        aoiB20 = tarAOI.BufferByDistance(B20Dist, 1, False, True)
        aoiB20.SaveAs(aoiBuff20FName)

        ' Begin with DEM, this will be the reference file for nudging
        Dim demFinal As New Grid
        Dim demFinalFName As String
        refGridName = ""
        demFinalFName = PrepOneRaster(demFName, aoiBuff20FName, refGridName, dirDataPrep, "DEM")
        Try
            File.Copy(Path.ChangeExtension(demFName, "mwleg"), Path.ChangeExtension(demFinalFName, "mwleg"))
        Catch ex As Exception
            MsgBox("It looks like there is no MapWindow color file for your original DEM file.  You will get a default set of colors.")
        End Try
        If (cbLoadFinal.Checked) Then
            ' Load DEM into MapWindow
            If (Not AddFeatureLayerToMapFromFileName(demFinalFName)) Then
                MsgBox("ERROR: Final DEM raster not found: " & vbLf & txtAOI.Text.ToString)
            End If
        End If

        ' Now land Cover
        Dim lcFinal As New Grid
        Dim lcFinalFName As String
        lcFinalFName = PrepOneRaster(lcFName, aoiBuff20FName, refGridName, dirDataPrep, "LULC")
        Try
            File.Copy(Path.ChangeExtension(lcFName, "mwleg"), Path.ChangeExtension(lcFinalFName, "mwleg"))
        Catch ex As Exception
            MsgBox("It looks like there is no MapWindow color file for your original Land Cover file.  You just get a default set of colors.")
        End Try
        If (cbLoadFinal.Checked) Then
            ' Load LULC into MapWindow
            If (Not AddFeatureLayerToMapFromFileName(lcFinalFName)) Then
                MsgBox("ERROR: Final LULC raster not found: " & vbLf & txtAOI.Text.ToString)
            End If
        End If

        ' And Precip
        Dim precipFinal As New Grid
        Dim precipFinalFName As String
        precipFinalFName = PrepOneRaster(precipFName, aoiBuff20FName, refGridName, dirDataPrep, "Precip")
        Try
            File.Copy(Path.ChangeExtension(precipFName, "mwleg"), Path.ChangeExtension(precipFinalFName, "mwleg"))
        Catch ex As Exception
            MsgBox("It looks like there is no MapWindow color file for your original Precip file.  You will get a default set of colors.")
        End Try
        If (cbLoadFinal.Checked) Then
            ' Load DEM into MapWindow
            If (Not AddFeatureLayerToMapFromFileName(precipFinalFName)) Then
                MsgBox("ERROR: Final Precip raster not found: " & vbLf & txtAOI.Text.ToString)
            End If
        End If

        'If everything ran well, clean up, delete temporary files as needed, and close elegantly:

        If (Not cbKeep.Checked) Then
            'Close all the temporary files:
            aoiB20.Close()
            ' And the files to be kept:
            demFinal.Close()
            lcFinal.Close()
            precipFinal.Close()
            ' Then recursively delete the temprorary directories and the files they contain:
            Try
                Directory.Delete(dirOtherProj, True)
            Catch ex As Exception
                MsgBox("Error deleting " & dirOtherProj & ": " & vbCrLf & ex.ToString, MsgBoxStyle.Critical)
            End Try
            Try
                Directory.Delete(dirTarProj, True)
            Catch ex As Exception
                MsgBox("Error deleting " & dirTarProj & ": " & vbCrLf & ex.ToString, MsgBoxStyle.Critical)
            End Try
        End If
        Return True
    End Function
    ''' <summary>
    ''' Takes a given raster and clips it to the extent of a target shapefile, reprojecting the shapefile if needed.
    ''' </summary>
    ''' <param name="rawGridName"></param>
    ''' <param name="aoi20SFName"></param>
    ''' <param name="refGridFName"></param>
    ''' <param name="dpRoot"></param>
    ''' <param name="rasterSfx"></param>
    ''' <remarks>
    '''  PrepOneRaster takes a given raster and clips it to the extent of a target shapefile, reprojecting 
    '''   the shapefile if needed.  This makes the reprojection faster.
    '''  Steps are: 1) reproject AOI shapefile, if needed, 2) Clip Raw Raster, 3) Reproject clipped Raster 
    '''    back to original AOI projection, 4) Bin reprojected raster to desired cell size, 5) align 
    '''</remarks>
    Private Function PrepOneRaster(ByVal rawGridName As String, ByVal aoi20SFName As String, _
                                   ByRef refGridFName As String, _
                                   ByVal dpRoot As String, ByVal rasterSfx As String) As String
        Dim rawGrid As New Grid
        Dim tarAOI As New Shapefile
        Dim aoiRasterB20 As New Shapefile
        Dim rawGridB20 As New Grid
        Dim rawGridB20Fname As String
        Dim aoiB20 As New Shapefile
        Dim tarGeoProj As New MapWinGIS.GeoProjection
        Dim tarProj4 As String
        Dim tarGridB20 As New Grid  ' A buffered grid in final projection, ready for binning nudging and final clipping
        Dim tarBinned As New Grid   ' Projected raster that has been rebinned (if needed) to target bin size
        Dim tarNudged As New Grid   ' Nudged grid, ready for final clipping
        Dim tarFinal As New Grid    ' Final projected, nudged and clipped grid
        Dim tarFinalFName As String    ' Final projected, nudged and clipped grid Name
        Dim clippedFName As String    ' Clipped grid Name, may be in correct projection or may not be
        Dim tarBinFName As String
        Dim tarNudgedFName As String
        Dim tmpFName As String    ' Just a temporary file name variable
        Dim onlyGDAL As Boolean = False  'Only use GDAL Warp routine for reprojection and rebinning
        Dim statusReproject As Boolean = False

        rawGrid.Open(rawGridName)
        aoiB20.Open(aoi20SFName)
        tarGeoProj = aoiB20.GeoProjection
        tarProj4 = tarGeoProj.ExportToProj4

        Dim tmpMessage As String
        tmpMessage = "Data Prep on " & Path.GetFileNameWithoutExtension(rawGridName)

        Dim progress As New SynchronousProgressDialog("Begin processing...", tmpMessage, 7, Me)

        ' Check projections and cell sizes and change as needed:

        If (Not rawGrid.Header.GeoProjection.IsSame(tarGeoProj)) Then
            ' Mismatch between projections.  
            ' Reproject Buffered Target AOI shapefile to  Raw Raster Projection and CLIP the RAW Raster
            progress.Increment("Reprojecting buffered AOI shapefile...")
            aoiRasterB20 = aoiB20.Reproject(rawGrid.Header.GeoProjection, 1)
            tmpFName = dirOtherProj & "aoiRasterB20" & rasterSfx & ".shp"
            aoiRasterB20.SaveAs(tmpFName)
            'MsgBox("New Buffered AOI is  = " & tmpFName.ToString)

            ' Now clip by aoiRasterB20 polygon with fast, rectangular clipping
            rawGridB20Fname = dirOtherProj & Path.GetFileNameWithoutExtension(rawGridName) & "B20" & rasterSfx & ".tif"
            progress.Increment("Clipping to reprojected, buffered AOI")
            Try
                rawGridB20 = ClipBySelectedPolyExtents(rawGrid, aoiRasterB20.Shape(0), rawGridB20Fname)
                'MsgBox("Clipped raw file = " & rawGridB20Fname)
            Catch ex As Exception
                MsgBox("Error on clipping to buffered AOI.  Check that AOI and file overlap:" & rawGridName)
            End Try

            'raw data are now clipped, so reproject this smaller raster back into the AOI Projection
            If (onlyGDAL) Then
                clippedFName = dirTarProj & Path.GetFileNameWithoutExtension(rawGridB20Fname) & ".tif"
            Else
                'Test if reproject worked and save, if so.  If not, error message
                clippedFName = dirTarProj & Path.GetFileNameWithoutExtension(rawGridB20Fname) & ".tif"
                progress.Increment("Reprojecting the clipped, buffered grid to target projection")
                statusReproject = SpatialReference.ProjectGrid(rawGridB20.Header.GeoProjection.ExportToProj4, _
                                                               tarProj4, rawGridB20Fname, clippedFName, False)
                If (statusReproject) Then
                    ' MsgBox("Reprojected buffered file is " & clippedFName)
                    tarGridB20.Open(clippedFName)
                Else
                    MsgBox("Error in reprojecting " & rawGridB20Fname.ToString, MsgBoxStyle.OkCancel)
                End If
            End If
        Else
            ' Reprojection not needed so make clipping Aoi same as B20
            aoiRasterB20 = aoiB20  ' Clip by aoiRasterB20 polygon with fast, rectangular clipping
            rawGridB20Fname = dirTarProj & Path.GetFileNameWithoutExtension(rawGridName) & "B20" & rasterSfx & ".tif"
            ' MsgBox("Clipped raw file = " & rawGridB20Fname)
            progress.Increment("Reprojection not needed.  Clipping to buffered AOI...")
            clippedFName = dirTarProj & Path.GetFileNameWithoutExtension(rawGridB20Fname) & ".tif"
            ' tarGridB20.Open(clippedFName)
            tarGridB20 = ClipBySelectedPolyExtents(rawGrid, aoiRasterB20.Shape(0), rawGridB20Fname)
            If (Not tarGridB20.Save(clippedFName)) Then
                MsgBox("Error trying to save " & clippedFName)
            End If
        End If

        ' Raster is now clipped to buffered size.  Ready for
        '   1) Reprojecting and Binning to specified cell size
        '   2) Nudging to align grid with DEM and
        '   3) Clipping to final AOI
        '
        '  Internal Grid names are:
        '    tarGridB20   ' The buffered grid in final projection, ready for binning, nudging and final clipping
        '    tarBinned    ' Projected raster that has been rebinned (if needed) to target bin size
        '    tarNudged    ' Nudged grid, ready for final clipping
        '    tarFinal     ' Final projected, nudged and clipped grid

        tarBinFName = dirTarProj & Path.GetFileNameWithoutExtension(clippedFName) & "_B.tif"
        tarNudgedFName = dirTarProj & Path.GetFileNameWithoutExtension(tarBinFName) & "_N.tif"
        progress.Increment("Binning to target cell size...")

        statusReproject = DoResample_DLE(tarGridB20, tarBinFName, finalCellSize)
        If (statusReproject) Then
            ' MsgBox("DoResample_DLE reproject worked on " & tarBinFName)
            tarBinned.Open(tarBinFName)
        Else
            MsgBox("Binning failed on " & tarBinFName)
        End If

        ' Raster is now Binned.  If it is NOT the DEM/Reference files, nudge it to align properly with that file.
        ' NOTE BENE: The DEM/Reference Raster must be processed first!

        If (rasterSfx = "DEM") Then ' Set the reference coordinates
            progress.Increment("Reference grid, no nudging needed, copying original grid...")
            refGridFName = tarBinFName
            tarNudged = CopyRaster(tarBinned, tarNudgedFName)
            tarNudged.Save(tarNudgedFName)
        Else
            progress.Increment("Nudging to reference grid...")
            If (NudgeGrid(tarBinFName, refGridFName, tarNudgedFName)) Then
            Else
            End If

        End If

        ' The grid is now nudged, so clip to final shapefile and return it.
        tarFinalFName = dirDataPrep & "\" & Path.GetFileNameWithoutExtension(rawGridName) & "_DP.tif"
        tarFinal.Open(tarFinalFName)
        tarNudged.Open(tarNudgedFName)
        tarAOI.Open(aoiFName)
        progress.Increment("Clipping to final non-buffered, projected AOI!")
        'SpatialOperations.ClipGridWithPolygon(tarNudgedFName, tarAOI.Shape(0), tarFinalFName, False)
        tarFinal = ClipBySelectedPoly(tarNudged, tarAOI.Shape(0), tarFinalFName)
        tarFinal.Save()
        tarFinal.Save(tarFinalFName)

        ' Close all temporary files to keep everything clean and tidy!
        aoiRasterB20.Close()
        aoiB20.Close()
        rawGrid.Close()
        rawGridB20.Close()
        tarGridB20.Close()
        tarBinned.Close()
        tarNudged.Close()
        tarFinal.Close()
        Return tarFinalFName
    End Function

    ''' <summary>
    ''' Copies a target grid while shifting Lower left X and Y coordinates to make the new grid line up with a reference grid
    ''' </summary>
    ''' <param name="tarGridFName"></param>
    ''' <param name="refGridName"></param>
    ''' <param name="nudgedGridFName"></param>
    ''' <remarks></remarks>
    Private Function NudgeGrid(ByVal tarGridFName As String, ByVal refGridName As String, ByVal nudgedGridFName As String) As Boolean
        Dim tarGrid As New Grid
        Dim refGrid As New Grid
        Dim nudgedGrid As New Grid

        Dim shiftXll, shiftYll, delX, delY As New Double
        Dim tarXll, tarYll, tarDX, tarDY As New Double
        Dim refXll, refYll, refDX, refDY As New Double

        refGrid.Open(refGridName)
        refXll = refGrid.Header.XllCenter
        refYll = refGrid.Header.YllCenter
        refDX = refGrid.Header.dX
        refDY = refGrid.Header.dY

        tarGrid.Open(tarGridFName)
        tarXll = tarGrid.Header.XllCenter
        tarYll = tarGrid.Header.YllCenter
        tarDX = tarGrid.Header.dX
        tarDY = tarGrid.Header.dY

        If (refDX.Equals(tarDX) And refDY.Equals(tarDY)) Then

            delX = refXll - tarXll
            delY = refYll - tarYll
            shiftXll = (refXll - tarXll) Mod refDX
            shiftYll = (refYll - tarYll) Mod refDY

            Try
                'Check to see if any needed X shift is less than 1/2 the X cell size
                'and correct as needed.
                If (Math.Abs(shiftXll) > refDX / 2.0) Then
                    If (shiftXll < 0) Then
                        shiftXll = shiftXll + refDX
                    Else
                        shiftXll = shiftXll - refDX
                    End If
                End If

                'Check to see if any needed Y shift is less than 1/2 the Y cell size
                'and correct as needed.
                If (Math.Abs(shiftYll) > refDY / 2.0) Then
                    If (shiftYll < 0) Then
                        shiftYll = shiftYll + refDY
                    Else
                        shiftYll = shiftYll - refDY
                    End If
                End If

            Catch ex As Exception
                MsgBox("Something broke in calculating the X and Y shifts")
                Return False
            End Try

            '           If (Not shiftXll.Equals(0.0) Or Not shiftYll.Equals(0.0)) Then
            If (shiftXll <> 0.0 Or shiftYll <> 0.0) Then
                Try
                    'Dim tarNudged As New Grid
                    nudgedGrid = CopyRaster(tarGrid, nudgedGridFName)
                    nudgedGrid.Header.XllCenter = nudgedGrid.Header.XllCenter + shiftXll
                    nudgedGrid.Header.YllCenter = nudgedGrid.Header.YllCenter + shiftYll
                    nudgedGrid.Save(nudgedGridFName)
                    'nudgedGrid.Close()
                Catch ex As Exception
                    MsgBox("Something failed in saving nudged file " & nudgedGridFName)
                    Return False
                End Try
            Else
                ' Cells already line up, so just copy the grid
                nudgedGrid = CopyRaster(tarGrid, nudgedGridFName)
                nudgedGrid.Save(nudgedGridFName)
                MsgBox("On " & tarGrid.Filename & ", no shift required")
                Return True
            End If
        Else
            MsgBox("Cell sizes do not match.  Nudging not possible on grid" & _
                   ControlChars.CrLf & tarGrid.Filename, MsgBoxStyle.Exclamation)
            Return False
        End If
        refGrid.Close()
        tarGrid.Close()
        nudgedGrid.Close()
        Return True
    End Function

    ''' <summary>
    ''' Resamples a grid to create a new gird with teh specified cell size
    ''' </summary>
    ''' <param name="grd"></param>
    ''' <param name="newGridFName"></param>
    ''' <param name="CellSize"></param>
    ''' <remarks>This code is copied from the MapWindow DoResample function, whic I don't seem to ba able to access.</remarks>
    Public Function DoResample_DLE(ByRef grd As MapWinGIS.Grid, ByVal newGridFName As String, ByVal CellSize As Double) As Boolean
        Dim i, j As Integer
        Dim newGrid As New MapWinGIS.Grid
        Dim newHeader As New MapWinGIS.GridHeader
        Dim numCols, numRows As Integer
        Dim absLeft, absRight, absBottom, absTop As Double
        Dim halfDX, halfDY As Double
        Dim tX, tY, oldX, oldY, nDX, cDX As Double

        Try
            With newHeader
                numCols = Int((grd.Header.dX * grd.Header.NumberCols) / CellSize)
                numRows = Int((grd.Header.dY * grd.Header.NumberRows) / CellSize)

                absLeft = grd.Header.XllCenter - (grd.Header.dX / 2)
                absBottom = grd.Header.YllCenter - (grd.Header.dY / 2)
                absRight = absLeft + (grd.Header.dX * grd.Header.NumberCols)
                absTop = absBottom + (grd.Header.dY * grd.Header.NumberRows)

                newHeader.NumberCols = numCols
                newHeader.NumberRows = numRows
                newHeader.dX = CellSize
                newHeader.dY = CellSize
                newHeader.XllCenter = absLeft + (CellSize / 2)
                newHeader.YllCenter = absBottom + (CellSize / 2)
                newHeader.NodataValue = grd.Header.NodataValue
                newHeader.Notes = grd.Header.Notes
                newHeader.Key = grd.Header.Key
                newHeader.Projection = grd.Header.Projection

                If newGrid.CreateNew(newGridFName, newHeader, grd.DataType, grd.Header.NodataValue, True) = False Then
                    Return False
                End If

                halfDX = newHeader.dX * 0.5
                halfDY = newHeader.dY * 0.5

                For j = 0 To numRows - 1
                    tY = absTop - (j * newHeader.dY) - halfDY

                    nDX = newHeader.dX
                    cDX = grd.Header.dX

                    oldY = Int(grd.Header.NumberRows - ((tY - absBottom) / grd.Header.dY))

                    For i = 0 To numCols - 1
                        tX = absLeft + (i * nDX) + halfDX
                        oldX = Int((tX - absLeft) / cDX)
                        newGrid.Value(i, j) = grd.Value(oldX, oldY)
                    Next i
                Next j
            End With

            grd.Close()
            grd = newGrid
            grd.Save(newGridFName)
            Return True
        Catch ex As Exception
            MapWinUtility.Logger.Msg(ex.Message & vbCrLf & ex.StackTrace, MsgBoxStyle.Critical Or MsgBoxStyle.Information, "DoResample_DLE - Error")
            Return False
        End Try
    End Function

End Class