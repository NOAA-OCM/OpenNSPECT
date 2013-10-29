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
'      Initial commit to repository
Imports System.IO
Imports MapWindow.Interfaces
Imports MapWinGIS
Imports MapWinGeoProc
Imports MapWindow
Imports MapWinUtility
Imports MapWindow.Controls.GisToolbox
Imports MapWindow.Controls.Data



Public Class DataPrepForm
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


    Private aoiBuffFName As String
    Private demBuffFname As String
    Private lcBuffFname As String
    Private precipBuffFName As String
    Private dirDPRoot As String
    Private dirTarProj As String
    Private dirDataPrep As String
    Private dirOtherProj As String
    Private dirFinal As String

    Private finalCellSize As Double

    Private Sub btnAOI_Click(sender As Object, e As EventArgs) Handles btnAOI.Click
        Dim AOI As New MapWinGIS.Shapefile
        diaOpenPrep.Reset()
        diaOpenPrep.Title = "Open AOI Shapefile"
        diaOpenPrep.Filter = "Shapefiles|*.shp"
        If diaOpenPrep.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txtAOI.Text = diaOpenPrep.FileName
            aoiFName = txtAOI.Text
            'MsgBox("Shapefile name is " & aoiFName)
            AOI.Open(aoiFName)
            txtProjParams.Text = AOI.Projection.ToString
            txtProjName.Text = AOI.GeoProjection.ProjectionName
            txtFinalCellUnits.Text = ProjectionUnits(txtProjParams.Text, AOI.GeoProjection)
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
            g.Close()
        End If
    End Sub

    Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
        Try
            'If btnRun.DialogResult = Windows.Forms.DialogResult.OK Then

            'End If
            ' Check for valid data and set up main variables
            Try
                finalCellSize = Convert.ToDouble(txtFinalCell.Text)
                'MsgBox("Cell size is " & finalCellSize)
            Catch ex As Exception
                MsgBox("Final Cell size invalid.  Please correct.")
            End Try
            txtFinalCell.Focus()
            'MsgBox("txtFinalCell.Text = '" & txtFinalCell.Text & "'")
            'If (aoiFName = "" Or demFName = "" Or lcFName = "" Or precipFName = "" _
            '    Or txtFinalCell.Text = "") Then
            '    If aoiFName = "" Then
            '        txtAOI.Focus()
            '    ElseIf demFName = "" Then
            '        txtDEMName.Focus()
            '    ElseIf lcFName = "" Then
            '        txtLCName.Focus()
            '    ElseIf precipFName = "" Then
            '        txtPrecipName.Focus()
            If (aoiFName = "" _
                Or txtFinalCell.Text = "") Then
                If aoiFName = "" Then
                    txtAOI.Focus()
                Else
                    txtFinalCell.Focus()
                End If
                Dim msgResult As MsgBoxResult = MsgBox("Please specify all input items", MsgBoxStyle.RetryCancel)
            Else
                '  MsgBox("txtFinalCell.Text = '" & txtFinalCell.Text & "'")
                PrepRawData()
            End If

        Catch ex As Exception
            MsgBox("Error: " & ex.ToString, MsgBoxStyle.Critical)
        Finally
            'TODO DLE, 10/23/2013: Fix teh cose section so it only closes when there is not an error.
            Close()
        End Try
    End Sub

    Private Sub PrepRawData()
        '  MsgBox("More processing")
        ' Create intermediate subdirectories
        Dim fileInfoDPRoot As New FileInfo(aoiFName)
        Dim removeOld As New Boolean
        removeOld = False

        If cbKeep.Checked = True Then
            dirDPRoot = fileInfoDPRoot.DirectoryName.ToString
            dirDataPrep = dirDPRoot & "\ON_DataPrep"
            dirTarProj = dirDataPrep & "\TargetProj\"
            dirOtherProj = dirDataPrep & "\OtherProj\"
            If (Directory.Exists(dirDataPrep)) Then
                Dim response As MsgBoxResult = MsgBox("Root data directory exists.  Do you want to proceed, deleting any duplicate files?", _
                                                      MsgBoxStyle.OkCancel)
                If response = MsgBoxResult.Cancel Then
                    MsgBox("Canceling.  To rerun, put AOI Shapefile in a new directory.")
                    Return
                ElseIf response = MsgBoxResult.Ok Then
                    removeOld = True
                    ' Return
                End If
                'Else
                '    Directory.CreateDirectory(dirDataPrep)
                '    Directory.CreateDirectory(dirTarProj)
                '    Directory.CreateDirectory(dirOtherProj)
            End If
        Else
            'TODO: Add logic to use a tmp directory to be deleted later.
            ' Temporary directories here
            dirDPRoot = fileInfoDPRoot.DirectoryName.ToString
            dirDataPrep = dirDPRoot & "\ON_DataPrep"
            dirTarProj = dirDataPrep & "\TargetProj\"
            dirOtherProj = dirDataPrep & "\OtherProj\"
        End If

        Directory.CreateDirectory(dirDataPrep)
        Directory.CreateDirectory(dirTarProj)
        Directory.CreateDirectory(dirOtherProj)


        ' Buffer AOI Shapefile by 2cells -> AOIB20
        aoiBuffFName = dirTarProj & Path.GetFileNameWithoutExtension(aoiFName) & "B20.shp"
        'MsgBox("aoiBuffName is " & aoiBuffFName)
        If (File.Exists(aoiBuffFName) And removeOld) Then
            ' DelShapefile(aoiBuffFName)
            MapWinGeoProc.DataManagement.DeleteShapefile(aoiBuffFName)
        End If
        'Dim aoi As New Shapefile
        Dim aoiB20 As New Shapefile
        Dim B20Dist As New Double
        B20Dist = 20 * finalCellSize
        tarAOI.Open(aoiFName)
        aoiB20.Open(aoiBuffFName)
        aoiB20 = tarAOI.BufferByDistance(B20Dist, 1, False, True)
        aoiB20.SaveAs(aoiBuffFName)

        ' Begin with DEM, this will be the reference file for nudging
        Dim demFinal As New Grid
        Dim demFinalFName As String
        demFinalFName = PrepOneRaster(demFName, aoiBuffFName, dirDataPrep, "DEM")
        demFinal.Open(demFinalFName)
  
        ' Now land Cover
        Dim lcFinal As New Grid
        Dim lcFinalFName As String
        lcFinalFName = PrepOneRaster(lcFName, aoiBuffFName, dirDataPrep, "LULC")

        ' And Precip
        Dim precipFinal As New Grid
        Dim precipFinalFName As String
        precipFinalFName = PrepOneRaster(precipFName, aoiBuffFName, dirDataPrep, "Precip")
        'precipFinal.Save(precipFName)
        'precipFinal.Close()

        'If everything ran well, delete temporary files as needed:
        ' add stuff here...
        If (Not cbKeep.Checked) Then
            'Close all the temporary files:
            aoiB20.Close()
            ' And the files to be kept:
            demFinal.Close()
            lcFinal.Close()
            precipFinal.Close()

            ' Then recursively delte teh temprorary directories adn the files they contain:
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



    End Sub
    'PrepOneRaster takes a given raster and clips it to the extent of a target shapefile, reprojecting 
    '   the shapefile if needed.  This makes the reprojection faster.
    '  Steps are: 1) reproject AOI shapefile, if needed, 2) Clip Raw Raster, 3) Reproject clipped Raster 
    '    back to original AOI projection, 4) Bin reprojected raster to desired cell size, 5) align 
    Private Function PrepOneRaster(ByVal rawGridName As String, ByVal aoiSFName As String, _
                                   ByVal dpRoot As String, ByVal rasterSfx As String) As String
        Dim rawGrid As New Grid
        Dim rawProj4 As String
        Dim aoiRasterB20 As New Shapefile
        Dim rawGridB20 As New Grid
        Dim rawGridB20Fname As String
        Dim clipPoly As New MapWinGIS.Shape
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
        Dim warpNeeded As Boolean
        Dim statusReproject As Boolean = False
        Dim tarAOI As New Shapefile

        Dim tarXll As Double
        Dim tarYll As Double
        Dim tardX As Double
        Dim tardY As Double

        rawGrid.Open(rawGridName)
        rawProj4 = rawGrid.Header.GeoProjection.ExportToProj4
        aoiB20.Open(aoiSFName)
        tarGeoProj = aoiB20.GeoProjection
        tarProj4 = tarGeoProj.ExportToProj4

        Dim tmpMessage As String
        tmpMessage = "Begin processing on " & rawGridName

        Dim progress As New SynchronousProgressDialog(tmpMessage, "Data Preparation on " & rawGridName, 6, Me)


        ' Check projections and cell sizes and change as needed:
        '       If (Not rawGrid.Header.GeoProjection.IsSame(tarGeoProj) Or rawGrid.Header.dX <> finalCellSize) Then

        If (Not rawGrid.Header.GeoProjection.IsSame(tarGeoProj)) Then
            progress.Increment("Reprojecting buffered AOI shapefile...")
            ' Mismatch between projections.  
            warpNeeded = True
            ' Reproject Buffered Target AOI shapefile to  Raw Raster Projection and CLIP the RAW Raster
            aoiRasterB20 = aoiB20.Reproject(rawGrid.Header.GeoProjection, 1)
            tmpFName = dirOtherProj & "aoiRasterB20" & rasterSfx & ".shp"
            aoiRasterB20.SaveAs(tmpFName)
            'MsgBox("New Buffered AOI is  = " & tmpFName.ToString)
            ' Now clip by aoiRasterB20 polygon with fast, rectangular clipping
            rawGridB20Fname = dirOtherProj & Path.GetFileNameWithoutExtension(rawGridName) & "B20" & rasterSfx & ".tif"
            'MsgBox("Clipped raw file = " & rawGridB20Fname)
            rawGridB20 = ClipBySelectedPolyExtents(rawGrid, aoiRasterB20.Shape(0), rawGridB20Fname)

            'raw data are now clipped, so reproject this smaller raster back into the AOI Projection
            If (onlyGDAL) Then
                clippedFName = dirTarProj & Path.GetFileNameWithoutExtension(rawGridB20Fname) & ".tif"
            Else
                'Test if reproject worked and save, if so.  If not, error message
                progress.Increment("Clipping to reprojected, buffered AOI")
                clippedFName = dirTarProj & Path.GetFileNameWithoutExtension(rawGridB20Fname) & ".tif"
                statusReproject = SpatialReference.ProjectGrid(rawGridB20.Header.GeoProjection.ExportToProj4, _
                                                               tarProj4, rawGridB20Fname, clippedFName, False)
                If (statusReproject) Then
                    ' If (cbKeep.Checked) Then
                    ' MsgBox("Reprojected buffered file is " & clippedFName)
                    tarGridB20.Open(clippedFName)
                    'End If
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
            tarGridB20.Open(clippedFName)
            tarGridB20 = ClipBySelectedPolyExtents(rawGrid, aoiRasterB20.Shape(0), rawGridB20Fname)
            If (Not tarGridB20.Save(clippedFName)) Then
                MsgBox("Error trying to save " & clippedFName)
            End If
            If (rawGrid.Header.dX <> finalCellSize) Then
                warpNeeded = True
            Else
                warpNeeded = False
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
            MsgBox("DoResample_DLE reproject worked on " & tarBinFName)
            'tarBinned.Save(tarBinFName)
        Else
            MsgBox("Binning failed on " & tarBinFName)
        End If

        ' Raster is now Binned.  If it is NOT the DEM/Reference files, nudge it to align properly with that file.
        ' If it IS the DEM/Ref raster, set the reference coordinates  for subsiequent nudging.
        ' NOTE BENE: The DEM/Reference Raster must be processed first!

        Dim tarGrid As New Grid
        Dim shiftXll, shiftYll, delX, delY As Double

        tarGrid.Open(tarBinFName)
        tarXll = tarGrid.Header.XllCenter
        tarYll = tarGrid.Header.YllCenter
        tardX = tarGrid.Header.dX
        tardY = tarGrid.Header.dY
        If (rasterSfx = "DEM") Then ' Set the reference coordinates
            progress.Increment("Reference grid, no nudging needed...")
            tarNudged = CopyRaster(tarGrid, tarNudgedFName)
            refXll = tarXll
            refYll = tarYll
            refdX = tardX
            refdY = tarDy
        Else
            If (refdX.Equals(tardX) And refdY.Equals(tardY)) Then
                progress.Increment("Nudging to align with reference grid...")
                delX = refXll - tarXll
                delY = refYll - tarYll
                shiftXll = (refXll - tarXll) Mod refdX
                shiftYll = (refYll - tarYll) Mod refdY

                Try
                    'Check to see if any needed X shift is less than 1/2 the X cell size
                    'and correct as needed.
                    If (Math.Abs(shiftXll) > refdX / 2.0) Then
                        If (shiftXll < 0) Then
                            shiftXll = shiftXll + refdX
                        Else
                            shiftXll = shiftXll - refdX
                        End If
                    End If

                    'Check to see if any needed Y shift is less than 1/2 the Y cell size
                    'and correct as needed.
                    If (Math.Abs(shiftYll) > refdY / 2.0) Then
                        If (shiftYll < 0) Then
                            shiftYll = shiftYll + refdY
                        Else
                            shiftYll = shiftYll - refdY
                        End If
                    End If

                Catch ex As Exception
                    MsgBox("Something broke in calculating the X and Y shifts")
                End Try

                If (Not shiftXll.Equals(0.0) Or Not shiftYll.Equals(0.0)) Then
                    Try
                        'Dim tarNudged As New Grid
                        tarNudged = CopyRaster(tarGrid, tarNudgedFName)
                        tarNudged.Header.XllCenter = tarNudged.Header.XllCenter + shiftXll
                        tarNudged.Header.YllCenter = tarNudged.Header.YllCenter + shiftYll
                        tarNudged.Save()
                        tarNudged.Close()

                    Catch ex As Exception
                        MsgBox("Something in new file " & tarNudgedFName & " failed!")
                    End Try

                Else
                    MsgBox("On " & tarGrid.Filename & ", no shift required")

                End If

            Else
                MsgBox("Cell sizes do not match.  Nudging not possible on grid" & _
                       ControlChars.CrLf & tarGrid.Filename, MsgBoxStyle.Exclamation)
            End If

        End If
        ' The grid is now nudged, so clip to final shapefile and return it.
        'TODO Not working here.  CHeck shapefile...

        tarFinalFName = dirDataPrep & "\" & Path.GetFileNameWithoutExtension(rawGridName) & "_DP.tif"
        tarFinal.Open(tarFinalFName)
        tarNudged.Open(tarNudgedFName)
        tarAOI.Open(aoiFName)
        progress.Increment("Clipping to final non-buffered, projected AOI!")
        tarFinal = ClipBySelectedPoly(tarNudged, tarAOI.Shape(0), tarFinalFName)
        tarFinal.Save()

        ' Close all temporary files to keep everything clean and tidy!
        aoiRasterB20.Close()
        aoiB20.Close()
        rawGrid.Close()
        rawGridB20.Close()
        tarGridB20.Close()
        tarBinned.Close()
        tarNudged.Close()
        tarFinal.Close()
        tarGrid.Close()
        tarAOI.Close()


        Return tarFinalFName

    End Function

    Private Function BinProjRaster(ByVal unbinRasFName As String, ByVal BinRasFName As String, _
                                   ByVal srcProj4 As String, ByVal tarProj4 As String, _
                                   ByVal cellSize As Double) As Boolean
        Dim binnedRas As New Grid
        Dim unbinRas As New Grid
        unbinRas.Open(unbinRasFName)
        Try
            Dim strTransform As String
            strTransform = "-overwrite -of GTiff -r near" _
                & " -s_srs '" & srcProj4 & "' -t_srs '" & tarProj4 & "' " _
                & " -tr " & Convert.ToString(cellSize) & " " & Convert.ToString(cellSize)
            MsgBox("Transform string is " & strTransform)
            Dim u = New MapWinGIS.UtilsClass
 
            If (Not u.GDALWarp(unbinRasFName, BinRasFName, strTransform)) Then
                MsgBox("Error, transform did not work", MsgBoxStyle.Critical)
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.ToString, MsgBoxStyle.Critical)
            Return False
        End Try

    End Function
    Public Function DoResample_DLE(ByRef grd As MapWinGIS.Grid, ByVal newGridFName As String, ByVal CellSize As Double) As Boolean
        Dim i, j As Integer
        Dim newGrid As New MapWinGIS.Grid
        Dim newHeader As New MapWinGIS.GridHeader
        Dim numCols, numRows As Integer
        Dim absLeft, absRight, absBottom, absTop As Double
        Dim halfDX, halfDY As Double
        Dim tX, tY, oldX, oldY, nDX, cDX As Double

        ' Dim newFilen As String = System.IO.Path.GetFileName(grd.Filename)

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
                    'Progress.Progress(grd.Filename, j / numRows * 100, "Resampling " & grd.Filename & " row " & j)

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
            MapWinUtility.Logger.Msg(ex.Message & vbCrLf & ex.StackTrace, MsgBoxStyle.Critical Or MsgBoxStyle.Information, "Grid Wizard 2.0 - Error")
            Return False
        End Try
    End Function

   
End Class