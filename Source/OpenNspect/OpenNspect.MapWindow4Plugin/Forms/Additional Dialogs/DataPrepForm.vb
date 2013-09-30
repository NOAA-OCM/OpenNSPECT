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


Public Class DataPrepForm
    Public Property Filter As String

    Public aoiFName As String
    Public demFName As String
    Public lcFName As String
    Public precipFName As String

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


        ' Buffer AOI Shapefile by 10 cells -> AOIB10
        aoiBuffFName = dirTarProj & Path.GetFileNameWithoutExtension(aoiFName) & "B10.shp"
        MsgBox("aoiBuffName is " & aoiBuffFName)
        If (File.Exists(aoiBuffFName) And removeOld) Then
            ' DelShapefile(aoiBuffFName)
            MapWinGeoProc.DataManagement.DeleteShapefile(aoiBuffFName)
        End If
        Dim aoi As New Shapefile
        Dim aoiB10 As New Shapefile
        Dim b10Dist As New Double
        b10Dist = 10.8 * finalCellSize
        aoi.Open(aoiFName)
        aoiB10.Open(aoiBuffFName)
        aoiB10 = aoi.BufferByDistance(b10Dist, 1, False, True)
        aoiB10.SaveAs(aoiBuffFName)
        'TODO: DLE 9/30/2013:  All these clip teh same raster!
        ' Begin with DEM
        Dim demFinal As New Grid
        demFinal = PrepOneRaster(demFName, aoiBuffFName, dirDataPrep, "DEM")

        ' Now land Cover
        Dim lcFinal As New Grid
        lcFinal = PrepOneRaster(lcFName, aoiBuffFName, dirDataPrep, "LULC")

        ' And Precip
        Dim precipFinal As New Grid
        precipFinal = PrepOneRaster(precipFName, aoiBuffFName, dirDataPrep, "Precip")

    End Sub
    'PrepOneRaster takes a given raster and clips it to the extent of a target shapefile, reprojecting 
    '   the shapefile if needed.  This makes the reprojection faster.
    '  Steps are: 1) reproject AOI shapefile, if needed, 2) Clip Raw Raster, 3) Reproject clipped Raster 
    '    back to original AOI projection
    Private Function PrepOneRaster(ByVal rawGridName As String, ByVal aoiSFName As String, _
                                   ByVal dpRoot As String, ByVal rasterSfx As String) As Grid
        Dim rawGrid As New Grid
        Dim rawProj4 As String
        Dim aoiRasterB10 As New Shapefile
        Dim rawGridB10 As New Grid
        Dim rawGridB10Fname As String
        Dim clipPoly As New MapWinGIS.Shape
        Dim aoiB10 As New Shapefile
        Dim tarGeoProj As New MapWinGIS.GeoProjection
        Dim tarProj4 As String

        rawGrid.Open(demFName)
        rawProj4 = rawGrid.Header.GeoProjection.ExportToProj4
        aoiB10.Open(aoiSFName)
        tarGeoProj = aoiB10.GeoProjection
        tarProj4 = tarGeoProj.ExportToProj4

        ' Check projections
        If (Not rawGrid.Header.GeoProjection.IsSame(tarGeoProj)) Then
            ' Reproject Buffered AOI shapefile to  DEM Projection
            aoiRasterB10 = aoiB10.Reproject(rawGrid.Header.GeoProjection, 1)
            Dim tmpName As String = dirOtherProj & "aoiRasterB10" & rasterSfx & ".shp"

            MsgBox("tmpName = " & tmpName.ToString)
            ' Now clip by aoiRasterB10 polygon with fast, rectangular clipping
            rawGridB10Fname = dirOtherProj & Path.GetFileNameWithoutExtension(demFName) & "B10" & rasterSfx & ".tif"
            MsgBox("Clipped raw file = " & rawGridB10Fname)
            '        rawGridB10.Open(rawGridB10Fname)
            rawGridB10 = ClipBySelectedPolyExtents(rawGrid, aoiRasterB10.Shape(0), rawGridB10Fname)

            'raw data are now clipped, so reproject this smaller raster back into the AOI Projection


            If (cbKeep.Checked) Then
                aoiRasterB10.SaveAs(tmpName)
            End If
        Else
            ' Reprojection not needed so make clipping Aoi same as B10
            aoiRasterB10 = aoiB10
            ' Now clip by aoiRasterB10 polygon with fast, rectangular clipping
            rawGridB10Fname = dirOtherProj & Path.GetFileNameWithoutExtension(demFName) & "B10" & rasterSfx & ".tif"
            MsgBox("Clipped raw file = " & rawGridB10Fname)
            '        rawGridB10.Open(rawGridB10Fname)
            rawGridB10 = ClipBySelectedPolyExtents(rawGrid, aoiRasterB10.Shape(0), rawGridB10Fname)
        End If
        rawGrid.Close()
        Return rawGridB10
  
    End Function

    'Private Sub txtFinalCell_TextChanged(sender As Object, e As EventArgs) Handles txtFinalCell.TextChanged
    '    Dim ch(20) As Char
    '    Dim len As Integer = txtFinalCell.Text.Length
    '    ch = txtFinalCell.Text.ToCharArray
    '    For i As Integer = 0 To len - 1
    '        If (Not IsNumeric(ch(i)) Or ch(i) = "." Or ch(i) = "," Or ch(i) = "-" Or ch(i) = "+") Then
    '            MsgBox("Final Cell Size can only contain numbers.  Please correct.")
    '        End If
    '    Next

    'End Sub
End Class