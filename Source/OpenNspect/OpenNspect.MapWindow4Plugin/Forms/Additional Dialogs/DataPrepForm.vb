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
            If (aoiFName = "" Or demFName = "" Or lcFName = "" Or precipFName = "" _
                Or txtFinalCell.Text = "") Then
                If aoiFName = "" Then
                    txtAOI.Focus()
                ElseIf demFName = "" Then
                    txtDEMName.Focus()
                ElseIf lcFName = "" Then
                    txtLCName.Focus()
                ElseIf precipFName = "" Then
                    txtPrecipName.Focus()
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
        dirDPRoot = fileInfoDPRoot.DirectoryName.ToString
        dirDataPrep = dirDPRoot & "\ON_DataPrep"
        dirTarProj = dirDataPrep & "\TargetProj\"
        dirOtherProj = dirDataPrep & "\OtherProj\"
        If (Directory.Exists(dirDataPrep)) Then
            Dim response As MsgBoxResult = MsgBox("Root data directory exists.  Do you want to proceed?", _
                                                  MsgBoxStyle.AbortRetryIgnore)
            If response = MsgBoxResult.Abort Then
                Return
            ElseIf response = MsgBoxResult.Retry Then
                Return
            End If
        Else
            Directory.CreateDirectory(dirDataPrep)
            Directory.CreateDirectory(dirTarProj)
            Directory.CreateDirectory(dirOtherProj)
        End If

        ' Buffer AOI Shapefile by 10 cells -> AOIB10
        aoiBuffFName = dirTarProj & Path.GetFileNameWithoutExtension(aoiFName) & "B10.shp"
        MsgBox("aoiBuffName is " & aoiBuffFName)
        Dim aoi As New Shapefile
        Dim aoiB10 As New Shapefile
        Dim b10Dist As New Double
        b10Dist = 10.8 * finalCellSize
        aoi.Open(aoiFName)
        aoiB10.Open(aoiBuffFName)
        aoiB10 = aoi.BufferByDistance(b10Dist, 1, False, True)
        aoiB10.SaveAs(aoiBuffFName)


        ' Clip raw data to AOIB10, reprojecting AOI as needed
        'Define target Projection
        Dim tarGeoProj As New MapWinGIS.GeoProjection
        tarGeoProj = aoiB10.GeoProjection
        Dim tarProj4 As String
        tarProj4 = tarGeoProj.ExportToProj4

        ' Begin with DEM
        Dim demRaw As New Grid
        Dim rawProj4 As String
        Dim aoiDEMB10 As New Shapefile
        Dim demRawB10 As New Grid
        Dim demB10Fname As String
        Dim clipPoly As New MapWinGIS.Shape

        demRaw.Open(demFName)
        rawProj4 = demRaw.Header.GeoProjection.ExportToProj4

        ' Check projections
        If (Not demRaw.Header.GeoProjection.IsSame(tarGeoProj)) Then
            '            If Not String.Equals(rawProj4, tarProj4) Then
            ' Reprojection is needed
            MsgBox("Reprojection Needed" & Environment.NewLine _
                   & "Target Proj4 = " & tarProj4 & Environment.NewLine _
                   & "Raw Proj4    = " & rawProj4)
            ' Reproject Buffered AOI shapefile to  DEM Projection
            aoiDEMB10 = aoiB10.Reproject(demRaw.Header.GeoProjection, 1)
            'XXXXXXXXXXXXXXXXXXX This didn't work  DLE 9/27/2013
        Else
            ' Reprojection not needed so make clipping Aoi same as B10
            aoiDEMB10 = aoiB10
        End If
        ' Now clip by aoiDEMB10 polygon with fast, rectangular clipping
        clipPoly = aoiB10.Shape(0)
        demB10Fname = dirOtherProj & Path.GetFileNameWithoutExtension(demFName) & "B10.tif"
        If Not SpatialOperations.ClipGridWithPolygon(demFName, clipPoly, demB10Fname) Then
            MsgBox("CClip failed")
        Else
            MsgBox("ClipBySelectedPoly worked!")
        End If

        'If (Not demRaw.Header.GeoProjection.IsSame(tarGeoProj)) Then
        '    MsgBox("Reprojection Needed" & Environment.NewLine _
        '            & "TarGeoProj Name  = " & tarGeoProj.Name & Environment.NewLine _
        '            & "RawGeoProj Name    = " & demRaw.Header.GeoProjection.Name)
        'End If

 
    End Sub

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