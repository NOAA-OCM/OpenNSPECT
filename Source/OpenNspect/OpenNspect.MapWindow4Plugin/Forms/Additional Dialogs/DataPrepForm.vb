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

    Private Sub btnRun_Click(sender As Object, e As EventArgs) Handles btnRun.Click
        ' Check for valid data and set up main variables
        If (aoiFName = "" Or demFName = "" Or lcFName = "" Or precipFName = "") Then
            MsgBox("Please specify all input items")
            '            MsgBox("Please specify all input items", MsgBoxStyle.RetryCancel)
            Return
        End If

        ' Create intermediate subdirectories
        ' Buffer AOI Shapefile by 10 cells -> AOIB10
        ' Clip raw data to AOIB10, reprojecting AOI as needed
        ' Reproject to Target Projection

    End Sub
End Class