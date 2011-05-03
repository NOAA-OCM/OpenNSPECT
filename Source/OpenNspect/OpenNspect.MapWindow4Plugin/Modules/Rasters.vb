Imports System.IO
Imports MapWinGIS

Module Rasters

    ''' <summary>
    ''' Gets the raster full path.
    ''' </summary>
    ''' <param name="pathOrDirectory">The full path or directory where the raster resides.</param><returns></returns>
    Public Function GetRasterFullPath(pathOrDirectory As String) As String
        If Path.HasExtension(pathOrDirectory) Then
            Return pathOrDirectory
        Else
            Return Path.Combine(pathOrDirectory, "sta.adf")
        End If
    End Function
    Public Function RasterExists(ByRef strRasterFileName As String) As Boolean
        Return File.Exists(GetRasterFullPath(strRasterFileName))
    End Function

    Public Function ReturnRaster(pathOrDirectory As String) As Grid
        Dim path = GetRasterFullPath(pathOrDirectory)

        If File.Exists(path) Then
            Dim g As New Grid
            If g.Open(path) Then
                Return g
            End If
        End If

        Return Nothing
    End Function

    Public Function ReturnPermanentRaster(ByRef pRaster As Grid, ByRef sOutputName As String) As Grid

        pRaster.Save()
        pRaster.Save(sOutputName)
        pRaster.Header.Projection = MapWindowPlugin.MapWindowInstance.Project.ProjectProjection

        Dim tmpraster As New Grid
        tmpraster.Open(sOutputName)
        Return tmpraster
    End Function

    Public Function AddRasterLayerToMapFromFileName(ByRef fileOrDirectory As String) As Boolean
        If Not Path.HasExtension(fileOrDirectory) Then
            ' we were passed a directory which represents an Esri Grid.
            fileOrDirectory = fileOrDirectory + "\sta.adf"
        End If

        MapWindowPlugin.MapWindowInstance.Layers.Add(fileOrDirectory)

        Return True
    End Function

End Module
