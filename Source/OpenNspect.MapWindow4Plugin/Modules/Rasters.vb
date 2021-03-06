﻿Imports System.IO
Imports MapWinGIS

Module Rasters

    ''' <summary>
    ''' Gets the raster full path.
    ''' </summary>
    ''' <param name="pathOrDirectory">The full path or directory where the raster resides.</param>
    ''' <returns>The full path with sta.adf</returns>
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

    ''' <summary>
    ''' Gets the absolute raster table path including ...dbf
    ''' </summary>
    ''' <param name="raster">The raster.</param><returns></returns>
    Public Function GetRasterTablePath(ByRef raster As Grid) As String
        Dim tablepath As String
        Dim lcPath As String = raster.Filename
        'Dim TableExist As Boolean
        If Path.GetFileName(lcPath) = "sta.adf" Or Path.GetFileName(lcPath) = "sta.bmp" Then
            tablepath = Path.GetDirectoryName(lcPath) + ".dbf"
        Else
            tablepath = Path.ChangeExtension(lcPath, ".dbf")
        End If
        If Not File.Exists(tablepath) Then
            '    TableExist = True
            'Else
            BuildTable(raster, tablepath)
            'If File.Exists(tablepath) Then
            '    TableExist = True
            'Else
            '    TableExist = BuildTable(raster, tablepath)
            'End If
        End If
        Return tablepath
    End Function
    'Public Function GetRasterTablePath(ByRef raster As Grid) As String
    '    Dim tablepath As String
    '    Dim lcPath As String = raster.Filename
    '    If Path.GetFileName(lcPath) = "sta.adf" Or Path.GetFileName(lcPath) = "sta.bmp" Then
    '        tablepath = Path.GetDirectoryName(lcPath) + ".dbf"
    '    Else
    '        tablepath = Path.ChangeExtension(lcPath, ".dbf")
    '    End If
    '    Return tablepath
    'End Function

    Public Function GetFieldIndex(ByVal mwTable As Table) As Short
        Dim FieldIndex As Short = -1
        For fidx As Integer = 0 To mwTable.NumFields - 1
            If mwTable.Field(fidx).Name.ToLower = "value" Then
                FieldIndex = fidx
                Exit For
            End If
        Next
        Return FieldIndex
    End Function
    Public Function CopyRaster(ByRef pRaster As Grid, ByRef fileName As String) As Grid

        pRaster.Save()
        pRaster.Save(fileName)

        Dim newRaster As New Grid
        newRaster.Open(fileName)
        Return newRaster
    End Function

    Public Function AddRasterLayerToMapFromFileName(ByRef fileOrDirectory As String) As Boolean
        If Not Path.HasExtension(fileOrDirectory) Then
            ' we were passed a directory which represents an Esri Grid.
            fileOrDirectory = Path.Combine(fileOrDirectory, "sta.adf")
        End If

        MapWindowPlugin.MapWindowInstance.Layers.Add(fileOrDirectory)

        Return True
    End Function

End Module
