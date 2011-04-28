﻿Imports System.IO
Imports MapWinGIS

Module Features

    Public Function FeatureExists(ByRef strFeatureFileName As String) As Boolean
        Return File.Exists(GetFeatureFullPath(strFeatureFileName))
    End Function

    Private Function GetFeatureFullPath(pathOrDirectory As String) As String
        If Path.HasExtension(pathOrDirectory) Then
            Return pathOrDirectory
        Else
            Return pathOrDirectory + ".shp"
        End If
    End Function
    Public Function ReturnFeature(ByRef pathOrDirectory As String) As Shapefile
        Dim path = GetFeatureFullPath(pathOrDirectory)

        If File.Exists(path) Then
            Dim sf As New Shapefile
            If sf.Open(path) Then
                Return sf
            End If
        End If

        Return Nothing

    End Function

    Public Function AddFeatureLayerToMapFromFileName(ByRef pathOrDirectory As String, Optional ByRef strLyrName As String = "") _
        As Boolean
        Try
            Dim path = GetFeatureFullPath(pathOrDirectory)

            If File.Exists(path) Then
                If strLyrName <> "" Then
                    MapWindowPlugin.MapWindowInstance.Layers.Add(pathOrDirectory, strLyrName)
                Else
                    MapWindowPlugin.MapWindowInstance.Layers.Add(pathOrDirectory)
                End If
            Else
                Return False
            End If

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
End Module
