Imports System.IO

Module MapWindowUI
    Public Function AddRasterLayerToMapFromFileName(ByRef directory As String) As Boolean
        MapWindowPlugin.MapWindowInstance.Layers.Add(directory + "\sta.adf")
        Return True
        ' todo: was this for backwards compatibility?
        'Try
        '    If Path.GetExtension(directory) <> "" Then ' if this is a file
        '        MapWindowPlugin.MapWindowInstance.Layers.Add(directory)
        '    Else ' we were passed a directory
        '        MapWindowPlugin.MapWindowInstance.Layers.Add(directory + "\sta.adf")
        '    End If
        '    Return True
        'Catch ex As Exception
        '    Return False
        'End Try

    End Function
End Module
