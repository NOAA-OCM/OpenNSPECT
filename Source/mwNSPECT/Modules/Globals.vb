'********************************************************************************************************
'File Name: Globals.vb
'Description: This screen is used to specify parameters for adding a new field.
'********************************************************************************************************
'The contents of this file are subject to the Mozilla Public License Version 1.1 (the "License"); 
'you may not use this file except in compliance with the License. You may obtain a copy of the License at 
'http://www.mozilla.org/MPL/ 
'Software distributed under the License is distributed on an "AS IS" basis, WITHOUT WARRANTY OF 
'ANY KIND, either express or implied. See the License for the specificlanguage governing rights and 
'limitations under the License. 
'
'Contributor(s): (Open source contributors should list themselves and their modifications here). 
'Nov 23, 2009:  Allen Anselmo allen.anselmo@gmail.com - 
'               Added licensing and comments to original code

Module Globals
#Region "Base Globals"
    Public g_MapWin As MapWindow.Interfaces.IMapWin
    Public g_handle As Integer
    Public g_StatusBar As System.Windows.Forms.StatusBarPanel

#End Region

#Region "Menu Globals"
    Public g_mnuNSPECTMain As String = "mnunspectMainMenu"
    Public g_mnuNSPECTAnalysis As String = "mnunspectAnalysis"
    Public g_mnuNSPECTAdvSettings As String = "mnunspectAdvancedSettings"
    Public g_mnuNSPECTAdvLand As String = "mnunspectLandCover"
    Public g_mnuNSPECTAdvPolutants As String = "mnunspectPollutants"
    Public g_mnuNSPECTAdvWQ As String = "mnunspectWaterQuality"
    Public g_mnuNSPECTAdvPrecip As String = "mnunspectPrecipitation"
    Public g_mnuNSPECTAdvWSDelin As String = "mnunspectWatershedDelineations"
    Public g_mnuNSPECTAdvSoils As String = "mnunspectSoils"
    Public g_mnuNSPECTHelp As String = "mnunspectHelp"

#End Region

#Region "Toolbar Globals"
    Public g_ToolBarName As String = "nspectToolbar"

#End Region

#Region "Global functions"

    ''' <summary>
    ''' Gets a field index from a shapefile by its name
    ''' </summary>
    ''' <param name="sf"></param>
    ''' <param name="currField"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getShapefileFieldByName(ByVal sf As MapWinGIS.Shapefile, ByVal currField As String) As Integer
        For i As Integer = 0 To sf.NumFields - 1
            If sf.Field(i).Name = currField Then
                Return i
            End If
        Next
        Return -1
    End Function

    ''' <summary>
    ''' Gets a layer index by a given path
    ''' </summary>
    ''' <param name="currPath"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getLayerIndexByPath(ByVal currPath As String) As Integer
        For i As Integer = 0 To g_MapWin.Layers.NumLayers - 1
            If g_MapWin.Layers.Item(i).FileName = currPath Then
                Return i
            End If
        Next
        Return -1
    End Function

    ''' <summary>
    ''' Function used to get a layer path by its name
    ''' </summary>
    ''' <param name="strName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getLayerPathByName(ByVal strName As String) As String
        Dim i As Integer
        For i = 0 To g_MapWin.Layers.NumLayers - 1
            If g_MapWin.Layers.Item(g_MapWin.Layers.GetHandle(i)).Name = strName Then
                Return g_MapWin.Layers.Item(g_MapWin.Layers.GetHandle(i)).FileName
            End If
        Next
        Return ""
    End Function

    ''' <summary>
    ''' Function used to get a layer name by its path
    ''' </summary>
    ''' <param name="strPath"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getLayerNameByPath(ByVal strPath As String) As String
        Dim i As Integer
        For i = 0 To g_MapWin.Layers.NumLayers - 1
            If g_MapWin.Layers.Item(g_MapWin.Layers.GetHandle(i)).FileName = strPath Then
                Return g_MapWin.Layers.Item(g_MapWin.Layers.GetHandle(i)).Name
            End If
        Next
        Return ""
    End Function

    ''' <summary>
    ''' Function used to get a layer index by its name
    ''' </summary>
    ''' <param name="strName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getLayerIndexByName(ByVal strName As String) As Integer
        Dim i As Integer
        For i = 0 To g_MapWin.Layers.NumLayers - 1
            If g_MapWin.Layers.Item(g_MapWin.Layers.GetHandle(i)).Name = strName Then
                Return g_MapWin.Layers.GetHandle(i)
            End If
        Next
        Return -1
    End Function

    ''' <summary>
    ''' Function used to check if a layer of a given path exists
    ''' </summary>
    ''' <param name="strPath"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function layerExists(ByVal strPath As String) As Boolean
        Dim i As Integer
        For i = 0 To g_MapWin.Layers.NumLayers - 1
            If g_MapWin.Layers.Item(g_MapWin.Layers.GetHandle(i)).FileName = strPath Then
                Return True
            End If
        Next
        Return False
    End Function
#End Region

End Module
