'********************************************************************************************************
'File Name: modUtil.vb
'Description: Many global utility functions used throughout the plugin
'********************************************************************************************************
'The contents of this file are subject to the Mozilla Public License Version 1.1 (the "License"); 
'you may not use this file except in compliance with the License. You may obtain a copy of the License at 
'http://www.mozilla.org/MPL/ 
'Software distributed under the License is distributed on an "AS IS" basis, WITHOUT WARRANTY OF 
'ANY KIND, either express or implied. See the License for the specificlanguage governing rights and 
'limitations under the License. 
'
'Note: This code was converted from the vb6 NSPECT ArcGIS extension and so bears many of the old comments
'in the files where it was possible to leave them.
'
'Contributor(s): (Open source contributors should list themselves and their modifications here). 
'Oct 20, 2010:  Allen Anselmo allen.anselmo@gmail.com - 
'               Added licensing and comments to code
Imports System.Collections.Generic
Imports System.Windows.Forms
Imports System.IO
Imports MapWindow.Interfaces
Imports MapWinGeoProc
Imports MapWinGIS
Imports OpenNspect.Xml

Module Utilities
    Public g_nspectPath As String
    Public g_nspectDocPath As String

    Public g_cbMainForm As MainForm
    Public g_MainForm As Form

    Public g_comp As CompareOutputsForm
    Public g_luscen As EditLandUseScenario

    'Intermediate and output extensions
    Public Const OutputGridExt As String = ".tif"
    Public Const FinalOutputGridExt As String = ".tif"
    Public Const TAUDEMGridExt As String = ".tif"

    Public g_TempFilesToDel As New List(Of String)

    Public g_boolAddCoeff As Boolean
    Public g_boolCopyCoeff As Boolean
    'True: called frmPollutants, False: called frmNewPollutants
    Public g_boolNewWShed As Boolean
    'True: New WaterShed form called from frmPrj

    'Management Scenario variables::frmPrjCalc
    Public g_ManagementScenarioLUScenFileName As String
    Public g_ManagementScenarioRowNumber As String

    'Returns a filename given for example C:\temp\dataset returns dataset
    Public Function SplitFileName(ByRef sWholeName As String) As String
        SplitFileName = ""
        Try
            Dim pos As Short
            Dim sT As Object
            Dim sName As String
            pos = InStrRev(sWholeName, "\")
            If pos > 0 Then
                sT = Mid(sWholeName, 1, pos - 1)
                If pos = Len(sWholeName) Then
                    Exit Function
                End If
                sName = Mid(sWholeName, pos + 1, Len(sWholeName) - Len(sT))
                pos = InStr(sName, ".")
                If pos > 0 Then
                    SplitFileName = Mid(sName, 1, pos - 1)
                Else
                    SplitFileName = sName
                End If
            End If
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Function

    Public Function GetRasterDistanceUnits(ByRef strLayerName As String) As Short
        For i As Integer = 0 To MapWindowPlugin.MapWindowInstance.Layers.NumLayers - 1
            If MapWindowPlugin.MapWindowInstance.Layers(i).Name = strLayerName Then
                Dim proj As String = MapWindowPlugin.MapWindowInstance.Layers(i).Projection
                If Not String.IsNullOrEmpty(proj) Then
                    If proj.Contains("units=m") Then
                        Return 0
                    ElseIf proj.Contains("units=ft") Then
                        Return 1
                    End If
                Else
                    MsgBox("The GRID you have choosen has no spatial reference information.  Please define a projection before continuing.", MsgBoxStyle.Exclamation, "No Project Information Detected")
                    Return -1
                End If
            End If
        Next
        Return -1
    End Function

    Public Function GetIndexOfEntry(ByRef strList As String, ByRef cbo As ComboBox) As Short
        Return cbo.Items.IndexOf(strList)
    End Function

    Public Function LayerLoadedInMap(ByRef layerName As String) As Boolean
        Try
            For Each layer As Layer In MapWindowPlugin.MapWindowInstance.Layers
                If layer.Name = layerName Then
                    Return True
                End If
            Next

            Return False
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Function

    Public Function ExistsLayerInMapByFileName(ByRef strName As String) As Boolean
        For layerIndex As Integer = 0 To MapWindowPlugin.MapWindowInstance.Layers.NumLayers - 1
            Dim pLayer As Layer = MapWindowPlugin.MapWindowInstance.Layers(MapWindowPlugin.MapWindowInstance.Layers.GetHandle(layerIndex))

            If Trim(LCase(pLayer.FileName)) = Trim(LCase(strName)) Then Return True
        Next
        Return False
    End Function

    Public Function GetLayerIndex(ByRef strLayerName As String) As Integer
        GetLayerIndex = -1
        Try
            For lngLyrIndex As Integer = 0 To MapWindowPlugin.MapWindowInstance.Layers.NumLayers - 1
                Dim pLayer As Layer = MapWindowPlugin.MapWindowInstance.Layers(MapWindowPlugin.MapWindowInstance.Layers.GetHandle(lngLyrIndex))

                If Trim(LCase(pLayer.Name)) = Trim(LCase(strLayerName)) Then
                    GetLayerIndex = lngLyrIndex
                    Exit For
                End If
            Next
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Function

    Public Function GetLayerIndexByFilename(ByRef strLayerFileName As String) As Integer
        GetLayerIndexByFilename = -1
        Try
            For lngLyrIndex As Integer = 0 To MapWindowPlugin.MapWindowInstance.Layers.NumLayers - 1
                Dim pLayer As Layer = MapWindowPlugin.MapWindowInstance.Layers(MapWindowPlugin.MapWindowInstance.Layers.GetHandle(lngLyrIndex))

                If Trim(LCase(pLayer.FileName)) = Trim(LCase(strLayerFileName)) Then
                    GetLayerIndexByFilename = lngLyrIndex
                    Exit For
                End If
            Next
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Function

    Public Function GetLayerFilename(ByRef strLayerName As String) As String
        GetLayerFilename = ""
        Try
            For lngLyrIndex As Integer = 0 To MapWindowPlugin.MapWindowInstance.Layers.NumLayers - 1
                Dim pLayer As Layer = MapWindowPlugin.MapWindowInstance.Layers(MapWindowPlugin.MapWindowInstance.Layers.GetHandle(lngLyrIndex))

                If Trim(LCase(pLayer.Name)) = Trim(LCase(strLayerName)) Then
                    GetLayerFilename = pLayer.FileName

                    If GetLayerFilename.EndsWith("sta.adf") Then
                        GetLayerFilename = Path.GetDirectoryName(GetLayerFilename)
                    End If

                    Exit For
                End If
            Next
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Function


    Public Function AddInputFromGxBrowser(ByRef strType As String) As String
        Try
            Dim pFilter As String = ""

            Select Case strType
                Case "Feature"
                    Dim shp As New Shapefile
                    pFilter = shp.CdlgFilter
                Case "Raster"
                    Dim g As New Grid
                    pFilter = g.CdlgFilter
            End Select

            Using dlgopen As New OpenFileDialog()
                dlgopen.Title = String.Format("Open a {0} file", strType)
                dlgopen.Filter = pFilter

                If dlgopen.ShowDialog = DialogResult.OK Then
                    Return dlgopen.FileName
                End If
            End Using
        Catch ex As Exception
            HandleError(ex)
        End Try
        Return ""
    End Function

    Public Function AddInputFromGxBrowserText(ByRef txtInput As TextBox, ByRef strTitle As String) As Grid
        AddInputFromGxBrowserText = Nothing
        Try
            Using dlgOpen As New OpenFileDialog()
                Dim g As New Grid()
                dlgOpen.Filter = g.CdlgFilter
                dlgOpen.Title = strTitle
                If dlgOpen.ShowDialog = DialogResult.OK Then
                    txtInput.Text = dlgOpen.FileName
                    g.Open(dlgOpen.FileName)
                    Return g
                Else
                    Return Nothing
                End If
            End Using
        Catch ex As Exception
            MsgBox("The file you have choosen is not a valid GRID dataset.  Please select another.", MsgBoxStyle.Critical, "Invalid Data Type")
        End Try

    End Function

    Public Function CheckSpatialReference(ByRef pRasGeoDataset As Grid) As String
        Try
            If Not pRasGeoDataset Is Nothing Then
                Dim strprj As String = pRasGeoDataset.Header.Projection
                If Not String.IsNullOrEmpty(strprj) Then
                    Return strprj
                Else
                    If Path.GetFileName(pRasGeoDataset.Filename) = "sta.adf" Then
                        Dim rasPath As String = Path.Combine(Path.GetDirectoryName(pRasGeoDataset.Filename), "prj.adf")
                        If File.Exists(rasPath) Then
                            Using infile As New StreamReader(rasPath)
                                'TODO: Temporary measure that allows at least units to be recognized
                                If infile.ReadToEnd.Contains("METERS") Then
                                    Return "units=m"
                                Else
                                    Return "units=ft"
                                End If
                            End Using
                            Return "convert this prj.adf to real projection"
                        End If
                    End If
                End If
            End If

        Catch ex As Exception
            HandleError(ex)
        End Try
        Return ""
    End Function

    Public Function ClipBySelectedPoly(ByRef pGridToClip As Grid, ByVal pSelectedPolyClip As MapWinGIS.Shape, ByVal outputFileName As String) As Grid
        Dim strtmp1 As String = Path.GetTempFileName
        g_TempFilesToDel.Add(strtmp1)
        strtmp1 = strtmp1 + OutputGridExt
        g_TempFilesToDel.Add(strtmp1)
        DataManagement.DeleteGrid(strtmp1)
        pGridToClip.Save()
        pGridToClip.Save(strtmp1)
        pGridToClip.Header.Projection = MapWindowPlugin.MapWindowInstance.Project.ProjectProjection

        SpatialOperations.ClipGridWithPolygon(strtmp1, pSelectedPolyClip, outputFileName)

        Dim out As New Grid
        out.Open(outputFileName)
        Return out
    End Function

    Public Function BrowseForFileName(ByRef strType As String) As String

        Dim pfilter As String

        Select Case strType
            Case "Feature"
                Dim shp As New Shapefile
                pfilter = shp.CdlgFilter
            Case "Raster"
                Dim g As New Grid
                pfilter = g.CdlgFilter
            Case Else
                pfilter = ""
        End Select

        Using dlg As New OpenFileDialog()
            dlg.Filter = pfilter
            dlg.Title = "Open " + strType
            If dlg.ShowDialog() = DialogResult.OK Then
                Return dlg.FileName
            Else
                Return ""
            End If
        End Using

    End Function

    Public Function ExportSelectedFeatures(ByVal SelectLyrPath As String, ByRef SelectedShapes As List(Of Integer)) As String
        ' Modified from http://www.mapwindow.org/wiki/index.php/MapWinGIS:SampleCode-VB_Net:ExportSelectedShapes
        Dim Result As Boolean

        Dim sFileName, sLayerType As String

        Dim myShapeFile, newShapefile As Shapefile
        Dim myShape As MapWinGIS.Shape
        Dim ShapefileType As ShpfileType
        Dim iShapeHandle, iFieldCnt As Integer

        ExportSelectedFeatures = Nothing
        'First check to see if any features have been selected
        If SelectedShapes Is Nothing OrElse SelectedShapes.Count = 0 Then
            Exit Function
        End If

        Try
            myShapeFile = New Shapefile
            myShapeFile.Open(SelectLyrPath)

            'Determine if shape is polygon, line, or point
            sLayerType = LCase(MapWindowPlugin.MapWindowInstance.Layers(MapWindowPlugin.MapWindowInstance.Layers.CurrentLayer).LayerType.ToString)
            If InStr(sLayerType, "line", CompareMethod.Text) > 0 Then
                ShapefileType = ShpfileType.SHP_POLYLINE
            ElseIf InStr(sLayerType, "polygon", CompareMethod.Text) > 0 Then
                ShapefileType = ShpfileType.SHP_POLYGON
            ElseIf InStr(sLayerType, "point", CompareMethod.Text) > 0 Then
                ShapefileType = ShpfileType.SHP_POINT
            End If

            sFileName = GetUniqueFileName("selpoly", g_XmlPrjFile.ProjectWorkspace, ".shp")

            'Create the new shapefile
            newShapefile = New Shapefile
            newShapefile.CreateNew(sFileName, ShapefileType)
            newShapefile.Projection = MapWindowPlugin.MapWindowInstance.Project.ProjectProjection

            'The new shapefile has no fields at this point
            For iFieldCnt = 0 To myShapeFile.NumFields - 1
                newShapefile.EditInsertField(myShapeFile.Field(iFieldCnt), iFieldCnt)
            Next iFieldCnt

            'Start an edit session in the shapefile
            newShapefile.StartEditingShapes(True, Nothing)

            'Iterate through each of the selected feature
            For i As Integer = 0 To SelectedShapes.Count - 1
                'Set to the selected shape
                myShape = myShapeFile.Shape(SelectedShapes(i))

                'insert the selected shape
                iShapeHandle = newShapefile.NumShapes
                Result = newShapefile.EditInsertShape(myShape, iShapeHandle)

                'Populate the aspatial data
                For iFieldCnt = 0 To myShapeFile.NumFields - 1
                    newShapefile.EditCellValue(iFieldCnt, iShapeHandle, myShapeFile.CellValue(iFieldCnt, SelectedShapes(i)))
                Next iFieldCnt
            Next i

            newShapefile.StopEditingShapes()
            newShapefile.Close()
            myShapeFile.Close()
            Return sFileName
        Catch ex As Exception
            MsgBox("Error in exporting selected features.", MsgBoxStyle.Exclamation, "Exporting Selected Error")
        End Try

    End Function

    Public Function AddOutputGridLayer(ByRef outRast As Grid, ByVal ColorString As String, ByVal UseStretch As Boolean, ByVal LayerName As String, ByVal OutputType As String, ByVal OutputGroup As Integer, ByRef OutputItems As OutputItems) As Boolean
        Dim cs As GridColorScheme
        If UseStretch = True Then
            cs = GetRasterStretchColorRampCS(outRast, ColorString)
        Else
            cs = GetUniqueRasterRenderer()
        End If

        Dim lyr As Layer = MapWindowPlugin.MapWindowInstance.Layers.Add(outRast, cs, LayerName)
        lyr.Visible = False
        If OutputGroup <> -1 Then
            lyr.MoveTo(0, OutputGroup)
        Else
            lyr.MoveTo(0, g_pGroupLayer)
        End If
        If Not OutputItems Is Nothing Then
            Dim item As New OutputItem
            item.strPath = outRast.Filename
            item.strName = LayerName
            item.strType = OutputType
            item.strColor = ColorString
            item.booUseStretch = UseStretch
            OutputItems.Add(item)
        End If
    End Function
    Public Function GetUniqueFileName(ByRef Name As String, ByRef folderPath As String, ByVal Extension As String) As String
        Dim i As Integer = 0
        Dim nameAttempt As String

        Do
            i = i + 1
            nameAttempt = folderPath + Path.DirectorySeparatorChar + Name + i.ToString + Extension
        Loop While File.Exists(nameAttempt) And i < 1000

        If i < 1000 Then Return nameAttempt Else Throw New InvalidOperationException("Workspace is too large. Could not get Unique Filename.")
    End Function
End Module
