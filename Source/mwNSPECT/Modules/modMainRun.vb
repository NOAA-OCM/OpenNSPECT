Imports System.Data.OleDb
Module modMainRun
    ' *************************************************************************************
    ' *  Perot Systems Government Services
    ' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
    ' *  modMainRun
    ' *************************************************************************************
    ' *  Description:  Sets the analysis enviroment and many of the 'global' rasters used
    ' *  througout the calculations.
    ' *
    ' *
    ' *  Called By: frmPrj::CmdOK
    ' *
    ' *  Confession:  N-SPECT started back in 2003 and has lived a long life.  Not sure how,
    ' *  but it has.  It is a cobbled together piece of code now going on 6 years of age.
    ' *  Whomever takes it over, I offer my apologies.  It's a bit overly complex, no doubt about
    ' *  that, but have fun.  On the bright side, modUtil contains nearly ever function you
    ' *  would ever need in the raster world of ArcGIS.
    ' *************************************************************************************

    'Following represent variables that will be used for all analysis
    'Garnered primarily from the DEM Dataset

    'Public g_pSpatEnv As ESRI.ArcGIS.GeoAnalyst.IRasterAnalysisEnvironment 'Global Raster Analysis Env
    'Public g_pRasWorkspace As ESRI.ArcGIS.DataSourcesRaster.IRasterWorkspace 'Global Workspace- Raster Variety
    'Public g_pFeatWorkspace As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace 'Global Workspace-
    Public g_intDistanceUnits As Short 'Global Units 0 = meters, 1 = feet
    Public g_dblCellSize As Double 'Global Cell Size, again taken from DEM

    Public g_booLocalEffects As Boolean 'Did they check local effects?
    Public g_booSelectedPolys As Boolean 'Did they select n polygons for limiting analysis?
    'Public g_pSelectedPolyClip As ESRI.ArcGIS.Geometry.IGeometry 'Clip Geometry if they chose above
    Public g_intPrecipType As Short 'Precipitation Type; I, IA, II, III

    'The Public member datasets, to be used quite a bit
    'Public g_pDEMRaster As ESRI.ArcGIS.Geodatabase.IRaster 'DEM Raster
    'Public g_pFlowAccRaster As ESRI.ArcGIS.Geodatabase.IRaster 'Flow Accumulation
    'Public g_pFlowDirRaster As ESRI.ArcGIS.Geodatabase.IRaster 'Flow Direction
    Public g_strFlowDirFilename As String 'Flow Direction file name
    Public g_strLSFileName As String 'LS file name
    'Public g_pLSRaster As ESRI.ArcGIS.Geodatabase.IRaster 'LS Raster
    'Public g_pWaterShedFeatClass As ESRI.ArcGIS.Geodatabase.IFeatureClass 'WaterShed Poly featureclass
    'Public g_KFactorRaster As ESRI.ArcGIS.Geodatabase.IRaster 'K Factor DS, used in RUSLE or MUSLE

    'Public g_pGroupLayer As ESRI.ArcGIS.Carto.IGroupLayer 'Group Layer to handle and hold all output rasters
    Public g_dicMetadata As Generic.Dictionary(Of String, String) 'Global dictionary to hold name of layer, metadata process string
    Public g_clsXMLPrjFile As clsXMLPrjFile 'Global xml project file for metadata support

    Public Sub SetGlobalEnvironment(ByRef cmdWShed As OleDbCommand, ByRef strWorkspace As String, Optional ByRef pWShedlayer As MapWindow.Interfaces.Layer = Nothing)
        ''GOAL:  Set the analysis environment based on the properties of the DEM, and establish
        ''the other contributing datasets: Watersheds, flow direction, flow accumulation, length/slope
        ''Incoming Parameters:
        ''rsWShed: the recordset of the selected ws delineation...used to get paths to datasets
        ''strWorkspace: Workspace identified by user in main project window
        ''pMap: current map
        ''pWShedlayer: watershedlayer
        'Dim pEnv As ESRI.ArcGIS.GeoAnalyst.IRasterAnalysisEnvironment 'Env
        'Dim dblCellSize As Double 'Cell Size
        'Dim pEnvelope As ESRI.ArcGIS.Geometry.IEnvelope 'Extent
        'Dim pRasterProps As ESRI.ArcGIS.DataSourcesRaster.IRasterProps 'Raster Props
        'Dim pMaskGeoDataset As ESRI.ArcGIS.Geodatabase.IGeoDataset
        'Dim strError As String 'Error Handling String
        'Dim strDEM As String
        'Dim strWS As String
        'Dim strFlowDir As String
        'Dim strFlowAcc As String
        'Dim strLS As String
        'Dim intDistUnits As Short

        'strDEM = rsWShed.Fields("FilledDEMFileName").Value
        'strWS = rsWShed.Fields("wsfilename").Value
        'strFlowDir = rsWShed.Fields("FlowDirFileName").Value
        'strFlowAcc = rsWShed.Fields("FlowAccumFileName").Value
        'strLS = rsWShed.Fields("LSFileName").Value
        'intDistUnits = rsWShed.Fields("DEMGridUnits").Value


        ''STEP 1: Get the workspaces set
        'g_pRasWorkspace = modUtil.SetRasterWorkspace(strWorkspace)
        'g_pFeatWorkspace = modUtil.SetFeatureShapeWorkspace(strWorkspace)

        'g_strFlowDirFilename = strFlowDir

        ''STEP 2: Establish the environment
        'If modUtil.RasterExists(strDEM) Then
        '    g_pDEMRaster = modUtil.ReturnRaster(strDEM)
        'Else
        '    strError = "DEM Raster Does Not Exist: " & strDEM
        'End If

        ''STEP 6: Set the other Datasets
        ''Begin with Water shed, let a featureclass
        'g_pWaterShedFeatClass = modUtil.ReturnFeatureClass(strWS)

        'If g_booSelectedPolys Then
        '    'UPGRADE_WARNING: Couldn't resolve default property of object g_pFeatWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '    pMaskGeoDataset = ReturnAnalysisMask(pWShedlayer, pMap, g_pFeatWorkspace, strWS)
        'Else
        '    pMaskGeoDataset = g_pDEMRaster
        'End If

        ''STEP 3: With the Rasterdataset set, get its properties
        'pRasterProps = g_pDEMRaster
        ''Get cell size and envelope
        'dblCellSize = pRasterProps.MeanCellSize.X
        ''Set the global cell size
        'g_dblCellSize = dblCellSize
        ''Set pEnvelope = pRasterProps.Extent
        ''Either pMaskGeoDataset will be the DEM or the mask
        'pEnvelope = pMaskGeoDataset.Extent

        ''Init pEnv
        ''Set pEnv = New RasterAnalysis
        'g_pSpatEnv = New ESRI.ArcGIS.GeoAnalyst.RasterAnalysis

        ''STEP 4: Now set the analysis environment
        'With g_pSpatEnv
        '    .SetCellSize(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvValue, dblCellSize) 'Cell Size
        '    .SetExtent(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvValue, pEnvelope, True) 'Extent
        '    .Mask = pMaskGeoDataset 'Mask
        '    'UPGRADE_WARNING: Couldn't resolve default property of object g_pSpatEnv.OutSpatialReference. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '    .OutSpatialReference = pRasterProps.SpatialReference 'Projection
        '    'UPGRADE_WARNING: Couldn't resolve default property of object g_pSpatEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '    .OutWorkspace = g_pRasWorkspace 'Outworkspace
        '    .SetAsNewDefaultEnvironment()
        'End With

        ''Set g_pSpatEnv = pEnv

        ''STEP 5: Set global units
        'Select Case intDistUnits
        '    Case 0 'Meters
        '        g_intDistanceUnits = 0
        '    Case 1 'Feed
        '        g_intDistanceUnits = 1
        'End Select

        ''Flow Direction
        'If modUtil.RasterExists(strFlowDir) Then
        '    g_pFlowDirRaster = modUtil.ReturnRaster(strFlowDir)
        'Else
        '    strError = "Flow Direction Raster Does Not Exist: " & strFlowDir
        'End If

        ''FlowAccumulation
        'If modUtil.RasterExists(strFlowAcc) Then
        '    g_pFlowAccRaster = modUtil.ReturnRaster(strFlowAcc)
        'Else
        '    strError = "Flow Accumulation Raster Does Not Exist: " & strFlowAcc
        'End If

        ''Length Slope
        'If modUtil.RasterExists(strLS) Then
        '    g_pLSRaster = modUtil.ReturnRaster(strLS)
        '    g_strLSFileName = strLS
        'Else
        '    strError = "Length Slope raster does not Exist: " & strLS
        'End If

        'If Len(strError) > 0 Then
        '    MsgBox(strError, MsgBoxStyle.Critical, "Missing Data")
        'End If

        ''Init the group layer
        'g_pGroupLayer = New ESRI.ArcGIS.Carto.GroupLayer

        ''Cleanup
        ''UPGRADE_NOTE: Object pEnv may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        'pEnv = Nothing
        ''UPGRADE_NOTE: Object pRasterProps may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        'pRasterProps = Nothing
        ''UPGRADE_NOTE: Object pEnvelope may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        'pEnvelope = Nothing

    End Sub


    Private Function ReturnAnalysisMask(ByRef pLayer As MapWindow.Interfaces.Layer, ByRef strBasinFeatClass As String) As MapWinGIS.Shapefile
        ''Incoming
        ''pLayer: Layer user has chosen as being the one from which the selected polys will come
        ''pMap: current map
        ''pWorkspace:  place to put the exported selected sheds
        ''strBasinFeatClass: string file location of BasinPoly.shp

        'Dim pFLayer As ESRI.ArcGIS.Carto.IFeatureLayer
        'Dim pFC As ESRI.ArcGIS.Geodatabase.IFeatureClass
        'Dim pINFeatureClassName As ESRI.ArcGIS.Geodatabase.IFeatureClassName
        'Dim pDataset As ESRI.ArcGIS.Geodatabase.IDataset
        'Dim pInDsName As ESRI.ArcGIS.Geodatabase.IDatasetName
        'Dim pFSel As ESRI.ArcGIS.Carto.IFeatureSelection
        'Dim pSelSet As ESRI.ArcGIS.Geodatabase.ISelectionSet
        'Dim pFeatureClassName As ESRI.ArcGIS.Geodatabase.IFeatureClassName
        'Dim pOutDatasetName As ESRI.ArcGIS.Geodatabase.IDatasetName
        'Dim pWorkspaceName As ESRI.ArcGIS.Geodatabase.IWorkspaceName
        'Dim pExportOp As ESRI.ArcGIS.GeoDatabaseUI.IExportOperation
        'Dim pFeatWS As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace
        'Dim pFeatClass As ESRI.ArcGIS.Geodatabase.IFeatureClass

        'Dim pSelectGeometry As ESRI.ArcGIS.Geometry.IGeometry
        'Dim pBasinFeatClass As ESRI.ArcGIS.Geodatabase.IFeatureClass
        'Dim pSpatialFilter As ESRI.ArcGIS.Geodatabase.ISpatialFilter
        'Dim pBasinSelSet As ESRI.ArcGIS.Geodatabase.ISelectionSet

        ''Set the layer to the incoming one and get featureclass
        'pFLayer = pLayer

        ''Get the selection set
        'pFSel = pFLayer
        'pSelSet = pFSel.SelectionSet

        ''Make a call to the function below to return a unioned geometry of any and all selected features
        'pSelectGeometry = ReturnSelectGeometry(pSelSet)
        'g_pSelectedPolyClip = pSelectGeometry

        ''Make a call to get the BasinPoly featureclass using the name sent over
        'pBasinFeatClass = modUtil.ReturnFeatureClass(strBasinFeatClass)

        ''Now we select all basinpolys that intersect the unioned geometry we got earlier
        'pSpatialFilter = New ESRI.ArcGIS.Geodatabase.SpatialFilter

        'With pSpatialFilter
        '    .Geometry = pSelectGeometry
        '    .GeometryField = pBasinFeatClass.ShapeFieldName
        '    .SpatialRel = ESRI.ArcGIS.Geodatabase.esriSpatialRelEnum.esriSpatialRelIntersects
        'End With

        ''UPGRADE_WARNING: Couldn't resolve default property of object pSpatialFilter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'pBasinSelSet = pBasinFeatClass.Select(pSpatialFilter, ESRI.ArcGIS.Geodatabase.esriSelectionType.esriSelectionTypeIDSet, ESRI.ArcGIS.Geodatabase.esriSelectionOption.esriSelectionOptionNormal, pWorkspace)

        ''Get the FcName from the featureclass
        'pDataset = pBasinFeatClass

        ''Set classname and dataset
        'pINFeatureClassName = pDataset.FullName
        'pInDsName = pINFeatureClassName

        ''Create a new feature class name
        '' Define the output feature class name
        'pFeatureClassName = New ESRI.ArcGIS.Geodatabase.FeatureClassName
        'pOutDatasetName = pFeatureClassName

        ''UPGRADE_WARNING: Couldn't resolve default property of object pWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'pOutDatasetName.Name = modUtil.GetUniqueFeatureClassName(pWorkspace, "selshed")

        'pWorkspaceName = New ESRI.ArcGIS.Geodatabase.WorkspaceName
        'pWorkspaceName.PathName = pWorkspace.PathName
        'pWorkspaceName.WorkspaceFactoryProgID = "esriDataSourcesFile.shapefileworkspacefactory.1"

        'pOutDatasetName.WorkspaceName = pWorkspaceName

        'With pFeatureClassName
        '    .FeatureType = ESRI.ArcGIS.Geodatabase.esriFeatureType.esriFTSimple
        '    .ShapeType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolygon
        '    .ShapeFieldName = "Shape"
        'End With

        ''Export
        'pExportOp = New ESRI.ArcGIS.GeoDatabaseUI.ExportOperation
        ''UPGRADE_WARNING: Couldn't resolve default property of object pOutDatasetName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'pExportOp.ExportFeatureClass(pInDsName, Nothing, pBasinSelSet, Nothing, pOutDatasetName, 0)

        'pFeatWS = pWorkspace
        'pFeatClass = pFeatWS.OpenFeatureClass(pOutDatasetName.Name)
        'ReturnAnalysisMask = pFeatClass

        ''Cleanup
        ''UPGRADE_NOTE: Object pFLayer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        'pFLayer = Nothing
        ''UPGRADE_NOTE: Object pFC may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        'pFC = Nothing
        ''UPGRADE_NOTE: Object pINFeatureClassName may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        'pINFeatureClassName = Nothing
        ''UPGRADE_NOTE: Object pDataset may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        'pDataset = Nothing
        ''UPGRADE_NOTE: Object pInDsName may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        'pInDsName = Nothing
        ''UPGRADE_NOTE: Object pFSel may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        'pFSel = Nothing
        ''UPGRADE_NOTE: Object pSelSet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        'pSelSet = Nothing
        ''UPGRADE_NOTE: Object pFeatureClassName may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        'pFeatureClassName = Nothing
        ''UPGRADE_NOTE: Object pOutDatasetName may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        'pOutDatasetName = Nothing
        ''UPGRADE_NOTE: Object pWorkspaceName may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        'pWorkspaceName = Nothing
        ''UPGRADE_NOTE: Object pExportOp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        'pExportOp = Nothing
        ''UPGRADE_NOTE: Object pFeatWS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        'pFeatWS = Nothing

    End Function

    Public Function CheckMultiPartPolygon(ByVal pPolygon As MapWinGIS.Shape) As Boolean

        ''UPGRADE_WARNING: Couldn't resolve default property of object pPolygon.ExteriorRingCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'If pPolygon.ExteriorRingCount > 1 Then
        '    CheckMultiPartPolygon = True
        'Else
        '    CheckMultiPartPolygon = False
        'End If

    End Function


    Public Function ReturnSelectGeometry(ByRef pInSelectionSet As MapWindow.Interfaces.SelectInfo) As MapWinGIS.Shape

        'Dim pFeature As ESRI.ArcGIS.Geodatabase.IFeature
        'Dim pTopoSimple As ESRI.ArcGIS.Geometry.ITopologicalOperator 'Simplifier
        'Dim pNewTopo As ESRI.ArcGIS.Geometry.ITopologicalOperator 'Union
        'Dim pCurrPolygon As ESRI.ArcGIS.Geometry.IPolygon

        'Dim pSelectionSet As ESRI.ArcGIS.Geodatabase.ISelectionSet
        'pSelectionSet = pInSelectionSet 'pFeatureSelection.SelectionSet

        'Dim pFeatureCursor As ESRI.ArcGIS.Geodatabase.IFeatureCursor
        ''UPGRADE_WARNING: Couldn't resolve default property of object pFeatureCursor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'pSelectionSet.Search(Nothing, True, pFeatureCursor)

        'If pSelectionSet.Count = 0 Then
        '    Exit Function
        'End If

        'pTopoSimple = New ESRI.ArcGIS.Geometry.Polygon
        'pNewTopo = New ESRI.ArcGIS.Geometry.Polygon

        'pFeature = pFeatureCursor.NextFeature

        'Do While Not pFeature Is Nothing

        '    pCurrPolygon = pFeature.Shape
        '    pTopoSimple = pCurrPolygon
        '    pTopoSimple.Simplify()

        '    'UPGRADE_WARNING: Couldn't resolve default property of object pTopoSimple. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '    pNewTopo = pNewTopo.Union(pTopoSimple)
        '    pFeature = pFeatureCursor.NextFeature

        'Loop

        'ReturnSelectGeometry = pNewTopo

        ''Cleanup
        ''UPGRADE_NOTE: Object pFeature may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        'pFeature = Nothing
        ''UPGRADE_NOTE: Object pTopoSimple may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        'pTopoSimple = Nothing
        ''UPGRADE_NOTE: Object pCurrPolygon may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        'pCurrPolygon = Nothing
        ''UPGRADE_NOTE: Object pSelectionSet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        'pSelectionSet = Nothing

    End Function
End Module