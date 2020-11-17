Attribute VB_Name = "modMainRun"
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
Option Explicit

Public g_pSpatEnv As IRasterAnalysisEnvironment     'Global Raster Analysis Env
Public g_pRasWorkspace As IRasterWorkspace          'Global Workspace- Raster Variety
Public g_pFeatWorkspace As IFeatureWorkspace        'Global Workspace-
Public g_intDistanceUnits As Integer                'Global Units 0 = meters, 1 = feet
Public g_dblCellSize As Double                      'Global Cell Size, again taken from DEM

Public g_booLocalEffects As Boolean                 'Did they check local effects?
Public g_booSelectedPolys As Boolean                'Did they select n polygons for limiting analysis?
Public g_pSelectedPolyClip As IGeometry             'Clip Geometry if they chose above
Public g_intPrecipType As Integer                   'Precipitation Type; I, IA, II, III

'The Public member datasets, to be used quite a bit
Public g_pDEMRaster As IRaster                      'DEM Raster
Public g_pFlowAccRaster As IRaster                  'Flow Accumulation
Public g_pFlowDirRaster As IRaster                  'Flow Direction
Public g_strFlowDirFilename As String               'Flow Direction file name
Public g_strLSFileName As String                    'LS file name
Public g_pLSRaster As IRaster                       'LS Raster
Public g_pWaterShedFeatClass As IFeatureClass       'WaterShed Poly featureclass
Public g_KFactorRaster As IRaster                   'K Factor DS, used in RUSLE or MUSLE

Public g_pGroupLayer As IGroupLayer                 'Group Layer to handle and hold all output rasters
Public g_dicMetadata As Dictionary                  'Global dictionary to hold name of layer, metadata process string
Public g_clsXMLPrjFile As clsXMLPrjFile             'Global xml project file for metadata support

Public Sub SetGlobalEnvironment(rsWShed As ADODB.Recordset, strWorkspace As String, _
                                pMap As IMap, Optional pWShedlayer As ILayer)
    'GOAL:  Set the analysis environment based on the properties of the DEM, and establish
    'the other contributing datasets: Watersheds, flow direction, flow accumulation, length/slope
    'Incoming Parameters:
    'rsWShed: the recordset of the selected ws delineation...used to get paths to datasets
    'strWorkspace: Workspace identified by user in main project window
    'pMap: current map
    'pWShedlayer: watershedlayer
    Dim pEnv As IRasterAnalysisEnvironment      'Env
    Dim dblCellSize As Double                   'Cell Size
    Dim pEnvelope As IEnvelope                  'Extent
    Dim pRasterProps As IRasterProps            'Raster Props
    Dim pMaskGeoDataset As IGeoDataset
    Dim strError As String                      'Error Handling String
    Dim strDEM As String
    Dim strWS As String
    Dim strFlowDir As String
    Dim strFlowAcc As String
    Dim strLS As String
    Dim intDistUnits As Integer
    
    strDEM = rsWShed!FilledDEMFileName
    strWS = rsWShed!wsfilename
    strFlowDir = rsWShed!FlowDirFileName
    strFlowAcc = rsWShed!FlowAccumFileName
    strLS = rsWShed!LSFileName
    intDistUnits = rsWShed!DEMGridUnits
    

    'STEP 1: Get the workspaces set
    Set g_pRasWorkspace = modUtil.SetRasterWorkspace(strWorkspace)
    Set g_pFeatWorkspace = modUtil.SetFeatureShapeWorkspace(strWorkspace)
    
    g_strFlowDirFilename = strFlowDir
    
    'STEP 2: Establish the environment
    If modUtil.RasterExists(strDEM) Then
        Set g_pDEMRaster = modUtil.ReturnRaster(strDEM)
    Else
        strError = "DEM Raster Does Not Exist: " & strDEM
    End If
    
    'STEP 6: Set the other Datasets
    'Begin with Water shed, let a featureclass
    Set g_pWaterShedFeatClass = modUtil.ReturnFeatureClass(strWS)
    
    If g_booSelectedPolys Then
        Set pMaskGeoDataset = ReturnAnalysisMask(pWShedlayer, pMap, g_pFeatWorkspace, strWS)
    Else
        Set pMaskGeoDataset = g_pDEMRaster
    End If
    
    'STEP 3: With the Rasterdataset set, get its properties
    Set pRasterProps = g_pDEMRaster
    'Get cell size and envelope
    dblCellSize = pRasterProps.MeanCellSize.X
    'Set the global cell size
    g_dblCellSize = dblCellSize
    'Set pEnvelope = pRasterProps.Extent
    'Either pMaskGeoDataset will be the DEM or the mask
    Set pEnvelope = pMaskGeoDataset.Extent
    
    'Init pEnv
    'Set pEnv = New RasterAnalysis
    Set g_pSpatEnv = New RasterAnalysis
    
    'STEP 4: Now set the analysis environment
    With g_pSpatEnv
        .SetCellSize esriRasterEnvValue, dblCellSize                'Cell Size
        .SetExtent esriRasterEnvValue, pEnvelope, True              'Extent
        Set .Mask = pMaskGeoDataset                                 'Mask
        Set .OutSpatialReference = pRasterProps.SpatialReference    'Projection
        Set .OutWorkspace = g_pRasWorkspace                         'Outworkspace
        .SetAsNewDefaultEnvironment
    End With
    
    'Set g_pSpatEnv = pEnv
    
    'STEP 5: Set global units
    Select Case intDistUnits
        Case 0    'Meters
            g_intDistanceUnits = 0
        Case 1    'Feed
            g_intDistanceUnits = 1
    End Select
    
    'Flow Direction
    If modUtil.RasterExists(strFlowDir) Then
        Set g_pFlowDirRaster = modUtil.ReturnRaster(strFlowDir)
    Else
        strError = "Flow Direction Raster Does Not Exist: " & strFlowDir
    End If
    
    'FlowAccumulation
    If modUtil.RasterExists(strFlowAcc) Then
        Set g_pFlowAccRaster = modUtil.ReturnRaster(strFlowAcc)
    Else
        strError = "Flow Accumulation Raster Does Not Exist: " & strFlowAcc
    End If
    
    'Length Slope
    If modUtil.RasterExists(strLS) Then
        Set g_pLSRaster = modUtil.ReturnRaster(strLS)
        g_strLSFileName = strLS
    Else
        strError = "Length Slope raster does not Exist: " & strLS
    End If
    
    If Len(strError) > 0 Then
        MsgBox strError, vbCritical, "Missing Data"
    End If
    
    'Init the group layer
    Set g_pGroupLayer = New GroupLayer
    
    'Cleanup
    Set pEnv = Nothing
    Set pRasterProps = Nothing
    Set pEnvelope = Nothing
    
End Sub


Private Function ReturnAnalysisMask(pLayer As ILayer, pMap As IMap, pWorkspace As IWorkspace, strBasinFeatClass As String) As IFeatureClass
    'Incoming
    'pLayer: Layer user has chosen as being the one from which the selected polys will come
    'pMap: current map
    'pWorkspace:  place to put the exported selected sheds
    'strBasinFeatClass: string file location of BasinPoly.shp
      
    Dim pFLayer As IFeatureLayer
    Dim pFC As IFeatureClass
    Dim pINFeatureClassName As IFeatureClassName
    Dim pDataset As IDataset
    Dim pInDsName As IDatasetName
    Dim pFSel As IFeatureSelection
    Dim pSelSet As ISelectionSet
    Dim pFeatureClassName As IFeatureClassName
    Dim pOutDatasetName As IDatasetName
    Dim pWorkspaceName As IWorkspaceName
    Dim pExportOp As IExportOperation
    Dim pFeatWS As IFeatureWorkspace
    Dim pFeatClass As IFeatureClass
    
    Dim pSelectGeometry As IGeometry
    Dim pBasinFeatClass As IFeatureClass
    Dim pSpatialFilter As ISpatialFilter
    Dim pBasinSelSet As ISelectionSet
        
    'Set the layer to the incoming one and get featureclass
    Set pFLayer = pLayer
    
    'Get the selection set
    Set pFSel = pFLayer
    Set pSelSet = pFSel.SelectionSet
    
    'Make a call to the function below to return a unioned geometry of any and all selected features
    Set pSelectGeometry = ReturnSelectGeometry(pSelSet)
    Set g_pSelectedPolyClip = pSelectGeometry
    
    'Make a call to get the BasinPoly featureclass using the name sent over
    Set pBasinFeatClass = modUtil.ReturnFeatureClass(strBasinFeatClass)
    
    'Now we select all basinpolys that intersect the unioned geometry we got earlier
    Set pSpatialFilter = New SpatialFilter
    
    With pSpatialFilter
        Set .Geometry = pSelectGeometry
        .GeometryField = pBasinFeatClass.ShapeFieldName
        .SpatialRel = esriSpatialRelIntersects
    End With
    
    Set pBasinSelSet = pBasinFeatClass.Select(pSpatialFilter, esriSelectionTypeIDSet, _
                                              esriSelectionOptionNormal, pWorkspace)
    
    'Get the FcName from the featureclass
    Set pDataset = pBasinFeatClass
    
    'Set classname and dataset
    Set pINFeatureClassName = pDataset.FullName
    Set pInDsName = pINFeatureClassName

    'Create a new feature class name
    ' Define the output feature class name
    Set pFeatureClassName = New FeatureClassName
    Set pOutDatasetName = pFeatureClassName

    pOutDatasetName.Name = modUtil.GetUniqueFeatureClassName(pWorkspace, "selshed")

    Set pWorkspaceName = New WorkspaceName
    pWorkspaceName.PathName = pWorkspace.PathName
    pWorkspaceName.WorkspaceFactoryProgID = "esriDataSourcesFile.shapefileworkspacefactory.1"

    Set pOutDatasetName.WorkspaceName = pWorkspaceName
    
    With pFeatureClassName
        .FeatureType = esriFTSimple
        .ShapeType = esriGeometryPolygon
        .ShapeFieldName = "Shape"
    End With
    
    'Export
    Set pExportOp = New ExportOperation
    pExportOp.ExportFeatureClass pInDsName, Nothing, pBasinSelSet, Nothing, pOutDatasetName, 0

    Set pFeatWS = pWorkspace
    Set pFeatClass = pFeatWS.OpenFeatureClass(pOutDatasetName.Name)
    Set ReturnAnalysisMask = pFeatClass
    
    'Cleanup
    Set pFLayer = Nothing
    Set pFC = Nothing
    Set pINFeatureClassName = Nothing
    Set pDataset = Nothing
    Set pInDsName = Nothing
    Set pFSel = Nothing
    Set pSelSet = Nothing
    Set pFeatureClassName = Nothing
    Set pOutDatasetName = Nothing
    Set pWorkspaceName = Nothing
    Set pExportOp = Nothing
    Set pFeatWS = Nothing
    
End Function

Public Function CheckMultiPartPolygon(ByVal pPolygon As IPolygon4) As Boolean

        If pPolygon.ExteriorRingCount > 1 Then
            CheckMultiPartPolygon = True
        Else
            CheckMultiPartPolygon = False
        End If

    End Function


Public Function ReturnSelectGeometry(pInSelectionSet As ISelectionSet) As IGeometry
    
    Dim pFeature As IFeature
    Dim pTopoSimple As ITopologicalOperator         'Simplifier
    Dim pNewTopo As ITopologicalOperator            'Union
    Dim pCurrPolygon As IPolygon
        
    Dim pSelectionSet As ISelectionSet
    Set pSelectionSet = pInSelectionSet 'pFeatureSelection.SelectionSet
    
    Dim pFeatureCursor As IFeatureCursor
    pSelectionSet.Search Nothing, True, pFeatureCursor
        
    If pSelectionSet.Count = 0 Then
        Exit Function
    End If
    
    Set pTopoSimple = New Polygon
    Set pNewTopo = New Polygon
    
    Set pFeature = pFeatureCursor.NextFeature
    
    Do While Not pFeature Is Nothing
               
        Set pCurrPolygon = pFeature.Shape
        Set pTopoSimple = pCurrPolygon
        pTopoSimple.Simplify
        
        Set pNewTopo = pNewTopo.Union(pTopoSimple)
        Set pFeature = pFeatureCursor.NextFeature
      
    Loop
    
    Set ReturnSelectGeometry = pNewTopo
    
    'Cleanup
    Set pFeature = Nothing
    Set pTopoSimple = Nothing
    Set pCurrPolygon = Nothing
    Set pSelectionSet = Nothing
    
End Function




