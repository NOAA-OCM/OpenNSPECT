Attribute VB_Name = "modMgmtScen"
' *************************************************************************************
' *  Perot Systems Government Services
' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
' *  modMgmtScen
' *************************************************************************************
' *  Description: Code for handling the Management Scenarios
' *
' *
' *  Called By:
' *************************************************************************************
Option Explicit

Private m_strLCClass As String
Private m_pEnv As IRasterAnalysisEnvironment
Private m_pRasterProps As IRasterProps
Private m_pLandCoverRaster As IRaster
Private m_pWS As IWorkspace
Public g_booLCChange As Boolean

Public Sub MgmtScenSetup(clsMgmtScens As clsXMLMgmtScenItems, strLCClass As String, strLCFileName As String, _
                         strWorkspace As String)
'Main Sub for setting everything up
'clsMgmtScens: XML wrapper for the management scenarios created by the user
'strLCClass: Name of the LandCover being used, CCAP
'strLCFileName: filename of location of LandCover file

    Dim strExpress As String
    Dim strExpression As String
    Dim strOutLandCover As String
    Dim booLandScen As Boolean
    Dim pNewLandCoverRaster As IRaster
    
    Dim i As Integer
        
    Dim pMapAlgebraOp As IMapAlgebraOp
    Dim pArray As IArray
    
On Error GoTo ErrHandler:

    'init everything
    m_strLCClass = strLCClass
    Set pArray = New esriSystem.Array
    Set m_pEnv = New RasterAnalysis
        
    'Make sure the landcoverraster exists..it better if they get to this point, ED!
    If modUtil.RasterExists(strLCFileName) Then
        Set m_pLandCoverRaster = modUtil.ReturnRaster(strLCFileName)
    Else
        Exit Sub
    End If
    
    'Get the rasterprops of the landcover raster; we use it this time and this time only
    Set m_pRasterProps = m_pLandCoverRaster
    
    'Get the workspace
    Set m_pWS = modUtil.SetRasterWorkspace(strWorkspace)
    
    'Set the environment
    With m_pEnv
        .SetCellSize esriRasterEnvValue, m_pRasterProps.MeanCellSize.X
        .SetExtent esriRasterEnvValue, m_pRasterProps.Extent
        Set .OutSpatialReference = m_pRasterProps.SpatialReference
        Set .OutWorkspace = m_pWS
    End With
    
    'Going to now take each entry in the landuse scenarios, if they've choosen 'apply', go
    'to ReclassRaster to make a new raster of the Value indicated.  Then add the returned
    'raster to m_RasterArray.  We'll then make a map algebra statement out of that collection
    'of rasters to merge back into the original landcover raster...voila.
    If clsMgmtScens.Count > 0 Then
        For i = 1 To clsMgmtScens.Count
            If clsMgmtScens.Item(i).intApply = 1 Then
                modProgDialog.ProgDialog "Adding new landclass...", "Creating Management Scenario", 0, CLng(clsMgmtScens.Count), CLng(i), 0
                If modProgDialog.g_boolCancel Then
                    pArray.Add ReclassRaster(clsMgmtScens.Item(i), m_strLCClass)
                    booLandScen = True
                End If
            End If
        Next i
    End If
    
    If Not booLandScen Then
        Set g_LandCoverRaster = m_pLandCoverRaster
        Exit Sub
    End If
    
    
    Set pMapAlgebraOp = New RasterMapAlgebraOp
    Set m_pEnv = pMapAlgebraOp
    
    With m_pEnv
        .SetCellSize esriRasterEnvValue, m_pRasterProps.MeanCellSize.X
        .SetExtent esriRasterEnvValue, m_pRasterProps.Extent
        Set .OutSpatialReference = m_pRasterProps.SpatialReference
        Set .OutWorkspace = m_pWS
    End With
    'Now, if the raster array has some rasters in it, let's merge the fuckers together
    If pArray.Count > 0 Then
        For i = 0 To pArray.Count - 1
            pMapAlgebraOp.BindRaster pArray.element(i), "raster" & CStr(i)
            If strExpress = "" Then
                strExpress = "Merge([" & "raster" & CStr(i) & "], "
            Else
                strExpress = strExpress & "[" & "raster" & CStr(i) & "], "
            End If
        Next i
        
        pMapAlgebraOp.BindRaster m_pLandCoverRaster, "landcover"
        
        strExpression = strExpress & "[landcover])"
        
        modProgDialog.ProgDialog "Adding new landclass...", "Creating Management Scenario", 0, CLng(clsMgmtScens.Count + 1), CLng(clsMgmtScens.Count + 1), 0
                If modProgDialog.g_boolCancel Then
                Set pNewLandCoverRaster = pMapAlgebraOp.Execute(strExpression)
                End If
    Else
        Exit Sub
    End If
    
    strOutLandCover = modUtil.GetUniqueName("landcover", modUtil.SplitWorkspaceName(m_pEnv.OutWorkspace.PathName))
    
    If modUtil.MakePermanentRaster(pNewLandCoverRaster, m_pEnv.OutWorkspace.PathName, strOutLandCover) Then
        Set g_LandCoverRaster = pNewLandCoverRaster
    End If
    
    modProgDialog.KillDialog
    
    'Cleanup
    Set pMapAlgebraOp = Nothing
    Set m_pWS = Nothing
    
Exit Sub
ErrHandler:
    MsgBox "error in MSSetup " & Err.Number & ": " & Err.Description
    modProgDialog.KillDialog
            
End Sub

Public Function ReclassRaster(clsMgmtScen As clsXMLMgmtScenItem, strLCClass As String) As IRaster
'We're passing over a single management scenarios in the form of the xml
'class clsXMLmgmtScenItem, seems to be the easiest way to do this.
    Dim strSelect As String                         'ADO selections string
    Dim rsValue As ADODB.Recordset                  'ADO recordset
    Dim dblValue As Double                          'value
    Dim strOutName As String                        'String for poly's name
    Dim strExpression As String                     'Map calculator expression
    
    Dim pConversionOp As IConversionOp              'Conversion for poly to raster
    Dim pPolyFeatureClass As IFeatureClass          'The polygon featureclass
    Dim pPolyDS As IGeoDataset
    Dim pOutPolyRDS As IRasterDataset
    Dim pMapAlgebraOp As IMapAlgebraOp
    Dim pPolyValueRaster As IRaster
    Dim pFinalPolyValueRaster As IRaster

On Error GoTo ErrHandler
    
    'STEP 1: Open the landclass Value Value -------------------------------------------------------------------------
    'This is the value user's landclass will change to
    strSelect = "SELECT LCTYPE.LCTYPEID, LCCLASS.NAME, LCCLASS.VALUE FROM " & _
        "LCTYPE INNER JOIN LCCLASS ON LCTYPE.LCTYPEID = LCCLASS.LCTYPEID " & _
        "WHERE LCTYPE.NAME LIKE '" & strLCClass & "' AND LCCLASS.NAME LIKE '" & clsMgmtScen.strChangeToClass & "'"
    
    Set rsValue = New ADODB.Recordset
    rsValue.Open strSelect, g_ADOConn, adOpenKeyset
        
    dblValue = rsValue!Value
    
    rsValue.Close
    'END STEP 1: ----------------------------------------------------------------------------------------------------
    
    'STEP 2: --------------------------------------------------------------------------------------------------------
    'Init
    Set pConversionOp = New RasterConversionOp
    'Set the spat Environment
    Set m_pEnv = pConversionOp
    
    With m_pEnv
        .SetCellSize esriRasterEnvValue, m_pRasterProps.MeanCellSize.X
        .SetExtent esriRasterEnvValue, m_pRasterProps.Extent
        Set .OutSpatialReference = m_pRasterProps.SpatialReference
        Set .OutWorkspace = m_pWS
    End With
    'END STEP 2: ----------------------------------------------------------------------------------------------------
    
    'STEP 3: --------------------------------------------------------------------------------------------------------
    'Convert the polygon featureclass into a Raster for sending back out
    'Get the featureclass
    Set pPolyFeatureClass = modUtil.ReturnFeatureClass(clsMgmtScen.strAreaFileName)
    'Have to use a GeoDataset, so QI
    Set pPolyDS = pPolyFeatureClass
    'Init the out dataset
    Set pOutPolyRDS = New RasterDataset
    strOutName = modUtil.GetUniqueName("poly", m_pWS.PathName)
    Set pOutPolyRDS = pConversionOp.ToRasterDataset(pPolyDS, "GRID", m_pWS, strOutName)
    'END STEP 3 ------------------------------------------------------------------------------------------------------
    
    'STEP 4: ---------------------------------------------------------------------------------------------------------
    'Need to now make that raster the value the user choose way back when
    Set pMapAlgebraOp = New RasterMapAlgebraOp
    Set m_pEnv = pMapAlgebraOp
    
    With m_pEnv
        .SetCellSize esriRasterEnvValue, m_pRasterProps.MeanCellSize.X
        .SetExtent esriRasterEnvValue, m_pRasterProps.Extent
        Set .OutSpatialReference = m_pRasterProps.SpatialReference
        Set .OutWorkspace = m_pWS
    End With
    
    Set pPolyValueRaster = pOutPolyRDS.CreateDefaultRaster
        
    pMapAlgebraOp.BindRaster pPolyValueRaster, "poly"
       
    strExpression = "Con([poly] <> -9999, " & dblValue & ", 0)"
    
    Set pFinalPolyValueRaster = pMapAlgebraOp.Execute(strExpression)
    'End STEP 4: -----------------------------------------------------------------------------------------------------
    
    Set ReclassRaster = pFinalPolyValueRaster
    
'Cleanup
    Set rsValue = Nothing
    Set pConversionOp = Nothing
    Set pPolyFeatureClass = Nothing
    Set pPolyDS = Nothing
    Set pOutPolyRDS = Nothing
    Set pMapAlgebraOp = Nothing
    Set pPolyValueRaster = Nothing
    Set pFinalPolyValueRaster = Nothing
    
    
Exit Function
ErrHandler:
    MsgBox Err.Number & ": " & Err.Description
    
    
End Function





