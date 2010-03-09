Module modMgmtScen
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

    Private m_strLCClass As String
    'Private m_pEnv As ESRI.ArcGIS.GeoAnalyst.IRasterAnalysisEnvironment
    'Private m_pRasterProps As ESRI.ArcGIS.DataSourcesRaster.IRasterProps
    'Private m_pLandCoverRaster As ESRI.ArcGIS.Geodatabase.IRaster
    'Private m_pWS As ESRI.ArcGIS.Geodatabase.IWorkspace
    Public g_booLCChange As Boolean

    Public Sub MgmtScenSetup(ByRef clsMgmtScens As clsXMLMgmtScenItems, ByRef strLCClass As String, ByRef strLCFileName As String, ByRef strWorkspace As String)
        '        'Main Sub for setting everything up
        '        'clsMgmtScens: XML wrapper for the management scenarios created by the user
        '        'strLCClass: Name of the LandCover being used, CCAP
        '        'strLCFileName: filename of location of LandCover file

        '        Dim strExpress As String
        '        Dim strExpression As String
        '        Dim strOutLandCover As String
        '        Dim booLandScen As Boolean
        '        Dim pNewLandCoverRaster As ESRI.ArcGIS.Geodatabase.IRaster

        '        Dim i As Short

        '        Dim pMapAlgebraOp As ESRI.ArcGIS.SpatialAnalyst.IMapAlgebraOp
        '        Dim pArray As ESRI.ArcGIS.esriSystem.IArray

        '        On Error GoTo ErrHandler

        '        'init everything
        '        m_strLCClass = strLCClass
        '        pArray = New ESRI.ArcGIS.esriSystem.Array
        '        m_pEnv = New ESRI.ArcGIS.GeoAnalyst.RasterAnalysis

        '        'Make sure the landcoverraster exists..it better if they get to this point, ED!
        '        If modUtil.RasterExists(strLCFileName) Then
        '            m_pLandCoverRaster = modUtil.ReturnRaster(strLCFileName)
        '        Else
        '            Exit Sub
        '        End If

        '        'Get the rasterprops of the landcover raster; we use it this time and this time only
        '        m_pRasterProps = m_pLandCoverRaster

        '        'Get the workspace
        '        m_pWS = modUtil.SetRasterWorkspace(strWorkspace)

        '        'Set the environment
        '        With m_pEnv
        '            .SetCellSize(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvValue, m_pRasterProps.MeanCellSize.X)
        '            .SetExtent(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvValue, m_pRasterProps.Extent)
        '            'UPGRADE_WARNING: Couldn't resolve default property of object m_pEnv.OutSpatialReference. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            .OutSpatialReference = m_pRasterProps.SpatialReference
        '            'UPGRADE_WARNING: Couldn't resolve default property of object m_pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            .OutWorkspace = m_pWS
        '        End With

        '        'Going to now take each entry in the landuse scenarios, if they've choosen 'apply', go
        '        'to ReclassRaster to make a new raster of the Value indicated.  Then add the returned
        '        'raster to m_RasterArray.  We'll then make a map algebra statement out of that collection
        '        'of rasters to merge back into the original landcover raster...voila.
        '        If clsMgmtScens.Count > 0 Then
        '            For i = 1 To clsMgmtScens.Count
        '                If clsMgmtScens.Item(i).intApply = 1 Then
        '                    modProgDialog.ProgDialog("Adding new landclass...", "Creating Management Scenario", 0, CInt(clsMgmtScens.Count), CInt(i), 0)
        '                    If modProgDialog.g_boolCancel Then
        '                        pArray.Add(ReclassRaster(clsMgmtScens.Item(i), m_strLCClass))
        '                        booLandScen = True
        '                    End If
        '                End If
        '            Next i
        '        End If

        '        If Not booLandScen Then
        '            g_LandCoverRaster = m_pLandCoverRaster
        '            Exit Sub
        '        End If


        '        pMapAlgebraOp = New ESRI.ArcGIS.SpatialAnalyst.RasterMapAlgebraOp
        '        m_pEnv = pMapAlgebraOp

        '        With m_pEnv
        '            .SetCellSize(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvValue, m_pRasterProps.MeanCellSize.X)
        '            .SetExtent(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvValue, m_pRasterProps.Extent)
        '            'UPGRADE_WARNING: Couldn't resolve default property of object m_pEnv.OutSpatialReference. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            .OutSpatialReference = m_pRasterProps.SpatialReference
        '            'UPGRADE_WARNING: Couldn't resolve default property of object m_pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            .OutWorkspace = m_pWS
        '        End With
        '        'Now, if the raster array has some rasters in it, let's merge the fuckers together
        '        If pArray.Count > 0 Then
        '            For i = 0 To pArray.Count - 1
        '                'UPGRADE_WARNING: Couldn't resolve default property of object pArray.element(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                pMapAlgebraOp.BindRaster(pArray.element(i), "raster" & CStr(i))
        '                If strExpress = "" Then
        '                    strExpress = "Merge([" & "raster" & CStr(i) & "], "
        '                Else
        '                    strExpress = strExpress & "[" & "raster" & CStr(i) & "], "
        '                End If
        '            Next i

        '            'UPGRADE_WARNING: Couldn't resolve default property of object m_pLandCoverRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            pMapAlgebraOp.BindRaster(m_pLandCoverRaster, "landcover")

        '            strExpression = strExpress & "[landcover])"

        '            modProgDialog.ProgDialog("Adding new landclass...", "Creating Management Scenario", 0, CInt(clsMgmtScens.Count + 1), CInt(clsMgmtScens.Count + 1), 0)
        '            If modProgDialog.g_boolCancel Then
        '                pNewLandCoverRaster = pMapAlgebraOp.Execute(strExpression)
        '            End If
        '        Else
        '            Exit Sub
        '        End If

        '        'UPGRADE_WARNING: Couldn't resolve default property of object m_pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        strOutLandCover = modUtil.GetUniqueName("landcover", modUtil.SplitWorkspaceName(m_pEnv.OutWorkspace.PathName))

        '        'UPGRADE_WARNING: Couldn't resolve default property of object m_pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        If modUtil.MakePermanentRaster(pNewLandCoverRaster, m_pEnv.OutWorkspace.PathName, strOutLandCover) Then
        '            g_LandCoverRaster = pNewLandCoverRaster
        '        End If

        '        modProgDialog.KillDialog()

        '        'Cleanup
        '        'UPGRADE_NOTE: Object pMapAlgebraOp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pMapAlgebraOp = Nothing
        '        'UPGRADE_NOTE: Object m_pWS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        m_pWS = Nothing

        '        Exit Sub
        'ErrHandler:
        '        MsgBox("error in MSSetup " & Err.Number & ": " & Err.Description)
        '        modProgDialog.KillDialog()

    End Sub

    Public Function ReclassRaster(ByRef clsMgmtScen As clsXMLMgmtScenItem, ByRef strLCClass As String) As MapWinGIS.Grid
        '        'We're passing over a single management scenarios in the form of the xml
        '        'class clsXMLmgmtScenItem, seems to be the easiest way to do this.
        '        Dim strSelect As String 'ADO selections string
        '        Dim rsValue As ADODB.Recordset 'ADO recordset
        '        Dim dblValue As Double 'value
        '        Dim strOutName As String 'String for poly's name
        '        Dim strExpression As String 'Map calculator expression

        '        Dim pConversionOp As ESRI.ArcGIS.GeoAnalyst.IConversionOp 'Conversion for poly to raster
        '        Dim pPolyFeatureClass As ESRI.ArcGIS.Geodatabase.IFeatureClass 'The polygon featureclass
        '        Dim pPolyDS As ESRI.ArcGIS.Geodatabase.IGeoDataset
        '        Dim pOutPolyRDS As ESRI.ArcGIS.Geodatabase.IRasterDataset
        '        Dim pMapAlgebraOp As ESRI.ArcGIS.SpatialAnalyst.IMapAlgebraOp
        '        Dim pPolyValueRaster As ESRI.ArcGIS.Geodatabase.IRaster
        '        Dim pFinalPolyValueRaster As ESRI.ArcGIS.Geodatabase.IRaster

        '        On Error GoTo ErrHandler

        '        'STEP 1: Open the landclass Value Value -------------------------------------------------------------------------
        '        'This is the value user's landclass will change to
        '        strSelect = "SELECT LCTYPE.LCTYPEID, LCCLASS.NAME, LCCLASS.VALUE FROM " & "LCTYPE INNER JOIN LCCLASS ON LCTYPE.LCTYPEID = LCCLASS.LCTYPEID " & "WHERE LCTYPE.NAME LIKE '" & strLCClass & "' AND LCCLASS.NAME LIKE '" & clsMgmtScen.strChangeToClass & "'"

        '        rsValue = New ADODB.Recordset
        '        rsValue.Open(strSelect, g_ADOConn, ADODB.CursorTypeEnum.adOpenKeyset)

        '        dblValue = rsValue.Fields("Value").Value

        '        rsValue.Close()
        '        'END STEP 1: ----------------------------------------------------------------------------------------------------

        '        'STEP 2: --------------------------------------------------------------------------------------------------------
        '        'Init
        '        pConversionOp = New ESRI.ArcGIS.GeoAnalyst.RasterConversionOp
        '        'Set the spat Environment
        '        m_pEnv = pConversionOp

        '        With m_pEnv
        '            .SetCellSize(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvValue, m_pRasterProps.MeanCellSize.X)
        '            .SetExtent(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvValue, m_pRasterProps.Extent)
        '            'UPGRADE_WARNING: Couldn't resolve default property of object m_pEnv.OutSpatialReference. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            .OutSpatialReference = m_pRasterProps.SpatialReference
        '            'UPGRADE_WARNING: Couldn't resolve default property of object m_pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            .OutWorkspace = m_pWS
        '        End With
        '        'END STEP 2: ----------------------------------------------------------------------------------------------------

        '        'STEP 3: --------------------------------------------------------------------------------------------------------
        '        'Convert the polygon featureclass into a Raster for sending back out
        '        'Get the featureclass
        '        pPolyFeatureClass = modUtil.ReturnFeatureClass((clsMgmtScen.strAreaFileName))
        '        'Have to use a GeoDataset, so QI
        '        pPolyDS = pPolyFeatureClass
        '        'Init the out dataset
        '        pOutPolyRDS = New ESRI.ArcGIS.DataSourcesRaster.RasterDataset
        '        strOutName = modUtil.GetUniqueName("poly", (m_pWS.PathName))
        '        pOutPolyRDS = pConversionOp.ToRasterDataset(pPolyDS, "GRID", m_pWS, strOutName)
        '        'END STEP 3 ------------------------------------------------------------------------------------------------------

        '        'STEP 4: ---------------------------------------------------------------------------------------------------------
        '        'Need to now make that raster the value the user choose way back when
        '        pMapAlgebraOp = New ESRI.ArcGIS.SpatialAnalyst.RasterMapAlgebraOp
        '        m_pEnv = pMapAlgebraOp

        '        With m_pEnv
        '            .SetCellSize(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvValue, m_pRasterProps.MeanCellSize.X)
        '            .SetExtent(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvValue, m_pRasterProps.Extent)
        '            'UPGRADE_WARNING: Couldn't resolve default property of object m_pEnv.OutSpatialReference. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            .OutSpatialReference = m_pRasterProps.SpatialReference
        '            'UPGRADE_WARNING: Couldn't resolve default property of object m_pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            .OutWorkspace = m_pWS
        '        End With

        '        pPolyValueRaster = pOutPolyRDS.CreateDefaultRaster

        '        'UPGRADE_WARNING: Couldn't resolve default property of object pPolyValueRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        pMapAlgebraOp.BindRaster(pPolyValueRaster, "poly")

        '        strExpression = "Con([poly] <> -9999, " & dblValue & ", 0)"

        '        pFinalPolyValueRaster = pMapAlgebraOp.Execute(strExpression)
        '        'End STEP 4: -----------------------------------------------------------------------------------------------------

        '        ReclassRaster = pFinalPolyValueRaster

        '        'Cleanup
        '        'UPGRADE_NOTE: Object rsValue may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        rsValue = Nothing
        '        'UPGRADE_NOTE: Object pConversionOp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pConversionOp = Nothing
        '        'UPGRADE_NOTE: Object pPolyFeatureClass may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pPolyFeatureClass = Nothing
        '        'UPGRADE_NOTE: Object pPolyDS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pPolyDS = Nothing
        '        'UPGRADE_NOTE: Object pOutPolyRDS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pOutPolyRDS = Nothing
        '        'UPGRADE_NOTE: Object pMapAlgebraOp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pMapAlgebraOp = Nothing
        '        'UPGRADE_NOTE: Object pPolyValueRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pPolyValueRaster = Nothing
        '        'UPGRADE_NOTE: Object pFinalPolyValueRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pFinalPolyValueRaster = Nothing


        '        Exit Function
        'ErrHandler:
        '        MsgBox(Err.Number & ": " & Err.Description)


    End Function
End Module