Imports System.Data.OleDb
Module modRunoff
    ' *************************************************************************************
    ' *  Perot Systems Government Services
    ' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
    ' *  modRunoff
    ' *************************************************************************************
    ' *  Description:  Code for all runoff goodies
    ' *
    ' *
    ' *  Called By:  Various
    ' *************************************************************************************

    'Public g_pSoilsRaster As ESRI.ArcGIS.Geodatabase.IRasterDataset 'Soils Raster GRID
    'Public g_SoilsRasterDataset As ESRI.ArcGIS.Geodatabase.IRasterDataset
    'Public g_LandCoverRaster As ESRI.ArcGIS.Geodatabase.IRaster 'LandCover GLOBAL
    'Public g_pRunoffRaster As ESRI.ArcGIS.Geodatabase.IRaster 'Runoff for use elsewhere
    'Public g_pMetRunoffRaster As ESRI.ArcGIS.Geodatabase.IRaster 'Metric Runoff
    'Public g_pSCS100Raster As ESRI.ArcGIS.Geodatabase.IRaster 'SCS GRID for use in RUSLE
    'Public g_pAbstractRaster As ESRI.ArcGIS.Geodatabase.IRaster 'Abstraction Raster used in MUSLE
    'Public g_pPrecipRaster As ESRI.ArcGIS.Geodatabase.IRaster 'Precipitation Raster
    'Public g_pCellAreaSqMiRaster As ESRI.ArcGIS.Geodatabase.IRaster 'Square Mile area raster used in MUSLE
    'Public g_pRunoffInchRaster As ESRI.ArcGIS.Geodatabase.IRaster 'Runoff Inches used in MUSLE
    'Public g_pRunoffAFRaster As ESRI.ArcGIS.Geodatabase.IRaster 'Runoff AF used in MUSLE
    'Public g_pRunoffCFRaster As ESRI.ArcGIS.Geodatabase.IRaster 'MUSLE
    Public g_strPrecipFileName As String 'Precipitation File Name
    Public g_strLandCoverParameters As String 'Global string to hold formatted LC params for metadata

    Private m_strRunoffMetadata As String 'Metatdata string for runoff
    Private m_intPrecipType As Short 'Precip Event Type: 0=Annual; 1=Event
    Private m_intRainingDays As Short '# of Raining Days for Annual Precip
    ' Variables used by the Error handler function - DO NOT REMOVE
    Const c_sModuleFileName As String = "modRunoff.bas"
    Private m_ParentHWND As Integer ' Set this to get correct parenting of Error handler forms



    Private Function ConstructPickStatment(ByRef strLandClass As String, ByRef pLCRaster As MapWinGIS.Grid) As String
        '        'Creates the large initial pick statement using the name of the the LandCass [CCAP, for example]
        '        'and the Land Class Raster.  Returns a string

        '        Dim strRS As String
        '        Dim rsLandClass As ADODB.Recordset
        '        Dim strPick(3) As String 'Array of strings that hold 'pick' numbers
        '        Dim strCurveCalc As String 'Full String

        '        Dim pRasterCol As ESRI.ArcGIS.DataSourcesRaster.IRasterBandCollection
        '        Dim pBand As ESRI.ArcGIS.DataSourcesRaster.IRasterBand
        '        Dim pRasterStats As ESRI.ArcGIS.DataSourcesRaster.IRasterStatistics
        '        Dim dblMaxValue As Double
        '        Dim i As Short
        '        Dim pTable As ESRI.ArcGIS.Geodatabase.ITable
        '        Dim TableExist As Boolean
        '        Dim pCursor As ESRI.ArcGIS.Geodatabase.ICursor
        '        Dim pRow As ESRI.ArcGIS.Geodatabase.IRow
        '        Dim FieldIndex As Short
        '        Dim booValueFound As Boolean
        '        Dim j As Short

        '        On Error GoTo ErrHandler

        '        'STEP 1:  get the records from the database -----------------------------------------------
        '        strRS = "SELECT LCCLASS.LCClassID, Value, LCCLASS.Name as Name2, LCCLASS.LCTypeID, [CN-A], [CN-B], [CN-C], [CN-D], CoverFactor, W_WL FROM LCCLASS " & "LEFT OUTER JOIN LCTYPE ON LCCLASS.LCTYPEID = LCTYPE.LCTYPEID " & "Where LCTYPE.NAME = '" & strLandClass & "' ORDER BY LCCLASS.VALUE"

        '        rsLandClass = New ADODB.Recordset
        '        rsLandClass.Open(strRS, g_ADOConn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)


        '        'while here, make metadata
        '        m_strRunoffMetadata = CreateMetadata(rsLandClass, strLandClass, g_booLocalEffects)
        '        'End Database stuff

        '        'STEP 2: Raster Values ---------------------------------------------------------------------
        '        'Now Get the RASTER values
        '        ' Get Rasterband from the incoming raster
        '        pRasterCol = pLCRaster
        '        pBand = pRasterCol.Item(0)

        '        'Get the max value
        '        pRasterStats = pBand.Statistics
        '        dblMaxValue = pRasterStats.Maximum

        '        'Get the raster table
        '        pBand.HasTable(TableExist)
        '        If Not TableExist Then Exit Function

        '        pTable = pBand.AttributeTable
        '        'Get All rows
        '        pCursor = pTable.Search(Nothing, True)
        '        'Init pRow
        '        pRow = pCursor.NextRow

        '        'Get index of Value Field
        '        'UPGRADE_WARNING: Couldn't resolve default property of object pTable.FindField. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        FieldIndex = pTable.FindField("Value")

        '        'STEP 3: Table values vs. Raster Values Count - if not equal bark -------------------------
        '        'If pTable.RowCount(Nothing) <> rsLandClass.RecordCount Then
        '        '   MsgBox "Error: The number of records in your Land Class Table do not match your GRID dataset."
        '        '   Exit Function
        '        'End If
        '        rsLandClass.MoveFirst()

        '        'STEP 4: Create the strings
        '        'Loop through and get all values
        '        For i = 1 To dblMaxValue
        '            'Check first value, usually they won't have a [1], but if there check against database,
        '            'else move on
        '            If i = 1 Then
        '                'UPGRADE_WARNING: Couldn't resolve default property of object pRow.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                If (pRow.Value(FieldIndex) = i) Then 'And (pRow.Value(FieldIndex) = rsLandClass!Value) Then
        '                    rsLandClass.MoveFirst()
        '                    booValueFound = False
        '                    For j = 0 To rsLandClass.RecordCount - 0
        '                        'UPGRADE_WARNING: Couldn't resolve default property of object pRow.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                        If pRow.Value(FieldIndex) = rsLandClass.Fields("Value").Value Then
        '                            booValueFound = True
        '                            strPick(0) = CStr(rsLandClass.Fields("CN-A").Value)
        '                            strPick(1) = CStr(rsLandClass.Fields("CN-B").Value)
        '                            strPick(2) = CStr(rsLandClass.Fields("CN-C").Value)
        '                            strPick(3) = CStr(rsLandClass.Fields("CN-D").Value)
        '                            'rsLandClass.MoveNext
        '                            pRow = pCursor.NextRow
        '                            Exit For
        '                        Else
        '                            booValueFound = False
        '                            rsLandClass.MoveNext()
        '                        End If
        '                    Next j
        '                    If booValueFound = False Then
        '                        MsgBox("Error: Your N-SPECT Land Class Table is missing values found in your landcover GRID dataset.")
        '                        ConstructPickStatment = ""
        '                        Exit Function
        '                    End If
        '                Else
        '                    strPick(0) = "-9999"
        '                    strPick(1) = "-9999"
        '                    strPick(2) = "-9999"
        '                    strPick(3) = "-9999"
        '                End If
        '            Else
        '                'UPGRADE_WARNING: Couldn't resolve default property of object pRow.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                If (pRow.Value(FieldIndex) = i) Then 'And (pRow.Value(FieldIndex) = rsLandClass!Value) Then
        '                    rsLandClass.MoveFirst()
        '                    booValueFound = False
        '                    For j = 0 To rsLandClass.RecordCount - 0
        '                        'UPGRADE_WARNING: Couldn't resolve default property of object pRow.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                        If pRow.Value(FieldIndex) = rsLandClass.Fields("Value").Value Then
        '                            booValueFound = True
        '                            strPick(0) = strPick(0) & ", " & CStr(rsLandClass.Fields("CN-A").Value)
        '                            strPick(1) = strPick(1) & ", " & CStr(rsLandClass.Fields("CN-B").Value)
        '                            strPick(2) = strPick(2) & ", " & CStr(rsLandClass.Fields("CN-C").Value)
        '                            strPick(3) = strPick(3) & ", " & CStr(rsLandClass.Fields("CN-D").Value)
        '                            'rsLandClass.MoveNext
        '                            pRow = pCursor.NextRow
        '                            Exit For
        '                        Else
        '                            booValueFound = False
        '                            rsLandClass.MoveNext()
        '                        End If
        '                    Next j
        '                    If booValueFound = False Then
        '                        MsgBox("Error: Your N-SPECT Land Class Table is missing values found in your landcover GRID dataset.")
        '                        ConstructPickStatment = ""
        '                        Exit Function
        '                    End If
        '                Else
        '                    strPick(0) = strPick(0) & ", 0"
        '                    strPick(1) = strPick(1) & ", 0"
        '                    strPick(2) = strPick(2) & ", 0"
        '                    strPick(3) = strPick(3) & ", 0"
        '                End If
        '            End If
        '        Next i

        '        strCurveCalc = "con([rev_soils] == 1, pick([ccap], " & strPick(0) & "), con([rev_soils] == 2, pick([ccap], " & strPick(1) & "), con([rev_soils] == 3, pick([ccap], " & strPick(2) & "), con([rev_soils] == 4, pick([ccap], " & strPick(3) & ")))))"

        '        ConstructPickStatment = strCurveCalc



        '        'Cleanup:
        '        'UPGRADE_NOTE: Object pRasterCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pRasterCol = Nothing
        '        'UPGRADE_NOTE: Object pRasterStats may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pRasterStats = Nothing
        '        'UPGRADE_NOTE: Object pBand may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pBand = Nothing
        '        'UPGRADE_NOTE: Object pTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pTable = Nothing
        '        'UPGRADE_NOTE: Object pCursor may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pCursor = Nothing
        '        'UPGRADE_NOTE: Object pRow may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pRow = Nothing

        '        Exit Function
        'ErrHandler:
        '        MsgBox(Err.Number & ": " & Err.Description & " " & "ConstructPickStatemnt")


    End Function

    Public Function CreateRunoffGrid(ByRef strLCFileName As String, ByRef strLCCLassType As String, ByRef cmdPrecip As oledbcommand, ByRef strSoilsFileName As String) As Boolean
        '        'This sub serves as a link between frmPrj and the actual calculation of Runoff
        '        'It establishes the Rasters being used

        '        'strLCFileName: Path to land class file
        '        'strLCCLassType: type of landclass
        '        'rsPrecip: ADO recordset of precip scenario being used
        '        Dim pRainFallRaster As ESRI.ArcGIS.Geodatabase.IRaster
        '        Dim pLandCoverRaster As ESRI.ArcGIS.Geodatabase.IRaster
        '        Dim pSoilsRaster As ESRI.ArcGIS.Geodatabase.IRaster
        '        Dim strPickStatement As String
        '        Dim strError As Object

        '        'Get the LandCover Raster
        '        'if no management scenarios were applied then g_landcoverRaster will be nothing
        '        If g_LandCoverRaster Is Nothing Then
        '            If modUtil.RasterExists(strLCFileName) Then
        '                g_LandCoverRaster = modUtil.ReturnRaster(strLCFileName)
        '                pLandCoverRaster = g_LandCoverRaster
        '            Else
        '                'UPGRADE_WARNING: Couldn't resolve default property of object strError. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                strError = strLCFileName
        '                GoTo Missing
        '            End If
        '        Else
        '            pLandCoverRaster = g_LandCoverRaster
        '        End If

        '        'Get the Precip Raster
        '        If modUtil.RasterExists(rsPrecip.Fields("PrecipFileName").Value) Then
        '            pRainFallRaster = modUtil.ReturnRaster(rsPrecip.Fields("PrecipFileName").Value)

        '            'If in cm, then convert to a GRID in inches.
        '            If rsPrecip.Fields("PrecipUnits").Value = 0 Then
        '                g_pPrecipRaster = ConvertRainGridCMToInches(pRainFallRaster)
        '            Else 'if already in inches then just use this one.
        '                g_pPrecipRaster = pRainFallRaster 'Global Precip
        '            End If

        '            g_strPrecipFileName = rsPrecip.Fields("PrecipFileName").Value
        '            m_intPrecipType = rsPrecip.Fields("Type").Value 'Set the mod level precip type
        '            m_intRainingDays = rsPrecip.Fields("RainingDays").Value 'Set the mod level rainingdays
        '        Else
        '            'UPGRADE_WARNING: Couldn't resolve default property of object strError. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            strError = strLCFileName
        '            GoTo Missing
        '        End If

        '        '------ 1.1.4 Change -----
        '        'if they select cm as incoming precip units, convert GRID

        '        'Get the Soils Raster
        '        If modUtil.RasterExists(strSoilsFileName) Then
        '            pSoilsRaster = modUtil.ReturnRaster(strSoilsFileName)
        '        Else
        '            'UPGRADE_WARNING: Couldn't resolve default property of object strError. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            strError = strSoilsFileName
        '            GoTo Missing
        '        End If

        '        'Now construct the Pick Statement
        '        strPickStatement = ConstructPickStatment(strLCCLassType, pLandCoverRaster)

        '        If Len(strPickStatement) = 0 Then
        '            Exit Function
        '        End If

        '        'Call the Runoff Calculation using the string and rasters
        '        If RunoffCalculation(strPickStatement, g_pPrecipRaster, pLandCoverRaster, pSoilsRaster) Then
        '            CreateRunoffGrid = True
        '        Else
        '            CreateRunoffGrid = False
        '            Exit Function
        '        End If

        '        CreateRunoffGrid = True

        '        Exit Function
        'Missing:
        '        'UPGRADE_WARNING: Couldn't resolve default property of object strError. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        MsgBox("Error: The following dataset is missing: " & strError, MsgBoxStyle.Critical, "Missing Data")
        '        CreateRunoffGrid = False
        '        Exit Function

    End Function

    Private Function ConvertRainGridCMToInches(ByRef pInRaster As MapWinGIS.Grid) As MapWinGIS.Grid

        'Dim pMapAlgebraOp As ESRI.ArcGIS.SpatialAnalyst.IMapAlgebraOp
        'Dim pEnv As ESRI.ArcGIS.GeoAnalyst.IRasterAnalysisEnvironment
        'Dim strExpression As String

        'pEnv = g_pSpatEnv
        'pMapAlgebraOp = New ESRI.ArcGIS.SpatialAnalyst.RasterMapAlgebraOp
        'pEnv = pMapAlgebraOp

        ''UPGRADE_WARNING: Couldn't resolve default property of object pInRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'pMapAlgebraOp.BindRaster(pInRaster, "precip")
        'strExpression = "[precip] / 2.54"

        'ConvertRainGridCMToInches = pMapAlgebraOp.Execute(strExpression)

        'pMapAlgebraOp.UnbindRaster("precip")


    End Function


    Public Function RunoffCalculation(ByRef strPickStatement As String, ByRef pInRainRaster As MapWinGIS.Grid, ByRef pInLandCoverRaster As MapWinGIS.Grid, ByRef pInSoilsRaster As MapWinGIS.Grid) As Boolean
        '        'strPickStatement: our friend the dynamic pick statemnt
        '        'pInRainRaster: the precip grid
        '        'pInLandCoverRaster: landcover grid
        '        On Error GoTo ErrHandler
        '        'The rasters                             'Associated Steps
        '        Dim pRasterProps As ESRI.ArcGIS.DataSourcesRaster.IRasterProps
        '        Dim pLandCoverRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 1: LandCover Raster
        '        Dim pLandSampleRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 1: Properly sized landcover
        '        Dim pHydSoilsRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 2: Soils Raster
        '        Dim pPrecipRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 3: Create Precip Grid
        '        Dim pSCSRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 1: Create Curve Number GRID
        '        Dim pDailyRainRaster As ESRI.ArcGIS.Geodatabase.IRaster 'Rain
        '        Dim pSCS100Raster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 2: SCS * 100
        '        Dim pRetentionRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 3: Retention
        '        Dim pAbstractRaster As ESRI.ArcGIS.Geodatabase.IRaster
        '        Dim pTemp1RunoffRaster As ESRI.ArcGIS.Geodatabase.IRaster
        '        Dim pTemp2RunoffRaster As ESRI.ArcGIS.Geodatabase.IRaster
        '        Dim pRunoffInRaster As ESRI.ArcGIS.Geodatabase.IRaster
        '        Dim pMaskRaster As ESRI.ArcGIS.Geodatabase.IRaster
        '        Dim pMetRunoffRaster As ESRI.ArcGIS.Geodatabase.IRaster
        '        Dim pMetRunoffNoNullRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 6a:  no nulls
        '        Dim pAccumRunoffRaster As ESRI.ArcGIS.Geodatabase.IRaster
        '        Dim pAccumRunoffRaster1 As ESRI.ArcGIS.Geodatabase.IRaster
        '        Dim pCellAreaSMRaster As ESRI.ArcGIS.Geodatabase.IRaster
        '        Dim pCellAreaSqKMRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 5a1: Square km
        '        Dim pCellAreaSqMileRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 5a2: Square miles
        '        Dim pCellAreaSFRaster As ESRI.ArcGIS.Geodatabase.IRaster
        '        Dim pCellAreaSIRaster As ESRI.ArcGIS.Geodatabase.IRaster
        '        Dim pCellAreaAcreRaster As ESRI.ArcGIS.Geodatabase.IRaster
        '        Dim pCellAreaCIRaster As ESRI.ArcGIS.Geodatabase.IRaster
        '        Dim pCellAreaCFRaster As ESRI.ArcGIS.Geodatabase.IRaster
        '        Dim pCellAreaAFRaster As ESRI.ArcGIS.Geodatabase.IRaster
        '        Dim pPermAccumRunoffRaster As ESRI.ArcGIS.Geodatabase.IRaster
        '        Dim pRunoffRasterLayer As ESRI.ArcGIS.Carto.IRasterLayer 'Rasterlayer of runoff
        '        Dim pPermAccumLocRunoffRaster As ESRI.ArcGIS.Geodatabase.IRaster
        '        Dim pLocRunoffRasterLayer As ESRI.ArcGIS.Carto.IRasterLayer 'Rasterlayer of local effects runoff

        '        Dim pEnv As ESRI.ArcGIS.GeoAnalyst.IRasterAnalysisEnvironment

        '        'Create Map Algebra Operator
        '        Dim pMapAlgebraOp As ESRI.ArcGIS.SpatialAnalyst.IMapAlgebraOp 'Workhorse

        '        'String to hold calculations
        '        Dim strExpression As String
        '        Dim strOutAccum As String
        '        Const strTitle As String = "Processing Runoff Calculation..."

        '        'TEMP CODE
        '        Dim strTemp As String
        '        Dim pTempRaster As ESRI.ArcGIS.Geodatabase.IRaster

        '        'Set the enviornment stuff
        '        pEnv = g_pSpatEnv
        '        pMapAlgebraOp = New ESRI.ArcGIS.SpatialAnalyst.RasterMapAlgebraOp
        '        pEnv = pMapAlgebraOp

        '        'Assign rasters
        '        pHydSoilsRaster = pInSoilsRaster
        '        pLandCoverRaster = pInLandCoverRaster
        '        pPrecipRaster = pInRainRaster

        '        modProgDialog.ProgDialog("Checking landcover cell size...", strTitle, 0, 10, 1, (frmPrj.m_App.hwnd))

        '        If modProgDialog.g_boolCancel Then
        '            'Step 1: ----------------------------------------------------------------------------
        '            'Make sure LandCover is in the same cellsize as the global environment

        '            pRasterProps = pLandCoverRaster

        '            If pRasterProps.MeanCellSize.X <> g_dblCellSize Then

        '                'UPGRADE_WARNING: Couldn't resolve default property of object pLandCoverRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                pMapAlgebraOp.BindRaster(pLandCoverRaster, "landcover")

        '                strExpression = "resample([landcover]," & CStr(g_dblCellSize) & ")"

        '                pLandSampleRaster = pMapAlgebraOp.Execute(strExpression)

        '                pMapAlgebraOp.UnbindRaster("landcover")
        '            Else
        '                pLandSampleRaster = pLandCoverRaster
        '            End If
        '        Else
        '            Exit Function
        '        End If

        '        'TEMP Code added 2008 to help w/ debugging
        '        'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        '        'Set pTempRaster = modUtil.ReturnPermanentRaster(pLandSampleRaster, pEnv.OutWorkspace.PathName, strTemp)


        '        modProgDialog.ProgDialog("Creating precipitation raster...", strTitle, 0, 10, 1, (frmPrj.m_App.hwnd))

        '        If modProgDialog.g_boolCancel Then
        '            'Step 2a: ----------------------------------------------------------------------------
        '            'Make the daily rain raster
        '            'UPGRADE_WARNING: Couldn't resolve default property of object pPrecipRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            pMapAlgebraOp.BindRaster(pPrecipRaster, "precip")

        '            strExpression = "[precip]"

        '            pDailyRainRaster = pMapAlgebraOp.Execute(strExpression)

        '            pMapAlgebraOp.UnbindRaster("precip")
        '        Else
        '            'GoTo ProgCancel
        '            Exit Function
        '        End If
        '        'END STEP 2a: -------------------------------------------------------------------------


        '        'TEMP Code added 2008 to help w/ debugging
        '        'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        '        'Set pTempRaster = modUtil.ReturnPermanentRaster(pDailyRainRaster, pEnv.OutWorkspace.PathName, strTemp)





        '        'UPGRADE_NOTE: Object pPrecipRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pPrecipRaster = Nothing

        '        modProgDialog.ProgDialog("Creating curve number GRID...", strTitle, 0, 10, 2, (frmPrj.m_App.hwnd))

        '        If modProgDialog.g_boolCancel Then

        '            'STEP 1: ----------------------------------------------------------------------------
        '            'Create Curve Number Grid
        '            With pMapAlgebraOp
        '                'UPGRADE_WARNING: Couldn't resolve default property of object pHydSoilsRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                .BindRaster(pHydSoilsRaster, "rev_soils")
        '                'UPGRADE_WARNING: Couldn't resolve default property of object pLandSampleRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                .BindRaster(pLandSampleRaster, "ccap")
        '            End With

        '            'Comes in
        '            strExpression = strPickStatement

        '            pSCSRaster = pMapAlgebraOp.Execute(strExpression)

        '            With pMapAlgebraOp
        '                .UnbindRaster("rev_soils")
        '                .UnbindRaster("ccap")
        '            End With
        '        Else
        '            GoTo ProgCancel
        '        End If


        '        'TEMP Code added 2008 to help w/ debugging
        '        'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        '        'Set pTempRaster = modUtil.ReturnPermanentRaster(pSCSRaster, pEnv.OutWorkspace.PathName, strTemp)


        '        'END STEP 1: --------------------------------------------------------------------------
        '        'UPGRADE_NOTE: Object pLandSampleRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pLandSampleRaster = Nothing
        '        'UPGRADE_NOTE: Object pLandCoverRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pLandCoverRaster = Nothing
        '        'UPGRADE_NOTE: Object pRasterProps may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pRasterProps = Nothing
        '        'UPGRADE_NOTE: Object pHydSoilsRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pHydSoilsRaster = Nothing

        '        'STEP 1a: Added January 2008 to combat null values in wet areas -----------------------

        '        modProgDialog.ProgDialog("Calculating maximum potential retention...", strTitle, 0, 10, 3, (frmPrj.m_App.hwnd))

        '        If modProgDialog.g_boolCancel Then
        '            'STEP 2: ------------------------------------------------------------------------------
        '            'Calculate maxiumum potential retention
        '            'UPGRADE_WARNING: Couldn't resolve default property of object pSCSRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            pMapAlgebraOp.BindRaster(pSCSRaster, "scsgrid")

        '            strExpression = "([scsgrid] * 100)"

        '            pSCS100Raster = pMapAlgebraOp.Execute(strExpression)
        '            g_pSCS100Raster = pSCS100Raster

        '            pMapAlgebraOp.UnbindRaster("scsgrid")
        '        Else
        '            GoTo ProgCancel
        '        End If
        '        'END STEP 2: ---------------------------------------------------------------------------


        '        'TEMP Code added 2008 to help w/ debugging
        '        'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        '        'Set pTempRaster = modUtil.ReturnPermanentRaster(g_pSCS100Raster, pEnv.OutWorkspace.PathName, strTemp)



        '        'UPGRADE_NOTE: Object pSCSRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pSCSRaster = Nothing

        '        modProgDialog.ProgDialog("Calculating maximum potential retention...", strTitle, 0, 10, 4, (frmPrj.m_App.hwnd))

        '        If modProgDialog.g_boolCancel Then
        '            'STEP 2A: ------------------------------------------------------------------------------
        '            'UPGRADE_WARNING: Couldn't resolve default property of object pSCS100Raster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            pMapAlgebraOp.BindRaster(pSCS100Raster, "scsgrid100")

        '            'Added 7/23 to account for rainman days
        '            If m_intPrecipType = 0 Then
        '                strExpression = "Float((1000. / [scsgrid100]) - 10) * " & m_intRainingDays
        '            Else
        '                strExpression = "Float((1000. / [scsgrid100]) - 10)"
        '            End If

        '            pRetentionRaster = pMapAlgebraOp.Execute(strExpression)

        '            pMapAlgebraOp.UnbindRaster("scsgrid100")
        '        Else
        '            GoTo ProgCancel
        '        End If
        '        'END STEP 2A: --------------------------------------------------------------------------


        '        'TEMP Code added 2008 to help w/ debugging
        '        'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        '        'Set pTempRaster = modUtil.ReturnPermanentRaster(pRetentionRaster, pEnv.OutWorkspace.PathName, strTemp)




        '        'UPGRADE_NOTE: Object pSCS100Raster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pSCS100Raster = Nothing

        '        modProgDialog.ProgDialog("Calculating initial abstraction...", strTitle, 0, 10, 5, (frmPrj.m_App.hwnd))

        '        If modProgDialog.g_boolCancel Then

        '            'STEP 3: -------------------------------------------------------------------------------
        '            'Calculate initial abstraction
        '            'UPGRADE_WARNING: Couldn't resolve default property of object pRetentionRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            pMapAlgebraOp.BindRaster(pRetentionRaster, "retention")

        '            strExpression = "(0.2 * [retention])"

        '            pAbstractRaster = pMapAlgebraOp.Execute(strExpression)
        '            g_pAbstractRaster = pAbstractRaster

        '            pMapAlgebraOp.UnbindRaster("retention")
        '        Else
        '            GoTo ProgCancel
        '        End If
        '        'END STEP 3: ---------------------------------------------------------------------------


        '        'TEMP Code added 2008 to help w/ debugging
        '        'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        '        'Set pTempRaster = modUtil.ReturnPermanentRaster(pAbstractRaster, pEnv.OutWorkspace.PathName, strTemp)

        '        modProgDialog.ProgDialog("Calculating runoff...", strTitle, 0, 10, 6, (frmPrj.m_App.hwnd))

        '        If modProgDialog.g_boolCancel Then
        '            'STEP 4: -------------------------------------------------------------------------------
        '            'Calculate Runoff
        '            With pMapAlgebraOp
        '                'UPGRADE_WARNING: Couldn't resolve default property of object pDailyRainRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                .BindRaster(pDailyRainRaster, "rainsix")
        '                'UPGRADE_WARNING: Couldn't resolve default property of object pAbstractRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                .BindRaster(pAbstractRaster, "abstract")
        '            End With

        '            strExpression = "Con(([rainsix] - [abstract]) > 0, ([rainsix] - [abstract]), 0.0)"

        '            'con( ([precip] - [abstraction] > 0.0 ), [precip] - [abstraction], 0.0 )

        '            pTemp1RunoffRaster = pMapAlgebraOp.Execute(strExpression)

        '            With pMapAlgebraOp
        '                .UnbindRaster("rainsix")
        '                .UnbindRaster("abstract")
        '            End With
        '            'END STEP 4: ----------------------------------------------------------------------------


        '            'TEMP Code added 2008 to help w/ debugging
        '            'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        '            'Set pTempRaster = modUtil.ReturnPermanentRaster(pTemp1RunoffRaster, pEnv.OutWorkspace.PathName, strTemp)

        '            'UPGRADE_NOTE: Object pDailyRainRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '            pDailyRainRaster = Nothing

        '            'STEP 4a: --------------------------------------------------------------------------------
        '            'Temp2 calc in runoff
        '            'UPGRADE_WARNING: Couldn't resolve default property of object pTemp1RunoffRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            pMapAlgebraOp.BindRaster(pTemp1RunoffRaster, "temp1")

        '            strExpression = "Pow([temp1], 2)"

        '            pTemp2RunoffRaster = pMapAlgebraOp.Execute(strExpression)

        '            pMapAlgebraOp.UnbindRaster("temp1")
        '            'END STEP 4a -----------------------------------------------------------------------------

        '            'TEMP Code added 2008 to help w/ debugging
        '            'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        '            'Set pTempRaster = modUtil.ReturnPermanentRaster(pTemp2RunoffRaster, pEnv.OutWorkspace.PathName, strTemp)



        '            'STEP 4b: --------------------------------------------------------------------------------
        '            With pMapAlgebraOp
        '                'UPGRADE_WARNING: Couldn't resolve default property of object pTemp1RunoffRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                .BindRaster(pTemp1RunoffRaster, "temp1")
        '                'UPGRADE_WARNING: Couldn't resolve default property of object pTemp2RunoffRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                .BindRaster(pTemp2RunoffRaster, "temp2")
        '                'UPGRADE_WARNING: Couldn't resolve default property of object pRetentionRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                .BindRaster(pRetentionRaster, "retention")
        '            End With

        '            strExpression = "Con([temp1] > 0, ([temp2] / ([temp1] + [retention])), 0)"

        '            pRunoffInRaster = pMapAlgebraOp.Execute(strExpression)
        '            g_pRunoffInchRaster = pRunoffInRaster

        '            'modUtil.AddRasterLayer frmPrj.m_App, pRunoffInRaster, "RunInRaster"

        '            With pMapAlgebraOp
        '                .UnbindRaster("temp1")
        '                .UnbindRaster("temp2")
        '                .UnbindRaster("retention")
        '            End With
        '        Else
        '            GoTo ProgCancel
        '        End If
        '        'END STEP 4b: -----------------------------------------------------------------------------



        '        'TEMP Code added 2008 to help w/ debugging
        '        'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        '        'Set pTempRaster = modUtil.ReturnPermanentRaster(pRunoffInRaster, pEnv.OutWorkspace.PathName, strTemp)



        '        'UPGRADE_NOTE: Object pTemp1RunoffRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pTemp1RunoffRaster = Nothing
        '        'UPGRADE_NOTE: Object pTemp2RunoffRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pTemp2RunoffRaster = Nothing
        '        'UPGRADE_NOTE: Object pRetentionRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pRetentionRaster = Nothing

        '        modProgDialog.ProgDialog("Calculating maximum potential retention...", strTitle, 0, 10, 7, (frmPrj.m_App.hwnd))

        '        If modProgDialog.g_boolCancel Then
        '            'STEP 5a: ---------------------------------------------------------------------------------
        '            'Create mask
        '            'UPGRADE_WARNING: Couldn't resolve default property of object g_pDEMRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            pMapAlgebraOp.BindRaster(g_pDEMRaster, "DEM")

        '            strExpression = "Con([DEM] >= 0, 1, 0)"

        '            pMaskRaster = pMapAlgebraOp.Execute(strExpression)

        '            pMapAlgebraOp.UnbindRaster("DEM")


        '            'TEMP Code added 2008 to help w/ debugging
        '            'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        '            'Set pTempRaster = modUtil.ReturnPermanentRaster(pMaskRaster, pEnv.OutWorkspace.PathName, strTemp)



        '            'Convert Runoff
        '            'Square Meters
        '            'UPGRADE_WARNING: Couldn't resolve default property of object pMaskRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            pMapAlgebraOp.BindRaster(pMaskRaster, "cellmask")

        '            strExpression = "Pow([cellmask] * " & g_dblCellSize & ", 2)"

        '            pCellAreaSMRaster = pMapAlgebraOp.Execute(strExpression)

        '            pMapAlgebraOp.UnbindRaster("cellmask")
        '            'END STEP 5a ------------------------------------------------------------------------------


        '            'TEMP Code added 2008 to help w/ debugging
        '            'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        '            'Set pTempRaster = modUtil.ReturnPermanentRaster(pCellAreaSMRaster, pEnv.OutWorkspace.PathName, strTemp)




        '            'UPGRADE_NOTE: Object pMaskRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '            pMaskRaster = Nothing

        '            'STEP 5a1: --------------------------------------------------------------------------------
        '            'Square Kilometers - Added for MUSLE use
        '            'UPGRADE_WARNING: Couldn't resolve default property of object pCellAreaSMRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            pMapAlgebraOp.BindRaster(pCellAreaSMRaster, "cellarea_sqm")

        '            strExpression = "[cellarea_sqm] * 0.000001"

        '            pCellAreaSqKMRaster = pMapAlgebraOp.Execute(strExpression)

        '            pMapAlgebraOp.UnbindRaster("cellarea_sqm")
        '            'END STEP 5a1: -----------------------------------------------------------------------------


        '            'TEMP Code added 2008 to help w/ debugging
        '            'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        '            'Set pTempRaster = modUtil.ReturnPermanentRaster(pCellAreaSqKMRaster, pEnv.OutWorkspace.PathName, strTemp)



        '            'STEP 5a2
        '            'Square Kilometers - Added for MUSLE use ---------------------------------------------------
        '            'UPGRADE_WARNING: Couldn't resolve default property of object pCellAreaSqKMRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            pMapAlgebraOp.BindRaster(pCellAreaSqKMRaster, "cellarea_sqkm")

        '            strExpression = "[cellarea_sqkm] * 0.386102"

        '            pCellAreaSqMileRaster = pMapAlgebraOp.Execute(strExpression)
        '            g_pCellAreaSqMiRaster = pCellAreaSqMileRaster

        '            pMapAlgebraOp.UnbindRaster("cellarea_sqkm")
        '            'END STEP 5a1: -----------------------------------------------------------------------------

        '            'TEMP Code added 2008 to help w/ debugging
        '            'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        '            'Set pTempRaster = modUtil.ReturnPermanentRaster(pCellAreaSqMileRaster, pEnv.OutWorkspace.PathName, strTemp)




        '            'UPGRADE_NOTE: Object pCellAreaSqKMRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '            pCellAreaSqKMRaster = Nothing

        '            'STEP 5b: ----------------------------------------------------------------------------------
        '            'Square Feet
        '            'UPGRADE_WARNING: Couldn't resolve default property of object pCellAreaSMRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            pMapAlgebraOp.BindRaster(pCellAreaSMRaster, "cellarea_sm")

        '            strExpression = "[cellarea_sm] * 10.76"

        '            pCellAreaSFRaster = pMapAlgebraOp.Execute(strExpression)

        '            pMapAlgebraOp.UnbindRaster("cellarea_sm")
        '            'END STEP 5b--------------------------------------------------------------------------------


        '            'TEMP Code added 2008 to help w/ debugging
        '            'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        '            'Set pTempRaster = modUtil.ReturnPermanentRaster(pCellAreaSFRaster, pEnv.OutWorkspace.PathName, strTemp)



        '            'STEP 5c: ----------------------------------------------------------------------------------
        '            'Square Inches
        '            'UPGRADE_WARNING: Couldn't resolve default property of object pCellAreaSFRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            pMapAlgebraOp.BindRaster(pCellAreaSFRaster, "cellarea_sf")

        '            strExpression = "[cellarea_sf] * 144"

        '            pCellAreaSIRaster = pMapAlgebraOp.Execute(strExpression)

        '            pMapAlgebraOp.UnbindRaster("cellarea_sf")
        '            'END STEP 5c--------------------------------------------------------------------------------

        '            'TEMP Code added 2008 to help w/ debugging
        '            'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        '            'Set pTempRaster = modUtil.ReturnPermanentRaster(pCellAreaSIRaster, pEnv.OutWorkspace.PathName, strTemp)




        '            'UPGRADE_NOTE: Object pCellAreaSFRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '            pCellAreaSFRaster = Nothing

        '            'STEP 5d: ----------------------------------------------------------------------------------
        '            'Acres
        '            'UPGRADE_WARNING: Couldn't resolve default property of object pCellAreaSMRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            pMapAlgebraOp.BindRaster(pCellAreaSMRaster, "cell_sm")

        '            strExpression = "[cell_sm] * 0.000247104369"

        '            pCellAreaAcreRaster = pMapAlgebraOp.Execute(strExpression)

        '            pMapAlgebraOp.UnbindRaster("cell_sm")
        '            'END STEP 5d--------------------------------------------------------------------------------


        '            'TEMP Code added 2008 to help w/ debugging
        '            'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        '            'Set pTempRaster = modUtil.ReturnPermanentRaster(pCellAreaAcreRaster, pEnv.OutWorkspace.PathName, strTemp)





        '            'UPGRADE_NOTE: Object pCellAreaSMRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '            pCellAreaSMRaster = Nothing

        '            'STEP 5e: ----------------------------------------------------------------------------------
        '            'Cubic Inches
        '            With pMapAlgebraOp
        '                'UPGRADE_WARNING: Couldn't resolve default property of object pCellAreaSIRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                .BindRaster(pCellAreaSIRaster, "cell_si")
        '                'UPGRADE_WARNING: Couldn't resolve default property of object pRunoffInRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                .BindRaster(pRunoffInRaster, "runoff_in")
        '            End With

        '            strExpression = "([cell_si] * [runoff_in])"

        '            pCellAreaCIRaster = pMapAlgebraOp.Execute(strExpression)

        '            With pMapAlgebraOp
        '                .UnbindRaster("cell_si")
        '                .UnbindRaster("runoff_in")
        '            End With
        '            'END STEP 5e -------------------------------------------------------------------------------

        '            'TEMP Code added 2008 to help w/ debugging
        '            'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        '            'Set pTempRaster = modUtil.ReturnPermanentRaster(pCellAreaCIRaster, pEnv.OutWorkspace.PathName, strTemp)




        '            'UPGRADE_NOTE: Object pCellAreaSIRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '            pCellAreaSIRaster = Nothing

        '            'STEP 5f: -----------------------------------------------------------------------------------
        '            'Cubic feet
        '            'UPGRADE_WARNING: Couldn't resolve default property of object pCellAreaCIRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            pMapAlgebraOp.BindRaster(pCellAreaCIRaster, "runoff_ci")

        '            strExpression = "([runoff_ci] * 0.0005787)"

        '            pCellAreaCFRaster = pMapAlgebraOp.Execute(strExpression)
        '            g_pRunoffCFRaster = pCellAreaCFRaster

        '            pMapAlgebraOp.UnbindRaster("runoff_ci")

        '            'END STEP 5f: --------------------------------------------------------------------------------

        '            'TEMP Code added 2008 to help w/ debugging
        '            'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        '            'Set pTempRaster = modUtil.ReturnPermanentRaster(pCellAreaCFRaster, pEnv.OutWorkspace.PathName, strTemp)




        '            'STEP 5g: ------------------------------------------------------------------------------------
        '            'Runoff acre-feet
        '            'UPGRADE_WARNING: Couldn't resolve default property of object pCellAreaCFRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            pMapAlgebraOp.BindRaster(pCellAreaCFRaster, "runoff_cf")

        '            strExpression = "([runoff_cf] * 0.000022957)"

        '            pCellAreaAFRaster = pMapAlgebraOp.Execute(strExpression)
        '            g_pRunoffAFRaster = pCellAreaAFRaster

        '            pMapAlgebraOp.UnbindRaster("runoff_cf")
        '            'END STEP 5f: --------------------------------------------------------------------------------
        '        Else
        '            GoTo ProgCancel
        '        End If


        '        'TEMP Code added 2008 to help w/ debugging
        '        'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        '        'Set pTempRaster = modUtil.ReturnPermanentRaster(pCellAreaAFRaster, pEnv.OutWorkspace.PathName, strTemp)



        '        modProgDialog.ProgDialog("Converting to correct units...", strTitle, 0, 10, 8, (frmPrj.m_App.hwnd))

        '        If modProgDialog.g_boolCancel Then
        '            'STEP 6: -------------------------------------------------------------------------------------
        '            'Convert cubic inches to liters
        '            If pCellAreaCIRaster Is Nothing Then
        '                MsgBox("nothing")
        '            End If

        '            'UPGRADE_WARNING: Couldn't resolve default property of object pCellAreaCIRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            pMapAlgebraOp.BindRaster(pCellAreaCIRaster, "run_ci")

        '            strExpression = "[run_ci] * 0.016387064"

        '            pMetRunoffRaster = pMapAlgebraOp.Execute(strExpression)
        '            g_pMetRunoffRaster = pMetRunoffRaster

        '            pMapAlgebraOp.UnbindRaster("run_ci")
        '            'END STEP 6: ---------------------------------------------------------------------------------
        '        Else
        '            GoTo ProgCancel
        '        End If

        '        'TEMP Code added 2008 to help w/ debugging
        '        'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        '        'Set pTempRaster = modUtil.ReturnPermanentRaster(pMetRunoffRaster, pEnv.OutWorkspace.PathName, strTemp)



        '        'STEP 6a: -------------------------------------------------------------------------------------
        '        'Eliminate nulls: Added 1/28/07
        '        'UPGRADE_WARNING: Couldn't resolve default property of object pMetRunoffRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        pMapAlgebraOp.BindRaster(pMetRunoffRaster, "runoffgrid")

        '        strExpression = "Con(IsNull([runoffgrid]),0,[runoffgrid])"

        '        pMetRunoffNoNullRaster = pMapAlgebraOp.Execute(strExpression)

        '        g_pMetRunoffRaster = pMetRunoffNoNullRaster

        '        pMapAlgebraOp.UnbindRaster("runoffgrid")


        '        'TEMP Code added 2008 to help w/ debugging
        '        'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        '        'Set pTempRaster = modUtil.ReturnPermanentRaster(pMetRunoffNoNullRaster, pEnv.OutWorkspace.PathName, strTemp)


        '        Dim pClipLocAccumRaster As ESRI.ArcGIS.Geodatabase.IRaster
        '        If g_booLocalEffects Then
        '            modProgDialog.ProgDialog("Creating data layer for local effects...", strTitle, 0, 10, 10, (frmPrj.m_App.hwnd))
        '            If modProgDialog.g_boolCancel Then

        '                'STEP 12: Local Effects -------------------------------------------------
        '                'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                strOutAccum = modUtil.GetUniqueName("locaccum", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        '                'Added 7/23/04 to account for clip by selected polys functionality
        '                If g_booSelectedPolys Then
        '                    pClipLocAccumRaster = modUtil.ClipBySelectedPoly(pMetRunoffNoNullRaster, g_pSelectedPolyClip, pEnv)
        '                    'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                    pPermAccumLocRunoffRaster = modUtil.ReturnPermanentRaster(pClipLocAccumRaster, pEnv.OutWorkspace.PathName, strOutAccum)
        '                    'UPGRADE_NOTE: Object pClipLocAccumRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '                    pClipLocAccumRaster = Nothing
        '                Else
        '                    'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                    pPermAccumLocRunoffRaster = modUtil.ReturnPermanentRaster(pMetRunoffNoNullRaster, pEnv.OutWorkspace.PathName, strOutAccum)
        '                End If

        '                pLocRunoffRasterLayer = modUtil.ReturnRasterLayer((frmPrj.m_App), pPermAccumLocRunoffRaster, "Runoff Local Effects (L)")
        '                pLocRunoffRasterLayer.Renderer = modUtil.ReturnRasterStretchColorRampRender(pLocRunoffRasterLayer, "Blue")
        '                'UPGRADE_WARNING: Couldn't resolve default property of object pLocRunoffRasterLayer.Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                pLocRunoffRasterLayer.Visible = False

        '                'UPGRADE_WARNING: Couldn't resolve default property of object pLocRunoffRasterLayer.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                g_dicMetadata.Add(pLocRunoffRasterLayer.Name, m_strRunoffMetadata)

        '                'UPGRADE_WARNING: Couldn't resolve default property of object pLocRunoffRasterLayer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                g_pGroupLayer.Add(pLocRunoffRasterLayer)

        '                RunoffCalculation = True
        '                modProgDialog.KillDialog()
        '                Exit Function

        '            End If
        '        End If

        '        modProgDialog.ProgDialog("Creating flow accumulation...", strTitle, 0, 10, 9, (frmPrj.m_App.hwnd))
        '        If modProgDialog.g_boolCancel Then
        '            'STEP 7: ------------------------------------------------------------------------------------
        '            'Derive Accumulated Runoff

        '            With pMapAlgebraOp
        '                'UPGRADE_WARNING: Couldn't resolve default property of object g_pFlowDirRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                .BindRaster(g_pFlowDirRaster, "flowdir")
        '                'UPGRADE_WARNING: Couldn't resolve default property of object pMetRunoffNoNullRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                .BindRaster(pMetRunoffNoNullRaster, "met_run")
        '            End With

        '            strExpression = "FlowAccumulation([flowdir], [met_run], FLOAT)"

        '            pAccumRunoffRaster1 = pMapAlgebraOp.Execute(strExpression)

        '            With pMapAlgebraOp
        '                .UnbindRaster("flowdir")
        '                .UnbindRaster("met_run")
        '            End With
        '            'END STEP 7: ----------------------------------------------------------------------------------
        '        Else
        '            GoTo ProgCancel
        '        End If

        '        'TEMP Code added 2008 to help w/ debugging
        '        'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        '        'Set pTempRaster = modUtil.ReturnPermanentRaster(pAccumRunoffRaster1, pEnv.OutWorkspace.PathName, strTemp)


        '        modProgDialog.ProgDialog("Creating flow accumulation...", strTitle, 0, 10, 9, (frmPrj.m_App.hwnd))
        '        If modProgDialog.g_boolCancel Then
        '            'STEP 7: ------------------------------------------------------------------------------------
        '            'Derive the real Accumulated Runoff

        '            With pMapAlgebraOp
        '                'UPGRADE_WARNING: Couldn't resolve default property of object pAccumRunoffRaster1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                .BindRaster(pAccumRunoffRaster1, "accum")
        '                'UPGRADE_WARNING: Couldn't resolve default property of object pMetRunoffNoNullRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                .BindRaster(pMetRunoffNoNullRaster, "met_run")
        '            End With

        '            strExpression = "[accum] + [met_run]"

        '            pAccumRunoffRaster = pMapAlgebraOp.Execute(strExpression)

        '            With pMapAlgebraOp
        '                .UnbindRaster("accum")
        '                .UnbindRaster("met_run")
        '            End With
        '            'END STEP 7: ----------------------------------------------------------------------------------
        '        Else
        '            GoTo ProgCancel
        '        End If

        '        'TEMP Code added 2008 to help w/ debugging
        '        'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        '        'Set pTempRaster = modUtil.ReturnPermanentRaster(pAccumRunoffRaster, pEnv.OutWorkspace.PathName, strTemp)


        '        'Add this then map as our runoff grid
        '        modProgDialog.ProgDialog("Creating Runoff Layer...", strTitle, 0, 10, 10, (frmPrj.m_App.hwnd))
        '        Dim pClipAccumRaster As ESRI.ArcGIS.Geodatabase.IRaster
        '        If modProgDialog.g_boolCancel Then
        '            'Get a unique name for accumulation GRID
        '            'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            strOutAccum = modUtil.GetUniqueName("runoff", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))

        '            'Clip to selected polys if chosen
        '            If g_booSelectedPolys Then
        '                pClipAccumRaster = modUtil.ClipBySelectedPoly(pAccumRunoffRaster, g_pSelectedPolyClip, pEnv)
        '                'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                pPermAccumRunoffRaster = modUtil.ReturnPermanentRaster(pClipAccumRaster, pEnv.OutWorkspace.PathName, strOutAccum)
        '                'UPGRADE_NOTE: Object pClipAccumRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '                pClipAccumRaster = Nothing
        '            Else
        '                'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                pPermAccumRunoffRaster = modUtil.ReturnPermanentRaster(pAccumRunoffRaster, pEnv.OutWorkspace.PathName, strOutAccum)
        '            End If

        '            'Add Completed raster to the g_pGroupLayer
        '            pRunoffRasterLayer = modUtil.ReturnRasterLayer((frmPrj.m_App), pPermAccumRunoffRaster, "Accumulated Runoff (L)")
        '            pRunoffRasterLayer.Renderer = ReturnRasterStretchColorRampRender(pRunoffRasterLayer, "Blue")
        '            'UPGRADE_WARNING: Couldn't resolve default property of object pRunoffRasterLayer.Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            pRunoffRasterLayer.Visible = False
        '            'UPGRADE_WARNING: Couldn't resolve default property of object pRunoffRasterLayer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            g_pGroupLayer.Add(pRunoffRasterLayer)

        '            'UPGRADE_WARNING: Couldn't resolve default property of object pRunoffRasterLayer.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            g_dicMetadata.Add(pRunoffRasterLayer.Name, m_strRunoffMetadata)

        '            'Global Runoff
        '            g_pRunoffRaster = pPermAccumRunoffRaster
        '        Else
        '            GoTo ProgCancel
        '        End If

        '        RunoffCalculation = True

        '        modProgDialog.KillDialog()

        '        'Cleanup
        '        'UPGRADE_NOTE: Object pCellAreaSMRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pCellAreaSMRaster = Nothing
        '        'UPGRADE_NOTE: Object pCellAreaAcreRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pCellAreaAcreRaster = Nothing
        '        'UPGRADE_NOTE: Object pCellAreaCIRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pCellAreaCIRaster = Nothing
        '        'UPGRADE_NOTE: Object pAccumRunoffRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pAccumRunoffRaster = Nothing
        '        'UPGRADE_NOTE: Object pAccumRunoffRaster1 may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pAccumRunoffRaster1 = Nothing

        '        Exit Function

        'ErrHandler:
        '        If Err.Number = -2147217297 Then 'User cancelled operation
        '            modProgDialog.g_boolCancel = False
        '            RunoffCalculation = False
        '        ElseIf Err.Number = -2147467259 Then
        '            MsgBox("ArcMap has reached the maximum number of GRIDs allowed in memory.  " & "Please exit N-SPECT and restart ArcMap.", MsgBoxStyle.Information, "Maximum GRID Number Encountered")
        '            modProgDialog.g_boolCancel = False
        '            modProgDialog.KillDialog()
        '            RunoffCalculation = False
        '        Else
        '            MsgBox("Error: " & Err.Number & " on RunoffCalculation: " & strExpression)
        '            modProgDialog.g_boolCancel = False
        '            modProgDialog.KillDialog()
        '            RunoffCalculation = False
        '        End If

        'ProgCancel:


    End Function

    Private Function CreateMetadata(ByRef cmdLandClass As OleDbCommand, ByRef strLandClass As String, ByRef booLocal As Boolean) As String
        '        On Error GoTo ErrorHandler


        '        Dim i As Short
        '        Dim strHeader As String
        '        Dim strLCHeader As String
        '        Dim strCoeffValues As String

        '        If booLocal Then
        '            strHeader = vbTab & "Input Datasets:" & vbNewLine & vbTab & vbTab & "Hydrologic soils grid: " & g_clsXMLPrjFile.strSoilsHydFileName & vbNewLine & vbTab & vbTab & "Landcover grid: " & g_clsXMLPrjFile.strLCGridFileName & vbNewLine & vbTab & vbTab & "Precipitation grid: " & g_strPrecipFileName & vbNewLine
        '        Else
        '            strHeader = vbTab & "Input Datasets:" & vbNewLine & vbTab & vbTab & "Hydrologic soils grid: " & g_clsXMLPrjFile.strSoilsHydFileName & vbNewLine & vbTab & vbTab & "Landcover grid: " & g_clsXMLPrjFile.strLCGridFileName & vbNewLine & vbTab & vbTab & "Precipitation grid: " & g_strPrecipFileName & vbNewLine & vbTab & vbTab & "Flow direction grid: " & g_strFlowDirFilename & vbNewLine
        '        End If

        '        strLCHeader = vbTab & "Landcover parameters: " & vbNewLine & vbTab & vbTab & "Landcover type name: " & strLandClass & vbNewLine & vbTab & vbTab & "Coefficients: " & vbNewLine

        '        rsLandClass.MoveFirst()

        '        For i = 1 To rsLandClass.RecordCount
        '            If i = 1 Then
        '                strCoeffValues = vbTab & vbTab & vbTab & rsLandClass.Fields("Name2").Value & ":" & vbNewLine & vbTab & vbTab & vbTab & vbTab & "CN-A: " & CStr(rsLandClass.Fields("CN-A").Value) & vbNewLine & vbTab & vbTab & vbTab & vbTab & "CN-B: " & CStr(rsLandClass.Fields("CN-B").Value) & vbNewLine & vbTab & vbTab & vbTab & vbTab & "CN-C: " & CStr(rsLandClass.Fields("CN-C").Value) & vbNewLine & vbTab & vbTab & vbTab & vbTab & "CN-D: " & CStr(rsLandClass.Fields("CN-D").Value) & vbNewLine & vbTab & vbTab & vbTab & vbTab & "Cover factor: " & rsLandClass.Fields("CoverFactor").Value & vbNewLine

        '            Else
        '                strCoeffValues = strCoeffValues & vbTab & vbTab & vbTab & rsLandClass.Fields("Name2").Value & ":" & vbNewLine & vbTab & vbTab & vbTab & vbTab & "CN-A: " & CStr(rsLandClass.Fields("CN-A").Value) & vbNewLine & vbTab & vbTab & vbTab & vbTab & "CN-B: " & CStr(rsLandClass.Fields("CN-B").Value) & vbNewLine & vbTab & vbTab & vbTab & vbTab & "CN-C: " & CStr(rsLandClass.Fields("CN-C").Value) & vbNewLine & vbTab & vbTab & vbTab & vbTab & "CN-D: " & CStr(rsLandClass.Fields("CN-D").Value) & vbNewLine & vbTab & vbTab & vbTab & vbTab & "Cover factor: " & rsLandClass.Fields("CoverFactor").Value & vbNewLine
        '            End If
        '            rsLandClass.MoveNext()
        '        Next i

        '        CreateMetadata = strHeader & strLCHeader & strCoeffValues
        '        g_strLandCoverParameters = strLCHeader & strCoeffValues



        '        Exit Function
        'ErrorHandler:
        '        HandleError(False, "CreateMetadata " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
    End Function
End Module