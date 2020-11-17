Attribute VB_Name = "modRunoff"
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

Option Explicit
Public g_pSoilsRaster As IRasterDataset             'Soils Raster GRID
Public g_SoilsRasterDataset As IRasterDataset
Public g_LandCoverRaster As IRaster                 'LandCover GLOBAL
Public g_pRunoffRaster As IRaster                   'Runoff for use elsewhere
Public g_pMetRunoffRaster As IRaster                'Metric Runoff
Public g_pSCS100Raster As IRaster                   'SCS GRID for use in RUSLE
Public g_pAbstractRaster As IRaster                 'Abstraction Raster used in MUSLE
Public g_pPrecipRaster As IRaster                   'Precipitation Raster
Public g_pCellAreaSqMiRaster As IRaster             'Square Mile area raster used in MUSLE
Public g_pRunoffInchRaster As IRaster               'Runoff Inches used in MUSLE
Public g_pRunoffAFRaster As IRaster                 'Runoff AF used in MUSLE
Public g_pRunoffCFRaster As IRaster                 'MUSLE
Public g_strPrecipFileName As String                'Precipitation File Name
Public g_strLandCoverParameters As String           'Global string to hold formatted LC params for metadata

Private m_strRunoffMetadata As String               'Metatdata string for runoff
Private m_intPrecipType As Integer                  'Precip Event Type: 0=Annual; 1=Event
Private m_intRainingDays As Integer                 '# of Raining Days for Annual Precip
' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "modRunoff.bas"
Private m_ParentHWND As Long          ' Set this to get correct parenting of Error handler forms



Private Function ConstructPickStatment(strLandClass As String, pLCRaster As IRaster) As String
'Creates the large initial pick statement using the name of the the LandCass [CCAP, for example]
'and the Land Class Raster.  Returns a string

    Dim strRS As String
    Dim rsLandClass As ADODB.Recordset
    Dim strPick(3) As String     'Array of strings that hold 'pick' numbers
    Dim strCurveCalc As String   'Full String
    
    Dim pRasterCol As IRasterBandCollection
    Dim pBand As IRasterBand
    Dim pRasterStats As IRasterStatistics
    Dim dblMaxValue As Double
    Dim i As Integer
    Dim pTable As ITable
    Dim TableExist As Boolean
    Dim pCursor As ICursor
    Dim pRow As iRow
    Dim FieldIndex As Integer
    Dim booValueFound As Boolean
    Dim j As Integer
    
On Error GoTo ErrHandler:

    'STEP 1:  get the records from the database -----------------------------------------------
    strRS = "SELECT LCCLASS.LCClassID, Value, LCCLASS.Name as Name2, LCCLASS.LCTypeID, [CN-A], [CN-B], [CN-C], [CN-D], CoverFactor, W_WL FROM LCCLASS " & _
     "LEFT OUTER JOIN LCTYPE ON LCCLASS.LCTYPEID = LCTYPE.LCTYPEID " & _
     "Where LCTYPE.NAME = '" & strLandClass & "' ORDER BY LCCLASS.VALUE"
    
    Set rsLandClass = New ADODB.Recordset
    rsLandClass.Open strRS, g_ADOConn, adOpenKeyset, adLockOptimistic
    

    'while here, make metadata
    m_strRunoffMetadata = CreateMetadata(rsLandClass, strLandClass, g_booLocalEffects)
    'End Database stuff
    
    'STEP 2: Raster Values ---------------------------------------------------------------------
    'Now Get the RASTER values
    ' Get Rasterband from the incoming raster
    Set pRasterCol = pLCRaster
    Set pBand = pRasterCol.Item(0)
    
    'Get the max value
    Set pRasterStats = pBand.Statistics
    dblMaxValue = pRasterStats.Maximum

    'Get the raster table
    pBand.HasTable TableExist
    If Not TableExist Then Exit Function
    
    Set pTable = pBand.AttributeTable
    'Get All rows
    Set pCursor = pTable.Search(Nothing, True)
    'Init pRow
    Set pRow = pCursor.NextRow
    
    'Get index of Value Field
    FieldIndex = pTable.FindField("Value")
    
    'STEP 3: Table values vs. Raster Values Count - if not equal bark -------------------------
    'If pTable.RowCount(Nothing) <> rsLandClass.RecordCount Then
     '   MsgBox "Error: The number of records in your Land Class Table do not match your GRID dataset."
     '   Exit Function
   'End If
    rsLandClass.MoveFirst
    
    'STEP 4: Create the strings
    'Loop through and get all values
    For i = 1 To dblMaxValue
        'Check first value, usually they won't have a [1], but if there check against database,
        'else move on
        If i = 1 Then
            If (pRow.Value(FieldIndex) = i) Then 'And (pRow.Value(FieldIndex) = rsLandClass!Value) Then
                rsLandClass.MoveFirst
                booValueFound = False
                For j = 0 To rsLandClass.RecordCount - 0
                    If pRow.Value(FieldIndex) = rsLandClass!Value Then
                        booValueFound = True
                        strPick(0) = CStr(rsLandClass![CN-A])
                        strPick(1) = CStr(rsLandClass![CN-B])
                        strPick(2) = CStr(rsLandClass![CN-C])
                        strPick(3) = CStr(rsLandClass![CN-D])
                        'rsLandClass.MoveNext
                        Set pRow = pCursor.NextRow
                        Exit For
                    Else
                        booValueFound = False
                        rsLandClass.MoveNext
                    End If
                Next j
                If booValueFound = False Then
                    MsgBox "Error: Your N-SPECT Land Class Table is missing values found in your landcover GRID dataset."
                    ConstructPickStatment = ""
                    Exit Function
                End If
            Else
                strPick(0) = "-9999"
                strPick(1) = "-9999"
                strPick(2) = "-9999"
                strPick(3) = "-9999"
            End If
        Else
            If (pRow.Value(FieldIndex) = i) Then 'And (pRow.Value(FieldIndex) = rsLandClass!Value) Then
                rsLandClass.MoveFirst
                booValueFound = False
                For j = 0 To rsLandClass.RecordCount - 0
                    If pRow.Value(FieldIndex) = rsLandClass!Value Then
                        booValueFound = True
                        strPick(0) = strPick(0) & ", " & CStr(rsLandClass![CN-A])
                        strPick(1) = strPick(1) & ", " & CStr(rsLandClass![CN-B])
                        strPick(2) = strPick(2) & ", " & CStr(rsLandClass![CN-C])
                        strPick(3) = strPick(3) & ", " & CStr(rsLandClass![CN-D])
                        'rsLandClass.MoveNext
                        Set pRow = pCursor.NextRow
                        Exit For
                    Else
                        booValueFound = False
                        rsLandClass.MoveNext
                    End If
                Next j
                If booValueFound = False Then
                    MsgBox "Error: Your N-SPECT Land Class Table is missing values found in your landcover GRID dataset."
                    ConstructPickStatment = ""
                    Exit Function
                End If
            Else
                strPick(0) = strPick(0) & ", 0"
                strPick(1) = strPick(1) & ", 0"
                strPick(2) = strPick(2) & ", 0"
                strPick(3) = strPick(3) & ", 0"
            End If
        End If
    Next i
    
    strCurveCalc = "con([rev_soils] == 1, pick([ccap], " & strPick(0) & "), con([rev_soils] == 2, pick([ccap], " & _
    strPick(1) & "), con([rev_soils] == 3, pick([ccap], " & strPick(2) & "), con([rev_soils] == 4, pick([ccap], " & _
    strPick(3) & ")))))"

    ConstructPickStatment = strCurveCalc
    
    
    
    'Cleanup:
    Set pRasterCol = Nothing
    Set pRasterStats = Nothing
    Set pBand = Nothing
    Set pTable = Nothing
    Set pCursor = Nothing
    Set pRow = Nothing

Exit Function
ErrHandler:
    MsgBox Err.Number & ": " & Err.Description & " " & "ConstructPickStatemnt"
   
    
End Function

Public Function CreateRunoffGrid(strLCFileName As String, strLCCLassType As String, rsPrecip As ADODB.Recordset, strSoilsFileName As String) As Boolean
'This sub serves as a link between frmPrj and the actual calculation of Runoff
'It establishes the Rasters being used

    'strLCFileName: Path to land class file
    'strLCCLassType: type of landclass
    'rsPrecip: ADO recordset of precip scenario being used
    Dim pRainFallRaster As IRaster
    Dim pLandCoverRaster As IRaster
    Dim pSoilsRaster As IRaster
    Dim strPickStatement As String
    Dim strError
    
    'Get the LandCover Raster
    'if no management scenarios were applied then g_landcoverRaster will be nothing
    If g_LandCoverRaster Is Nothing Then
        If modUtil.RasterExists(strLCFileName) Then
            Set g_LandCoverRaster = modUtil.ReturnRaster(strLCFileName)
            Set pLandCoverRaster = g_LandCoverRaster
        Else
            strError = strLCFileName
            GoTo Missing
        End If
    Else
        Set pLandCoverRaster = g_LandCoverRaster
    End If
    
    'Get the Precip Raster
    If modUtil.RasterExists(rsPrecip!PrecipFileName) Then
        Set pRainFallRaster = modUtil.ReturnRaster(rsPrecip!PrecipFileName)
        
        'If in cm, then convert to a GRID in inches.
        If rsPrecip!PrecipUnits = 0 Then
            Set g_pPrecipRaster = ConvertRainGridCMToInches(pRainFallRaster)
        Else 'if already in inches then just use this one.
            Set g_pPrecipRaster = pRainFallRaster       'Global Precip
        End If
        
        g_strPrecipFileName = rsPrecip!PrecipFileName
        m_intPrecipType = rsPrecip!Type             'Set the mod level precip type
        m_intRainingDays = rsPrecip!RainingDays     'Set the mod level rainingdays
    Else
        strError = strLCFileName
        GoTo Missing
    End If
    
    '------ 1.1.4 Change -----
    'if they select cm as incoming precip units, convert GRID
    
    'Get the Soils Raster
    If modUtil.RasterExists(strSoilsFileName) Then
        Set pSoilsRaster = modUtil.ReturnRaster(strSoilsFileName)
    Else
        strError = strSoilsFileName
        GoTo Missing
    End If
    
    'Now construct the Pick Statement
    strPickStatement = ConstructPickStatment(strLCCLassType, pLandCoverRaster)
    
    If Len(strPickStatement) = 0 Then
        Exit Function
    End If
            
    'Call the Runoff Calculation using the string and rasters
    If RunoffCalculation(strPickStatement, g_pPrecipRaster, pLandCoverRaster, pSoilsRaster) Then
        CreateRunoffGrid = True
    Else
        CreateRunoffGrid = False
        Exit Function
    End If
    
    CreateRunoffGrid = True

Exit Function
Missing:
    MsgBox "Error: The following dataset is missing: " & strError, vbCritical, "Missing Data"
    CreateRunoffGrid = False
    Exit Function

End Function

Private Function ConvertRainGridCMToInches(pInRaster As IRaster) As IRaster
    
    Dim pMapAlgebraOp As IMapAlgebraOp
    Dim pEnv As IRasterAnalysisEnvironment
    Dim strExpression As String
    
    Set pEnv = g_pSpatEnv
    Set pMapAlgebraOp = New RasterMapAlgebraOp
    Set pEnv = pMapAlgebraOp
    
    pMapAlgebraOp.BindRaster pInRaster, "precip"
    strExpression = "[precip] / 2.54"
    
    Set ConvertRainGridCMToInches = pMapAlgebraOp.Execute(strExpression)
    
    pMapAlgebraOp.UnbindRaster "precip"
       
     
End Function


Public Function RunoffCalculation(strPickStatement As String, pInRainRaster As IRaster, pInLandCoverRaster As IRaster, pInSoilsRaster As IRaster) As Boolean
'strPickStatement: our friend the dynamic pick statemnt
'pInRainRaster: the precip grid
'pInLandCoverRaster: landcover grid
On Error GoTo ErrHandler:
   'The rasters                             'Associated Steps
    Dim pRasterProps As IRasterProps
    Dim pLandCoverRaster As IRaster         'STEP 1: LandCover Raster
    Dim pLandSampleRaster As IRaster        'STEP 1: Properly sized landcover
    Dim pHydSoilsRaster As IRaster          'STEP 2: Soils Raster
    Dim pPrecipRaster As IRaster            'STEP 3: Create Precip Grid
    Dim pSCSRaster As IRaster               'STEP 1: Create Curve Number GRID
    Dim pDailyRainRaster As IRaster         'Rain
    Dim pSCS100Raster As IRaster            'STEP 2: SCS * 100
    Dim pRetentionRaster As IRaster         'STEP 3: Retention
    Dim pAbstractRaster As IRaster
    Dim pTemp1RunoffRaster As IRaster
    Dim pTemp2RunoffRaster As IRaster
    Dim pRunoffInRaster As IRaster
    Dim pMaskRaster As IRaster
    Dim pMetRunoffRaster As IRaster
    Dim pMetRunoffNoNullRaster As IRaster    'STEP 6a:  no nulls
    Dim pAccumRunoffRaster As IRaster
    Dim pAccumRunoffRaster1 As IRaster
    Dim pCellAreaSMRaster As IRaster
    Dim pCellAreaSqKMRaster As IRaster       'STEP 5a1: Square km
    Dim pCellAreaSqMileRaster As IRaster     'STEP 5a2: Square miles
    Dim pCellAreaSFRaster As IRaster
    Dim pCellAreaSIRaster As IRaster
    Dim pCellAreaAcreRaster As IRaster
    Dim pCellAreaCIRaster As IRaster
    Dim pCellAreaCFRaster As IRaster
    Dim pCellAreaAFRaster As IRaster
    Dim pPermAccumRunoffRaster As IRaster
    Dim pRunoffRasterLayer As IRasterLayer      'Rasterlayer of runoff
    Dim pPermAccumLocRunoffRaster As IRaster
    Dim pLocRunoffRasterLayer As IRasterLayer   'Rasterlayer of local effects runoff
    
    Dim pEnv As IRasterAnalysisEnvironment

    'Create Map Algebra Operator
    Dim pMapAlgebraOp As IMapAlgebraOp      'Workhorse

    'String to hold calculations
    Dim strExpression As String
    Dim strOutAccum As String
    Const strTitle As String = "Processing Runoff Calculation..."
    
    'TEMP CODE
    Dim strTemp As String
    Dim pTempRaster As IRaster

    'Set the enviornment stuff
    Set pEnv = g_pSpatEnv
    Set pMapAlgebraOp = New RasterMapAlgebraOp
    Set pEnv = pMapAlgebraOp

    'Assign rasters
    Set pHydSoilsRaster = pInSoilsRaster
    Set pLandCoverRaster = pInLandCoverRaster
    Set pPrecipRaster = pInRainRaster
    
    modProgDialog.ProgDialog "Checking landcover cell size...", strTitle, 0, 10, 1, frmPrj.m_App.hwnd
    
    If modProgDialog.g_boolCancel Then
    'Step 1: ----------------------------------------------------------------------------
    'Make sure LandCover is in the same cellsize as the global environment
    
    Set pRasterProps = pLandCoverRaster
        
        If pRasterProps.MeanCellSize.X <> g_dblCellSize Then
            
            pMapAlgebraOp.BindRaster pLandCoverRaster, "landcover"
    
            strExpression = "resample([landcover]," & CStr(g_dblCellSize) & ")"
    
            Set pLandSampleRaster = pMapAlgebraOp.Execute(strExpression)
    
            pMapAlgebraOp.UnbindRaster "landcover"
        Else
            Set pLandSampleRaster = pLandCoverRaster
        End If
    Else
        Exit Function
    End If
    
    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pLandSampleRaster, pEnv.OutWorkspace.PathName, strTemp)
    

    modProgDialog.ProgDialog "Creating precipitation raster...", strTitle, 0, 10, 1, frmPrj.m_App.hwnd
    
    If modProgDialog.g_boolCancel Then
    'Step 2a: ----------------------------------------------------------------------------
    'Make the daily rain raster
        pMapAlgebraOp.BindRaster pPrecipRaster, "precip"
    
        strExpression = "[precip]"
    
        Set pDailyRainRaster = pMapAlgebraOp.Execute(strExpression)
        
        pMapAlgebraOp.UnbindRaster "precip"
    Else
        'GoTo ProgCancel
        Exit Function
    End If
    'END STEP 2a: -------------------------------------------------------------------------
    
    
    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pDailyRainRaster, pEnv.OutWorkspace.PathName, strTemp)
    

    
    
    
    Set pPrecipRaster = Nothing
    
    modProgDialog.ProgDialog "Creating curve number GRID...", strTitle, 0, 10, 2, frmPrj.m_App.hwnd
    
    If modProgDialog.g_boolCancel Then
    
    'STEP 1: ----------------------------------------------------------------------------
    'Create Curve Number Grid
        With pMapAlgebraOp
            .BindRaster pHydSoilsRaster, "rev_soils"
            .BindRaster pLandSampleRaster, "ccap"
        End With
    
        'Comes in
        strExpression = strPickStatement
    
        Set pSCSRaster = pMapAlgebraOp.Execute(strExpression)
    
        With pMapAlgebraOp
            .UnbindRaster "rev_soils"
            .UnbindRaster "ccap"
        End With
    Else
        GoTo ProgCancel
    End If
    
    
    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pSCSRaster, pEnv.OutWorkspace.PathName, strTemp)
    
    
    'END STEP 1: --------------------------------------------------------------------------
    Set pLandSampleRaster = Nothing
    Set pLandCoverRaster = Nothing
    Set pRasterProps = Nothing
    Set pHydSoilsRaster = Nothing
    
    'STEP 1a: Added January 2008 to combat null values in wet areas -----------------------
    
    modProgDialog.ProgDialog "Calculating maximum potential retention...", strTitle, 0, 10, 3, frmPrj.m_App.hwnd
    
    If modProgDialog.g_boolCancel Then
    'STEP 2: ------------------------------------------------------------------------------
    'Calculate maxiumum potential retention
        pMapAlgebraOp.BindRaster pSCSRaster, "scsgrid"
    
        strExpression = "([scsgrid] * 100)"
    
        Set pSCS100Raster = pMapAlgebraOp.Execute(strExpression)
        Set g_pSCS100Raster = pSCS100Raster
    
        pMapAlgebraOp.UnbindRaster "scsgrid"
    Else
        GoTo ProgCancel
    End If
    'END STEP 2: ---------------------------------------------------------------------------
    
    
    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(g_pSCS100Raster, pEnv.OutWorkspace.PathName, strTemp)
    


    Set pSCSRaster = Nothing

    modProgDialog.ProgDialog "Calculating maximum potential retention...", strTitle, 0, 10, 4, frmPrj.m_App.hwnd
    
    If modProgDialog.g_boolCancel Then
    'STEP 2A: ------------------------------------------------------------------------------
        pMapAlgebraOp.BindRaster pSCS100Raster, "scsgrid100"
        
        'Added 7/23 to account for rainman days
        If m_intPrecipType = 0 Then
            strExpression = "Float((1000. / [scsgrid100]) - 10) * " & m_intRainingDays
        Else
            strExpression = "Float((1000. / [scsgrid100]) - 10)"
        End If
        
        Set pRetentionRaster = pMapAlgebraOp.Execute(strExpression)
    
        pMapAlgebraOp.UnbindRaster "scsgrid100"
     Else
        GoTo ProgCancel
    End If
    'END STEP 2A: --------------------------------------------------------------------------
    
    
    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pRetentionRaster, pEnv.OutWorkspace.PathName, strTemp)
    

    
    
    Set pSCS100Raster = Nothing

    modProgDialog.ProgDialog "Calculating initial abstraction...", strTitle, 0, 10, 5, frmPrj.m_App.hwnd
    
    If modProgDialog.g_boolCancel Then

    'STEP 3: -------------------------------------------------------------------------------
    'Calculate initial abstraction
        pMapAlgebraOp.BindRaster pRetentionRaster, "retention"
    
        strExpression = "(0.2 * [retention])"
        
        Set pAbstractRaster = pMapAlgebraOp.Execute(strExpression)
        Set g_pAbstractRaster = pAbstractRaster
    
        pMapAlgebraOp.UnbindRaster "retention"
    Else
        GoTo ProgCancel
    End If
    'END STEP 3: ---------------------------------------------------------------------------
    
    
    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pAbstractRaster, pEnv.OutWorkspace.PathName, strTemp)
    
    modProgDialog.ProgDialog "Calculating runoff...", strTitle, 0, 10, 6, frmPrj.m_App.hwnd
    
    If modProgDialog.g_boolCancel Then
    'STEP 4: -------------------------------------------------------------------------------
    'Calculate Runoff
        With pMapAlgebraOp
            .BindRaster pDailyRainRaster, "rainsix"
            .BindRaster pAbstractRaster, "abstract"
        End With
    
        strExpression = "Con(([rainsix] - [abstract]) > 0, ([rainsix] - [abstract]), 0.0)"
        
        'con( ([precip] - [abstraction] > 0.0 ), [precip] - [abstraction], 0.0 )
    
        Set pTemp1RunoffRaster = pMapAlgebraOp.Execute(strExpression)
    
        With pMapAlgebraOp
            .UnbindRaster "rainsix"
            .UnbindRaster "abstract"
        End With
    'END STEP 4: ----------------------------------------------------------------------------
    
    
    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pTemp1RunoffRaster, pEnv.OutWorkspace.PathName, strTemp)
    
     Set pDailyRainRaster = Nothing
    
    'STEP 4a: --------------------------------------------------------------------------------
    'Temp2 calc in runoff
        pMapAlgebraOp.BindRaster pTemp1RunoffRaster, "temp1"
    
        strExpression = "Pow([temp1], 2)"
    
        Set pTemp2RunoffRaster = pMapAlgebraOp.Execute(strExpression)
    
        pMapAlgebraOp.UnbindRaster "temp1"
        'END STEP 4a -----------------------------------------------------------------------------
        
        'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pTemp2RunoffRaster, pEnv.OutWorkspace.PathName, strTemp)
    

    
        'STEP 4b: --------------------------------------------------------------------------------
        With pMapAlgebraOp
            .BindRaster pTemp1RunoffRaster, "temp1"
            .BindRaster pTemp2RunoffRaster, "temp2"
            .BindRaster pRetentionRaster, "retention"
        End With
    
        strExpression = "Con([temp1] > 0, ([temp2] / ([temp1] + [retention])), 0)"
    
        Set pRunoffInRaster = pMapAlgebraOp.Execute(strExpression)
        Set g_pRunoffInchRaster = pRunoffInRaster
        
        'modUtil.AddRasterLayer frmPrj.m_App, pRunoffInRaster, "RunInRaster"
        
        With pMapAlgebraOp
            .UnbindRaster "temp1"
            .UnbindRaster "temp2"
            .UnbindRaster "retention"
        End With
    Else
        GoTo ProgCancel
    End If
    'END STEP 4b: -----------------------------------------------------------------------------
    
    
    
    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pRunoffInRaster, pEnv.OutWorkspace.PathName, strTemp)
    

    
    Set pTemp1RunoffRaster = Nothing
    Set pTemp2RunoffRaster = Nothing
    Set pRetentionRaster = Nothing
    
    modProgDialog.ProgDialog "Calculating maximum potential retention...", strTitle, 0, 10, 7, frmPrj.m_App.hwnd
    
    If modProgDialog.g_boolCancel Then
    'STEP 5a: ---------------------------------------------------------------------------------
    'Create mask
        pMapAlgebraOp.BindRaster g_pDEMRaster, "DEM"
        
        strExpression = "Con([DEM] >= 0, 1, 0)"
        
        Set pMaskRaster = pMapAlgebraOp.Execute(strExpression)
        
        pMapAlgebraOp.UnbindRaster "DEM"
        
        
    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pMaskRaster, pEnv.OutWorkspace.PathName, strTemp)
    


    'Convert Runoff
    'Square Meters
        pMapAlgebraOp.BindRaster pMaskRaster, "cellmask"
    
        strExpression = "Pow([cellmask] * " & g_dblCellSize & ", 2)"
    
        Set pCellAreaSMRaster = pMapAlgebraOp.Execute(strExpression)
    
        pMapAlgebraOp.UnbindRaster "cellmask"
    'END STEP 5a ------------------------------------------------------------------------------
    
    
    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pCellAreaSMRaster, pEnv.OutWorkspace.PathName, strTemp)
    

    
    
    Set pMaskRaster = Nothing
    
    'STEP 5a1: --------------------------------------------------------------------------------
    'Square Kilometers - Added for MUSLE use
        pMapAlgebraOp.BindRaster pCellAreaSMRaster, "cellarea_sqm"
        
        strExpression = "[cellarea_sqm] * 0.000001"
        
        Set pCellAreaSqKMRaster = pMapAlgebraOp.Execute(strExpression)
        
        pMapAlgebraOp.UnbindRaster "cellarea_sqm"
    'END STEP 5a1: -----------------------------------------------------------------------------
    
    
    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pCellAreaSqKMRaster, pEnv.OutWorkspace.PathName, strTemp)
    

    
    'STEP 5a2
    'Square Kilometers - Added for MUSLE use ---------------------------------------------------
        pMapAlgebraOp.BindRaster pCellAreaSqKMRaster, "cellarea_sqkm"
        
        strExpression = "[cellarea_sqkm] * 0.386102"
        
        Set pCellAreaSqMileRaster = pMapAlgebraOp.Execute(strExpression)
        Set g_pCellAreaSqMiRaster = pCellAreaSqMileRaster
        
        pMapAlgebraOp.UnbindRaster "cellarea_sqkm"
    'END STEP 5a1: -----------------------------------------------------------------------------
    
    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pCellAreaSqMileRaster, pEnv.OutWorkspace.PathName, strTemp)
    

    
    
    Set pCellAreaSqKMRaster = Nothing

    'STEP 5b: ----------------------------------------------------------------------------------
    'Square Feet
        pMapAlgebraOp.BindRaster pCellAreaSMRaster, "cellarea_sm"
    
        strExpression = "[cellarea_sm] * 10.76"
    
        Set pCellAreaSFRaster = pMapAlgebraOp.Execute(strExpression)
    
        pMapAlgebraOp.UnbindRaster "cellarea_sm"
    'END STEP 5b--------------------------------------------------------------------------------
    
    
    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pCellAreaSFRaster, pEnv.OutWorkspace.PathName, strTemp)
    


    'STEP 5c: ----------------------------------------------------------------------------------
    'Square Inches
        pMapAlgebraOp.BindRaster pCellAreaSFRaster, "cellarea_sf"
    
        strExpression = "[cellarea_sf] * 144"
    
        Set pCellAreaSIRaster = pMapAlgebraOp.Execute(strExpression)
    
        pMapAlgebraOp.UnbindRaster "cellarea_sf"
    'END STEP 5c--------------------------------------------------------------------------------

    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pCellAreaSIRaster, pEnv.OutWorkspace.PathName, strTemp)
    

    
    
    Set pCellAreaSFRaster = Nothing

    'STEP 5d: ----------------------------------------------------------------------------------
    'Acres
        pMapAlgebraOp.BindRaster pCellAreaSMRaster, "cell_sm"
    
        strExpression = "[cell_sm] * 0.000247104369"
    
        Set pCellAreaAcreRaster = pMapAlgebraOp.Execute(strExpression)
    
        pMapAlgebraOp.UnbindRaster "cell_sm"
    'END STEP 5d--------------------------------------------------------------------------------


    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pCellAreaAcreRaster, pEnv.OutWorkspace.PathName, strTemp)
    




    Set pCellAreaSMRaster = Nothing

    'STEP 5e: ----------------------------------------------------------------------------------
    'Cubic Inches
        With pMapAlgebraOp
            .BindRaster pCellAreaSIRaster, "cell_si"
            .BindRaster pRunoffInRaster, "runoff_in"
        End With
    
        strExpression = "([cell_si] * [runoff_in])"
    
        Set pCellAreaCIRaster = pMapAlgebraOp.Execute(strExpression)
    
        With pMapAlgebraOp
            .UnbindRaster "cell_si"
            .UnbindRaster "runoff_in"
        End With
    'END STEP 5e -------------------------------------------------------------------------------

    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pCellAreaCIRaster, pEnv.OutWorkspace.PathName, strTemp)
    

    
    
    Set pCellAreaSIRaster = Nothing
    
    'STEP 5f: -----------------------------------------------------------------------------------
    'Cubic feet
        pMapAlgebraOp.BindRaster pCellAreaCIRaster, "runoff_ci"
    
        strExpression = "([runoff_ci] * 0.0005787)"
    
        Set pCellAreaCFRaster = pMapAlgebraOp.Execute(strExpression)
        Set g_pRunoffCFRaster = pCellAreaCFRaster
    
        pMapAlgebraOp.UnbindRaster "runoff_ci"
    
    'END STEP 5f: --------------------------------------------------------------------------------
        
    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pCellAreaCFRaster, pEnv.OutWorkspace.PathName, strTemp)
    



    'STEP 5g: ------------------------------------------------------------------------------------
    'Runoff acre-feet
        pMapAlgebraOp.BindRaster pCellAreaCFRaster, "runoff_cf"
    
        strExpression = "([runoff_cf] * 0.000022957)"
    
        Set pCellAreaAFRaster = pMapAlgebraOp.Execute(strExpression)
        Set g_pRunoffAFRaster = pCellAreaAFRaster
    
        pMapAlgebraOp.UnbindRaster "runoff_cf"
    'END STEP 5f: --------------------------------------------------------------------------------
    Else
        GoTo ProgCancel
    End If
    
    
    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pCellAreaAFRaster, pEnv.OutWorkspace.PathName, strTemp)
    

    
    modProgDialog.ProgDialog "Converting to correct units...", strTitle, 0, 10, 8, frmPrj.m_App.hwnd
    
    If modProgDialog.g_boolCancel Then
    'STEP 6: -------------------------------------------------------------------------------------
    'Convert cubic inches to liters
        If pCellAreaCIRaster Is Nothing Then
            MsgBox "nothing"
        End If
        
        pMapAlgebraOp.BindRaster pCellAreaCIRaster, "run_ci"
    
        strExpression = "[run_ci] * 0.016387064"
    
        Set pMetRunoffRaster = pMapAlgebraOp.Execute(strExpression)
        Set g_pMetRunoffRaster = pMetRunoffRaster
        
        pMapAlgebraOp.UnbindRaster "run_ci"
    'END STEP 6: ---------------------------------------------------------------------------------
    Else
        GoTo ProgCancel
    End If
    
    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pMetRunoffRaster, pEnv.OutWorkspace.PathName, strTemp)
    

    
    'STEP 6a: -------------------------------------------------------------------------------------
    'Eliminate nulls: Added 1/28/07
    pMapAlgebraOp.BindRaster pMetRunoffRaster, "runoffgrid"
        
    strExpression = "Con(IsNull([runoffgrid]),0,[runoffgrid])"
    
    Set pMetRunoffNoNullRaster = pMapAlgebraOp.Execute(strExpression)
    
    Set g_pMetRunoffRaster = pMetRunoffNoNullRaster
    
    pMapAlgebraOp.UnbindRaster "runoffgrid"
    
    
    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pMetRunoffNoNullRaster, pEnv.OutWorkspace.PathName, strTemp)
    

    If g_booLocalEffects Then
        modProgDialog.ProgDialog "Creating data layer for local effects...", strTitle, 0, 10, 10, frmPrj.m_App.hwnd
        If modProgDialog.g_boolCancel Then
                       
            'STEP 12: Local Effects -------------------------------------------------
            strOutAccum = modUtil.GetUniqueName("locaccum", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
            'Added 7/23/04 to account for clip by selected polys functionality
            If g_booSelectedPolys Then
                Dim pClipLocAccumRaster As IRaster
                Set pClipLocAccumRaster = modUtil.ClipBySelectedPoly(pMetRunoffNoNullRaster, g_pSelectedPolyClip, pEnv)
                Set pPermAccumLocRunoffRaster = modUtil.ReturnPermanentRaster(pClipLocAccumRaster, pEnv.OutWorkspace.PathName, strOutAccum)
                Set pClipLocAccumRaster = Nothing
            Else
                Set pPermAccumLocRunoffRaster = modUtil.ReturnPermanentRaster(pMetRunoffNoNullRaster, pEnv.OutWorkspace.PathName, strOutAccum)
            End If
            
            Set pLocRunoffRasterLayer = modUtil.ReturnRasterLayer(frmPrj.m_App, pPermAccumLocRunoffRaster, "Runoff Local Effects (L)")
            Set pLocRunoffRasterLayer.Renderer = modUtil.ReturnRasterStretchColorRampRender(pLocRunoffRasterLayer, "Blue")
            pLocRunoffRasterLayer.Visible = False
            
            g_dicMetadata.Add pLocRunoffRasterLayer.Name, m_strRunoffMetadata
            
            g_pGroupLayer.Add pLocRunoffRasterLayer
            
            RunoffCalculation = True
            modProgDialog.KillDialog
            Exit Function
           
        End If
    End If
        
    modProgDialog.ProgDialog "Creating flow accumulation...", strTitle, 0, 10, 9, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
    'STEP 7: ------------------------------------------------------------------------------------
    'Derive Accumulated Runoff
    
        With pMapAlgebraOp
            .BindRaster g_pFlowDirRaster, "flowdir"
            .BindRaster pMetRunoffNoNullRaster, "met_run"
        End With
    
        strExpression = "FlowAccumulation([flowdir], [met_run], FLOAT)"
    
        Set pAccumRunoffRaster1 = pMapAlgebraOp.Execute(strExpression)
            
        With pMapAlgebraOp
            .UnbindRaster "flowdir"
            .UnbindRaster "met_run"
        End With
    'END STEP 7: ----------------------------------------------------------------------------------
    Else
        GoTo ProgCancel
    End If
    
    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pAccumRunoffRaster1, pEnv.OutWorkspace.PathName, strTemp)
        
    
    modProgDialog.ProgDialog "Creating flow accumulation...", strTitle, 0, 10, 9, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
    'STEP 7: ------------------------------------------------------------------------------------
    'Derive the real Accumulated Runoff
    
        With pMapAlgebraOp
            .BindRaster pAccumRunoffRaster1, "accum"
            .BindRaster pMetRunoffNoNullRaster, "met_run"
        End With
    
        strExpression = "[accum] + [met_run]"
    
        Set pAccumRunoffRaster = pMapAlgebraOp.Execute(strExpression)
            
        With pMapAlgebraOp
            .UnbindRaster "accum"
            .UnbindRaster "met_run"
        End With
    'END STEP 7: ----------------------------------------------------------------------------------
    Else
        GoTo ProgCancel
    End If
    
    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("runstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pAccumRunoffRaster, pEnv.OutWorkspace.PathName, strTemp)
    
    
    'Add this then map as our runoff grid
    modProgDialog.ProgDialog "Creating Runoff Layer...", strTitle, 0, 10, 10, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
        'Get a unique name for accumulation GRID
        strOutAccum = modUtil.GetUniqueName("runoff", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        
        'Clip to selected polys if chosen
        If g_booSelectedPolys Then
            Dim pClipAccumRaster As IRaster
            Set pClipAccumRaster = modUtil.ClipBySelectedPoly(pAccumRunoffRaster, g_pSelectedPolyClip, pEnv)
            Set pPermAccumRunoffRaster = modUtil.ReturnPermanentRaster(pClipAccumRaster, pEnv.OutWorkspace.PathName, strOutAccum)
            Set pClipAccumRaster = Nothing
        Else
            Set pPermAccumRunoffRaster = modUtil.ReturnPermanentRaster(pAccumRunoffRaster, pEnv.OutWorkspace.PathName, strOutAccum)
        End If
        
        'Add Completed raster to the g_pGroupLayer
        Set pRunoffRasterLayer = modUtil.ReturnRasterLayer(frmPrj.m_App, pPermAccumRunoffRaster, "Accumulated Runoff (L)")
        Set pRunoffRasterLayer.Renderer = ReturnRasterStretchColorRampRender(pRunoffRasterLayer, "Blue")
        pRunoffRasterLayer.Visible = False
        g_pGroupLayer.Add pRunoffRasterLayer
        
        g_dicMetadata.Add pRunoffRasterLayer.Name, m_strRunoffMetadata
            
        'Global Runoff
        Set g_pRunoffRaster = pPermAccumRunoffRaster
    Else
        GoTo ProgCancel
    End If
    
    RunoffCalculation = True
    
    modProgDialog.KillDialog
    
    'Cleanup
    Set pCellAreaSMRaster = Nothing
    Set pCellAreaAcreRaster = Nothing
    Set pCellAreaCIRaster = Nothing
    Set pAccumRunoffRaster = Nothing
    Set pAccumRunoffRaster1 = Nothing
    
Exit Function

ErrHandler:
    If Err.Number = -2147217297 Then 'User cancelled operation
        modProgDialog.g_boolCancel = False
        RunoffCalculation = False
    ElseIf Err.Number = -2147467259 Then
        MsgBox "ArcMap has reached the maximum number of GRIDs allowed in memory.  " & _
            "Please exit N-SPECT and restart ArcMap.", vbInformation, "Maximum GRID Number Encountered"
        modProgDialog.g_boolCancel = False
        modProgDialog.KillDialog
        RunoffCalculation = False
    Else
        MsgBox "Error: " & Err.Number & " on RunoffCalculation: " & strExpression
        modProgDialog.g_boolCancel = False
        modProgDialog.KillDialog
        RunoffCalculation = False
    End If

ProgCancel:
    
    
End Function

Private Function CreateMetadata(rsLandClass As ADODB.Recordset, strLandClass As String, booLocal As Boolean) As String
  On Error GoTo ErrorHandler

    
    Dim i As Integer
    Dim strHeader As String
    Dim strLCHeader As String
    Dim strCoeffValues As String
    
    If booLocal Then
        strHeader = vbTab & "Input Datasets:" & vbNewLine & _
                vbTab & vbTab & "Hydrologic soils grid: " & g_clsXMLPrjFile.strSoilsHydFileName & vbNewLine & _
                vbTab & vbTab & "Landcover grid: " & g_clsXMLPrjFile.strLCGridFileName & vbNewLine & _
                vbTab & vbTab & "Precipitation grid: " & g_strPrecipFileName & vbNewLine
    Else
        strHeader = vbTab & "Input Datasets:" & vbNewLine & _
                vbTab & vbTab & "Hydrologic soils grid: " & g_clsXMLPrjFile.strSoilsHydFileName & vbNewLine & _
                vbTab & vbTab & "Landcover grid: " & g_clsXMLPrjFile.strLCGridFileName & vbNewLine & _
                vbTab & vbTab & "Precipitation grid: " & g_strPrecipFileName & vbNewLine & _
                vbTab & vbTab & "Flow direction grid: " & g_strFlowDirFilename & vbNewLine
    End If
    
    strLCHeader = vbTab & "Landcover parameters: " & vbNewLine & _
        vbTab & vbTab & "Landcover type name: " & strLandClass & vbNewLine & _
        vbTab & vbTab & "Coefficients: " & vbNewLine

    rsLandClass.MoveFirst
    
    For i = 1 To rsLandClass.RecordCount
        If i = 1 Then
            strCoeffValues = vbTab & vbTab & vbTab & rsLandClass!Name2 & ":" & vbNewLine & _
                vbTab & vbTab & vbTab & vbTab & "CN-A: " & CStr(rsLandClass![CN-A]) & vbNewLine & _
                vbTab & vbTab & vbTab & vbTab & "CN-B: " & CStr(rsLandClass![CN-B]) & vbNewLine & _
                vbTab & vbTab & vbTab & vbTab & "CN-C: " & CStr(rsLandClass![CN-C]) & vbNewLine & _
                vbTab & vbTab & vbTab & vbTab & "CN-D: " & CStr(rsLandClass![CN-D]) & vbNewLine & _
                vbTab & vbTab & vbTab & vbTab & "Cover factor: " & rsLandClass!CoverFactor & vbNewLine
            
        Else
            strCoeffValues = strCoeffValues & vbTab & vbTab & vbTab & rsLandClass!Name2 & ":" & vbNewLine & _
                vbTab & vbTab & vbTab & vbTab & "CN-A: " & CStr(rsLandClass![CN-A]) & vbNewLine & _
                vbTab & vbTab & vbTab & vbTab & "CN-B: " & CStr(rsLandClass![CN-B]) & vbNewLine & _
                vbTab & vbTab & vbTab & vbTab & "CN-C: " & CStr(rsLandClass![CN-C]) & vbNewLine & _
                vbTab & vbTab & vbTab & vbTab & "CN-D: " & CStr(rsLandClass![CN-D]) & vbNewLine & _
                vbTab & vbTab & vbTab & vbTab & "Cover factor: " & rsLandClass!CoverFactor & vbNewLine
        End If
        rsLandClass.MoveNext
    Next i
    
    CreateMetadata = strHeader & strLCHeader & strCoeffValues
    g_strLandCoverParameters = strLCHeader & strCoeffValues
    


  Exit Function
ErrorHandler:
  HandleError False, "CreateMetadata " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Function
