Attribute VB_Name = "modMUSLE"
' *************************************************************************************
' *  Perot Systems Government Services
' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
' *  modRUSLE
' *************************************************************************************
' *  Description:  Modified Universal Soil Loss Equation
' *
' *  Subs/Functions: Defined in each.
' *
' *  Called By:  frmPrj
' *************************************************************************************
Option Explicit
Private m_strCFactorConStatement As String          'C Factor Con Statement
Private m_strPondConStatement As String             'Pond Factor con statement
Private m_dblMUSLEVal As Double                     'Customizable MUSLE value in Equation
Private m_dblMUSLEExp As Double                     'Customizable musle exponent in equation

Private m_strMusleMetadata As String                'MUSLE metadata string

' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "modMUSLE.bas"


Public Function MUSLESetup(strSoilsDefName As String, strKfactorFileName As String, strLandClass As String) As Boolean

'Sub takes incoming parameters from the project file and then parses them out
'strSoilsDefName: Name of the Soils Definition being used
'strKFactorFileName: K Factor FileName
'strLandClass: Name of the Landclass we're using
    
    'RecordSet to get the coverfactor
    Dim rsCoverFactor As New ADODB.Recordset
    Dim strTempLCType As String                     'Our potential holder for a temp landtype
    
    Dim rsSoilsDef As New ADODB.Recordset
    Dim strSoilsDef As String
    
    'Open Strings
    Dim strCovFactor As String
    Dim strError As String
    
    'STEP 1: Get the MUSLE Values
    strSoilsDef = "SELECT * FROM SOILS WHERE NAME LIKE '" & strSoilsDefName & "'"
        
    rsSoilsDef.Open strSoilsDef, modUtil.g_ADOConn, adOpenForwardOnly, adLockReadOnly
    
    m_dblMUSLEVal = rsSoilsDef!MUSLEVal
    m_dblMUSLEExp = rsSoilsDef!MUSLEExp
    
        
    'STEP 2: Set the K Factor Raster
    If modUtil.RasterExists(strKfactorFileName) Then
        Set g_KFactorRaster = modUtil.ReturnRaster(strKfactorFileName)
    Else
        strError = "K Factor Raster Does Not Exist: " & strKfactorFileName
    End If
    'END STEP 1: -----------------------------------------------------------------------------------
       
    If Len(g_DictTempNames.Item(strLandClass)) > 0 Then
        strTempLCType = g_DictTempNames.Item(strLandClass)
    Else
        strTempLCType = strLandClass
    End If
       
    'Get the landclasses of type strLandClass
    strCovFactor = "SELECT LCTYPE.LCTYPEID, LCCLASS.NAME, LCCLASS.VALUE, LCCLASS.COVERFACTOR, LCCLASS.W_WL FROM " & _
        "LCTYPE INNER JOIN LCCLASS ON LCTYPE.LCTYPEID = LCCLASS.LCTYPEID " & _
        "WHERE LCTYPE.NAME LIKE '" & strTempLCType & "' ORDER BY LCCLASS.VALUE"
        
    m_strMusleMetadata = CreateMetadata(g_booLocalEffects) ', rsCoverFactor)
        
    rsCoverFactor.Open strCovFactor, g_ADOConn, adOpenKeyset, adLockOptimistic
    
    If Len(strError) > 0 Then
        MsgBox strError
        Exit Function
    End If
        
    'Get the con statement for the cover factor calculation
    m_strCFactorConStatement = ConstructConStatment(rsCoverFactor, g_LandCoverRaster)
    m_strPondConStatement = ConstructPondConStatement(rsCoverFactor, g_LandCoverRaster)
    
    'Calc rusle using the con
    If CalcMUSLE(m_strCFactorConStatement, m_strPondConStatement) Then
        MUSLESetup = True
    Else
        MUSLESetup = False
    End If
    
    'Clean that bubba up
    rsCoverFactor.Close
    Set rsCoverFactor = Nothing
    
    rsSoilsDef.Close
    Set rsSoilsDef = Nothing
    
    
    
End Function

Private Function CalcMUSLE(strConStatement As String, strConPondStatement As String) As Boolean
'Incoming strings: strConStatment: the monster con statement
'strConPondstatement: the con for the pond stuff
'Calculates the MUSLE erosion model

    Dim pWeightRaster As IRaster                'STEP 1: Weight Raster
    Dim pWSLengthRaster As IRaster              'STEP 2: Watershed length
    Dim pWSLengthUnitsRaster As IRaster         'STEP 3: Units conversion
    Dim pSlopePRRaster As IRaster               'STEP 4: Average Slope
    Dim pSlopeModRaster As IRaster              'STEP 4a: Mod Slope
    Dim pTemp1LagRaster As IRaster              'STEP 5a: Lag temp1
    Dim pTemp2LagRaster As IRaster              'STEP 5b: Lag temp2
    Dim pTemp3LagRaster As IRaster              'STEP 5c: Lag temp3
    Dim pTemp4LagRaster As IRaster              'STEP 5d: Lag temp4
    Dim pLagRaster As IRaster                   'STEP 5e: Lag
    Dim pTOCRaster As IRaster                   'STEP 6a: TOC
    Dim pTOCTempRaster As IRaster               'STEP 6b: TOC temp
    Dim pModTOCRaster As IRaster                'STEP 6c: Mod TOC
    Dim pAbPrecipRaster As IRaster              'STEP 7: Abstraction-Precip Ratio
    Dim pLogTOCRaster As IRaster                'STEP 8a: Calc unit peak discharge
    Dim pTempLogTOCRaster As IRaster            'STEP 8b:
    Dim pCZeroRaster As IRaster                 'STEP 9: CZero
    Dim pConeRaster As IRaster                  'STEP 10: Cone
    Dim pCTwoRaster As IRaster                  'STEP 11: CTwo
    Dim pLogQuRaster As IRaster                 'STEP 12a: Who knows
    Dim pQuRaster As IRaster                    'STep 12b: Who Knows(b)
    Dim pPondFactorRaster As IRaster            'STEP 13: Pond Factor
    Dim pQPRaster As IRaster                    'STEP 14: QP factor
    Dim pCFactorRaster As IRaster               'STEP 15: Yee old CFactor Raster
    Dim pSYTempRaster As IRaster                'Step 16a Temp yield
    Dim pSYRaster As IRaster                    'Step 16b Yield
    Dim pHISYTempRaster As IRaster              'STEP 16c: HI Specific temp yield
    Dim pHISYRaster As IRaster                  'STEP 16d: HI Specific yield
    'Dim pSYMGRaster As IRaster                  'STEP 17a: tons to milligrams
    Dim pHISYMGRaster As IRaster                'STEP 17b: tons to milligrams, HI Specific
    Dim pPermMUSLERaster As IRaster             'STEP 17c: Local Effects permanent raster
    Dim pTempFlowDir1Raster As IRaster          'flowDir Temp Raster
    'Dim pTempFlowDir2Raster As IRaster          'FlowDir temp raster
    Dim pTempflowDir3raster As IRaster          'FlowDir temp raster
    Dim pLiterRunoffRaster As IRaster           'STEP 18: Runoff to liters
    'Dim pAccSedRaster As IRaster                'STEP 19: Acc sediment
    Dim pAccSedHIRaster  As IRaster             'STEP 19HI: acc sediment
    Dim pFlowAccumOp As IHydrologyOp            'STEP 19: IHydroOp and friends
    Dim pFlowDirRDS As IGeoDataset
    Dim pLiterRDS As IGeoDataset
    Dim pOutRDS As IGeoDataset
    Dim pAccRunLiterRaster As IRaster           'STEP 19a: Acc runoff liter
    Dim pFlowAccumOp1 As IHydrologyOp           'STEP 19HI: IHydroOp
    Dim pFlowDirRDS1 As IGeoDataset
    Dim pHISYMGRDS As IGeoDataset
    Dim pOutRDS1 As IGeoDataset
    Dim pTotSedMassRaster As IRaster            'STEP 20: Tot sed mass
    Dim pTotSedMassHIRaster As IRaster          'STEP 20HI: tot sed mass HI
    Dim pTotSedTempRaster As IRaster            'STEP 21: First step
    Dim pPermTotSedConcHIraster As IRaster      'Permanent MUSLE
    Dim pSedConcRaster As IRaster               'Sed Concentration
    Dim pPermSedConcRaster As IRaster           'Permanent Sed Concentration
    Dim pSedConcRasterLayer As IRasterLayer     'Sed Concentration layer
    Dim pMUSLERasterLayer As IRasterLayer
    Dim pMUSLERasterLocLayer As IRasterLayer    'MUSLE Local Effects
    Dim strMUSLE As String
    
    
    Dim clsPrecip As New clsPrecipType
    
    Dim pEnv As IRasterAnalysisEnvironment      'Raster Environment

    'Create Map Algebra Operator
    Dim pMapAlgebraOp As IMapAlgebraOp          'Workhorse

    'String to hold calculations
    Dim strExpression As String
    Const strTitle As String = "Processing MUSLE Calculation..."
    
    'TEMP code
    Dim strTemp As String
    Dim pTempRaster As IRaster
    
    'Set the enviornment stuff
    Set pEnv = g_pSpatEnv
    Set pMapAlgebraOp = New RasterMapAlgebraOp
    Set pEnv = pMapAlgebraOp
    
On Error GoTo ErrHandler
  
    modProgDialog.ProgDialog "Computing length/distance...", strTitle, 0, 27, 1, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
        'STEP 1: ------------------------------------------------------------------------------------
        'Create weight grid that represents cell length/distance
        pMapAlgebraOp.BindRaster g_pFlowDirRaster, "flowdir"
        
        strExpression = "Con([flowdir] eq 2, 1.41421, " & _
                        "Con([flowdir] eq 8, 1.41421, " & _
                        "Con([flowdir] eq 32, 1.41421, " & _
                        "Con([flowdir] eq 128, 1.41421, 1.0))))"
        
        Set pWeightRaster = pMapAlgebraOp.Execute(strExpression)
        
        pMapAlgebraOp.UnbindRaster "flowdir"
        'END STEP 2: -------------------------------------------------------------------------------
    End If
    
    'TEMP Code added 8/29/06 to help w/ Hawaii's checking of the RUSLE equation
    'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pWeightRaster, pEnv.OutWorkspace.PathName, strTemp)
    
    modProgDialog.ProgDialog "Calculating Watershed Length...", strTitle, 0, 27, 2, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
        'STEP 2: ------------------------------------------------------------------------------------
        'Calculate Watershed Length
        With pMapAlgebraOp
            .BindRaster g_pFlowDirRaster, "flowdir"
            .BindRaster pWeightRaster, "weight"
        End With
        
        strExpression = "flowlength([flowdir], [weight], upstream)"
        
        Set pWSLengthRaster = pMapAlgebraOp.Execute(strExpression)
        
        With pMapAlgebraOp
            .UnbindRaster "flowdir"
            .UnbindRaster "weight"
        End With
        'End STEP 2 ----------------------------------------------------------------------------------
    End If
    
    'TEMP Code added 8/29/06 to help w/ Hawaii's checking of the RUSLE equation
    'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pWSLengthRaster, pEnv.OutWorkspace.PathName, strTemp)
    
    'Notice throughout that we cleanup before the end...don't want to exceed the max number allowed rasters
    'in memory.
    Set pWeightRaster = Nothing
    
    modProgDialog.ProgDialog "Converting units...", strTitle, 0, 27, 3, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
        'STEP 3: --------------------------------------------------------------------------------------
        'Convert Metric Units
        pMapAlgebraOp.BindRaster pWSLengthRaster, "cell_wslength"
        
        strExpression = "([cell_wslength] * 3.28084)"
        
        Set pWSLengthUnitsRaster = pMapAlgebraOp.Execute(strExpression)
        
        pMapAlgebraOp.UnbindRaster "cell_wslength"
        'END STEP 3: -----------------------------------------------------------------------------------
    End If
    
    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pWSLengthUnitsRaster, pEnv.OutWorkspace.PathName, strTemp)
    
    Set pWSLengthRaster = Nothing
    
    modProgDialog.ProgDialog "Calculating Average Slope...", strTitle, 0, 27, 4, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
        'STEP 4a: ---------------------------------------------------------------------------------------
        'Calculate Average Slope
        pMapAlgebraOp.BindRaster g_pDEMRaster, "dem"
        
        strExpression = "slope([dem], percentrise)"
        
        Set pSlopePRRaster = pMapAlgebraOp.Execute(strExpression)
        
        pMapAlgebraOp.UnbindRaster "dem"
        'END STEP 4a ------------------------------------------------------------------------------------
    End If
    
    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pSlopePRRaster, pEnv.OutWorkspace.PathName, strTemp)
    
    modProgDialog.ProgDialog "Calculating Mod Slope...", strTitle, 0, 27, 5, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
        'STEP 4b: ---------------------------------------------------------------------------------------
        'Calculate modslope
        pMapAlgebraOp.BindRaster pSlopePRRaster, "slopepr"
        
        strExpression = "Con([slopepr] eq 0, 0.1, [slopepr])"
        
        Set pSlopeModRaster = pMapAlgebraOp.Execute(strExpression)
        
        pMapAlgebraOp.UnbindRaster "slopepr"
        'END STEP 4b: -----------------------------------------------------------------------------------
    End If
    
    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pSlopeModRaster, pEnv.OutWorkspace.PathName, strTemp)
    
    Set pSlopePRRaster = Nothing
    
    
    modProgDialog.ProgDialog "Calculating Lag...", strTitle, 0, 27, 6, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
        'STEP 5a: ---------------------------------------------------------------------------------------
        'Calculate Lag
        pMapAlgebraOp.BindRaster pWSLengthUnitsRaster, "cell_wslngft"
        
        strExpression = "Pow([cell_wslngft], 0.8)"
        
        Set pTemp1LagRaster = pMapAlgebraOp.Execute(strExpression)
        
        pMapAlgebraOp.UnbindRaster "cell_wslngft"
        'END STEP 5a: ----------------------------------------------------------------------------------
    End If
    
    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pTemp1LagRaster, pEnv.OutWorkspace.PathName, strTemp)
    
    Set pWSLengthUnitsRaster = Nothing
    
    modProgDialog.ProgDialog "Checking SCS GRID...", strTitle, 0, 27, 7, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
        'STEP 5b: ---------------------------------------------------------------------------------------
        pMapAlgebraOp.BindRaster g_pSCS100Raster, "scsgrid100"
        
        strExpression = "(1000 / [scsgrid100]) - 9"
        
        Set pTemp2LagRaster = pMapAlgebraOp.Execute(strExpression)
        
        pMapAlgebraOp.UnbindRaster "scsgrid100"
        'END STEP 9b: ----------------------------------------------------------------------------------
    End If
    
    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pTemp2LagRaster, pEnv.OutWorkspace.PathName, strTemp)
    
    
    modProgDialog.ProgDialog "Multiplying GRIDs...", strTitle, 0, 27, 8, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
        'STEP 5c: --------------------------------------------------------------------------------------
        pMapAlgebraOp.BindRaster pTemp2LagRaster, "temp4"
        
        strExpression = "Pow([temp4], 0.7)"
        
        Set pTemp3LagRaster = pMapAlgebraOp.Execute(strExpression)
        
        pMapAlgebraOp.UnbindRaster "temp4"
        'END STEP 5c: ----------------------------------------------------------------------------------
    End If
    
    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pTemp3LagRaster, pEnv.OutWorkspace.PathName, strTemp)
    
    
    Set pTemp2LagRaster = Nothing
    
    modProgDialog.ProgDialog "Pow([modslope], 0.5...", strTitle, 0, 27, 9, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
        'STEP 5d: --------------------------------------------------------------------------------------
        pMapAlgebraOp.BindRaster pSlopeModRaster, "modslope"
        
        strExpression = "Pow([modslope], 0.5)"
        
        Set pTemp4LagRaster = pMapAlgebraOp.Execute(strExpression)
        
        pMapAlgebraOp.UnbindRaster "modslope"
        'END STEP 5d: ----------------------------------------------------------------------------------
    End If
    
    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pTemp4LagRaster, pEnv.OutWorkspace.PathName, strTemp)
    
    Set pSlopeModRaster = Nothing

    modProgDialog.ProgDialog "Lag calculation...", strTitle, 0, 30, 27, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
        'STEP 5e: --------------------------------------------------------------------------------------
        'Finally, the lag calculation
        With pMapAlgebraOp
            .BindRaster pTemp1LagRaster, "temp3"
            .BindRaster pTemp3LagRaster, "temp5"
            .BindRaster pTemp4LagRaster, "temp6"
        End With
        
        strExpression = "([temp3] * [temp5]) / (1900 * [temp6])"
        
        Set pLagRaster = pMapAlgebraOp.Execute(strExpression)
        
        With pMapAlgebraOp
            .UnbindRaster "temp3"
            .UnbindRaster "temp5"
            .UnbindRaster "temp6"
        End With
        'STEP 5e: ---------------------------------------------------------------------------------------
    End If
    
    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pLagRaster, pEnv.OutWorkspace.PathName, strTemp)
    
    Set pTemp1LagRaster = Nothing
    Set pTemp3LagRaster = Nothing
    Set pTemp4LagRaster = Nothing
    
    
    modProgDialog.ProgDialog "Calculating time of concentration...", strTitle, 0, 27, 11, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
    
        'STEP 6a: ---------------------------------------------------------------------------------------
        'Calculate the time of concentration
        pMapAlgebraOp.BindRaster pLagRaster, "lag"
        
        strExpression = "[lag] / 0.6"
        
        Set pTOCRaster = pMapAlgebraOp.Execute(strExpression)
        
        pMapAlgebraOp.UnbindRaster "lag"
        'END STEP 6a: -----------------------------------------------------------------------------------
        
        'TEMP Code added 2008 to help w/ debugging
        'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        'Set pTempRaster = modUtil.ReturnPermanentRaster(pTOCRaster, pEnv.OutWorkspace.PathName, strTemp)
        
        Set pLagRaster = Nothing
        
        'STEP 6b: --------------------------------------------------------------------------------------
        pMapAlgebraOp.BindRaster pTOCRaster, "toc"
        
        strExpression = "Con([toc] lt 0.1, 0.1, [toc])"
        
        Set pTOCTempRaster = pMapAlgebraOp.Execute(strExpression)
        
        pMapAlgebraOp.UnbindRaster "toc"
        'END STEP 6b: ----------------------------------------------------------------------------------
        
        'TEMP Code added 2008 to help w/ debugging
        'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        'Set pTempRaster = modUtil.ReturnPermanentRaster(pTOCTempRaster, pEnv.OutWorkspace.PathName, strTemp)
        
        Set pTOCRaster = Nothing
        
        'STEP 6c: --------------------------------------------------------------------------------------
        pMapAlgebraOp.BindRaster pTOCTempRaster, "temp7"
        
        strExpression = "Con([temp7] gt 10, 10, [temp7])"
        
        Set pModTOCRaster = pMapAlgebraOp.Execute(strExpression)
        
        pMapAlgebraOp.UnbindRaster "temp7"
        'END STEP 6c: ----------------------------------------------------------------------------------
        
        'TEMP Code added 2008 to help w/ debugging
        'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        'Set pTempRaster = modUtil.ReturnPermanentRaster(pModTOCRaster, pEnv.OutWorkspace.PathName, strTemp)
        
    End If
    
    Set pTOCTempRaster = Nothing
    
    modProgDialog.ProgDialog "Abstraction Precipitation Ratio...", strTitle, 0, 27, 12, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
        'STEP 7: ---------------------------------------------------------------------------------------
        'Find the Abstraction Precipitation Ratio
        With pMapAlgebraOp
            .BindRaster g_pAbstractRaster, "abstract"
            .BindRaster g_pPrecipRaster, "rain"
        End With
        
        strExpression = "[abstract] / [rain]"
        
        Set pAbPrecipRaster = pMapAlgebraOp.Execute(strExpression)
        
        With pMapAlgebraOp
            .UnbindRaster "abstract"
            .UnbindRaster "rain"
        End With
        'END STEP 7: ----------------------------------------------------------------------------------
    End If
    
    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pAbPrecipRaster, pEnv.OutWorkspace.PathName, strTemp)
    
    modProgDialog.ProgDialog "Calculating the peak unit discharge...", strTitle, 0, 27, 13, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
        'STEP 8a: ----------------------------------------------------------------------------------------
        'Calculate the unit peak discharge
        pMapAlgebraOp.BindRaster pModTOCRaster, "modtoc"
        
        strExpression = "log10([modtoc])"
        
        Set pLogTOCRaster = pMapAlgebraOp.Execute(strExpression)
        
        pMapAlgebraOp.UnbindRaster "modtoc"
        'END STEP 8a: -----------------------------------------------------------------------------------
        
        'TEMP Code added 2008 to help w/ debugging
        'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        'Set pTempRaster = modUtil.ReturnPermanentRaster(pLogTOCRaster, pEnv.OutWorkspace.PathName, strTemp)
        
        'STEP 8b: ---------------------------------------------------------------------------------------
        '2nd part of it
        pMapAlgebraOp.BindRaster pLogTOCRaster, "logtoc"
        
        strExpression = "Pow([logtoc], 2)"
        
        Set pTempLogTOCRaster = pMapAlgebraOp.Execute(strExpression)
        
        pMapAlgebraOp.UnbindRaster "logtoc"
        'END STEP 8b: -----------------------------------------------------------------------------------
        
        'TEMP Code added 2008 to help w/ debugging
        'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        'Set pTempRaster = modUtil.ReturnPermanentRaster(pTempLogTOCRaster, pEnv.OutWorkspace.PathName, strTemp)
        
    End If
    
    Set pModTOCRaster = Nothing
    
    modProgDialog.ProgDialog "Creating C-Zero GRID...", strTitle, 0, 27, 14, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
        'STEP 9: ----------------------------------------------------------------------------------------
        'CZERO GRID
        pMapAlgebraOp.BindRaster pAbPrecipRaster, "ip"
        
        'Call to clsPrecipType(called clsprecip here init above) to get string based on Precip Type
        'g_intPrecipType
        strExpression = clsPrecip.CZero(g_intPrecipType)
        
        Set pCZeroRaster = pMapAlgebraOp.Execute(strExpression)
        
        pMapAlgebraOp.UnbindRaster "ip"
        'END STEP 9: ------------------------------------------------------------------------------------
    End If
    
    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pCZeroRaster, pEnv.OutWorkspace.PathName, strTemp)
        
    
    modProgDialog.ProgDialog "Creating Cone GRID...", strTitle, 0, 27, 15, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
        'STEP 10: ---------------------------------------------------------------------------------------
        'CONE grid
        pMapAlgebraOp.BindRaster pAbPrecipRaster, "ip"
        
        strExpression = clsPrecip.Cone(g_intPrecipType)
        
        Set pConeRaster = pMapAlgebraOp.Execute(strExpression)
        
        pMapAlgebraOp.UnbindRaster "ip"
        'END STEP 10 ------------------------------------------------------------------------------------
    End If
    
    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pConeRaster, pEnv.OutWorkspace.PathName, strTemp)
        
    
    
    modProgDialog.ProgDialog "Creating C2 GRID...", strTitle, 0, 27, 16, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
        'STEP 11: ---------------------------------------------------------------------------------------
        'CTwo GRID
        pMapAlgebraOp.BindRaster pAbPrecipRaster, "ip"
        
        strExpression = clsPrecip.CTwo(g_intPrecipType)
                
        Set pCTwoRaster = pMapAlgebraOp.Execute(strExpression)
        
        pMapAlgebraOp.UnbindRaster "ip"
        'END STEP 11: ------------------------------------------------------------------------------------
    End If
    
    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pCTwoRaster, pEnv.OutWorkspace.PathName, strTemp)
        
    
            
    Set pAbPrecipRaster = Nothing
            
    modProgDialog.ProgDialog "More math...", strTitle, 0, 27, 17, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
        'STEP 12a: ----------------------------------------------------------------------------------------
        'Logqu
        With pMapAlgebraOp
            .BindRaster pCZeroRaster, "czero"
            .BindRaster pConeRaster, "cone"
            .BindRaster pLogTOCRaster, "logtoc"
            .BindRaster pCTwoRaster, "ctwo"
            .BindRaster pTempLogTOCRaster, "temp8"
        End With
        
        strExpression = "[czero] + ([cone] * [logtoc]) + ([ctwo] * [temp8])"
        
        Set pLogQuRaster = pMapAlgebraOp.Execute(strExpression)
        
        'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pLogQuRaster, pEnv.OutWorkspace.PathName, strTemp)
        
    
        
        With pMapAlgebraOp
            .UnbindRaster "czero"
            .UnbindRaster "cone"
            .UnbindRaster "logtoc"
            .UnbindRaster "ctwo"
            .UnbindRaster "temp8"
        End With
        'END STEP 12a: -------------------------------------------------------------------------------------
        
        Set pLogTOCRaster = Nothing
        Set pCZeroRaster = Nothing
        Set pConeRaster = Nothing
        Set pCTwoRaster = Nothing
                
        'STEP 12b: -----------------------------------------------------------------------------------------
        pMapAlgebraOp.BindRaster pLogQuRaster, "logqu"
        
        strExpression = "Pow(10, [logqu])"
        
        Set pQuRaster = pMapAlgebraOp.Execute(strExpression)
        
        pMapAlgebraOp.UnbindRaster "logqu"
        'END STEP 12b: -------------------------------------------------------------------------------------
    End If
    
      'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pQuRaster, pEnv.OutWorkspace.PathName, strTemp)
    
    Set pTempLogTOCRaster = Nothing
    Set pLogQuRaster = Nothing
    
    modProgDialog.ProgDialog "Creating Pond Factor GRID...", strTitle, 0, 27, 16, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
        'STEP 13: ------------------------------------------------------------------------------------------
        'Create pond factor grid
        pMapAlgebraOp.BindRaster g_LandCoverRaster, "nu_lulc"
        
        strExpression = strConPondStatement
        
        Set pPondFactorRaster = pMapAlgebraOp.Execute(strExpression)
        
        pMapAlgebraOp.UnbindRaster "nu_lulc"
        'END STEP 13: ---------------------------------------------------------------------------------------
    End If
    
    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pPondFactorRaster, pEnv.OutWorkspace.PathName, strTemp)
    
    modProgDialog.ProgDialog "Calculating peak discharge; cubic feet per second...", strTitle, 0, 27, 17, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
        'STEP 14: -------------------------------------------------------------------------------------------
        'Calculate peak discharge: cubic feet per second
        With pMapAlgebraOp
            .BindRaster pQuRaster, "qu"
            .BindRaster g_pCellAreaSqMiRaster, "cellarea_sqmi"
            .BindRaster g_pRunoffInchRaster, "runoff_in"
            .BindRaster pPondFactorRaster, "pondfact"
        End With
        
        strExpression = "[qu] * [cellarea_sqmi] * [runoff_in] * [pondfact]"
        
        Set pQPRaster = pMapAlgebraOp.Execute(strExpression)
        
        With pMapAlgebraOp
            .UnbindRaster "qu"
            .UnbindRaster "cellarea_sqmi"
            .UnbindRaster "runoff_in"
            .UnbindRaster "pondfact"
        End With
        'END STEP 14: ----------------------------------------------------------------------------------------
    End If
    
    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pQPRaster, pEnv.OutWorkspace.PathName, strTemp)

    Set pPondFactorRaster = Nothing
    Set pQuRaster = Nothing
    
    modProgDialog.ProgDialog "Creating cover factor GRID...", strTitle, 0, 27, 18, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
        'STEP 15: --------------------------------------------------------------------------------------------
        'Cover Factor GRID
        pMapAlgebraOp.BindRaster g_LandCoverRaster, "nu_lulc"
        
        strExpression = strConStatement
        
        Set pCFactorRaster = pMapAlgebraOp.Execute(strExpression)
        
        pMapAlgebraOp.UnbindRaster "nu_lulc"
        'END STEP 15 -----------------------------------------------------------------------------------------
    End If
    
    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pCFactorRaster, pEnv.OutWorkspace.PathName, strTemp)
    
'    modProgDialog.ProgDialog "Temporary Sediment Yield...", strTitle, 0, 30, 19, frmPrj.m_App.hWnd
'    If modProgDialog.g_boolCancel Then
'        'STEP 16a: --------------------------------------------------------------------------------------------
'        'Temp Sediment Yield
'        With pMapAlgebraOp
'            .BindRaster g_pRunoffAFRaster, "runoff_af"
'            .BindRaster pQPRaster, "qp"
'        End With
'
'        strExpression = "Pow(([runoff_af] * [qp]), 0.56)"
'
'        Set pSYTempRaster = pMapAlgebraOp.Execute(strExpression)
'
'        With pMapAlgebraOp
'            .UnbindRaster "runoff_af"
'            .UnbindRaster "qp"
'        End With
'        'END STEP 16a: ------------------------------------------------------------------------------------------------------
'    End If
'
'    modProgDialog.ProgDialog "Sediment Yield...", strTitle, 0, 30, 20, frmPrj.m_App.hWnd
'    If modProgDialog.g_boolCancel Then
'        'STEP 16b: ---------------------------------------------------------------------------------------------
'        'Sediment Yield
'        With pMapAlgebraOp
'            .BindRaster pCFactorRaster, "cfactor"
'            .BindRaster g_KFactorRaster, "kfactor"
'            .BindRaster g_pLSRaster, "lsfactor"
'            .BindRaster pSYTempRaster, "temp9"
'        End With
'
'        strExpression = "95 * ([cfactor] * [kfactor] * [lsfactor] * [temp9])"
'
'        Set pSYRaster = pMapAlgebraOp.Execute(strExpression)
'
'        With pMapAlgebraOp
'            .UnbindRaster "cfactor"
'            .UnbindRaster "kfactor"
'            .UnbindRaster "lsfactor"
'            .UnbindRaster "temp9"
'        End With
'        'END STEP 16b: ----------------------------------------------------------------------------------------
'    End If
'
'        Set pSYTempRaster = Nothing
        
    modProgDialog.ProgDialog "Sediment Yield...", strTitle, 0, 27, 19, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
        'STEP 16 c: -------------------------------------------------------------------------------------------
        'Temp Sediment Yield: Note: MUSLE Exponent inserted to allow for global use, cuz other's will use
        With pMapAlgebraOp
            .BindRaster g_pRunoffAFRaster, "runoff_af"
            .BindRaster pQPRaster, "qp"
        End With
        
        strExpression = "Pow(([runoff_af] * [qp]), " & m_dblMUSLEExp & ")"
        
        Set pHISYTempRaster = pMapAlgebraOp.Execute(strExpression)
        
        With pMapAlgebraOp
            .UnbindRaster "runoff_af"
            .UnbindRaster "qp"
        End With
        'END STEP 16c: ------------------------------------------------------------------------------------------------------
    End If
    
    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pHISYTempRaster, pEnv.OutWorkspace.PathName, strTemp)
    
    Set pQPRaster = Nothing
        
    modProgDialog.ProgDialog "Sediment Yield...", strTitle, 0, 30, 27, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
        'STEP 16d: ---------------------------------------------------------------------------------------------
        'Sediment Yield: Note m_dblMUSLEVal now in for universal
        With pMapAlgebraOp
            .BindRaster pCFactorRaster, "cfactor"
            .BindRaster g_KFactorRaster, "kfactor"
            .BindRaster g_pLSRaster, "lsfactor"
            .BindRaster pHISYTempRaster, "temp9"
        End With
        
        strExpression = m_dblMUSLEVal & " * ([cfactor] * [kfactor] * [lsfactor] * [temp9])"
        
        Set pHISYRaster = pMapAlgebraOp.Execute(strExpression)
        
        With pMapAlgebraOp
            .UnbindRaster "cfactor"
            .UnbindRaster "kfactor"
            .UnbindRaster "lsfactor"
            .UnbindRaster "temp9"
        End With
        'END STEP 16b: ----------------------------------------------------------------------------------------
    End If
    
    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pHISYRaster, pEnv.OutWorkspace.PathName, strTemp)
        
    Set pCFactorRaster = Nothing
    Set pHISYTempRaster = Nothing
        
    modProgDialog.ProgDialog "Converting tons to milligrams...", strTitle, 0, 27, 21, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
        
        pMapAlgebraOp.BindRaster pHISYRaster, "sy"
                
        strExpression = "[sy] * 907.184740"
        
        Set pHISYMGRaster = pMapAlgebraOp.Execute(strExpression)
        
        pMapAlgebraOp.UnbindRaster "sy"
        'END STEP 17b: ----------------------------------------------------------------------------------------
    End If
    
    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pHISYMGRaster, pEnv.OutWorkspace.PathName, strTemp)
    
    Set pHISYRaster = Nothing
    
    If g_booLocalEffects Then
        
        modProgDialog.ProgDialog "Creating data layer for local effects...", strTitle, 0, 27, 27, frmPrj.m_App.hwnd
        If modProgDialog.g_boolCancel Then
                       
            strMUSLE = modUtil.GetUniqueName("locmusle", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
            'Added 7/23/04 to account for clip by selected polys functionality
            If g_booSelectedPolys Then
                Dim pClipMusleRaster As IRaster
                Set pClipMusleRaster = modUtil.ClipBySelectedPoly(pHISYMGRaster, g_pSelectedPolyClip, pEnv)
                Set pPermMUSLERaster = modUtil.ReturnPermanentRaster(pClipMusleRaster, pEnv.OutWorkspace.PathName, strMUSLE)
                Set pClipMusleRaster = Nothing
            Else
                Set pPermMUSLERaster = modUtil.ReturnPermanentRaster(pHISYMGRaster, pEnv.OutWorkspace.PathName, strMUSLE)
            End If
            
            Set pMUSLERasterLocLayer = modUtil.ReturnRasterLayer(frmPrj.m_App, pPermMUSLERaster, "MUSLE Local Effects (mg)")
            Set pMUSLERasterLocLayer.Renderer = modUtil.ReturnRasterStretchColorRampRender(pMUSLERasterLocLayer, "Brown")
            pMUSLERasterLocLayer.Visible = False
            
            'metadata time
            g_dicMetadata.Add pMUSLERasterLocLayer.Name, m_strMusleMetadata
            
            g_pGroupLayer.Add pMUSLERasterLocLayer
            
            CalcMUSLE = True
            modProgDialog.KillDialog
            Set pPermMUSLERaster = Nothing
            Exit Function
           
        End If
        
    End If
    
    
'    modProgDialog.ProgDialog "Converting runoff to liters...", strTitle, 0, 30, 23, frmPrj.m_App.hWnd
'    If modProgDialog.g_boolCancel Then
'
'        'STEP 18: ---------------------------------------------------------------------------------------------
'        'Convert runoff to liters
'        pMapAlgebraOp.BindRaster g_pRunoffCFRaster, "runoff_cf"
'
'        strExpression = "[runoff_cf] * 28.31685"
'
'        Set pLiterRunoffRaster = pMapAlgebraOp.Execute(strExpression)
'
'        pMapAlgebraOp.UnbindRaster "runoff_cf"
'        'END STEP 18: -----------------------------------------------------------------------------------------
'    End If
    
'    modProgDialog.ProgDialog "Calculating accumulated sediment...", strTitle, 0, 27, 22, frmPrj.m_App.hWnd
'    If modProgDialog.g_boolCancel Then
'        'STEP 19: ---------------------------------------------------------------------------------------------
'        'Find accumulated sediment & runoff
'
'        If g_pFlowDirRaster Is Nothing Then
'            MsgBox "Moose Knuckles"
'        Else
'            Set pTempFlowDir1Raster = modUtil.ReturnRaster(g_strFlowDirFilename)
'        End If
'
'        With pMapAlgebraOp
'            .BindRaster pTempFlowDir1Raster, "flowdir1"
'            .BindRaster pHISYMGRaster, "sy_mg"
'        End With
'
'        strExpression = "flowaccumulation([flowdir1], [sy_mg])"
'
'        Set pAccSedRaster = pMapAlgebraOp.Execute(strExpression)
'
'        With pMapAlgebraOp
'            .UnbindRaster "flowdir1"
'            .UnbindRaster "sy_mg"
'        End With
'
'    End If

    'Because of a very fun ESRI error, the MapAlgebraOp had to be jettisened in favor of  the IHydroOp, only
    'thing that seems to work.  Ergo, we do this...
'    modProgDialog.ProgDialog "Calculating the accumulated sediment...", strTitle, 0, 30, 25, frmPrj.m_App.hWnd
'    If modProgDialog.g_boolCancel Then
'        'STEP 19a: ---------------------------------------------------------------------------------------------
'        'Find accumulated runoff liters
'
'        Set pTempFlowDir2Raster = modUtil.ReturnRaster(g_strFlowDirFilename)
'        Set pFlowAccumOp = New RasterHydrologyOp
'
'        Set pFlowDirRDS = pTempFlowDir2Raster
'        Set pLiterRDS = pLiterRunoffRaster
'
'        Set pEnv = pFlowAccumOp
'        Set pOutRDS = pFlowAccumOp.FlowAccumulation(pFlowDirRDS, pLiterRDS)
'
'        Set pAccRunLiterRaster = pOutRDS
'        'END STEP 19a: -----------------------------------------------------------------------------------------
'    End If
    
'    Set pTempFlowDir2Raster = Nothing
    
    modProgDialog.ProgDialog "Calculating the accumulated sediment...", strTitle, 0, 27, 23, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
    
        Set pTempflowDir3raster = modUtil.ReturnRaster(g_strFlowDirFilename)
        Set pFlowAccumOp1 = New RasterHydrologyOp
        
        Set pFlowDirRDS1 = pTempflowDir3raster
        Set pHISYMGRDS = pHISYMGRaster
        
        Set pEnv = pFlowAccumOp1
        Set pOutRDS1 = pFlowAccumOp1.FlowAccumulation(pFlowDirRDS1, pHISYMGRDS)
        
        Set pAccSedHIRaster = pOutRDS1
        
    End If
    
    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pAccSedHIRaster, pEnv.OutWorkspace.PathName, strTemp)
    
    Set pTempflowDir3raster = Nothing
    
    modProgDialog.ProgDialog "Calculating Total Sediment Mass...", strTitle, 0, 27, 24, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
        'STEP 20: -----------------------------------------------------------------------------------------------
        'Total Sediment Mass
'        With pMapAlgebraOp
'            .BindRaster pHISYMGRaster, "sy_mg"
'            .BindRaster pAccSedRaster, "accsed"
'        End With
'
'        strExpression = "[sy_mg] + [accsed]"
'
'        Set pTotSedMassRaster = pMapAlgebraOp.Execute(strExpression)
'
'        With pMapAlgebraOp
'            .UnbindRaster "sy_mg"
'            .UnbindRaster "accsed"
'        End With
        'END STEP 20: -------------------------------------------------------------------------------------------
        
        'Set pSYMGRaster = Nothing
        'Set pAccSedRaster = Nothing
        
        'STEP 20HI: -----------------------------------------------------------------------------------------------
        'Total Sediment Mass
        With pMapAlgebraOp
            .BindRaster pHISYMGRaster, "sy_mg_HI"
            .BindRaster pAccSedHIRaster, "accsed_HI"
        End With
        
        'old / 10000000
        strExpression = "[sy_mg_HI] + [accsed_HI]"
        
        Set pTotSedMassHIRaster = pMapAlgebraOp.Execute(strExpression)
        
        With pMapAlgebraOp
            .UnbindRaster "sy_mg_HI"
            .UnbindRaster "accsed_HI"
        End With
        'END STEP 20HI: -------------------------------------------------------------------------------------------
    End If
    
    'TEMP Code added 2008 to help w/ debugging
    'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pTotSedMassHIRaster, pEnv.OutWorkspace.PathName, strTemp)
    
    Set pHISYMGRaster = Nothing
    Set pAccSedHIRaster = Nothing
    
    
    modProgDialog.ProgDialog "Adding Sediment Mass to Group Layer...", strTitle, 0, 27, 25, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
        'STEP 21: Created the Sediment Mass Raster layer and add to Group Layer -----------------------------------
        'Get a unique name for MUSLE and return the permanently made raster
        strMUSLE = modUtil.GetUniqueName("MUSLEmass", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        
        'Clip to selected polys if chosen
        If g_booSelectedPolys Then
            Dim pClipMusleMassRaster As IRaster
            Set pClipMusleMassRaster = modUtil.ClipBySelectedPoly(pTotSedMassHIRaster, g_pSelectedPolyClip, pEnv)
            Set pPermTotSedConcHIraster = modUtil.ReturnPermanentRaster(pClipMusleMassRaster, pEnv.OutWorkspace.PathName, strMUSLE)
            Set pClipMusleMassRaster = Nothing
        Else
            Set pPermTotSedConcHIraster = modUtil.ReturnPermanentRaster(pTotSedMassHIRaster, pEnv.OutWorkspace.PathName, strMUSLE)
        End If
        
        'Now create the MUSLE layer
        Set pMUSLERasterLayer = modUtil.ReturnRasterLayer(frmPrj.m_App, pPermTotSedConcHIraster, "MUSLE Sediment Mass (kg)")
        Set pMUSLERasterLayer.Renderer = modUtil.ReturnRasterStretchColorRampRender(pMUSLERasterLayer, "Brown")
        pMUSLERasterLayer.Visible = False
        
        'Metadata:
        g_dicMetadata.Add pMUSLERasterLayer.Name, m_strMusleMetadata
        
        'Add the MUSLE Layer to the final group layer
        g_pGroupLayer.Add pMUSLERasterLayer
        
        'end STEP 21: Created the Sediment Mass Raster layer and add to Group Layer -----------------------------------
        
    End If
    
    '******************************************************************************************
    'Sediment conentration removed per change request: 11/20/07
'    modProgDialog.ProgDialog "Calculating Sediment Concentration...", strTitle, 0, 27, 26, frmPrj.m_App.hWnd
'    If modProgDialog.g_boolCancel Then
'
'        'STEP 22: Calc the Sediment Conc Raster -----------------------------------------------------------------------
'        With pMapAlgebraOp
'            .BindRaster pTotSedMassHIRaster, "sedmass"
'            .BindRaster g_pRunoffRaster, "runoff"
'        End With
'
'        strExpression = "Con(([sedmass] / [runoff]) >= 0, [sedmass] / [runoff], 0)"
'
'        Set pSedConcRaster = pMapAlgebraOp.Execute(strExpression)
'
'        With pMapAlgebraOp
'            .UnbindRaster "sedmass"
'            .UnbindRaster "runoff"
'        End With
'        'END STEP 22
'
'    End If
    
    
'    modProgDialog.ProgDialog "Adding Sediment Mass to Group Layer...", strTitle, 0, 27, 27, frmPrj.m_App.hWnd
'    If modProgDialog.g_boolCancel Then
'        'STEP 23: Created the Sediment Concentration Raster layer and add to Group Layer -----------------------------------
'        'Get a unique name for MUSLE and return the permanently made raster
'        strMUSLE = modUtil.GetUniqueName("MUSLEconc", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
'
'        'Clip to selected polys if chosen
'        If g_booSelectedPolys Then
'            Dim pClipMusleConcRaster As IRaster
'            Set pClipMusleConcRaster = modUtil.ClipBySelectedPoly(pSedConcRaster, g_pSelectedPolyClip, pEnv)
'            Set pPermSedConcRaster = modUtil.ReturnPermanentRaster(pClipMusleConcRaster, pEnv.OutWorkspace.PathName, strMUSLE)
'            Set pClipMusleConcRaster = Nothing
'        Else
'            Set pPermSedConcRaster = modUtil.ReturnPermanentRaster(pSedConcRaster, pEnv.OutWorkspace.PathName, strMUSLE)
'        End If
'
'        'Now create the MUSLE layer
'        Set pSedConcRasterLayer = modUtil.ReturnRasterLayer(frmPrj.m_App, pPermSedConcRaster, "MUSLE Sediment Concentration (kg/L)")
'        Set pSedConcRasterLayer.Renderer = modUtil.ReturnRasterStretchColorRampRender(pSedConcRasterLayer, "Brown")
'        pSedConcRasterLayer.Visible = False
'
'        'Metadata:
'        g_dicMetadata.Add pSedConcRasterLayer.Name, m_strMusleMetadata
'
'        'Add the MUSLE Layer to the final group layer
'        g_pGroupLayer.Add pSedConcRasterLayer
'
'        'end STEP 23: Created the Sediment Concentration Raster layer and add to Group Layer -----------------------------------
'
'    End If
    'END REMOVE ***************************************************************************************
    
    CalcMUSLE = True
    
    modProgDialog.KillDialog
        
    'Clean up all existing rasters that have not yet been destroyed.
    Set pMapAlgebraOp = Nothing
    Set pFlowAccumOp = Nothing
    Set pFlowAccumOp1 = Nothing
    Set pFlowDirRDS = Nothing
    Set pLiterRDS = Nothing
    Set pOutRDS = Nothing
    Set pFlowDirRDS1 = Nothing
    Set pHISYMGRDS = Nothing
    Set pOutRDS1 = Nothing
    Set clsPrecip = Nothing
    Set pPermTotSedConcHIraster = Nothing
    Set pPermSedConcRaster = Nothing
    
Exit Function

ErrHandler:
    If Err.Number = -2147217297 Then 'S.A. constant for User cancelled operation
        modProgDialog.g_boolCancel = False
    ElseIf Err.Number = -2147467259 Then 'S.A. constant for crappy ESRI stupid GRID error
        MsgBox "ArcMap has reached the maximum number of GRIDs allowed in memory.  " & _
        "Please exit N-SPECT and restart ArcMap.", vbInformation, "Maximum GRID Number Encountered"
        CalcMUSLE = False
        modProgDialog.g_boolCancel = False
        modProgDialog.KillDialog
    Else
        MsgBox "MUSLE Error: " & Err.Number & " on MUSLE Calculation: " & strExpression
        CalcMUSLE = False
        modProgDialog.g_boolCancel = False
        modProgDialog.KillDialog
    End If
End Function


Private Function ConstructConStatment(rsCF As ADODB.Recordset, pLCRaster As IRaster) As String
'Creates the Cover Factor con statement using the name of the the LandCass Recordset, and the Landclass Raster
'Returns: String
'Looks like: con(([nu_lulc] eq 2), 0.000, con((nu_lulc eq 3), 0.030....

    Dim strCon As String            'Con statement base
    Dim strParens As String         'String of trailing parens
    Dim strCompleteCon As String    'Concatenate of strCon & strParens
    
    Dim pRasterCol As IRasterBandCollection
    Dim pBand As IRasterBand
    Dim pTable As ITable
    Dim TableExist As Boolean
    Dim pCursor As ICursor
    Dim pRow As iRow
    Dim FieldIndex As Integer
    Dim booValueFound As Boolean
    Dim i As Integer
    
    'STEP 1:  get the records from the database -----------------------------------------------
    rsCF.MoveFirst
    'End Database stuff
    
    'STEP 2: Raster Values ---------------------------------------------------------------------
    'Now Get the RASTER values
    ' Get Rasterband from the incoming raster
    Set pRasterCol = pLCRaster
    Set pBand = pRasterCol.Item(0)

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
    
    'STEP 3: Table values vs. Raster Values Count - if not equal bark --------------------------
'    If pTable.RowCount(Nothing) <> rsCF.RecordCount Then
'        MsgBox "Error: The number of records in your Land Class Table do not match your GRID dataset."
'        Exit Function
'    End If
    
    'STEP 4: Create the strings
    'Loop through and get all values
    Do While Not pRow Is Nothing
        booValueFound = False
        rsCF.MoveFirst
    
        For i = 0 To rsCF.RecordCount - 1
            If pRow.Value(FieldIndex) = rsCF!Value Then
            
                booValueFound = True
            
                If strCon = "" Then
                    strCon = "Con(([nu_lulc] eq " & pRow.Value(FieldIndex) & "), " & rsCF!CoverFactor & ", "
                Else
                    strCon = strCon & "Con(([nu_lulc] eq " & pRow.Value(FieldIndex) & "), " & rsCF!CoverFactor & ", "
                End If
                
                If strParens = "" Then
                    strParens = "-9999)"
                Else
                    strParens = strParens & ")"
                End If
                Exit For
            Else
                booValueFound = False
            End If
            rsCF.MoveNext
        Next i
        
        If booValueFound = False Then
            MsgBox "Error in MUSLE ConstructConStatment Function: Your N-SPECT Land Class Table is missing values found in your landcover GRID dataset."
            ConstructConStatment = ""
            Exit Function
        Else
            Set pRow = pCursor.NextRow
            i = 0
        End If
        
    Loop
    
    
     'Remove 11/30/2007 in favor of check above.
'    '==========================================================================
'    Do While Not pRow Is Nothing
'
'        booValueFound = False
'        rsType.MoveFirst
'
'        For i = 0 To rsType.RecordCount - 1
'            If rsType!Value = pRow.Value Then
'
'                booValueFound = True
'
'                If strCon = "" Then
'                    strCon = "Con(([nu_lulc] eq " & pRow.Value(FieldIndex) & "), " & rsType!CoeffType & ", "
'                Else
'                    strCon = strCon & "Con(([nu_lulc] eq " & pRow.Value(FieldIndex) & "), " & rsType!CoeffType & ", "
'                End If
'
'                If strParens = "" Then
'                    strParens = "-9999)"
'                Else
'                    strParens = strParens & ")"
'                End If
'
'                Exit For
'            Else
'                booValueFound = False
'            End If
'            rsType.MoveNext
'        Next i
'
'        If booValueFound = False Then
'            MsgBox "Values in table LCClass table not equal to values in landclass dataset."
'            ConstructConStatment = ""
'            Exit Function
'        Else
'            Set pRow = pCursor.NextRow
'            i = 0
'        End If
'    Loop
    '==================================================================================
               
    strCompleteCon = strCon & strParens
    ConstructConStatment = strCompleteCon
    
    'Cleanup:
    'Set pLCRaster = Nothing
    Set pRasterCol = Nothing
    Set pBand = Nothing
    Set pTable = Nothing
    Set pCursor = Nothing
    Set pRow = Nothing
    
End Function



Private Function ConstructPondConStatement(rsCF As ADODB.Recordset, pLCRaster As IRaster) As String
'Creates the Con Statement used in the Pond Factor GRID
'Returns: String
'Looks like: con(([nu_lulc] eq 16), 0, con((nu_lulc eq 17), 0...

    Dim strCon As String            'Con statement base
    Dim strParens As String         'String of trailing parens
    Dim strCompleteCon As String    'Concatenate of strCon & strParens
    
    Dim pRasterCol As IRasterBandCollection
    Dim pBand As IRasterBand
    Dim pTable As ITable
    Dim TableExist As Boolean
    Dim pCursor As ICursor
    Dim pRow As iRow
    Dim FieldIndex As Integer
    Dim booValueFound As Boolean
    Dim i As Integer
    
    'STEP 1:  get the records from the database -----------------------------------------------
    rsCF.MoveFirst
    'End Database stuff
    
    'STEP 2: Raster Values ---------------------------------------------------------------------
    'Now Get the RASTER values
    ' Get Rasterband from the incoming raster
    Set pRasterCol = pLCRaster
    Set pBand = pRasterCol.Item(0)

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
    
    'REMOVED 11/30/2007
    'STEP 3: Table values vs. Raster Values Count - if not equal bark --------------------------
'    If pTable.RowCount(Nothing) <> rsCF.RecordCount Then
'        MsgBox "Error: The number of records in your Land Class Table do not match your GRID dataset."
'        Exit Function
'    End If
'
    'STEP 4: Create the strings
    'Loop through and get all values
    Do While Not pRow Is Nothing
    
        booValueFound = False
        rsCF.MoveFirst
        
        For i = 0 To rsCF.RecordCount - 1
            If pRow.Value(FieldIndex) = rsCF!Value Then
                booValueFound = True
                Select Case rsCF!W_WL
                Case 0 'Means the current landclass is NOT Water or Wetland, therefore gets a 1
                    If strCon = "" Then
                        strCon = "Con(([nu_lulc] eq " & pRow.Value(FieldIndex) & "), 1, "
                    Else
                        strCon = strCon & "Con(([nu_lulc] eq " & pRow.Value(FieldIndex) & "), 1, "
                    End If
                    
                    If strParens = "" Then
                        strParens = "1)"
                    Else
                        strParens = strParens & ")"
                    End If
                    
                Case 1 'Means the current landclass IS Water or Wetland, therefore gets a 0
                    If strCon = "" Then
                        strCon = "Con(([nu_lulc] eq " & pRow.Value(FieldIndex) & "), 0, "
                    Else
                        strCon = strCon & "Con(([nu_lulc] eq " & pRow.Value(FieldIndex) & "), 0, "
                    End If
                    
                    If strParens = "" Then
                        strParens = ")"
                    Else
                        strParens = strParens & ")"
                    End If
                
                End Select
                
                'rsCF.MoveNext
                'Set pRow = pCursor.NextRow
                Exit For
                
            Else
                booValueFound = False
            End If
            
            rsCF.MoveNext
        
        Next i
        
        If booValueFound = False Then
            MsgBox "Error in MUSLEConstructPondConStatement: Your N-SPECT Land Class Table is missing values found in your landcover GRID dataset."
            ConstructPondConStatement = ""
            Exit Function
        Else
            Set pRow = pCursor.NextRow
            i = 0
        End If
    Loop
        
    strCompleteCon = strCon & strParens
    ConstructPondConStatement = strCompleteCon
    
    'Cleanup:
    'Set pLCRaster = Nothing
    Set pRasterCol = Nothing
    Set pBand = Nothing
    Set pTable = Nothing
    Set pCursor = Nothing
    Set pRow = Nothing
    
End Function



Private Function CreateMetadata(booLocal As Boolean) As String
    
    Dim strHeader As String
    'Dim i As Integer
    'Dim strCFactor As String
    
    'Set up the header w/or without flow direction
    If booLocal = True Then
        strHeader = vbTab & "Input Datasets:" & vbNewLine & _
                vbTab & vbTab & "Hydrologic soils grid: " & g_clsXMLPrjFile.strSoilsHydFileName & vbNewLine & _
                vbTab & vbTab & "Landcover grid: " & g_clsXMLPrjFile.strLCGridFileName & vbNewLine & _
                vbTab & vbTab & "Precipitation grid: " & g_strPrecipFileName & vbNewLine & _
                vbTab & vbTab & "Soils K Factor: " & g_clsXMLPrjFile.strSoilsKFileName & vbNewLine & _
                vbTab & vbTab & "LS Factor grid: " & g_strLSFileName & vbNewLine & _
                g_strLandCoverParameters & vbNewLine 'append the g_strLandCoverParameters that was set up during runoff
    Else
        strHeader = vbTab & "Input Datasets:" & vbNewLine & _
                    vbTab & vbTab & "Hydrologic soils grid: " & g_clsXMLPrjFile.strSoilsHydFileName & vbNewLine & _
                    vbTab & vbTab & "Landcover grid: " & g_clsXMLPrjFile.strLCGridFileName & vbNewLine & _
                    vbTab & vbTab & "Precipitation grid: " & g_strPrecipFileName & vbNewLine & _
                    vbTab & vbTab & "Soils K Factor: " & g_clsXMLPrjFile.strSoilsKFileName & vbNewLine & _
                    vbTab & vbTab & "LS Factor grid: " & g_strLSFileName & vbNewLine & _
                    vbTab & vbTab & "Flow direction grid: " & g_strFlowDirFilename & vbNewLine & _
                    g_strLandCoverParameters & vbNewLine
    End If
    
    'Now report the C:Factor figures for the landcover
'    rsCFactor.MoveFirst
'
'    For i = 1 To rsCFactor.RecordCount
'        If i = 1 Then
'            strCFactor = vbTab & vbTab & rsCFactor!Name & ": " & rsCFactor!CoverFactor & vbNewLine
'        Else
'            strCFactor = strCFactor & vbTab & vbTab & rsCFactor!Name & ": " & rsCFactor!CoverFactor & vbNewLine
'        End If
'        rsCFactor.MoveNext
'    Next i

    CreateMetadata = strHeader '& vbTab & "C-Factor values: " & vbNewLine & strCFactor

End Function



