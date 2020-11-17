Option Strict Off
Option Explicit On
Module modMUSLE
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
	Private m_strCFactorConStatement As String 'C Factor Con Statement
	Private m_strPondConStatement As String 'Pond Factor con statement
	Private m_dblMUSLEVal As Double 'Customizable MUSLE value in Equation
	Private m_dblMUSLEExp As Double 'Customizable musle exponent in equation
	
	Private m_strMusleMetadata As String 'MUSLE metadata string
	
	' Variables used by the Error handler function - DO NOT REMOVE
	Const c_sModuleFileName As String = "modMUSLE.bas"
	
	
	Public Function MUSLESetup(ByRef strSoilsDefName As String, ByRef strKfactorFileName As String, ByRef strLandClass As String) As Boolean
		
		'Sub takes incoming parameters from the project file and then parses them out
		'strSoilsDefName: Name of the Soils Definition being used
		'strKFactorFileName: K Factor FileName
		'strLandClass: Name of the Landclass we're using
		
		'RecordSet to get the coverfactor
		Dim rsCoverFactor As New ADODB.Recordset
		Dim strTempLCType As String 'Our potential holder for a temp landtype
		
		Dim rsSoilsDef As New ADODB.Recordset
		Dim strSoilsDef As String
		
		'Open Strings
		Dim strCovFactor As String
		Dim strError As String
		
		'STEP 1: Get the MUSLE Values
		strSoilsDef = "SELECT * FROM SOILS WHERE NAME LIKE '" & strSoilsDefName & "'"
		
		rsSoilsDef.Open(strSoilsDef, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		
		m_dblMUSLEVal = rsSoilsDef.Fields("MUSLEVal").Value
		m_dblMUSLEExp = rsSoilsDef.Fields("MUSLEExp").Value
		
		
		'STEP 2: Set the K Factor Raster
		If modUtil.RasterExists(strKfactorFileName) Then
			g_KFactorRaster = modUtil.ReturnRaster(strKfactorFileName)
		Else
			strError = "K Factor Raster Does Not Exist: " & strKfactorFileName
		End If
		'END STEP 1: -----------------------------------------------------------------------------------
		
		If Len(g_DictTempNames.Item(strLandClass)) > 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object g_DictTempNames.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			strTempLCType = g_DictTempNames.Item(strLandClass)
		Else
			strTempLCType = strLandClass
		End If
		
		'Get the landclasses of type strLandClass
		strCovFactor = "SELECT LCTYPE.LCTYPEID, LCCLASS.NAME, LCCLASS.VALUE, LCCLASS.COVERFACTOR, LCCLASS.W_WL FROM " & "LCTYPE INNER JOIN LCCLASS ON LCTYPE.LCTYPEID = LCCLASS.LCTYPEID " & "WHERE LCTYPE.NAME LIKE '" & strTempLCType & "' ORDER BY LCCLASS.VALUE"
		
		m_strMusleMetadata = CreateMetadata(g_booLocalEffects) ', rsCoverFactor)
		
		rsCoverFactor.Open(strCovFactor, g_ADOConn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		
		If Len(strError) > 0 Then
			MsgBox(strError)
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
		rsCoverFactor.Close()
		'UPGRADE_NOTE: Object rsCoverFactor may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsCoverFactor = Nothing
		
		rsSoilsDef.Close()
		'UPGRADE_NOTE: Object rsSoilsDef may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsSoilsDef = Nothing
		
		
		
	End Function
	
	Private Function CalcMUSLE(ByRef strConStatement As String, ByRef strConPondStatement As String) As Boolean
		'Incoming strings: strConStatment: the monster con statement
		'strConPondstatement: the con for the pond stuff
		'Calculates the MUSLE erosion model
		
		Dim pWeightRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 1: Weight Raster
		Dim pWSLengthRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 2: Watershed length
		Dim pWSLengthUnitsRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 3: Units conversion
		Dim pSlopePRRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 4: Average Slope
		Dim pSlopeModRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 4a: Mod Slope
		Dim pTemp1LagRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 5a: Lag temp1
		Dim pTemp2LagRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 5b: Lag temp2
		Dim pTemp3LagRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 5c: Lag temp3
		Dim pTemp4LagRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 5d: Lag temp4
		Dim pLagRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 5e: Lag
		Dim pTOCRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 6a: TOC
		Dim pTOCTempRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 6b: TOC temp
		Dim pModTOCRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 6c: Mod TOC
		Dim pAbPrecipRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 7: Abstraction-Precip Ratio
		Dim pLogTOCRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 8a: Calc unit peak discharge
		Dim pTempLogTOCRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 8b:
		Dim pCZeroRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 9: CZero
		Dim pConeRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 10: Cone
		Dim pCTwoRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 11: CTwo
		Dim pLogQuRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 12a: Who knows
		Dim pQuRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STep 12b: Who Knows(b)
		Dim pPondFactorRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 13: Pond Factor
		Dim pQPRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 14: QP factor
		Dim pCFactorRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 15: Yee old CFactor Raster
		Dim pSYTempRaster As ESRI.ArcGIS.Geodatabase.IRaster 'Step 16a Temp yield
		Dim pSYRaster As ESRI.ArcGIS.Geodatabase.IRaster 'Step 16b Yield
		Dim pHISYTempRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 16c: HI Specific temp yield
		Dim pHISYRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 16d: HI Specific yield
		'Dim pSYMGRaster As IRaster                  'STEP 17a: tons to milligrams
		Dim pHISYMGRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 17b: tons to milligrams, HI Specific
		Dim pPermMUSLERaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 17c: Local Effects permanent raster
		Dim pTempFlowDir1Raster As ESRI.ArcGIS.Geodatabase.IRaster 'flowDir Temp Raster
		'Dim pTempFlowDir2Raster As IRaster          'FlowDir temp raster
		Dim pTempflowDir3raster As ESRI.ArcGIS.Geodatabase.IRaster 'FlowDir temp raster
		Dim pLiterRunoffRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 18: Runoff to liters
		'Dim pAccSedRaster As IRaster                'STEP 19: Acc sediment
		Dim pAccSedHIRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 19HI: acc sediment
		Dim pFlowAccumOp As ESRI.ArcGIS.SpatialAnalyst.IHydrologyOp 'STEP 19: IHydroOp and friends
		Dim pFlowDirRDS As ESRI.ArcGIS.Geodatabase.IGeoDataset
		Dim pLiterRDS As ESRI.ArcGIS.Geodatabase.IGeoDataset
		Dim pOutRDS As ESRI.ArcGIS.Geodatabase.IGeoDataset
		Dim pAccRunLiterRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 19a: Acc runoff liter
		Dim pFlowAccumOp1 As ESRI.ArcGIS.SpatialAnalyst.IHydrologyOp 'STEP 19HI: IHydroOp
		Dim pFlowDirRDS1 As ESRI.ArcGIS.Geodatabase.IGeoDataset
		Dim pHISYMGRDS As ESRI.ArcGIS.Geodatabase.IGeoDataset
		Dim pOutRDS1 As ESRI.ArcGIS.Geodatabase.IGeoDataset
		Dim pTotSedMassRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 20: Tot sed mass
		Dim pTotSedMassHIRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 20HI: tot sed mass HI
		Dim pTotSedTempRaster As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 21: First step
		Dim pPermTotSedConcHIraster As ESRI.ArcGIS.Geodatabase.IRaster 'Permanent MUSLE
		Dim pSedConcRaster As ESRI.ArcGIS.Geodatabase.IRaster 'Sed Concentration
		Dim pPermSedConcRaster As ESRI.ArcGIS.Geodatabase.IRaster 'Permanent Sed Concentration
		Dim pSedConcRasterLayer As ESRI.ArcGIS.Carto.IRasterLayer 'Sed Concentration layer
		Dim pMUSLERasterLayer As ESRI.ArcGIS.Carto.IRasterLayer
		Dim pMUSLERasterLocLayer As ESRI.ArcGIS.Carto.IRasterLayer 'MUSLE Local Effects
		Dim strMUSLE As String
		
		
		Dim clsPrecip As New clsPrecipType
		
		Dim pEnv As ESRI.ArcGIS.GeoAnalyst.IRasterAnalysisEnvironment 'Raster Environment
		
		'Create Map Algebra Operator
		Dim pMapAlgebraOp As ESRI.ArcGIS.SpatialAnalyst.IMapAlgebraOp 'Workhorse
		
		'String to hold calculations
		Dim strExpression As String
		Const strTitle As String = "Processing MUSLE Calculation..."
		
		'TEMP code
		Dim strTemp As String
		Dim pTempRaster As ESRI.ArcGIS.Geodatabase.IRaster
		
		'Set the enviornment stuff
		pEnv = g_pSpatEnv
		pMapAlgebraOp = New ESRI.ArcGIS.SpatialAnalyst.RasterMapAlgebraOp
		pEnv = pMapAlgebraOp
		
		On Error GoTo ErrHandler
		
		modProgDialog.ProgDialog("Computing length/distance...", strTitle, 0, 27, 1, (frmPrj.m_App.hwnd))
		If modProgDialog.g_boolCancel Then
			'STEP 1: ------------------------------------------------------------------------------------
			'Create weight grid that represents cell length/distance
			'UPGRADE_WARNING: Couldn't resolve default property of object g_pFlowDirRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pMapAlgebraOp.BindRaster(g_pFlowDirRaster, "flowdir")
			
			strExpression = "Con([flowdir] eq 2, 1.41421, " & "Con([flowdir] eq 8, 1.41421, " & "Con([flowdir] eq 32, 1.41421, " & "Con([flowdir] eq 128, 1.41421, 1.0))))"
			
			pWeightRaster = pMapAlgebraOp.Execute(strExpression)
			
			pMapAlgebraOp.UnbindRaster("flowdir")
			'END STEP 2: -------------------------------------------------------------------------------
		End If
		
		'TEMP Code added 8/29/06 to help w/ Hawaii's checking of the RUSLE equation
		'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
		'Set pTempRaster = modUtil.ReturnPermanentRaster(pWeightRaster, pEnv.OutWorkspace.PathName, strTemp)
		
		modProgDialog.ProgDialog("Calculating Watershed Length...", strTitle, 0, 27, 2, (frmPrj.m_App.hwnd))
		If modProgDialog.g_boolCancel Then
			'STEP 2: ------------------------------------------------------------------------------------
			'Calculate Watershed Length
			With pMapAlgebraOp
				'UPGRADE_WARNING: Couldn't resolve default property of object g_pFlowDirRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(g_pFlowDirRaster, "flowdir")
				'UPGRADE_WARNING: Couldn't resolve default property of object pWeightRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(pWeightRaster, "weight")
			End With
			
			strExpression = "flowlength([flowdir], [weight], upstream)"
			
			pWSLengthRaster = pMapAlgebraOp.Execute(strExpression)
			
			With pMapAlgebraOp
				.UnbindRaster("flowdir")
				.UnbindRaster("weight")
			End With
			'End STEP 2 ----------------------------------------------------------------------------------
		End If
		
		'TEMP Code added 8/29/06 to help w/ Hawaii's checking of the RUSLE equation
		'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
		'Set pTempRaster = modUtil.ReturnPermanentRaster(pWSLengthRaster, pEnv.OutWorkspace.PathName, strTemp)
		
		'Notice throughout that we cleanup before the end...don't want to exceed the max number allowed rasters
		'in memory.
		'UPGRADE_NOTE: Object pWeightRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pWeightRaster = Nothing
		
		modProgDialog.ProgDialog("Converting units...", strTitle, 0, 27, 3, (frmPrj.m_App.hwnd))
		If modProgDialog.g_boolCancel Then
			'STEP 3: --------------------------------------------------------------------------------------
			'Convert Metric Units
			'UPGRADE_WARNING: Couldn't resolve default property of object pWSLengthRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pMapAlgebraOp.BindRaster(pWSLengthRaster, "cell_wslength")
			
			strExpression = "([cell_wslength] * 3.28084)"
			
			pWSLengthUnitsRaster = pMapAlgebraOp.Execute(strExpression)
			
			pMapAlgebraOp.UnbindRaster("cell_wslength")
			'END STEP 3: -----------------------------------------------------------------------------------
		End If
		
		'TEMP Code added 2008 to help w/ debugging
		'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
		'Set pTempRaster = modUtil.ReturnPermanentRaster(pWSLengthUnitsRaster, pEnv.OutWorkspace.PathName, strTemp)
		
		'UPGRADE_NOTE: Object pWSLengthRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pWSLengthRaster = Nothing
		
		modProgDialog.ProgDialog("Calculating Average Slope...", strTitle, 0, 27, 4, (frmPrj.m_App.hwnd))
		If modProgDialog.g_boolCancel Then
			'STEP 4a: ---------------------------------------------------------------------------------------
			'Calculate Average Slope
			'UPGRADE_WARNING: Couldn't resolve default property of object g_pDEMRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pMapAlgebraOp.BindRaster(g_pDEMRaster, "dem")
			
			strExpression = "slope([dem], percentrise)"
			
			pSlopePRRaster = pMapAlgebraOp.Execute(strExpression)
			
			pMapAlgebraOp.UnbindRaster("dem")
			'END STEP 4a ------------------------------------------------------------------------------------
		End If
		
		'TEMP Code added 2008 to help w/ debugging
		'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
		'Set pTempRaster = modUtil.ReturnPermanentRaster(pSlopePRRaster, pEnv.OutWorkspace.PathName, strTemp)
		
		modProgDialog.ProgDialog("Calculating Mod Slope...", strTitle, 0, 27, 5, (frmPrj.m_App.hwnd))
		If modProgDialog.g_boolCancel Then
			'STEP 4b: ---------------------------------------------------------------------------------------
			'Calculate modslope
			'UPGRADE_WARNING: Couldn't resolve default property of object pSlopePRRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pMapAlgebraOp.BindRaster(pSlopePRRaster, "slopepr")
			
			strExpression = "Con([slopepr] eq 0, 0.1, [slopepr])"
			
			pSlopeModRaster = pMapAlgebraOp.Execute(strExpression)
			
			pMapAlgebraOp.UnbindRaster("slopepr")
			'END STEP 4b: -----------------------------------------------------------------------------------
		End If
		
		'TEMP Code added 2008 to help w/ debugging
		'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
		'Set pTempRaster = modUtil.ReturnPermanentRaster(pSlopeModRaster, pEnv.OutWorkspace.PathName, strTemp)
		
		'UPGRADE_NOTE: Object pSlopePRRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pSlopePRRaster = Nothing
		
		
		modProgDialog.ProgDialog("Calculating Lag...", strTitle, 0, 27, 6, (frmPrj.m_App.hwnd))
		If modProgDialog.g_boolCancel Then
			'STEP 5a: ---------------------------------------------------------------------------------------
			'Calculate Lag
			'UPGRADE_WARNING: Couldn't resolve default property of object pWSLengthUnitsRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pMapAlgebraOp.BindRaster(pWSLengthUnitsRaster, "cell_wslngft")
			
			strExpression = "Pow([cell_wslngft], 0.8)"
			
			pTemp1LagRaster = pMapAlgebraOp.Execute(strExpression)
			
			pMapAlgebraOp.UnbindRaster("cell_wslngft")
			'END STEP 5a: ----------------------------------------------------------------------------------
		End If
		
		'TEMP Code added 2008 to help w/ debugging
		'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
		'Set pTempRaster = modUtil.ReturnPermanentRaster(pTemp1LagRaster, pEnv.OutWorkspace.PathName, strTemp)
		
		'UPGRADE_NOTE: Object pWSLengthUnitsRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pWSLengthUnitsRaster = Nothing
		
		modProgDialog.ProgDialog("Checking SCS GRID...", strTitle, 0, 27, 7, (frmPrj.m_App.hwnd))
		If modProgDialog.g_boolCancel Then
			'STEP 5b: ---------------------------------------------------------------------------------------
			'UPGRADE_WARNING: Couldn't resolve default property of object g_pSCS100Raster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pMapAlgebraOp.BindRaster(g_pSCS100Raster, "scsgrid100")
			
			strExpression = "(1000 / [scsgrid100]) - 9"
			
			pTemp2LagRaster = pMapAlgebraOp.Execute(strExpression)
			
			pMapAlgebraOp.UnbindRaster("scsgrid100")
			'END STEP 9b: ----------------------------------------------------------------------------------
		End If
		
		'TEMP Code added 2008 to help w/ debugging
		'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
		'Set pTempRaster = modUtil.ReturnPermanentRaster(pTemp2LagRaster, pEnv.OutWorkspace.PathName, strTemp)
		
		
		modProgDialog.ProgDialog("Multiplying GRIDs...", strTitle, 0, 27, 8, (frmPrj.m_App.hwnd))
		If modProgDialog.g_boolCancel Then
			'STEP 5c: --------------------------------------------------------------------------------------
			'UPGRADE_WARNING: Couldn't resolve default property of object pTemp2LagRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pMapAlgebraOp.BindRaster(pTemp2LagRaster, "temp4")
			
			strExpression = "Pow([temp4], 0.7)"
			
			pTemp3LagRaster = pMapAlgebraOp.Execute(strExpression)
			
			pMapAlgebraOp.UnbindRaster("temp4")
			'END STEP 5c: ----------------------------------------------------------------------------------
		End If
		
		'TEMP Code added 2008 to help w/ debugging
		'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
		'Set pTempRaster = modUtil.ReturnPermanentRaster(pTemp3LagRaster, pEnv.OutWorkspace.PathName, strTemp)
		
		
		'UPGRADE_NOTE: Object pTemp2LagRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pTemp2LagRaster = Nothing
		
		modProgDialog.ProgDialog("Pow([modslope], 0.5...", strTitle, 0, 27, 9, (frmPrj.m_App.hwnd))
		If modProgDialog.g_boolCancel Then
			'STEP 5d: --------------------------------------------------------------------------------------
			'UPGRADE_WARNING: Couldn't resolve default property of object pSlopeModRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pMapAlgebraOp.BindRaster(pSlopeModRaster, "modslope")
			
			strExpression = "Pow([modslope], 0.5)"
			
			pTemp4LagRaster = pMapAlgebraOp.Execute(strExpression)
			
			pMapAlgebraOp.UnbindRaster("modslope")
			'END STEP 5d: ----------------------------------------------------------------------------------
		End If
		
		'TEMP Code added 2008 to help w/ debugging
		'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
		'Set pTempRaster = modUtil.ReturnPermanentRaster(pTemp4LagRaster, pEnv.OutWorkspace.PathName, strTemp)
		
		'UPGRADE_NOTE: Object pSlopeModRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pSlopeModRaster = Nothing
		
		modProgDialog.ProgDialog("Lag calculation...", strTitle, 0, 30, 27, (frmPrj.m_App.hwnd))
		If modProgDialog.g_boolCancel Then
			'STEP 5e: --------------------------------------------------------------------------------------
			'Finally, the lag calculation
			With pMapAlgebraOp
				'UPGRADE_WARNING: Couldn't resolve default property of object pTemp1LagRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(pTemp1LagRaster, "temp3")
				'UPGRADE_WARNING: Couldn't resolve default property of object pTemp3LagRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(pTemp3LagRaster, "temp5")
				'UPGRADE_WARNING: Couldn't resolve default property of object pTemp4LagRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(pTemp4LagRaster, "temp6")
			End With
			
			strExpression = "([temp3] * [temp5]) / (1900 * [temp6])"
			
			pLagRaster = pMapAlgebraOp.Execute(strExpression)
			
			With pMapAlgebraOp
				.UnbindRaster("temp3")
				.UnbindRaster("temp5")
				.UnbindRaster("temp6")
			End With
			'STEP 5e: ---------------------------------------------------------------------------------------
		End If
		
		'TEMP Code added 2008 to help w/ debugging
		'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
		'Set pTempRaster = modUtil.ReturnPermanentRaster(pLagRaster, pEnv.OutWorkspace.PathName, strTemp)
		
		'UPGRADE_NOTE: Object pTemp1LagRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pTemp1LagRaster = Nothing
		'UPGRADE_NOTE: Object pTemp3LagRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pTemp3LagRaster = Nothing
		'UPGRADE_NOTE: Object pTemp4LagRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pTemp4LagRaster = Nothing
		
		
		modProgDialog.ProgDialog("Calculating time of concentration...", strTitle, 0, 27, 11, (frmPrj.m_App.hwnd))
		If modProgDialog.g_boolCancel Then
			
			'STEP 6a: ---------------------------------------------------------------------------------------
			'Calculate the time of concentration
			'UPGRADE_WARNING: Couldn't resolve default property of object pLagRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pMapAlgebraOp.BindRaster(pLagRaster, "lag")
			
			strExpression = "[lag] / 0.6"
			
			pTOCRaster = pMapAlgebraOp.Execute(strExpression)
			
			pMapAlgebraOp.UnbindRaster("lag")
			'END STEP 6a: -----------------------------------------------------------------------------------
			
			'TEMP Code added 2008 to help w/ debugging
			'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
			'Set pTempRaster = modUtil.ReturnPermanentRaster(pTOCRaster, pEnv.OutWorkspace.PathName, strTemp)
			
			'UPGRADE_NOTE: Object pLagRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			pLagRaster = Nothing
			
			'STEP 6b: --------------------------------------------------------------------------------------
			'UPGRADE_WARNING: Couldn't resolve default property of object pTOCRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pMapAlgebraOp.BindRaster(pTOCRaster, "toc")
			
			strExpression = "Con([toc] lt 0.1, 0.1, [toc])"
			
			pTOCTempRaster = pMapAlgebraOp.Execute(strExpression)
			
			pMapAlgebraOp.UnbindRaster("toc")
			'END STEP 6b: ----------------------------------------------------------------------------------
			
			'TEMP Code added 2008 to help w/ debugging
			'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
			'Set pTempRaster = modUtil.ReturnPermanentRaster(pTOCTempRaster, pEnv.OutWorkspace.PathName, strTemp)
			
			'UPGRADE_NOTE: Object pTOCRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			pTOCRaster = Nothing
			
			'STEP 6c: --------------------------------------------------------------------------------------
			'UPGRADE_WARNING: Couldn't resolve default property of object pTOCTempRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pMapAlgebraOp.BindRaster(pTOCTempRaster, "temp7")
			
			strExpression = "Con([temp7] gt 10, 10, [temp7])"
			
			pModTOCRaster = pMapAlgebraOp.Execute(strExpression)
			
			pMapAlgebraOp.UnbindRaster("temp7")
			'END STEP 6c: ----------------------------------------------------------------------------------
			
			'TEMP Code added 2008 to help w/ debugging
			'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
			'Set pTempRaster = modUtil.ReturnPermanentRaster(pModTOCRaster, pEnv.OutWorkspace.PathName, strTemp)
			
		End If
		
		'UPGRADE_NOTE: Object pTOCTempRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pTOCTempRaster = Nothing
		
		modProgDialog.ProgDialog("Abstraction Precipitation Ratio...", strTitle, 0, 27, 12, (frmPrj.m_App.hwnd))
		If modProgDialog.g_boolCancel Then
			'STEP 7: ---------------------------------------------------------------------------------------
			'Find the Abstraction Precipitation Ratio
			With pMapAlgebraOp
				'UPGRADE_WARNING: Couldn't resolve default property of object g_pAbstractRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(g_pAbstractRaster, "abstract")
				'UPGRADE_WARNING: Couldn't resolve default property of object g_pPrecipRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(g_pPrecipRaster, "rain")
			End With
			
			strExpression = "[abstract] / [rain]"
			
			pAbPrecipRaster = pMapAlgebraOp.Execute(strExpression)
			
			With pMapAlgebraOp
				.UnbindRaster("abstract")
				.UnbindRaster("rain")
			End With
			'END STEP 7: ----------------------------------------------------------------------------------
		End If
		
		'TEMP Code added 2008 to help w/ debugging
		'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
		'Set pTempRaster = modUtil.ReturnPermanentRaster(pAbPrecipRaster, pEnv.OutWorkspace.PathName, strTemp)
		
		modProgDialog.ProgDialog("Calculating the peak unit discharge...", strTitle, 0, 27, 13, (frmPrj.m_App.hwnd))
		If modProgDialog.g_boolCancel Then
			'STEP 8a: ----------------------------------------------------------------------------------------
			'Calculate the unit peak discharge
			'UPGRADE_WARNING: Couldn't resolve default property of object pModTOCRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pMapAlgebraOp.BindRaster(pModTOCRaster, "modtoc")
			
			strExpression = "log10([modtoc])"
			
			pLogTOCRaster = pMapAlgebraOp.Execute(strExpression)
			
			pMapAlgebraOp.UnbindRaster("modtoc")
			'END STEP 8a: -----------------------------------------------------------------------------------
			
			'TEMP Code added 2008 to help w/ debugging
			'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
			'Set pTempRaster = modUtil.ReturnPermanentRaster(pLogTOCRaster, pEnv.OutWorkspace.PathName, strTemp)
			
			'STEP 8b: ---------------------------------------------------------------------------------------
			'2nd part of it
			'UPGRADE_WARNING: Couldn't resolve default property of object pLogTOCRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pMapAlgebraOp.BindRaster(pLogTOCRaster, "logtoc")
			
			strExpression = "Pow([logtoc], 2)"
			
			pTempLogTOCRaster = pMapAlgebraOp.Execute(strExpression)
			
			pMapAlgebraOp.UnbindRaster("logtoc")
			'END STEP 8b: -----------------------------------------------------------------------------------
			
			'TEMP Code added 2008 to help w/ debugging
			'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
			'Set pTempRaster = modUtil.ReturnPermanentRaster(pTempLogTOCRaster, pEnv.OutWorkspace.PathName, strTemp)
			
		End If
		
		'UPGRADE_NOTE: Object pModTOCRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pModTOCRaster = Nothing
		
		modProgDialog.ProgDialog("Creating C-Zero GRID...", strTitle, 0, 27, 14, (frmPrj.m_App.hwnd))
		If modProgDialog.g_boolCancel Then
			'STEP 9: ----------------------------------------------------------------------------------------
			'CZERO GRID
			'UPGRADE_WARNING: Couldn't resolve default property of object pAbPrecipRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pMapAlgebraOp.BindRaster(pAbPrecipRaster, "ip")
			
			'Call to clsPrecipType(called clsprecip here init above) to get string based on Precip Type
			'g_intPrecipType
			strExpression = clsPrecip.CZero(g_intPrecipType)
			
			pCZeroRaster = pMapAlgebraOp.Execute(strExpression)
			
			pMapAlgebraOp.UnbindRaster("ip")
			'END STEP 9: ------------------------------------------------------------------------------------
		End If
		
		'TEMP Code added 2008 to help w/ debugging
		'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
		'Set pTempRaster = modUtil.ReturnPermanentRaster(pCZeroRaster, pEnv.OutWorkspace.PathName, strTemp)
		
		
		modProgDialog.ProgDialog("Creating Cone GRID...", strTitle, 0, 27, 15, (frmPrj.m_App.hwnd))
		If modProgDialog.g_boolCancel Then
			'STEP 10: ---------------------------------------------------------------------------------------
			'CONE grid
			'UPGRADE_WARNING: Couldn't resolve default property of object pAbPrecipRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pMapAlgebraOp.BindRaster(pAbPrecipRaster, "ip")
			
			strExpression = clsPrecip.Cone(g_intPrecipType)
			
			pConeRaster = pMapAlgebraOp.Execute(strExpression)
			
			pMapAlgebraOp.UnbindRaster("ip")
			'END STEP 10 ------------------------------------------------------------------------------------
		End If
		
		'TEMP Code added 2008 to help w/ debugging
		'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
		'Set pTempRaster = modUtil.ReturnPermanentRaster(pConeRaster, pEnv.OutWorkspace.PathName, strTemp)
		
		
		
		modProgDialog.ProgDialog("Creating C2 GRID...", strTitle, 0, 27, 16, (frmPrj.m_App.hwnd))
		If modProgDialog.g_boolCancel Then
			'STEP 11: ---------------------------------------------------------------------------------------
			'CTwo GRID
			'UPGRADE_WARNING: Couldn't resolve default property of object pAbPrecipRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pMapAlgebraOp.BindRaster(pAbPrecipRaster, "ip")
			
			strExpression = clsPrecip.CTwo(g_intPrecipType)
			
			pCTwoRaster = pMapAlgebraOp.Execute(strExpression)
			
			pMapAlgebraOp.UnbindRaster("ip")
			'END STEP 11: ------------------------------------------------------------------------------------
		End If
		
		'TEMP Code added 2008 to help w/ debugging
		'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
		'Set pTempRaster = modUtil.ReturnPermanentRaster(pCTwoRaster, pEnv.OutWorkspace.PathName, strTemp)
		
		
		
		'UPGRADE_NOTE: Object pAbPrecipRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pAbPrecipRaster = Nothing
		
		modProgDialog.ProgDialog("More math...", strTitle, 0, 27, 17, (frmPrj.m_App.hwnd))
		If modProgDialog.g_boolCancel Then
			'STEP 12a: ----------------------------------------------------------------------------------------
			'Logqu
			With pMapAlgebraOp
				'UPGRADE_WARNING: Couldn't resolve default property of object pCZeroRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(pCZeroRaster, "czero")
				'UPGRADE_WARNING: Couldn't resolve default property of object pConeRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(pConeRaster, "cone")
				'UPGRADE_WARNING: Couldn't resolve default property of object pLogTOCRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(pLogTOCRaster, "logtoc")
				'UPGRADE_WARNING: Couldn't resolve default property of object pCTwoRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(pCTwoRaster, "ctwo")
				'UPGRADE_WARNING: Couldn't resolve default property of object pTempLogTOCRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(pTempLogTOCRaster, "temp8")
			End With
			
			strExpression = "[czero] + ([cone] * [logtoc]) + ([ctwo] * [temp8])"
			
			pLogQuRaster = pMapAlgebraOp.Execute(strExpression)
			
			'TEMP Code added 2008 to help w/ debugging
			'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
			'Set pTempRaster = modUtil.ReturnPermanentRaster(pLogQuRaster, pEnv.OutWorkspace.PathName, strTemp)
			
			
			
			With pMapAlgebraOp
				.UnbindRaster("czero")
				.UnbindRaster("cone")
				.UnbindRaster("logtoc")
				.UnbindRaster("ctwo")
				.UnbindRaster("temp8")
			End With
			'END STEP 12a: -------------------------------------------------------------------------------------
			
			'UPGRADE_NOTE: Object pLogTOCRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			pLogTOCRaster = Nothing
			'UPGRADE_NOTE: Object pCZeroRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			pCZeroRaster = Nothing
			'UPGRADE_NOTE: Object pConeRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			pConeRaster = Nothing
			'UPGRADE_NOTE: Object pCTwoRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			pCTwoRaster = Nothing
			
			'STEP 12b: -----------------------------------------------------------------------------------------
			'UPGRADE_WARNING: Couldn't resolve default property of object pLogQuRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pMapAlgebraOp.BindRaster(pLogQuRaster, "logqu")
			
			strExpression = "Pow(10, [logqu])"
			
			pQuRaster = pMapAlgebraOp.Execute(strExpression)
			
			pMapAlgebraOp.UnbindRaster("logqu")
			'END STEP 12b: -------------------------------------------------------------------------------------
		End If
		
		'TEMP Code added 2008 to help w/ debugging
		'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
		'Set pTempRaster = modUtil.ReturnPermanentRaster(pQuRaster, pEnv.OutWorkspace.PathName, strTemp)
		
		'UPGRADE_NOTE: Object pTempLogTOCRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pTempLogTOCRaster = Nothing
		'UPGRADE_NOTE: Object pLogQuRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pLogQuRaster = Nothing
		
		modProgDialog.ProgDialog("Creating Pond Factor GRID...", strTitle, 0, 27, 16, (frmPrj.m_App.hwnd))
		If modProgDialog.g_boolCancel Then
			'STEP 13: ------------------------------------------------------------------------------------------
			'Create pond factor grid
			'UPGRADE_WARNING: Couldn't resolve default property of object g_LandCoverRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pMapAlgebraOp.BindRaster(g_LandCoverRaster, "nu_lulc")
			
			strExpression = strConPondStatement
			
			pPondFactorRaster = pMapAlgebraOp.Execute(strExpression)
			
			pMapAlgebraOp.UnbindRaster("nu_lulc")
			'END STEP 13: ---------------------------------------------------------------------------------------
		End If
		
		'TEMP Code added 2008 to help w/ debugging
		'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
		'Set pTempRaster = modUtil.ReturnPermanentRaster(pPondFactorRaster, pEnv.OutWorkspace.PathName, strTemp)
		
		modProgDialog.ProgDialog("Calculating peak discharge; cubic feet per second...", strTitle, 0, 27, 17, (frmPrj.m_App.hwnd))
		If modProgDialog.g_boolCancel Then
			'STEP 14: -------------------------------------------------------------------------------------------
			'Calculate peak discharge: cubic feet per second
			With pMapAlgebraOp
				'UPGRADE_WARNING: Couldn't resolve default property of object pQuRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(pQuRaster, "qu")
				'UPGRADE_WARNING: Couldn't resolve default property of object g_pCellAreaSqMiRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(g_pCellAreaSqMiRaster, "cellarea_sqmi")
				'UPGRADE_WARNING: Couldn't resolve default property of object g_pRunoffInchRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(g_pRunoffInchRaster, "runoff_in")
				'UPGRADE_WARNING: Couldn't resolve default property of object pPondFactorRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(pPondFactorRaster, "pondfact")
			End With
			
			strExpression = "[qu] * [cellarea_sqmi] * [runoff_in] * [pondfact]"
			
			pQPRaster = pMapAlgebraOp.Execute(strExpression)
			
			With pMapAlgebraOp
				.UnbindRaster("qu")
				.UnbindRaster("cellarea_sqmi")
				.UnbindRaster("runoff_in")
				.UnbindRaster("pondfact")
			End With
			'END STEP 14: ----------------------------------------------------------------------------------------
		End If
		
		'TEMP Code added 2008 to help w/ debugging
		'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
		'Set pTempRaster = modUtil.ReturnPermanentRaster(pQPRaster, pEnv.OutWorkspace.PathName, strTemp)
		
		'UPGRADE_NOTE: Object pPondFactorRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pPondFactorRaster = Nothing
		'UPGRADE_NOTE: Object pQuRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pQuRaster = Nothing
		
		modProgDialog.ProgDialog("Creating cover factor GRID...", strTitle, 0, 27, 18, (frmPrj.m_App.hwnd))
		If modProgDialog.g_boolCancel Then
			'STEP 15: --------------------------------------------------------------------------------------------
			'Cover Factor GRID
			'UPGRADE_WARNING: Couldn't resolve default property of object g_LandCoverRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pMapAlgebraOp.BindRaster(g_LandCoverRaster, "nu_lulc")
			
			strExpression = strConStatement
			
			pCFactorRaster = pMapAlgebraOp.Execute(strExpression)
			
			pMapAlgebraOp.UnbindRaster("nu_lulc")
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
		
		modProgDialog.ProgDialog("Sediment Yield...", strTitle, 0, 27, 19, (frmPrj.m_App.hwnd))
		If modProgDialog.g_boolCancel Then
			'STEP 16 c: -------------------------------------------------------------------------------------------
			'Temp Sediment Yield: Note: MUSLE Exponent inserted to allow for global use, cuz other's will use
			With pMapAlgebraOp
				'UPGRADE_WARNING: Couldn't resolve default property of object g_pRunoffAFRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(g_pRunoffAFRaster, "runoff_af")
				'UPGRADE_WARNING: Couldn't resolve default property of object pQPRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(pQPRaster, "qp")
			End With
			
			strExpression = "Pow(([runoff_af] * [qp]), " & m_dblMUSLEExp & ")"
			
			pHISYTempRaster = pMapAlgebraOp.Execute(strExpression)
			
			With pMapAlgebraOp
				.UnbindRaster("runoff_af")
				.UnbindRaster("qp")
			End With
			'END STEP 16c: ------------------------------------------------------------------------------------------------------
		End If
		
		'TEMP Code added 2008 to help w/ debugging
		'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
		'Set pTempRaster = modUtil.ReturnPermanentRaster(pHISYTempRaster, pEnv.OutWorkspace.PathName, strTemp)
		
		'UPGRADE_NOTE: Object pQPRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pQPRaster = Nothing
		
		modProgDialog.ProgDialog("Sediment Yield...", strTitle, 0, 30, 27, (frmPrj.m_App.hwnd))
		If modProgDialog.g_boolCancel Then
			'STEP 16d: ---------------------------------------------------------------------------------------------
			'Sediment Yield: Note m_dblMUSLEVal now in for universal
			With pMapAlgebraOp
				'UPGRADE_WARNING: Couldn't resolve default property of object pCFactorRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(pCFactorRaster, "cfactor")
				'UPGRADE_WARNING: Couldn't resolve default property of object g_KFactorRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(g_KFactorRaster, "kfactor")
				'UPGRADE_WARNING: Couldn't resolve default property of object g_pLSRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(g_pLSRaster, "lsfactor")
				'UPGRADE_WARNING: Couldn't resolve default property of object pHISYTempRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(pHISYTempRaster, "temp9")
			End With
			
			strExpression = m_dblMUSLEVal & " * ([cfactor] * [kfactor] * [lsfactor] * [temp9])"
			
			pHISYRaster = pMapAlgebraOp.Execute(strExpression)
			
			With pMapAlgebraOp
				.UnbindRaster("cfactor")
				.UnbindRaster("kfactor")
				.UnbindRaster("lsfactor")
				.UnbindRaster("temp9")
			End With
			'END STEP 16b: ----------------------------------------------------------------------------------------
		End If
		
		'TEMP Code added 2008 to help w/ debugging
		'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
		'Set pTempRaster = modUtil.ReturnPermanentRaster(pHISYRaster, pEnv.OutWorkspace.PathName, strTemp)
		
		'UPGRADE_NOTE: Object pCFactorRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pCFactorRaster = Nothing
		'UPGRADE_NOTE: Object pHISYTempRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pHISYTempRaster = Nothing
		
		modProgDialog.ProgDialog("Converting tons to milligrams...", strTitle, 0, 27, 21, (frmPrj.m_App.hwnd))
		If modProgDialog.g_boolCancel Then
			
			'UPGRADE_WARNING: Couldn't resolve default property of object pHISYRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pMapAlgebraOp.BindRaster(pHISYRaster, "sy")
			
			strExpression = "[sy] * 907.184740"
			
			pHISYMGRaster = pMapAlgebraOp.Execute(strExpression)
			
			pMapAlgebraOp.UnbindRaster("sy")
			'END STEP 17b: ----------------------------------------------------------------------------------------
		End If
		
		'TEMP Code added 2008 to help w/ debugging
		'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
		'Set pTempRaster = modUtil.ReturnPermanentRaster(pHISYMGRaster, pEnv.OutWorkspace.PathName, strTemp)
		
		'UPGRADE_NOTE: Object pHISYRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pHISYRaster = Nothing
		
		Dim pClipMusleRaster As ESRI.ArcGIS.Geodatabase.IRaster
		If g_booLocalEffects Then
			
			modProgDialog.ProgDialog("Creating data layer for local effects...", strTitle, 0, 27, 27, (frmPrj.m_App.hwnd))
			If modProgDialog.g_boolCancel Then
				
				'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				strMUSLE = modUtil.GetUniqueName("locmusle", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
				'Added 7/23/04 to account for clip by selected polys functionality
				If g_booSelectedPolys Then
					pClipMusleRaster = modUtil.ClipBySelectedPoly(pHISYMGRaster, g_pSelectedPolyClip, pEnv)
					'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					pPermMUSLERaster = modUtil.ReturnPermanentRaster(pClipMusleRaster, pEnv.OutWorkspace.PathName, strMUSLE)
					'UPGRADE_NOTE: Object pClipMusleRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					pClipMusleRaster = Nothing
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					pPermMUSLERaster = modUtil.ReturnPermanentRaster(pHISYMGRaster, pEnv.OutWorkspace.PathName, strMUSLE)
				End If
				
				pMUSLERasterLocLayer = modUtil.ReturnRasterLayer((frmPrj.m_App), pPermMUSLERaster, "MUSLE Local Effects (mg)")
				pMUSLERasterLocLayer.Renderer = modUtil.ReturnRasterStretchColorRampRender(pMUSLERasterLocLayer, "Brown")
				'UPGRADE_WARNING: Couldn't resolve default property of object pMUSLERasterLocLayer.Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pMUSLERasterLocLayer.Visible = False
				
				'metadata time
				'UPGRADE_WARNING: Couldn't resolve default property of object pMUSLERasterLocLayer.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				g_dicMetadata.Add(pMUSLERasterLocLayer.Name, m_strMusleMetadata)
				
				'UPGRADE_WARNING: Couldn't resolve default property of object pMUSLERasterLocLayer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				g_pGroupLayer.Add(pMUSLERasterLocLayer)
				
				CalcMUSLE = True
				modProgDialog.KillDialog()
				'UPGRADE_NOTE: Object pPermMUSLERaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				pPermMUSLERaster = Nothing
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
		
		modProgDialog.ProgDialog("Calculating the accumulated sediment...", strTitle, 0, 27, 23, (frmPrj.m_App.hwnd))
		If modProgDialog.g_boolCancel Then
			
			pTempflowDir3raster = modUtil.ReturnRaster(g_strFlowDirFilename)
			pFlowAccumOp1 = New ESRI.ArcGIS.SpatialAnalyst.RasterHydrologyOp
			
			pFlowDirRDS1 = pTempflowDir3raster
			pHISYMGRDS = pHISYMGRaster
			
			pEnv = pFlowAccumOp1
			pOutRDS1 = pFlowAccumOp1.FlowAccumulation(pFlowDirRDS1, pHISYMGRDS)
			
			pAccSedHIRaster = pOutRDS1
			
		End If
		
		'TEMP Code added 2008 to help w/ debugging
		'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
		'Set pTempRaster = modUtil.ReturnPermanentRaster(pAccSedHIRaster, pEnv.OutWorkspace.PathName, strTemp)
		
		'UPGRADE_NOTE: Object pTempflowDir3raster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pTempflowDir3raster = Nothing
		
		modProgDialog.ProgDialog("Calculating Total Sediment Mass...", strTitle, 0, 27, 24, (frmPrj.m_App.hwnd))
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
				'UPGRADE_WARNING: Couldn't resolve default property of object pHISYMGRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(pHISYMGRaster, "sy_mg_HI")
				'UPGRADE_WARNING: Couldn't resolve default property of object pAccSedHIRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(pAccSedHIRaster, "accsed_HI")
			End With
			
			'old / 10000000
			strExpression = "[sy_mg_HI] + [accsed_HI]"
			
			pTotSedMassHIRaster = pMapAlgebraOp.Execute(strExpression)
			
			With pMapAlgebraOp
				.UnbindRaster("sy_mg_HI")
				.UnbindRaster("accsed_HI")
			End With
			'END STEP 20HI: -------------------------------------------------------------------------------------------
		End If
		
		'TEMP Code added 2008 to help w/ debugging
		'strTemp = modUtil.GetUniqueName("mslstep", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
		'Set pTempRaster = modUtil.ReturnPermanentRaster(pTotSedMassHIRaster, pEnv.OutWorkspace.PathName, strTemp)
		
		'UPGRADE_NOTE: Object pHISYMGRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pHISYMGRaster = Nothing
		'UPGRADE_NOTE: Object pAccSedHIRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pAccSedHIRaster = Nothing
		
		
		modProgDialog.ProgDialog("Adding Sediment Mass to Group Layer...", strTitle, 0, 27, 25, (frmPrj.m_App.hwnd))
		Dim pClipMusleMassRaster As ESRI.ArcGIS.Geodatabase.IRaster
		If modProgDialog.g_boolCancel Then
			'STEP 21: Created the Sediment Mass Raster layer and add to Group Layer -----------------------------------
			'Get a unique name for MUSLE and return the permanently made raster
			'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			strMUSLE = modUtil.GetUniqueName("MUSLEmass", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
			
			'Clip to selected polys if chosen
			If g_booSelectedPolys Then
				pClipMusleMassRaster = modUtil.ClipBySelectedPoly(pTotSedMassHIRaster, g_pSelectedPolyClip, pEnv)
				'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pPermTotSedConcHIraster = modUtil.ReturnPermanentRaster(pClipMusleMassRaster, pEnv.OutWorkspace.PathName, strMUSLE)
				'UPGRADE_NOTE: Object pClipMusleMassRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				pClipMusleMassRaster = Nothing
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pPermTotSedConcHIraster = modUtil.ReturnPermanentRaster(pTotSedMassHIRaster, pEnv.OutWorkspace.PathName, strMUSLE)
			End If
			
			'Now create the MUSLE layer
			pMUSLERasterLayer = modUtil.ReturnRasterLayer((frmPrj.m_App), pPermTotSedConcHIraster, "MUSLE Sediment Mass (kg)")
			pMUSLERasterLayer.Renderer = modUtil.ReturnRasterStretchColorRampRender(pMUSLERasterLayer, "Brown")
			'UPGRADE_WARNING: Couldn't resolve default property of object pMUSLERasterLayer.Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pMUSLERasterLayer.Visible = False
			
			'Metadata:
			'UPGRADE_WARNING: Couldn't resolve default property of object pMUSLERasterLayer.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			g_dicMetadata.Add(pMUSLERasterLayer.Name, m_strMusleMetadata)
			
			'Add the MUSLE Layer to the final group layer
			'UPGRADE_WARNING: Couldn't resolve default property of object pMUSLERasterLayer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			g_pGroupLayer.Add(pMUSLERasterLayer)
			
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
		
		modProgDialog.KillDialog()
		
		'Clean up all existing rasters that have not yet been destroyed.
		'UPGRADE_NOTE: Object pMapAlgebraOp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pMapAlgebraOp = Nothing
		'UPGRADE_NOTE: Object pFlowAccumOp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFlowAccumOp = Nothing
		'UPGRADE_NOTE: Object pFlowAccumOp1 may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFlowAccumOp1 = Nothing
		'UPGRADE_NOTE: Object pFlowDirRDS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFlowDirRDS = Nothing
		'UPGRADE_NOTE: Object pLiterRDS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pLiterRDS = Nothing
		'UPGRADE_NOTE: Object pOutRDS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pOutRDS = Nothing
		'UPGRADE_NOTE: Object pFlowDirRDS1 may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFlowDirRDS1 = Nothing
		'UPGRADE_NOTE: Object pHISYMGRDS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pHISYMGRDS = Nothing
		'UPGRADE_NOTE: Object pOutRDS1 may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pOutRDS1 = Nothing
		'UPGRADE_NOTE: Object clsPrecip may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		clsPrecip = Nothing
		'UPGRADE_NOTE: Object pPermTotSedConcHIraster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pPermTotSedConcHIraster = Nothing
		'UPGRADE_NOTE: Object pPermSedConcRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pPermSedConcRaster = Nothing
		
		Exit Function
		
ErrHandler: 
		If Err.Number = -2147217297 Then 'S.A. constant for User cancelled operation
			modProgDialog.g_boolCancel = False
		ElseIf Err.Number = -2147467259 Then  'S.A. constant for crappy ESRI stupid GRID error
			MsgBox("ArcMap has reached the maximum number of GRIDs allowed in memory.  " & "Please exit N-SPECT and restart ArcMap.", MsgBoxStyle.Information, "Maximum GRID Number Encountered")
			CalcMUSLE = False
			modProgDialog.g_boolCancel = False
			modProgDialog.KillDialog()
		Else
			MsgBox("MUSLE Error: " & Err.Number & " on MUSLE Calculation: " & strExpression)
			CalcMUSLE = False
			modProgDialog.g_boolCancel = False
			modProgDialog.KillDialog()
		End If
	End Function
	
	
	Private Function ConstructConStatment(ByRef rsCF As ADODB.Recordset, ByRef pLCRaster As ESRI.ArcGIS.Geodatabase.IRaster) As String
		'Creates the Cover Factor con statement using the name of the the LandCass Recordset, and the Landclass Raster
		'Returns: String
		'Looks like: con(([nu_lulc] eq 2), 0.000, con((nu_lulc eq 3), 0.030....
		
		Dim strCon As String 'Con statement base
		Dim strParens As String 'String of trailing parens
		Dim strCompleteCon As String 'Concatenate of strCon & strParens
		
		Dim pRasterCol As ESRI.ArcGIS.DataSourcesRaster.IRasterBandCollection
		Dim pBand As ESRI.ArcGIS.DataSourcesRaster.IRasterBand
		Dim pTable As ESRI.ArcGIS.Geodatabase.ITable
		Dim TableExist As Boolean
		Dim pCursor As ESRI.ArcGIS.Geodatabase.ICursor
		Dim pRow As ESRI.ArcGIS.Geodatabase.IRow
		Dim FieldIndex As Short
		Dim booValueFound As Boolean
		Dim i As Short
		
		'STEP 1:  get the records from the database -----------------------------------------------
		rsCF.MoveFirst()
		'End Database stuff
		
		'STEP 2: Raster Values ---------------------------------------------------------------------
		'Now Get the RASTER values
		' Get Rasterband from the incoming raster
		pRasterCol = pLCRaster
		pBand = pRasterCol.Item(0)
		
		'Get the raster table
		pBand.HasTable(TableExist)
		If Not TableExist Then Exit Function
		
		pTable = pBand.AttributeTable
		'Get All rows
		pCursor = pTable.Search(Nothing, True)
		'Init pRow
		pRow = pCursor.NextRow
		
		'Get index of Value Field
		'UPGRADE_WARNING: Couldn't resolve default property of object pTable.FindField. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
			rsCF.MoveFirst()
			
			For i = 0 To rsCF.RecordCount - 1
				'UPGRADE_WARNING: Couldn't resolve default property of object pRow.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If pRow.Value(FieldIndex) = rsCF.Fields("Value").Value Then
					
					booValueFound = True
					
					If strCon = "" Then
						'UPGRADE_WARNING: Couldn't resolve default property of object pRow.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						strCon = "Con(([nu_lulc] eq " & pRow.Value(FieldIndex) & "), " & rsCF.Fields("CoverFactor").Value & ", "
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object pRow.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						strCon = strCon & "Con(([nu_lulc] eq " & pRow.Value(FieldIndex) & "), " & rsCF.Fields("CoverFactor").Value & ", "
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
				rsCF.MoveNext()
			Next i
			
			If booValueFound = False Then
				MsgBox("Error in MUSLE ConstructConStatment Function: Your N-SPECT Land Class Table is missing values found in your landcover GRID dataset.")
				ConstructConStatment = ""
				Exit Function
			Else
				pRow = pCursor.NextRow
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
		'UPGRADE_NOTE: Object pRasterCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRasterCol = Nothing
		'UPGRADE_NOTE: Object pBand may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pBand = Nothing
		'UPGRADE_NOTE: Object pTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pTable = Nothing
		'UPGRADE_NOTE: Object pCursor may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pCursor = Nothing
		'UPGRADE_NOTE: Object pRow may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRow = Nothing
		
	End Function
	
	
	
	Private Function ConstructPondConStatement(ByRef rsCF As ADODB.Recordset, ByRef pLCRaster As ESRI.ArcGIS.Geodatabase.IRaster) As String
		'Creates the Con Statement used in the Pond Factor GRID
		'Returns: String
		'Looks like: con(([nu_lulc] eq 16), 0, con((nu_lulc eq 17), 0...
		
		Dim strCon As String 'Con statement base
		Dim strParens As String 'String of trailing parens
		Dim strCompleteCon As String 'Concatenate of strCon & strParens
		
		Dim pRasterCol As ESRI.ArcGIS.DataSourcesRaster.IRasterBandCollection
		Dim pBand As ESRI.ArcGIS.DataSourcesRaster.IRasterBand
		Dim pTable As ESRI.ArcGIS.Geodatabase.ITable
		Dim TableExist As Boolean
		Dim pCursor As ESRI.ArcGIS.Geodatabase.ICursor
		Dim pRow As ESRI.ArcGIS.Geodatabase.IRow
		Dim FieldIndex As Short
		Dim booValueFound As Boolean
		Dim i As Short
		
		'STEP 1:  get the records from the database -----------------------------------------------
		rsCF.MoveFirst()
		'End Database stuff
		
		'STEP 2: Raster Values ---------------------------------------------------------------------
		'Now Get the RASTER values
		' Get Rasterband from the incoming raster
		pRasterCol = pLCRaster
		pBand = pRasterCol.Item(0)
		
		'Get the raster table
		pBand.HasTable(TableExist)
		If Not TableExist Then Exit Function
		
		pTable = pBand.AttributeTable
		'Get All rows
		pCursor = pTable.Search(Nothing, True)
		'Init pRow
		pRow = pCursor.NextRow
		
		'Get index of Value Field
		'UPGRADE_WARNING: Couldn't resolve default property of object pTable.FindField. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
			rsCF.MoveFirst()
			
			For i = 0 To rsCF.RecordCount - 1
				'UPGRADE_WARNING: Couldn't resolve default property of object pRow.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If pRow.Value(FieldIndex) = rsCF.Fields("Value").Value Then
					booValueFound = True
					Select Case rsCF.Fields("W_WL").Value
						Case 0 'Means the current landclass is NOT Water or Wetland, therefore gets a 1
							If strCon = "" Then
								'UPGRADE_WARNING: Couldn't resolve default property of object pRow.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								strCon = "Con(([nu_lulc] eq " & pRow.Value(FieldIndex) & "), 1, "
							Else
								'UPGRADE_WARNING: Couldn't resolve default property of object pRow.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								strCon = strCon & "Con(([nu_lulc] eq " & pRow.Value(FieldIndex) & "), 1, "
							End If
							
							If strParens = "" Then
								strParens = "1)"
							Else
								strParens = strParens & ")"
							End If
							
						Case 1 'Means the current landclass IS Water or Wetland, therefore gets a 0
							If strCon = "" Then
								'UPGRADE_WARNING: Couldn't resolve default property of object pRow.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								strCon = "Con(([nu_lulc] eq " & pRow.Value(FieldIndex) & "), 0, "
							Else
								'UPGRADE_WARNING: Couldn't resolve default property of object pRow.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
				
				rsCF.MoveNext()
				
			Next i
			
			If booValueFound = False Then
				MsgBox("Error in MUSLEConstructPondConStatement: Your N-SPECT Land Class Table is missing values found in your landcover GRID dataset.")
				ConstructPondConStatement = ""
				Exit Function
			Else
				pRow = pCursor.NextRow
				i = 0
			End If
		Loop 
		
		strCompleteCon = strCon & strParens
		ConstructPondConStatement = strCompleteCon
		
		'Cleanup:
		'Set pLCRaster = Nothing
		'UPGRADE_NOTE: Object pRasterCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRasterCol = Nothing
		'UPGRADE_NOTE: Object pBand may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pBand = Nothing
		'UPGRADE_NOTE: Object pTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pTable = Nothing
		'UPGRADE_NOTE: Object pCursor may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pCursor = Nothing
		'UPGRADE_NOTE: Object pRow may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRow = Nothing
		
	End Function
	
	
	
	Private Function CreateMetadata(ByRef booLocal As Boolean) As String
		
		Dim strHeader As String
		'Dim i As Integer
		'Dim strCFactor As String
		
		'Set up the header w/or without flow direction
		If booLocal = True Then
			strHeader = vbTab & "Input Datasets:" & vbNewLine & vbTab & vbTab & "Hydrologic soils grid: " & g_clsXMLPrjFile.strSoilsHydFileName & vbNewLine & vbTab & vbTab & "Landcover grid: " & g_clsXMLPrjFile.strLCGridFileName & vbNewLine & vbTab & vbTab & "Precipitation grid: " & g_strPrecipFileName & vbNewLine & vbTab & vbTab & "Soils K Factor: " & g_clsXMLPrjFile.strSoilsKFileName & vbNewLine & vbTab & vbTab & "LS Factor grid: " & g_strLSFileName & vbNewLine & g_strLandCoverParameters & vbNewLine 'append the g_strLandCoverParameters that was set up during runoff
		Else
			strHeader = vbTab & "Input Datasets:" & vbNewLine & vbTab & vbTab & "Hydrologic soils grid: " & g_clsXMLPrjFile.strSoilsHydFileName & vbNewLine & vbTab & vbTab & "Landcover grid: " & g_clsXMLPrjFile.strLCGridFileName & vbNewLine & vbTab & vbTab & "Precipitation grid: " & g_strPrecipFileName & vbNewLine & vbTab & vbTab & "Soils K Factor: " & g_clsXMLPrjFile.strSoilsKFileName & vbNewLine & vbTab & vbTab & "LS Factor grid: " & g_strLSFileName & vbNewLine & vbTab & vbTab & "Flow direction grid: " & g_strFlowDirFilename & vbNewLine & g_strLandCoverParameters & vbNewLine
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
End Module