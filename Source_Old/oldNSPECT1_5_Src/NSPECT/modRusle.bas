Attribute VB_Name = "modRusle"
' *************************************************************************************
' *  Perot Systems Government Services
' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
' *  modRUSLE
' *************************************************************************************
' *  Description:  Revised Universal Soil Loss Equation
' *
' *
' *  Called By:  Various
' *************************************************************************************

Option Explicit
Public g_NibbleRaster As IRaster        'Nibble Raster
Public g_DEMTwoCellRaster As IRaster    'Two Cell buffer of the DEM
Public g_RFactorRaster As IRaster       'R Factor Raster
Private m_strRusleMetadata As String    'Metadata holder
Private m_booUsingConstantValue As Boolean 'Boolean value
Private m_dblRFactorConstant As Double  'Constant for R-factor
Private m_strSDRFileName As String      'If user provides own SDR GRid, store path here


Public Function RUSLESetup(strNibbleFileName As String, _
                            strDEMTwoCellFileName As String, _
                            strRFactorFileName As String, _
                            strKfactorFileName As String, _
                            strSDRFileName As String, _
                            strLandClass As String, _
                            Optional dblRFactorConstant As Double) As Boolean
'Sub takes incoming parameters from the project file and then parses them out
'strNibbleFileName: FileName of the nibble GRID
'strDEMTwoCellFileName: FileName of the two cell buffered DEM
'strRFactorFileName: FileName of the R Factor GRID
'strLandClass: Name of the Landclass we're using
    
    'RecordSet to get the coverfactor
    Dim rsCoverFactor As New ADODB.Recordset
    
    'Open Strings
    Dim strCovFactor As String
    Dim strConStatement As String
    Dim strError As String
    Dim strTempLCType As String
    
    'STEP 1: Set the Nibble Raster ----------------------------------------------------------------
    If modUtil.RasterExists(strNibbleFileName) Then
        Set g_NibbleRaster = modUtil.ReturnRaster(strNibbleFileName)
    Else
        strError = "Nibble Raster Does Not Exist: " & strNibbleFileName
    End If
    'END STEP 1: -----------------------------------------------------------------------------------
    
    'STEP 2: Set the DEMTwoCell Raster -------------------------------------------------------------
    If modUtil.RasterExists(strDEMTwoCellFileName) Then
        Set g_DEMTwoCellRaster = modUtil.ReturnRaster(strDEMTwoCellFileName)
    Else
        strError = "DEM Two Cell Buffer Raster Does Not Exist: " & strDEMTwoCellFileName
    End If
    'END STEP 2: -----------------------------------------------------------------------------------
    
    'STEP 3: Set the R Factor Raster ---------------------------------------------------------------
    If Len(strRFactorFileName) > 0 Then
        If modUtil.RasterExists(strRFactorFileName) Then
            Set g_RFactorRaster = modUtil.ReturnRaster(strRFactorFileName)
        Else
            strError = "R Factor Raster Does Not Exist: " & strRFactorFileName
        End If
        m_booUsingConstantValue = False
    Else
        m_booUsingConstantValue = True
        m_dblRFactorConstant = dblRFactorConstant
    End If
    
    'END STEP 3: -----------------------------------------------------------------------------------
    
    'STEP 4: Set the K Factor Raster
    If modUtil.RasterExists(strKfactorFileName) Then
        Set g_KFactorRaster = modUtil.ReturnRaster(strKfactorFileName)
    Else
        strError = "K Factor Raster Does Not Exist: " & strKfactorFileName
    End If
    'END STEP 3: -----------------------------------------------------------------------------------
       
    'Get the landclasses of type strLandClass
    'Check first for temp name
    If Len(g_DictTempNames.Item(strLandClass)) > 0 Then
        strTempLCType = g_DictTempNames.Item(strLandClass)
    Else
        strTempLCType = strLandClass
    End If
    
    strCovFactor = "SELECT LCTYPE.LCTYPEID, LCCLASS.NAME, LCCLASS.VALUE, LCCLASS.COVERFACTOR FROM " & _
        "LCTYPE INNER JOIN LCCLASS ON LCTYPE.LCTYPEID = LCCLASS.LCTYPEID " & _
        "WHERE LCTYPE.NAME LIKE '" & strTempLCType & "' ORDER BY LCCLASS.VALUE"
        
    rsCoverFactor.Open strCovFactor, g_ADOConn, adOpenKeyset, adLockOptimistic
    
    If Len(strError) > 0 Then
        MsgBox strError
        Exit Function
    End If
    
    'Get the con statement for the cover factor calculation
    strConStatement = ConstructConStatment(rsCoverFactor, g_LandCoverRaster)
    
    'Are they using SDR
    m_strSDRFileName = Trim(strSDRFileName)
    
    'Metadata time
    m_strRusleMetadata = CreateMetadata(g_booLocalEffects)

    'Calc rusle using the con
    If CalcRUSLE(strConStatement) Then
        RUSLESetup = True
    Else
        RUSLESetup = False
    End If
    
    
End Function

Private Function ConstructConStatment(rsCF As ADODB.Recordset, pLCRaster As IRaster) As String
'Creates the initial pick statement using the name of the the LandCass [CCAP, for example]
'and the Land Class Raster.  Returns a string

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
    'Get Rasterband from the incoming raster
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
    
    'REMOVED 11/30/2007 in favor of method below.
    'STEP 3: Table values vs. Raster Values Count - if not equal bark --------------------------
'    If pTable.RowCount(Nothing) <> rsCF.RecordCount Then
'        MsgBox "Error: The number of records in your Land Class Table do not match your GRID dataset."
'        Exit Function
'    End If
    
    
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
            MsgBox "Error: Your N-SPECT Land Class Table is missing values found in your landcover GRID dataset."
            ConstructConStatment = ""
            Exit Function
        Else
            Set pRow = pCursor.NextRow
            i = 0
        End If
        
    Loop
    
    'REMOVED 11/30/2007 in favor of method above
'    'STEP 4: Create the strings
'    'Loop through and get all values
'    Do While Not pRow Is Nothing
'        If pRow.Value(FieldIndex) = rsCF!Value Then
'            If strCon = "" Then
'                strCon = "Con(([nu_lulc] eq " & pRow.Value(FieldIndex) & "), " & rsCF!CoverFactor & ", "
'            Else
'                strCon = strCon & "Con(([nu_lulc] eq " & pRow.Value(FieldIndex) & "), " & rsCF!CoverFactor & ", "
'            End If
'
'            If strParens = "" Then
'                strParens = "-9999)"
'            Else
'                strParens = strParens & ")"
'            End If
'
'            rsCF.MoveNext
'            Set pRow = pCursor.NextRow
'
'        Else
'            MsgBox "Values in table LCClass table not equal to values in landclass dataset."
'            ConstructConStatment = ""
'            Exit Function
'        End If
'    Loop
    
    
    strCompleteCon = strCon & strParens
    Debug.Print strCompleteCon
        
    ConstructConStatment = strCompleteCon
    
    'Cleanup:
    'Set pLCRaster = Nothing
    Set pRasterCol = Nothing
    Set pBand = Nothing
    Set pTable = Nothing
    Set pCursor = Nothing
    Set pRow = Nothing
    
End Function

Private Function CalcRUSLE(strConStatement As String) As Boolean
    
    Dim pRasterProps As IRasterProps            'Raster props
    Dim pLandSampleRaster As IRaster            'LC Cover sampleized
    Dim pCFactorRaster As IRaster               'C Factor
    Dim pSoilLossRaster As IRaster              'Soil Loss
    Dim pSoilLossAcres As IRaster               'Soil Loss Acres
    Dim pZSedDelRaster As IRaster               'And I quote, Dave's Whacky Sediment Delivery Ratio
    Dim pCellAreaKMRaster As IRaster            'Cell Area KM
    Dim pDASedDelRaster As IRaster              'da_sed_del
    Dim pTemp3Raster As IRaster                 'Temp3
    Dim pTemp4Raster As IRaster                 'Temp4
    Dim pTemp5Raster As IRaster                 'Temp5
    Dim pTemp6Raster As IRaster                 'Temp6
    Dim pSDRRaster As IRaster                   'SDR
    Dim pSedYieldRaster As IRaster              'Sediment Yield
    Dim pAccumSedRaster As IRaster              'Accumulated Sed Raster
    Dim pTotalAccumSedRaster As IRaster         'Total accumulated sediment raster
    Dim pPermAccumSedRaster As IRaster          'Permanent accumulated sediment raster
    Dim pConcSedRaster As IRaster               'Concentration sed raster
    Dim pPermConcSedRaster As IRaster           'Permanent sediment concentration raster
    Dim pRUSLERasterLayer As IRasterLayer       'RUSLE accumulated raster layer
    Dim pRUSLEConcRasterLayer As IRasterLayer   'RUSLE Concentration raster layer
    Dim pPermRUSLELocRaster As IRaster          'RUSLE Local Effects raster
    Dim pRUSLELocRasterLayer As IRasterLayer    'RUSLE Local Effects Raster Layer
        
    Dim pEnv As IRasterAnalysisEnvironment

    Dim pTempRaster As IRaster  'Temp raster added 8/29/06
    'Create Map Algebra Operator
    Dim pMapAlgebraOp As IMapAlgebraOp      'Workhorse
    Dim strTemp As String
    
    'String to hold calculations
    Dim strExpression As String
    Const strTitle As String = "Processing RUSLE Calculation..."
    Dim strOutYield As String
    Dim strOutConc As String
    
    'Set the enviornment stuff
    Set pEnv = g_pSpatEnv
    Set pMapAlgebraOp = New RasterMapAlgebraOp
    Set pEnv = pMapAlgebraOp

On Error GoTo ErrHandler

    modProgDialog.ProgDialog "Checking landcover cell size...", strTitle, 0, 13, 1, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
        'Step 1a: ----------------------------------------------------------------------------
        'Make sure LandCover is in the same cellsize as the global environment
        Set pRasterProps = g_LandCoverRaster
        
        If pRasterProps.MeanCellSize.X <> g_dblCellSize Then
            pMapAlgebraOp.BindRaster g_LandCoverRaster, "landcover"
    
            strExpression = "resample([landcover], " & CStr(g_dblCellSize) & ")"
    
            Set pLandSampleRaster = pMapAlgebraOp.Execute(strExpression)
        
            pMapAlgebraOp.UnbindRaster "landcover"
        Else
            Set pLandSampleRaster = g_LandCoverRaster
        End If
        
    End If

    modProgDialog.ProgDialog "Calculating Cover Magagement Factor GRID...", strTitle, 0, 13, 2, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then

        'Step 1: CREATE Management cover Factor GRID ---------------------------------------------
        pMapAlgebraOp.BindRaster pLandSampleRaster, "nu_lulc"
        
        strExpression = strConStatement
        
        Set pCFactorRaster = pMapAlgebraOp.Execute(strExpression)
        
        pMapAlgebraOp.UnbindRaster "nu_lulc"
        'END STEP 1: ------------------------------------------------------------------------------
    End If
    
    'TEMP Code added 8/29/06 to help w/ Hawaii's checking of the RUSLE equation
    'strTemp = modUtil.GetUniqueName("cfactor", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pCFactorRaster, pEnv.OutWorkspace.PathName, strTemp)
    
    Set pLandSampleRaster = Nothing
    
    modProgDialog.ProgDialog "Solving RUSLE Equation...", strTitle, 0, 13, 3, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
        'STEP 2: SOLVE RUSLE EQUATION -------------------------------------------------------------
        
        If Not m_booUsingConstantValue Then
            If g_RFactorRaster Is Nothing Then
                MsgBox "Rfactor is missing"
            End If
        End If
        
        If g_pLSRaster Is Nothing Then
            MsgBox "LS Grid is missing."
        End If
        
        If g_KFactorRaster Is Nothing Then
            MsgBox "Kfactor is missing."
        End If
        
        If pCFactorRaster Is Nothing Then
            MsgBox "Cfactor is missing."
        End If
        
        With pMapAlgebraOp
            .BindRaster g_pLSRaster, "lsfactor"
            .BindRaster g_KFactorRaster, "kfactor"
            .BindRaster pCFactorRaster, "cfactor"
        End With
        
        If Not m_booUsingConstantValue Then     'If not using a constant
            pMapAlgebraOp.BindRaster g_RFactorRaster, "rfactor"
            strExpression = "[rfactor] * [kfactor] * [lsfactor] * [cfactor]"
        Else 'if using a constant
            strExpression = m_dblRFactorConstant & " * [kfactor] * [lsfactor] * [cfactor]"
        End If
                
        Set pSoilLossRaster = pMapAlgebraOp.Execute(strExpression)
        
        If Not m_booUsingConstantValue Then
            pMapAlgebraOp.UnbindRaster "rfactor"
        End If
        
        With pMapAlgebraOp
            .UnbindRaster "lsfactor"
            .UnbindRaster "kfactor"
            .UnbindRaster "cfactor"
        End With
        'END STEP 2: -------------------------------------------------------------------------------
    End If
    
    Set pCFactorRaster = Nothing
    
    'TEMP Code added 8/29/06 to help w/ Hawaii's checking of the RUSLE equation
    'strTemp = modUtil.GetUniqueName("soilloss", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pSoilLossRaster, pEnv.OutWorkspace.PathName, strTemp)
    
    modProgDialog.ProgDialog "Converting Soil Loss...", strTitle, 0, 13, 4, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
        'STEP 3: DERIVE ACCUMULATED POLLUTANT ------------------------------------------------------
        With pMapAlgebraOp
            .BindRaster pSoilLossRaster, "soil_loss"
        End With
        
        strExpression = "(Pow(" & g_dblCellSize & ", 2) * 0.000247104369) * [soil_loss]"
        
        Set pSoilLossAcres = pMapAlgebraOp.Execute(strExpression)
        
        With pMapAlgebraOp
            .UnbindRaster "soil_loss"
        End With
        'END STEP 3: ------------------------------------------------------------------------------
    End If
    
    'TEMP Code added 8/29/06 to help w/ Hawaii's checking of the RUSLE equation
    'strTemp = modUtil.GetUniqueName("slacres", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pSoilLossAcres, pEnv.OutWorkspace.PathName, strTemp)
    
    Set pSoilLossRaster = Nothing
    
    '***********************************************
    'BEGIN SDR CODE......
    '***********************************************
    If Len(Trim(m_strSDRFileName)) = 0 Then
        modProgDialog.ProgDialog "Calculating Relief-Length Ratio for Sediment Delivery...", strTitle, 0, 13, 5, frmPrj.m_App.hwnd
        If modProgDialog.g_boolCancel Then
            'STEP 4: DAVE'S WACKY CALCULATION OF RELIEF-LENGTH RATIO FOR SEDIMENT DELIVERY RATIO-------
            With pMapAlgebraOp
                .BindRaster g_NibbleRaster, "fdrnib"
                .BindRaster g_DEMTwoCellRaster, "dem_2b"
            End With
            
            strExpression = "Con(([fdrnib] ge 0.5 and [fdrnib] lt 1.5), (([dem_2b] - [dem_2b](1,0)) / (" & g_dblCellSize & " * 0.001))," & _
                    "Con(([fdrnib] ge 1.5 and [fdrnib] lt 3.0), (([dem_2b] - [dem_2b](1,1)) / (" & g_dblCellSize & " * 0.0014142))," & _
                    "Con(([fdrnib] ge 3.0 and [fdrnib] lt 6.0), (([dem_2b] - [dem_2b](0,1)) / (" & g_dblCellSize & " * 0.001))," & _
                    "Con(([fdrnib] ge 6.0 and [fdrnib] lt 12.0), (([dem_2b] - [dem_2b](-1,1)) / (" & g_dblCellSize & " * 0.0014142))," & _
                    "Con(([fdrnib] ge 12.0 and [fdrnib] lt 24.0), (([dem_2b] - [dem_2b](-1,0)) / (" & g_dblCellSize & " * 0.001))," & _
                    "Con(([fdrnib] ge 24.0 and [fdrnib] lt 48.0), (([dem_2b] - [dem_2b](-1,-1)) / (" & g_dblCellSize & " * 0.0014142))," & _
                    "Con(([fdrnib] ge 48.0 and [fdrnib] lt 96.0), (([dem_2b] - [dem_2b](0,-1)) / (" & g_dblCellSize & " * 0.001))," & _
                    "Con(([fdrnib] ge 96.0 and [fdrnib] lt 192.0), (([dem_2b] - [dem_2b](1,-1)) / (" & g_dblCellSize & " * 0.0014142))," & _
                    "Con(([fdrnib] ge 192.0 and [fdrnib] le 255.0), (([dem_2b] - [dem_2b](1,0)) / (" & g_dblCellSize & " * 0.001))," & _
                    "0.1)))))))))"
    
            Set pZSedDelRaster = pMapAlgebraOp.Execute(strExpression)
            
            With pMapAlgebraOp
                .UnbindRaster "fdrnib"
                .UnbindRaster "dem_2b"
            End With
            'END STEP 4: ------------------------------------------------------------------------------
        End If
        
        'TEMP Code added 8/29/06 to help w/ Hawaii's checking of the RUSLE equation
        'strTemp = modUtil.GetUniqueName("rellen", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        'Set pTempRaster = modUtil.ReturnPermanentRaster(pZSedDelRaster, pEnv.OutWorkspace.PathName, strTemp)
        
        modProgDialog.ProgDialog "Calculating Sediment Delivery Ratio...", strTitle, 0, 13, 6, frmPrj.m_App.hwnd
        If modProgDialog.g_boolCancel Then
            'STEP 5: CALCULATE SEDIMENT DELIVERY RATIO ------------------------------------------------
            With pMapAlgebraOp
                .BindRaster g_pDEMRaster, "DEM"
            End With
            
            'NOTE: Original equation is cellarea_km = ([cellsize] / 1000).  To get cell_size I do a CON on the DEM.
            'Notice the float...if you don't do that, screw ville...GRID comes back = 0
            strExpression = "(float(con([DEM] >= 0, " & g_dblCellSize & ", 0))) / 1000"
            
            Set pCellAreaKMRaster = pMapAlgebraOp.Execute(strExpression)
            
            With pMapAlgebraOp
                .UnbindRaster "DEM"
            End With
            'END STEP 5: -------------------------------------------------------------------------------
        End If
        
        'TEMP Code added 8/29/06 to help w/ Hawaii's checking of the RUSLE equation
        'strTemp = modUtil.GetUniqueName("sedone", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        'Set pTempRaster = modUtil.ReturnPermanentRaster(pCellAreaKMRaster, pEnv.OutWorkspace.PathName, strTemp)
        
        modProgDialog.ProgDialog "Step 2 Sediment Delivery Ratio...", strTitle, 0, 13, 7, frmPrj.m_App.hwnd
        If modProgDialog.g_boolCancel Then
            'STEP 6: -----------------------------------------------------------------------------------
            With pMapAlgebraOp
                .BindRaster pCellAreaKMRaster, "cellarea_km"
            End With
            
            strExpression = "Pow([cellarea_km], 2)"
            
            Set pDASedDelRaster = pMapAlgebraOp.Execute(strExpression)
            
            With pMapAlgebraOp
                .UnbindRaster "cellarea_km"
            End With
            'END STEP 6: ------------------------------------------------------------------------------
        End If
        
        Set pCellAreaKMRaster = Nothing
        
        'TEMP Code added 8/29/06 to help w/ Hawaii's checking of the RUSLE equation
        'strTemp = modUtil.GetUniqueName("sedtwo", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        'Set pTempRaster = modUtil.ReturnPermanentRaster(pDASedDelRaster, pEnv.OutWorkspace.PathName, strTemp)
        
        modProgDialog.ProgDialog "Step 3 Sediment Delivery Ratio...", strTitle, 0, 13, 8, frmPrj.m_App.hwnd
        If modProgDialog.g_boolCancel Then
        
            'STEP 7: temp3 = Pow([da_sed_del], -0.998) ---------------------------------------------------
            With pMapAlgebraOp
                .BindRaster pDASedDelRaster, "da_sed_del"
            End With
            
            strExpression = "Pow([da_sed_del], -0.0998)"
            
            Set pTemp3Raster = pMapAlgebraOp.Execute(strExpression)
            
            With pMapAlgebraOp
                .UnbindRaster "da_sed_del"
            End With
            'END STEP 7: --------------------------------------------------------------------------------
        End If
        
        'TEMP Code added 8/29/06 to help w/ Hawaii's checking of the RUSLE equation
        'strTemp = modUtil.GetUniqueName("sedthr", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        'Set pTempRaster = modUtil.ReturnPermanentRaster(pTemp3Raster, pEnv.OutWorkspace.PathName, strTemp)
            
        Set pDASedDelRaster = Nothing
        
        modProgDialog.ProgDialog "Step 4 Sediment Delivery Ratio...", strTitle, 0, 13, 9, frmPrj.m_App.hwnd
        If modProgDialog.g_boolCancel Then
        
            'STEP 8: temp4 = Pow([zl_sed_del], -0.0998) ---------------------------------------------------
            With pMapAlgebraOp
                .BindRaster pZSedDelRaster, "zl_sed_del"
            End With
            
            strExpression = "Pow([zl_sed_del], 0.3629)"
            
            Set pTemp4Raster = pMapAlgebraOp.Execute(strExpression)
            
            With pMapAlgebraOp
                .UnbindRaster "zl_sed_del"
            End With
            'END STEP 8: --------------------------------------------------------------------------------
        End If
        
        'TEMP Code added 8/29/06 to help w/ Hawaii's checking of the RUSLE equation
        'strTemp = modUtil.GetUniqueName("sedfr", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        'Set pTempRaster = modUtil.ReturnPermanentRaster(pTemp4Raster, pEnv.OutWorkspace.PathName, strTemp)
            
        Set pZSedDelRaster = Nothing
        
        modProgDialog.ProgDialog "Step 5 Sediment Delivery Ratio...", strTitle, 0, 13, 10, frmPrj.m_App.hwnd
        If modProgDialog.g_boolCancel Then
        
            'STEP 9: temp5 = Pow([scsgrid100], 5.444) ---------------------------------------------------
            With pMapAlgebraOp
                .BindRaster g_pSCS100Raster, "scsgrid100"
            End With
            
            strExpression = "Pow([scsgrid100], 5.444)"
            
            Set pTemp5Raster = pMapAlgebraOp.Execute(strExpression)
            
            With pMapAlgebraOp
                .UnbindRaster "scsgrid100"
            End With
            'END STEP 9: --------------------------------------------------------------------------------
        End If
        
        'TEMP Code added 8/29/06 to help w/ Hawaii's checking of the RUSLE equation
        'strTemp = modUtil.GetUniqueName("sedfv", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        'Set pTempRaster = modUtil.ReturnPermanentRaster(pTemp4Raster, pEnv.OutWorkspace.PathName, strTemp)
         
        modProgDialog.ProgDialog "Step 6 Sediment Delivery Ratio...", strTitle, 0, 13, 11, frmPrj.m_App.hwnd
        If modProgDialog.g_boolCancel Then
        
            'STEP 9: temp6 = 1.366 * [temp2] * [temp3] * [temp4] * [temp5] -------------------------------
            With pMapAlgebraOp
                .BindRaster pTemp3Raster, "temp3"
                .BindRaster pTemp4Raster, "temp4"
                .BindRaster pTemp5Raster, "temp5"
            End With
            
            strExpression = "1.366 * (Pow(10, -11)) * [temp3] * [temp4] * [temp5]"
            
            Set pTemp6Raster = pMapAlgebraOp.Execute(strExpression)
            
            With pMapAlgebraOp
                .UnbindRaster "temp3"
                .UnbindRaster "temp4"
                .UnbindRaster "temp5"
            End With
            'END STEP 9: --------------------------------------------------------------------------------
        End If
        
        'TEMP Code added 8/29/06 to help w/ Hawaii's checking of the RUSLE equation
        'strTemp = modUtil.GetUniqueName("sedsx", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        'Set pTempRaster = modUtil.ReturnPermanentRaster(pTemp4Raster, pEnv.OutWorkspace.PathName, strTemp)
        
        Set pTemp3Raster = Nothing
        Set pTemp4Raster = Nothing
        Set pTemp5Raster = Nothing
        
        modProgDialog.ProgDialog "Final Calculation Sediment Delivery Ratio...", strTitle, 0, 13, 12, frmPrj.m_App.hwnd
        If modProgDialog.g_boolCancel Then
        
            'STEP 10:
            With pMapAlgebraOp
                .BindRaster pTemp6Raster, "temp6"
            End With
            
            strExpression = "Con(([temp6] gt 1), 1, [temp6])"
            
            Set pSDRRaster = pMapAlgebraOp.Execute(strExpression)
            
            With pMapAlgebraOp
                .UnbindRaster "temp6"
            End With
            'END STEP 10: --------------------------------------------------------------------------------
        End If
        
        'TEMP Code added 8/29/06 to help w/ Hawaii's checking of the RUSLE equation
        'strTemp = modUtil.GetUniqueName("sedfin", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        'Set pTempRaster = modUtil.ReturnPermanentRaster(pSDRRaster, pEnv.OutWorkspace.PathName, strTemp)
            
        Set pTemp6Raster = Nothing
    
    Else
        Set pSDRRaster = modUtil.ReturnRaster(m_strSDRFileName)
    End If
    '********************************************************************
    'END SDR CALC
    '********************************************************************
    
    modProgDialog.ProgDialog "Applying Sediment Delivery Ratio...", strTitle, 0, 13, 13, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
    
        'STEP 11: sed_yield = [soil_loss_ac] * [sdr] -------------------------------------------------
        With pMapAlgebraOp
            .BindRaster pSDRRaster, "sdr"
            .BindRaster pSoilLossAcres, "soil_loss_ac"
        End With
        
        'local effects
        strExpression = "([soil_loss_ac] * [sdr]) * 907.18474"
        
        Set pSedYieldRaster = pMapAlgebraOp.Execute(strExpression)
        
        With pMapAlgebraOp
            .UnbindRaster "sdr"
            .UnbindRaster "soil_loss_ac"
        End With
        'END STEP 11: --------------------------------------------------------------------------------
    End If
    
    'TEMP Code added 8/29/06 to help w/ Hawaii's checking of the RUSLE equation
    'strTemp = modUtil.GetUniqueName("sedyed", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    'Set pTempRaster = modUtil.ReturnPermanentRaster(pSedYieldRaster, pEnv.OutWorkspace.PathName, strTemp)
    
    Set pSDRRaster = Nothing
    Set pSoilLossAcres = Nothing
    
    If g_booLocalEffects Then
        
        modProgDialog.ProgDialog "Creating data layer for local effects...", strTitle, 0, 13, 13, frmPrj.m_App.hwnd
        If modProgDialog.g_boolCancel Then
                       
            'STEP 12: Local Effects -------------------------------------------------
            strOutYield = modUtil.GetUniqueName("locrusle", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
            'Added 7/23/04 to account for clip by selected polys functionality
            If g_booSelectedPolys Then
                Dim pClipLocRusleRaster As IRaster
                Set pClipLocRusleRaster = modUtil.ClipBySelectedPoly(pSedYieldRaster, g_pSelectedPolyClip, pEnv)
                Set pPermRUSLELocRaster = modUtil.ReturnPermanentRaster(pClipLocRusleRaster, pEnv.OutWorkspace.PathName, strOutYield)
                Set pClipLocRusleRaster = Nothing
            Else
                Set pPermRUSLELocRaster = modUtil.ReturnPermanentRaster(pSedYieldRaster, pEnv.OutWorkspace.PathName, strOutYield)
            End If
            
            Set pRUSLELocRasterLayer = modUtil.ReturnRasterLayer(frmPrj.m_App, pPermRUSLELocRaster, "Sediment Local Effects (mg)")
            Set pRUSLELocRasterLayer.Renderer = modUtil.ReturnRasterStretchColorRampRender(pRUSLELocRasterLayer, "Brown")
            pRUSLELocRasterLayer.Visible = False
            
            'Metadata
            g_dicMetadata.Add pRUSLELocRasterLayer.Name, m_strRusleMetadata
                       
            'Add to group layer
            g_pGroupLayer.Add pRUSLELocRasterLayer
            
            CalcRUSLE = True
            modProgDialog.KillDialog
            Set pPermRUSLELocRaster = Nothing
            Set pSedYieldRaster = Nothing
            Exit Function
           
        End If
        
    End If
    
     
    modProgDialog.ProgDialog "Calculating Accumulated Sediment...", strTitle, 0, 13, 13, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
    
        'STEP 12: accum_sed = flowaccumulation([flowdir], [sedyield]) -------------------------------------------------
        With pMapAlgebraOp
            .BindRaster g_pFlowDirRaster, "flowdir"
            .BindRaster pSedYieldRaster, "sedyield"
        End With
        
        strExpression = "flowaccumulation([flowdir], [sedyield], FLOAT)"
        
        Set pAccumSedRaster = pMapAlgebraOp.Execute(strExpression)
        
        With pMapAlgebraOp
            .UnbindRaster "flowdir"
            .UnbindRaster "sedyield"
        End With
        'END STEP 12: --------------------------------------------------------------------------------
    End If
    
    modProgDialog.ProgDialog "Calculating Total Accumulated Sediment...", strTitle, 0, 13, 13, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
    
        'STEP 13: accum_sed = flowaccumulation([flowdir], [sedyield]) -------------------------------------------------
        With pMapAlgebraOp
            .BindRaster pAccumSedRaster, "accumsed"
            .BindRaster pSedYieldRaster, "sedyield"
        End With
        
        strExpression = "[accumsed] + [sedyield]"
        
        Set pTotalAccumSedRaster = pMapAlgebraOp.Execute(strExpression)
        
        With pMapAlgebraOp
            .UnbindRaster "accumsed"
            .UnbindRaster "sedyield"
        End With
        'END STEP 13: --------------------------------------------------------------------------------
    End If
    
    Set pSedYieldRaster = Nothing
    
    modProgDialog.ProgDialog "Adding accumulated sediment layer to the data group layer...", strTitle, 0, 13, 13, frmPrj.m_App.hwnd
    
    If modProgDialog.g_boolCancel Then
        strOutYield = modUtil.GetUniqueName("RUSLE", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        
        'Clip to selected polys if chosen
        If g_booSelectedPolys Then
            Dim pClipRUSLERaster As IRaster
            Set pClipRUSLERaster = modUtil.ClipBySelectedPoly(pTotalAccumSedRaster, g_pSelectedPolyClip, pEnv)
            Set pPermAccumSedRaster = modUtil.ReturnPermanentRaster(pClipRUSLERaster, pEnv.OutWorkspace.PathName, strOutYield)
            Set pClipRUSLERaster = Nothing
        Else
            Set pPermAccumSedRaster = modUtil.ReturnPermanentRaster(pTotalAccumSedRaster, pEnv.OutWorkspace.PathName, strOutYield)
        End If
        
        Set pRUSLERasterLayer = modUtil.ReturnRasterLayer(frmPrj.m_App, pPermAccumSedRaster, "Accumulated Sediment (kg)")
        Set pRUSLERasterLayer.Renderer = modUtil.ReturnRasterStretchColorRampRender(pRUSLERasterLayer, "Brown")
        pRUSLERasterLayer.Visible = False
        
        'Metadata
        g_dicMetadata.Add pRUSLERasterLayer.Name, m_strRusleMetadata
        
        g_pGroupLayer.Add pRUSLERasterLayer
    End If
    
'***********************************************************
'Sediment concentration removed 11/20/07 per change request
'***********************************************************
'    modProgDialog.ProgDialog "Calculating Sediment Concentration...", strTitle, 0, 13, 13, frmPrj.m_App.hWnd
'    If modProgDialog.g_boolCancel Then
'
'        'STEP 13: accum_sed = flowaccumulation([flowdir], [sedyield]) -------------------------------------------------
'        With pMapAlgebraOp
'            .BindRaster g_pRunoffRaster, "runoff"
'            .BindRaster pAccumSedRaster, "accumsed"
'        End With
'
'        strExpression = "[accumsed] / [runoff]"
'
'        Set pConcSedRaster = pMapAlgebraOp.Execute(strExpression)
'
'        With pMapAlgebraOp
'            .UnbindRaster "runoff"
'            .UnbindRaster "accumsed"
'        End With
'        'END STEP 13: --------------------------------------------------------------------------------
'    End If
'    If modProgDialog.g_boolCancel Then
'        strOutConc = modUtil.GetUniqueName("RUSLE", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
'
'        'Clip to selected polys if chosen
'        If g_booSelectedPolys Then
'            Dim pClipRUSLEConcRaster As IRaster
'            Set pClipRUSLEConcRaster = modUtil.ClipBySelectedPoly(pConcSedRaster, g_pSelectedPolyClip, pEnv)
'            Set pPermConcSedRaster = modUtil.ReturnPermanentRaster(pClipRUSLEConcRaster, pEnv.OutWorkspace.PathName, strOutConc)
'            Set pClipRUSLEConcRaster = Nothing
'        Else
'            Set pPermConcSedRaster = modUtil.ReturnPermanentRaster(pConcSedRaster, pEnv.OutWorkspace.PathName, strOutConc)
'        End If
'
'        Set pRUSLEConcRasterLayer = modUtil.ReturnRasterLayer(frmPrj.m_App, pPermConcSedRaster, "Sediment Concentration (kg/L)")
'        Set pRUSLEConcRasterLayer.Renderer = modUtil.ReturnRasterStretchColorRampRender(pRUSLEConcRasterLayer, "Brown")
'        pRUSLEConcRasterLayer.Visible = False
'
'        g_dicMetadata.Add pRUSLEConcRasterLayer.Name, m_strRusleMetadata
'
'        g_pGroupLayer.Add pRUSLEConcRasterLayer
'    End If
'*****************************************************************
'End Remove
'*****************************************************************
      
    CalcRUSLE = True
    
    modProgDialog.KillDialog
         
    'Cleanup
    Set pSDRRaster = Nothing                   'SDR
    Set pMapAlgebraOp = Nothing
    Set pConcSedRaster = Nothing
    Set pAccumSedRaster = Nothing
    Set pPermConcSedRaster = Nothing
    Set pPermAccumSedRaster = Nothing
    Set pEnv = Nothing

Exit Function

ErrHandler:
    If Err.Number = -2147217297 Then 'User cancelled operation
        modProgDialog.g_boolCancel = False
        CalcRUSLE = False
    ElseIf Err.Number = -2147467259 Then
        MsgBox "ArcMap has reached the maximum number of GRIDs allowed in memory.  " & _
        "Please exit N-SPECT and restart ArcMap.", vbInformation, "Maximum GRID Number Encountered"
        modProgDialog.g_boolCancel = False
        modProgDialog.KillDialog
        CalcRUSLE = False
    Else
        MsgBox "RUSLE Error: " & Err.Number & " on RUSLE Calculation: " & strExpression
        MsgBox Err.Number & ": " & Err.Description
        modProgDialog.g_boolCancel = False
        modProgDialog.KillDialog
        Screen.MousePointer = vbNormal
        CalcRUSLE = False
    End If

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
                vbTab & vbTab & "R Factor grid: " & g_clsXMLPrjFile.strRainGridFileName & vbNewLine & _
                g_strLandCoverParameters & vbNewLine 'append the g_strLandCoverParameters that was set up during runoff
    Else
        strHeader = vbTab & "Input Datasets:" & vbNewLine & _
                    vbTab & vbTab & "Hydrologic soils grid: " & g_clsXMLPrjFile.strSoilsHydFileName & vbNewLine & _
                    vbTab & vbTab & "Landcover grid: " & g_clsXMLPrjFile.strLCGridFileName & vbNewLine & _
                    vbTab & vbTab & "Precipitation grid: " & g_strPrecipFileName & vbNewLine & _
                    vbTab & vbTab & "Soils K Factor: " & g_clsXMLPrjFile.strSoilsKFileName & vbNewLine & _
                    vbTab & vbTab & "LS Factor grid: " & g_strLSFileName & vbNewLine & _
                    vbTab & vbTab & "R Factor grid: " & g_clsXMLPrjFile.strRainGridFileName & vbNewLine & _
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




