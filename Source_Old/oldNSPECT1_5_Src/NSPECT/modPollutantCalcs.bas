Attribute VB_Name = "modPollutantCalcs"
' *************************************************************************************
' *  Perot Systems Government Services
' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
' *  modPollutantCalcs
' *************************************************************************************
' *  Description:  Callculation of pollutant concentration
' *     * Sub PollutantConcentrationSetup: called from frmPrj, sort of the main sub in
' *       this module.  Uses a clsXMLPollutantItem and the landclass to get things started.
' *       Calls: ConstructConStatement, CalcPollutantConcentration
' *     * Function ConstructConStatment: Constructs the initial Con statement in the pollutant
' *       concentration calculations.
' *     * Sub CalcPollutantConcentration: The big workhorse. Contains all the map algebra that
' *       gets this turkey finished
' *
' *  Called By:  frmPrj
' *************************************************************************************

Private m_strPollName As String             'Mod level variable for name of pollutant being used
Private m_strWQName As String               'Mod level variable for name of water quality standard
Private m_strColor As String                'Mod level variable holding the string of the pollutant color
Private m_strPollCoeffMetadata As String    'Variable to hold coeffs for use in metadata

' Constant used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "modPollutantCalcs.bas"
Private m_ParentHWND As Long    ' Set this to get correct parenting of Error handler forms

Public Function PollutantConcentrationSetup(clsPollutant As clsXMLPollutantItem, strLandClass As String, strWQName As String) As Boolean
  On Error GoTo ErrorHandler

'Sub takes incoming parameters (in the form of a pollutant item) from the project file
'and then parses them out

    'RS
    Dim rsPoll As New ADODB.Recordset
    Dim rsType As New ADODB.Recordset
    Dim rsPollColor As New ADODB.Recordset
    
    'Open Strings
    Dim strPoll As String
    Dim strType As String
    Dim strField As String
    Dim strConStatement As String
    Dim strTempCoeffSet As String       'Again, because of landuse, we have to check for 'temp' coeff sets and their use
    Dim strPollColor As String
    
    
    'Get the name of the pollutant
    m_strPollName = clsPollutant.strPollName
    
    'Get the name of the Water Quality Standard
    m_strWQName = strWQName
    
    'Figure out what coeff user wants
    Select Case clsPollutant.strCoeff
        Case "Type 1"
            strField = "Coeff1"
        Case "Type 2"
            strField = "Coeff2"
        Case "Type 3"
            strField = "Coeff3"
        Case "Type 4"
            strField = "Coeff4"
        Case ""
    End Select
    
    'Find out the name of the Coefficient set, could be a temporary one due to landuses
    If Len(g_DictTempNames.Item(clsPollutant.strCoeffSet)) > 0 Then
        strTempCoeffSet = g_DictTempNames.Item(clsPollutant.strCoeffSet)
    Else
        strTempCoeffSet = clsPollutant.strCoeffSet
    End If
        
    If Len(strField) > 0 Then
        strPoll = "SELECT * FROM COEFFICIENTSET WHERE NAME LIKE '" & strTempCoeffSet & "'"
        rsPoll.Open strPoll, g_ADOConn, adOpenDynamic, adLockOptimistic
            
        strType = "SELECT LCCLASS.Value, LCCLASS.Name, COEFFICIENT." & strField & " As CoeffType, COEFFICIENT.CoeffID, COEFFICIENT.LCCLASSID " & _
                   "FROM LCCLASS LEFT OUTER JOIN COEFFICIENT " & _
                   "ON LCCLASS.LCCLASSID = COEFFICIENT.LCCLASSID " & _
                   "WHERE COEFFICIENT.COEFFSETID = " & rsPoll!CoeffSetID & " ORDER BY LCCLASS.VALUE"
        
        rsType.Open strType, g_ADOConn, adOpenKeyset, adLockOptimistic
                
        strConStatement = ConstructConStatment(rsType, g_LandCoverRaster)
        m_strPollCoeffMetadata = ConstructMetaData(rsType, clsPollutant.strCoeff, g_booLocalEffects)
        
    End If
    
    'Find out the color of the pollutant
    strPollColor = "Select Color from Pollutant where NAME LIKE '" & m_strPollName & "'"
    rsPollColor.Open strPollColor, g_ADOConn, adOpenDynamic, adLockReadOnly
    
    m_strColor = CStr(rsPollColor!Color)
    
    If CalcPollutantConcentration(strConStatement) Then
        PollutantConcentrationSetup = True
    Else
        PollutantConcentrationSetup = False
    End If
    
    'Cleanup
    rsPoll.Close
    rsType.Close
    
    Set rsPoll = Nothing
    Set rsType = Nothing

  Exit Function
ErrorHandler:
    HandleError True, "PollutantConcentrationSetup " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
    PollutantConcentrationSetup = False
End Function

Private Function ConstructMetaData(rsType As ADODB.Recordset, strCoeffSet As String, booLocal As Boolean) As String
'Takes the rs and creates a string describing the pollutants and coefficients used in this run, will
'later be added to the global dictionary

    Dim strConstructMetaData As String
    Dim strLandClassCoeff As String
    Dim strHeader As String
    Dim i As Integer
    
    If booLocal Then
    
        strHeader = vbTab & "Input Datasets:" & vbNewLine & _
                    vbTab & vbTab & "Hydrologic soils grid: " & g_clsXMLPrjFile.strSoilsHydFileName & vbNewLine & _
                    vbTab & vbTab & "Landcover grid: " & g_clsXMLPrjFile.strLCGridFileName & vbNewLine & _
                    vbTab & vbTab & "Landcover grid type: " & g_clsXMLPrjFile.strLCGridType & vbNewLine & _
                    vbTab & vbTab & "Landcover grid units: " & g_clsXMLPrjFile.strLCGridUnits & vbNewLine & _
                    vbTab & vbTab & "Precipitation grid: " & g_strPrecipFileName & vbNewLine
                    
    Else
        strHeader = vbTab & "Input Datasets:" & vbNewLine & _
                    vbTab & vbTab & "Hydrologic soils grid: " & g_clsXMLPrjFile.strSoilsHydFileName & vbNewLine & _
                    vbTab & vbTab & "Landcover grid: " & g_clsXMLPrjFile.strLCGridFileName & vbNewLine & _
                    vbTab & vbTab & "Precipitation grid: " & g_strPrecipFileName & vbNewLine & _
                    vbTab & vbTab & "Flow direction grid: " & g_strFlowDirFilename & vbNewLine
    End If
    
    strConstructMetaData = vbTab & "Pollutant Coefficients:" & vbNewLine & _
                           vbTab & vbTab & "Pollutant: " & m_strPollName & vbNewLine & _
                           vbTab & vbTab & "Coefficient Set: " & strCoeffSet & vbNewLine & _
                           vbTab & vbTab & "The following lists the landcover classes and associated coefficients used" & vbNewLine & _
                           vbTab & vbTab & "in the N-SPECT analysis run that created this dataset: " & vbNewLine
    
    rsType.MoveFirst
    
    For i = 1 To rsType.RecordCount
        If i = 1 Then
            strLandClassCoeff = vbTab & vbTab & vbTab & rsType!Name & ": " & rsType!CoeffType & vbNewLine
        Else
            strLandClassCoeff = strLandClassCoeff & vbTab & vbTab & vbTab & rsType!Name & ": " & rsType!CoeffType & vbNewLine
        End If
        rsType.MoveNext
    Next i
    
    ConstructMetaData = strHeader & g_strLandCoverParameters & strConstructMetaData & strLandClassCoeff
    
End Function

Private Function ConstructConStatment(rsType As ADODB.Recordset, pLCRaster As IRaster) As String
'Creates the initial pick statement using the name of the the LandCass [CCAP, for example]
'and the Land Class Raster.  Returns a string

On Error GoTo ErrHandler:

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
    rsType.MoveFirst
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
    
    'REMOVED 11/30/2007 in favor of code below
    'STEP 3: Table values vs. Raster Values Count - if not equal bark --------------------------
'    If pTable.RowCount(Nothing) <> rsType.RecordCount Then
'        MsgBox "Error: The number of records in your Land Class Table do not match your GRID dataset."
'        Exit Function
'    End If
    
    'STEP 4: Create the strings
    'Loop through and get all values
    'Code changed 11/30/2007 to account for mismatches in landclass table vs. a clipped grid.
    Do While Not pRow Is Nothing
    
        booValueFound = False
        rsType.MoveFirst
        
        For i = 0 To rsType.RecordCount - 1
            If rsType!Value = pRow.Value(FieldIndex) Then
                
                booValueFound = True
            
                If strCon = "" Then
                    strCon = "Con(([nu_lulc] eq " & pRow.Value(FieldIndex) & "), " & rsType!CoeffType & ", "
                Else
                    strCon = strCon & "Con(([nu_lulc] eq " & pRow.Value(FieldIndex) & "), " & rsType!CoeffType & ", "
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
            rsType.MoveNext
        Next i
    
        If booValueFound = False Then
            MsgBox "Values in table LCClass table not equal to values in landclass dataset."
            ConstructConStatment = ""
            Exit Function
        Else
            Set pRow = pCursor.NextRow
            i = 0
        End If
    Loop
        
    strCompleteCon = strCon & strParens
    
    ConstructConStatment = strCompleteCon
    
    'Cleanup:
    'Set pLCRaster = Nothing
    Set pRasterCol = Nothing
    Set pBand = Nothing
    Set pTable = Nothing
    Set pCursor = Nothing
    Set pRow = Nothing
    
Exit Function
ErrHandler:
    MsgBox "Error in Con Statement: " & Err.Number & ": " & Err.Description
End Function

Private Function CalcPollutantConcentration(strConStatement As String) As Boolean
    
    Dim pRasterProps As IRasterProps
    Dim pLandSampleRaster As IRaster
    Dim pPollMassRaster As IRaster
    Dim pMassVolumeRaster As IRaster
    Dim pPermMassVolumeRaster As IRaster
    Dim pAccumPollRaster As IRaster
    Dim pTemp1PollRaster As IRaster
    Dim pTemp2PollRaster As IRaster
    Dim pTotalPollConcRaster As IRaster
    Dim pPermAccPollRaster As IRaster
    Dim pPermTotalConcRaster As IRaster
    Dim pTotalPollConc0Raster As IRaster 'gets rid of no data...replace with 0
    Dim pPollRasterLayer As IRasterLayer
    Dim pAccPollRasterLayer As IRasterLayer
    
    Dim pEnv As IRasterAnalysisEnvironment

    'Create Map Algebra Operator
    Dim pMapAlgebraOp As IMapAlgebraOp      'Workhorse

    'String to hold calculations
    Dim strExpression As String
    Dim strTitle As String
    strTitle = "Processing " & m_strPollName & " Conc. Calculation..."
    Dim strOutConc As String
    Dim strAccPoll As String
    
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
    
                strExpression = "resample([landcover]," & CStr(g_dblCellSize) & ")"
    
                Set pLandSampleRaster = pMapAlgebraOp.Execute(strExpression)
        
                pMapAlgebraOp.UnbindRaster "landcover"
            Else
                Set pLandSampleRaster = g_LandCoverRaster
            End If
    End If

    modProgDialog.ProgDialog "Calculating EMC GRID...", strTitle, 0, 13, 1, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then

        'Step 1: CREATE PHOSPHORUS EMC GRID AT CELL LEVEL -----------------------------------------
        pMapAlgebraOp.BindRaster pLandSampleRaster, "nu_lulc"
        
        strExpression = strConStatement
        
        Set pPollMassRaster = pMapAlgebraOp.Execute(strExpression)
        
        pMapAlgebraOp.UnbindRaster "nu_lulc"
        'END STEP 1: ------------------------------------------------------------------------------
    End If
    
    'modUtil.AddRasterLayer frmPrj.m_App, pPollMassRaster, "PollMass"
    
    modProgDialog.ProgDialog "Calculating Mass Volume...", strTitle, 0, 13, 2, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
        'STEP 2: MASS OF PHOSPHORUS PRODUCED BY EACH CELL -----------------------------------------
        With pMapAlgebraOp
            .BindRaster g_pMetRunoffRaster, "met_runoff"
            .BindRaster pPollMassRaster, "pollmass"
        End With
        
        strExpression = "[met_runoff] * [pollmass]"
        
        Set pMassVolumeRaster = pMapAlgebraOp.Execute(strExpression)
        
        With pMapAlgebraOp
            .UnbindRaster "met_runoff"
            .UnbindRaster "pollmass"
        End With
        'END STEP 2: -------------------------------------------------------------------------------
    End If
    
    'LOCAL EFFECTS ONLY...
    'At this point the above grid will satisfy 'local effects only' people so...
    If g_booLocalEffects Then
        
        modProgDialog.ProgDialog "Creating data layer for local effects...", strTitle, 0, 13, 13, frmPrj.m_App.hwnd
        If modProgDialog.g_boolCancel Then
                       
            strOutConc = modUtil.GetUniqueName("locconc", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
            'Added 7/23/04 to account for clip by selected polys functionality
            If g_booSelectedPolys Then
                Dim pClipAccPollRaster As IRaster
                Set pClipAccPollRaster = modUtil.ClipBySelectedPoly(pMassVolumeRaster, g_pSelectedPolyClip, pEnv)
                Set pPermMassVolumeRaster = modUtil.ReturnPermanentRaster(pClipAccPollRaster, pEnv.OutWorkspace.PathName, strOutConc)
                Set pClipAccPollRaster = Nothing
            Else
                Set pPermMassVolumeRaster = modUtil.ReturnPermanentRaster(pMassVolumeRaster, pEnv.OutWorkspace.PathName, strOutConc)
            End If
            
            Set pPollRasterLayer = modUtil.ReturnRasterLayer(frmPrj.m_App, pPermMassVolumeRaster, m_strPollName & " Local Effects (mg)")
            Set pPollRasterLayer.Renderer = modUtil.ReturnRasterStretchColorRampRender(pPollRasterLayer, m_strColor)
            pPollRasterLayer.Visible = False
            
            g_dicMetadata.Add pPollRasterLayer.Name, m_strPollCoeffMetadata
            
            g_pGroupLayer.Add pPollRasterLayer
            
            CalcPollutantConcentration = True
           
        End If
        
        modProgDialog.KillDialog
        Exit Function
        
    End If
    
    modProgDialog.ProgDialog "Deriving accumulated pollutant...", strTitle, 0, 13, 3, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
        'STEP 3: DERIVE ACCUMULATED POLLUTANT ------------------------------------------------------
        With pMapAlgebraOp
            .BindRaster g_pFlowDirRaster, "flowdir"
            .BindRaster pMassVolumeRaster, "massvolume"
        End With
        
        strExpression = "(FlowAccumulation([flowdir], [massvolume], FLOAT)) * 1.0e-6"
        
        Set pAccumPollRaster = pMapAlgebraOp.Execute(strExpression)
        
        With pMapAlgebraOp
            .UnbindRaster "flowdir"
            .UnbindRaster "massvolume"
        End With
        'END STEP 3: ------------------------------------------------------------------------------
    End If
    
    'STEP 3a: Added 7/26: ADD ACCUMULATED POLLUTANT TO GROUP LAYER-----------------------------------
    modProgDialog.ProgDialog "Creating accumlated pollutant layer...", strTitle, 0, 13, 4, frmPrj.hwnd
    If modProgDialog.g_boolCancel Then
        
        strAccPoll = modUtil.GetUniqueName("accpoll", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        'Added 7/23/04 to account for clip by selected polys functionality
        If g_booSelectedPolys Then
            Dim pClipAccPoll2Raster As IRaster
            Set pClipAccPoll2Raster = modUtil.ClipBySelectedPoly(pAccumPollRaster, g_pSelectedPolyClip, pEnv)
            Set pPermAccPollRaster = modUtil.ReturnPermanentRaster(pClipAccPoll2Raster, pEnv.OutWorkspace.PathName, strAccPoll)
            Set pClipAccPoll2Raster = Nothing
        Else
            Set pPermAccPollRaster = modUtil.ReturnPermanentRaster(pAccumPollRaster, pEnv.OutWorkspace.PathName, strAccPoll)
        End If
        
        Set pAccPollRasterLayer = modUtil.ReturnRasterLayer(frmPrj.m_App, pPermAccPollRaster, "Accumulated " & m_strPollName & " (kg)")
        Set pAccPollRasterLayer.Renderer = modUtil.ReturnRasterStretchColorRampRender(pAccPollRasterLayer, m_strColor)
        
        pAccPollRasterLayer.Visible = False
        
        g_dicMetadata.Add pAccPollRasterLayer.Name, m_strPollCoeffMetadata
            
        g_pGroupLayer.Add pAccPollRasterLayer
    
    End If
    'END STEP 3a: ---------------------------------------------------------------------------------
        
    
    modProgDialog.ProgDialog "Deriving total concentration at each cell...", strTitle, 0, 13, 6, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
        'STEP 5: DERIVE TOTAL CONCENTRATION (AT EACH CELL) ----------------------------------------
        With pMapAlgebraOp
            .BindRaster pMassVolumeRaster, "massvolume"
            .BindRaster pAccumPollRaster, "accpoll"
        End With
        
        strExpression = "[massvolume] + ([accpoll] / 1.0e-6)"
        
        Set pTemp1PollRaster = pMapAlgebraOp.Execute(strExpression)
        
        With pMapAlgebraOp
            .UnbindRaster "massvolume"
            .UnbindRaster "accpoll"
        End With
        'END STEP 5: -------------------------------------------------------------------------------
    End If
    
    modProgDialog.ProgDialog "Adding metric runoff...", strTitle, 0, 13, 7, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
        'STEP 6: -----------------------------------------------------------------------------------
        With pMapAlgebraOp
            .BindRaster g_pMetRunoffRaster, "met_run"
            .BindRaster g_pRunoffRaster, "accrun"
        End With
        
        If g_pMetRunoffRaster Is Nothing Then
            MsgBox "g_pMetRunoff is crap"
        End If
        
        If g_pRunoffRaster Is Nothing Then
            MsgBox "g_pRunoff is crap"
        End If
        
        strExpression = "[met_run] + [accrun]"
        
        Set pTemp2PollRaster = pMapAlgebraOp.Execute(strExpression)
        
        With pMapAlgebraOp
            .UnbindRaster "met_run"
            .UnbindRaster "accrun"
        End With
        'END STEP 6: ------------------------------------------------------------------------------
    End If
        
    modProgDialog.ProgDialog "Calculating final concentration...", strTitle, 0, 13, 8, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
    
        'STEP 7: FINAL CONCENTRATION ---------------------------------------------------------------
        With pMapAlgebraOp
            .BindRaster pTemp1PollRaster, "temp1"
            .BindRaster pTemp2PollRaster, "temp2"
        End With
        
        strExpression = "[temp1] / [temp2]"
        
        Set pTotalPollConcRaster = pMapAlgebraOp.Execute(strExpression)
        
        With pMapAlgebraOp
            .UnbindRaster "temp1"
            .UnbindRaster "temp2"
        End With
        'END STEP 7: --------------------------------------------------------------------------------
    End If
         
    modProgDialog.ProgDialog "Calculating final concentration...", strTitle, 0, 13, 9, frmPrj.m_App.hwnd
    If modProgDialog.g_boolCancel Then
    
        'STEP 8: FINAL CONCENTRATION -remove all noData values ---------------------------------------
        With pMapAlgebraOp
            .BindRaster g_pDEMRaster, "dem"
            .BindRaster pTotalPollConcRaster, "totalConc"
        End With
        
        strExpression = "Merge([totalConc], Con([dem] >= 0, 0))"
        
        Set pTotalPollConc0Raster = pMapAlgebraOp.Execute(strExpression)
        
        With pMapAlgebraOp
            .UnbindRaster "dem"
            .UnbindRaster "totalConc"
        End With
        'END STEP 7: --------------------------------------------------------------------------------
    End If
     
    modProgDialog.ProgDialog "Converting to correct units...", strTitle, 0, 13, 10, frmPrj.m_App.hwnd
    
    If modProgDialog.g_boolCancel Then
        modProgDialog.ProgDialog "Creating data layer...", strTitle, 0, 13, 11, frmPrj.m_App.hwnd
        
        strOutConc = modUtil.GetUniqueName("conc", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        
        If g_booSelectedPolys Then
            Dim pClipTotalConcRaster As IRaster
            Set pClipTotalConcRaster = modUtil.ClipBySelectedPoly(pTotalPollConc0Raster, g_pSelectedPolyClip, pEnv)
            Set pPermTotalConcRaster = modUtil.ReturnPermanentRaster(pClipTotalConcRaster, pEnv.OutWorkspace.PathName, strOutConc)
            Set pClipTotalConcRaster = Nothing
        Else
            Set pPermTotalConcRaster = modUtil.ReturnPermanentRaster(pTotalPollConc0Raster, pEnv.OutWorkspace.PathName, strOutConc)
        End If
        
        Set pPollRasterLayer = modUtil.ReturnRasterLayer(frmPrj.m_App, pPermTotalConcRaster, m_strPollName & " Conc. (mg/L)")
        Set pPollRasterLayer.Renderer = modUtil.ReturnRasterStretchColorRampRender(pPollRasterLayer, m_strColor)

        pPollRasterLayer.Visible = False
        
        g_dicMetadata.Add pPollRasterLayer.Name, m_strPollCoeffMetadata
            
        g_pGroupLayer.Add pPollRasterLayer
        
    End If
    
    'Cleanup
    Set pLandSampleRaster = Nothing
    Set pMassVolumeRaster = Nothing
    Set pPollMassRaster = Nothing
    Set pAccumPollRaster = Nothing
    Set pTemp1PollRaster = Nothing
    Set pTemp2PollRaster = Nothing
    Set pTotalPollConcRaster = Nothing
    Set pMapAlgebraOp = Nothing
    Set pEnv = Nothing
    Set pTotalPollConc0Raster = Nothing

    modProgDialog.ProgDialog "Comparing to water quality standard...", strTitle, 0, 13, 13, frmPrj.m_App.hwnd
    
    If modProgDialog.g_boolCancel Then
        If Not CompareWaterQuality(g_pWaterShedFeatClass, pPermTotalConcRaster) Then
            CalcPollutantConcentration = False
            Exit Function
        End If
    End If
    
    'if we get to the end
    CalcPollutantConcentration = True
    
    Set pPermTotalConcRaster = Nothing
    
    modProgDialog.KillDialog
    
Exit Function


ErrHandler:
    If Err.Number = -2147217297 Then 'User cancelled operation
        modProgDialog.g_boolCancel = False
        CalcPollutantConcentration = False
        Exit Function
    ElseIf Err.Number = -2147467259 Then
        MsgBox "ArcMap has reached the maximum number of GRIDs allowed in memory.  " & _
            "Please exit N-SPECT and restart ArcMap.", vbInformation, "Maximum GRID Number Encountered"
        modProgDialog.g_boolCancel = False
        modProgDialog.KillDialog
        CalcPollutantConcentration = False
    Else
        HandleError False, "CalcPollutantConcentration " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
        modProgDialog.g_boolCancel = False
        modProgDialog.KillDialog
        CalcPollutantConcentration = False
    End If

End Function

Private Function CompareWaterQuality(pWSFeatureClass As IFeatureClass, pPollutantRaster As IRaster) As Boolean
  
    'Get the zone dataset from the first layer in ArcMap
    Dim pGeoPollDS As IGeoDataset
    Dim pLocalOp As ILocalOp
    Dim pEnv As IRasterAnalysisEnvironment
    Dim pWS As IWorkspace
    Dim pMaxRaster As IRaster
    Dim pConRaster As IRaster
    Dim pPermWQRaster As IRaster
    Dim pWQRasterLayer As IRasterLayer
    Dim pMapAlgebraOp As IMapAlgebraOp
    Dim dblConvertValue As Double
    Dim strOutWQ As String
    Dim strExpression As String
    Dim strMetadata As String
   
On Error GoTo ErrorHandler:
   
   'Create the value raster from the pollutant Raster
    Set pGeoPollDS = pPollutantRaster

   'Create a Spatial operator
    Set pLocalOp = New RasterLocalOp
    
    'Set output workspace and environ
    Set pEnv = g_pSpatEnv
    Set pEnv = pLocalOp
    Set pWS = pEnv.OutWorkspace
    
    ' Perform Spatial operation
    Set pMaxRaster = pLocalOp.LocalStatistics(pGeoPollDS, esriGeoAnalysisStatsMaximum)
    
    strWQVAlue = ReturnWQValue(m_strPollName, m_strWQName)
    
    dblConvertValue = (CDbl(strWQVAlue)) / 1000
    
    'Now run water quality
    Set pEnv = g_pSpatEnv
    Set pMapAlgebraOp = New RasterMapAlgebraOp
    Set pEnv = pMapAlgebraOp
    
    With pMapAlgebraOp
        .BindRaster pMaxRaster, "Max"
        .BindRaster g_pFlowAccRaster, "flowacc"
    End With
    
    'This rather ugly expression was set up to check for meets/exceed water quality standards for
    'only the streams.  It takes the values of flowaccumulation from watershed delineation fame that
    'exceed values of greater than 1%.  Then multiplies the result (all cells representing streams) times
    'the water quality grid.
    strExpression = "(Con([Max] gt " & CStr(dblConvertValue) & ", 1, 2)) * (con([flowacc] > (" & _
                    CStr(modUtil.ReturnRasterMax(g_pFlowAccRaster)) & " * 0.01), 1))"
    Set pConRaster = pMapAlgebraOp.Execute(strExpression)
    
    pMapAlgebraOp.UnbindRaster "Max"
    pMapAlgebraOp.UnbindRaster "flowacc"
    
    strOutWQ = modUtil.GetUniqueName("wq", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
    
    'Clip if selectedpolys
    If g_booSelectedPolys Then
        Dim pClipWQRaster As IRaster
        Set pClipWQRaster = modUtil.ClipBySelectedPoly(pConRaster, g_pSelectedPolyClip, pEnv)
        Set pPermWQRaster = modUtil.ReturnPermanentRaster(pClipWQRaster, pEnv.OutWorkspace.PathName, strOutWQ)
        Set pClipWQRaster = Nothing
    Else
        Set pPermWQRaster = modUtil.ReturnPermanentRaster(pConRaster, pEnv.OutWorkspace.PathName, strOutWQ)
    End If
    
    Set pWQRasterLayer = modUtil.ReturnRasterLayer(frmPrj.m_App, pPermWQRaster, m_strPollName & " Standard: " & CStr(dblConvertValue) & " mg/L")
    Set pWQRasterLayer.Renderer = modUtil.ReturnUniqueRasterRenderer(pWQRasterLayer, m_strWQName)
    pWQRasterLayer.Visible = False

    
    strMetadata = vbTab & "Water Quality Standard:" & vbNewLine & _
                  vbTab & vbTab & "Criteria Name: " & m_strWQName & vbNewLine & _
                  vbTab & vbTab & "Standard: " & dblConvertValue & " mg/L"
                  

    g_dicMetadata.Add pWQRasterLayer.Name, m_strPollCoeffMetadata & strMetadata
            
    g_pGroupLayer.Add pWQRasterLayer
    
    CompareWaterQuality = True
    
    'Cleanup
    Set pGeoPollDS = Nothing
    Set pLocalOp = Nothing
    Set pEnv = Nothing
    Set pWS = Nothing
    Set pMaxRaster = Nothing
    Set pConRaster = Nothing
    Set pPermWQRaster = Nothing
    Set pMapAlgebraOp = Nothing

  Exit Function

ErrorHandler:
    If Err.Number = -2147467259 Then
        MsgBox "ArcMap has reached the maximum number of GRIDs allowed in memory.  " & _
            "Please exit N-SPECT and restart ArcMap.", vbInformation, "Maximum GRID Number Encountered"
        CompareWaterQuality = False
        modProgDialog.g_boolCancel = False
        modProgDialog.KillDialog
    Else
        HandleError False, "CompareWaterQuality " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
        CompareWaterQuality = False
        modProgDialog.g_boolCancel = False
             modProgDialog.KillDialog
    End If
End Function


Private Function ReturnWQValue(strPollName As String, strWQstdName As String) As String

On Error GoTo ErrHandler:
    Dim rsPoll As New ADODB.Recordset
    Dim rsWQstd As New ADODB.Recordset
    
    Dim strPoll As String
    Dim strWQStd As String
    
    strPoll = "Select * from Pollutant where name like '" & strPollName & "'"
    
    rsPoll.Open strPoll, g_ADOConn, adOpenStatic, adLockReadOnly
    
    strWQStd = "SELECT * FROM WQCRITERIA INNER JOIN POLL_WQCRITERIA ON WQCRITERIA.WQCRITID = POLL_WQCRITERIA.WQCRITID " & _
    "WHERE WQCRITERIA.NAME Like '" & strWQstdName & "' AND POLL_WQCRITERIA.POLLID = " & rsPoll!POLLID
    
    rsWQstd.Open strWQStd, g_ADOConn, adOpenStatic, adLockReadOnly
    
    ReturnWQValue = CStr(rsWQstd!Threshold)
    
    rsPoll.Close
    rsWQstd.Close
    
    Set rsPoll = Nothing
    Set rsWQstd = Nothing

Exit Function
     
ErrHandler:
    MsgBox "Error in ADO pollutant part: " & Err.Number & vbNewLine & Err.Description & vbNewLine & strWQStd
End Function
