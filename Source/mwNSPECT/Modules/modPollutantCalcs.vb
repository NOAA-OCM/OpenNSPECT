Imports System.Data.OleDb
Module modPollutantCalcs
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

    Private m_strPollName As String 'Mod level variable for name of pollutant being used
    Private m_strWQName As String 'Mod level variable for name of water quality standard
    Private m_strColor As String 'Mod level variable holding the string of the pollutant color
    Private m_strPollCoeffMetadata As String 'Variable to hold coeffs for use in metadata

    ' Constant used by the Error handler function - DO NOT REMOVE
    Const c_sModuleFileName As String = "modPollutantCalcs.bas"
    Private m_ParentHWND As Integer ' Set this to get correct parenting of Error handler forms

    Public Function PollutantConcentrationSetup(ByRef clsPollutant As clsXMLPollutantItem, ByRef strLandClass As String, ByRef strWQName As String) As Boolean
        '        On Error GoTo ErrorHandler

        '        'Sub takes incoming parameters (in the form of a pollutant item) from the project file
        '        'and then parses them out

        '        'RS
        '        Dim rsPoll As New ADODB.Recordset
        '        Dim rsType As New ADODB.Recordset
        '        Dim rsPollColor As New ADODB.Recordset

        '        'Open Strings
        '        Dim strPoll As String
        '        Dim strType As String
        '        Dim strField As String
        '        Dim strConStatement As String
        '        Dim strTempCoeffSet As String 'Again, because of landuse, we have to check for 'temp' coeff sets and their use
        '        Dim strPollColor As String


        '        'Get the name of the pollutant
        '        m_strPollName = clsPollutant.strPollName

        '        'Get the name of the Water Quality Standard
        '        m_strWQName = strWQName

        '        'Figure out what coeff user wants
        '        Select Case clsPollutant.strCoeff
        '            Case "Type 1"
        '                strField = "Coeff1"
        '            Case "Type 2"
        '                strField = "Coeff2"
        '            Case "Type 3"
        '                strField = "Coeff3"
        '            Case "Type 4"
        '                strField = "Coeff4"
        '            Case ""
        '        End Select

        '        'Find out the name of the Coefficient set, could be a temporary one due to landuses
        '        If Len(g_DictTempNames.Item(clsPollutant.strCoeffSet)) > 0 Then
        '            'UPGRADE_WARNING: Couldn't resolve default property of object g_DictTempNames.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            strTempCoeffSet = g_DictTempNames.Item(clsPollutant.strCoeffSet)
        '        Else
        '            strTempCoeffSet = clsPollutant.strCoeffSet
        '        End If

        '        If Len(strField) > 0 Then
        '            strPoll = "SELECT * FROM COEFFICIENTSET WHERE NAME LIKE '" & strTempCoeffSet & "'"
        '            rsPoll.Open(strPoll, g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        '            strType = "SELECT LCCLASS.Value, LCCLASS.Name, COEFFICIENT." & strField & " As CoeffType, COEFFICIENT.CoeffID, COEFFICIENT.LCCLASSID " & "FROM LCCLASS LEFT OUTER JOIN COEFFICIENT " & "ON LCCLASS.LCCLASSID = COEFFICIENT.LCCLASSID " & "WHERE COEFFICIENT.COEFFSETID = " & rsPoll.Fields("CoeffSetID").Value & " ORDER BY LCCLASS.VALUE"

        '            rsType.Open(strType, g_ADOConn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)

        '            strConStatement = ConstructConStatment(rsType, g_LandCoverRaster)
        '            m_strPollCoeffMetadata = ConstructMetaData(rsType, (clsPollutant.strCoeff), g_booLocalEffects)

        '        End If

        '        'Find out the color of the pollutant
        '        strPollColor = "Select Color from Pollutant where NAME LIKE '" & m_strPollName & "'"
        '        rsPollColor.Open(strPollColor, g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockReadOnly)

        '        m_strColor = CStr(rsPollColor.Fields("Color").Value)

        '        If CalcPollutantConcentration(strConStatement) Then
        '            PollutantConcentrationSetup = True
        '        Else
        '            PollutantConcentrationSetup = False
        '        End If

        '        'Cleanup
        '        rsPoll.Close()
        '        rsType.Close()

        '        'UPGRADE_NOTE: Object rsPoll may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        rsPoll = Nothing
        '        'UPGRADE_NOTE: Object rsType may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        rsType = Nothing

        '        Exit Function
        'ErrorHandler:
        '        HandleError(True, "PollutantConcentrationSetup " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        '        PollutantConcentrationSetup = False
    End Function

    Private Function ConstructMetaData(ByRef cmdType As oledbcommand, ByRef strCoeffSet As String, ByRef booLocal As Boolean) As String
        ''Takes the rs and creates a string describing the pollutants and coefficients used in this run, will
        ''later be added to the global dictionary

        'Dim strConstructMetaData As String
        'Dim strLandClassCoeff As String
        'Dim strHeader As String
        'Dim i As Short

        'If booLocal Then

        '    strHeader = vbTab & "Input Datasets:" & vbNewLine & vbTab & vbTab & "Hydrologic soils grid: " & g_clsXMLPrjFile.strSoilsHydFileName & vbNewLine & vbTab & vbTab & "Landcover grid: " & g_clsXMLPrjFile.strLCGridFileName & vbNewLine & vbTab & vbTab & "Landcover grid type: " & g_clsXMLPrjFile.strLCGridType & vbNewLine & vbTab & vbTab & "Landcover grid units: " & g_clsXMLPrjFile.strLCGridUnits & vbNewLine & vbTab & vbTab & "Precipitation grid: " & g_strPrecipFileName & vbNewLine

        'Else
        '    strHeader = vbTab & "Input Datasets:" & vbNewLine & vbTab & vbTab & "Hydrologic soils grid: " & g_clsXMLPrjFile.strSoilsHydFileName & vbNewLine & vbTab & vbTab & "Landcover grid: " & g_clsXMLPrjFile.strLCGridFileName & vbNewLine & vbTab & vbTab & "Precipitation grid: " & g_strPrecipFileName & vbNewLine & vbTab & vbTab & "Flow direction grid: " & g_strFlowDirFilename & vbNewLine
        'End If

        'strConstructMetaData = vbTab & "Pollutant Coefficients:" & vbNewLine & vbTab & vbTab & "Pollutant: " & m_strPollName & vbNewLine & vbTab & vbTab & "Coefficient Set: " & strCoeffSet & vbNewLine & vbTab & vbTab & "The following lists the landcover classes and associated coefficients used" & vbNewLine & vbTab & vbTab & "in the N-SPECT analysis run that created this dataset: " & vbNewLine

        'rsType.MoveFirst()

        'For i = 1 To rsType.RecordCount
        '    If i = 1 Then
        '        strLandClassCoeff = vbTab & vbTab & vbTab & rsType.Fields("Name").Value & ": " & rsType.Fields("CoeffType").Value & vbNewLine
        '    Else
        '        strLandClassCoeff = strLandClassCoeff & vbTab & vbTab & vbTab & rsType.Fields("Name").Value & ": " & rsType.Fields("CoeffType").Value & vbNewLine
        '    End If
        '    rsType.MoveNext()
        'Next i

        'ConstructMetaData = strHeader & g_strLandCoverParameters & strConstructMetaData & strLandClassCoeff

    End Function

    Private Function ConstructConStatment(ByRef cmdType As OleDbCommand, ByRef pLCRaster As MapWinGIS.Grid) As String
        '        'Creates the initial pick statement using the name of the the LandCass [CCAP, for example]
        '        'and the Land Class Raster.  Returns a string

        '        On Error GoTo ErrHandler

        '        Dim strCon As String 'Con statement base
        '        Dim strParens As String 'String of trailing parens
        '        Dim strCompleteCon As String 'Concatenate of strCon & strParens

        '        Dim pRasterCol As ESRI.ArcGIS.DataSourcesRaster.IRasterBandCollection
        '        Dim pBand As ESRI.ArcGIS.DataSourcesRaster.IRasterBand
        '        Dim pTable As ESRI.ArcGIS.Geodatabase.ITable
        '        Dim TableExist As Boolean
        '        Dim pCursor As ESRI.ArcGIS.Geodatabase.ICursor
        '        Dim pRow As ESRI.ArcGIS.Geodatabase.IRow
        '        Dim FieldIndex As Short
        '        Dim booValueFound As Boolean
        '        Dim i As Short

        '        'STEP 1:  get the records from the database -----------------------------------------------
        '        rsType.MoveFirst()
        '        'End Database stuff

        '        'STEP 2: Raster Values ---------------------------------------------------------------------
        '        'Now Get the RASTER values
        '        ' Get Rasterband from the incoming raster
        '        pRasterCol = pLCRaster
        '        pBand = pRasterCol.Item(0)

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

        '        'REMOVED 11/30/2007 in favor of code below
        '        'STEP 3: Table values vs. Raster Values Count - if not equal bark --------------------------
        '        '    If pTable.RowCount(Nothing) <> rsType.RecordCount Then
        '        '        MsgBox "Error: The number of records in your Land Class Table do not match your GRID dataset."
        '        '        Exit Function
        '        '    End If

        '        'STEP 4: Create the strings
        '        'Loop through and get all values
        '        'Code changed 11/30/2007 to account for mismatches in landclass table vs. a clipped grid.
        '        Do While Not pRow Is Nothing

        '            booValueFound = False
        '            rsType.MoveFirst()

        '            For i = 0 To rsType.RecordCount - 1
        '                'UPGRADE_WARNING: Couldn't resolve default property of object pRow.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                If rsType.Fields("Value").Value = pRow.Value(FieldIndex) Then

        '                    booValueFound = True

        '                    If strCon = "" Then
        '                        'UPGRADE_WARNING: Couldn't resolve default property of object pRow.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                        strCon = "Con(([nu_lulc] eq " & pRow.Value(FieldIndex) & "), " & rsType.Fields("CoeffType").Value & ", "
        '                    Else
        '                        'UPGRADE_WARNING: Couldn't resolve default property of object pRow.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                        strCon = strCon & "Con(([nu_lulc] eq " & pRow.Value(FieldIndex) & "), " & rsType.Fields("CoeffType").Value & ", "
        '                    End If

        '                    If strParens = "" Then
        '                        strParens = "-9999)"
        '                    Else
        '                        strParens = strParens & ")"
        '                    End If

        '                    Exit For
        '                Else
        '                    booValueFound = False
        '                End If
        '                rsType.MoveNext()
        '            Next i

        '            If booValueFound = False Then
        '                MsgBox("Values in table LCClass table not equal to values in landclass dataset.")
        '                ConstructConStatment = ""
        '                Exit Function
        '            Else
        '                pRow = pCursor.NextRow
        '                i = 0
        '            End If
        '        Loop

        '        strCompleteCon = strCon & strParens

        '        ConstructConStatment = strCompleteCon

        '        'Cleanup:
        '        'Set pLCRaster = Nothing
        '        'UPGRADE_NOTE: Object pRasterCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pRasterCol = Nothing
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
        '        MsgBox("Error in Con Statement: " & Err.Number & ": " & Err.Description)
    End Function

    Private Function CalcPollutantConcentration(ByRef strConStatement As String) As Boolean

        '        Dim pRasterProps As ESRI.ArcGIS.DataSourcesRaster.IRasterProps
        '        Dim pLandSampleRaster As ESRI.ArcGIS.Geodatabase.IRaster
        '        Dim pPollMassRaster As ESRI.ArcGIS.Geodatabase.IRaster
        '        Dim pMassVolumeRaster As ESRI.ArcGIS.Geodatabase.IRaster
        '        Dim pPermMassVolumeRaster As ESRI.ArcGIS.Geodatabase.IRaster
        '        Dim pAccumPollRaster As ESRI.ArcGIS.Geodatabase.IRaster
        '        Dim pTemp1PollRaster As ESRI.ArcGIS.Geodatabase.IRaster
        '        Dim pTemp2PollRaster As ESRI.ArcGIS.Geodatabase.IRaster
        '        Dim pTotalPollConcRaster As ESRI.ArcGIS.Geodatabase.IRaster
        '        Dim pPermAccPollRaster As ESRI.ArcGIS.Geodatabase.IRaster
        '        Dim pPermTotalConcRaster As ESRI.ArcGIS.Geodatabase.IRaster
        '        Dim pTotalPollConc0Raster As ESRI.ArcGIS.Geodatabase.IRaster 'gets rid of no data...replace with 0
        '        Dim pPollRasterLayer As ESRI.ArcGIS.Carto.IRasterLayer
        '        Dim pAccPollRasterLayer As ESRI.ArcGIS.Carto.IRasterLayer

        '        Dim pEnv As ESRI.ArcGIS.GeoAnalyst.IRasterAnalysisEnvironment

        '        'Create Map Algebra Operator
        '        Dim pMapAlgebraOp As ESRI.ArcGIS.SpatialAnalyst.IMapAlgebraOp 'Workhorse

        '        'String to hold calculations
        '        Dim strExpression As String
        '        Dim strTitle As String
        '        strTitle = "Processing " & m_strPollName & " Conc. Calculation..."
        '        Dim strOutConc As String
        '        Dim strAccPoll As String

        '        'Set the enviornment stuff
        '        pEnv = g_pSpatEnv
        '        pMapAlgebraOp = New ESRI.ArcGIS.SpatialAnalyst.RasterMapAlgebraOp
        '        pEnv = pMapAlgebraOp

        '        On Error GoTo ErrHandler

        '        modProgDialog.ProgDialog("Checking landcover cell size...", strTitle, 0, 13, 1, (frmPrj.m_App.hwnd))
        '        If modProgDialog.g_boolCancel Then
        '            'Step 1a: ----------------------------------------------------------------------------
        '            'Make sure LandCover is in the same cellsize as the global environment
        '            pRasterProps = g_LandCoverRaster

        '            If pRasterProps.MeanCellSize.X <> g_dblCellSize Then

        '                'UPGRADE_WARNING: Couldn't resolve default property of object g_LandCoverRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                pMapAlgebraOp.BindRaster(g_LandCoverRaster, "landcover")

        '                strExpression = "resample([landcover]," & CStr(g_dblCellSize) & ")"

        '                pLandSampleRaster = pMapAlgebraOp.Execute(strExpression)

        '                pMapAlgebraOp.UnbindRaster("landcover")
        '            Else
        '                pLandSampleRaster = g_LandCoverRaster
        '            End If
        '        End If

        '        modProgDialog.ProgDialog("Calculating EMC GRID...", strTitle, 0, 13, 1, (frmPrj.m_App.hwnd))
        '        If modProgDialog.g_boolCancel Then

        '            'Step 1: CREATE PHOSPHORUS EMC GRID AT CELL LEVEL -----------------------------------------
        '            'UPGRADE_WARNING: Couldn't resolve default property of object pLandSampleRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            pMapAlgebraOp.BindRaster(pLandSampleRaster, "nu_lulc")

        '            strExpression = strConStatement

        '            pPollMassRaster = pMapAlgebraOp.Execute(strExpression)

        '            pMapAlgebraOp.UnbindRaster("nu_lulc")
        '            'END STEP 1: ------------------------------------------------------------------------------
        '        End If

        '        'modUtil.AddRasterLayer frmPrj.m_App, pPollMassRaster, "PollMass"

        '        modProgDialog.ProgDialog("Calculating Mass Volume...", strTitle, 0, 13, 2, (frmPrj.m_App.hwnd))
        '        If modProgDialog.g_boolCancel Then
        '            'STEP 2: MASS OF PHOSPHORUS PRODUCED BY EACH CELL -----------------------------------------
        '            With pMapAlgebraOp
        '                'UPGRADE_WARNING: Couldn't resolve default property of object g_pMetRunoffRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                .BindRaster(g_pMetRunoffRaster, "met_runoff")
        '                'UPGRADE_WARNING: Couldn't resolve default property of object pPollMassRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                .BindRaster(pPollMassRaster, "pollmass")
        '            End With

        '            strExpression = "[met_runoff] * [pollmass]"

        '            pMassVolumeRaster = pMapAlgebraOp.Execute(strExpression)

        '            With pMapAlgebraOp
        '                .UnbindRaster("met_runoff")
        '                .UnbindRaster("pollmass")
        '            End With
        '            'END STEP 2: -------------------------------------------------------------------------------
        '        End If

        '        'LOCAL EFFECTS ONLY...
        '        'At this point the above grid will satisfy 'local effects only' people so...
        '        Dim pClipAccPollRaster As ESRI.ArcGIS.Geodatabase.IRaster
        '        If g_booLocalEffects Then

        '            modProgDialog.ProgDialog("Creating data layer for local effects...", strTitle, 0, 13, 13, (frmPrj.m_App.hwnd))
        '            If modProgDialog.g_boolCancel Then

        '                'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                strOutConc = modUtil.GetUniqueName("locconc", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        '                'Added 7/23/04 to account for clip by selected polys functionality
        '                If g_booSelectedPolys Then
        '                    pClipAccPollRaster = modUtil.ClipBySelectedPoly(pMassVolumeRaster, g_pSelectedPolyClip, pEnv)
        '                    'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                    pPermMassVolumeRaster = modUtil.ReturnPermanentRaster(pClipAccPollRaster, pEnv.OutWorkspace.PathName, strOutConc)
        '                    'UPGRADE_NOTE: Object pClipAccPollRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '                    pClipAccPollRaster = Nothing
        '                Else
        '                    'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                    pPermMassVolumeRaster = modUtil.ReturnPermanentRaster(pMassVolumeRaster, pEnv.OutWorkspace.PathName, strOutConc)
        '                End If

        '                pPollRasterLayer = modUtil.ReturnRasterLayer((frmPrj.m_App), pPermMassVolumeRaster, m_strPollName & " Local Effects (mg)")
        '                pPollRasterLayer.Renderer = modUtil.ReturnRasterStretchColorRampRender(pPollRasterLayer, m_strColor)
        '                'UPGRADE_WARNING: Couldn't resolve default property of object pPollRasterLayer.Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                pPollRasterLayer.Visible = False

        '                'UPGRADE_WARNING: Couldn't resolve default property of object pPollRasterLayer.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                g_dicMetadata.Add(pPollRasterLayer.Name, m_strPollCoeffMetadata)

        '                'UPGRADE_WARNING: Couldn't resolve default property of object pPollRasterLayer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                g_pGroupLayer.Add(pPollRasterLayer)

        '                CalcPollutantConcentration = True

        '            End If

        '            modProgDialog.KillDialog()
        '            Exit Function

        '        End If

        '        modProgDialog.ProgDialog("Deriving accumulated pollutant...", strTitle, 0, 13, 3, (frmPrj.m_App.hwnd))
        '        If modProgDialog.g_boolCancel Then
        '            'STEP 3: DERIVE ACCUMULATED POLLUTANT ------------------------------------------------------
        '            With pMapAlgebraOp
        '                'UPGRADE_WARNING: Couldn't resolve default property of object g_pFlowDirRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                .BindRaster(g_pFlowDirRaster, "flowdir")
        '                'UPGRADE_WARNING: Couldn't resolve default property of object pMassVolumeRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                .BindRaster(pMassVolumeRaster, "massvolume")
        '            End With

        '            strExpression = "(FlowAccumulation([flowdir], [massvolume], FLOAT)) * 1.0e-6"

        '            pAccumPollRaster = pMapAlgebraOp.Execute(strExpression)

        '            With pMapAlgebraOp
        '                .UnbindRaster("flowdir")
        '                .UnbindRaster("massvolume")
        '            End With
        '            'END STEP 3: ------------------------------------------------------------------------------
        '        End If

        '        'STEP 3a: Added 7/26: ADD ACCUMULATED POLLUTANT TO GROUP LAYER-----------------------------------
        '        modProgDialog.ProgDialog("Creating accumlated pollutant layer...", strTitle, 0, 13, 4, frmPrj.Handle.ToInt32)
        '        Dim pClipAccPoll2Raster As ESRI.ArcGIS.Geodatabase.IRaster
        '        If modProgDialog.g_boolCancel Then

        '            'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            strAccPoll = modUtil.GetUniqueName("accpoll", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
        '            'Added 7/23/04 to account for clip by selected polys functionality
        '            If g_booSelectedPolys Then
        '                pClipAccPoll2Raster = modUtil.ClipBySelectedPoly(pAccumPollRaster, g_pSelectedPolyClip, pEnv)
        '                'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                pPermAccPollRaster = modUtil.ReturnPermanentRaster(pClipAccPoll2Raster, pEnv.OutWorkspace.PathName, strAccPoll)
        '                'UPGRADE_NOTE: Object pClipAccPoll2Raster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '                pClipAccPoll2Raster = Nothing
        '            Else
        '                'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                pPermAccPollRaster = modUtil.ReturnPermanentRaster(pAccumPollRaster, pEnv.OutWorkspace.PathName, strAccPoll)
        '            End If

        '            pAccPollRasterLayer = modUtil.ReturnRasterLayer((frmPrj.m_App), pPermAccPollRaster, "Accumulated " & m_strPollName & " (kg)")
        '            pAccPollRasterLayer.Renderer = modUtil.ReturnRasterStretchColorRampRender(pAccPollRasterLayer, m_strColor)

        '            'UPGRADE_WARNING: Couldn't resolve default property of object pAccPollRasterLayer.Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            pAccPollRasterLayer.Visible = False

        '            'UPGRADE_WARNING: Couldn't resolve default property of object pAccPollRasterLayer.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            g_dicMetadata.Add(pAccPollRasterLayer.Name, m_strPollCoeffMetadata)

        '            'UPGRADE_WARNING: Couldn't resolve default property of object pAccPollRasterLayer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            g_pGroupLayer.Add(pAccPollRasterLayer)

        '        End If
        '        'END STEP 3a: ---------------------------------------------------------------------------------


        '        modProgDialog.ProgDialog("Deriving total concentration at each cell...", strTitle, 0, 13, 6, (frmPrj.m_App.hwnd))
        '        If modProgDialog.g_boolCancel Then
        '            'STEP 5: DERIVE TOTAL CONCENTRATION (AT EACH CELL) ----------------------------------------
        '            With pMapAlgebraOp
        '                'UPGRADE_WARNING: Couldn't resolve default property of object pMassVolumeRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                .BindRaster(pMassVolumeRaster, "massvolume")
        '                'UPGRADE_WARNING: Couldn't resolve default property of object pAccumPollRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                .BindRaster(pAccumPollRaster, "accpoll")
        '            End With

        '            strExpression = "[massvolume] + ([accpoll] / 1.0e-6)"

        '            pTemp1PollRaster = pMapAlgebraOp.Execute(strExpression)

        '            With pMapAlgebraOp
        '                .UnbindRaster("massvolume")
        '                .UnbindRaster("accpoll")
        '            End With
        '            'END STEP 5: -------------------------------------------------------------------------------
        '        End If

        '        modProgDialog.ProgDialog("Adding metric runoff...", strTitle, 0, 13, 7, (frmPrj.m_App.hwnd))
        '        If modProgDialog.g_boolCancel Then
        '            'STEP 6: -----------------------------------------------------------------------------------
        '            With pMapAlgebraOp
        '                'UPGRADE_WARNING: Couldn't resolve default property of object g_pMetRunoffRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                .BindRaster(g_pMetRunoffRaster, "met_run")
        '                'UPGRADE_WARNING: Couldn't resolve default property of object g_pRunoffRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                .BindRaster(g_pRunoffRaster, "accrun")
        '            End With

        '            If g_pMetRunoffRaster Is Nothing Then
        '                MsgBox("g_pMetRunoff is crap")
        '            End If

        '            If g_pRunoffRaster Is Nothing Then
        '                MsgBox("g_pRunoff is crap")
        '            End If

        '            strExpression = "[met_run] + [accrun]"

        '            pTemp2PollRaster = pMapAlgebraOp.Execute(strExpression)

        '            With pMapAlgebraOp
        '                .UnbindRaster("met_run")
        '                .UnbindRaster("accrun")
        '            End With
        '            'END STEP 6: ------------------------------------------------------------------------------
        '        End If

        '        modProgDialog.ProgDialog("Calculating final concentration...", strTitle, 0, 13, 8, (frmPrj.m_App.hwnd))
        '        If modProgDialog.g_boolCancel Then

        '            'STEP 7: FINAL CONCENTRATION ---------------------------------------------------------------
        '            With pMapAlgebraOp
        '                'UPGRADE_WARNING: Couldn't resolve default property of object pTemp1PollRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                .BindRaster(pTemp1PollRaster, "temp1")
        '                'UPGRADE_WARNING: Couldn't resolve default property of object pTemp2PollRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                .BindRaster(pTemp2PollRaster, "temp2")
        '            End With

        '            strExpression = "[temp1] / [temp2]"

        '            pTotalPollConcRaster = pMapAlgebraOp.Execute(strExpression)

        '            With pMapAlgebraOp
        '                .UnbindRaster("temp1")
        '                .UnbindRaster("temp2")
        '            End With
        '            'END STEP 7: --------------------------------------------------------------------------------
        '        End If

        '        modProgDialog.ProgDialog("Calculating final concentration...", strTitle, 0, 13, 9, (frmPrj.m_App.hwnd))
        '        If modProgDialog.g_boolCancel Then

        '            'STEP 8: FINAL CONCENTRATION -remove all noData values ---------------------------------------
        '            With pMapAlgebraOp
        '                'UPGRADE_WARNING: Couldn't resolve default property of object g_pDEMRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                .BindRaster(g_pDEMRaster, "dem")
        '                'UPGRADE_WARNING: Couldn't resolve default property of object pTotalPollConcRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                .BindRaster(pTotalPollConcRaster, "totalConc")
        '            End With

        '            strExpression = "Merge([totalConc], Con([dem] >= 0, 0))"

        '            pTotalPollConc0Raster = pMapAlgebraOp.Execute(strExpression)

        '            With pMapAlgebraOp
        '                .UnbindRaster("dem")
        '                .UnbindRaster("totalConc")
        '            End With
        '            'END STEP 7: --------------------------------------------------------------------------------
        '        End If

        '        modProgDialog.ProgDialog("Converting to correct units...", strTitle, 0, 13, 10, (frmPrj.m_App.hwnd))

        '        Dim pClipTotalConcRaster As ESRI.ArcGIS.Geodatabase.IRaster
        '        If modProgDialog.g_boolCancel Then
        '            modProgDialog.ProgDialog("Creating data layer...", strTitle, 0, 13, 11, (frmPrj.m_App.hwnd))

        '            'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            strOutConc = modUtil.GetUniqueName("conc", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))

        '            If g_booSelectedPolys Then
        '                pClipTotalConcRaster = modUtil.ClipBySelectedPoly(pTotalPollConc0Raster, g_pSelectedPolyClip, pEnv)
        '                'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                pPermTotalConcRaster = modUtil.ReturnPermanentRaster(pClipTotalConcRaster, pEnv.OutWorkspace.PathName, strOutConc)
        '                'UPGRADE_NOTE: Object pClipTotalConcRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '                pClipTotalConcRaster = Nothing
        '            Else
        '                'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                pPermTotalConcRaster = modUtil.ReturnPermanentRaster(pTotalPollConc0Raster, pEnv.OutWorkspace.PathName, strOutConc)
        '            End If

        '            pPollRasterLayer = modUtil.ReturnRasterLayer((frmPrj.m_App), pPermTotalConcRaster, m_strPollName & " Conc. (mg/L)")
        '            pPollRasterLayer.Renderer = modUtil.ReturnRasterStretchColorRampRender(pPollRasterLayer, m_strColor)

        '            'UPGRADE_WARNING: Couldn't resolve default property of object pPollRasterLayer.Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            pPollRasterLayer.Visible = False

        '            'UPGRADE_WARNING: Couldn't resolve default property of object pPollRasterLayer.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            g_dicMetadata.Add(pPollRasterLayer.Name, m_strPollCoeffMetadata)

        '            'UPGRADE_WARNING: Couldn't resolve default property of object pPollRasterLayer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            g_pGroupLayer.Add(pPollRasterLayer)

        '        End If

        '        'Cleanup
        '        'UPGRADE_NOTE: Object pLandSampleRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pLandSampleRaster = Nothing
        '        'UPGRADE_NOTE: Object pMassVolumeRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pMassVolumeRaster = Nothing
        '        'UPGRADE_NOTE: Object pPollMassRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pPollMassRaster = Nothing
        '        'UPGRADE_NOTE: Object pAccumPollRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pAccumPollRaster = Nothing
        '        'UPGRADE_NOTE: Object pTemp1PollRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pTemp1PollRaster = Nothing
        '        'UPGRADE_NOTE: Object pTemp2PollRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pTemp2PollRaster = Nothing
        '        'UPGRADE_NOTE: Object pTotalPollConcRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pTotalPollConcRaster = Nothing
        '        'UPGRADE_NOTE: Object pMapAlgebraOp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pMapAlgebraOp = Nothing
        '        'UPGRADE_NOTE: Object pEnv may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pEnv = Nothing
        '        'UPGRADE_NOTE: Object pTotalPollConc0Raster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pTotalPollConc0Raster = Nothing

        '        modProgDialog.ProgDialog("Comparing to water quality standard...", strTitle, 0, 13, 13, (frmPrj.m_App.hwnd))

        '        If modProgDialog.g_boolCancel Then
        '            If Not CompareWaterQuality(g_pWaterShedFeatClass, pPermTotalConcRaster) Then
        '                CalcPollutantConcentration = False
        '                Exit Function
        '            End If
        '        End If

        '        'if we get to the end
        '        CalcPollutantConcentration = True

        '        'UPGRADE_NOTE: Object pPermTotalConcRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pPermTotalConcRaster = Nothing

        '        modProgDialog.KillDialog()

        '        Exit Function


        'ErrHandler:
        '        If Err.Number = -2147217297 Then 'User cancelled operation
        '            modProgDialog.g_boolCancel = False
        '            CalcPollutantConcentration = False
        '            Exit Function
        '        ElseIf Err.Number = -2147467259 Then
        '            MsgBox("ArcMap has reached the maximum number of GRIDs allowed in memory.  " & "Please exit N-SPECT and restart ArcMap.", MsgBoxStyle.Information, "Maximum GRID Number Encountered")
        '            modProgDialog.g_boolCancel = False
        '            modProgDialog.KillDialog()
        '            CalcPollutantConcentration = False
        '        Else
        '            HandleError(False, "CalcPollutantConcentration " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        '            modProgDialog.g_boolCancel = False
        '            modProgDialog.KillDialog()
        '            CalcPollutantConcentration = False
        '        End If

    End Function

    Private Function CompareWaterQuality(ByRef pWSFeatureClass As MapWinGIS.Shapefile, ByRef pPollutantRaster As MapWinGIS.Grid) As Boolean
        '        Dim strWQVAlue As Object

        '        'Get the zone dataset from the first layer in ArcMap
        '        Dim pGeoPollDS As ESRI.ArcGIS.Geodatabase.IGeoDataset
        '        Dim pLocalOp As ESRI.ArcGIS.SpatialAnalyst.ILocalOp
        '        Dim pEnv As ESRI.ArcGIS.GeoAnalyst.IRasterAnalysisEnvironment
        '        Dim pWS As ESRI.ArcGIS.Geodatabase.IWorkspace
        '        Dim pMaxRaster As ESRI.ArcGIS.Geodatabase.IRaster
        '        Dim pConRaster As ESRI.ArcGIS.Geodatabase.IRaster
        '        Dim pPermWQRaster As ESRI.ArcGIS.Geodatabase.IRaster
        '        Dim pWQRasterLayer As ESRI.ArcGIS.Carto.IRasterLayer
        '        Dim pMapAlgebraOp As ESRI.ArcGIS.SpatialAnalyst.IMapAlgebraOp
        '        Dim dblConvertValue As Double
        '        Dim strOutWQ As String
        '        Dim strExpression As String
        '        Dim strMetadata As String

        '        On Error GoTo ErrorHandler

        '        'Create the value raster from the pollutant Raster
        '        pGeoPollDS = pPollutantRaster

        '        'Create a Spatial operator
        '        pLocalOp = New ESRI.ArcGIS.SpatialAnalyst.RasterLocalOp

        '        'Set output workspace and environ
        '        pEnv = g_pSpatEnv
        '        pEnv = pLocalOp
        '        'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        pWS = pEnv.OutWorkspace

        '        ' Perform Spatial operation
        '        pMaxRaster = pLocalOp.LocalStatistics(pGeoPollDS, ESRI.ArcGIS.GeoAnalyst.esriGeoAnalysisStatisticsEnum.esriGeoAnalysisStatsMaximum)

        '        'UPGRADE_WARNING: Couldn't resolve default property of object strWQVAlue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        strWQVAlue = ReturnWQValue(m_strPollName, m_strWQName)

        '        'UPGRADE_WARNING: Couldn't resolve default property of object strWQVAlue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        dblConvertValue = (CDbl(strWQVAlue)) / 1000

        '        'Now run water quality
        '        pEnv = g_pSpatEnv
        '        pMapAlgebraOp = New ESRI.ArcGIS.SpatialAnalyst.RasterMapAlgebraOp
        '        pEnv = pMapAlgebraOp

        '        With pMapAlgebraOp
        '            'UPGRADE_WARNING: Couldn't resolve default property of object pMaxRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            .BindRaster(pMaxRaster, "Max")
        '            'UPGRADE_WARNING: Couldn't resolve default property of object g_pFlowAccRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            .BindRaster(g_pFlowAccRaster, "flowacc")
        '        End With

        '        'This rather ugly expression was set up to check for meets/exceed water quality standards for
        '        'only the streams.  It takes the values of flowaccumulation from watershed delineation fame that
        '        'exceed values of greater than 1%.  Then multiplies the result (all cells representing streams) times
        '        'the water quality grid.
        '        strExpression = "(Con([Max] gt " & CStr(dblConvertValue) & ", 1, 2)) * (con([flowacc] > (" & CStr(modUtil.ReturnRasterMax(g_pFlowAccRaster)) & " * 0.01), 1))"
        '        pConRaster = pMapAlgebraOp.Execute(strExpression)

        '        pMapAlgebraOp.UnbindRaster("Max")
        '        pMapAlgebraOp.UnbindRaster("flowacc")

        '        'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        strOutWQ = modUtil.GetUniqueName("wq", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))

        '        'Clip if selectedpolys
        '        Dim pClipWQRaster As ESRI.ArcGIS.Geodatabase.IRaster
        '        If g_booSelectedPolys Then
        '            pClipWQRaster = modUtil.ClipBySelectedPoly(pConRaster, g_pSelectedPolyClip, pEnv)
        '            'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            pPermWQRaster = modUtil.ReturnPermanentRaster(pClipWQRaster, pEnv.OutWorkspace.PathName, strOutWQ)
        '            'UPGRADE_NOTE: Object pClipWQRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '            pClipWQRaster = Nothing
        '        Else
        '            'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            pPermWQRaster = modUtil.ReturnPermanentRaster(pConRaster, pEnv.OutWorkspace.PathName, strOutWQ)
        '        End If

        '        pWQRasterLayer = modUtil.ReturnRasterLayer((frmPrj.m_App), pPermWQRaster, m_strPollName & " Standard: " & CStr(dblConvertValue) & " mg/L")
        '        pWQRasterLayer.Renderer = modUtil.ReturnUniqueRasterRenderer(pWQRasterLayer, m_strWQName)
        '        'UPGRADE_WARNING: Couldn't resolve default property of object pWQRasterLayer.Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        pWQRasterLayer.Visible = False


        '        strMetadata = vbTab & "Water Quality Standard:" & vbNewLine & vbTab & vbTab & "Criteria Name: " & m_strWQName & vbNewLine & vbTab & vbTab & "Standard: " & dblConvertValue & " mg/L"


        '        'UPGRADE_WARNING: Couldn't resolve default property of object pWQRasterLayer.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        g_dicMetadata.Add(pWQRasterLayer.Name, m_strPollCoeffMetadata & strMetadata)

        '        'UPGRADE_WARNING: Couldn't resolve default property of object pWQRasterLayer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        g_pGroupLayer.Add(pWQRasterLayer)

        '        CompareWaterQuality = True

        '        'Cleanup
        '        'UPGRADE_NOTE: Object pGeoPollDS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pGeoPollDS = Nothing
        '        'UPGRADE_NOTE: Object pLocalOp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pLocalOp = Nothing
        '        'UPGRADE_NOTE: Object pEnv may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pEnv = Nothing
        '        'UPGRADE_NOTE: Object pWS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pWS = Nothing
        '        'UPGRADE_NOTE: Object pMaxRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pMaxRaster = Nothing
        '        'UPGRADE_NOTE: Object pConRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pConRaster = Nothing
        '        'UPGRADE_NOTE: Object pPermWQRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pPermWQRaster = Nothing
        '        'UPGRADE_NOTE: Object pMapAlgebraOp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        pMapAlgebraOp = Nothing

        '        Exit Function

        'ErrorHandler:
        '        If Err.Number = -2147467259 Then
        '            MsgBox("ArcMap has reached the maximum number of GRIDs allowed in memory.  " & "Please exit N-SPECT and restart ArcMap.", MsgBoxStyle.Information, "Maximum GRID Number Encountered")
        '            CompareWaterQuality = False
        '            modProgDialog.g_boolCancel = False
        '            modProgDialog.KillDialog()
        '        Else
        '            HandleError(False, "CompareWaterQuality " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        '            CompareWaterQuality = False
        '            modProgDialog.g_boolCancel = False
        '            modProgDialog.KillDialog()
        '        End If
    End Function


    Private Function ReturnWQValue(ByRef strPollName As String, ByRef strWQstdName As String) As String

        '        On Error GoTo ErrHandler
        '        Dim rsPoll As New ADODB.Recordset
        '        Dim rsWQstd As New ADODB.Recordset

        '        Dim strPoll As String
        '        Dim strWQStd As String

        '        strPoll = "Select * from Pollutant where name like '" & strPollName & "'"

        '        rsPoll.Open(strPoll, g_ADOConn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

        '        strWQStd = "SELECT * FROM WQCRITERIA INNER JOIN POLL_WQCRITERIA ON WQCRITERIA.WQCRITID = POLL_WQCRITERIA.WQCRITID " & "WHERE WQCRITERIA.NAME Like '" & strWQstdName & "' AND POLL_WQCRITERIA.POLLID = " & rsPoll.Fields("POLLID").Value

        '        rsWQstd.Open(strWQStd, g_ADOConn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

        '        ReturnWQValue = CStr(rsWQstd.Fields("Threshold").Value)

        '        rsPoll.Close()
        '        rsWQstd.Close()

        '        'UPGRADE_NOTE: Object rsPoll may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        rsPoll = Nothing
        '        'UPGRADE_NOTE: Object rsWQstd may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        rsWQstd = Nothing

        '        Exit Function

        'ErrHandler:
        '        MsgBox("Error in ADO pollutant part: " & Err.Number & vbNewLine & Err.Description & vbNewLine & strWQStd)
    End Function
End Module