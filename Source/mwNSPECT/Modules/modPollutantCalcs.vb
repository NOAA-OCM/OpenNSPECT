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

    Private _strPollName As String 'Mod level variable for name of pollutant being used
    Private _strWQName As String 'Mod level variable for name of water quality standard
    Private _strColor As String 'Mod level variable holding the string of the pollutant color
    Private _strPollCoeffMetadata As String 'Variable to hold coeffs for use in metadata

    ' Constant used by the Error handler function - DO NOT REMOVE
    Const c_sModuleFileName As String = "modPollutantCalcs.bas"
    Private _ParentHWND As Integer ' Set this to get correct parenting of Error handler forms

    Private _picks() As String

    Public Function PollutantConcentrationSetup(ByRef clsPollutant As clsXMLPollutantItem, ByRef strLandClass As String, ByRef strWQName As String) As Boolean
        'Sub takes incoming parameters (in the form of a pollutant item) from the project file
        'and then parses them out
        Try
            'Open Strings
            Dim strPoll As String
            Dim strType As String
            Dim strField As String = ""
            Dim strConStatement As String = ""
            Dim strTempCoeffSet As String 'Again, because of landuse, we have to check for 'temp' coeff sets and their use
            Dim strPollColor As String

            'Get the name of the pollutant
            _strPollName = clsPollutant.strPollName

            'Get the name of the Water Quality Standard
            _strWQName = strWQName

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
            If g_DictTempNames.Count > 0 Then
                If Len(g_DictTempNames.Item(clsPollutant.strCoeffSet)) > 0 Then
                    strTempCoeffSet = g_DictTempNames.Item(clsPollutant.strCoeffSet)
                Else
                    strTempCoeffSet = clsPollutant.strCoeffSet
                End If
            Else
                strTempCoeffSet = clsPollutant.strCoeffSet
            End If

            If Len(strField) > 0 Then
                strPoll = "SELECT * FROM COEFFICIENTSET WHERE NAME LIKE '" & strTempCoeffSet & "'"
                Dim cmdPoll As New OleDbCommand(strPoll, g_DBConn)
                Dim dataPoll As OleDbDataReader = cmdPoll.ExecuteReader()
                dataPoll.Read()
                strType = "SELECT LCCLASS.Value, LCCLASS.Name, COEFFICIENT." & strField & " As CoeffType, COEFFICIENT.CoeffID, COEFFICIENT.LCCLASSID " & "FROM LCCLASS LEFT OUTER JOIN COEFFICIENT " & "ON LCCLASS.LCCLASSID = COEFFICIENT.LCCLASSID " & "WHERE COEFFICIENT.COEFFSETID = " & dataPoll("CoeffSetID") & " ORDER BY LCCLASS.VALUE"
                dataPoll.Close()
                Dim cmdType As New OleDbCommand(strType, g_DBConn)

                strConStatement = ConstructPickStatment(cmdType, g_LandCoverRaster)
                _strPollCoeffMetadata = ConstructMetaData(cmdType, (clsPollutant.strCoeff), g_booLocalEffects)

            End If

            'Find out the color of the pollutant
            strPollColor = "Select Color from Pollutant where NAME LIKE '" & _strPollName & "'"
            Dim cmdPollColor As New OleDbCommand(strPollColor, g_DBConn)
            Dim datapollcolor As OleDbDataReader = cmdPollColor.ExecuteReader()
            datapollcolor.Read()
            _strColor = CStr(datapollcolor("Color"))
            datapollcolor.Close()

            If CalcPollutantConcentration(strConStatement) Then
                PollutantConcentrationSetup = True
            Else
                PollutantConcentrationSetup = False
            End If

        Catch ex As Exception
            HandleError(True, "PollutantConcentrationSetup " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, _ParentHWND)
            PollutantConcentrationSetup = False
        End Try
    End Function

    Private Function ConstructMetaData(ByRef cmdType As oledbcommand, ByRef strCoeffSet As String, ByRef booLocal As Boolean) As String
        'Takes the rs and creates a string describing the pollutants and coefficients used in this run, will
        'later be added to the global dictionary

        Dim strConstructMetaData As String
        Dim strLandClassCoeff As String = ""
        Dim strHeader As String
        Dim i As Short

        If booLocal Then
            strHeader = vbTab & "Input Datasets:" & vbNewLine & vbTab & vbTab & "Hydrologic soils grid: " & g_clsXMLPrjFile.strSoilsHydFileName & vbNewLine & vbTab & vbTab & "Landcover grid: " & g_clsXMLPrjFile.strLCGridFileName & vbNewLine & vbTab & vbTab & "Landcover grid type: " & g_clsXMLPrjFile.strLCGridType & vbNewLine & vbTab & vbTab & "Landcover grid units: " & g_clsXMLPrjFile.strLCGridUnits & vbNewLine & vbTab & vbTab & "Precipitation grid: " & g_strPrecipFileName & vbNewLine

        Else
            strHeader = vbTab & "Input Datasets:" & vbNewLine & vbTab & vbTab & "Hydrologic soils grid: " & g_clsXMLPrjFile.strSoilsHydFileName & vbNewLine & vbTab & vbTab & "Landcover grid: " & g_clsXMLPrjFile.strLCGridFileName & vbNewLine & vbTab & vbTab & "Precipitation grid: " & g_strPrecipFileName & vbNewLine & vbTab & vbTab & "Flow direction grid: " & g_strFlowDirFilename & vbNewLine
        End If

        strConstructMetaData = vbTab & "Pollutant Coefficients:" & vbNewLine & vbTab & vbTab & "Pollutant: " & _strPollName & vbNewLine & vbTab & vbTab & "Coefficient Set: " & strCoeffSet & vbNewLine & vbTab & vbTab & "The following lists the landcover classes and associated coefficients used" & vbNewLine & vbTab & vbTab & "in the N-SPECT analysis run that created this dataset: " & vbNewLine

        Dim dataType As OleDbDataReader = cmdType.ExecuteReader()
        While dataType.Read()
            If i = 1 Then
                strLandClassCoeff = vbTab & vbTab & vbTab & dataType("Name") & ": " & dataType("CoeffType") & vbNewLine
            Else
                strLandClassCoeff = strLandClassCoeff & vbTab & vbTab & vbTab & dataType("Name") & ": " & dataType("CoeffType") & vbNewLine
            End If
        End While
        dataType.Close()

        ConstructMetaData = strHeader & g_strLandCoverParameters & strConstructMetaData & strLandClassCoeff

    End Function

    Private Function ConstructPickStatment(ByRef cmdType As OleDbCommand, ByRef pLCRaster As MapWinGIS.Grid) As String
        'Creates the initial pick statement using the name of the the LandCass [CCAP, for example]
        'and the Land Class Raster.  Returns a string

        ConstructPickStatment = ""
        Try
            Dim strCon As String = "" 'Con statement base
            Dim strParens As String = "" 'String of trailing parens
            'Dim strCompleteCon As String 'Concatenate of strCon & strParens

            Dim TableExist As Boolean
            Dim FieldIndex As Short
            Dim booValueFound As Boolean
            Dim i As Short
            Dim maxVal As Integer = pLCRaster.Maximum

            Dim tablepath As String = ""
            'Get the raster table
            Dim lcPath As String = pLCRaster.Filename
            If IO.Path.GetFileName(lcPath) = "sta.adf" Then
                tablepath = IO.Path.GetDirectoryName(lcPath) + ".dbf"
                If IO.File.Exists(tablepath) Then

                    TableExist = True
                Else
                    TableExist = False
                End If
            Else
                tablepath = IO.Path.ChangeExtension(lcPath, ".dbf")
                If IO.File.Exists(tablepath) Then
                    TableExist = True
                Else
                    TableExist = False
                End If
            End If

            Dim strpick As String = ""

            Dim mwTable As New MapWinGIS.Table
            If Not TableExist Then
                MsgBox("No MapWindow-readable raster table was found. To create one using ArcMap 9.3+, add the raster to the default project, right click on its layer and select Open Attribute Table. Now click on the options button in the lower right and select Export. In the export path, navigate to the directory of the grid folder and give the export the name of the raster folder with the .dbf extension. i.e. if you are exporting a raster attribute table from a raster named landcover, export landcover.dbf into the same level directory as the folder.", MsgBoxStyle.Exclamation, "Raster Attribute Table Not Found")

                Return ""
            Else
                mwTable.Open(tablepath)

                'Get index of Value Field
                FieldIndex = -1
                For fidx As Integer = 0 To mwTable.NumFields - 1
                    If mwTable.Field(fidx).Name.ToLower = "value" Then
                        FieldIndex = fidx
                        Exit For
                    End If
                Next

                Dim rowidx As Integer = 0
                Dim dataType As OleDbDataReader
                For i = 0 To maxVal

                    If i = 1 Then
                        If (mwTable.CellValue(FieldIndex, rowidx) = i) Then
                            dataType = cmdType.ExecuteReader
                            booValueFound = False
                            While dataType.Read()
                                If mwTable.CellValue(FieldIndex, rowidx) = dataType("Value") Then
                                    booValueFound = True
                                    strpick = CStr(dataType("CoeffType"))
                                    rowidx = rowidx + 1
                                    Exit While
                                Else
                                    booValueFound = False
                                End If
                            End While
                            If booValueFound = False Then
                                MsgBox("Error: Your N-SPECT Land Class Table is missing values found in your landcover GRID dataset.")
                                ConstructPickStatment = Nothing
                                dataType.Close()
                                mwTable.Close()
                                Exit Function
                            End If
                            dataType.Close()
                        Else
                            strpick = "-9999"
                        End If
                    Else
                        If (mwTable.CellValue(FieldIndex, rowidx) = i) Then 'And (pRow.Value(FieldIndex) = rsLandClass!Value) Then
                            dataType = cmdType.ExecuteReader

                            booValueFound = False
                            While dataType.Read()
                                If mwTable.CellValue(FieldIndex, rowidx) = dataType("Value") Then
                                    booValueFound = True
                                    strpick = strpick & ", " & CStr(dataType("CoeffType"))
                                    rowidx = rowidx + 1
                                    Exit While
                                Else
                                    booValueFound = False
                                End If
                            End While
                            If booValueFound = False Then
                                MsgBox("Error: Your N-SPECT Land Class Table is missing values found in your landcover GRID dataset.")
                                ConstructPickStatment = Nothing
                                dataType.Close()
                                mwTable.Close()
                                Exit Function
                            End If
                            dataType.Close()

                        Else
                            strpick = strpick & ", 0"
                        End If
                    End If
                Next

                'For row As Integer = 0 To mwTable.NumRows - 1
                '    booValueFound = False
                '    While dataType.Read()
                '        If dataType("Value") = mwTable.CellValue(FieldIndex, row) Then

                '            booValueFound = True

                '            If strCon = "" Then
                '                strCon = "Con(([nu_lulc] eq " & mwTable.CellValue(FieldIndex, row) & "), " & dataType("CoeffType") & ", "
                '            Else
                '                strCon = strCon & "Con(([nu_lulc] eq " & mwTable.CellValue(FieldIndex, row) & "), " & dataType("CoeffType") & ", "
                '            End If

                '            If strParens = "" Then
                '                strParens = "-9999)"
                '            Else
                '                strParens = strParens & ")"
                '            End If

                '            Exit While
                '        Else
                '            booValueFound = False
                '        End If
                '    End While

                '    If booValueFound = False Then
                '        MsgBox("Values in table LCClass table not equal to values in landclass dataset.")
                '        ConstructPickStatment = ""
                '        Exit Function
                '    Else
                '        i = 0
                '    End If
                'Next
                'dataType.Close()
            End If

            'strCompleteCon = strCon & strParens
            'ConstructPickStatment = strCompleteCon
            ConstructPickStatment = strpick

        Catch ex As Exception
            MsgBox("Error in pick Statement: " & Err.Number & ": " & Err.Description)
        End Try
    End Function

    Private Function CalcPollutantConcentration(ByRef strConStatement As String) As Boolean

        Dim pLandSampleRaster As MapWinGIS.Grid = Nothing
        Dim pPollMassRaster As MapWinGIS.Grid = Nothing
        Dim pMassVolumeRaster As MapWinGIS.Grid = Nothing
        Dim pPermMassVolumeRaster As MapWinGIS.Grid = Nothing
        Dim pAccumPollRaster As MapWinGIS.Grid = Nothing
        Dim pTemp1PollRaster As MapWinGIS.Grid = Nothing
        Dim pTemp2PollRaster As MapWinGIS.Grid = Nothing
        Dim pTotalPollConcRaster As MapWinGIS.Grid = Nothing
        Dim pPermAccPollRaster As MapWinGIS.Grid = Nothing
        Dim pPermTotalConcRaster As MapWinGIS.Grid = Nothing
        Dim pTotalPollConc0Raster As MapWinGIS.Grid = Nothing  'gets rid of no data...replace with 0

        'Dim pPollRasterLayer As ESRI.ArcGIS.Carto.IRasterLayer
        'Dim pAccPollRasterLayer As ESRI.ArcGIS.Carto.IRasterLayer

        'String to hold calculations
        Dim strTitle As String
        strTitle = "Processing " & _strPollName & " Conc. Calculation..."
        Dim strOutConc As String
        Dim strAccPoll As String

        Try
            modProgDialog.ProgDialog("Checking landcover cell size...", strTitle, 0, 13, 1, 0)
            If modProgDialog.g_boolCancel Then
                'Step 1a: ----------------------------------------------------------------------------
                'Make sure LandCover is in the same cellsize as the global environment
                Dim lchead As MapWinGIS.GridHeader = g_LandCoverRaster.Header

                If lchead.dX <> g_dblCellSize Then
                    'TODO: Resample g_LandCoverRaster to correct cellsize
                Else
                    pLandSampleRaster = g_LandCoverRaster
                End If
            End If

            modProgDialog.ProgDialog("Calculating EMC GRID...", strTitle, 0, 13, 1, 0)
            If modProgDialog.g_boolCancel Then

                'Step 1: CREATE PHOSPHORUS EMC GRID AT CELL LEVEL -----------------------------------------

                ReDim _picks(strConStatement.Split(",").Length)
                _picks = strConStatement.Split(",")

                Dim pollmasscalc As New RasterMathCellCalc(AddressOf pollmassCellCalc)
                RasterMath(pLandSampleRaster, Nothing, Nothing, Nothing, Nothing, pPollMassRaster, pollmasscalc)

                'END STEP 1: ------------------------------------------------------------------------------
            End If


            modProgDialog.ProgDialog("Calculating Mass Volume...", strTitle, 0, 13, 2, 0)
            If modProgDialog.g_boolCancel Then
                'STEP 2: MASS OF PHOSPHORUS PRODUCED BY EACH CELL -----------------------------------------
                Dim massvolcalc As New RasterMathCellCalc(AddressOf massvolCellCalc)
                RasterMath(g_pMetRunoffRaster, pPollMassRaster, Nothing, Nothing, Nothing, pMassVolumeRaster, massvolcalc)

                'END STEP 2: -------------------------------------------------------------------------------
            End If

            'LOCAL EFFECTS ONLY...
            'At this point the above grid will satisfy 'local effects only' people so...
            Dim pClipAccPollRaster As MapWinGIS.Grid
            If g_booLocalEffects Then

                modProgDialog.ProgDialog("Creating data layer for local effects...", strTitle, 0, 13, 13, 0)
                If modProgDialog.g_boolCancel Then

                    strOutConc = modUtil.GetUniqueName("locconc", g_strWorkspace, ".tif")
                    'Added 7/23/04 to account for clip by selected polys functionality
                    If g_booSelectedPolys Then
                        'TODO: get these functions working
                        'pClipAccPollRaster = modUtil.ClipBySelectedPoly(pMassVolumeRaster, g_pSelectedPolyClip, pEnv)
                        'pPermMassVolumeRaster = modUtil.ReturnPermanentRaster(pClipAccPollRaster, pEnv.OutWorkspace.PathName, strOutConc)
                    Else
                        'TODO get this working
                        'pPermMassVolumeRaster = modUtil.ReturnPermanentRaster(pMassVolumeRaster, pEnv.OutWorkspace.PathName, strOutConc)
                    End If

                    g_dicMetadata.Add(_strPollName & "Local Effects (mg)", _strPollCoeffMetadata)
                    'TODO: Add to the group map as nonvisible
                    'pPollRasterLayer = modUtil.ReturnRasterLayer((frmPrj.m_App), pPermMassVolumeRaster, _strPollName & " Local Effects (mg)")
                    'pPollRasterLayer.Renderer = modUtil.ReturnRasterStretchColorRampRender(pPollRasterLayer, _strColor)
                    'pPollRasterLayer.Visible = False
                    'g_pGroupLayer.Add(pPollRasterLayer)

                    CalcPollutantConcentration = True

                End If

                modProgDialog.KillDialog()
                Exit Function

            End If

            modProgDialog.ProgDialog("Deriving accumulated pollutant...", strTitle, 0, 13, 3, 0)
            If modProgDialog.g_boolCancel Then
                'STEP 3: DERIVE ACCUMULATED POLLUTANT ------------------------------------------------------

                'TODO: Use weightedaread8 from geoproc to accum, then rastercalc to multiply this out
                'With pMapAlgebraOp
                '    .BindRaster(g_pFlowDirRaster, "flowdir")
                '    .BindRaster(pMassVolumeRaster, "massvolume")
                'End With

                'strExpression = "(FlowAccumulation([flowdir], [massvolume], FLOAT)) * 1.0e-6"

                'pAccumPollRaster = pMapAlgebraOp.Execute(strExpression)

                'END STEP 3: ------------------------------------------------------------------------------
            End If

            'STEP 3a: Added 7/26: ADD ACCUMULATED POLLUTANT TO GROUP LAYER-----------------------------------
            modProgDialog.ProgDialog("Creating accumlated pollutant layer...", strTitle, 0, 13, 4, 0)
            Dim pClipAccPoll2Raster As MapWinGIS.Grid
            If modProgDialog.g_boolCancel Then
                strAccPoll = modUtil.GetUniqueName("accpoll", g_strWorkspace, ".tif")
                'Added 7/23/04 to account for clip by selected polys functionality
                If g_booSelectedPolys Then
                    'TODO: Get this working
                    'pClipAccPoll2Raster = modUtil.ClipBySelectedPoly(pAccumPollRaster, g_pSelectedPolyClip, pEnv)
                    'pPermAccPollRaster = modUtil.ReturnPermanentRaster(pClipAccPoll2Raster, pEnv.OutWorkspace.PathName, strAccPoll)
                Else
                    'TODO
                    'pPermAccPollRaster = modUtil.ReturnPermanentRaster(pAccumPollRaster, pEnv.OutWorkspace.PathName, strAccPoll)
                End If


                g_dicMetadata.Add("Accumulated " & _strPollName & " (kg)", _strPollCoeffMetadata)

                'TODO: Add raster to group and etc.
                'pAccPollRasterLayer = modUtil.ReturnRasterLayer((frmPrj.m_App), pPermAccPollRaster, "Accumulated " & _strPollName & " (kg)")
                'pAccPollRasterLayer.Renderer = modUtil.ReturnRasterStretchColorRampRender(pAccPollRasterLayer, m_strColor)
                'pAccPollRasterLayer.Visible = False
                'g_pGroupLayer.Add(pAccPollRasterLayer)

            End If
            'END STEP 3a: ---------------------------------------------------------------------------------


            modProgDialog.ProgDialog("Deriving total concentration at each cell...", strTitle, 0, 13, 6, 0)
            If modProgDialog.g_boolCancel Then
                'STEP 5: DERIVE TOTAL CONCENTRATION (AT EACH CELL) ----------------------------------------
                Dim temp1calc As New RasterMathCellCalc(AddressOf temp1CellCalc)
                RasterMath(pMassVolumeRaster, pAccumPollRaster, Nothing, Nothing, Nothing, pTemp1PollRaster, temp1calc)
                'END STEP 5: -------------------------------------------------------------------------------
            End If

            modProgDialog.ProgDialog("Adding metric runoff...", strTitle, 0, 13, 7, 0)
            If modProgDialog.g_boolCancel Then
                'STEP 6: -----------------------------------------------------------------------------------
                Dim temp2calc As New RasterMathCellCalc(AddressOf temp2CellCalc)
                RasterMath(g_pMetRunoffRaster, g_pRunoffRaster, Nothing, Nothing, Nothing, pTemp2PollRaster, temp2calc)
               
                'END STEP 6: ------------------------------------------------------------------------------
            End If

            modProgDialog.ProgDialog("Calculating final concentration...", strTitle, 0, 13, 8, 0)
            If modProgDialog.g_boolCancel Then
                'STEP 7: FINAL CONCENTRATION ---------------------------------------------------------------
                Dim totconcalc As New RasterMathCellCalc(AddressOf totconCellCalc)
                RasterMath(pTemp1PollRaster, pTemp2PollRaster, Nothing, Nothing, Nothing, pTotalPollConcRaster, totconcalc)
                'END STEP 7: --------------------------------------------------------------------------------
            End If

            modProgDialog.ProgDialog("Calculating final concentration...", strTitle, 0, 13, 9, 0)
            If modProgDialog.g_boolCancel Then
                'STEP 8: FINAL CONCENTRATION -remove all noData values ---------------------------------------
                Dim totconnonullcalc As New RasterMathCellCalcNulls(AddressOf totconnonullCellCalc)
                RasterMath(g_pDEMRaster, pTotalPollConcRaster, Nothing, Nothing, Nothing, pTotalPollConc0Raster, Nothing, False, totconnonullcalc)
                'END STEP 7: --------------------------------------------------------------------------------
            End If

            modProgDialog.ProgDialog("Converting to correct units...", strTitle, 0, 13, 10, 0)

            Dim pClipTotalConcRaster As MapWinGIS.Grid
            If modProgDialog.g_boolCancel Then
                modProgDialog.ProgDialog("Creating data layer...", strTitle, 0, 13, 11, 0)

                strOutConc = modUtil.GetUniqueName("conc", g_strWorkspace, "tif")

                If g_booSelectedPolys Then
                    'TODO: again
                    'pClipTotalConcRaster = modUtil.ClipBySelectedPoly(pTotalPollConc0Raster, g_pSelectedPolyClip, pEnv)
                    'pPermTotalConcRaster = modUtil.ReturnPermanentRaster(pClipTotalConcRaster, pEnv.OutWorkspace.PathName, strOutConc)
                Else
                    'TODO: Again
                    'pPermTotalConcRaster = modUtil.ReturnPermanentRaster(pTotalPollConc0Raster, pEnv.OutWorkspace.PathName, strOutConc)
                End If

                g_dicMetadata.Add(_strPollName & " Conc. (mg/L)", _strPollCoeffMetadata)
                'TODO: Add layer
                'pPollRasterLayer = modUtil.ReturnRasterLayer((frmPrj.m_App), pPermTotalConcRaster, m_strPollName & " Conc. (mg/L)")
                'pPollRasterLayer.Renderer = modUtil.ReturnRasterStretchColorRampRender(pPollRasterLayer, m_strColor)
                'pPollRasterLayer.Visible = False
                'g_pGroupLayer.Add(pPollRasterLayer)
            End If

            modProgDialog.ProgDialog("Comparing to water quality standard...", strTitle, 0, 13, 13, 0)

            If modProgDialog.g_boolCancel Then
                If Not CompareWaterQuality(g_pWaterShedFeatClass, pPermTotalConcRaster) Then
                    CalcPollutantConcentration = False
                    Exit Function
                End If
            End If

            'if we get to the end
            CalcPollutantConcentration = True

            modProgDialog.KillDialog()

        Catch ex As Exception
            If Err.Number = -2147217297 Then 'User cancelled operation
                modProgDialog.g_boolCancel = False
                CalcPollutantConcentration = False
                Exit Function
            ElseIf Err.Number = -2147467259 Then
                MsgBox("ArcMap has reached the maximum number of GRIDs allowed in memory.  " & "Please exit N-SPECT and restart ArcMap.", MsgBoxStyle.Information, "Maximum GRID Number Encountered")
                modProgDialog.g_boolCancel = False
                modProgDialog.KillDialog()
                CalcPollutantConcentration = False
            Else
                HandleError(False, "CalcPollutantConcentration " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, _ParentHWND)
                modProgDialog.g_boolCancel = False
                modProgDialog.KillDialog()
                CalcPollutantConcentration = False
            End If
        End Try
    End Function

    Private Function CompareWaterQuality(ByRef pWSFeatureClass As MapWinGIS.Shapefile, ByRef pPollutantRaster As MapWinGIS.Grid) As Boolean
        Dim strWQVAlue As Object

        'Get the zone dataset from the first layer in ArcMap
        Dim pMaxRaster As MapWinGIS.Grid = Nothing
        Dim pConRaster As MapWinGIS.Grid = Nothing
        Dim pPermWQRaster As MapWinGIS.Grid = Nothing
        Dim pWQRasterLayer As MapWinGIS.Grid = Nothing
        Dim dblConvertValue As Double
        Dim strOutWQ As String
        Dim strExpression As String = ""
        Dim strMetadata As String

        Try

            ' Perform Spatial operation
            'TODO: Find what this does
            'pMaxRaster = pLocalOp.LocalStatistics(pPollutantRaster, ESRI.ArcGIS.GeoAnalyst.esriGeoAnalysisStatisticsEnum.esriGeoAnalysisStatsMaximum)

            strWQVAlue = ReturnWQValue(_strPollName, _strWQName)

            dblConvertValue = (CDbl(strWQVAlue)) / 1000

    
            'TODO: once the above max thing is figured out
            'With pMapAlgebraOp
            '    .BindRaster(pMaxRaster, "Max")
            '    .BindRaster(g_pFlowAccRaster, "flowacc")
            'End With

            ''This rather ugly expression was set up to check for meets/exceed water quality standards for
            ''only the streams.  It takes the values of flowaccumulation from watershed delineation fame that
            ''exceed values of greater than 1%.  Then multiplies the result (all cells representing streams) times
            ''the water quality grid.
            'strExpression = "(Con([Max] gt " & CStr(dblConvertValue) & ", 1, 2)) * (con([flowacc] > (" & CStr(modUtil.ReturnRasterMax(g_pFlowAccRaster)) & " * 0.01), 1))"
            'pConRaster = pMapAlgebraOp.Execute(strExpression)


            strOutWQ = modUtil.GetUniqueName("wq", g_strWorkspace, ".tif")

            'Clip if selectedpolys
            Dim pClipWQRaster As MapWinGIS.Grid
            If g_booSelectedPolys Then
                'TODO: again
                'pClipWQRaster = modUtil.ClipBySelectedPoly(pConRaster, g_pSelectedPolyClip, pEnv)
                'pPermWQRaster = modUtil.ReturnPermanentRaster(pClipWQRaster, pEnv.OutWorkspace.PathName, strOutWQ)
            Else
                'TODO
                'pPermWQRaster = modUtil.ReturnPermanentRaster(pConRaster, pEnv.OutWorkspace.PathName, strOutWQ)
            End If

            strMetadata = vbTab & "Water Quality Standard:" & vbNewLine & vbTab & vbTab & "Criteria Name: " & _strWQName & vbNewLine & vbTab & vbTab & "Standard: " & dblConvertValue & " mg/L"
            g_dicMetadata.Add(_strPollName & " Standard: " & CStr(dblConvertValue) & " mg/L", _strPollCoeffMetadata & strMetadata)

            'TODO: Add raster
            'pWQRasterLayer = modUtil.ReturnRasterLayer((frmPrj.m_App), pPermWQRaster, m_strPollName & " Standard: " & CStr(dblConvertValue) & " mg/L")
            'pWQRasterLayer.Renderer = modUtil.ReturnUniqueRasterRenderer(pWQRasterLayer, m_strWQName)
            'pWQRasterLayer.Visible = False
            'g_pGroupLayer.Add(pWQRasterLayer)

            CompareWaterQuality = True


        Catch ex As Exception
            If Err.Number = -2147467259 Then
                MsgBox("ArcMap has reached the maximum number of GRIDs allowed in memory.  " & "Please exit N-SPECT and restart ArcMap.", MsgBoxStyle.Information, "Maximum GRID Number Encountered")
                CompareWaterQuality = False
                modProgDialog.g_boolCancel = False
                modProgDialog.KillDialog()
            Else
                HandleError(False, "CompareWaterQuality " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, 0)
                CompareWaterQuality = False
                modProgDialog.g_boolCancel = False
                modProgDialog.KillDialog()
            End If
        End Try
    End Function

    Private Function ReturnWQValue(ByRef strPollName As String, ByRef strWQstdName As String) As String
        Dim strPoll As String
        Dim strWQStd As String = ""
        ReturnWQValue = ""
        Try

            strPoll = "Select * from Pollutant where name like '" & strPollName & "'"
            Dim cmdpoll As New OleDbCommand(strPoll, g_DBConn)
            Dim datapoll As OleDbDataReader = cmdpoll.ExecuteReader
            datapoll.Read()
            strWQStd = "SELECT * FROM WQCRITERIA INNER JOIN POLL_WQCRITERIA ON WQCRITERIA.WQCRITID = POLL_WQCRITERIA.WQCRITID " & "WHERE WQCRITERIA.NAME Like '" & strWQstdName & "' AND POLL_WQCRITERIA.POLLID = " & datapoll("POLLID")
            datapoll.Close()

            Dim cmdWQ As New OleDbCommand(strWQStd, g_DBConn)
            Dim datawq As OleDbDataReader = cmdWQ.ExecuteReader()
            datawq.Read()
            ReturnWQValue = CStr(datawq("Threshold"))
            datawq.Close()
        Catch ex As Exception
            MsgBox("Error in ADO pollutant part: " & Err.Number & vbNewLine & Err.Description & vbNewLine & strWQStd)
        End Try
    End Function



#Region "Raster Math"
    Private Function pollmassCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        For i As Integer = 0 To _picks.Length - 1
            If Input1 = i + 1 Then
                Return _picks(i)
            End If
        Next
    End Function

    Private Function massvolCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'strExpression = "[met_runoff] * [pollmass]"
        Return Input1 * Input2
    End Function

    Private Function temp1CellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'strExpression = "[massvolume] + ([accpoll] / 1.0e-6)"
        Return Input1 + (Input2 / 0.000001)
    End Function

    Private Function temp2CellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'strExpression = "[met_run] + [accrun]"
        Return Input1 + Input2
    End Function

    Private Function totconCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'strExpression = "[temp1] / [temp2]"
        Return Input1 / Input2
    End Function

    Private Function totconnonullCellCalc(ByVal Input1 As Single, ByVal Input1Null As Single, ByVal Input2 As Single, ByVal Input2Null As Single, ByVal Input3 As Single, ByVal Input3Null As Single, ByVal Input4 As Single, ByVal Input4Null As Single, ByVal Input5 As Single, ByVal Input5Null As Single) As Single
        'strExpression = "Merge([totalConc], Con([dem] >= 0, 0))"
        If Input1 <> Input1Null Then
            Return Input1
        Else
            If Input2 > 0 Then
                Return 0
            Else
                Return Input1
            End If
        End If
    End Function

#End Region


End Module