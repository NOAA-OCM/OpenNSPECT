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

    Private _WQValue As Single
    Private _FlowMax As Single

    ' Constant used by the Error handler function - DO NOT REMOVE
    Const c_sModuleFileName As String = "modPollutantCalcs.vb"
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
                For i = 1 To maxVal

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
                            strpick = "0"
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

        Dim pMassVolumeRaster As MapWinGIS.Grid = Nothing
        Dim pPermMassVolumeRaster As MapWinGIS.Grid = Nothing
        Dim pAccumPollRaster As MapWinGIS.Grid = Nothing
        Dim pPermAccPollRaster As MapWinGIS.Grid = Nothing
        Dim pPermTotalConcRaster As MapWinGIS.Grid = Nothing
        Dim pTotalPollConc0Raster As MapWinGIS.Grid = Nothing  'gets rid of no data...replace with 0

        'String to hold calculations
        Dim strTitle As String
        strTitle = "Processing " & _strPollName & " Conc. Calculation..."
        Dim strOutConc As String
        Dim strAccPoll As String

        Try
            modProgDialog.ProgDialog("Calculating Mass Volume...", strTitle, 0, 13, 2, 0)
            If modProgDialog.g_boolCancel Then
                'STEP 2: MASS OF PHOSPHORUS PRODUCED BY EACH CELL -----------------------------------------
                ReDim _picks(strConStatement.Split(",").Length)
                _picks = strConStatement.Split(",")
                Dim massvolcalc As New RasterMathCellCalc(AddressOf massvolCellCalc)
                RasterMath(g_LandCoverRaster, g_pMetRunoffRaster, Nothing, Nothing, Nothing, pMassVolumeRaster, massvolcalc)

                'END STEP 2: -------------------------------------------------------------------------------
            End If

            'LOCAL EFFECTS ONLY...
            'At this point the above grid will satisfy 'local effects only' people so...
            If g_booLocalEffects Then

                modProgDialog.ProgDialog("Creating data layer for local effects...", strTitle, 0, 13, 13, 0)
                If modProgDialog.g_boolCancel Then

                    strOutConc = modUtil.GetUniqueName("locconc", g_strWorkspace, ".bgd")
                    'Added 7/23/04 to account for clip by selected polys functionality
                    If g_booSelectedPolys Then
                        pPermMassVolumeRaster = modUtil.ClipBySelectedPoly(pMassVolumeRaster, g_pSelectedPolyClip, strOutConc)
                    Else
                        pPermMassVolumeRaster = modUtil.ReturnPermanentRaster(pMassVolumeRaster, strOutConc)
                    End If

                    g_dicMetadata.Add(_strPollName & "Local Effects (mg)", _strPollCoeffMetadata)

                    Dim cs As MapWinGIS.GridColorScheme = ReturnRasterStretchColorRampCS(pPermMassVolumeRaster, _strColor)
                    Dim lyr As MapWindow.Interfaces.Layer = g_MapWin.Layers.Add(pPermMassVolumeRaster, cs, _strPollName & " Local Effects (mg)")
                    lyr.Visible = False
                    lyr.MoveTo(0, g_pGroupLayer)

                    CalcPollutantConcentration = True

                End If

                modProgDialog.KillDialog()
                Exit Function

            End If

            modProgDialog.ProgDialog("Deriving accumulated pollutant...", strTitle, 0, 13, 3, 0)
            If modProgDialog.g_boolCancel Then
                'STEP 3: DERIVE ACCUMULATED POLLUTANT ------------------------------------------------------

                'Use weightedaread8 from geoproc to accum, then rastercalc to multiply this out
                Dim pTauD8Flow As MapWinGIS.Grid = Nothing

                Dim tauD8calc As New RasterMathCellCalcNulls(AddressOf tauD8CellCalc)
                RasterMath(g_pFlowDirRaster, Nothing, Nothing, Nothing, Nothing, pTauD8Flow, Nothing, False, tauD8calc)
                pTauD8Flow.Header.NodataValue = -1

                Dim strtmp1 As String = IO.Path.GetTempFileName + ".bgd"
                MapWinGeoProc.DataManagement.DeleteGrid(strtmp1)
                pTauD8Flow.Save(strtmp1)

                Dim strtmp2 As String = IO.Path.GetTempFileName + ".bgd"
                MapWinGeoProc.DataManagement.DeleteGrid(strtmp2)
                pMassVolumeRaster.Save(strtmp2)

                Dim strtmpout As String = IO.Path.GetTempFileName + "out.bgd"
                MapWinGeoProc.DataManagement.DeleteGrid(strtmpout)


                'Use geoproc weightedAreaD8 after converting the D8 grid to taudem format bgd if needed
                MapWinGeoProc.Hydrology.WeightedAreaD8(strtmp1, strtmp2, "", strtmpout, False, False, Nothing)
                'strExpression = "FlowAccumulation([flowdir], [met_run], FLOAT)"

                Dim tmpGrid As New MapWinGIS.Grid
                tmpGrid.Open(strtmpout)

                Dim multAccumcalc As New RasterMathCellCalc(AddressOf multAccumCellCalc)
                RasterMath(tmpGrid, Nothing, Nothing, Nothing, Nothing, pAccumPollRaster, multAccumcalc)

                pTauD8Flow.Close()
                MapWinGeoProc.DataManagement.DeleteGrid(strtmp1)

                'END STEP 3: ------------------------------------------------------------------------------
            End If

            'STEP 3a: Added 7/26: ADD ACCUMULATED POLLUTANT TO GROUP LAYER-----------------------------------
            modProgDialog.ProgDialog("Creating accumlated pollutant layer...", strTitle, 0, 13, 4, 0)
            If modProgDialog.g_boolCancel Then
                strAccPoll = modUtil.GetUniqueName("accpoll", g_strWorkspace, ".bgd")
                'Added 7/23/04 to account for clip by selected polys functionality
                If g_booSelectedPolys Then
                    pPermAccPollRaster = modUtil.ClipBySelectedPoly(pAccumPollRaster, g_pSelectedPolyClip, strAccPoll)
                Else
                    pPermAccPollRaster = modUtil.ReturnPermanentRaster(pAccumPollRaster, strAccPoll)
                End If

                g_dicMetadata.Add("Accumulated " & _strPollName & " (kg)", _strPollCoeffMetadata)

                Dim cs As MapWinGIS.GridColorScheme = ReturnRasterStretchColorRampCS(pPermAccPollRaster, _strColor)
                Dim lyr As MapWindow.Interfaces.Layer = g_MapWin.Layers.Add(pPermAccPollRaster, cs, "Accumulated " & _strPollName & " (kg)")
                lyr.Visible = False
                lyr.MoveTo(0, g_pGroupLayer)

            End If
            'END STEP 3a: ---------------------------------------------------------------------------------


            modProgDialog.ProgDialog("Calculating final concentration...", strTitle, 0, 13, 9, 0)
            If modProgDialog.g_boolCancel Then
                Dim AllConCalc As New RasterMathCellCalcNulls(AddressOf AllConCellCalc)
                RasterMath(pMassVolumeRaster, pAccumPollRaster, g_pMetRunoffRaster, g_pRunoffRaster, g_pDEMRaster, pTotalPollConc0Raster, Nothing, False, AllConCalc)
            End If


            If modProgDialog.g_boolCancel Then
                modProgDialog.ProgDialog("Creating data layer...", strTitle, 0, 13, 11, 0)

                strOutConc = modUtil.GetUniqueName("conc", g_strWorkspace, ".bgd")

                If g_booSelectedPolys Then
                    pPermTotalConcRaster = modUtil.ClipBySelectedPoly(pTotalPollConc0Raster, g_pSelectedPolyClip, strOutConc)
                Else
                    pPermTotalConcRaster = modUtil.ReturnPermanentRaster(pTotalPollConc0Raster, strOutConc)
                End If

                g_dicMetadata.Add(_strPollName & " Conc. (mg/L)", _strPollCoeffMetadata)
                Dim cs As MapWinGIS.GridColorScheme = ReturnRasterStretchColorRampCS(pPermTotalConcRaster, _strColor)
                Dim lyr As MapWindow.Interfaces.Layer = g_MapWin.Layers.Add(pPermTotalConcRaster, cs, _strPollName & " Conc. (mg/L)")
                lyr.Visible = False
                lyr.MoveTo(0, g_pGroupLayer)
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
        Dim pClipWQRaster As MapWinGIS.Grid = Nothing
        Dim pPermWQRaster As MapWinGIS.Grid = Nothing
        Dim pWQRasterLayer As MapWinGIS.Grid = Nothing
        Dim dblConvertValue As Double
        Dim strOutWQ As String
        Dim strExpression As String = ""
        Dim strMetadata As String

        Try

            ' Perform Spatial operation
            'TODO: This seems useless on a singleband thing, otherwise, seems random. so skipping it.
            'pMaxRaster = pLocalOp.LocalStatistics(pPollutantRaster, ESRI.ArcGIS.GeoAnalyst.esriGeoAnalysisStatisticsEnum.esriGeoAnalysisStatsMaximum)

            strWQVAlue = ReturnWQValue(_strPollName, _strWQName)

            _WQValue = (CDbl(strWQVAlue)) / 1000
            _FlowMax = g_pFlowAccRaster.Maximum
    
            Dim concalc As New RasterMathCellCalc(AddressOf concompCellCalc)
            RasterMath(pPollutantRaster, g_pFlowAccRaster, Nothing, Nothing, Nothing, pConRaster, concalc)


            strOutWQ = modUtil.GetUniqueName("wq", g_strWorkspace, ".bgd")

            'Clip if selectedpolys
            If g_booSelectedPolys Then
                pPermWQRaster = modUtil.ClipBySelectedPoly(pConRaster, g_pSelectedPolyClip, strOutWQ)
            Else
                pPermWQRaster = modUtil.ReturnPermanentRaster(pConRaster, strOutWQ)
            End If

            strMetadata = vbTab & "Water Quality Standard:" & vbNewLine & vbTab & vbTab & "Criteria Name: " & _strWQName & vbNewLine & vbTab & vbTab & "Standard: " & dblConvertValue & " mg/L"
            g_dicMetadata.Add(_strPollName & " Standard: " & CStr(dblConvertValue) & " mg/L", _strPollCoeffMetadata & strMetadata)

            Dim cs As MapWinGIS.GridColorScheme = modUtil.ReturnUniqueRasterRenderer(pPermWQRaster, _strWQName)
            Dim lyr As MapWindow.Interfaces.Layer = g_MapWin.Layers.Add(pPermWQRaster, cs, _strPollName & " Standard: " & CStr(dblConvertValue) & " mg/L")
            lyr.Visible = False
            lyr.MoveTo(0, g_pGroupLayer)

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
    Private Function massvolCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        Dim tmpval As Single
        'strexpression = pick([pLandSampleRaster], _picks)"
        For i As Integer = 0 To _picks.Length - 1
            If Input1 = i + 1 Then
                tmpval = _picks(i)
                Exit For
            End If
        Next

        'strExpression = "[met_runoff] * [pollmass]"
        Return Input2 * tmpval
    End Function

    Private Function multAccumCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        'strExpression = "(FlowAccumulation([flowdir], [massvolume], FLOAT)) * 1.0e-6"
        Return Input1 * 0.000001
    End Function

    Private Function AllConCellCalc(ByVal Input1 As Single, ByVal Input1Null As Single, ByVal Input2 As Single, ByVal Input2Null As Single, ByVal Input3 As Single, ByVal Input3Null As Single, ByVal Input4 As Single, ByVal Input4Null As Single, ByVal Input5 As Single, ByVal Input5Null As Single, ByVal OutNull As Single) As Single
        If Input1 = Input1Null Or Input2 = Input2Null Or Input3 = Input3Null Or Input4 = Input4Null Then
            If Input5 = Input5Null Then
                Return OutNull
            Else
                If Input5 > 0 Then
                    Return 0
                Else
                    Return OutNull
                End If
            End If
        Else
            Return (Input1 + (Input2 / 0.000001)) / (Input3 + Input4)
        End If

    End Function


    Private Function concompCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        'This rather ugly expression was set up to check for meets/exceed water quality standards for
        'only the streams.  It takes the values of flowaccumulation from watershed delineation fame that
        'exceed values of greater than 1%.  Then multiplies the result (all cells representing streams) times
        'the water quality grid.
        'strExpression = "(Con([Max] gt " & CStr(dblConvertValue) & ", 1, 2)) * (con([flowacc] > (" & CStr(modUtil.ReturnRasterMax(g_pFlowAccRaster)) & " * 0.01), 1))"
        'strExpression = "(Con([Max] gt _WQValue, 1, 2)) * (con([flowacc] > (_FlowMax * 0.01), 1))"
        If Input1 > _WQValue Then
            '(con([flowacc] > (" & CStr(modUtil.ReturnRasterMax(g_pFlowAccRaster)) & " * 0.01), 1))
            If Input2 > (_FlowMax * 0.01) Then
                Return 1
            Else
                Return OutNull
            End If
        Else
            If Input2 > (_FlowMax * 0.01) Then
                Return 2
            Else
                Return OutNull
            End If
        End If
    End Function

#End Region


End Module