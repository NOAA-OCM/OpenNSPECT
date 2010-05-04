Imports System.Data.OleDb
Module modRusle
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

    Public g_NibbleRaster As MapWinGIS.Grid 'Nibble Raster
    Public g_DEMTwoCellRaster As MapWinGIS.Grid 'Two Cell buffer of the DEM
    Public g_RFactorRaster As MapWinGIS.Grid 'R Factor Raster
    Private _strRusleMetadata As String 'Metadata holder
    Private _booUsingConstantValue As Boolean 'Boolean value
    Private _dblRFactorConstant As Double 'Constant for R-factor
    Private _strSDRFileName As String 'If user provides own SDR GRid, store path here
    Private _picks As String()

    Public Function RUSLESetup(ByRef strNibbleFileName As String, ByRef strDEMTwoCellFileName As String, ByRef strRFactorFileName As String, ByRef strKfactorFileName As String, ByRef strSDRFileName As String, ByRef strLandClass As String, Optional ByRef dblRFactorConstant As Double = 0) As Boolean
        'Sub takes incoming parameters from the project file and then parses them out
        'strNibbleFileName: FileName of the nibble GRID
        'strDEMTwoCellFileName: FileName of the two cell buffered DEM
        'strRFactorFileName: FileName of the R Factor GRID
        'strLandClass: Name of the Landclass we're using

        'Open Strings
        Dim strCovFactor As String
        Dim strConStatement As String
        Dim strError As String = ""
        Dim strTempLCType As String

        'STEP 1: Set the Nibble Raster ----------------------------------------------------------------
        If modUtil.RasterExists(strNibbleFileName) Then
            g_NibbleRaster = modUtil.ReturnRaster(strNibbleFileName)
        Else
            strError = "Nibble Raster Does Not Exist: " & strNibbleFileName
        End If
        'END STEP 1: -----------------------------------------------------------------------------------

        'STEP 2: Set the DEMTwoCell Raster -------------------------------------------------------------
        If modUtil.RasterExists(strDEMTwoCellFileName) Then
            g_DEMTwoCellRaster = modUtil.ReturnRaster(strDEMTwoCellFileName)
        Else
            strError = "DEM Two Cell Buffer Raster Does Not Exist: " & strDEMTwoCellFileName
        End If
        'END STEP 2: -----------------------------------------------------------------------------------

        'STEP 3: Set the R Factor Raster ---------------------------------------------------------------
        If Len(strRFactorFileName) > 0 Then
            If modUtil.RasterExists(strRFactorFileName) Then
                g_RFactorRaster = modUtil.ReturnRaster(strRFactorFileName)
            Else
                strError = "R Factor Raster Does Not Exist: " & strRFactorFileName
            End If
            _booUsingConstantValue = False
        Else
            _booUsingConstantValue = True
            _dblRFactorConstant = dblRFactorConstant
        End If

        'END STEP 3: -----------------------------------------------------------------------------------

        'STEP 4: Set the K Factor Raster
        If modUtil.RasterExists(strKfactorFileName) Then
            g_KFactorRaster = modUtil.ReturnRaster(strKfactorFileName)
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

        strCovFactor = "SELECT LCTYPE.LCTYPEID, LCCLASS.NAME, LCCLASS.VALUE, LCCLASS.COVERFACTOR FROM " & "LCTYPE INNER JOIN LCCLASS ON LCTYPE.LCTYPEID = LCCLASS.LCTYPEID " & "WHERE LCTYPE.NAME LIKE '" & strTempLCType & "' ORDER BY LCCLASS.VALUE"
        Dim cmdCov As New OleDbCommand(strCovFactor, g_DBConn)

        If Len(strError) > 0 Then
            MsgBox(strError)
            Exit Function
        End If

        'Get the con statement for the cover factor calculation
        strConStatement = ConstructPickStatment(cmdCov, g_LandCoverRaster)

        'Are they using SDR
        _strSDRFileName = Trim(strSDRFileName)

        'Metadata time
        _strRusleMetadata = CreateMetadata(g_booLocalEffects)

        'Calc rusle using the con
        If CalcRUSLE(strConStatement) Then
            RUSLESetup = True
        Else
            RUSLESetup = False
        End If


    End Function

    Private Function CreateMetadata(ByRef booLocal As Boolean) As String

        Dim strHeader As String
        'Dim i As Integer
        'Dim strCFactor As String

        'Set up the header w/or without flow direction
        If booLocal = True Then
            strHeader = vbTab & "Input Datasets:" & vbNewLine & vbTab & vbTab & "Hydrologic soils grid: " & g_clsXMLPrjFile.strSoilsHydFileName & vbNewLine & vbTab & vbTab & "Landcover grid: " & g_clsXMLPrjFile.strLCGridFileName & vbNewLine & vbTab & vbTab & "Precipitation grid: " & g_strPrecipFileName & vbNewLine & vbTab & vbTab & "Soils K Factor: " & g_clsXMLPrjFile.strSoilsKFileName & vbNewLine & vbTab & vbTab & "LS Factor grid: " & g_strLSFileName & vbNewLine & vbTab & vbTab & "R Factor grid: " & g_clsXMLPrjFile.strRainGridFileName & vbNewLine & g_strLandCoverParameters & vbNewLine 'append the g_strLandCoverParameters that was set up during runoff
        Else
            strHeader = vbTab & "Input Datasets:" & vbNewLine & vbTab & vbTab & "Hydrologic soils grid: " & g_clsXMLPrjFile.strSoilsHydFileName & vbNewLine & vbTab & vbTab & "Landcover grid: " & g_clsXMLPrjFile.strLCGridFileName & vbNewLine & vbTab & vbTab & "Precipitation grid: " & g_strPrecipFileName & vbNewLine & vbTab & vbTab & "Soils K Factor: " & g_clsXMLPrjFile.strSoilsKFileName & vbNewLine & vbTab & vbTab & "LS Factor grid: " & g_strLSFileName & vbNewLine & vbTab & vbTab & "R Factor grid: " & g_clsXMLPrjFile.strRainGridFileName & vbNewLine & vbTab & vbTab & "Flow direction grid: " & g_strFlowDirFilename & vbNewLine & g_strLandCoverParameters & vbNewLine
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
                    If (mwTable.CellValue(FieldIndex, rowidx) = i) Then 'And (pRow.Value(FieldIndex) = rsLandClass!Value) Then
                        dataType = cmdType.ExecuteReader

                        booValueFound = False
                        While dataType.Read()
                            If mwTable.CellValue(FieldIndex, rowidx) = dataType("Value") Then
                                booValueFound = True
                                If strpick = "" Then
                                    strpick = CStr(dataType("CoverFactor"))
                                Else
                                    strpick = strpick & ", " & CStr(dataType("CoverFactor"))
                                End If
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
                        If strpick = "" Then
                            strpick = "-9999"
                        Else
                            strpick = strpick & ", 0"
                        End If
                    End If

                Next
            End If

            'strCompleteCon = strCon & strParens
            'ConstructPickStatment = strCompleteCon
            ConstructPickStatment = strpick

        Catch ex As Exception
            MsgBox("Error in pick Statement: " & Err.Number & ": " & Err.Description)
        End Try
    End Function

    'OBSOLETE
    'Private Function ConstructConStatment(ByRef rsCF As ADODB.Recordset, ByRef pLCRaster As ESRI.ArcGIS.Geodatabase.IRaster) As String
    '    'Creates the initial pick statement using the name of the the LandCass [CCAP, for example]
    '    'and the Land Class Raster.  Returns a string

    '    Dim strCon As String 'Con statement base
    '    Dim strParens As String 'String of trailing parens
    '    Dim strCompleteCon As String 'Concatenate of strCon & strParens

    '    Dim pRasterCol As ESRI.ArcGIS.DataSourcesRaster.IRasterBandCollection
    '    Dim pBand As ESRI.ArcGIS.DataSourcesRaster.IRasterBand
    '    Dim pTable As ESRI.ArcGIS.Geodatabase.ITable
    '    Dim TableExist As Boolean
    '    Dim pCursor As ESRI.ArcGIS.Geodatabase.ICursor
    '    Dim pRow As ESRI.ArcGIS.Geodatabase.IRow
    '    Dim FieldIndex As Short
    '    Dim booValueFound As Boolean
    '    Dim i As Short

    '    'STEP 1:  get the records from the database -----------------------------------------------
    '    rsCF.MoveFirst()
    '    'End Database stuff

    '    'STEP 2: Raster Values ---------------------------------------------------------------------
    '    'Now Get the RASTER values
    '    'Get Rasterband from the incoming raster
    '    pRasterCol = pLCRaster
    '    pBand = pRasterCol.Item(0)

    '    'Get the raster table
    '    pBand.HasTable(TableExist)
    '    If Not TableExist Then Exit Function

    '    pTable = pBand.AttributeTable
    '    'Get All rows
    '    pCursor = pTable.Search(Nothing, True)
    '    'Init pRow
    '    pRow = pCursor.NextRow

    '    'Get index of Value Field
    '    FieldIndex = pTable.FindField("Value")

    '    'REMOVED 11/30/2007 in favor of method below.
    '    'STEP 3: Table values vs. Raster Values Count - if not equal bark --------------------------
    '    '    If pTable.RowCount(Nothing) <> rsCF.RecordCount Then
    '    '        MsgBox "Error: The number of records in your Land Class Table do not match your GRID dataset."
    '    '        Exit Function
    '    '    End If


    '    Do While Not pRow Is Nothing
    '        booValueFound = False
    '        rsCF.MoveFirst()

    '        For i = 0 To rsCF.RecordCount - 1
    '            If pRow.Value(FieldIndex) = rsCF.Fields("Value").Value Then

    '                booValueFound = True

    '                If strCon = "" Then
    '                    strCon = "Con(([nu_lulc] eq " & pRow.Value(FieldIndex) & "), " & rsCF.Fields("CoverFactor").Value & ", "
    '                Else
    '                    strCon = strCon & "Con(([nu_lulc] eq " & pRow.Value(FieldIndex) & "), " & rsCF.Fields("CoverFactor").Value & ", "
    '                End If

    '                If strParens = "" Then
    '                    strParens = "-9999)"
    '                Else
    '                    strParens = strParens & ")"
    '                End If

    '                Exit For
    '            Else
    '                booValueFound = False
    '            End If
    '            rsCF.MoveNext()
    '        Next i

    '        If booValueFound = False Then
    '            MsgBox("Error: Your N-SPECT Land Class Table is missing values found in your landcover GRID dataset.")
    '            ConstructConStatment = ""
    '            Exit Function
    '        Else
    '            pRow = pCursor.NextRow
    '            i = 0
    '        End If

    '    Loop

    '    'REMOVED 11/30/2007 in favor of method above
    '    '    'STEP 4: Create the strings
    '    '    'Loop through and get all values
    '    '    Do While Not pRow Is Nothing
    '    '        If pRow.Value(FieldIndex) = rsCF!Value Then
    '    '            If strCon = "" Then
    '    '                strCon = "Con(([nu_lulc] eq " & pRow.Value(FieldIndex) & "), " & rsCF!CoverFactor & ", "
    '    '            Else
    '    '                strCon = strCon & "Con(([nu_lulc] eq " & pRow.Value(FieldIndex) & "), " & rsCF!CoverFactor & ", "
    '    '            End If
    '    '
    '    '            If strParens = "" Then
    '    '                strParens = "-9999)"
    '    '            Else
    '    '                strParens = strParens & ")"
    '    '            End If
    '    '
    '    '            rsCF.MoveNext
    '    '            Set pRow = pCursor.NextRow
    '    '
    '    '        Else
    '    '            MsgBox "Values in table LCClass table not equal to values in landclass dataset."
    '    '            ConstructConStatment = ""
    '    '            Exit Function
    '    '        End If
    '    '    Loop


    '    strCompleteCon = strCon & strParens
    '    Debug.Print(strCompleteCon)

    '    ConstructConStatment = strCompleteCon

    '    'Cleanup:
    '    'Set pLCRaster = Nothing
    '    pRasterCol = Nothing
    '    pBand = Nothing
    '    pTable = Nothing
    '    pCursor = Nothing
    '    pRow = Nothing

    'End Function

    Private Function CalcRUSLE(ByRef strConStatement As String) As Boolean

        Dim pLandSampleRaster As MapWinGIS.Grid = Nothing 'LC Cover sampleized
        Dim pCFactorRaster As MapWinGIS.Grid = Nothing  'C Factor
        Dim pSoilLossRaster As MapWinGIS.Grid = Nothing  'Soil Loss
        Dim pSoilLossAcres As MapWinGIS.Grid = Nothing  'Soil Loss Acres
        Dim pZSedDelRaster As MapWinGIS.Grid = Nothing  'And I quote, Dave's Whacky Sediment Delivery Ratio
        Dim pCellAreaKMRaster As MapWinGIS.Grid = Nothing  'Cell Area KM
        Dim pDASedDelRaster As MapWinGIS.Grid = Nothing  'da_sed_del
        Dim pTemp3Raster As MapWinGIS.Grid = Nothing  'Temp3
        Dim pTemp4Raster As MapWinGIS.Grid = Nothing  'Temp4
        Dim pTemp5Raster As MapWinGIS.Grid = Nothing  'Temp5
        Dim pTemp6Raster As MapWinGIS.Grid = Nothing  'Temp6
        Dim pSDRRaster As MapWinGIS.Grid = Nothing  'SDR
        Dim pSedYieldRaster As MapWinGIS.Grid = Nothing  'Sediment Yield
        Dim pAccumSedRaster As MapWinGIS.Grid = Nothing  'Accumulated Sed Raster
        Dim pTotalAccumSedRaster As MapWinGIS.Grid = Nothing  'Total accumulated sediment raster
        Dim pPermAccumSedRaster As MapWinGIS.Grid = Nothing  'Permanent accumulated sediment raster
        Dim pPermRUSLELocRaster As MapWinGIS.Grid = Nothing  'RUSLE Local Effects raster

        'String to hold calculations
        Dim strExpression As String = ""
        Const strTitle As String = "Processing RUSLE Calculation..."
        Dim strOutYield As String


        Try
            modProgDialog.ProgDialog("Checking landcover cell size...", strTitle, 0, 13, 1, 0)
            If modProgDialog.g_boolCancel Then
                'Step 1a: ----------------------------------------------------------------------------
                'Make sure LandCover is in the same cellsize as the global environment
                Dim headlc As MapWinGIS.GridHeader = g_LandCoverRaster.Header

                If headlc.dX <> g_dblCellSize Then
                    'TODO: resample landcover raster
                Else
                    pLandSampleRaster = g_LandCoverRaster
                End If
            End If

            modProgDialog.ProgDialog("Calculating Cover Magagement Factor GRID...", strTitle, 0, 13, 2, 0)
            If modProgDialog.g_boolCancel Then

                'Step 1: CREATE Management cover Factor GRID ---------------------------------------------

                ReDim _picks(strConStatement.Split(",").Length)
                _picks = strConStatement.Split(",")

                Dim factcalc As New RasterMathCellCalc(AddressOf factCellCalc)
                RasterMath(pLandSampleRaster, Nothing, Nothing, Nothing, Nothing, pCFactorRaster, factcalc)

                'END STEP 1: ------------------------------------------------------------------------------
            End If

            modProgDialog.ProgDialog("Solving RUSLE Equation...", strTitle, 0, 13, 3, 0)
            If modProgDialog.g_boolCancel Then
                'STEP 2: SOLVE RUSLE EQUATION -------------------------------------------------------------

                Dim soillosscalc As New RasterMathCellCalc(AddressOf soillossCellCalc)
                If Not _booUsingConstantValue Then 'If not using a constant
                    RasterMath(g_pLSRaster, g_KFactorRaster, pCFactorRaster, g_RFactorRaster, Nothing, pSoilLossRaster, soillosscalc)
                    'strExpression = "[rfactor] * [kfactor] * [lsfactor] * [cfactor]"
                Else 'if using a constant
                    RasterMath(g_pLSRaster, g_KFactorRaster, pCFactorRaster, Nothing, Nothing, pSoilLossRaster, soillosscalc)
                    'strExpression = _dblRFactorConstant & " * [kfactor] * [lsfactor] * [cfactor]"
                End If

                'END STEP 2: -------------------------------------------------------------------------------
            End If

            modProgDialog.ProgDialog("Converting Soil Loss...", strTitle, 0, 13, 4, 0)
            If modProgDialog.g_boolCancel Then
                'STEP 3: DERIVE ACCUMULATED POLLUTANT ------------------------------------------------------
                Dim lossaccalc As New RasterMathCellCalc(AddressOf lossacCellCalc)
                RasterMath(pSoilLossRaster, Nothing, Nothing, Nothing, Nothing, pSoilLossAcres, lossaccalc)

                'END STEP 3: ------------------------------------------------------------------------------
            End If

            '***********************************************
            'BEGIN SDR CODE......
            '***********************************************
            If Len(Trim(_strSDRFileName)) = 0 Then
                modProgDialog.ProgDialog("Calculating Relief-Length Ratio for Sediment Delivery...", strTitle, 0, 13, 5, 0)
                If modProgDialog.g_boolCancel Then
                    'STEP 4: DAVE'S WACKY CALCULATION OF RELIEF-LENGTH RATIO FOR SEDIMENT DELIVERY RATIO-------
                    Dim pZSedcalc As New RasterMathCellCalc(AddressOf pZSedCellCalc)
                    RasterMath(g_NibbleRaster, g_DEMTwoCellRaster, Nothing, Nothing, Nothing, pZSedDelRaster, pZSedcalc)

                    'strExpression = "Con(([fdrnib] ge 0.5 and [fdrnib] lt 1.5), (([dem_2b] - [dem_2b](1,0)) / (" & g_dblCellSize & " * 0.001))," & "Con(([fdrnib] ge 1.5 and [fdrnib] lt 3.0), (([dem_2b] - [dem_2b](1,1)) / (" & g_dblCellSize & " * 0.0014142))," & "Con(([fdrnib] ge 3.0 and [fdrnib] lt 6.0), (([dem_2b] - [dem_2b](0,1)) / (" & g_dblCellSize & " * 0.001))," & "Con(([fdrnib] ge 6.0 and [fdrnib] lt 12.0), (([dem_2b] - [dem_2b](-1,1)) / (" & g_dblCellSize & " * 0.0014142))," & "Con(([fdrnib] ge 12.0 and [fdrnib] lt 24.0), (([dem_2b] - [dem_2b](-1,0)) / (" & g_dblCellSize & " * 0.001))," & "Con(([fdrnib] ge 24.0 and [fdrnib] lt 48.0), (([dem_2b] - [dem_2b](-1,-1)) / (" & g_dblCellSize & " * 0.0014142))," & "Con(([fdrnib] ge 48.0 and [fdrnib] lt 96.0), (([dem_2b] - [dem_2b](0,-1)) / (" & g_dblCellSize & " * 0.001))," & "Con(([fdrnib] ge 96.0 and [fdrnib] lt 192.0), (([dem_2b] - [dem_2b](1,-1)) / (" & g_dblCellSize & " * 0.0014142))," & "Con(([fdrnib] ge 192.0 and [fdrnib] le 255.0), (([dem_2b] - [dem_2b](1,0)) / (" & g_dblCellSize & " * 0.001))," & "0.1)))))))))"

                    'END STEP 4: ------------------------------------------------------------------------------
                End If

                modProgDialog.ProgDialog("Calculating Sediment Delivery Ratio...", strTitle, 0, 13, 6, 0)
                If modProgDialog.g_boolCancel Then
                    'STEP 5: CALCULATE SEDIMENT DELIVERY RATIO ------------------------------------------------
                    Dim areaKMcalc As New RasterMathCellCalc(AddressOf areaKMCellCalc)
                    RasterMath(g_pDEMRaster, Nothing, Nothing, Nothing, Nothing, pCellAreaKMRaster, areaKMcalc)
                    
                    'NOTE: Original equation is cellarea_km = ([cellsize] / 1000).  To get cell_size I do a CON on the DEM.
                    'Notice the float...if you don't do that, screw ville...GRID comes back = 0
                    'strExpression = "(float(con([DEM] >= 0, " & g_dblCellSize & ", 0))) / 1000"

                    'END STEP 5: -------------------------------------------------------------------------------
                End If

                modProgDialog.ProgDialog("Step 2 Sediment Delivery Ratio...", strTitle, 0, 13, 7, 0)
                If modProgDialog.g_boolCancel Then
                    'STEP 6: -----------------------------------------------------------------------------------
                    Dim pDAcalc As New RasterMathCellCalc(AddressOf pDACellCalc)
                    RasterMath(pCellAreaKMRaster, Nothing, Nothing, Nothing, Nothing, pDASedDelRaster, pDAcalc)
                    'strExpression = "Pow([cellarea_km], 2)"

                    'END STEP 6: ------------------------------------------------------------------------------
                End If

                modProgDialog.ProgDialog("Step 3 Sediment Delivery Ratio...", strTitle, 0, 13, 8, 0)
                If modProgDialog.g_boolCancel Then

                    'STEP 7: temp3 = Pow([da_sed_del], -0.998) ---------------------------------------------------
                    Dim temp3calc As New RasterMathCellCalc(AddressOf temp3CellCalc)
                    RasterMath(pDASedDelRaster, Nothing, Nothing, Nothing, Nothing, pTemp3Raster, temp3calc)
                    'strExpression = "Pow([da_sed_del], -0.0998)"
                    'END STEP 7: --------------------------------------------------------------------------------
                End If

                modProgDialog.ProgDialog("Step 4 Sediment Delivery Ratio...", strTitle, 0, 13, 9, 0)
                If modProgDialog.g_boolCancel Then

                    'STEP 8: temp4 = Pow([zl_sed_del], -0.0998) ---------------------------------------------------
                    Dim temp4calc As New RasterMathCellCalc(AddressOf temp4CellCalc)
                    RasterMath(pZSedDelRaster, Nothing, Nothing, Nothing, Nothing, pTemp4Raster, temp4calc)
                    'strExpression = "Pow([zl_sed_del], 0.3629)"

                    'END STEP 8: --------------------------------------------------------------------------------
                End If

                modProgDialog.ProgDialog("Step 5 Sediment Delivery Ratio...", strTitle, 0, 13, 10, 0)
                If modProgDialog.g_boolCancel Then

                    'STEP 9: temp5 = Pow([scsgrid100], 5.444) ---------------------------------------------------
                    Dim temp5calc As New RasterMathCellCalc(AddressOf temp5CellCalc)
                    RasterMath(g_pSCS100Raster, Nothing, Nothing, Nothing, Nothing, pTemp5Raster, temp5calc)
                    'strExpression = "Pow([scsgrid100], 5.444)"
                    'END STEP 9: --------------------------------------------------------------------------------
                End If

                modProgDialog.ProgDialog("Step 6 Sediment Delivery Ratio...", strTitle, 0, 13, 11, 0)
                If modProgDialog.g_boolCancel Then
                    'STEP 9: temp6 = 1.366 * [temp2] * [temp3] * [temp4] * [temp5] -------------------------------
                    Dim temp6calc As New RasterMathCellCalc(AddressOf temp6CellCalc)
                    RasterMath(pTemp3Raster, pTemp4Raster, pTemp5Raster, Nothing, Nothing, pTemp6Raster, temp6calc)
                    'strExpression = "1.366 * (Pow(10, -11)) * [temp3] * [temp4] * [temp5]"
                    'END STEP 9: --------------------------------------------------------------------------------
                End If

                modProgDialog.ProgDialog("Final Calculation Sediment Delivery Ratio...", strTitle, 0, 13, 12, 0)
                If modProgDialog.g_boolCancel Then

                    'STEP 10:
                    Dim SDRcalc As New RasterMathCellCalc(AddressOf SDRCellCalc)
                    RasterMath(pTemp6Raster, Nothing, Nothing, Nothing, Nothing, pSDRRaster, SDRcalc)
                    'strExpression = "Con(([temp6] gt 1), 1, [temp6])"
                    'END STEP 10: --------------------------------------------------------------------------------
                End If

            Else
                pSDRRaster = modUtil.ReturnRaster(_strSDRFileName)
            End If
            '********************************************************************
            'END SDR CALC
            '********************************************************************

            modProgDialog.ProgDialog("Applying Sediment Delivery Ratio...", strTitle, 0, 13, 13, 0)
            If modProgDialog.g_boolCancel Then
                'STEP 11: sed_yield = [soil_loss_ac] * [sdr] -------------------------------------------------
                Dim SedYieldcalc As New RasterMathCellCalc(AddressOf SedYieldCellCalc)
                RasterMath(pSDRRaster, pSoilLossAcres, Nothing, Nothing, Nothing, pSedYieldRaster, SedYieldcalc)
                'strExpression = "([soil_loss_ac] * [sdr]) * 907.18474"
                'END STEP 11: --------------------------------------------------------------------------------
            End If

            Dim pClipLocRusleRaster As MapWinGIS.Grid
            If g_booLocalEffects Then

                modProgDialog.ProgDialog("Creating data layer for local effects...", strTitle, 0, 13, 13, 0)
                If modProgDialog.g_boolCancel Then

                    'STEP 12: Local Effects -------------------------------------------------

                    strOutYield = modUtil.GetUniqueName("locrusle", g_strWorkspace, ".tif")
                    'Added 7/23/04 to account for clip by selected polys functionality
                    If g_booSelectedPolys Then
                        'TODO again
                        'pClipLocRusleRaster = modUtil.ClipBySelectedPoly(pSedYieldRaster, g_pSelectedPolyClip, pEnv)
                        'pPermRUSLELocRaster = modUtil.ReturnPermanentRaster(pClipLocRusleRaster, pEnv.OutWorkspace.PathName, strOutYield)
                    Else
                        'TODO: again
                        'pPermRUSLELocRaster = modUtil.ReturnPermanentRaster(pSedYieldRaster, pEnv.OutWorkspace.PathName, strOutYield)
                    End If

                    'Metadata
                    g_dicMetadata.Add("Sediment Local Effects (mg)", _strRusleMetadata)
                    'TODO
                    'pRUSLELocRasterLayer = modUtil.ReturnRasterLayer((frmPrj.m_App), pPermRUSLELocRaster, "Sediment Local Effects (mg)")
                    'pRUSLELocRasterLayer.Renderer = modUtil.ReturnRasterStretchColorRampRender(pRUSLELocRasterLayer, "Brown")
                    'pRUSLELocRasterLayer.Visible = False
                    'g_pGroupLayer.Add(pRUSLELocRasterLayer)

                    CalcRUSLE = True
                    modProgDialog.KillDialog()

                    Exit Function

                End If

            End If


            modProgDialog.ProgDialog("Calculating Accumulated Sediment...", strTitle, 0, 13, 13, 0)
            If modProgDialog.g_boolCancel Then

                'STEP 12: accum_sed = flowaccumulation([flowdir], [sedyield]) -------------------------------------------------

                'TODO: Run weighted accumulation from geoproc hydrology

                'With pMapAlgebraOp
                '    .BindRaster(g_pFlowDirRaster, "flowdir")
                '    .BindRaster(pSedYieldRaster, "sedyield")
                'End With

                'strExpression = "flowaccumulation([flowdir], [sedyield], FLOAT)"

                'pAccumSedRaster = pMapAlgebraOp.Execute(strExpression)

                'END STEP 12: --------------------------------------------------------------------------------
            End If

            modProgDialog.ProgDialog("Calculating Total Accumulated Sediment...", strTitle, 0, 13, 13, 0)
            If modProgDialog.g_boolCancel Then

                'STEP 13: accum_sed = flowaccumulation([flowdir], [sedyield]) -------------------------------------------------
                Dim totacccalc As New RasterMathCellCalc(AddressOf totaccCellCalc)
                RasterMath(pAccumSedRaster, pSedYieldRaster, Nothing, Nothing, Nothing, pTotalAccumSedRaster, totacccalc)
                'strExpression = "[accumsed] + [sedyield]"

                'END STEP 13: --------------------------------------------------------------------------------
            End If


            modProgDialog.ProgDialog("Adding accumulated sediment layer to the data group layer...", strTitle, 0, 13, 13, 0)

            Dim pClipRUSLERaster As MapWinGIS.Grid
            If modProgDialog.g_boolCancel Then
                strOutYield = modUtil.GetUniqueName("RUSLE", g_strWorkspace, ".tif")

                'Clip to selected polys if chosen
                If g_booSelectedPolys Then
                    'TODO
                    'pClipRUSLERaster = modUtil.ClipBySelectedPoly(pTotalAccumSedRaster, g_pSelectedPolyClip, pEnv)
                    'pPermAccumSedRaster = modUtil.ReturnPermanentRaster(pClipRUSLERaster, pEnv.OutWorkspace.PathName, strOutYield)
                Else
                    'pPermAccumSedRaster = modUtil.ReturnPermanentRaster(pTotalAccumSedRaster, pEnv.OutWorkspace.PathName, strOutYield)
                End If

                'Metadata
                g_dicMetadata.Add("Accumulated Sediment (kg)", _strRusleMetadata)

                'TODO
                'pRUSLERasterLayer = modUtil.ReturnRasterLayer((frmPrj.m_App), pPermAccumSedRaster, "Accumulated Sediment (kg)")
                'pRUSLERasterLayer.Renderer = modUtil.ReturnRasterStretchColorRampRender(pRUSLERasterLayer, "Brown")
                'pRUSLERasterLayer.Visible = False
                'g_pGroupLayer.Add(pRUSLERasterLayer)
            End If


            CalcRUSLE = True

            modProgDialog.KillDialog()


        Catch ex As Exception

            If Err.Number = -2147217297 Then 'User cancelled operation
                modProgDialog.g_boolCancel = False
                CalcRUSLE = False
            ElseIf Err.Number = -2147467259 Then
                MsgBox("ArcMap has reached the maximum number of GRIDs allowed in memory.  " & "Please exit N-SPECT and restart ArcMap.", MsgBoxStyle.Information, "Maximum GRID Number Encountered")
                modProgDialog.g_boolCancel = False
                modProgDialog.KillDialog()
                CalcRUSLE = False
            Else
                MsgBox("RUSLE Error: " & Err.Number & " on RUSLE Calculation: " & strExpression)
                MsgBox(Err.Number & ": " & Err.Description)
                modProgDialog.g_boolCancel = False
                modProgDialog.KillDialog()
                System.Windows.Forms.Cursor.Current = Windows.Forms.Cursors.Default
                CalcRUSLE = False
            End If
        End Try

    End Function


#Region "Raster Math"
    Private Function factCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        For i As Integer = 0 To _picks.Length - 1
            If Input1 = i + 1 Then
                Return _picks(i)
            End If
        Next
    End Function

    Private Function soillossCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        If Not _booUsingConstantValue Then 'If not using a constant
            'strExpression = "[rfactor] * [kfactor] * [lsfactor] * [cfactor]"
            Return Input1 * Input2 * Input3 * Input4
        Else 'if using a constant
            'strExpression = _dblRFactorConstant & " * [kfactor] * [lsfactor] * [cfactor]"
            Return _dblRFactorConstant * Input1 * Input2 * Input3
        End If

    End Function

    Private Function lossacCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'strExpression = "(Pow(" & g_dblCellSize & ", 2) * 0.000247104369) * [soil_loss]"
        Return (Math.Pow(g_dblCellSize, 2) * 0.000247104369) * Input1
    End Function

    Private Function pZSedCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'strExpression = "Con(([fdrnib] ge 0.5 and [fdrnib] lt 1.5), (([dem_2b] - [dem_2b](1,0)) / (" & g_dblCellSize & " * 0.001))," & "Con(([fdrnib] ge 1.5 and [fdrnib] lt 3.0), (([dem_2b] - [dem_2b](1,1)) / (" & g_dblCellSize & " * 0.0014142))," & "Con(([fdrnib] ge 3.0 and [fdrnib] lt 6.0), (([dem_2b] - [dem_2b](0,1)) / (" & g_dblCellSize & " * 0.001))," & "Con(([fdrnib] ge 6.0 and [fdrnib] lt 12.0), (([dem_2b] - [dem_2b](-1,1)) / (" & g_dblCellSize & " * 0.0014142))," & "Con(([fdrnib] ge 12.0 and [fdrnib] lt 24.0), (([dem_2b] - [dem_2b](-1,0)) / (" & g_dblCellSize & " * 0.001))," & "Con(([fdrnib] ge 24.0 and [fdrnib] lt 48.0), (([dem_2b] - [dem_2b](-1,-1)) / (" & g_dblCellSize & " * 0.0014142))," & "Con(([fdrnib] ge 48.0 and [fdrnib] lt 96.0), (([dem_2b] - [dem_2b](0,-1)) / (" & g_dblCellSize & " * 0.001))," & "Con(([fdrnib] ge 96.0 and [fdrnib] lt 192.0), (([dem_2b] - [dem_2b](1,-1)) / (" & g_dblCellSize & " * 0.0014142))," & "Con(([fdrnib] ge 192.0 and [fdrnib] le 255.0), (([dem_2b] - [dem_2b](1,0)) / (" & g_dblCellSize & " * 0.001))," & "0.1)))))))))"
        'Con(
        '  ([fdrnib] ge 0.5 and [fdrnib] lt 1.5), 
        '  (([dem_2b] - [dem_2b](1,0)) / (" & g_dblCellSize & " * 0.001)),
        '  Con(
        '    ([fdrnib] ge 1.5 and [fdrnib] lt 3.0), 
        '    (([dem_2b] - [dem_2b](1,1)) / (" & g_dblCellSize & " * 0.0014142)),
        '    Con(
        '      ([fdrnib] ge 3.0 and [fdrnib] lt 6.0), 
        '      (([dem_2b] - [dem_2b](0,1)) / (" & g_dblCellSize & " * 0.001)),
        '      Con(
        '        ([fdrnib] ge 6.0 and [fdrnib] lt 12.0), 
        '        (([dem_2b] - [dem_2b](-1,1)) / (" & g_dblCellSize & " * 0.0014142)),
        '        Con(
        '          ([fdrnib] ge 12.0 and [fdrnib] lt 24.0), 
        '          (([dem_2b] - [dem_2b](-1,0)) / (" & g_dblCellSize & " * 0.001)),
        '          Con(
        '            ([fdrnib] ge 24.0 and [fdrnib] lt 48.0), 
        '            (([dem_2b] - [dem_2b](-1,-1)) / (" & g_dblCellSize & " * 0.0014142)),
        '              Con(
        '                ([fdrnib] ge 48.0 and [fdrnib] lt 96.0), 
        '                (([dem_2b] - [dem_2b](0,-1)) / (" & g_dblCellSize & " * 0.001)),
        '                Con(
        '                  ([fdrnib] ge 96.0 and [fdrnib] lt 192.0), 
        '                  (([dem_2b] - [dem_2b](1,-1)) / (" & g_dblCellSize & " * 0.0014142)),
        '                    Con(
        '                      ([fdrnib] ge 192.0 and [fdrnib] le 255.0), 
        '                      (([dem_2b] - [dem_2b](1,0)) / (" & g_dblCellSize & " * 0.001)),
        '                      0.1
        '                    )
        '                 )
        '              )
        '           )
        '        )
        '      )
        '    )
        '  )
        ')
        If Input1 >= 0.5 And Input1 < 1.5 Then

        End If
    End Function

    Private Function areaKMCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'strExpression = "(float(con([DEM] >= 0, " & g_dblCellSize & ", 0))) / 1000"
        If Input1 >= 0 Then
            Return g_dblCellSize / 1000.0
        Else
            Return 0
        End If
    End Function

    Private Function pDACellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'strExpression = "Pow([cellarea_km], 2)"
        Return Math.Pow(Input1, 2)
    End Function

    Private Function temp3CellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'strExpression = "Pow([da_sed_del], -0.0998)"
        Return Math.Pow(Input1, -0.0998)
    End Function

    Private Function temp4CellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'strExpression = "Pow([zl_sed_del], 0.3629)"
        Return Math.Pow(Input1, 0.3629)
    End Function

    Private Function temp5CellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'strExpression = "Pow([scsgrid100], 5.444)"
        Return Math.Pow(Input1, 5.444)
    End Function

    Private Function temp6CellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'strExpression = "1.366 * (Pow(10, -11)) * [temp3] * [temp4] * [temp5]"
        Return 1.366 * Math.Pow(10, -11) * Input1 * Input2 * Input3
    End Function

    Private Function SDRCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'strExpression = "Con(([temp6] gt 1), 1, [temp6])"
        If Input1 > 1 Then
            Return 1
        Else
            Return Input1
        End If
    End Function

    Private Function sedYieldCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'strExpression = "([soil_loss_ac] * [sdr]) * 907.18474"
        Return Input1 * Input2 * 907.18474
    End Function

    Private Function totaccCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'strExpression = "[accumsed] + [sedyield]"
        Return Input1 + Input2
    End Function
    
#End Region



End Module