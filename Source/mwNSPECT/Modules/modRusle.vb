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
        If g_DictTempNames.Count > 0 AndAlso Len(g_DictTempNames.Item(strLandClass)) > 0 Then
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
                For i = 1 To maxVal
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
                            strpick = "0"
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

    Private Function CalcRUSLE(ByRef strConStatement As String) As Boolean

        Dim pSoilLossAcres As MapWinGIS.Grid = Nothing  'Soil Loss Acres
        Dim pZSedDelRaster As MapWinGIS.Grid = Nothing  'And I quote, Dave's Whacky Sediment Delivery Ratio
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
            modProgDialog.ProgDialog("Solving RUSLE Equation...", strTitle, 0, 13, 3, 0)
            If modProgDialog.g_boolCancel Then
                'STEP 2: SOLVE RUSLE EQUATION -------------------------------------------------------------
                ReDim _picks(strConStatement.Split(",").Length)
                _picks = strConStatement.Split(",")
                Dim AllSoilLossCalc As New RasterMathCellCalc(AddressOf AllSoilLossCellCalc)
                If Not _booUsingConstantValue Then 'If not using a constant
                    RasterMath(g_pLSRaster, g_KFactorRaster, g_LandCoverRaster, g_RFactorRaster, Nothing, pSoilLossAcres, AllSoilLossCalc)
                    'strExpression = "[rfactor] * [kfactor] * [lsfactor] * [cfactor]"
                Else 'if using a constant
                    RasterMath(g_pLSRaster, g_KFactorRaster, g_LandCoverRaster, Nothing, Nothing, pSoilLossAcres, AllSoilLossCalc)
                    'strExpression = _dblRFactorConstant & " * [kfactor] * [lsfactor] * [cfactor]"
                End If
                'END STEP 2: -------------------------------------------------------------------------------
            End If


            '***********************************************
            'BEGIN SDR CODE......
            '***********************************************
            If Len(Trim(_strSDRFileName)) = 0 Then
                modProgDialog.ProgDialog("Calculating Relief-Length Ratio for Sediment Delivery...", strTitle, 0, 13, 5, 0)
                If modProgDialog.g_boolCancel Then
                    'STEP 4: DAVE'S WACKY CALCULATION OF RELIEF-LENGTH RATIO FOR SEDIMENT DELIVERY RATIO-------
                    Dim pZSedcalc As New RasterMathCellCalcWindowNulls(AddressOf pZSedCellCalc)
                    RasterMathWindow(g_NibbleRaster, g_DEMTwoCellRaster, Nothing, Nothing, Nothing, pZSedDelRaster, Nothing, False, pZSedcalc)
                    'strExpression = "Con(([fdrnib] ge 0.5 and [fdrnib] lt 1.5), (([dem_2b] - [dem_2b](1,0)) / (" & g_dblCellSize & " * 0.001))," & "Con(([fdrnib] ge 1.5 and [fdrnib] lt 3.0), (([dem_2b] - [dem_2b](1,1)) / (" & g_dblCellSize & " * 0.0014142))," & "Con(([fdrnib] ge 3.0 and [fdrnib] lt 6.0), (([dem_2b] - [dem_2b](0,1)) / (" & g_dblCellSize & " * 0.001))," & "Con(([fdrnib] ge 6.0 and [fdrnib] lt 12.0), (([dem_2b] - [dem_2b](-1,1)) / (" & g_dblCellSize & " * 0.0014142))," & "Con(([fdrnib] ge 12.0 and [fdrnib] lt 24.0), (([dem_2b] - [dem_2b](-1,0)) / (" & g_dblCellSize & " * 0.001))," & "Con(([fdrnib] ge 24.0 and [fdrnib] lt 48.0), (([dem_2b] - [dem_2b](-1,-1)) / (" & g_dblCellSize & " * 0.0014142))," & "Con(([fdrnib] ge 48.0 and [fdrnib] lt 96.0), (([dem_2b] - [dem_2b](0,-1)) / (" & g_dblCellSize & " * 0.001))," & "Con(([fdrnib] ge 96.0 and [fdrnib] lt 192.0), (([dem_2b] - [dem_2b](1,-1)) / (" & g_dblCellSize & " * 0.0014142))," & "Con(([fdrnib] ge 192.0 and [fdrnib] le 255.0), (([dem_2b] - [dem_2b](1,0)) / (" & g_dblCellSize & " * 0.001))," & "0.1)))))))))"

                    'END STEP 4: ------------------------------------------------------------------------------
                End If

                modProgDialog.ProgDialog("Calculating Sediment Delivery Ratio...", strTitle, 0, 13, 6, 0)
                If modProgDialog.g_boolCancel Then
                    Dim AllSDRCalc As New RasterMathCellCalc(AddressOf AllSDRCellCalc)
                    RasterMath(g_pDEMRaster, pZSedDelRaster, g_pSCS100Raster, Nothing, Nothing, pSDRRaster, AllSDRCalc)
                    pZSedDelRaster.Close()
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
                pSDRRaster.Close()
                pSoilLossAcres.Close()
                'strExpression = "([soil_loss_ac] * [sdr]) * 907.18474"
                'END STEP 11: --------------------------------------------------------------------------------
            End If

            If g_booLocalEffects Then
                modProgDialog.ProgDialog("Creating data layer for local effects...", strTitle, 0, 13, 13, 0)
                If modProgDialog.g_boolCancel Then

                    'STEP 12: Local Effects -------------------------------------------------

                    strOutYield = modUtil.GetUniqueName("locrusle", g_strWorkspace, ".bgd")
                    If g_booSelectedPolys Then
                        pPermRUSLELocRaster = modUtil.ClipBySelectedPoly(pSedYieldRaster, g_pSelectedPolyClip, strOutYield)
                    Else
                        pPermRUSLELocRaster = modUtil.ReturnPermanentRaster(pSedYieldRaster, strOutYield)
                    End If

                    'Metadata
                    g_dicMetadata.Add("Sediment Local Effects (mg)", _strRusleMetadata)

                    Dim cs As MapWinGIS.GridColorScheme = ReturnRasterStretchColorRampCS(pPermRUSLELocRaster, "Brown")
                    Dim lyr As MapWindow.Interfaces.Layer = g_MapWin.Layers.Add(pPermRUSLELocRaster, cs, "Sediment Local Effects (mg)")
                    lyr.Visible = False
                    lyr.MoveTo(0, g_pGroupLayer)


                    CalcRUSLE = True
                    modProgDialog.KillDialog()

                    Exit Function

                End If

            End If


            modProgDialog.ProgDialog("Calculating Accumulated Sediment...", strTitle, 0, 13, 13, 0)
            If modProgDialog.g_boolCancel Then

                'STEP 12: accum_sed = flowaccumulation([flowdir], [sedyield]) -------------------------------------------------

                Dim pTauD8Flow As MapWinGIS.Grid = Nothing

                Dim tauD8calc As New RasterMathCellCalcNulls(AddressOf tauD8CellCalc)
                RasterMath(g_pFlowDirRaster, Nothing, Nothing, Nothing, Nothing, pTauD8Flow, Nothing, False, tauD8calc)
                pTauD8Flow.Header.NodataValue = -1

                Dim strtmp1 As String = IO.Path.GetTempFileName + ".bgd"
                MapWinGeoProc.DataManagement.DeleteGrid(strtmp1)
                pTauD8Flow.Save(strtmp1)

                Dim strtmp2 As String = IO.Path.GetTempFileName + ".bgd"
                MapWinGeoProc.DataManagement.DeleteGrid(strtmp2)
                pSedYieldRaster.Save(strtmp2)

                Dim strtmpout As String = IO.Path.GetTempFileName + "out.bgd"
                MapWinGeoProc.DataManagement.DeleteGrid(strtmpout)


                'Use geoproc weightedAreaD8 after converting the D8 grid to taudem format bgd if needed
                MapWinGeoProc.Hydrology.WeightedAreaD8(strtmp1, strtmp2, "", strtmpout, False, False, Nothing)
                'strExpression = "flowaccumulation([flowdir], [sedyield], FLOAT)"

                pTotalAccumSedRaster = New MapWinGIS.Grid
                pTotalAccumSedRaster.Open(strtmpout)

                pTauD8Flow.Close()
                MapWinGeoProc.DataManagement.DeleteGrid(strtmp1)
                pSedYieldRaster.Close()
                MapWinGeoProc.DataManagement.DeleteGrid(strtmp2)
                'END STEP 12: --------------------------------------------------------------------------------
            End If


            modProgDialog.ProgDialog("Adding accumulated sediment layer to the data group layer...", strTitle, 0, 13, 13, 0)

            If modProgDialog.g_boolCancel Then
                strOutYield = modUtil.GetUniqueName("RUSLE", g_strWorkspace, ".bgd")

                'Clip to selected polys if chosen
                If g_booSelectedPolys Then
                    pPermAccumSedRaster = modUtil.ClipBySelectedPoly(pTotalAccumSedRaster, g_pSelectedPolyClip, strOutYield)
                Else
                    pPermAccumSedRaster = modUtil.ReturnPermanentRaster(pTotalAccumSedRaster, strOutYield)
                End If

                'Metadata
                g_dicMetadata.Add("Accumulated Sediment (kg)", _strRusleMetadata)

                Dim cs As MapWinGIS.GridColorScheme = ReturnRasterStretchColorRampCS(pPermAccumSedRaster, "Brown")
                Dim lyr As MapWindow.Interfaces.Layer = g_MapWin.Layers.Add(pPermAccumSedRaster, cs, "Accumulated Sediment (kg)")
                lyr.Visible = False
                lyr.MoveTo(0, g_pGroupLayer)
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
    Private Function AllSoilLossCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        Dim tmpval As Single
        For i As Integer = 0 To _picks.Length - 1
            If Input3 = i + 1 Then
                tmpval = _picks(i)
                Exit For
            End If
        Next

        If Not _booUsingConstantValue Then 'If not using a constant
            'strExpression = "[rfactor] * [kfactor] * [lsfactor] * [cfactor]"
            Return (Math.Pow(g_dblCellSize, 2) * 0.000247104369) * Input1 * Input2 * tmpval * Input4
        Else 'if using a constant
            'strExpression = _dblRFactorConstant & " * [kfactor] * [lsfactor] * [cfactor]"
            Return (Math.Pow(g_dblCellSize, 2) * 0.000247104369) * _dblRFactorConstant * Input1 * Input2 * tmpval
        End If

    End Function

    Private Function pZSedCellCalc(ByRef InputBox1(,) As Single, ByVal Input1Null As Single, ByRef InputBox2(,) As Single, ByVal Input2Null As Single, ByRef InputBox3(,) As Single, ByVal Input3Null As Single, ByRef InputBox4(,) As Single, ByVal Input4Null As Single, ByRef InputBox5(,) As Single, ByVal Input5Null As Single, ByVal OutNull As Single) As Single
        'strExpression = "Con(([fdrnib] ge 0.5 and [fdrnib] lt 1.5), (([dem_2b] - [dem_2b](1,0)) / (" & g_dblCellSize & " * 0.001))," & "Con(([fdrnib] ge 1.5 and [fdrnib] lt 3.0), (([dem_2b] - [dem_2b](1,1)) / (" & g_dblCellSize & " * 0.0014142))," & "Con(([fdrnib] ge 3.0 and [fdrnib] lt 6.0), (([dem_2b] - [dem_2b](0,1)) / (" & g_dblCellSize & " * 0.001))," & "Con(([fdrnib] ge 6.0 and [fdrnib] lt 12.0), (([dem_2b] - [dem_2b](-1,1)) / (" & g_dblCellSize & " * 0.0014142))," & "Con(([fdrnib] ge 12.0 and [fdrnib] lt 24.0), (([dem_2b] - [dem_2b](-1,0)) / (" & g_dblCellSize & " * 0.001))," & "Con(([fdrnib] ge 24.0 and [fdrnib] lt 48.0), (([dem_2b] - [dem_2b](-1,-1)) / (" & g_dblCellSize & " * 0.0014142))," & "Con(([fdrnib] ge 48.0 and [fdrnib] lt 96.0), (([dem_2b] - [dem_2b](0,-1)) / (" & g_dblCellSize & " * 0.001))," & "Con(([fdrnib] ge 96.0 and [fdrnib] lt 192.0), (([dem_2b] - [dem_2b](1,-1)) / (" & g_dblCellSize & " * 0.0014142))," & "Con(([fdrnib] ge 192.0 and [fdrnib] le 255.0), (([dem_2b] - [dem_2b](1,0)) / (" & g_dblCellSize & " * 0.001))," & "0.1)))))))))"

        If InputBox1(1, 1) <> Input1Null Then
            If InputBox2(1, 1) <> Input2Null Then
                If InputBox1(1, 1) >= 0.5 And InputBox1(1, 1) < 1.5 Then
                    'Con(
                    '  ([fdrnib] ge 0.5 and [fdrnib] lt 1.5), 
                    '  (([dem_2b] - [dem_2b](1,0)) / (" & g_dblCellSize & " * 0.001)),
                    If InputBox2(1, 2) <> Input2Null Then
                        Return (InputBox2(1, 1) - InputBox2(1, 2)) / (g_dblCellSize * 0.001)
                    Else
                        Return OutNull
                    End If
                ElseIf InputBox1(1, 1) >= 1.5 And InputBox1(1, 1) < 3.0 Then
                    '  Con(
                    '    ([fdrnib] ge 1.5 and [fdrnib] lt 3.0), 
                    '    (([dem_2b] - [dem_2b](1,1)) / (" & g_dblCellSize & " * 0.0014142)),
                    If InputBox2(2, 2) <> Input2Null Then
                        Return (InputBox2(1, 1) - InputBox2(2, 2)) / (g_dblCellSize * 0.0014142)
                    Else
                        Return OutNull
                    End If
                ElseIf InputBox1(1, 1) >= 3.0 And InputBox1(1, 1) < 6.0 Then
                    '    Con(
                    '      ([fdrnib] ge 3.0 and [fdrnib] lt 6.0), 
                    '      (([dem_2b] - [dem_2b](0,1)) / (" & g_dblCellSize & " * 0.001)),
                    If InputBox2(2, 1) <> Input2Null Then
                        Return (InputBox2(1, 1) - InputBox2(2, 1)) / (g_dblCellSize * 0.001)
                    Else
                        Return OutNull
                    End If
                ElseIf InputBox1(1, 1) >= 6.0 And InputBox1(1, 1) < 12.0 Then
                    '      Con(
                    '        ([fdrnib] ge 6.0 and [fdrnib] lt 12.0), 
                    '        (([dem_2b] - [dem_2b](-1,1)) / (" & g_dblCellSize & " * 0.0014142)),
                    If InputBox2(2, 0) <> Input2Null Then
                        Return (InputBox2(1, 1) - InputBox2(2, 0)) / (g_dblCellSize * 0.0014142)
                    Else
                        Return OutNull
                    End If
                ElseIf InputBox1(1, 1) >= 12.0 And InputBox1(1, 1) < 24.0 Then
                    '        Con(
                    '          ([fdrnib] ge 12.0 and [fdrnib] lt 24.0), 
                    '          (([dem_2b] - [dem_2b](-1,0)) / (" & g_dblCellSize & " * 0.001)),
                    If InputBox2(1, 0) <> Input2Null Then
                        Return (InputBox2(1, 1) - InputBox2(1, 0)) / (g_dblCellSize * 0.001)
                    Else
                        Return OutNull
                    End If
                ElseIf InputBox1(1, 1) >= 24.0 And InputBox1(1, 1) < 48.0 Then
                    '          Con(
                    '            ([fdrnib] ge 24.0 and [fdrnib] lt 48.0), 
                    '            (([dem_2b] - [dem_2b](-1,-1)) / (" & g_dblCellSize & " * 0.0014142)),
                    If InputBox2(0, 0) <> Input2Null Then
                        Return (InputBox2(1, 1) - InputBox2(0, 0)) / (g_dblCellSize * 0.0014142)
                    Else
                        Return OutNull
                    End If
                ElseIf InputBox1(1, 1) >= 48.0 And InputBox1(1, 1) < 96.0 Then
                    '              Con(
                    '                ([fdrnib] ge 48.0 and [fdrnib] lt 96.0), 
                    '                (([dem_2b] - [dem_2b](0,-1)) / (" & g_dblCellSize & " * 0.001)),
                    If InputBox2(0, 1) <> Input2Null Then
                        Return (InputBox2(1, 1) - InputBox2(0, 1)) / (g_dblCellSize * 0.001)
                    Else
                        Return OutNull
                    End If
                ElseIf InputBox1(1, 1) >= 96.0 And InputBox1(1, 1) < 192.0 Then
                    '                Con(
                    '                  ([fdrnib] ge 96.0 and [fdrnib] lt 192.0), 
                    '                  (([dem_2b] - [dem_2b](1,-1)) / (" & g_dblCellSize & " * 0.0014142)),
                    If InputBox2(0, 2) <> Input2Null Then
                        Return (InputBox2(1, 1) - InputBox2(0, 2)) / (g_dblCellSize * 0.0014142)
                    Else
                        Return OutNull
                    End If
                ElseIf InputBox1(1, 1) >= 192.0 And InputBox1(1, 1) < 255.0 Then
                    '                    Con(
                    '                      ([fdrnib] ge 192.0 and [fdrnib] le 255.0), 
                    '                      (([dem_2b] - [dem_2b](1,0)) / (" & g_dblCellSize & " * 0.001)),
                    If InputBox2(1, 2) <> Input2Null Then
                        Return (InputBox2(1, 1) - InputBox2(1, 2)) / (g_dblCellSize * 0.001)
                    Else
                        Return OutNull
                    End If
                Else
                    Return 0.1
                End If
            Else
                Return OutNull
            End If
        Else
            Return OutNull
        End If


    End Function

    Private Function SDRCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        'strExpression = "Con(([temp6] gt 1), 1, [temp6])"
        If Input1 > 1 Then
            Return 1
        Else
            Return Input1
        End If
    End Function

    Private Function AllSDRCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        Dim kmval, daval, tmp3val, tmp4val, tmp5val, tmp6val As Single

        'strExpression = "(float(con([DEM] >= 0, " & g_dblCellSize & ", 0))) / 1000"
        If Input1 >= 0 Then
            kmval = g_dblCellSize / 1000.0
        Else
            kmval = 0
        End If

        'strExpression = "Pow([cellarea_km], 2)"
        daval = Math.Pow(kmval, 2)

        'strExpression = "Pow([da_sed_del], -0.0998)"
        tmp3val = Math.Pow(daval, -0.0998)

        'strExpression = "Pow([zl_sed_del], 0.3629)"
        tmp4val = Math.Pow(Input2, 0.3629)

        'strExpression = "Pow([scsgrid100], 5.444)"
        tmp5val = Math.Pow(Input3, 5.444)

        'strExpression = "1.366 * (Pow(10, -11)) * [temp3] * [temp4] * [temp5]"
        tmp6val = 1.366 * Math.Pow(10, -11) * tmp3val * tmp4val * tmp5val

        'strExpression = "Con(([temp6] gt 1), 1, [temp6])"
        If tmp6val > 1 Then
            Return 1
        Else
            Return tmp6val
        End If
    End Function

    Private Function sedYieldCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        'strExpression = "([soil_loss_ac] * [sdr]) * 907.18474"
        Return Input1 * Input2 * 907.18474
    End Function

    
#End Region



End Module