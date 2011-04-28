'********************************************************************************************************
'File Name: modRusle.vb
'Description: Functions handling the RUSLE portion of the model
'********************************************************************************************************
'The contents of this file are subject to the Mozilla Public License Version 1.1 (the "License"); 
'you may not use this file except in compliance with the License. You may obtain a copy of the License at 
'http://www.mozilla.org/MPL/ 
'Software distributed under the License is distributed on an "AS IS" basis, WITHOUT WARRANTY OF 
'ANY KIND, either express or implied. See the License for the specificlanguage governing rights and 
'limitations under the License. 
'
'Note: This code was converted from the vb6 NSPECT ArcGIS extension and so bears many of the old comments
'in the files where it was possible to leave them.
'
'Contributor(s): (Open source contributors should list themselves and their modifications here). 
'Oct 20, 2010:  Allen Anselmo allen.anselmo@gmail.com - 
'               Added licensing and comments to code
Imports System.IO
Imports System.Windows.Forms
Imports MapWinGeoProc
Imports MapWinGIS
Imports System.Data.OleDb
Imports OpenNspect.Xml

Module modRusleSoilLossEquation
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

    Public g_NibbleRaster As Grid
    'Nibble Raster
    Public g_DEMTwoCellRaster As Grid
    'Two Cell buffer of the DEM
    Public g_RFactorRaster As Grid
    'R Factor Raster
    Private _strRusleMetadata As String
    'Metadata holder
    Private _booUsingConstantValue As Boolean
    'Boolean value
    Private _dblRFactorConstant As Double
    'Constant for R-factor
    Private _strSDRFileName As String
    'If user provides own SDR GRid, store path here
    Private _picks As String()

    Public Function RUSLESetup(ByRef strNibbleFileName As String, ByRef strDEMTwoCellFileName As String, _
                                ByRef strRFactorFileName As String, ByRef strKfactorFileName As String, _
                                ByRef strSDRFileName As String, ByRef strLandClass As String, _
                                ByRef OutputItems As OutputItems, Optional ByRef dblRFactorConstant As Double = 0) _
        As Boolean
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
        If RasterExists(strNibbleFileName) Then
            g_NibbleRaster = ReturnRaster(strNibbleFileName)
        Else
            strError = "Nibble Raster Does Not Exist: " & strNibbleFileName
        End If
        'END STEP 1: -----------------------------------------------------------------------------------

        'STEP 2: Set the DEMTwoCell Raster -------------------------------------------------------------
        If RasterExists(strDEMTwoCellFileName) Then
            g_DEMTwoCellRaster = ReturnRaster(strDEMTwoCellFileName)
        Else
            strError = "DEM Two Cell Buffer Raster Does Not Exist: " & strDEMTwoCellFileName
        End If
        'END STEP 2: -----------------------------------------------------------------------------------

        'STEP 3: Set the R Factor Raster ---------------------------------------------------------------
        If Len(strRFactorFileName) > 0 Then
            If RasterExists(strRFactorFileName) Then
                g_RFactorRaster = ReturnRaster(strRFactorFileName)
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
        If RasterExists(strKfactorFileName) Then
            g_KFactorRaster = ReturnRaster(strKfactorFileName)
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

        strCovFactor = "SELECT LCTYPE.LCTYPEID, LCCLASS.NAME, LCCLASS.VALUE, LCCLASS.COVERFACTOR FROM " & _
                       "LCTYPE INNER JOIN LCCLASS ON LCTYPE.LCTYPEID = LCCLASS.LCTYPEID " & "WHERE LCTYPE.NAME LIKE '" & _
                       strTempLCType & "' ORDER BY LCCLASS.VALUE"
        Dim cmdCov As New DataHelper(strCovFactor)

        If Len(strError) > 0 Then
            MsgBox(strError)
            Exit Function
        End If

        'Get the con statement for the cover factor calculation
        strConStatement = ConstructPickStatment(cmdCov.GetCommand(), g_LandCoverRaster)

        'Are they using SDR
        _strSDRFileName = Trim(strSDRFileName)

        'Metadata time
        _strRusleMetadata = CreateMetadata(g_booLocalEffects)

        'Calc rusle using the con
        If CalcRUSLE(strConStatement, OutputItems) Then
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
            strHeader = vbTab & "Input Datasets:" & vbNewLine & vbTab & vbTab & "Hydrologic soils grid: " & _
                        g_XmlPrjFile.strSoilsHydFileName & vbNewLine & vbTab & vbTab & "Landcover grid: " & _
                        g_XmlPrjFile.strLCGridFileName & vbNewLine & vbTab & vbTab & "Precipitation grid: " & _
                        g_strPrecipFileName & vbNewLine & vbTab & vbTab & "Soils K Factor: " & _
                        g_XmlPrjFile.strSoilsKFileName & vbNewLine & vbTab & vbTab & "LS Factor grid: " & _
                        g_strLSFileName & vbNewLine & vbTab & vbTab & "R Factor grid: " & _
                        g_XmlPrjFile.strRainGridFileName & vbNewLine & g_strLandCoverParameters & vbNewLine
            'append the g_strLandCoverParameters that was set up during runoff
        Else
            strHeader = vbTab & "Input Datasets:" & vbNewLine & vbTab & vbTab & "Hydrologic soils grid: " & _
                        g_XmlPrjFile.strSoilsHydFileName & vbNewLine & vbTab & vbTab & "Landcover grid: " & _
                        g_XmlPrjFile.strLCGridFileName & vbNewLine & vbTab & vbTab & "Precipitation grid: " & _
                        g_strPrecipFileName & vbNewLine & vbTab & vbTab & "Soils K Factor: " & _
                        g_XmlPrjFile.strSoilsKFileName & vbNewLine & vbTab & vbTab & "LS Factor grid: " & _
                        g_strLSFileName & vbNewLine & vbTab & vbTab & "R Factor grid: " & _
                        g_XmlPrjFile.strRainGridFileName & vbNewLine & vbTab & vbTab & "Flow direction grid: " & _
                        g_strFlowDirFilename & vbNewLine & g_strLandCoverParameters & vbNewLine
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

        CreateMetadata = strHeader
        '& vbTab & "C-Factor values: " & vbNewLine & strCFactor

    End Function

    Private Function ConstructPickStatment(ByRef cmdType As OleDbCommand, ByRef pLCRaster As Grid) As String
        'Creates the initial pick statement using the name of the the LandCass [CCAP, for example]
        'and the Land Class Raster.  Returns a string
        ConstructPickStatment = ""
        Try
            Dim FieldIndex As Short
            Dim booValueFound As Boolean
            Dim i As Short
            Dim maxVal As Integer = pLCRaster.Maximum
             'TODO: it looks like some of this code is copied, refactor it.
            Dim tablepath As String = ""
            'Get the raster table
            Dim lcPath As String = pLCRaster.Filename
            If Path.GetFileName(lcPath) = "sta.adf" Then
                tablepath = Path.GetDirectoryName(lcPath) + ".dbf"
            Else
                tablepath = Path.ChangeExtension(lcPath, ".dbf")
            End If
            Dim TableExist As Boolean = File.Exists(tablepath)

            Dim strpick As String = ""

            Dim mwTable As New Table
            If Not TableExist Then
                MsgBox( _
                        "No MapWindow-readable raster table was found. To create one using ArcMap 9.3+, add the raster to the default project, right click on its layer and select Open Attribute Table. Now click on the options button in the lower right and select Export. In the export path, navigate to the directory of the grid folder and give the export the name of the raster folder with the .dbf extension. i.e. if you are exporting a raster attribute table from a raster named landcover, export landcover.dbf into the same level directory as the folder.", _
                        MsgBoxStyle.Exclamation, "Raster Attribute Table Not Found")

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
                    If (mwTable.CellValue(FieldIndex, rowidx) = i) Then _
'And (pRow.Value(FieldIndex) = rsLandClass!Value) Then
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
                            MsgBox( _
                                    "Error: Your OpenNSPECT Land Class Table is missing values found in your landcover GRID dataset.")
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
                mwTable.Close()
            End If

            'strCompleteCon = strCon & strParens
            'ConstructPickStatment = strCompleteCon
            ConstructPickStatment = strpick

        Catch ex As Exception
            MsgBox("Error in pick Statement: " & Err.Number & ": " & Err.Description)
        End Try
    End Function

    Private Function CalcRUSLE(ByRef strConStatement As String, ByRef OutputItems As OutputItems) As Boolean

        Dim pSoilLossAcres As Grid = Nothing
        'Soil Loss Acres
        Dim pZSedDelRaster As Grid = Nothing
        'And I quote, Dave's Whacky Sediment Delivery Ratio
        Dim pSDRRaster As Grid = Nothing
        'SDR
        Dim pSedYieldRaster As Grid = Nothing
        'Sediment Yield
        Dim pAccumSedRaster As Grid = Nothing
        'Accumulated Sed Raster
        Dim pTotalAccumSedRaster As Grid = Nothing
        'Total accumulated sediment raster
        Dim pPermAccumSedRaster As Grid = Nothing
        'Permanent accumulated sediment raster
        Dim pPermRUSLELocRaster As Grid = Nothing
        'RUSLE Local Effects raster

        Const strTitle As String = "Processing RUSLE Calculation..."
        Dim strOutYield As String

        Try
            ShowProgress("Solving RUSLE Equation...", strTitle, 13, 3, g_frmProjectSetup)
            If g_KeepRunning Then
                'STEP 2: SOLVE RUSLE EQUATION -------------------------------------------------------------
                ReDim _picks(strConStatement.Split(",").Length)
                _picks = strConStatement.Split(",")
                Dim AllSoilLossCalc As New RasterMathCellCalc(AddressOf AllSoilLossCellCalc)
                If Not _booUsingConstantValue Then 'If not using a constant
                    RasterMath(g_pLSRaster, g_KFactorRaster, g_LandCoverRaster, g_RFactorRaster, Nothing, _
                                pSoilLossAcres, AllSoilLossCalc)
                    'strExpression = "[rfactor] * [kfactor] * [lsfactor] * [cfactor]"
                Else 'if using a constant
                    RasterMath(g_pLSRaster, g_KFactorRaster, g_LandCoverRaster, Nothing, Nothing, pSoilLossAcres, _
                                AllSoilLossCalc)
                    'strExpression = _dblRFactorConstant & " * [kfactor] * [lsfactor] * [cfactor]"
                End If
                'END STEP 2: -------------------------------------------------------------------------------
            End If

            '***********************************************
            'BEGIN SDR CODE......
            '***********************************************
            If Len(Trim(_strSDRFileName)) = 0 Then
                ShowProgress("Calculating Relief-Length Ratio for Sediment Delivery...", strTitle, 13, 5, g_frmProjectSetup)
                If g_KeepRunning Then
                    'STEP 4: DAVE'S WACKY CALCULATION OF RELIEF-LENGTH RATIO FOR SEDIMENT DELIVERY RATIO-------
                    Dim pZSedcalc As New RasterMathCellCalcWindowNulls(AddressOf pZSedCellCalc)
                    'ARA 10/29/2010 Using base dem and flow dir instead of expanded grids
                    'RasterMathWindow(g_NibbleRaster, g_DEMTwoCellRaster, Nothing, Nothing, Nothing, pZSedDelRaster, Nothing, False, pZSedcalc)
                    RasterMathWindow(g_pFlowDirRaster, g_pDEMRaster, Nothing, Nothing, Nothing, pZSedDelRaster, Nothing, _
                                      False, pZSedcalc)
                    'strExpression = "Con(([fdrnib] ge 0.5 and [fdrnib] lt 1.5), (([dem_2b] - [dem_2b](1,0)) / (" & g_dblCellSize & " * 0.001))," & "Con(([fdrnib] ge 1.5 and [fdrnib] lt 3.0), (([dem_2b] - [dem_2b](1,1)) / (" & g_dblCellSize & " * 0.0014142))," & "Con(([fdrnib] ge 3.0 and [fdrnib] lt 6.0), (([dem_2b] - [dem_2b](0,1)) / (" & g_dblCellSize & " * 0.001))," & "Con(([fdrnib] ge 6.0 and [fdrnib] lt 12.0), (([dem_2b] - [dem_2b](-1,1)) / (" & g_dblCellSize & " * 0.0014142))," & "Con(([fdrnib] ge 12.0 and [fdrnib] lt 24.0), (([dem_2b] - [dem_2b](-1,0)) / (" & g_dblCellSize & " * 0.001))," & "Con(([fdrnib] ge 24.0 and [fdrnib] lt 48.0), (([dem_2b] - [dem_2b](-1,-1)) / (" & g_dblCellSize & " * 0.0014142))," & "Con(([fdrnib] ge 48.0 and [fdrnib] lt 96.0), (([dem_2b] - [dem_2b](0,-1)) / (" & g_dblCellSize & " * 0.001))," & "Con(([fdrnib] ge 96.0 and [fdrnib] lt 192.0), (([dem_2b] - [dem_2b](1,-1)) / (" & g_dblCellSize & " * 0.0014142))," & "Con(([fdrnib] ge 192.0 and [fdrnib] le 255.0), (([dem_2b] - [dem_2b](1,0)) / (" & g_dblCellSize & " * 0.001))," & "0.1)))))))))"

                    'END STEP 4: ------------------------------------------------------------------------------
                End If

                ShowProgress("Calculating Sediment Delivery Ratio...", strTitle, 13, 6, g_frmProjectSetup)
                If g_KeepRunning Then
                    Dim AllSDRCalc As New RasterMathCellCalc(AddressOf AllSDRCellCalc)
                    RasterMath(g_pDEMRaster, pZSedDelRaster, g_pSCS100Raster, Nothing, Nothing, pSDRRaster, AllSDRCalc)
                    pZSedDelRaster.Close()
                End If

            Else
                pSDRRaster = ReturnRaster(_strSDRFileName)
            End If
            '********************************************************************
            'END SDR CALC
            '********************************************************************

            ShowProgress("Applying Sediment Delivery Ratio...", strTitle, 13, 13, g_frmProjectSetup)
            If g_KeepRunning Then
                'STEP 11: sed_yield = [soil_loss_ac] * [sdr] -------------------------------------------------
                Dim SedYieldcalc As New RasterMathCellCalc(AddressOf sedYieldCellCalc)
                RasterMath(pSDRRaster, pSoilLossAcres, Nothing, Nothing, Nothing, pSedYieldRaster, SedYieldcalc)
                pSDRRaster.Close()
                pSoilLossAcres.Close()
                'strExpression = "([soil_loss_ac] * [sdr]) * 907.18474"
                'END STEP 11: --------------------------------------------------------------------------------
            End If

            If g_booLocalEffects Then
                ShowProgress("Creating data layer for local effects...", strTitle, 13, 13, g_frmProjectSetup)
                If g_KeepRunning Then

                    'STEP 12: Local Effects -------------------------------------------------

                    strOutYield = GetUniqueName("locrusle", g_strWorkspace, g_FinalOutputGridExt)
                    If g_booSelectedPolys Then
                        pPermRUSLELocRaster = _
                            ClipBySelectedPoly(pSedYieldRaster, g_pSelectedPolyClip, strOutYield)
                    Else
                        pPermRUSLELocRaster = ReturnPermanentRaster(pSedYieldRaster, strOutYield)
                    End If

                    'Metadata
                    g_dicMetadata.Add("Sediment Local Effects (mg)", _strRusleMetadata)

                    AddOutputGridLayer(pPermRUSLELocRaster, "Brown", True, "Sediment Local Effects (mg)", "RUSLE Local", _
                                        -1, OutputItems)

                    CalcRUSLE = True
                    CloseProgressDialog()

                    Exit Function

                End If

            End If

            ShowProgress("Calculating Accumulated Sediment...", strTitle, 13, 13, g_frmProjectSetup)
            If g_KeepRunning Then

                'STEP 12: accum_sed = flowaccumulation([flowdir], [sedyield]) -------------------------------------------------

                Dim pTauD8Flow As Grid = Nothing

                Dim tauD8calc = GetConverterToTauDemFromEsri()
                RasterMath(g_pFlowDirRaster, Nothing, Nothing, Nothing, Nothing, pTauD8Flow, Nothing, False, tauD8calc)
                pTauD8Flow.Header.NodataValue = -1

                Dim strtmp1 As String = Path.GetTempFileName
                g_TempFilesToDel.Add(strtmp1)
                strtmp1 = strtmp1 + g_TAUDEMGridExt
                g_TempFilesToDel.Add(strtmp1)
                DataManagement.DeleteGrid(strtmp1)
                pTauD8Flow.Save(strtmp1)

                Dim strtmp2 As String = Path.GetTempFileName
                g_TempFilesToDel.Add(strtmp2)
                strtmp2 = strtmp2 + g_TAUDEMGridExt
                g_TempFilesToDel.Add(strtmp2)
                DataManagement.DeleteGrid(strtmp2)
                pSedYieldRaster.Save(strtmp2)

                Dim strtmpout As String = Path.GetTempFileName
                g_TempFilesToDel.Add(strtmpout)
                strtmpout = strtmpout + "out" + g_TAUDEMGridExt
                g_TempFilesToDel.Add(strtmpout)
                DataManagement.DeleteGrid(strtmpout)

                'Use geoproc weightedAreaD8 after converting the D8 grid to taudem format bgd if needed
                Dim result = Hydrology.WeightedAreaD8(strtmp1, strtmp2, "", strtmpout, False, False, _
                                           Environment.ProcessorCount, Nothing)
                If result <> 0 Then
                    g_KeepRunning = False
                End If
                'strExpression = "flowaccumulation([flowdir], [sedyield], FLOAT)"

                pTotalAccumSedRaster = New Grid
                pTotalAccumSedRaster.Open(strtmpout)

                pTauD8Flow.Close()
                DataManagement.DeleteGrid(strtmp1)
                pSedYieldRaster.Close()
                DataManagement.DeleteGrid(strtmp2)
                'END STEP 12: --------------------------------------------------------------------------------
            End If

            ShowProgress("Adding accumulated sediment layer to the data group layer...", strTitle, 13, 13, g_frmProjectSetup)

            If g_KeepRunning Then
                strOutYield = GetUniqueName("RUSLE", g_strWorkspace, g_FinalOutputGridExt)

                'Clip to selected polys if chosen
                If g_booSelectedPolys Then
                    pPermAccumSedRaster = _
                        ClipBySelectedPoly(pTotalAccumSedRaster, g_pSelectedPolyClip, strOutYield)
                Else
                    pPermAccumSedRaster = ReturnPermanentRaster(pTotalAccumSedRaster, strOutYield)
                End If

                'Metadata
                g_dicMetadata.Add("Accumulated Sediment (kg)", _strRusleMetadata)

                AddOutputGridLayer(pPermAccumSedRaster, "Brown", True, "Accumulated Sediment (kg)", "RUSLE Accum", -1, _
                                    OutputItems)
            End If

            CalcRUSLE = True

            CloseProgressDialog()

        Catch ex As Exception

            If Err.Number = -2147217297 Then 'User cancelled operation
                g_KeepRunning = False
                CalcRUSLE = False
            ElseIf Err.Number = -2147467259 Then
                MsgBox( _
                        "ArcMap has reached the maximum number of GRIDs allowed in memory.  " & _
                        "Please exit OpenNSPECT and restart ArcMap.", MsgBoxStyle.Information, _
                        "Maximum GRID Number Encountered")
                g_KeepRunning = False
                CloseProgressDialog()
                CalcRUSLE = False
            Else
                MsgBox("RUSLE Error: " & Err.Number & " on RUSLE Calculation")
                MsgBox(Err.Number & ": " & Err.Description)
                g_KeepRunning = False
                CloseProgressDialog()
                Cursor.Current = Cursors.Default
                CalcRUSLE = False
            End If
        End Try

    End Function

#Region "Raster Math"

    Private Function AllSoilLossCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, _
                                          ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) _
        As Single
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

    Private Function pZSedCellCalc(ByRef InputBox1(,) As Single, ByVal Input1Null As Single, _
                                    ByRef InputBox2(,) As Single, ByVal Input2Null As Single, _
                                    ByRef InputBox3(,) As Single, ByVal Input3Null As Single, _
                                    ByRef InputBox4(,) As Single, ByVal Input4Null As Single, _
                                    ByRef InputBox5(,) As Single, ByVal Input5Null As Single, ByVal OutNull As Single) _
        As Single
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

    Private Function AllSDRCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, _
                                     ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        Dim kmval, daval, tmp3val, tmp4val, tmp5val, tmp6val As Single

        'strExpression = "(float(con([DEM] >= 0, " & g_dblCellSize & ", 0))) / 1000"
        If Input1 >= 0 Then
            kmval = g_dblCellSize / 1000.0
        Else
            kmval = 0
        End If

        daval = Math.Pow(kmval, 2)
        tmp3val = Math.Pow(daval, -0.0998)
        tmp4val = Math.Pow(Input2, 0.3629)
        tmp5val = Math.Pow(Input3, 5.444)
        tmp6val = 1.366 * Math.Pow(10, -11) * tmp3val * tmp4val * tmp5val

        If tmp6val > 1 Then
            Return 1
        Else
            Return tmp6val
        End If
    End Function

    Private Function sedYieldCellCalc(ByVal soilLossAC As Single, ByVal Sdr As Single, ByVal Input3 As Single, _
                                       ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) _
        As Single
        'strExpression = "([soil_loss_ac] * [sdr]) * 907.18474"
        Return soilLossAC * Sdr * 907.18474
    End Function

#End Region
End Module