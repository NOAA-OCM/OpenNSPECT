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
Imports MapWinGeoProc
Imports MapWinGIS
Imports OpenNspect.Xml

Module RevisedUniversalSoilLossEquation
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

    Private _RFactorRaster As Grid
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

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="strRFactorFileName">FileName of the R Factor GRID.</param>
    ''' <param name="strKfactorFileName">.</param>
    ''' <param name="strSDRFileName">.</param>
    ''' <param name="strLandClass">Name of the Landclass we're using.</param>
    ''' <param name="OutputItems">The output items.</param>
    ''' <param name="dblRFactorConstant">The R factor constant.</param>
    ''' <returns></returns>
    Public Function RUSLESetup(ByRef strRFactorFileName As String, ByRef strKfactorFileName As String, ByRef strSDRFileName As String, ByRef strLandClass As String, ByRef OutputItems As OutputItems, Optional ByRef dblRFactorConstant As Double = 0) As Boolean

        Dim strError As String = ""

        'STEP 3: Set the R Factor Raster ---------------------------------------------------------------
        If Len(strRFactorFileName) > 0 Then
            If RasterExists(strRFactorFileName) Then
                _RFactorRaster = ReturnRaster(strRFactorFileName)
            Else
                strError = "R Factor Raster Does Not Exist: " & strRFactorFileName
            End If
            _booUsingConstantValue = False
        Else
            _booUsingConstantValue = True
            _dblRFactorConstant = dblRFactorConstant
        End If

        'STEP 4: Set the K Factor Raster
        If RasterExists(strKfactorFileName) Then
            g_KFactorRaster = ReturnRaster(strKfactorFileName)
        Else
            strError = "K Factor Raster Does Not Exist: " & strKfactorFileName
        End If

        If Len(strError) > 0 Then
            MsgBox(strError)
            Exit Function
        End If

        'Get the landclasses of type strLandClass
        'Check first for temp name
        Dim strTempLCType As String
        If g_LandUse_DictTempNames.Count > 0 AndAlso Len(g_LandUse_DictTempNames.Item(strLandClass)) > 0 Then
            strTempLCType = g_LandUse_DictTempNames.Item(strLandClass)
        Else
            strTempLCType = strLandClass
        End If



        'Get the con statement for the cover factor calculation
        Dim strCovFactor As String = "SELECT LCTYPE.LCTYPEID, LCCLASS.NAME, LCCLASS.VALUE, LCCLASS.COVERFACTOR FROM " & "LCTYPE INNER JOIN LCCLASS ON LCTYPE.LCTYPEID = LCCLASS.LCTYPEID " & "WHERE LCTYPE.NAME LIKE '" & strTempLCType & "' ORDER BY LCCLASS.VALUE"
        Using cmdCov As New DataHelper(strCovFactor)
            Dim strConStatement As String = ConstructPickStatmentUsingLandClass(cmdCov.GetCommand(), g_LandCoverRaster)
            'Are they using SDR
            _strSDRFileName = Trim(strSDRFileName)
            'Metadata time
            _strRusleMetadata = CreateMetadata(g_Project.IncludeLocalEffects)
            'Calc rusle using the con
            Return CalcRUSLE(strConStatement, OutputItems)
        End Using

    End Function

    Private Function CreateMetadata(ByRef localEffects As Boolean) As String

        Dim strHeader As String
        'Dim i As Integer
        'Dim strCFactor As String

        'Set up the header w/or without flow direction
        If localEffects = True Then
            strHeader = vbTab & "Input Datasets:" & vbNewLine & vbTab & vbTab & "Hydrologic soils grid: " & g_Project.SoilsHydDirectory & vbNewLine & vbTab & vbTab & "Landcover grid: " & g_Project.LandCoverGridDirectory & vbNewLine & vbTab & vbTab & "Precipitation grid: " & g_strPrecipFileName & vbNewLine & vbTab & vbTab & "Soils K Factor: " & g_Project.SoilsKFileName & vbNewLine & vbTab & vbTab & "LS Factor grid: " & g_strLSFileName & vbNewLine & vbTab & vbTab & "R Factor grid: " & g_Project.StrRainGridFileName & vbNewLine & g_strLandCoverParameters & vbNewLine
            'append the g_strLandCoverParameters that was set up during runoff
        Else
            strHeader = vbTab & "Input Datasets:" & vbNewLine & vbTab & vbTab & "Hydrologic soils grid: " & g_Project.SoilsHydDirectory & vbNewLine & vbTab & vbTab & "Landcover grid: " & g_Project.LandCoverGridDirectory & vbNewLine & vbTab & vbTab & "Precipitation grid: " & g_strPrecipFileName & vbNewLine & vbTab & vbTab & "Soils K Factor: " & g_Project.SoilsKFileName & vbNewLine & vbTab & vbTab & "LS Factor grid: " & g_strLSFileName & vbNewLine & vbTab & vbTab & "R Factor grid: " & g_Project.StrRainGridFileName & vbNewLine & vbTab & vbTab & "Flow direction grid: " & g_strFlowDirFilename & vbNewLine & g_strLandCoverParameters & vbNewLine
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

    Private Function CalcRUSLE(ByRef strConStatement As String, ByRef OutputItems As OutputItems) As Boolean

        Dim pSoilLossAcres As Grid = Nothing
        'Soil Loss Acres
        Dim pZSedDelRaster As Grid = Nothing
        'And I quote, Dave's Whacky Sediment Delivery Ratio
        Dim pSDRRaster As Grid = Nothing
        'SDR
        Dim pSedYieldRaster As Grid = Nothing
        'Sediment Yield
        Dim pTotalAccumSedRaster As Grid = Nothing
        'Total accumulated sediment raster
        Dim pPermAccumSedRaster As Grid = Nothing
        'Permanent accumulated sediment raster
        Dim pPermRUSLELocRaster As Grid = Nothing
        'RUSLE Local Effects raster


        Dim strOutYield As String
        Dim progress = New SynchronousProgressDialog("Processing RUSLE Calculation...", 13, g_MainForm)
        Try
            progress.Increment("Solving RUSLE Equation...")
            If SynchronousProgressDialog.KeepRunning Then
                'STEP 2: SOLVE RUSLE EQUATION -------------------------------------------------------------
                ReDim _picks(strConStatement.Split(",").Length)
                _picks = strConStatement.Split(",")
                Dim AllSoilLossCalc As New RasterMathCellCalc(AddressOf AllSoilLossCellCalc)
                If _booUsingConstantValue Then
                    RasterMath(g_pLSRaster, g_KFactorRaster, g_LandCoverRaster, Nothing, Nothing, pSoilLossAcres, AllSoilLossCalc)
                    'strExpression = _dblRFactorConstant & " * [kfactor] * [lsfactor] * [cfactor]"
                Else
                    RasterMath(g_pLSRaster, g_KFactorRaster, g_LandCoverRaster, _RFactorRaster, Nothing, pSoilLossAcres, AllSoilLossCalc)
                    'strExpression = "[rfactor] * [kfactor] * [lsfactor] * [cfactor]"
                End If
                'END STEP 2: -------------------------------------------------------------------------------
            End If

            '***********************************************
            'BEGIN SDR CODE......
            '***********************************************
            If Len(Trim(_strSDRFileName)) = 0 Then
                progress.Increment("Calculating Relief-Length Ratio for Sediment Delivery...")
                If SynchronousProgressDialog.KeepRunning Then
                    'STEP 4: DAVE'S WACKY CALCULATION OF RELIEF-LENGTH RATIO FOR SEDIMENT DELIVERY RATIO-------
                    Dim pZSedcalc As New RasterMathCellCalcWindowNulls(AddressOf pZSedCellCalc)
                    'ARA 10/29/2010 Using base dem and flow dir instead of expanded grids
                    'RasterMathWindow(g_NibbleRaster, g_DEMTwoCellRaster, Nothing, Nothing, Nothing, pZSedDelRaster, Nothing, False, pZSedcalc)
                    RasterMathWindow(g_pFlowDirRaster, g_pDEMRaster, Nothing, Nothing, Nothing, pZSedDelRaster, Nothing, False, pZSedcalc)
                    'strExpression = "Con(([fdrnib] ge 0.5 and [fdrnib] lt 1.5), (([dem_2b] - [dem_2b](1,0)) / (" & g_dblCellSize & " * 0.001))," & "Con(([fdrnib] ge 1.5 and [fdrnib] lt 3.0), (([dem_2b] - [dem_2b](1,1)) / (" & g_dblCellSize & " * 0.0014142))," & "Con(([fdrnib] ge 3.0 and [fdrnib] lt 6.0), (([dem_2b] - [dem_2b](0,1)) / (" & g_dblCellSize & " * 0.001))," & "Con(([fdrnib] ge 6.0 and [fdrnib] lt 12.0), (([dem_2b] - [dem_2b](-1,1)) / (" & g_dblCellSize & " * 0.0014142))," & "Con(([fdrnib] ge 12.0 and [fdrnib] lt 24.0), (([dem_2b] - [dem_2b](-1,0)) / (" & g_dblCellSize & " * 0.001))," & "Con(([fdrnib] ge 24.0 and [fdrnib] lt 48.0), (([dem_2b] - [dem_2b](-1,-1)) / (" & g_dblCellSize & " * 0.0014142))," & "Con(([fdrnib] ge 48.0 and [fdrnib] lt 96.0), (([dem_2b] - [dem_2b](0,-1)) / (" & g_dblCellSize & " * 0.001))," & "Con(([fdrnib] ge 96.0 and [fdrnib] lt 192.0), (([dem_2b] - [dem_2b](1,-1)) / (" & g_dblCellSize & " * 0.0014142))," & "Con(([fdrnib] ge 192.0 and [fdrnib] le 255.0), (([dem_2b] - [dem_2b](1,0)) / (" & g_dblCellSize & " * 0.001))," & "0.1)))))))))"

                    'END STEP 4: ------------------------------------------------------------------------------
                End If

                progress.Increment("Calculating Sediment Delivery Ratio...")
                If SynchronousProgressDialog.KeepRunning Then
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

            progress.Increment("Applying Sediment Delivery Ratio...")
            If SynchronousProgressDialog.KeepRunning Then
                'STEP 11: sed_yield = [soil_loss_ac] * [sdr] -------------------------------------------------
                Dim SedYieldcalc As New RasterMathCellCalc(AddressOf sedYieldCellCalc)
                RasterMath(pSDRRaster, pSoilLossAcres, Nothing, Nothing, Nothing, pSedYieldRaster, SedYieldcalc)
                pSDRRaster.Close()
                pSoilLossAcres.Close()
                'END STEP 11: --------------------------------------------------------------------------------
            End If

            ' this was the one of two places where local effects would not calculate accumulations. For consistency we will keep running.
            If g_Project.IncludeLocalEffects Then

                If Not progress.Increment("Creating data layer for local effects...") Then Return False

                'STEP 12: Local Effects -------------------------------------------------

                strOutYield = GetUniqueFileName("locrusle", g_Project.ProjectWorkspace, OutputGridExt)
                If g_Project.UseSelectedPolygons Then
                    pPermRUSLELocRaster = ClipBySelectedPoly(pSedYieldRaster, g_pSelectedPolyClip, strOutYield)
                Else
                    pPermRUSLELocRaster = CopyRaster(pSedYieldRaster, strOutYield)
                End If

                'Metadata
                g_dicMetadata.Add("Sediment Local Effects (mg)", _strRusleMetadata)

                AddOutputGridLayer(pPermRUSLELocRaster, "Brown", True, "Sediment Local Effects (mg)", "RUSLE Local", -1, OutputItems)

            Else
                pSedYieldRaster.Save()
            End If

            progress.Increment("Calculating Accumulated Sediment...")
            If SynchronousProgressDialog.KeepRunning Then

                Dim pTauD8Flow As Grid = Nothing

                Dim tauD8calc = GetConverterToTauDemFromEsri()
                RasterMath(g_pFlowDirRaster, Nothing, Nothing, Nothing, Nothing, pTauD8Flow, Nothing, False, tauD8calc)
                pTauD8Flow.Header.NodataValue = -1

                Dim flowdir As String = GetTempFileNameTauDemGridExt()
                pTauD8Flow.Save()  'Saving because it seemed necessary.
                pTauD8Flow.Save(flowdir)

                Dim sedyield As String = GetTempFileNameTauDemGridExt()
                pSedYieldRaster.Save()  'Saving because it seemed necessary.
                pSedYieldRaster.Save(sedyield)

                Dim strtmpout As String = GetTempFileNameTauDemGridExt()

                Dim result = Hydrology.WeightedAreaD8(flowdir, sedyield, Nothing, strtmpout, False, False, Environment.ProcessorCount, False, Nothing)
                If result <> 0 Then
                    SynchronousProgressDialog.KeepRunning = False
                End If

                pTotalAccumSedRaster = New Grid
                pTotalAccumSedRaster.Open(strtmpout)

                pTauD8Flow.Close()
                pSedYieldRaster.Close()
            End If

            progress.Increment("Adding accumulated sediment layer to the data group layer...")

            If SynchronousProgressDialog.KeepRunning Then
                strOutYield = GetUniqueFileName("RUSLE", g_Project.ProjectWorkspace, OutputGridExt)

                'Clip to selected polys if chosen
                If g_Project.UseSelectedPolygons Then
                    pPermAccumSedRaster = ClipBySelectedPoly(pTotalAccumSedRaster, g_pSelectedPolyClip, strOutYield)
                Else
                    pPermAccumSedRaster = CopyRaster(pTotalAccumSedRaster, strOutYield)
                End If

                'Metadata
                g_dicMetadata.Add("Accumulated Sediment (kg)", _strRusleMetadata)

                AddOutputGridLayer(pPermAccumSedRaster, "Brown", True, "Accumulated Sediment (kg)", "RUSLE Accum", -1, OutputItems)
            End If

            CalcRUSLE = True

        Catch ex As Exception
            HandleError(ex)
            Return False
        Finally
            progress.Dispose()
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
            Return (Math.Pow(g_CellSize, 2) * 0.000247104369) * Input1 * Input2 * tmpval * Input4
        Else 'if using a constant
            'strExpression = _dblRFactorConstant & " * [kfactor] * [lsfactor] * [cfactor]"
            Return (Math.Pow(g_CellSize, 2) * 0.000247104369) * _dblRFactorConstant * Input1 * Input2 * tmpval
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
                        Return (InputBox2(1, 1) - InputBox2(1, 2)) / (g_CellSize * 0.001)
                    Else
                        Return OutNull
                    End If
                ElseIf InputBox1(1, 1) >= 1.5 And InputBox1(1, 1) < 3.0 Then
                    '  Con(
                    '    ([fdrnib] ge 1.5 and [fdrnib] lt 3.0), 
                    '    (([dem_2b] - [dem_2b](1,1)) / (" & g_dblCellSize & " * 0.0014142)),
                    If InputBox2(2, 2) <> Input2Null Then
                        Return (InputBox2(1, 1) - InputBox2(2, 2)) / (g_CellSize * 0.0014142)
                    Else
                        Return OutNull
                    End If
                ElseIf InputBox1(1, 1) >= 3.0 And InputBox1(1, 1) < 6.0 Then
                    '    Con(
                    '      ([fdrnib] ge 3.0 and [fdrnib] lt 6.0), 
                    '      (([dem_2b] - [dem_2b](0,1)) / (" & g_dblCellSize & " * 0.001)),
                    If InputBox2(2, 1) <> Input2Null Then
                        Return (InputBox2(1, 1) - InputBox2(2, 1)) / (g_CellSize * 0.001)
                    Else
                        Return OutNull
                    End If
                ElseIf InputBox1(1, 1) >= 6.0 And InputBox1(1, 1) < 12.0 Then
                    '      Con(
                    '        ([fdrnib] ge 6.0 and [fdrnib] lt 12.0), 
                    '        (([dem_2b] - [dem_2b](-1,1)) / (" & g_dblCellSize & " * 0.0014142)),
                    If InputBox2(2, 0) <> Input2Null Then
                        Return (InputBox2(1, 1) - InputBox2(2, 0)) / (g_CellSize * 0.0014142)
                    Else
                        Return OutNull
                    End If
                ElseIf InputBox1(1, 1) >= 12.0 And InputBox1(1, 1) < 24.0 Then
                    '        Con(
                    '          ([fdrnib] ge 12.0 and [fdrnib] lt 24.0), 
                    '          (([dem_2b] - [dem_2b](-1,0)) / (" & g_dblCellSize & " * 0.001)),
                    If InputBox2(1, 0) <> Input2Null Then
                        Return (InputBox2(1, 1) - InputBox2(1, 0)) / (g_CellSize * 0.001)
                    Else
                        Return OutNull
                    End If
                ElseIf InputBox1(1, 1) >= 24.0 And InputBox1(1, 1) < 48.0 Then
                    '          Con(
                    '            ([fdrnib] ge 24.0 and [fdrnib] lt 48.0), 
                    '            (([dem_2b] - [dem_2b](-1,-1)) / (" & g_dblCellSize & " * 0.0014142)),
                    If InputBox2(0, 0) <> Input2Null Then
                        Return (InputBox2(1, 1) - InputBox2(0, 0)) / (g_CellSize * 0.0014142)
                    Else
                        Return OutNull
                    End If
                ElseIf InputBox1(1, 1) >= 48.0 And InputBox1(1, 1) < 96.0 Then
                    '              Con(
                    '                ([fdrnib] ge 48.0 and [fdrnib] lt 96.0), 
                    '                (([dem_2b] - [dem_2b](0,-1)) / (" & g_dblCellSize & " * 0.001)),
                    If InputBox2(0, 1) <> Input2Null Then
                        Return (InputBox2(1, 1) - InputBox2(0, 1)) / (g_CellSize * 0.001)
                    Else
                        Return OutNull
                    End If
                ElseIf InputBox1(1, 1) >= 96.0 And InputBox1(1, 1) < 192.0 Then
                    '                Con(
                    '                  ([fdrnib] ge 96.0 and [fdrnib] lt 192.0), 
                    '                  (([dem_2b] - [dem_2b](1,-1)) / (" & g_dblCellSize & " * 0.0014142)),
                    If InputBox2(0, 2) <> Input2Null Then
                        Return (InputBox2(1, 1) - InputBox2(0, 2)) / (g_CellSize * 0.0014142)
                    Else
                        Return OutNull
                    End If
                ElseIf InputBox1(1, 1) >= 192.0 And InputBox1(1, 1) < 255.0 Then
                    '                    Con(
                    '                      ([fdrnib] ge 192.0 and [fdrnib] le 255.0), 
                    '                      (([dem_2b] - [dem_2b](1,0)) / (" & g_dblCellSize & " * 0.001)),
                    If InputBox2(1, 2) <> Input2Null Then
                        Return (InputBox2(1, 1) - InputBox2(1, 2)) / (g_CellSize * 0.001)
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

    Private Function AllSDRCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        Dim kmval, daval, tmp3val, tmp4val, tmp5val, tmp6val As Single

        'strExpression = "(float(con([DEM] >= 0, " & g_dblCellSize & ", 0))) / 1000"
        If Input1 >= 0 Then
            kmval = g_CellSize / 1000.0
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

    Private Function sedYieldCellCalc(ByVal soilLossAC As Single, ByVal Sdr As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        'strExpression = "([soil_loss_ac] * [sdr]) * 907.18474"
        Return soilLossAC * Sdr * 907.18474
    End Function

#End Region
End Module