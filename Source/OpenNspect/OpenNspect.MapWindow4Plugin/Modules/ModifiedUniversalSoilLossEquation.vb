'********************************************************************************************************
'File Name: modMUSLE.vb
'Description: Functions handling the MUSLE portion of the model
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
Imports System.Data.OleDb
Imports System.IO
Imports MapWinGeoProc
Imports MapWinGIS
Imports OpenNspect.Xml

Module ModifiedUniversalSoilLossEquation
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
    Private _strCFactorConStatement As String
    'C Factor Con Statement
    Private _strPondConStatement As String
    'Pond Factor con statement
    Private _dblMUSLEVal As Double
    'Customizable MUSLE value in Equation
    Private _dblMUSLEExp As Double
    'Customizable musle exponent in equation

    Private _strMusleMetadata As String
    'MUSLE metadata string

    ' Variables used by the Error handler function - DO NOT REMOVE
    Private _picks As String()
    Private _pondpicks As String()

    Public Function MUSLESetup(ByRef strSoilsDefName As String, ByRef strKfactorFileName As String, ByRef strLandClass As String, ByRef OutputItems As OutputItems) As Boolean
        'Sub takes incoming parameters from the project file and then parses them out
        'strSoilsDefName: Name of the Soils Definition being used
        'strKFactorFileName: K Factor FileName
        'strLandClass: Name of the Landclass we're using

        Dim strTempLCType As String
        'Our potential holder for a temp landtype

        'Open Strings
        Dim strError As String = ""

        'STEP 1: Get the MUSLE Values
        Dim cmdsoils As New DataHelper("SELECT * FROM SOILS WHERE NAME LIKE '" & strSoilsDefName & "'")
        Dim datasoils As OleDbDataReader = cmdsoils.ExecuteReader
        datasoils.Read()

        _dblMUSLEVal = datasoils("MUSLEVal")
        _dblMUSLEExp = datasoils("MUSLEExp")
        datasoils.Close()

        'STEP 2: Set the K Factor Raster
        If RasterExists(strKfactorFileName) Then
            g_KFactorRaster = ReturnRaster(strKfactorFileName)
        Else
            strError = "K Factor Raster Does Not Exist: " & strKfactorFileName
        End If

        '
        If g_LandUse_DictTempNames.Count > 0 AndAlso Len(g_LandUse_DictTempNames.Item(strLandClass)) > 0 Then
            strTempLCType = g_LandUse_DictTempNames.Item(strLandClass)
        Else
            strTempLCType = strLandClass
        End If

        'Get the landclasses of type strLandClass
        Dim cmdCovfact As New DataHelper("SELECT LCTYPE.LCTYPEID, LCCLASS.NAME, LCCLASS.VALUE, LCCLASS.COVERFACTOR, LCCLASS.W_WL FROM " & "LCTYPE INNER JOIN LCCLASS ON LCTYPE.LCTYPEID = LCCLASS.LCTYPEID " & "WHERE LCTYPE.NAME LIKE '" & strTempLCType & "' ORDER BY LCCLASS.VALUE")

        _strMusleMetadata = CreateMetadata(g_Project.IncludeLocalEffects)

        If Len(strError) > 0 Then
            MsgBox(strError)
            Exit Function
        End If

        'Get the con statement for the cover factor calculation
        Dim command As OleDbCommand = cmdCovfact.GetCommand()
        _strCFactorConStatement = ConstructPickStatmentUsingLandClass(command, g_LandCoverRaster)
        _strPondConStatement = ConstructPondPickStatement(command, g_LandCoverRaster)

        'Calc rusle using the con
        If CalcMUSLE(_strCFactorConStatement, _strPondConStatement, OutputItems) Then
            MUSLESetup = True
        Else
            MUSLESetup = False
        End If

    End Function

    Private Function CreateMetadata(ByRef includeLocalEffects As Boolean) As String

        Dim strHeader As String
        'Dim i As Integer
        'Dim strCFactor As String

        'Set up the header w/or without flow direction
        If includeLocalEffects = True Then
            strHeader = vbTab & "Input Datasets:" & vbNewLine & vbTab & vbTab & "Hydrologic soils grid: " & g_Project.SoilsHydDirectory & vbNewLine & vbTab & vbTab & "Landcover grid: " & g_Project.LandCoverGridDirectory & vbNewLine & vbTab & vbTab & "Precipitation grid: " & g_strPrecipFileName & vbNewLine & vbTab & vbTab & "Soils K Factor: " & g_Project.SoilsKFileName & vbNewLine & vbTab & vbTab & "LS Factor grid: " & g_strLSFileName & vbNewLine & g_strLandCoverParameters & vbNewLine
            'append the g_strLandCoverParameters that was set up during runoff
        Else
            strHeader = vbTab & "Input Datasets:" & vbNewLine & vbTab & vbTab & "Hydrologic soils grid: " & g_Project.SoilsHydDirectory & vbNewLine & vbTab & vbTab & "Landcover grid: " & g_Project.LandCoverGridDirectory & vbNewLine & vbTab & vbTab & "Precipitation grid: " & g_strPrecipFileName & vbNewLine & vbTab & vbTab & "Soils K Factor: " & g_Project.SoilsKFileName & vbNewLine & vbTab & vbTab & "LS Factor grid: " & g_strLSFileName & vbNewLine & vbTab & vbTab & "Flow direction grid: " & g_strFlowDirFilename & vbNewLine & g_strLandCoverParameters & vbNewLine
        End If

        CreateMetadata = strHeader
    End Function

    Private Function ConstructPondPickStatement(ByRef cmdCF As OleDbCommand, ByRef pLCRaster As Grid) As String
        'Creates the Con Statement used in the Pond Factor GRID
        'Returns: String
        'Looks like: con(([nu_lulc] eq 16), 0, con((nu_lulc eq 17), 0...
        ConstructPondPickStatement = ""

        Dim FieldIndex As Short
        Dim booValueFound As Boolean
        Dim i As Short

        Dim maxVal As Integer = pLCRaster.Maximum
        Dim minVal As Integer = pLCRaster.Minimum
        Dim nodata As Single = pLCRaster.Header.NodataValue
        'TODO: it looks like some of this code is copied, refactor it.
        Dim tablepath = GetRasterTablePath(pLCRaster)
        Dim TableExist As Boolean = File.Exists(tablepath)
        Dim strpick As String = ""

        Dim mwTable As New Table
        If Not TableExist Then
            MsgBox("No MapWindow-readable raster table was found. To create one using ArcMap 9.3+, add the raster to the default project, right click on its layer and select Open Attribute Table. Now click on the options button in the lower right and select Export. In the export path, navigate to the directory of the grid folder and give the export the name of the raster folder with the .dbf extension. i.e. if you are exporting a raster attribute table from a raster named landcover, export landcover.dbf into the same level directory as the folder.", MsgBoxStyle.Exclamation, "Raster Attribute Table Not Found")

            Return ""
        Else
            mwTable.Open(tablepath)

            FieldIndex = GetFieldIndex(mwTable)

            Dim rowidx As Integer = 0
            Dim dataCF As OleDbDataReader
            ' DLE: 109/12: LC rasters may have 0 as min value, but if minval > 1, then start at 1
            If (minVal > 1) Then minVal = 1
            For i = minVal To maxVal  ' DLE 10/9/2012: LC values may begin at 0, not 1
                'For i = 1 To maxVal  ' DLE 10/9/2012: LC values may begin at 0, not 1
                If (mwTable.CellValue(FieldIndex, rowidx) = i) Then 'And (pRow.Value(FieldIndex) = rsLandClass!Value) Then
                    dataCF = cmdCF.ExecuteReader

                    booValueFound = False
                    While dataCF.Read()
                        If mwTable.CellValue(FieldIndex, rowidx) = dataCF("Value") Then
                            booValueFound = True
                            If dataCF("W_WL") = 0 Then 'Means the current landclass is NOT Water or Wetland, therefore gets a 1
                                If strpick = "" Then
                                    strpick = "1"
                                Else
                                    strpick = strpick & ", " & "1"
                                End If
                            ElseIf dataCF("W_WL") = 1 Then 'Means the current landclass is Water or Wetland, therefore gets a 0
                                If strpick = "" Then
                                    strpick = "0"
                                Else
                                    strpick = strpick & ", " & "0"
                                End If
                            End If
                            rowidx = rowidx + 1
                            Exit While
                        Else
                            booValueFound = False
                        End If
                    End While
                    If booValueFound = False Then
                        MsgBox("Error: Your OpenNSPECT Land Class Table is missing values found in your landcover GRID dataset.")
                        ConstructPondPickStatement = Nothing
                        dataCF.Close()
                        mwTable.Close()
                        Exit Function
                    End If
                    dataCF.Close()

                Else
                    If strpick = "" Then
                        strpick = "0"
                    Else
                        strpick = strpick & ", " & nodata
                    End If
                End If

            Next
            mwTable.Close()
        End If

        Return strpick

    End Function


    ''' <summary>
    ''' Calculates the MUSLE erosion model
    ''' </summary>
    ''' <param name="concentrationStatement">The concentration statement.</param>
    ''' <param name="strConPondStatement">The STR con pond statement.</param>
    ''' <param name="OutputItems">The output items.</param><returns></returns>
    Private Function CalcMUSLE(ByRef concentrationStatement As String, ByRef strConPondStatement As String, ByRef OutputItems As OutputItems) As Boolean

        Dim pWSLengthRaster As Grid = Nothing
        'STEP 2: Watershed length
        Dim pWSLengthUnitsRaster As Grid = Nothing
        'STEP 3: Units conversion
        Dim pSlopePRRaster As Grid = Nothing
        'STEP 4: Average Slope
        Dim pSlopeModRaster As Grid = Nothing
        'STEP 4a: Mod Slope
        Dim pQuRaster As Grid = Nothing
        'STep 12b: Who Knows(b)
        Dim pHISYTempRaster As Grid = Nothing
        'STEP 16c: HI Specific temp yield
        Dim pHISYMGRaster As Grid = Nothing
        'STEP 17b: tons to milligrams, HI Specific
        Dim pPermMUSLERaster As Grid = Nothing
        'STEP 17c: Local Effects permanent raster
        Dim pTotSedMassHIRaster As Grid = Nothing
        'STEP 20HI: tot sed mass HI
        Dim pPermTotSedConcHIraster As Grid = Nothing
        'Permanent MUSLE
        Dim strMUSLE As String

        Dim progress = New SynchronousProgressDialog("Processing MUSLE Calculation...", 10, g_MainForm)

        Try
            progress.Increment("Calculating Watershed Length...")
            If SynchronousProgressDialog.KeepRunning Then
                'STEP 2: ------------------------------------------------------------------------------------
                'Calculate Watershed Length

                Dim pTauD8Flow As Grid = Nothing

                Dim tauD8calc = GetConverterToTauDemFromEsri()
                RasterMath(g_pFlowDirRaster, Nothing, Nothing, Nothing, Nothing, pTauD8Flow, Nothing, False, tauD8calc)
                pTauD8Flow.Header.NodataValue = -1

                Dim flowDir As String = GetTempFileNameTauDemGridExt()
                pTauD8Flow.Save()  'Saving because it seemed necessary.
                pTauD8Flow.Save(flowDir)

                Dim strLongestOut As String = GetTempFileNameTauDemGridExt()
                Dim strTotalOut As String = GetTempFileNameTauDemGridExt()
                Dim strStrahlOut As String = GetTempFileNameTauDemGridExt()

                'Use geoproc weightedAreaD8 after converting the D8 grid to taudem format bgd if needed
                Hydrology.PathLength(flowDir, strStrahlOut, strLongestOut, strTotalOut, Environment.ProcessorCount, False, Nothing)
                'strExpression = "flowlength([flowdir], [weight], upstream)"

                pTauD8Flow.Close()
                DataManagement.DeleteGrid(flowDir)

                pWSLengthRaster = New Grid
                pWSLengthRaster.Open(strLongestOut)

                'TODO: determine if this commented code is important
                'Dim gTemp As New MapWinGIS.Grid
                'gTemp.Open(strLongestOut)
                'Because taudem's flowlength is based on cell size, need to divide by cell size to put it back to where it should be.
                'Dim wslengthfixcalc As New RasterMathCellCalc(AddressOf wslengthfixCellCalc)
                'RasterMath(gTemp, Nothing, Nothing, Nothing, Nothing, pWSLengthRaster, wslengthfixcalc)
                'gTemp.Close()
                DataManagement.DeleteGrid(strTotalOut)
                DataManagement.DeleteGrid(strStrahlOut)
                'End STEP 2 ----------------------------------------------------------------------------------

                'STEP 3: --------------------------------------------------------------------------------------
                'Convert Metric Units
                Dim wslengthcalc As New RasterMathCellCalc(AddressOf wslengthCellCalc)
                RasterMath(pWSLengthRaster, Nothing, Nothing, Nothing, Nothing, pWSLengthUnitsRaster, wslengthcalc)
                'strExpression = "([cell_wslength] * 3.28084)"

                pWSLengthRaster.Close()
                DataManagement.DeleteGrid(strLongestOut)

                'END STEP 3: -----------------------------------------------------------------------------------
            End If

            progress.Increment("Calculating Mod Slope...")
            If SynchronousProgressDialog.KeepRunning Then
                'STEP 4a: ---------------------------------------------------------------------------------------
                'Calculate Average Slope
                Dim strtmpslpout As String = GetTempFileNameOutputGridExt()

                TerrainAnalysis.Slope2(g_pDEMRaster.Filename, 1, strtmpslpout, True, Nothing)
                'strExpression = "slope([dem], percentrise)"

                pSlopePRRaster = New Grid
                pSlopePRRaster.Open(strtmpslpout)

                'Calculate modslope
                Dim slpmodcalc As New RasterMathCellCalc(AddressOf slpmodCellCalc)
                RasterMath(pSlopePRRaster, Nothing, Nothing, Nothing, Nothing, pSlopeModRaster, slpmodcalc)
                'strExpression = "Con([slopepr] eq 0, 0.1, [slopepr])"
                pSlopePRRaster.Close()
                DataManagement.DeleteGrid(strtmpslpout)
                'END STEP 4a ------------------------------------------------------------------------------------
            End If

            progress.Increment("Calculating MUSLE...")
            If SynchronousProgressDialog.KeepRunning Then
                Dim AllMUSLECalc As New RasterMathCellCalc(AddressOf AllMUSLECellCalc)
                RasterMath(pWSLengthUnitsRaster, g_pSCS100Raster, pSlopeModRaster, g_pPrecipRaster, g_LandCoverRaster, pQuRaster, AllMUSLECalc)
            End If
            'modUtil.ReturnPermanentRaster(pQuRaster, modUtil.GetUniqueName("qu", g_XmlPrjFile.ProjectWorkspace, g_OutputGridExt))

            progress.Increment("Calculating MUSLE...")
            If SynchronousProgressDialog.KeepRunning Then
                _pondpicks = strConPondStatement.Split(",")
                Dim AllMUSLECalc2 As New RasterMathCellCalc(AddressOf AllMUSLECellCalc2)
                RasterMath(pQuRaster, g_LandCoverRaster, g_pDEMRaster, g_pMetRunoffRaster, Nothing, pHISYTempRaster, AllMUSLECalc2)
                pQuRaster.Close()
            End If
            'modUtil.ReturnPermanentRaster(pHISYTempRaster, modUtil.GetUniqueName("hisytmp", g_XmlPrjFile.ProjectWorkspace, g_OutputGridExt))

            progress.Increment("Calculating MUSLE...")
            If SynchronousProgressDialog.KeepRunning Then
                _picks = concentrationStatement.Split(",")
                Dim AllMUSLECalc3 As New RasterMathCellCalc(AddressOf AllMUSLECellCalc3)
                RasterMath(pHISYTempRaster, g_LandCoverRaster, g_KFactorRaster, g_pLSRaster, Nothing, pHISYMGRaster, AllMUSLECalc3)
                pHISYTempRaster.Close()
            End If
            'modUtil.ReturnPermanentRaster(pHISYMGRaster, modUtil.GetUniqueName("hisymg", g_XmlPrjFile.ProjectWorkspace, g_OutputGridExt))

            Dim pHISYMGRasterNoNull As Grid = Nothing
            Dim hisymgrnonullcalc As New RasterMathCellCalcNulls(AddressOf hisymgrnonullCellCalc)
            RasterMath(pHISYMGRaster, g_pDEMRaster, Nothing, Nothing, Nothing, pHISYMGRasterNoNull, Nothing, False, hisymgrnonullcalc)

            ' this was the one of two places where local effects would not calculate accumulations. For consistency we will keep running.
            If g_Project.IncludeLocalEffects Then

                If Not progress.Increment("Creating data layer for local effects...") Then Return False

                strMUSLE = GetUniqueFileName("locmusle", g_Project.ProjectWorkspace, OutputGridExt)
                'Added 7/23/04 to account for clip by selected polys functionality
                If g_Project.UseSelectedPolygons Then
                    pPermMUSLERaster = ClipBySelectedPoly(pHISYMGRasterNoNull, g_pSelectedPolyClip, strMUSLE)
                Else
                    pPermMUSLERaster = CopyRaster(pHISYMGRasterNoNull, strMUSLE)
                End If

                'metadata time
                g_dicMetadata.Add("MUSLE Local Effects (mg)", _strMusleMetadata)
                writeMetadata(g_Project.ProjectName, "MUSLE Local Effects (mg)", _strMusleMetadata, pPermMUSLERaster.Filename)

                AddOutputGridLayer(pPermMUSLERaster, "Brown", True, "MUSLE Local Effects (mg)", "MUSLE Local", -1, OutputItems)

            Else
                pHISYMGRasterNoNull.Save()
            End If

            progress.Increment("Calculating the accumulated sediment...")
            If SynchronousProgressDialog.KeepRunning Then
                Dim pTauD8Flow As Grid = Nothing

                Dim tauD8calc = GetConverterToTauDemFromEsri()
                RasterMath(g_pFlowDirRaster, Nothing, Nothing, Nothing, Nothing, pTauD8Flow, Nothing, False, tauD8calc)
                pTauD8Flow.Header.NodataValue = -1

                Dim flowDir As String = GetTempFileNameTauDemGridExt()
                pTauD8Flow.Save()  'Saving because it seemed necessary.
                pTauD8Flow.Save(flowDir)

                Dim pHISYMGRasterTmp As String = GetTempFileNameTauDemGridExt()
                pHISYMGRasterNoNull.Save()  'Saving because it seemed necessary.
                pHISYMGRasterNoNull.Save(pHISYMGRasterTmp)

                Dim strtmpout As String = GetTempFileNameTauDemGridExt()

                Hydrology.WeightedAreaD8(flowDir, pHISYMGRasterTmp, "", strtmpout, False, False, Environment.ProcessorCount, False, Nothing)

                pTotSedMassHIRaster = New Grid
                pTotSedMassHIRaster.Open(strtmpout)

                pTauD8Flow.Close()
                DataManagement.DeleteGrid(flowDir)
                pHISYMGRaster.Close()
                DataManagement.DeleteGrid(pHISYMGRasterTmp)
            End If

            progress.Increment("Adding Sediment Mass to Group Layer...")
            If SynchronousProgressDialog.KeepRunning Then
                'STEP 21: Created the Sediment Mass Raster layer and add to Group Layer -----------------------------------
                'Get a unique name for MUSLE and return the permanently made raster
                strMUSLE = GetUniqueFileName("MUSLEmass", g_Project.ProjectWorkspace, OutputGridExt)

                'Clip to selected polys if chosen
                If g_Project.UseSelectedPolygons Then
                    pPermTotSedConcHIraster = ClipBySelectedPoly(pTotSedMassHIRaster, g_pSelectedPolyClip, strMUSLE)
                Else
                    pPermTotSedConcHIraster = CopyRaster(pTotSedMassHIRaster, strMUSLE)
                End If

                'Metadata:
                g_dicMetadata.Add("MUSLE Sediment Mass (kg)", _strMusleMetadata)
                writeMetadata(g_Project.ProjectName, "MUSLE Sediment Mass (kg)", _strMusleMetadata, pPermTotSedConcHIraster.Filename)

                AddOutputGridLayer(pPermTotSedConcHIraster, "Brown", True, "MUSLE Sediment Mass (kg)", "MUSLE Accum", -1, OutputItems)

            End If

            CalcMUSLE = True
        Catch ex As Exception
            HandleError(ex)
            CalcMUSLE = False
        Finally
            progress.Dispose()
        End Try
    End Function

#Region "Raster Math"

    'Private Function weightCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
    '    'strExpression = "Con([flowdir] eq 2, 1.41421, " & "Con([flowdir] eq 8, 1.41421, " & "Con([flowdir] eq 32, 1.41421, " & "Con([flowdir] eq 128, 1.41421, 1.0))))"
    '    'Con(
    '    '  [flowdir] eq 2, 
    '    '  1.41421, 
    '    '  Con(
    '    '    [flowdir] eq 8
    '    '    1.41421
    '    '    Con(
    '    '      [flowdir] eq 32
    '    '      1.41421
    '    '      Con(
    '    '        [flowdir] eq 128
    '    '        1.41421
    '    '        1.0))))
    '    'TODO: add the taudem equivalents
    '    If Input1 = 2 Or Input1 = 8 Or Input1 = 32 Or Input1 = 128 Then
    '        Return 1.41421
    '    Else
    '        Return 1
    '    End If

    'End Function

    'Private Function wslengthfixCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
    '    'Because taudem's flowlength is based on cell size, need to divide by cell size to put it back to where it should be.
    '    Return Input1 / g_dblCellSize
    'End Function

    Private Function wslengthCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        'strExpression = "([cell_wslength] * 3.28084)"
        Return Input1 * 3.28084
    End Function

    Private Function slpmodCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        'strExpression = "Con([slopepr] eq 0, 0.1, [slopepr])"
        If Input1 = 0 Then
            Return 0.1
        Else
            Return Input1
        End If
    End Function

    Private Function c0CellCalc0(ByVal Input1 As Single) As Single
        If Input1 <= 0.1 Then
            Return 2.3055
        ElseIf Input1 > 0.1 And Input1 < 0.2 Then
            Return 2.23537
        ElseIf Input1 >= 0.2 And Input1 < 0.25 Then
            Return 2.18219
        ElseIf Input1 >= 0.25 And Input1 < 0.3 Then
            Return 2.10624
        ElseIf Input1 >= 0.3 And Input1 < 0.35 Then
            Return 2.00303
        ElseIf Input1 >= 0.35 And Input1 < 0.4 Then
            Return 1.87733
        ElseIf Input1 >= 0.4 And Input1 < 0.45 Then
            Return 1.76312
        Else
            Return 1.67889
        End If

    End Function

    Private Function c0CellCalc1(ByVal Input1 As Single) As Single
        If Input1 <= 0.1 Then
            Return 2.0325
        ElseIf Input1 > 0.1 And Input1 < 0.2 Then
            Return 1.91978
        ElseIf Input1 >= 0.2 And Input1 < 0.25 Then
            Return 1.83842
        ElseIf Input1 >= 0.25 And Input1 < 0.3 Then
            Return 1.72657
        Else
            Return 1.63417
        End If
    End Function

    Private Function c0CellCalc2(ByVal Input1 As Single) As Single
        If Input1 <= 0.1 Then
            Return 2.55323
        ElseIf Input1 > 0.1 And Input1 < 0.3 Then
            Return 2.46532
        ElseIf Input1 >= 0.3 And Input1 < 0.35 Then
            Return 2.41896
        ElseIf Input1 >= 0.35 And Input1 < 0.4 Then
            Return 2.36409
        ElseIf Input1 >= 0.4 And Input1 < 0.45 Then
            Return 2.29238
        Else
            Return 2.20282
        End If
    End Function

    Private Function c0CellCalc3(ByVal Input1 As Single) As Single
        If Input1 <= 0.1 Then
            Return 2.47317
        ElseIf Input1 > 0.1 And Input1 < 0.3 Then
            Return 2.39628
        ElseIf Input1 >= 0.3 And Input1 < 0.35 Then
            Return 2.35477
        ElseIf Input1 >= 0.35 And Input1 < 0.4 Then
            Return 2.30726
        ElseIf Input1 >= 0.4 And Input1 < 0.45 Then
            Return 2.24876
        Else
            Return 2.17772
        End If
    End Function

    Private Function c1CellCalc0(ByVal Input1 As Single) As Single
        If Input1 <= 0.1 Then
            Return -0.51429
        ElseIf Input1 > 0.1 And Input1 < 0.2 Then
            Return -0.50387
        ElseIf Input1 >= 0.2 And Input1 < 0.25 Then
            Return -0.48488
        ElseIf Input1 >= 0.25 And Input1 < 0.3 Then
            Return -0.45695
        ElseIf Input1 >= 0.3 And Input1 < 0.35 Then
            Return -0.40769
        ElseIf Input1 >= 0.35 And Input1 < 0.4 Then
            Return -0.32274
        ElseIf Input1 >= 0.4 And Input1 < 0.45 Then
            Return -0.15644
        Else
            Return -0.0693
        End If
    End Function

    Private Function c1CellCalc1(ByVal Input1 As Single) As Single
        If Input1 <= 0.1 Then
            Return 2.0325
        ElseIf Input1 > 0.1 And Input1 < 0.2 Then
            Return 1.91978
        ElseIf Input1 >= 0.2 And Input1 < 0.25 Then
            Return 1.83842
        ElseIf Input1 >= 0.25 And Input1 < 0.3 Then
            Return 1.72657
        Else
            Return 1.63417
        End If
    End Function

    Private Function c1CellCalc2(ByVal Input1 As Single) As Single
        If Input1 <= 0.1 Then
            Return -0.31583
        ElseIf Input1 > 0.1 And Input1 < 0.2 Then
            Return -0.28215
        ElseIf Input1 >= 0.2 And Input1 < 0.25 Then
            Return -0.25543
        ElseIf Input1 >= 0.25 And Input1 < 0.3 Then
            Return -0.19826
        Else
            Return -0.091
        End If

    End Function

    Private Function c1CellCalc3(ByVal Input1 As Single) As Single
        If Input1 <= 0.1 Then
            Return -0.51848
        ElseIf Input1 > 0.1 And Input1 < 0.3 Then
            Return -0.51202
        ElseIf Input1 >= 0.3 And Input1 < 0.35 Then
            Return -0.49735
        ElseIf Input1 >= 0.35 And Input1 < 0.4 Then
            Return -0.46541
        ElseIf Input1 >= 0.4 And Input1 < 0.45 Then
            Return -0.41314
        Else
            Return -0.36803
        End If

    End Function

    Private Function c2CellCalc0(ByVal Input1 As Single) As Single
        If Input1 <= 0.1 Then
            Return -0.1175
        ElseIf Input1 > 0.1 And Input1 < 0.2 Then
            Return -0.08929
        ElseIf Input1 >= 0.2 And Input1 < 0.25 Then
            Return -0.06589
        ElseIf Input1 >= 0.25 And Input1 < 0.3 Then
            Return -0.02835
        ElseIf Input1 <= 0.3 And Input1 < 0.35 Then
            Return 0.01983
        ElseIf Input1 >= 0.35 And Input1 < 0.4 Then
            Return 0.05754
        ElseIf Input1 >= 0.4 And Input1 < 0.45 Then
            Return 0.00453
        Else
            Return 0
        End If

    End Function

    Private Function c2CellCalc1(ByVal Input1 As Single) As Single
        If Input1 <= 0.1 Then
            Return -0.13748
        ElseIf Input1 > 0.1 And Input1 < 0.2 Then
            Return -0.0702
        ElseIf Input1 >= 0.2 And Input1 < 0.25 Then
            Return -0.02597
        ElseIf Input1 >= 0.25 And Input1 < 0.3 Then
            Return -0.02633
        Else
            Return 0
        End If
    End Function

    Private Function c2CellCalc2(ByVal Input1 As Single) As Single

        If Input1 <= 0.1 Then
            Return -0.16403
        ElseIf Input1 > 0.1 And Input1 < 0.3 Then
            Return -0.11657
        ElseIf Input1 >= 0.3 And Input1 < 0.35 Then
            Return -0.0882
        ElseIf Input1 >= 0.35 And Input1 < 0.4 Then
            Return -0.05621
        ElseIf Input1 >= 0.4 And Input1 < 0.45 Then
            Return -0.02281
        Else
            Return -0.01259
        End If
    End Function

    Private Function c2CellCalc3(ByVal Input1 As Single) As Single
        If Input1 <= 0.1 Then
            Return -0.17083
        ElseIf Input1 > 0.1 And Input1 < 0.3 Then
            Return -0.13245
        ElseIf Input1 >= 0.3 And Input1 < 0.35 Then
            Return -0.11985
        ElseIf Input1 >= 0.35 And Input1 < 0.4 Then
            Return -0.11094
        ElseIf Input1 >= 0.4 And Input1 < 0.45 Then
            Return -0.11508
        Else
            Return -0.09525
        End If
    End Function

    Private Function AllMUSLECellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        Dim tmp1val, tmp2val, tmp3val, tmp4val, lagval, tocval, toctmpval, modtocval, logtocval, logtoctmpval, RetentionVal, AbstractVal, abprecval, c0calc, c1calc, c2calc, logquval As Single

        'strExpression = "Pow([cell_wslngft], 0.8)"
        tmp1val = Math.Pow(Input1, 0.8)

        'strExpression = "(1000 / [scsgrid100]) - 9"
        If Input2 = 0 Then
            Return OutNull
        Else
            tmp2val = (1000 / Input2) - 9

            'strExpression = "Pow([temp4], 0.7)"
            tmp3val = Math.Pow(tmp2val, 0.7)

            'strExpression = "Pow([modslope], 0.5)"
            tmp4val = Math.Pow(Input3, 0.5)

            'strExpression = "([temp3] * [temp5]) / (1900 * [temp6])"
            If tmp4val = 0 Then
                Return OutNull
            Else
                lagval = (tmp1val * tmp3val) / (1900 * tmp4val)
                'strExpression = "[lag] / 0.6"
                tocval = lagval / 0.6

                'strExpression = "Con([toc] lt 0.1, 0.1, [toc])"
                If tocval < 0.1 Then
                    toctmpval = 0.1
                Else
                    toctmpval = tocval
                End If

                'strExpression = "Con([temp7] gt 10, 10, [temp7])"
                If toctmpval > 10 Then
                    modtocval = 10
                Else
                    modtocval = toctmpval
                End If

                'strExpression = "log10([modtoc])"
                logtocval = Math.Log10(modtocval)

                'strExpression = "Pow([logtoc], 2)"
                logtoctmpval = Math.Pow(logtocval, 2)

                If Input2 = 0 Then
                    AbstractVal = OutNull
                Else
                    If g_intRunoffPrecipType = 0 Then
                        RetentionVal = ((1000.0 / Input2) - 10) * g_intRainingDays
                    Else
                        RetentionVal = ((1000.0 / Input2) - 10)
                    End If

                    AbstractVal = 0.2 * RetentionVal
                End If

                'strExpression = "[abstract] / [rain]"
                If Input4 = 0 Then
                    Return OutNull
                Else
                    abprecval = AbstractVal / Input4
                    If abprecval = OutNull Then
                        c0calc = OutNull
                        c1calc = OutNull
                        c2calc = OutNull
                    Else
                        Select Case g_intPrecipType
                            Case 0
                                c0calc = c0CellCalc0(abprecval)
                                c1calc = c1CellCalc0(abprecval)
                                c2calc = c2CellCalc0(abprecval)
                            Case 1
                                c0calc = c0CellCalc1(abprecval)
                                c1calc = c1CellCalc1(abprecval)
                                c2calc = c2CellCalc1(abprecval)
                            Case 2
                                c0calc = c0CellCalc2(abprecval)
                                c1calc = c1CellCalc2(abprecval)
                                c2calc = c2CellCalc2(abprecval)
                            Case 3
                                c0calc = c0CellCalc3(abprecval)
                                c1calc = c1CellCalc3(abprecval)
                                c2calc = c2CellCalc3(abprecval)
                        End Select
                    End If

                    logquval = c0calc + (c1calc * logtocval) + (c2calc * logtoctmpval)

                    Return Math.Pow(10, logquval)
                End If
            End If
        End If
    End Function

    Private Function AllMUSLECellCalc2(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        Dim pondval, qpval, sqmival, afval, runoff_inches As Single

        For i As Integer = 0 To _pondpicks.Length - 1
            If Input2 = i + 1 Then
                pondval = _pondpicks(i)
                Exit For
            End If
        Next

        If Input3 > 0 Then
            'sq meters to square miles area
            sqmival = ((g_CellSize * g_CellSize) * 0.000001 * 0.386102)
            'cubic meters to cubic inches to inches
            runoff_inches = (Input4 / 0.016387064) / (g_CellSize * g_CellSize * 1550.0031)
        Else
            sqmival = 0
            runoff_inches = 0
        End If

        'strExpression = "[qu] * [cellarea_sqmi] * [runoff_in] * [pondfact]"
        qpval = Input1 * sqmival * runoff_inches * pondval

        'cubic meters to cubic inches to cubic feet to acresfeet
        afval = (Input4 / 0.016387064) * 0.0005787 * 0.000022957

        'strExpression = "Pow(([runoff_af] * [qp]), " & _dblMUSLEExp & ")"
        Return Math.Pow((afval * qpval), _dblMUSLEExp)

    End Function

    Private Function AllMUSLECellCalc3(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        Dim cfactval, hisyval As Single

        For i As Integer = 0 To _picks.Length - 1
            If Input2 = i + 1 Then
                cfactval = _picks(i)
                Exit For
            End If
        Next

        'strExpression = _dblMUSLEVal & " * ([cfactor] * [kfactor] * [lsfactor] * [temp9])"
        hisyval = _dblMUSLEVal * (cfactval * Input3 * Input4 * Input1)

        'strExpression = "[sy] * 907.184740"
        Return hisyval * 907.18474

    End Function

    Function hisymgrnonullCellCalc(ByVal Input1 As Single, ByVal Input1Null As Single, ByVal Input2 As Single, ByVal Input2Null As Single, ByVal Input3 As Single, ByVal Input3Null As Single, ByVal Input4 As Single, ByVal Input4Null As Single, ByVal Input5 As Single, ByVal Input5Null As Single, ByVal OutNull As Single) As Single
        If Input1 <> Input1Null Then
            Return Input1
        Else
            If Input2 >= 0 Then
                Return 0
            End If
        End If
    End Function

#End Region
End Module