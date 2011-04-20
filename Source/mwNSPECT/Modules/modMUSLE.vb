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

Module modMUSLE
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
    Private _strCFactorConStatement As String 'C Factor Con Statement
    Private _strPondConStatement As String 'Pond Factor con statement
    Private _dblMUSLEVal As Double 'Customizable MUSLE value in Equation
    Private _dblMUSLEExp As Double 'Customizable musle exponent in equation

    Private _strMusleMetadata As String 'MUSLE metadata string

    ' Variables used by the Error handler function - DO NOT REMOVE
    Const c_sModuleFileName As String = "modMUSLE.vb"
    Private _picks As String()
    Private _pondpicks As String()

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="strSoilsDefName"></param>
    ''' <param name="strKfactorFileName"></param>
    ''' <param name="strLandClass"></param>
    ''' <param name="OutputItems"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function MUSLESetup(ByRef strSoilsDefName As String, ByRef strKfactorFileName As String, ByRef strLandClass As String, ByRef OutputItems As clsXMLOutputItems) As Boolean
        'Sub takes incoming parameters from the project file and then parses them out
        'strSoilsDefName: Name of the Soils Definition being used
        'strKFactorFileName: K Factor FileName
        'strLandClass: Name of the Landclass we're using

        Dim strTempLCType As String 'Our potential holder for a temp landtype

        Dim strSoilsDef As String

        'Open Strings
        Dim strCovFactor As String
        Dim strError As String = ""

        'STEP 1: Get the MUSLE Values
        strSoilsDef = "SELECT * FROM SOILS WHERE NAME LIKE '" & strSoilsDefName & "'"
        Dim cmdsoils As New OleDbCommand(strSoilsDef, g_DBConn)
        Dim datasoils As OleDbDataReader = cmdsoils.ExecuteReader
        datasoils.Read()

        _dblMUSLEVal = datasoils("MUSLEVal")
        _dblMUSLEExp = datasoils("MUSLEExp")
        datasoils.Close()

        'STEP 2: Set the K Factor Raster
        If modUtil.RasterExists(strKfactorFileName) Then
            g_KFactorRaster = modUtil.ReturnRaster(strKfactorFileName)
        Else
            strError = "K Factor Raster Does Not Exist: " & strKfactorFileName
        End If
        'END STEP 1: -----------------------------------------------------------------------------------

        If g_DictTempNames.Count > 0 AndAlso Len(g_DictTempNames.Item(strLandClass)) > 0 Then
            strTempLCType = g_DictTempNames.Item(strLandClass)
        Else
            strTempLCType = strLandClass
        End If

        'Get the landclasses of type strLandClass
        strCovFactor = "SELECT LCTYPE.LCTYPEID, LCCLASS.NAME, LCCLASS.VALUE, LCCLASS.COVERFACTOR, LCCLASS.W_WL FROM " & "LCTYPE INNER JOIN LCCLASS ON LCTYPE.LCTYPEID = LCCLASS.LCTYPEID " & "WHERE LCTYPE.NAME LIKE '" & strTempLCType & "' ORDER BY LCCLASS.VALUE"
        Dim cmdCovfact As New OleDbCommand(strCovFactor, g_DBConn)

        _strMusleMetadata = CreateMetadata(g_booLocalEffects) ', rsCoverFactor)

        If Len(strError) > 0 Then
            MsgBox(strError)
            Exit Function
        End If

        'Get the con statement for the cover factor calculation
        _strCFactorConStatement = ConstructPickStatment(cmdCovfact, g_LandCoverRaster)
        _strPondConStatement = ConstructPondPickStatement(cmdCovfact, g_LandCoverRaster)

        'Calc rusle using the con
        If CalcMUSLE(_strCFactorConStatement, _strPondConStatement, OutputItems) Then
            MUSLESetup = True
        Else
            MUSLESetup = False
        End If

    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="booLocal"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CreateMetadata(ByRef booLocal As Boolean) As String

        Dim strHeader As String
        'Dim i As Integer
        'Dim strCFactor As String

        'Set up the header w/or without flow direction
        If booLocal = True Then
            strHeader = vbTab & "Input Datasets:" & vbNewLine & vbTab & vbTab & "Hydrologic soils grid: " & g_clsXMLPrjFile.strSoilsHydFileName & vbNewLine & vbTab & vbTab & "Landcover grid: " & g_clsXMLPrjFile.strLCGridFileName & vbNewLine & vbTab & vbTab & "Precipitation grid: " & g_strPrecipFileName & vbNewLine & vbTab & vbTab & "Soils K Factor: " & g_clsXMLPrjFile.strSoilsKFileName & vbNewLine & vbTab & vbTab & "LS Factor grid: " & g_strLSFileName & vbNewLine & g_strLandCoverParameters & vbNewLine 'append the g_strLandCoverParameters that was set up during runoff
        Else
            strHeader = vbTab & "Input Datasets:" & vbNewLine & vbTab & vbTab & "Hydrologic soils grid: " & g_clsXMLPrjFile.strSoilsHydFileName & vbNewLine & vbTab & vbTab & "Landcover grid: " & g_clsXMLPrjFile.strLCGridFileName & vbNewLine & vbTab & vbTab & "Precipitation grid: " & g_strPrecipFileName & vbNewLine & vbTab & vbTab & "Soils K Factor: " & g_clsXMLPrjFile.strSoilsKFileName & vbNewLine & vbTab & vbTab & "LS Factor grid: " & g_strLSFileName & vbNewLine & vbTab & vbTab & "Flow direction grid: " & g_strFlowDirFilename & vbNewLine & g_strLandCoverParameters & vbNewLine
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

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="cmdType"></param>
    ''' <param name="pLCRaster"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
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
                mwTable.Close()
            End If

            'strCompleteCon = strCon & strParens
            'ConstructPickStatment = strCompleteCon
            ConstructPickStatment = strpick

        Catch ex As Exception
            MsgBox("Error in pick Statement: " & Err.Number & ": " & Err.Description)
        End Try
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="cmdCF"></param>
    ''' <param name="pLCRaster"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ConstructPondPickStatement(ByRef cmdCF As OleDbCommand, ByRef pLCRaster As MapWinGIS.Grid) As String
        'Creates the Con Statement used in the Pond Factor GRID
        'Returns: String
        'Looks like: con(([nu_lulc] eq 16), 0, con((nu_lulc eq 17), 0...
        ConstructPondPickStatement = ""

        Dim TableExist As Boolean
        Dim FieldIndex As Short
        Dim booValueFound As Boolean
        Dim i As Short

        Dim maxVal As Integer = pLCRaster.Maximum
        Dim nodata As Single = pLCRaster.Header.NodataValue

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
            Dim dataCF As OleDbDataReader
            For i = 1 To maxVal
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
                        MsgBox("Error: Your N-SPECT Land Class Table is missing values found in your landcover GRID dataset.")
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


        'strCompleteCon = strCon & strParens
        ConstructPondPickStatement = strpick

    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="strConStatement"></param>
    ''' <param name="strConPondStatement"></param>
    ''' <param name="OutputItems"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CalcMUSLE(ByRef strConStatement As String, ByRef strConPondStatement As String, ByRef OutputItems As clsXMLOutputItems) As Boolean
        'Incoming strings: strConStatment: the monster con statement
        'strConPondstatement: the con for the pond stuff
        'Calculates the MUSLE erosion model

        Dim pWSLengthRaster As MapWinGIS.Grid = Nothing  'STEP 2: Watershed length
        Dim pWSLengthUnitsRaster As MapWinGIS.Grid = Nothing  'STEP 3: Units conversion
        Dim pSlopePRRaster As MapWinGIS.Grid = Nothing  'STEP 4: Average Slope
        Dim pSlopeModRaster As MapWinGIS.Grid = Nothing  'STEP 4a: Mod Slope
        Dim pQuRaster As MapWinGIS.Grid = Nothing 'STep 12b: Who Knows(b)
        Dim pHISYTempRaster As MapWinGIS.Grid = Nothing  'STEP 16c: HI Specific temp yield
        Dim pHISYMGRaster As MapWinGIS.Grid = Nothing  'STEP 17b: tons to milligrams, HI Specific
        Dim pPermMUSLERaster As MapWinGIS.Grid = Nothing  'STEP 17c: Local Effects permanent raster
        Dim pTotSedMassHIRaster As MapWinGIS.Grid = Nothing  'STEP 20HI: tot sed mass HI
        Dim pPermTotSedConcHIraster As MapWinGIS.Grid = Nothing  'Permanent MUSLE
        Dim strMUSLE As String


        'String to hold calculations
        Dim strExpression As String = ""
        Const strTitle As String = "Processing MUSLE Calculation..."

        Try

            modProgDialog.ProgDialog("Calculating Watershed Length...", strTitle, 0, 27, 2, g_frmProjectSetup)
            If modProgDialog.g_boolCancel Then
                'STEP 2: ------------------------------------------------------------------------------------
                'Calculate Watershed Length

                Dim pTauD8Flow As MapWinGIS.Grid = Nothing

                Dim tauD8calc As New RasterMathCellCalcNulls(AddressOf tauD8CellCalc)
                RasterMath(g_pFlowDirRaster, Nothing, Nothing, Nothing, Nothing, pTauD8Flow, Nothing, False, tauD8calc)
                pTauD8Flow.Header.NodataValue = -1

                Dim strtmp1 As String = IO.Path.GetTempFileName
                g_TempFilesToDel.Add(strtmp1)
                strtmp1 = strtmp1 + g_TAUDEMGridExt
                g_TempFilesToDel.Add(strtmp1)
                MapWinGeoProc.DataManagement.DeleteGrid(strtmp1)
                pTauD8Flow.Save(strtmp1)

                Dim strLongestOut As String = IO.Path.GetTempFileName
                g_TempFilesToDel.Add(strLongestOut)
                strLongestOut = strLongestOut + "out" + g_TAUDEMGridExt
                g_TempFilesToDel.Add(strLongestOut)
                MapWinGeoProc.DataManagement.DeleteGrid(strLongestOut)

                Dim strTotalOut As String = IO.Path.GetTempFileName
                g_TempFilesToDel.Add(strTotalOut)
                strTotalOut = strTotalOut + "out" + g_TAUDEMGridExt
                g_TempFilesToDel.Add(strTotalOut)
                MapWinGeoProc.DataManagement.DeleteGrid(strTotalOut)

                Dim strStrahlOut As String = IO.Path.GetTempFileName
                g_TempFilesToDel.Add(strStrahlOut)
                strStrahlOut = strStrahlOut + "out" + g_TAUDEMGridExt
                g_TempFilesToDel.Add(strStrahlOut)
                MapWinGeoProc.DataManagement.DeleteGrid(strStrahlOut)

                'Use geoproc weightedAreaD8 after converting the D8 grid to taudem format bgd if needed
                MapWinGeoProc.Hydrology.PathLength(strtmp1, strStrahlOut, strLongestOut, strTotalOut, Environment.ProcessorCount, Nothing)
                'strExpression = "flowlength([flowdir], [weight], upstream)"

                pTauD8Flow.Close()
                MapWinGeoProc.DataManagement.DeleteGrid(strtmp1)

                pWSLengthRaster = New MapWinGIS.Grid
                pWSLengthRaster.Open(strLongestOut)

                'Dim gTemp As New MapWinGIS.Grid
                'gTemp.Open(strLongestOut)
                ''Because taudem's flowlength is based on cell size, need to divide by cell size to put it back to where it should be.
                'Dim wslengthfixcalc As New RasterMathCellCalc(AddressOf wslengthfixCellCalc)
                'RasterMath(gTemp, Nothing, Nothing, Nothing, Nothing, pWSLengthRaster, wslengthfixcalc)
                'gTemp.Close()
                MapWinGeoProc.DataManagement.DeleteGrid(strTotalOut)
                MapWinGeoProc.DataManagement.DeleteGrid(strStrahlOut)
                'End STEP 2 ----------------------------------------------------------------------------------

                'STEP 3: --------------------------------------------------------------------------------------
                'Convert Metric Units
                Dim wslengthcalc As New RasterMathCellCalc(AddressOf wslengthCellCalc)
                RasterMath(pWSLengthRaster, Nothing, Nothing, Nothing, Nothing, pWSLengthUnitsRaster, wslengthcalc)
                'strExpression = "([cell_wslength] * 3.28084)"

                pWSLengthRaster.Close()
                MapWinGeoProc.DataManagement.DeleteGrid(strLongestOut)

                'END STEP 3: -----------------------------------------------------------------------------------
            End If

            modProgDialog.ProgDialog("Calculating Mod Slope...", strTitle, 0, 27, 4, g_frmProjectSetup)
            If modProgDialog.g_boolCancel Then
                'STEP 4a: ---------------------------------------------------------------------------------------
                'Calculate Average Slope
                Dim strtmpslpout As String = IO.Path.GetTempFileName
                g_TempFilesToDel.Add(strtmpslpout)
                strtmpslpout = strtmpslpout + g_OutputGridExt
                g_TempFilesToDel.Add(strtmpslpout)
                MapWinGeoProc.DataManagement.DeleteGrid(strtmpslpout)
                MapWinGeoProc.TerrainAnalysis.Slope2(g_pDEMRaster.Filename, 1, strtmpslpout, True, Nothing)
                'strExpression = "slope([dem], percentrise)"

                pSlopePRRaster = New MapWinGIS.Grid
                pSlopePRRaster.Open(strtmpslpout)

                'Calculate modslope
                Dim slpmodcalc As New RasterMathCellCalc(AddressOf slpmodCellCalc)
                RasterMath(pSlopePRRaster, Nothing, Nothing, Nothing, Nothing, pSlopeModRaster, slpmodcalc)
                'strExpression = "Con([slopepr] eq 0, 0.1, [slopepr])"
                pSlopePRRaster.Close()
                MapWinGeoProc.DataManagement.DeleteGrid(strtmpslpout)
                'END STEP 4a ------------------------------------------------------------------------------------
            End If



            modProgDialog.ProgDialog("Calculating MUSLE...", strTitle, 0, 27, 18, g_frmProjectSetup)
            If modProgDialog.g_boolCancel Then
                Dim AllMUSLECalc As New RasterMathCellCalc(AddressOf AllMUSLECellCalc)
                RasterMath(pWSLengthUnitsRaster, g_pSCS100Raster, pSlopeModRaster, g_pPrecipRaster, g_LandCoverRaster, pQuRaster, AllMUSLECalc)
            End If
            'modUtil.ReturnPermanentRaster(pQuRaster, modUtil.GetUniqueName("qu", g_strWorkspace, g_OutputGridExt))

            modProgDialog.ProgDialog("Calculating MUSLE...", strTitle, 0, 27, 22, g_frmProjectSetup)
            If modProgDialog.g_boolCancel Then
                ReDim _pondpicks(strConPondStatement.Split(",").Length)
                _pondpicks = strConPondStatement.Split(",")
                Dim AllMUSLECalc2 As New RasterMathCellCalc(AddressOf AllMUSLECellCalc2)
                RasterMath(pQuRaster, g_LandCoverRaster, g_pDEMRaster, g_pMetRunoffRaster, Nothing, pHISYTempRaster, AllMUSLECalc2)
                pQuRaster.Close()
            End If
            'modUtil.ReturnPermanentRaster(pHISYTempRaster, modUtil.GetUniqueName("hisytmp", g_strWorkspace, g_OutputGridExt))

            modProgDialog.ProgDialog("Calculating MUSLE...", strTitle, 0, 27, 25, g_frmProjectSetup)
            If modProgDialog.g_boolCancel Then
                ReDim _picks(strConStatement.Split(",").Length)
                _picks = strConStatement.Split(",")
                Dim AllMUSLECalc3 As New RasterMathCellCalc(AddressOf AllMUSLECellCalc3)
                RasterMath(pHISYTempRaster, g_LandCoverRaster, g_KFactorRaster, g_pLSRaster, Nothing, pHISYMGRaster, AllMUSLECalc3)
                pHISYTempRaster.Close()
            End If
            'modUtil.ReturnPermanentRaster(pHISYMGRaster, modUtil.GetUniqueName("hisymg", g_strWorkspace, g_OutputGridExt))


            Dim pHISYMGRasterNoNull As MapWinGIS.Grid = Nothing
            Dim hisymgrnonullcalc As New RasterMathCellCalcNulls(AddressOf hisymgrnonullCellCalc)
            RasterMath(pHISYMGRaster, g_pDEMRaster, Nothing, Nothing, Nothing, pHISYMGRasterNoNull, Nothing, False, hisymgrnonullcalc)

            If g_booLocalEffects Then

                modProgDialog.ProgDialog("Creating data layer for local effects...", strTitle, 0, 27, 27, g_frmProjectSetup)
                If modProgDialog.g_boolCancel Then

                    strMUSLE = modUtil.GetUniqueName("locmusle", g_strWorkspace, g_FinalOutputGridExt)
                    'Added 7/23/04 to account for clip by selected polys functionality
                    If g_booSelectedPolys Then
                        pPermMUSLERaster = modUtil.ClipBySelectedPoly(pHISYMGRasterNoNull, g_pSelectedPolyClip, strMUSLE)
                        'pPermMUSLERaster = modUtil.ClipBySelectedPoly(pHISYMGRaster, g_pSelectedPolyClip, strMUSLE)
                    Else
                        pPermMUSLERaster = modUtil.ReturnPermanentRaster(pHISYMGRasterNoNull, strMUSLE)
                        'pPermMUSLERaster = modUtil.ReturnPermanentRaster(pHISYMGRaster, strMUSLE)
                    End If

                    'metadata time
                    g_dicMetadata.Add("MUSLE Local Effects (mg)", _strMusleMetadata)

                    AddOutputGridLayer(pPermMUSLERaster, "Brown", True, "MUSLE Local Effects (mg)", "MUSLE Local", -1, OutputItems)

                    CalcMUSLE = True
                    modProgDialog.KillDialog()
                    Exit Function

                End If

            End If


            modProgDialog.ProgDialog("Calculating the accumulated sediment...", strTitle, 0, 27, 23, g_frmProjectSetup)
            If modProgDialog.g_boolCancel Then
                Dim pTauD8Flow As MapWinGIS.Grid = Nothing

                Dim tauD8calc As New RasterMathCellCalcNulls(AddressOf tauD8CellCalc)
                RasterMath(g_pFlowDirRaster, Nothing, Nothing, Nothing, Nothing, pTauD8Flow, Nothing, False, tauD8calc)
                pTauD8Flow.Header.NodataValue = -1

                Dim strtmp1 As String = IO.Path.GetTempFileName
                g_TempFilesToDel.Add(strtmp1)
                strtmp1 = strtmp1 + g_TAUDEMGridExt
                g_TempFilesToDel.Add(strtmp1)
                MapWinGeoProc.DataManagement.DeleteGrid(strtmp1)
                pTauD8Flow.Save(strtmp1)

                Dim strtmp2 As String = IO.Path.GetTempFileName
                g_TempFilesToDel.Add(strtmp2)
                strtmp2 = strtmp2 + g_TAUDEMGridExt
                g_TempFilesToDel.Add(strtmp2)
                MapWinGeoProc.DataManagement.DeleteGrid(strtmp2)
                pHISYMGRasterNoNull.Save(strtmp2)
                'pHISYMGRaster.Save(strtmp2)

                Dim strtmpout As String = IO.Path.GetTempFileName
                g_TempFilesToDel.Add(strtmpout)
                strtmpout = strtmpout + "out" + g_TAUDEMGridExt
                g_TempFilesToDel.Add(strtmpout)
                MapWinGeoProc.DataManagement.DeleteGrid(strtmpout)


                'Use geoproc weightedAreaD8 after converting the D8 grid to taudem format bgd if needed
                MapWinGeoProc.Hydrology.WeightedAreaD8(strtmp1, strtmp2, "", strtmpout, False, False, Environment.ProcessorCount, Nothing)
                'strExpression = "FlowAccumulation([flowdir], [pHISYMGRaster], FLOAT)"

                pTotSedMassHIRaster = New MapWinGIS.Grid
                pTotSedMassHIRaster.Open(strtmpout)

                pTauD8Flow.Close()
                MapWinGeoProc.DataManagement.DeleteGrid(strtmp1)
                pHISYMGRaster.Close()
                MapWinGeoProc.DataManagement.DeleteGrid(strtmp2)
            End If


            modProgDialog.ProgDialog("Adding Sediment Mass to Group Layer...", strTitle, 0, 27, 25, g_frmProjectSetup)
            If modProgDialog.g_boolCancel Then
                'STEP 21: Created the Sediment Mass Raster layer and add to Group Layer -----------------------------------
                'Get a unique name for MUSLE and return the permanently made raster
                strMUSLE = modUtil.GetUniqueName("MUSLEmass", g_strWorkspace, g_FinalOutputGridExt)

                'Clip to selected polys if chosen
                If g_booSelectedPolys Then
                    pPermTotSedConcHIraster = modUtil.ClipBySelectedPoly(pTotSedMassHIRaster, g_pSelectedPolyClip, strMUSLE)
                Else
                    pPermTotSedConcHIraster = modUtil.ReturnPermanentRaster(pTotSedMassHIRaster, strMUSLE)
                End If


                'Metadata:
                g_dicMetadata.Add("MUSLE Sediment Mass (kg)", _strMusleMetadata)

                AddOutputGridLayer(pPermTotSedConcHIraster, "Brown", True, "MUSLE Sediment Mass (kg)", "MUSLE Accum", -1, OutputItems)

            End If


            CalcMUSLE = True

            modProgDialog.KillDialog()


        Catch ex As Exception
            If Err.Number = -2147217297 Then 'S.A. constant for User cancelled operation
                modProgDialog.g_boolCancel = False
            ElseIf Err.Number = -2147467259 Then  'S.A. constant for crappy ESRI stupid GRID error
                MsgBox("ArcMap has reached the maximum number of GRIDs allowed in memory.  " & "Please exit N-SPECT and restart ArcMap.", MsgBoxStyle.Information, "Maximum GRID Number Encountered")
                CalcMUSLE = False
                modProgDialog.g_boolCancel = False
                modProgDialog.KillDialog()
            Else
                MsgBox("MUSLE Error: " & Err.Number & " on MUSLE Calculation: " & strExpression)
                CalcMUSLE = False
                modProgDialog.g_boolCancel = False
                modProgDialog.KillDialog()
            End If
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

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Input1"></param>
    ''' <param name="Input2"></param>
    ''' <param name="Input3"></param>
    ''' <param name="Input4"></param>
    ''' <param name="Input5"></param>
    ''' <param name="OutNull"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function wslengthCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        'strExpression = "([cell_wslength] * 3.28084)"
        Return Input1 * 3.28084
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Input1"></param>
    ''' <param name="Input2"></param>
    ''' <param name="Input3"></param>
    ''' <param name="Input4"></param>
    ''' <param name="Input5"></param>
    ''' <param name="OutNull"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function slpmodCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        'strExpression = "Con([slopepr] eq 0, 0.1, [slopepr])"
        If Input1 = 0 Then
            Return 0.1
        Else
            Return Input1
        End If
    End Function

    'Private Function tmp1lagCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
    '    'strExpression = "Pow([cell_wslngft], 0.8)"
    '    Return Math.Pow(Input1, 0.8)
    'End Function

    'Private Function tmp2lagCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
    '    'strExpression = "(1000 / [scsgrid100]) - 9"
    '    If Input1 = 0 Then
    '        Return OutNull
    '    Else
    '        Return (1000 / Input1) - 9
    '    End If
    'End Function

    'Private Function tmp3lagCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
    '    'strExpression = "Pow([temp4], 0.7)"
    '    Return Math.Pow(Input1, 0.7)
    'End Function

    'Private Function tmp4lagCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
    '    'strExpression = "Pow([modslope], 0.5)"
    '    Return Math.Pow(Input1, 0.5)
    'End Function

    'Private Function lagCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
    '    'strExpression = "([temp3] * [temp5]) / (1900 * [temp6])"
    '    If Input3 = 0 Then
    '        Return OutNull
    '    Else
    '        Return (Input1 * Input2) / (1900 * Input3)
    '    End If
    'End Function

    'Private Function tocCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
    '    'strExpression = "[lag] / 0.6"
    '    Return Input1 / 0.6
    'End Function

    'Private Function toctmpCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
    '    'strExpression = "Con([toc] lt 0.1, 0.1, [toc])"
    '    If Input1 < 0.1 Then
    '        Return 0.1
    '    Else
    '        Return Input1
    '    End If
    'End Function

    'Private Function modtocCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
    '    'strExpression = "Con([temp7] gt 10, 10, [temp7])"
    '    If Input1 > 10 Then
    '        Return 10
    '    Else
    '        Return Input1
    '    End If
    'End Function

    'Private Function abprecipCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
    '    'strExpression = "[abstract] / [rain]"
    '    If Input2 = 0 Then
    '        Return OutNull
    '    Else
    '        Return Input1 / Input2
    '    End If
    'End Function

    'Private Function logtocCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
    '    'strExpression = "log10([modtoc])"
    '    Return Math.Log10(Input1)
    'End Function

    'Private Function tmplogtocCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
    '    'strExpression = "Pow([logtoc], 2)"
    '    Return Math.Pow(Input1, 2)
    'End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Input1"></param>
    ''' <param name="Input2"></param>
    ''' <param name="Input3"></param>
    ''' <param name="Input4"></param>
    ''' <param name="Input5"></param>
    ''' <param name="OutNull"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function c0CellCalc0(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        'm_strCZero = "Con(([ip] le 0.10), 2.30550," & "Con(([ip] gt 0.10 and [ip] lt 0.20), 2.23537," & "Con(([ip] ge 0.20 and [ip] lt 0.25), 2.18219," & "Con(([ip] ge 0.25 and [ip] lt 0.30), 2.10624," & "Con(([ip] ge 0.30 and [ip] lt 0.35), 2.00303," & "Con(([ip] ge 0.35 and [ip] lt 0.40), 1.87733," & "Con(([ip] ge 0.40 and [ip] lt 0.45), 1.76312, 1.67889)))))))"
        'Con(
        '  ([ip] le 0.10),
        '  2.30550," & "
        '  Con(
        '    ([ip] gt 0.10 and [ip] lt 0.20),
        '    2.23537," & "
        '    Con(
        '      ([ip] ge 0.20 and [ip] lt 0.25),
        '      2.18219," & "
        '      Con(
        '        ([ip] ge 0.25 and [ip] lt 0.30),
        '        2.10624," & "
        '        Con(
        '          ([ip] ge 0.30 and [ip] lt 0.35),
        '          2.00303," & "
        '          Con(
        '            ([ip] ge 0.35 and [ip] lt 0.40),
        '            1.87733," & "
        '            Con(
        '              ([ip] ge 0.40 and [ip] lt 0.45),
        '              1.76312,
        '              1.67889)))))))
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

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Input1"></param>
    ''' <param name="Input2"></param>
    ''' <param name="Input3"></param>
    ''' <param name="Input4"></param>
    ''' <param name="Input5"></param>
    ''' <param name="OutNull"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function c0CellCalc1(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        'm_strCZero = "Con(([ip] le 0.10), 2.03250," & "Con(([ip] gt 0.10 and [ip] lt 0.20), 1.91978," & "Con(([ip] ge 0.20 and [ip] lt 0.25), 1.83842," & "Con(([ip] ge 0.25 and [ip] lt 0.30), 1.72657, 1.63417))))"
        'Con(
        '  ([ip] le 0.10),
        '  2.03250," & "
        '  Con(
        '    ([ip] gt 0.10 and [ip] lt 0.20),
        '    1.91978," & "
        '    Con(
        '      ([ip] ge 0.20 and [ip] lt 0.25),
        '      1.83842," & "
        '      Con(
        '        ([ip] ge 0.25 and [ip] lt 0.30),
        '        1.72657,
        '        1.63417))))
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

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Input1"></param>
    ''' <param name="Input2"></param>
    ''' <param name="Input3"></param>
    ''' <param name="Input4"></param>
    ''' <param name="Input5"></param>
    ''' <param name="OutNull"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function c0CellCalc2(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        'm_strCZero = "Con(([ip] le 0.10), 2.55323," & "Con(([ip] gt 0.10 and [ip] lt 0.30), 2.46532," & "Con(([ip] ge 0.30 and [ip] lt 0.35), 2.41896," & "Con(([ip] ge 0.35 and [ip] lt 0.40), 2.36409," & "Con(([ip] ge 0.40 and [ip] lt 0.45), 2.29238, 2.20282)))))"
        'Con(
        '  ([ip] le 0.10),
        '  2.55323," & "
        '  Con(
        '    ([ip] gt 0.10 and [ip] lt 0.30),
        '    2.46532," & "
        '    Con(
        '      ([ip] ge 0.30 and [ip] lt 0.35),
        '      2.41896," & "
        '      Con(
        '        ([ip] ge 0.35 and [ip] lt 0.40),
        '        2.36409," & "
        '        Con(
        '          ([ip] ge 0.40 and [ip] lt 0.45),
        '          2.29238,
        '          2.20282)))))
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

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Input1"></param>
    ''' <param name="Input2"></param>
    ''' <param name="Input3"></param>
    ''' <param name="Input4"></param>
    ''' <param name="Input5"></param>
    ''' <param name="OutNull"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function c0CellCalc3(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        'm_strCZero = "Con(([ip] le 0.10), 2.47317," & "Con(([ip] ge 0.10 and [ip] lt 0.30), 2.39628," & "Con(([ip] ge 0.30 and [ip] lt 0.35), 2.35477," & "Con(([ip] ge 0.35 and [ip] lt 0.40), 2.30726," & "Con(([ip] ge 0.40 and [ip] lt 0.45), 2.24876, 2.17772)))))"
        'Con(
        '  ([ip] le 0.10),
        '  2.47317," & "
        '  Con(
        '    ([ip] ge 0.10 and [ip] lt 0.30),
        '    2.39628," & "
        '    Con(
        '      ([ip] ge 0.30 and [ip] lt 0.35),
        '      2.35477," & "
        '      Con(
        '        ([ip] ge 0.35 and [ip] lt 0.40),
        '        2.30726," & "
        '        Con(
        '          ([ip] ge 0.40 and [ip] lt 0.45),
        '          2.24876,
        '          2.17772)))))
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

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Input1"></param>
    ''' <param name="Input2"></param>
    ''' <param name="Input3"></param>
    ''' <param name="Input4"></param>
    ''' <param name="Input5"></param>
    ''' <param name="OutNull"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function c1CellCalc0(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        'm_strCone = "Con(([ip] le 0.10), -0.51429," & "Con(([ip] gt 0.10 and [ip] lt 0.20), -0.50387," & "Con(([ip] ge 0.20 and [ip] lt 0.25), -0.48488," & "Con(([ip] ge 0.25 and [ip] lt 0.30), -0.45695," & "Con(([ip] ge 0.30 and [ip] lt 0.35), -0.40769," & "Con(([ip] ge 0.35 and [ip] lt 0.40), -0.32274," & "Con(([ip] ge 0.40 and [ip] lt 0.45), -0.15644, -0.06930)))))))"
        'Con(
        '  ([ip] le 0.10),
        '  -0.51429," & "
        '  Con(
        '    ([ip] gt 0.10 and [ip] lt 0.20),
        '    -0.50387," & "
        '    Con(
        '      ([ip] ge 0.20 and [ip] lt 0.25),
        '      -0.48488," & "
        '      Con(
        '        ([ip] ge 0.25 and [ip] lt 0.30),
        '        -0.45695," & "
        '        Con(
        '          ([ip] ge 0.30 and [ip] lt 0.35),
        '          -0.40769," & "
        '          Con(
        '            ([ip] ge 0.35 and [ip] lt 0.40),
        '            -0.32274," & "
        '            Con(
        '              ([ip] ge 0.40 and [ip] lt 0.45),
        '              -0.15644,
        '              -0.06930)))))))
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

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Input1"></param>
    ''' <param name="Input2"></param>
    ''' <param name="Input3"></param>
    ''' <param name="Input4"></param>
    ''' <param name="Input5"></param>
    ''' <param name="OutNull"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function c1CellCalc1(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        'm_strCone = "Con(([ip] le 0.10), 2.03250," & "Con(([ip] gt 0.10 and [ip] lt 0.20), 1.91978," & "Con(([ip] ge 0.20 and [ip] lt 0.25), 1.83842," & "Con(([ip] ge 0.25 and [ip] lt 0.30), 1.72657, 1.63417))))"
        'Con(
        '  ([ip] le 0.10),
        '  2.03250," & "
        '  Con(
        '    ([ip] gt 0.10 and [ip] lt 0.20),
        '    1.91978," & "
        '    Con(
        '      ([ip] ge 0.20 and [ip] lt 0.25),
        '      1.83842," & "
        '      Con(
        '        ([ip] ge 0.25 and [ip] lt 0.30),
        '        1.72657,
        '        1.63417))))
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

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Input1"></param>
    ''' <param name="Input2"></param>
    ''' <param name="Input3"></param>
    ''' <param name="Input4"></param>
    ''' <param name="Input5"></param>
    ''' <param name="OutNull"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function c1CellCalc2(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        'm_strCone = "Con(([ip] le 0.10), -0.31583," & "Con(([ip] gt 0.10 and [ip] lt 0.20), -0.28215," & "Con(([ip] ge 0.20 and [ip] lt 0.25), -0.25543," & "Con(([ip] ge 0.25 and [ip] lt 0.30), -0.19826, -0.09100))))"
        'Con(
        '  ([ip] le 0.10), 
        '  -0.31583,
        '  Con(
        '    ([ip] gt 0.10 and [ip] lt 0.20), 
        '    -0.28215,
        '    Con(
        '      ([ip] ge 0.20 and [ip] lt 0.25), 
        '      -0.25543," & "
        '      Con(
        '        ([ip] ge 0.25 and [ip] lt 0.30), 
        '        -0.19826, 
        '        -0.09100))))
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

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Input1"></param>
    ''' <param name="Input2"></param>
    ''' <param name="Input3"></param>
    ''' <param name="Input4"></param>
    ''' <param name="Input5"></param>
    ''' <param name="OutNull"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function c1CellCalc3(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        'm_strCone = "Con(([ip] le 0.10), -0.51848," & "Con(([ip] ge 0.10 and [ip] lt 0.30), -0.51202," & "Con(([ip] ge 0.30 and [ip] lt 0.35), -0.49735," & "Con(([ip] ge 0.35 and [ip] lt 0.40), -0.46541," & "Con(([ip] ge 0.40 and [ip] lt 0.45), -0.41314, -0.36803)))))"
        'Con(
        '  ([ip] le 0.10), 
        '  -0.51848," & "
        '  Con(
        '    ([ip] ge 0.10 and [ip] lt 0.30), 
        '    -0.51202," & "
        '    Con(
        '      ([ip] ge 0.30 and [ip] lt 0.35), 
        '      -0.49735," & "
        '      Con(
        '        ([ip] ge 0.35 and [ip] lt 0.40), 
        '        -0.46541," & "
        '        Con(
        '          ([ip] ge 0.40 and [ip] lt 0.45), 
        '          -0.41314, 
        '          -0.36803)))))
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

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Input1"></param>
    ''' <param name="Input2"></param>
    ''' <param name="Input3"></param>
    ''' <param name="Input4"></param>
    ''' <param name="Input5"></param>
    ''' <param name="OutNull"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function c2CellCalc0(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        'm_strCTwo = "Con(([ip] le 0.10), -0.11750," & "Con(([ip] gt 0.10 and [ip] lt 0.20), -0.08929," & "Con(([ip] ge 0.20 and [ip] lt 0.25), -0.06589," & "Con(([ip] ge 0.25 and [ip] lt 0.30), -0.02835," & "Con(([ip] ge 0.30 and [ip] lt 0.35), 0.01983," & "Con(([ip] ge 0.35 and [ip] lt 0.40), 0.05754," & "Con(([ip] ge 0.40 and [ip] lt 0.45), 0.00453, 0.00000)))))))"
        'Con(
        '  ([ip] le 0.10), 
        '  -0.11750," & "
        '  Con(
        '    ([ip] gt 0.10 and [ip] lt 0.20), 
        '    -0.08929," & "
        '    Con(
        '      ([ip] ge 0.20 and [ip] lt 0.25), 
        '      -0.06589," & "
        '      Con(
        '        ([ip] ge 0.25 and [ip] lt 0.30), 
        '        -0.02835," & "
        '        Con(
        '          ([ip] ge 0.30 and [ip] lt 0.35), 
        '          0.01983," & "
        '          Con(
        '            ([ip] ge 0.35 and [ip] lt 0.40), 
        '            0.05754," & "
        '            Con(
        '              ([ip] ge 0.40 and [ip] lt 0.45), 
        '              0.00453, 
        '              0.00000)))))))
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

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Input1"></param>
    ''' <param name="Input2"></param>
    ''' <param name="Input3"></param>
    ''' <param name="Input4"></param>
    ''' <param name="Input5"></param>
    ''' <param name="OutNull"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function c2CellCalc1(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        'm_strCTwo = "Con(([ip] le 0.10), -0.13748," & "Con(([ip] gt 0.10 and [ip] lt 0.20), -0.07020," & "Con(([ip] ge 0.20 and [ip] lt 0.25), -0.02597," & "Con(([ip] ge 0.25 and [ip] lt 0.30), -0.02633, -0.0))))"
        'Con(
        '  ([ip] le 0.10),
        '  -0.13748," & "
        '  Con(
        '    ([ip] gt 0.10 and [ip] lt 0.20),
        '    -0.07020," & "
        '    Con(
        '      ([ip] ge 0.20 and [ip] lt 0.25),
        '      -0.02597," & "
        '      Con(
        '        ([ip] ge 0.25 and [ip] lt 0.30),
        '        -0.02633,
        '        -0.0))))
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

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Input1"></param>
    ''' <param name="Input2"></param>
    ''' <param name="Input3"></param>
    ''' <param name="Input4"></param>
    ''' <param name="Input5"></param>
    ''' <param name="OutNull"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function c2CellCalc2(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        'm_strCTwo = "Con(([ip] le 0.10), -0.16403," & "Con(([ip] gt 0.10 and [ip] lt 0.30), -0.11657," & "Con(([ip] ge 0.30 and [ip] lt 0.35), -0.08820," & "Con(([ip] ge 0.35 and [ip] lt 0.40), -0.05621," & "Con(([ip] ge 0.40 and [ip] lt 0.45), -0.02281, -0.01259)))))"
        'Con(
        '  ([ip] le 0.10),
        '  -0.16403," & "
        '  Con(
        '    ([ip] gt 0.10 and [ip] lt 0.30),
        '    -0.11657," & "
        '    Con(
        '      ([ip] ge 0.30 and [ip] lt 0.35),
        '      -0.08820," & "
        '      Con(
        '        ([ip] ge 0.35 and [ip] lt 0.40),
        '        -0.05621," & "
        '        Con(
        '          ([ip] ge 0.40 and [ip] lt 0.45),
        '          -0.02281,
        '          -0.01259)))))
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

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Input1"></param>
    ''' <param name="Input2"></param>
    ''' <param name="Input3"></param>
    ''' <param name="Input4"></param>
    ''' <param name="Input5"></param>
    ''' <param name="OutNull"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function c2CellCalc3(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        'm_strCTwo = "Con(([ip] le 0.10), -0.17083," & "Con(([ip] ge 0.10 and [ip] lt 0.30), -0.13245," & "Con(([ip] ge 0.30 and [ip] lt 0.35), -0.11985," & "Con(([ip] ge 0.35 and [ip] lt 0.40), -0.11094," & "Con(([ip] ge 0.40 and [ip] lt 0.45), -0.11508, -0.09525)))))"
        'Con(
        '  ([ip] le 0.10),
        '  -0.17083," & "
        '  Con(
        '    ([ip] ge 0.10 and [ip] lt 0.30),
        '    -0.13245," & "
        '    Con(
        '      ([ip] ge 0.30 and [ip] lt 0.35),
        '      -0.11985," & "
        '      Con(
        '        ([ip] ge 0.35 and [ip] lt 0.40),
        '        -0.11094," & "
        '        Con(
        '          ([ip] ge 0.40 and [ip] lt 0.45),
        '          -0.11508,
        '          -0.09525)))))
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


    'pWSLengthUnitsRaster, g_pSCS100Raster, pSlopeModRaster, g_pPrecipRaster, g_LandCoverRaster
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Input1"></param>
    ''' <param name="Input2"></param>
    ''' <param name="Input3"></param>
    ''' <param name="Input4"></param>
    ''' <param name="Input5"></param>
    ''' <param name="OutNull"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function AllMUSLECellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        Dim tmp1val, tmp2val, tmp3val, tmp4val, lagval, tocval, toctmpval, modtocval, logtocval, _
             logtoctmpval, RetentionVal, AbstractVal, abprecval, c0calc, c1calc, c2calc, logquval As Single

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
                                c0calc = c0CellCalc0(abprecval, Nothing, Nothing, Nothing, Nothing, OutNull)
                            Case 1
                                c0calc = c0CellCalc1(abprecval, Nothing, Nothing, Nothing, Nothing, OutNull)
                            Case 2
                                c0calc = c0CellCalc2(abprecval, Nothing, Nothing, Nothing, Nothing, OutNull)
                            Case 3
                                c0calc = c0CellCalc3(abprecval, Nothing, Nothing, Nothing, Nothing, OutNull)
                        End Select

                        Select Case g_intPrecipType
                            Case 0
                                c1calc = c1CellCalc0(abprecval, Nothing, Nothing, Nothing, Nothing, OutNull)
                            Case 1
                                c1calc = c1CellCalc1(abprecval, Nothing, Nothing, Nothing, Nothing, OutNull)
                            Case 2
                                c1calc = c1CellCalc2(abprecval, Nothing, Nothing, Nothing, Nothing, OutNull)
                            Case 3
                                c1calc = c1CellCalc3(abprecval, Nothing, Nothing, Nothing, Nothing, OutNull)
                        End Select

                        Select Case g_intPrecipType
                            Case 0
                                c2calc = c2CellCalc0(abprecval, Nothing, Nothing, Nothing, Nothing, OutNull)
                            Case 1
                                c2calc = c2CellCalc1(abprecval, Nothing, Nothing, Nothing, Nothing, OutNull)
                            Case 2
                                c2calc = c2CellCalc2(abprecval, Nothing, Nothing, Nothing, Nothing, OutNull)
                            Case 3
                                c2calc = c2CellCalc3(abprecval, Nothing, Nothing, Nothing, Nothing, OutNull)
                        End Select
                    End If

                    'strExpression = "[czero] + ([cone] * [logtoc]) + ([ctwo] * [temp8])"
                    logquval = c0calc + (c1calc * logtocval) + (c2calc * logtoctmpval)

                    'strExpression = "Pow(10, [logqu])"
                    Return Math.Pow(10, logquval)
                End If
            End If
        End If
    End Function

    'quval   g_LandCoverRaster    g_pDEMRaster  g_pMetRunoffRaster  g_KFactorRaster, g_pLSRaster
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Input1"></param>
    ''' <param name="Input2"></param>
    ''' <param name="Input3"></param>
    ''' <param name="Input4"></param>
    ''' <param name="Input5"></param>
    ''' <param name="OutNull"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
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
            sqmival = ((g_dblCellSize * g_dblCellSize) * 0.000001 * 0.386102)
            'cubic meters to cubic inches to inches
            runoff_inches = (Input4 / 0.016387064) / (g_dblCellSize * g_dblCellSize * 1550.0031)
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

    'hisytmpval   g_LandCoverRaster   g_KFactorRaster, g_pLSRaster
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Input1"></param>
    ''' <param name="Input2"></param>
    ''' <param name="Input3"></param>
    ''' <param name="Input4"></param>
    ''' <param name="Input5"></param>
    ''' <param name="OutNull"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
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