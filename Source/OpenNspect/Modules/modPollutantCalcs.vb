'********************************************************************************************************
'File Name: modPollutantCalcs.vb
'Description: Functions handling the Pollutant calculations for the model
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

    Private _strPollName As String
    'Mod level variable for name of pollutant being used
    Private _strWQName As String
    'Mod level variable for name of water quality standard
    Private _strColor As String
    'Mod level variable holding the string of the pollutant color
    Private _strPollCoeffMetadata As String
    'Variable to hold coeffs for use in metadata

    Private _WQValue As Single
    Private _FlowMax As Single

    ' Constant used by the Error handler function - DO NOT REMOVE
    Private _ParentHWND As Integer
    ' Set this to get correct parenting of Error handler forms

    Private _picks() As String

    Public Function PollutantConcentrationSetup (ByRef clsPollutant As clsXMLPollutantItem, ByRef strLandClass As String, _
                                                 ByRef strWQName As String, ByRef OutputItems As clsXMLOutputItems) _
        As Boolean
        'Sub takes incoming parameters (in the form of a pollutant item) from the project file
        'and then parses them out
        Try
            'Open Strings
            Dim strPoll As String
            Dim strType As String
            Dim strField As String = ""
            Dim strConStatement As String = ""
            Dim strTempCoeffSet As String
            'Again, because of landuse, we have to check for 'temp' coeff sets and their use
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
                If Len (g_DictTempNames.Item (clsPollutant.strCoeffSet)) > 0 Then
                    strTempCoeffSet = g_DictTempNames.Item (clsPollutant.strCoeffSet)
                Else
                    strTempCoeffSet = clsPollutant.strCoeffSet
                End If
            Else
                strTempCoeffSet = clsPollutant.strCoeffSet
            End If

            If Len (strField) > 0 Then
                strPoll = "SELECT * FROM COEFFICIENTSET WHERE NAME LIKE '" & strTempCoeffSet & "'"
                Dim cmdPoll As New DataHelper (strPoll)
                Dim dataPoll As OleDbDataReader = cmdPoll.ExecuteReader()
                dataPoll.Read()
                strType = "SELECT LCCLASS.Value, LCCLASS.Name, COEFFICIENT." & strField & _
                          " As CoeffType, COEFFICIENT.CoeffID, COEFFICIENT.LCCLASSID " & _
                          "FROM LCCLASS LEFT OUTER JOIN COEFFICIENT " & "ON LCCLASS.LCCLASSID = COEFFICIENT.LCCLASSID " & _
                          "WHERE COEFFICIENT.COEFFSETID = " & dataPoll ("CoeffSetID") & " ORDER BY LCCLASS.VALUE"
                dataPoll.Close()
                Dim cmdType As New DataHelper (strType)

                Dim command As OleDbCommand = cmdType.GetCommand()
                strConStatement = ConstructPickStatment (command, g_LandCoverRaster)
                _strPollCoeffMetadata = ConstructMetaData (command, (clsPollutant.strCoeff), g_booLocalEffects)

            End If

            'Find out the color of the pollutant
            strPollColor = "Select Color from Pollutant where NAME LIKE '" & _strPollName & "'"
            Dim cmdPollColor As New DataHelper (strPollColor)
            Dim datapollcolor As OleDbDataReader = cmdPollColor.ExecuteReader()
            datapollcolor.Read()
            _strColor = CStr (datapollcolor ("Color"))
            datapollcolor.Close()

            If CalcPollutantConcentration (strConStatement, OutputItems) Then
                PollutantConcentrationSetup = True
            Else
                PollutantConcentrationSetup = False
            End If

        Catch ex As Exception
            HandleError (ex)
            'True, "PollutantConcentrationSetup " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, _ParentHWND)
            PollutantConcentrationSetup = False
        End Try
    End Function

    Private Function ConstructMetaData (ByRef cmdType As oledbcommand, ByRef strCoeffSet As String, _
                                        ByRef booLocal As Boolean) As String
        'Takes the rs and creates a string describing the pollutants and coefficients used in this run, will
        'later be added to the global dictionary

        Dim strConstructMetaData As String
        Dim strLandClassCoeff As String = ""
        Dim strHeader As String
        Dim i As Short

        If booLocal Then
            strHeader = vbTab & "Input Datasets:" & vbNewLine & vbTab & vbTab & "Hydrologic soils grid: " & _
                        g_clsXMLPrjFile.strSoilsHydFileName & vbNewLine & vbTab & vbTab & "Landcover grid: " & _
                        g_clsXMLPrjFile.strLCGridFileName & vbNewLine & vbTab & vbTab & "Landcover grid type: " & _
                        g_clsXMLPrjFile.strLCGridType & vbNewLine & vbTab & vbTab & "Landcover grid units: " & _
                        g_clsXMLPrjFile.strLCGridUnits & vbNewLine & vbTab & vbTab & "Precipitation grid: " & _
                        g_strPrecipFileName & vbNewLine

        Else
            strHeader = vbTab & "Input Datasets:" & vbNewLine & vbTab & vbTab & "Hydrologic soils grid: " & _
                        g_clsXMLPrjFile.strSoilsHydFileName & vbNewLine & vbTab & vbTab & "Landcover grid: " & _
                        g_clsXMLPrjFile.strLCGridFileName & vbNewLine & vbTab & vbTab & "Precipitation grid: " & _
                        g_strPrecipFileName & vbNewLine & vbTab & vbTab & "Flow direction grid: " & g_strFlowDirFilename & _
                        vbNewLine
        End If

        strConstructMetaData = vbTab & "Pollutant Coefficients:" & vbNewLine & vbTab & vbTab & "Pollutant: " & _
                               _strPollName & vbNewLine & vbTab & vbTab & "Coefficient Set: " & strCoeffSet & vbNewLine & _
                               vbTab & vbTab & _
                               "The following lists the landcover classes and associated coefficients used" & vbNewLine & _
                               vbTab & vbTab & "in the OpenNSPECT analysis run that created this dataset: " & vbNewLine

        Dim dataType As OleDbDataReader = cmdType.ExecuteReader()
        While dataType.Read()
            If i = 1 Then
                strLandClassCoeff = vbTab & vbTab & vbTab & dataType ("Name") & ": " & dataType ("CoeffType") & _
                                    vbNewLine
            Else
                strLandClassCoeff = strLandClassCoeff & vbTab & vbTab & vbTab & dataType ("Name") & ": " & _
                                    dataType ("CoeffType") & vbNewLine
            End If
        End While
        dataType.Close()

        ConstructMetaData = strHeader & g_strLandCoverParameters & strConstructMetaData & strLandClassCoeff

    End Function

    Private Function ConstructPickStatment (ByRef cmdType As OleDbCommand, ByRef pLCRaster As Grid) As String
        'Creates the initial pick statement using the name of the the LandCass [CCAP, for example]
        'and the Land Class Raster.  Returns a string

        ConstructPickStatment = ""
        Try
            Dim TableExist As Boolean
            Dim FieldIndex As Short
            Dim booValueFound As Boolean
            Dim i As Short
            Dim maxVal As Integer = pLCRaster.Maximum

            Dim tablepath As String = ""
            'Get the raster table
            Dim lcPath As String = pLCRaster.Filename
            If Path.GetFileName (lcPath) = "sta.adf" Then
                tablepath = Path.GetDirectoryName (lcPath) + ".dbf"
                If File.Exists (tablepath) Then

                    TableExist = True
                Else
                    TableExist = False
                End If
            Else
                tablepath = Path.ChangeExtension (lcPath, ".dbf")
                If File.Exists (tablepath) Then
                    TableExist = True
                Else
                    TableExist = False
                End If
            End If

            Dim strpick As String = ""

            Dim mwTable As New Table
            If Not TableExist Then
                MsgBox ( _
                        "No MapWindow-readable raster table was found. To create one using ArcMap 9.3+, add the raster to the default project, right click on its layer and select Open Attribute Table. Now click on the options button in the lower right and select Export. In the export path, navigate to the directory of the grid folder and give the export the name of the raster folder with the .dbf extension. i.e. if you are exporting a raster attribute table from a raster named landcover, export landcover.dbf into the same level directory as the folder.", _
                        MsgBoxStyle.Exclamation, "Raster Attribute Table Not Found")

                Return ""
            Else
                mwTable.Open (tablepath)

                'Get index of Value Field
                FieldIndex = - 1
                For fidx As Integer = 0 To mwTable.NumFields - 1
                    If mwTable.Field (fidx).Name.ToLower = "value" Then
                        FieldIndex = fidx
                        Exit For
                    End If
                Next

                Dim rowidx As Integer = 0
                Dim dataType As OleDbDataReader
                For i = 1 To maxVal

                    If i = 1 Then
                        If (mwTable.CellValue (FieldIndex, rowidx) = i) Then
                            dataType = cmdType.ExecuteReader
                            booValueFound = False
                            While dataType.Read()
                                If mwTable.CellValue (FieldIndex, rowidx) = dataType ("Value") Then
                                    booValueFound = True
                                    strpick = CStr (dataType ("CoeffType"))
                                    rowidx = rowidx + 1
                                    Exit While
                                Else
                                    booValueFound = False
                                End If
                            End While
                            If booValueFound = False Then
                                MsgBox ( _
                                        "Error: Your OpenNSPECT Land Class Table is missing values found in your landcover GRID dataset.")
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
                        If (mwTable.CellValue (FieldIndex, rowidx) = i) Then _
'And (pRow.Value(FieldIndex) = rsLandClass!Value) Then
                            dataType = cmdType.ExecuteReader

                            booValueFound = False
                            While dataType.Read()
                                If mwTable.CellValue (FieldIndex, rowidx) = dataType ("Value") Then
                                    booValueFound = True
                                    strpick = strpick & ", " & CStr (dataType ("CoeffType"))
                                    rowidx = rowidx + 1
                                    Exit While
                                Else
                                    booValueFound = False
                                End If
                            End While
                            If booValueFound = False Then
                                MsgBox ( _
                                        "Error: Your OpenNSPECT Land Class Table is missing values found in your landcover GRID dataset.")
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

                mwTable.Close()
            End If
            'strCompleteCon = strCon & strParens
            'ConstructPickStatment = strCompleteCon
            ConstructPickStatment = strpick

        Catch ex As Exception
            MsgBox ("Error in pick Statement: " & Err.Number & ": " & Err.Description)
        End Try
    End Function

    Private Function CalcPollutantConcentration (ByRef strConStatement As String, ByRef OutputItems As clsXMLOutputItems) _
        As Boolean

        Dim pMassVolumeRaster As Grid = Nothing
        Dim pPermMassVolumeRaster As Grid = Nothing
        Dim pAccumPollRaster As Grid = Nothing
        Dim pPermAccPollRaster As Grid = Nothing
        Dim pPermTotalConcRaster As Grid = Nothing
        Dim pTotalPollConc0Raster As Grid = Nothing
        'gets rid of no data...replace with 0

        'String to hold calculations
        Dim strTitle As String
        strTitle = "Processing " & _strPollName & " Conc. Calculation..."
        Dim strOutConc As String
        Dim strAccPoll As String

        Try
            ShowProgress ("Calculating Mass Volume...", strTitle, 0, 13, 2, g_frmProjectSetup)
            If g_KeepRunning Then
                'STEP 2: MASS OF PHOSPHORUS PRODUCED BY EACH CELL -----------------------------------------
                ReDim _picks(strConStatement.Split (",").Length)
                _picks = strConStatement.Split (",")
                Dim massvolcalc As New RasterMathCellCalc (AddressOf massvolCellCalc)
                RasterMath (g_LandCoverRaster, g_pMetRunoffRaster, Nothing, Nothing, Nothing, pMassVolumeRaster, _
                            massvolcalc)

                'END STEP 2: -------------------------------------------------------------------------------
            End If

            'LOCAL EFFECTS ONLY...
            'At this point the above grid will satisfy 'local effects only' people so...
            If g_booLocalEffects Then

                ShowProgress ("Creating data layer for local effects...", strTitle, 0, 13, 13, _
                              g_frmProjectSetup)
                If g_KeepRunning Then

                    strOutConc = GetUniqueName ("locconc", g_strWorkspace, g_FinalOutputGridExt)
                    'Added 7/23/04 to account for clip by selected polys functionality
                    If g_booSelectedPolys Then
                        pPermMassVolumeRaster = _
                            ClipBySelectedPoly (pMassVolumeRaster, g_pSelectedPolyClip, strOutConc)
                    Else
                        pPermMassVolumeRaster = ReturnPermanentRaster (pMassVolumeRaster, strOutConc)
                    End If

                    g_dicMetadata.Add (_strPollName & "Local Effects (mg)", _strPollCoeffMetadata)

                    AddOutputGridLayer (pPermMassVolumeRaster, _strColor, True, _strPollName & " Local Effects (mg)", _
                                        "Pollutant " & _strPollName & " Local", - 1, OutputItems)

                    CalcPollutantConcentration = True

                End If

                CloseDialog()
                Exit Function

            End If

            ShowProgress ("Deriving accumulated pollutant...", strTitle, 0, 13, 3, g_frmProjectSetup)
            If g_KeepRunning Then
                'STEP 3: DERIVE ACCUMULATED POLLUTANT ------------------------------------------------------

                'Use weightedaread8 from geoproc to accum, then rastercalc to multiply this out
                Dim pTauD8Flow As Grid = Nothing

                Dim tauD8calc As New RasterMathCellCalcNulls (AddressOf tauD8CellCalc)
                RasterMath (g_pFlowDirRaster, Nothing, Nothing, Nothing, Nothing, pTauD8Flow, Nothing, False, tauD8calc)
                pTauD8Flow.Header.NodataValue = - 1

                Dim strtmp1 As String = Path.GetTempFileName
                g_TempFilesToDel.Add (strtmp1)
                strtmp1 = strtmp1 + g_TAUDEMGridExt
                g_TempFilesToDel.Add (strtmp1)
                DataManagement.DeleteGrid (strtmp1)
                pTauD8Flow.Save (strtmp1)

                Dim strtmp2 As String = Path.GetTempFileName
                g_TempFilesToDel.Add (strtmp2)
                strtmp2 = strtmp2 + g_TAUDEMGridExt
                g_TempFilesToDel.Add (strtmp2)
                DataManagement.DeleteGrid (strtmp2)
                pMassVolumeRaster.Save (strtmp2)

                Dim strtmpout As String = Path.GetTempFileName
                g_TempFilesToDel.Add (strtmpout)
                strtmpout = strtmpout + "out" + g_TAUDEMGridExt
                g_TempFilesToDel.Add (strtmpout)
                DataManagement.DeleteGrid (strtmpout)

                'Use geoproc weightedAreaD8 after converting the D8 grid to taudem format bgd if needed
                Hydrology.WeightedAreaD8 (strtmp1, strtmp2, "", strtmpout, False, False, _
                                          Environment.ProcessorCount, Nothing)
                'strExpression = "FlowAccumulation([flowdir], [met_run], FLOAT)"

                Dim tmpGrid As New Grid
                tmpGrid.Open (strtmpout)

                Dim multAccumcalc As New RasterMathCellCalc (AddressOf multAccumCellCalc)
                RasterMath (tmpGrid, Nothing, Nothing, Nothing, Nothing, pAccumPollRaster, multAccumcalc)

                pTauD8Flow.Close()
                DataManagement.DeleteGrid (strtmp1)

                'END STEP 3: ------------------------------------------------------------------------------
            End If

            'STEP 3a: Added 7/26: ADD ACCUMULATED POLLUTANT TO GROUP LAYER-----------------------------------
            ShowProgress ("Creating accumlated pollutant layer...", strTitle, 0, 13, 4, g_frmProjectSetup)
            If g_KeepRunning Then
                strAccPoll = GetUniqueName ("accpoll", g_strWorkspace, g_FinalOutputGridExt)
                'Added 7/23/04 to account for clip by selected polys functionality
                If g_booSelectedPolys Then
                    pPermAccPollRaster = ClipBySelectedPoly (pAccumPollRaster, g_pSelectedPolyClip, strAccPoll)
                Else
                    pPermAccPollRaster = ReturnPermanentRaster (pAccumPollRaster, strAccPoll)
                End If

                g_dicMetadata.Add ("Accumulated " & _strPollName & " (kg)", _strPollCoeffMetadata)

                AddOutputGridLayer (pPermAccPollRaster, _strColor, True, "Accumulated " & _strPollName & " (kg)", _
                                    "Pollutant " & _strPollName & " Accum", - 1, OutputItems)
            End If
            'END STEP 3a: ---------------------------------------------------------------------------------

            ShowProgress ("Calculating final concentration...", strTitle, 0, 13, 9, g_frmProjectSetup)
            If g_KeepRunning Then
                Dim AllConCalc As New RasterMathCellCalcNulls (AddressOf AllConCellCalc)
                RasterMath (pMassVolumeRaster, pAccumPollRaster, g_pMetRunoffRaster, g_pRunoffRaster, g_pDEMRaster, _
                            pTotalPollConc0Raster, Nothing, False, AllConCalc)
            End If

            If g_KeepRunning Then
                ShowProgress ("Creating data layer...", strTitle, 0, 13, 11, g_frmProjectSetup)

                strOutConc = GetUniqueName ("conc", g_strWorkspace, g_FinalOutputGridExt)

                If g_booSelectedPolys Then
                    pPermTotalConcRaster = _
                        ClipBySelectedPoly (pTotalPollConc0Raster, g_pSelectedPolyClip, strOutConc)
                Else
                    pPermTotalConcRaster = ReturnPermanentRaster (pTotalPollConc0Raster, strOutConc)
                End If

                g_dicMetadata.Add (_strPollName & " Conc. (mg/L)", _strPollCoeffMetadata)

                AddOutputGridLayer (pPermTotalConcRaster, _strColor, True, _strPollName & " Conc. (mg/L)", _
                                    "Pollutant " & _strPollName & " Conc", - 1, OutputItems)
            End If

            ShowProgress ("Comparing to water quality standard...", strTitle, 0, 13, 13, g_frmProjectSetup)

            If g_KeepRunning Then
                If Not CompareWaterQuality (g_pWaterShedFeatClass, pTotalPollConc0Raster, OutputItems) Then
                    CalcPollutantConcentration = False
                    Exit Function
                End If
            End If

            'if we get to the end
            CalcPollutantConcentration = True

            CloseDialog()

        Catch ex As Exception
            If Err.Number = - 2147217297 Then 'User cancelled operation
                g_KeepRunning = False
                CalcPollutantConcentration = False
                Exit Function
            ElseIf Err.Number = - 2147467259 Then
                MsgBox ( _
                        "ArcMap has reached the maximum number of GRIDs allowed in memory.  " & _
                        "Please exit OpenNSPECT and restart ArcMap.", MsgBoxStyle.Information, _
                        "Maximum GRID Number Encountered")
                g_KeepRunning = False
                CloseDialog()
                CalcPollutantConcentration = False
            Else
                HandleError (ex)
                'False, "CalcPollutantConcentration " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, _ParentHWND)
                g_KeepRunning = False
                CloseDialog()
                CalcPollutantConcentration = False
            End If
        End Try
    End Function

    Private Function CompareWaterQuality (ByRef pWSFeatureClass As Shapefile, _
                                          ByRef pPollutantRaster As Grid, _
                                          ByRef OutputItems As clsXMLOutputItems) As Boolean
        Dim strWQVAlue As Object

        'Get the zone dataset from the first layer in ArcMap
        Dim pMaxRaster As Grid = Nothing
        Dim pConRaster As Grid = Nothing
        Dim pClipWQRaster As Grid = Nothing
        Dim pPermWQRaster As Grid = Nothing
        Dim pWQRasterLayer As Grid = Nothing
        Dim dblConvertValue As Double
        Dim strOutWQ As String
        Dim strExpression As String = ""
        Dim strMetadata As String

        Try

            ' Perform Spatial operation
            'TODO: This seems useless on a singleband thing, otherwise, seems random. so skipping it.
            'pMaxRaster = pLocalOp.LocalStatistics(pPollutantRaster, ESRI.ArcGIS.GeoAnalyst.esriGeoAnalysisStatisticsEnum.esriGeoAnalysisStatsMaximum)

            strWQVAlue = ReturnWQValue (_strPollName, _strWQName)

            _WQValue = (CDbl (strWQVAlue))/1000
            _FlowMax = g_pFlowAccRaster.Maximum

            Dim concalc As New RasterMathCellCalc (AddressOf concompCellCalc)
            RasterMath (pPollutantRaster, g_pFlowAccRaster, Nothing, Nothing, Nothing, pConRaster, concalc)

            strOutWQ = GetUniqueName ("wq", g_strWorkspace, g_FinalOutputGridExt)

            'Clip if selectedpolys
            If g_booSelectedPolys Then
                pPermWQRaster = ClipBySelectedPoly (pConRaster, g_pSelectedPolyClip, strOutWQ)
            Else
                pPermWQRaster = ReturnPermanentRaster (pConRaster, strOutWQ)
            End If

            strMetadata = vbTab & "Water Quality Standard:" & vbNewLine & vbTab & vbTab & "Criteria Name: " & _strWQName & _
                          vbNewLine & vbTab & vbTab & "Standard: " & dblConvertValue & " mg/L"
            g_dicMetadata.Add (_strPollName & " Standard: " & CStr (dblConvertValue) & " mg/L", _
                               _strPollCoeffMetadata & strMetadata)

            AddOutputGridLayer (pPermWQRaster, _strWQName, False, _
                                _strPollName & " Standard: " & CStr (dblConvertValue) & " mg/L", _
                                "Pollutant " & _strPollName & " WQ", - 1, OutputItems)

            CompareWaterQuality = True

        Catch ex As Exception
            If Err.Number = - 2147467259 Then
                MsgBox ( _
                        "ArcMap has reached the maximum number of GRIDs allowed in memory.  " & _
                        "Please exit OpenNSPECT and restart ArcMap.", MsgBoxStyle.Information, _
                        "Maximum GRID Number Encountered")
                CompareWaterQuality = False
                g_KeepRunning = False
                CloseDialog()
            Else
                HandleError (ex)
                'False, "CompareWaterQuality " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, 0)
                CompareWaterQuality = False
                g_KeepRunning = False
                CloseDialog()
            End If
        End Try
    End Function

    Private Function ReturnWQValue (ByRef strPollName As String, ByRef strWQstdName As String) As String
        Dim strPoll As String
        Dim strWQStd As String = ""
        ReturnWQValue = ""
        Try

            strPoll = "Select * from Pollutant where name like '" & strPollName & "'"
            Dim cmdpoll As New DataHelper (strPoll)
            Dim datapoll As OleDbDataReader = cmdpoll.ExecuteReader
            datapoll.Read()
            strWQStd = _
                "SELECT * FROM WQCRITERIA INNER JOIN POLL_WQCRITERIA ON WQCRITERIA.WQCRITID = POLL_WQCRITERIA.WQCRITID " & _
                "WHERE WQCRITERIA.NAME Like '" & strWQstdName & "' AND POLL_WQCRITERIA.POLLID = " & datapoll ("POLLID")
            datapoll.Close()

            Dim cmdWQ As New DataHelper (strWQStd)
            Dim datawq As OleDbDataReader = cmdWQ.ExecuteReader()
            datawq.Read()
            ReturnWQValue = CStr (datawq ("Threshold"))
            datawq.Close()
        Catch ex As Exception
            MsgBox ("Error in ADO pollutant part: " & Err.Number & vbNewLine & Err.Description & vbNewLine & strWQStd)
        End Try
    End Function

#Region "Raster Math"

    Private Function massvolCellCalc (ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, _
                                      ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        Dim tmpval As Single
        'strexpression = pick([pLandSampleRaster], _picks)"
        For i As Integer = 0 To _picks.Length - 1
            If Input1 = i + 1 Then
                tmpval = _picks (i)
                Exit For
            End If
        Next

        'strExpression = "[met_runoff] * [pollmass]"
        Return Input2*tmpval
    End Function

    Private Function multAccumCellCalc (ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, _
                                        ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) _
        As Single
        'strExpression = "(FlowAccumulation([flowdir], [massvolume], FLOAT)) * 1.0e-6"
        Return Input1*0.000001
    End Function

    Private Function AllConCellCalc (ByVal Input1 As Single, ByVal Input1Null As Single, ByVal Input2 As Single, _
                                     ByVal Input2Null As Single, ByVal Input3 As Single, ByVal Input3Null As Single, _
                                     ByVal Input4 As Single, ByVal Input4Null As Single, ByVal Input5 As Single, _
                                     ByVal Input5Null As Single, ByVal OutNull As Single) As Single
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
            Return (Input1 + (Input2/0.000001))/(Input3 + Input4)
        End If

    End Function

    Private Function concompCellCalc (ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, _
                                      ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        'This rather ugly expression was set up to check for meets/exceed water quality standards for
        'only the streams.  It takes the values of flowaccumulation from watershed delineation fame that
        'exceed values of greater than 1%.  Then multiplies the result (all cells representing streams) times
        'the water quality grid.
        'strExpression = "(Con([Max] gt " & CStr(dblConvertValue) & ", 1, 2)) * (con([flowacc] > (" & CStr(modUtil.ReturnRasterMax(g_pFlowAccRaster)) & " * 0.01), 1))"
        'strExpression = "(Con([Max] gt _WQValue, 1, 2)) * (con([flowacc] > (_FlowMax * 0.01), 1))"
        If Input1 > _WQValue Then
            '(con([flowacc] > (" & CStr(modUtil.ReturnRasterMax(g_pFlowAccRaster)) & " * 0.01), 1))
            If Input2 > (_FlowMax*0.01) Then
                Return 1
            Else
                Return OutNull
            End If
        Else
            If Input2 > (_FlowMax*0.01) Then
                Return 2
            Else
                Return OutNull
            End If
        End If
    End Function

#End Region
End Module