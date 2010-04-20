Imports System.Data.OleDb
Module modRunoff
    ' *************************************************************************************
    ' *  Perot Systems Government Services
    ' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
    ' *  modRunoff
    ' *************************************************************************************
    ' *  Description:  Code for all runoff goodies
    ' *
    ' *
    ' *  Called By:  Various
    ' *************************************************************************************

    'Public g_pSoilsRaster As mapwingis.GridDataset 'Soils Raster GRID
    'Public g_SoilsRasterDataset As mapwingis.GridDataset
    Public g_LandCoverRaster As MapWinGIS.Grid 'LandCover GLOBAL
    Public g_pRunoffRaster As MapWinGIS.Grid 'Runoff for use elsewhere
    Public g_pMetRunoffRaster As MapWinGIS.Grid 'Metric Runoff
    Public g_pSCS100Raster As MapWinGIS.Grid 'SCS GRID for use in RUSLE
    Public g_pAbstractRaster As MapWinGIS.Grid 'Abstraction Raster used in MUSLE
    Public g_pPrecipRaster As MapWinGIS.Grid 'Precipitation Raster
    Public g_pCellAreaSqMiRaster As MapWinGIS.Grid 'Square Mile area raster used in MUSLE
    Public g_pRunoffInchRaster As MapWinGIS.Grid 'Runoff Inches used in MUSLE
    Public g_pRunoffAFRaster As MapWinGIS.Grid 'Runoff AF used in MUSLE
    Public g_pRunoffCFRaster As MapWinGIS.Grid 'MUSLE
    Public g_strPrecipFileName As String 'Precipitation File Name
    Public g_strLandCoverParameters As String 'Global string to hold formatted LC params for metadata

    Private _strRunoffMetadata As String 'Metatdata string for runoff
    Private _intPrecipType As Short 'Precip Event Type: 0=Annual; 1=Event
    Private _intRainingDays As Short '# of Raining Days for Annual Precip
    ' Variables used by the Error handler function - DO NOT REMOVE
    Const c_sModuleFileName As String = "modRunoff.bas"
    Private _ParentHWND As Integer ' Set this to get correct parenting of Error handler forms

    Private _picks()() As String

    Public Function CreateRunoffGrid(ByRef strLCFileName As String, ByRef strLCCLassType As String, ByRef cmdPrecip As OleDbCommand, ByRef strSoilsFileName As String) As Boolean
        'This sub serves as a link between frmPrj and the actual calculation of Runoff
        'It establishes the Rasters being used

        'strLCFileName: Path to land class file
        'strLCCLassType: type of landclass
        'rsPrecip: ADO recordset of precip scenario being used
        Dim pRainFallRaster As MapWinGIS.Grid
        Dim pLandCoverRaster As MapWinGIS.Grid
        Dim pSoilsRaster As MapWinGIS.Grid
        Dim strError As Object

        'Get the LandCover Raster
        'if no management scenarios were applied then g_landcoverRaster will be nothing
        If g_LandCoverRaster Is Nothing Then
            If modUtil.RasterExists(strLCFileName) Then
                g_LandCoverRaster = modUtil.ReturnRaster(strLCFileName)
                pLandCoverRaster = g_LandCoverRaster
            Else
                strError = strLCFileName
                MsgBox("Error: The following dataset is missing: " & strError, MsgBoxStyle.Critical, "Missing Data")
                Return False
            End If
        Else
            pLandCoverRaster = g_LandCoverRaster
        End If

        Dim dataPrecip As OleDbDataReader = cmdPrecip.ExecuteReader()
        dataPrecip.Read()

        'Get the Precip Raster
        If modUtil.RasterExists(dataPrecip("PrecipFileName")) Then
            pRainFallRaster = modUtil.ReturnRaster(dataPrecip("PrecipFileName"))

            'If in cm, then convert to a GRID in inches.
            If dataPrecip("PrecipUnits") = 0 Then
                g_pPrecipRaster = ConvertRainGridCMToInches(pRainFallRaster)
            Else 'if already in inches then just use this one.
                g_pPrecipRaster = pRainFallRaster 'Global Precip
            End If

            g_strPrecipFileName = dataPrecip("PrecipFileName")
            _intPrecipType = dataPrecip("Type") 'Set the mod level precip type
            _intRainingDays = dataPrecip("RainingDays") 'Set the mod level rainingdays
        Else
            strError = strLCFileName
            dataPrecip.Close()
            MsgBox("Error: The following dataset is missing: " & strError, MsgBoxStyle.Critical, "Missing Data")
            Return False
        End If
        dataPrecip.Close()

        '------ 1.1.4 Change -----
        'if they select cm as incoming precip units, convert GRID

        'Get the Soils Raster
        If modUtil.RasterExists(strSoilsFileName) Then
            pSoilsRaster = modUtil.ReturnRaster(strSoilsFileName)
        Else
            strError = strSoilsFileName
            MsgBox("Error: The following dataset is missing: " & strError, MsgBoxStyle.Critical, "Missing Data")
            Return False
        End If

        'Now construct the Pick Statement
        Dim strPick() As String = Nothing
        strPick = ConstructPickStatment(strLCCLassType, pLandCoverRaster)

        If strPick Is Nothing Then
            Exit Function
        End If

        'Call the Runoff Calculation using the string and rasters
        If RunoffCalculation(strPick, g_pPrecipRaster, pLandCoverRaster, pSoilsRaster) Then
            CreateRunoffGrid = True
        Else
            CreateRunoffGrid = False
            Exit Function
        End If

        CreateRunoffGrid = True

    End Function

    Private Function ConvertRainGridCMToInches(ByRef pInRaster As MapWinGIS.Grid) As MapWinGIS.Grid

        Dim head As MapWinGIS.GridHeader = pInRaster.Header
        Dim ncol As Integer = head.NumberCols - 1
        Dim nrow As Integer = head.NumberRows - 1
        Dim nodata As Single = head.NodataValue
        Dim rowvals() As Single
        ReDim rowvals(ncol)
        For row As Integer = 0 To nrow
            pInRaster.GetRow(row, rowvals(0))

            For col As Integer = 0 To ncol
                If rowvals(col) <> nodata Then
                    rowvals(col) = rowvals(col) / 2.54
                End If
            Next

            pInRaster.PutRow(row, rowvals(0))
        Next

        Return pInRaster
    End Function

    Private Function ConstructPickStatment(ByRef strLandClass As String, ByRef pLCRaster As MapWinGIS.Grid) As String()
        'Creates the large initial pick statement using the name of the the LandCass [CCAP, for example]
        'and the Land Class Raster.  Returns a string
        ConstructPickStatment = Nothing
        Try
            Dim strRS As String
            Dim strPick(3) As String 'Array of strings that hold 'pick' numbers

            Dim dblMaxValue As Double
            Dim i As Short
            Dim TableExist As Boolean
            Dim FieldIndex As Short
            Dim booValueFound As Boolean

            'STEP 1:  get the records from the database -----------------------------------------------
            strRS = "SELECT LCCLASS.LCClassID, Value, LCCLASS.Name as Name2, LCCLASS.LCTypeID, [CN-A], [CN-B], [CN-C], [CN-D], CoverFactor, W_WL FROM LCCLASS " & "LEFT OUTER JOIN LCTYPE ON LCCLASS.LCTYPEID = LCTYPE.LCTYPEID " & "Where LCTYPE.NAME = '" & strLandClass & "' ORDER BY LCCLASS.VALUE"

            Dim cmdLandClass As New OleDbCommand(strRS, g_DBConn)

            'while here, make metadata
            _strRunoffMetadata = CreateMetadata(cmdLandClass, strLandClass, g_booLocalEffects)
            'End Database stuff

            'STEP 2: Raster Values ---------------------------------------------------------------------
            'Now Get the RASTER values

            'Get the max value
            dblMaxValue = pLCRaster.Maximum


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

            Dim mwTable As New MapWinGIS.Table
            If Not TableExist Then
                MsgBox("No MapWindow-readable raster table was found. To create one using ArcMap 9.3+, add the raster to the default project, right click on its layer and select Open Attribute Table. Now click on the options button in the lower right and select Export. In the export path, navigate to the directory of the grid folder and give the export the name of the raster folder with the .dbf extension. i.e. if you are exporting a raster attribute table from a raster named landcover, export landcover.dbf into the same level directory as the folder.", MsgBoxStyle.Exclamation, "Raster Attribute Table Not Found")

                Exit Function
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


                'STEP 3: Table values vs. Raster Values Count - if not equal bark -------------------------
                Dim dataLandClass As OleDbDataReader

                Dim rowidx As Integer = 0

                'STEP 4: Create the strings
                'Loop through and get all values
                For i = 1 To dblMaxValue
                    
                    If (mwTable.CellValue(FieldIndex, rowidx) = i) Then 'And (pRow.Value(FieldIndex) = rsLandClass!Value) Then
                        dataLandClass = cmdLandClass.ExecuteReader

                        booValueFound = False
                        While dataLandClass.Read()
                            If mwTable.CellValue(FieldIndex, rowidx) = dataLandClass("Value") Then
                                booValueFound = True
                                If strPick(0) = "" Then
                                    strPick(0) = CStr(dataLandClass("CN-A"))
                                    strPick(1) = CStr(dataLandClass("CN-B"))
                                    strPick(2) = CStr(dataLandClass("CN-C"))
                                    strPick(3) = CStr(dataLandClass("CN-D"))
                                Else
                                    strPick(0) = strPick(0) & ", " & CStr(dataLandClass("CN-A"))
                                    strPick(1) = strPick(1) & ", " & CStr(dataLandClass("CN-B"))
                                    strPick(2) = strPick(2) & ", " & CStr(dataLandClass("CN-C"))
                                    strPick(3) = strPick(3) & ", " & CStr(dataLandClass("CN-D"))
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
                            dataLandClass.Close()
                            mwTable.Close()
                            Exit Function
                        End If
                        dataLandClass.Close()

                    Else
                        If strPick(0) = "" Then
                            strPick(0) = strPick(0) & "-9999"
                            strPick(1) = strPick(1) & "-9999"
                            strPick(2) = strPick(2) & "-9999"
                            strPick(3) = strPick(3) & "-9999"
                        Else
                            strPick(0) = strPick(0) & ", 0"
                            strPick(1) = strPick(1) & ", 0"
                            strPick(2) = strPick(2) & ", 0"
                            strPick(3) = strPick(3) & ", 0"
                        End If
                    End If

                Next i

                ConstructPickStatment = strPick

            End If


        Catch ex As Exception
            MsgBox(Err.Number & ": " & Err.Description & " " & "ConstructPickStatemnt")
        End Try
    End Function

    Private Function CreateMetadata(ByRef cmdLandClass As OleDbCommand, ByRef strLandClass As String, ByRef booLocal As Boolean) As String
        CreateMetadata = ""
        Try
            Dim i As Short
            Dim strHeader As String
            Dim strLCHeader As String
            Dim strCoeffValues As String = ""

            If booLocal Then
                strHeader = vbTab & "Input Datasets:" & vbNewLine & vbTab & vbTab & "Hydrologic soils grid: " & g_clsXMLPrjFile.strSoilsHydFileName & vbNewLine & vbTab & vbTab & "Landcover grid: " & g_clsXMLPrjFile.strLCGridFileName & vbNewLine & vbTab & vbTab & "Precipitation grid: " & g_strPrecipFileName & vbNewLine
            Else
                strHeader = vbTab & "Input Datasets:" & vbNewLine & vbTab & vbTab & "Hydrologic soils grid: " & g_clsXMLPrjFile.strSoilsHydFileName & vbNewLine & vbTab & vbTab & "Landcover grid: " & g_clsXMLPrjFile.strLCGridFileName & vbNewLine & vbTab & vbTab & "Precipitation grid: " & g_strPrecipFileName & vbNewLine & vbTab & vbTab & "Flow direction grid: " & g_strFlowDirFilename & vbNewLine
            End If

            strLCHeader = vbTab & "Landcover parameters: " & vbNewLine & vbTab & vbTab & "Landcover type name: " & strLandClass & vbNewLine & vbTab & vbTab & "Coefficients: " & vbNewLine

            Dim dataLandClass As OleDbDataReader = cmdLandClass.ExecuteReader()

            While dataLandClass.Read()
                If i = 0 Then
                    strCoeffValues = vbTab & vbTab & vbTab & dataLandClass("Name2") & ":" & vbNewLine & vbTab & vbTab & vbTab & vbTab & "CN-A: " & CStr(dataLandClass("CN-A")) & vbNewLine & vbTab & vbTab & vbTab & vbTab & "CN-B: " & CStr(dataLandClass("CN-B")) & vbNewLine & vbTab & vbTab & vbTab & vbTab & "CN-C: " & CStr(dataLandClass("CN-C")) & vbNewLine & vbTab & vbTab & vbTab & vbTab & "CN-D: " & CStr(dataLandClass("CN-D")) & vbNewLine & vbTab & vbTab & vbTab & vbTab & "Cover factor: " & dataLandClass("CoverFactor") & vbNewLine
                    i = i + 1
                Else
                    strCoeffValues = strCoeffValues & vbTab & vbTab & vbTab & dataLandClass("Name2") & ":" & vbNewLine & vbTab & vbTab & vbTab & vbTab & "CN-A: " & CStr(dataLandClass("CN-A")) & vbNewLine & vbTab & vbTab & vbTab & vbTab & "CN-B: " & CStr(dataLandClass("CN-B")) & vbNewLine & vbTab & vbTab & vbTab & vbTab & "CN-C: " & CStr(dataLandClass("CN-C")) & vbNewLine & vbTab & vbTab & vbTab & vbTab & "CN-D: " & CStr(dataLandClass("CN-D")) & vbNewLine & vbTab & vbTab & vbTab & vbTab & "Cover factor: " & dataLandClass("CoverFactor") & vbNewLine
                End If
            End While

            CreateMetadata = strHeader & strLCHeader & strCoeffValues
            g_strLandCoverParameters = strLCHeader & strCoeffValues

            dataLandClass.Close()
        Catch ex As Exception
            HandleError(False, "CreateMetadata " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, _ParentHWND)
        End Try
    End Function

    Public Function RunoffCalculation(ByRef strPick As String(), ByRef pInRainRaster As MapWinGIS.Grid, ByRef pInLandCoverRaster As MapWinGIS.Grid, ByRef pInSoilsRaster As MapWinGIS.Grid) As Boolean
        'strPickStatement: our friend the dynamic pick statemnt
        'pInRainRaster: the precip grid
        'pInLandCoverRaster: landcover grid
        'The rasters                             'Associated Steps

        Try
            Dim pLandCoverRaster As MapWinGIS.Grid = Nothing  'STEP 1: LandCover Raster
            Dim pLandSampleRaster As MapWinGIS.Grid = Nothing  'STEP 1: Properly sized landcover
            Dim pHydSoilsRaster As MapWinGIS.Grid = Nothing  'STEP 2: Soils Raster
            Dim pPrecipRaster As MapWinGIS.Grid = Nothing  'STEP 3: Create Precip Grid
            Dim pDailyRainRaster As MapWinGIS.Grid = Nothing  'Rain
            Dim pSCS100Raster As MapWinGIS.Grid = Nothing  'STEP 2: SCS * 100
            Dim pRetentionRaster As MapWinGIS.Grid = Nothing  'STEP 3: Retention
            Dim pAbstractRaster As MapWinGIS.Grid = Nothing
            Dim pRunoffInRaster As MapWinGIS.Grid = Nothing
            Dim pMetRunoffRaster As MapWinGIS.Grid = Nothing
            Dim pMetRunoffNoNullRaster As MapWinGIS.Grid = Nothing  'STEP 6a:  no nulls
            Dim pAccumRunoffRaster As MapWinGIS.Grid = Nothing
            Dim pAccumRunoffRaster1 As MapWinGIS.Grid = Nothing
            Dim pCellAreaSMRaster As MapWinGIS.Grid = Nothing
            Dim pCellAreaSqMileRaster As MapWinGIS.Grid = Nothing  'STEP 5a2: Square miles
            Dim pCellAreaSIRaster As MapWinGIS.Grid = Nothing
            Dim pCellAreaAcreRaster As MapWinGIS.Grid = Nothing
            Dim pCellAreaCIRaster As MapWinGIS.Grid = Nothing
            Dim pCellAreaCFRaster As MapWinGIS.Grid = Nothing
            Dim pCellAreaAFRaster As MapWinGIS.Grid = Nothing
            Dim pPermAccumRunoffRaster As MapWinGIS.Grid = Nothing
            Dim pPermAccumLocRunoffRaster As MapWinGIS.Grid = Nothing


            'String to hold calculations
            Dim strExpression As String = ""
            Dim strOutAccum As String
            Const strTitle As String = "Processing Runoff Calculation..."

           
            'Assign rasters
            pHydSoilsRaster = pInSoilsRaster
            pLandCoverRaster = pInLandCoverRaster
            pPrecipRaster = pInRainRaster

            modProgDialog.ProgDialog("Checking landcover cell size...", strTitle, 0, 10, 1, 0)

            If modProgDialog.g_boolCancel Then
                'Step 1: ----------------------------------------------------------------------------
                'Make sure LandCover is in the same cellsize as the global environment

                If pLandCoverRaster.Header.dX <> g_dblCellSize Then
                    'TODO: Resample land cover raster to g_dblCellSize
                    MsgBox("The Landcover raster is not the same cellsize as the other project data, please resample this raster to a cellsize of " + g_dblCellSize.ToString, MsgBoxStyle.Exclamation, "Landcover Cell Size Mismatch")
                    Exit Function
                Else
                    pLandSampleRaster = pLandCoverRaster
                End If
            Else
                Exit Function
            End If

            modProgDialog.ProgDialog("Creating precipitation raster...", strTitle, 0, 10, 1, 0)

            If modProgDialog.g_boolCancel Then
                'Step 2a: ----------------------------------------------------------------------------
                'Make the daily rain raster
                pDailyRainRaster = pPrecipRaster
            Else
                Exit Function
            End If
            'END STEP 2a: -------------------------------------------------------------------------


            'STEP 2: ------------------------------------------------------------------------------
            modProgDialog.ProgDialog("Calculating maximum potential retention...", strTitle, 0, 10, 3, 0)

            If modProgDialog.g_boolCancel Then
                'Calculate maxiumum potential retention

                Dim picksLength As Integer = strPick(0).Split(",").Length
                ReDim _picks(strPick.Length - 1)
                For i As Integer = 0 To strPick.Length - 1
                    ReDim _picks(i)(picksLength)
                    _picks(i) = strPick(i).Split(",")
                Next
                Dim sc100calc As New RasterMathCellCalc(AddressOf SC100CellCalc)
                pSCS100Raster = Nothing
                RasterMath(pHydSoilsRaster, pLandSampleRaster, Nothing, Nothing, Nothing, pSCS100Raster, sc100calc)
                g_pSCS100Raster = pSCS100Raster
            Else
                Exit Function
            End If
            'END STEP 2: ---------------------------------------------------------------------------


            'STEP 2A: ------------------------------------------------------------------------------
            modProgDialog.ProgDialog("Calculating maximum potential retention...", strTitle, 0, 10, 4, 0)

            If modProgDialog.g_boolCancel Then
                Dim retentioncalc As New RasterMathCellCalc(AddressOf RetentionCellCalc)
                pRetentionRaster = Nothing
                RasterMath(pSCS100Raster, Nothing, Nothing, Nothing, Nothing, pRetentionRaster, retentioncalc)
            Else
                Exit Function
            End If
            'END STEP 2A: --------------------------------------------------------------------------


            'STEP 3: -------------------------------------------------------------------------------
            modProgDialog.ProgDialog("Calculating initial abstraction...", strTitle, 0, 10, 5, 0)

            If modProgDialog.g_boolCancel Then
                'Calculate initial abstraction
                Dim abstractcalc As New RasterMathCellCalc(AddressOf AbstractCellCalc)
                pAbstractRaster = Nothing
                RasterMath(pRetentionRaster, Nothing, Nothing, Nothing, Nothing, pAbstractRaster, abstractcalc)
                g_pAbstractRaster = pAbstractRaster
            Else
                Exit Function
            End If
            'END STEP 3: ---------------------------------------------------------------------------

            'STEP 4: -------------------------------------------------------------------------------
            modProgDialog.ProgDialog("Calculating runoff...", strTitle, 0, 10, 6, 0)

            If modProgDialog.g_boolCancel Then
                'Calculate Runoff
                Dim runoffcalc As New RasterMathCellCalc(AddressOf RunoffCellCalc)
                pRunoffInRaster = Nothing
                RasterMath(pDailyRainRaster, pAbstractRaster, pRetentionRaster, Nothing, Nothing, pRunoffInRaster, runoffcalc)

                g_pRunoffInchRaster = pRunoffInRaster
            Else
                Exit Function
            End If
            'END STEP 4b: -----------------------------------------------------------------------------



            modProgDialog.ProgDialog("Calculating maximum potential retention...", strTitle, 0, 10, 7, 0)

            If modProgDialog.g_boolCancel Then
                'STEP 5a: ---------------------------------------------------------------------------------
                ''Convert Runoff
                ''Square Meters
                Dim cellSMcalc As New RasterMathCellCalc(AddressOf CellSMCellCalc)
                pCellAreaSMRaster = Nothing
                RasterMath(g_pDEMRaster, Nothing, Nothing, Nothing, Nothing, pCellAreaSMRaster, cellSMcalc)
                'END STEP 5a ------------------------------------------------------------------------------


                'STEP 5a2 sqmiles
                Dim cellSMicalc As New RasterMathCellCalc(AddressOf CellSMiCellCalc)
                pCellAreaSqMileRaster = Nothing
                RasterMath(pCellAreaSMRaster, Nothing, Nothing, Nothing, Nothing, pCellAreaSqMileRaster, cellSMicalc)

                g_pCellAreaSqMiRaster = pCellAreaSqMileRaster

                'END STEP 5a1: -----------------------------------------------------------------------------



                'STEP 5c: ----------------------------------------------------------------------------------
                'Square Inches
                Dim cellSIcalc As New RasterMathCellCalc(AddressOf CellSICellCalc)
                pCellAreaSIRaster = Nothing
                RasterMath(pCellAreaSMRaster, Nothing, Nothing, Nothing, Nothing, pCellAreaSIRaster, cellSIcalc)
                'END STEP 5c--------------------------------------------------------------------------------

                'STEP 5d: ----------------------------------------------------------------------------------
                'Acres
                Dim cellAreaAcrecalc As New RasterMathCellCalc(AddressOf CellAreaAcreCellCalc)
                pCellAreaAcreRaster = Nothing
                RasterMath(pCellAreaSMRaster, Nothing, Nothing, Nothing, Nothing, pCellAreaAcreRaster, cellAreaAcrecalc)
                'END STEP 5d--------------------------------------------------------------------------------


                'STEP 5e: ----------------------------------------------------------------------------------
                'Cubic Inches
                Dim cellAreaCIcalc As New RasterMathCellCalc(AddressOf CellAreaCICellCalc)
                pCellAreaCIRaster = Nothing
                RasterMath(pCellAreaSIRaster, pRunoffInRaster, Nothing, Nothing, Nothing, pCellAreaCIRaster, cellAreaCIcalc)
                'END STEP 5e -------------------------------------------------------------------------------


                'STEP 5f: -----------------------------------------------------------------------------------
                'Cubic feet
                Dim cellAreaCFcalc As New RasterMathCellCalc(AddressOf cellAreaCFCellCalc)
                pCellAreaCFRaster = Nothing
                RasterMath(pCellAreaCIRaster, Nothing, Nothing, Nothing, Nothing, pCellAreaCFRaster, cellAreaCFcalc)
                g_pRunoffCFRaster = pCellAreaCFRaster
                'END STEP 5f: --------------------------------------------------------------------------------


                'STEP 5g: ------------------------------------------------------------------------------------
                'Runoff acre-feet
                Dim cellAreaAFcalc As New RasterMathCellCalc(AddressOf cellAreaAFCellCalc)
                pCellAreaAFRaster = Nothing
                RasterMath(pCellAreaCFRaster, Nothing, Nothing, Nothing, Nothing, pCellAreaAFRaster, cellAreaAFcalc)

                g_pRunoffAFRaster = pCellAreaAFRaster
                'END STEP 5f: --------------------------------------------------------------------------------
            Else
                Exit Function
            End If


            'STEP 6: -------------------------------------------------------------------------------------
            modProgDialog.ProgDialog("Converting to correct units...", strTitle, 0, 10, 8, 0)

            If modProgDialog.g_boolCancel Then
                'Convert cubic inches to liters
                Dim metRunoffcalc As New RasterMathCellCalc(AddressOf metRunoffCellCalc)
                pMetRunoffRaster = Nothing
                RasterMath(pCellAreaCIRaster, Nothing, Nothing, Nothing, Nothing, pMetRunoffRaster, metRunoffcalc)

                g_pMetRunoffRaster = pMetRunoffRaster
            Else
                Exit Function
            End If
            'END STEP 6: ---------------------------------------------------------------------------------

            'STEP 6a: -------------------------------------------------------------------------------------
            'Eliminate nulls: Added 1/28/07
            Dim metRunoffNoNullcalc As New RasterMathCellCalcNulls(AddressOf metRunoffNoNullCellCalc)
            pMetRunoffNoNullRaster = Nothing
            RasterMath(pMetRunoffRaster, Nothing, Nothing, Nothing, Nothing, pMetRunoffNoNullRaster, Nothing, False, metRunoffNoNullcalc)

            g_pMetRunoffRaster = pMetRunoffNoNullRaster


            If g_booLocalEffects Then
                modProgDialog.ProgDialog("Creating data layer for local effects...", strTitle, 0, 10, 10, 0)
                If modProgDialog.g_boolCancel Then

                    'STEP 12: Local Effects -------------------------------------------------
                    strOutAccum = modUtil.GetUniqueName("locaccum", g_strWorkspace, ".bgd")
                    'Added 7/23/04 to account for clip by selected polys functionality
                    If g_booSelectedPolys Then
                        pPermAccumLocRunoffRaster = modUtil.ClipBySelectedPoly(pMetRunoffNoNullRaster, g_pSelectedPolyClip, strOutAccum)
                    Else
                        pPermAccumLocRunoffRaster = modUtil.ReturnPermanentRaster(pMetRunoffNoNullRaster, strOutAccum)
                    End If

                    g_dicMetadata.Add("Runoff Local Effects (L)", _strRunoffMetadata)

                    Dim cs As MapWinGIS.GridColorScheme = ReturnRasterStretchColorRampCS(pPermAccumLocRunoffRaster, "Blue")
                    Dim lyr As MapWindow.Interfaces.Layer = g_MapWin.Layers.Add(pPermAccumLocRunoffRaster, cs, "Runoff Local Effects (L)")
                    lyr.Visible = False
                    lyr.MoveTo(0, g_pGroupLayer)

                    RunoffCalculation = True
                    modProgDialog.KillDialog()
                    Exit Function
                End If
            End If

            modProgDialog.ProgDialog("Creating flow accumulation...", strTitle, 0, 10, 9, 0)
            If modProgDialog.g_boolCancel Then
                'STEP 7: ------------------------------------------------------------------------------------
                'Derive Accumulated Runoff
                'TODO: do flow accumulation of g_pFlowDirRaster d8 and pMetRunoffNoNullRaster and assign to pAccumRunoffRaster1
                'TODO: Use geoproc weightedAreaD8 after converting the D8 grid to taudem format bgd if needed

                'With pMapAlgebraOp
                '    .BindRaster(g_pFlowDirRaster, "flowdir")
                '    .BindRaster(pMetRunoffNoNullRaster, "met_run")
                'End With

                'strExpression = "FlowAccumulation([flowdir], [met_run], FLOAT)"

                'pAccumRunoffRaster1 = pMapAlgebraOp.Execute(strExpression)



                'END STEP 7: ----------------------------------------------------------------------------------
            Else
                Exit Function
            End If


            modProgDialog.ProgDialog("Creating flow accumulation...", strTitle, 0, 10, 9, 0)
            If modProgDialog.g_boolCancel Then
                'STEP 7: ------------------------------------------------------------------------------------
                'Derive the real Accumulated Runoff
                Dim accumcalc As New RasterMathCellCalc(AddressOf accumCellCalc)
                'TODO: Wait til accumulation is actually working to uncomment this
                'RasterMath(pAccumRunoffRaster1, pMetRunoffNoNullRaster, Nothing, pAccumRunoffRaster, accumcalc)

                'END STEP 7: ----------------------------------------------------------------------------------
            Else
                Exit Function
            End If

            'TODO: wait for the accum to be finished
            'Add this then map as our runoff grid
            'modProgDialog.ProgDialog("Creating Runoff Layer...", strTitle, 0, 10, 10, 0)
            'If modProgDialog.g_boolCancel Then
            '    'Get a unique name for accumulation GRID
            '    strOutAccum = modUtil.GetUniqueName("runoff", g_strWorkspace, ".tif")

            '    'Clip to selected polys if chosen
            '    If g_booSelectedPolys Then
            '        pPermAccumRunoffRaster = modUtil.ClipBySelectedPoly(pAccumRunoffRaster, g_pSelectedPolyClip, strOutAccum)
            '    Else
            '        pPermAccumRunoffRaster = modUtil.ReturnPermanentRaster(pAccumRunoffRaster, strOutAccum)
            '    End If

            '    Dim cs As MapWinGIS.GridColorScheme = ReturnRasterStretchColorRampCS(pPermAccumRunoffRaster, "Blue")
            '    Dim lyr As MapWindow.Interfaces.Layer = g_MapWin.Layers.Add(pPermAccumRunoffRaster, cs, "Accumulated Runoff (L)")
            '    lyr.Visible = False
            '    lyr.MoveTo(0, g_pGroupLayer)

            '    g_dicMetadata.Add("Accumulated Runoff (L)", _strRunoffMetadata)

            '    'Global Runoff
            '    g_pRunoffRaster = pPermAccumRunoffRaster
            'Else
            '    Exit Function
            'End If

            RunoffCalculation = True

            modProgDialog.KillDialog()

        Catch ex As Exception
            If Err.Number = -2147217297 Then 'User cancelled operation
                modProgDialog.g_boolCancel = False
                RunoffCalculation = False
            ElseIf Err.Number = -2147467259 Then
                MsgBox("ArcMap has reached the maximum number of GRIDs allowed in memory.  " & "Please exit N-SPECT and restart ArcMap.", MsgBoxStyle.Information, "Maximum GRID Number Encountered")
                modProgDialog.g_boolCancel = False
                modProgDialog.KillDialog()
                RunoffCalculation = False
            Else
                MsgBox("Error: " & Err.Number & " on RunoffCalculation")
                modProgDialog.g_boolCancel = False
                modProgDialog.KillDialog()
                RunoffCalculation = False
            End If
        End Try
    End Function

#Region "Raster Math"
    Private Function SC100CellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'pSCS100Raster = 100 * con([pHydSoilsRaster] == 1, pick([pLandSampleRaster], " & strPick(0) & "), con([pHydSoilsRaster] == 2, pick([pLandSampleRaster], " & strPick(1) & "), con([pHydSoilsRaster] == 3, pick([pLandSampleRaster], " & strPick(2) & "), con([pHydSoilsRaster] == 4, pick([pLandSampleRaster], " & strPick(3) & ")))))

        If Input1 = 1 Then
            For i As Integer = 0 To _picks(0).Length - 1
                If Input2 = i + 1 Then
                    Return _picks(0)(i)
                End If
            Next
        ElseIf Input1 = 2 Then
            For i As Integer = 0 To _picks(1).Length - 1
                If Input2 = i + 1 Then
                    Return _picks(1)(i)
                End If
            Next
        ElseIf Input1 = 3 Then
            For i As Integer = 0 To _picks(2).Length - 1
                If Input2 = i + 1 Then
                    Return _picks(2)(i)
                End If
            Next
        ElseIf Input1 = 4 Then
            For i As Integer = 0 To _picks(3).Length - 1
                If Input2 = i + 1 Then
                    Return _picks(3)(i)
                End If
            Next
        End If
    End Function

    Private Function RetentionCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'If m_intPrecipType = 0 Then
        '    pRetentionRaster = "0.2 * Float((1000. / [pSCS100Raster]) - 10) * " & m_intRainingDays
        'Else
        '    pRetentionRaster = "0.2 * Float((1000. / [pSCS100Raster]) - 10)"
        'End If
        If _intPrecipType = 0 Then
            Return ((1000.0 / Input1) - 10) * _intRainingDays
        Else
            Return ((1000.0 / Input1) - 10) * _intRainingDays
        End If
    End Function

    Private Function AbstractCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'strExpression = "(0.2 * [retention])"
        Return 0.2 * Input1
    End Function

    Private Function RunoffCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'pRunoffInRaster = "Con([Con(([pDailyRainRaster] - [pAbstractRaster]) > 0, ([pDailyRainRaster] - [pAbstractRaster]), 0.0)] > 0, (Pow([Con(([pDailyRainRaster] - [pAbstractRaster]) > 0, ([pDailyRainRaster] - [pAbstractRaster]), 0.0)], 2) / ([Con(([pDailyRainRaster] - [pAbstractRaster]) > 0, ([pDailyRainRaster] - [pAbstractRaster]), 0.0)] + [pRetentionRaster])), 0)"
        Dim tmpVal As Single = Input1 - Input2
        If tmpVal > 0 Then
            Return Math.Pow(tmpVal, 2) / (tmpVal + input3)
        Else
            Return 0
        End If
    End Function

    Private Function CellSMCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'pCellAreaSMRaster = "Pow(Con([g_pDEMRaster] >= 0, 1, 0) * " & g_dblCellSize & ", 2)"
        If Input1 > 0 Then
            Return Math.Pow(g_dblCellSize, 2)
        Else
            Input1 = 0
        End If
    End Function

    Private Function CellSMiCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'pCellAreaSqMileRaster = "pCellAreaSMRaster * 0.000001 * 0.386102"
        Return Input1 * 0.000001 * 0.386102
    End Function

    Private Function CellSICellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'pCellAreaSIRaster = "[pCellAreaSMRaster] * 10.76 * 144"
        Return Input1 * 10.76 * 144
    End Function

    Private Function CellAreaAcreCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'pCellAreaAcreRaster = "[pCellAreaSMRaster] * 0.000247104369"
        Return Input1 * 0.000247104369
    End Function

    Private Function cellAreaCICellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'pCellAreaCIRaster = "([pCellAreaSIRaster] * [pRunoffInRaster])"
        Return Input1 * Input2
    End Function

    Private Function cellAreaCFCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'pCellAreaCFRaster = "([pCellAreaCIRaster] * 0.0005787)"
        Return Input1 * 0.0005787
    End Function

    Private Function cellAreaAFCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'pCellAreaAFRaster = "([pCellAreaCFRaster] * 0.000022957)"
        Return Input1 * 0.000022957
    End Function

    Private Function metRunoffCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        Return Input1 * 0.016387064
    End Function

    Private Function metRunoffNoNullCellCalc(ByVal Input1 As Single, ByVal Input1Null As Single, ByVal Input2 As Single, ByVal Input2Null As Single, ByVal Input3 As Single, ByVal Input3Null As Single, ByVal Input4 As Single, ByVal Input4Null As Single, ByVal Input5 As Single, ByVal Input5Null As Single) As Single
        'strExpression = "Con(IsNull([runoffgrid]),0,[runoffgrid])"
        If Input1 <> Input1Null Then
            Return Input1
        Else
            Return 0
        End If
    End Function

    Private Function accumCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'strExpression = "[accum] + [met_run]"
        Return Input1 + Input2
    End Function


#End Region

End Module