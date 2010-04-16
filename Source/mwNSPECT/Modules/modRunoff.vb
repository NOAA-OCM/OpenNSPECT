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

    Private m_strRunoffMetadata As String 'Metatdata string for runoff
    Private m_intPrecipType As Short 'Precip Event Type: 0=Annual; 1=Event
    Private m_intRainingDays As Short '# of Raining Days for Annual Precip
    ' Variables used by the Error handler function - DO NOT REMOVE
    Const c_sModuleFileName As String = "modRunoff.bas"
    Private m_ParentHWND As Integer ' Set this to get correct parenting of Error handler forms



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
            m_intPrecipType = dataPrecip("Type") 'Set the mod level precip type
            m_intRainingDays = dataPrecip("RainingDays") 'Set the mod level rainingdays
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
            Dim strCurveCalc As String 'Full String

            Dim dblMaxValue As Double
            Dim i As Short
            Dim TableExist As Boolean
            Dim FieldIndex As Short
            Dim booValueFound As Boolean

            'STEP 1:  get the records from the database -----------------------------------------------
            strRS = "SELECT LCCLASS.LCClassID, Value, LCCLASS.Name as Name2, LCCLASS.LCTypeID, [CN-A], [CN-B], [CN-C], [CN-D], CoverFactor, W_WL FROM LCCLASS " & "LEFT OUTER JOIN LCTYPE ON LCCLASS.LCTYPEID = LCTYPE.LCTYPEID " & "Where LCTYPE.NAME = '" & strLandClass & "' ORDER BY LCCLASS.VALUE"

            Dim cmdLandClass As New OleDbCommand(strRS, g_DBConn)



            'while here, make metadata
            m_strRunoffMetadata = CreateMetadata(cmdLandClass, strLandClass, g_booLocalEffects)
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
                    'Check first value, usually they won't have a [1], but if there check against database,
                    'else move on
                    If i = 1 Then
                        If (mwTable.CellValue(FieldIndex, rowidx) = i) Then 'And (pRow.Value(FieldIndex) = rsLandClass!Value) Then
                            dataLandClass = cmdLandClass.ExecuteReader

                            booValueFound = False
                            While dataLandClass.Read()
                                If mwTable.CellValue(FieldIndex, rowidx) = dataLandClass("Value") Then
                                    booValueFound = True
                                    strPick(0) = CStr(dataLandClass("CN-A"))
                                    strPick(1) = CStr(dataLandClass("CN-B"))
                                    strPick(2) = CStr(dataLandClass("CN-C"))
                                    strPick(3) = CStr(dataLandClass("CN-D"))
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
                            strPick(0) = "-9999"
                            strPick(1) = "-9999"
                            strPick(2) = "-9999"
                            strPick(3) = "-9999"
                        End If
                    Else
                        If (mwTable.CellValue(FieldIndex, rowidx) = i) Then 'And (pRow.Value(FieldIndex) = rsLandClass!Value) Then
                            dataLandClass = cmdLandClass.ExecuteReader

                            booValueFound = False
                            While dataLandClass.Read()
                                If mwTable.CellValue(FieldIndex, rowidx) = dataLandClass("Value") Then
                                    booValueFound = True
                                    strPick(0) = strPick(0) & ", " & CStr(dataLandClass("CN-A"))
                                    strPick(1) = strPick(1) & ", " & CStr(dataLandClass("CN-B"))
                                    strPick(2) = strPick(2) & ", " & CStr(dataLandClass("CN-C"))
                                    strPick(3) = strPick(3) & ", " & CStr(dataLandClass("CN-D"))
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
            HandleError(False, "CreateMetadata " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Function


    Public Function RunoffCalculation(ByRef strPick As String(), ByRef pInRainRaster As MapWinGIS.Grid, ByRef pInLandCoverRaster As MapWinGIS.Grid, ByRef pInSoilsRaster As MapWinGIS.Grid) As Boolean
        'strPickStatement: our friend the dynamic pick statemnt
        'pInRainRaster: the precip grid
        'pInLandCoverRaster: landcover grid
        'The rasters                             'Associated Steps

        Try
            Dim pLandCoverRaster As MapWinGIS.Grid 'STEP 1: LandCover Raster
            Dim pLandSampleRaster As MapWinGIS.Grid 'STEP 1: Properly sized landcover
            Dim pHydSoilsRaster As MapWinGIS.Grid 'STEP 2: Soils Raster
            Dim pPrecipRaster As MapWinGIS.Grid 'STEP 3: Create Precip Grid
            Dim pSCSRaster As MapWinGIS.Grid 'STEP 1: Create Curve Number GRID
            Dim pDailyRainRaster As MapWinGIS.Grid 'Rain
            Dim pSCS100Raster As MapWinGIS.Grid 'STEP 2: SCS * 100
            Dim pRetentionRaster As MapWinGIS.Grid 'STEP 3: Retention
            Dim pAbstractRaster As MapWinGIS.Grid
            Dim pTemp1RunoffRaster As MapWinGIS.Grid
            Dim pTemp2RunoffRaster As MapWinGIS.Grid
            Dim pRunoffInRaster As MapWinGIS.Grid
            Dim pMaskRaster As MapWinGIS.Grid
            Dim pMetRunoffRaster As MapWinGIS.Grid
            Dim pMetRunoffNoNullRaster As MapWinGIS.Grid 'STEP 6a:  no nulls
            Dim pAccumRunoffRaster As MapWinGIS.Grid
            Dim pAccumRunoffRaster1 As MapWinGIS.Grid
            Dim pCellAreaSMRaster As MapWinGIS.Grid
            Dim pCellAreaSqKMRaster As MapWinGIS.Grid 'STEP 5a1: Square km
            Dim pCellAreaSqMileRaster As MapWinGIS.Grid 'STEP 5a2: Square miles
            Dim pCellAreaSFRaster As MapWinGIS.Grid
            Dim pCellAreaSIRaster As MapWinGIS.Grid
            Dim pCellAreaAcreRaster As MapWinGIS.Grid
            Dim pCellAreaCIRaster As MapWinGIS.Grid
            Dim pCellAreaCFRaster As MapWinGIS.Grid
            Dim pCellAreaAFRaster As MapWinGIS.Grid
            Dim pPermAccumRunoffRaster As MapWinGIS.Grid
            Dim pRunoffRasterLayer As MapWindow.Interfaces.Layer 'Rasterlayer of runoff
            Dim pPermAccumLocRunoffRaster As MapWinGIS.Grid
            Dim pLocRunoffRasterLayer As MapWindow.Interfaces.Layer 'Rasterlayer of local effects runoff


            'String to hold calculations
            Dim strExpression As String
            Dim strOutAccum As String
            Const strTitle As String = "Processing Runoff Calculation..."

            'TEMP CODE
            Dim strTemp As String
            Dim pTempRaster As MapWinGIS.Grid

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

            Dim head1, head2 As MapWinGIS.GridHeader
            Dim ncol As Integer
            Dim nrow As Integer
            Dim nodata1, nodata2 As Single
            Dim rowvals1(), rowvals2(), rowvalsout() As Single

            modProgDialog.ProgDialog("Calculating maximum potential retention...", strTitle, 0, 10, 3, 0)

            If modProgDialog.g_boolCancel Then
                'STEP 2: ------------------------------------------------------------------------------
                'Calculate maxiumum potential retention

                'pSCS100Raster = 100 * con([pHydSoilsRaster] == 1, pick([pLandSampleRaster], " & strPick(0) & "), con([pHydSoilsRaster] == 2, pick([pLandSampleRaster], " & strPick(1) & "), con([pHydSoilsRaster] == 3, pick([pLandSampleRaster], " & strPick(2) & "), con([pHydSoilsRaster] == 4, pick([pLandSampleRaster], " & strPick(3) & ")))))
                Dim picksLength As Integer = strPick(0).Split(",").Length
                Dim picks(strPick.Length - 1)() As String

                For i As Integer = 0 To strPick.Length - 1
                    ReDim picks(i)(picksLength)
                    picks(i) = strPick(i).Split(",")
                Next


                head1 = pHydSoilsRaster.Header
                head2 = pLandSampleRaster.Header
                Dim head100 As New MapWinGIS.GridHeader
                head100.CopyFrom(head1)
                pSCS100Raster = New MapWinGIS.Grid()
                pSCS100Raster.CreateNew("", head100, MapWinGIS.GridDataType.DoubleDataType, head100.NodataValue)
                ncol = head1.NumberCols - 1
                nrow = head1.NumberRows - 1
                nodata1 = head1.NodataValue
                nodata2 = head2.NodataValue
                ReDim rowvals1(ncol)
                ReDim rowvals2(ncol)
                ReDim rowvalsout(ncol)

                For row As Integer = 0 To nrow
                    pHydSoilsRaster.GetRow(row, rowvals1(0))
                    pLandSampleRaster.GetRow(row, rowvals2(0))

                    For col As Integer = 0 To ncol
                        If rowvals1(col) <> nodata1 And rowvals2(col) <> nodata2 Then
                            If rowvals1(col) = 1 Then
                                For i As Integer = 0 To picks.Length - 1
                                    If rowvals2(col) = i + 1 Then
                                        rowvalsout(col) = picks(0)(i)
                                        Exit For
                                    End If
                                Next
                            ElseIf rowvals1(col) = 2 Then
                                For i As Integer = 0 To picks.Length - 1
                                    If rowvals2(col) = i + 1 Then
                                        rowvalsout(col) = picks(1)(i)
                                        Exit For
                                    End If
                                Next
                            ElseIf rowvals1(col) = 3 Then
                                For i As Integer = 0 To picks.Length - 1
                                    If rowvals2(col) = i + 1 Then
                                        rowvalsout(col) = picks(2)(i)
                                        Exit For
                                    End If
                                Next
                            ElseIf rowvals1(col) = 4 Then
                                For i As Integer = 0 To picks.Length - 1
                                    If rowvals2(col) = i + 1 Then
                                        rowvalsout(col) = picks(3)(i)
                                        Exit For
                                    End If
                                Next
                            End If
                        Else
                            rowvalsout(col) = nodata1
                        End If
                    Next

                    pSCS100Raster.PutRow(row, rowvalsout(0))
                Next


                'pSCS100Raster = pMapAlgebraOp.Execute(strExpression)


                g_pSCS100Raster = pSCS100Raster

            Else
                Exit Function
            End If
            'END STEP 2: ---------------------------------------------------------------------------


            'modProgDialog.ProgDialog("Calculating maximum potential retention...", strTitle, 0, 10, 4, 0)

            'If modProgDialog.g_boolCancel Then
            '    'STEP 2A: ------------------------------------------------------------------------------
            '    pMapAlgebraOp.BindRaster(pSCS100Raster, "scsgrid100")

            '    'Added 7/23 to account for rainman days
            '    If m_intPrecipType = 0 Then
            '        strExpression = "Float((1000. / [scsgrid100]) - 10) * " & m_intRainingDays
            '    Else
            '        strExpression = "Float((1000. / [scsgrid100]) - 10)"
            '    End If

            '    pRetentionRaster = pMapAlgebraOp.Execute(strExpression)
            'Else
            '    Exit Function
            'End If
            ''END STEP 2A: --------------------------------------------------------------------------


            'modProgDialog.ProgDialog("Calculating initial abstraction...", strTitle, 0, 10, 5, 0)

            'If modProgDialog.g_boolCancel Then

            '    'STEP 3: -------------------------------------------------------------------------------
            '    'Calculate initial abstraction
            '    pMapAlgebraOp.BindRaster(pRetentionRaster, "retention")

            '    strExpression = "(0.2 * [retention])"

            '    pAbstractRaster = pMapAlgebraOp.Execute(strExpression)
            '    g_pAbstractRaster = pAbstractRaster
            'Else
            '    Exit Function
            'End If
            ''END STEP 3: ---------------------------------------------------------------------------

            'modProgDialog.ProgDialog("Calculating runoff...", strTitle, 0, 10, 6, 0)

            'If modProgDialog.g_boolCancel Then
            '    'STEP 4: -------------------------------------------------------------------------------
            '    'Calculate Runoff
            '    With pMapAlgebraOp
            '        .BindRaster(pDailyRainRaster, "rainsix")
            '        .BindRaster(pAbstractRaster, "abstract")
            '    End With

            '    strExpression = "Con(([rainsix] - [abstract]) > 0, ([rainsix] - [abstract]), 0.0)"

            '    pTemp1RunoffRaster = pMapAlgebraOp.Execute(strExpression)

            '    'END STEP 4: ----------------------------------------------------------------------------



            '    'STEP 4a: --------------------------------------------------------------------------------
            '    'Temp2 calc in runoff
            '    pMapAlgebraOp.BindRaster(pTemp1RunoffRaster, "temp1")

            '    strExpression = "Pow([temp1], 2)"

            '    pTemp2RunoffRaster = pMapAlgebraOp.Execute(strExpression)
            '    'END STEP 4a -----------------------------------------------------------------------------



            '    'STEP 4b: --------------------------------------------------------------------------------
            '    With pMapAlgebraOp
            '        .BindRaster(pTemp1RunoffRaster, "temp1")
            '        .BindRaster(pTemp2RunoffRaster, "temp2")
            '        .BindRaster(pRetentionRaster, "retention")
            '    End With

            '    strExpression = "Con([temp1] > 0, ([temp2] / ([temp1] + [retention])), 0)"

            '    pRunoffInRaster = pMapAlgebraOp.Execute(strExpression)
            '    g_pRunoffInchRaster = pRunoffInRaster

            'Else
            '    Exit Function
            'End If
            ''END STEP 4b: -----------------------------------------------------------------------------

            'modProgDialog.ProgDialog("Calculating maximum potential retention...", strTitle, 0, 10, 7, 0)

            'If modProgDialog.g_boolCancel Then
            '    'STEP 5a: ---------------------------------------------------------------------------------
            '    'Create mask
            '    pMapAlgebraOp.BindRaster(g_pDEMRaster, "DEM")

            '    strExpression = "Con([DEM] >= 0, 1, 0)"

            '    pMaskRaster = pMapAlgebraOp.Execute(strExpression)


            '    'Convert Runoff
            '    'Square Meters
            '    pMapAlgebraOp.BindRaster(pMaskRaster, "cellmask")

            '    strExpression = "Pow([cellmask] * " & g_dblCellSize & ", 2)"

            '    pCellAreaSMRaster = pMapAlgebraOp.Execute(strExpression)

            '    'END STEP 5a ------------------------------------------------------------------------------


            '    'STEP 5a1: --------------------------------------------------------------------------------
            '    'Square Kilometers - Added for MUSLE use
            '    pMapAlgebraOp.BindRaster(pCellAreaSMRaster, "cellarea_sqm")

            '    strExpression = "[cellarea_sqm] * 0.000001"

            '    pCellAreaSqKMRaster = pMapAlgebraOp.Execute(strExpression)

            '    'END STEP 5a1: -----------------------------------------------------------------------------


            '    'STEP 5a2
            '    'Square Kilometers - Added for MUSLE use ---------------------------------------------------
            '    pMapAlgebraOp.BindRaster(pCellAreaSqKMRaster, "cellarea_sqkm")

            '    strExpression = "[cellarea_sqkm] * 0.386102"

            '    pCellAreaSqMileRaster = pMapAlgebraOp.Execute(strExpression)
            '    g_pCellAreaSqMiRaster = pCellAreaSqMileRaster

            '    'END STEP 5a1: -----------------------------------------------------------------------------


            '    'STEP 5b: ----------------------------------------------------------------------------------
            '    'Square Feet
            '    pMapAlgebraOp.BindRaster(pCellAreaSMRaster, "cellarea_sm")

            '    strExpression = "[cellarea_sm] * 10.76"

            '    pCellAreaSFRaster = pMapAlgebraOp.Execute(strExpression)
            '    'END STEP 5b--------------------------------------------------------------------------------


            '    'STEP 5c: ----------------------------------------------------------------------------------
            '    'Square Inches
            '    pMapAlgebraOp.BindRaster(pCellAreaSFRaster, "cellarea_sf")

            '    strExpression = "[cellarea_sf] * 144"

            '    pCellAreaSIRaster = pMapAlgebraOp.Execute(strExpression)
            '    'END STEP 5c--------------------------------------------------------------------------------

            '    'STEP 5d: ----------------------------------------------------------------------------------
            '    'Acres
            '    pMapAlgebraOp.BindRaster(pCellAreaSMRaster, "cell_sm")

            '    strExpression = "[cell_sm] * 0.000247104369"

            '    pCellAreaAcreRaster = pMapAlgebraOp.Execute(strExpression)
            '    'END STEP 5d--------------------------------------------------------------------------------


            '    'STEP 5e: ----------------------------------------------------------------------------------
            '    'Cubic Inches
            '    With pMapAlgebraOp
            '        .BindRaster(pCellAreaSIRaster, "cell_si")
            '        .BindRaster(pRunoffInRaster, "runoff_in")
            '    End With

            '    strExpression = "([cell_si] * [runoff_in])"

            '    pCellAreaCIRaster = pMapAlgebraOp.Execute(strExpression)

            '    'END STEP 5e -------------------------------------------------------------------------------


            '    'STEP 5f: -----------------------------------------------------------------------------------
            '    'Cubic feet
            '    pMapAlgebraOp.BindRaster(pCellAreaCIRaster, "runoff_ci")

            '    strExpression = "([runoff_ci] * 0.0005787)"

            '    pCellAreaCFRaster = pMapAlgebraOp.Execute(strExpression)
            '    g_pRunoffCFRaster = pCellAreaCFRaster

            '    'END STEP 5f: --------------------------------------------------------------------------------


            '    'STEP 5g: ------------------------------------------------------------------------------------
            '    'Runoff acre-feet
            '    pMapAlgebraOp.BindRaster(pCellAreaCFRaster, "runoff_cf")

            '    strExpression = "([runoff_cf] * 0.000022957)"

            '    pCellAreaAFRaster = pMapAlgebraOp.Execute(strExpression)
            '    g_pRunoffAFRaster = pCellAreaAFRaster

            '    'END STEP 5f: --------------------------------------------------------------------------------
            'Else
            '    Exit Function
            'End If


            'modProgDialog.ProgDialog("Converting to correct units...", strTitle, 0, 10, 8, 0)

            'If modProgDialog.g_boolCancel Then
            '    'STEP 6: -------------------------------------------------------------------------------------
            '    'Convert cubic inches to liters
            '    If pCellAreaCIRaster Is Nothing Then
            '        MsgBox("pCellAreaCIRaster nothing")
            '    End If

            '    pMapAlgebraOp.BindRaster(pCellAreaCIRaster, "run_ci")

            '    strExpression = "[run_ci] * 0.016387064"

            '    pMetRunoffRaster = pMapAlgebraOp.Execute(strExpression)
            '    g_pMetRunoffRaster = pMetRunoffRaster

            '    'END STEP 6: ---------------------------------------------------------------------------------
            'Else
            '    Exit Function
            'End If

            ''STEP 6a: -------------------------------------------------------------------------------------
            ''Eliminate nulls: Added 1/28/07
            'pMapAlgebraOp.BindRaster(pMetRunoffRaster, "runoffgrid")

            'strExpression = "Con(IsNull([runoffgrid]),0,[runoffgrid])"

            'pMetRunoffNoNullRaster = pMapAlgebraOp.Execute(strExpression)

            'g_pMetRunoffRaster = pMetRunoffNoNullRaster


            'Dim pClipLocAccumRaster As MapWinGIS.Grid
            'If g_booLocalEffects Then
            '    modProgDialog.ProgDialog("Creating data layer for local effects...", strTitle, 0, 10, 10, 0)
            '    If modProgDialog.g_boolCancel Then

            '        'STEP 12: Local Effects -------------------------------------------------
            '        strOutAccum = modUtil.GetUniqueName("locaccum", g_strWorkspace, ".tif")
            '        'Added 7/23/04 to account for clip by selected polys functionality
            '        If g_booSelectedPolys Then
            '            'TODO: Implement these modutil
            '            'pClipLocAccumRaster = modUtil.ClipBySelectedPoly(pMetRunoffNoNullRaster, g_pSelectedPolyClip, pEnv)
            '            'pPermAccumLocRunoffRaster = modUtil.ReturnPermanentRaster(pClipLocAccumRaster, pEnv.OutWorkspace.PathName, strOutAccum)
            '            pClipLocAccumRaster = Nothing
            '        Else
            '            'TODO
            '            'pPermAccumLocRunoffRaster = modUtil.ReturnPermanentRaster(pMetRunoffNoNullRaster, pEnv.OutWorkspace.PathName, strOutAccum)
            '        End If

            '        pLocRunoffRasterLayer = modUtil.ReturnRaster((frmPrj.m_App), pPermAccumLocRunoffRaster, "Runoff Local Effects (L)")
            '        'TODO: implement stretched thingy
            '        'pLocRunoffRasterLayer.Renderer = modUtil.ReturnRasterStretchColorRampRender(pLocRunoffRasterLayer, "Blue")
            '        pLocRunoffRasterLayer.Visible = False

            '        g_dicMetadata.Add(pLocRunoffRasterLayer.Name, m_strRunoffMetadata)

            '        'TODO: find g_pgrouplayer
            '        'g_pGroupLayer.Add(pLocRunoffRasterLayer)

            '        RunoffCalculation = True
            '        modProgDialog.KillDialog()
            '        Exit Function
            '    End If
            'End If

            'modProgDialog.ProgDialog("Creating flow accumulation...", strTitle, 0, 10, 9, 0)
            'If modProgDialog.g_boolCancel Then
            '    'STEP 7: ------------------------------------------------------------------------------------
            '    'Derive Accumulated Runoff

            '    With pMapAlgebraOp
            '        .BindRaster(g_pFlowDirRaster, "flowdir")
            '        .BindRaster(pMetRunoffNoNullRaster, "met_run")
            '    End With

            '    strExpression = "FlowAccumulation([flowdir], [met_run], FLOAT)"

            '    pAccumRunoffRaster1 = pMapAlgebraOp.Execute(strExpression)

            '    'END STEP 7: ----------------------------------------------------------------------------------
            'Else
            '    Exit Function
            'End If


            'modProgDialog.ProgDialog("Creating flow accumulation...", strTitle, 0, 10, 9, 0)
            'If modProgDialog.g_boolCancel Then
            '    'STEP 7: ------------------------------------------------------------------------------------
            '    'Derive the real Accumulated Runoff

            '    With pMapAlgebraOp
            '        .BindRaster(pAccumRunoffRaster1, "accum")
            '        .BindRaster(pMetRunoffNoNullRaster, "met_run")
            '    End With

            '    strExpression = "[accum] + [met_run]"

            '    pAccumRunoffRaster = pMapAlgebraOp.Execute(strExpression)

            '    'END STEP 7: ----------------------------------------------------------------------------------
            'Else
            '    Exit Function
            'End If

            ''Add this then map as our runoff grid
            'modProgDialog.ProgDialog("Creating Runoff Layer...", strTitle, 0, 10, 10, 0)
            'Dim pClipAccumRaster As MapWinGIS.Grid
            'If modProgDialog.g_boolCancel Then
            '    'Get a unique name for accumulation GRID
            '    strOutAccum = modUtil.GetUniqueName("runoff", g_strWorkspace, ".tif")

            '    'Clip to selected polys if chosen
            '    If g_booSelectedPolys Then
            '        'TODO: Implement
            '        'pClipAccumRaster = modUtil.ClipBySelectedPoly(pAccumRunoffRaster, g_pSelectedPolyClip, pEnv)
            '        'pPermAccumRunoffRaster = modUtil.ReturnPermanentRaster(pClipAccumRaster, pEnv.OutWorkspace.PathName, strOutAccum)
            '    Else
            '        'pPermAccumRunoffRaster = modUtil.ReturnPermanentRaster(pAccumRunoffRaster, pEnv.OutWorkspace.PathName, strOutAccum)
            '    End If

            '    'TODO
            '    'Add Completed raster to the g_pGroupLayer
            '    'pRunoffRasterLayer = modUtil.ReturnRasterLayer((frmPrj.m_App), pPermAccumRunoffRaster, "Accumulated Runoff (L)")
            '    'pRunoffRasterLayer.Renderer = ReturnRasterStretchColorRampRender(pRunoffRasterLayer, "Blue")


            '    'g_pGroupLayer.Add(pRunoffRasterLayer)

            '    'UPGRADE_WARNING: Couldn't resolve default property of object pRunoffRasterLayer.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            '    g_dicMetadata.Add(pRunoffRasterLayer.Name, m_strRunoffMetadata)

            '    'Global Runoff
            '    g_pRunoffRaster = pPermAccumRunoffRaster
            'Else
            '    Exit Function
            'End If

            'RunoffCalculation = True

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
End Module