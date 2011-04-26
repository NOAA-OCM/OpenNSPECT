'********************************************************************************************************
'File Name: modRunoff
'Description: Functions handling the runoff portion of the model
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
Imports System.Collections.Generic
Imports MapWinGeoProc
Imports MapWinGIS
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

    Public g_LandCoverRaster As Grid
    'LandCover GLOBAL
    Public g_pRunoffRaster As Grid
    'Runoff for use elsewhere
    Public g_pMetRunoffRaster As Grid
    'Metric Runoff
    Public g_pSCS100Raster As Grid
    'SCS GRID for use in RUSLE
    Public g_pPrecipRaster As Grid
    'Precipitation Raster
    Public g_strPrecipFileName As String
    'Precipitation File Name
    Public g_strLandCoverParameters As String
    'Global string to hold formatted LC params for metadata

    Private _strRunoffMetadata As String
    'Metatdata string for runoff
    Public g_intRainingDays As Short
    '# of Raining Days for Annual Precip
    Public g_intRunoffPrecipType As Short
    'Precip Event Type: 0=Annual; 1=Event

    Private _picks()() As String

    Public Function CreateRunoffGrid (ByRef strLCFileName As String, ByRef strLCCLassType As String, _
                                      ByRef cmdPrecip As OleDbCommand, ByRef strSoilsFileName As String, _
                                      ByRef OutputItems As clsXMLOutputItems) As Boolean
        'This sub serves as a link between frmPrj and the actual calculation of Runoff
        'It establishes the Rasters being used

        'strLCFileName: Path to land class file
        'strLCCLassType: type of landclass
        'rsPrecip: ADO recordset of precip scenario being used
        Dim pRainFallRaster As Grid
        Dim pLandCoverRaster As Grid
        Dim pSoilsRaster As Grid
        Dim strError As Object

        'Get the LandCover Raster
        'if no management scenarios were applied then g_landcoverRaster will be nothing
        If g_LandCoverRaster Is Nothing Then
            If RasterExists (strLCFileName) Then
                g_LandCoverRaster = ReturnRaster (strLCFileName)
                pLandCoverRaster = g_LandCoverRaster
            Else
                strError = strLCFileName
                MsgBox ("Error: The following dataset is missing: " & strError, MsgBoxStyle.Critical, "Missing Data")
                Return False
            End If
        Else
            pLandCoverRaster = g_LandCoverRaster
        End If

        Dim dataPrecip As OleDbDataReader = cmdPrecip.ExecuteReader()
        dataPrecip.Read()

        'Get the Precip Raster
        If RasterExists (dataPrecip ("PrecipFileName")) Then
            pRainFallRaster = ReturnRaster (dataPrecip ("PrecipFileName"))

            'If in cm, then convert to a GRID in inches.
            If dataPrecip ("PrecipUnits") = 0 Then
                g_pPrecipRaster = ConvertRainGridCMToInches (pRainFallRaster)
            Else 'if already in inches then just use this one.
                g_pPrecipRaster = pRainFallRaster
                'Global Precip
            End If

            g_strPrecipFileName = dataPrecip ("PrecipFileName")
            g_intRunoffPrecipType = dataPrecip ("Type")
            'Set the mod level precip type
            g_intRainingDays = dataPrecip ("RainingDays")
            'Set the mod level rainingdays
        Else
            strError = strLCFileName
            dataPrecip.Close()
            MsgBox ("Error: The following dataset is missing: " & strError, MsgBoxStyle.Critical, "Missing Data")
            Return False
        End If
        dataPrecip.Close()

        '------ 1.1.4 Change -----
        'if they select cm as incoming precip units, convert GRID

        'Get the Soils Raster
        If RasterExists (strSoilsFileName) Then
            pSoilsRaster = ReturnRaster (strSoilsFileName)
        Else
            strError = strSoilsFileName
            MsgBox ("Error: The following dataset is missing: " & strError, MsgBoxStyle.Critical, "Missing Data")
            Return False
        End If

        'Now construct the Pick Statement
        Dim strPick() As String = Nothing
        strPick = ConstructPickStatment (strLCCLassType, pLandCoverRaster)

        If strPick Is Nothing Then
            Exit Function
        End If

        'Call the Runoff Calculation using the string and rasters
        If RunoffCalculation (strPick, g_pPrecipRaster, pLandCoverRaster, pSoilsRaster, OutputItems) Then
            CreateRunoffGrid = True
        Else
            CreateRunoffGrid = False
            Exit Function
        End If

        CreateRunoffGrid = True

    End Function

    Private Function ConvertRainGridCMToInches (ByRef pInRaster As Grid) As Grid

        Dim head As GridHeader = pInRaster.Header
        Dim ncol As Integer = head.NumberCols - 1
        Dim nrow As Integer = head.NumberRows - 1
        Dim nodata As Single = head.NodataValue
        Dim rowvals() As Single
        ReDim rowvals(ncol)
        For row As Integer = 0 To nrow
            pInRaster.GetRow (row, rowvals (0))

            For col As Integer = 0 To ncol
                If rowvals (col) <> nodata Then
                    rowvals (col) = rowvals (col)/2.54
                End If
            Next

            pInRaster.PutRow (row, rowvals (0))
        Next

        Return pInRaster
    End Function

    Private Function ConstructPickStatment (ByRef strLandClass As String, ByRef pLCRaster As Grid) As String()
        'Creates the large initial pick statement using the name of the the LandCass [CCAP, for example]
        'and the Land Class Raster.  Returns a string
        ConstructPickStatment = Nothing
        Try
            Dim strRS As String
            Dim strPick(3) As String
            'Array of strings that hold 'pick' numbers

            Dim dblMaxValue As Double
            Dim i As Short
            Dim TableExist As Boolean
            Dim FieldIndex As Short
            Dim booValueFound As Boolean

            'STEP 1:  get the records from the database -----------------------------------------------
            strRS = _
                "SELECT LCCLASS.LCClassID, Value, LCCLASS.Name as Name2, LCCLASS.LCTypeID, [CN-A], [CN-B], [CN-C], [CN-D], CoverFactor, W_WL FROM LCCLASS " & _
                "LEFT OUTER JOIN LCTYPE ON LCCLASS.LCTYPEID = LCTYPE.LCTYPEID " & "Where LCTYPE.NAME = '" & strLandClass & _
                "' ORDER BY LCCLASS.VALUE"

            Dim cmdLandClass As New DataHelper (strRS)

            'while here, make metadata
            _strRunoffMetadata = CreateMetadata (cmdLandClass.GetCommand(), strLandClass, g_booLocalEffects)
            'End Database stuff

            'STEP 2: Raster Values ---------------------------------------------------------------------
            'Now Get the RASTER values

            'Get the max value
            dblMaxValue = pLCRaster.Maximum

            Dim tablepath As String = ""
            'Get the raster table
            Dim lcPath As String = pLCRaster.Filename
            If Path.GetFileName (lcPath) = "sta.adf" Then
                tablepath = Path.GetDirectoryName (lcPath) + ".dbf"
                If File.Exists (tablepath) Then

                    TableExist = True
                Else
                    TableExist = BuildTable (pLCRaster, tablepath)
                End If
            Else
                tablepath = Path.ChangeExtension (lcPath, ".dbf")
                If File.Exists (tablepath) Then
                    TableExist = True
                Else
                    TableExist = BuildTable (pLCRaster, tablepath)
                End If
            End If

            Dim mwTable As New Table
            If Not TableExist Then
                MsgBox ( _
                        "No MapWindow-readable raster table was found. To create one using ArcMap 9.3+, add the raster to the default project, right click on its layer and select Open Attribute Table. Now click on the options button in the lower right and select Export. In the export path, navigate to the directory of the grid folder and give the export the name of the raster folder with the .dbf extension. i.e. if you are exporting a raster attribute table from a raster named landcover, export landcover.dbf into the same level directory as the folder.", _
                        MsgBoxStyle.Exclamation, "Raster Attribute Table Not Found")

                Exit Function
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

                'STEP 3: Table values vs. Raster Values Count - if not equal bark -------------------------
                Dim dataLandClass = cmdLandClass.ExecuteReader()

                Dim rowidx As Integer = 0

                'STEP 4: Create the strings
                'Loop through and get all values
                For i = 1 To dblMaxValue

                    If (mwTable.CellValue (FieldIndex, rowidx) = i) Then _
'And (pRow.Value(FieldIndex) = rsLandClass!Value) Then
                        'MK a new reader was created each time this loop ran, but I couldn't see why as it was passed the same query.
                        booValueFound = False
                        While dataLandClass.Read()
                            If mwTable.CellValue (FieldIndex, rowidx) = dataLandClass ("Value") Then
                                booValueFound = True
                                If strPick (0) = "" Then
                                    strPick (0) = CStr (dataLandClass ("CN-A"))
                                    strPick (1) = CStr (dataLandClass ("CN-B"))
                                    strPick (2) = CStr (dataLandClass ("CN-C"))
                                    strPick (3) = CStr (dataLandClass ("CN-D"))
                                Else
                                    strPick (0) = strPick (0) & ", " & CStr (dataLandClass ("CN-A"))
                                    strPick (1) = strPick (1) & ", " & CStr (dataLandClass ("CN-B"))
                                    strPick (2) = strPick (2) & ", " & CStr (dataLandClass ("CN-C"))
                                    strPick (3) = strPick (3) & ", " & CStr (dataLandClass ("CN-D"))
                                End If
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
                            dataLandClass.Close()
                            mwTable.Close()
                            Exit Function
                        End If

                    Else
                        If strPick (0) = "" Then
                            strPick (0) = strPick (0) & "0"
                            strPick (1) = strPick (1) & "0"
                            strPick (2) = strPick (2) & "0"
                            strPick (3) = strPick (3) & "0"
                        Else
                            strPick (0) = strPick (0) & ", 0"
                            strPick (1) = strPick (1) & ", 0"
                            strPick (2) = strPick (2) & ", 0"
                            strPick (3) = strPick (3) & ", 0"
                        End If
                    End If

                Next i
                dataLandClass.Close()

                ConstructPickStatment = strPick

            End If

        Catch ex As Exception
            MsgBox (Err.Number & ": " & Err.Description & " " & "ConstructPickStatemnt")
        End Try
    End Function

    Private Function BuildTable (ByRef pLCRaster As Grid, ByVal tablepath As String) As Boolean
        Dim mwTable As New Table

        Dim result As Boolean = mwTable.CreateNew (tablepath)
        mwTable.StartEditingTable()
        Dim valfield As New Field()
        valfield.Name = "VALUE"
        valfield.Type = FieldType.INTEGER_FIELD
        mwTable.EditInsertField (valfield, mwTable.NumFields)

        Dim countfield As New Field()
        countfield.Name = "COUNT"
        countfield.Type = FieldType.INTEGER_FIELD
        mwTable.EditInsertField (countfield, mwTable.NumFields)

        Dim descfield As New Field()
        descfield.Name = "NAME"
        descfield.Type = FieldType.STRING_FIELD
        descfield.Width = 100
        mwTable.EditInsertField (descfield, mwTable.NumFields)

        Dim vallist As New List(Of Integer)
        Dim countlist As New List(Of Integer)

        Dim head As GridHeader = pLCRaster.Header
        Dim nr As Integer = head.NumberRows - 1
        Dim nc As Integer = head.NumberCols - 1
        Dim nodata As Single = head.NodataValue
        Dim rowvals(nc) As Single
        Dim itemfound As Boolean
        For row As Integer = 0 To nr
            pLCRaster.GetRow (row, rowvals (0))
            For col As Integer = 0 To nc
                If rowvals (col) <> nodata Then
                    itemfound = False
                    For i As Integer = 0 To vallist.Count - 1
                        If vallist (i) = rowvals (col) Then
                            itemfound = True
                            countlist (i) = countlist (i) + 1
                            Exit For
                        ElseIf vallist (i) > rowvals (col) Then
                            itemfound = True
                            vallist.Insert (i, rowvals (col))
                            countlist.Insert (i, 1)
                            Exit For
                        End If
                    Next
                    If Not itemfound Then
                        vallist.Add (rowvals (col))
                        countlist.Add (1)
                    End If
                End If
            Next
        Next
        Dim rowidx As Integer
        For i As Integer = 0 To vallist.Count - 1
            rowidx = mwTable.NumRows
            mwTable.EditInsertRow (rowidx)
            mwTable.EditCellValue (0, rowidx, vallist (i))
            mwTable.EditCellValue (1, rowidx, countlist (i))
        Next
        mwTable.StopEditingTable (True)
        mwTable.Close()
        If mwTable.NumRows > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function CreateMetadata (ByRef cmdLandClass As OleDbCommand, ByRef strLandClass As String, _
                                     ByRef booLocal As Boolean) As String
        CreateMetadata = ""
        Try
            Dim i As Short
            Dim strHeader As String
            Dim strLCHeader As String
            Dim strCoeffValues As String = ""

            If booLocal Then
                strHeader = vbTab & "Input Datasets:" & vbNewLine & vbTab & vbTab & "Hydrologic soils grid: " & _
                            g_clsXMLPrjFile.strSoilsHydFileName & vbNewLine & vbTab & vbTab & "Landcover grid: " & _
                            g_clsXMLPrjFile.strLCGridFileName & vbNewLine & vbTab & vbTab & "Precipitation grid: " & _
                            g_strPrecipFileName & vbNewLine
            Else
                strHeader = vbTab & "Input Datasets:" & vbNewLine & vbTab & vbTab & "Hydrologic soils grid: " & _
                            g_clsXMLPrjFile.strSoilsHydFileName & vbNewLine & vbTab & vbTab & "Landcover grid: " & _
                            g_clsXMLPrjFile.strLCGridFileName & vbNewLine & vbTab & vbTab & "Precipitation grid: " & _
                            g_strPrecipFileName & vbNewLine & vbTab & vbTab & "Flow direction grid: " & _
                            g_strFlowDirFilename & vbNewLine
            End If

            strLCHeader = vbTab & "Landcover parameters: " & vbNewLine & vbTab & vbTab & "Landcover type name: " & _
                          strLandClass & vbNewLine & vbTab & vbTab & "Coefficients: " & vbNewLine

            Dim dataLandClass As OleDbDataReader = cmdLandClass.ExecuteReader()

            While dataLandClass.Read()
                If i = 0 Then
                    strCoeffValues = vbTab & vbTab & vbTab & dataLandClass ("Name2") & ":" & vbNewLine & vbTab & vbTab & _
                                     vbTab & vbTab & "CN-A: " & CStr (dataLandClass ("CN-A")) & vbNewLine & vbTab & _
                                     vbTab & vbTab & vbTab & "CN-B: " & CStr (dataLandClass ("CN-B")) & vbNewLine & _
                                     vbTab & vbTab & vbTab & vbTab & "CN-C: " & CStr (dataLandClass ("CN-C")) & _
                                     vbNewLine & vbTab & vbTab & vbTab & vbTab & "CN-D: " & _
                                     CStr (dataLandClass ("CN-D")) & vbNewLine & vbTab & vbTab & vbTab & vbTab & _
                                     "Cover factor: " & dataLandClass ("CoverFactor") & vbNewLine
                    i = i + 1
                Else
                    strCoeffValues = strCoeffValues & vbTab & vbTab & vbTab & dataLandClass ("Name2") & ":" & vbNewLine & _
                                     vbTab & vbTab & vbTab & vbTab & "CN-A: " & CStr (dataLandClass ("CN-A")) & _
                                     vbNewLine & vbTab & vbTab & vbTab & vbTab & "CN-B: " & _
                                     CStr (dataLandClass ("CN-B")) & vbNewLine & vbTab & vbTab & vbTab & vbTab & _
                                     "CN-C: " & CStr (dataLandClass ("CN-C")) & vbNewLine & vbTab & vbTab & vbTab & _
                                     vbTab & "CN-D: " & CStr (dataLandClass ("CN-D")) & vbNewLine & vbTab & vbTab & _
                                     vbTab & vbTab & "Cover factor: " & dataLandClass ("CoverFactor") & vbNewLine
                End If
            End While

            CreateMetadata = strHeader & strLCHeader & strCoeffValues
            g_strLandCoverParameters = strLCHeader & strCoeffValues

            dataLandClass.Close()
        Catch ex As Exception
            HandleError (ex)
            'False, "CreateMetadata " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, _ParentHWND)
        End Try
    End Function

    Const ProgressTitle As String = "Processing Runoff Calculation..."

    ''' <summary>
    ''' Runoffs the calculation.
    ''' </summary>
    ''' <param name="strPick">our friend the dynamic pick statemnt.</param>
    ''' <param name="pInRainRaster">the precip grid.</param>
    ''' <param name="pInLandCoverRaster">landcover grid.</param>
    ''' <param name="pInSoilsRaster">soils grid.</param>
    ''' <param name="OutputItems">The output items.</param><returns></returns>
    Public Function RunoffCalculation (ByRef strPick As String(), ByRef pInRainRaster As Grid, _
                                       ByRef pInLandCoverRaster As Grid, _
                                       ByRef pInSoilsRaster As Grid, ByRef OutputItems As clsXMLOutputItems) _
        As Boolean

        Try
            Dim pSCS100Raster As Grid = Nothing
            'STEP 2: SCS * 100
            Dim pMetRunoffRaster As Grid = Nothing
            Dim pMetRunoffNoNullRaster As Grid = Nothing
            'STEP 6a:  no nulls
            Dim pAccumRunoffRaster As Grid = Nothing
            Dim pPermAccumRunoffRaster As Grid = Nothing
            Dim pPermAccumLocRunoffRaster As Grid = Nothing
            Dim strOutAccum As String

            ShowProgress ("Calculating maximum potential retention...", ProgressTitle, 0, 10, 3, _
                          g_frmProjectSetup)

            If g_KeepRunning Then
                'Calculate maxiumum potential retention

                Dim picksLength As Integer = strPick (0).Split (",").Length
                ReDim _picks(strPick.Length - 1)
                For i As Integer = 0 To strPick.Length - 1
                    ReDim _picks (i)(picksLength)
                    _picks (i) = strPick (i).Split (",")
                Next
                Dim sc100calc As New RasterMathCellCalc (AddressOf SC100CellCalc)
                RasterMath (pInSoilsRaster, pInLandCoverRaster, Nothing, Nothing, Nothing, pSCS100Raster, sc100calc)
                g_pSCS100Raster = pSCS100Raster
            Else
                Exit Function
            End If

            If g_KeepRunning Then
                ShowProgress ("Calculating runoff...", ProgressTitle, 0, 10, 6, g_frmProjectSetup)
                Dim AllRunOffCalc As New RasterMathCellCalcNulls (AddressOf AllRunoffCellCalc)
                RasterMath (pSCS100Raster, pInRainRaster, g_pDEMRaster, Nothing, Nothing, pMetRunoffRaster, Nothing, _
                            False, AllRunOffCalc)

                'STEP 6a: -------------------------------------------------------------------------------------
                'Eliminate nulls: Added 1/28/07
                Dim metRunoffNoNullcalc As New RasterMathCellCalcNulls (AddressOf metRunoffNoNullCellCalc)
                RasterMath (pMetRunoffRaster, g_pDEMRaster, Nothing, Nothing, Nothing, pMetRunoffNoNullRaster, Nothing, _
                            False, metRunoffNoNullcalc)

                g_pMetRunoffRaster = pMetRunoffNoNullRaster
                pMetRunoffRaster.Close()
            Else
                Exit Function
            End If

            If g_booLocalEffects Then
                ShowProgress ("Creating data layer for local effects...", ProgressTitle, 0, 10, 10, _
                              g_frmProjectSetup)
                If g_KeepRunning Then

                    'STEP 12: Local Effects -------------------------------------------------
                    strOutAccum = GetUniqueName ("locaccum", g_strWorkspace, g_FinalOutputGridExt)
                    'Added 7/23/04 to account for clip by selected polys functionality
                    If g_booSelectedPolys Then
                        pPermAccumLocRunoffRaster = _
                            ClipBySelectedPoly (g_pMetRunoffRaster, g_pSelectedPolyClip, strOutAccum)
                    Else
                        pPermAccumLocRunoffRaster = ReturnPermanentRaster (g_pMetRunoffRaster, strOutAccum)
                    End If

                    g_dicMetadata.Add ("Runoff Local Effects (L)", _strRunoffMetadata)

                    AddOutputGridLayer (pPermAccumLocRunoffRaster, "Blue", True, "Runoff Local Effects (L)", _
                                        "Runoff Local", - 1, OutputItems)

                    RunoffCalculation = True
                    CloseDialog()
                    Exit Function
                End If
            End If

            ShowProgress ("Creating flow accumulation...", ProgressTitle, 0, 10, 9, g_frmProjectSetup)
            If g_KeepRunning Then
                'STEP 7: ------------------------------------------------------------------------------------
                'Derive Accumulated Runoff

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
                g_pMetRunoffRaster.Save (strtmp2)

                Dim strtmpout As String = Path.GetTempFileName
                g_TempFilesToDel.Add (strtmpout)
                strtmpout = String.Format ("{0}out{1}", strtmpout, g_TAUDEMGridExt)
                g_TempFilesToDel.Add (strtmpout)
                DataManagement.DeleteGrid (strtmpout)

                'Use geoproc weightedAreaD8 after converting the D8 grid to taudem format bgd if needed
                Hydrology.WeightedAreaD8 (strtmp1, strtmp2, Nothing, strtmpout, False, False, _
                                          Environment.ProcessorCount, Nothing)
                'strExpression = "FlowAccumulation([flowdir], [met_run], FLOAT)"

                pAccumRunoffRaster = New Grid
                pAccumRunoffRaster.Open (strtmpout)

                pTauD8Flow.Close()
                DataManagement.DeleteGrid (strtmp1)

                'END STEP 7: ----------------------------------------------------------------------------------
            Else
                Exit Function
            End If

            'Add this then map as our runoff grid
            ShowProgress ("Creating Runoff Layer...", ProgressTitle, 0, 10, 10, g_frmProjectSetup)
            If g_KeepRunning Then
                'Get a unique name for accumulation GRID
                strOutAccum = GetUniqueName ("runoff", g_strWorkspace, g_FinalOutputGridExt)

                'Clip to selected polys if chosen
                If g_booSelectedPolys Then
                    pPermAccumRunoffRaster = _
                        ClipBySelectedPoly (pAccumRunoffRaster, g_pSelectedPolyClip, strOutAccum)
                Else
                    pPermAccumRunoffRaster = ReturnPermanentRaster (pAccumRunoffRaster, strOutAccum)
                End If

                AddOutputGridLayer (pPermAccumRunoffRaster, "Blue", True, "Accumulated Runoff (L)", "Runoff Accum", - 1, _
                                    OutputItems)

                g_dicMetadata.Add ("Accumulated Runoff (L)", _strRunoffMetadata)

                'Global Runoff
                g_pRunoffRaster = pAccumRunoffRaster
            Else
                Exit Function
            End If

            RunoffCalculation = True

            CloseDialog()

        Catch ex As Exception
            If Err.Number = - 2147217297 Then 'User cancelled operation
                g_KeepRunning = False
                RunoffCalculation = False
            ElseIf Err.Number = - 2147467259 Then
                MsgBox ( _
                        "ArcMap has reached the maximum number of GRIDs allowed in memory.  " & _
                        "Please exit OpenNSPECT and restart ArcMap.", MsgBoxStyle.Information, _
                        "Maximum GRID Number Encountered")
                g_KeepRunning = False
                CloseDialog()
                RunoffCalculation = False
            Else
                MsgBox ("Error: " & Err.Number & " on RunoffCalculation")
                g_KeepRunning = False
                CloseDialog()
                RunoffCalculation = False
            End If
        End Try
    End Function

#Region "Raster Math"

    Private Function SC100CellCalc (ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, _
                                    ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        'pSCS100Raster = 100 * con([pHydSoilsRaster] == 1, pick([pLandSampleRaster], " & strPick(0) & "), con([pHydSoilsRaster] == 2, pick([pLandSampleRaster], " & strPick(1) & "), con([pHydSoilsRaster] == 3, pick([pLandSampleRaster], " & strPick(2) & "), con([pHydSoilsRaster] == 4, pick([pLandSampleRaster], " & strPick(3) & ")))))

        If Input1 = 1 Then
            For i As Integer = 0 To _picks (0).Length - 1
                If Input2 = i + 1 Then
                    Return 100*_picks (0) (i)
                End If
            Next
        ElseIf Input1 = 2 Then
            For i As Integer = 0 To _picks (1).Length - 1
                If Input2 = i + 1 Then
                    Return 100*_picks (1) (i)
                End If
            Next
        ElseIf Input1 = 3 Then
            For i As Integer = 0 To _picks (2).Length - 1
                If Input2 = i + 1 Then
                    Return 100*_picks (2) (i)
                End If
            Next
        ElseIf Input1 = 4 Then
            For i As Integer = 0 To _picks (3).Length - 1
                If Input2 = i + 1 Then
                    Return 100*_picks (3) (i)
                End If
            Next
        End If
    End Function

    Private Function AllRunoffCellCalc (ByVal Input1 As Single, ByVal Input1Null As Single, ByVal Input2 As Single, _
                                        ByVal Input2Null As Single, ByVal Input3 As Single, ByVal Input3Null As Single, _
                                        ByVal Input4 As Single, ByVal Input4Null As Single, ByVal Input5 As Single, _
                                        ByVal Input5Null As Single, ByVal OutNull As Single) As Single
        Dim tmpVal, RetentionVal, AbstractVal, RunoffInches, AreaSquareMeters As Single

        If Input1 = Input1Null Or Input2 = Input2Null Or Input3 = Input3Null Then
            Return OutNull
        Else
            If Input1 = 0 Then 'avoid #inf, treat as null is correct, but 0 comes out looking better irritatingly
                Return OutNull
            Else
                If g_intRunoffPrecipType = 0 Then
                    RetentionVal = ((1000.0/Input1) - 10)*g_intRainingDays
                Else
                    RetentionVal = ((1000.0/Input1) - 10)
                End If

                AbstractVal = 0.2*RetentionVal

                tmpVal = Input2 - AbstractVal
                If tmpVal > 0 Then
                    RunoffInches = Math.Pow (tmpVal, 2)/(tmpVal + RetentionVal)
                Else
                    RunoffInches = 0
                End If

                If Input3 >= 0 Then
                    AreaSquareMeters = Math.Pow (g_dblCellSize, 2)
                Else
                    AreaSquareMeters = 0
                End If

                Return RunoffInches*(AreaSquareMeters*10.76*144)*0.016387064

            End If
        End If

    End Function

    Private Function metRunoffNoNullCellCalc (ByVal Input1 As Single, ByVal Input1Null As Single, ByVal Input2 As Single, _
                                              ByVal Input2Null As Single, ByVal Input3 As Single, _
                                              ByVal Input3Null As Single, ByVal Input4 As Single, _
                                              ByVal Input4Null As Single, ByVal Input5 As Single, _
                                              ByVal Input5Null As Single, ByVal OutNull As Single) As Single
        'strExpression = "Con(IsNull([runoffgrid]),0,[runoffgrid])"
        If Input1 <> Input1Null Then
            Return Input1
        Else
            If Input2 >= 0 Then
                Return 0
            Else
                Return OutNull
            End If
        End If
    End Function

    Public Function tauD8CellCalc (ByVal Input1 As Single, ByVal Input1Null As Single, ByVal Input2 As Single, _
                                   ByVal Input2Null As Single, ByVal Input3 As Single, ByVal Input3Null As Single, _
                                   ByVal Input4 As Single, ByVal Input4Null As Single, ByVal Input5 As Single, _
                                   ByVal Input5Null As Single, ByVal OutNull As Single) As Single
        'ESRI is clockwise 1-128 from east. TAUDEM is 1-8 counter-clockwise from east
        If Input1 = 1 Then
            Return 1
        ElseIf Input1 = 2 Then
            Return 8
        ElseIf Input1 = 4 Then
            Return 7
        ElseIf Input1 = 8 Then
            Return 6
        ElseIf Input1 = 16 Then
            Return 5
        ElseIf Input1 = 32 Then
            Return 4
        ElseIf Input1 = 64 Then
            Return 3
        ElseIf Input1 = 128 Then
            Return 2
        ElseIf Input1 = Input1Null Then
            Return - 1
        Else
            Return - 1
        End If
    End Function

    Public Function tauD8ToESRICellCalc (ByVal Input1 As Single, ByVal Input1Null As Single, ByVal Input2 As Single, _
                                         ByVal Input2Null As Single, ByVal Input3 As Single, ByVal Input3Null As Single, _
                                         ByVal Input4 As Single, ByVal Input4Null As Single, ByVal Input5 As Single, _
                                         ByVal Input5Null As Single, ByVal OutNull As Single) As Single
        'ESRI is clockwise 1-128 from east. TAUDEM is 1-8 counter-clockwise from east
        If Input1 = 1 Then
            Return 1
        ElseIf Input1 = 8 Then
            Return 2
        ElseIf Input1 = 7 Then
            Return 4
        ElseIf Input1 = 6 Then
            Return 8
        ElseIf Input1 = 5 Then
            Return 16
        ElseIf Input1 = 4 Then
            Return 32
        ElseIf Input1 = 3 Then
            Return 64
        ElseIf Input1 = 2 Then
            Return 128
        ElseIf Input1 = Input1Null Then
            Return - 1
        Else
            Return - 1
        End If
    End Function

#End Region
End Module