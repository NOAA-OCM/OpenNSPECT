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
Imports OpenNspect.Xml

Module Runoff
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

    ''' <summary>
    ''' Creates the runoff grid.
    ''' This sub serves as a link between frmPrj and the actual calculation of Runoff
    ''' It establishes the Rasters being used
    ''' </summary>
    ''' <param name="strLCFileName">Path to land class file.</param>
    ''' <param name="strLCCLassType">type of landclass.</param>
    ''' <param name="cmdPrecip">ADO recordset of precip scenario being used.</param>
    ''' <param name="strSoilsFileName">Name of the STR soils file.</param>
    ''' <param name="OutputItems">The output items.</param><returns></returns>
    Public Function CreateRunoffGrid(ByRef strLCFileName As String, ByRef strLCCLassType As String,
                                      ByRef cmdPrecip As OleDbCommand, ByRef strSoilsFileName As String,
                                      ByRef OutputItems As OutputItems) As Boolean

        'Get the LandCover Raster
        'if no management scenarios were applied then g_landcoverRaster will be nothing
        If g_LandCoverRaster Is Nothing Then
            If RasterExists(strLCFileName) Then
                g_LandCoverRaster = ReturnRaster(strLCFileName)
            Else
                MsgBox("Error: The following dataset is missing: " & strLCFileName, MsgBoxStyle.Critical, "Missing Data")
                Return False
            End If
        End If

        Using dataPrecip As OleDbDataReader = cmdPrecip.ExecuteReader()
            dataPrecip.Read()
            'Get the Precip Raster
            If RasterExists(dataPrecip("PrecipFileName")) Then
                Dim pRainFallRaster As Grid = ReturnRaster(dataPrecip("PrecipFileName"))
                'If in cm, then convert to a GRID in inches.
                If dataPrecip("PrecipUnits") = 0 Then
                    g_pPrecipRaster = ConvertRainGridCMToInches(pRainFallRaster)
                Else
                    'if already in inches then just use this one.
                    g_pPrecipRaster = pRainFallRaster
                End If
                g_strPrecipFileName = dataPrecip("PrecipFileName")
                g_intRunoffPrecipType = dataPrecip("Type")
                g_intRainingDays = dataPrecip("RainingDays")
            Else
                MsgBox("Error: The following dataset is missing: " & strLCFileName, MsgBoxStyle.Critical, "Missing Data")
                Return False
            End If
        End Using

        'Get the Soils Raster
        Dim pSoilsRaster As Grid
        If RasterExists(strSoilsFileName) Then
            pSoilsRaster = ReturnRaster(strSoilsFileName)
        Else
            MsgBox("Error: The following dataset is missing: " & strSoilsFileName, MsgBoxStyle.Critical, "Missing Data")
            Return False
        End If

        'Now construct the Pick Statement
        Dim strPick() As String = ConstructPickStatment(strLCCLassType, g_LandCoverRaster)

        If strPick Is Nothing Then
            Return False
        End If

        Return RunoffCalculation(strPick, g_pPrecipRaster, g_LandCoverRaster, pSoilsRaster, OutputItems)

    End Function

    Private Function ConvertRainGridCMToInches(ByRef pInRaster As Grid) As Grid

        Dim head As GridHeader = pInRaster.Header
        Dim ncol As Integer = head.NumberCols - 1
        Dim nrow As Integer = head.NumberRows - 1
        Dim nodata As Single = head.NodataValue
        Dim rowvals(ncol) As Single
        
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

    Private Function ConstructPickStatment(ByRef strLandClass As String, ByRef pLCRaster As Grid) As String()
        'Creates the large initial pick statement using the name of the the LandCass [CCAP, for example]
        'and the Land Class Raster.  Returns a string
        ConstructPickStatment = Nothing
        Try
            Dim strRS As String
            Dim strPick(3) As String
            'Array of strings that hold 'pick' numbers

            Dim dblMaxValue As Double
            Dim dblMinValue As Double
            Dim i As Short
            Dim FieldIndex As Short
            Dim booValueFound As Boolean

            'STEP 1:  get the records from the database -----------------------------------------------
            strRS =
                "SELECT LCCLASS.LCClassID, Value, LCCLASS.Name as Name2, LCCLASS.LCTypeID, [CN-A], [CN-B], [CN-C], [CN-D], CoverFactor, W_WL FROM LCCLASS " &
                "LEFT OUTER JOIN LCTYPE ON LCCLASS.LCTYPEID = LCTYPE.LCTYPEID " & "Where LCTYPE.NAME = '" & strLandClass &
                "' ORDER BY LCCLASS.VALUE"

            Dim cmdLandClass As New DataHelper(strRS)

            'while here, make metadata
            _strRunoffMetadata = CreateMetadata(cmdLandClass.GetCommand(), strLandClass, g_Project.IncludeLocalEffects)
            'End Database stuff

            'STEP 2: Raster Values ---------------------------------------------------------------------
            'Now Get the RASTER values

            ' pLCRaster.Maximum not updating correctly when adding land Use scenarios 7/30/2012
            ' FIXED 8/30/2012: Temp LC raster updates .Maximum field correctly.
            dblMaxValue = pLCRaster.Maximum
            dblMinValue = pLCRaster.Minimum

            'TODO: it looks like some of this code is almost copied, refactor it.
            Dim tablepath = GetRasterTablePath(pLCRaster)
            Dim TableExist As Boolean = File.Exists(tablepath)
            ' Refactored!

            Dim mwTable As New Table
            If Not TableExist Then
                MsgBox("No MapWindow-readable raster table was found. (3)", MsgBoxStyle.Exclamation, "Raster Attribute Table Not Found")

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
                Dim dataLandClass = cmdLandClass.ExecuteReader()

                Dim rowidx As Integer = 0

                'STEP 4: Create the strings
                'Loop through and get all values
                ' DLE: 109/12: LC rasters may have 0 as min value, but if minval > 1, then start at 1
                If (dblMinValue > 1) Then dblMinValue = 1
                For i = dblMinValue To dblMaxValue  ' FIXED DLE 10/9/2012 Start at min value, not 1.  Min may be 0.
                    ' For i = 1 To dblMaxValue  ' FIXED DLE 10/9/2012 Start at min value, not 1.  Min may be 0.

                    If (mwTable.CellValue(FieldIndex, rowidx) = i) Then
                        'MK a new reader was created each time this loop ran, but I couldn't see why as it was passed the same query.
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
                            MsgBox("Error: Your OpenNSPECT Land Class Table is missing values found in your landcover GRID dataset.")
                            ConstructPickStatment = Nothing
                            dataLandClass.Close()
                            mwTable.Close()
                            Exit Function
                        End If

                    Else
                        If strPick(0) = "" Then
                            strPick(0) = strPick(0) & "0"
                            strPick(1) = strPick(1) & "0"
                            strPick(2) = strPick(2) & "0"
                            strPick(3) = strPick(3) & "0"
                        Else
                            strPick(0) = strPick(0) & ", 0"
                            strPick(1) = strPick(1) & ", 0"
                            strPick(2) = strPick(2) & ", 0"
                            strPick(3) = strPick(3) & ", 0"
                        End If
                    End If

                Next i
                dataLandClass.Close()

                ConstructPickStatment = strPick

            End If

        Catch ex As Exception
            HandleError(ex)
        End Try
    End Function

    Public Function BuildTable(ByRef pLCRaster As Grid, ByVal tablepath As String) As Boolean
        Dim mwTable As New Table

        Dim result As Boolean = mwTable.CreateNew(tablepath)
        If Not result Then
            Throw New ArgumentException("Could not create new table.")
        End If

        mwTable.StartEditingTable()
        Dim valfield As New Field()
        valfield.Name = "VALUE"
        valfield.Type = FieldType.INTEGER_FIELD
        mwTable.EditInsertField(valfield, mwTable.NumFields)

        Dim countfield As New Field()
        countfield.Name = "COUNT"
        countfield.Type = FieldType.INTEGER_FIELD
        mwTable.EditInsertField(countfield, mwTable.NumFields)

        Dim descfield As New Field()
        descfield.Name = "NAME"
        descfield.Type = FieldType.STRING_FIELD
        descfield.Width = 100
        mwTable.EditInsertField(descfield, mwTable.NumFields)

        Dim vallist As New List(Of Integer)
        Dim countlist As New List(Of Integer)

        Dim head As GridHeader = pLCRaster.Header
        Dim nr As Integer = head.NumberRows - 1
        Dim nc As Integer = head.NumberCols - 1
        Dim nodata As Single = head.NodataValue
        Dim rowvals(nc) As Single
        Dim itemfound As Boolean
        For row As Integer = 0 To nr
            pLCRaster.GetRow(row, rowvals(0))
            For col As Integer = 0 To nc
                If rowvals(col) <> nodata Then
                    itemfound = False
                    For i As Integer = 0 To vallist.Count - 1
                        If vallist(i) = rowvals(col) Then
                            itemfound = True
                            countlist(i) = countlist(i) + 1
                            Exit For
                        ElseIf vallist(i) > rowvals(col) Then
                            itemfound = True
                            vallist.Insert(i, rowvals(col))
                            countlist.Insert(i, 1)
                            Exit For
                        End If
                    Next
                    If Not itemfound Then
                        vallist.Add(rowvals(col))
                        countlist.Add(1)
                    End If
                End If
            Next
        Next
        Dim rowidx As Integer
        For i As Integer = 0 To vallist.Count - 1
            rowidx = mwTable.NumRows
            mwTable.EditInsertRow(rowidx)
            mwTable.EditCellValue(0, rowidx, vallist(i))
            mwTable.EditCellValue(1, rowidx, countlist(i))
        Next
        mwTable.StopEditingTable(True)
        'mwTable.Close()  ' In 4.8.8 you can't close a table then access a property of that table.  Apparently you could in 4.8.6
        If mwTable.NumRows > 0 Then '
            mwTable.Close()
            Return True
        Else
            mwTable.Close()
            Return False
        End If
    End Function

    Private Function CreateMetadata(ByRef cmdLandClass As OleDbCommand, ByRef strLandClass As String,
                                     ByRef includeLocalEffects As Boolean) As String
        CreateMetadata = ""
        Try
            Dim i As Short
            Dim strHeader As String
            Dim strLCHeader As String
            Dim strCoeffValues As String = ""
            strHeader = vbTab & "Input Datasets:" & vbNewLine & vbTab & vbTab & "Hydrologic soils grid: " &
                                        g_Project.SoilsHydDirectory & vbNewLine & vbTab & vbTab & "Landcover grid: " &
                                        g_Project.LandCoverGridDirectory & vbNewLine & vbTab & vbTab & "Precipitation grid: " &
                                        g_strPrecipFileName & vbNewLine
            If includeLocalEffects Then

            Else
                strHeader += vbTab & vbTab & "Flow direction grid: " & g_strFlowDirFilename & vbNewLine
            End If

            strLCHeader = vbTab & "Landcover parameters: " & vbNewLine & vbTab & vbTab & "Landcover type name: " &
                          strLandClass & vbNewLine & vbTab & vbTab & "Coefficients: " & vbNewLine

            Dim dataLandClass As OleDbDataReader = cmdLandClass.ExecuteReader()

            While dataLandClass.Read()
                Dim info As String = vbTab & vbTab & vbTab & dataLandClass("Name2") & ":" & vbNewLine & vbTab & vbTab &
                                                     vbTab & vbTab & "CN-A: " & CStr(dataLandClass("CN-A")) & vbNewLine & vbTab &
                                                     vbTab & vbTab & vbTab & "CN-B: " & CStr(dataLandClass("CN-B")) & vbNewLine &
                                                     vbTab & vbTab & vbTab & vbTab & "CN-C: " & CStr(dataLandClass("CN-C")) &
                                                     vbNewLine & vbTab & vbTab & vbTab & vbTab & "CN-D: " &
                                                     CStr(dataLandClass("CN-D")) & vbNewLine & vbTab & vbTab & vbTab & vbTab &
                                                     "Cover factor: " & dataLandClass("CoverFactor") & vbNewLine
                If i = 0 Then
                    strCoeffValues = info
                    i = i + 1
                Else
                    strCoeffValues += info
                End If
            End While

            CreateMetadata = strHeader & strLCHeader & strCoeffValues
            g_strLandCoverParameters = strLCHeader & strCoeffValues

            dataLandClass.Close()
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Function


    Private Sub CalculateMaxiumumPotentialRetention(ByRef strPick As String(), ByRef pInLandCoverRaster As Grid, ByRef pInSoilsRaster As Grid)
        ReDim _picks(strPick.Length - 1)
        For i As Integer = 0 To strPick.Length - 1
            _picks(i) = strPick(i).Split(",")
        Next
        Dim sc100calc As New RasterMathCellCalc(AddressOf SC100CellCalc)
        Dim pSCS100Raster As Grid = Nothing
        RasterMath(pInSoilsRaster, pInLandCoverRaster, Nothing, Nothing, Nothing, pSCS100Raster, sc100calc)
        g_pSCS100Raster = pSCS100Raster
        g_pSCS100Raster.Header.Projection = g_pDEMRaster.Header.Projection 'DLE 7/25/2012: Fix Local Effect runoff not projected issue.
    End Sub
    Private Sub CalculateRunoff(ByRef pInRainRaster As Grid)
        Dim AllRunOffCalc As New RasterMathCellCalcNulls(AddressOf AllRunoffCellCalc)
        Dim pMetRunoffRaster As Grid = Nothing
        RasterMath(g_pSCS100Raster, pInRainRaster, g_pDEMRaster, Nothing, Nothing, pMetRunoffRaster, Nothing,
                                    False, AllRunOffCalc)

        'Eliminate nulls: Added 1/28/07
        Dim metRunoffNoNullcalc As New RasterMathCellCalcNulls(AddressOf metRunoffNoNullCellCalc)
        Dim pMetRunoffNoNullRaster As Grid = Nothing
        RasterMath(pMetRunoffRaster, g_pDEMRaster, Nothing, Nothing, Nothing, pMetRunoffNoNullRaster, Nothing,
                                    False, metRunoffNoNullcalc)

        g_pMetRunoffRaster = pMetRunoffNoNullRaster
        g_pMetRunoffRaster.Header.Projection = g_pDEMRaster.Header.Projection 'DLE 10/5/2012: Make sure grid is projected.

        pMetRunoffRaster.Close()
    End Sub
    Private Function CreateDataLayerForLocalEffects(ByRef OutputItems As OutputItems) As String
        Dim strOutAccum As String = GetUniqueFileName("locaccum", g_Project.ProjectWorkspace, OutputGridExt)
        'Added 7/23/04 to account for clip by selected polys functionality
        Dim pPermAccumLocRunoffRaster As Grid = Nothing
        If g_Project.UseSelectedPolygons Then
            pPermAccumLocRunoffRaster = ClipBySelectedPoly(g_pMetRunoffRaster, g_pSelectedPolyClip, strOutAccum)
        Else
            pPermAccumLocRunoffRaster = CopyRaster(g_pMetRunoffRaster, strOutAccum)
        End If

        g_dicMetadata.Add("Runoff Local Effects (L)", _strRunoffMetadata)
        writeMetadata(g_Project.ProjectName, "Runoff Local Effects (L)", _strRunoffMetadata, pPermAccumLocRunoffRaster.Filename)

        AddOutputGridLayer(pPermAccumLocRunoffRaster, "Blue", True, "Runoff Local Effects (L)",
                            "Runoff Local", -1, OutputItems)
        Return strOutAccum
    End Function

    Private Function DeriveAccumulatedRunoff() As Grid
        Dim tauD8calc = GetConverterToTauDemFromEsri()
        Dim pTauD8Flow As Grid = Nothing
        RasterMath(g_pFlowDirRaster, Nothing, Nothing, Nothing, Nothing, pTauD8Flow, Nothing, False, tauD8calc)
        pTauD8Flow.Header.NodataValue = -1

        Dim flowFile As String = GetTempFileNameTauDemGridExt()
        pTauD8Flow.Save() 'Saving because it seemed necessary.
        If Not pTauD8Flow.Save(flowFile) Then Return Nothing

        Dim metRunoffFile As String = GetTempFileNameTauDemGridExt()
        g_pMetRunoffRaster.Save() 'Saving because it seemed necessary.
        If Not g_pMetRunoffRaster.Save(metRunoffFile) Then Return Nothing

        Dim outFile As String = GetTempFileNameTauDemGridExt()

        Dim result = Hydrology.WeightedAreaD8(flowFile, metRunoffFile, Nothing, outFile, False, False,
                                  Environment.ProcessorCount, False, Nothing)
        If result <> 0 Then
            SynchronousProgressDialog.KeepRunning = False
        End If
        Dim pAccumRunoffRaster As Grid = New Grid
        pAccumRunoffRaster.Open(outFile)

        Return pAccumRunoffRaster
    End Function
    Private Sub CreateRunoffGridFileAndLayer(ByRef OutputItems As OutputItems, ByVal pAccumRunoffRaster As Grid)
        'Get a unique name for accumulation GRID
        Dim strOutAccum = GetUniqueFileName("runoff", g_Project.ProjectWorkspace, OutputGridExt)

        'Clip to selected polys if chosen
        Dim pPermAccumRunoffRaster As Grid = Nothing
        If g_Project.UseSelectedPolygons Then
            pPermAccumRunoffRaster = ClipBySelectedPoly(pAccumRunoffRaster, g_pSelectedPolyClip, strOutAccum)
        Else
            pPermAccumRunoffRaster = CopyRaster(pAccumRunoffRaster, strOutAccum)
        End If

        g_dicMetadata.Add("Accumulated Runoff (L)", _strRunoffMetadata)
        writeMetadata(g_Project.ProjectName, "Accumulated Runoff (L)", _strRunoffMetadata, pPermAccumRunoffRaster.Filename)

        AddOutputGridLayer(pPermAccumRunoffRaster, "Blue", True, "Accumulated Runoff (L)", "Runoff Accum", -1, OutputItems)

        'Global Runoff
        g_pRunoffRaster = pAccumRunoffRaster
    End Sub
    ''' <summary>
    ''' Runoffs the calculation.
    ''' </summary>
    ''' <param name="picks">our friend the dynamic pick statemnt.</param>
    ''' <param name="pInRainRaster">the precip grid.</param>
    ''' <param name="pInLandCoverRaster">landcover grid.</param>
    ''' <param name="pInSoilsRaster">soils grid.</param>
    ''' <param name="OutputItems">The output items.</param><returns></returns>
    Public Function RunoffCalculation(ByRef picks As String(), ByRef pInRainRaster As Grid,
                                       ByRef pInLandCoverRaster As Grid,
                                       ByRef pInSoilsRaster As Grid, ByRef OutputItems As OutputItems) As Boolean

        Try
            Using progress = New SynchronousProgressDialog("Calculating maximum potential retention...", "Processing Runoff Calculation...", 5, g_MainForm)
                CalculateMaxiumumPotentialRetention(picks, pInLandCoverRaster, pInSoilsRaster)

                If Not progress.Increment("Calculating runoff...") Then Return False
                CalculateRunoff(pInRainRaster)

                If g_Project.IncludeLocalEffects Then
                    If Not progress.Increment("Creating data layer for local effects...") Then Return False
                    CreateDataLayerForLocalEffects(OutputItems)
                Else
                    Dim pRaster = g_pMetRunoffRaster
                    pRaster.Save()
                End If

                If Not progress.Increment("Creating flow accumulation...") Then Return False 'already failed at this point
                Dim pAccumRunoffRaster As Grid = DeriveAccumulatedRunoff()

                If Not progress.Increment("Creating Runoff Layer...") Then Return False
                CreateRunoffGridFileAndLayer(OutputItems, pAccumRunoffRaster)
            End Using

            Return True

        Catch ex As Exception
            HandleError(ex)
            Return False
        End Try
    End Function

#Region "Raster Math"

    Private Function SC100CellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single,
                                    ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single

        If Input1 = 1 OrElse Input1 = 2 OrElse Input1 = 3 OrElse Input1 = 4 Then
            Dim position = Input1 - 1
            For i As Integer = 0 To _picks(position).Length - 1
                If Input2 = i + 1 Then
                    Return 100 * _picks(position)(i)
                End If
            Next
        End If

    End Function

    Private Function AllRunoffCellCalc(ByVal Input1 As Single, ByVal Input1Null As Single, ByVal Input2 As Single,
                                        ByVal Input2Null As Single, ByVal Input3 As Single, ByVal Input3Null As Single,
                                        ByVal Input4 As Single, ByVal Input4Null As Single, ByVal Input5 As Single,
                                        ByVal Input5Null As Single, ByVal OutNull As Single) As Single
        Dim tmpVal, RetentionVal, AbstractVal, RunoffInches, AreaSquareMeters As Single

        If Input1 = Input1Null Or Input2 = Input2Null Or Input3 = Input3Null Then
            Return OutNull
        Else
            If Input1 = 0 Then 'avoid #inf, treat as null is correct, but 0 comes out looking better irritatingly
                Return OutNull
            Else
                If g_intRunoffPrecipType = 0 Then
                    RetentionVal = ((1000.0 / Input1) - 10) * g_intRainingDays
                Else
                    RetentionVal = ((1000.0 / Input1) - 10)
                End If

                AbstractVal = 0.2 * RetentionVal

                tmpVal = Input2 - AbstractVal
                If tmpVal > 0 Then
                    RunoffInches = Math.Pow(tmpVal, 2) / (tmpVal + RetentionVal)
                Else
                    RunoffInches = 0
                End If

                If Input3 >= 0 Then
                    AreaSquareMeters = Math.Pow(g_CellSize, 2)
                Else
                    AreaSquareMeters = 0
                End If

                Return RunoffInches * (AreaSquareMeters * 10.76 * 144) * 0.016387064

            End If
        End If

    End Function

    Private Function metRunoffNoNullCellCalc(ByVal runnOffGrid As Single, ByVal Input1Null As Single, ByVal Input2 As Single,
                                              ByVal Input2Null As Single, ByVal Input3 As Single,
                                              ByVal Input3Null As Single, ByVal Input4 As Single,
                                              ByVal Input4Null As Single, ByVal Input5 As Single,
                                              ByVal Input5Null As Single, ByVal OutNull As Single) As Single
        If runnOffGrid <> Input1Null Then
            Return runnOffGrid
        Else
            If Input2 >= 0 Then
                Return 0
            Else
                Return OutNull
            End If
        End If
    End Function

#End Region
End Module