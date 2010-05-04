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
    Const c_sModuleFileName As String = "modMUSLE.bas"
    Private _picks As String()
    Private _pondpicks As String()


    Public Function MUSLESetup(ByRef strSoilsDefName As String, ByRef strKfactorFileName As String, ByRef strLandClass As String) As Boolean
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

        If Len(g_DictTempNames.Item(strLandClass)) > 0 Then
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
        If CalcMUSLE(_strCFactorConStatement, _strPondConStatement) Then
            MUSLESetup = True
        Else
            MUSLESetup = False
        End If

    End Function

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
                For i = 0 To maxVal
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
                            strpick = "-9999"
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

    'OBSOLETE
    'Private Function ConstructConStatment(ByRef cmdCF As OleDbCommand, ByRef pLCRaster As MapWinGIS.Grid) As String
    '    ''Creates the Cover Factor con statement using the name of the the LandCass Recordset, and the Landclass Raster
    '    ''Returns: String
    '    ''Looks like: con(([nu_lulc] eq 2), 0.000, con((nu_lulc eq 3), 0.030....

    '    'Dim strCon As String 'Con statement base
    '    'Dim strParens As String 'String of trailing parens
    '    'Dim strCompleteCon As String 'Concatenate of strCon & strParens

    '    'Dim pRasterCol As ESRI.ArcGIS.DataSourcesRaster.IRasterBandCollection
    '    'Dim pBand As ESRI.ArcGIS.DataSourcesRaster.IRasterBand
    '    'Dim pTable As ESRI.ArcGIS.Geodatabase.ITable
    '    'Dim TableExist As Boolean
    '    'Dim pCursor As ESRI.ArcGIS.Geodatabase.ICursor
    '    'Dim pRow As ESRI.ArcGIS.Geodatabase.IRow
    '    'Dim FieldIndex As Short
    '    'Dim booValueFound As Boolean
    '    'Dim i As Short

    '    ''STEP 1:  get the records from the database -----------------------------------------------
    '    'rsCF.MoveFirst()
    '    ''End Database stuff

    '    ''STEP 2: Raster Values ---------------------------------------------------------------------
    '    ''Now Get the RASTER values
    '    '' Get Rasterband from the incoming raster
    '    'pRasterCol = pLCRaster
    '    'pBand = pRasterCol.Item(0)

    '    ''Get the raster table
    '    'pBand.HasTable(TableExist)
    '    'If Not TableExist Then Exit Function

    '    'pTable = pBand.AttributeTable
    '    ''Get All rows
    '    'pCursor = pTable.Search(Nothing, True)
    '    ''Init pRow
    '    'pRow = pCursor.NextRow

    '    ''Get index of Value Field
    '    'FieldIndex = pTable.FindField("Value")

    '    ''STEP 3: Table values vs. Raster Values Count - if not equal bark --------------------------
    '    ''    If pTable.RowCount(Nothing) <> rsCF.RecordCount Then
    '    ''        MsgBox "Error: The number of records in your Land Class Table do not match your GRID dataset."
    '    ''        Exit Function
    '    ''    End If

    '    ''STEP 4: Create the strings
    '    ''Loop through and get all values
    '    'Do While Not pRow Is Nothing
    '    '    booValueFound = False
    '    '    rsCF.MoveFirst()

    '    '    For i = 0 To rsCF.RecordCount - 1
    '    '        If pRow.Value(FieldIndex) = rsCF.Fields("Value").Value Then

    '    '            booValueFound = True

    '    '            If strCon = "" Then
    '    '                strCon = "Con(([nu_lulc] eq " & pRow.Value(FieldIndex) & "), " & rsCF.Fields("CoverFactor").Value & ", "
    '    '            Else
    '    '                strCon = strCon & "Con(([nu_lulc] eq " & pRow.Value(FieldIndex) & "), " & rsCF.Fields("CoverFactor").Value & ", "
    '    '            End If

    '    '            If strParens = "" Then
    '    '                strParens = "-9999)"
    '    '            Else
    '    '                strParens = strParens & ")"
    '    '            End If
    '    '            Exit For
    '    '        Else
    '    '            booValueFound = False
    '    '        End If
    '    '        rsCF.MoveNext()
    '    '    Next i

    '    '    If booValueFound = False Then
    '    '        MsgBox("Error in MUSLE ConstructConStatment Function: Your N-SPECT Land Class Table is missing values found in your landcover GRID dataset.")
    '    '        ConstructConStatment = ""
    '    '        Exit Function
    '    '    Else
    '    '        pRow = pCursor.NextRow
    '    '        i = 0
    '    '    End If

    '    'Loop


    '    ''Remove 11/30/2007 in favor of check above.
    '    ''    '==========================================================================
    '    ''    Do While Not pRow Is Nothing
    '    ''
    '    ''        booValueFound = False
    '    ''        rsType.MoveFirst
    '    ''
    '    ''        For i = 0 To rsType.RecordCount - 1
    '    ''            If rsType!Value = pRow.Value Then
    '    ''
    '    ''                booValueFound = True
    '    ''
    '    ''                If strCon = "" Then
    '    ''                    strCon = "Con(([nu_lulc] eq " & pRow.Value(FieldIndex) & "), " & rsType!CoeffType & ", "
    '    ''                Else
    '    ''                    strCon = strCon & "Con(([nu_lulc] eq " & pRow.Value(FieldIndex) & "), " & rsType!CoeffType & ", "
    '    ''                End If
    '    ''
    '    ''                If strParens = "" Then
    '    ''                    strParens = "-9999)"
    '    ''                Else
    '    ''                    strParens = strParens & ")"
    '    ''                End If
    '    ''
    '    ''                Exit For
    '    ''            Else
    '    ''                booValueFound = False
    '    ''            End If
    '    ''            rsType.MoveNext
    '    ''        Next i
    '    ''
    '    ''        If booValueFound = False Then
    '    ''            MsgBox "Values in table LCClass table not equal to values in landclass dataset."
    '    ''            ConstructConStatment = ""
    '    ''            Exit Function
    '    ''        Else
    '    ''            Set pRow = pCursor.NextRow
    '    ''            i = 0
    '    ''        End If
    '    ''    Loop
    '    ''==================================================================================

    '    'strCompleteCon = strCon & strParens
    '    'ConstructConStatment = strCompleteCon

    '    ''Cleanup:
    '    ''Set pLCRaster = Nothing
    '    'pRasterCol = Nothing
    '    'pBand = Nothing
    '    'pTable = Nothing
    '    'pCursor = Nothing
    '    'pRow = Nothing

    'End Function

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
            MsgBox("No MapWindow-readable raster table was found for the landcover raster. To create one using ArcMap 9.3+, add the raster to the default project, right click on its layer and select Open Attribute Table. Now click on the options button in the lower right and select Export. In the export path, navigate to the directory of the grid folder and give the export the name of the raster folder with the .dbf extension. i.e. if you are exporting a raster attribute table from a raster named landcover, export landcover.dbf into the same level directory as the folder.", MsgBoxStyle.Exclamation, "Raster Attribute Table Not Found")

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
            For i = 0 To maxVal
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
                        strpick = "-9999"
                    Else
                        strpick = strpick & ", " & nodata
                    End If
                End If

            Next
        End If


        'strCompleteCon = strCon & strParens
        ConstructPondPickStatement = strpick

    End Function

    'OBSOLETE
    'Private Function ConstructPondConStatement(ByRef rsCF As ADODB.Recordset, ByRef pLCRaster As ESRI.ArcGIS.Geodatabase.IRaster) As String
    '    'Creates the Con Statement used in the Pond Factor GRID
    '    'Returns: String
    '    'Looks like: con(([nu_lulc] eq 16), 0, con((nu_lulc eq 17), 0...

    '    Dim strCon As String 'Con statement base
    '    Dim strParens As String 'String of trailing parens
    '    Dim strCompleteCon As String 'Concatenate of strCon & strParens

    '    Dim pRasterCol As ESRI.ArcGIS.DataSourcesRaster.IRasterBandCollection
    '    Dim pBand As ESRI.ArcGIS.DataSourcesRaster.IRasterBand
    '    Dim pTable As ESRI.ArcGIS.Geodatabase.ITable
    '    Dim TableExist As Boolean
    '    Dim pCursor As ESRI.ArcGIS.Geodatabase.ICursor
    '    Dim pRow As ESRI.ArcGIS.Geodatabase.IRow
    '    Dim FieldIndex As Short
    '    Dim booValueFound As Boolean
    '    Dim i As Short

    '    'STEP 1:  get the records from the database -----------------------------------------------
    '    rsCF.MoveFirst()
    '    'End Database stuff

    '    'STEP 2: Raster Values ---------------------------------------------------------------------
    '    'Now Get the RASTER values
    '    ' Get Rasterband from the incoming raster
    '    pRasterCol = pLCRaster
    '    pBand = pRasterCol.Item(0)

    '    'Get the raster table
    '    pBand.HasTable(TableExist)
    '    If Not TableExist Then Exit Function

    '    pTable = pBand.AttributeTable
    '    'Get All rows
    '    pCursor = pTable.Search(Nothing, True)
    '    'Init pRow
    '    pRow = pCursor.NextRow

    '    'Get index of Value Field
    '    FieldIndex = pTable.FindField("Value")

    '    'REMOVED 11/30/2007
    '    'STEP 3: Table values vs. Raster Values Count - if not equal bark --------------------------
    '    '    If pTable.RowCount(Nothing) <> rsCF.RecordCount Then
    '    '        MsgBox "Error: The number of records in your Land Class Table do not match your GRID dataset."
    '    '        Exit Function
    '    '    End If
    '    '
    '    'STEP 4: Create the strings
    '    'Loop through and get all values
    '    Do While Not pRow Is Nothing

    '        booValueFound = False
    '        rsCF.MoveFirst()

    '        For i = 0 To rsCF.RecordCount - 1
    '            If pRow.Value(FieldIndex) = rsCF.Fields("Value").Value Then
    '                booValueFound = True
    '                Select Case rsCF.Fields("W_WL").Value
    '                    Case 0 'Means the current landclass is NOT Water or Wetland, therefore gets a 1
    '                        If strCon = "" Then
    '                            strCon = "Con(([nu_lulc] eq " & pRow.Value(FieldIndex) & "), 1, "
    '                        Else
    '                            strCon = strCon & "Con(([nu_lulc] eq " & pRow.Value(FieldIndex) & "), 1, "
    '                        End If

    '                        If strParens = "" Then
    '                            strParens = "1)"
    '                        Else
    '                            strParens = strParens & ")"
    '                        End If

    '                    Case 1 'Means the current landclass IS Water or Wetland, therefore gets a 0
    '                        If strCon = "" Then
    '                            strCon = "Con(([nu_lulc] eq " & pRow.Value(FieldIndex) & "), 0, "
    '                        Else
    '                            strCon = strCon & "Con(([nu_lulc] eq " & pRow.Value(FieldIndex) & "), 0, "
    '                        End If

    '                        If strParens = "" Then
    '                            strParens = ")"
    '                        Else
    '                            strParens = strParens & ")"
    '                        End If

    '                End Select

    '                'rsCF.MoveNext
    '                'Set pRow = pCursor.NextRow
    '                Exit For

    '            Else
    '                booValueFound = False
    '            End If

    '            rsCF.MoveNext()

    '        Next i

    '        If booValueFound = False Then
    '            MsgBox("Error in MUSLEConstructPondConStatement: Your N-SPECT Land Class Table is missing values found in your landcover GRID dataset.")
    '            ConstructPondConStatement = ""
    '            Exit Function
    '        Else
    '            pRow = pCursor.NextRow
    '            i = 0
    '        End If
    '    Loop

    '    strCompleteCon = strCon & strParens
    '    ConstructPondConStatement = strCompleteCon

    '    'Cleanup:
    '    'Set pLCRaster = Nothing
    '    pRasterCol = Nothing
    '    pBand = Nothing
    '    pTable = Nothing
    '    pCursor = Nothing
    '    pRow = Nothing

    'End Function


    Private Function CalcMUSLE(ByRef strConStatement As String, ByRef strConPondStatement As String) As Boolean
        'Incoming strings: strConStatment: the monster con statement
        'strConPondstatement: the con for the pond stuff
        'Calculates the MUSLE erosion model

        Dim pWeightRaster As MapWinGIS.Grid = Nothing 'STEP 1: Weight Raster
        Dim pWSLengthRaster As MapWinGIS.Grid = Nothing  'STEP 2: Watershed length
        Dim pWSLengthUnitsRaster As MapWinGIS.Grid = Nothing  'STEP 3: Units conversion
        Dim pSlopePRRaster As MapWinGIS.Grid = Nothing  'STEP 4: Average Slope
        Dim pSlopeModRaster As MapWinGIS.Grid = Nothing  'STEP 4a: Mod Slope
        Dim pTemp1LagRaster As MapWinGIS.Grid = Nothing  'STEP 5a: Lag temp1
        Dim pTemp2LagRaster As MapWinGIS.Grid = Nothing  'STEP 5b: Lag temp2
        Dim pTemp3LagRaster As MapWinGIS.Grid = Nothing  'STEP 5c: Lag temp3
        Dim pTemp4LagRaster As MapWinGIS.Grid = Nothing  'STEP 5d: Lag temp4
        Dim pLagRaster As MapWinGIS.Grid = Nothing  'STEP 5e: Lag
        Dim pTOCRaster As MapWinGIS.Grid = Nothing  'STEP 6a: TOC
        Dim pTOCTempRaster As MapWinGIS.Grid = Nothing  'STEP 6b: TOC temp
        Dim pModTOCRaster As MapWinGIS.Grid = Nothing  'STEP 6c: Mod TOC
        Dim pAbPrecipRaster As MapWinGIS.Grid = Nothing  'STEP 7: Abstraction-Precip Ratio
        Dim pLogTOCRaster As MapWinGIS.Grid = Nothing  'STEP 8a: Calc unit peak discharge
        Dim pTempLogTOCRaster As MapWinGIS.Grid = Nothing  'STEP 8b:
        Dim pCZeroRaster As MapWinGIS.Grid = Nothing  'STEP 9: CZero
        Dim pConeRaster As MapWinGIS.Grid = Nothing  'STEP 10: Cone
        Dim pCTwoRaster As MapWinGIS.Grid = Nothing  'STEP 11: CTwo
        Dim pLogQuRaster As MapWinGIS.Grid = Nothing  'STEP 12a: Who knows
        Dim pQuRaster As MapWinGIS.Grid = Nothing 'STep 12b: Who Knows(b)
        Dim pPondFactorRaster As MapWinGIS.Grid = Nothing  'STEP 13: Pond Factor
        Dim pQPRaster As MapWinGIS.Grid = Nothing  'STEP 14: QP factor
        Dim pCFactorRaster As MapWinGIS.Grid = Nothing  'STEP 15: Yee old CFactor Raster
        Dim pHISYTempRaster As MapWinGIS.Grid = Nothing  'STEP 16c: HI Specific temp yield
        Dim pHISYRaster As MapWinGIS.Grid = Nothing  'STEP 16d: HI Specific yield
        Dim pHISYMGRaster As MapWinGIS.Grid = Nothing  'STEP 17b: tons to milligrams, HI Specific
        Dim pPermMUSLERaster As MapWinGIS.Grid = Nothing  'STEP 17c: Local Effects permanent raster
        Dim pTempflowDir3raster As MapWinGIS.Grid = Nothing  'FlowDir temp raster
        Dim pAccSedHIRaster As MapWinGIS.Grid = Nothing  'STEP 19HI: acc sediment
        Dim pTotSedMassHIRaster As MapWinGIS.Grid = Nothing  'STEP 20HI: tot sed mass HI
        Dim pPermTotSedConcHIraster As MapWinGIS.Grid = Nothing  'Permanent MUSLE
        Dim strMUSLE As String


        'String to hold calculations
        Dim strExpression As String = ""
        Const strTitle As String = "Processing MUSLE Calculation..."


        Try
            modProgDialog.ProgDialog("Computing length/distance...", strTitle, 0, 27, 1, 0)
            If modProgDialog.g_boolCancel Then
                'STEP 1: ------------------------------------------------------------------------------------
                'Create weight grid that represents cell length/distance

                Dim weightcalc As New RasterMathCellCalc(AddressOf weightCellCalc)
                RasterMath(g_pFlowDirRaster, Nothing, Nothing, Nothing, Nothing, pWeightRaster, weightcalc)
                'strExpression = "Con([flowdir] eq 2, 1.41421, " & "Con([flowdir] eq 8, 1.41421, " & "Con([flowdir] eq 32, 1.41421, " & "Con([flowdir] eq 128, 1.41421, 1.0))))"

                'END STEP 2: -------------------------------------------------------------------------------
            End If

            modProgDialog.ProgDialog("Calculating Watershed Length...", strTitle, 0, 27, 2, 0)
            If modProgDialog.g_boolCancel Then
                'STEP 2: ------------------------------------------------------------------------------------
                'Calculate Watershed Length

                'TODO: Find access to the Taudem flowlength function and see if it would work for this
                'With pMapAlgebraOp
                '    .BindRaster(g_pFlowDirRaster, "flowdir")
                '    .BindRaster(pWeightRaster, "weight")
                'End With

                'strExpression = "flowlength([flowdir], [weight], upstream)"

                'pWSLengthRaster = pMapAlgebraOp.Execute(strExpression)

                'End STEP 2 ----------------------------------------------------------------------------------
            End If

            modProgDialog.ProgDialog("Converting units...", strTitle, 0, 27, 3, 0)
            If modProgDialog.g_boolCancel Then
                'STEP 3: --------------------------------------------------------------------------------------
                'Convert Metric Units
                Dim wslengthcalc As New RasterMathCellCalc(AddressOf wslengthCellCalc)
                RasterMath(pWSLengthRaster, Nothing, Nothing, Nothing, Nothing, pWSLengthUnitsRaster, wslengthcalc)
                'strExpression = "([cell_wslength] * 3.28084)"

                'END STEP 3: -----------------------------------------------------------------------------------
            End If

            modProgDialog.ProgDialog("Calculating Average Slope...", strTitle, 0, 27, 4, 0)
            If modProgDialog.g_boolCancel Then
                'STEP 4a: ---------------------------------------------------------------------------------------
                'Calculate Average Slope

                'TODO: use my slope moving window to calculate slope as percent rise
                'pMapAlgebraOp.BindRaster(g_pDEMRaster, "dem")

                'strExpression = "slope([dem], percentrise)"

                'pSlopePRRaster = pMapAlgebraOp.Execute(strExpression)

                'END STEP 4a ------------------------------------------------------------------------------------
            End If

            modProgDialog.ProgDialog("Calculating Mod Slope...", strTitle, 0, 27, 5, 0)
            If modProgDialog.g_boolCancel Then
                'STEP 4b: ---------------------------------------------------------------------------------------
                'Calculate modslope
                Dim slpmodcalc As New RasterMathCellCalc(AddressOf slpmodCellCalc)
                RasterMath(pSlopePRRaster, Nothing, Nothing, Nothing, Nothing, pSlopeModRaster, slpmodcalc)
                'strExpression = "Con([slopepr] eq 0, 0.1, [slopepr])"

                'END STEP 4b: -----------------------------------------------------------------------------------
            End If


            modProgDialog.ProgDialog("Calculating Lag...", strTitle, 0, 27, 6, 0)
            If modProgDialog.g_boolCancel Then
                'STEP 5a: ---------------------------------------------------------------------------------------
                'Calculate Lag
                Dim tmp1lagcalc As New RasterMathCellCalc(AddressOf tmp1lagCellCalc)
                RasterMath(pWSLengthUnitsRaster, Nothing, Nothing, Nothing, Nothing, pTemp1LagRaster, tmp1lagcalc)
                'strExpression = "Pow([cell_wslngft], 0.8)"

                'END STEP 5a: ----------------------------------------------------------------------------------
            End If

            modProgDialog.ProgDialog("Checking SCS GRID...", strTitle, 0, 27, 7, 0)
            If modProgDialog.g_boolCancel Then
                'STEP 5b: ---------------------------------------------------------------------------------------
                Dim tmp2lagcalc As New RasterMathCellCalc(AddressOf tmp2lagCellCalc)
                RasterMath(g_pSCS100Raster, Nothing, Nothing, Nothing, Nothing, pTemp2LagRaster, tmp2lagcalc)
                'strExpression = "(1000 / [scsgrid100]) - 9"

                'END STEP 9b: ----------------------------------------------------------------------------------
            End If


            modProgDialog.ProgDialog("Multiplying GRIDs...", strTitle, 0, 27, 8, 0)
            If modProgDialog.g_boolCancel Then
                'STEP 5c: --------------------------------------------------------------------------------------
                Dim tmp3lagcalc As New RasterMathCellCalc(AddressOf tmp3lagCellCalc)
                RasterMath(pTemp2LagRaster, Nothing, Nothing, Nothing, Nothing, pTemp3LagRaster, tmp3lagcalc)
                'strExpression = "Pow([temp4], 0.7)"

                'END STEP 5c: ----------------------------------------------------------------------------------
            End If

            modProgDialog.ProgDialog("Pow([modslope], 0.5...", strTitle, 0, 27, 9, 0)
            If modProgDialog.g_boolCancel Then
                'STEP 5d: --------------------------------------------------------------------------------------
                Dim tmp4lagcalc As New RasterMathCellCalc(AddressOf tmp4lagCellCalc)
                RasterMath(pSlopeModRaster, Nothing, Nothing, Nothing, Nothing, pTemp4LagRaster, tmp4lagcalc)
                'strExpression = "Pow([modslope], 0.5)"

                'END STEP 5d: ----------------------------------------------------------------------------------
            End If

            modProgDialog.ProgDialog("Lag calculation...", strTitle, 0, 30, 27, 0)
            If modProgDialog.g_boolCancel Then
                'STEP 5e: --------------------------------------------------------------------------------------
                'Finally, the lag calculation
                Dim lagcalc As New RasterMathCellCalc(AddressOf lagCellCalc)
                RasterMath(pTemp1LagRaster, pTemp3LagRaster, pTemp4LagRaster, Nothing, Nothing, pLagRaster, lagcalc)
                'strExpression = "([temp3] * [temp5]) / (1900 * [temp6])"

                'STEP 5e: ---------------------------------------------------------------------------------------
            End If


            modProgDialog.ProgDialog("Calculating time of concentration...", strTitle, 0, 27, 11, 0)
            If modProgDialog.g_boolCancel Then

                'STEP 6a: ---------------------------------------------------------------------------------------
                'Calculate the time of concentration
                Dim toccalc As New RasterMathCellCalc(AddressOf tocCellCalc)
                RasterMath(pLagRaster, Nothing, Nothing, Nothing, Nothing, pTOCRaster, toccalc)
                'strExpression = "[lag] / 0.6"
                'END STEP 6a: -----------------------------------------------------------------------------------


                'STEP 6b: --------------------------------------------------------------------------------------
                Dim toctmpcalc As New RasterMathCellCalc(AddressOf toctmpCellCalc)
                RasterMath(pTOCRaster, Nothing, Nothing, Nothing, Nothing, pTOCTempRaster, toctmpcalc)
                'strExpression = "Con([toc] lt 0.1, 0.1, [toc])"

                'END STEP 6b: ----------------------------------------------------------------------------------


                'STEP 6c: --------------------------------------------------------------------------------------
                Dim modtoccalc As New RasterMathCellCalc(AddressOf modtocCellCalc)
                RasterMath(pTOCTempRaster, Nothing, Nothing, Nothing, Nothing, pModTOCRaster, modtoccalc)
                'strExpression = "Con([temp7] gt 10, 10, [temp7])"

                'END STEP 6c: ----------------------------------------------------------------------------------

            End If

            modProgDialog.ProgDialog("Abstraction Precipitation Ratio...", strTitle, 0, 27, 12, 0)
            If modProgDialog.g_boolCancel Then
                'STEP 7: ---------------------------------------------------------------------------------------
                'Find the Abstraction Precipitation Ratio
                Dim abprecipcalc As New RasterMathCellCalc(AddressOf abprecipCellCalc)
                RasterMath(g_pAbstractRaster, g_pPrecipRaster, Nothing, Nothing, Nothing, pAbPrecipRaster, abprecipcalc)
                'strExpression = "[abstract] / [rain]"

                'END STEP 7: ----------------------------------------------------------------------------------
            End If

            modProgDialog.ProgDialog("Calculating the peak unit discharge...", strTitle, 0, 27, 13, 0)
            If modProgDialog.g_boolCancel Then
                'STEP 8a: ----------------------------------------------------------------------------------------
                'Calculate the unit peak discharge
                Dim logtoccalc As New RasterMathCellCalc(AddressOf logtocCellCalc)
                RasterMath(pModTOCRaster, Nothing, Nothing, Nothing, Nothing, pLogTOCRaster, logtoccalc)
                'strExpression = "log10([modtoc])"

                'END STEP 8a: -----------------------------------------------------------------------------------


                'STEP 8b: ---------------------------------------------------------------------------------------
                '2nd part of it
                Dim tmplogtoccalc As New RasterMathCellCalc(AddressOf tmplogtocCellCalc)
                RasterMath(pLogTOCRaster, Nothing, Nothing, Nothing, Nothing, pTempLogTOCRaster, tmplogtoccalc)
                'strExpression = "Pow([logtoc], 2)"

                'END STEP 8b: -----------------------------------------------------------------------------------


            End If

            modProgDialog.ProgDialog("Creating C-Zero GRID...", strTitle, 0, 27, 14, 0)
            If modProgDialog.g_boolCancel Then
                'STEP 9: ----------------------------------------------------------------------------------------
                'CZERO GRID
                Dim c0calc As RasterMathCellCalc = Nothing
                Select Case g_intPrecipType
                    Case 0
                        c0calc = New RasterMathCellCalc(AddressOf c0CellCalc0)
                    Case 1
                        c0calc = New RasterMathCellCalc(AddressOf c0CellCalc1)
                    Case 2
                        c0calc = New RasterMathCellCalc(AddressOf c0CellCalc2)
                    Case 3
                        c0calc = New RasterMathCellCalc(AddressOf c0CellCalc3)
                End Select
                RasterMath(pAbPrecipRaster, Nothing, Nothing, Nothing, Nothing, pCZeroRaster, c0calc)
                'Call to clsPrecipType(called clsprecip here init above) to get string based on Precip Type
                'strExpression = clsPrecip.CZero(g_intPrecipType)

                'END STEP 9: ------------------------------------------------------------------------------------
            End If


            modProgDialog.ProgDialog("Creating C1 GRID...", strTitle, 0, 27, 15, 0)
            If modProgDialog.g_boolCancel Then
                'STEP 10: ---------------------------------------------------------------------------------------
                'CONE grid
                Dim c1calc As RasterMathCellCalc = Nothing
                Select Case g_intPrecipType
                    Case 0
                        c1calc = New RasterMathCellCalc(AddressOf c1CellCalc0)
                    Case 1
                        c1calc = New RasterMathCellCalc(AddressOf c1CellCalc1)
                    Case 2
                        c1calc = New RasterMathCellCalc(AddressOf c1CellCalc2)
                    Case 3
                        c1calc = New RasterMathCellCalc(AddressOf c1CellCalc3)
                End Select
                RasterMath(pAbPrecipRaster, Nothing, Nothing, Nothing, Nothing, pConeRaster, c1calc)
                'strExpression = clsPrecip.Cone(g_intPrecipType)

                'END STEP 10 ------------------------------------------------------------------------------------
            End If


            modProgDialog.ProgDialog("Creating C2 GRID...", strTitle, 0, 27, 16, 0)
            If modProgDialog.g_boolCancel Then
                'STEP 11: ---------------------------------------------------------------------------------------
                'CTwo GRID
                Dim c2calc As RasterMathCellCalc = Nothing
                Select Case g_intPrecipType
                    Case 0
                        c2calc = New RasterMathCellCalc(AddressOf c2CellCalc0)
                    Case 1
                        c2calc = New RasterMathCellCalc(AddressOf c2CellCalc1)
                    Case 2
                        c2calc = New RasterMathCellCalc(AddressOf c2CellCalc2)
                    Case 3
                        c2calc = New RasterMathCellCalc(AddressOf c2CellCalc3)
                End Select
                RasterMath(pAbPrecipRaster, Nothing, Nothing, Nothing, Nothing, pCTwoRaster, c2calc)
                'strExpression = clsPrecip.CTwo(g_intPrecipType)

                'END STEP 11: ------------------------------------------------------------------------------------
            End If


            modProgDialog.ProgDialog("More math...", strTitle, 0, 27, 17, 0)
            If modProgDialog.g_boolCancel Then
                'STEP 12a: ----------------------------------------------------------------------------------------
                'Logqu
                Dim logqucalc As New RasterMathCellCalc(AddressOf logquCellCalc)
                RasterMath(pCZeroRaster, pConeRaster, pLogTOCRaster, pCTwoRaster, pTempLogTOCRaster, pLogQuRaster, logqucalc)
                'strExpression = "[czero] + ([cone] * [logtoc]) + ([ctwo] * [temp8])"

                'END STEP 12a: -------------------------------------------------------------------------------------


                'STEP 12b: -----------------------------------------------------------------------------------------
                Dim qucalc As New RasterMathCellCalc(AddressOf quCellCalc)
                RasterMath(pLogQuRaster, Nothing, Nothing, Nothing, Nothing, pQuRaster, qucalc)
                'strExpression = "Pow(10, [logqu])"

                'END STEP 12b: -------------------------------------------------------------------------------------
            End If


            modProgDialog.ProgDialog("Creating Pond Factor GRID...", strTitle, 0, 27, 16, 0)
            If modProgDialog.g_boolCancel Then
                'STEP 13: ------------------------------------------------------------------------------------------
                'Create pond factor grid
                ReDim _pondpicks(strConPondStatement.Split(",").Length)
                _pondpicks = strConPondStatement.Split(",")

                Dim pondfactcalc As New RasterMathCellCalc(AddressOf pondfactCellCalc)
                RasterMath(g_LandCoverRaster, Nothing, Nothing, Nothing, Nothing, pPondFactorRaster, pondfactcalc)
                'strExpression = strConPondStatement

                'END STEP 13: ---------------------------------------------------------------------------------------
            End If


            modProgDialog.ProgDialog("Calculating peak discharge; cubic feet per second...", strTitle, 0, 27, 17, 0)
            If modProgDialog.g_boolCancel Then
                'STEP 14: -------------------------------------------------------------------------------------------
                'Calculate peak discharge: cubic feet per second
                Dim qpcalc As New RasterMathCellCalc(AddressOf qpCellCalc)
                RasterMath(pQuRaster, g_pCellAreaSqMiRaster, g_pRunoffInchRaster, pPondFactorRaster, Nothing, pQPRaster, qpcalc)
                'strExpression = "[qu] * [cellarea_sqmi] * [runoff_in] * [pondfact]"

                'END STEP 14: ----------------------------------------------------------------------------------------
            End If


            modProgDialog.ProgDialog("Creating cover factor GRID...", strTitle, 0, 27, 18, 0)
            If modProgDialog.g_boolCancel Then
                'STEP 15: --------------------------------------------------------------------------------------------
                'Cover Factor GRID
                ReDim _picks(strConStatement.Split(",").Length)
                _picks = strConStatement.Split(",")

                Dim cfactcalc As New RasterMathCellCalc(AddressOf cfactCellCalc)
                RasterMath(g_LandCoverRaster, Nothing, Nothing, Nothing, Nothing, pCFactorRaster, cfactcalc)

                'END STEP 15 -----------------------------------------------------------------------------------------
            End If

            modProgDialog.ProgDialog("Sediment Yield...", strTitle, 0, 27, 19, 0)
            If modProgDialog.g_boolCancel Then
                'STEP 16 c: -------------------------------------------------------------------------------------------
                'Temp Sediment Yield: Note: MUSLE Exponent inserted to allow for global use, cuz other's will use
                Dim hisytmpcalc As New RasterMathCellCalc(AddressOf hisytmpCellCalc)
                RasterMath(g_pRunoffAFRaster, pQPRaster, Nothing, Nothing, Nothing, pHISYTempRaster, hisytmpcalc)
                'strExpression = "Pow(([runoff_af] * [qp]), " & _dblMUSLEExp & ")"

                'END STEP 16c: ------------------------------------------------------------------------------------------------------
            End If


            modProgDialog.ProgDialog("Sediment Yield...", strTitle, 0, 30, 27, 0)
            If modProgDialog.g_boolCancel Then
                'STEP 16d: ---------------------------------------------------------------------------------------------
                'Sediment Yield: Note m_dblMUSLEVal now in for universal
                Dim hisycalc As New RasterMathCellCalc(AddressOf hisyCellCalc)
                RasterMath(pCFactorRaster, g_KFactorRaster, g_pLSRaster, pHISYTempRaster, Nothing, pHISYRaster, hisycalc)
                'strExpression = _dblMUSLEVal & " * ([cfactor] * [kfactor] * [lsfactor] * [temp9])"

                'END STEP 16b: ----------------------------------------------------------------------------------------
            End If


            modProgDialog.ProgDialog("Converting tons to milligrams...", strTitle, 0, 27, 21, 0)
            If modProgDialog.g_boolCancel Then
                Dim hisymgrcalc As New RasterMathCellCalc(AddressOf hisymgrCellCalc)
                RasterMath(pHISYRaster, Nothing, Nothing, Nothing, Nothing, pHISYMGRaster, hisymgrcalc)
                'strExpression = "[sy] * 907.184740"
                'END STEP 17b: ----------------------------------------------------------------------------------------
            End If

            
            Dim pClipMusleRaster As MapWinGIS.Grid
            If g_booLocalEffects Then

                modProgDialog.ProgDialog("Creating data layer for local effects...", strTitle, 0, 27, 27, 0)
                If modProgDialog.g_boolCancel Then

                    strMUSLE = modUtil.GetUniqueName("locmusle", g_strWorkspace, ".tif")
                    'Added 7/23/04 to account for clip by selected polys functionality
                    If g_booSelectedPolys Then
                        'TODO
                        'pClipMusleRaster = modUtil.ClipBySelectedPoly(pHISYMGRaster, g_pSelectedPolyClip, pEnv)
                        'pPermMUSLERaster = modUtil.ReturnPermanentRaster(pClipMusleRaster, pEnv.OutWorkspace.PathName, strMUSLE)
                    Else
                        'pPermMUSLERaster = modUtil.ReturnPermanentRaster(pHISYMGRaster, pEnv.OutWorkspace.PathName, strMUSLE)
                    End If

                    'metadata time
                    g_dicMetadata.Add("MUSLE Local Effects (mg)", _strMusleMetadata)

                    'TODO
                    'pMUSLERasterLocLayer = modUtil.ReturnRasterLayer((frmPrj.m_App), pPermMUSLERaster, "MUSLE Local Effects (mg)")
                    'pMUSLERasterLocLayer.Renderer = modUtil.ReturnRasterStretchColorRampRender(pMUSLERasterLocLayer, "Brown")
                    'pMUSLERasterLocLayer.Visible = False
                    'g_pGroupLayer.Add(pMUSLERasterLocLayer)

                    CalcMUSLE = True
                    modProgDialog.KillDialog()
                    Exit Function

                End If

            End If


            modProgDialog.ProgDialog("Calculating the accumulated sediment...", strTitle, 0, 27, 23, 0)
            If modProgDialog.g_boolCancel Then
                'TODO: replace this with taudem flow accum
                'pTempflowDir3raster = modUtil.ReturnRaster(g_strFlowDirFilename)
                'pFlowAccumOp1 = New ESRI.ArcGIS.SpatialAnalyst.RasterHydrologyOp

                'pFlowDirRDS1 = pTempflowDir3raster
                'pHISYMGRDS = pHISYMGRaster

                'pEnv = pFlowAccumOp1
                'pOutRDS1 = pFlowAccumOp1.FlowAccumulation(pFlowDirRDS1, pHISYMGRDS)

                'pAccSedHIRaster = pOutRDS1

            End If

            modProgDialog.ProgDialog("Calculating Total Sediment Mass...", strTitle, 0, 27, 24, 0)
            If modProgDialog.g_boolCancel Then
                'STEP 20HI: -----------------------------------------------------------------------------------------------
                'Total Sediment Mass
                Dim totsedmasscalc As New RasterMathCellCalc(AddressOf totsedmassCellCalc)
                RasterMath(pHISYMGRaster, pAccSedHIRaster, Nothing, Nothing, Nothing, pTotSedMassHIRaster, totsedmasscalc)
                'old / 10000000
                'strExpression = "[sy_mg_HI] + [accsed_HI]"

                'END STEP 20HI: -------------------------------------------------------------------------------------------
            End If


            modProgDialog.ProgDialog("Adding Sediment Mass to Group Layer...", strTitle, 0, 27, 25, 0)
            Dim pClipMusleMassRaster As MapWinGIS.Grid
            If modProgDialog.g_boolCancel Then
                'STEP 21: Created the Sediment Mass Raster layer and add to Group Layer -----------------------------------
                'Get a unique name for MUSLE and return the permanently made raster
                strMUSLE = modUtil.GetUniqueName("MUSLEmass", g_strWorkspace, ".tif")

                'Clip to selected polys if chosen
                If g_booSelectedPolys Then
                    'TODO
                    'pClipMusleMassRaster = modUtil.ClipBySelectedPoly(pTotSedMassHIRaster, g_pSelectedPolyClip, pEnv)
                    'pPermTotSedConcHIraster = modUtil.ReturnPermanentRaster(pClipMusleMassRaster, pEnv.OutWorkspace.PathName, strMUSLE)
                Else
                    'pPermTotSedConcHIraster = modUtil.ReturnPermanentRaster(pTotSedMassHIRaster, pEnv.OutWorkspace.PathName, strMUSLE)
                End If


                'Metadata:
                g_dicMetadata.Add("MUSLE Sediment Mass (kg)", _strMusleMetadata)

                'Now create the MUSLE layer
                'TODO
                'pMUSLERasterLayer = modUtil.ReturnRasterLayer((frmPrj.m_App), pPermTotSedConcHIraster, "MUSLE Sediment Mass (kg)")
                'pMUSLERasterLayer.Renderer = modUtil.ReturnRasterStretchColorRampRender(pMUSLERasterLayer, "Brown")
                'pMUSLERasterLayer.Visible = False
                'g_pGroupLayer.Add(pMUSLERasterLayer)

                'end STEP 21: Created the Sediment Mass Raster layer and add to Group Layer -----------------------------------

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
    Private Function weightCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'strExpression = "Con([flowdir] eq 2, 1.41421, " & "Con([flowdir] eq 8, 1.41421, " & "Con([flowdir] eq 32, 1.41421, " & "Con([flowdir] eq 128, 1.41421, 1.0))))"
        'Con(
        '  [flowdir] eq 2, 
        '  1.41421, 
        '  Con(
        '    [flowdir] eq 8
        '    1.41421
        '    Con(
        '      [flowdir] eq 32
        '      1.41421
        '      Con(
        '        [flowdir] eq 128
        '        1.41421
        '        1.0))))
        'TODO: add the taudem equivalents
        If Input1 = 2 Or Input1 = 8 Or Input1 = 32 Or Input1 = 128 Then
            Return 1.41421
        Else
            Return 1
        End If

    End Function

    Private Function wslengthCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'strExpression = "([cell_wslength] * 3.28084)"
        Return Input1 * 3.28084
    End Function

    Private Function slpmodCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'strExpression = "Con([slopepr] eq 0, 0.1, [slopepr])"
        If Input1 = 0 Then
            Return 0.1
        Else
            Return Input1
        End If
    End Function

    Private Function tmp1lagCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'strExpression = "Pow([cell_wslngft], 0.8)"
        Return Math.Pow(Input1, 0.8)
    End Function

    Private Function tmp2lagCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'strExpression = "(1000 / [scsgrid100]) - 9"
        Return (1000 / Input1) - 9
    End Function

    Private Function tmp3lagCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'strExpression = "Pow([temp4], 0.7)"
        Return Math.Pow(Input1, 0.7)
    End Function

    Private Function tmp4lagCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'strExpression = "Pow([modslope], 0.5)"
        Return Math.Pow(Input1, 0.5)
    End Function

    Private Function lagCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'strExpression = "([temp3] * [temp5]) / (1900 * [temp6])"
        Return (Input1 * Input2) / (1900 * Input3)
    End Function

    Private Function tocCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'strExpression = "[lag] / 0.6"
        Return Input1 / 0.6
    End Function

    Private Function toctmpCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'strExpression = "Con([toc] lt 0.1, 0.1, [toc])"
        If Input1 < 0.1 Then
            Return 0.1
        Else
            Return Input1
        End If
    End Function

    Private Function modtocCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'strExpression = "Con([temp7] gt 10, 10, [temp7])"
        If Input1 > 10 Then
            Return 10
        Else
            Return Input1
        End If
    End Function

    Private Function abprecipCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'strExpression = "[abstract] / [rain]"
        Return Input1 / Input2
    End Function

    Private Function logtocCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'strExpression = "log10([modtoc])"
        Return Math.Log10(Input1)
    End Function

    Private Function tmplogtocCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'strExpression = "Pow([logtoc], 2)"
        Return Math.Pow(Input1, 2)
    End Function

    Private Function c0CellCalc0(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
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

    Private Function c0CellCalc1(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
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

    Private Function c0CellCalc2(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
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

    Private Function c0CellCalc3(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
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

    Private Function c1CellCalc0(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
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

    Private Function c1CellCalc1(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
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

    Private Function c1CellCalc2(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
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

    Private Function c1CellCalc3(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
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

    Private Function c2CellCalc0(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
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

    Private Function c2CellCalc1(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
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

    Private Function c2CellCalc2(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
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

    Private Function c2CellCalc3(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
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

    Private Function logquCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'strExpression = "[czero] + ([cone] * [logtoc]) + ([ctwo] * [temp8])"
        Return Input1 + (Input2 * Input3) + (Input4 * Input5)
    End Function

    Private Function quCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'strExpression = "Pow(10, [logqu])"
        Return Math.Pow(10, Input1)
    End Function

    Private Function pondfactCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        For i As Integer = 0 To _pondpicks.Length - 1
            If Input1 = i + 1 Then
                Return _pondpicks(i)
            End If
        Next
    End Function

    Private Function qpCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'strExpression = "[qu] * [cellarea_sqmi] * [runoff_in] * [pondfact]"
        Return Input1 * Input2 * Input3 * Input4
    End Function

    Private Function cfactCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        For i As Integer = 0 To _picks.Length - 1
            If Input1 = i + 1 Then
                Return _picks(i)
            End If
        Next
    End Function

    Private Function hisytmpCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'strExpression = "Pow(([runoff_af] * [qp]), " & _dblMUSLEExp & ")"
        Return Math.Pow((Input1 * Input2), _dblMUSLEExp)
    End Function

    Private Function hisyCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'strExpression = _dblMUSLEVal & " * ([cfactor] * [kfactor] * [lsfactor] * [temp9])"
        Return _dblMUSLEExp * (Input1 * Input2 * Input3 * Input4)
    End Function

    Private Function hisymgrCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'strExpression = "[sy] * 907.184740"
        Return Input1 * 907.18474
    End Function

    Private Function totsedmassCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'old / 10000000
        'strExpression = "[sy_mg_HI] + [accsed_HI]"
        Return Input1 + Input2
    End Function


#End Region

End Module