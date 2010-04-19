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
        Dim strError As String

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

        Dim pWeightRaster As MapWinGIS.Grid 'STEP 1: Weight Raster
        Dim pWSLengthRaster As MapWinGIS.Grid 'STEP 2: Watershed length
        Dim pWSLengthUnitsRaster As MapWinGIS.Grid 'STEP 3: Units conversion
        Dim pSlopePRRaster As MapWinGIS.Grid 'STEP 4: Average Slope
        Dim pSlopeModRaster As MapWinGIS.Grid 'STEP 4a: Mod Slope
        Dim pTemp1LagRaster As MapWinGIS.Grid 'STEP 5a: Lag temp1
        Dim pTemp2LagRaster As MapWinGIS.Grid 'STEP 5b: Lag temp2
        Dim pTemp3LagRaster As MapWinGIS.Grid 'STEP 5c: Lag temp3
        Dim pTemp4LagRaster As MapWinGIS.Grid 'STEP 5d: Lag temp4
        Dim pLagRaster As MapWinGIS.Grid 'STEP 5e: Lag
        Dim pTOCRaster As MapWinGIS.Grid 'STEP 6a: TOC
        Dim pTOCTempRaster As MapWinGIS.Grid 'STEP 6b: TOC temp
        Dim pModTOCRaster As MapWinGIS.Grid 'STEP 6c: Mod TOC
        Dim pAbPrecipRaster As MapWinGIS.Grid 'STEP 7: Abstraction-Precip Ratio
        Dim pLogTOCRaster As MapWinGIS.Grid 'STEP 8a: Calc unit peak discharge
        Dim pTempLogTOCRaster As MapWinGIS.Grid 'STEP 8b:
        Dim pCZeroRaster As MapWinGIS.Grid 'STEP 9: CZero
        Dim pConeRaster As MapWinGIS.Grid 'STEP 10: Cone
        Dim pCTwoRaster As MapWinGIS.Grid 'STEP 11: CTwo
        Dim pLogQuRaster As MapWinGIS.Grid 'STEP 12a: Who knows
        Dim pQuRaster As MapWinGIS.Grid 'STep 12b: Who Knows(b)
        Dim pPondFactorRaster As MapWinGIS.Grid 'STEP 13: Pond Factor
        Dim pQPRaster As MapWinGIS.Grid 'STEP 14: QP factor
        Dim pCFactorRaster As MapWinGIS.Grid 'STEP 15: Yee old CFactor Raster
        Dim pSYTempRaster As MapWinGIS.Grid 'Step 16a Temp yield
        Dim pSYRaster As MapWinGIS.Grid 'Step 16b Yield
        Dim pHISYTempRaster As MapWinGIS.Grid 'STEP 16c: HI Specific temp yield
        Dim pHISYRaster As MapWinGIS.Grid 'STEP 16d: HI Specific yield
        'Dim pSYMGRaster As IRaster                  'STEP 17a: tons to milligrams
        Dim pHISYMGRaster As MapWinGIS.Grid 'STEP 17b: tons to milligrams, HI Specific
        Dim pPermMUSLERaster As MapWinGIS.Grid 'STEP 17c: Local Effects permanent raster
        Dim pTempFlowDir1Raster As MapWinGIS.Grid 'flowDir Temp Raster
        'Dim pTempFlowDir2Raster As IRaster          'FlowDir temp raster
        Dim pTempflowDir3raster As MapWinGIS.Grid 'FlowDir temp raster
        Dim pLiterRunoffRaster As MapWinGIS.Grid 'STEP 18: Runoff to liters
        'Dim pAccSedRaster As IRaster                'STEP 19: Acc sediment
        Dim pAccSedHIRaster As MapWinGIS.Grid 'STEP 19HI: acc sediment
        Dim pAccRunLiterRaster As MapWinGIS.Grid 'STEP 19a: Acc runoff liter
        Dim pTotSedMassRaster As MapWinGIS.Grid 'STEP 20: Tot sed mass
        Dim pTotSedMassHIRaster As MapWinGIS.Grid 'STEP 20HI: tot sed mass HI
        Dim pTotSedTempRaster As MapWinGIS.Grid 'STEP 21: First step
        Dim pPermTotSedConcHIraster As MapWinGIS.Grid 'Permanent MUSLE
        Dim pSedConcRaster As MapWinGIS.Grid 'Sed Concentration
        Dim pPermSedConcRaster As MapWinGIS.Grid 'Permanent Sed Concentration
        Dim strMUSLE As String


        'String to hold calculations
        Dim strExpression As String
        Const strTitle As String = "Processing MUSLE Calculation..."


        Try
            modProgDialog.ProgDialog("Computing length/distance...", strTitle, 0, 27, 1, 0)
            If modProgDialog.g_boolCancel Then
                'STEP 1: ------------------------------------------------------------------------------------
                'Create weight grid that represents cell length/distance

                Dim weightcalc As New RasterMathCellCalc(AddressOf weightCellCalc)
                pWeightRaster = Nothing
                RasterMath(g_pFlowDirRaster, Nothing, Nothing, Nothing, Nothing, pWeightRaster, weightcalc)
                'strExpression = "Con([flowdir] eq 2, 1.41421, " & "Con([flowdir] eq 8, 1.41421, " & "Con([flowdir] eq 32, 1.41421, " & "Con([flowdir] eq 128, 1.41421, 1.0))))"

                'END STEP 2: -------------------------------------------------------------------------------
            End If

            'modProgDialog.ProgDialog("Calculating Watershed Length...", strTitle, 0, 27, 2, 0)
            'If modProgDialog.g_boolCancel Then
            '    'STEP 2: ------------------------------------------------------------------------------------
            '    'Calculate Watershed Length
            '    With pMapAlgebraOp
            '        .BindRaster(g_pFlowDirRaster, "flowdir")
            '        .BindRaster(pWeightRaster, "weight")
            '    End With

            '    strExpression = "flowlength([flowdir], [weight], upstream)"

            '    pWSLengthRaster = pMapAlgebraOp.Execute(strExpression)

            '    With pMapAlgebraOp
            '        .UnbindRaster("flowdir")
            '        .UnbindRaster("weight")
            '    End With
            '    'End STEP 2 ----------------------------------------------------------------------------------
            'End If

            'modProgDialog.ProgDialog("Converting units...", strTitle, 0, 27, 3, 0)
            'If modProgDialog.g_boolCancel Then
            '    'STEP 3: --------------------------------------------------------------------------------------
            '    'Convert Metric Units
            '    pMapAlgebraOp.BindRaster(pWSLengthRaster, "cell_wslength")

            '    strExpression = "([cell_wslength] * 3.28084)"

            '    pWSLengthUnitsRaster = pMapAlgebraOp.Execute(strExpression)

            '    pMapAlgebraOp.UnbindRaster("cell_wslength")
            '    'END STEP 3: -----------------------------------------------------------------------------------
            'End If

            'modProgDialog.ProgDialog("Calculating Average Slope...", strTitle, 0, 27, 4, 0)
            'If modProgDialog.g_boolCancel Then
            '    'STEP 4a: ---------------------------------------------------------------------------------------
            '    'Calculate Average Slope
            '    pMapAlgebraOp.BindRaster(g_pDEMRaster, "dem")

            '    strExpression = "slope([dem], percentrise)"

            '    pSlopePRRaster = pMapAlgebraOp.Execute(strExpression)

            '    pMapAlgebraOp.UnbindRaster("dem")
            '    'END STEP 4a ------------------------------------------------------------------------------------
            'End If

            'modProgDialog.ProgDialog("Calculating Mod Slope...", strTitle, 0, 27, 5, 0)
            'If modProgDialog.g_boolCancel Then
            '    'STEP 4b: ---------------------------------------------------------------------------------------
            '    'Calculate modslope
            '    pMapAlgebraOp.BindRaster(pSlopePRRaster, "slopepr")

            '    strExpression = "Con([slopepr] eq 0, 0.1, [slopepr])"

            '    pSlopeModRaster = pMapAlgebraOp.Execute(strExpression)

            '    pMapAlgebraOp.UnbindRaster("slopepr")
            '    'END STEP 4b: -----------------------------------------------------------------------------------
            'End If


            'modProgDialog.ProgDialog("Calculating Lag...", strTitle, 0, 27, 6, 0)
            'If modProgDialog.g_boolCancel Then
            '    'STEP 5a: ---------------------------------------------------------------------------------------
            '    'Calculate Lag
            '    pMapAlgebraOp.BindRaster(pWSLengthUnitsRaster, "cell_wslngft")

            '    strExpression = "Pow([cell_wslngft], 0.8)"

            '    pTemp1LagRaster = pMapAlgebraOp.Execute(strExpression)

            '    pMapAlgebraOp.UnbindRaster("cell_wslngft")
            '    'END STEP 5a: ----------------------------------------------------------------------------------
            'End If

            'modProgDialog.ProgDialog("Checking SCS GRID...", strTitle, 0, 27, 7, 0)
            'If modProgDialog.g_boolCancel Then
            '    'STEP 5b: ---------------------------------------------------------------------------------------
            '    pMapAlgebraOp.BindRaster(g_pSCS100Raster, "scsgrid100")

            '    strExpression = "(1000 / [scsgrid100]) - 9"

            '    pTemp2LagRaster = pMapAlgebraOp.Execute(strExpression)

            '    pMapAlgebraOp.UnbindRaster("scsgrid100")
            '    'END STEP 9b: ----------------------------------------------------------------------------------
            'End If


            'modProgDialog.ProgDialog("Multiplying GRIDs...", strTitle, 0, 27, 8, 0)
            'If modProgDialog.g_boolCancel Then
            '    'STEP 5c: --------------------------------------------------------------------------------------
            '    pMapAlgebraOp.BindRaster(pTemp2LagRaster, "temp4")

            '    strExpression = "Pow([temp4], 0.7)"

            '    pTemp3LagRaster = pMapAlgebraOp.Execute(strExpression)

            '    pMapAlgebraOp.UnbindRaster("temp4")
            '    'END STEP 5c: ----------------------------------------------------------------------------------
            'End If

            'modProgDialog.ProgDialog("Pow([modslope], 0.5...", strTitle, 0, 27, 9, 0)
            'If modProgDialog.g_boolCancel Then
            '    'STEP 5d: --------------------------------------------------------------------------------------
            '    pMapAlgebraOp.BindRaster(pSlopeModRaster, "modslope")

            '    strExpression = "Pow([modslope], 0.5)"

            '    pTemp4LagRaster = pMapAlgebraOp.Execute(strExpression)

            '    pMapAlgebraOp.UnbindRaster("modslope")
            '    'END STEP 5d: ----------------------------------------------------------------------------------
            'End If

            'modProgDialog.ProgDialog("Lag calculation...", strTitle, 0, 30, 27, 0)
            'If modProgDialog.g_boolCancel Then
            '    'STEP 5e: --------------------------------------------------------------------------------------
            '    'Finally, the lag calculation
            '    With pMapAlgebraOp
            '        .BindRaster(pTemp1LagRaster, "temp3")
            '        .BindRaster(pTemp3LagRaster, "temp5")
            '        .BindRaster(pTemp4LagRaster, "temp6")
            '    End With

            '    strExpression = "([temp3] * [temp5]) / (1900 * [temp6])"

            '    pLagRaster = pMapAlgebraOp.Execute(strExpression)

            '    With pMapAlgebraOp
            '        .UnbindRaster("temp3")
            '        .UnbindRaster("temp5")
            '        .UnbindRaster("temp6")
            '    End With
            '    'STEP 5e: ---------------------------------------------------------------------------------------
            'End If


            'modProgDialog.ProgDialog("Calculating time of concentration...", strTitle, 0, 27, 11, 0)
            'If modProgDialog.g_boolCancel Then

            '    'STEP 6a: ---------------------------------------------------------------------------------------
            '    'Calculate the time of concentration
            '    pMapAlgebraOp.BindRaster(pLagRaster, "lag")

            '    strExpression = "[lag] / 0.6"

            '    pTOCRaster = pMapAlgebraOp.Execute(strExpression)

            '    pMapAlgebraOp.UnbindRaster("lag")
            '    'END STEP 6a: -----------------------------------------------------------------------------------


            '    'STEP 6b: --------------------------------------------------------------------------------------
            '    pMapAlgebraOp.BindRaster(pTOCRaster, "toc")

            '    strExpression = "Con([toc] lt 0.1, 0.1, [toc])"

            '    pTOCTempRaster = pMapAlgebraOp.Execute(strExpression)

            '    pMapAlgebraOp.UnbindRaster("toc")
            '    'END STEP 6b: ----------------------------------------------------------------------------------

            
            '    'STEP 6c: --------------------------------------------------------------------------------------
            '    pMapAlgebraOp.BindRaster(pTOCTempRaster, "temp7")

            '    strExpression = "Con([temp7] gt 10, 10, [temp7])"

            '    pModTOCRaster = pMapAlgebraOp.Execute(strExpression)

            '    pMapAlgebraOp.UnbindRaster("temp7")
            '    'END STEP 6c: ----------------------------------------------------------------------------------

            'End If

            'modProgDialog.ProgDialog("Abstraction Precipitation Ratio...", strTitle, 0, 27, 12, 0)
            'If modProgDialog.g_boolCancel Then
            '    'STEP 7: ---------------------------------------------------------------------------------------
            '    'Find the Abstraction Precipitation Ratio
            '    With pMapAlgebraOp
            '        .BindRaster(g_pAbstractRaster, "abstract")
            '        .BindRaster(g_pPrecipRaster, "rain")
            '    End With

            '    strExpression = "[abstract] / [rain]"

            '    pAbPrecipRaster = pMapAlgebraOp.Execute(strExpression)

            '    With pMapAlgebraOp
            '        .UnbindRaster("abstract")
            '        .UnbindRaster("rain")
            '    End With
            '    'END STEP 7: ----------------------------------------------------------------------------------
            'End If

            'modProgDialog.ProgDialog("Calculating the peak unit discharge...", strTitle, 0, 27, 13, 0)
            'If modProgDialog.g_boolCancel Then
            '    'STEP 8a: ----------------------------------------------------------------------------------------
            '    'Calculate the unit peak discharge
            '    pMapAlgebraOp.BindRaster(pModTOCRaster, "modtoc")

            '    strExpression = "log10([modtoc])"

            '    pLogTOCRaster = pMapAlgebraOp.Execute(strExpression)

            '    pMapAlgebraOp.UnbindRaster("modtoc")
            '    'END STEP 8a: -----------------------------------------------------------------------------------


            '    'STEP 8b: ---------------------------------------------------------------------------------------
            '    '2nd part of it
            '    pMapAlgebraOp.BindRaster(pLogTOCRaster, "logtoc")

            '    strExpression = "Pow([logtoc], 2)"

            '    pTempLogTOCRaster = pMapAlgebraOp.Execute(strExpression)

            '    pMapAlgebraOp.UnbindRaster("logtoc")
            '    'END STEP 8b: -----------------------------------------------------------------------------------


            'End If

            'modProgDialog.ProgDialog("Creating C-Zero GRID...", strTitle, 0, 27, 14, 0)
            'If modProgDialog.g_boolCancel Then
            '    'STEP 9: ----------------------------------------------------------------------------------------
            '    'CZERO GRID
            '    pMapAlgebraOp.BindRaster(pAbPrecipRaster, "ip")

            '    'Call to clsPrecipType(called clsprecip here init above) to get string based on Precip Type
            '    'g_intPrecipType
            '    strExpression = clsPrecip.CZero(g_intPrecipType)

            '    pCZeroRaster = pMapAlgebraOp.Execute(strExpression)

            '    pMapAlgebraOp.UnbindRaster("ip")
            '    'END STEP 9: ------------------------------------------------------------------------------------
            'End If


            'modProgDialog.ProgDialog("Creating Cone GRID...", strTitle, 0, 27, 15, 0)
            'If modProgDialog.g_boolCancel Then
            '    'STEP 10: ---------------------------------------------------------------------------------------
            '    'CONE grid
            '    pMapAlgebraOp.BindRaster(pAbPrecipRaster, "ip")

            '    strExpression = clsPrecip.Cone(g_intPrecipType)

            '    pConeRaster = pMapAlgebraOp.Execute(strExpression)

            '    pMapAlgebraOp.UnbindRaster("ip")
            '    'END STEP 10 ------------------------------------------------------------------------------------
            'End If


            'modProgDialog.ProgDialog("Creating C2 GRID...", strTitle, 0, 27, 16, 0)
            'If modProgDialog.g_boolCancel Then
            '    'STEP 11: ---------------------------------------------------------------------------------------
            '    'CTwo GRID
            '    pMapAlgebraOp.BindRaster(pAbPrecipRaster, "ip")

            '    strExpression = clsPrecip.CTwo(g_intPrecipType)

            '    pCTwoRaster = pMapAlgebraOp.Execute(strExpression)

            '    pMapAlgebraOp.UnbindRaster("ip")
            '    'END STEP 11: ------------------------------------------------------------------------------------
            'End If


            'modProgDialog.ProgDialog("More math...", strTitle, 0, 27, 17, 0)
            'If modProgDialog.g_boolCancel Then
            '    'STEP 12a: ----------------------------------------------------------------------------------------
            '    'Logqu
            '    With pMapAlgebraOp
            '        .BindRaster(pCZeroRaster, "czero")
            '        .BindRaster(pConeRaster, "cone")
            '        .BindRaster(pLogTOCRaster, "logtoc")
            '        .BindRaster(pCTwoRaster, "ctwo")
            '        .BindRaster(pTempLogTOCRaster, "temp8")
            '    End With

            '    strExpression = "[czero] + ([cone] * [logtoc]) + ([ctwo] * [temp8])"

            '    pLogQuRaster = pMapAlgebraOp.Execute(strExpression)

            '    With pMapAlgebraOp
            '        .UnbindRaster("czero")
            '        .UnbindRaster("cone")
            '        .UnbindRaster("logtoc")
            '        .UnbindRaster("ctwo")
            '        .UnbindRaster("temp8")
            '    End With
            '    'END STEP 12a: -------------------------------------------------------------------------------------


            '    'STEP 12b: -----------------------------------------------------------------------------------------
            '    pMapAlgebraOp.BindRaster(pLogQuRaster, "logqu")

            '    strExpression = "Pow(10, [logqu])"

            '    pQuRaster = pMapAlgebraOp.Execute(strExpression)

            '    pMapAlgebraOp.UnbindRaster("logqu")
            '    'END STEP 12b: -------------------------------------------------------------------------------------
            'End If


            'modProgDialog.ProgDialog("Creating Pond Factor GRID...", strTitle, 0, 27, 16, 0)
            'If modProgDialog.g_boolCancel Then
            '    'STEP 13: ------------------------------------------------------------------------------------------
            '    'Create pond factor grid
            '    pMapAlgebraOp.BindRaster(g_LandCoverRaster, "nu_lulc")

            '    strExpression = strConPondStatement

            '    pPondFactorRaster = pMapAlgebraOp.Execute(strExpression)

            '    pMapAlgebraOp.UnbindRaster("nu_lulc")
            '    'END STEP 13: ---------------------------------------------------------------------------------------
            'End If


            'modProgDialog.ProgDialog("Calculating peak discharge; cubic feet per second...", strTitle, 0, 27, 17, 0)
            'If modProgDialog.g_boolCancel Then
            '    'STEP 14: -------------------------------------------------------------------------------------------
            '    'Calculate peak discharge: cubic feet per second
            '    With pMapAlgebraOp
            '        .BindRaster(pQuRaster, "qu")
            '        .BindRaster(g_pCellAreaSqMiRaster, "cellarea_sqmi")
            '        .BindRaster(g_pRunoffInchRaster, "runoff_in")
            '        .BindRaster(pPondFactorRaster, "pondfact")
            '    End With

            '    strExpression = "[qu] * [cellarea_sqmi] * [runoff_in] * [pondfact]"

            '    pQPRaster = pMapAlgebraOp.Execute(strExpression)

            '    With pMapAlgebraOp
            '        .UnbindRaster("qu")
            '        .UnbindRaster("cellarea_sqmi")
            '        .UnbindRaster("runoff_in")
            '        .UnbindRaster("pondfact")
            '    End With
            '    'END STEP 14: ----------------------------------------------------------------------------------------
            'End If


            'modProgDialog.ProgDialog("Creating cover factor GRID...", strTitle, 0, 27, 18, 0)
            'If modProgDialog.g_boolCancel Then
            '    'STEP 15: --------------------------------------------------------------------------------------------
            '    'Cover Factor GRID
            '    pMapAlgebraOp.BindRaster(g_LandCoverRaster, "nu_lulc")

            '    strExpression = strConStatement

            '    pCFactorRaster = pMapAlgebraOp.Execute(strExpression)

            '    pMapAlgebraOp.UnbindRaster("nu_lulc")
            '    'END STEP 15 -----------------------------------------------------------------------------------------
            'End If

            'modProgDialog.ProgDialog("Sediment Yield...", strTitle, 0, 27, 19, 0)
            'If modProgDialog.g_boolCancel Then
            '    'STEP 16 c: -------------------------------------------------------------------------------------------
            '    'Temp Sediment Yield: Note: MUSLE Exponent inserted to allow for global use, cuz other's will use
            '    With pMapAlgebraOp
            '        .BindRaster(g_pRunoffAFRaster, "runoff_af")
            '        .BindRaster(pQPRaster, "qp")
            '    End With

            '    strExpression = "Pow(([runoff_af] * [qp]), " & m_dblMUSLEExp & ")"

            '    pHISYTempRaster = pMapAlgebraOp.Execute(strExpression)

            '    With pMapAlgebraOp
            '        .UnbindRaster("runoff_af")
            '        .UnbindRaster("qp")
            '    End With
            '    'END STEP 16c: ------------------------------------------------------------------------------------------------------
            'End If


            'modProgDialog.ProgDialog("Sediment Yield...", strTitle, 0, 30, 27, 0)
            'If modProgDialog.g_boolCancel Then
            '    'STEP 16d: ---------------------------------------------------------------------------------------------
            '    'Sediment Yield: Note m_dblMUSLEVal now in for universal
            '    With pMapAlgebraOp
            '        .BindRaster(pCFactorRaster, "cfactor")
            '        .BindRaster(g_KFactorRaster, "kfactor")
            '        .BindRaster(g_pLSRaster, "lsfactor")
            '        .BindRaster(pHISYTempRaster, "temp9")
            '    End With

            '    strExpression = m_dblMUSLEVal & " * ([cfactor] * [kfactor] * [lsfactor] * [temp9])"

            '    pHISYRaster = pMapAlgebraOp.Execute(strExpression)

            '    With pMapAlgebraOp
            '        .UnbindRaster("cfactor")
            '        .UnbindRaster("kfactor")
            '        .UnbindRaster("lsfactor")
            '        .UnbindRaster("temp9")
            '    End With
            '    'END STEP 16b: ----------------------------------------------------------------------------------------
            'End If


            'modProgDialog.ProgDialog("Converting tons to milligrams...", strTitle, 0, 27, 21, 0)
            'If modProgDialog.g_boolCancel Then

            '    pMapAlgebraOp.BindRaster(pHISYRaster, "sy")

            '    strExpression = "[sy] * 907.184740"

            '    pHISYMGRaster = pMapAlgebraOp.Execute(strExpression)

            '    pMapAlgebraOp.UnbindRaster("sy")
            '    'END STEP 17b: ----------------------------------------------------------------------------------------
            'End If

            
            'Dim pClipMusleRaster As ESRI.ArcGIS.Geodatabase.IRaster
            'If g_booLocalEffects Then

            '    modProgDialog.ProgDialog("Creating data layer for local effects...", strTitle, 0, 27, 27, 0)
            '    If modProgDialog.g_boolCancel Then

            '        strMUSLE = modUtil.GetUniqueName("locmusle", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))
            '        'Added 7/23/04 to account for clip by selected polys functionality
            '        If g_booSelectedPolys Then
            '            pClipMusleRaster = modUtil.ClipBySelectedPoly(pHISYMGRaster, g_pSelectedPolyClip, pEnv)
            '            pPermMUSLERaster = modUtil.ReturnPermanentRaster(pClipMusleRaster, pEnv.OutWorkspace.PathName, strMUSLE)
            '        Else
            '            pPermMUSLERaster = modUtil.ReturnPermanentRaster(pHISYMGRaster, pEnv.OutWorkspace.PathName, strMUSLE)
            '        End If

            '        pMUSLERasterLocLayer = modUtil.ReturnRasterLayer((frmPrj.m_App), pPermMUSLERaster, "MUSLE Local Effects (mg)")
            '        pMUSLERasterLocLayer.Renderer = modUtil.ReturnRasterStretchColorRampRender(pMUSLERasterLocLayer, "Brown")
            '        pMUSLERasterLocLayer.Visible = False

            '        'metadata time
            '        g_dicMetadata.Add(pMUSLERasterLocLayer.Name, m_strMusleMetadata)

            '        g_pGroupLayer.Add(pMUSLERasterLocLayer)

            '        CalcMUSLE = True
            '        modProgDialog.KillDialog()
            '        Exit Function

            '    End If

            'End If


            'modProgDialog.ProgDialog("Calculating the accumulated sediment...", strTitle, 0, 27, 23, 0)
            'If modProgDialog.g_boolCancel Then

            '    pTempflowDir3raster = modUtil.ReturnRaster(g_strFlowDirFilename)
            '    pFlowAccumOp1 = New ESRI.ArcGIS.SpatialAnalyst.RasterHydrologyOp

            '    pFlowDirRDS1 = pTempflowDir3raster
            '    pHISYMGRDS = pHISYMGRaster

            '    pEnv = pFlowAccumOp1
            '    pOutRDS1 = pFlowAccumOp1.FlowAccumulation(pFlowDirRDS1, pHISYMGRDS)

            '    pAccSedHIRaster = pOutRDS1

            'End If

            'modProgDialog.ProgDialog("Calculating Total Sediment Mass...", strTitle, 0, 27, 24, 0)
            'If modProgDialog.g_boolCancel Then
            '    'STEP 20HI: -----------------------------------------------------------------------------------------------
            '    'Total Sediment Mass
            '    With pMapAlgebraOp
            '        .BindRaster(pHISYMGRaster, "sy_mg_HI")
            '        .BindRaster(pAccSedHIRaster, "accsed_HI")
            '    End With

            '    'old / 10000000
            '    strExpression = "[sy_mg_HI] + [accsed_HI]"

            '    pTotSedMassHIRaster = pMapAlgebraOp.Execute(strExpression)

            '    With pMapAlgebraOp
            '        .UnbindRaster("sy_mg_HI")
            '        .UnbindRaster("accsed_HI")
            '    End With
            '    'END STEP 20HI: -------------------------------------------------------------------------------------------
            'End If


            'modProgDialog.ProgDialog("Adding Sediment Mass to Group Layer...", strTitle, 0, 27, 25, 0)
            'Dim pClipMusleMassRaster As ESRI.ArcGIS.Geodatabase.IRaster
            'If modProgDialog.g_boolCancel Then
            '    'STEP 21: Created the Sediment Mass Raster layer and add to Group Layer -----------------------------------
            '    'Get a unique name for MUSLE and return the permanently made raster
            '    strMUSLE = modUtil.GetUniqueName("MUSLEmass", modUtil.SplitWorkspaceName(pEnv.OutWorkspace.PathName))

            '    'Clip to selected polys if chosen
            '    If g_booSelectedPolys Then
            '        pClipMusleMassRaster = modUtil.ClipBySelectedPoly(pTotSedMassHIRaster, g_pSelectedPolyClip, pEnv)
            '        pPermTotSedConcHIraster = modUtil.ReturnPermanentRaster(pClipMusleMassRaster, pEnv.OutWorkspace.PathName, strMUSLE)
            '    Else
            '        pPermTotSedConcHIraster = modUtil.ReturnPermanentRaster(pTotSedMassHIRaster, pEnv.OutWorkspace.PathName, strMUSLE)
            '    End If

            '    'Now create the MUSLE layer
            '    pMUSLERasterLayer = modUtil.ReturnRasterLayer((frmPrj.m_App), pPermTotSedConcHIraster, "MUSLE Sediment Mass (kg)")
            '    pMUSLERasterLayer.Renderer = modUtil.ReturnRasterStretchColorRampRender(pMUSLERasterLayer, "Brown")
            '    pMUSLERasterLayer.Visible = False

            '    'Metadata:
            '    g_dicMetadata.Add(pMUSLERasterLayer.Name, m_strMusleMetadata)

            '    'Add the MUSLE Layer to the final group layer
            '    g_pGroupLayer.Add(pMUSLERasterLayer)

            '    'end STEP 21: Created the Sediment Mass Raster layer and add to Group Layer -----------------------------------

            'End If

            
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

    End Function

    Private Function tmpCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        For i As Integer = 0 To _picks.Length - 1
            If Input1 = i + 1 Then
                Return _picks(i)
            End If
        Next
    End Function

#End Region

End Module