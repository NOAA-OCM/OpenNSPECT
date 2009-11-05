Attribute VB_Name = "modPollConc"
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





'CREATE PHOSPHORUS EMC GRID AT CELL LEVEL        (these values will change)
'phosmass =
'pick([ccap],-9999,.33,.52,.37,.20,.11,.11,.11,.14,0.08,0.08,0.08,0.08,0.08,0.08,0.08,.26,0.08,0.08,0.08)
'Units: milligrams per liter
'
'
'MASS OF PHOSPHORUS PRODUCED BY EACH CELL
'massvolume = [met_runoff] * [phosmass]
'Units: milligrams per day
'
'
'DERIVE ACCUMULATED PHOSPHORUS
'accphosphorus = FlowAccumulation([flowdir], [massvolume])
'Units: milligrams per day
'
'
'DERIVE ACCUMULATED CONCENTRATION
'accphos_conc = [accphosphorus] / [accrunoff]
'Units: milligrams per liter
'
'
'DERIVE TOTAL CONCENTRATION (AT EACH CELL)
'[temp1] = [massvolume] + [accphosphorus]
'[temp2] = [met_runoff] + [accrunoff]
'totphos_conc = [temp1] / [temp2]
'Units: milligrams per liter

Private Function ConstructPickStatment(strLandClass As String, pLCRaster As IRaster) As String
    
    Dim strRS As String
    Dim rsLandClass As ADODB.Recordset
    Dim strPick(3) As String     'Array of strings that hold 'pick' numbers
    Dim strCurveCalc As String   'Full String
    
    Dim pRasterCol As IRasterBandCollection
    Dim pBand As IRasterBand
    Dim pTable As ITable
    Dim TableExist As Boolean
    Dim pCursor As ICursor
    Dim pRow As iRow
    Dim FieldIndex As Integer
    
    'STEP 1:  get the records from the database -----------------------------------------------
    strRS = "SELECT * FROM LCCLASS " & _
     "LEFT OUTER JOIN LCTYPE ON LCCLASS.LCTYPEID = LCTYPE.LCTYPEID " & _
     "Where LCTYPE.NAME = '" & strLandClass & "' ORDER BY LCCLASS.VALUE"
    
    Set rsLandClass = New ADODB.Recordset
    rsLandClass.Open strRS, g_ADOConn, adOpenKeyset, adLockOptimistic
    
    rsLandClass.MoveFirst
    'End Database stuff
    
    'STEP 2: Raster Values ---------------------------------------------------------------------
    'Now Get the RASTER values
    ' Get Rasterband from the incoming raster
    Set pRasterCol = pLCRaster
    Set pBand = pRasterCol.Item(0)

    'Get the raster table
    pBand.HasTable TableExist
    If Not TableExist Then Exit Function
    
    Set pTable = pBand.AttributeTable
    'Get All rows
    Set pCursor = pTable.Search(Nothing, True)
    'Init pRow
    Set pRow = pCursor.NextRow
    
    'Get index of Value Field
    FieldIndex = pTable.FindField("Value")
    
    'STEP 3: Table values vs. Raster Values Count - if not equal bark --------------------------
    If pTable.RowCount(Nothing) <> rsLandClass.RecordCount Then
        MsgBox "Error: The number of records in your Land Class Table do not match your GRID dataset."
        Exit Function
    End If
    
    'STEP 4: Create the strings
    'Loop through and get all values
    Do While Not pRow Is Nothing
        If pRow.Value(FieldIndex) = rsLandClass!Value Then
            If strPick(0) = "" Then
                strPick(0) = "-9999, " & CStr(rsLandClass![CN-A])
            Else
                strPick(0) = strPick(0) & ", " & CStr(rsLandClass![CN-A])
            End If
            
            If strPick(1) = "" Then
                strPick(1) = "-9999, " & CStr(rsLandClass![CN-B])
            Else
                strPick(1) = strPick(1) & ", " & CStr(rsLandClass![CN-B])
            End If
            
            If strPick(2) = "" Then
                strPick(2) = "-9999, " & CStr(rsLandClass![CN-C])
            Else
                strPick(2) = strPick(2) & ", " & CStr(rsLandClass![CN-C])
            End If
            
            If strPick(3) = "" Then
                strPick(3) = "-9999, " & CStr(rsLandClass![CN-D])
            Else
                strPick(3) = strPick(3) & ", " & CStr(rsLandClass![CN-D])
            End If
            
            rsLandClass.MoveNext
            Set pRow = pCursor.NextRow
        Else
            MsgBox "Values in table LCClass table not equal to values in landclass dataset."
            ConstructPickStatment = ""
            Exit Function
        End If
    Loop
    
    strCurveCalc = "con([rev_soils] == 1, pick([ccap], " & strPick(0) & "), con([rev_soils] == 2, pick([ccap], " & _
    strPick(1) & "), con([rev_soils] == 3, pick([ccap], " & strPick(2) & "), con([rev_soils] == 4, pick([ccap], " & _
    strPick(3) & ")))))"

    ConstructPickStatment = strCurveCalc
    
    'Cleanup:
    Set pRasterCol = Nothing
    Set pBand = Nothing
    Set pTable = Nothing
    Set pCursor = Nothing
    Set pRow = Nothing
    
End Function


