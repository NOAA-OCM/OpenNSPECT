Attribute VB_Name = "modUtil"
' *************************************************************************************
' *  modUtil
' *  Perot Systems Government Services
' *  Contact: Ed Dempsey  -  ed.dempsey@noaa.gov
' *************************************************************************************
' *  Description:  Contains a number of widely accessed variables, functions, and subs
' *  for use throughout N-SPECT
' *
' *  Called By:  Various
' *************************************************************************************

Option Explicit

' Path to N-SPECT's application and document folder
Public g_nspectPath         As String
Public g_nspectDocPath      As String

'Database Variables
Public g_ADOConn            As ADODB.Connection         'Connection
Public g_strConn            As String                   'Connection String
Public g_ADORS              As ADODB.Recordset          'ADO Recordset
Public g_adoRSCoeff         As New ADODB.Recordset
Public g_boolConnected      As Boolean                  'Bool: connected
Public g_intLCClassid       As Integer                  'LCClassid - used for adding 'NEW' records

'Bool: frmAddCoeff is used twice; called from frmNewPollutants(False), frmPollutants(True)
Public g_boolAddCoeff       As Boolean
Public g_boolCopyCoeff      As Boolean                  'True: called frmPollutants, False: called frmNewPollutants
Public g_boolAgree          As Boolean                  'True: use the Agree Function on Streams.
Public g_boolHydCorr        As Boolean                  'True: Hyrdologically Correct DEM, no fill needed
Public g_boolNewWShed       As Boolean                  'True: New WaterShed form called from frmPrj

'WqStd
Public rsWQStdLoad As ADODB.Recordset                   'Water Quality RecordSet

'Agree DEM Stuff
Public g_boolParams         As Boolean                  'Flag to indicate Agree params have been entered

'Help API
Declare Function HtmlHelp Lib "HHCtrl.ocx" Alias "HtmlHelpA" _
   (ByVal hwndCaller As Long, _
   ByVal pszFile As String, _
   ByVal uCommand As Long, _
   dwData As Any) As Long
   
Public Const HH_DISPLAY_TOPIC = &H0
Public Const HH_HELP_CONTEXT = &HF

'Project Form Variables
Public g_strPrjFileName As String                       'Project file name

'Management Scenario variables::frmPrjCalc
Public g_strLUScenFileName As String                    'Management scenario file name
Public g_intManScenRow As String                        'Management scenario ROW number

'Pollutant Coefficient variable::frmPrjCalc
Public g_intCoeffRow As Integer                         'Coeff Row Number
Public g_strCoeffCalc As String                         'if the Calc option is chosen, hold results in string
Const c_sModuleFileName As String = "modUtil.bas"
Private m_ParentHWND As Long          ' Set this to get correct parenting of Error handler forms

Public Sub AddFeatureLayer(pApp As esriFramework.IApplication, pFClass As esriGeoDatabase.IFeatureClass, Optional sName As String)
  On Error GoTo ErrorHandler

    Dim pFLayer As IFeatureLayer
    Set pFLayer = New FeatureLayer
    Set pFLayer.FeatureClass = pFClass
    If (sName = "") Then
      Dim pFDataset As IFeatureDataset
      Set pFDataset = pFClass.FeatureDataset
      Dim pDS As IDataset
      Set pDS = pFDataset
      pFLayer.Name = pDS.BrowseName
    Else
      pFLayer.Name = sName
    End If
    Dim pMap As IBasicMap
    Dim pMxDoc As IMxDocument
    Dim pActView As IActiveView
    Set pMxDoc = pApp.Document
    Set pMap = pMxDoc.FocusMap
    Set pActView = pMxDoc.ActiveView
    pMap.AddLayer pFLayer
    pActView.Refresh
    pMxDoc.UpdateContents
    Exit Sub
        
    Set pMxDoc = Nothing
    Set pActView = Nothing
    Set pMap = Nothing
    Set pFLayer = Nothing
    
  Exit Sub
ErrorHandler:
  HandleError True, "AddFeatureLayer " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Public Sub AddFeatureLayerToComboBox(cboBox As ComboBox, pMap As esriCarto.IMap, strType As String)
'Populate combobox of choice with RasterLayers in map
On Error GoTo ErrHandler
    cboBox.Clear
    Dim iLyrIndex As Long
    Dim pFeatClass As IFeatureClass
    Dim pLyr As ILayer
    Dim pFeatLayer As IFeatureLayer
    ' Add raster layers into  Combobox
    Dim iLayerCount As Integer
    iLayerCount = pMap.LayerCount
    If iLayerCount > 0 Then
        cboBox.Enabled = True
        For iLyrIndex = 0 To iLayerCount - 1
            Set pLyr = pMap.Layer(iLyrIndex)
            If (TypeOf pLyr Is IFeatureLayer) Then
                Set pFeatLayer = pLyr
                Set pFeatClass = pFeatLayer.FeatureClass
                
                Select Case strType
                
                Case "line"
                
                    If pFeatClass.ShapeType = esriGeometryPolyline Then
                        cboBox.AddItem pLyr.Name
                        cboBox.ItemData(cboBox.ListCount - 1) = iLyrIndex
                    End If
                    
                Case "poly"
                
                    If pFeatClass.ShapeType = esriGeometryPolygon Then
                        cboBox.AddItem pLyr.Name
                        cboBox.ItemData(cboBox.ListCount - 1) = iLyrIndex
                    End If
                    
                Case "point"
                    
                    If pFeatClass.ShapeType = esriGeometryPoint Then
                        cboBox.AddItem pLyr.Name
                        cboBox.ItemData(cboBox.ListCount - 1) = iLyrIndex
                    End If
                    
                End Select
            End If
        Next iLyrIndex
        If (cboBox.ListCount > 0) Then
            cboBox.ListIndex = 0
            cboBox.Text = pMap.Layer(cboBox.ItemData(0)).Name
        End If
    End If
    
    Set pFeatClass = Nothing
    Set pLyr = Nothing
    Set pFeatLayer = Nothing
    
Exit Sub
ErrHandler:
    MsgBox "Add Layer to ComboBox:" & Err.Description

End Sub
'General Function used to simply get the index of combobox entries
Public Function GetCboIndex(strList As String, cbo As ComboBox) As Integer
  On Error GoTo ErrorHandler

    
    Dim i As Integer
    i = 0
    
    For i = 0 To cbo.ListCount - 1
        If cbo.List(i) = strList Then
            GetCboIndex = i
        End If
    Next i
        


  Exit Function
ErrorHandler:
  HandleError True, "GetCboIndex " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Function

'Function for connection to NSPECT.mdb: fires on dll load
Public Sub ADOConnection()
    
    On Error GoTo ErrHandler:
    If Not g_boolConnected Then
        
        Set g_ADOConn = New ADODB.Connection
        
        g_strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=" & modUtil.g_nspectPath & "\nspect.mdb"
        
        g_ADOConn.Open g_strConn
        g_ADOConn.CursorLocation = adUseServer
        
        g_boolConnected = True
     
    End If
        Exit Sub
ErrHandler:
        MsgBox Err.Number & Err.Description & " Error connecting to database, please check NSPECTDAT enviornment variable.  Current value of NSPECTDAT: " & _
        g_strConn, vbCritical, "Error Connecting"
    
End Sub

'Need to check the text file coming in from the import menu of the pollutant form.
'Bringing the Text File itself, and the name of the LCType as picked by John User
Public Function ValidateCoeffTextFile(strFileName As String, strLCTypeName As String) As Boolean
  On Error GoTo ErrorHandler


    Dim fso As New Scripting.FileSystemObject
    Dim fl As TextStream
    Dim strLine As String
    Dim intLine As Integer
    Dim strValue As String
    Dim strParams(7) As Variant
    Dim strLCType As String
    Dim rsLCType As New ADODB.Recordset
    Dim strLCTypeNum As String
    Dim rsCheck As New ADODB.Recordset
    Dim strCheck As String
    Dim i As Integer, j As Integer
    Dim lstValues() As Integer
    
    'Gameplan is to find number of records(landclasses) in the chosen LCType. Then
    'compare that to the number of lines in the text file, and the [Value] field to
    'make sure both jive.  If not, bark at them...ruff, ruff
    
    strLCTypeNum = "SELECT LCTYPE.LCTYPEID, LCCLASS.NAME, LCCLASS.VALUE, LCCLASS.LCCLASSID FROM " & _
        "LCTYPE INNER JOIN LCCLASS ON LCTYPE.LCTYPEID = LCCLASS.LCTYPEID " & _
        "WHERE LCTYPE.NAME LIKE '" & strLCTypeName & "'"
    
    rsLCType.Open strLCTypeNum, g_ADOConn, adOpenKeyset, adLockOptimistic

    If fso.FileExists(strFileName) Then
    
        Set fl = fso.OpenTextFile(strFileName, ForReading, True, TristateFalse)
        
        intLine = 0
        
        'CHECK 1: loop through the text file, compare value to [Value], make sure exists
        Do While Not fl.AtEndOfStream
        
            strLine = fl.ReadLine
            'Value exits??
            strValue = Split(strLine, ",")(0)
    
            j = 0
            rsLCType.MoveFirst
            
            For i = 0 To rsLCType.RecordCount - 1
                
                If rsLCType!Value = strValue Then
                     j = j + 1
                End If
                rsLCType.MoveNext
                
            Next i
            
            If j = 0 Then
                MsgBox "There is a value in your text file that does not exist in the Land Class Type: '" & _
                    strLCTypeName & "' Please check your text file in line: " & intLine + 1, vbOKOnly, "Data Import Error"
                
                 ValidateCoeffTextFile = False
                GoTo Cleanup
            ElseIf j > 1 Then
                MsgBox "There are records in your text file that contain the same value.  Please check line " & intLine, vbCritical, _
                    "Multiple values found"
                ValidateCoeffTextFile = False
                GoTo Cleanup
            ElseIf j = 1 Then
                ValidateCoeffTextFile = True
            End If
            
            intLine = intLine + 1
            Debug.Print intLine
            
        Loop
                
        'Final check, make sure same number of records in text file vs the
        If rsLCType.RecordCount = intLine Then
                ValidateCoeffTextFile = True
        Else
            MsgBox "The number of records in your import file do not match the number of records in the " & _
                "Landclass '" & strLCTypeName & "'.  Your file should contain " & rsLCType.RecordCount & " records.", _
                vbCritical, "Error Importing File"
                ValidateCoeffTextFile = False
                GoTo Cleanup
        End If
                

    Else
        MsgBox "The file you are pointing to does not exist. Please select another.", vbCritical, "File Not Found"
    'Cleanup
    End If
    
    If ValidateCoeffTextFile Then
        Set g_adoRSCoeff = rsLCType
    End If
    
Exit Function
    
Cleanup:
                
    Set fso = Nothing
    Set fl = Nothing
    rsLCType.Close
    Set rsLCType = Nothing
    


  Exit Function
ErrorHandler:
  HandleError True, "ValidateCoeffTextFile " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Function

'Check for spatial reference of a Raster Dataset, returns the Coordinate System
Public Function CheckSpatialReference(pRasGeoDataset As esriGeoDatabase.IRasterDataset) As esriGeometry.IProjectedCoordinateSystem
  On Error GoTo ErrorHandler

    
    'Code to call ArcMap Functionality....
    Dim pGeoDataset As IGeoDataset
    Dim pDEMSpatRef As ISpatialReference
    Dim pRasterProps As IRasterProps
    Dim pDistUnit As ILinearUnit
    Dim pPrjCoordSys As IProjectedCoordinateSystem
    
    Set pGeoDataset = pRasGeoDataset.CreateDefaultRaster  'Create a default to get props
    Set pRasterProps = pGeoDataset
    
    'Get the Spat Reference, check for: if there, send it back, else Nothing
    If Not (pRasterProps.SpatialReference Is Nothing) Then
        Set CheckSpatialReference = pRasterProps.SpatialReference
    Else
        Set CheckSpatialReference = Nothing
    End If
    
    'Cleanup
    Set pGeoDataset = Nothing
    Set pDEMSpatRef = Nothing
    Set pRasterProps = Nothing
    Set pDistUnit = Nothing
    Set pPrjCoordSys = Nothing



  Exit Function
ErrorHandler:
  HandleError True, "CheckSpatialReference " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Function

Public Function CheckLandCoverFields(strLandClass As String, pLCRaster As esriGeoDatabase.IRaster) As Boolean

    Dim strRS As String
    Dim rsLandClass As ADODB.Recordset
    Dim strPick(3) As String     'Array of strings that hold 'pick' numbers
    Dim strCurveCalc As String   'Full String
    
    Dim pRasterCol As IRasterBandCollection
    Dim pBand As IRasterBand
    Dim pRasterStats As IRasterStatistics
    Dim dblMaxValue As Double
    Dim i As Integer
    Dim pTable As ITable
    Dim TableExist As Boolean
    
On Error GoTo ErrHandler:

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
    
    'Get the max value
    Set pRasterStats = pBand.Statistics
    dblMaxValue = pRasterStats.Maximum

    'Get the raster table
    pBand.HasTable TableExist
    If Not TableExist Then
        CheckLandCoverFields = False
        Exit Function
    End If
    
    Set pTable = pBand.AttributeTable
    
    'STEP 3: Table values vs. Raster Values Count - if not equal bark --------------------------
    If pTable.RowCount(Nothing) <> rsLandClass.RecordCount Then
        CheckLandCoverFields = False
    Else
        CheckLandCoverFields = True
    End If
    
    'Cleanup
    rsLandClass.Close
    Set rsLandClass = Nothing
Exit Function

ErrHandler:
    MsgBox "Error Occurred on CheckFields"
    CheckLandCoverFields = False
    
End Function



'add grid column headers for Coefficient tab
Public Sub InitPollDefGrid(grdPollDef As MSHFlexGrid)
  On Error GoTo ErrorHandler
    
    
   
   With grdPollDef
      .col = .FixedCols
      .row = .FixedRows
      .ColWidth(0) = 400
      .ColWidth(1) = 850
      .ColWidth(2) = 2600
      .ColWidth(3) = 840
      .ColWidth(4) = 840
      .ColWidth(5) = 840
      .ColWidth(6) = 840
      .ColWidth(7) = 0      'Coefficient Set ID
      .ColWidth(8) = 0      'LCClass ID - both hidden but necessary for other deeds.
      .row = 0
      .col = 1
      .Text = "Value"
      .CellAlignment = flexAlignCenterCenter
      .col = 2
      .Text = "Name"
      .CellAlignment = flexAlignCenterCenter
      .col = 3
      .Text = "Type 1"
      .CellAlignment = flexAlignCenterCenter
      .col = 4
      .Text = "Type 2"
      .CellAlignment = flexAlignCenterCenter
      .col = 5
      .Text = "Type 3"
      .CellAlignment = flexAlignCenterCenter
      .col = 6
      .Text = "Type 4"
      .CellAlignment = flexAlignCenterCenter
    End With
   
    


  Exit Sub
ErrorHandler:
  HandleError True, "InitPollDefGrid " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

'add grid column headers for WQ Standards tab
Public Sub InitPollWQStdGrid(grdWQStd As MSHFlexGrid)
  On Error GoTo ErrorHandler

  
   With grdWQStd
      .col = .FixedCols
      .row = .FixedRows
      .BackColorSel = &HC0FFFF   'lt. yellow
      .Width = 7100
      .ColWidth(0) = 400
      .ColWidth(1) = 2000
      .ColWidth(2) = 3150
      .ColWidth(3) = 1450
      .row = 0
      .col = 1
      .Text = "Standards Name"
      .CellAlignment = flexAlignCenterCenter
      .col = 2
      .Text = "Description"
      .CellAlignment = flexAlignCenterCenter
      .col = 3
      .Text = "Threshold (ug/L)"
      .CellAlignment = flexAlignRightCenter
      .col = 4
      .ColWidth(4) = 0
   End With



  Exit Sub
ErrorHandler:
  HandleError True, "InitPollWQStdGrid " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

'add grid column headers
Public Sub InitWQStdGrid(grdWQStd As MSHFlexGrid)
  On Error GoTo ErrorHandler

  
   With grdWQStd
      .col = .FixedCols
      .row = .FixedRows
      .BackColorSel = &HC0FFFF   'lt. yellow
      .Width = 5460
      .ColWidth(0) = 300
      .ColWidth(1) = 2900
      .ColWidth(2) = 2140
      .ColWidth(3) = 0
      .row = 0
      .col = 1
      .Text = "Pollutant"
      .CellAlignment = flexAlignCenterCenter
      .col = 2
      .Text = "Threshold (ug/L)"
      .CellAlignment = flexAlignCenterCenter
      
   End With



  Exit Sub
ErrorHandler:
  HandleError True, "InitWQStdGrid " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Public Sub InitLCClassesGrid(grdLCClasses As MSHFlexGrid)
  On Error GoTo ErrorHandler

   'add grid column headers
   With grdLCClasses
      .col = .FixedCols
      .row = .FixedRows
      .BackColorSel = &HC0FFFF   'lt. yellow
      .ColWidth(0) = 400
      .ColWidth(1) = 800
      .ColWidth(2) = 2200
      .ColWidth(3) = 800
      .ColWidth(4) = 800
      .ColWidth(5) = 800
      .ColWidth(6) = 800
      .ColWidth(7) = 1350
      .ColWidth(8) = 800
      .Width = 9050
      .row = 0
      .col = 1
      .Text = "Value"
      .CellAlignment = flexAlignRightCenter
      .col = 2
      .Text = "Name"
      .CellAlignment = flexAlignLeftCenter
      .col = 3
      .Text = "CN-A"
      .CellAlignment = flexAlignRightCenter
      .col = 4
      .Text = "CN-B"
      .CellAlignment = flexAlignRightCenter
      .col = 5
      .Text = "CN-C"
      .CellAlignment = flexAlignRightCenter
      .col = 6
      .Text = "CN-D"
      .CellAlignment = flexAlignRightCenter
      .col = 7
      .Text = "Cover-Factor"
      .CellAlignment = flexAlignRightCenter
      .col = 8
      .Text = "Wet"
      .CellAlignment = flexAlignCenterCenter
      .col = 9
      .ColWidth(9) = 0
      .col = 10
      .ColWidth(10) = 0
   End With



  Exit Sub
ErrorHandler:
  HandleError True, "InitLCClassesGrid " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Public Function IsLoaded(FormName As String) As Boolean
  On Error GoTo ErrorHandler

'Used to see if particular forms are/are not loaded at various time
    Dim sFormName As String
    Dim f As Form
    sFormName = UCase$(FormName)
    
    For Each f In Forms
       If UCase$(f.Name) = sFormName Then
         IsLoaded = True
         Exit Function
       End If
    Next



  Exit Function
ErrorHandler:
  HandleError True, "IsLoaded " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Function

Public Sub InitComboBox(cbo As ComboBox, strName As String)
  On Error GoTo ErrorHandler

'Loads the variety of comboboxes throught the project using combobox and name of table
    Dim rsNames As ADODB.Recordset
    Dim strSelectStatement As String
    Set rsNames = New ADODB.Recordset
   
    strSelectStatement = "SELECT NAME FROM " & strName & " ORDER BY NAME ASC"
    
    'Check thrown in to make sure g_ADOconn is something, in v9.1 we started having problems.
    If Not g_boolConnected Then
        ADOConnection
    End If
    
    rsNames.CursorLocation = adUseClient
    rsNames.Open strSelectStatement, g_ADOConn, adOpenDynamic, adLockOptimistic
    
    If rsNames.RecordCount > 0 Then
        With cbo
            Do Until rsNames.EOF
              .AddItem rsNames!Name
            rsNames.MoveNext
            Loop
        End With
        
        cbo.ListIndex = 0
    Else
        MsgBox "Warning.  There are no records remaining.  Please add a new one.", vbCritical, "Recordset Empty"
        Exit Sub
    End If
    
    'Cleanup
    rsNames.Close
    Set rsNames = Nothing
    


  Exit Sub
ErrorHandler:
  HandleError True, "InitComboBox " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

'Tests name inputs to insure unique values
Public Function UniqueName(strTableName As String, strName As String) As Boolean
  On Error GoTo ErrorHandler

    
    Dim strCmdText As String
    Dim rsName As ADODB.Recordset
    
    strCmdText = "SELECT * FROM " & strTableName & " WHERE NAME LIKE '" & strName & "'"
    Set rsName = New ADODB.Recordset
    
    rsName.Open strCmdText, modUtil.g_ADOConn, adOpenStatic, adLockReadOnly
    
    If rsName.RecordCount > 0 Then
        UniqueName = False
    Else
        UniqueName = True
    End If
    
    'Cleanjeans
    rsName.Close
    Set rsName = Nothing
      


  Exit Function
ErrorHandler:
  HandleError True, "UniqueName " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Function

'Tests name inputs to insure unique values for databases
Public Function CreateUniqueName(strTableName As String, strName As String) As String
  On Error GoTo ErrorHandler

    
    Dim strCmdText As String
    Dim rsName As ADODB.Recordset
    Dim i As Integer
    Dim sCurrNum As String
    Dim strCurrNameRecord As String
    strCmdText = "SELECT * FROM " & strTableName '& " WHERE NAME LIKE '" & strName & "'"
    Set rsName = New ADODB.Recordset
    
    rsName.Open strCmdText, modUtil.g_ADOConn, adOpenStatic, adLockReadOnly
    sCurrNum = "0"
    
    rsName.MoveFirst
    
    For i = 1 To rsName.RecordCount
    strCurrNameRecord = CStr(rsName!Name)
    If InStr(1, strCurrNameRecord, strName, 1) > 0 Then
        If IsNumeric(Right(strCurrNameRecord, 2)) Then
          If (CInt(Right(strCurrNameRecord, 2)) > CInt(sCurrNum)) Then
            sCurrNum = Right(strCurrNameRecord, 2)
          Else
            Exit For
          End If
        Else
          If IsNumeric(Right(strCurrNameRecord, 1)) Then
            If (CInt(Right(strCurrNameRecord, 1)) > CInt(sCurrNum)) Then
              sCurrNum = Right(strCurrNameRecord, 1)
            End If
          End If
        End If
      End If
    rsName.MoveNext
    Next i
    If sCurrNum = "0" Then
      CreateUniqueName = strName + "1"
    Else
      CreateUniqueName = strName & CStr(CInt(sCurrNum) + 1)
    End If
        
    'Cleanjeans
    rsName.Close
    Set rsName = Nothing
      


  Exit Function
ErrorHandler:
  HandleError True, "CreateUniqueName " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Function


Private Sub SetGridColumnWidth(grd As MSFlexGrid)
  On Error GoTo ErrorHandler

    'params:    ms flexgrid control
    'purpose:   sets the column widths to the
    '           lengths of the longest string in the column
    'requirements:  the grid must have the same
    '               font as the underlying form

    Dim InnerLoopCount As Long
    Dim OuterLoopCount As Long
    Dim lngLongestLen As Long
    Dim sLongestString As String
    Dim lngColWidth As Long
    Dim szCellText As String

    For OuterLoopCount = 0 To grd.Cols - 1
        sLongestString = ""
        lngLongestLen = 0

        'grd.Col = OuterLoopCount
        For InnerLoopCount = 0 To grd.Rows - 1
            szCellText = grd.TextMatrix(InnerLoopCount, OuterLoopCount)
            'grd.Row = InnerLoopCount
            'szCellText = Trim$(grd.Text)
            If Len(szCellText) > lngLongestLen Then
                lngLongestLen = Len(szCellText)
                sLongestString = szCellText
            End If
        Next
        lngColWidth = grd.Parent.TextWidth(sLongestString)

        'add 100 for more readable spreadsheet
        grd.ColWidth(OuterLoopCount) = lngColWidth + 200
    Next


  Exit Sub
ErrorHandler:
  HandleError False, "SetGridColumnWidth " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Public Sub HighlightGridRow(grd As MSHFlexGrid, iRow As Integer)
  On Error GoTo ErrorHandler

    With grd
        If .Rows > 1 Then
            .row = iRow
            .col = 1
            .ColSel = .Cols - 1
            .RowSel = iRow
        End If
    End With


  Exit Sub
ErrorHandler:
  HandleError True, "HighlightGridRow " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub


'*********************************************************************************************
'** Following series of functions courtesy of ESRI, provide a variety of useful
'** actions
'*********************************************************************************************
Public Function OpenRasterDataset(sPath As String, sRasterName As String) As IRasterDataset
    'Return RasterDataset Object given a file name and its directory
    On Error GoTo ERH
    Dim pWSFact As IWorkspaceFactory
    Dim pRasterWS As IRasterWorkspace
    
    Set pWSFact = New RasterWorkspaceFactory
    If pWSFact.IsWorkspace(sPath) Then
        Set pRasterWS = pWSFact.OpenFromFile(sPath, 0)
        Set OpenRasterDataset = pRasterWS.OpenRasterDataset(sRasterName)
        
    End If
    Exit Function
ERH:
    MsgBox "Failed in opening raster dataset. " & Err.Description
End Function
Public Function IsValidWorkspace(sPath As String) As Boolean
' Given a pathname, determines if workspace is valid
On Error GoTo ErrorHandler:
    Dim pWSF As IWorkspaceFactory
    Set pWSF = New RasterWorkspaceFactory
    If pWSF.IsWorkspace(sPath) Then
        IsValidWorkspace = True
    Else
        IsValidWorkspace = False
    End If
        
    Set pWSF = Nothing
    Exit Function

ErrorHandler:
    IsValidWorkspace = False
    Set pWSF = Nothing
    
End Function

Public Function SetRasterWorkspace(sPath As String) As IRasterWorkspace

' Given a pathname, returns the raster workspace object for that path
    On Error GoTo ErrorSetWorkspace
    Dim pWSF As IWorkspaceFactory
    Set pWSF = New RasterWorkspaceFactory
    If pWSF.IsWorkspace(sPath) Then
        Set SetRasterWorkspace = pWSF.OpenFromFile(sPath, 0)
        Set pWSF = Nothing
    Else
        Dim pPropSet As IPropertySet
        Set pPropSet = New PropertySet
        pPropSet.setProperty "DATABASE", sPath
        Set SetRasterWorkspace = pWSF.Open(pPropSet, 0)
        
        Set pPropSet = Nothing
        
    End If
    Exit Function
ErrorSetWorkspace:
    Set SetRasterWorkspace = Nothing

End Function

Public Function SetFeatureShapeWorkspace(sPath As String) As IWorkspace

' Given a pathname, returns the shapefile workspace object for that path
On Error GoTo ErrorSetWorkspace
    Dim pWSF As IWorkspaceFactory
    Set pWSF = New ShapefileWorkspaceFactory
    If pWSF.IsWorkspace(sPath) Then
        Set SetFeatureShapeWorkspace = pWSF.OpenFromFile(sPath, 0)
        Set pWSF = Nothing
    Else
        Dim pPropSet As IPropertySet
        Set pPropSet = New PropertySet
        pPropSet.setProperty "DATABASE", sPath
        Set SetFeatureShapeWorkspace = pWSF.Open(pPropSet, 0)
        
        Set pPropSet = Nothing
    End If
    Exit Function

ErrorSetWorkspace:
    Set SetFeatureShapeWorkspace = Nothing

End Function

'Function to get unique filename in directory of choice
Public Function GetUniqueName(Name As String, folderPath As String) As String

On Error GoTo EH
    ' Find file, if exists, increment number
    Dim fso 'As FileSystemObject
    Dim pFolder 'As Folder
    Dim pFile 'As File
    Dim sCurrNum As String
    
    Set fso = CreateObject("Scripting.FileSystemObject") ' New FileSystemObject
    Set pFolder = fso.GetFolder(folderPath)
    sCurrNum = "0"
    For Each pFile In pFolder.SubFolders
      If InStr(1, pFile.Name, Name, 1) > 0 Then
        Dim s() As String
        s = Split(pFile.Name, ".")
        If IsNumeric(Right(s(0), 2)) Then
          If (CInt(Right(s(0), 2)) > CInt(sCurrNum)) Then
            sCurrNum = Right(s(0), 2)
          Else
            Exit For
          End If
        Else
          If IsNumeric(Right(s(0), 1)) Then
            If (CInt(Right(s(0), 1)) > CInt(sCurrNum)) Then
              sCurrNum = Right(s(0), 1)
            End If
          End If
        End If
      End If
    Next pFile
    If sCurrNum = "0" Then
      GetUniqueName = Name + "1"
    Else
      GetUniqueName = Name & CStr(CInt(sCurrNum) + 1)
    End If
    
    Set pFile = Nothing
    Set pFolder = Nothing
    Set fso = Nothing
    Exit Function
EH:
    MsgBox Err.Number & vbLf & Err.Description, , "Error in GetUniqueName "

End Function

Public Sub AddRasterLayer(pApp As IApplication, pRaster As IRaster, Optional sName As String)
  On Error GoTo EH

    Dim pRasterlayer As IRasterLayer
    Set pRasterlayer = New RasterLayer
    pRasterlayer.CreateFromRaster pRaster
    
    If (sName = "") Then
      Dim pRasterBands As IRasterBandCollection
      Set pRasterBands = pRaster
      Dim pRasterBand As IRasterBand
      Set pRasterBand = pRasterBands.Item(0)
      Dim pDS As IDataset
      Set pDS = pRasterBand.RasterDataset
      pRasterlayer.Name = pDS.BrowseName
    Else
      pRasterlayer.Name = sName
    End If
    
    Dim pMap As IBasicMap
    Dim pMxDoc As IMxDocument
    Dim pActView As IActiveView
    Set pMxDoc = pApp.Document
    Set pMap = pMxDoc.FocusMap
    Set pActView = pMxDoc.ActiveView
    pMap.AddLayer pRasterlayer
    pActView.Refresh
    pMxDoc.UpdateContents
    Exit Sub
    
    Set pMxDoc = Nothing
    Set pActView = Nothing
    Set pMap = Nothing
    Set pRasterlayer = Nothing
EH:
    MsgBox "Util.AddRasterLayer: " & Err.Description
End Sub

Public Function ReturnRasterLayer(pApp As IApplication, pRaster As IRaster, Optional sName As String) As IRasterLayer

    Dim pRasterlayer As IRasterLayer

On Error GoTo EH:
    
    Set pRasterlayer = New RasterLayer
    pRasterlayer.CreateFromRaster pRaster
   
    If (sName = "") Then
      Dim pRasterBands As IRasterBandCollection
      Set pRasterBands = pRaster
      Dim pRasterBand As IRasterBand
      Set pRasterBand = pRasterBands.Item(0)
      Dim pDS As IDataset
      Set pDS = pRasterBand.RasterDataset
      pRasterlayer.Name = pDS.BrowseName
    Else
      pRasterlayer.Name = sName
    End If
    
    Set ReturnRasterLayer = pRasterlayer
    
    Set pRasterBands = Nothing
    Set pRasterBand = Nothing
    Set pDS = Nothing
    
Exit Function
EH:
    MsgBox "Util.AddRasterLayer: " & Err.Description
End Function

Public Function GetFeatureClass(strFeatClass As String, pApp As IApplication) As IFeatureClass
  On Error GoTo ErrorHandler

    
    Dim pMap As IMap
  
    If (TypeOf pApp Is IMxApplication) Then
        Dim pMxDoc As IMxDocument
        Dim pLayer As IFeatureLayer
        
        Set pMxDoc = pApp.Document
        Set pMap = pMxDoc.ActiveView.FocusMap
    
        Set pLayer = pMap.Layer(GetLayerIndex(strFeatClass, pApp))
        Set GetFeatureClass = pLayer.FeatureClass
    
    End If
  


  Exit Function
ErrorHandler:
  HandleError True, "GetFeatureClass " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Function

Public Function GetLayerIndex(strLayerName As String, pApp As IApplication) As Long
  On Error GoTo ErrorHandler

    
    Dim pMxDoc As IMxDocument
    Dim pMap As IMap
    Dim i As Integer
    
    Set pMxDoc = pApp.Document
    Set pMap = pMxDoc.FocusMap
    
    i = 0
    
    For i = 0 To pMap.LayerCount - 1
        If pMap.Layer(i).Name = strLayerName Then
            GetLayerIndex = i
        End If
    Next i
    
    'Cleanup
    Set pMxDoc = Nothing
    Set pMap = Nothing
    
  Exit Function
ErrorHandler:
  HandleError True, "GetLayerIndex " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Function

Public Sub AddLayer(pApp As IApplication, pLayer As ILayer)

    On Error GoTo EH
    
    Dim pMap As IBasicMap
    
    If (TypeOf pApp Is IMxApplication) Then
      Dim pMxDoc As IMxDocument
      Set pMxDoc = pApp.Document
      Set pMap = pMxDoc.ActiveView.FocusMap
    End If
    pMap.AddLayer pLayer
    Exit Sub
EH:
    MsgBox "Util.AddLayer: " & Err.Description
End Sub

Public Function BrowseForFileName(strType As String, frm As Form, strTitle As String) As String

On Error GoTo ERH
    Dim pGxObject As IGxObject
    '  Dim pFilter As GxFilterRasterDatasets
    Dim pFilter As IGxObjectFilter
    Dim pMiniBrowser As IGxDialog
    Dim pEnumGxObject As IEnumGxObject
    Dim pDatasetName As IDatasetName
    Dim strSlash As String
    Set pMiniBrowser = New GxDialog
    
    Select Case strType
        Case "Feature"
            Set pFilter = New GxFilterShapefiles
            strSlash = "\"
        Case "Raster"
            Set pFilter = New GxFilterRasterDatasets
            strSlash = ""
    End Select
    
    Set pMiniBrowser.ObjectFilter = pFilter
    pMiniBrowser.Title = strTitle
    
    If (pMiniBrowser.DoModalOpen(frm.hwnd, pEnumGxObject)) Then
      Set pGxObject = pEnumGxObject.Next
      Dim pGxDataset As IGxDataset
      Set pGxDataset = pGxObject
      Dim pDataset As esriGeoDatabase.IDataset
      Set pDataset = pGxDataset.Dataset
      Set pDatasetName = pDataset.FullName
    End If
    
    If strType = "Feature" Then
        BrowseForFileName = pDataset.Workspace.PathName & strSlash & pDataset.Name & ".shp"
    Else
        BrowseForFileName = pDataset.Workspace.PathName & strSlash & pDataset.Name
    End If
    
    
      Set pGxObject = Nothing
      Set pMiniBrowser = Nothing
    Exit Function
ERH:
      MsgBox "AddInputfromBrowser:" & Err.Descript
End Function


Public Function AddInputFromGxBrowser(cboInput As ComboBox, frm As Form, strType As String) As String
On Error GoTo ERH
    Dim pGxObject As IGxObject
    '  Dim pFilter As GxFilterRasterDatasets
    Dim pFilter As IGxObjectFilter
    Dim pMiniBrowser As IGxDialog
    Dim pEnumGxObject As IEnumGxObject
    
    Set pMiniBrowser = New GxDialog
    
    Select Case strType
        Case "Feature"
            Set pFilter = New GxFilterFeatureClasses
        Case "Raster"
            Set pFilter = New GxFilterRasterDatasets
    End Select
    
    Set pMiniBrowser.ObjectFilter = pFilter
    pMiniBrowser.Title = "Select Dataset"
    
    If (pMiniBrowser.DoModalOpen(frm.hwnd, pEnumGxObject)) Then
      Set pGxObject = pEnumGxObject.Next
      Dim pGxDataset As IGxDataset
      Set pGxDataset = pGxObject
      Dim pDataset As esriGeoDatabase.IDataset
      Set pDataset = pGxDataset.Dataset
      cboInput.AddItem pDataset.Name
    Else
        AddInputFromGxBrowser = ""
        Exit Function
    End If
    
    Select Case strType
        Case "Feature"
            AddInputFromGxBrowser = pDataset.Workspace.PathName & "\" & pDataset.Name
        Case "Raster"
            AddInputFromGxBrowser = pDataset.Workspace.PathName & pDataset.Name
    End Select
        
      Set pGxObject = Nothing
      Set pMiniBrowser = Nothing
    Exit Function
ERH:
      MsgBox "AddInputfromBrowser:" & Err.Description
End Function

' Used in "Browse..." buttons on various forms to pick out rasters.
' Have to love this function, by the way.  Not only does it return a rasterdataset but it also populates
' a text box that you send it w/ said rasterdataset path.  More like a Sunction this is.
Public Function AddInputFromGxBrowserText(txtInput As TextBox, strTitle As String, frm As Form, intType As Integer) As IRasterDataset

On Error GoTo ErrHandler:
  
    Dim pGxObject As IGxObject
    Dim pFilter As IGxObjectFilter
    Dim pMiniBrowser As IGxDialog
    Dim pEnumGxObject As IEnumGxObject
    
    Set pMiniBrowser = New GxDialog
    
    Select Case intType 'Set up for future use case we add other filters.
        Case 0
          Set pFilter = New GxFilterRasterDatasets 'Filter for Rasters

    End Select
            
    Set pMiniBrowser.ObjectFilter = pFilter
    pMiniBrowser.Title = strTitle
    
    If (pMiniBrowser.DoModalOpen(frm.hwnd, pEnumGxObject)) Then
      Set pGxObject = pEnumGxObject.Next
      Dim pGxDataset As IGxDataset
      Set pGxDataset = pGxObject
      
      Dim pDataset As esriGeoDatabase.IDataset
      Set pDataset = pGxDataset.Dataset
      txtInput.Text = pDataset.Workspace.PathName & pDataset.Name
    
    Else
        
        Set AddInputFromGxBrowserText = Nothing
        Exit Function
    
    End If
        
        Set AddInputFromGxBrowserText = pDataset
        Set pGxObject = Nothing
        Set pMiniBrowser = Nothing
    
    Exit Function
    
ErrHandler:
    MsgBox "The file you have choosen is not a valid GRID dataset.  Please select another.", vbCritical, "Invalid Data Type"

End Function


' Brings up GxDialog to get output name for specified type. Returns true if user defines output name or
' false if operation cancelled.
Public Function BrowseForOutputName(ByRef sName As String, ByRef sLocation As String, eType As esriDatasetType, hParentHwnd As OLE_HANDLE) As Boolean
    On Error GoTo EH
    Dim pGxDialog As IGxDialog
    Set pGxDialog = New GxDialog
    Dim pFilterCol As IGxObjectFilterCollection
    Set pFilterCol = pGxDialog
    
    Select Case eType
      Case esriDTFeatureClass
        pGxDialog.Title = "Save Features As"
        Dim pFilter As IGxObjectFilter
        Set pFilter = New GxFilterShapefiles
        pFilterCol.AddFilter pFilter, True
      Case esriDTRasterDataset
        pGxDialog.Title = "Save Raster As"
        Dim pFormatFilter As New clsRasterFilter
        pFormatFilter.format = "TIFF"
        pFilterCol.AddFilter pFormatFilter, True
        Set pFormatFilter = New clsRasterFilter
        pFormatFilter.format = "IMAGINE Image"
        pFilterCol.AddFilter pFormatFilter, True
        Set pFormatFilter = New clsRasterFilter
        pFormatFilter.format = "GRID"
        pFilterCol.AddFilter pFormatFilter, True
      Case Else
    End Select
    
    pGxDialog.StartingLocation = sLocation
    
    If (pGxDialog.DoModalSave(hParentHwnd)) Then ' using this form as parent keeps dialog forward
      Dim pGxObject As IGxObject             ' of parent application when pressing save key on browser
      Set pGxObject = pGxDialog.FinalLocation
      sLocation = pGxObject.FullName
      
      sName = pGxDialog.Name
      Dim iPos As Integer
      iPos = InStr(sName, ".")
      If iPos > 0 Then sName = Left(sName, iPos - 1)
      Dim sFormat As String
      sFormat = pGxDialog.ObjectFilter.Name
      Select Case sFormat
          Case "IMAGINE Image"
              sName = sName & ".img"
          Case "TIFF"
              sName = sName & ".tif"
      End Select
      sName = sLocation & "\" & sName
      'MsgBox pGxObject.Category
      BrowseForOutputName = True
    Else
      BrowseForOutputName = False
    End If
    
    Exit Function
EH:
    BrowseForOutputName = False
End Function
Public Function GetUniqueFeatureClassName(pWS As IFeatureWorkspace, Optional sPrefix) As String
    
   On Error GoTo EH
    Dim sPreName As String
    If (sPrefix <> "") Then
      sPreName = sPrefix
    Else
      sPreName = "ai"
    End If
    Dim i As Long
    i = 1
    
    ' if shapefile workspace add .shp extension
    Dim pDataset As IDataset
    Set pDataset = pWS
    If (InStr(UCase(pDataset.Category), "SHAPEFILE") > 0) Then
      Dim sExt As String
      sExt = ".shp"
    Else
      sExt = ""
    End If
    
    Dim done As Boolean
    Do While Not done
      Dim pFC As IFeatureClass
      Dim sName As String
      sName = sPreName & Right(Str(i), Len(Str(i) - 1)) & sExt
      Set pFC = pWS.OpenFeatureClass(sName) ' this can raise an error if dataset doesn't exist so error handler should use it
      If (pFC Is Nothing) Then
        done = True
      Else
        i = i + 1
      End If
    Loop
    GetUniqueFeatureClassName = sName
    Exit Function
EH:
    GetUniqueFeatureClassName = sName
End Function

' sSuffix should include '.'
Public Function GetUniqueFileName(sDir As String, Optional sPrefix As String = "ai", Optional sSuffix As String = "") As String
  On Error GoTo ErrorHandler

  Dim fs As FileSystemObject
  Set fs = New FileSystemObject
  Dim i As Long
  Dim done As Boolean
  Dim Name As String
  done = False
  i = 1
  ' work whether or not input dir has "\" on end - like raster analysis workspace will
  Dim sDirNew As String
  sDirNew = sDir
  If (Right(sDir, 1) = "\") Then
    sDirNew = Left(sDir, Len(sDir) - 1)
  End If
  Do While Not done
    Name = sPrefix & Right(Str(i), Len(Str(i)) - 1) & sSuffix ' make sure to remove space the 'Str' function places before number
    If ((Not fs.FolderExists(sDirNew + "\" + Name)) And _
        (Not fs.FileExists(sDirNew + "\" + Name))) Then
      GetUniqueFileName = Name
      Exit Function
    End If
    i = i + 1
  Loop


  Exit Function
ErrorHandler:
  HandleError True, "GetUniqueFileName " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Function
'Returns a Workspace given for example C:\temp\dataset returns C:\temp
Public Function SplitWorkspaceName(sWholeName As String) As String
    On Error GoTo ERH
    Dim pos As Integer
    pos = InStrRev(sWholeName, "\")
    If pos > 0 Then
        SplitWorkspaceName = Mid(sWholeName, 1, pos - 1)
    Else
        Exit Function
    End If
    Exit Function
ERH:
    MsgBox "Workspace Split:" & Err.Description
End Function

'Returns a filename given for example C:\temp\dataset returns dataset
Public Function SplitFileName(sWholeName As String) As String
    On Error GoTo ERH
    Dim pos As Integer
    Dim sT, sName As String
    pos = InStrRev(sWholeName, "\")
    If pos > 0 Then
        sT = Mid(sWholeName, 1, pos - 1)
        If pos = Len(sWholeName) Then
            Exit Function
        End If
        sName = Mid(sWholeName, pos + 1, Len(sWholeName) - Len(sT))
        pos = InStr(sName, ".")
        If pos > 0 Then
            SplitFileName = Mid(sName, 1, pos - 1)
        Else
            SplitFileName = sName
        End If
    End If
    Exit Function
ERH:
    MsgBox "Workspace Split:" & Err.Description
End Function

Public Function MakePerminentGrid(pRaster As IRaster, sOutPath As String, sOutName As String) As String
    
   On Error GoTo ERH
    Dim pWS As IWorkspace
    Set pWS = SetRasterWorkspace(sOutPath)
    Dim pBandC As IRasterBandCollection
    Dim pBand As IRasterBand
    Set pBandC = pRaster
    Set pBand = pBandC.Item(0)
    Dim pRDS As IRasterDataset
    Set pRDS = pBand.RasterDataset
    Dim pDS As IDataset
    If Not pRDS.CanCopy Then
        Exit Function
    End If
    pRDS.Copy sOutName, pWS
    MakePerminentGrid = sOutPath & sOutName
    Set pRDS = Nothing
    Set pDS = Nothing
    Set pWS = Nothing
    Exit Function
ERH:
    MsgBox "MakePerminebtGrid:" & Err.Description
End Function

Public Function ReturnPermanentRaster(pRaster As IRaster, sOutputPath As String, sOutputName As String) As IRaster

On Error GoTo ERH

    If sOutputPath = "" Or sOutputName = "" Then
        Set ReturnPermanentRaster = Nothing
        Exit Function
    End If
    Dim iPos As Integer
    iPos = InStr(sOutputName, ".")
    Dim sExt As String
    If iPos > 0 Then
        sExt = Mid(sOutputName, iPos + 1)
    Else
        sExt = ""
    End If
    Dim sFormat As String
    Select Case sExt
        Case ""
            sFormat = "GRID"
        Case "tif"
            sFormat = "TIFF"
        Case "img"
            sFormat = "IMAGINE Image"
        Case Else
            MsgBox "Make Permanent Raster: Unsupported file extension"
            Exit Function
    End Select

    Dim pWS As IRasterWorkspace
    Set pWS = modUtil.SetRasterWorkspace(sOutputPath)
    Dim pBandC As IRasterBandCollection
    Dim pRasterDataset As IRasterDataset
    
    Set pBandC = pRaster
    pBandC.SaveAs sOutputName, pWS, sFormat
    
    Set pRasterDataset = pWS.OpenRasterDataset(sOutputName)
    
    Set ReturnPermanentRaster = pRasterDataset.CreateDefaultRaster
    Exit Function
ERH:
    MsgBox "Return Permanent Raster:" & Err.Description
End Function

Public Function MakePermanentRaster(pRaster As IRaster, sOutputPath As String, sOutputName As String) As Boolean

On Error GoTo ERH

    Dim pWS As IWorkspace
    Dim pBandC As IRasterBandCollection
    
    Set pWS = modUtil.SetRasterWorkspace(sOutputPath)
    Set pBandC = pRaster
    
    pBandC.SaveAs sOutputName, pWS, "GRID"
    
    MakePermanentRaster = True
    
    Set pWS = Nothing
    Set pBandC = Nothing
    
    Exit Function
ERH:
    MsgBox "Make Permanent Raster:" & Err.Description
    MakePermanentRaster = False
End Function


Public Function CheckSpatialAnalystLicense() As Boolean

On Error GoTo ErrHandler
    
    Dim pLicManager As IExtensionManager
    Dim pLicAdmin As IExtensionManagerAdmin
    Dim saUID As Variant
    Dim pUID As New UID
    Dim v As Variant
    Dim pExtension As IExtension
    Dim pExtensionConfig As IExtensionConfig
    
    Set pLicManager = New ExtensionManager
    Set pLicAdmin = pLicManager
    
    saUID = "esriSpatialAnalystUI.SAExtension.1"
    pUID.Value = saUID
    
    Call pLicAdmin.AddExtension(pUID, v)

    Set pExtension = pLicManager.FindExtension(pUID)
    Set pExtensionConfig = pExtension
    
    pExtensionConfig.State = esriESEnabled
    CheckSpatialAnalystLicense = True
    
Exit Function

ErrHandler:
    MsgBox "Failed in License Checking" & Err.Description
    CheckSpatialAnalystLicense = False
    
End Function
'Populate combobox of choice with RasterLayers in map
Public Sub AddRasterLayerToComboBox(cboBox As ComboBox, pMap As IMap)

On Error GoTo ErrHandler
    cboBox.Clear
    Dim iLyrIndex As Long
    Dim pLyr As ILayer
    ' Add raster layers into  Combobox
    Dim iLayerCount As Integer
    iLayerCount = pMap.LayerCount
    If iLayerCount > 0 Then
        cboBox.Enabled = True
        For iLyrIndex = 0 To iLayerCount - 1
            Set pLyr = pMap.Layer(iLyrIndex)
            If (TypeOf pLyr Is IRasterLayer) Then
                cboBox.AddItem pLyr.Name
                cboBox.ItemData(cboBox.ListCount - 1) = iLyrIndex
            End If
        Next iLyrIndex
        If (cboBox.ListCount > 0) Then
            cboBox.ListIndex = 0
            cboBox.Text = pMap.Layer(cboBox.ItemData(0)).Name
        End If
    End If

    Set pLyr = Nothing
    
Exit Sub
ErrHandler:
    MsgBox "Add Layer to ComboBox:" & Err.Description

End Sub

Public Sub SetFeatureClassName(pFeatClass As IFeatureClass, strName As String)
  On Error GoTo ErrorHandler

    
    Dim pDataset As IDataset
    Set pDataset = pFeatClass
    
    pDataset.Rename strName
    
  Exit Sub
ErrorHandler:
  HandleError True, "SetFeatureClassName " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Public Function LayerInMap(strName As String, pMap As IMap) As Boolean
  On Error GoTo ErrorHandler

    Dim lngLyrIndex As Long
    Dim pLayer As ILayer
    Dim intLayerCount As Integer
    
    intLayerCount = pMap.LayerCount
    If intLayerCount > 0 Then
        For lngLyrIndex = 0 To intLayerCount - 1
            Set pLayer = pMap.Layer(lngLyrIndex)
            If pLayer.Name = strName Then
                LayerInMap = True
                Exit Function
            Else
                LayerInMap = False
            End If
        Next lngLyrIndex
    End If
    
    Set pLayer = Nothing


  Exit Function
ErrorHandler:
  HandleError True, "LayerInMap " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Function


Public Function LayerInMapByFileName(strName As String, pMap As IMap) As Boolean
    
    Dim i As Integer
    Dim pFeatureLayer As IFeatureLayer
    Dim pDataset As IDataset
    Dim strDatasetname As String
    
    For i = 0 To pMap.LayerCount - 1
        
        If TypeOf pMap.Layer(i) Is IFeatureLayer Then
            Set pFeatureLayer = pMap.Layer(i)
            Set pDataset = pFeatureLayer
            
            strDatasetname = Trim(pDataset.Workspace.PathName & "\" & pFeatureLayer.FeatureClass.AliasName)
            
            If Trim(LCase(strDatasetname)) <> Trim(LCase(strName)) Then
                LayerInMapByFileName = False
            Else
                LayerInMapByFileName = True
                Exit For
            End If
            
        End If
    Next i
    

End Function

Public Sub AddLayerFileToMap(strLyr As String, pMap As IMap)
  
    Dim pGxLayer As IGxLayer
    Dim pGxFile As IGxFile
    
    Set pGxLayer = New GxLayer
    Set pGxFile = pGxLayer
  
    pGxFile.path = strLyr
    pMap.AddLayer pGxLayer.Layer

    Set pGxLayer = Nothing
    Set pGxFile = Nothing

End Sub




Public Function AddRasterLayerToMapFromFileName(strName As String, pMap As IMap) As Boolean

On Error GoTo ErrHandler

    Dim pRasterlayer As IRasterLayer
    Set pRasterlayer = New RasterLayer
    
    pRasterlayer.CreateFromFilePath strName
    
    pMap.AddLayer pRasterlayer
    AddRasterLayerToMapFromFileName = True
    
    Set pRasterlayer = Nothing

Exit Function
        
ErrHandler:
AddRasterLayerToMapFromFileName = False

End Function


Public Function AddFeatureLayerToMapFromFileName(strName As String, pMap As IMap, Optional strLyrName As String) As Boolean

On Error GoTo ErrHandler
    
    Dim pWorkspaceFactory As IWorkspaceFactory
    Dim pFeatureWorkspace As IFeatureWorkspace
    Dim pFeatureLayer As IFeatureLayer
    
    Dim strWorkspace As String
    Dim strFeatClass As String
    
    strWorkspace = SplitWorkspaceName(strName)
    strFeatClass = SplitFileName(strName)
    
    'Create a new ShapefileWorkspaceFactory object and open a shapefile folder
    Set pWorkspaceFactory = New ShapefileWorkspaceFactory
    Set pFeatureWorkspace = pWorkspaceFactory.OpenFromFile(strWorkspace, 0)
    
    'Create a new FeatureLayer and assign a shapefile to it
    Set pFeatureLayer = New FeatureLayer
    Set pFeatureLayer.FeatureClass = pFeatureWorkspace.OpenFeatureClass(strFeatClass)
    
    If Len(strLyrName) > 0 Then
        pFeatureLayer.Name = strLyrName
    Else
        pFeatureLayer.Name = pFeatureLayer.FeatureClass.AliasName
    End If
    
    'Add the FeatureLayer to the focus map
    pMap.AddLayer pFeatureLayer
    AddFeatureLayerToMapFromFileName = True
    
    Set pWorkspaceFactory = Nothing
    Set pFeatureWorkspace = Nothing
    Set pFeatureLayer = Nothing
    
Exit Function
        
ErrHandler:

AddFeatureLayerToMapFromFileName = False

End Function

'Public Function ReturnCellCount(pRaster As IRaster) As Double
'  On Error GoTo ErrorHandler
'
'    Dim pTable As ITable
'    Dim pRasterBandCollection As IRasterBandCollection
'    Dim pRasterBand As IRasterBand
'    Dim pQueryFilter As IQueryFilter
'    Dim pCursor As ICursor
'
'    Dim pDataStatistics As IDataStatistics
'    Dim pStatisticsResult As IStatisticsResults
'
'    'QI the feature layer for the table interface
'    Set pRasterBandCollection = pRaster
'    Set pRasterBand = pRasterBandCollection.Item(0)
'
'    Set pTable = pRasterBand
'
'    'Initialise a query and get a cursor to the first row
'    Set pQueryFilter = New QueryFilter
'    pQueryFilter.AddField "Count"
'    Set pCursor = pTable.Search(pQueryFilter, True)
'
'    'Use the statistics objects to calculate the sum of the count field, hence the number of cells
'    Set pDataStatistics = New DataStatistics
'    Set pDataStatistics.Cursor = pCursor
'    pDataStatistics.Field = "Count"
'
'    Set pStatisticsResult = pDataStatistics.Statistics
'    If pStatisticsResult Is Nothing Then
'        MsgBox "Failed to gather stats on the feature class"
'        Exit Function
'    End If
'
'    ReturnCellCount = pStatisticsResult.Sum
'
'    'Cleanup
'    Set pTable = Nothing
'    Set pRasterBandCollection = Nothing
'    Set pRasterBand = Nothing
'    Set pQueryFilter = Nothing
'    Set pCursor = Nothing
'    Set pDataStatistics = Nothing
'    Set pStatisticsResult = Nothing
'
'  Exit Function
'ErrorHandler:
'  HandleError True, "ReturnCellCount " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
'End Function
Public Function ReturnCellCount(pRaster As IRaster) As Double
  On Error GoTo ErrorHandler

    Dim pRasterProps As IRasterProps
    Set pRasterProps = pRaster
    
    ReturnCellCount = pRasterProps.Height * pRasterProps.Width
   
    'Cleanup
    Set pRasterProps = Nothing

  Exit Function
ErrorHandler:
  HandleError True, "ReturnCellCount " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Function

Public Function GetRasterDistanceUnits(strLayerName As String, pApp As IApplication) As Integer
    
    Dim pMxDoc As IMxDocument
    Dim pMap As IMap
    Dim pRasterlayer As IRasterLayer
    Dim pRasterDataset As IRasterDataset
    Dim pDistUnit As ILinearUnit
    Dim intUnit As Integer
    Dim pProjCoord As IProjectedCoordinateSystem
    
On Error GoTo ErrHandler:
    Set pMxDoc = pApp.Document
    Set pMap = pMxDoc.FocusMap
    
    Set pRasterlayer = pMap.Layer(modUtil.GetLayerIndex(strLayerName, pApp))
    
    If pRasterlayer Is Nothing Then
        Exit Function
    Else
    
    Set pRasterDataset = pRasterlayer.Raster
        
        If CheckSpatialReference(pRasterDataset) Is Nothing Then
            
            MsgBox "The GRID you have choosen has no spatial reference information.  Please define a projection before continuing.", vbExclamation, "No Project Information Detected"
            
            Exit Function
        
        Else
        
            Set pProjCoord = CheckSpatialReference(pRasterDataset)
            Set pDistUnit = pProjCoord.CoordinateUnit
            intUnit = pDistUnit.MetersPerUnit
            
                If intUnit = 1 Then
                    GetRasterDistanceUnits = 0
                Else
                    GetRasterDistanceUnits = 1
                End If
    
        End If
    
    End If
    
    Set pRasterlayer = Nothing
    Set pRasterDataset = Nothing
    Set pDistUnit = Nothing
    Set pProjCoord = Nothing

Exit Function
ErrHandler:
    Exit Function

End Function

'Using the layer name of a Raster, chases down the Filename of the dataset
Public Function GetRasterFileName(strRasterName As String, pApp As IApplication) As String
  On Error GoTo ErrorHandler

        
    Dim pMxDoc As IMxDocument
    Dim pMap As IMap
    Dim pRasterlayer As IRasterLayer
    
    Set pMxDoc = pApp.Document
    Set pMap = pMxDoc.FocusMap
    
    Set pRasterlayer = pMap.Layer(GetLayerIndex(strRasterName, pApp))
    
    GetRasterFileName = pRasterlayer.FilePath
    
    Set pMxDoc = Nothing
    Set pMap = Nothing
    Set pRasterlayer = Nothing
    


  Exit Function
ErrorHandler:
  HandleError True, "GetRasterFileName " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Function

'Takes a layer name, returns filename
Public Function GetFeatureFileName(strFeatureName As String, pApp As IApplication) As String
  On Error GoTo ErrorHandler


    Dim pMxDoc As IMxDocument
    Dim pMap As IMap
    Dim pFeatureLayer As IFeatureLayer
    Dim pDataset As IDataset
    Dim pFeatureClass As IFeatureClass
    
    Set pMxDoc = pApp.Document
    Set pMap = pMxDoc.FocusMap
    
    If Len(strFeatureName) > 0 Then
        Set pFeatureLayer = pMap.Layer(GetLayerIndex(strFeatureName, pApp))
        Set pDataset = pFeatureLayer
              Set pFeatureClass = pFeatureLayer.FeatureClass
        GetFeatureFileName = pDataset.Workspace.PathName & "\" & pFeatureClass.AliasName
    Else
        Exit Function
    End If
    
    'Cleanup
    Set pMxDoc = Nothing
    Set pMap = Nothing
    Set pFeatureLayer = Nothing
    Set pDataset = Nothing



  Exit Function
ErrorHandler:
  HandleError True, "GetFeatureFileName " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Function

Public Function ReturnRaster(strRasterFileName As String) As IRaster
    
     'Takes in FileName, checks for existance of data.
On Error GoTo ErrHandler:

    ' Create RasterWorkSpaceFactory
    Dim pWSF As IWorkspaceFactory
    Set pWSF = New RasterWorkspaceFactory
    
    ' Get RasterWorkspace
    Dim pRasterWS As IRasterWorkspace
    Dim pRasterDS As IRasterDataset
    
    If pWSF.IsWorkspace(modUtil.SplitWorkspaceName(strRasterFileName)) Then
        Set pRasterWS = pWSF.OpenFromFile(modUtil.SplitWorkspaceName(strRasterFileName), 0)
        Set pRasterDS = pRasterWS.OpenRasterDataset(modUtil.SplitFileName(strRasterFileName))
        Set ReturnRaster = pRasterDS.CreateDefaultRaster
    End If
        
    Set pWSF = Nothing
    Set pRasterWS = Nothing
    Set pRasterDS = Nothing
    
    Exit Function

ErrHandler:
    Set ReturnRaster = Nothing
    MsgBox Err.Description
End Function

Public Function RasterExists(strRasterFileName As String) As Boolean
    
    'Takes in FileName, checks for existance of data.
On Error GoTo ErrHandler:

    ' Create RasterWorkSpaceFactory
    Dim pWSF As IWorkspaceFactory
    Set pWSF = New RasterWorkspaceFactory
    
    ' Get RasterWorkspace
    Dim pRasterWS As IRasterWorkspace
    Dim pRasterDS As IRasterDataset
    
    If pWSF.IsWorkspace(modUtil.SplitWorkspaceName(strRasterFileName)) Then
        Set pRasterWS = pWSF.OpenFromFile(modUtil.SplitWorkspaceName(strRasterFileName), 0)
        Set pRasterDS = pRasterWS.OpenRasterDataset(modUtil.SplitFileName(strRasterFileName))
    End If
    
    RasterExists = (Not pRasterDS Is Nothing)
    
    Set pWSF = Nothing
    Set pRasterWS = Nothing
    Set pRasterDS = Nothing
    
    Exit Function

ErrHandler:
    RasterExists = False
    MsgBox Err.Description
End Function

Public Function FeatureExists(strFeatureFileName As String) As Boolean

On Error GoTo ErrHandler:
    
    Dim pWSF As IWorkspaceFactory
    Set pWSF = New ShapefileWorkspaceFactory
    
    Dim pFeatWS As IFeatureWorkspace
    Dim pFeatDS As IFeatureClass

    Dim strWorkspace As String
    Dim strFeatDS As String
    strWorkspace = modUtil.SplitWorkspaceName(strFeatureFileName) & "\"
    strFeatDS = modUtil.SplitFileName(strFeatureFileName)
    
    If pWSF.IsWorkspace(strWorkspace) Then
        Set pFeatWS = pWSF.OpenFromFile(strWorkspace, 0)
        Set pFeatDS = pFeatWS.OpenFeatureClass(strFeatDS)
    End If
    
    FeatureExists = (Not pFeatDS Is Nothing)
    
    Set pWSF = Nothing
    Set pFeatWS = Nothing
    Set pFeatDS = Nothing

Exit Function

ErrHandler:
    FeatureExists = False
    
End Function

Public Function ReturnFeatureClass(strFeatureFileName As String) As IFeatureClass
On Error GoTo ErrHandler:
    
    Dim pWSF As IWorkspaceFactory
    Set pWSF = New ShapefileWorkspaceFactory
    
    Dim pFeatWS As IFeatureWorkspace
    Dim pFeatDS As IFeatureClass

    Dim strWorkspace As String
    Dim strFeatDS As String
    strWorkspace = modUtil.SplitWorkspaceName(strFeatureFileName) & "\"
    strFeatDS = modUtil.SplitFileName(strFeatureFileName)
    
    If pWSF.IsWorkspace(strWorkspace) Then
        Set pFeatWS = pWSF.OpenFromFile(strWorkspace, 0)
        Set pFeatDS = pFeatWS.OpenFeatureClass(strFeatDS)
    End If
    
    If Not pFeatDS Is Nothing Then
        Set ReturnFeatureClass = pFeatDS
    Else
        MsgBox "Featureclass " & strFeatureFileName & "does not exist.", vbCritical, "Featureclass Not Found"
        Exit Function
    End If
    
    
    Set pWSF = Nothing
    Set pFeatWS = Nothing
    Set pFeatDS = Nothing

Exit Function

ErrHandler:
    MsgBox "Error Returning Featureclass.", vbCritical, "Error"
    
    
End Function


Public Function CreateMask(pMaskRaster As IGeoDataset) As IRaster
  On Error GoTo ErrorHandler

'General function for creating analysis masks.  Uses an incoming
'raster as the mask and a con statement.

    Dim pMapAlgebraOp As IMapAlgebraOp
    Dim strExpression As String
    
    Set pMapAlgebraOp = New RasterMapAlgebraOp
    
    With pMapAlgebraOp
        .BindRaster pMaskRaster, "mask"
    End With
    
    strExpression = "con(isnull([mask]),0,1)"
    
    Set CreateMask = pMapAlgebraOp.Execute(strExpression)
    
    'Cleanup
    Set pMapAlgebraOp = Nothing
    Set pMaskRaster = Nothing
    


  Exit Function
ErrorHandler:
  HandleError True, "CreateMask " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Function

Public Function ReturnRasterStretchColorRampRender(pRasterlayer As IRasterLayer, strColor As String) As IRasterStretchColorRampRenderer
  On Error GoTo ErrorHandler

    
    ' Get raster input from layer
    Dim pRLayer As IRasterLayer
    Dim pRaster As IRaster
    
    ' Create renderer and QI RasterRenderer
    Dim pRasStretchMM As IRasterStretch
    Dim pStretchRen As IRasterStretchColorRampRenderer
    Dim pRasRen As IRasterRenderer
    
    Dim strHighValue As String
    Dim strLowValue As String
    
    Set pRLayer = pRasterlayer
    Set pRaster = pRLayer.Raster
    
    'init the stretch renderer
    Set pStretchRen = New RasterStretchColorRampRenderer
    Set pRasRen = pStretchRen
    Set pRasStretchMM = pStretchRen
    
    'Added 8/16, Andrew wanted Min/Max not standard deviation
    With pRasStretchMM
        .StretchType = esriRasterStretch_StandardDeviations 'esriRasterStretch_MinimumMaximum: changed 10/19/2007 per D. Eslinger
    End With
    
    ' Set raster for the renderer and update
    Set pRasRen.Raster = pRaster
    pRasRen.Update

     ' Define two colors
    Dim pFromColor As IHsvColor
    Dim pToColor As IHsvColor
    
    Dim lngHueTo As Long
    Dim lngSatTo As Long
    Dim lngValueTo As Long
    
    Dim lngHueFrom As Long
    Dim lngSatFrom As Long
    Dim lngValueFrom As Long
    
    Set pFromColor = New HsvColor
    Set pToColor = New HsvColor
    
    Select Case strColor
        Case "Blue"
            
            lngHueFrom = 226
            lngSatFrom = 5
            lngValueFrom = 100
            
            lngHueTo = 226
            lngSatTo = 93
            lngValueTo = 100
        
        Case "Brown"
            
            lngHueFrom = 39
            lngSatFrom = 15
            lngValueFrom = 100
            
            lngHueTo = 40
            lngSatTo = 100
            lngValueTo = 69
            
        Case Else
            
            lngHueTo = Split(strColor, ",")(0)
            lngSatTo = Split(strColor, ",")(1)
            lngValueTo = Split(strColor, ",")(2)
            
            lngHueFrom = Split(strColor, ",")(3)
            lngSatFrom = Split(strColor, ",")(4)
            lngValueFrom = Split(strColor, ",")(5)
            
    End Select
    
    With pFromColor
        .Hue = lngHueFrom
        .Saturation = lngSatFrom
        .Value = lngValueFrom
    End With
    
    With pToColor
        .Hue = lngHueTo
        .Saturation = lngSatTo
        .Value = lngValueTo
    End With
        
        
    ' Create color ramp
    Dim pRamp As IAlgorithmicColorRamp
    Set pRamp = New AlgorithmicColorRamp
    
    With pRamp
        .Size = 2
        .Algorithm = esriCIELabAlgorithm
        .FromColor = pFromColor
        .ToColor = pToColor
        .CreateRamp True
    End With
    
    
    ' Plug this colorramp into renderer and select a band
    pStretchRen.BandIndex = 0
    pStretchRen.ColorRamp = pRamp
    
    'Format the labels
    
    strHighValue = Right(pStretchRen.LabelHigh, Len(pStretchRen.LabelHigh) - 7)
    strLowValue = Right(pStretchRen.LabelLow, Len(pStretchRen.LabelLow) - 6)
           
    pStretchRen.LabelHigh = "High : " & format(strHighValue, "###,###,###,###,###,##0.00")
    pStretchRen.LabelLow = "Low : " & format(strLowValue, "###,###,###,###,###,##0.00")
    
    ' Update the renderer with new settings and plug into layer
    pRasRen.Update
    Set ReturnRasterStretchColorRampRender = pStretchRen
    
    'Release memeory
    Set pRLayer = Nothing
    Set pRaster = Nothing
    Set pStretchRen = Nothing
    Set pRasRen = Nothing
    Set pRamp = Nothing
    Set pToColor = Nothing
    Set pFromColor = Nothing
    


  Exit Function
ErrorHandler:
  HandleError True, "ReturnRasterStretchColorRampRender " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Function

Public Function ReturnUniqueRasterRenderer(pRasterlayer As IRasterLayer, strStandardName As String)

On Error GoTo ErrHandler:
  
   ' Get raster input from layer
    Dim pRLayer As IRasterLayer
    Dim pRaster As IRaster
    Dim pTable As ITable
    Dim pBand As IRasterBand
    Dim pBandCol As IRasterBandCollection
    Dim booTableExist As Boolean
    Dim intNumValues As Integer
    Dim strFieldName As String
    Dim lngFieldIndex As Long
    Dim pColorRed As IColor
    Dim pColorGreen As IColor
    Dim pFSymbol As ISimpleFillSymbol
    Dim pUVRen As IRasterUniqueValueRenderer
    Dim pRasRen As IRasterRenderer
    Dim i As Long
    Dim pRow As iRow
    Dim LabelValue As Variant
         
    Set pRLayer = pRasterlayer
    Set pRaster = pRLayer.Raster

   'Get the number of rows from raster table
    Set pBandCol = pRaster
    Set pBand = pBandCol.Item(0)
    pBand.HasTable booTableExist
    If Not booTableExist Then Exit Function
    
    Set pTable = pBand.AttributeTable
    intNumValues = pTable.RowCount(Nothing)
  
    'Specified a field and get the field index for the specified field to be rendered.     Dim FieldIndex As Integer
    strFieldName = "Value"   'Value is the default field, you can specify other field here..
    lngFieldIndex = pTable.FindField(strFieldName)
    
    'Create two colors, red, green
    Set pColorRed = New RgbColor
    Set pColorGreen = New RgbColor
    
    pColorRed.RGB = RGB(214, 71, 0)
    pColorGreen.RGB = RGB(56, 168, 0)
  
   ' Create UniqueValue renderer and QI RasterRenderer
    Set pUVRen = New RasterUniqueValueRenderer
    Set pRasRen = pUVRen
  
   ' Connect renderer and raster
    Set pRasRen.Raster = pRaster
    pRasRen.Update
  
   ' Set UniqueValue renerer
    pUVRen.HeadingCount = 1   ' Use one heading
    pUVRen.Heading(0) = strStandardName
    pUVRen.ClassCount(0) = intNumValues
    pUVRen.Field = strFieldName
    
    For i = 0 To intNumValues - 1
       Set pRow = pTable.GetRow(i) 'Get a row from the table
        If pRow.Value(lngFieldIndex) = 1 Then
            LabelValue = pRow.Value(lngFieldIndex)  ' Get value of the given index
            pUVRen.AddValue 0, i, LabelValue  'Set value for the renderer
            pUVRen.Label(0, i) = "Exceeds Standard"  ' Set label
            Set pFSymbol = New SimpleFillSymbol
            pFSymbol.Color = pColorRed
            pUVRen.Symbol(0, i) = pFSymbol  'Set symbol
        Else
            LabelValue = pRow.Value(lngFieldIndex)   ' Get value of the given index
            pUVRen.AddValue 0, i, LabelValue  'Set value for the renderer
            pUVRen.Label(0, i) = "Below Standard"  ' Set label
            Set pFSymbol = New SimpleFillSymbol
            pFSymbol.Color = pColorGreen
            pUVRen.Symbol(0, i) = pFSymbol  'Set symbol
        End If
    Next i
  
    'Update render and refresh layer
    pRasRen.Update
    Set ReturnUniqueRasterRenderer = pUVRen
    
    'Clean up
    Set pRLayer = Nothing
    Set pUVRen = Nothing
    Set pRasRen = Nothing
    Set pFSymbol = Nothing
    Set pRaster = Nothing
    Set pRLayer = Nothing
    Set pBand = Nothing
    Set pBandCol = Nothing
    Set pTable = Nothing
    Set pRow = Nothing
    Exit Function
ErrHandler:
    MsgBox Err.Description
End Function

Public Function ReturnHSVColorString() As String
  On Error GoTo ErrorHandler

'Returns a comma delimited string of 6 values.  1st 3 a 'To Color' - HIGH, 2nd 3 a 'From Color' - LOW
    Dim intHue As Integer
    
    'Hue is a value from 1 to 360 so find a random one
    intHue = Int((360 * Rnd) + 1)
    
    'Value will be a constant of 97, 100 in the SV and 5, 100..
    ReturnHSVColorString = CStr(intHue) & ",97,100," & CStr(intHue) & ",5,100"
    
  Exit Function
ErrorHandler:
  HandleError True, "ReturnHSVColorString " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Function


Public Sub CleanupRasterFolder(strWorkspacePath As String)
  On Error GoTo ErrorHandler

'Used to cleanup the User's workspace and avoid the dreaded -2147467259 error

    Dim pWorkspaceFactory As IWorkspaceFactory
    Dim pWorkspace As IWorkspace
    Dim pDataset As IDataset
    Dim pEnumRasterDataset As IEnumDataset
    Dim pEnv As IRasterAnalysisEnvironment
    
    Set pWorkspaceFactory = New RasterWorkspaceFactory
    Set pWorkspace = pWorkspaceFactory.OpenFromFile(strWorkspacePath, 0)
    
    Set pEnumRasterDataset = pWorkspace.Datasets(esriDTRasterDataset)
    Set pEnv = New RasterAnalysis
        
    Set pDataset = pEnumRasterDataset.Next
    
    Do While Not pDataset Is Nothing
        If InStr(1, pDataset.Name, pEnv.DefaultOutputRasterPrefix, vbTextCompare) > 0 Then
            If pDataset.CanDelete Then
                pDataset.Delete
            End If
        End If
        Set pDataset = pEnumRasterDataset.Next
    Loop
    
    'Cleanup
    Set pWorkspaceFactory = Nothing
    Set pWorkspace = Nothing
    Set pDataset = Nothing
    Set pEnumRasterDataset = Nothing
    Set pEnv = Nothing
      
  Exit Sub
ErrorHandler:
  HandleError True, "CleanupRasterFolder " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Public Sub CleanGlobals()
  On Error GoTo ErrorHandler

'Sub to rid the world of stragling GRIDS, i.e. the ones established for global usse

    If Not modMainRun.g_pFeatWorkspace Is Nothing Then
        Set modMainRun.g_pFeatWorkspace = Nothing
    End If
    
    If Not modMainRun.g_pRasWorkspace Is Nothing Then
        Set modMainRun.g_pRasWorkspace = Nothing
    End If
    
    'Had an 'elegant' solution using an Iarray to hold global rasters, but that didn't seem to do the
    'job, so we have to manually set each and everyone to nothing

    If Not g_pSCS100Raster Is Nothing Then
        Set g_pSCS100Raster = Nothing
    End If

    If Not g_pAbstractRaster Is Nothing Then
        Set g_pAbstractRaster = Nothing
    End If

    If Not g_pRunoffRaster Is Nothing Then
        Set g_pRunoffRaster = Nothing
    End If

    If Not g_pRunoffInchRaster Is Nothing Then
        Set g_pRunoffInchRaster = Nothing
    End If

    If Not g_pCellAreaSqMiRaster Is Nothing Then
        Set g_pCellAreaSqMiRaster = Nothing
    End If

    If Not g_pRunoffCFRaster Is Nothing Then
        Set g_pRunoffCFRaster = Nothing
    End If

    If Not g_pRunoffAFRaster Is Nothing Then
        Set g_pRunoffAFRaster = Nothing
    End If

    If Not g_pMetRunoffRaster Is Nothing Then
        Set g_pMetRunoffRaster = Nothing
    End If

    If Not g_pRunoffRaster Is Nothing Then
        Set g_pRunoffRaster = Nothing
    End If
    
    If Not g_pDEMRaster Is Nothing Then
        Set g_pDEMRaster = Nothing
    End If
    
    If Not g_pFlowAccRaster Is Nothing Then
        Set g_pFlowAccRaster = Nothing
    End If
    
    If Not g_pFlowDirRaster Is Nothing Then
        Set g_pFlowDirRaster = Nothing
    End If
    
    If Not g_pLSRaster Is Nothing Then
        Set g_pLSRaster = Nothing
    End If
    
    If Not g_pWaterShedFeatClass Is Nothing Then
        Set g_pWaterShedFeatClass = Nothing
    End If
    
    If Not g_KFactorRaster Is Nothing Then
        Set g_KFactorRaster = Nothing
    End If
    
    If Not g_pPrecipRaster Is Nothing Then
        Set g_pPrecipRaster = Nothing
    End If
    
    If Not g_pSoilsRaster Is Nothing Then
        Set g_pSoilsRaster = Nothing
    End If
    
    If Not g_LandCoverRaster Is Nothing Then
        Set g_LandCoverRaster = Nothing
    End If
    
    If Not g_pSelectedPolyClip Is Nothing Then
        Set g_pSelectedPolyClip = Nothing
    End If
        


  Exit Sub
ErrorHandler:
  HandleError True, "CleanGlobals " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Public Function GetSelectedFeatureCount(pLayer As ILayer, pMap As IMap) As Integer
  On Error GoTo ErrorHandler

    'Get the # of selected features from a specific layer
    'strName: name of the layer you want to find out # of selected features
    'pMap: Current Map
    
    Dim pFeatLayer As IFeatureLayer
    Dim pFeatureSelection As IFeatureSelection
    
    If TypeOf pLayer Is IFeatureLayer Then
        Set pFeatLayer = pLayer
    Else
        GetSelectedFeatureCount = 0
        Exit Function
    End If
    
    Set pFeatureSelection = pFeatLayer
        
    Dim pSelectionSet As ISelectionSet
    Set pSelectionSet = pFeatureSelection.SelectionSet
    
    Dim pFeatureCursor As IFeatureCursor
    pSelectionSet.Search Nothing, True, pFeatureCursor
        
    GetSelectedFeatureCount = pSelectionSet.Count
    
    'Cleanup
    Set pFeatLayer = Nothing
    Set pFeatureSelection = Nothing
    Set pSelectionSet = Nothing
    Set pFeatureCursor = Nothing
    


  Exit Function
ErrorHandler:
  HandleError True, "GetSelectedFeatureCount " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Function

Public Function ClipBySelectedPoly(pRaster As IRaster, pClipPoly As IGeometry, pEnv As IRasterAnalysisEnvironment) As IRaster
  On Error GoTo ErrorHandler

    'Clip the raster by analysis extent poly
    'pRaster: Incoming raster to be clipped
    'pClipPoly: the polygon doing the clipping
    'pEnv: the Raster analysis environmnet
    
    ' Create the RasterExtractionOp object
    Dim pExtractionOp As IExtractionOp
    Dim pClipPolygon As IPolygon
    
    Set pExtractionOp = New RasterExtractionOp
    Set pClipPolygon = pClipPoly
    
    'Set pEnv = g_pSpatEnv
    Set g_pSpatEnv = pExtractionOp
    
    ' Call the method
    Set ClipBySelectedPoly = pExtractionOp.Polygon(pRaster, pClipPolygon, True)
    
    'Cleanup
    Set pExtractionOp = Nothing
    Set pClipPolygon = Nothing
    'Set pRaster = Nothing
    


  Exit Function
ErrorHandler:
  HandleError True, "ClipBySelectedPoly " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Function

Public Sub ExportLayerToPath(pLayer As ILayer, sPath As String)

    On Error GoTo ExportLayerToPath_ERR
    
    Dim pGxLayer As IGxLayer
    Dim pGxFile As IGxFile
    
    Set pGxLayer = New GxLayer
    Set pGxFile = pGxLayer
    
    pGxFile.path = sPath
    
    Set pGxLayer.Layer = pLayer
    Set pGxFile = Nothing
    Set pGxLayer = Nothing

    Exit Sub
    
ExportLayerToPath_ERR:
    Debug.Print "ExportLayerToPath_ERR: " & Err.Description
    Debug.Assert 0
    
End Sub


Public Sub CreateMetadata(pGroupLayer As IGroupLayer, strProjectInfo As String)
  On Error GoTo ErrorHandler

'This sub will create metadata for the final group layer of raster outputs.  The idea is to use
'a xml template shipped with N-SPECT to provide the base.  Next the metadata is synchronized which fills
'in all of the spatial particulars.  Finally, specific parameters are set.  The xml file is shipped in
'NSPECTDAT\metadata.
    
    Dim fs
    Dim booExists As Boolean

    'Metadata objects
    Dim pGxFile As IGxFile
    Dim pMetaData As IMetadata
    Dim pPropSet As IPropertySet
    Dim pPropXMLSet As IXmlPropertySet
    Dim pSelMD As IMetadata
    Dim pXPS As IXmlPropertySet2
    Dim pXPS2 As IXmlPropertySet2
    
    Dim sPath As String
    Dim strMeta As String
    Dim i As Integer
    Dim lyrCount As Long
    
    'Group layer and raster stuff
    Dim pLayer As ILayer
    Dim pCompositeLayer As ICompositeLayer
    Dim pGroupLayerPiece As ILayer
    Dim pRasterlayer As IRasterLayer
    Dim pRasterDataset As IRasterDataset
    Dim pRasterWorkSpaceFact As IWorkspaceFactory
    Dim pRasterWorkSpace As IRasterWorkspace
    
    sPath = modUtil.g_nspectPath & "\metadata\metadata.xml"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    booExists = fs.FileExists(sPath)
    
    If booExists = False Then
        MsgBox ("The N-SPECT metadata template: " & sPath & " was not found." & vbNewLine & "Please provide an existing Xml Document"), vbExclamation
        Exit Sub
    End If
    
    'Set up the metadata
    Set pMetaData = New GxMetadata
    Set pGxFile = pMetaData
    pGxFile.path = sPath
    
    Set pPropSet = pMetaData.Metadata
    Set pXPS2 = pPropSet
    
    'Get all the Xml from the template
    strMeta = pXPS2.GetXml("")
    
    'Now get ready to get a hold of the Rasterdatasets
    Set pRasterWorkSpaceFact = New RasterWorkspaceFactory
    
    'set up player
    Set pLayer = pGroupLayer
    
    lyrCount = 0
      
    If Not TypeOf pLayer Is IGroupLayer Then
        Exit Sub
    End If
    
    Set pCompositeLayer = pLayer
    
    'We know this grouplayer will have rasters in it, so set up pRasterLayer so we
    'can go from there.
    Set pRasterlayer = pCompositeLayer.Layer(0)
    Set pRasterWorkSpace = pRasterWorkSpaceFact.OpenFromFile(SplitWorkspaceName(pRasterlayer.FilePath), 0)
    
    'Now with the workspace set, go to work on each of the layers
    For i = 0 To pCompositeLayer.Count - 1
                  
        Set pGroupLayerPiece = pCompositeLayer.Layer(i)
        If TypeOf pGroupLayerPiece Is IRasterLayer Then
                
            modProgDialog.ProgDialog "Creating metadata for " & pGroupLayerPiece.Name & "...", "Completing Analysis", 0, pCompositeLayer.Count, lyrCount, frmPrj.m_App.hwnd
            If modProgDialog.g_boolCancel Then
                
                
                Set pRasterlayer = pGroupLayerPiece
                'Have to get the RasterDataset in order to get metadata
                Set pRasterDataset = pRasterWorkSpace.OpenRasterDataset(SplitFileName(pRasterlayer.FilePath))

                Set pSelMD = pRasterDataset
                Set pPropXMLSet = pSelMD.Metadata
                Set pXPS = pPropXMLSet
    
                'Write the template Xml to the selected dataset
                'This will delete any existing xml metadata in the selected dataset
                pXPS.SetXml (strMeta)
                
                'Now set seperate properties
                With pXPS
                    .SetPropertyX "idinfo/citation/citeinfo/pubdate", Now, esriXPTText, esriXSPAAddOrReplace, False
                    .SetPropertyX "idinfo/citation/citeinfo/title", pRasterlayer.Name, esriXPTText, esriXSPAAddOrReplace, False
                    .SetPropertyX "idinfo/descript/supplinf", strProjectInfo, esriXPTText, esriXSPAAddOrReplace, False
                    .SetPropertyX "dataqual/lineage/procstep/procdesc", g_dicMetadata.Item(pRasterlayer.Name), esriXPTText, esriXSPAAddOrReplace, False
                    .SetPropertyX "dataqual/lineage/procstep/procdate", Now, esriXPTText, esriXSPAAddOrReplace, False
                    .SetPropertyX "metainfo/metd", Now, esriXPTText, esriXSPAAddOrReplace, False
                End With
                
                'Save  the metadata
                pSelMD.Metadata = pXPS
                pSelMD.Synchronize esriMSAAccessed, 1
                
                DoEvents
            
            End If
                            
        End If
        lyrCount = lyrCount + 1
    Next i
    
    
    'kill dialog
    modProgDialog.KillDialog
    
    'Cleanup
    Set pGxFile = Nothing
    Set pMetaData = Nothing
    Set pPropSet = Nothing
    Set pPropXMLSet = Nothing
    Set pSelMD = Nothing
    Set pXPS = Nothing
    Set pXPS2 = Nothing
    Set pLayer = Nothing
    Set pCompositeLayer = Nothing
    Set pGroupLayerPiece = Nothing
    Set pRasterlayer = Nothing
    Set pRasterDataset = Nothing
    Set pRasterWorkSpaceFact = Nothing
    Set pRasterWorkSpace = Nothing
    
  Exit Sub
ErrorHandler:
  HandleError True, "CreateMetadata " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Public Function ReturnRasterMax(pRaster As IRaster) As Double
  On Error GoTo ErrorHandler

    Dim pRasterBandCollection As IRasterBandCollection
    Dim pRasterBand As IRasterBand
    Dim pRasterStats As IRasterStatistics
    
    Set pRasterBandCollection = pRaster
    Set pRasterBand = pRasterBandCollection.Item(0)
    
    Set pRasterStats = pRasterBand.Statistics
    
    ReturnRasterMax = pRasterStats.Maximum
    
    'Clean
    Set pRasterBandCollection = Nothing
    Set pRasterBand = Nothing
    Set pRasterStats = Nothing
    
  Exit Function
ErrorHandler:
  HandleError True, "ReturnRasterMax " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Function

Public Function ParseProjectforMetadata(clsPrjXML As clsXMLPrjFile, strPrjFileName As String) As String
    
    Dim strFormattedProject As String
    Dim i As Integer
    
    'Start by getting all the particulars that must be there...
    strFormattedProject = "N-SPECT Project Information: " & vbNewLine & _
                          vbTab & "N-SPECT Project Name: " & clsPrjXML.strProjectName & vbNewLine & _
                          vbTab & "N-SPECT Project File Location: " & strPrjFileName & vbNewLine & _
                          vbTab & "N-SPECT Parameter Database: " & modUtil.g_nspectPath & "\nspect.mdb" & vbNewLine & _
                          vbTab & "Project Working Directory: " & clsPrjXML.strProjectWorkspace & vbNewLine & _
                          vbTab & "Watershed Delineation: " & clsPrjXML.strWaterShedDelin
    ParseProjectforMetadata = strFormattedProject

End Function

Public Function ExportSelectedFeatures(pLayer As ILayer, pMap As IMap, pWorkspace As IWorkspace, strFeatClass As String) As IFeatureClass
  On Error GoTo ErrorHandler

    'Incoming
    'pLayer: Layer user has chosen as being the one from which the selected polys will come
    'pMap: current map
    'pWorkspace:  place to put the exported selected polys
    'strFeatClass: string file location of featureclass
      
    Dim pFLayer As IFeatureLayer
    Dim pFC As IFeatureClass
    Dim pINFeatureClassName As IFeatureClassName
    Dim pDataset As IDataset
    Dim pInDsName As IDatasetName
    Dim pFSel As IFeatureSelection
    Dim pSelSet As ISelectionSet
    Dim pFeatureClassName As IFeatureClassName
    Dim pOutDatasetName As IDatasetName
    Dim pWorkspaceName As IWorkspaceName
    Dim pExportOp As IExportOperation
    Dim pFeatWS As IFeatureWorkspace
    Dim pFeatClass As IFeatureClass
    
    Dim pSelectGeometry As IGeometry
    Dim pBasinFeatClass As IFeatureClass
    Dim pSpatialFilter As ISpatialFilter
    Dim pBasinSelSet As ISelectionSet
        
    'Set the layer to the incoming one and get featureclass
    Set pFLayer = pLayer
    
    'Get the selection set
    Set pFSel = pFLayer
    Set pSelSet = pFSel.SelectionSet
    
    'Make a call to the function below to return a unioned geometry of any and all selected features
    Set pSelectGeometry = modMainRun.ReturnSelectGeometry(pSelSet)
    Set g_pSelectedPolyClip = pSelectGeometry
    
    'Make a call to get the BasinPoly featureclass using the name sent over
    Set pBasinFeatClass = modUtil.ReturnFeatureClass(strFeatClass)
    
    'Now we select all basinpolys that intersect the unioned geometry we got earlier
    Set pSpatialFilter = New SpatialFilter
    
    With pSpatialFilter
        Set .Geometry = pSelectGeometry
        .GeometryField = pBasinFeatClass.ShapeFieldName
        .SpatialRel = esriSpatialRelContains
    End With
    
    Set pBasinSelSet = pBasinFeatClass.Select(pSpatialFilter, esriSelectionTypeIDSet, _
                                              esriSelectionOptionNormal, pWorkspace)
    
    'Get the FcName from the featureclass
    Set pDataset = pBasinFeatClass
    
    'Set classname and dataset
    Set pINFeatureClassName = pDataset.FullName
    Set pInDsName = pINFeatureClassName

    'Create a new feature class name
    ' Define the output feature class name
    Set pFeatureClassName = New FeatureClassName
    Set pOutDatasetName = pFeatureClassName

    pOutDatasetName.Name = modUtil.GetUniqueFeatureClassName(pWorkspace, "selpoly")

    Set pWorkspaceName = New WorkspaceName
    pWorkspaceName.PathName = pWorkspace.PathName
    pWorkspaceName.WorkspaceFactoryProgID = "esriDataSourcesFile.shapefileworkspacefactory.1"

    Set pOutDatasetName.WorkspaceName = pWorkspaceName
    
    With pFeatureClassName
        .FeatureType = esriFTSimple
        .ShapeType = esriGeometryPolygon
        .ShapeFieldName = "Shape"
    End With
    
    'Export
    Set pExportOp = New ExportOperation
    pExportOp.ExportFeatureClass pInDsName, Nothing, pBasinSelSet, Nothing, pOutDatasetName, 0

    Set pFeatWS = pWorkspace
    Set pFeatClass = pFeatWS.OpenFeatureClass(pOutDatasetName.Name)
    Set ExportSelectedFeatures = pFeatClass
    
    'Cleanup
    Set pFLayer = Nothing
    Set pFC = Nothing
    Set pINFeatureClassName = Nothing
    Set pDataset = Nothing
    Set pInDsName = Nothing
    Set pFSel = Nothing
    Set pSelSet = Nothing
    Set pFeatureClassName = Nothing
    Set pOutDatasetName = Nothing
    Set pWorkspaceName = Nothing
    Set pExportOp = Nothing
    Set pFeatWS = Nothing
    
  Exit Function
ErrorHandler:
  HandleError True, "ExportSelectedFeatures " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Function



