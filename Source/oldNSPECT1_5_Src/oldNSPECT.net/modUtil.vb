Option Strict Off
Option Explicit On
Module modUtil
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
	
	
	' Path to N-SPECT's application and document folder
	Public g_nspectPath As String
	Public g_nspectDocPath As String
	
	'Database Variables
	Public g_ADOConn As ADODB.Connection 'Connection
	Public g_strConn As String 'Connection String
	Public g_ADORS As ADODB.Recordset 'ADO Recordset
	Public g_adoRSCoeff As New ADODB.Recordset
	Public g_boolConnected As Boolean 'Bool: connected
	Public g_intLCClassid As Short 'LCClassid - used for adding 'NEW' records
	
	'Bool: frmAddCoeff is used twice; called from frmNewPollutants(False), frmPollutants(True)
	Public g_boolAddCoeff As Boolean
	Public g_boolCopyCoeff As Boolean 'True: called frmPollutants, False: called frmNewPollutants
	Public g_boolAgree As Boolean 'True: use the Agree Function on Streams.
	Public g_boolHydCorr As Boolean 'True: Hyrdologically Correct DEM, no fill needed
	Public g_boolNewWShed As Boolean 'True: New WaterShed form called from frmPrj
	
	'WqStd
	Public rsWQStdLoad As ADODB.Recordset 'Water Quality RecordSet
	
	'Agree DEM Stuff
	Public g_boolParams As Boolean 'Flag to indicate Agree params have been entered
	
	'Help API
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Declare Function HtmlHelp Lib "HHCtrl.ocx"  Alias "HtmlHelpA"(ByVal hwndCaller As Integer, ByVal pszFile As String, ByVal uCommand As Integer, ByRef dwData As Any) As Integer
	
	Public Const HH_DISPLAY_TOPIC As Integer = &H0
	Public Const HH_HELP_CONTEXT As Integer = &HF
	
	'Project Form Variables
	Public g_strPrjFileName As String 'Project file name
	
	'Management Scenario variables::frmPrjCalc
	Public g_strLUScenFileName As String 'Management scenario file name
	Public g_intManScenRow As String 'Management scenario ROW number
	
	'Pollutant Coefficient variable::frmPrjCalc
	Public g_intCoeffRow As Short 'Coeff Row Number
	Public g_strCoeffCalc As String 'if the Calc option is chosen, hold results in string
	Const c_sModuleFileName As String = "modUtil.bas"
	Private m_ParentHWND As Integer ' Set this to get correct parenting of Error handler forms
	
	Public Sub AddFeatureLayer(ByRef pApp As ESRI.ArcGIS.Framework.IApplication, ByRef pFClass As ESRI.ArcGIS.Geodatabase.IFeatureClass, Optional ByRef sName As String = "")
		On Error GoTo ErrorHandler
		
		Dim pFLayer As ESRI.ArcGIS.Carto.IFeatureLayer
		pFLayer = New ESRI.ArcGIS.Carto.FeatureLayer
		pFLayer.FeatureClass = pFClass
		Dim pFDataset As ESRI.ArcGIS.Geodatabase.IFeatureDataset
		Dim pDS As ESRI.ArcGIS.Geodatabase.IDataset
		If (sName = "") Then
			pFDataset = pFClass.FeatureDataset
			pDS = pFDataset
			'UPGRADE_WARNING: Couldn't resolve default property of object pFLayer.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pFLayer.Name = pDS.BrowseName
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object pFLayer.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pFLayer.Name = sName
		End If
		Dim pMap As ESRI.ArcGIS.Carto.IBasicMap
		Dim pMxDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
		Dim pActView As ESRI.ArcGIS.Carto.IActiveView
		pMxDoc = pApp.Document
		pMap = pMxDoc.FocusMap
		pActView = pMxDoc.ActiveView
		'UPGRADE_WARNING: Couldn't resolve default property of object pFLayer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pMap.AddLayer(pFLayer)
		pActView.Refresh()
		pMxDoc.UpdateContents()
		Exit Sub
		
		'UPGRADE_NOTE: Object pMxDoc may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pMxDoc = Nothing
		'UPGRADE_NOTE: Object pActView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pActView = Nothing
		'UPGRADE_NOTE: Object pMap may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pMap = Nothing
		'UPGRADE_NOTE: Object pFLayer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFLayer = Nothing
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "AddFeatureLayer " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	Public Sub AddFeatureLayerToComboBox(ByRef cboBox As System.Windows.Forms.ComboBox, ByRef pMap As ESRI.ArcGIS.Carto.IMap, ByRef strType As String)
		'Populate combobox of choice with RasterLayers in map
		On Error GoTo ErrHandler
		cboBox.Items.Clear()
		Dim iLyrIndex As Integer
		Dim pFeatClass As ESRI.ArcGIS.Geodatabase.IFeatureClass
		Dim pLyr As ESRI.ArcGIS.Carto.ILayer
		Dim pFeatLayer As ESRI.ArcGIS.Carto.IFeatureLayer
		' Add raster layers into  Combobox
		Dim iLayerCount As Short
		iLayerCount = pMap.LayerCount
		If iLayerCount > 0 Then
			cboBox.Enabled = True
			For iLyrIndex = 0 To iLayerCount - 1
				pLyr = pMap.Layer(iLyrIndex)
				'UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				If (TypeOf pLyr Is ESRI.ArcGIS.Carto.IFeatureLayer) Then
					pFeatLayer = pLyr
					pFeatClass = pFeatLayer.FeatureClass
					
					Select Case strType
						
						Case "line"
							
							If pFeatClass.ShapeType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolyline Then
								cboBox.Items.Add(pLyr.Name)
								VB6.SetItemData(cboBox, cboBox.Items.Count - 1, iLyrIndex)
							End If
							
						Case "poly"
							
							If pFeatClass.ShapeType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolygon Then
								cboBox.Items.Add(pLyr.Name)
								VB6.SetItemData(cboBox, cboBox.Items.Count - 1, iLyrIndex)
							End If
							
						Case "point"
							
							If pFeatClass.ShapeType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPoint Then
								cboBox.Items.Add(pLyr.Name)
								VB6.SetItemData(cboBox, cboBox.Items.Count - 1, iLyrIndex)
							End If
							
					End Select
				End If
			Next iLyrIndex
			If (cboBox.Items.Count > 0) Then
				cboBox.SelectedIndex = 0
				cboBox.Text = pMap.Layer(VB6.GetItemData(cboBox, 0)).Name
			End If
		End If
		
		'UPGRADE_NOTE: Object pFeatClass may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFeatClass = Nothing
		'UPGRADE_NOTE: Object pLyr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pLyr = Nothing
		'UPGRADE_NOTE: Object pFeatLayer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFeatLayer = Nothing
		
		Exit Sub
ErrHandler: 
		MsgBox("Add Layer to ComboBox:" & Err.Description)
		
	End Sub
	'General Function used to simply get the index of combobox entries
	Public Function GetCboIndex(ByRef strList As String, ByRef cbo As System.Windows.Forms.ComboBox) As Short
		On Error GoTo ErrorHandler
		
		
		Dim i As Short
		i = 0
		
		For i = 0 To cbo.Items.Count - 1
			If VB6.GetItemString(cbo, i) = strList Then
				GetCboIndex = i
			End If
		Next i
		
		
		
		Exit Function
ErrorHandler: 
		HandleError(True, "GetCboIndex " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Function
	
	'Function for connection to NSPECT.mdb: fires on dll load
	Public Sub ADOConnection()
		
		On Error GoTo ErrHandler
		If Not g_boolConnected Then
			
			g_ADOConn = New ADODB.Connection
			
			g_strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & modUtil.g_nspectPath & "\nspect.mdb"
			
			g_ADOConn.Open(g_strConn)
			g_ADOConn.CursorLocation = ADODB.CursorLocationEnum.adUseServer
			
			g_boolConnected = True
			
		End If
		Exit Sub
ErrHandler: 
		MsgBox(Err.Number & Err.Description & " Error connecting to database, please check NSPECTDAT enviornment variable.  Current value of NSPECTDAT: " & g_strConn, MsgBoxStyle.Critical, "Error Connecting")
		
	End Sub
	
	'Need to check the text file coming in from the import menu of the pollutant form.
	'Bringing the Text File itself, and the name of the LCType as picked by John User
	Public Function ValidateCoeffTextFile(ByRef strFileName As String, ByRef strLCTypeName As String) As Boolean
		On Error GoTo ErrorHandler
		
		
		Dim fso As New Scripting.FileSystemObject
		Dim fl As Scripting.TextStream
		Dim strLine As String
		Dim intLine As Short
		Dim strValue As String
		Dim strParams(7) As Object
		Dim strLCType As String
		Dim rsLCType As New ADODB.Recordset
		Dim strLCTypeNum As String
		Dim rsCheck As New ADODB.Recordset
		Dim strCheck As String
		Dim i, j As Short
		Dim lstValues() As Short
		
		'Gameplan is to find number of records(landclasses) in the chosen LCType. Then
		'compare that to the number of lines in the text file, and the [Value] field to
		'make sure both jive.  If not, bark at them...ruff, ruff
		
		strLCTypeNum = "SELECT LCTYPE.LCTYPEID, LCCLASS.NAME, LCCLASS.VALUE, LCCLASS.LCCLASSID FROM " & "LCTYPE INNER JOIN LCCLASS ON LCTYPE.LCTYPEID = LCCLASS.LCTYPEID " & "WHERE LCTYPE.NAME LIKE '" & strLCTypeName & "'"
		
		rsLCType.Open(strLCTypeNum, g_ADOConn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		
		If fso.FileExists(strFileName) Then
			
			fl = fso.OpenTextFile(strFileName, Scripting.IOMode.ForReading, True, Scripting.Tristate.TristateFalse)
			
			intLine = 0
			
			'CHECK 1: loop through the text file, compare value to [Value], make sure exists
			Do While Not fl.AtEndOfStream
				
				strLine = fl.ReadLine
				'Value exits??
				strValue = Split(strLine, ",")(0)
				
				j = 0
				rsLCType.MoveFirst()
				
				For i = 0 To rsLCType.RecordCount - 1
					
					If rsLCType.Fields("Value").Value = strValue Then
						j = j + 1
					End If
					rsLCType.MoveNext()
					
				Next i
				
				If j = 0 Then
					MsgBox("There is a value in your text file that does not exist in the Land Class Type: '" & strLCTypeName & "' Please check your text file in line: " & intLine + 1, MsgBoxStyle.OKOnly, "Data Import Error")
					
					ValidateCoeffTextFile = False
					GoTo Cleanup
				ElseIf j > 1 Then 
					MsgBox("There are records in your text file that contain the same value.  Please check line " & intLine, MsgBoxStyle.Critical, "Multiple values found")
					ValidateCoeffTextFile = False
					GoTo Cleanup
				ElseIf j = 1 Then 
					ValidateCoeffTextFile = True
				End If
				
				intLine = intLine + 1
				Debug.Print(intLine)
				
			Loop 
			
			'Final check, make sure same number of records in text file vs the
			If rsLCType.RecordCount = intLine Then
				ValidateCoeffTextFile = True
			Else
				MsgBox("The number of records in your import file do not match the number of records in the " & "Landclass '" & strLCTypeName & "'.  Your file should contain " & rsLCType.RecordCount & " records.", MsgBoxStyle.Critical, "Error Importing File")
				ValidateCoeffTextFile = False
				GoTo Cleanup
			End If
			
			
		Else
			MsgBox("The file you are pointing to does not exist. Please select another.", MsgBoxStyle.Critical, "File Not Found")
			'Cleanup
		End If
		
		If ValidateCoeffTextFile Then
			g_adoRSCoeff = rsLCType
		End If
		
		Exit Function
		
Cleanup: 
		
		'UPGRADE_NOTE: Object fso may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		fso = Nothing
		'UPGRADE_NOTE: Object fl may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		fl = Nothing
		rsLCType.Close()
		'UPGRADE_NOTE: Object rsLCType may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsLCType = Nothing
		
		
		
		Exit Function
ErrorHandler: 
		HandleError(True, "ValidateCoeffTextFile " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Function
	
	'Check for spatial reference of a Raster Dataset, returns the Coordinate System
	Public Function CheckSpatialReference(ByRef pRasGeoDataset As ESRI.ArcGIS.Geodatabase.IRasterDataset) As ESRI.ArcGIS.Geometry.IProjectedCoordinateSystem
		On Error GoTo ErrorHandler
		
		
		'Code to call ArcMap Functionality....
		Dim pGeoDataset As ESRI.ArcGIS.Geodatabase.IGeoDataset
		Dim pDEMSpatRef As ESRI.ArcGIS.Geometry.ISpatialReference
		Dim pRasterProps As ESRI.ArcGIS.DataSourcesRaster.IRasterProps
		Dim pDistUnit As ESRI.ArcGIS.Geometry.ILinearUnit
		Dim pPrjCoordSys As ESRI.ArcGIS.Geometry.IProjectedCoordinateSystem
		
		pGeoDataset = pRasGeoDataset.CreateDefaultRaster 'Create a default to get props
		pRasterProps = pGeoDataset
		
		'Get the Spat Reference, check for: if there, send it back, else Nothing
		If Not (pRasterProps.SpatialReference Is Nothing) Then
			CheckSpatialReference = pRasterProps.SpatialReference
		Else
			'UPGRADE_NOTE: Object CheckSpatialReference may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			CheckSpatialReference = Nothing
		End If
		
		'Cleanup
		'UPGRADE_NOTE: Object pGeoDataset may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pGeoDataset = Nothing
		'UPGRADE_NOTE: Object pDEMSpatRef may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pDEMSpatRef = Nothing
		'UPGRADE_NOTE: Object pRasterProps may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRasterProps = Nothing
		'UPGRADE_NOTE: Object pDistUnit may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pDistUnit = Nothing
		'UPGRADE_NOTE: Object pPrjCoordSys may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pPrjCoordSys = Nothing
		
		
		
		Exit Function
ErrorHandler: 
		HandleError(True, "CheckSpatialReference " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Function
	
	Public Function CheckLandCoverFields(ByRef strLandClass As String, ByRef pLCRaster As ESRI.ArcGIS.Geodatabase.IRaster) As Boolean
		
		Dim strRS As String
		Dim rsLandClass As ADODB.Recordset
		Dim strPick(3) As String 'Array of strings that hold 'pick' numbers
		Dim strCurveCalc As String 'Full String
		
		Dim pRasterCol As ESRI.ArcGIS.DataSourcesRaster.IRasterBandCollection
		Dim pBand As ESRI.ArcGIS.DataSourcesRaster.IRasterBand
		Dim pRasterStats As ESRI.ArcGIS.DataSourcesRaster.IRasterStatistics
		Dim dblMaxValue As Double
		Dim i As Short
		Dim pTable As ESRI.ArcGIS.Geodatabase.ITable
		Dim TableExist As Boolean
		
		On Error GoTo ErrHandler
		
		'STEP 1:  get the records from the database -----------------------------------------------
		strRS = "SELECT * FROM LCCLASS " & "LEFT OUTER JOIN LCTYPE ON LCCLASS.LCTYPEID = LCTYPE.LCTYPEID " & "Where LCTYPE.NAME = '" & strLandClass & "' ORDER BY LCCLASS.VALUE"
		
		rsLandClass = New ADODB.Recordset
		rsLandClass.Open(strRS, g_ADOConn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		
		rsLandClass.MoveFirst()
		'End Database stuff
		
		'STEP 2: Raster Values ---------------------------------------------------------------------
		'Now Get the RASTER values
		' Get Rasterband from the incoming raster
		pRasterCol = pLCRaster
		pBand = pRasterCol.Item(0)
		
		'Get the max value
		pRasterStats = pBand.Statistics
		dblMaxValue = pRasterStats.Maximum
		
		'Get the raster table
		pBand.HasTable(TableExist)
		If Not TableExist Then
			CheckLandCoverFields = False
			Exit Function
		End If
		
		pTable = pBand.AttributeTable
		
		'STEP 3: Table values vs. Raster Values Count - if not equal bark --------------------------
		If pTable.RowCount(Nothing) <> rsLandClass.RecordCount Then
			CheckLandCoverFields = False
		Else
			CheckLandCoverFields = True
		End If
		
		'Cleanup
		rsLandClass.Close()
		'UPGRADE_NOTE: Object rsLandClass may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsLandClass = Nothing
		Exit Function
		
ErrHandler: 
		MsgBox("Error Occurred on CheckFields")
		CheckLandCoverFields = False
		
	End Function
	
	
	
	'add grid column headers for Coefficient tab
	Public Sub InitPollDefGrid(ByRef grdPollDef As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid)
		On Error GoTo ErrorHandler
		
		
		
		With grdPollDef
			.col = .FixedCols
			.row = .FixedRows
			.set_ColWidth(0,  , 400)
			.set_ColWidth(1,  , 850)
			.set_ColWidth(2,  , 2600)
			.set_ColWidth(3,  , 840)
			.set_ColWidth(4,  , 840)
			.set_ColWidth(5,  , 840)
			.set_ColWidth(6,  , 840)
			.set_ColWidth(7,  , 0) 'Coefficient Set ID
			.set_ColWidth(8,  , 0) 'LCClass ID - both hidden but necessary for other deeds.
			.row = 0
			.col = 1
			.Text = "Value"
			.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
			.col = 2
			.Text = "Name"
			.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
			.col = 3
			.Text = "Type 1"
			.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
			.col = 4
			.Text = "Type 2"
			.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
			.col = 5
			.Text = "Type 3"
			.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
			.col = 6
			.Text = "Type 4"
			.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
		End With
		
		
		
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "InitPollDefGrid " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	'add grid column headers for WQ Standards tab
	Public Sub InitPollWQStdGrid(ByRef grdWQStd As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid)
		On Error GoTo ErrorHandler
		
		
		With grdWQStd
			.col = .FixedCols
			.row = .FixedRows
			.BackColorSel = System.Drawing.ColorTranslator.FromOle(&HC0FFFF) 'lt. yellow
			.Width = VB6.TwipsToPixelsX(7100)
			.set_ColWidth(0,  , 400)
			.set_ColWidth(1,  , 2000)
			.set_ColWidth(2,  , 3150)
			.set_ColWidth(3,  , 1450)
			.row = 0
			.col = 1
			.Text = "Standards Name"
			.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
			.col = 2
			.Text = "Description"
			.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
			.col = 3
			.Text = "Threshold (ug/L)"
			.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter
			.col = 4
			.set_ColWidth(4,  , 0)
		End With
		
		
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "InitPollWQStdGrid " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	'add grid column headers
	Public Sub InitWQStdGrid(ByRef grdWQStd As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid)
		On Error GoTo ErrorHandler
		
		
		With grdWQStd
			.col = .FixedCols
			.row = .FixedRows
			.BackColorSel = System.Drawing.ColorTranslator.FromOle(&HC0FFFF) 'lt. yellow
			.Width = VB6.TwipsToPixelsX(5460)
			.set_ColWidth(0,  , 300)
			.set_ColWidth(1,  , 2900)
			.set_ColWidth(2,  , 2140)
			.set_ColWidth(3,  , 0)
			.row = 0
			.col = 1
			.Text = "Pollutant"
			.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
			.col = 2
			.Text = "Threshold (ug/L)"
			.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
			
		End With
		
		
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "InitWQStdGrid " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	Public Sub InitLCClassesGrid(ByRef grdLCClasses As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid)
		On Error GoTo ErrorHandler
		
		'add grid column headers
		With grdLCClasses
			.col = .FixedCols
			.row = .FixedRows
			.BackColorSel = System.Drawing.ColorTranslator.FromOle(&HC0FFFF) 'lt. yellow
			.set_ColWidth(0,  , 400)
			.set_ColWidth(1,  , 800)
			.set_ColWidth(2,  , 2200)
			.set_ColWidth(3,  , 800)
			.set_ColWidth(4,  , 800)
			.set_ColWidth(5,  , 800)
			.set_ColWidth(6,  , 800)
			.set_ColWidth(7,  , 1350)
			.set_ColWidth(8,  , 800)
			.Width = VB6.TwipsToPixelsX(9050)
			.row = 0
			.col = 1
			.Text = "Value"
			.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter
			.col = 2
			.Text = "Name"
			.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter
			.col = 3
			.Text = "CN-A"
			.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter
			.col = 4
			.Text = "CN-B"
			.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter
			.col = 5
			.Text = "CN-C"
			.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter
			.col = 6
			.Text = "CN-D"
			.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter
			.col = 7
			.Text = "Cover-Factor"
			.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter
			.col = 8
			.Text = "Wet"
			.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
			.col = 9
			.set_ColWidth(9,  , 0)
			.col = 10
			.set_ColWidth(10,  , 0)
		End With
		
		
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "InitLCClassesGrid " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	Public Function IsLoaded(ByRef FormName As String) As Boolean
		On Error GoTo ErrorHandler
		
		'Used to see if particular forms are/are not loaded at various time
		Dim sFormName As String
		Dim f As System.Windows.Forms.Form
		sFormName = UCase(FormName)
		
		For	Each f In My.Application.OpenForms
			If UCase(f.Name) = sFormName Then
				IsLoaded = True
				Exit Function
			End If
		Next f
		
		
		
		Exit Function
ErrorHandler: 
		HandleError(True, "IsLoaded " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Function
	
	Public Sub InitComboBox(ByRef cbo As System.Windows.Forms.ComboBox, ByRef strName As String)
		On Error GoTo ErrorHandler
		
		'Loads the variety of comboboxes throught the project using combobox and name of table
		Dim rsNames As ADODB.Recordset
		Dim strSelectStatement As String
		rsNames = New ADODB.Recordset
		
		strSelectStatement = "SELECT NAME FROM " & strName & " ORDER BY NAME ASC"
		
		'Check thrown in to make sure g_ADOconn is something, in v9.1 we started having problems.
		If Not g_boolConnected Then
			ADOConnection()
		End If
		
		rsNames.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		rsNames.Open(strSelectStatement, g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
		
		If rsNames.RecordCount > 0 Then
			With cbo
				Do Until rsNames.EOF
					.Items.Add(rsNames.Fields("Name").Value)
					rsNames.MoveNext()
				Loop 
			End With
			
			cbo.SelectedIndex = 0
		Else
			MsgBox("Warning.  There are no records remaining.  Please add a new one.", MsgBoxStyle.Critical, "Recordset Empty")
			Exit Sub
		End If
		
		'Cleanup
		rsNames.Close()
		'UPGRADE_NOTE: Object rsNames may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsNames = Nothing
		
		
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "InitComboBox " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	'Tests name inputs to insure unique values
	Public Function UniqueName(ByRef strTableName As String, ByRef strName As String) As Boolean
		On Error GoTo ErrorHandler
		
		
		Dim strCmdText As String
		Dim rsName As ADODB.Recordset
		
		strCmdText = "SELECT * FROM " & strTableName & " WHERE NAME LIKE '" & strName & "'"
		rsName = New ADODB.Recordset
		
		rsName.Open(strCmdText, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
		
		If rsName.RecordCount > 0 Then
			UniqueName = False
		Else
			UniqueName = True
		End If
		
		'Cleanjeans
		rsName.Close()
		'UPGRADE_NOTE: Object rsName may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsName = Nothing
		
		
		
		Exit Function
ErrorHandler: 
		HandleError(True, "UniqueName " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Function
	
	'Tests name inputs to insure unique values for databases
	Public Function CreateUniqueName(ByRef strTableName As String, ByRef strName As String) As String
		On Error GoTo ErrorHandler
		
		
		Dim strCmdText As String
		Dim rsName As ADODB.Recordset
		Dim i As Short
		Dim sCurrNum As String
		Dim strCurrNameRecord As String
		strCmdText = "SELECT * FROM " & strTableName '& " WHERE NAME LIKE '" & strName & "'"
		rsName = New ADODB.Recordset
		
		rsName.Open(strCmdText, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
		sCurrNum = "0"
		
		rsName.MoveFirst()
		
		For i = 1 To rsName.RecordCount
			strCurrNameRecord = CStr(rsName.Fields("Name").Value)
			If InStr(1, strCurrNameRecord, strName, 1) > 0 Then
				If IsNumeric(Right(strCurrNameRecord, 2)) Then
					If (CShort(Right(strCurrNameRecord, 2)) > CShort(sCurrNum)) Then
						sCurrNum = Right(strCurrNameRecord, 2)
					Else
						Exit For
					End If
				Else
					If IsNumeric(Right(strCurrNameRecord, 1)) Then
						If (CShort(Right(strCurrNameRecord, 1)) > CShort(sCurrNum)) Then
							sCurrNum = Right(strCurrNameRecord, 1)
						End If
					End If
				End If
			End If
			rsName.MoveNext()
		Next i
		If sCurrNum = "0" Then
			CreateUniqueName = strName & "1"
		Else
			CreateUniqueName = strName & CStr(CShort(sCurrNum) + 1)
		End If
		
		'Cleanjeans
		rsName.Close()
		'UPGRADE_NOTE: Object rsName may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsName = Nothing
		
		
		
		Exit Function
ErrorHandler: 
		HandleError(True, "CreateUniqueName " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Function
	
	
	Private Sub SetGridColumnWidth(ByRef grd As AxMSFlexGridLib.AxMSFlexGrid)
		On Error GoTo ErrorHandler
		
		'params:    ms flexgrid control
		'purpose:   sets the column widths to the
		'           lengths of the longest string in the column
		'requirements:  the grid must have the same
		'               font as the underlying form
		
		Dim InnerLoopCount As Integer
		Dim OuterLoopCount As Integer
		Dim lngLongestLen As Integer
		Dim sLongestString As String
		Dim lngColWidth As Integer
		Dim szCellText As String
		
		For OuterLoopCount = 0 To grd.Cols - 1
			sLongestString = ""
			lngLongestLen = 0
			
			'grd.Col = OuterLoopCount
			For InnerLoopCount = 0 To grd.Rows - 1
				szCellText = grd.get_TextMatrix(InnerLoopCount, OuterLoopCount)
				'grd.Row = InnerLoopCount
				'szCellText = Trim$(grd.Text)
				If Len(szCellText) > lngLongestLen Then
					lngLongestLen = Len(szCellText)
					sLongestString = szCellText
				End If
			Next 
			'UPGRADE_ISSUE: Form method Parent.TextWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			lngColWidth = grd.Parent.TextWidth(sLongestString)
			
			'add 100 for more readable spreadsheet
			grd.set_ColWidth(OuterLoopCount, lngColWidth + 200)
		Next 
		
		
		Exit Sub
ErrorHandler: 
		HandleError(False, "SetGridColumnWidth " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	Public Sub HighlightGridRow(ByRef grd As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid, ByRef iRow As Short)
		On Error GoTo ErrorHandler
		
		With grd
			If .Rows > 1 Then
				.row = iRow
				.col = 1
				.ColSel = .get_Cols() - 1
				.RowSel = iRow
			End If
		End With
		
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "HighlightGridRow " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	
	'*********************************************************************************************
	'** Following series of functions courtesy of ESRI, provide a variety of useful
	'** actions
	'*********************************************************************************************
	Public Function OpenRasterDataset(ByRef sPath As String, ByRef sRasterName As String) As ESRI.ArcGIS.Geodatabase.IRasterDataset
		'Return RasterDataset Object given a file name and its directory
		On Error GoTo ERH
		Dim pWSFact As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory
		Dim pRasterWS As ESRI.ArcGIS.DataSourcesRaster.IRasterWorkspace
		
		pWSFact = New ESRI.ArcGIS.DataSourcesRaster.RasterWorkspaceFactory
		If pWSFact.IsWorkspace(sPath) Then
			pRasterWS = pWSFact.OpenFromFile(sPath, 0)
			OpenRasterDataset = pRasterWS.OpenRasterDataset(sRasterName)
			
		End If
		Exit Function
ERH: 
		MsgBox("Failed in opening raster dataset. " & Err.Description)
	End Function
	Public Function IsValidWorkspace(ByRef sPath As String) As Boolean
		' Given a pathname, determines if workspace is valid
		On Error GoTo ErrorHandler
		Dim pWSF As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory
		pWSF = New ESRI.ArcGIS.DataSourcesRaster.RasterWorkspaceFactory
		If pWSF.IsWorkspace(sPath) Then
			IsValidWorkspace = True
		Else
			IsValidWorkspace = False
		End If
		
		'UPGRADE_NOTE: Object pWSF may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pWSF = Nothing
		Exit Function
		
ErrorHandler: 
		IsValidWorkspace = False
		'UPGRADE_NOTE: Object pWSF may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pWSF = Nothing
		
	End Function
	
	Public Function SetRasterWorkspace(ByRef sPath As String) As ESRI.ArcGIS.DataSourcesRaster.IRasterWorkspace
		
		' Given a pathname, returns the raster workspace object for that path
		On Error GoTo ErrorSetWorkspace
		Dim pWSF As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory
		pWSF = New ESRI.ArcGIS.DataSourcesRaster.RasterWorkspaceFactory
		Dim pPropSet As ESRI.ArcGIS.esriSystem.IPropertySet
		If pWSF.IsWorkspace(sPath) Then
			SetRasterWorkspace = pWSF.OpenFromFile(sPath, 0)
			'UPGRADE_NOTE: Object pWSF may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			pWSF = Nothing
		Else
			pPropSet = New ESRI.ArcGIS.esriSystem.PropertySet
			pPropSet.setProperty("DATABASE", sPath)
			SetRasterWorkspace = pWSF.Open(pPropSet, 0)
			
			'UPGRADE_NOTE: Object pPropSet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			pPropSet = Nothing
			
		End If
		Exit Function
ErrorSetWorkspace: 
		'UPGRADE_NOTE: Object SetRasterWorkspace may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		SetRasterWorkspace = Nothing
		
	End Function
	
	Public Function SetFeatureShapeWorkspace(ByRef sPath As String) As ESRI.ArcGIS.Geodatabase.IWorkspace
		
		' Given a pathname, returns the shapefile workspace object for that path
		On Error GoTo ErrorSetWorkspace
		Dim pWSF As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory
		pWSF = New ESRI.ArcGIS.DataSourcesFile.ShapefileWorkspaceFactory
		Dim pPropSet As ESRI.ArcGIS.esriSystem.IPropertySet
		If pWSF.IsWorkspace(sPath) Then
			SetFeatureShapeWorkspace = pWSF.OpenFromFile(sPath, 0)
			'UPGRADE_NOTE: Object pWSF may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			pWSF = Nothing
		Else
			pPropSet = New ESRI.ArcGIS.esriSystem.PropertySet
			pPropSet.setProperty("DATABASE", sPath)
			SetFeatureShapeWorkspace = pWSF.Open(pPropSet, 0)
			
			'UPGRADE_NOTE: Object pPropSet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			pPropSet = Nothing
		End If
		Exit Function
		
ErrorSetWorkspace: 
		'UPGRADE_NOTE: Object SetFeatureShapeWorkspace may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		SetFeatureShapeWorkspace = Nothing
		
	End Function
	
	'Function to get unique filename in directory of choice
	Public Function GetUniqueName(ByRef Name As String, ByRef folderPath As String) As String
		
		On Error GoTo EH
		' Find file, if exists, increment number
		Dim fso As Object 'As FileSystemObject
		Dim pFolder As Object 'As Folder
		Dim pFile As Object 'As File
		Dim sCurrNum As String
		
		fso = CreateObject("Scripting.FileSystemObject") ' New FileSystemObject
		'UPGRADE_WARNING: Couldn't resolve default property of object fso.GetFolder. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pFolder = fso.GetFolder(folderPath)
		sCurrNum = "0"
		Dim s() As String
		'UPGRADE_WARNING: Couldn't resolve default property of object pFolder.SubFolders. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each pFile In pFolder.SubFolders
			'UPGRADE_WARNING: Couldn't resolve default property of object pFile.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If InStr(1, pFile.Name, Name, 1) > 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object pFile.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				s = Split(pFile.Name, ".")
				If IsNumeric(Right(s(0), 2)) Then
					If (CShort(Right(s(0), 2)) > CShort(sCurrNum)) Then
						sCurrNum = Right(s(0), 2)
					Else
						Exit For
					End If
				Else
					If IsNumeric(Right(s(0), 1)) Then
						If (CShort(Right(s(0), 1)) > CShort(sCurrNum)) Then
							sCurrNum = Right(s(0), 1)
						End If
					End If
				End If
			End If
		Next pFile
		If sCurrNum = "0" Then
			GetUniqueName = Name & "1"
		Else
			GetUniqueName = Name & CStr(CShort(sCurrNum) + 1)
		End If
		
		'UPGRADE_NOTE: Object pFile may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFile = Nothing
		'UPGRADE_NOTE: Object pFolder may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFolder = Nothing
		'UPGRADE_NOTE: Object fso may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		fso = Nothing
		Exit Function
EH: 
		MsgBox(Err.Number & vbLf & Err.Description,  , "Error in GetUniqueName ")
		
	End Function
	
	Public Sub AddRasterLayer(ByRef pApp As ESRI.ArcGIS.Framework.IApplication, ByRef pRaster As ESRI.ArcGIS.Geodatabase.IRaster, Optional ByRef sName As String = "")
		On Error GoTo EH
		
		Dim pRasterlayer As ESRI.ArcGIS.Carto.IRasterLayer
		pRasterlayer = New ESRI.ArcGIS.Carto.RasterLayer
		pRasterlayer.CreateFromRaster(pRaster)
		
		Dim pRasterBands As ESRI.ArcGIS.DataSourcesRaster.IRasterBandCollection
		Dim pRasterBand As ESRI.ArcGIS.DataSourcesRaster.IRasterBand
		Dim pDS As ESRI.ArcGIS.Geodatabase.IDataset
		If (sName = "") Then
			pRasterBands = pRaster
			pRasterBand = pRasterBands.Item(0)
			pDS = pRasterBand.RasterDataset
			'UPGRADE_WARNING: Couldn't resolve default property of object pRasterlayer.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pRasterlayer.Name = pDS.BrowseName
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object pRasterlayer.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pRasterlayer.Name = sName
		End If
		
		Dim pMap As ESRI.ArcGIS.Carto.IBasicMap
		Dim pMxDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
		Dim pActView As ESRI.ArcGIS.Carto.IActiveView
		pMxDoc = pApp.Document
		pMap = pMxDoc.FocusMap
		pActView = pMxDoc.ActiveView
		'UPGRADE_WARNING: Couldn't resolve default property of object pRasterlayer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pMap.AddLayer(pRasterlayer)
		pActView.Refresh()
		pMxDoc.UpdateContents()
		Exit Sub
		
		'UPGRADE_NOTE: Object pMxDoc may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pMxDoc = Nothing
		'UPGRADE_NOTE: Object pActView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pActView = Nothing
		'UPGRADE_NOTE: Object pMap may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pMap = Nothing
		'UPGRADE_NOTE: Object pRasterlayer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRasterlayer = Nothing
EH: 
		MsgBox("Util.AddRasterLayer: " & Err.Description)
	End Sub
	
	Public Function ReturnRasterLayer(ByRef pApp As ESRI.ArcGIS.Framework.IApplication, ByRef pRaster As ESRI.ArcGIS.Geodatabase.IRaster, Optional ByRef sName As String = "") As ESRI.ArcGIS.Carto.IRasterLayer
		
		Dim pRasterlayer As ESRI.ArcGIS.Carto.IRasterLayer
		
		On Error GoTo EH
		
		pRasterlayer = New ESRI.ArcGIS.Carto.RasterLayer
		pRasterlayer.CreateFromRaster(pRaster)
		
		Dim pRasterBands As ESRI.ArcGIS.DataSourcesRaster.IRasterBandCollection
		Dim pRasterBand As ESRI.ArcGIS.DataSourcesRaster.IRasterBand
		Dim pDS As ESRI.ArcGIS.Geodatabase.IDataset
		If (sName = "") Then
			pRasterBands = pRaster
			pRasterBand = pRasterBands.Item(0)
			pDS = pRasterBand.RasterDataset
			'UPGRADE_WARNING: Couldn't resolve default property of object pRasterlayer.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pRasterlayer.Name = pDS.BrowseName
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object pRasterlayer.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pRasterlayer.Name = sName
		End If
		
		ReturnRasterLayer = pRasterlayer
		
		'UPGRADE_NOTE: Object pRasterBands may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRasterBands = Nothing
		'UPGRADE_NOTE: Object pRasterBand may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRasterBand = Nothing
		'UPGRADE_NOTE: Object pDS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pDS = Nothing
		
		Exit Function
EH: 
		MsgBox("Util.AddRasterLayer: " & Err.Description)
	End Function
	
	Public Function GetFeatureClass(ByRef strFeatClass As String, ByRef pApp As ESRI.ArcGIS.Framework.IApplication) As ESRI.ArcGIS.Geodatabase.IFeatureClass
		On Error GoTo ErrorHandler
		
		
		Dim pMap As ESRI.ArcGIS.Carto.IMap
		
		'UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Dim pMxDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
		Dim pLayer As ESRI.ArcGIS.Carto.IFeatureLayer
		If (TypeOf pApp Is ESRI.ArcGIS.ArcMapUI.IMxApplication) Then
			
			pMxDoc = pApp.Document
			pMap = pMxDoc.ActiveView.FocusMap
			
			pLayer = pMap.Layer(GetLayerIndex(strFeatClass, pApp))
			GetFeatureClass = pLayer.FeatureClass
			
		End If
		
		
		
		Exit Function
ErrorHandler: 
		HandleError(True, "GetFeatureClass " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Function
	
	Public Function GetLayerIndex(ByRef strLayerName As String, ByRef pApp As ESRI.ArcGIS.Framework.IApplication) As Integer
		On Error GoTo ErrorHandler
		
		
		Dim pMxDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
		Dim pMap As ESRI.ArcGIS.Carto.IMap
		Dim i As Short
		
		pMxDoc = pApp.Document
		pMap = pMxDoc.FocusMap
		
		i = 0
		
		For i = 0 To pMap.LayerCount - 1
			If pMap.Layer(i).Name = strLayerName Then
				GetLayerIndex = i
			End If
		Next i
		
		'Cleanup
		'UPGRADE_NOTE: Object pMxDoc may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pMxDoc = Nothing
		'UPGRADE_NOTE: Object pMap may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pMap = Nothing
		
		Exit Function
ErrorHandler: 
		HandleError(True, "GetLayerIndex " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Function
	
	Public Sub AddLayer(ByRef pApp As ESRI.ArcGIS.Framework.IApplication, ByRef pLayer As ESRI.ArcGIS.Carto.ILayer)
		
		On Error GoTo EH
		
		Dim pMap As ESRI.ArcGIS.Carto.IBasicMap
		
		'UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Dim pMxDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
		If (TypeOf pApp Is ESRI.ArcGIS.ArcMapUI.IMxApplication) Then
			pMxDoc = pApp.Document
			pMap = pMxDoc.ActiveView.FocusMap
		End If
		pMap.AddLayer(pLayer)
		Exit Sub
EH: 
		MsgBox("Util.AddLayer: " & Err.Description)
	End Sub
	
	Public Function BrowseForFileName(ByRef strType As String, ByRef frm As System.Windows.Forms.Form, ByRef strTitle As String) As String
		
		On Error GoTo ERH
		Dim pGxObject As ESRI.ArcGIS.Catalog.IGxObject
		'  Dim pFilter As GxFilterRasterDatasets
		Dim pFilter As ESRI.ArcGIS.Catalog.IGxObjectFilter
		Dim pMiniBrowser As ESRI.ArcGIS.CatalogUI.IGxDialog
		Dim pEnumGxObject As ESRI.ArcGIS.Catalog.IEnumGxObject
		Dim pDatasetName As ESRI.ArcGIS.Geodatabase.IDatasetName
		Dim strSlash As String
		pMiniBrowser = New ESRI.ArcGIS.CatalogUI.GxDialog
		
		Select Case strType
			Case "Feature"
				pFilter = New ESRI.ArcGIS.Catalog.GxFilterShapefiles
				strSlash = "\"
			Case "Raster"
				pFilter = New ESRI.ArcGIS.Catalog.GxFilterRasterDatasets
				strSlash = ""
		End Select
		
		pMiniBrowser.ObjectFilter = pFilter
		pMiniBrowser.Title = strTitle
		
		Dim pGxDataset As ESRI.ArcGIS.Catalog.IGxDataset
		Dim pDataset As ESRI.ArcGIS.Geodatabase.IDataset
		If (pMiniBrowser.DoModalOpen(frm.Handle.ToInt32, pEnumGxObject)) Then
			pGxObject = pEnumGxObject.Next
			pGxDataset = pGxObject
			pDataset = pGxDataset.Dataset
			pDatasetName = pDataset.FullName
		End If
		
		If strType = "Feature" Then
			BrowseForFileName = pDataset.Workspace.PathName & strSlash & pDataset.Name & ".shp"
		Else
			BrowseForFileName = pDataset.Workspace.PathName & strSlash & pDataset.Name
		End If
		
		
		'UPGRADE_NOTE: Object pGxObject may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pGxObject = Nothing
		'UPGRADE_NOTE: Object pMiniBrowser may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pMiniBrowser = Nothing
		Exit Function
ERH: 
		'UPGRADE_WARNING: Couldn't resolve default property of object Err.Descript. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		MsgBox("AddInputfromBrowser:" & Err.Descript)
	End Function
	
	
	Public Function AddInputFromGxBrowser(ByRef cboInput As System.Windows.Forms.ComboBox, ByRef frm As System.Windows.Forms.Form, ByRef strType As String) As String
		On Error GoTo ERH
		Dim pGxObject As ESRI.ArcGIS.Catalog.IGxObject
		'  Dim pFilter As GxFilterRasterDatasets
		Dim pFilter As ESRI.ArcGIS.Catalog.IGxObjectFilter
		Dim pMiniBrowser As ESRI.ArcGIS.CatalogUI.IGxDialog
		Dim pEnumGxObject As ESRI.ArcGIS.Catalog.IEnumGxObject
		
		pMiniBrowser = New ESRI.ArcGIS.CatalogUI.GxDialog
		
		Select Case strType
			Case "Feature"
				pFilter = New ESRI.ArcGIS.Catalog.GxFilterFeatureClasses
			Case "Raster"
				pFilter = New ESRI.ArcGIS.Catalog.GxFilterRasterDatasets
		End Select
		
		pMiniBrowser.ObjectFilter = pFilter
		pMiniBrowser.Title = "Select Dataset"
		
		Dim pGxDataset As ESRI.ArcGIS.Catalog.IGxDataset
		Dim pDataset As ESRI.ArcGIS.Geodatabase.IDataset
		If (pMiniBrowser.DoModalOpen(frm.Handle.ToInt32, pEnumGxObject)) Then
			pGxObject = pEnumGxObject.Next
			pGxDataset = pGxObject
			pDataset = pGxDataset.Dataset
			cboInput.Items.Add(pDataset.Name)
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
		
		'UPGRADE_NOTE: Object pGxObject may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pGxObject = Nothing
		'UPGRADE_NOTE: Object pMiniBrowser may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pMiniBrowser = Nothing
		Exit Function
ERH: 
		MsgBox("AddInputfromBrowser:" & Err.Description)
	End Function
	
	' Used in "Browse..." buttons on various forms to pick out rasters.
	' Have to love this function, by the way.  Not only does it return a rasterdataset but it also populates
	' a text box that you send it w/ said rasterdataset path.  More like a Sunction this is.
	Public Function AddInputFromGxBrowserText(ByRef txtInput As System.Windows.Forms.TextBox, ByRef strTitle As String, ByRef frm As System.Windows.Forms.Form, ByRef intType As Short) As ESRI.ArcGIS.Geodatabase.IRasterDataset
		
		On Error GoTo ErrHandler
		
		Dim pGxObject As ESRI.ArcGIS.Catalog.IGxObject
		Dim pFilter As ESRI.ArcGIS.Catalog.IGxObjectFilter
		Dim pMiniBrowser As ESRI.ArcGIS.CatalogUI.IGxDialog
		Dim pEnumGxObject As ESRI.ArcGIS.Catalog.IEnumGxObject
		
		pMiniBrowser = New ESRI.ArcGIS.CatalogUI.GxDialog
		
		Select Case intType 'Set up for future use case we add other filters.
			Case 0
				pFilter = New ESRI.ArcGIS.Catalog.GxFilterRasterDatasets 'Filter for Rasters
				
		End Select
		
		pMiniBrowser.ObjectFilter = pFilter
		pMiniBrowser.Title = strTitle
		
		Dim pGxDataset As ESRI.ArcGIS.Catalog.IGxDataset
		Dim pDataset As ESRI.ArcGIS.Geodatabase.IDataset
		If (pMiniBrowser.DoModalOpen(frm.Handle.ToInt32, pEnumGxObject)) Then
			pGxObject = pEnumGxObject.Next
			pGxDataset = pGxObject
			
			pDataset = pGxDataset.Dataset
			txtInput.Text = pDataset.Workspace.PathName & pDataset.Name
			
		Else
			
			'UPGRADE_NOTE: Object AddInputFromGxBrowserText may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			AddInputFromGxBrowserText = Nothing
			Exit Function
			
		End If
		
		AddInputFromGxBrowserText = pDataset
		'UPGRADE_NOTE: Object pGxObject may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pGxObject = Nothing
		'UPGRADE_NOTE: Object pMiniBrowser may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pMiniBrowser = Nothing
		
		Exit Function
		
ErrHandler: 
		MsgBox("The file you have choosen is not a valid GRID dataset.  Please select another.", MsgBoxStyle.Critical, "Invalid Data Type")
		
	End Function
	
	
	' Brings up GxDialog to get output name for specified type. Returns true if user defines output name or
	' false if operation cancelled.
	Public Function BrowseForOutputName(ByRef sName As String, ByRef sLocation As String, ByRef eType As ESRI.ArcGIS.Geodatabase.esriDatasetType, ByRef hParentHwnd As Integer) As Boolean
		On Error GoTo EH
		Dim pGxDialog As ESRI.ArcGIS.CatalogUI.IGxDialog
		pGxDialog = New ESRI.ArcGIS.CatalogUI.GxDialog
		Dim pFilterCol As ESRI.ArcGIS.Catalog.IGxObjectFilterCollection
		pFilterCol = pGxDialog
		
		Dim pFilter As ESRI.ArcGIS.Catalog.IGxObjectFilter
		Dim pFormatFilter As New clsRasterFilter
		Select Case eType
			Case ESRI.ArcGIS.Geodatabase.esriDatasetType.esriDTFeatureClass
				pGxDialog.Title = "Save Features As"
				pFilter = New ESRI.ArcGIS.Catalog.GxFilterShapefiles
				pFilterCol.AddFilter(pFilter, True)
			Case ESRI.ArcGIS.Geodatabase.esriDatasetType.esriDTRasterDataset
				pGxDialog.Title = "Save Raster As"
				pFormatFilter.format_Renamed = "TIFF"
				'UPGRADE_WARNING: Couldn't resolve default property of object pFormatFilter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pFilterCol.AddFilter(pFormatFilter, True)
				pFormatFilter = New clsRasterFilter
				pFormatFilter.format_Renamed = "IMAGINE Image"
				'UPGRADE_WARNING: Couldn't resolve default property of object pFormatFilter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pFilterCol.AddFilter(pFormatFilter, True)
				pFormatFilter = New clsRasterFilter
				pFormatFilter.format_Renamed = "GRID"
				'UPGRADE_WARNING: Couldn't resolve default property of object pFormatFilter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pFilterCol.AddFilter(pFormatFilter, True)
			Case Else
		End Select
		
		pGxDialog.StartingLocation = sLocation
		
		Dim pGxObject As ESRI.ArcGIS.Catalog.IGxObject
		Dim iPos As Short
		Dim sFormat As String ' of parent application when pressing save key on browser
		If (pGxDialog.DoModalSave(hParentHwnd)) Then ' using this form as parent keeps dialog forward
			pGxObject = pGxDialog.FinalLocation
			sLocation = pGxObject.FullName
			
			sName = pGxDialog.Name
			iPos = InStr(sName, ".")
			If iPos > 0 Then sName = Left(sName, iPos - 1)
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
	Public Function GetUniqueFeatureClassName(ByRef pWS As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace, Optional ByRef sPrefix As Object = Nothing) As String
		
		On Error GoTo EH
		Dim sPreName As String
		'UPGRADE_WARNING: Couldn't resolve default property of object sPrefix. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (sPrefix <> "") Then
			'UPGRADE_WARNING: Couldn't resolve default property of object sPrefix. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sPreName = sPrefix
		Else
			sPreName = "ai"
		End If
		Dim i As Integer
		i = 1
		
		' if shapefile workspace add .shp extension
		Dim pDataset As ESRI.ArcGIS.Geodatabase.IDataset
		pDataset = pWS
		Dim sExt As String
		If (InStr(UCase(pDataset.Category), "SHAPEFILE") > 0) Then
			sExt = ".shp"
		Else
			sExt = ""
		End If
		
		Dim done As Boolean
		Dim pFC As ESRI.ArcGIS.Geodatabase.IFeatureClass
		Dim sName As String
		Do While Not done
			sName = sPreName & Right(Str(i), Len(CDbl(Str(i)) - 1)) & sExt
			pFC = pWS.OpenFeatureClass(sName) ' this can raise an error if dataset doesn't exist so error handler should use it
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
	Public Function GetUniqueFileName(ByRef sDir As String, Optional ByRef sPrefix As String = "ai", Optional ByRef sSuffix As String = "") As String
		On Error GoTo ErrorHandler
		
		Dim fs As Scripting.FileSystemObject
		fs = New Scripting.FileSystemObject
		Dim i As Integer
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
			If ((Not fs.FolderExists(sDirNew & "\" & Name)) And (Not fs.FileExists(sDirNew & "\" & Name))) Then
				GetUniqueFileName = Name
				Exit Function
			End If
			i = i + 1
		Loop 
		
		
		Exit Function
ErrorHandler: 
		HandleError(True, "GetUniqueFileName " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Function
	'Returns a Workspace given for example C:\temp\dataset returns C:\temp
	Public Function SplitWorkspaceName(ByRef sWholeName As String) As String
		On Error GoTo ERH
		Dim pos As Short
		pos = InStrRev(sWholeName, "\")
		If pos > 0 Then
			SplitWorkspaceName = Mid(sWholeName, 1, pos - 1)
		Else
			Exit Function
		End If
		Exit Function
ERH: 
		MsgBox("Workspace Split:" & Err.Description)
	End Function
	
	'Returns a filename given for example C:\temp\dataset returns dataset
	Public Function SplitFileName(ByRef sWholeName As String) As String
		On Error GoTo ERH
		Dim pos As Short
		Dim sT As Object
		Dim sName As String
		pos = InStrRev(sWholeName, "\")
		If pos > 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object sT. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
		MsgBox("Workspace Split:" & Err.Description)
	End Function
	
	Public Function MakePerminentGrid(ByRef pRaster As ESRI.ArcGIS.Geodatabase.IRaster, ByRef sOutPath As String, ByRef sOutName As String) As String
		
		On Error GoTo ERH
		Dim pWS As ESRI.ArcGIS.Geodatabase.IWorkspace
		pWS = SetRasterWorkspace(sOutPath)
		Dim pBandC As ESRI.ArcGIS.DataSourcesRaster.IRasterBandCollection
		Dim pBand As ESRI.ArcGIS.DataSourcesRaster.IRasterBand
		pBandC = pRaster
		pBand = pBandC.Item(0)
		Dim pRDS As ESRI.ArcGIS.Geodatabase.IRasterDataset
		pRDS = pBand.RasterDataset
		Dim pDS As ESRI.ArcGIS.Geodatabase.IDataset
		If Not pRDS.CanCopy Then
			Exit Function
		End If
		pRDS.Copy(sOutName, pWS)
		MakePerminentGrid = sOutPath & sOutName
		'UPGRADE_NOTE: Object pRDS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRDS = Nothing
		'UPGRADE_NOTE: Object pDS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pDS = Nothing
		'UPGRADE_NOTE: Object pWS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pWS = Nothing
		Exit Function
ERH: 
		MsgBox("MakePerminebtGrid:" & Err.Description)
	End Function
	
	Public Function ReturnPermanentRaster(ByRef pRaster As ESRI.ArcGIS.Geodatabase.IRaster, ByRef sOutputPath As String, ByRef sOutputName As String) As ESRI.ArcGIS.Geodatabase.IRaster
		
		On Error GoTo ERH
		
		If sOutputPath = "" Or sOutputName = "" Then
			'UPGRADE_NOTE: Object ReturnPermanentRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			ReturnPermanentRaster = Nothing
			Exit Function
		End If
		Dim iPos As Short
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
				MsgBox("Make Permanent Raster: Unsupported file extension")
				Exit Function
		End Select
		
		Dim pWS As ESRI.ArcGIS.DataSourcesRaster.IRasterWorkspace
		pWS = modUtil.SetRasterWorkspace(sOutputPath)
		Dim pBandC As ESRI.ArcGIS.DataSourcesRaster.IRasterBandCollection
		Dim pRasterDataset As ESRI.ArcGIS.Geodatabase.IRasterDataset
		
		pBandC = pRaster
		'UPGRADE_WARNING: Couldn't resolve default property of object pWS. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pBandC.SaveAs(sOutputName, pWS, sFormat)
		
		pRasterDataset = pWS.OpenRasterDataset(sOutputName)
		
		ReturnPermanentRaster = pRasterDataset.CreateDefaultRaster
		Exit Function
ERH: 
		MsgBox("Return Permanent Raster:" & Err.Description)
	End Function
	
	Public Function MakePermanentRaster(ByRef pRaster As ESRI.ArcGIS.Geodatabase.IRaster, ByRef sOutputPath As String, ByRef sOutputName As String) As Boolean
		
		On Error GoTo ERH
		
		Dim pWS As ESRI.ArcGIS.Geodatabase.IWorkspace
		Dim pBandC As ESRI.ArcGIS.DataSourcesRaster.IRasterBandCollection
		
		pWS = modUtil.SetRasterWorkspace(sOutputPath)
		pBandC = pRaster
		
		pBandC.SaveAs(sOutputName, pWS, "GRID")
		
		MakePermanentRaster = True
		
		'UPGRADE_NOTE: Object pWS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pWS = Nothing
		'UPGRADE_NOTE: Object pBandC may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pBandC = Nothing
		
		Exit Function
ERH: 
		MsgBox("Make Permanent Raster:" & Err.Description)
		MakePermanentRaster = False
	End Function
	
	
	Public Function CheckSpatialAnalystLicense() As Boolean
		
		On Error GoTo ErrHandler
		
		Dim pLicManager As ESRI.ArcGIS.esriSystem.IExtensionManager
		Dim pLicAdmin As ESRI.ArcGIS.esriSystem.IExtensionManagerAdmin
		Dim saUID As Object
		Dim pUID As New ESRI.ArcGIS.esriSystem.UID
		Dim v As Object
		Dim pExtension As ESRI.ArcGIS.esriSystem.IExtension
		Dim pExtensionConfig As ESRI.ArcGIS.esriSystem.IExtensionConfig
		
		pLicManager = New ESRI.ArcGIS.esriSystem.ExtensionManager
		pLicAdmin = pLicManager
		
		'UPGRADE_WARNING: Couldn't resolve default property of object saUID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		saUID = "esriSpatialAnalystUI.SAExtension.1"
		'UPGRADE_WARNING: Couldn't resolve default property of object saUID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pUID.Value = saUID
		
		Call pLicAdmin.AddExtension(pUID, v)
		
		pExtension = pLicManager.FindExtension(pUID)
		pExtensionConfig = pExtension
		
		pExtensionConfig.State = ESRI.ArcGIS.esriSystem.esriExtensionState.esriESEnabled
		CheckSpatialAnalystLicense = True
		
		Exit Function
		
ErrHandler: 
		MsgBox("Failed in License Checking" & Err.Description)
		CheckSpatialAnalystLicense = False
		
	End Function
	'Populate combobox of choice with RasterLayers in map
	Public Sub AddRasterLayerToComboBox(ByRef cboBox As System.Windows.Forms.ComboBox, ByRef pMap As ESRI.ArcGIS.Carto.IMap)
		
		On Error GoTo ErrHandler
		cboBox.Items.Clear()
		Dim iLyrIndex As Integer
		Dim pLyr As ESRI.ArcGIS.Carto.ILayer
		' Add raster layers into  Combobox
		Dim iLayerCount As Short
		iLayerCount = pMap.LayerCount
		If iLayerCount > 0 Then
			cboBox.Enabled = True
			For iLyrIndex = 0 To iLayerCount - 1
				pLyr = pMap.Layer(iLyrIndex)
				'UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				If (TypeOf pLyr Is ESRI.ArcGIS.Carto.IRasterLayer) Then
					cboBox.Items.Add(pLyr.Name)
					VB6.SetItemData(cboBox, cboBox.Items.Count - 1, iLyrIndex)
				End If
			Next iLyrIndex
			If (cboBox.Items.Count > 0) Then
				cboBox.SelectedIndex = 0
				cboBox.Text = pMap.Layer(VB6.GetItemData(cboBox, 0)).Name
			End If
		End If
		
		'UPGRADE_NOTE: Object pLyr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pLyr = Nothing
		
		Exit Sub
ErrHandler: 
		MsgBox("Add Layer to ComboBox:" & Err.Description)
		
	End Sub
	
	Public Sub SetFeatureClassName(ByRef pFeatClass As ESRI.ArcGIS.Geodatabase.IFeatureClass, ByRef strName As String)
		On Error GoTo ErrorHandler
		
		
		Dim pDataset As ESRI.ArcGIS.Geodatabase.IDataset
		pDataset = pFeatClass
		
		pDataset.Rename(strName)
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "SetFeatureClassName " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	Public Function LayerInMap(ByRef strName As String, ByRef pMap As ESRI.ArcGIS.Carto.IMap) As Boolean
		On Error GoTo ErrorHandler
		
		Dim lngLyrIndex As Integer
		Dim pLayer As ESRI.ArcGIS.Carto.ILayer
		Dim intLayerCount As Short
		
		intLayerCount = pMap.LayerCount
		If intLayerCount > 0 Then
			For lngLyrIndex = 0 To intLayerCount - 1
				pLayer = pMap.Layer(lngLyrIndex)
				If pLayer.Name = strName Then
					LayerInMap = True
					Exit Function
				Else
					LayerInMap = False
				End If
			Next lngLyrIndex
		End If
		
		'UPGRADE_NOTE: Object pLayer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pLayer = Nothing
		
		
		Exit Function
ErrorHandler: 
		HandleError(True, "LayerInMap " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Function
	
	
	Public Function LayerInMapByFileName(ByRef strName As String, ByRef pMap As ESRI.ArcGIS.Carto.IMap) As Boolean
		
		Dim i As Short
		Dim pFeatureLayer As ESRI.ArcGIS.Carto.IFeatureLayer
		Dim pDataset As ESRI.ArcGIS.Geodatabase.IDataset
		Dim strDatasetname As String
		
		For i = 0 To pMap.LayerCount - 1
			
			'UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If TypeOf pMap.Layer(i) Is ESRI.ArcGIS.Carto.IFeatureLayer Then
				pFeatureLayer = pMap.Layer(i)
				pDataset = pFeatureLayer
				
				'UPGRADE_WARNING: Couldn't resolve default property of object pFeatureLayer.FeatureClass.AliasName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
	
	Public Sub AddLayerFileToMap(ByRef strLyr As String, ByRef pMap As ESRI.ArcGIS.Carto.IMap)
		
		Dim pGxLayer As ESRI.ArcGIS.Catalog.IGxLayer
		Dim pGxFile As ESRI.ArcGIS.Catalog.IGxFile
		
		pGxLayer = New ESRI.ArcGIS.Catalog.GxLayer
		pGxFile = pGxLayer
		
		pGxFile.path = strLyr
		pMap.AddLayer(pGxLayer.Layer)
		
		'UPGRADE_NOTE: Object pGxLayer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pGxLayer = Nothing
		'UPGRADE_NOTE: Object pGxFile may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pGxFile = Nothing
		
	End Sub
	
	
	
	
	Public Function AddRasterLayerToMapFromFileName(ByRef strName As String, ByRef pMap As ESRI.ArcGIS.Carto.IMap) As Boolean
		
		On Error GoTo ErrHandler
		
		Dim pRasterlayer As ESRI.ArcGIS.Carto.IRasterLayer
		pRasterlayer = New ESRI.ArcGIS.Carto.RasterLayer
		
		pRasterlayer.CreateFromFilePath(strName)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object pRasterlayer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pMap.AddLayer(pRasterlayer)
		AddRasterLayerToMapFromFileName = True
		
		'UPGRADE_NOTE: Object pRasterlayer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRasterlayer = Nothing
		
		Exit Function
		
ErrHandler: 
		AddRasterLayerToMapFromFileName = False
		
	End Function
	
	
	Public Function AddFeatureLayerToMapFromFileName(ByRef strName As String, ByRef pMap As ESRI.ArcGIS.Carto.IMap, Optional ByRef strLyrName As String = "") As Boolean
		
		On Error GoTo ErrHandler
		
		Dim pWorkspaceFactory As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory
		Dim pFeatureWorkspace As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace
		Dim pFeatureLayer As ESRI.ArcGIS.Carto.IFeatureLayer
		
		Dim strWorkspace As String
		Dim strFeatClass As String
		
		strWorkspace = SplitWorkspaceName(strName)
		strFeatClass = SplitFileName(strName)
		
		'Create a new ShapefileWorkspaceFactory object and open a shapefile folder
		pWorkspaceFactory = New ESRI.ArcGIS.DataSourcesFile.ShapefileWorkspaceFactory
		pFeatureWorkspace = pWorkspaceFactory.OpenFromFile(strWorkspace, 0)
		
		'Create a new FeatureLayer and assign a shapefile to it
		pFeatureLayer = New ESRI.ArcGIS.Carto.FeatureLayer
		pFeatureLayer.FeatureClass = pFeatureWorkspace.OpenFeatureClass(strFeatClass)
		
		If Len(strLyrName) > 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object pFeatureLayer.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pFeatureLayer.Name = strLyrName
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object pFeatureLayer.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object pFeatureLayer.FeatureClass.AliasName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pFeatureLayer.Name = pFeatureLayer.FeatureClass.AliasName
		End If
		
		'Add the FeatureLayer to the focus map
		'UPGRADE_WARNING: Couldn't resolve default property of object pFeatureLayer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pMap.AddLayer(pFeatureLayer)
		AddFeatureLayerToMapFromFileName = True
		
		'UPGRADE_NOTE: Object pWorkspaceFactory may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pWorkspaceFactory = Nothing
		'UPGRADE_NOTE: Object pFeatureWorkspace may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFeatureWorkspace = Nothing
		'UPGRADE_NOTE: Object pFeatureLayer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFeatureLayer = Nothing
		
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
	Public Function ReturnCellCount(ByRef pRaster As ESRI.ArcGIS.Geodatabase.IRaster) As Double
		On Error GoTo ErrorHandler
		
		Dim pRasterProps As ESRI.ArcGIS.DataSourcesRaster.IRasterProps
		pRasterProps = pRaster
		
		ReturnCellCount = pRasterProps.Height * pRasterProps.Width
		
		'Cleanup
		'UPGRADE_NOTE: Object pRasterProps may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRasterProps = Nothing
		
		Exit Function
ErrorHandler: 
		HandleError(True, "ReturnCellCount " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Function
	
	Public Function GetRasterDistanceUnits(ByRef strLayerName As String, ByRef pApp As ESRI.ArcGIS.Framework.IApplication) As Short
		
		Dim pMxDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
		Dim pMap As ESRI.ArcGIS.Carto.IMap
		Dim pRasterlayer As ESRI.ArcGIS.Carto.IRasterLayer
		Dim pRasterDataset As ESRI.ArcGIS.Geodatabase.IRasterDataset
		Dim pDistUnit As ESRI.ArcGIS.Geometry.ILinearUnit
		Dim intUnit As Short
		Dim pProjCoord As ESRI.ArcGIS.Geometry.IProjectedCoordinateSystem
		
		On Error GoTo ErrHandler
		pMxDoc = pApp.Document
		pMap = pMxDoc.FocusMap
		
		pRasterlayer = pMap.Layer(modUtil.GetLayerIndex(strLayerName, pApp))
		
		If pRasterlayer Is Nothing Then
			Exit Function
		Else
			
			pRasterDataset = pRasterlayer.Raster
			
			If CheckSpatialReference(pRasterDataset) Is Nothing Then
				
				MsgBox("The GRID you have choosen has no spatial reference information.  Please define a projection before continuing.", MsgBoxStyle.Exclamation, "No Project Information Detected")
				
				Exit Function
				
			Else
				
				pProjCoord = CheckSpatialReference(pRasterDataset)
				pDistUnit = pProjCoord.CoordinateUnit
				intUnit = pDistUnit.MetersPerUnit
				
				If intUnit = 1 Then
					GetRasterDistanceUnits = 0
				Else
					GetRasterDistanceUnits = 1
				End If
				
			End If
			
		End If
		
		'UPGRADE_NOTE: Object pRasterlayer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRasterlayer = Nothing
		'UPGRADE_NOTE: Object pRasterDataset may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRasterDataset = Nothing
		'UPGRADE_NOTE: Object pDistUnit may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pDistUnit = Nothing
		'UPGRADE_NOTE: Object pProjCoord may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pProjCoord = Nothing
		
		Exit Function
ErrHandler: 
		Exit Function
		
	End Function
	
	'Using the layer name of a Raster, chases down the Filename of the dataset
	Public Function GetRasterFileName(ByRef strRasterName As String, ByRef pApp As ESRI.ArcGIS.Framework.IApplication) As String
		On Error GoTo ErrorHandler
		
		
		Dim pMxDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
		Dim pMap As ESRI.ArcGIS.Carto.IMap
		Dim pRasterlayer As ESRI.ArcGIS.Carto.IRasterLayer
		
		pMxDoc = pApp.Document
		pMap = pMxDoc.FocusMap
		
		pRasterlayer = pMap.Layer(GetLayerIndex(strRasterName, pApp))
		
		GetRasterFileName = pRasterlayer.FilePath
		
		'UPGRADE_NOTE: Object pMxDoc may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pMxDoc = Nothing
		'UPGRADE_NOTE: Object pMap may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pMap = Nothing
		'UPGRADE_NOTE: Object pRasterlayer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRasterlayer = Nothing
		
		
		
		Exit Function
ErrorHandler: 
		HandleError(True, "GetRasterFileName " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Function
	
	'Takes a layer name, returns filename
	Public Function GetFeatureFileName(ByRef strFeatureName As String, ByRef pApp As ESRI.ArcGIS.Framework.IApplication) As String
		On Error GoTo ErrorHandler
		
		
		Dim pMxDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
		Dim pMap As ESRI.ArcGIS.Carto.IMap
		Dim pFeatureLayer As ESRI.ArcGIS.Carto.IFeatureLayer
		Dim pDataset As ESRI.ArcGIS.Geodatabase.IDataset
		Dim pFeatureClass As ESRI.ArcGIS.Geodatabase.IFeatureClass
		
		pMxDoc = pApp.Document
		pMap = pMxDoc.FocusMap
		
		If Len(strFeatureName) > 0 Then
			pFeatureLayer = pMap.Layer(GetLayerIndex(strFeatureName, pApp))
			pDataset = pFeatureLayer
			pFeatureClass = pFeatureLayer.FeatureClass
			'UPGRADE_WARNING: Couldn't resolve default property of object pFeatureClass.AliasName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetFeatureFileName = pDataset.Workspace.PathName & "\" & pFeatureClass.AliasName
		Else
			Exit Function
		End If
		
		'Cleanup
		'UPGRADE_NOTE: Object pMxDoc may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pMxDoc = Nothing
		'UPGRADE_NOTE: Object pMap may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pMap = Nothing
		'UPGRADE_NOTE: Object pFeatureLayer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFeatureLayer = Nothing
		'UPGRADE_NOTE: Object pDataset may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pDataset = Nothing
		
		
		
		Exit Function
ErrorHandler: 
		HandleError(True, "GetFeatureFileName " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Function
	
	Public Function ReturnRaster(ByRef strRasterFileName As String) As ESRI.ArcGIS.Geodatabase.IRaster
		
		'Takes in FileName, checks for existance of data.
		On Error GoTo ErrHandler
		
		' Create RasterWorkSpaceFactory
		Dim pWSF As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory
		pWSF = New ESRI.ArcGIS.DataSourcesRaster.RasterWorkspaceFactory
		
		' Get RasterWorkspace
		Dim pRasterWS As ESRI.ArcGIS.DataSourcesRaster.IRasterWorkspace
		Dim pRasterDS As ESRI.ArcGIS.Geodatabase.IRasterDataset
		
		If pWSF.IsWorkspace(modUtil.SplitWorkspaceName(strRasterFileName)) Then
			pRasterWS = pWSF.OpenFromFile(modUtil.SplitWorkspaceName(strRasterFileName), 0)
			pRasterDS = pRasterWS.OpenRasterDataset(modUtil.SplitFileName(strRasterFileName))
			ReturnRaster = pRasterDS.CreateDefaultRaster
		End If
		
		'UPGRADE_NOTE: Object pWSF may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pWSF = Nothing
		'UPGRADE_NOTE: Object pRasterWS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRasterWS = Nothing
		'UPGRADE_NOTE: Object pRasterDS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRasterDS = Nothing
		
		Exit Function
		
ErrHandler: 
		'UPGRADE_NOTE: Object ReturnRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		ReturnRaster = Nothing
		MsgBox(Err.Description)
	End Function
	
	Public Function RasterExists(ByRef strRasterFileName As String) As Boolean
		
		'Takes in FileName, checks for existance of data.
		On Error GoTo ErrHandler
		
		' Create RasterWorkSpaceFactory
		Dim pWSF As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory
		pWSF = New ESRI.ArcGIS.DataSourcesRaster.RasterWorkspaceFactory
		
		' Get RasterWorkspace
		Dim pRasterWS As ESRI.ArcGIS.DataSourcesRaster.IRasterWorkspace
		Dim pRasterDS As ESRI.ArcGIS.Geodatabase.IRasterDataset
		
		If pWSF.IsWorkspace(modUtil.SplitWorkspaceName(strRasterFileName)) Then
			pRasterWS = pWSF.OpenFromFile(modUtil.SplitWorkspaceName(strRasterFileName), 0)
			pRasterDS = pRasterWS.OpenRasterDataset(modUtil.SplitFileName(strRasterFileName))
		End If
		
		RasterExists = (Not pRasterDS Is Nothing)
		
		'UPGRADE_NOTE: Object pWSF may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pWSF = Nothing
		'UPGRADE_NOTE: Object pRasterWS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRasterWS = Nothing
		'UPGRADE_NOTE: Object pRasterDS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRasterDS = Nothing
		
		Exit Function
		
ErrHandler: 
		RasterExists = False
		MsgBox(Err.Description)
	End Function
	
	Public Function FeatureExists(ByRef strFeatureFileName As String) As Boolean
		
		On Error GoTo ErrHandler
		
		Dim pWSF As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory
		pWSF = New ESRI.ArcGIS.DataSourcesFile.ShapefileWorkspaceFactory
		
		Dim pFeatWS As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace
		Dim pFeatDS As ESRI.ArcGIS.Geodatabase.IFeatureClass
		
		Dim strWorkspace As String
		Dim strFeatDS As String
		strWorkspace = modUtil.SplitWorkspaceName(strFeatureFileName) & "\"
		strFeatDS = modUtil.SplitFileName(strFeatureFileName)
		
		If pWSF.IsWorkspace(strWorkspace) Then
			pFeatWS = pWSF.OpenFromFile(strWorkspace, 0)
			pFeatDS = pFeatWS.OpenFeatureClass(strFeatDS)
		End If
		
		FeatureExists = (Not pFeatDS Is Nothing)
		
		'UPGRADE_NOTE: Object pWSF may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pWSF = Nothing
		'UPGRADE_NOTE: Object pFeatWS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFeatWS = Nothing
		'UPGRADE_NOTE: Object pFeatDS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFeatDS = Nothing
		
		Exit Function
		
ErrHandler: 
		FeatureExists = False
		
	End Function
	
	Public Function ReturnFeatureClass(ByRef strFeatureFileName As String) As ESRI.ArcGIS.Geodatabase.IFeatureClass
		On Error GoTo ErrHandler
		
		Dim pWSF As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory
		pWSF = New ESRI.ArcGIS.DataSourcesFile.ShapefileWorkspaceFactory
		
		Dim pFeatWS As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace
		Dim pFeatDS As ESRI.ArcGIS.Geodatabase.IFeatureClass
		
		Dim strWorkspace As String
		Dim strFeatDS As String
		strWorkspace = modUtil.SplitWorkspaceName(strFeatureFileName) & "\"
		strFeatDS = modUtil.SplitFileName(strFeatureFileName)
		
		If pWSF.IsWorkspace(strWorkspace) Then
			pFeatWS = pWSF.OpenFromFile(strWorkspace, 0)
			pFeatDS = pFeatWS.OpenFeatureClass(strFeatDS)
		End If
		
		If Not pFeatDS Is Nothing Then
			ReturnFeatureClass = pFeatDS
		Else
			MsgBox("Featureclass " & strFeatureFileName & "does not exist.", MsgBoxStyle.Critical, "Featureclass Not Found")
			Exit Function
		End If
		
		
		'UPGRADE_NOTE: Object pWSF may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pWSF = Nothing
		'UPGRADE_NOTE: Object pFeatWS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFeatWS = Nothing
		'UPGRADE_NOTE: Object pFeatDS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFeatDS = Nothing
		
		Exit Function
		
ErrHandler: 
		MsgBox("Error Returning Featureclass.", MsgBoxStyle.Critical, "Error")
		
		
	End Function
	
	
	Public Function CreateMask(ByRef pMaskRaster As ESRI.ArcGIS.Geodatabase.IGeoDataset) As ESRI.ArcGIS.Geodatabase.IRaster
		On Error GoTo ErrorHandler
		
		'General function for creating analysis masks.  Uses an incoming
		'raster as the mask and a con statement.
		
		Dim pMapAlgebraOp As ESRI.ArcGIS.SpatialAnalyst.IMapAlgebraOp
		Dim strExpression As String
		
		pMapAlgebraOp = New ESRI.ArcGIS.SpatialAnalyst.RasterMapAlgebraOp
		
		With pMapAlgebraOp
			.BindRaster(pMaskRaster, "mask")
		End With
		
		strExpression = "con(isnull([mask]),0,1)"
		
		CreateMask = pMapAlgebraOp.Execute(strExpression)
		
		'Cleanup
		'UPGRADE_NOTE: Object pMapAlgebraOp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pMapAlgebraOp = Nothing
		'UPGRADE_NOTE: Object pMaskRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pMaskRaster = Nothing
		
		
		
		Exit Function
ErrorHandler: 
		HandleError(True, "CreateMask " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Function
	
	Public Function ReturnRasterStretchColorRampRender(ByRef pRasterlayer As ESRI.ArcGIS.Carto.IRasterLayer, ByRef strColor As String) As ESRI.ArcGIS.Carto.IRasterStretchColorRampRenderer
		On Error GoTo ErrorHandler
		
		
		' Get raster input from layer
		Dim pRLayer As ESRI.ArcGIS.Carto.IRasterLayer
		Dim pRaster As ESRI.ArcGIS.Geodatabase.IRaster
		
		' Create renderer and QI RasterRenderer
		Dim pRasStretchMM As ESRI.ArcGIS.Carto.IRasterStretch
		Dim pStretchRen As ESRI.ArcGIS.Carto.IRasterStretchColorRampRenderer
		Dim pRasRen As ESRI.ArcGIS.Carto.IRasterRenderer
		
		Dim strHighValue As String
		Dim strLowValue As String
		
		pRLayer = pRasterlayer
		pRaster = pRLayer.Raster
		
		'init the stretch renderer
		pStretchRen = New ESRI.ArcGIS.Carto.RasterStretchColorRampRenderer
		pRasRen = pStretchRen
		pRasStretchMM = pStretchRen
		
		'Added 8/16, Andrew wanted Min/Max not standard deviation
		With pRasStretchMM
			.StretchType = ESRI.ArcGIS.Carto.esriRasterStretchTypesEnum.esriRasterStretch_StandardDeviations 'esriRasterStretch_MinimumMaximum: changed 10/19/2007 per D. Eslinger
		End With
		
		' Set raster for the renderer and update
		pRasRen.Raster = pRaster
		pRasRen.Update()
		
		' Define two colors
		Dim pFromColor As ESRI.ArcGIS.Display.IHsvColor
		Dim pToColor As ESRI.ArcGIS.Display.IHsvColor
		
		Dim lngHueTo As Integer
		Dim lngSatTo As Integer
		Dim lngValueTo As Integer
		
		Dim lngHueFrom As Integer
		Dim lngSatFrom As Integer
		Dim lngValueFrom As Integer
		
		pFromColor = New ESRI.ArcGIS.Display.HsvColor
		pToColor = New ESRI.ArcGIS.Display.HsvColor
		
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
				
				lngHueTo = CInt(Split(strColor, ",")(0))
				lngSatTo = CInt(Split(strColor, ",")(1))
				lngValueTo = CInt(Split(strColor, ",")(2))
				
				lngHueFrom = CInt(Split(strColor, ",")(3))
				lngSatFrom = CInt(Split(strColor, ",")(4))
				lngValueFrom = CInt(Split(strColor, ",")(5))
				
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
		Dim pRamp As ESRI.ArcGIS.Display.IAlgorithmicColorRamp
		pRamp = New ESRI.ArcGIS.Display.AlgorithmicColorRamp
		
		With pRamp
			'UPGRADE_WARNING: Couldn't resolve default property of object pRamp.Size. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.Size = 2
			.Algorithm = ESRI.ArcGIS.Display.esriColorRampAlgorithm.esriCIELabAlgorithm
			'UPGRADE_WARNING: Couldn't resolve default property of object pFromColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.FromColor = pFromColor
			'UPGRADE_WARNING: Couldn't resolve default property of object pToColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.ToColor = pToColor
			'UPGRADE_WARNING: Couldn't resolve default property of object pRamp.CreateRamp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.CreateRamp(True)
		End With
		
		
		' Plug this colorramp into renderer and select a band
		pStretchRen.BandIndex = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object pRamp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pStretchRen.ColorRamp = pRamp
		
		'Format the labels
		
		strHighValue = Right(pStretchRen.LabelHigh, Len(pStretchRen.LabelHigh) - 7)
		strLowValue = Right(pStretchRen.LabelLow, Len(pStretchRen.LabelLow) - 6)
		
		pStretchRen.LabelHigh = "High : " & VB6.Format(strHighValue, "###,###,###,###,###,##0.00")
		pStretchRen.LabelLow = "Low : " & VB6.Format(strLowValue, "###,###,###,###,###,##0.00")
		
		' Update the renderer with new settings and plug into layer
		pRasRen.Update()
		ReturnRasterStretchColorRampRender = pStretchRen
		
		'Release memeory
		'UPGRADE_NOTE: Object pRLayer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRLayer = Nothing
		'UPGRADE_NOTE: Object pRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRaster = Nothing
		'UPGRADE_NOTE: Object pStretchRen may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pStretchRen = Nothing
		'UPGRADE_NOTE: Object pRasRen may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRasRen = Nothing
		'UPGRADE_NOTE: Object pRamp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRamp = Nothing
		'UPGRADE_NOTE: Object pToColor may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pToColor = Nothing
		'UPGRADE_NOTE: Object pFromColor may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFromColor = Nothing
		
		
		
		Exit Function
ErrorHandler: 
		HandleError(True, "ReturnRasterStretchColorRampRender " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Function
	
	Public Function ReturnUniqueRasterRenderer(ByRef pRasterlayer As ESRI.ArcGIS.Carto.IRasterLayer, ByRef strStandardName As String) As Object
		
		On Error GoTo ErrHandler
		
		' Get raster input from layer
		Dim pRLayer As ESRI.ArcGIS.Carto.IRasterLayer
		Dim pRaster As ESRI.ArcGIS.Geodatabase.IRaster
		Dim pTable As ESRI.ArcGIS.Geodatabase.ITable
		Dim pBand As ESRI.ArcGIS.DataSourcesRaster.IRasterBand
		Dim pBandCol As ESRI.ArcGIS.DataSourcesRaster.IRasterBandCollection
		Dim booTableExist As Boolean
		Dim intNumValues As Short
		Dim strFieldName As String
		Dim lngFieldIndex As Integer
		Dim pColorRed As ESRI.ArcGIS.Display.IColor
		Dim pColorGreen As ESRI.ArcGIS.Display.IColor
		Dim pFSymbol As ESRI.ArcGIS.Display.ISimpleFillSymbol
		Dim pUVRen As ESRI.ArcGIS.Carto.IRasterUniqueValueRenderer
		Dim pRasRen As ESRI.ArcGIS.Carto.IRasterRenderer
		Dim i As Integer
		Dim pRow As ESRI.ArcGIS.Geodatabase.IRow
		Dim LabelValue As Object
		
		pRLayer = pRasterlayer
		pRaster = pRLayer.Raster
		
		'Get the number of rows from raster table
		pBandCol = pRaster
		pBand = pBandCol.Item(0)
		pBand.HasTable(booTableExist)
		If Not booTableExist Then Exit Function
		
		pTable = pBand.AttributeTable
		intNumValues = pTable.RowCount(Nothing)
		
		'Specified a field and get the field index for the specified field to be rendered.     Dim FieldIndex As Integer
		strFieldName = "Value" 'Value is the default field, you can specify other field here..
		'UPGRADE_WARNING: Couldn't resolve default property of object pTable.FindField. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		lngFieldIndex = pTable.FindField(strFieldName)
		
		'Create two colors, red, green
		pColorRed = New ESRI.ArcGIS.Display.RgbColor
		pColorGreen = New ESRI.ArcGIS.Display.RgbColor
		
		pColorRed.RGB = System.Convert.ToUInt32(RGB(214, 71, 0))
		pColorGreen.RGB = System.Convert.ToUInt32(RGB(56, 168, 0))
		
		' Create UniqueValue renderer and QI RasterRenderer
		pUVRen = New ESRI.ArcGIS.Carto.RasterUniqueValueRenderer
		pRasRen = pUVRen
		
		' Connect renderer and raster
		pRasRen.Raster = pRaster
		pRasRen.Update()
		
		' Set UniqueValue renerer
		pUVRen.HeadingCount = 1 ' Use one heading
		pUVRen.Heading(0) = strStandardName
		pUVRen.ClassCount(0) = intNumValues
		pUVRen.Field = strFieldName
		
		For i = 0 To intNumValues - 1
			pRow = pTable.GetRow(i) 'Get a row from the table
			'UPGRADE_WARNING: Couldn't resolve default property of object pRow.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If pRow.Value(lngFieldIndex) = 1 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object pRow.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object LabelValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				LabelValue = pRow.Value(lngFieldIndex) ' Get value of the given index
				pUVRen.AddValue(0, i, LabelValue) 'Set value for the renderer
				pUVRen.Label(0, i) = "Exceeds Standard" ' Set label
				pFSymbol = New ESRI.ArcGIS.Display.SimpleFillSymbol
				'UPGRADE_WARNING: Couldn't resolve default property of object pFSymbol.Color. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object pColorRed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pFSymbol.Color = pColorRed
				'UPGRADE_WARNING: Couldn't resolve default property of object pFSymbol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object pUVRen.Symbol(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pUVRen.Symbol(0, i) = pFSymbol 'Set symbol
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object pRow.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object LabelValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				LabelValue = pRow.Value(lngFieldIndex) ' Get value of the given index
				pUVRen.AddValue(0, i, LabelValue) 'Set value for the renderer
				pUVRen.Label(0, i) = "Below Standard" ' Set label
				pFSymbol = New ESRI.ArcGIS.Display.SimpleFillSymbol
				'UPGRADE_WARNING: Couldn't resolve default property of object pFSymbol.Color. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object pColorGreen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pFSymbol.Color = pColorGreen
				'UPGRADE_WARNING: Couldn't resolve default property of object pFSymbol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object pUVRen.Symbol(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pUVRen.Symbol(0, i) = pFSymbol 'Set symbol
			End If
		Next i
		
		'Update render and refresh layer
		pRasRen.Update()
		ReturnUniqueRasterRenderer = pUVRen
		
		'Clean up
		'UPGRADE_NOTE: Object pRLayer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRLayer = Nothing
		'UPGRADE_NOTE: Object pUVRen may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pUVRen = Nothing
		'UPGRADE_NOTE: Object pRasRen may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRasRen = Nothing
		'UPGRADE_NOTE: Object pFSymbol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFSymbol = Nothing
		'UPGRADE_NOTE: Object pRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRaster = Nothing
		'UPGRADE_NOTE: Object pRLayer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRLayer = Nothing
		'UPGRADE_NOTE: Object pBand may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pBand = Nothing
		'UPGRADE_NOTE: Object pBandCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pBandCol = Nothing
		'UPGRADE_NOTE: Object pTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pTable = Nothing
		'UPGRADE_NOTE: Object pRow may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRow = Nothing
		Exit Function
ErrHandler: 
		MsgBox(Err.Description)
	End Function
	
	Public Function ReturnHSVColorString() As String
		On Error GoTo ErrorHandler
		
		'Returns a comma delimited string of 6 values.  1st 3 a 'To Color' - HIGH, 2nd 3 a 'From Color' - LOW
		Dim intHue As Short
		
		'Hue is a value from 1 to 360 so find a random one
		intHue = Int((360 * Rnd()) + 1)
		
		'Value will be a constant of 97, 100 in the SV and 5, 100..
		ReturnHSVColorString = CStr(intHue) & ",97,100," & CStr(intHue) & ",5,100"
		
		Exit Function
ErrorHandler: 
		HandleError(True, "ReturnHSVColorString " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Function
	
	
	Public Sub CleanupRasterFolder(ByRef strWorkspacePath As String)
		On Error GoTo ErrorHandler
		
		'Used to cleanup the User's workspace and avoid the dreaded -2147467259 error
		
		Dim pWorkspaceFactory As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory
		Dim pWorkspace As ESRI.ArcGIS.Geodatabase.IWorkspace
		Dim pDataset As ESRI.ArcGIS.Geodatabase.IDataset
		Dim pEnumRasterDataset As ESRI.ArcGIS.Geodatabase.IEnumDataset
		Dim pEnv As ESRI.ArcGIS.GeoAnalyst.IRasterAnalysisEnvironment
		
		pWorkspaceFactory = New ESRI.ArcGIS.DataSourcesRaster.RasterWorkspaceFactory
		pWorkspace = pWorkspaceFactory.OpenFromFile(strWorkspacePath, 0)
		
		pEnumRasterDataset = pWorkspace.Datasets(ESRI.ArcGIS.Geodatabase.esriDatasetType.esriDTRasterDataset)
		pEnv = New ESRI.ArcGIS.GeoAnalyst.RasterAnalysis
		
		pDataset = pEnumRasterDataset.Next
		
		Do While Not pDataset Is Nothing
			If InStr(1, pDataset.Name, pEnv.DefaultOutputRasterPrefix, CompareMethod.Text) > 0 Then
				If pDataset.CanDelete Then
					pDataset.Delete()
				End If
			End If
			pDataset = pEnumRasterDataset.Next
		Loop 
		
		'Cleanup
		'UPGRADE_NOTE: Object pWorkspaceFactory may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pWorkspaceFactory = Nothing
		'UPGRADE_NOTE: Object pWorkspace may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pWorkspace = Nothing
		'UPGRADE_NOTE: Object pDataset may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pDataset = Nothing
		'UPGRADE_NOTE: Object pEnumRasterDataset may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pEnumRasterDataset = Nothing
		'UPGRADE_NOTE: Object pEnv may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pEnv = Nothing
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "CleanupRasterFolder " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	Public Sub CleanGlobals()
		On Error GoTo ErrorHandler
		
		'Sub to rid the world of stragling GRIDS, i.e. the ones established for global usse
		
		If Not modMainRun.g_pFeatWorkspace Is Nothing Then
			'UPGRADE_NOTE: Object modMainRun.g_pFeatWorkspace may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			modMainRun.g_pFeatWorkspace = Nothing
		End If
		
		If Not modMainRun.g_pRasWorkspace Is Nothing Then
			'UPGRADE_NOTE: Object modMainRun.g_pRasWorkspace may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			modMainRun.g_pRasWorkspace = Nothing
		End If
		
		'Had an 'elegant' solution using an Iarray to hold global rasters, but that didn't seem to do the
		'job, so we have to manually set each and everyone to nothing
		
		If Not g_pSCS100Raster Is Nothing Then
			'UPGRADE_NOTE: Object g_pSCS100Raster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			g_pSCS100Raster = Nothing
		End If
		
		If Not g_pAbstractRaster Is Nothing Then
			'UPGRADE_NOTE: Object g_pAbstractRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			g_pAbstractRaster = Nothing
		End If
		
		If Not g_pRunoffRaster Is Nothing Then
			'UPGRADE_NOTE: Object g_pRunoffRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			g_pRunoffRaster = Nothing
		End If
		
		If Not g_pRunoffInchRaster Is Nothing Then
			'UPGRADE_NOTE: Object g_pRunoffInchRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			g_pRunoffInchRaster = Nothing
		End If
		
		If Not g_pCellAreaSqMiRaster Is Nothing Then
			'UPGRADE_NOTE: Object g_pCellAreaSqMiRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			g_pCellAreaSqMiRaster = Nothing
		End If
		
		If Not g_pRunoffCFRaster Is Nothing Then
			'UPGRADE_NOTE: Object g_pRunoffCFRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			g_pRunoffCFRaster = Nothing
		End If
		
		If Not g_pRunoffAFRaster Is Nothing Then
			'UPGRADE_NOTE: Object g_pRunoffAFRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			g_pRunoffAFRaster = Nothing
		End If
		
		If Not g_pMetRunoffRaster Is Nothing Then
			'UPGRADE_NOTE: Object g_pMetRunoffRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			g_pMetRunoffRaster = Nothing
		End If
		
		If Not g_pRunoffRaster Is Nothing Then
			'UPGRADE_NOTE: Object g_pRunoffRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			g_pRunoffRaster = Nothing
		End If
		
		If Not g_pDEMRaster Is Nothing Then
			'UPGRADE_NOTE: Object g_pDEMRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			g_pDEMRaster = Nothing
		End If
		
		If Not g_pFlowAccRaster Is Nothing Then
			'UPGRADE_NOTE: Object g_pFlowAccRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			g_pFlowAccRaster = Nothing
		End If
		
		If Not g_pFlowDirRaster Is Nothing Then
			'UPGRADE_NOTE: Object g_pFlowDirRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			g_pFlowDirRaster = Nothing
		End If
		
		If Not g_pLSRaster Is Nothing Then
			'UPGRADE_NOTE: Object g_pLSRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			g_pLSRaster = Nothing
		End If
		
		If Not g_pWaterShedFeatClass Is Nothing Then
			'UPGRADE_NOTE: Object g_pWaterShedFeatClass may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			g_pWaterShedFeatClass = Nothing
		End If
		
		If Not g_KFactorRaster Is Nothing Then
			'UPGRADE_NOTE: Object g_KFactorRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			g_KFactorRaster = Nothing
		End If
		
		If Not g_pPrecipRaster Is Nothing Then
			'UPGRADE_NOTE: Object g_pPrecipRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			g_pPrecipRaster = Nothing
		End If
		
		If Not g_pSoilsRaster Is Nothing Then
			'UPGRADE_NOTE: Object g_pSoilsRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			g_pSoilsRaster = Nothing
		End If
		
		If Not g_LandCoverRaster Is Nothing Then
			'UPGRADE_NOTE: Object g_LandCoverRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			g_LandCoverRaster = Nothing
		End If
		
		If Not g_pSelectedPolyClip Is Nothing Then
			'UPGRADE_NOTE: Object g_pSelectedPolyClip may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			g_pSelectedPolyClip = Nothing
		End If
		
		
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "CleanGlobals " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	Public Function GetSelectedFeatureCount(ByRef pLayer As ESRI.ArcGIS.Carto.ILayer, ByRef pMap As ESRI.ArcGIS.Carto.IMap) As Short
		On Error GoTo ErrorHandler
		
		'Get the # of selected features from a specific layer
		'strName: name of the layer you want to find out # of selected features
		'pMap: Current Map
		
		Dim pFeatLayer As ESRI.ArcGIS.Carto.IFeatureLayer
		Dim pFeatureSelection As ESRI.ArcGIS.Carto.IFeatureSelection
		
		'UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If TypeOf pLayer Is ESRI.ArcGIS.Carto.IFeatureLayer Then
			pFeatLayer = pLayer
		Else
			GetSelectedFeatureCount = 0
			Exit Function
		End If
		
		pFeatureSelection = pFeatLayer
		
		Dim pSelectionSet As ESRI.ArcGIS.Geodatabase.ISelectionSet
		pSelectionSet = pFeatureSelection.SelectionSet
		
		Dim pFeatureCursor As ESRI.ArcGIS.Geodatabase.IFeatureCursor
		'UPGRADE_WARNING: Couldn't resolve default property of object pFeatureCursor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pSelectionSet.Search(Nothing, True, pFeatureCursor)
		
		GetSelectedFeatureCount = pSelectionSet.Count
		
		'Cleanup
		'UPGRADE_NOTE: Object pFeatLayer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFeatLayer = Nothing
		'UPGRADE_NOTE: Object pFeatureSelection may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFeatureSelection = Nothing
		'UPGRADE_NOTE: Object pSelectionSet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pSelectionSet = Nothing
		'UPGRADE_NOTE: Object pFeatureCursor may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFeatureCursor = Nothing
		
		
		
		Exit Function
ErrorHandler: 
		HandleError(True, "GetSelectedFeatureCount " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Function
	
	Public Function ClipBySelectedPoly(ByRef pRaster As ESRI.ArcGIS.Geodatabase.IRaster, ByRef pClipPoly As ESRI.ArcGIS.Geometry.IGeometry, ByRef pEnv As ESRI.ArcGIS.GeoAnalyst.IRasterAnalysisEnvironment) As ESRI.ArcGIS.Geodatabase.IRaster
		On Error GoTo ErrorHandler
		
		'Clip the raster by analysis extent poly
		'pRaster: Incoming raster to be clipped
		'pClipPoly: the polygon doing the clipping
		'pEnv: the Raster analysis environmnet
		
		' Create the RasterExtractionOp object
		Dim pExtractionOp As ESRI.ArcGIS.SpatialAnalyst.IExtractionOp
		Dim pClipPolygon As ESRI.ArcGIS.Geometry.IPolygon
		
		pExtractionOp = New ESRI.ArcGIS.SpatialAnalyst.RasterExtractionOp
		pClipPolygon = pClipPoly
		
		'Set pEnv = g_pSpatEnv
		g_pSpatEnv = pExtractionOp
		
		' Call the method
		'UPGRADE_WARNING: Couldn't resolve default property of object pRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ClipBySelectedPoly = pExtractionOp.Polygon(pRaster, pClipPolygon, True)
		
		'Cleanup
		'UPGRADE_NOTE: Object pExtractionOp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pExtractionOp = Nothing
		'UPGRADE_NOTE: Object pClipPolygon may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pClipPolygon = Nothing
		'Set pRaster = Nothing
		
		
		
		Exit Function
ErrorHandler: 
		HandleError(True, "ClipBySelectedPoly " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Function
	
	Public Sub ExportLayerToPath(ByRef pLayer As ESRI.ArcGIS.Carto.ILayer, ByRef sPath As String)
		
		On Error GoTo ExportLayerToPath_ERR
		
		Dim pGxLayer As ESRI.ArcGIS.Catalog.IGxLayer
		Dim pGxFile As ESRI.ArcGIS.Catalog.IGxFile
		
		pGxLayer = New ESRI.ArcGIS.Catalog.GxLayer
		pGxFile = pGxLayer
		
		pGxFile.path = sPath
		
		pGxLayer.Layer = pLayer
		'UPGRADE_NOTE: Object pGxFile may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pGxFile = Nothing
		'UPGRADE_NOTE: Object pGxLayer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pGxLayer = Nothing
		
		Exit Sub
		
ExportLayerToPath_ERR: 
		Debug.Print("ExportLayerToPath_ERR: " & Err.Description)
		System.Diagnostics.Debug.Assert(0, "")
		
	End Sub
	
	
	Public Sub CreateMetadata(ByRef pGroupLayer As ESRI.ArcGIS.Carto.IGroupLayer, ByRef strProjectInfo As String)
		On Error GoTo ErrorHandler
		
		'This sub will create metadata for the final group layer of raster outputs.  The idea is to use
		'a xml template shipped with N-SPECT to provide the base.  Next the metadata is synchronized which fills
		'in all of the spatial particulars.  Finally, specific parameters are set.  The xml file is shipped in
		'NSPECTDAT\metadata.
		
		Dim fs As Object
		Dim booExists As Boolean
		
		'Metadata objects
		Dim pGxFile As ESRI.ArcGIS.Catalog.IGxFile
		Dim pMetaData As ESRI.ArcGIS.Geodatabase.IMetadata
		Dim pPropSet As ESRI.ArcGIS.esriSystem.IPropertySet
		Dim pPropXMLSet As ESRI.ArcGIS.Geodatabase.IXmlPropertySet
		Dim pSelMD As ESRI.ArcGIS.Geodatabase.IMetadata
		Dim pXPS As ESRI.ArcGIS.Geodatabase.IXmlPropertySet2
		Dim pXPS2 As ESRI.ArcGIS.Geodatabase.IXmlPropertySet2
		
		Dim sPath As String
		Dim strMeta As String
		Dim i As Short
		Dim lyrCount As Integer
		
		'Group layer and raster stuff
		Dim pLayer As ESRI.ArcGIS.Carto.ILayer
		Dim pCompositeLayer As ESRI.ArcGIS.Carto.ICompositeLayer
		Dim pGroupLayerPiece As ESRI.ArcGIS.Carto.ILayer
		Dim pRasterlayer As ESRI.ArcGIS.Carto.IRasterLayer
		Dim pRasterDataset As ESRI.ArcGIS.Geodatabase.IRasterDataset
		Dim pRasterWorkSpaceFact As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory
		Dim pRasterWorkSpace As ESRI.ArcGIS.DataSourcesRaster.IRasterWorkspace
		
		sPath = modUtil.g_nspectPath & "\metadata\metadata.xml"
		
		fs = CreateObject("Scripting.FileSystemObject")
		'UPGRADE_WARNING: Couldn't resolve default property of object fs.FileExists. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		booExists = fs.FileExists(sPath)
		
		If booExists = False Then
			MsgBox("The N-SPECT metadata template: " & sPath & " was not found." & vbNewLine & "Please provide an existing Xml Document", MsgBoxStyle.Exclamation)
			Exit Sub
		End If
		
		'Set up the metadata
		pMetaData = New ESRI.ArcGIS.Catalog.GxMetadata
		pGxFile = pMetaData
		pGxFile.path = sPath
		
		pPropSet = pMetaData.Metadata
		pXPS2 = pPropSet
		
		'Get all the Xml from the template
		strMeta = pXPS2.GetXml("")
		
		'Now get ready to get a hold of the Rasterdatasets
		pRasterWorkSpaceFact = New ESRI.ArcGIS.DataSourcesRaster.RasterWorkspaceFactory
		
		'set up player
		pLayer = pGroupLayer
		
		lyrCount = 0
		
		'UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If Not TypeOf pLayer Is ESRI.ArcGIS.Carto.IGroupLayer Then
			Exit Sub
		End If
		
		pCompositeLayer = pLayer
		
		'We know this grouplayer will have rasters in it, so set up pRasterLayer so we
		'can go from there.
		pRasterlayer = pCompositeLayer.Layer(0)
		pRasterWorkSpace = pRasterWorkSpaceFact.OpenFromFile(SplitWorkspaceName((pRasterlayer.FilePath)), 0)
		
		'Now with the workspace set, go to work on each of the layers
		For i = 0 To pCompositeLayer.Count - 1
			
			pGroupLayerPiece = pCompositeLayer.Layer(i)
			'UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If TypeOf pGroupLayerPiece Is ESRI.ArcGIS.Carto.IRasterLayer Then
				
				modProgDialog.ProgDialog("Creating metadata for " & pGroupLayerPiece.Name & "...", "Completing Analysis", 0, (pCompositeLayer.Count), lyrCount, (frmPrj.m_App.hwnd))
				If modProgDialog.g_boolCancel Then
					
					
					pRasterlayer = pGroupLayerPiece
					'Have to get the RasterDataset in order to get metadata
					pRasterDataset = pRasterWorkSpace.OpenRasterDataset(SplitFileName((pRasterlayer.FilePath)))
					
					pSelMD = pRasterDataset
					pPropXMLSet = pSelMD.Metadata
					pXPS = pPropXMLSet
					
					'Write the template Xml to the selected dataset
					'This will delete any existing xml metadata in the selected dataset
					pXPS.SetXml((strMeta))
					
					'Now set seperate properties
					With pXPS
						.SetPropertyX("idinfo/citation/citeinfo/pubdate", Now, ESRI.ArcGIS.Geodatabase.esriXmlPropertyType.esriXPTText, ESRI.ArcGIS.Geodatabase.esriXmlSetPropertyAction.esriXSPAAddOrReplace, False)
						'UPGRADE_WARNING: Couldn't resolve default property of object pRasterlayer.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.SetPropertyX("idinfo/citation/citeinfo/title", pRasterlayer.Name, ESRI.ArcGIS.Geodatabase.esriXmlPropertyType.esriXPTText, ESRI.ArcGIS.Geodatabase.esriXmlSetPropertyAction.esriXSPAAddOrReplace, False)
						.SetPropertyX("idinfo/descript/supplinf", strProjectInfo, ESRI.ArcGIS.Geodatabase.esriXmlPropertyType.esriXPTText, ESRI.ArcGIS.Geodatabase.esriXmlSetPropertyAction.esriXSPAAddOrReplace, False)
						'UPGRADE_WARNING: Couldn't resolve default property of object pRasterlayer.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.SetPropertyX("dataqual/lineage/procstep/procdesc", g_dicMetadata.Item(pRasterlayer.Name), ESRI.ArcGIS.Geodatabase.esriXmlPropertyType.esriXPTText, ESRI.ArcGIS.Geodatabase.esriXmlSetPropertyAction.esriXSPAAddOrReplace, False)
						.SetPropertyX("dataqual/lineage/procstep/procdate", Now, ESRI.ArcGIS.Geodatabase.esriXmlPropertyType.esriXPTText, ESRI.ArcGIS.Geodatabase.esriXmlSetPropertyAction.esriXSPAAddOrReplace, False)
						.SetPropertyX("metainfo/metd", Now, ESRI.ArcGIS.Geodatabase.esriXmlPropertyType.esriXPTText, ESRI.ArcGIS.Geodatabase.esriXmlSetPropertyAction.esriXSPAAddOrReplace, False)
					End With
					
					'Save  the metadata
					'UPGRADE_WARNING: Couldn't resolve default property of object pXPS. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					pSelMD.Metadata = pXPS
					pSelMD.Synchronize(ESRI.ArcGIS.Geodatabase.esriMetadataSyncAction.esriMSAAccessed, 1)
					
					System.Windows.Forms.Application.DoEvents()
					
				End If
				
			End If
			lyrCount = lyrCount + 1
		Next i
		
		
		'kill dialog
		modProgDialog.KillDialog()
		
		'Cleanup
		'UPGRADE_NOTE: Object pGxFile may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pGxFile = Nothing
		'UPGRADE_NOTE: Object pMetaData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pMetaData = Nothing
		'UPGRADE_NOTE: Object pPropSet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pPropSet = Nothing
		'UPGRADE_NOTE: Object pPropXMLSet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pPropXMLSet = Nothing
		'UPGRADE_NOTE: Object pSelMD may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pSelMD = Nothing
		'UPGRADE_NOTE: Object pXPS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pXPS = Nothing
		'UPGRADE_NOTE: Object pXPS2 may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pXPS2 = Nothing
		'UPGRADE_NOTE: Object pLayer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pLayer = Nothing
		'UPGRADE_NOTE: Object pCompositeLayer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pCompositeLayer = Nothing
		'UPGRADE_NOTE: Object pGroupLayerPiece may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pGroupLayerPiece = Nothing
		'UPGRADE_NOTE: Object pRasterlayer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRasterlayer = Nothing
		'UPGRADE_NOTE: Object pRasterDataset may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRasterDataset = Nothing
		'UPGRADE_NOTE: Object pRasterWorkSpaceFact may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRasterWorkSpaceFact = Nothing
		'UPGRADE_NOTE: Object pRasterWorkSpace may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRasterWorkSpace = Nothing
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "CreateMetadata " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	Public Function ReturnRasterMax(ByRef pRaster As ESRI.ArcGIS.Geodatabase.IRaster) As Double
		On Error GoTo ErrorHandler
		
		Dim pRasterBandCollection As ESRI.ArcGIS.DataSourcesRaster.IRasterBandCollection
		Dim pRasterBand As ESRI.ArcGIS.DataSourcesRaster.IRasterBand
		Dim pRasterStats As ESRI.ArcGIS.DataSourcesRaster.IRasterStatistics
		
		pRasterBandCollection = pRaster
		pRasterBand = pRasterBandCollection.Item(0)
		
		pRasterStats = pRasterBand.Statistics
		
		ReturnRasterMax = pRasterStats.Maximum
		
		'Clean
		'UPGRADE_NOTE: Object pRasterBandCollection may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRasterBandCollection = Nothing
		'UPGRADE_NOTE: Object pRasterBand may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRasterBand = Nothing
		'UPGRADE_NOTE: Object pRasterStats may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRasterStats = Nothing
		
		Exit Function
ErrorHandler: 
		HandleError(True, "ReturnRasterMax " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Function
	
	Public Function ParseProjectforMetadata(ByRef clsPrjXML As clsXMLPrjFile, ByRef strPrjFileName As String) As String
		
		Dim strFormattedProject As String
		Dim i As Short
		
		'Start by getting all the particulars that must be there...
		strFormattedProject = "N-SPECT Project Information: " & vbNewLine & vbTab & "N-SPECT Project Name: " & clsPrjXML.strProjectName & vbNewLine & vbTab & "N-SPECT Project File Location: " & strPrjFileName & vbNewLine & vbTab & "N-SPECT Parameter Database: " & modUtil.g_nspectPath & "\nspect.mdb" & vbNewLine & vbTab & "Project Working Directory: " & clsPrjXML.strProjectWorkspace & vbNewLine & vbTab & "Watershed Delineation: " & clsPrjXML.strWaterShedDelin
		ParseProjectforMetadata = strFormattedProject
		
	End Function
	
	Public Function ExportSelectedFeatures(ByRef pLayer As ESRI.ArcGIS.Carto.ILayer, ByRef pMap As ESRI.ArcGIS.Carto.IMap, ByRef pWorkspace As ESRI.ArcGIS.Geodatabase.IWorkspace, ByRef strFeatClass As String) As ESRI.ArcGIS.Geodatabase.IFeatureClass
		On Error GoTo ErrorHandler
		
		'Incoming
		'pLayer: Layer user has chosen as being the one from which the selected polys will come
		'pMap: current map
		'pWorkspace:  place to put the exported selected polys
		'strFeatClass: string file location of featureclass
		
		Dim pFLayer As ESRI.ArcGIS.Carto.IFeatureLayer
		Dim pFC As ESRI.ArcGIS.Geodatabase.IFeatureClass
		Dim pINFeatureClassName As ESRI.ArcGIS.Geodatabase.IFeatureClassName
		Dim pDataset As ESRI.ArcGIS.Geodatabase.IDataset
		Dim pInDsName As ESRI.ArcGIS.Geodatabase.IDatasetName
		Dim pFSel As ESRI.ArcGIS.Carto.IFeatureSelection
		Dim pSelSet As ESRI.ArcGIS.Geodatabase.ISelectionSet
		Dim pFeatureClassName As ESRI.ArcGIS.Geodatabase.IFeatureClassName
		Dim pOutDatasetName As ESRI.ArcGIS.Geodatabase.IDatasetName
		Dim pWorkspaceName As ESRI.ArcGIS.Geodatabase.IWorkspaceName
		Dim pExportOp As ESRI.ArcGIS.GeoDatabaseUI.IExportOperation
		Dim pFeatWS As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace
		Dim pFeatClass As ESRI.ArcGIS.Geodatabase.IFeatureClass
		
		Dim pSelectGeometry As ESRI.ArcGIS.Geometry.IGeometry
		Dim pBasinFeatClass As ESRI.ArcGIS.Geodatabase.IFeatureClass
		Dim pSpatialFilter As ESRI.ArcGIS.Geodatabase.ISpatialFilter
		Dim pBasinSelSet As ESRI.ArcGIS.Geodatabase.ISelectionSet
		
		'Set the layer to the incoming one and get featureclass
		pFLayer = pLayer
		
		'Get the selection set
		pFSel = pFLayer
		pSelSet = pFSel.SelectionSet
		
		'Make a call to the function below to return a unioned geometry of any and all selected features
		pSelectGeometry = modMainRun.ReturnSelectGeometry(pSelSet)
		g_pSelectedPolyClip = pSelectGeometry
		
		'Make a call to get the BasinPoly featureclass using the name sent over
		pBasinFeatClass = modUtil.ReturnFeatureClass(strFeatClass)
		
		'Now we select all basinpolys that intersect the unioned geometry we got earlier
		pSpatialFilter = New ESRI.ArcGIS.Geodatabase.SpatialFilter
		
		With pSpatialFilter
			.Geometry = pSelectGeometry
			.GeometryField = pBasinFeatClass.ShapeFieldName
			.SpatialRel = ESRI.ArcGIS.Geodatabase.esriSpatialRelEnum.esriSpatialRelContains
		End With
		
		'UPGRADE_WARNING: Couldn't resolve default property of object pSpatialFilter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pBasinSelSet = pBasinFeatClass.Select(pSpatialFilter, ESRI.ArcGIS.Geodatabase.esriSelectionType.esriSelectionTypeIDSet, ESRI.ArcGIS.Geodatabase.esriSelectionOption.esriSelectionOptionNormal, pWorkspace)
		
		'Get the FcName from the featureclass
		pDataset = pBasinFeatClass
		
		'Set classname and dataset
		pINFeatureClassName = pDataset.FullName
		pInDsName = pINFeatureClassName
		
		'Create a new feature class name
		' Define the output feature class name
		pFeatureClassName = New ESRI.ArcGIS.Geodatabase.FeatureClassName
		pOutDatasetName = pFeatureClassName
		
		'UPGRADE_WARNING: Couldn't resolve default property of object pWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pOutDatasetName.Name = modUtil.GetUniqueFeatureClassName(pWorkspace, "selpoly")
		
		pWorkspaceName = New ESRI.ArcGIS.Geodatabase.WorkspaceName
		pWorkspaceName.PathName = pWorkspace.PathName
		pWorkspaceName.WorkspaceFactoryProgID = "esriDataSourcesFile.shapefileworkspacefactory.1"
		
		pOutDatasetName.WorkspaceName = pWorkspaceName
		
		With pFeatureClassName
			.FeatureType = ESRI.ArcGIS.Geodatabase.esriFeatureType.esriFTSimple
			.ShapeType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolygon
			.ShapeFieldName = "Shape"
		End With
		
		'Export
		pExportOp = New ESRI.ArcGIS.GeoDatabaseUI.ExportOperation
		'UPGRADE_WARNING: Couldn't resolve default property of object pOutDatasetName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pExportOp.ExportFeatureClass(pInDsName, Nothing, pBasinSelSet, Nothing, pOutDatasetName, 0)
		
		pFeatWS = pWorkspace
		pFeatClass = pFeatWS.OpenFeatureClass(pOutDatasetName.Name)
		ExportSelectedFeatures = pFeatClass
		
		'Cleanup
		'UPGRADE_NOTE: Object pFLayer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFLayer = Nothing
		'UPGRADE_NOTE: Object pFC may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFC = Nothing
		'UPGRADE_NOTE: Object pINFeatureClassName may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pINFeatureClassName = Nothing
		'UPGRADE_NOTE: Object pDataset may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pDataset = Nothing
		'UPGRADE_NOTE: Object pInDsName may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pInDsName = Nothing
		'UPGRADE_NOTE: Object pFSel may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFSel = Nothing
		'UPGRADE_NOTE: Object pSelSet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pSelSet = Nothing
		'UPGRADE_NOTE: Object pFeatureClassName may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFeatureClassName = Nothing
		'UPGRADE_NOTE: Object pOutDatasetName may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pOutDatasetName = Nothing
		'UPGRADE_NOTE: Object pWorkspaceName may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pWorkspaceName = Nothing
		'UPGRADE_NOTE: Object pExportOp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pExportOp = Nothing
		'UPGRADE_NOTE: Object pFeatWS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFeatWS = Nothing
		
		Exit Function
ErrorHandler: 
		HandleError(True, "ExportSelectedFeatures " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Function
End Module