Option Strict Off
Option Explicit On
Module modLanduse
	' *************************************************************************************************
	' *  Perot Systems Government Services
	' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
	' *  modLanduse
	' *************************************************************************************************
	' *  Description: Code mod for handling Landuse scenarios of frmPrj
	' *         Sub Begin:: The main sub; take in params from project form and busts
	' *                     out the good stuff.  Adds new classes
	' *         Sub CopyCoefficient:: Essentially finds the current coefficient set and copies it for use
	' *                               in the new land use
	' *         Function ReclassLanduse::  Adds the 'new' landclasses to the land cover dataset
	' *         Function ReclassRaster::  Creates the new landcover raster
	' *         Function Cleanup:: Rid the Access database of temporary entries holding the land cover
	' *                            and new coefficient stuff
	' *
	' *  Called By:  frmPrj::CmdOK
	' **************************************************************************************************
	
	Private m_intLCTypeID As Short
	Private m_intLCClassID As Short
	Private m_intCoeffSetID As Short
	Private m_strLCFileName As String
	
	Private m_strLCClass As String
	Private m_pEnv As ESRI.ArcGIS.GeoAnalyst.IRasterAnalysisEnvironment
	Private m_pRasterProps As ESRI.ArcGIS.DataSourcesRaster.IRasterProps
	Private m_pLandCoverRaster As ESRI.ArcGIS.Geodatabase.IRaster
	'Private m_pWSFactory As IWorkspaceFactory
	Private m_pWS As ESRI.ArcGIS.Geodatabase.IWorkspace
	Private m_pMap As ESRI.ArcGIS.Carto.IMap
	
	Public g_booLCChange As Boolean
	Public g_strLCTypeName As String 'the temp name of Land Cover type, if indeed landuses are applied.
	Public g_DictTempNames As New Scripting.Dictionary 'Array holding the temp names of the LCType, and subsequent Coefficient Set Names
	' Constant used by the Error handler function - DO NOT REMOVE
	Const c_sModuleFileName As String = "modLanduse.bas"
	Private m_ParentHWND As Integer ' Set this to get correct parenting of Error handler forms
	
	
	
	
	Public Sub Begin(ByRef strLCClassName As String, ByRef clsLUScenItems As clsXMLLandUseItems, ByRef dictPollutants As Scripting.Dictionary, ByRef strLCFileName As String, ByRef pMap As ESRI.ArcGIS.Carto.IMap, ByRef strWorkspace As String)
		'strLCClassName: str of current land cover class
		'clsLUScenItems: XML class that holds the params of the user's Land Use Scenario
		'dictPollutants: dictionary created to hold the pollutants of this particular project
		'strLCFileName: FileName of land cover grid
		
		On Error GoTo ErrHandler
		
		Dim rsCurrentLCType As New ADODB.Recordset 'Current LCTYpe
		Dim strCurrentLCType As String
		Dim rsCurrentLCClass As New ADODB.Recordset 'Current Landclasses of The LCType
		Dim strCurrentLCClass As String
		Dim strInsertTempLCType As String
		Dim strTempLCTypeName As String
		Dim rsInsertLCType As New ADODB.Recordset 'the inserted LCTYPE
		Dim intLCTypeID As Short
		Dim strrsInsertLCType As String
		Dim rsCloneLCClass As New ADODB.Recordset 'new temp landclasses
		Dim rsPermLandClass As New ADODB.Recordset 'the landclass table currently in the database
		Dim strrsPermLandClass As String
		Dim rsNewLandClass As New ADODB.Recordset 'Newly added landclass - to get its new ID
		Dim strNewLandClass As String
		
		Dim i As Short
		Dim j As Short
		Dim k As Short
		Dim intValue As Short 'Temp new landclasses fake value
		Dim clsLUScen As New clsXMLLUScen 'XML Land use scenario
		Dim strCoeffSetTempName As String 'New temp name for the coefficient set
		Dim strCoeffSetOrigName As String 'Original name for the coefficient set
		Dim pollArray() As Object 'Array to hold pollutants...get itself from the dictionary
		
		'init the file name of the LandClass File
77: m_strLCFileName = strLCFileName
78: m_pMap = pMap
		'init the workspace
		m_pWS = modUtil.SetRasterWorkspace(strWorkspace)
		
		'STEP 1: Get the current LCTYPE
81: strCurrentLCType = "SELECT * FROM LCTYPE WHERE NAME LIKE '" & strLCClassName & "'"
82: rsCurrentLCType.Open(strCurrentLCType, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
		
		'STEP 2: Get the current LCCLASSES of the current LCTYPE
85: strCurrentLCClass = "SELECT * FROM LCCLASS WHERE" & " LCTYPEID = " & rsCurrentLCType.Fields("LCTypeID").Value & " ORDER BY LCCLass.Value"
87: rsCurrentLCClass.Open(strCurrentLCClass, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic)
		
		'STEP 3: Now INSERT a copy of current LCTYPE
		'First, get a temp name... like CCAPLUTemp1, CCAPLUTemp2 etc
91: strTempLCTypeName = modUtil.CreateUniqueName("LCTYPE", strLCClassName & "LUTemp")
92: strInsertTempLCType = "INSERT INTO LCTYPE (NAME,DESCRIPTION) VALUES ('" & Replace(strTempLCTypeName, "'", "''") & "', '" & Replace(strTempLCTypeName, "'", "''") & " Description')"
		
		'Add the name to the Dictionary for storage; used for cleanup and in pollutants
95: g_DictTempNames.Add(strLCClassName, strTempLCTypeName)
		
		'STEP 4: INSERT the copy of the LCTYPE in
98: modUtil.g_ADOConn.Execute(strInsertTempLCType)
		
		'STEP 5: Get it back now so you can use its ID for inserting the landclasses
101: strrsInsertLCType = "SELECT LCTYPEID FROM LCTYPE " & "WHERE LCTYPE.NAME LIKE '" & strTempLCTypeName & "'"
		
104: rsInsertLCType.Open(strrsInsertLCType, g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockReadOnly)
105: m_intLCTypeID = rsInsertLCType.Fields("LCTypeID").Value
		
		'STEP 6: Now clone the current landclasses into a new recordset
108: rsCloneLCClass = rsCurrentLCClass.Clone
		
		'Prepare the landclass table to accept the copies of landclass
111: strrsPermLandClass = "SELECT * FROM LCCLASS"
112: rsPermLandClass.Open(strrsPermLandClass, g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
		
		'STEP 7: loop through the clsXMLLandUseItems to add new land uses to the copy rs
115: rsCloneLCClass.MoveFirst()
		'Now add all the landclasses.
117: Do While Not rsCloneLCClass.EOF
			'Add new
119: rsPermLandClass.AddNew()
			'Add the necessary components
121: rsPermLandClass.Fields("LCTypeID").Value = m_intLCTypeID
122: rsPermLandClass.Fields("Value").Value = rsCloneLCClass.Fields("Value").Value
123: rsPermLandClass.Fields("Name").Value = rsCloneLCClass.Fields("Name").Value
124: rsPermLandClass.Fields("CN-A").Value = rsCloneLCClass.Fields("CN-A").Value
125: rsPermLandClass.Fields("CN-B").Value = rsCloneLCClass.Fields("CN-B").Value
126: rsPermLandClass.Fields("CN-C").Value = rsCloneLCClass.Fields("CN-C").Value
127: rsPermLandClass.Fields("CN-D").Value = rsCloneLCClass.Fields("CN-D").Value
128: rsPermLandClass.Fields("CoverFactor").Value = rsCloneLCClass.Fields("CoverFactor").Value
129: rsPermLandClass.Fields("W_WL").Value = rsCloneLCClass.Fields("W_WL").Value
			'Update
131: rsPermLandClass.Update()
			'move to next record
133: intValue = rsCloneLCClass.Fields("Value").Value 'This keeps track of the max
134: rsCloneLCClass.MoveNext()
135: Loop 
		
		'STEP 8: Now add the new landclass
		
		Dim rsTemp As ADODB.Recordset
		
		Dim intLCClassIDs() As Short
		ReDim intLCClassIDs(clsLUScenItems.Count - 1)
		
144: For i = 1 To clsLUScenItems.Count
145: If clsLUScenItems.Item(i).intApply = 1 Then
				'init the fake value: will be max value + 1
147: intValue = intValue + 1
				
				'Init the clsLUScen
150: clsLUScen.XML = clsLUScenItems.Item(i).strLUScenXMLFile
				
152: rsPermLandClass.AddNew()
153: rsPermLandClass.Fields("LCTypeID").Value = m_intLCTypeID
154: rsPermLandClass.Fields("Value").Value = intValue
155: rsPermLandClass.Fields("Name").Value = clsLUScen.strLUScenName
156: rsPermLandClass.Fields("CN-A").Value = clsLUScen.intSCSCurveA
157: rsPermLandClass.Fields("CN-B").Value = clsLUScen.intSCSCurveB
158: rsPermLandClass.Fields("CN-C").Value = clsLUScen.intSCSCurveC
159: rsPermLandClass.Fields("CN-D").Value = clsLUScen.intSCSCurveD
160: rsPermLandClass.Fields("CoverFactor").Value = clsLUScen.lngCoverFactor
161: rsPermLandClass.Fields("W_WL").Value = clsLUScen.intWaterWetlands
162: rsPermLandClass.Update()
				
164: End If
			
			'Gather the newly added LCClassIds in an array for use later
167: strNewLandClass = "SELECT LCCLASSID FROM LCCLASS WHERE NAME LIKE '" & clsLUScen.strLUScenName & "'"
168: rsNewLandClass.Open(strNewLandClass, g_ADOConn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
169: intLCClassIDs(i - 1) = rsNewLandClass.Fields("LCClassID").Value
170: rsNewLandClass.Close()
171: Next i
		
		'STEP 10: Parse the incoming dictionary
		'Remember, it's in this form:
		'Key_________Item___________
		'Pollutant , CoefficientSet
		Dim l As Short
178: 'UPGRADE_WARNING: Couldn't resolve default property of object dictPollutants.Keys. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pollArray = dictPollutants.Keys
		
		'Now to the pollutants
		'Loop through the pollutants coming from the XML class, as well as those in the project that are being used
182: For j = 1 To clsLUScen.clsPollItems.Count
183: For k = 0 To dictPollutants.Count - 1
184: 'UPGRADE_WARNING: Couldn't resolve default property of object pollArray(k). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If InStr(1, clsLUScen.clsPollItems.Item(j).strPollName, pollArray(k), CompareMethod.Text) > 0 Then
185: 'UPGRADE_WARNING: Couldn't resolve default property of object dictPollutants.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					strCoeffSetOrigName = dictPollutants.Item(pollArray(k)) 'Original Name
186: strCoeffSetTempName = modUtil.CreateUniqueName("COEFFICIENTSET", strCoeffSetOrigName & "_Temp") 'Make a new temp name
					
					'Now add names of the coefficient sets to the dictionary
189: g_DictTempNames.Add(strCoeffSetOrigName, strCoeffSetTempName)
					
191: rsTemp = New ADODB.Recordset
192: rsTemp = CopyCoefficient(strCoeffSetTempName, strCoeffSetOrigName) 'Call to function that returns the
					'copied record set.
					
					'Now add the new values using the info in the XML file and the array of new LCClass IDs
196: For l = 0 To UBound(intLCClassIDs)
197: rsTemp.AddNew()
198: clsLUScen.XML = clsLUScenItems.Item(l + 1).strLUScenXMLFile
						
200: rsTemp.Fields("Coeff1").Value = clsLUScen.clsPollItems.Item(j).intType1
201: rsTemp.Fields("Coeff2").Value = clsLUScen.clsPollItems.Item(j).intType2
202: rsTemp.Fields("Coeff3").Value = clsLUScen.clsPollItems.Item(j).intType3
203: rsTemp.Fields("Coeff4").Value = clsLUScen.clsPollItems.Item(j).intType4
204: rsTemp.Fields("CoeffSetID").Value = m_intCoeffSetID
205: rsTemp.Fields("LCClassID").Value = intLCClassIDs(l)
						
207: rsTemp.Update()
						
209: Next l
210: rsTemp.Close()
211: End If
212: Next k
213: Next j
		
		'Mop up after yourself
216: rsCurrentLCType.Close()
217: rsCurrentLCClass.Close()
218: rsInsertLCType.Close()
219: rsPermLandClass.Close()
		'rsNewLandClass.Close
		
222: 'UPGRADE_NOTE: Object rsCurrentLCType may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsCurrentLCType = Nothing
223: 'UPGRADE_NOTE: Object rsCurrentLCClass may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsCurrentLCClass = Nothing
224: 'UPGRADE_NOTE: Object rsInsertLCType may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsInsertLCType = Nothing
225: 'UPGRADE_NOTE: Object rsPermLandClass may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsPermLandClass = Nothing
226: 'UPGRADE_NOTE: Object rsNewLandClass may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsNewLandClass = Nothing
		
228: g_strLCTypeName = strTempLCTypeName
229: ReclassLanduse(clsLUScenItems, strTempLCTypeName, m_strLCFileName)
		
		Exit Sub
ErrHandler: 
233: MsgBox("Error Number: " & Err.Number & vbNewLine & "Error Description: " & Err.Description)
		
	End Sub
	
	Private Function CopyCoefficient(ByRef strNewCoeffName As String, ByRef strCoeffSet As String) As ADODB.Recordset
		On Error GoTo ErrHandler
		'General gist:  First we add new record to the Coefficient Set table using strNewCoeffName as
		'the name, PollID, LCTYPEID.  Once that's done, we'll add the coefficients
		'from the set being copied
		Dim strCopySet As String 'The Recordset of existing coefficients being copied
		Dim strNewLcType As String 'CmdString for inserting new coefficientset               '
		Dim strNewCoeff As String
		Dim strNewCoeffID As String 'Holder for the CoefficientSetID
		Dim intCoeffSetID As Short
		Dim i As Short
		
		Dim rsCopySet As New ADODB.Recordset 'First RS
		Dim rsCoeffSetID As New ADODB.Recordset '1 record recordset to get a hold of CoeffsetID
		Dim rsNewCoeff As New ADODB.Recordset
		
253: strCopySet = "SELECT * FROM COEFFICIENTSET INNER JOIN COEFFICIENT ON COEFFICIENTSET.COEFFSETID = " & "COEFFICIENT.COEFFSETID WHERE COEFFICIENTSET.NAME LIKE '" & strCoeffSet & "'"
		
256: rsCopySet.Open(strCopySet, g_ADOConn, ADODB.CursorTypeEnum.adOpenKeyset)
		
258: Debug.Print(rsCopySet.RecordCount)
259: Debug.Print(rsCopySet.Fields("POLLID").Value)
		
		'INSERT: new Coefficient set taking the PollID and LCType ID from rsCopySet
262: strNewLcType = "INSERT INTO COEFFICIENTSET(NAME, POLLID, LCTYPEID) VALUES ('" & Replace(strNewCoeffName, "'", "''") & "'," & rsCopySet.Fields("POLLID").Value & "," & m_intLCTypeID & ")"
		
		'First need to add the coefficient set to that table
266: g_ADOConn.Execute(strNewLcType, ADODB.AffectEnum.adAffectCurrent)
		
		'Get the Coefficient Set ID of the newly created coefficient set
269: strNewCoeffID = "SELECT COEFFSETID FROM COEFFICIENTSET " & "WHERE COEFFICIENTSET.NAME LIKE '" & strNewCoeffName & "'"
		
272: rsCoeffSetID.Open(strNewCoeffID, g_ADOConn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockPessimistic)
273: m_intCoeffSetID = rsCoeffSetID.Fields("CoeffSetID").Value
		
		'Now loopy loo to populate values.
		'Get the coefficient table
		Dim strNewCoeff1 As String
278: strNewCoeff1 = "SELECT * FROM COEFFICIENT"
279: rsNewCoeff.Open(strNewCoeff1, g_ADOConn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		
281: rsCopySet.MoveFirst()
		
		'Actually add the records to the new set
284: For i = 1 To rsCopySet.RecordCount
			
			'Add New one
287: rsNewCoeff.AddNew()
			
			'Add the necessary components
290: rsNewCoeff.Fields("Coeff1").Value = rsCopySet.Fields("Coeff1").Value
291: rsNewCoeff.Fields("Coeff2").Value = rsCopySet.Fields("Coeff2").Value
292: rsNewCoeff.Fields("Coeff3").Value = rsCopySet.Fields("Coeff3").Value
293: rsNewCoeff.Fields("Coeff4").Value = rsCopySet.Fields("Coeff4").Value
294: rsNewCoeff.Fields("CoeffSetID").Value = m_intCoeffSetID
295: rsNewCoeff.Fields("LCClassID").Value = rsCopySet.Fields("LCClassID").Value
			
297: rsNewCoeff.Update()
298: rsCopySet.MoveNext()
			
300: Next i
		
302: CopyCoefficient = rsNewCoeff
		
		'Cleanup
305: rsCopySet.Close()
306: rsCoeffSetID.Close()
		'rsNewCoeff.Close
		
309: 'UPGRADE_NOTE: Object rsCopySet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsCopySet = Nothing
310: 'UPGRADE_NOTE: Object rsCoeffSetID may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsCoeffSetID = Nothing
		'Set rsNewCoeff = Nothing
		
		Exit Function
ErrHandler: 
		
316: MsgBox("Error Number: " & Err.Number & vbNewLine & "Error Description: " & Err.Description, MsgBoxStyle.Critical, "Error in modutil.copycoefficient")
		
	End Function
	
	Private Sub ReclassLanduse(ByRef clsLUScenItems As clsXMLLandUseItems, ByRef strLCClass As String, ByRef strLCFileName As String)
		'clsLUScenItems: which is a collection of the landuse entered
		'strLCClass: Name of the LCTYPE being altered
		'strLCFileName: path to which the landcover grid exists
		
		Dim strExpress As String
		Dim strExpression As String
		Dim strOutLandCover As String
		Dim pNewLandCoverRaster As ESRI.ArcGIS.Geodatabase.IRaster
		Dim pTempDeleteRaster As ESRI.ArcGIS.Geodatabase.IRaster
		
		Dim i As Short
		
		Dim pMapAlgebraOp As ESRI.ArcGIS.SpatialAnalyst.IMapAlgebraOp
		Dim pArray As ESRI.ArcGIS.esriSystem.IArray
		On Error GoTo ErrHandler
		
		'init everything
338: m_strLCClass = strLCClass
339: pArray = New ESRI.ArcGIS.esriSystem.Array
340: m_pEnv = New ESRI.ArcGIS.GeoAnalyst.RasterAnalysis
		
		'Make sure the landcoverraster exists..it better if they get to this point, ED!
343: If g_LandCoverRaster Is Nothing Then
344: If modUtil.RasterExists(strLCFileName) Then
345: m_pLandCoverRaster = modUtil.ReturnRaster(strLCFileName)
346: Else
				Exit Sub
348: End If
349: Else
350: m_pLandCoverRaster = g_LandCoverRaster
351: End If
		
		'Get the rasterprops of the landcover raster
354: m_pRasterProps = m_pLandCoverRaster
		
		'Get the workspace
357: 'Set m_pWSFactory = New RasterWorkspaceFactory
358: 'Set m_pWS = m_pWSFactory.OpenFromFile(modUtil.SplitWorkspaceName(strLCFileName), frmPrj.m_App.hWnd)
		'Set m_pWS = modUtil.SetRasterWorkspace(strWorkspace)
		'Set the environment
361: With m_pEnv
362: .SetCellSize(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvValue, m_pRasterProps.MeanCellSize.X)
363: .SetExtent(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvValue, m_pRasterProps.Extent)
364: 'UPGRADE_WARNING: Couldn't resolve default property of object m_pEnv.OutSpatialReference. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.OutSpatialReference = m_pRasterProps.SpatialReference
365: 'UPGRADE_WARNING: Couldn't resolve default property of object m_pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.OutWorkspace = m_pWS
366: End With
		
		
		'Going to now take each entry in the landuse scenarios, if they're not ignored, go
		'to ReclassRaster to make a new raster of the Value indicated.  Then add the returned
		'raster to m_RasterArray.  We'll then make a map algebra statement out of that collection
		'of rasters to merge back into the original landcover raster...voila.
		
374: If clsLUScenItems.Count > 0 Then
375: For i = 1 To clsLUScenItems.Count
376: If clsLUScenItems.Item(i).intApply = 1 Then
377: modProgDialog.ProgDialog("Processing Landuse scenario...", "Landuse Scenario", 0, CInt(clsLUScenItems.Count + 1), CInt(i), 0)
378: If modProgDialog.g_boolCancel Then
379: pArray.Add(ReclassRaster(clsLUScenItems.Item(i), m_strLCClass))
380: End If
					
382: End If
383: Next i
384: End If
		
386: pMapAlgebraOp = New ESRI.ArcGIS.SpatialAnalyst.RasterMapAlgebraOp
387: m_pEnv = pMapAlgebraOp
		
389: With m_pEnv
390: .SetCellSize(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvValue, m_pRasterProps.MeanCellSize.X)
391: .SetExtent(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvValue, m_pRasterProps.Extent)
392: 'UPGRADE_WARNING: Couldn't resolve default property of object m_pEnv.OutSpatialReference. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.OutSpatialReference = m_pRasterProps.SpatialReference
393: 'UPGRADE_WARNING: Couldn't resolve default property of object m_pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.OutWorkspace = m_pWS
394: End With
		'Now, if the raster array has some rasters in it, let's merge the fuckers together
396: If pArray.Count > 0 Then
397: For i = 0 To pArray.Count - 1
398: 'UPGRADE_WARNING: Couldn't resolve default property of object pArray.element(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pMapAlgebraOp.BindRaster(pArray.element(i), "raster" & CStr(i))
399: If strExpress = "" Then
400: strExpress = "Merge([" & "raster" & CStr(i) & "], "
401: Else
402: strExpress = strExpress & "[" & "raster" & CStr(i) & "], "
403: End If
404: Next i
			
406: 'UPGRADE_WARNING: Couldn't resolve default property of object m_pLandCoverRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pMapAlgebraOp.BindRaster(m_pLandCoverRaster, "landcover")
407: strExpression = strExpress & "[landcover])"
			
409: modProgDialog.ProgDialog("Creating new landcover GRID...", "Landuse Scenario", 0, CInt(clsLUScenItems.Count + 1), CInt(i + 1), 0)
			
411: If modProgDialog.g_boolCancel Then
412: pNewLandCoverRaster = pMapAlgebraOp.Execute(strExpression)
413: End If
414: Else
			Exit Sub
416: End If
		
418: 'UPGRADE_WARNING: Couldn't resolve default property of object m_pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		strOutLandCover = modUtil.GetUniqueName("landcover", modUtil.SplitWorkspaceName(m_pEnv.OutWorkspace.PathName))
		
		
		
422: 'UPGRADE_WARNING: Couldn't resolve default property of object m_pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If modUtil.MakePermanentRaster(pNewLandCoverRaster, m_pEnv.OutWorkspace.PathName, strOutLandCover) Then
423: g_LandCoverRaster = pNewLandCoverRaster 'MsgBox "done"
424: End If
		
		'Cleanup
427: For i = 0 To pArray.Count - 1
428: pTempDeleteRaster = pArray.element(i)
429: 'UPGRADE_NOTE: Object pTempDeleteRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			pTempDeleteRaster = Nothing
430: Next i
		
432: modProgDialog.KillDialog()
		
434: 'UPGRADE_NOTE: Object clsLUScenItems may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		clsLUScenItems = Nothing
435: 'UPGRADE_NOTE: Object pMapAlgebraOp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pMapAlgebraOp = Nothing
436: 'UPGRADE_NOTE: Object m_pWS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_pWS = Nothing
437: 'Set m_pWSFactory = Nothing
		
		Exit Sub
ErrHandler: 
441: MsgBox("error in MSSetup " & Err.Number & ": " & Err.Description)
		
	End Sub
	
	Private Function ReclassRaster(ByRef clsLUItem As clsXMLLandUseItem, ByRef strLCClass As String) As ESRI.ArcGIS.Geodatabase.IRaster
		On Error GoTo ErrorHandler
		
		'We're passing over a single land use scenario in the form of the xml
		'class clsXMLLandUseItem, seems to be the easiest way to do this.
		Dim strSelect As String 'ADO selections string
		Dim rsValue As ADODB.Recordset 'ADO recordset
		Dim dblValue As Double 'value
		Dim strOutName As String 'String for poly's name
		Dim strExpression As String 'Map calculator expression
		Dim clsLUItemDetails As New clsXMLLUScen 'The particulars in the landuse
		
		Dim pConversionOp As ESRI.ArcGIS.GeoAnalyst.IConversionOp 'Conversion for poly to raster
		Dim pPolyFeatureClass As ESRI.ArcGIS.Geodatabase.IFeatureClass 'The polygon featureclass
		Dim pPolyFeatLayer As ESRI.ArcGIS.Carto.ILayer 'Polygon featurelayer
		Dim pPolyWS As ESRI.ArcGIS.Geodatabase.IWorkspace
		Dim lngPolyFeatIndex As Integer
		Dim pPolyDS As ESRI.ArcGIS.Geodatabase.IGeoDataset
		Dim pOutPolyRDS As ESRI.ArcGIS.Geodatabase.IRasterDataset
		Dim pMapAlgebraOp As ESRI.ArcGIS.SpatialAnalyst.IMapAlgebraOp
		Dim pPolyValueRaster As ESRI.ArcGIS.Geodatabase.IRaster
		Dim pFinalPolyValueRaster As ESRI.ArcGIS.Geodatabase.IRaster 'Final poly new raster value
		
		
		
		'STEP 1: Open the landclass Value Value -------------------------------------------------------------------------
		'This is the value user's landclass will change to
471: strSelect = "SELECT LCTYPE.LCTYPEID, LCCLASS.NAME, LCCLASS.VALUE FROM " & "LCTYPE INNER JOIN LCCLASS ON LCTYPE.LCTYPEID = LCCLASS.LCTYPEID " & "WHERE LCTYPE.NAME LIKE '" & strLCClass & "' AND LCCLASS.NAME LIKE '" & clsLUItem.strLUScenName & "'"
		
475: rsValue = New ADODB.Recordset
476: rsValue.Open(strSelect, g_ADOConn, ADODB.CursorTypeEnum.adOpenKeyset)
		
478: dblValue = rsValue.Fields("Value").Value
479: rsValue.Close()
		
		'init the landuse xml stuff
482: clsLUItemDetails.XML = clsLUItem.strLUScenXMLFile
		
		'END STEP 1: ----------------------------------------------------------------------------------------------------
		
		'STEP 2: --------------------------------------------------------------------------------------------------------
		'Init
488: pConversionOp = New ESRI.ArcGIS.GeoAnalyst.RasterConversionOp
		'Set the spat Environment
490: m_pEnv = pConversionOp
		
492: With m_pEnv
493: .SetCellSize(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvValue, m_pRasterProps.MeanCellSize.X)
494: .SetExtent(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvValue, m_pRasterProps.Extent)
495: 'UPGRADE_WARNING: Couldn't resolve default property of object m_pEnv.OutSpatialReference. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.OutSpatialReference = m_pRasterProps.SpatialReference
496: 'UPGRADE_WARNING: Couldn't resolve default property of object m_pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.OutWorkspace = m_pWS
497: End With
		'END STEP 2: ----------------------------------------------------------------------------------------------------
		
		'STEP 3: --------------------------------------------------------------------------------------------------------
		'Convert the polygon featureclass into a Raster for sending back out
		'Get the featureclass, check for selected features
		
504: If clsLUItemDetails.intLUScenSelectedPoly = 1 And m_pMap.SelectionCount > 0 Then
505: lngPolyFeatIndex = modUtil.GetLayerIndex((clsLUItemDetails.strLUScenLyrName), (frmPrj.m_App))
506: pPolyFeatLayer = m_pMap.Layer(lngPolyFeatIndex)
			pPolyWS = modUtil.SetFeatureShapeWorkspace(modUtil.SplitWorkspaceName((clsLUItemDetails.strLUScenFileName)))
507: pPolyFeatureClass = modUtil.ExportSelectedFeatures(pPolyFeatLayer, m_pMap, pPolyWS, (clsLUItemDetails.strLUScenFileName))
508: Else
509: pPolyFeatureClass = modUtil.ReturnFeatureClass((clsLUItemDetails.strLUScenFileName))
510: End If
		'Have to use a GeoDataset, so QI
512: pPolyDS = pPolyFeatureClass
		'Init the out dataset
514: pOutPolyRDS = New ESRI.ArcGIS.DataSourcesRaster.RasterDataset
515: strOutName = modUtil.GetUniqueName("poly", (m_pWS.PathName))
516: pOutPolyRDS = pConversionOp.ToRasterDataset(pPolyDS, "GRID", m_pWS, strOutName)
		'END STEP 3 ------------------------------------------------------------------------------------------------------
		
		'STEP 4: ---------------------------------------------------------------------------------------------------------
		'Need to now make that raster the value the user choose way back when
521: pMapAlgebraOp = New ESRI.ArcGIS.SpatialAnalyst.RasterMapAlgebraOp
522: m_pEnv = pMapAlgebraOp
		
524: With m_pEnv
525: .SetCellSize(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvValue, m_pRasterProps.MeanCellSize.X)
526: .SetExtent(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvValue, m_pRasterProps.Extent)
527: 'UPGRADE_WARNING: Couldn't resolve default property of object m_pEnv.OutSpatialReference. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.OutSpatialReference = m_pRasterProps.SpatialReference
528: 'UPGRADE_WARNING: Couldn't resolve default property of object m_pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.OutWorkspace = m_pWS
529: End With
		
531: pPolyValueRaster = pOutPolyRDS.CreateDefaultRaster
		
533: 'UPGRADE_WARNING: Couldn't resolve default property of object pPolyValueRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pMapAlgebraOp.BindRaster(pPolyValueRaster, "poly")
		
535: strExpression = "Con([poly] <> -9999, " & dblValue & ", 0)"
		
		
538: pFinalPolyValueRaster = pMapAlgebraOp.Execute(strExpression)
		'End STEP 4: -----------------------------------------------------------------------------------------------------
		
541: ReclassRaster = pFinalPolyValueRaster
		
		'Cleanup
544: 'UPGRADE_NOTE: Object rsValue may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsValue = Nothing
545: 'UPGRADE_NOTE: Object pConversionOp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pConversionOp = Nothing
546: 'UPGRADE_NOTE: Object pPolyFeatureClass may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pPolyFeatureClass = Nothing
547: 'UPGRADE_NOTE: Object pPolyDS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pPolyDS = Nothing
548: 'UPGRADE_NOTE: Object pOutPolyRDS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pOutPolyRDS = Nothing
549: 'UPGRADE_NOTE: Object pMapAlgebraOp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pMapAlgebraOp = Nothing
550: 'UPGRADE_NOTE: Object pPolyValueRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pPolyValueRaster = Nothing
551: 'UPGRADE_NOTE: Object pFinalPolyValueRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFinalPolyValueRaster = Nothing
552: 'UPGRADE_NOTE: Object clsLUItemDetails may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		clsLUItemDetails = Nothing
		
		Exit Function
		
		
		
		
		
		
		Exit Function
ErrorHandler: 
		HandleError(False, "ReclassRaster " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Function
	
	Public Sub Cleanup(ByRef dictNames As Scripting.Dictionary, ByRef clsPollItems As clsXMLPollutantItems, ByRef strLCTypeName As String)
		On Error GoTo ErrorHandler
		
		
		Dim strDeleteCoeffSet As String
		Dim strDeleteLCType As String
		Dim strCoeffDeleteName As String
		Dim strLCDeleteName As String
		Dim i As Short
		
		
577: 'UPGRADE_WARNING: Couldn't resolve default property of object dictNames.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		strLCDeleteName = dictNames.Item(strLCTypeName)
		
579: If Len(strLCDeleteName) = 0 Then
			Exit Sub
581: End If
		
		Dim strLCTypeDelete As String
		Dim strLCClassDelete As String
		
		Dim rsLCTypeDelete As ADODB.Recordset
		
588: strLCTypeDelete = "SELECT * FROM LCTYPE WHERE NAME LIKE '" & strLCDeleteName & "'"
		
590: rsLCTypeDelete = New ADODB.Recordset
		
592: rsLCTypeDelete.CursorLocation = ADODB.CursorLocationEnum.adUseClient
593: rsLCTypeDelete.Open(strLCTypeDelete, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockOptimistic)
		
595: strLCClassDelete = "Delete * FROM LCCLASS WHERE LCTYPEID =" & rsLCTypeDelete.Fields("LCTypeID").Value
		
597: modUtil.g_ADOConn.Execute(strLCClassDelete)
		
		'Set up a delete rs and get rid of it
600: rsLCTypeDelete.Delete(ADODB.AffectEnum.adAffectCurrent)
601: rsLCTypeDelete.Update()
		
603: For i = 1 To clsPollItems.Count
604: 'UPGRADE_WARNING: Couldn't resolve default property of object dictNames.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			strCoeffDeleteName = dictNames.Item(clsPollItems.Item(i).strCoeffSet)
605: If Len(strCoeffDeleteName) > 0 Then
606: strDeleteCoeffSet = "DELETE * FROM COEFFICIENTSET WHERE NAME LIKE '" & strCoeffDeleteName & "'"
607: modUtil.g_ADOConn.Execute(strDeleteCoeffSet)
608: End If
609: Next i
		
611: rsLCTypeDelete.Close()
		
		
614: 'UPGRADE_NOTE: Object rsLCTypeDelete may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsLCTypeDelete = Nothing
		
		
		
		
		
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "Cleanup " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
End Module