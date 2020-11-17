Attribute VB_Name = "modLanduse"
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

Option Explicit
Private m_intLCTypeID As Integer
Private m_intLCClassID As Integer
Private m_intCoeffSetID As Integer
Private m_strLCFileName As String

Private m_strLCClass As String
Private m_pEnv As IRasterAnalysisEnvironment
Private m_pRasterProps As IRasterProps
Private m_pLandCoverRaster As IRaster
'Private m_pWSFactory As IWorkspaceFactory
Private m_pWS As IWorkspace
Private m_pMap As IMap

Public g_booLCChange As Boolean
Public g_strLCTypeName As String                    'the temp name of Land Cover type, if indeed landuses are applied.
Public g_DictTempNames As New Dictionary            'Array holding the temp names of the LCType, and subsequent Coefficient Set Names
' Constant used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "modLanduse.bas"
Private m_ParentHWND As Long    ' Set this to get correct parenting of Error handler forms




Public Sub Begin(strLCClassName As String, clsLUScenItems As clsXMLLandUseItems, _
                dictPollutants As Dictionary, strLCFileName As String, pMap As IMap, strWorkspace As String)
'strLCClassName: str of current land cover class
'clsLUScenItems: XML class that holds the params of the user's Land Use Scenario
'dictPollutants: dictionary created to hold the pollutants of this particular project
'strLCFileName: FileName of land cover grid
    
On Error GoTo ErrHandler:

    Dim rsCurrentLCType As New ADODB.Recordset      'Current LCTYpe
    Dim strCurrentLCType As String
    Dim rsCurrentLCClass As New ADODB.Recordset     'Current Landclasses of The LCType
    Dim strCurrentLCClass As String
    Dim strInsertTempLCType As String
    Dim strTempLCTypeName As String
    Dim rsInsertLCType As New ADODB.Recordset       'the inserted LCTYPE
    Dim intLCTypeID As Integer
    Dim strrsInsertLCType As String
    Dim rsCloneLCClass As New ADODB.Recordset       'new temp landclasses
    Dim rsPermLandClass As New ADODB.Recordset      'the landclass table currently in the database
    Dim strrsPermLandClass As String
    Dim rsNewLandClass As New ADODB.Recordset       'Newly added landclass - to get its new ID
    Dim strNewLandClass As String
    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim intValue As Integer                         'Temp new landclasses fake value
    Dim clsLUScen As New clsXMLLUScen               'XML Land use scenario
    Dim strCoeffSetTempName As String               'New temp name for the coefficient set
    Dim strCoeffSetOrigName As String               'Original name for the coefficient set
    Dim pollArray() As Variant                      'Array to hold pollutants...get itself from the dictionary
    
    'init the file name of the LandClass File
77:     m_strLCFileName = strLCFileName
78:     Set m_pMap = pMap
    'init the workspace
    Set m_pWS = modUtil.SetRasterWorkspace(strWorkspace)
    
    'STEP 1: Get the current LCTYPE
81:     strCurrentLCType = "SELECT * FROM LCTYPE WHERE NAME LIKE '" & strLCClassName & "'"
82:     rsCurrentLCType.Open strCurrentLCType, modUtil.g_ADOConn, adOpenDynamic, adLockOptimistic
    
    'STEP 2: Get the current LCCLASSES of the current LCTYPE
85:     strCurrentLCClass = "SELECT * FROM LCCLASS WHERE" & _
            " LCTYPEID = " & rsCurrentLCType!LCTypeID & " ORDER BY LCCLass.Value"
87:     rsCurrentLCClass.Open strCurrentLCClass, modUtil.g_ADOConn, adOpenStatic, adLockBatchOptimistic
    
    'STEP 3: Now INSERT a copy of current LCTYPE
    'First, get a temp name... like CCAPLUTemp1, CCAPLUTemp2 etc
91:     strTempLCTypeName = modUtil.CreateUniqueName("LCTYPE", strLCClassName & "LUTemp")
92:     strInsertTempLCType = "INSERT INTO LCTYPE (NAME,DESCRIPTION) VALUES ('" & _
            Replace(strTempLCTypeName, "'", "''") & "', '" & _
            Replace(strTempLCTypeName, "'", "''") & " Description')"
    
    'Add the name to the Dictionary for storage; used for cleanup and in pollutants
95:     g_DictTempNames.Add strLCClassName, strTempLCTypeName
    
    'STEP 4: INSERT the copy of the LCTYPE in
98:     modUtil.g_ADOConn.Execute strInsertTempLCType
    
    'STEP 5: Get it back now so you can use its ID for inserting the landclasses
101:     strrsInsertLCType = "SELECT LCTYPEID FROM LCTYPE " & _
        "WHERE LCTYPE.NAME LIKE '" & strTempLCTypeName & "'"
            
104:     rsInsertLCType.Open strrsInsertLCType, g_ADOConn, adOpenDynamic, adLockReadOnly
105:     m_intLCTypeID = rsInsertLCType!LCTypeID
    
    'STEP 6: Now clone the current landclasses into a new recordset
108:     Set rsCloneLCClass = rsCurrentLCClass.Clone
    
    'Prepare the landclass table to accept the copies of landclass
111:     strrsPermLandClass = "SELECT * FROM LCCLASS"
112:     rsPermLandClass.Open strrsPermLandClass, g_ADOConn, adOpenDynamic, adLockOptimistic
    
    'STEP 7: loop through the clsXMLLandUseItems to add new land uses to the copy rs
115:     rsCloneLCClass.MoveFirst
    'Now add all the landclasses.
117:     Do While Not rsCloneLCClass.EOF
        'Add new
119:         rsPermLandClass.AddNew
        'Add the necessary components
121:         rsPermLandClass!LCTypeID = m_intLCTypeID
122:         rsPermLandClass!Value = rsCloneLCClass!Value
123:         rsPermLandClass!Name = rsCloneLCClass!Name
124:         rsPermLandClass![CN-A] = rsCloneLCClass![CN-A]
125:         rsPermLandClass![CN-B] = rsCloneLCClass![CN-B]
126:         rsPermLandClass![CN-C] = rsCloneLCClass![CN-C]
127:         rsPermLandClass![CN-D] = rsCloneLCClass![CN-D]
128:         rsPermLandClass!CoverFactor = rsCloneLCClass!CoverFactor
129:         rsPermLandClass![W_WL] = rsCloneLCClass![W_WL]
        'Update
131:         rsPermLandClass.Update
        'move to next record
133:         intValue = rsCloneLCClass!Value     'This keeps track of the max
134:         rsCloneLCClass.MoveNext
135:     Loop
        
    'STEP 8: Now add the new landclass
    
    Dim rsTemp As ADODB.Recordset
    
    Dim intLCClassIDs() As Integer
    ReDim intLCClassIDs(clsLUScenItems.Count - 1)
    
144:     For i = 1 To clsLUScenItems.Count
145:         If clsLUScenItems.Item(i).intApply = 1 Then
            'init the fake value: will be max value + 1
147:             intValue = intValue + 1
            
            'Init the clsLUScen
150:             clsLUScen.XML = clsLUScenItems.Item(i).strLUScenXMLFile
            
152:             rsPermLandClass.AddNew
153:             rsPermLandClass!LCTypeID = m_intLCTypeID
154:             rsPermLandClass!Value = intValue
155:             rsPermLandClass!Name = clsLUScen.strLUScenName
156:             rsPermLandClass![CN-A] = clsLUScen.intSCSCurveA
157:             rsPermLandClass![CN-B] = clsLUScen.intSCSCurveB
158:             rsPermLandClass![CN-C] = clsLUScen.intSCSCurveC
159:             rsPermLandClass![CN-D] = clsLUScen.intSCSCurveD
160:             rsPermLandClass!CoverFactor = clsLUScen.lngCoverFactor
161:             rsPermLandClass![W_WL] = clsLUScen.intWaterWetlands
162:             rsPermLandClass.Update
            
164:         End If
        
            'Gather the newly added LCClassIds in an array for use later
167:             strNewLandClass = "SELECT LCCLASSID FROM LCCLASS WHERE NAME LIKE '" & clsLUScen.strLUScenName & "'"
168:             rsNewLandClass.Open strNewLandClass, g_ADOConn, adOpenStatic, adLockReadOnly
169:             intLCClassIDs(i - 1) = rsNewLandClass!LCClassID
170:             rsNewLandClass.Close
171:     Next i
            
    'STEP 10: Parse the incoming dictionary
    'Remember, it's in this form:
       'Key_________Item___________
       'Pollutant , CoefficientSet
    Dim l As Integer
178:     pollArray = dictPollutants.Keys

    'Now to the pollutants
    'Loop through the pollutants coming from the XML class, as well as those in the project that are being used
182:     For j = 1 To clsLUScen.clsPollItems.Count
183:         For k = 0 To dictPollutants.Count - 1
184:             If InStr(1, clsLUScen.clsPollItems.Item(j).strPollName, pollArray(k), vbTextCompare) > 0 Then
185:                 strCoeffSetOrigName = dictPollutants.Item(pollArray(k)) 'Original Name
186:                 strCoeffSetTempName = modUtil.CreateUniqueName("COEFFICIENTSET", strCoeffSetOrigName & "_Temp") 'Make a new temp name
                
                'Now add names of the coefficient sets to the dictionary
189:                 g_DictTempNames.Add strCoeffSetOrigName, strCoeffSetTempName
                
191:                 Set rsTemp = New ADODB.Recordset
192:                 Set rsTemp = CopyCoefficient(strCoeffSetTempName, strCoeffSetOrigName) 'Call to function that returns the
                'copied record set.
                
                'Now add the new values using the info in the XML file and the array of new LCClass IDs
196:                 For l = 0 To UBound(intLCClassIDs)
197:                 rsTemp.AddNew
198:                     clsLUScen.XML = clsLUScenItems.Item(l + 1).strLUScenXMLFile
                    
200:                     rsTemp!Coeff1 = clsLUScen.clsPollItems.Item(j).intType1
201:                     rsTemp!Coeff2 = clsLUScen.clsPollItems.Item(j).intType2
202:                     rsTemp!Coeff3 = clsLUScen.clsPollItems.Item(j).intType3
203:                     rsTemp!Coeff4 = clsLUScen.clsPollItems.Item(j).intType4
204:                     rsTemp!CoeffSetID = m_intCoeffSetID
205:                     rsTemp!LCClassID = intLCClassIDs(l)

207:                 rsTemp.Update
                
209:                 Next l
210:                 rsTemp.Close
211:             End If
212:         Next k
213:     Next j

    'Mop up after yourself
216:     rsCurrentLCType.Close
217:     rsCurrentLCClass.Close
218:     rsInsertLCType.Close
219:     rsPermLandClass.Close
    'rsNewLandClass.Close
    
222:     Set rsCurrentLCType = Nothing
223:     Set rsCurrentLCClass = Nothing
224:     Set rsInsertLCType = Nothing
225:     Set rsPermLandClass = Nothing
226:     Set rsNewLandClass = Nothing
    
228:     g_strLCTypeName = strTempLCTypeName
229:     ReclassLanduse clsLUScenItems, strTempLCTypeName, m_strLCFileName
    
Exit Sub
ErrHandler:
233: MsgBox "Error Number: " & Err.Number & vbNewLine & "Error Description: " & Err.Description

End Sub

Private Function CopyCoefficient(strNewCoeffName As String, strCoeffSet As String) As ADODB.Recordset
On Error GoTo ErrHandler:
'General gist:  First we add new record to the Coefficient Set table using strNewCoeffName as
'the name, PollID, LCTYPEID.  Once that's done, we'll add the coefficients
'from the set being copied
    Dim strCopySet As String                'The Recordset of existing coefficients being copied
    Dim strNewLcType As String              'CmdString for inserting new coefficientset               '
    Dim strNewCoeff As String
    Dim strNewCoeffID As String             'Holder for the CoefficientSetID
    Dim intCoeffSetID As Integer
    Dim i As Integer

    Dim rsCopySet As New ADODB.Recordset        'First RS
    Dim rsCoeffSetID As New ADODB.Recordset '1 record recordset to get a hold of CoeffsetID
    Dim rsNewCoeff As New ADODB.Recordset

253:     strCopySet = "SELECT * FROM COEFFICIENTSET INNER JOIN COEFFICIENT ON COEFFICIENTSET.COEFFSETID = " & _
        "COEFFICIENT.COEFFSETID WHERE COEFFICIENTSET.NAME LIKE '" & strCoeffSet & "'"

256:     rsCopySet.Open strCopySet, g_ADOConn, adOpenKeyset

258:     Debug.Print rsCopySet.RecordCount
259:     Debug.Print rsCopySet!POLLID

    'INSERT: new Coefficient set taking the PollID and LCType ID from rsCopySet
262:     strNewLcType = "INSERT INTO COEFFICIENTSET(NAME, POLLID, LCTYPEID) VALUES ('" & _
                Replace(strNewCoeffName, "'", "''") & "'," & _
                rsCopySet!POLLID & "," & _
                m_intLCTypeID & ")"

    'First need to add the coefficient set to that table
266:     g_ADOConn.Execute strNewLcType, adAffectCurrent

    'Get the Coefficient Set ID of the newly created coefficient set
269:     strNewCoeffID = "SELECT COEFFSETID FROM COEFFICIENTSET " & _
        "WHERE COEFFICIENTSET.NAME LIKE '" & strNewCoeffName & "'"

272:     rsCoeffSetID.Open strNewCoeffID, g_ADOConn, adOpenKeyset, adLockPessimistic
273:     m_intCoeffSetID = rsCoeffSetID!CoeffSetID

    'Now loopy loo to populate values.
    'Get the coefficient table
    Dim strNewCoeff1 As String
278:     strNewCoeff1 = "SELECT * FROM COEFFICIENT"
279:     rsNewCoeff.Open strNewCoeff1, g_ADOConn, adOpenKeyset, adLockOptimistic

281:     rsCopySet.MoveFirst

    'Actually add the records to the new set
284:     For i = 1 To rsCopySet.RecordCount

        'Add New one
287:         rsNewCoeff.AddNew

        'Add the necessary components
290:         rsNewCoeff!Coeff1 = rsCopySet!Coeff1
291:         rsNewCoeff!Coeff2 = rsCopySet!Coeff2
292:         rsNewCoeff!Coeff3 = rsCopySet!Coeff3
293:         rsNewCoeff!Coeff4 = rsCopySet!Coeff4
294:         rsNewCoeff!CoeffSetID = m_intCoeffSetID
295:         rsNewCoeff!LCClassID = rsCopySet!LCClassID

297:         rsNewCoeff.Update
298:         rsCopySet.MoveNext

300:     Next i

302:     Set CopyCoefficient = rsNewCoeff

    'Cleanup
305:     rsCopySet.Close
306:     rsCoeffSetID.Close
    'rsNewCoeff.Close

309:     Set rsCopySet = Nothing
310:     Set rsCoeffSetID = Nothing
    'Set rsNewCoeff = Nothing

  Exit Function
ErrHandler:

316: MsgBox "Error Number: " & Err.Number & vbNewLine & "Error Description: " & Err.Description, vbCritical, "Error in modutil.copycoefficient"

End Function

Private Sub ReclassLanduse(clsLUScenItems As clsXMLLandUseItems, strLCClass As String, strLCFileName As String)
    'clsLUScenItems: which is a collection of the landuse entered
    'strLCClass: Name of the LCTYPE being altered
    'strLCFileName: path to which the landcover grid exists
    
    Dim strExpress As String
    Dim strExpression As String
    Dim strOutLandCover As String
    Dim pNewLandCoverRaster As IRaster
    Dim pTempDeleteRaster As IRaster
    
    Dim i As Integer
        
    Dim pMapAlgebraOp As IMapAlgebraOp
    Dim pArray As IArray
On Error GoTo ErrHandler:

    'init everything
338:     m_strLCClass = strLCClass
339:     Set pArray = New esriSystem.Array
340:     Set m_pEnv = New RasterAnalysis
        
    'Make sure the landcoverraster exists..it better if they get to this point, ED!
343:     If g_LandCoverRaster Is Nothing Then
344:         If modUtil.RasterExists(strLCFileName) Then
345:             Set m_pLandCoverRaster = modUtil.ReturnRaster(strLCFileName)
346:         Else
            Exit Sub
348:         End If
349:     Else
350:         Set m_pLandCoverRaster = g_LandCoverRaster
351:     End If
    
    'Get the rasterprops of the landcover raster
354:     Set m_pRasterProps = m_pLandCoverRaster
    
    'Get the workspace
357:     'Set m_pWSFactory = New RasterWorkspaceFactory
358:     'Set m_pWS = m_pWSFactory.OpenFromFile(modUtil.SplitWorkspaceName(strLCFileName), frmPrj.m_App.hWnd)
    'Set m_pWS = modUtil.SetRasterWorkspace(strWorkspace)
    'Set the environment
361:     With m_pEnv
362:         .SetCellSize esriRasterEnvValue, m_pRasterProps.MeanCellSize.X
363:         .SetExtent esriRasterEnvValue, m_pRasterProps.Extent
364:         Set .OutSpatialReference = m_pRasterProps.SpatialReference
365:         Set .OutWorkspace = m_pWS
366:     End With
    

    'Going to now take each entry in the landuse scenarios, if they're not ignored, go
    'to ReclassRaster to make a new raster of the Value indicated.  Then add the returned
    'raster to m_RasterArray.  We'll then make a map algebra statement out of that collection
    'of rasters to merge back into the original landcover raster...voila.
    
374:     If clsLUScenItems.Count > 0 Then
375:         For i = 1 To clsLUScenItems.Count
376:             If clsLUScenItems.Item(i).intApply = 1 Then
377:                 modProgDialog.ProgDialog "Processing Landuse scenario...", "Landuse Scenario", 0, CLng(clsLUScenItems.Count + 1), CLng(i), 0
378:                 If modProgDialog.g_boolCancel Then
379:                     pArray.Add ReclassRaster(clsLUScenItems.Item(i), m_strLCClass)
380:                 End If
                
382:             End If
383:         Next i
384:     End If
    
386:     Set pMapAlgebraOp = New RasterMapAlgebraOp
387:     Set m_pEnv = pMapAlgebraOp
    
389:     With m_pEnv
390:         .SetCellSize esriRasterEnvValue, m_pRasterProps.MeanCellSize.X
391:         .SetExtent esriRasterEnvValue, m_pRasterProps.Extent
392:         Set .OutSpatialReference = m_pRasterProps.SpatialReference
393:         Set .OutWorkspace = m_pWS
394:     End With
    'Now, if the raster array has some rasters in it, let's merge the fuckers together
396:     If pArray.Count > 0 Then
397:         For i = 0 To pArray.Count - 1
398:             pMapAlgebraOp.BindRaster pArray.element(i), "raster" & CStr(i)
399:             If strExpress = "" Then
400:                 strExpress = "Merge([" & "raster" & CStr(i) & "], "
401:             Else
402:                 strExpress = strExpress & "[" & "raster" & CStr(i) & "], "
403:             End If
404:         Next i
        
406:         pMapAlgebraOp.BindRaster m_pLandCoverRaster, "landcover"
407:         strExpression = strExpress & "[landcover])"
        
409:         modProgDialog.ProgDialog "Creating new landcover GRID...", "Landuse Scenario", 0, CLng(clsLUScenItems.Count + 1), CLng(i + 1), 0
        
411:         If modProgDialog.g_boolCancel Then
412:             Set pNewLandCoverRaster = pMapAlgebraOp.Execute(strExpression)
413:         End If
414:     Else
        Exit Sub
416:     End If
    
418:     strOutLandCover = modUtil.GetUniqueName("landcover", modUtil.SplitWorkspaceName(m_pEnv.OutWorkspace.PathName))
    

    
422:     If modUtil.MakePermanentRaster(pNewLandCoverRaster, m_pEnv.OutWorkspace.PathName, strOutLandCover) Then
423:         Set g_LandCoverRaster = pNewLandCoverRaster 'MsgBox "done"
424:     End If
    
    'Cleanup
427:     For i = 0 To pArray.Count - 1
428:         Set pTempDeleteRaster = pArray.element(i)
429:         Set pTempDeleteRaster = Nothing
430:     Next i
    
432:     modProgDialog.KillDialog
        
434:     Set clsLUScenItems = Nothing
435:     Set pMapAlgebraOp = Nothing
436:     Set m_pWS = Nothing
437:     'Set m_pWSFactory = Nothing
    
Exit Sub
ErrHandler:
441:     MsgBox "error in MSSetup " & Err.Number & ": " & Err.Description
         
End Sub

Private Function ReclassRaster(clsLUItem As clsXMLLandUseItem, strLCClass As String) As IRaster
  On Error GoTo ErrorHandler

'We're passing over a single land use scenario in the form of the xml
'class clsXMLLandUseItem, seems to be the easiest way to do this.
    Dim strSelect As String                         'ADO selections string
    Dim rsValue As ADODB.Recordset                  'ADO recordset
    Dim dblValue As Double                          'value
    Dim strOutName As String                        'String for poly's name
    Dim strExpression As String                     'Map calculator expression
    Dim clsLUItemDetails As New clsXMLLUScen        'The particulars in the landuse
    
    Dim pConversionOp As IConversionOp              'Conversion for poly to raster
    Dim pPolyFeatureClass As IFeatureClass          'The polygon featureclass
    Dim pPolyFeatLayer As ILayer                    'Polygon featurelayer
    Dim pPolyWS As IWorkspace
    Dim lngPolyFeatIndex As Long
    Dim pPolyDS As IGeoDataset
    Dim pOutPolyRDS As IRasterDataset
    Dim pMapAlgebraOp As IMapAlgebraOp
    Dim pPolyValueRaster As IRaster
    Dim pFinalPolyValueRaster As IRaster            'Final poly new raster value


    
    'STEP 1: Open the landclass Value Value -------------------------------------------------------------------------
    'This is the value user's landclass will change to
471:     strSelect = "SELECT LCTYPE.LCTYPEID, LCCLASS.NAME, LCCLASS.VALUE FROM " & _
        "LCTYPE INNER JOIN LCCLASS ON LCTYPE.LCTYPEID = LCCLASS.LCTYPEID " & _
        "WHERE LCTYPE.NAME LIKE '" & strLCClass & "' AND LCCLASS.NAME LIKE '" & clsLUItem.strLUScenName & "'"
    
475:     Set rsValue = New ADODB.Recordset
476:     rsValue.Open strSelect, g_ADOConn, adOpenKeyset
        
478:     dblValue = rsValue!Value
479:     rsValue.Close
    
    'init the landuse xml stuff
482:     clsLUItemDetails.XML = clsLUItem.strLUScenXMLFile
    
    'END STEP 1: ----------------------------------------------------------------------------------------------------
    
    'STEP 2: --------------------------------------------------------------------------------------------------------
    'Init
488:     Set pConversionOp = New RasterConversionOp
    'Set the spat Environment
490:     Set m_pEnv = pConversionOp
    
492:     With m_pEnv
493:         .SetCellSize esriRasterEnvValue, m_pRasterProps.MeanCellSize.X
494:         .SetExtent esriRasterEnvValue, m_pRasterProps.Extent
495:         Set .OutSpatialReference = m_pRasterProps.SpatialReference
496:         Set .OutWorkspace = m_pWS
497:     End With
    'END STEP 2: ----------------------------------------------------------------------------------------------------
    
    'STEP 3: --------------------------------------------------------------------------------------------------------
    'Convert the polygon featureclass into a Raster for sending back out
    'Get the featureclass, check for selected features
    
504:     If clsLUItemDetails.intLUScenSelectedPoly = 1 And m_pMap.SelectionCount > 0 Then
505:         lngPolyFeatIndex = modUtil.GetLayerIndex(clsLUItemDetails.strLUScenLyrName, frmPrj.m_App)
506:         Set pPolyFeatLayer = m_pMap.Layer(lngPolyFeatIndex)
             Set pPolyWS = modUtil.SetFeatureShapeWorkspace(modUtil.SplitWorkspaceName(clsLUItemDetails.strLUScenFileName))
507:         Set pPolyFeatureClass = modUtil.ExportSelectedFeatures(pPolyFeatLayer, m_pMap, pPolyWS, clsLUItemDetails.strLUScenFileName)
508:     Else
509:        Set pPolyFeatureClass = modUtil.ReturnFeatureClass(clsLUItemDetails.strLUScenFileName)
510:     End If
    'Have to use a GeoDataset, so QI
512:     Set pPolyDS = pPolyFeatureClass
    'Init the out dataset
514:     Set pOutPolyRDS = New RasterDataset
515:     strOutName = modUtil.GetUniqueName("poly", m_pWS.PathName)
516:     Set pOutPolyRDS = pConversionOp.ToRasterDataset(pPolyDS, "GRID", m_pWS, strOutName)
    'END STEP 3 ------------------------------------------------------------------------------------------------------
    
    'STEP 4: ---------------------------------------------------------------------------------------------------------
    'Need to now make that raster the value the user choose way back when
521:     Set pMapAlgebraOp = New RasterMapAlgebraOp
522:     Set m_pEnv = pMapAlgebraOp
    
524:     With m_pEnv
525:         .SetCellSize esriRasterEnvValue, m_pRasterProps.MeanCellSize.X
526:         .SetExtent esriRasterEnvValue, m_pRasterProps.Extent
527:         Set .OutSpatialReference = m_pRasterProps.SpatialReference
528:         Set .OutWorkspace = m_pWS
529:     End With
    
531:     Set pPolyValueRaster = pOutPolyRDS.CreateDefaultRaster
        
533:     pMapAlgebraOp.BindRaster pPolyValueRaster, "poly"
       
535:     strExpression = "Con([poly] <> -9999, " & dblValue & ", 0)"
    
    
538:     Set pFinalPolyValueRaster = pMapAlgebraOp.Execute(strExpression)
    'End STEP 4: -----------------------------------------------------------------------------------------------------
    
541:     Set ReclassRaster = pFinalPolyValueRaster
    
'Cleanup
544:     Set rsValue = Nothing
545:     Set pConversionOp = Nothing
546:     Set pPolyFeatureClass = Nothing
547:     Set pPolyDS = Nothing
548:     Set pOutPolyRDS = Nothing
549:     Set pMapAlgebraOp = Nothing
550:     Set pPolyValueRaster = Nothing
551:     Set pFinalPolyValueRaster = Nothing
552:     Set clsLUItemDetails = Nothing
    
Exit Function


    
    


  Exit Function
ErrorHandler:
  HandleError False, "ReclassRaster " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Function

Public Sub Cleanup(dictNames As Dictionary, clsPollItems As clsXMLPollutantItems, strLCTypeName As String)
  On Error GoTo ErrorHandler

    
    Dim strDeleteCoeffSet As String
    Dim strDeleteLCType As String
    Dim strCoeffDeleteName As String
    Dim strLCDeleteName As String
    Dim i As Integer
    

577:     strLCDeleteName = dictNames.Item(strLCTypeName)
    
579:     If Len(strLCDeleteName) = 0 Then
        Exit Sub
581:     End If
    
    Dim strLCTypeDelete As String
    Dim strLCClassDelete As String
    
    Dim rsLCTypeDelete As ADODB.Recordset
    
588:     strLCTypeDelete = "SELECT * FROM LCTYPE WHERE NAME LIKE '" & strLCDeleteName & "'"
    
590:     Set rsLCTypeDelete = New ADODB.Recordset
    
592:     rsLCTypeDelete.CursorLocation = adUseClient
593:     rsLCTypeDelete.Open strLCTypeDelete, modUtil.g_ADOConn, adOpenForwardOnly, adLockOptimistic
       
595:     strLCClassDelete = "Delete * FROM LCCLASS WHERE LCTYPEID =" & rsLCTypeDelete!LCTypeID
             
597:     modUtil.g_ADOConn.Execute strLCClassDelete
        
        'Set up a delete rs and get rid of it
600:     rsLCTypeDelete.Delete adAffectCurrent
601:     rsLCTypeDelete.Update
        
603:     For i = 1 To clsPollItems.Count
604:         strCoeffDeleteName = dictNames.Item(clsPollItems.Item(i).strCoeffSet)
605:         If Len(strCoeffDeleteName) > 0 Then
606:             strDeleteCoeffSet = "DELETE * FROM COEFFICIENTSET WHERE NAME LIKE '" & strCoeffDeleteName & "'"
607:             modUtil.g_ADOConn.Execute strDeleteCoeffSet
608:         End If
609:     Next i
    
611:     rsLCTypeDelete.Close

    
614:     Set rsLCTypeDelete = Nothing
    
    

    


  Exit Sub
ErrorHandler:
  HandleError True, "Cleanup " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub



