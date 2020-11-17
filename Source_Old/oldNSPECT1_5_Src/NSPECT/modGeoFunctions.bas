Attribute VB_Name = "modGeoFunctions"
' *************************************************************************************************
' *  Perot Systems Government Services
' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
' *  modGeoFunctions
' *************************************************************************************************
' *  Description: Form for browsing and maintaining watershed delineation
' *  Scenarios within NSPECT
' *  Called By:  frmWSDelin menu item New...
' *********************************************************************************
' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "C:\Projects\HI\code\060203\modGeoFunctions.bas"

Public Function SetNewDefaultEnvironment(pGeoGrid As IGeoDataset) As IRasterAnalysisEnvironment
On Error GoTo ErrorHandler
                       
'    If Not TypeOf pGeoGrid Is IRaster Then
'        GoTo errorHandler
'    End If
'
    Dim pEnv As IRasterAnalysisEnvironment
    Set pEnv = New RasterAnalysis
    
    Dim pRaster As IRaster
    Dim pRasterDataset As IRasterDataset
    
    Set pRasterDataset = pGeoGrid
    Set pRaster = pRasterDataset.CreateDefaultRaster
    
    
    'Get the cell size
    Dim pRProp As IRasterProps
    Set pRProp = pRaster
    Dim tRange
    tRange = pRProp.MeanCellSize.X
    'Set the cell size
    If tRange <> 0 Then pEnv.SetCellSize esriRasterEnvValue, tRange
    
    'Set the out put workspace
   'Dim pRaster As IRaster
    Dim pWS As IWorkspace
    Dim pDataset As IDataset
    
    Dim pRasterBandCollection As IRasterBandCollection
    Set pRasterBandCollection = pRProp
    Dim pRasterBand As IRasterBand
    Set pRasterBand = pRasterBandCollection.Item(0)
    Set pDataset = pRasterBand
    Set pWS = pDataset.Workspace
    If Not pWS Is Nothing Then Set pEnv.OutWorkspace = pWS
    
    'Set the extent
    If Not pRProp.Extent Is Nothing Then pEnv.SetExtent esriRasterEnvValue, pRProp.Extent
    'Set the spatial reference
    If Not pRProp.SpatialReference Is Nothing Then Set pEnv.OutSpatialReference = pRProp.SpatialReference
    
    ' Set it as the default settings
    pEnv.SetAsNewDefaultEnvironment
    
    ' Return reference to the default environment setting
    Set SetNewDefaultEnvironment = pEnv
    
    Exit Function
ErrorHandler:
    Set SetNewDefaultEnvironment = Nothing
End Function

Public Function DEMAgree(pDEMGDSin As IGeoDataset, pStreamFeatClass As IFeatureClass, _
                          intBuffDistance As Integer, intSmoothDist As Integer, _
                          intSharpDist As Integer, pInWS As IWorkspace) As IGeoDataset
  On Error GoTo ErrorHandler

'
''*****************************************************************
'' Copy Right, 2001 ESRI, All Rights Reserved. Revised 10/10
''*****************************************************************
'' Description:   Generates a reconditioned DEM for an input (DEM) GRID.
''                This script implements the Agree method developed by
''                F.Hellweger at the University of Texas at Austin in 1997.
''                Summary:  The method replaces the DEM in the buffer zone
''                          of the vector theme, with a DEM that slopes
''                          towards the vector theme.  The vector theme is
''                          burned into the DEM.
''                Required information:
''                    1) Original DEM.
''                    2) Vector theme representing streams (burning) or
''                        boundaries (fencing).
''                    3) Buffer distance (cells).  Width around the vector theme
''                        for which the DEM will be "reconditioned".
''                        Outside the buffer the elevations are the same as
''                        for the original DEM.
''                    4) Smooth distance (elevation).  Depth at the vector theme
''                        (change from the existing DEM at that location) that
''                        will be used for smooting within the buffer zone.
''                    5) Sharp distance (elevation).  Depth at the vector theme
''                        (change from the smooth DEM at that location) that
''                        will be used for deepening of the smooth DEM (further
''                        burning).
''*****************************************************************
'
''Step 1: Compute the vector grid
''Step 2: Compute the smooth drop/raise grid
''Step 3: Compute the vector distance grid
''Step 4: Compute the vector Allocation grid"
''Step 5: Compute the elevation grid outside of the buffer (Elevation Doughnut grid)
''Step 6: Compute the distance grid for the Elevation Doughnut grid
''Step 7: Compute the Allocation grid for the Elevation Doughnut grid
''Step 8: Compute the smooth interpolated elevation grid
''Step 9: Compute the sharp drop/raise grid
''Step 10:Compute the modified elevation grid
''Step 11: If there are negative elevations, offer to raise the DEM


    'Get the RawDEM Grid

    Dim pDEMRaster As IRaster
    Dim pRasterDataset As IRasterDataset
    
    Set pRasterDataset = pDEMGDSin
    Set pDEMRaster = pRasterDataset.CreateDefaultRaster
    
    Dim pAgreeRivFC As IFeatureClass
    Set pAgreeRivFC = pStreamFeatClass

    Dim pWS As IWorkspace
    Set pWS = pInWS

    
    LoadSALicense

    'Get the cell size
    Dim pRProp As IRasterProps
    Set pRProp = pDEMRaster
    Dim tRange
    tRange = pRProp.MeanCellSize.X
    bInteger = pRProp.IsInteger
    'Set the Analysis Environment as that of the Input Grid
    'Set this Environment as the deafult for all the Ops to be created later
    'Call pEnv.Reset to reset the Environment back to the original
    
    Dim pEnv As IRasterAnalysisEnvironment
    Set pEnv = SetNewDefaultEnvironment(pDEMGDSin)
    
    Dim pConversionOp As IConversionOp
    Set pConversionOp = New RasterConversionOp

    Dim pIrastermakerop  As IRasterMakerOp
    Set pIrastermakerop = New RasterMakerOp
    
''Step 1: Compute the vector grid
    Dim pStreamRasterDataset As IRasterDataset
    Dim pAgreeRivGeoDataset As IGeoDataset
    Set pAgreeRivGeoDataset = pAgreeRivFC

    Set pStreamRasterDataset = pConversionOp.ToRasterDataset(pAgreeRivGeoDataset, "GRID", pWS, "rivgrid")
    Dim pStreamRaster As IRaster
    Set pStreamRaster = pStreamRasterDataset.CreateDefaultRaster
    Dim pGeoStream As IGeoDataset
    Set pGeoStream = pStreamRaster

    Dim dBuffer As Double
    dBuffer = CDbl(intBuffDistance) * tRange - tRange / 2

    'Make a Constant Grid for the buffer
    Dim pGeoBuffer As IGeoDataset
    Set pGeoBuffer = pIrastermakerop.MakeConstant(CDbl(dBuffer), False)

    'Make a Constant Grid for the Smooth Drop
    Dim pGeoSmoothDrop As IGeoDataset
    Set pGeoSmoothDrop = pIrastermakerop.MakeConstant(CDbl(intSmoothDist), False)

    'Make a Constant Grid for the Sharp Drop
    Dim pGeoSharpDrop As IGeoDataset
    Set pGeoSharpDrop = pIrastermakerop.MakeConstant(CDbl(intSharpDist), False)

    Dim pCondRaster As IGeoDataset
    Dim pLogicalOp As ILogicalOp
    Set pLogicalOp = New RasterMathOps
    Dim pMathOp As IMathOp
    Set pMathOp = New RasterMathOps

    ''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Added to take care of the Zero and negative DEMS
    'Lift the DEM by Abs(Minimum) + Smooth Drop + Sharp Drop
    'And drop it back by the same amount after done
    'Get the minimum from raster statistics
    Dim pRasterBandCollection As IRasterBandCollection
    Dim dMin As Double
    Set pRasterBandCollection = pRProp
    Dim pRasterBand As IRasterBand
    Set pRasterBand = pRasterBandCollection.Item(0)
    Dim pRasterStats As IRasterStatistics
    Set pRasterStats = pRasterBand.Statistics
    dMin = pRasterStats.Maximum
    If dMin <= 0 Then
      bNegative = True
      dMin = Abs(dMin) + 1
      Dim pGeoTotalDrop As IGeoDataset
      Set pGeoTotalDrop = pMathOp.Plus(pGeoSmoothDrop, pGeoSharpDrop)
      Set pGeoTotalDrop = pMathOp.Plus(pGeoTotalDrop, pIrastermakerop.MakeConstant(dMin, False))
      Set pDEMGDSin = pMathOp.Plus(pDEMGDSin, pGeoTotalDrop)
    End If
    
    Set pCondRaster = pLogicalOp.IsNull(pGeoStream)
    Dim pConditionalOp As IConditionalOp
    Set pConditionalOp = New RasterConditionalOp

'Step 2: Compute the smooth drop/raise grid
    Dim pGeoSmoothDropFromDEM As IGeoDataset
    Set pGeoSmoothDropFromDEM = pMathOp.Minus(pDEMGDSin, pGeoSmoothDrop)
    Dim pGeoSmooth As IGeoDataset
    Set pGeoSmooth = pConditionalOp.SetNull(pCondRaster, pGeoSmoothDropFromDEM)

'Step 3: Compute the vector distance grid
    Dim pDistOP As IDistanceOp
    Set pDistOP = New RasterDistanceOp

    Dim pDistOp2 As IDistanceOp
    Set pDistOp2 = New RasterDistanceOp

    Dim pGeoVectEucDist As IGeoDataset
    Set pGeoVectEucDist = pDistOP.EucDistance(pGeoSmooth, Nothing, Nothing)

'Step 4: Compute the vector Allocation grid"
    Set pGeoSmooth = pMathOp.Int(pGeoSmooth)
    Dim pGeoVectEucAlloc As IGeoDataset
    Set pGeoVectEucAlloc = pDistOp2.EucAllocation(pGeoSmooth, Nothing, Nothing)

'Step 5: Compute the elevation grid outside of the buffer (Elevation Doughnut grid)
    'Get the Condition Raster where the vect Dist is greater than BufferGrid"
    Set pCondRaster = pLogicalOp.GreaterThan(pGeoVectEucDist, pGeoBuffer)
    Dim pGeoFalse As IGeoDataset
    Set pGeoFalse = pConditionalOp.SetNull(pDEMGDSin, pDEMGDSin)
    Dim pGeoBuffElv As IGeoDataset
    Set pGeoBuffElv = pConditionalOp.Con(pCondRaster, pDEMGDSin, Nothing)

'Step 6: Compute the distance grid for the Elevation Doughnut grid
    
    Dim pGeoBuffEucDist As IGeoDataset
    Set pGeoBuffElv = pMathOp.Int(pGeoBuffElv)
    Set pGeoBuffEucDist = pDistOP.EucDistance(pGeoBuffElv, Nothing, Nothing)
    
    LoadSALicense
    
'Step 7: Compute the Allocation grid for the Elevation Doughnut grid
    Dim pGeoBuffEucAlloc As IGeoDataset
    Set pGeoBuffEucAlloc = pDistOp2.EucAllocation(pGeoBuffElv, Nothing, Nothing)

'Step 8: Compute the smooth interpolated elevation grid
    'Avenue logic: smoelev = gOrigDEM.Con((vectallo + (((bufallo - vectallo) / (bufdist + vectdist)) * vectdist)), gOrigDEM)
    Dim pGeoBuffAlloMinusVectAllo As IGeoDataset
    Set pGeoBuffAlloMinusVectAllo = pMathOp.Minus(pGeoBuffEucAlloc, pGeoVectEucAlloc)
    Dim pGeoBuffDistPlusVectDist As IGeoDataset
    Set pGeoBuffDistPlusVectDist = pMathOp.Plus(pGeoBuffEucDist, pGeoVectEucDist)
    Dim pGeoSmoothMod As IGeoDataset
    Set pGeoSmoothMod = pMathOp.Plus(pGeoVectEucAlloc, pMathOp.Times(pMathOp.Divide(pGeoBuffAlloMinusVectAllo, pGeoBuffDistPlusVectDist), pGeoVectEucDist))
    Dim pGeoSmoothDEM As IGeoDataset
    Set pCondRaster = pLogicalOp.IsNull(pDEMGDSin)
    Set pGeoSmoothDEM = pConditionalOp.Con(pCondRaster, pDEMGDSin, pGeoSmoothMod)

   

'Step 9: Compute the sharp drop/raise grid
    'Avenue logic: shagrid = (vectgrid.IsNull).SetNull(smoelev - nSharpDrop.AsGrid).Int
    Dim pGeoSharpMod As IGeoDataset
    Set pCondRaster = pLogicalOp.IsNull(pGeoStream)
    
    'pStatusBar.Message(0) = "Computing the sharp drop/raise grid .."
    Set pGeoSharpMod = pConditionalOp.SetNull(pCondRaster, pMathOp.Minus(pGeoSmoothDEM, pGeoSharpDrop))

'Step 10:   Compute the modified elevation grid
    'Avenue logic: elevgrid = (vectgrid.IsNull).Con(smoelev, shagrid)
    Dim pGeoAgree As IGeoDataset
    Set pGeoAgree = pConditionalOp.Con(pCondRaster, pGeoSmoothDEM, pGeoSharpMod)
    
    'If the RawDEM had 0 or negative values, then drop the grid by the equivalent Raise
    If bNegative Then
      Set pGeoAgree = pMathOp.Minus(pGeoAgree, pGeoTotalDrop)
      Set pDEMGDSin = pMathOp.Minus(pGeoAgree, pGeoTotalDrop)
    End If

    'Comoute the final Agree DEM
    Set pCondRaster = pLogicalOp.IsNull(pGeoBuffElv)
    Set pGeoAgree = pConditionalOp.Con(pCondRaster, pGeoAgree, pDEMGDSin)
    Set pRProp = pGeoAgree

    'Get the minimum from raster statistics
    Set pRasterBandCollection = pRProp
    Set pRasterBand = pRasterBandCollection.Item(0)
    Set pRasterStats = pRasterBand.Statistics
    dMin = pRasterStats.Minimum

    Dim Response

'Step 11: If there are negative elevations, offer to raise the DEM

    If (dMin < 0) Then
        ' we have negative "elevations"
        strMsg = "The resulting Agree DEM contains negative elevation values."
        strMsg = strMsg & vbLf & "This will cause problems if filling sinks."
        strMsg = strMsg & vbLf & "Do you want to raise the DEM so there are no negative elevations?"
        Response = MsgBox(strMsg, vbYesNo, sTitle)
        
        Dim pRasterMakerOp As IRasterMakerOp
        Set pRasterMakerOp = New RasterMakerOp
        
        If (Response = vbYes) Then
            
            ' raise the dem by the largest negative value plus about a 10
            dMin = Abs((Int(dMin) - 10))
            LoadSALicense
            Dim pGeoRaise As IGeoDataset
            Set pGeoRaise = pMathOp.Plus(pGeoAgree, pRasterMakerOp.MakeConstant(dMin, False))
            Set pGeoAgree = pGeoRaise
            
            strMsg = "The Agree DEM was raised " & dMin & " vertical units."
            strMsg = strMsg & vbLf & "Remember to use the original DEM for parameter extraction!"
            MsgBox strMsg, vbOKOnly, sTitle
        
        Else
            
            strMsg = "The Agree DEM was not raised!"
            strMsg = strMsg & vbLf & "Remember to eliminate the negative values before using this DEM for filling!"
            MsgBox strMsg, vbOKOnly, sTitle
        
        End If
    End If


    'If the original DEM is integer, then int the Agree
    If bInteger Then
        Set pGeoAgree = pMathOp.Int(pGeoAgree)
        strMsg = "The AgreeDEM will be generated as integer"
        strMsg = strMsg & vbLf & "You may lose some of the smoothening depending upon the buffer/drop ratio!"
        MsgBox strMsg, vbOKOnly, "Data Warning"
    End If
    
    Set DEMAgree = pGeoAgree
    

Exit Function

ErrorHandler:
  HandleError True, "DEMAgree " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4


End Function

Private Sub LoadSALicense()
    ' Get Spatial Analyst Extension UID
    Dim pUID As New UID
    pUID.Value = "esriCore.SAExtension.1"
    ' Add Spatial Analyst extension to the license manager
    Dim v As Variant
    Dim pLicAdmin As IExtensionManagerAdmin
    Set pLicAdmin = New ExtensionManager
    Call pLicAdmin.AddExtension(pUID, v)
    
    ' Enable the license
    Dim pLicManager As IExtensionManager
    Set pLicManager = pLicAdmin
    Dim pExtensionConfig As IExtensionConfig
    Set pExtensionConfig = pLicManager.FindExtension(pUID)
    If Not pExtensionConfig.State = esriESUnavailable Then
        pExtensionConfig.State = esriESEnabled
    End If
End Sub

