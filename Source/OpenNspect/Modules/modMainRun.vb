'********************************************************************************************************
'File Name: modMainRun.vb
'Description: Primary initialization of the model process
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

Imports System.Data.OleDb
Module modMainRun
    ' *************************************************************************************
    ' *  Perot Systems Government Services
    ' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
    ' *  modMainRun
    ' *************************************************************************************
    ' *  Description:  Sets the analysis enviroment and many of the 'global' rasters used
    ' *  througout the calculations.
    ' *
    ' *
    ' *  Called By: frmPrj::CmdOK
    ' *
    ' *  Confession:  OpenNSPECT started back in 2003 and has lived a long life.  Not sure how,
    ' *  but it has.  It is a cobbled together piece of code now going on 6 years of age.
    ' *  Whomever takes it over, I offer my apologies.  It's a bit overly complex, no doubt about
    ' *  that, but have fun.  On the bright side, modUtil contains nearly ever function you
    ' *  would ever need in the raster world of ArcGIS.
    ' *************************************************************************************

    'Following represent variables that will be used for all analysis
    'Garnered primarily from the DEM Dataset

    Public g_intDistanceUnits As Short 'Global Units 0 = meters, 1 = feet
    Public g_dblCellSize As Double 'Global Cell Size, again taken from DEM
    Public g_intPrecipType As Short 'Precipitation Type; I, IA, II, III
    Public g_booLocalEffects As Boolean 'Did they check local effects?
    Public g_booSelectedPolys As Boolean 'Did they select n polygons for limiting analysis?

    'The Public member datasets, to be used quite a bit
    Public g_pDEMRaster As MapWinGIS.Grid 'DEM Raster
    Public g_pFlowAccRaster As MapWinGIS.Grid 'Flow Accumulation
    Public g_pFlowDirRaster As MapWinGIS.Grid 'Flow Direction
    Public g_strFlowDirFilename As String 'Flow Direction file name
    Public g_strLSFileName As String 'LS file name
    Public g_pLSRaster As MapWinGIS.Grid 'LS Raster
    Public g_pWaterShedFeatClass As MapWinGIS.Shapefile 'WaterShed Poly featureclass
    Public g_KFactorRaster As MapWinGIS.Grid 'K Factor DS, used in RUSLE or MUSLE

    Public g_dicMetadata As Generic.Dictionary(Of String, String) 'Global dictionary to hold name of layer, metadata process string
    Public g_clsXMLPrjFile As clsXMLPrjFile 'Global xml project file for metadata support

    Public g_pSelectedPolyClip As MapWinGIS.Shape

    Public g_pGroupLayer As Integer = -1


    Public Sub SetGlobalEnvironment(ByRef cmdWShed As OleDbCommand, Optional ByVal SelectedPath As String = "", Optional ByRef SelectedShapes As Collections.Generic.List(Of Integer) = Nothing)
        'GOAL:  Set the analysis environment based on the properties of the DEM, and establish
        'the other contributing datasets: Watersheds, flow direction, flow accumulation, length/slope
        'Incoming Parameters:
        'rsWShed: the recordset of the selected ws delineation...used to get paths to datasets
        'strWorkspace: Workspace identified by user in main project window
        'pWShedlayer: selection layer, usually a watershed selection
        Dim dblCellSize As Double 'Cell Size
        Dim strError As String = "" 'Error Handling String
        Dim strDEM As String
        Dim strWS As String
        Dim strFlowDir As String
        Dim strFlowAcc As String
        Dim strLS As String
        Dim intDistUnits As Short
        Dim pMaskGeoDataset As MapWinGIS.Shapefile

        Dim dataWshed As OleDbDataReader = cmdWShed.ExecuteReader()
        dataWshed.Read()

        strDEM = dataWshed("FilledDEMFileName")
        strWS = dataWshed("wsfilename")
        strFlowDir = dataWshed("FlowDirFileName")
        strFlowAcc = dataWshed("FlowAccumFileName")
        strLS = dataWshed("LSFileName")
        intDistUnits = dataWshed("DEMGridUnits")


        'STEP 1: Get the workspaces set
        g_strFlowDirFilename = strFlowDir

        'STEP 2: Establish the environment
        If modUtil.RasterExists(strDEM) Then
            g_pDEMRaster = modUtil.ReturnRaster(strDEM)
        Else
            strError = "DEM Raster Does Not Exist: " & strDEM
        End If

        'STEP 6: Set the other Datasets
        'Begin with Water shed, let a featureclass
        g_pWaterShedFeatClass = New MapWinGIS.Shapefile
        g_pWaterShedFeatClass = modUtil.ReturnFeature(strWS)

        If g_booSelectedPolys Then
            pMaskGeoDataset = ReturnAnalysisMask(SelectedPath, SelectedShapes, strWS)
            g_MapWin.View.Extents = pMaskGeoDataset.Extents
        Else
            pMaskGeoDataset = Nothing
            g_MapWin.View.ZoomToMaxExtents()
        End If

        'STEP 3: With the Rasterdataset set, get its properties
        Dim pRasterProps As MapWinGIS.GridHeader = g_pDEMRaster.Header
        'Get cell size and envelope
        dblCellSize = pRasterProps.dX
        'Set the global cell size
        g_dblCellSize = dblCellSize

        'STEP 5: Set global units
        Select Case intDistUnits
            Case 0 'Meters
                g_intDistanceUnits = 0
            Case 1 'Feed
                g_intDistanceUnits = 1
        End Select

        'Flow Direction
        If modUtil.RasterExists(strFlowDir) Then
            g_pFlowDirRaster = modUtil.ReturnRaster(strFlowDir)
        Else
            strError = "Flow Direction Raster Does Not Exist: " & strFlowDir
        End If

        'FlowAccumulation
        If modUtil.RasterExists(strFlowAcc) Then
            g_pFlowAccRaster = modUtil.ReturnRaster(strFlowAcc)
        Else
            strError = "Flow Accumulation Raster Does Not Exist: " & strFlowAcc
        End If

        'Length Slope
        If modUtil.RasterExists(strLS) Then
            g_pLSRaster = modUtil.ReturnRaster(strLS)
            g_strLSFileName = strLS
        Else
            strError = "Length Slope raster does not Exist: " & strLS
        End If


        If Len(strError) > 0 Then
            MsgBox(strError, MsgBoxStyle.Critical, "Missing Data")
        End If

    End Sub


    Private Function ReturnAnalysisMask(ByVal SelectedPath As String, ByRef SelectedShapes As Collections.Generic.List(Of Integer), ByRef strBasinFeatClass As String) As MapWinGIS.Shapefile
        'Incoming
        'pLayer: Layer user has chosen as being the one from which the selected polys will come
        'pMap: current map
        'pWorkspace:  place to put the exported selected sheds
        'strBasinFeatClass: string file location of BasinPoly.shp
        ReturnAnalysisMask = Nothing

        g_strSelectedExportPath = modUtil.ExportSelectedFeatures(SelectedPath, SelectedShapes)
        g_pSelectedPolyClip = ReturnSelectGeometry(g_strSelectedExportPath)
        Dim sfSelected As MapWinGIS.Shapefile = modUtil.ReturnFeature(g_strSelectedExportPath)

        'ARA 12/5/2010 Since this is purely used for zoom, intersecting with the basins is kind of pointless and ExportShapesWithPolygons isn't working anyways, so just returning the extents of the selection area.
        Return sfSelected


        ''Make a call to get the BasinPoly featureclass using the name sent over
        'Dim basinSF As MapWinGIS.Shapefile = modUtil.ReturnFeature(strBasinFeatClass)

        ''Get unique name for output
        'Dim strOutPath As String = modUtil.GetUniqueName("selshed", g_strWorkspace, ".shp")
        'Dim sfOut As New MapWinGIS.Shapefile

        ''Intersect watersheds by selection shapefile
        'If MapWinGeoProc.Selection.ExportShapesWithPolygons(basinSF, sfSelected, sfOut) Then
        '    sfOut.SaveAs(strOutPath)
        'End If

        'Return sfOut
    End Function


    Public Function CheckMultiPartPolygon(ByVal pPolygon As MapWinGIS.Shape) As Boolean
        If pPolygon.NumParts > 1 Then
            CheckMultiPartPolygon = True
        Else
            CheckMultiPartPolygon = False
        End If
    End Function


    Public Function ReturnSelectGeometry(ByVal strInputSF As String) As MapWinGIS.Shape
        ReturnSelectGeometry = Nothing

        Dim sfSelected As MapWinGIS.Shapefile = modUtil.ReturnFeature(strInputSF)
        If Not sfSelected Is Nothing Then
            Dim unionShape As MapWinGIS.Shape = sfSelected.Shape(0)
            For i As Integer = 1 To sfSelected.NumShapes - 1
                unionShape = MapWinGeoProc.SpatialOperations.Union(unionShape, sfSelected.Shape(i))
            Next
            sfSelected.Close()
            Return unionShape
        End If
    End Function

End Module