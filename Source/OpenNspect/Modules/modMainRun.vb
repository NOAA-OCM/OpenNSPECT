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
Imports System.Collections.Generic
Imports MapWinGeoProc
Imports MapWinGIS
Imports System.Data.OleDb
Imports OpenNspect.Xml

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

    Public g_intDistanceUnits As Short
    'Global Units 0 = meters, 1 = feet
    Public g_dblCellSize As Double
    'Global Cell Size, again taken from DEM
    Public g_intPrecipType As Short
    'Precipitation Type; I, IA, II, III
    Public g_booLocalEffects As Boolean
    'Did they check local effects?
    Public g_booSelectedPolys As Boolean
    'Did they select n polygons for limiting analysis?

    'The Public member datasets, to be used quite a bit
    Public g_pDEMRaster As Grid
    'DEM Raster
    Public g_pFlowAccRaster As Grid
    'Flow Accumulation
    Public g_pFlowDirRaster As Grid
    'Flow Direction
    Public g_strFlowDirFilename As String
    'Flow Direction file name
    Public g_strLSFileName As String
    'LS file name
    Public g_pLSRaster As Grid
    'LS Raster
    Public g_pWaterShedFeatClass As Shapefile
    'WaterShed Poly featureclass
    Public g_KFactorRaster As Grid
    'K Factor DS, used in RUSLE or MUSLE

    Public g_dicMetadata As Dictionary(Of String, String)
    'Global dictionary to hold name of layer, metadata process string
    Public g_XmlPrjFile As ProjectFile
    'Global xml project file for metadata support

    Public g_pSelectedPolyClip As Shape

    Public g_pGroupLayer As Integer = -1

    ''' <summary>
    '''Set the analysis environment based on the properties of the DEM, and establish
    '''the other contributing datasets: Watersheds, flow direction, flow accumulation, length/slope
    ''' </summary>
    ''' <param name="cmdWShed">the recordset of the selected ws delineation...used to get paths to datasets.</param>
    ''' <param name="SelectedPath">selection layer, usually a watershed selection.</param>
    ''' <param name="SelectedShapes">The selected shapes.</param>
    Public Sub SetGlobalEnvironment(ByRef cmdWShed As OleDbCommand, Optional ByVal SelectedPath As String = "", _
                                     Optional ByRef SelectedShapes As List(Of Integer) = Nothing)

        Dim strError As String
        Dim dataWshed As OleDbDataReader = cmdWShed.ExecuteReader()
        dataWshed.Read()

        Dim strDEM As String = dataWshed("FilledDEMFileName")
        Dim strWS As String = dataWshed("wsfilename")
        Dim strFlowDir As String = dataWshed("FlowDirFileName")
        Dim strFlowAcc As String = dataWshed("FlowAccumFileName")
        Dim strLS As String = dataWshed("LSFileName")
        Dim intDistUnits As Short = dataWshed("DEMGridUnits")

        dataWshed.Close()

        'STEP 1: Get the workspaces set
        g_strFlowDirFilename = strFlowDir

        'STEP 2: Establish the environment
        If RasterExists(strDEM) Then
            g_pDEMRaster = ReturnRaster(strDEM)
        Else
            strError = "DEM Raster Does Not Exist: " & strDEM
            MsgBox(strError, MsgBoxStyle.Critical, "Missing Data")
        End If

        'STEP 6: Set the other Datasets
        'Begin with Water shed, let a featureclass
        g_pWaterShedFeatClass = ReturnFeature(strWS)

        Dim pMaskGeoDataset As Shapefile
        If g_booSelectedPolys Then
            pMaskGeoDataset = ReturnAnalysisMask(SelectedPath, SelectedShapes, strWS)
            MapWindowPlugin.MapWindowInstance.View.Extents = pMaskGeoDataset.Extents
        Else
            pMaskGeoDataset = Nothing
            MapWindowPlugin.MapWindowInstance.View.ZoomToMaxExtents()
        End If

        'STEP 3: With the Rasterdataset set, get its properties
        Dim pRasterProps As GridHeader = g_pDEMRaster.Header
        'Set the global cell size
        g_dblCellSize = pRasterProps.dX

        'STEP 5: Set global units
        Select Case intDistUnits
            Case 0 'Meters
                g_intDistanceUnits = 0
            Case 1 'Feed
                g_intDistanceUnits = 1
        End Select

        'Flow Direction
        If RasterExists(strFlowDir) Then
            g_pFlowDirRaster = ReturnRaster(strFlowDir)
        Else
            strError = "Flow Direction Raster Does Not Exist: " & strFlowDir
            MsgBox(strError, MsgBoxStyle.Critical, "Missing Data")
        End If

        'FlowAccumulation
        If RasterExists(strFlowAcc) Then
            g_pFlowAccRaster = ReturnRaster(strFlowAcc)
        Else
            strError = "Flow Accumulation Raster Does Not Exist: " & strFlowAcc
            MsgBox(strError, MsgBoxStyle.Critical, "Missing Data")
        End If

        'Length Slope
        If RasterExists(strLS) Then
            g_pLSRaster = ReturnRaster(strLS)
            g_strLSFileName = strLS
        Else
            strError = "Length Slope raster does not Exist: " & strLS
            MsgBox(strError, MsgBoxStyle.Critical, "Missing Data")
        End If

    End Sub

    Private Function ReturnAnalysisMask(ByVal SelectedPath As String, _
                                         ByRef SelectedShapes As List(Of Integer), _
                                         ByRef strBasinFeatClass As String) As Shapefile
        'Incoming
        'pLayer: Layer user has chosen as being the one from which the selected polys will come
        'pMap: current map
        'pWorkspace:  place to put the exported selected sheds
        'strBasinFeatClass: string file location of BasinPoly.shp
        ReturnAnalysisMask = Nothing

        g_strSelectedExportPath = ExportSelectedFeatures(SelectedPath, SelectedShapes)
        g_pSelectedPolyClip = ReturnSelectGeometry(g_strSelectedExportPath)
        Dim sfSelected As Shapefile = ReturnFeature(g_strSelectedExportPath)

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

    Public Function CheckMultiPartPolygon(ByVal pPolygon As Shape) As Boolean
        If pPolygon.NumParts > 1 Then
            CheckMultiPartPolygon = True
        Else
            CheckMultiPartPolygon = False
        End If
    End Function

    Public Function ReturnSelectGeometry(ByVal strInputSF As String) As Shape
        ReturnSelectGeometry = Nothing

        Dim sfSelected As Shapefile = ReturnFeature(strInputSF)
        If Not sfSelected Is Nothing Then
            Dim unionShape As Shape = sfSelected.Shape(0)
            For i As Integer = 1 To sfSelected.NumShapes - 1
                unionShape = SpatialOperations.Union(unionShape, sfSelected.Shape(i))
            Next
            sfSelected.Close()
            Return unionShape
        End If
    End Function
End Module