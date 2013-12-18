'********************************************************************************************************
'File Name: modPollutantCalcs.vb
'Description: Functions handling the Pollutant calculations for the model
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
Imports MapWinGeoProc
Imports MapWinGIS
Imports OpenNspect.Xml


Module Pollutants
    ' *************************************************************************************
    ' *  Perot Systems Government Services
    ' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
    ' *  modPollutantCalcs
    ' *************************************************************************************
    ' *  Description:  Callculation of pollutant concentration
    ' *     * Sub PollutantConcentrationSetup: called from frmPrj, sort of the main sub in
    ' *       this module.  Uses a XmlPollutantItem and the landclass to get things started.
    ' *       Calls: ConstructConStatement, CalcPollutantConcentration
    ' *     * Function ConstructConStatment: Constructs the initial Con statement in the pollutant
    ' *       concentration calculations.
    ' *     * Sub CalcPollutantConcentration: The big workhorse. Contains all the map algebra that
    ' *       gets this turkey finished
    ' *************************************************************************************

    Private _PollutantName As String
    Private _WaterQualityStandardName As String
    Private _PollutantColor As String
    Private _PollutantCoeffMetadata As String 'Variable to hold coeffs for use in metadata
    Private _WQValue As Single
    Private _FlowMax As Single
    Private _picks() As String
    Private _picksArray(1, 1) As String


    ''' <summary>
    ''' Gets the name of the coefficient set; could be a temporary one due to landuses
    ''' </summary>
    ''' <param name="Pollutant">The pollutant.</param><returns></returns>
    Private Function GetCoefficientSetName(ByRef Pollutant As PollutantItem) As String
        If g_LandUse_DictTempNames.ContainsKey(Pollutant.strCoeffSet) AndAlso Len(g_LandUse_DictTempNames.Item(Pollutant.strCoeffSet)) > 0 Then
            Return g_LandUse_DictTempNames.Item(Pollutant.strCoeffSet)
        End If

        Return Pollutant.strCoeffSet
    End Function
    Public Function PollutantConcentrationSetup(ByRef Pollutant As PollutantItem, ByRef strWQName As String, ByRef OutputItems As OutputItems) As Boolean
        'Sub takes incoming parameters (in the form of a pollutant item) from the project file
        'and then parses them out
        Try
            'Open Strings
            Dim strType As String
            Dim strField As String = ""
            Dim concentrationStatement As String = ""
            Dim coeffTypes() As String = {"Coeff1", "Coeff2", "Coeff3", "Coeff4"}
            Dim concentrationStateArray() As String = {"", "", "", ""}
            'Again, because of landuse, we have to check for 'temp' coeff sets and their use
            'Get the name of the pollutant
            _PollutantName = Pollutant.strPollName

            'Get the name of the Water Quality Standard
            _WaterQualityStandardName = strWQName

            'Figure out what coeff user wants
            Select Case Pollutant.strCoeff
                Case "Type 1"
                    strField = "Coeff1"
                Case "Type 2"
                    strField = "Coeff2"
                Case "Type 3"
                    strField = "Coeff3"
                Case "Type 4"
                    strField = "Coeff4"
                Case "Use shapefile..."
                    strField = "Pick"
                Case ""
            End Select


            If Len(strField) > 0 Then
                Using cmdPoll As New DataHelper(String.Format("SELECT * FROM COEFFICIENTSET WHERE NAME LIKE '{0}'", GetCoefficientSetName(Pollutant)))
                    Using dataPoll As OleDbDataReader = cmdPoll.ExecuteReader()
                        dataPoll.Read()
                        If strField = "Pick" Then
                             'concentrationStateArray(4) = "foo"
                            concentrationStatement = ""
                            For icoeff As Integer = 0 To 3
                                'For Each coeffType As String In coeffTypes
                                ' strType = String.Format("SELECT LCCLASS.Value, LCCLASS.Name, COEFFICIENT.{0} As CoeffType, COEFFICIENT.CoeffID, COEFFICIENT.LCCLASSID FROM LCCLASS LEFT OUTER JOIN COEFFICIENT ON LCCLASS.LCCLASSID = COEFFICIENT.LCCLASSID WHERE COEFFICIENT.COEFFSETID = {1} ORDER BY LCCLASS.VALUE", coeffTypes(icoeff), dataPoll("CoeffSetID"))
                                strType = String.Format("SELECT LCCLASS.Value, LCCLASS.Name, COEFFICIENT.{0} As CoeffType, COEFFICIENT.CoeffID, COEFFICIENT.LCCLASSID FROM LCCLASS INNER JOIN COEFFICIENT ON LCCLASS.LCCLASSID = COEFFICIENT.LCCLASSID WHERE COEFFICIENT.COEFFSETID = {1} ORDER BY LCCLASS.VALUE", coeffTypes(icoeff), dataPoll("CoeffSetID"))
                                Debug.WriteLine(strType)
                                Using cmdType As New DataHelper(strType)
                                    Dim command As OleDbCommand = cmdType.GetCommand()
                                    'concentrationStatement = concentrationStatement & ConstructPickStatmentUsingLandClass(command, g_LandCoverRaster, "CoeffType") ' Coefficient list picked here: array with 1 coeff/LC class
                                    concentrationStateArray(icoeff) = concentrationStatement & ConstructPickStatmentUsingLandClass(command, g_LandCoverRaster, "CoeffType") ' Coefficient list picked here: array with 1 coeff/LC class
                                    _PollutantCoeffMetadata = ConstructMetaData(command, (Pollutant.strCoeff), g_Project.IncludeLocalEffects)
                                End Using
                            Next
                            '                          _PollutantCoeffMetadata = ConstructMetaData(cmdType.GetCommand(), (Pollutant.strCoeff), g_Project.IncludeLocalEffects)
                        Else
                            strType = String.Format("SELECT LCCLASS.Value, LCCLASS.Name, COEFFICIENT.{0} As CoeffType, COEFFICIENT.CoeffID, COEFFICIENT.LCCLASSID FROM LCCLASS LEFT OUTER JOIN COEFFICIENT ON LCCLASS.LCCLASSID = COEFFICIENT.LCCLASSID WHERE COEFFICIENT.COEFFSETID = {1} ORDER BY LCCLASS.VALUE", strField, dataPoll("CoeffSetID"))
                            Using cmdType As New DataHelper(strType)
                                Dim command As OleDbCommand = cmdType.GetCommand()
                                concentrationStatement = ConstructPickStatmentUsingLandClass(command, g_LandCoverRaster, "CoeffType") ' Coefficient list picked here: array with 1 coeff/LC class
                                _PollutantCoeffMetadata = ConstructMetaData(command, (Pollutant.strCoeff), g_Project.IncludeLocalEffects)
                            End Using
                        End If
                    End Using
                End Using

            End If

            'Find out the color of the pollutant
            Using cmdPollColor As New DataHelper("Select Color from Pollutant where NAME LIKE '" & _PollutantName & "'")
                Using datapollcolor As OleDbDataReader = cmdPollColor.ExecuteReader()
                    datapollcolor.Read()
                    _PollutantColor = CStr(datapollcolor("Color"))
                End Using
            End Using

            If strField = "Pick" Then
                'MsgBox("Sorry, this doesn't work at present.  However, coefficients are:" & vbCrLf & _
                'concentrationStateArray(0) & vbCrLf & _
                'concentrationStateArray(1) & vbCrLf & _
                'concentrationStateArray(2) & vbCrLf & _
                'concentrationStateArray(3))

                'TODO TO DO: Fix this to read shapefile and field names from xml file
                'concentrationStatement = "Pick," & "C:\NSPECT\wsdelin\Test2\basinpoly.shp" & "," & "NitIndex" & "; " & _
                'concentrationStatement = "Pick," & "C:\NSPECT\TestSpect\TestShape.shp" & "," & "NitTest2" & "; " & _
                concentrationStatement = "Pick," & "C:\NSPECT\TestSpect\TestShape.shp" & "," & "NitIndex" & "; " & _
                    concentrationStateArray(0) & "; " & concentrationStateArray(1) & "; " & _
                    concentrationStateArray(2) & "; " & concentrationStateArray(3)
                MsgBox("new conc statement = " & concentrationStatement)
            Else

            End If
            PollutantConcentrationSetup = CalcPollutantConcentration(concentrationStatement, OutputItems)

        Catch ex As Exception
            HandleError(ex)
            Return False
        End Try
    End Function

    Private Function ConstructMetaData(ByRef cmdType As OleDbCommand, ByRef strCoeffSet As String, ByRef includeLocalEffects As Boolean) As String
        'Takes the rs and creates a string describing the pollutants and coefficients used in this run, will
        'later be added to the global dictionary

        Dim strConstructMetaData As String = vbTab & "Pollutant Coefficients:" & vbNewLine & vbTab & vbTab & "Pollutant: " & _PollutantName & vbNewLine & vbTab & vbTab & "Coefficient Set: " & strCoeffSet & vbNewLine & vbTab & vbTab & "The following lists the landcover classes and associated coefficients used" & vbNewLine & vbTab & vbTab & "in the OpenNSPECT analysis run that created this dataset: " & vbNewLine
        Dim strLandClassCoeff As String = ""
        Dim strHeader As String
        Dim i As Short

        If includeLocalEffects Then
            strHeader = vbTab & "Input Datasets:" & vbNewLine & vbTab & vbTab & "Hydrologic soils grid: " & g_Project.SoilsHydDirectory & vbNewLine & vbTab & vbTab & "Landcover grid: " & g_Project.LandCoverGridDirectory & vbNewLine & vbTab & vbTab & "Landcover grid type: " & g_Project.LandCoverGridType & vbNewLine & vbTab & vbTab & "Landcover grid units: " & g_Project.LandCoverGridUnits & vbNewLine & vbTab & vbTab & "Precipitation grid: " & g_strPrecipFileName & vbNewLine
        Else
            strHeader = vbTab & "Input Datasets:" & vbNewLine & vbTab & vbTab & "Hydrologic soils grid: " & g_Project.SoilsHydDirectory & vbNewLine & vbTab & vbTab & "Landcover grid: " & g_Project.LandCoverGridDirectory & vbNewLine & vbTab & vbTab & "Precipitation grid: " & g_strPrecipFileName & vbNewLine & vbTab & vbTab & "Flow direction grid: " & g_strFlowDirFilename & vbNewLine
        End If

        Dim dataType As OleDbDataReader = cmdType.ExecuteReader()
        While dataType.Read()
            If i = 1 Then
                strLandClassCoeff = vbTab & vbTab & vbTab & dataType("Name") & ": " & dataType("CoeffType") & vbNewLine
            Else
                strLandClassCoeff = strLandClassCoeff & vbTab & vbTab & vbTab & dataType("Name") & ": " & dataType("CoeffType") & vbNewLine
            End If
        End While
        dataType.Close()

        ConstructMetaData = strHeader & g_strLandCoverParameters & strConstructMetaData & strLandClassCoeff

    End Function

    Private Sub CalcMassOfPhosperous(ByRef concentrationStatement As String, ByRef pMassVolumeRaster As Grid)
        _picks = concentrationStatement.Split(",")
        Dim massvolcalc As New RasterMathCellCalc(AddressOf massvolCellCalc)
        RasterMath(g_LandCoverRaster, g_pMetRunoffRaster, Nothing, Nothing, Nothing, pMassVolumeRaster, massvolcalc)
    End Sub

    Private Sub CalcVariableMassOfPhosperous(ByRef concentrationStatement As String, ByRef pMassVolumeRaster As Grid)

        Dim idxHead As New GridHeader
        Dim noData As Integer = -9999
        Dim projX, projY As New Double
        Dim idxSelected As New Object
        Dim idxName As String
        Dim idxGrid As New Grid
        Dim ptExtent As New Extents
        Dim intIdx As Integer
        Dim tmpCoeffs() As String
        Dim indexShapefile As New Shapefile
        Dim indexField As Integer
        Dim fieldName As String

        ' Parse concentrationStatement into 5 subcomponents: Flag and Shapefile/Field names and 4 sets of coefficients.
        _picks = concentrationStatement.Split(";")
        ' Parse first subcomponent into target shapefile and index field names 
        tmpCoeffs = _picks(0).Split(",")
        Dim tmpString As String = tmpCoeffs(1)
        indexShapefile.Open(tmpString)
        fieldName = tmpCoeffs(2).ToString
        indexField = indexShapefile.Table.FieldIndexByName(fieldName)

        For i As Integer = 0 To 3
            tmpCoeffs = _picks(i + 1).Split(",")
            If i = 0 Then 'Now that we know how many coefficients there are, redimension and fill the new 2-D picks array.
                ReDim _picksArray(3, tmpCoeffs.Length - 1)
            End If
            For j = 0 To tmpCoeffs.Length - 1
                _picksArray(i, j) = tmpCoeffs(j)
            Next
        Next
        ' Create an index Grid, idxGrid from the shapefile.  This can be used in RasterMath calcs to pick correct coeff from _pickArray
        idxName = GetTempFileNameOutputGridExt()
        idxHead.CopyFrom(g_LandCoverRaster.Header)  'Copy the Index Grid header from that of th eLand Cover Grid.  They need to match.
        idxHead.NodataValue = noData
        If (idxGrid.CreateNew(idxName, idxHead, GridDataType.ShortDataType, noData, False, GridFileType.GeoTiff)) Then
            '
            ' Working!  12/17/2013 DLE
            Dim nr = idxHead.NumberRows - 1
            Dim nc = idxHead.NumberCols - 1
            Dim foundCells As Integer = 0
            Using progress2 = New SynchronousProgressDialog("Picking Pollutant Coefficients Using Shapefile...", "Processing Soils", nr + 2, g_MainForm)
                indexShapefile.BeginPointInShapefile()
                For row As Integer = 0 To nr
                    progress2.Increment("Picking Pollutant Coefficients Using Shapefile...")
                    For col As Integer = 0 To nc
                        idxGrid.CellToProj(col, row, projX, projY)
                        intIdx = indexShapefile.PointInShapefile(projX, projY)
                        If intIdx <> -1 Then
                            idxGrid.Value(col, row) = indexShapefile.CellValue(indexField, intIdx) - 1  'Subtract 1 because Index shapefile gives coefficents 1-4, but array is 0-3
                            foundCells = foundCells + 1
                        End If
                    Next
                Next
                indexShapefile.EndPointInShapefile()
            End Using
            If foundCells = 0 Then
                MsgBox("Error! Index grid not populated.  Where am the data?")
                'Else
                '    '    MsgBox("Found " & foundCells.ToString & " cells.")
                '    idxGrid.Save("c:\NSPECT\idxGrid.tif")
                '    '    'Exit Sub
            End If

        Else
            MsgBox("Error! Index grid not created properly!")
            Exit Sub
        End If

        Dim massvolvarcalc As New RasterMathCellCalc(AddressOf massvolVariableCellCalc)
        RasterMath(g_LandCoverRaster, g_pMetRunoffRaster, idxGrid, Nothing, Nothing, pMassVolumeRaster, massvolvarcalc)
    End Sub

    Private Sub CreateLayerForLocalEffect(ByRef OutputItems As OutputItems, ByVal pMassVolumeRaster As Grid, ByVal outputFileNameOutConc As String)
        Dim pPermMassVolumeRaster As Grid

        'Added 7/23/04 to account for clip by selected polys functionality
        If g_Project.UseSelectedPolygons Then
            pPermMassVolumeRaster = ClipBySelectedPoly(pMassVolumeRaster, g_pSelectedPolyClip, outputFileNameOutConc)
        Else
            pPermMassVolumeRaster = CopyRaster(pMassVolumeRaster, outputFileNameOutConc)
        End If

        g_dicMetadata.Add(_PollutantName & "Local Effects (mg)", _PollutantCoeffMetadata)
        writeMetadata(g_Project.ProjectName, _PollutantName & "Local Effects (mg)", _PollutantCoeffMetadata, pPermMassVolumeRaster.Filename)

        AddOutputGridLayer(pPermMassVolumeRaster, _PollutantColor, True, _PollutantName & " Local Effects (mg)", String.Format("Pollutant {0} Local", _PollutantName), -1, OutputItems)

    End Sub
    Private Sub DeriveAccumulatedPollutant(ByVal pMassVolumeRaster As Grid, ByRef pAccumPollRaster As Grid)
        'Use weightedaread8 from geoproc to accum, then rastercalc to multiply this out
        Dim pTauD8Flow As Grid = Nothing

        Dim tauD8calc = GetConverterToTauDemFromEsri()
        RasterMath(g_pFlowDirRaster, Nothing, Nothing, Nothing, Nothing, pTauD8Flow, Nothing, False, tauD8calc)
        pTauD8Flow.Header.NodataValue = -1

        Dim flowdir As String = GetTempFileNameTauDemGridExt()
        pTauD8Flow.Save()  'Saving because it seemed necessary.
        pTauD8Flow.Save(flowdir)

        Dim metRun As String = GetTempFileNameTauDemGridExt()
        pMassVolumeRaster.Save(metRun)

        Dim strtmpout As String = GetTempFileNameTauDemGridExt()

        Dim result = Hydrology.WeightedAreaD8(flowdir, metRun, "", strtmpout, False, False, Environment.ProcessorCount, False, Nothing)

        If result <> 0 Then
            SynchronousProgressDialog.KeepRunning = False
        End If
        Dim tmpGrid As New Grid
        tmpGrid.Open(strtmpout)

        Dim multAccumcalc As New RasterMathCellCalc(AddressOf multAccumCellCalc)
        RasterMath(tmpGrid, Nothing, Nothing, Nothing, Nothing, pAccumPollRaster, multAccumcalc)

        pTauD8Flow.Close()
        DataManagement.DeleteGrid(flowdir)
    End Sub
    Private Sub AddAccumulatedPollutantToGroupLayer(ByRef OutputItems As OutputItems, ByRef pAccumPollRaster As Grid)
        Dim strAccPoll As String = GetUniqueFileName("accpoll", g_Project.ProjectWorkspace, OutputGridExt)
        'Added 7/23/04 to account for clip by selected polys functionality
        Dim pPermAccPollRaster As Grid
        If g_Project.UseSelectedPolygons Then
            pPermAccPollRaster = ClipBySelectedPoly(pAccumPollRaster, g_pSelectedPolyClip, strAccPoll)
        Else
            pPermAccPollRaster = CopyRaster(pAccumPollRaster, strAccPoll)
        End If

        Dim layerName As String = String.Format("Accumulated {0} (kg)", _PollutantName)
        g_dicMetadata.Add(layerName, _PollutantCoeffMetadata)
        writeMetadata(g_Project.ProjectName, layerName, _PollutantCoeffMetadata, pPermAccPollRaster.Filename)

        AddOutputGridLayer(pPermAccPollRaster, _PollutantColor, True, layerName, String.Format("Pollutant {0} Accum", _PollutantName), -1, OutputItems)
    End Sub
    Private Sub CalcFinalConcentration(ByVal pMassVolumeRaster As Grid, ByVal pAccumPollRaster As Grid, ByRef pTotalPollConc0Raster As Grid)
        Dim AllConCalc As New RasterMathCellCalcNulls(AddressOf AllConCellCalc)
        RasterMath(pMassVolumeRaster, pAccumPollRaster, g_pMetRunoffRaster, g_pRunoffRaster, g_pDEMRaster, pTotalPollConc0Raster, Nothing, False, AllConCalc)
    End Sub
    Private Sub CreateDataLayer(ByRef OutputItems As OutputItems, ByVal pTotalPollConc0Raster As Grid, ByVal outputFileNameOutConc As Object)
        outputFileNameOutConc = GetUniqueFileName("conc", g_Project.ProjectWorkspace, OutputGridExt)
        Dim pPermTotalConcRaster As Grid
        If g_Project.UseSelectedPolygons Then
            pPermTotalConcRaster = ClipBySelectedPoly(pTotalPollConc0Raster, g_pSelectedPolyClip, outputFileNameOutConc)
        Else
            pPermTotalConcRaster = CopyRaster(pTotalPollConc0Raster, outputFileNameOutConc)
        End If

        g_dicMetadata.Add(_PollutantName & " Conc. (mg/L)", _PollutantCoeffMetadata)
        'Dim tmpString As String = _PollutantName & " Conc. (mg/L)"
        writeMetadata(g_Project.ProjectName, _PollutantName & " Conc. (mg/L)", _PollutantCoeffMetadata, pPermTotalConcRaster.Filename)

        AddOutputGridLayer(pPermTotalConcRaster, _PollutantColor, True, _PollutantName & " Conc. (mg/L)", String.Format("Pollutant {0} Conc", _PollutantName), -1, OutputItems)
    End Sub
    Private Function CalcPollutantConcentration(ByRef concentrationStatement As String, ByRef OutputItems As OutputItems) As Boolean
        If concentrationStatement = Nothing Then
            Throw New ArgumentNullException("concentrationStatement")
        End If
        _picks = concentrationStatement.Split(";")

        Dim massVolumeRaster As Grid = Nothing
        Dim pAccumPollRaster As Grid = Nothing
        Dim pTotalPollConc0Raster As Grid = Nothing

        Dim strTitle = String.Format("Processing {0} Conc. Calculation...", _PollutantName)
        Dim outputFileNameOutConc = GetUniqueFileName("locconc", g_Project.ProjectWorkspace, OutputGridExt)
        Dim progress = New SynchronousProgressDialog(strTitle, 13, g_MainForm)
        Try
            progress.Increment("Calculating Mass Volume...")
            If _picks(0).Split(",")(0) = "Pick" Then
                CalcVariableMassOfPhosperous(concentrationStatement, massVolumeRaster) 'New RasterMath setup for picking conc. from a shapefile.  
            Else
                CalcMassOfPhosperous(concentrationStatement, massVolumeRaster) 'Concentrations done here: Change for using polygon to pick
            End If

            'At this point the above grid will satisfy 'local effects only' people so...
            If g_Project.IncludeLocalEffects Then
                If Not progress.Increment("Creating data layer for local effects...") Then Return False
                CreateLayerForLocalEffect(OutputItems, massVolumeRaster, outputFileNameOutConc)
            Else
                massVolumeRaster.Save()
            End If

            If Not progress.Increment("Deriving accumulated pollutant...") Then Return False
            DeriveAccumulatedPollutant(massVolumeRaster, pAccumPollRaster)

            If Not progress.Increment("Creating accumlated pollutant layer...") Then Return False
            AddAccumulatedPollutantToGroupLayer(OutputItems, pAccumPollRaster)

            If Not progress.Increment("Calculating final concentration...") Then Return False
            CalcFinalConcentration(massVolumeRaster, pAccumPollRaster, pTotalPollConc0Raster)

            If Not progress.Increment("Creating data layer...") Then Return False
            CreateDataLayer(OutputItems, pTotalPollConc0Raster, outputFileNameOutConc)

            ' We don't do this anymore, so skip the calcualations entirely
            'If Not progress.Increment("Comparing to water quality standard...") Then Return False
            'If Not CompareWaterQuality(pTotalPollConc0Raster, OutputItems) Then Return False

            Return True

        Catch ex As Exception
            HandleError(ex)
            Return False
        Finally
            progress.Dispose()
        End Try

    End Function

    Private Function CompareWaterQuality(ByRef pPollutantRaster As Grid, ByRef OutputItems As OutputItems) As Boolean
        Dim strWQVAlue As Object

        'Get the zone dataset from the first layer in ArcMap

        Dim pConRaster As Grid = Nothing
        Dim pPermWQRaster As Grid = Nothing
        Dim dblConvertValue As Double
        Dim strOutWQ As String
        Dim strMetadata As String

        Try

            ' Perform Spatial operation
            strWQVAlue = ReturnWQValue(_PollutantName, _WaterQualityStandardName)

            _WQValue = (CDbl(strWQVAlue)) / 1000
            _FlowMax = g_pFlowAccRaster.Maximum

            Dim concalc As New RasterMathCellCalc(AddressOf concompCellCalc)
            RasterMath(pPollutantRaster, g_pFlowAccRaster, Nothing, Nothing, Nothing, pConRaster, concalc)

            strOutWQ = GetUniqueFileName("wq", g_Project.ProjectWorkspace, OutputGridExt)

            'Clip if selectedpolys
            If g_Project.UseSelectedPolygons Then
                pPermWQRaster = ClipBySelectedPoly(pConRaster, g_pSelectedPolyClip, strOutWQ)
            Else
                pPermWQRaster = CopyRaster(pConRaster, strOutWQ)
            End If

            strMetadata = vbTab & "Water Quality Standard:" & vbNewLine & vbTab & vbTab & "Criteria Name: " & _WaterQualityStandardName & vbNewLine & vbTab & vbTab & "Standard: " & dblConvertValue & " mg/L"
            Dim layerCaption As String = String.Format("{0} Standard: {1} mg/L", _PollutantName, CStr(dblConvertValue))
            g_dicMetadata.Add(layerCaption, _PollutantCoeffMetadata & strMetadata)
            writeMetadata(g_Project.ProjectName, layerCaption, _PollutantCoeffMetadata, pPermWQRaster.Filename)

            ' Hide standard output because we're not sure anyone uses it http://nspect.codeplex.com/workitem/20911
            If False Then
                AddOutputGridLayer(pPermWQRaster, _WaterQualityStandardName, False, layerCaption, "Pollutant " & _PollutantName & " WQ", -1, OutputItems)
            End If


            Return True
        Catch ex As Exception
            HandleError(ex)
            Return False
        End Try
    End Function

    Private Function ReturnWQValue(ByRef strPollName As String, ByRef strWQstdName As String) As String
        Dim strPoll As String
        Dim strWQStd As String = ""
        ReturnWQValue = ""
        Try
            strPoll = String.Format("Select * from Pollutant where name like '{0}'", strPollName)
            Using cmdpoll As New DataHelper(strPoll)
                Dim datapoll As OleDbDataReader = cmdpoll.ExecuteReader
                datapoll.Read()
                strWQStd = String.Format("SELECT * FROM WQCRITERIA INNER JOIN POLL_WQCRITERIA ON WQCRITERIA.WQCRITID = POLL_WQCRITERIA.WQCRITID WHERE WQCRITERIA.NAME Like '{0}' AND POLL_WQCRITERIA.POLLID = {1}", strWQstdName, datapoll("POLLID"))
                datapoll.Close()
                Using cmdWQ As New DataHelper(strWQStd)
                    Dim datawq As OleDbDataReader = cmdWQ.ExecuteReader()
                    datawq.Read()
                    ReturnWQValue = CStr(datawq("Threshold"))
                    datawq.Close()
                End Using
            End Using
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Function

#Region "Raster Math"

    Private Function massvolCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        Dim tmpval As Single
        'strexpression = pick([pLandSampleRaster], _picks)"
        For i As Integer = 0 To _picks.Length - 1
            If Input1 = i + 1 Then
                tmpval = _picks(i)
                Exit For
            End If
        Next

        'strExpression = "[met_runoff] * [pollmass]"
        Return Input2 * tmpval
    End Function

    Private Function massvolVariableCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        Dim tmpval As Single
        Dim numCoeff As New Single
        numCoeff = _picksArray.Length / 4
        'strexpression = pick([pLandSampleRaster], _picks)"
        For i As Integer = 0 To numCoeff - 1
            If Input1 = i + 1 Then
                tmpval = _picksArray(Input3, i)
                Exit For
            End If
        Next

        'strExpression = "[met_runoff] * [pollmass]"
        Return Input2 * tmpval
    End Function

    Private Function multAccumCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        'strExpression = "(FlowAccumulation([flowdir], [massvolume], FLOAT)) * 1.0e-6"
        Return Input1 * 0.000001
    End Function

    Private Function AllConCellCalc(ByVal Input1 As Single, ByVal Input1Null As Single, ByVal Input2 As Single, ByVal Input2Null As Single, ByVal Input3 As Single, ByVal Input3Null As Single, ByVal Input4 As Single, ByVal Input4Null As Single, ByVal Input5 As Single, ByVal Input5Null As Single, ByVal OutNull As Single) As Single
        If Input1 = Input1Null Or Input2 = Input2Null Or Input3 = Input3Null Or Input4 = Input4Null Then
            If Input5 = Input5Null Then
                Return OutNull
            Else
                If Input5 > 0 Then
                    Return 0
                Else
                    Return OutNull
                End If
            End If
        Else
            Return (Input1 + (Input2 / 0.000001)) / (Input3 + Input4)
        End If

    End Function

    Private Function concompCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        'This rather ugly expression was set up to check for meets/exceed water quality standards for
        'only the streams.  It takes the values of flowaccumulation from watershed delineation fame that
        'exceed values of greater than 1%.  Then multiplies the result (all cells representing streams) times
        'the water quality grid.
        'strExpression = "(Con([Max] gt " & CStr(dblConvertValue) & ", 1, 2)) * (con([flowacc] > (" & CStr(modUtil.ReturnRasterMax(g_pFlowAccRaster)) & " * 0.01), 1))"
        'strExpression = "(Con([Max] gt _WQValue, 1, 2)) * (con([flowacc] > (_FlowMax * 0.01), 1))"
        If Input1 > _WQValue Then
            '(con([flowacc] > (" & CStr(modUtil.ReturnRasterMax(g_pFlowAccRaster)) & " * 0.01), 1))
            If Input2 > (_FlowMax * 0.01) Then
                Return 1
            Else
                Return OutNull
            End If
        Else
            If Input2 > (_FlowMax * 0.01) Then
                Return 2
            Else
                Return OutNull
            End If
        End If
    End Function

#End Region
End Module