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
Imports System.IO
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

    Public Function PollutantConcentrationSetup(ByRef Pollutant As PollutantItem, ByRef strWQName As String, ByRef OutputItems As OutputItems) As Boolean
        'Sub takes incoming parameters (in the form of a pollutant item) from the project file
        'and then parses them out
        Try
            'Open Strings
            Dim strType As String
            Dim strField As String = ""
            Dim strConStatement As String = ""
            Dim strTempCoeffSet As String
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
                Case ""
            End Select

            'Find out the name of the Coefficient set, could be a temporary one due to landuses
            If g_DictTempNames.Count > 0 Then
                If Len(g_DictTempNames.Item(Pollutant.strCoeffSet)) > 0 Then
                    strTempCoeffSet = g_DictTempNames.Item(Pollutant.strCoeffSet)
                Else
                    strTempCoeffSet = Pollutant.strCoeffSet
                End If
            Else
                strTempCoeffSet = Pollutant.strCoeffSet
            End If

            If Len(strField) > 0 Then
                Using cmdPoll As New DataHelper("SELECT * FROM COEFFICIENTSET WHERE NAME LIKE '" & strTempCoeffSet & "'")
                    Using dataPoll As OleDbDataReader = cmdPoll.ExecuteReader()
                        dataPoll.Read()
                        strType = "SELECT LCCLASS.Value, LCCLASS.Name, COEFFICIENT." & strField & " As CoeffType, COEFFICIENT.CoeffID, COEFFICIENT.LCCLASSID " & "FROM LCCLASS LEFT OUTER JOIN COEFFICIENT " & "ON LCCLASS.LCCLASSID = COEFFICIENT.LCCLASSID " & "WHERE COEFFICIENT.COEFFSETID = " & dataPoll("CoeffSetID") & " ORDER BY LCCLASS.VALUE"
                        Using cmdType As New DataHelper(strType)
                            Dim command As OleDbCommand = cmdType.GetCommand()
                            strConStatement = ConstructPickStatmentUsingLandClass(command, g_LandCoverRaster, "CoeffType")
                            _PollutantCoeffMetadata = ConstructMetaData(command, (Pollutant.strCoeff), g_booLocalEffects)
                        End Using
                    End Using
                End Using

            End If

            'Find out the color of the pollutant
            Using cmdPollColor As New DataHelper("Select Color from Pollutant where NAME LIKE '" & _PollutantName & "'")
                Using datapollcolor As OleDbDataReader = cmdPollColor.ExecuteReader()
                    datapollcolor.Read()
                    _PollutantColor = CStr(datapollcolor("Color"))
                    datapollcolor.Close()
                End Using
            End Using

            PollutantConcentrationSetup = CalcPollutantConcentration(strConStatement, OutputItems)

        Catch ex As Exception
            HandleError(ex)
            Return False
        End Try
    End Function

    Private Function ConstructMetaData(ByRef cmdType As OleDbCommand, ByRef strCoeffSet As String, ByRef booLocal As Boolean) As String
        'Takes the rs and creates a string describing the pollutants and coefficients used in this run, will
        'later be added to the global dictionary

        Dim strConstructMetaData As String = vbTab & "Pollutant Coefficients:" & vbNewLine & vbTab & vbTab & "Pollutant: " & _PollutantName & vbNewLine & vbTab & vbTab & "Coefficient Set: " & strCoeffSet & vbNewLine & vbTab & vbTab & "The following lists the landcover classes and associated coefficients used" & vbNewLine & vbTab & vbTab & "in the OpenNSPECT analysis run that created this dataset: " & vbNewLine
        Dim strLandClassCoeff As String = ""
        Dim strHeader As String
        Dim i As Short

        If booLocal Then
            strHeader = vbTab & "Input Datasets:" & vbNewLine & vbTab & vbTab & "Hydrologic soils grid: " & g_XmlPrjFile.SoilsHydDirectory & vbNewLine & vbTab & vbTab & "Landcover grid: " & g_XmlPrjFile.LandCoverGridDirectory & vbNewLine & vbTab & vbTab & "Landcover grid type: " & g_XmlPrjFile.LandCoverGridType & vbNewLine & vbTab & vbTab & "Landcover grid units: " & g_XmlPrjFile.LandCoverGridUnits & vbNewLine & vbTab & vbTab & "Precipitation grid: " & g_strPrecipFileName & vbNewLine
        Else
            strHeader = vbTab & "Input Datasets:" & vbNewLine & vbTab & vbTab & "Hydrologic soils grid: " & g_XmlPrjFile.SoilsHydDirectory & vbNewLine & vbTab & vbTab & "Landcover grid: " & g_XmlPrjFile.LandCoverGridDirectory & vbNewLine & vbTab & vbTab & "Precipitation grid: " & g_strPrecipFileName & vbNewLine & vbTab & vbTab & "Flow direction grid: " & g_strFlowDirFilename & vbNewLine
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

    Private Sub CalcMassOfPhosperous(ByRef strConStatement As String, ByRef pMassVolumeRaster As Grid)
        ReDim _picks(strConStatement.Split(",").Length)
        _picks = strConStatement.Split(",")
        Dim massvolcalc As New RasterMathCellCalc(AddressOf massvolCellCalc)
        RasterMath(g_LandCoverRaster, g_pMetRunoffRaster, Nothing, Nothing, Nothing, pMassVolumeRaster, massvolcalc)
    End Sub
    Private Sub CreateLayerForLocalEffect(ByRef OutputItems As OutputItems, ByVal pMassVolumeRaster As Grid, ByVal outputFileNameOutConc As String)
        Dim pPermMassVolumeRaster As Grid

        'Added 7/23/04 to account for clip by selected polys functionality
        If g_booSelectedPolys Then
            pPermMassVolumeRaster = ClipBySelectedPoly(pMassVolumeRaster, g_pSelectedPolyClip, outputFileNameOutConc)
        Else
            pPermMassVolumeRaster = CopyRaster(pMassVolumeRaster, outputFileNameOutConc)
        End If

        g_dicMetadata.Add(_PollutantName & "Local Effects (mg)", _PollutantCoeffMetadata)

        AddOutputGridLayer(pPermMassVolumeRaster, _PollutantColor, True, _PollutantName & " Local Effects (mg)", String.Format("Pollutant {0} Local", _PollutantName), -1, OutputItems)

    End Sub
    Private Sub DeriveAccumulatedPollutant(ByVal pMassVolumeRaster As Grid, ByRef pAccumPollRaster As Grid)
        'Use weightedaread8 from geoproc to accum, then rastercalc to multiply this out
        Dim pTauD8Flow As Grid = Nothing

        Dim tauD8calc = GetConverterToTauDemFromEsri()
        RasterMath(g_pFlowDirRaster, Nothing, Nothing, Nothing, Nothing, pTauD8Flow, Nothing, False, tauD8calc)
        pTauD8Flow.Header.NodataValue = -1

        Dim strtmp1 As String = Path.GetTempFileName
        g_TempFilesToDel.Add(strtmp1)
        strtmp1 = strtmp1 + TAUDEMGridExt
        g_TempFilesToDel.Add(strtmp1)
        DataManagement.DeleteGrid(strtmp1)
        pTauD8Flow.Save(strtmp1)

        Dim strtmp2 As String = Path.GetTempFileName
        g_TempFilesToDel.Add(strtmp2)
        strtmp2 = strtmp2 + TAUDEMGridExt
        g_TempFilesToDel.Add(strtmp2)
        DataManagement.DeleteGrid(strtmp2)
        pMassVolumeRaster.Save(strtmp2)

        Dim strtmpout As String = Path.GetTempFileName
        g_TempFilesToDel.Add(strtmpout)
        strtmpout = String.Format("{0}out{1}", strtmpout, TAUDEMGridExt)
        g_TempFilesToDel.Add(strtmpout)
        DataManagement.DeleteGrid(strtmpout)

        'Use geoproc weightedAreaD8 after converting the D8 grid to taudem format bgd if needed
        Dim result = Hydrology.WeightedAreaD8(strtmp1, strtmp2, "", strtmpout, False, False, Environment.ProcessorCount, Nothing)
        'strExpression = "FlowAccumulation([flowdir], [met_run], FLOAT)"
        If result <> 0 Then
            SynchronousProgressDialog.KeepRunning = False
        End If
        Dim tmpGrid As New Grid
        tmpGrid.Open(strtmpout)

        Dim multAccumcalc As New RasterMathCellCalc(AddressOf multAccumCellCalc)
        RasterMath(tmpGrid, Nothing, Nothing, Nothing, Nothing, pAccumPollRaster, multAccumcalc)

        pTauD8Flow.Close()
        DataManagement.DeleteGrid(strtmp1)
    End Sub
    Private Sub AddAccumulatedPollutantToGroupLayer(ByRef OutputItems As OutputItems, ByRef pAccumPollRaster As Grid)
        Dim strAccPoll As String = GetUniqueFileName("accpoll", g_XmlPrjFile.ProjectWorkspace, FinalOutputGridExt)
        'Added 7/23/04 to account for clip by selected polys functionality
        Dim pPermAccPollRaster As Grid
        If g_booSelectedPolys Then
            pPermAccPollRaster = ClipBySelectedPoly(pAccumPollRaster, g_pSelectedPolyClip, strAccPoll)
        Else
            pPermAccPollRaster = CopyRaster(pAccumPollRaster, strAccPoll)
        End If

        Dim layerName As String = String.Format("Accumulated {0} (kg)", _PollutantName)
        g_dicMetadata.Add(layerName, _PollutantCoeffMetadata)

        AddOutputGridLayer(pPermAccPollRaster, _PollutantColor, True, layerName, String.Format("Pollutant {0} Accum", _PollutantName), -1, OutputItems)
    End Sub
    Private Sub CalcFinalConcentration(ByVal pMassVolumeRaster As Grid, ByVal pAccumPollRaster As Grid, ByRef pTotalPollConc0Raster As Grid)
        Dim AllConCalc As New RasterMathCellCalcNulls(AddressOf AllConCellCalc)
        RasterMath(pMassVolumeRaster, pAccumPollRaster, g_pMetRunoffRaster, g_pRunoffRaster, g_pDEMRaster, pTotalPollConc0Raster, Nothing, False, AllConCalc)
    End Sub
    Private Sub CreateDataLayer(ByRef OutputItems As OutputItems, ByVal pTotalPollConc0Raster As Grid, ByVal outputFileNameOutConc As Object)
        outputFileNameOutConc = GetUniqueFileName("conc", g_XmlPrjFile.ProjectWorkspace, FinalOutputGridExt)
        Dim pPermTotalConcRaster As Grid
        If g_booSelectedPolys Then
            pPermTotalConcRaster = ClipBySelectedPoly(pTotalPollConc0Raster, g_pSelectedPolyClip, outputFileNameOutConc)
        Else
            pPermTotalConcRaster = CopyRaster(pTotalPollConc0Raster, outputFileNameOutConc)
        End If

        g_dicMetadata.Add(_PollutantName & " Conc. (mg/L)", _PollutantCoeffMetadata)

        AddOutputGridLayer(pPermTotalConcRaster, _PollutantColor, True, _PollutantName & " Conc. (mg/L)", String.Format("Pollutant {0} Conc", _PollutantName), -1, OutputItems)
    End Sub
    Private Function CalcPollutantConcentration(ByRef strConStatement As String, ByRef OutputItems As OutputItems) As Boolean

        Dim pMassVolumeRaster As Grid = Nothing
        Dim pAccumPollRaster As Grid = Nothing
        Dim pTotalPollConc0Raster As Grid = Nothing

        Dim strTitle = String.Format("Processing {0} Conc. Calculation...", _PollutantName)
        Dim outputFileNameOutConc = GetUniqueFileName("locconc", g_XmlPrjFile.ProjectWorkspace, FinalOutputGridExt)
        Dim progress = New SynchronousProgressDialog(strTitle, 13, g_MainForm)
        Try
            progress.Increment("Calculating Mass Volume...")
            CalcMassOfPhosperous(strConStatement, pMassVolumeRaster)

            'At this point the above grid will satisfy 'local effects only' people so...
            If g_booLocalEffects Then
                If Not progress.Increment("Creating data layer for local effects...") Then Return False
                CreateLayerForLocalEffect(OutputItems, pMassVolumeRaster, outputFileNameOutConc)
            End If

            If Not progress.Increment("Deriving accumulated pollutant...") Then Return False
            DeriveAccumulatedPollutant(pMassVolumeRaster, pAccumPollRaster)

            If Not progress.Increment("Creating accumlated pollutant layer...") Then Return False
            AddAccumulatedPollutantToGroupLayer(OutputItems, pAccumPollRaster)

            If Not progress.Increment("Calculating final concentration...") Then Return False
            CalcFinalConcentration(pMassVolumeRaster, pAccumPollRaster, pTotalPollConc0Raster)

            If Not progress.Increment("Creating data layer...") Then Return False
            CreateDataLayer(OutputItems, pTotalPollConc0Raster, outputFileNameOutConc)

            If Not progress.Increment("Comparing to water quality standard...") Then Return False
            If Not CompareWaterQuality(pTotalPollConc0Raster, OutputItems) Then Return False

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

            strOutWQ = GetUniqueFileName("wq", g_XmlPrjFile.ProjectWorkspace, FinalOutputGridExt)

            'Clip if selectedpolys
            If g_booSelectedPolys Then
                pPermWQRaster = ClipBySelectedPoly(pConRaster, g_pSelectedPolyClip, strOutWQ)
            Else
                pPermWQRaster = CopyRaster(pConRaster, strOutWQ)
            End If

            strMetadata = vbTab & "Water Quality Standard:" & vbNewLine & vbTab & vbTab & "Criteria Name: " & _WaterQualityStandardName & vbNewLine & vbTab & vbTab & "Standard: " & dblConvertValue & " mg/L"
            g_dicMetadata.Add(_PollutantName & " Standard: " & CStr(dblConvertValue) & " mg/L", _PollutantCoeffMetadata & strMetadata)

            AddOutputGridLayer(pPermWQRaster, _WaterQualityStandardName, False, _PollutantName & " Standard: " & CStr(dblConvertValue) & " mg/L", "Pollutant " & _PollutantName & " WQ", -1, OutputItems)

            CompareWaterQuality = True

        Catch ex As Exception
            HandleError(ex)
            CompareWaterQuality = False
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