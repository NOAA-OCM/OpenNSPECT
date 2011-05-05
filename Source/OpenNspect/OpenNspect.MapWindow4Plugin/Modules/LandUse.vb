'********************************************************************************************************
'File Name: modLanduse.vb
'Description: Landuse functionality
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
Imports System.Data
Imports MapWinGIS
Imports System.Data.OleDb
Imports OpenNspect.Xml

Module LandUse
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

    Private _intLCTypeID As Short
    Private _intCoeffSetID As Short
    Private _strLCFileName As String

    Private _strLCClass As String
    Private _pLandCoverRaster As Grid

    Public g_booLCChange As Boolean
    Public g_strLCTypeName As String
    'the temp name of Land Cover type, if indeed landuses are applied.
    Public g_DictTempNames As New Dictionary(Of String, String)
    'Array holding the temp names of the LCType, and subsequent Coefficient Set Names
    ' Constant used by the Error handler function - DO NOT REMOVE

    ''' <summary>
    ''' Begins working.
    ''' </summary>
    ''' <param name="strLCClassName">str of current land cover class.</param>
    ''' <param name="LUScenItems">Xml class that holds the params of the user's Land Use Scenario.</param>
    ''' <param name="dictPollutants">dictionary created to hold the pollutants of this particular project.</param>
    ''' <param name="strLCFileName">FileName of land cover grid.</param>
    Public Sub Begin(ByRef strLCClassName As String, ByRef LUScenItems As LandUseItems, ByRef dictPollutants As Dictionary(Of String, String), ByRef strLCFileName As String)

        Try
            Dim strCurrentLCType As String
            'Current LCTYpe
            Dim strCurrentLCClass As String
            'Current Landclasses of The LCType
            Dim strInsertTempLCType As String
            Dim strTempLCTypeName As String
            Dim strrsInsertLCType As String
            'the inserted LCTYPE
            Dim strrsPermLandClass As String
            'the landclass table currently in the database
            Dim strNewLandClass As String
            'Newly added landclass - to get its new ID

            Dim i As Short
            Dim j As Short
            Dim k As Short
            Dim intValue As Short
            'Temp new landclasses fake value
            Dim LUScen As New LandUseMangementScenario
            'Xml Land use scenario
            Dim strCoeffSetTempName As String
            'New temp name for the coefficient set
            Dim strCoeffSetOrigName As String
            'Original name for the coefficient set
            Dim pollArray As Dictionary(Of String, String).KeyCollection.Enumerator
            'Array to hold pollutants...get itself from the dictionary

            'init the file name of the LandClass File
            _strLCFileName = strLCFileName

            'STEP 1: Get the current LCTYPE
            Dim lcTypeId As Object
            strCurrentLCType = String.Format("SELECT * FROM LCTYPE WHERE NAME LIKE '{0}'", strLCClassName)
            Using cmdCurrentLCType As New DataHelper(strCurrentLCType)
                Using dataCurrentLCType As OleDbDataReader = cmdCurrentLCType.ExecuteReader()
                    dataCurrentLCType.Read()
                    lcTypeId = dataCurrentLCType("LCTypeID")
                End Using
            End Using

            'MK - I'm not sure this works like they thought it did.
            'STEP 2: Get the current LCCLASSES of the current LCTYPE

            strCurrentLCClass = String.Format("SELECT * FROM LCCLASS WHERE LCTYPEID = {0} ORDER BY LCCLass.Value", lcTypeId)
            Dim cmdCurrentLCClass As New DataHelper(strCurrentLCClass)

            'STEP 3: Now INSERT a copy of current LCTYPE
            'First, get a temp name... like CCAPLUTemp1, CCAPLUTemp2 etc
            strTempLCTypeName = CreateUniqueName("LCTYPE", strLCClassName & "LUTemp")
            strInsertTempLCType = String.Format("INSERT INTO LCTYPE (NAME,DESCRIPTION) VALUES ('{0}', '{0} Description')", Replace(strTempLCTypeName, "'", "''"))

            'Add the name to the Dictionary for storage; used for cleanup and in pollutants
            g_DictTempNames.Add(strLCClassName, strTempLCTypeName)

            'STEP 4: INSERT the copy of the LCTYPE in
            Using cmdInsertCopy As New DataHelper(strInsertTempLCType)
                cmdInsertCopy.ExecuteNonQuery()
            End Using

            'STEP 5: Get it back now so you can use its ID for inserting the landclasses
            strrsInsertLCType = String.Format("SELECT LCTYPEID FROM LCTYPE WHERE LCTYPE.NAME LIKE '{0}'", strTempLCTypeName)
            Using cmdInsertType As New DataHelper(strrsInsertLCType)
                Using dataInsertType As OleDbDataReader = cmdInsertType.ExecuteReader()
                    dataInsertType.Read()
                    _intLCTypeID = dataInsertType("LCTypeID")
                End Using
            End Using

            'STEP 6: Now clone the current landclasses into a new recordset
            Dim cmdCloneLCClass As OleDbCommand = cmdCurrentLCClass.CloneCommand()

            'Prepare the landclass table to accept the copies of landclass
            strrsPermLandClass = "SELECT * FROM LCCLASS"
            Dim adaptPermLC As New OleDbDataAdapter(strrsPermLandClass, g_DBConn)
            Dim dataPermLC As New DataTable
            adaptPermLC.Fill(dataPermLC)

            'STEP 7: loop through the XmlLandUseItems to add new land uses to the copy rs
            Dim dataCloneLCClass As OleDbDataReader = cmdCloneLCClass.ExecuteReader()

            'Now add all the landclasses.
            While dataCloneLCClass.Read()
                'Add new
                Dim dataPermRow As DataRow = dataPermLC.NewRow()
                'Add the necessary components
                dataPermRow("LCTypeID") = _intLCTypeID
                dataPermRow("Value") = dataCloneLCClass("Value")
                dataPermRow("Name") = dataCloneLCClass("Name")
                dataPermRow("CN-A") = dataCloneLCClass("CN-A")
                dataPermRow("CN-B") = dataCloneLCClass("CN-B")
                dataPermRow("CN-C") = dataCloneLCClass("CN-C")
                dataPermRow("CN-D") = dataCloneLCClass("CN-D")
                dataPermRow("CoverFactor") = dataCloneLCClass("CoverFactor")
                dataPermRow("W_WL") = dataCloneLCClass("W_WL")

                'Update
                adaptPermLC.Update(dataPermLC)

                'move to next record
                intValue = dataCloneLCClass("Value")
                'This keeps track of the max
            End While
            dataCloneLCClass.Close()

            'STEP 8: Now add the new landclass
            Dim intLCClassIDs() As Short
            ReDim intLCClassIDs(LUScenItems.Count - 1)

            For i = 0 To LUScenItems.Count - 1
                If LUScenItems.Item(i).intApply = 1 Then
                    'init the fake value: will be max value + 1
                    intValue = intValue + 1

                    'Init the LUScen
                    LUScen.Xml = LUScenItems.Item(i).strLUScenXmlFile

                    Dim dataPermRow As DataRow = dataPermLC.NewRow()
                    dataPermRow("LCTypeID") = _intLCTypeID
                    dataPermRow("Value") = intValue
                    dataPermRow("Name") = LUScen.strLUScenName
                    dataPermRow("CN-A") = LUScen.intSCSCurveA
                    dataPermRow("CN-B") = LUScen.intSCSCurveB
                    dataPermRow("CN-C") = LUScen.intSCSCurveC
                    dataPermRow("CN-D") = LUScen.intSCSCurveD
                    dataPermRow("CoverFactor") = LUScen.lngCoverFactor
                    dataPermRow("W_WL") = LUScen.intWaterWetlands

                    adaptPermLC.Update(dataPermLC)
                End If

                'Gather the newly added LCClassIds in an array for use later
                strNewLandClass = String.Format("SELECT LCCLASSID FROM LCCLASS WHERE NAME LIKE '{0}'", LUScen.strLUScenName)
                Using cmdNewLC As New DataHelper(strNewLandClass)
                    Using dataNewLC As OleDbDataReader = cmdNewLC.ExecuteReader()
                        dataNewLC.Read()
                        intLCClassIDs(i - 1) = dataNewLC("LCClassID")
                    End Using
                End Using
            Next i

            'STEP 10: Parse the incoming dictionary
            'Remember, it's in this form:
            'Key_________Item__________            'Pollutant , CoefficientSet
            Dim l As Short
            pollArray = dictPollutants.Keys.GetEnumerator

            'Now to the pollutants
            'Loop through the pollutants coming from the Xml class, as well as those in the project that are being used
            For j = 0 To LUScen.PollItems.Count - 1
                For k = 0 To dictPollutants.Count - 1
                    If InStr(1, LUScen.PollItems.Item(j).strPollName, pollArray.Current, CompareMethod.Text) > 0 Then
                        strCoeffSetOrigName = dictPollutants.Item(pollArray.Current)
                        'Original Name
                        strCoeffSetTempName = CreateUniqueName("COEFFICIENTSET", strCoeffSetOrigName & "_Temp")
                        'Make a new temp name

                        'Now add names of the coefficient sets to the dictionary
                        g_DictTempNames.Add(strCoeffSetOrigName, strCoeffSetTempName)

                        Dim adaptTemp As OleDbDataAdapter = CopyCoefficient(strCoeffSetTempName, strCoeffSetOrigName)
                        'Call to function that returns the copied record set.
                        Dim dataTemp As New DataTable
                        adaptTemp.Fill(dataTemp)

                        'Now add the new values using the info in the Xml file and the array of new LCClass IDs
                        For l = 0 To UBound(intLCClassIDs)
                            Dim dataRow As DataRow = dataTemp.NewRow()

                            LUScen.Xml = LUScenItems.Item(l).strLUScenXmlFile

                            dataRow("Coeff1").Value = LUScen.PollItems.Item(j).intType1
                            dataRow("Coeff2").Value = LUScen.PollItems.Item(j).intType2
                            dataRow("Coeff3").Value = LUScen.PollItems.Item(j).intType3
                            dataRow("Coeff4").Value = LUScen.PollItems.Item(j).intType4
                            dataRow("CoeffSetID").Value = _intCoeffSetID
                            dataRow("LCClassID").Value = intLCClassIDs(l)

                            adaptTemp.Update(dataTemp)
                        Next l
                    End If
                    pollArray.MoveNext()
                Next k
            Next j

            g_strLCTypeName = strTempLCTypeName

            ReclassLanduse(LUScenItems, _strLCFileName)

        Catch ex As Exception
            HandleError(ex)
        End Try

    End Sub

    Private Function CopyCoefficient(ByRef strNewCoeffName As String, ByRef strCoeffSet As String) As OleDbDataAdapter
        'General gist:  First we add new record to the Coefficient Set table using strNewCoeffName as
        'the name, PollID, LCTYPEID.  Once that's done, we'll add the coefficients
        'from the set being copied

        CopyCoefficient = Nothing
        Try

            'CmdString for inserting new coefficientset               '
            'Holder for the CoefficientSetID

            'The Recordset of existing coefficients being copied
            Dim cmdCopySet As New DataHelper(String.Format("SELECT * FROM COEFFICIENTSET INNER JOIN COEFFICIENT ON COEFFICIENTSET.COEFFSETID = COEFFICIENT.COEFFSETID WHERE COEFFICIENTSET.NAME LIKE '{0}'", strCoeffSet))
            Dim dataCopySet As OleDbDataReader = cmdCopySet.ExecuteReader
            dataCopySet.Read()
            dataCopySet.Close()

            'INSERT: new Coefficient set taking the PollID and LCType ID from rsCopySet
            'First need to add the coefficient set to that table
            Dim cmdNewLCType As New DataHelper(String.Format("INSERT INTO COEFFICIENTSET(NAME, POLLID, LCTYPEID) VALUES ('{0}',{1},{2})", Replace(strNewCoeffName, "'", "''"), dataCopySet("POLLID"), _intLCTypeID))
            cmdNewLCType.ExecuteNonQuery()

            'Get the Coefficient Set ID of the newly created coefficient set
            Dim cmdNewCoeffId As New DataHelper(String.Format("SELECT COEFFSETID FROM COEFFICIENTSET WHERE COEFFICIENTSET.NAME LIKE '{0}'", strNewCoeffName))
            Dim datanewCoeffId As OleDbDataReader = cmdNewCoeffId.ExecuteReader()
            datanewCoeffId.Read()
            _intCoeffSetID = datanewCoeffId("CoeffSetID")
            datanewCoeffId.Close()

            'Now loopy loo to populate values.
            'Get the coefficient table
            Dim strNewCoeff1 As String
            strNewCoeff1 = "SELECT * FROM COEFFICIENT"
            Dim adaptNewCoeff As New OleDbDataAdapter(strNewCoeff1, g_DBConn)
            Dim dataNewCoeff As New DataTable
            adaptNewCoeff.Fill(dataNewCoeff)

            dataCopySet = cmdCopySet.ExecuteReader

            'Actually add the records to the new set
            While dataCopySet.Read()
                'Add New one
                Dim dataRow As DataRow = dataNewCoeff.NewRow()

                'Add the necessary components
                dataRow("Coeff1") = dataCopySet("Coeff1")
                dataRow("Coeff2") = dataCopySet("Coeff2")
                dataRow("Coeff3") = dataCopySet("Coeff3")
                dataRow("Coeff4") = dataCopySet("Coeff4")
                dataRow("CoeffSetID") = _intCoeffSetID
                dataRow("LCClassID") = dataCopySet("LCClassID")

                adaptNewCoeff.Update(dataNewCoeff)
            End While
            dataCopySet.Close()

            CopyCoefficient = adaptNewCoeff
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Function

    ''' <summary>
    ''' Reclasses the landuse.
    ''' </summary>
    ''' <param name="LUScenItems">which is a collection of the landuse entered.</param>
    ''' <param name="strLCFileName">path to which the landcover grid exists.</param>
    Private Sub ReclassLanduse(ByRef LUScenItems As LandUseItems, ByRef strLCFileName As String)
        'strLCClass: Name of the LCTYPE being altered

        Try
            Dim strOutLandCover As String

            '    'Make sure the landcoverraster exists..it better if they get to this point, ED!
            If g_LandCoverRaster Is Nothing Then
                If RasterExists(strLCFileName) Then
                    _pLandCoverRaster = ReturnRaster(strLCFileName)
                Else
                    Return
                End If
            Else
                _pLandCoverRaster = g_LandCoverRaster
            End If

            Dim booLandScen As Boolean
            Dim pNewLandCoverRaster As New Grid
            strOutLandCover = GetUniqueFileName("landcover", g_XmlPrjFile.ProjectWorkspace, OutputGridExt)

            'Going to now take each entry in the landuse scenarios, if they've choosen 'apply', we
            'will reclass that area of the output raster using reclass raster
            If LUScenItems.Count > 0 Then
                'There's at least one scenario, so copy the input grid to the output as is so that it can be modified
                _pLandCoverRaster.Save(strOutLandCover)
                _pLandCoverRaster.Close()
                pNewLandCoverRaster.Open(strOutLandCover)

                Using progress = New SynchronousProgressDialog("Landuse Scenario", CInt(LUScenItems.Count), g_MainForm)
                    Dim i As Short
                    For i = 0 To LUScenItems.Count - 1
                        If LUScenItems.Item(i).intApply = 1 Then
                            progress.Increment(String.Format("Processing Landuse scenario...{0}", i))
                            If SynchronousProgressDialog.KeepRunning Then
                                ReclassRaster(LUScenItems.Item(i), _strLCClass, pNewLandCoverRaster)
                                booLandScen = True
                            Else
                                pNewLandCoverRaster.Close()
                                booLandScen = False
                                Exit For
                            End If
                        End If
                    Next i
                End Using
            End If

            If booLandScen Then
                g_LandCoverRaster = pNewLandCoverRaster
            End If

        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub ReclassRaster(ByRef LUItem As LandUseItem, ByRef strLCClass As String, ByRef outputGrid As Grid)

        'We're passing over a single land use scenario in the form of the xml
        'class XmlLandUseItem, seems to be the easiest way to do this.

        Dim strSelect As String = String.Format("SELECT LCTYPE.LCTYPEID, LCCLASS.NAME, LCCLASS.VALUE FROM LCTYPE INNER JOIN LCCLASS ON LCTYPE.LCTYPEID = LCCLASS.LCTYPEID WHERE LCTYPE.NAME LIKE '{0}' AND LCCLASS.NAME LIKE '{1}'", strLCClass, LUItem.strLUScenName)
        Dim LCValue As Double
        Dim LUItemDetails As New LandUseMangementScenario
        'The particulars in the landuse

        'Open the landclass Value Value 
        'This is the value user's landclass will change to
        Using cmdLCVal As New OleDbCommand(strSelect, g_DBConn)
            Using readLCVal As OleDbDataReader = cmdLCVal.ExecuteReader()
                readLCVal.Read()
                LCValue = readLCVal("Value")
            End Using
        End Using

        'init the landuse xml stuff
        LUItemDetails.Xml = LUItem.strLUScenXmlFile

        'Convert the polygon featureclass into a Raster for sending back out
        'Get the featureclass, check for selected features
        Dim sf As New Shapefile
        Dim sfIndex As Long = GetLayerIndex(LUItemDetails.strLUScenLyrName)
        Dim shape As MapWinGIS.Shape
        If LUItemDetails.intLUScenSelectedPoly = 1 And MapWindowPlugin.MapWindowInstance.View.SelectedShapes.NumSelected > 0 And sfIndex <> -1 Then
            'Dim lyr As Layer = MapWindowPlugin.MapwindowInstance.Layers (sfIndex)
            Dim exportPath As String = ExportSelectedFeatures(LUItemDetails.strLUScenFileName, LUItemDetails.intLUScenSelectedPolyList)
            shape = ReturnSelectGeometry(exportPath)
            sf = ReturnFeature(exportPath)
        Else
            sf = ReturnFeature(LUItemDetails.strLUScenFileName)
            shape = sf.Shape(0)
        End If

        'classify the output grid cells under the area polygon to the correct value
        'Get minimum extents of the area file
        Dim sfExt As Extents = sf.Extents
        Dim startRow, startCol, endRow, endCol As Integer
        outputGrid.ProjToCell(sfExt.xMin, sfExt.yMax, startCol, startRow)
        outputGrid.ProjToCell(sfExt.xMax, sfExt.yMin, endCol, endRow)

        Dim x, y As Double
        sf.BeginPointInShapefile()
        'cycle and test cell center, then set the appropriate when found
        For row As Integer = startRow To endRow
            For col As Integer = startCol To endCol
                outputGrid.CellToProj(col, row, x, y)
                'If u.PointInPolygon(shape, pnt) Then
                If sf.PointInShapefile(x, y) <> -1 Then
                    outputGrid.Value(col, row) = LCValue
                End If
            Next
        Next
        sf.EndPointInShapefile()
        sf.Close()
        outputGrid.Save()
    End Sub

    Public Sub Cleanup(ByRef dictNames As Dictionary(Of String, String), ByRef PollItems As PollutantItems, ByRef strLCTypeName As String)
        Try
            Dim strDeleteCoeffSet As String
            Dim strCoeffDeleteName As String
            Dim strLCDeleteName As String
            Dim i As Short

            strLCDeleteName = dictNames.Item(strLCTypeName)

            If Len(strLCDeleteName) = 0 Then
                Return
            End If

            Dim strLCTypeDelete As String
            Dim strLCClassDelete As String

            strLCTypeDelete = String.Format("SELECT * FROM LCTYPE WHERE NAME LIKE '{0}'", strLCDeleteName)
            Using cmdDeleteList As New DataHelper(strLCTypeDelete)
                Dim dataDeleteList As OleDbDataReader = cmdDeleteList.ExecuteReader()
                dataDeleteList.Read()
                strLCClassDelete = "Delete * FROM LCCLASS WHERE LCTYPEID =" & dataDeleteList("LCTypeID")
                dataDeleteList.Close()
                Using cmdDelete As New DataHelper(strLCClassDelete)
                    cmdDelete.ExecuteNonQuery()
                End Using

                For i = 0 To PollItems.Count - 1
                    strCoeffDeleteName = dictNames.Item(PollItems.Item(i).strCoeffSet)
                    If Len(strCoeffDeleteName) > 0 Then
                        strDeleteCoeffSet = String.Format("DELETE * FROM COEFFICIENTSET WHERE NAME LIKE '{0}'", strCoeffDeleteName)
                        Using cmdDelete2 = New OleDbCommand(strDeleteCoeffSet, g_DBConn)
                            cmdDelete2.ExecuteNonQuery()
                        End Using
                    End If
                Next
            End Using

        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub
End Module