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

    Public g_LandUse_LCTypeName As String
    'the temp name of Land Cover type, if indeed landuses are applied.
    Public g_LandUse_DictTempNames As New Dictionary(Of String, String)
    'Array holding th'AddType, and subsequent Coefficient Set Names
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
            Dim currentLCType As String
            Dim strCurrentLCClass As String
            'Current Landclasses of The LCType
            Dim strInsertTempLCType As String
            Dim strrsInsertLCType As String
            'the inserted LCTYPE
            Dim strNewLandClass As String
            'Newly added landclass - to get its new ID

            Dim maxValue As Short
            'Temp new landclasses fake value
            Dim LUScen As New LandUseMangementScenario
            'Xml Land use scenario
            Dim coeffSetTempName As String
            'New temp name for the coefficient set
            Dim coeffSetOrigName As String
            'Original name for the coefficient set
            Dim pollArray As Dictionary(Of String, String).KeyCollection.Enumerator
            'Array to hold pollutants...get itself from the dictionary

            'init the file name of the LandClass File
            _strLCFileName = strLCFileName

            'STEP 1: Get the current LCTYPE
            Dim lcTypeId As Object
            currentLCType = String.Format("SELECT * FROM LCTYPE WHERE NAME LIKE '{0}'", strLCClassName)
            Using cmdCurrentLCType As New DataHelper(currentLCType)
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
            Dim tempLandClassTypeName As String = CreateUniqueName("LCTYPE", strLCClassName & "LUTemp")
            strInsertTempLCType = String.Format("INSERT INTO LCTYPE (NAME,DESCRIPTION) VALUES ('{0}', '{0} Description')", SqlEncode(tempLandClassTypeName))

            'Add the name to the Dictionary for storage; used for cleanup and in pollutants
            g_LandUse_DictTempNames(strLCClassName) = tempLandClassTypeName

            'STEP 4: INSERT the copy of the LCTYPE in
            Using cmdInsertCopy As New DataHelper(strInsertTempLCType)
                cmdInsertCopy.ExecuteNonQuery()
            End Using

            'STEP 5: Get it back now so you can use its ID for inserting the landclasses
            strrsInsertLCType = String.Format("SELECT LCTYPEID FROM LCTYPE WHERE LCTYPE.NAME LIKE '{0}'", tempLandClassTypeName)
            Using cmdInsertType As New DataHelper(strrsInsertLCType)
                Using dataInsertType As OleDbDataReader = cmdInsertType.ExecuteReader()
                    dataInsertType.Read()
                    _intLCTypeID = dataInsertType("LCTypeID")
                End Using
            End Using

            'STEP 6: Now clone the current landclasses into a new recordset
            Dim cmdCloneLCClass As OleDbCommand = cmdCurrentLCClass.CloneCommand()

            'Prepare the landclass table to accept the copies of landclass
            Dim adaptPermLC As New DataHelper("SELECT * FROM LCCLASS")
            Dim adaptNewCoeff = adaptPermLC.GetAdapter()
            ' creating this builder allows us to update.            
            Dim cbuilder As New OleDbCommandBuilder(adaptNewCoeff) With {.QuotePrefix = "[", .QuoteSuffix = "]"}
            Dim dt As New DataTable
            adaptNewCoeff.Fill(dt)

            'STEP 7: loop through the XmlLandUseItems to add new land uses to the copy rs
            Using dataCloneLCClass As OleDbDataReader = cmdCloneLCClass.ExecuteReader()
                'Now add all the landclasses.
                While dataCloneLCClass.Read()
                    'Add new
                    Dim dataPermRow As DataRow = dt.NewRow()
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
                    dt.Rows.Add(dataPermRow)

                    'move to next record
                    maxValue = dataCloneLCClass("Value")
                End While

                adaptNewCoeff.Update(dt)
                dataCloneLCClass.Close()
            End Using

            'STEP 8: Now add the new landclass
            Dim intLCClassIDs(LUScenItems.Count - 1) As Short

            For i = 0 To LUScenItems.Count - 1
                If LUScenItems.Item(i).Enabled Then

                    'Init the LUScen
                    LUScen.Xml = LUScenItems.Item(i).strLUScenXmlFile

                    Dim dataPermRow As DataRow = dt.NewRow()
                    dataPermRow("LCTypeID") = _intLCTypeID
                    dataPermRow("Value") = maxValue + 1 ' the max +1
                    dataPermRow("Name") = LUScen.strLUScenName
                    dataPermRow("CN-A") = LUScen.intSCSCurveA
                    dataPermRow("CN-B") = LUScen.intSCSCurveB
                    dataPermRow("CN-C") = LUScen.intSCSCurveC
                    dataPermRow("CN-D") = LUScen.intSCSCurveD
                    dataPermRow("CoverFactor") = LUScen.lngCoverFactor
                    dataPermRow("W_WL") = LUScen.intWaterWetlands
                    dt.Rows.Add(dataPermRow)
                    adaptNewCoeff.Update(dt)


                    'Gather the newly added LCClassIds in an array for use later
                    'This code assumes that there is only one scenario by that name
                    ' I've added in order by LCCLASSID DESC so that it gets the last id for something by that name.
                    strNewLandClass = String.Format("SELECT LCCLASSID FROM LCCLASS WHERE NAME LIKE '{0}' order by LCCLASSID DESC", LUScen.strLUScenName)
                    Using cmdNewLC As New DataHelper(strNewLandClass)
                        Using dataNewLC As OleDbDataReader = cmdNewLC.ExecuteReader()
                            dataNewLC.Read()
                            If dataNewLC.HasRows Then
                                intLCClassIDs(i) = dataNewLC("LCClassID")
                            End If
                        End Using
                    End Using

                End If
            Next i

            'STEP 10: Parse the incoming dictionary
            'Remember, it's in this form:
            'Key_________Item__________
            'Pollutant , CoefficientSet

            'Now to the pollutants
            'Loop through the pollutants coming from the Xml class, as well as those in the project that are being used
            For j = 0 To LUScen.PollItems.Count - 1
                pollArray = dictPollutants.Keys.GetEnumerator
                pollArray.MoveNext()
                For k = 0 To dictPollutants.Count - 1
                    If InStr(1, LUScen.PollItems.Item(j).strPollName, pollArray.Current, CompareMethod.Text) > 0 AndAlso pollArray.Current IsNot Nothing Then
                        coeffSetOrigName = dictPollutants.Item(pollArray.Current)
                        coeffSetTempName = CreateUniqueName("COEFFICIENTSET", coeffSetOrigName & "_Temp")

                        'Now add names of the coefficient sets to the dictionary
                        g_LandUse_DictTempNames(coeffSetOrigName) = coeffSetTempName

                        Dim adaptTemp As OleDbDataAdapter = CopyCoefficient(coeffSetTempName, coeffSetOrigName)
                        'Call to function that returns the copied record set.
                        ' creating this builder allows us to update.
                        Dim cbuilder2 As New OleDbCommandBuilder(adaptTemp) With {.QuotePrefix = "[", .QuoteSuffix = "]"}
                        Dim dataTemp As New DataTable
                        adaptTemp.Fill(dataTemp)

                        'Now add the new values using the info in the Xml file and the array of new LCClass IDs
                        For l = 0 To UBound(intLCClassIDs)
                            Dim dataRow As DataRow = dataTemp.NewRow()

                            LUScen.Xml = LUScenItems.Item(l).strLUScenXmlFile

                            dataRow("Coeff1") = LUScen.PollItems.Item(j).intType1
                            dataRow("Coeff2") = LUScen.PollItems.Item(j).intType2
                            dataRow("Coeff3") = LUScen.PollItems.Item(j).intType3
                            dataRow("Coeff4") = LUScen.PollItems.Item(j).intType4
                            dataRow("CoeffSetID") = _intCoeffSetID
                            dataRow("LCClassID") = intLCClassIDs(0)
                            dataTemp.Rows.Add(dataRow)

                        Next l
                        adaptTemp.Update(dataTemp)
                    End If
                    pollArray.MoveNext()
                Next k
            Next j

            g_LandUse_LCTypeName = tempLandClassTypeName

            ReclassLanduse(LUScenItems, _strLCFileName)

        Catch ex As Exception
            HandleError(ex)
        End Try

    End Sub

    Private Function SqlEncode(ByVal value As String) As String
        Return Replace(value, "'", "''")
    End Function
    ''' <summary>
    ''' Copies the coefficient. First we add new record to the Coefficient Set table using strNewCoeffName as the name, PollID, LCTYPEID.  Once that's done, we'll add the coefficients from the set being copied
    ''' </summary>
    ''' <param name="newCoeffName">Name of the new coeff.</param>
    ''' <param name="coeffSet">The coeff set.</param><returns></returns>
    Private Function CopyCoefficient(ByRef newCoeffName As String, ByRef coeffSet As String) As OleDbDataAdapter
        Try
            'The Recordset of existing coefficients being copied
            Dim cmdCopySet As New DataHelper(String.Format("SELECT * FROM COEFFICIENTSET INNER JOIN COEFFICIENT ON COEFFICIENTSET.COEFFSETID = COEFFICIENT.COEFFSETID WHERE COEFFICIENTSET.NAME LIKE '{0}'", coeffSet))
            Dim dataCopySet As OleDbDataReader = cmdCopySet.ExecuteReader
            dataCopySet.Read()

            'INSERT: new Coefficient set taking the PollID and LCType ID from rsCopySet
            'First need to add the coefficient set to that table
            Dim cmdNewLCType As New DataHelper(String.Format("INSERT INTO COEFFICIENTSET(NAME, POLLID, LCTYPEID) VALUES ('{0}',{1},{2})", Replace(newCoeffName, "'", "''"), dataCopySet("POLLID"), _intLCTypeID))
            cmdNewLCType.ExecuteNonQuery()
            dataCopySet.Close()

            'Get the Coefficient Set ID of the newly created coefficient set
            Dim cmdNewCoeffId As New DataHelper(String.Format("SELECT COEFFSETID FROM COEFFICIENTSET WHERE COEFFICIENTSET.NAME LIKE '{0}'", newCoeffName))
            Dim datanewCoeffId As OleDbDataReader = cmdNewCoeffId.ExecuteReader()
            datanewCoeffId.Read()
            _intCoeffSetID = datanewCoeffId("CoeffSetID")
            datanewCoeffId.Close()

            'Now loopy loo to populate values.
            'Get the coefficient table
            Dim adaptNewCoeff As New OleDbDataAdapter("SELECT * FROM COEFFICIENT", g_DBConn)
            ' creating this builder allows us to update.
            Dim builder As New OleDbCommandBuilder(adaptNewCoeff) With {.QuotePrefix = "[", .QuoteSuffix = "]"}
            Dim dataNewCoeff As New DataTable
            adaptNewCoeff.Fill(dataNewCoeff)

            'The Recordset of existing coefficients being copied
            Dim cmdCopySet2 As New DataHelper(String.Format("SELECT * FROM COEFFICIENTSET INNER JOIN COEFFICIENT ON COEFFICIENTSET.COEFFSETID = COEFFICIENT.COEFFSETID WHERE COEFFICIENTSET.NAME LIKE '{0}'", coeffSet))
            Dim dataCopySet2 As OleDbDataReader = cmdCopySet2.ExecuteReader

            'Actually add the records to the new set
            While dataCopySet2.Read()
                'Add New one
                Dim dataRow As DataRow = dataNewCoeff.NewRow()

                'Add the necessary components
                dataRow("Coeff1") = dataCopySet2("Coeff1")
                dataRow("Coeff2") = dataCopySet2("Coeff2")
                dataRow("Coeff3") = dataCopySet2("Coeff3")
                dataRow("Coeff4") = dataCopySet2("Coeff4")
                dataRow("CoeffSetID") = _intCoeffSetID
                dataRow("LCClassID") = dataCopySet2("LCClassID")

                dataNewCoeff.Rows.Add(dataRow)
            End While
            adaptNewCoeff.Update(dataNewCoeff)
            dataCopySet2.Close()

            Return adaptNewCoeff
        Catch ex As Exception
            HandleError(ex)
            Return Nothing
        End Try
    End Function

    Public Function GetLandCoverRaster(ByRef landCoverFileName As String) As Grid
        If g_LandCoverRaster IsNot Nothing Then
            Return g_LandCoverRaster
        End If

        If RasterExists(landCoverFileName) Then
            Return ReturnRaster(landCoverFileName)
        End If

        Return Nothing
    End Function
    ''' <summary>
    ''' Reclasses the landuse.
    ''' </summary>
    ''' <param name="LUScenItems">which is a collection of the landuse entered.</param>
    ''' <param name="landCoverFileName">path to which the landcover grid exists.</param>
    Private Sub ReclassLanduse(ByRef LUScenItems As LandUseItems, ByRef landCoverFileName As String)

        ' categorizes the all classes

        'Dim landCoverName As String = ""
        Dim landCoverRaster = GetLandCoverRaster(landCoverFileName)
        If landCoverRaster Is Nothing Then
            Throw New ArgumentException("cannot get landCoverRaster.")
        End If

        Try
            Dim strOutLandCover As String = GetUniqueFileName("landcover", g_Project.ProjectWorkspace, OutputGridExt)

            'Going to now take each entry in the landuse scenarios, if they've choosen 'apply', we
            'will reclass that area of the output raster using reclass raster
            Dim booLandScen As Boolean
            Dim pNewLandCoverRaster As New Grid
            If LUScenItems.Count > 0 Then
                'There's at least one scenario, so copy the input grid to the output as is so that it can be modified
                landCoverRaster.Save()  'Saving because it seemed necessary.
                '_____________________________________________________________________________________
                ' New way for creating tmp grid to try to fix updating of min/max issue.  DLE 8/29/2012
                Dim Flg As Boolean
                 Flg = pNewLandCoverRaster.CreateNew(strOutLandCover, landCoverRaster.Header, GridDataType.DoubleDataType, landCoverRaster.Header.NodataValue, False, GridFileType.GeoTiff)
                If Flg = False Then
                    MsgBox("ERROR in tmp LC grid initialization: " & strOutLandCover)
                End If
                ' Grid created, now populate it
                'pNewLandCoverRaster = landCoverRaster 'This method does not work: need to populate values individually.
                For icol As Integer = 0 To landCoverRaster.Header.NumberCols - 1
                    For jrow As Integer = 0 To landCoverRaster.Header.NumberRows - 1
                        pNewLandCoverRaster.Value(icol, jrow) = landCoverRaster.Value(icol, jrow)
                    Next
                Next
                '_________________________________  End new creation method ______________________
 
                Using progress = New SynchronousProgressDialog("Landuse Scenario", CInt(LUScenItems.Count), g_MainForm)
                    Dim i As Short
                    For i = 0 To LUScenItems.Count - 1
                        If LUScenItems.Item(i).Enabled Then
                            progress.Increment(String.Format("Processing Landuse scenario...{0}", i))
                            If SynchronousProgressDialog.KeepRunning Then
                                'TODO: This looks like a bug: landcovername is used to restrict the following select statement.
                                'Instead the select statement should allow anything.
                                'ReclassRaster(LUScenItems.Item(i), landCoverName, pNewLandCoverRaster)
                                ReclassRaster(LUScenItems.Item(i), pNewLandCoverRaster)
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

            'probably not missing an else clause
            If booLandScen Then
                g_LandCoverRaster = pNewLandCoverRaster
            End If

        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Public Function GetLandClassValue(name As String) As Double
        Dim strSelect As String = String.Format("SELECT LCTYPE.LCTYPEID, LCCLASS.NAME, LCCLASS.VALUE FROM LCTYPE INNER JOIN LCCLASS ON LCTYPE.LCTYPEID = LCCLASS.LCTYPEID WHERE LCCLASS.NAME LIKE '{0}'", name)
        Dim LCValue As Double


        'Open the landclass Value Value
        'This is the value user's landclass will change to
        Using cmdLCVal As New OleDbCommand(strSelect, g_DBConn)
            Using readLCVal As OleDbDataReader = cmdLCVal.ExecuteReader()
                readLCVal.Read()
                'HACK: this hasRows prevents an exception, but may cause a problem.
                If readLCVal.HasRows Then
                    LCValue = readLCVal("Value")
                End If
            End Using
        End Using
        Return LCValue
    End Function
    Private Sub ReclassRaster(ByRef LUItem As LandUseItem, ByRef outputGrid As Grid)

        'We're passing over a single land use scenario in the form of the xml
        'class XmlLandUseItem, seems to be the easiest way to do this.

        Dim LCValue As Double = GetLandClassValue(LUItem.strLUScenName)

        Dim LUItemDetails As New LandUseMangementScenario
        'The particulars in the landuse

        'init the landuse xml stuff
        LUItemDetails.Xml = LUItem.strLUScenXmlFile

        'Convert the polygon featureclass into a Raster for sending back out
        'Get the featureclass, check for selected features
        Dim sf As New Shapefile
        Dim sfIndex As Long = GetLayerIndex(LUItemDetails.strLUScenLyrName)
        Dim shape As Shape
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
        'MsgBox("Before Reclass: Maximum is" & outputGrid.Maximum.ToString & " should become " & LCValue.ToString)
        For row As Integer = startRow To endRow
            For col As Integer = startCol To endCol
                outputGrid.CellToProj(col, row, x, y)
                'If u.PointInPolygon(shape, pnt) Then
                If sf.PointInShapefile(x, y) <> -1 Then
                    outputGrid.Value(col, row) = LCValue
                    'outputGrid.Value(col, row) = 2
                    'Issue 20914: Set changed area to an existing land cover to see if this produces a result.  IT DOES 7/30/2012.
                    ' This makes me suspect the problem is a ...Grid.Maximum issue
                End If
            Next
        Next
        sf.EndPointInShapefile()
        sf.Close()
        outputGrid.Save()
        'MsgBox("New Maximum is" & outputGrid.Maximum.ToString) ' IS This is updated correctly ?
    End Sub

    Public Sub Cleanup(ByRef dictNames As Dictionary(Of String, String), ByRef PollItems As PollutantItems, ByRef strLCTypeName As String)
        Try

            Dim strLCDeleteName As String = dictNames.Item(strLCTypeName)

            If Len(strLCDeleteName) = 0 Then
                Return
            End If

            Using cmdDeleteList As New DataHelper(String.Format("SELECT * FROM LCTYPE WHERE NAME LIKE '{0}'", strLCDeleteName))
                Using dataDeleteList As OleDbDataReader = cmdDeleteList.ExecuteReader()
                    dataDeleteList.Read()
                    Using cmdDelete As New DataHelper("Delete * FROM LCCLASS WHERE LCTYPEID =" & dataDeleteList("LCTypeID"))
                        cmdDelete.ExecuteNonQuery()
                    End Using
                End Using
                'TODO: This section throwing errors.  Name not getting populated and items left in the database. DLE 8/30/2012
                For i = 0 To PollItems.Count - 1
                    Dim name As String = PollItems.Item(i).strCoeffSet
                    If dictNames.ContainsKey(name) Then
                        Dim strCoeffDeleteName As String
                        strCoeffDeleteName = dictNames.Item(name)
                        If Len(strCoeffDeleteName) > 0 Then
                            Using cmdDelete2 = New OleDbCommand(String.Format("DELETE * FROM COEFFICIENTSET WHERE NAME LIKE '{0}'", strCoeffDeleteName), g_DBConn)
                                cmdDelete2.ExecuteNonQuery()
                            End Using
                        End If
                    End If

                Next
            End Using

        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub
End Module