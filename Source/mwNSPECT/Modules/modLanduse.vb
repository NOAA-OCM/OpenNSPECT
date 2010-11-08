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

Imports System.Data.OleDb
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

    Private _intLCTypeID As Short
    Private _intLCClassID As Short
    Private _intCoeffSetID As Short
    Private _strLCFileName As String

    Private _strLCClass As String
    Private _pLandCoverRaster As MapWinGIS.Grid

    Public g_booLCChange As Boolean
    Public g_strLCTypeName As String 'the temp name of Land Cover type, if indeed landuses are applied.
    Public g_DictTempNames As New Generic.Dictionary(Of String, String) 'Array holding the temp names of the LCType, and subsequent Coefficient Set Names
    ' Constant used by the Error handler function - DO NOT REMOVE
    Const c_sModuleFileName As String = "modLanduse.vb"
    Private m_ParentHWND As Integer ' Set this to get correct parenting of Error handler forms




    Public Sub Begin(ByRef strLCClassName As String, ByRef clsLUScenItems As clsXMLLandUseItems, ByRef dictPollutants As Generic.Dictionary(Of String, String), ByRef strLCFileName As String, ByRef strWorkspace As String)
        'strLCClassName: str of current land cover class
        'clsLUScenItems: XML class that holds the params of the user's Land Use Scenario
        'dictPollutants: dictionary created to hold the pollutants of this particular project
        'strLCFileName: FileName of land cover grid

        Try
            Dim strCurrentLCType As String 'Current LCTYpe
            Dim strCurrentLCClass As String 'Current Landclasses of The LCType
            Dim strInsertTempLCType As String
            Dim strTempLCTypeName As String
            Dim strrsInsertLCType As String 'the inserted LCTYPE
            Dim strrsPermLandClass As String 'the landclass table currently in the database
            Dim strNewLandClass As String 'Newly added landclass - to get its new ID

            Dim i As Short
            Dim j As Short
            Dim k As Short
            Dim intValue As Short 'Temp new landclasses fake value
            Dim clsLUScen As New clsXMLLUScen 'XML Land use scenario
            Dim strCoeffSetTempName As String 'New temp name for the coefficient set
            Dim strCoeffSetOrigName As String 'Original name for the coefficient set
            Dim pollArray As Generic.Dictionary(Of String, String).KeyCollection.Enumerator 'Array to hold pollutants...get itself from the dictionary

            'init the file name of the LandClass File
            _strLCFileName = strLCFileName


            'STEP 1: Get the current LCTYPE
            strCurrentLCType = "SELECT * FROM LCTYPE WHERE NAME LIKE '" & strLCClassName & "'"
            Dim cmdCurrentLCType As New OleDbCommand(strCurrentLCType, g_DBConn)
            Dim dataCurrentLCType As OleDbDataReader = cmdCurrentLCType.ExecuteReader
            dataCurrentLCType.Read()


            'STEP 2: Get the current LCCLASSES of the current LCTYPE
            strCurrentLCClass = "SELECT * FROM LCCLASS WHERE" & " LCTYPEID = " & dataCurrentLCType("LCTypeID") & " ORDER BY LCCLass.Value"
            Dim cmdCurrentLCClass As New OleDbCommand(strCurrentLCClass, g_DBConn)

            'STEP 3: Now INSERT a copy of current LCTYPE
            'First, get a temp name... like CCAPLUTemp1, CCAPLUTemp2 etc
            strTempLCTypeName = modUtil.CreateUniqueName("LCTYPE", strLCClassName & "LUTemp")
            strInsertTempLCType = "INSERT INTO LCTYPE (NAME,DESCRIPTION) VALUES ('" & Replace(strTempLCTypeName, "'", "''") & "', '" & Replace(strTempLCTypeName, "'", "''") & " Description')"

            'Add the name to the Dictionary for storage; used for cleanup and in pollutants
            g_DictTempNames.Add(strLCClassName, strTempLCTypeName)

            'STEP 4: INSERT the copy of the LCTYPE in
            Dim cmdInsertCopy As New OleDbCommand(strInsertTempLCType, g_DBConn)
            cmdInsertCopy.ExecuteNonQuery()

            'STEP 5: Get it back now so you can use its ID for inserting the landclasses
            strrsInsertLCType = "SELECT LCTYPEID FROM LCTYPE " & "WHERE LCTYPE.NAME LIKE '" & strTempLCTypeName & "'"
            Dim cmdInsertType As New OleDbCommand(strrsInsertLCType, g_DBConn)
            Dim dataInsertType As OleDbDataReader = cmdInsertType.ExecuteReader()
            dataInsertType.Read()
            _intLCTypeID = dataInsertType("LCTypeID")
            dataInsertType.Close()

            'STEP 6: Now clone the current landclasses into a new recordset
            Dim cmdCloneLCClass As OleDbCommand = cmdCurrentLCClass.Clone()

            'Prepare the landclass table to accept the copies of landclass
            strrsPermLandClass = "SELECT * FROM LCCLASS"
            Dim adaptPermLC As New OleDbDataAdapter(strrsPermLandClass, g_DBConn)
            Dim dataPermLC As New Data.DataTable
            adaptPermLC.Fill(dataPermLC)

            'STEP 7: loop through the clsXMLLandUseItems to add new land uses to the copy rs
            Dim dataCloneLCClass As OleDbDataReader = cmdCloneLCClass.ExecuteReader()

            'Now add all the landclasses.
            While dataCloneLCClass.Read()
                'Add new
                Dim dataPermRow As Data.DataRow = dataPermLC.NewRow()
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
                intValue = dataCloneLCClass("Value") 'This keeps track of the max
            End While
            dataCloneLCClass.Close()


            'STEP 8: Now add the new landclass
            Dim intLCClassIDs() As Short
            ReDim intLCClassIDs(clsLUScenItems.Count - 1)

            For i = 0 To clsLUScenItems.Count - 1
                If clsLUScenItems.Item(i).intApply = 1 Then
                    'init the fake value: will be max value + 1
                    intValue = intValue + 1

                    'Init the clsLUScen
                    clsLUScen.XML = clsLUScenItems.Item(i).strLUScenXMLFile

                    Dim dataPermRow As Data.DataRow = dataPermLC.NewRow()
                    dataPermRow("LCTypeID") = _intLCTypeID
                    dataPermRow("Value") = intValue
                    dataPermRow("Name") = clsLUScen.strLUScenName
                    dataPermRow("CN-A") = clsLUScen.intSCSCurveA
                    dataPermRow("CN-B") = clsLUScen.intSCSCurveB
                    dataPermRow("CN-C") = clsLUScen.intSCSCurveC
                    dataPermRow("CN-D") = clsLUScen.intSCSCurveD
                    dataPermRow("CoverFactor") = clsLUScen.lngCoverFactor
                    dataPermRow("W_WL") = clsLUScen.intWaterWetlands

                    adaptPermLC.Update(dataPermLC)
                End If

                'Gather the newly added LCClassIds in an array for use later
                strNewLandClass = "SELECT LCCLASSID FROM LCCLASS WHERE NAME LIKE '" & clsLUScen.strLUScenName & "'"
                Dim cmdNewLC As New OleDbCommand(strNewLandClass, g_DBConn)
                Dim dataNewLC As OleDbDataReader = cmdNewLC.ExecuteReader()
                dataNewLC.Read()
                intLCClassIDs(i - 1) = dataNewLC("LCClassID")
                dataNewLC.Close()
            Next i

            'STEP 10: Parse the incoming dictionary
            'Remember, it's in this form:
            'Key_________Item___________
            'Pollutant , CoefficientSet
            Dim l As Short
            pollArray = dictPollutants.Keys.GetEnumerator

            'Now to the pollutants
            'Loop through the pollutants coming from the XML class, as well as those in the project that are being used
            For j = 0 To clsLUScen.clsPollItems.Count - 1
                For k = 0 To dictPollutants.Count - 1
                    If InStr(1, clsLUScen.clsPollItems.Item(j).strPollName, pollArray.Current, CompareMethod.Text) > 0 Then
                        strCoeffSetOrigName = dictPollutants.Item(pollArray.Current) 'Original Name
                        strCoeffSetTempName = modUtil.CreateUniqueName("COEFFICIENTSET", strCoeffSetOrigName & "_Temp") 'Make a new temp name

                        'Now add names of the coefficient sets to the dictionary
                        g_DictTempNames.Add(strCoeffSetOrigName, strCoeffSetTempName)

                        Dim adaptTemp As OleDbDataAdapter = CopyCoefficient(strCoeffSetTempName, strCoeffSetOrigName) 'Call to function that returns the
                        'copied record set.
                        Dim dataTemp As New Data.DataTable
                        adaptTemp.Fill(dataTemp)

                        'Now add the new values using the info in the XML file and the array of new LCClass IDs
                        For l = 0 To UBound(intLCClassIDs)
                            Dim dataRow As Data.DataRow = dataTemp.NewRow()

                            clsLUScen.XML = clsLUScenItems.Item(l).strLUScenXMLFile

                            dataRow("Coeff1").Value = clsLUScen.clsPollItems.Item(j).intType1
                            dataRow("Coeff2").Value = clsLUScen.clsPollItems.Item(j).intType2
                            dataRow("Coeff3").Value = clsLUScen.clsPollItems.Item(j).intType3
                            dataRow("Coeff4").Value = clsLUScen.clsPollItems.Item(j).intType4
                            dataRow("CoeffSetID").Value = _intCoeffSetID
                            dataRow("LCClassID").Value = intLCClassIDs(l)

                            adaptTemp.Update(dataTemp)
                        Next l
                    End If
                    pollArray.MoveNext()
                Next k
            Next j

            g_strLCTypeName = strTempLCTypeName
            ReclassLanduse(clsLUScenItems, strTempLCTypeName, _strLCFileName)


        Catch ex As Exception
            MsgBox("Error Number: " & Err.Number & vbNewLine & "Error Description: " & Err.Description)
        End Try

    End Sub

    Private Function CopyCoefficient(ByRef strNewCoeffName As String, ByRef strCoeffSet As String) As OleDbDataAdapter
        'General gist:  First we add new record to the Coefficient Set table using strNewCoeffName as
        'the name, PollID, LCTYPEID.  Once that's done, we'll add the coefficients
        'from the set being copied

        CopyCoefficient = Nothing
        Try
            Dim strCopySet As String 'The Recordset of existing coefficients being copied
            Dim strNewLcType As String 'CmdString for inserting new coefficientset               '
            Dim strNewCoeffID As String 'Holder for the CoefficientSetID

            strCopySet = "SELECT * FROM COEFFICIENTSET INNER JOIN COEFFICIENT ON COEFFICIENTSET.COEFFSETID = " & "COEFFICIENT.COEFFSETID WHERE COEFFICIENTSET.NAME LIKE '" & strCoeffSet & "'"
            Dim cmdCopySet As New OleDbCommand(strCopySet, g_DBConn)
            Dim dataCopySet As OleDbDataReader = cmdCopySet.ExecuteReader
            dataCopySet.Read()

            'INSERT: new Coefficient set taking the PollID and LCType ID from rsCopySet
            strNewLcType = "INSERT INTO COEFFICIENTSET(NAME, POLLID, LCTYPEID) VALUES ('" & Replace(strNewCoeffName, "'", "''") & "'," & dataCopySet("POLLID") & "," & _intLCTypeID & ")"

            dataCopySet.Close()

            'First need to add the coefficient set to that table
            Dim cmdNewLCType As New OleDbCommand(strNewLcType, g_DBConn)
            cmdNewLCType.ExecuteNonQuery()

            'Get the Coefficient Set ID of the newly created coefficient set
            strNewCoeffID = "SELECT COEFFSETID FROM COEFFICIENTSET " & "WHERE COEFFICIENTSET.NAME LIKE '" & strNewCoeffName & "'"
            Dim cmdNewCoeffId As New OleDbCommand(strNewCoeffID, g_DBConn)
            Dim datanewCoeffId As OleDbDataReader = cmdNewCoeffId.ExecuteReader()
            datanewCoeffId.Read()
            _intCoeffSetID = datanewCoeffId("CoeffSetID")
            datanewCoeffId.Close()

            'Now loopy loo to populate values.
            'Get the coefficient table
            Dim strNewCoeff1 As String
            strNewCoeff1 = "SELECT * FROM COEFFICIENT"
            Dim adaptNewCoeff As New OleDbDataAdapter(strNewCoeff1, g_DBConn)
            Dim dataNewCoeff As New Data.DataTable
            adaptNewCoeff.Fill(dataNewCoeff)

            dataCopySet = cmdCopySet.ExecuteReader

            'Actually add the records to the new set
            While dataCopySet.Read()
                'Add New one
                Dim dataRow As Data.DataRow = dataNewCoeff.NewRow()

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
            MsgBox("Error Number: " & Err.Number & vbNewLine & "Error Description: " & Err.Description, MsgBoxStyle.Critical, "Error in modutil.copycoefficient")
        End Try
    End Function

    Private Sub ReclassLanduse(ByRef clsLUScenItems As clsXMLLandUseItems, ByRef strLCClass As String, ByRef strLCFileName As String)
        'clsLUScenItems: which is a collection of the landuse entered
        'strLCClass: Name of the LCTYPE being altered
        'strLCFileName: path to which the landcover grid exists
        Try
            Dim strOutLandCover As String



            '    'Make sure the landcoverraster exists..it better if they get to this point, ED!
            If g_LandCoverRaster Is Nothing Then
                If modUtil.RasterExists(strLCFileName) Then
                    _pLandCoverRaster = modUtil.ReturnRaster(strLCFileName)
                Else
                    Exit Sub
                End If
            Else
                _pLandCoverRaster = g_LandCoverRaster
            End If

            Dim booLandScen As Boolean
            Dim pNewLandCoverRaster As New MapWinGIS.Grid
            strOutLandCover = modUtil.GetUniqueName("landcover", IO.Path.GetDirectoryName(_pLandCoverRaster.Filename), ".bgd")

            'Going to now take each entry in the landuse scenarios, if they've choosen 'apply', we
            'will reclass that area of the output raster using reclass raster
            Dim i As Short
            If clsLUScenItems.Count > 0 Then
                'There's at least one scenario, so copy the input grid to the output as is so that it can be modified
                _pLandCoverRaster.Save(strOutLandCover)
                _pLandCoverRaster.Close()
                pNewLandCoverRaster.Open(strOutLandCover)

                For i = 0 To clsLUScenItems.Count - 1
                    If clsLUScenItems.Item(i).intApply = 1 Then
                        modProgDialog.ProgDialog("Processing Landuse scenario...", "Landuse Scenario", 0, CInt(clsLUScenItems.Count), CInt(i), g_frmProjectSetup)
                        If modProgDialog.g_boolCancel Then
                            ReclassRaster(clsLUScenItems.Item(i), _strLCClass, pNewLandCoverRaster)
                            booLandScen = True
                        Else
                            pNewLandCoverRaster.Close()
                            booLandScen = False
                            Exit For
                        End If
                    End If
                Next i
            End If

            If booLandScen Then
                g_LandCoverRaster = pNewLandCoverRaster
            End If

            modProgDialog.KillDialog()

        Catch ex As Exception
            MsgBox("error in MSSetup " & Err.Number & ": " & Err.Description)
        End Try
    End Sub

    Private Sub ReclassRaster(ByRef clsLUItem As clsXMLLandUseItem, ByRef strLCClass As String, ByRef outputGrid As MapWinGIS.Grid)

        'We're passing over a single land use scenario in the form of the xml
        'class clsXMLLandUseItem, seems to be the easiest way to do this.

        Dim strSelect As String 'OLEDB selections string
        Dim LCValue As Double 'value
        Dim clsLUItemDetails As New clsXMLLUScen 'The particulars in the landuse

        'Open the landclass Value Value 
        'This is the value user's landclass will change to
        strSelect = "SELECT LCTYPE.LCTYPEID, LCCLASS.NAME, LCCLASS.VALUE FROM " & "LCTYPE INNER JOIN LCCLASS ON LCTYPE.LCTYPEID = LCCLASS.LCTYPEID " & "WHERE LCTYPE.NAME LIKE '" & strLCClass & "' AND LCCLASS.NAME LIKE '" & clsLUItem.strLUScenName & "'"
        Dim cmdLCVal As New OleDbCommand(strSelect, modUtil.g_DBConn)
        Dim readLCVal As OleDbDataReader = cmdLCVal.ExecuteReader()
        readLCVal.Read()
        LCValue = readLCVal("Value")
        readLCVal.Close()

        'init the landuse xml stuff
        clsLUItemDetails.XML = clsLUItem.strLUScenXMLFile

        'Convert the polygon featureclass into a Raster for sending back out
        'Get the featureclass, check for selected features
        Dim sf As New MapWinGIS.Shapefile
        Dim sfIndex As Long = modUtil.GetLayerIndex(clsLUItemDetails.strLUScenLyrName)
        If clsLUItemDetails.intLUScenSelectedPoly = 1 And g_MapWin.View.SelectedShapes.NumSelected > 0 And sfIndex <> -1 Then
            Dim lyr As MapWindow.Interfaces.Layer = g_MapWin.Layers(sfIndex)
            'TODO: This somehow needs to be modified to take into account that it may not be the current layer selection
            'Dim exportPath As String = modUtil.ExportSelectedFeatures(lyr)
            'sf.Open(exportPath)        
            sf = modUtil.ReturnFeature(clsLUItemDetails.strLUScenFileName)
        Else
            sf = modUtil.ReturnFeature(clsLUItemDetails.strLUScenFileName)
        End If


        'classify the output grid cells under the area polygon to the correct value
        'Get minimum extents of the area file
        Dim sfExt As MapWinGIS.Extents = sf.Extents
        Dim startRow, startCol, endRow, endCol As Integer
        outputGrid.ProjToCell(sfExt.xMin, sfExt.yMin, startCol, startRow)
        outputGrid.ProjToCell(sfExt.xMax, sfExt.yMax, endCol, endRow)

        Dim x, y As Double
        'cycle and test cell center, then set the appropriate when found
        For row As Integer = startRow To endRow
            For col As Integer = startCol To endCol
                outputGrid.CellToProj(col, row, x, y)
                If sf.PointInShapefile(x, y) <> -1 Then
                    outputGrid.Value(col, row) = LCValue
                End If
            Next
        Next
        sf.Close()
    End Sub

    Public Sub Cleanup(ByRef dictNames As Generic.Dictionary(Of String, String), ByRef clsPollItems As clsXMLPollutantItems, ByRef strLCTypeName As String)
        Try
            Dim strDeleteCoeffSet As String
            Dim strCoeffDeleteName As String
            Dim strLCDeleteName As String
            Dim i As Short

            strLCDeleteName = dictNames.Item(strLCTypeName)

            If Len(strLCDeleteName) = 0 Then
                Exit Sub
            End If

            Dim strLCTypeDelete As String
            Dim strLCClassDelete As String

            strLCTypeDelete = "SELECT * FROM LCTYPE WHERE NAME LIKE '" & strLCDeleteName & "'"
            Dim cmdDeleteList As New OleDbCommand(strLCTypeDelete, g_DBConn)
            Dim dataDeleteList As OleDbDataReader = cmdDeleteList.ExecuteReader()
            dataDeleteList.Read()
            strLCClassDelete = "Delete * FROM LCCLASS WHERE LCTYPEID =" & dataDeleteList("LCTypeID")
            dataDeleteList.Close()
            Dim cmdDelete As New OleDbCommand(strLCClassDelete, g_DBConn)
            cmdDelete.ExecuteNonQuery()

            For i = 0 To clsPollItems.Count - 1
                strCoeffDeleteName = dictNames.Item(clsPollItems.Item(i).strCoeffSet)
                If Len(strCoeffDeleteName) > 0 Then
                    strDeleteCoeffSet = "DELETE * FROM COEFFICIENTSET WHERE NAME LIKE '" & strCoeffDeleteName & "'"
                    cmdDelete = New OleDbCommand(strDeleteCoeffSet, g_DBConn)
                    cmdDelete.ExecuteNonQuery()
                End If
            Next i

        Catch ex As Exception
            HandleError(True, "Cleanup " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Sub
End Module