'********************************************************************************************************
'File Name: modMgmtScen.vb
'Description: Functions handling the management scenario additions of the model
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
Imports MapWinGIS
Imports System.Data.OleDb
Imports OpenNspect.Xml

Module ManagementScenarios
    ' *************************************************************************************
    ' *  Perot Systems Government Services
    ' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
    ' *  modMgmtScen
    ' *************************************************************************************
    ' *  Description: Code for handling the Management Scenarios
    ' *
    ' *
    ' *  Called By:
    ' *************************************************************************************

    Private _strLCClass As String
    Private _pLandCoverRaster As Grid
    Public g_booLCChange As Boolean

    Public Sub MgmtScenSetup(ByRef MgmtScens As ManagementScenarioItems, ByRef strLCClass As String, ByRef strLCFileName As String)
        'Main Sub for setting everything up
        'MgmtScens: Xml wrapper for the management scenarios created by the user
        'strLCClass: Name of the LandCover being used, CCAP
        'strLCFileName: filename of location of LandCover file
        Try
            Dim strOutLandCover As String
            Dim booLandScen As Boolean
            Dim pNewLandCoverRaster As New Grid

            'init everything
            _strLCClass = strLCClass

            'Make sure the landcoverraster exists..it better if they get to this point, ED!
            If RasterExists(strLCFileName) Then
                _pLandCoverRaster = ReturnRaster(strLCFileName)
            Else
                Return
            End If

            strOutLandCover = GetUniqueFileName("landcover", g_strWorkspace, OutputGridExt)

            'Going to now take each entry in the landuse scenarios, if they've choosen 'apply', we
            'will reclass that area of the output raster using reclass raster
            Dim i As Short
            If MgmtScens.Count > 0 Then
                'There's at least one scenario, so copy the input grid to the output as is so that it can be modified
                _pLandCoverRaster.Save(strOutLandCover)
                _pLandCoverRaster.Close()
                pNewLandCoverRaster.Open(strOutLandCover)

                For i = 0 To MgmtScens.Count - 1
                    If MgmtScens.Item(i).intApply = 1 Then
                        ShowProgress("Adding new landclass...", "Creating Management Scenario", CInt(MgmtScens.Count), CInt(i), g_MainForm)
                        If g_KeepRunning Then
                            Dim mgmtitem As ManagementScenarioItem = MgmtScens.Item(i)
                            ReclassRaster(mgmtitem, _strLCClass, pNewLandCoverRaster)
                            booLandScen = True
                        Else
                            pNewLandCoverRaster.Close()
                            booLandScen = False
                            Exit For
                        End If
                    End If
                Next i
            End If

            If Not booLandScen Then
                g_LandCoverRaster = _pLandCoverRaster
            Else
                g_LandCoverRaster = pNewLandCoverRaster
            End If

            CloseProgressDialog()

        Catch ex As Exception
            MsgBox("error in MSSetup " & Err.Number & ": " & Err.Description)
            CloseProgressDialog()
        End Try

    End Sub

    Public Sub ReclassRaster(ByRef MgmtScen As ManagementScenarioItem, ByVal strLCClass As String, ByRef outputGrid As Grid)
        'We're passing over a single management scenarios in the form of the xml
        'class XmlmgmtScenItem, seems to be the easiest way to do this.
        Dim strSelect As String
        'OLEDB selections string
        Dim LCValue As Double
        'value

        'Open the landclass Value Value 
        'This is the value user's landclass will change to
        strSelect = "SELECT LCTYPE.LCTYPEID, LCCLASS.NAME, LCCLASS.VALUE FROM " & "LCTYPE INNER JOIN LCCLASS ON LCTYPE.LCTYPEID = LCCLASS.LCTYPEID " & "WHERE LCTYPE.NAME LIKE '" & strLCClass & "' AND LCCLASS.NAME LIKE '" & MgmtScen.strChangeToClass & "'"
        Dim cmdLCVal As New OleDbCommand(strSelect, g_DBConn)
        Dim readLCVal As OleDbDataReader = cmdLCVal.ExecuteReader()
        readLCVal.Read()
        LCValue = readLCVal("Value")
        readLCVal.Close()

        'classify the output grid cells under the area polygon to the correct value
        Dim sf As New Shapefile
        sf = ReturnFeature(MgmtScen.strAreaFileName)

        'Get minimum extents of the area file
        Dim sfExt As Extents = sf.Extents
        Dim startRow, startCol, endRow, endCol As Integer
        outputGrid.ProjToCell(sfExt.xMin, sfExt.yMax, startCol, startRow)
        outputGrid.ProjToCell(sfExt.xMax, sfExt.yMin, endCol, endRow)

        Dim x, y As Double
        Dim pnt As New Point
        sf.BeginPointInShapefile()
        'cycle and test cell center, then set the appropriate when found
        For row As Integer = startRow To endRow
            For col As Integer = startCol To endCol
                outputGrid.CellToProj(col, row, x, y)
                pnt.x = x
                pnt.y = y
                'If u.PointInPolygon(sf.Shape(0), pnt) Then
                If sf.PointInShapefile(x, y) <> -1 Then
                    outputGrid.Value(col, row) = LCValue
                End If
            Next
        Next
        sf.EndPointInShapefile()
        sf.Close()
        outputGrid.Save()
    End Sub
End Module