Imports System.Data.OleDb
Module modMgmtScen
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
    Private _pLandCoverRaster As MapWinGIS.Grid
    Public g_booLCChange As Boolean

    Public Sub MgmtScenSetup(ByRef clsMgmtScens As clsXMLMgmtScenItems, ByRef strLCClass As String, ByRef strLCFileName As String, ByRef strWorkspace As String)
        'Main Sub for setting everything up
        'clsMgmtScens: XML wrapper for the management scenarios created by the user
        'strLCClass: Name of the LandCover being used, CCAP
        'strLCFileName: filename of location of LandCover file
        Try
            Dim strOutLandCover As String
            Dim booLandScen As Boolean
            Dim pNewLandCoverRaster As New MapWinGIS.Grid

            'init everything
            _strLCClass = strLCClass

            'Make sure the landcoverraster exists..it better if they get to this point, ED!
            If modUtil.RasterExists(strLCFileName) Then
                _pLandCoverRaster = modUtil.ReturnRaster(strLCFileName)
            Else
                Exit Sub
            End If

            strOutLandCover = modUtil.GetUniqueName("landcover", IO.Path.GetDirectoryName(_pLandCoverRaster.Filename), ".bgd")

            'Going to now take each entry in the landuse scenarios, if they've choosen 'apply', we
            'will reclass that area of the output raster using reclass raster
            Dim i As Short
            If clsMgmtScens.Count > 0 Then
                'There's at least one scenario, so copy the input grid to the output as is so that it can be modified
                _pLandCoverRaster.Save(strOutLandCover)
                _pLandCoverRaster.Close()
                pNewLandCoverRaster.Open(strOutLandCover)

                For i = 0 To clsMgmtScens.Count - 1
                    If clsMgmtScens.Item(i).intApply = 1 Then
                        modProgDialog.ProgDialog("Adding new landclass...", "Creating Management Scenario", 0, CInt(clsMgmtScens.Count), CInt(i), 0)
                        If modProgDialog.g_boolCancel Then
                            ReclassRaster(clsMgmtScens.Item(i), _strLCClass, pNewLandCoverRaster)
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

            modProgDialog.KillDialog()

        Catch ex As Exception
            MsgBox("error in MSSetup " & Err.Number & ": " & Err.Description)
            modProgDialog.KillDialog()
        End Try

    End Sub

    Public Sub ReclassRaster(ByRef clsMgmtScen As clsXMLMgmtScenItem, ByVal strLCClass As String, ByRef outputGrid As MapWinGIS.Grid)
        'We're passing over a single management scenarios in the form of the xml
        'class clsXMLmgmtScenItem, seems to be the easiest way to do this.
        Dim strSelect As String 'OLEDB selections string
        Dim LCValue As Double 'value

        'Open the landclass Value Value 
        'This is the value user's landclass will change to
        strSelect = "SELECT LCTYPE.LCTYPEID, LCCLASS.NAME, LCCLASS.VALUE FROM " & "LCTYPE INNER JOIN LCCLASS ON LCTYPE.LCTYPEID = LCCLASS.LCTYPEID " & "WHERE LCTYPE.NAME LIKE '" & strLCClass & "' AND LCCLASS.NAME LIKE '" & clsMgmtScen.strChangeToClass & "'"
        Dim cmdLCVal As New OleDbCommand(strSelect, modUtil.g_DBConn)
        Dim readLCVal As OleDbDataReader = cmdLCVal.ExecuteReader()
        readLCVal.Read()
        LCValue = readLCVal("Value")
        readLCVal.Close()

        'classify the output grid cells under the area polygon to the correct value
        Dim sf As New MapWinGIS.Shapefile
        sf = modUtil.ReturnFeature(clsMgmtScen.strAreaFileName)

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
End Module