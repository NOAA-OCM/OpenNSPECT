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

    ''' <summary>
    ''' setting everything up
    ''' </summary>
    ''' <param name="MgmtScens">Xml wrapper for the management scenarios created by the user.</param>
    ''' <param name="strLCClass">Name of the LandCover being used, CCAP.</param>
    ''' <param name="landCoverFileName">filename of location of LandCover file.</param>
    Public Sub MgmtScenSetup(ByRef MgmtScens As ManagementScenarioItems, ByRef strLCClass As String, ByRef landCoverFileName As String)

        'only categorizes the particular strLCClass in the managementScenario

        Dim landCoverRaster = GetLandCoverRaster(landCoverFileName)
        If landCoverRaster Is Nothing Then
            Throw New ArgumentException("cannot get landCoverRaseter.")
        End If
        Dim landCoverName As String
        Try
            'init everything
            landCoverName = strLCClass

            Dim strOutLandCover As String = GetUniqueFileName("landcover", g_Project.ProjectWorkspace, OutputGridExt)

            'Going to now take each entry in the landuse scenarios, if they've choosen 'apply', we
            'will reclass that area of the output raster using reclass raster
            Dim booLandScen As Boolean
            Dim pNewLandCoverRaster As New Grid
            If MgmtScens.Count > 0 Then
                'There's at least one scenario, so copy the input grid to the output as is so that it can be modified
                landCoverRaster.Save()  'Saving because it seemed necessary.
                landCoverRaster.Save(strOutLandCover)
                landCoverRaster.Close()
                pNewLandCoverRaster.Open(strOutLandCover)

                Using progress = New SynchronousProgressDialog("Creating Management Scenario", MgmtScens.Count, g_MainForm)
                    Dim i As Short
                    For i = 0 To MgmtScens.Count - 1
                        If MgmtScens.Item(i).intApply = 1 Then
                            progress.Increment("Adding new landclass...")
                            If SynchronousProgressDialog.KeepRunning Then
                                Dim mgmtitem As ManagementScenarioItem = MgmtScens.Item(i)
                                ReclassRaster(mgmtitem, landCoverName, pNewLandCoverRaster)
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

            If Not booLandScen Then
                g_LandCoverRaster = landCoverRaster
            Else
                g_LandCoverRaster = pNewLandCoverRaster
            End If

        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub
    Public Function GetLandClassValue(typename As String, className As String) As Double
        Dim strSelect As String = String.Format("SELECT LCTYPE.LCTYPEID, LCCLASS.NAME, LCCLASS.VALUE FROM LCTYPE INNER JOIN LCCLASS ON LCTYPE.LCTYPEID = LCCLASS.LCTYPEID WHERE LCTYPE.NAME LIKE '{0}' AND LCCLASS.NAME LIKE '{1}'", typename, className)
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
    Public Sub ReclassRaster(ByRef MgmtScen As ManagementScenarioItem, ByVal strLCClass As String, ByRef outputGrid As Grid)
        'We're passing over a single management scenarios in the form of the xml
        'class XmlmgmtScenItem, seems to be the easiest way to do this.
        Dim LCValue As Double = GetLandClassValue(strLCClass, MgmtScen.strChangeToClass)

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