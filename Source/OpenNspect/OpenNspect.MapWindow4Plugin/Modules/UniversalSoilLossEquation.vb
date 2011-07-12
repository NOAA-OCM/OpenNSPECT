
Imports System.Data.OleDb
Imports System.IO
Imports MapWinGIS
Module UniversalSoilLossEquation

    Public Function ConstructPickStatmentUsingLandClass(ByRef cmdType As OleDbCommand, ByRef pLCRaster As Grid, Optional ByVal nameOfColumn As String = "CoverFactor") As String
        'Creates the initial pick statement using the name of the the LandCass [CCAP, for example]
        'and the Land Class Raster.  Returns a string

        Try
            Dim FieldIndex As Short

            Dim i As Short
            Dim maxVal As Integer = pLCRaster.Maximum
            Dim tablepath = GetRasterTablePath(pLCRaster)
            Dim TableExist As Boolean = File.Exists(tablepath)
            Dim strpick As String = ""

            Dim mwTable As New Table
            If Not TableExist Then
                MsgBox("No MapWindow-readable raster table was found. To create one using ArcMap 9.3+, add the raster to the default project, right click on its layer and select Open Attribute Table. Now click on the options button in the lower right and select Export. In the export path, navigate to the directory of the grid folder and give the export the name of the raster folder with the .dbf extension. i.e. if you are exporting a raster attribute table from a raster named landcover, export landcover.dbf into the same level directory as the folder.", MsgBoxStyle.Exclamation, "Raster Attribute Table Not Found")

                Return ""
            Else
                mwTable.Open(tablepath)

                FieldIndex = GetFieldIndex(mwTable)

                Dim rowidx As Integer = 0
                Dim dataType As OleDbDataReader
                For i = 1 To maxVal
                    If (mwTable.CellValue(FieldIndex, rowidx) = i) Then 'And (pRow.Value(FieldIndex) = rsLandClass!Value) Then
                        dataType = cmdType.ExecuteReader


                        While dataType.Read()
                            If mwTable.CellValue(FieldIndex, rowidx) = dataType("Value") Then

                                If strpick = "" Then
                                    strpick = CStr(dataType(nameOfColumn))
                                Else
                                    strpick = String.Format("{0}, {1}", strpick, CStr(dataType(nameOfColumn)))
                                End If
                                rowidx = rowidx + 1
                                Exit While
                            Else
                                MsgBox("Error: Your OpenNSPECT Land Class Table is missing values found in your landcover GRID dataset.")
                                ConstructPickStatmentUsingLandClass = Nothing
                                dataType.Close()
                                mwTable.Close()
                                Exit Function
                            End If
                        End While

                        dataType.Close()

                    Else
                        If strpick = "" Then
                            strpick = "0"
                        Else
                            strpick = strpick & ", 0"
                        End If
                    End If

                Next
                mwTable.Close()
            End If

            Return strpick
        Catch ex As Exception
            HandleError(ex)
        End Try
        Return ""
    End Function
End Module
