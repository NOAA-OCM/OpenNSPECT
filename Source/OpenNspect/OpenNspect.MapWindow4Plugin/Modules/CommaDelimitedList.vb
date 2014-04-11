
Imports System.Data.OleDb
Imports System.IO
Imports MapWinGIS
Module CommaDelimitedList

    ''' <summary>
    ''' Creates the initial pick statement using the name of the the LandCass [CCAP, for example].
    ''' </summary>
    ''' <remarks>Assumes that the Values are sorted in increasing order.</remarks>
    ''' <param name="pLCRaster">The raster.</param>
    ''' <param name="nameOfColumn">The name of column.</param><returns></returns>
    Public Function ConstructPickStatmentUsingLandClass(ByRef command As OleDbCommand, ByRef pLCRaster As Grid, Optional ByVal nameOfColumn As String = "CoverFactor") As String

        Dim tablepath = GetRasterTablePath(pLCRaster)
        If Not File.Exists(tablepath) Then
            MsgBox("No MapWindow-readable raster table was found (1).", MsgBoxStyle.Exclamation, "Raster Attribute Table Not Found")

            Return String.Empty
        End If

        Dim attributeTable As New Table
        Try
            Dim strpick As String = String.Empty

            attributeTable.Open(tablepath)
            Dim valueColumn As Short = GetFieldIndex(attributeTable) ' Determine what column 'value' appears in.
            Dim currentTableRow As Integer = 0
            Dim maxVal As Integer = pLCRaster.Maximum
            Dim minVal As Integer = pLCRaster.Minimum


            ' Look for each value from minVal to the maximum, assuming values are integers.
            ' DLE: 109/12: LC rasters may have 0 as min value, but if minval > 1, then start at 1
            If (minVal > 1) Then minVal = 1
            If currentTableRow = pLCRaster.Header.NodataValue Or currentTableRow < minVal Then
                currentTableRow = minVal
            End If
            For i = minVal To maxVal
                'For i = 1 To maxVal
                If (attributeTable.CellValue(valueColumn, currentTableRow) = i) Then
                    Using dataType As OleDbDataReader = command.ExecuteReader

                        Dim valueFoundInDataBase As Boolean
                        While dataType.Read()

                            If attributeTable.CellValue(valueColumn, currentTableRow) = dataType("Value") Then
                                valueFoundInDataBase = True
                                strpick = String.Format(If(strpick = "", "{1}", "{0}, {1}"), strpick, CStr(dataType(nameOfColumn)))
                                currentTableRow = currentTableRow + 1
                                Exit While
                            Else
                                valueFoundInDataBase = False
                            End If
                        End While
                        If valueFoundInDataBase = False Then
                            MsgBox("Error: Your OpenNSPECT Land Class Table (database) is missing values found in your landcover GRID dataset (.dbf).")
                            Return Nothing
                        End If
                    End Using
                Else
                    strpick = If(strpick = "", "0", strpick & ", 0")
                End If
            Next

            Return strpick
        Catch ex As Exception
            HandleError(ex)
        Finally
            If attributeTable IsNot Nothing Then
                attributeTable.Close()
            End If
        End Try
        Return Nothing
    End Function
End Module
