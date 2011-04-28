Imports System.IO
Imports MapWinGIS

Module Rasters

    ''' <summary>
    ''' Gets the raster full path.
    ''' </summary>
    ''' <param name="pathOrDirectory">The full path or directory where the raster resides.</param><returns></returns>
    Public Function GetRasterFullPath(pathOrDirectory As String) As String
        If Path.HasExtension(pathOrDirectory) Then
            Return pathOrDirectory
        Else
            Return Path.Combine(pathOrDirectory, "sta.adf")
        End If
    End Function
    Public Function RasterExists(ByRef strRasterFileName As String) As Boolean
        Return File.Exists(GetRasterFullPath(strRasterFileName))
    End Function

    Public Function ReturnRaster(pathOrDirectory As String) As Grid
        Dim path = GetRasterFullPath(pathOrDirectory)

        If File.Exists(path) Then
            Dim g As New Grid
            If g.Open(path) Then
                Return g
            End If
        End If

        Return Nothing
    End Function

    Public Function ReturnPermanentRaster(ByRef pRaster As Grid, ByRef sOutputName As String) _
        As Grid
        pRaster.Save()
        pRaster.Save(sOutputName)
        pRaster.Header.Projection = MapWindowPlugin.MapWindowInstance.Project.ProjectProjection

        Dim tmpraster As New Grid
        tmpraster.Open(sOutputName)
        Return tmpraster
    End Function

    Public Function ReturnRasterStretchColorRampCS(ByRef pRaster As Grid, ByRef strColor As String) _
        As GridColorScheme
        ReturnRasterStretchColorRampCS = Nothing
        Try
            Dim rTo, bTo, gTo, rFrom, bFrom, gFrom, rTo2, bTo2, gTo2, rFrom2, bFrom2, gFrom2 As Integer

            Select Case strColor
                Case "Blue"
                    ' 242, 245, 255
                    rFrom = 242
                    gFrom = 245
                    bFrom = 255

                    '149, 173, 255
                    rTo = 91
                    gTo = 129
                    bTo = 255

                    '60, 103, 255
                    rFrom2 = 60
                    gFrom2 = 103
                    bFrom2 = 255

                    '18, 73, 255
                    rTo2 = 18
                    gTo2 = 73
                    bTo2 = 255
                Case "Brown"
                    '255, 242, 217
                    rFrom = 255
                    gFrom = 242
                    bFrom = 217

                    '255, 202, 102
                    rTo = 255
                    gTo = 202
                    bTo = 102

                    '255, 174, 23
                    rFrom2 = 255
                    gFrom2 = 174
                    bFrom2 = 23

                    '255, 193, 8
                    rTo2 = 255
                    gTo2 = 193
                    bTo2 = 8
                Case "81,100,100,81,5,100"
                    rFrom = 251
                    gFrom = 255
                    bFrom = 242

                    rTo = 187
                    gTo = 255
                    bTo = 100

                    rFrom2 = 187
                    gFrom2 = 255
                    bFrom2 = 60

                    rTo2 = 166
                    gTo2 = 255
                    bTo2 = 0
                Case "45,97,100,45,5,100"
                    rFrom = 255
                    gFrom = 242
                    bFrom = 217

                    '255, 202, 102
                    rTo = 255
                    gTo = 221
                    bTo = 157

                    '255, 174, 23
                    rFrom2 = 255
                    gFrom2 = 193
                    bFrom2 = 79

                    rTo2 = 255
                    gTo2 = 193
                    bTo2 = 8
                Case "273,97,100,273,5,100"
                    '144, 8, 255  
                    '249, 242, 255
                    rFrom = 249
                    gFrom = 242
                    bFrom = 255

                    rTo = 206
                    gTo = 147
                    bTo = 255

                    rFrom2 = 174
                    gFrom2 = 74
                    bFrom2 = 255

                    rTo2 = 144
                    gTo2 = 8
                    bTo2 = 255
                Case "198,97,100,198,5,100"
                    '242, 252, 255
                    '8, 181, 255
                    rFrom = 242
                    gFrom = 252
                    bFrom = 255

                    rTo = 149
                    gTo = 223
                    bTo = 255

                    rFrom2 = 85
                    gFrom2 = 204
                    bFrom2 = 255

                    rTo2 = 8
                    gTo2 = 181
                    bTo2 = 255
                Case "356,97,100,356,5,100"
                    '255, 242, 243
                    '255, 8, 24
                    rFrom = 255
                    gFrom = 242
                    bFrom = 243

                    rTo = 255
                    gTo = 140
                    bTo = 149

                    rFrom2 = 255
                    gFrom2 = 74
                    bFrom2 = 88

                    rTo2 = 255
                    gTo2 = 8
                    bTo2 = 24
                Case Else
                    'TODO: Convert HSV to RGB
                    rFrom = CInt(Split(strColor, ",")(0))
                    gFrom = CInt(Split(strColor, ",")(1))
                    bFrom = CInt(Split(strColor, ",")(2))

                    rTo = CInt(Split(strColor, ",")(3))
                    gTo = CInt(Split(strColor, ",")(4))
                    bTo = CInt(Split(strColor, ",")(5))

                    rFrom2 = CInt(Split(strColor, ",")(0))
                    gFrom2 = CInt(Split(strColor, ",")(1))
                    bFrom2 = CInt(Split(strColor, ",")(2))

                    rTo2 = CInt(Split(strColor, ",")(3))
                    gTo2 = CInt(Split(strColor, ",")(4))
                    bTo2 = CInt(Split(strColor, ",")(5))
            End Select

            Dim total As Double = 0.0
            Dim sqrtotal As Double = 0.0
            Dim stdDev As Double
            Dim count As Integer = 0
            Dim nc, nr As Integer
            Dim nodata As Double
            nc = pRaster.Header.NumberCols - 1
            nr = pRaster.Header.NumberRows - 1
            nodata = pRaster.Header.NodataValue
            Dim rowvals() As Single
            ReDim rowvals(nc)
            For row As Integer = 0 To nr
                pRaster.GetRow(row, rowvals(0))
                For col As Integer = 0 To nc
                    If rowvals(col) <> nodata And rowvals(col) < Double.MaxValue And rowvals(col) > Double.MinValue _
                        Then
                        total = total + rowvals(col)
                        sqrtotal = sqrtotal + (rowvals(col) * rowvals(col))
                        count = count + 1
                    End If
                Next
            Next
            stdDev = Math.Sqrt((sqrtotal / count) - (total / count) * (total / count))

            Const stdMult As Integer = 2

            'Dim percVal As Double = (stdDev * stdMult) / pRaster.Maximum
            'Dim rstep As Double = (stdMult) * (rTo - rFrom) * percVal
            'Dim gstep As Double = (stdMult) * (gTo - gFrom) * percVal
            'Dim bstep As Double = (stdMult) * (bTo - bFrom) * percVal

            Dim cs As New GridColorScheme
            Dim csbrk As New GridColorBreak
            csbrk.LowValue = pRaster.Minimum
            csbrk.HighValue = pRaster.Minimum + stdMult * stdDev
            csbrk.ColoringType = ColoringType.Gradient
            csbrk.LowColor = Convert.ToUInt32(RGB(CInt(rFrom), CInt(gFrom), CInt(bFrom)))
            csbrk.HighColor = Convert.ToUInt32(RGB(CInt(rTo), CInt(gTo), CInt(bTo)))
            csbrk.Caption = String.Format("{0} - {1}", csbrk.LowValue, csbrk.HighValue)
            cs.InsertBreak(csbrk)
            csbrk = New GridColorBreak
            csbrk.LowValue = pRaster.Minimum + stdMult * stdDev
            csbrk.HighValue = pRaster.Maximum - stdMult * stdDev
            csbrk.ColoringType = ColoringType.Gradient
            csbrk.LowColor = Convert.ToUInt32(RGB(CInt(rTo), CInt(gTo), CInt(bTo)))
            csbrk.HighColor = Convert.ToUInt32(RGB(CInt(rFrom2), CInt(gFrom2), CInt(bFrom2)))
            csbrk.Caption = String.Format("{0} - {1}", csbrk.LowValue, csbrk.HighValue)
            cs.InsertBreak(csbrk)
            csbrk = New GridColorBreak
            csbrk.LowValue = pRaster.Maximum - stdMult * stdDev
            csbrk.HighValue = pRaster.Maximum
            csbrk.ColoringType = ColoringType.Gradient
            csbrk.LowColor = Convert.ToUInt32(RGB(CInt(rFrom2), CInt(gFrom2), CInt(bFrom2)))
            csbrk.HighColor = Convert.ToUInt32(RGB(CInt(rTo2), CInt(gTo2), CInt(bTo2)))
            csbrk.Caption = String.Format("{0} - {1}", csbrk.LowValue, csbrk.HighValue)
            cs.InsertBreak(csbrk)
            Return cs
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Function

End Module
