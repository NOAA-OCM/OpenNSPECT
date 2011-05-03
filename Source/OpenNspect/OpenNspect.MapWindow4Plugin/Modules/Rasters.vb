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

    Public Function ReturnPermanentRaster(ByRef pRaster As Grid, ByRef sOutputName As String) As Grid

        pRaster.Save()
        pRaster.Save(sOutputName)
        pRaster.Header.Projection = MapWindowPlugin.MapWindowInstance.Project.ProjectProjection

        Dim tmpraster As New Grid
        tmpraster.Open(sOutputName)
        Return tmpraster
    End Function

    Public Function AddRasterLayerToMapFromFileName(ByRef fileOrDirectory As String) As Boolean
        If Not Path.HasExtension(fileOrDirectory) Then
            ' we were passed a directory which represents an Esri Grid.
            fileOrDirectory = fileOrDirectory + "\sta.adf"
        End If

        MapWindowPlugin.MapWindowInstance.Layers.Add(fileOrDirectory)

        Return True
    End Function

    Public Function GetRasterStretchColorRampCS(ByRef pRaster As Grid, ByRef strColor As String) As GridColorScheme
        ' The commented values look like other colors that could be substituted.
        Try
            Dim fromColor As Integer
            Dim toColor As Integer
            Dim fromColor2 As Integer
            Dim toColor2 As Integer

            Select Case strColor
                Case "Blue"
                    fromColor = RGB(242, 245, 255)
                    toColor = RGB(91, 129, 255)
                    '149, 173, 255

                    fromColor2 = RGB(60, 103, 255)
                    toColor2 = RGB(18, 73, 255)

                Case "Brown"
                    fromColor = RGB(255, 242, 217)
                    toColor = RGB(255, 202, 102)

                    fromColor2 = RGB(255, 174, 23)
                    toColor2 = RGB(255, 193, 8)

                Case "81,100,100,81,5,100"

                    fromColor = RGB(251, 255, 242)
                    toColor = RGB(187, 255, 100)

                    fromColor2 = RGB(187, 255, 60)
                    toColor2 = RGB(166, 255, 0)

                Case "45,97,100,45,5,100"
                    fromColor = RGB(255, 242, 217)
                    toColor = RGB(255, 221, 157)
                    '255, 202, 102

                    fromColor2 = RGB(255, 193, 79)
                    toColor2 = RGB(255, 193, 8)
                Case "273,97,100,273,5,100"
                    '144, 8, 255  
                    '249, 242, 255
                    fromColor = RGB(249, 242, 255)
                    toColor = RGB(206, 147, 255)

                    fromColor2 = RGB(174, 74, 255)
                    toColor2 = RGB(144, 8, 255)

                Case "198,97,100,198,5,100"
                    '242, 252, 255
                    '8, 181, 255
                    fromColor = RGB(242, 252, 255)
                    toColor = RGB(149, 223, 255)

                    fromColor2 = RGB(85, 204, 255)
                    toColor2 = RGB(8, 181, 255)

                Case "356,97,100,356,5,100"
                    '255, 242, 243
                    '255, 8, 24
                    fromColor = RGB(255, 242, 243)
                    toColor = RGB(255, 140, 149)

                    fromColor2 = RGB(255, 74, 88)
                    toColor2 = RGB(255, 8, 24)
                Case Else
                    'TODO: Convert HSV to RGB
                    Throw New NotImplementedException("Convert HSV to RGB")
                    'rFrom = CInt(Split(strColor, ",")(0))
                    'gFrom = CInt(Split(strColor, ",")(1))
                    'bFrom = CInt(Split(strColor, ",")(2))

                    'rTo = CInt(Split(strColor, ",")(3))
                    'gTo = CInt(Split(strColor, ",")(4))
                    'bTo = CInt(Split(strColor, ",")(5))

                    'rFrom2 = CInt(Split(strColor, ",")(0))
                    'gFrom2 = CInt(Split(strColor, ",")(1))
                    'bFrom2 = CInt(Split(strColor, ",")(2))

                    'rTo2 = CInt(Split(strColor, ",")(3))
                    'gTo2 = CInt(Split(strColor, ",")(4))
                    'bTo2 = CInt(Split(strColor, ",")(5))
            End Select

            Dim total As Double = 0.0
            Dim sqrtotal As Double = 0.0
            Dim stdDev As Double
            Dim count As Integer = 0
            Dim nc As Integer = pRaster.Header.NumberCols - 1
            Dim nr As Integer = pRaster.Header.NumberRows - 1
            Dim nodata As Double = pRaster.Header.NodataValue
            Dim rowvals(nc) As Single

            For row As Integer = 0 To nr
                pRaster.GetRow(row, rowvals(0))
                For col As Integer = 0 To nc
                    If rowvals(col) <> nodata And rowvals(col) < Double.MaxValue And rowvals(col) > Double.MinValue Then
                        total = total + rowvals(col)
                        sqrtotal = sqrtotal + (rowvals(col) * rowvals(col))
                        count = count + 1
                    End If
                Next
            Next
            stdDev = Math.Sqrt((sqrtotal / count) - (total / count) * (total / count))

            'Dim percVal As Double = (stdDev * stdMult) / pRaster.Maximum
            'Dim rstep As Double = (stdMult) * (rTo - rFrom) * percVal
            'Dim gstep As Double = (stdMult) * (gTo - gFrom) * percVal
            'Dim bstep As Double = (stdMult) * (bTo - bFrom) * percVal

            Const stdMult As Integer = 2
            Dim firstBoundary As Double = pRaster.Minimum + stdMult * stdDev
            Dim secondBoundary As Double = pRaster.Maximum - stdMult * stdDev

            Dim cs As New GridColorScheme
            Dim csbrk As New GridColorBreak
            csbrk.LowValue = pRaster.Minimum
            csbrk.HighValue = firstBoundary
            csbrk.ColoringType = ColoringType.Gradient
            csbrk.LowColor = fromColor
            csbrk.HighColor = toColor
            csbrk.Caption = String.Format("{0} - {1}", csbrk.LowValue, csbrk.HighValue)
            cs.InsertBreak(csbrk)
            csbrk = New GridColorBreak
            csbrk.LowValue = firstBoundary
            csbrk.HighValue = secondBoundary
            csbrk.ColoringType = ColoringType.Gradient
            csbrk.LowColor = toColor
            csbrk.HighColor = fromColor2
            csbrk.Caption = String.Format("{0} - {1}", csbrk.LowValue, csbrk.HighValue)
            cs.InsertBreak(csbrk)
            csbrk = New GridColorBreak
            csbrk.LowValue = secondBoundary
            csbrk.HighValue = pRaster.Maximum
            csbrk.ColoringType = ColoringType.Gradient
            csbrk.LowColor = fromColor2
            csbrk.HighColor = toColor2
            csbrk.Caption = String.Format("{0} - {1}", csbrk.LowValue, csbrk.HighValue)
            cs.InsertBreak(csbrk)
            Return cs

        Catch ex As Exception
            HandleError(ex)
            Return Nothing
        End Try

    End Function

End Module
