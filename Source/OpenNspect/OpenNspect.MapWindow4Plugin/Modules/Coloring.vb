Imports MapWinGIS

Module Coloring

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

            Const STR_LegendCaptionFormat As String = "{0:#,0.##} - {1:#,0.##}"

            Dim cs As New GridColorScheme
            Dim csbrk As New GridColorBreak
            csbrk.LowValue = pRaster.Minimum
            csbrk.HighValue = firstBoundary
            csbrk.ColoringType = ColoringType.Gradient
            csbrk.LowColor = fromColor
            csbrk.HighColor = toColor

            csbrk.Caption = String.Format(STR_LegendCaptionFormat, csbrk.LowValue, csbrk.HighValue)
            cs.InsertBreak(csbrk)
            csbrk = New GridColorBreak
            csbrk.LowValue = firstBoundary
            csbrk.HighValue = secondBoundary
            csbrk.ColoringType = ColoringType.Gradient
            csbrk.LowColor = toColor
            csbrk.HighColor = fromColor2
            csbrk.Caption = String.Format(STR_LegendCaptionFormat, csbrk.LowValue, csbrk.HighValue)
            cs.InsertBreak(csbrk)
            csbrk = New GridColorBreak
            csbrk.LowValue = secondBoundary
            csbrk.HighValue = pRaster.Maximum
            csbrk.ColoringType = ColoringType.Gradient
            csbrk.LowColor = fromColor2
            csbrk.HighColor = toColor2
            csbrk.Caption = String.Format(STR_LegendCaptionFormat, csbrk.LowValue, csbrk.HighValue)
            cs.InsertBreak(csbrk)
            Return cs

        Catch ex As Exception
            HandleError(ex)
            Return Nothing
        End Try

    End Function
    ''' <summary>
    ''' Generates two color representation of the grid or feature layer to show which flow paths were above the water quality standards for pollutant loads
    ''' </summary>
    ''' <returns></returns>
    Public Function GetUniqueRasterRenderer() As Object
        'Create two colors, red, green
        Dim red As UInteger = Convert.ToUInt32(RGB(214, 71, 0))
        Dim green As UInteger = Convert.ToUInt32(RGB(56, 168, 0))

        Dim colorBreak1 As New GridColorBreak
        colorBreak1.Caption = "Exceeds Standard"
        colorBreak1.ColoringType = ColoringType.Gradient
        colorBreak1.HighValue = 1
        colorBreak1.LowValue = 1
        colorBreak1.HighColor = red
        colorBreak1.LowColor = red

        Dim colorBreak2 As New GridColorBreak
        colorBreak2.Caption = "Below Standard"
        colorBreak2.ColoringType = ColoringType.Gradient
        colorBreak2.HighValue = 2
        colorBreak2.LowValue = 2
        colorBreak2.HighColor = green
        colorBreak2.LowColor = green

        Dim cs As New GridColorScheme
        cs.InsertBreak(colorBreak1)
        cs.InsertBreak(colorBreak2)

        Return cs
    End Function


    Public Function GetContinuousRampColorCS(ByRef grd As Grid, ByRef strColor As String) As GridColorScheme
        'Based on the Mapwindow Grid Coloring Scheme Editor MakeContinuousRamp function
        ' The most similar method I found in DotSpatial is named CreateRampColors and located in the Scheme.cs file.
        ' It is called as follows: CreateRampColors(numColors, .25f, .25f, 0, .75f, .75f, 360, 0, 255, 255);

        Dim arr(), val As Object, i, j As Integer
        Dim ht As New Hashtable
        Dim brk As GridColorBreak
        Dim gradientModel As GradientModel
        Dim coloringType As ColoringType
        Dim coloringscheme As New GridColorScheme

        'Dim gradient As String = "Linear"
        'Dim gradient As String = "Exponential"
        Const gradient As String = "Logarithmic"
        Const numBreaks As Integer = 5

        Try
            If grd Is Nothing Then Return Nothing

            Dim rTo As Integer
            Dim bTo As Integer
            Dim gTo As Integer

            Dim rFrom As Integer
            Dim bFrom As Integer
            Dim gFrom As Integer

            Select Case strColor
                Case "Blue"
                    ' 242, 245, 255
                    rFrom = 242
                    gFrom = 245
                    bFrom = 255

                    '18, 73, 255
                    rTo = 18
                    gTo = 73
                    bTo = 255
                Case "Brown"
                    '255, 242, 217
                    rFrom = 255
                    gFrom = 242
                    bFrom = 217

                    '176, 117, 0
                    rTo = 176
                    gTo = 117
                    bTo = 0
                Case "81,100,100,81,5,100"
                    rTo = 166
                    gTo = 255
                    bTo = 0

                    rFrom = 251
                    gFrom = 255
                    bFrom = 242
                Case "45,97,100,45,5,100"
                    rFrom = 255
                    gFrom = 252
                    bFrom = 242

                    rTo = 255
                    gTo = 193
                    bTo = 8
                Case Else
                    'TODO: Convert HSV to RGB
                    rFrom = CInt(Split(strColor, ",")(0))
                    gFrom = CInt(Split(strColor, ",")(1))
                    bFrom = CInt(Split(strColor, ",")(2))

                    rTo = CInt(Split(strColor, ",")(3))
                    gTo = CInt(Split(strColor, ",")(4))
                    bTo = CInt(Split(strColor, ",")(5))
            End Select

            For i = 0 To grd.Header.NumberRows - 1
                For j = 0 To grd.Header.NumberCols - 1
                    val = grd.Value(j, i)
                    If ht.ContainsKey(val) = False Then
                        ht.Add(val, val)
                    End If
                Next
            Next

            ReDim arr(ht.Count - 1)
            ht.Values().CopyTo(arr, 0)
            Array.Sort(arr)

            While coloringscheme.NumBreaks > 0
                coloringscheme.DeleteBreak(0)
            End While

            Select Case gradient
                Case "Linear"
                    gradientModel = MapWinGIS.GradientModel.Linear
                Case "Exponential"
                    gradientModel = MapWinGIS.GradientModel.Exponential
                Case "Logorithmic"
                    gradientModel = MapWinGIS.GradientModel.Logorithmic
                Case Else
                    Throw New NotImplementedException
            End Select

            coloringType = MapWinGIS.ColoringType.Gradient

            '' No Data break
            'brk = New MapWinGIS.GridColorBreak()
            'brk.Caption = "No Data"
            'brk.LowValue = grd.Header.NodataValue
            'brk.HighValue = brk.LowValue
            'brk.LowColor = MapWinUtility.Colors.ColorToUInteger(Color.Black)
            'brk.HighColor = brk.LowColor
            'brk.GradientModel = gradientModel
            'brk.ColoringType = coloringType
            'm_ColoringScheme.InsertBreak(brk)

            Dim sR, sG, sB As Integer
            Dim eR, eG, eB As Integer
            Dim r, g, b As Double
            Dim rStep, gStep, bStep As Double
            Dim startVal As Integer

            sR = rFrom
            sG = gFrom
            sB = bFrom
            eR = rTo
            eG = gTo
            eB = bTo

            r = sR
            g = sG
            b = sB

            If ht.Keys.Count <= numBreaks Then
                Dim brkArr(ht.Keys.Count - 1) As Object
                ht.Keys.CopyTo(brkArr, 0)
                Array.Sort(brkArr)

                rStep = (eR - sR) / brkArr.Length
                gStep = (eG - sG) / brkArr.Length
                bStep = (eB - sB) / brkArr.Length

                'This must be double.parse(convert.tostring) for handling of sbyte values - cdm 11/13/2005
                startVal = CInt(IIf(Double.Parse(Convert.ToString(grd.Header.NodataValue)) = Double.Parse(Convert.ToString(brkArr(0))), 1, 0))
                'startVal = CInt(IIf(CDbl(grd.Header.NodataValue) = CDbl(brkArr(0)), 1, 0))
                For i = startVal To brkArr.Length - 1
                    brk = New GridColorBreak
                    If IsNumeric(brkArr(i)) Then
                        brk.Caption = Double.Parse(Convert.ToString((brkArr(i)))).ToString()
                        'brk.Caption = CDbl(brkArr(i)).ToString(m_NumberFormat & m_Precision)
                        'tPrecision = CInt(Math.Floor(m_Precision - Math.Log10(CDbl(arr(i)))))
                        'If tPrecision < 0 Then tPrecision = 0
                        'brk.Caption = CStr(Math.Round(CDbl(brkArr(i)), tPrecision))
                    End If
                    brk.LowValue = Double.Parse(Convert.ToString(brkArr(i)))
                    brk.HighValue = brk.LowValue
                    brk.LowColor = Convert.ToUInt32(RGB(CInt(r), CInt(g), CInt(b)))
                    r += rStep
                    g += gStep
                    b += bStep
                    brk.HighColor = Convert.ToUInt32(RGB(CInt(r), CInt(g), CInt(b)))
                    brk.ColoringType = coloringType
                    brk.GradientModel = gradientModel
                    coloringscheme.InsertBreak(brk)
                Next
            Else
                rStep = (eR - sR) / numBreaks
                gStep = (eG - sG) / numBreaks
                bStep = (eB - sB) / numBreaks

                Dim min As Double, max As Double, range As Double
                startVal = CInt(IIf(Double.Parse(Convert.ToString(grd.Header.NodataValue)) = Double.Parse(Convert.ToString(arr(0))), 1, 0))

                min = Double.Parse(Convert.ToString(arr(startVal)))
                max = Double.Parse(Convert.ToString(arr(arr.Length() - 1)))
                range = max - min

                Dim prev As Double = min
                Dim t As Double = range / numBreaks

                For i = 1 To numBreaks - 1
                    brk = New GridColorBreak
                    brk.LowValue = prev
                    brk.HighValue = prev + t
                    brk.LowColor = Convert.ToUInt32(RGB(CInt(r), CInt(g), CInt(b)))
                    r += rStep
                    g += gStep
                    b += bStep
                    brk.HighColor = Convert.ToUInt32(RGB(CInt(r), CInt(g), CInt(b)))
                    If brk.HighValue = brk.LowValue Then
                        If brk.LowValue = min Then
                            brk.Caption = CStr(min)
                        Else
                            brk.Caption = brk.LowValue.ToString()
                            'tPrecision = CInt(Math.Floor(m_Precision - Math.Log10(CDbl(brk.LowValue))))
                            'If tPrecision < 0 Then tPrecision = 0
                            'brk.Caption = CStr(Math.Round(CDbl(brk.LowValue), tPrecision))
                        End If
                    Else
                        If brk.LowValue = min Then
                            brk.Caption = CStr(min)
                        Else
                            brk.Caption = brk.LowValue.ToString()
                            'tPrecision = CInt(Math.Floor(m_Precision - Math.Log10(CDbl(brk.LowValue))))
                            'If tPrecision < 0 Then tPrecision = 0
                            'brk.Caption = CStr(Math.Round(CDbl(brk.LowValue), tPrecision))
                        End If
                        brk.Caption &= " - "
                        brk.Caption = brk.HighValue.ToString()
                        'tPrecision = CInt(Math.Floor(m_Precision - Math.Log10(CDbl(brk.HighValue))))
                        'If tPrecision < 0 Then tPrecision = 0
                        'brk.Caption &= Math.Round(CDbl(brk.HighValue), tPrecision)
                    End If
                    brk.ColoringType = coloringType
                    brk.GradientModel = gradientModel
                    coloringscheme.InsertBreak(brk)
                    prev = brk.HighValue
                Next
                ' now do the last break
                brk = New GridColorBreak
                brk.LowValue = prev
                brk.HighValue = max
                brk.LowColor = Convert.ToUInt32(RGB(CInt(r), CInt(g), CInt(b)))
                r = eR
                g = eG
                b = eB
                brk.HighColor = Convert.ToUInt32(RGB(CInt(r), CInt(g), CInt(b)))
                If brk.HighValue = brk.LowValue Then
                    brk.Caption = CStr(brk.LowValue)
                Else
                    brk.Caption = brk.LowValue.ToString()
                    'tPrecision = CInt(Math.Floor(m_Precision - Math.Log10(CDbl(brk.LowValue))))
                    'If tPrecision < 0 Then tPrecision = 0
                    'brk.Caption = Math.Round(CDbl(brk.LowValue), tPrecision) & " - " & brk.HighValue
                End If
                brk.ColoringType = coloringType
                brk.GradientModel = gradientModel
                coloringscheme.InsertBreak(brk)
            End If
            Return coloringscheme
        Catch ex As Exception
            HandleError(ex)
            Return Nothing
        End Try
    End Function


    ''' <summary>
    ''' Returns the HSV color string.
    ''' </summary><returns>A comma delimited string of 6 values.  1st 3 a 'To Color' - HIGH, 2nd 3 a 'From Color' - LOW</returns>
    Public Function GetRandomHSVColorString() As String
        'Hue is a value from 1 to 360 so find a random one
        Dim intHue As Short = Int((360 * Rnd()) + 1)

        'Value will be a constant of 97, 100 in the SV and 5, 100..
        Return String.Format("{0},97,100,{0},5,100", CStr(intHue))

    End Function
End Module
