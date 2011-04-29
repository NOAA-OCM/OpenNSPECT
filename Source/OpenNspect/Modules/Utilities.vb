'********************************************************************************************************
'File Name: modUtil.vb
'Description: Many global utility functions used throughout the plugin
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
Imports System.Windows.Forms
Imports System.IO
Imports MapWindow.Interfaces
Imports MapWinGeoProc
Imports MapWinGIS
Imports OpenNspect.Xml

Module Utilities
    Public g_nspectPath As String
    Public g_nspectDocPath As String
    Public g_strWorkspace As String



    Public g_cbMainForm As MainForm
    Public g_MainForm As Form

    Public g_comp As CompareOutputsForm
    Public g_luscen As EditLandUseScenario

    Public g_strSelectedExportPath As String = ""

    'Intermediate and output extensions
    Public Const OutputGridExt As String = ".tif"
    Public Const FinalOutputGridExt As String = ".tif"
    Public Const TAUDEMGridExt As String = ".tif"

    Public g_TempFilesToDel As New List(Of String)


    Public g_boolAddCoeff As Boolean
    Public g_boolCopyCoeff As Boolean
    'True: called frmPollutants, False: called frmNewPollutants
    Public g_boolAgree As Boolean
    'True: use the Agree Function on Streams.
    Public g_boolHydCorr As Boolean
    'True: Hyrdologically Correct DEM, no fill needed
    Public g_boolNewWShed As Boolean
    'True: New WaterShed form called from frmPrj

    'WqStd

    'Agree DEM Stuff
    Public g_boolParams As Boolean
    'Flag to indicate Agree params have been entered

    'Project Form Variables
    Public g_strPrjFileName As String
    'Project file name

    'Management Scenario variables::frmPrjCalc
    Public g_strLUScenFileName As String
    'Management scenario file name
    Public g_intManScenRow As String
    'Management scenario ROW number

    'Pollutant Coefficient variable::frmPrjCalc
    Public g_intCoeffRow As Short
    'Coeff Row Number
    Public g_strCoeffCalc As String
    'if the Calc option is chosen, hold results in string



    'Returns a filename given for example C:\temp\dataset returns dataset

    Public Function SplitFileName(ByRef sWholeName As String) As String
        SplitFileName = ""
        Try
            Dim pos As Short
            Dim sT As Object
            Dim sName As String
            pos = InStrRev(sWholeName, "\")
            If pos > 0 Then
                sT = Mid(sWholeName, 1, pos - 1)
                If pos = Len(sWholeName) Then
                    Exit Function
                End If
                sName = Mid(sWholeName, pos + 1, Len(sWholeName) - Len(sT))
                pos = InStr(sName, ".")
                If pos > 0 Then
                    SplitFileName = Mid(sName, 1, pos - 1)
                Else
                    SplitFileName = sName
                End If
            End If
        Catch ex As Exception
            MsgBox("Workspace Split:" & Err.Description)
        End Try
    End Function

    Public Function GetRasterDistanceUnits(ByRef strLayerName As String) As Short
        For i As Integer = 0 To MapWindowPlugin.MapWindowInstance.Layers.NumLayers - 1
            If MapWindowPlugin.MapWindowInstance.Layers(i).Name = strLayerName Then
                Dim proj As String = MapWindowPlugin.MapWindowInstance.Layers(i).Projection
                If proj <> "" Then
                    If proj.Contains("units=m") Then
                        Return 0
                    ElseIf proj.Contains("units=ft") Then
                        Return 1
                    End If
                Else
                    MsgBox("The GRID you have choosen has no spatial reference information.  Please define a projection before continuing.", MsgBoxStyle.Exclamation, "No Project Information Detected")
                    Return -1
                End If
            End If
        Next
        Return -1
    End Function

    Public Function GetIndexOfEntry(ByRef strList As String, ByRef cbo As ComboBox) As Short
        Return cbo.Items.IndexOf(strList)
    End Function

    Public Function LayerLoadedInMap(ByRef layerName As String) As Boolean
        Try
            For Each layer As Layer In MapWindowPlugin.MapWindowInstance.Layers
                If layer.Name = layerName Then
                    Return True
                End If
            Next

            Return False
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Function

    Public Function LayerInMapByFileName(ByRef strName As String) As Boolean

        For lngLyrIndex As Integer = 0 To MapWindowPlugin.MapWindowInstance.Layers.NumLayers - 1
            Dim pLayer As Layer = MapWindowPlugin.MapWindowInstance.Layers(MapWindowPlugin.MapWindowInstance.Layers.GetHandle(lngLyrIndex))

            If Trim(LCase(pLayer.FileName)) <> Trim(LCase(strName)) Then
                LayerInMapByFileName = False
            Else
                LayerInMapByFileName = True
                Exit For
            End If
        Next
    End Function

    Public Function GetLayerIndex(ByRef strLayerName As String) As Integer
        GetLayerIndex = -1
        Try
            For lngLyrIndex As Integer = 0 To MapWindowPlugin.MapWindowInstance.Layers.NumLayers - 1
                Dim pLayer As Layer = MapWindowPlugin.MapWindowInstance.Layers(MapWindowPlugin.MapWindowInstance.Layers.GetHandle(lngLyrIndex))

                If Trim(LCase(pLayer.Name)) = Trim(LCase(strLayerName)) Then
                    GetLayerIndex = lngLyrIndex
                    Exit For
                End If
            Next
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Function

    Public Function GetLayerIndexByFilename(ByRef strLayerFileName As String) As Integer
        GetLayerIndexByFilename = -1
        Try
            For lngLyrIndex As Integer = 0 To MapWindowPlugin.MapWindowInstance.Layers.NumLayers - 1
                Dim pLayer As Layer = MapWindowPlugin.MapWindowInstance.Layers(MapWindowPlugin.MapWindowInstance.Layers.GetHandle(lngLyrIndex))

                If Trim(LCase(pLayer.FileName)) = Trim(LCase(strLayerFileName)) Then
                    GetLayerIndexByFilename = lngLyrIndex
                    Exit For
                End If
            Next
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Function

    Public Function GetLayerFilename(ByRef strLayerName As String) As String
        GetLayerFilename = ""
        Try
            For lngLyrIndex As Integer = 0 To MapWindowPlugin.MapWindowInstance.Layers.NumLayers - 1
                Dim pLayer As Layer = MapWindowPlugin.MapWindowInstance.Layers(MapWindowPlugin.MapWindowInstance.Layers.GetHandle(lngLyrIndex))

                If Trim(LCase(pLayer.Name)) = Trim(LCase(strLayerName)) Then
                    GetLayerFilename = pLayer.FileName

                    If GetLayerFilename.EndsWith("sta.adf") Then
                        GetLayerFilename = Path.GetDirectoryName(GetLayerFilename)
                    End If

                    Exit For
                End If
            Next
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Function

    Public Function AddRasterLayerToMapFromFileName(ByRef strName As String) As Boolean
        Try

            If Path.GetExtension(strName) <> "" Then
                MapWindowPlugin.MapWindowInstance.Layers.Add(strName)
            Else
                MapWindowPlugin.MapWindowInstance.Layers.Add(strName + "\sta.adf")
            End If
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function


    Public Function AddInputFromGxBrowser(ByRef strType As String) As String
        Try
            Dim pFilter As String = ""

            Select Case strType
                Case "Feature"
                    Dim shp As New Shapefile
                    pFilter = shp.CdlgFilter
                Case "Raster"
                    Dim g As New Grid
                    pFilter = g.CdlgFilter
            End Select

            Dim dlgopen As New OpenFileDialog
            dlgopen.Title = "Open a " + strType + " file"
            dlgopen.Filter = pFilter

            If dlgopen.ShowDialog = DialogResult.OK Then
                Return dlgopen.FileName
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox("AddInputfromBrowser:" & Err.Description)
        End Try
        Return ""
    End Function

    Public Function AddInputFromGxBrowserText(ByRef txtInput As TextBox, ByRef strTitle As String) As Grid
        AddInputFromGxBrowserText = Nothing
        Try
            Using dlgOpen As New OpenFileDialog()
                Dim g As New Grid()
                dlgOpen.Filter = g.CdlgFilter
                dlgOpen.Title = strTitle
                If dlgOpen.ShowDialog = DialogResult.OK Then
                    txtInput.Text = dlgOpen.FileName
                    g.Open(dlgOpen.FileName)
                    Return g
                Else
                    Return Nothing
                End If
            End Using
        Catch ex As Exception
            MsgBox("The file you have choosen is not a valid GRID dataset.  Please select another.", MsgBoxStyle.Critical, "Invalid Data Type")
        End Try

    End Function



    Public Function ReturnContinuousRampColorCS(ByRef grd As Grid, ByRef strColor As String) As GridColorScheme
        'Based on the Mapwindow Grid Coloring Scheme Editor MakeContinuousRamp function
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
                Dim brkArr() As Object
                ReDim brkArr(ht.Keys.Count - 1)
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

    Public Function ReturnUniqueRasterRenderer(ByRef pRaster As Grid, ByRef strStandardName As String) As Object
        Try
            'Create two colors, red, green
            'pColorRed.RGB = System.Convert.ToUInt32(RGB(214, 71, 0))
            'pColorGreen.RGB = System.Convert.ToUInt32(RGB(56, 168, 0))

            Dim cs As New GridColorScheme
            Dim csb1 As New GridColorBreak
            Dim csb2 As New GridColorBreak

            csb1.Caption = "Exceeds Standard"
            csb1.ColoringType = ColoringType.Gradient
            csb1.HighValue = 1
            csb1.LowValue = 1
            csb1.HighColor = Convert.ToUInt32(RGB(214, 71, 0))
            csb1.LowColor = Convert.ToUInt32(RGB(214, 71, 0))

            csb2.Caption = "Below Standard"
            csb2.ColoringType = ColoringType.Gradient
            csb2.HighValue = 2
            csb2.LowValue = 2
            csb2.HighColor = Convert.ToUInt32(RGB(56, 168, 0))
            csb2.LowColor = Convert.ToUInt32(RGB(56, 168, 0))

            cs.InsertBreak(csb1)
            cs.InsertBreak(csb2)

            Return cs
        Catch ex As Exception
            MsgBox(Err.Description)
            Return Nothing
        End Try
    End Function

    Public Function ReturnHSVColorString() As String
        ReturnHSVColorString = ""
        Try
            'Returns a comma delimited string of 6 values.  1st 3 a 'To Color' - HIGH, 2nd 3 a 'From Color' - LOW
            Dim intHue As Short

            'Hue is a value from 1 to 360 so find a random one
            intHue = Int((360 * Rnd()) + 1)

            'Value will be a constant of 97, 100 in the SV and 5, 100..
            ReturnHSVColorString = CStr(intHue) & ",97,100," & CStr(intHue) & ",5,100"

        Catch ex As Exception
            HandleError(ex)
        End Try
    End Function

    Public Function CheckSpatialReference(ByRef pRasGeoDataset As Grid) As String
        CheckSpatialReference = ""
        Try
            If Not pRasGeoDataset Is Nothing Then
                Dim strprj As String = pRasGeoDataset.Header.Projection
                If strprj <> "" Then
                    Return strprj
                Else
                    If Path.GetFileName(pRasGeoDataset.Filename) = "sta.adf" Then
                        If File.Exists(Path.GetDirectoryName(pRasGeoDataset.Filename) + Path.DirectorySeparatorChar + "prj.adf") Then
                            Dim infile As New StreamReader(Path.GetDirectoryName(pRasGeoDataset.Filename) + Path.DirectorySeparatorChar + "prj.adf")
                            'TODO: Temporary measure that allows at least units to be recognized
                            If infile.ReadToEnd.Contains("METERS") Then
                                Return "units=m"
                            Else
                                Return "units=ft"
                            End If
                            Return "convert this prj.adf to real projection"
                        Else
                            Return ""
                        End If
                    Else
                        Return ""
                    End If
                End If
            Else
                Return ""
            End If
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Function

    Public Function ClipBySelectedPoly(ByRef pGridToClip As Grid, ByVal pSelectedPolyClip As MapWinGIS.Shape, ByVal outputFileName As String) As Grid
        Dim strtmp1 As String = Path.GetTempFileName
        g_TempFilesToDel.Add(strtmp1)
        strtmp1 = strtmp1 + OutputGridExt
        g_TempFilesToDel.Add(strtmp1)
        DataManagement.DeleteGrid(strtmp1)
        pGridToClip.Save()
        pGridToClip.Save(strtmp1)
        pGridToClip.Header.Projection = MapWindowPlugin.MapWindowInstance.Project.ProjectProjection

        SpatialOperations.ClipGridWithPolygon(strtmp1, pSelectedPolyClip, outputFileName)

        Dim out As New Grid
        out.Open(outputFileName)
        Return out
    End Function

    Public Function BrowseForFileName(ByRef strType As String) As String

        Dim pfilter As String

        Select Case strType
            Case "Feature"
                Dim shp As New Shapefile
                pfilter = shp.CdlgFilter
            Case "Raster"
                Dim g As New Grid
                pfilter = g.CdlgFilter
            Case Else
                pfilter = ""
        End Select

        Using dlg As New OpenFileDialog()
            dlg.Filter = pfilter
            dlg.Title = "Open " + strType
            If dlg.ShowDialog() = DialogResult.OK Then
                Return dlg.FileName
            Else
                Return ""
            End If
        End Using

    End Function

    Public Function ExportSelectedFeatures(ByVal SelectLyrPath As String, ByRef SelectedShapes As List(Of Integer)) As String
        ' Modified from http://www.mapwindow.org/wiki/index.php/MapWinGIS:SampleCode-VB_Net:ExportSelectedShapes
        Dim Result As Boolean

        Dim sFileName, sLayerType As String

        Dim myShapeFile, newShapefile As Shapefile
        Dim myShape As MapWinGIS.Shape
        Dim ShapefileType As ShpfileType
        Dim iShapeHandle, iFieldCnt As Integer

        ExportSelectedFeatures = Nothing
        'First check to see if any features have been selected
        If SelectedShapes Is Nothing OrElse SelectedShapes.Count = 0 Then
            Exit Function
        End If

        Try
            myShapeFile = New Shapefile
            myShapeFile.Open(SelectLyrPath)

            'Determine if shape is polygon, line, or point
            sLayerType = LCase(MapWindowPlugin.MapWindowInstance.Layers(MapWindowPlugin.MapWindowInstance.Layers.CurrentLayer).LayerType.ToString)
            If InStr(sLayerType, "line", CompareMethod.Text) > 0 Then
                ShapefileType = ShpfileType.SHP_POLYLINE
            ElseIf InStr(sLayerType, "polygon", CompareMethod.Text) > 0 Then
                ShapefileType = ShpfileType.SHP_POLYGON
            ElseIf InStr(sLayerType, "point", CompareMethod.Text) > 0 Then
                ShapefileType = ShpfileType.SHP_POINT
            End If

            sFileName = GetUniqueFileName("selpoly", g_strWorkspace, ".shp")

            'Create the new shapefile
            newShapefile = New Shapefile
            newShapefile.CreateNew(sFileName, ShapefileType)
            newShapefile.Projection = MapWindowPlugin.MapWindowInstance.Project.ProjectProjection

            'The new shapefile has no fields at this point
            For iFieldCnt = 0 To myShapeFile.NumFields - 1
                newShapefile.EditInsertField(myShapeFile.Field(iFieldCnt), iFieldCnt)
            Next iFieldCnt

            'Start an edit session in the shapefile
            newShapefile.StartEditingShapes(True, Nothing)

            'Iterate through each of the selected feature
            For i As Integer = 0 To SelectedShapes.Count - 1
                'Set to the selected shape
                myShape = myShapeFile.Shape(SelectedShapes(i))

                'insert the selected shape
                iShapeHandle = newShapefile.NumShapes
                Result = newShapefile.EditInsertShape(myShape, iShapeHandle)

                'Populate the aspatial data
                For iFieldCnt = 0 To myShapeFile.NumFields - 1
                    newShapefile.EditCellValue(iFieldCnt, iShapeHandle, myShapeFile.CellValue(iFieldCnt, SelectedShapes(i)))
                Next iFieldCnt
            Next i

            newShapefile.StopEditingShapes()
            newShapefile.Close()
            myShapeFile.Close()
            Return sFileName
        Catch ex As Exception
            MsgBox("Error in exporting selected features.", MsgBoxStyle.Exclamation, "Exporting Selected Error")
        End Try

    End Function

    Public Function AddOutputGridLayer(ByRef outRast As Grid, ByVal ColorString As String, ByVal UseStretch As Boolean, ByVal LayerName As String, ByVal OutputType As String, ByVal OutputGroup As Integer, ByRef OutputItems As OutputItems) As Boolean
        Dim cs As GridColorScheme
        If UseStretch = True Then
            cs = ReturnRasterStretchColorRampCS(outRast, ColorString)
            'Dim cs As MapWinGIS.GridColorScheme = ReturnContinuousRampColorCS(pPermAccumLocRunoffRaster, "Blue")
        Else
            cs = ReturnUniqueRasterRenderer(outRast, ColorString)
        End If

        Dim lyr As Layer = MapWindowPlugin.MapWindowInstance.Layers.Add(outRast, cs, LayerName)
        lyr.Visible = False
        If OutputGroup <> -1 Then
            lyr.MoveTo(0, OutputGroup)
        Else
            lyr.MoveTo(0, g_pGroupLayer)
        End If
        If Not OutputItems Is Nothing Then
            Dim item As New OutputItem
            item.strPath = outRast.Filename
            item.strName = LayerName
            item.strType = OutputType
            item.strColor = ColorString
            item.booUseStretch = UseStretch
            OutputItems.Add(item)
        End If
    End Function

    Public Delegate Function RasterMathCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single

    Public Delegate Function RasterMathCellCalcNulls(ByVal Input1 As Single, ByVal Input1Null As Single, ByVal Input2 As Single, ByVal Input2Null As Single, ByVal Input3 As Single, ByVal Input3Null As Single, ByVal Input4 As Single, ByVal Input4Null As Single, ByVal Input5 As Single, ByVal Input5Null As Single, ByVal OutNull As Single) As Single

    Public Sub RasterMath(ByRef InputGrid1 As Grid, ByRef InputGrid2 As Grid, ByRef Inputgrid3 As Grid, ByRef Inputgrid4 As Grid, ByRef Inputgrid5 As Grid, ByRef outputGrid As Grid, ByRef CellCalc As RasterMathCellCalc, Optional ByVal checkNullFirst As Boolean = True, Optional ByRef CellCalcNull As RasterMathCellCalcNulls = Nothing)


        Dim nodata1, nodata2, nodata3, nodata4, nodata5 As Single

        Dim tmppath As String = GetTempFileNameOutputGridExt()

        Dim nodataout As Single = -9999.0

        Dim grid1Header As GridHeader = InputGrid1.Header
        Dim outputHeader As GridHeader = New GridHeader
        outputHeader.CopyFrom(grid1Header)
        outputHeader.NodataValue = nodataout
        outputGrid = New Grid()
        outputGrid.CreateNew(tmppath, outputHeader, GridDataType.FloatDataType, outputHeader.NodataValue)
        outputGrid.Header.Projection = grid1Header.Projection
        nodata1 = grid1Header.NodataValue
        nodata2 = If(InputGrid2 Is Nothing, nodata1, InputGrid2.Header.NodataValue)
        nodata3 = If(Inputgrid3 Is Nothing, nodata1, Inputgrid3.Header.NodataValue)
        nodata4 = If(Inputgrid4 Is Nothing, nodata1, Inputgrid4.Header.NodataValue)
        nodata5 = If(Inputgrid5 Is Nothing, nodata1, Inputgrid5.Header.NodataValue)

        Dim ncol As Integer = grid1Header.NumberCols - 1
        Dim rowvals1(ncol) As Single
        Dim rowvals2(ncol) As Single
        Dim rowvals3(ncol) As Single
        Dim rowvals4(ncol) As Single
        Dim rowvals5(ncol) As Single
        Dim rowvalsout(ncol) As Single

        Dim nrow As Integer = grid1Header.NumberRows - 1
        For row As Integer = 0 To nrow
            InputGrid1.GetRow(row, rowvals1(0))

            'Dim grid = If(InputGrid2, InputGrid1).GetRow(row, rowvals2(0))
            If Not InputGrid2 Is Nothing Then InputGrid2.GetRow(row, rowvals2(0)) Else InputGrid1.GetRow(row, rowvals2(0))
            If Not Inputgrid3 Is Nothing Then Inputgrid3.GetRow(row, rowvals3(0)) Else InputGrid1.GetRow(row, rowvals3(0))
            If Not Inputgrid4 Is Nothing Then Inputgrid4.GetRow(row, rowvals4(0)) Else InputGrid1.GetRow(row, rowvals4(0))
            If Not Inputgrid5 Is Nothing Then Inputgrid5.GetRow(row, rowvals5(0)) Else InputGrid1.GetRow(row, rowvals5(0))

            For col As Integer = 0 To ncol
                If checkNullFirst Then
                    If rowvals1(col) = nodata1 OrElse rowvals2(col) = nodata2 OrElse rowvals3(col) = nodata3 OrElse rowvals4(col) = nodata4 OrElse rowvals5(col) = nodata5 Then
                        rowvalsout(col) = nodataout
                    Else
                        rowvalsout(col) = CellCalc.Invoke(rowvals1(col), rowvals2(col), rowvals3(col), rowvals4(col), rowvals5(col), nodataout)
                    End If
                Else
                    rowvalsout(col) = CellCalcNull.Invoke(rowvals1(col), nodata1, rowvals2(col), nodata2, rowvals3(col), nodata3, rowvals4(col), nodata4, rowvals5(col), nodata5, nodataout)
                End If
            Next

            outputGrid.PutRow(row, rowvalsout(0))
        Next
    End Sub

    Public Delegate Function RasterMathCellCalcWindow(ByRef InputBox1(,) As Single, ByRef InputBox2(,) As Single, ByRef InputBox3(,) As Single, ByRef InputBox4(,) As Single, ByRef InputBox5(,) As Single, ByVal OutNull As Single) As Single

    Public Delegate Function RasterMathCellCalcWindowNulls(ByRef InputBox1(,) As Single, ByVal Input1Null As Single, ByRef InputBox2(,) As Single, ByVal Input2Null As Single, ByRef InputBox3(,) As Single, ByVal Input3Null As Single, ByRef InputBox4(,) As Single, ByVal Input4Null As Single, ByRef InputBox5(,) As Single, ByVal Input5Null As Single, ByVal OutNull As Single) As Single
    Public Function GetUniqueFileName(ByRef Name As String, ByRef folderPath As String, ByVal Extension As String) As String
        Dim i As Integer = 0
        Dim nameAttempt As String

        Do
            i = i + 1
            nameAttempt = folderPath + Path.DirectorySeparatorChar + Name + i.ToString + Extension
        Loop While File.Exists(nameAttempt) And i < 1000

        If i < 1000 Then
            GetUniqueFileName = nameAttempt
        Else
            GetUniqueFileName = ""
        End If
    End Function

    Public Sub RasterMathWindow(ByRef InputGrid1 As Grid, ByRef InputGrid2 As Grid, ByRef Inputgrid3 As Grid, ByRef Inputgrid4 As Grid, ByRef Inputgrid5 As Grid, ByRef Outputgrid As Grid, ByRef CellCalcWindow As RasterMathCellCalcWindow, Optional ByVal checkNullFirst As Boolean = True, Optional ByRef CellCalcWindowNull As RasterMathCellCalcWindowNulls = Nothing)
        Dim head1, head2, head3, head4, head5, headnew As GridHeader
        Dim ncol As Integer
        Dim nrow As Integer
        Dim nodata1, nodata2, nodata3, nodata4, nodata5, nodataout As Single
        Dim rowvals1(3)(), rowvals2(3)(), rowvals3(3)(), rowvals4(3)(), rowvals5(3)(), rowvalsout() As Single
        Dim InputBox1(3, 3), InputBox2(3, 3), InputBox3(3, 3), InputBox4(3, 3), InputBox5(3, 3) As Single
        Dim tmppath As String = Path.GetTempFileName
        g_TempFilesToDel.Add(tmppath)
        tmppath = tmppath + OutputGridExt
        g_TempFilesToDel.Add(tmppath)

        nodataout = -9999.0

        head1 = InputGrid1.Header
        headnew = New GridHeader
        headnew.CopyFrom(head1)
        headnew.NodataValue = nodataout
        Outputgrid = New Grid()
        Outputgrid.CreateNew(tmppath, headnew, GridDataType.FloatDataType, headnew.NodataValue)
        Outputgrid.Header.Projection = head1.Projection
        ncol = head1.NumberCols - 1
        nrow = head1.NumberRows - 1
        nodata1 = head1.NodataValue
        If Not InputGrid2 Is Nothing Then
            head2 = InputGrid2.Header
            nodata2 = head2.NodataValue
        Else
            nodata2 = nodata1
        End If
        If Not Inputgrid3 Is Nothing Then
            head3 = Inputgrid3.Header
            nodata3 = head3.NodataValue
        Else
            nodata3 = nodata1
        End If
        If Not Inputgrid4 Is Nothing Then
            head4 = Inputgrid4.Header
            nodata4 = head4.NodataValue
        Else
            nodata4 = nodata1
        End If
        If Not Inputgrid5 Is Nothing Then
            head5 = Inputgrid5.Header
            nodata5 = head5.NodataValue
        Else
            nodata5 = nodata1
        End If
        For i As Integer = 0 To 2
            ReDim rowvals1(i)(ncol)
            ReDim rowvals2(i)(ncol)
            ReDim rowvals3(i)(ncol)
            ReDim rowvals4(i)(ncol)
            ReDim rowvals5(i)(ncol)
        Next

        ReDim rowvalsout(ncol)

        For row As Integer = 0 To nrow
            InputGrid1.GetRow(row, rowvals1(1)(0))
            If Not InputGrid2 Is Nothing Then
                InputGrid2.GetRow(row, rowvals2(1)(0))
            Else
                InputGrid1.GetRow(row, rowvals2(1)(0))
            End If
            If Not Inputgrid3 Is Nothing Then
                Inputgrid3.GetRow(row, rowvals3(1)(0))
            Else
                InputGrid1.GetRow(row, rowvals3(1)(0))
            End If
            If Not Inputgrid4 Is Nothing Then
                Inputgrid4.GetRow(row, rowvals4(1)(0))
            Else
                InputGrid1.GetRow(row, rowvals4(1)(0))
            End If
            If Not Inputgrid5 Is Nothing Then
                Inputgrid5.GetRow(row, rowvals5(1)(0))
            Else
                InputGrid1.GetRow(row, rowvals5(1)(0))
            End If

            If row <> 0 Then
                InputGrid1.GetRow(row - 1, rowvals1(0)(0))
                If Not InputGrid2 Is Nothing Then
                    InputGrid2.GetRow(row - 1, rowvals2(0)(0))
                Else
                    InputGrid1.GetRow(row - 1, rowvals2(0)(0))
                End If
                If Not Inputgrid3 Is Nothing Then
                    Inputgrid3.GetRow(row - 1, rowvals3(0)(0))
                Else
                    InputGrid1.GetRow(row - 1, rowvals3(0)(0))
                End If
                If Not Inputgrid4 Is Nothing Then
                    Inputgrid4.GetRow(row - 1, rowvals4(0)(0))
                Else
                    InputGrid1.GetRow(row - 1, rowvals4(0)(0))
                End If
                If Not Inputgrid5 Is Nothing Then
                    Inputgrid5.GetRow(row - 1, rowvals5(0)(0))
                Else
                    InputGrid1.GetRow(row - 1, rowvals5(0)(0))
                End If
            Else
                ReDim rowvals1(0)(rowvals1(1).Length)
                ReDim rowvals2(0)(rowvals1(1).Length)
                ReDim rowvals3(0)(rowvals1(1).Length)
                ReDim rowvals4(0)(rowvals1(1).Length)
                ReDim rowvals5(0)(rowvals1(1).Length)
            End If

            If row <> nrow Then
                InputGrid1.GetRow(row + 1, rowvals1(2)(0))
                If Not InputGrid2 Is Nothing Then
                    InputGrid2.GetRow(row + 1, rowvals2(2)(0))
                Else
                    InputGrid1.GetRow(row + 1, rowvals2(2)(0))
                End If
                If Not Inputgrid3 Is Nothing Then
                    Inputgrid3.GetRow(row + 1, rowvals3(2)(0))
                Else
                    InputGrid1.GetRow(row + 1, rowvals3(2)(0))
                End If
                If Not Inputgrid4 Is Nothing Then
                    Inputgrid4.GetRow(row + 1, rowvals4(2)(0))
                Else
                    InputGrid1.GetRow(row + 1, rowvals4(2)(0))
                End If
                If Not Inputgrid5 Is Nothing Then
                    Inputgrid5.GetRow(row + 1, rowvals5(2)(0))
                Else
                    InputGrid1.GetRow(row + 1, rowvals5(2)(0))
                End If
            Else
                ReDim rowvals1(2)(rowvals1(1).Length)
                ReDim rowvals2(2)(rowvals1(1).Length)
                ReDim rowvals3(2)(rowvals1(1).Length)
                ReDim rowvals4(2)(rowvals1(1).Length)
                ReDim rowvals5(2)(rowvals1(1).Length)
            End If

            For col As Integer = 1 To ncol - 1
                If row <> 0 Then
                    InputBox1(0, 0) = rowvals1(0)(col - 1)
                    InputBox1(0, 1) = rowvals1(0)(col)
                    InputBox1(0, 2) = rowvals1(0)(col + 1)
                    InputBox2(0, 0) = rowvals2(0)(col - 1)
                    InputBox2(0, 1) = rowvals2(0)(col)
                    InputBox2(0, 2) = rowvals2(0)(col + 1)
                    InputBox3(0, 0) = rowvals3(0)(col - 1)
                    InputBox3(0, 1) = rowvals3(0)(col)
                    InputBox3(0, 2) = rowvals3(0)(col + 1)
                    InputBox4(0, 0) = rowvals4(0)(col - 1)
                    InputBox4(0, 1) = rowvals4(0)(col)
                    InputBox4(0, 2) = rowvals4(0)(col + 1)
                    InputBox5(0, 0) = rowvals5(0)(col - 1)
                    InputBox5(0, 1) = rowvals5(0)(col)
                    InputBox5(0, 2) = rowvals5(0)(col + 1)
                Else
                    InputBox1(0, 0) = nodata1
                    InputBox1(0, 1) = nodata1
                    InputBox1(0, 2) = nodata1
                    InputBox2(0, 0) = nodata2
                    InputBox2(0, 1) = nodata2
                    InputBox2(0, 2) = nodata2
                    InputBox3(0, 0) = nodata3
                    InputBox3(0, 1) = nodata3
                    InputBox3(0, 2) = nodata3
                    InputBox4(0, 0) = nodata4
                    InputBox4(0, 1) = nodata4
                    InputBox4(0, 2) = nodata4
                    InputBox5(0, 0) = nodata5
                    InputBox5(0, 1) = nodata5
                    InputBox5(0, 2) = nodata5
                End If

                InputBox1(1, 0) = rowvals1(1)(col - 1)
                InputBox1(1, 1) = rowvals1(1)(col)
                InputBox1(1, 2) = rowvals1(1)(col + 1)
                InputBox2(1, 0) = rowvals2(1)(col - 1)
                InputBox2(1, 1) = rowvals2(1)(col)
                InputBox2(1, 2) = rowvals2(1)(col + 1)
                InputBox3(1, 0) = rowvals3(1)(col - 1)
                InputBox3(1, 1) = rowvals3(1)(col)
                InputBox3(1, 2) = rowvals3(1)(col + 1)
                InputBox4(1, 0) = rowvals4(1)(col - 1)
                InputBox4(1, 1) = rowvals4(1)(col)
                InputBox4(1, 2) = rowvals4(1)(col + 1)
                InputBox5(1, 0) = rowvals5(1)(col - 1)
                InputBox5(1, 1) = rowvals5(1)(col)
                InputBox5(1, 2) = rowvals5(1)(col + 1)

                If row <> 0 Then
                    InputBox1(2, 0) = rowvals1(2)(col - 1)
                    InputBox1(2, 1) = rowvals1(2)(col)
                    InputBox1(2, 2) = rowvals1(2)(col + 1)
                    InputBox2(2, 0) = rowvals2(2)(col - 1)
                    InputBox2(2, 1) = rowvals2(2)(col)
                    InputBox2(2, 2) = rowvals2(2)(col + 1)
                    InputBox3(2, 0) = rowvals3(2)(col - 1)
                    InputBox3(2, 1) = rowvals3(2)(col)
                    InputBox3(2, 2) = rowvals3(2)(col + 1)
                    InputBox4(2, 0) = rowvals4(2)(col - 1)
                    InputBox4(2, 1) = rowvals4(2)(col)
                    InputBox4(2, 2) = rowvals4(2)(col + 1)
                    InputBox5(2, 0) = rowvals5(2)(col - 1)
                    InputBox5(2, 1) = rowvals5(2)(col)
                    InputBox5(2, 2) = rowvals5(2)(col + 1)
                Else
                    InputBox1(2, 0) = nodata1
                    InputBox1(2, 1) = nodata1
                    InputBox1(2, 2) = nodata1
                    InputBox2(2, 0) = nodata2
                    InputBox2(2, 1) = nodata2
                    InputBox2(2, 2) = nodata2
                    InputBox3(2, 0) = nodata3
                    InputBox3(2, 1) = nodata3
                    InputBox3(2, 2) = nodata3
                    InputBox4(2, 0) = nodata4
                    InputBox4(2, 1) = nodata4
                    InputBox4(2, 2) = nodata4
                    InputBox5(2, 0) = nodata5
                    InputBox5(2, 1) = nodata5
                    InputBox5(2, 2) = nodata5
                End If

                If checkNullFirst Then
                    If InputBox1(0, 0) = nodata1 OrElse InputBox1(0, 1) = nodata1 OrElse InputBox1(0, 2) = nodata1 OrElse InputBox2(0, 0) = nodata2 OrElse InputBox2(0, 1) = nodata2 OrElse InputBox2(0, 2) = nodata2 OrElse InputBox3(0, 0) = nodata3 OrElse InputBox3(0, 1) = nodata3 OrElse InputBox3(0, 2) = nodata3 OrElse InputBox4(0, 0) = nodata4 OrElse InputBox4(0, 1) = nodata4 OrElse InputBox4(0, 2) = nodata4 OrElse InputBox5(0, 0) = nodata5 OrElse InputBox5(0, 1) = nodata5 OrElse InputBox5(0, 2) = nodata5 OrElse InputBox1(1, 0) = nodata1 OrElse InputBox1(1, 1) = nodata1 OrElse InputBox1(1, 2) = nodata1 OrElse InputBox2(1, 0) = nodata2 OrElse InputBox2(1, 1) = nodata2 OrElse InputBox2(1, 2) = nodata2 OrElse InputBox3(1, 0) = nodata3 OrElse InputBox3(1, 1) = nodata3 OrElse InputBox3(1, 2) = nodata3 OrElse InputBox4(1, 0) = nodata4 OrElse InputBox4(1, 1) = nodata4 OrElse InputBox4(1, 2) = nodata4 OrElse InputBox5(1, 0) = nodata5 OrElse InputBox5(1, 1) = nodata5 OrElse InputBox5(1, 2) = nodata5 OrElse InputBox1(2, 0) = nodata1 OrElse InputBox1(2, 1) = nodata1 OrElse InputBox1(2, 2) = nodata1 OrElse InputBox2(2, 0) = nodata2 OrElse InputBox2(2, 1) = nodata2 OrElse InputBox2(2, 2) = nodata2 OrElse InputBox3(2, 0) = nodata3 OrElse InputBox3(2, 1) = nodata3 OrElse InputBox3(2, 2) = nodata3 OrElse InputBox4(2, 0) = nodata4 OrElse InputBox4(2, 1) = nodata4 OrElse InputBox4(2, 2) = nodata4 OrElse InputBox5(2, 0) = nodata5 OrElse InputBox5(2, 1) = nodata5 OrElse InputBox5(2, 2) = nodata5 Then
                        rowvalsout(col) = nodataout
                    Else
                        rowvalsout(col) = CellCalcWindow.Invoke(InputBox1, InputBox2, InputBox3, InputBox4, InputBox5, nodataout)
                    End If
                Else
                    rowvalsout(col) = CellCalcWindowNull.Invoke(InputBox1, nodata1, InputBox2, nodata2, InputBox3, nodata3, InputBox4, nodata4, InputBox5, nodata5, nodataout)
                End If
            Next

            Outputgrid.PutRow(row, rowvalsout(0))
        Next
    End Sub
End Module
