﻿Imports System.Data.OleDb
Imports System.Windows.Forms

Module modUtil
    Public g_nspectPath As String
    Public g_nspectDocPath As String
    Public g_strWorkspace As String

    Public g_cb As frmProjectSetup

    Public g_strSelectedExportPath As String = ""

    'Database Variables
    Public g_DBConn As OleDbConnection 'Connection
    Public g_strConn As String 'Connection String
    Public g_boolConnected As Boolean 'Bool: connected

    Public g_boolAddCoeff As Boolean
    Public g_boolCopyCoeff As Boolean 'True: called frmPollutants, False: called frmNewPollutants
    Public g_boolAgree As Boolean 'True: use the Agree Function on Streams.
    Public g_boolHydCorr As Boolean 'True: Hyrdologically Correct DEM, no fill needed
    Public g_boolNewWShed As Boolean 'True: New WaterShed form called from frmPrj


    'WqStd

    'Agree DEM Stuff
    Public g_boolParams As Boolean 'Flag to indicate Agree params have been entered


    'Project Form Variables
    Public g_strPrjFileName As String 'Project file name

    'Management Scenario variables::frmPrjCalc
    Public g_strLUScenFileName As String 'Management scenario file name
    Public g_intManScenRow As String 'Management scenario ROW number

    'Pollutant Coefficient variable::frmPrjCalc
    Public g_intCoeffRow As Short 'Coeff Row Number
    Public g_strCoeffCalc As String 'if the Calc option is chosen, hold results in string


    Const c_sModuleFileName As String = "modUtil.vb"
    Private m_ParentHWND As Integer ' Set this to get correct parenting of Error handler forms



    'Function for connection to NSPECT.mdb: fires on dll load
    Public Sub DBConnection()
        Try
            If Not g_boolConnected Then
                'TODO: check for location of file and prompt if not found
                g_strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & g_nspectPath & "\nspect.mdb"

                g_DBConn = New OleDbConnection(g_strConn)

                g_DBConn.Open()

                g_boolConnected = True

            End If
            Exit Sub
        Catch ex As Exception
            MsgBox(Err.Number & Err.Description & " Error connecting to database, please check NSPECTDAT enviornment variable.  Current value of NSPECTDAT: " & g_strConn, MsgBoxStyle.Critical, "Error Connecting")

        End Try
    End Sub



    Public Sub InitComboBox(ByRef cbo As System.Windows.Forms.ComboBox, ByRef strName As String)
        Try
            'Loads the variety of comboboxes throught the project using combobox and name of table
            Dim rsNamesCmd As OleDbCommand
            Dim rsNames As OleDbDataReader
            Dim strSelectStatement As String

            strSelectStatement = "SELECT NAME FROM " & strName & " ORDER BY NAME ASC"

            'Check thrown in to make sure g_ADOconn is something, in v9.1 we started having problems.
            If Not g_boolConnected Then
                DBConnection()
            End If

            rsNamesCmd = New OleDbCommand(strSelectStatement, g_DBConn)

            rsNames = rsNamesCmd.ExecuteReader()

            If rsNames.HasRows Then
                With cbo
                    Do While rsNames.Read()
                        .Items.Add(rsNames.Item("Name"))
                    Loop
                End With

                cbo.SelectedIndex = 0
            Else
                MsgBox("Warning.  There are no records remaining.  Please add a new one.", MsgBoxStyle.Critical, "Recordset Empty")
                Exit Sub
            End If

            'Cleanup
            rsNames.Close()
        Catch ex As Exception
            HandleError(True, "InitComboBox " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Sub


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
        For i As Integer = 0 To g_MapWin.Layers.NumLayers - 1
            If g_MapWin.Layers(i).Name = strLayerName Then
                Dim proj As String = g_MapWin.Layers(i).Projection
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


    'General Function used to simply get the index of combobox entries
    Public Function GetCboIndex(ByRef strList As String, ByRef cbo As System.Windows.Forms.ComboBox) As Short
        Try
            Dim i As Short
            i = 0

            For i = 0 To cbo.Items.Count - 1
                If cbo.Items(i) = strList Then
                    GetCboIndex = i
                End If
            Next i
        Catch ex As Exception
            HandleError(True, "GetCboIndex " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Function

    Public Function LayerInMap(ByRef strName As String) As Boolean
        Try


            For lngLyrIndex As Integer = 0 To g_MapWin.Layers.NumLayers - 1
                Dim pLayer As MapWindow.Interfaces.Layer = g_MapWin.Layers(lngLyrIndex)
                If pLayer.Name = strName Then
                    LayerInMap = True
                    Exit Function
                Else
                    LayerInMap = False
                End If
            Next
        Catch ex As Exception
            HandleError(True, "LayerInMap " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Function


    Public Function LayerInMapByFileName(ByRef strName As String) As Boolean

        For lngLyrIndex As Integer = 0 To g_MapWin.Layers.NumLayers - 1
            Dim pLayer As MapWindow.Interfaces.Layer = g_MapWin.Layers(lngLyrIndex)

            If Trim(LCase(pLayer.FileName)) <> Trim(LCase(strName)) Then
                LayerInMapByFileName = False
            Else
                LayerInMapByFileName = True
                Exit For
            End If
        Next
    End Function

    Public Function GetLayerIndex(ByRef strLayerName As String) As Integer
        Try
            For lngLyrIndex As Integer = 0 To g_MapWin.Layers.NumLayers - 1
                Dim pLayer As MapWindow.Interfaces.Layer = g_MapWin.Layers(lngLyrIndex)

                If Trim(LCase(pLayer.Name)) = Trim(LCase(strLayerName)) Then
                    GetLayerIndex = lngLyrIndex
                    Exit For
                End If
            Next
        Catch ex As Exception
            HandleError(True, "GetLayerIndex " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
        GetLayerIndex = -1
    End Function

    Public Function GetLayerFilename(ByRef strLayerName As String) As String
        GetLayerFilename = ""
        Try
            For lngLyrIndex As Integer = 0 To g_MapWin.Layers.NumLayers - 1
                Dim pLayer As MapWindow.Interfaces.Layer = g_MapWin.Layers(lngLyrIndex)

                If Trim(LCase(pLayer.Name)) = Trim(LCase(strLayerName)) Then
                    GetLayerFilename = pLayer.FileName

                    If GetLayerFilename.EndsWith("sta.adf") Then
                        GetLayerFilename = IO.Path.GetDirectoryName(GetLayerFilename)
                    End If

                    Exit For
                End If
            Next
        Catch ex As Exception
            HandleError(True, "GetLayerFilename " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Function


    Public Function AddRasterLayerToMapFromFileName(ByRef strName As String) As Boolean
        Try

            If IO.Path.GetExtension(strName) <> "" Then
                g_MapWin.Layers.Add(strName)
            Else
                g_MapWin.Layers.Add(strName + "\sta.adf")
            End If
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

    Public Function AddFeatureLayerToMapFromFileName(ByRef strName As String, Optional ByRef strLyrName As String = "") As Boolean
        Try
            If IO.Path.GetExtension(strName) <> "" Then
                If strLyrName <> "" Then
                    g_MapWin.Layers.Add(strName, strLyrName)
                Else
                    g_MapWin.Layers.Add(strName)
                End If
            Else
                If strLyrName <> "" Then
                    g_MapWin.Layers.Add(strName + ".shp", strLyrName)
                Else
                    g_MapWin.Layers.Add(strName + ".shp")
                End If
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
                    Dim shp As New MapWinGIS.Shapefile
                    pFilter = shp.CdlgFilter
                Case "Raster"
                    Dim g As New MapWinGIS.Grid
                    pFilter = g.CdlgFilter
            End Select

            Dim dlgopen As New System.Windows.Forms.OpenFileDialog
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

    Public Function FeatureExists(ByRef strFeatureFileName As String) As Boolean

        If IO.Path.GetExtension(strFeatureFileName) = "" Then
            Return IO.File.Exists(strFeatureFileName + ".shp")
        Else
            Return IO.File.Exists(strFeatureFileName)
        End If

    End Function

    Public Function RasterExists(ByRef strRasterFileName As String) As Boolean
        If IO.Path.GetExtension(strRasterFileName) = "" Then
            Return IO.File.Exists(strRasterFileName + "\sta.adf")
        Else
            Return IO.File.Exists(strRasterFileName)
        End If

    End Function

    Public Function ReturnFeature(ByRef strFeatureFileName As String) As MapWinGIS.Shapefile

        If IO.Path.GetExtension(strFeatureFileName) = "" Then
            If IO.File.Exists(strFeatureFileName + ".shp") Then
                Dim sf As New MapWinGIS.Shapefile
                If sf.Open(strFeatureFileName + ".shp") Then
                    Return sf
                Else
                    Return Nothing
                End If
            Else
                Return Nothing
            End If
        Else
            If IO.File.Exists(strFeatureFileName) Then
                Dim sf As New MapWinGIS.Shapefile
                If sf.Open(strFeatureFileName) Then
                    Return sf
                Else
                    Return Nothing
                End If
            Else
                Return Nothing
            End If
        End If

    End Function

    Public Function ReturnRaster(ByRef strRasterFileName As String) As MapWinGIS.Grid
        If IO.Path.GetExtension(strRasterFileName) = "" Then
            If IO.File.Exists(strRasterFileName + "\sta.adf") Then
                Dim g As New MapWinGIS.Grid
                If g.Open(strRasterFileName + "\sta.adf") Then
                    Return g
                Else
                    Return Nothing
                End If
            Else
                Return Nothing
            End If
        Else
            If IO.File.Exists(strRasterFileName) Then
                Dim g As New MapWinGIS.Grid
                If g.Open(strRasterFileName) Then
                    Return g
                Else
                    Return Nothing
                End If
            Else
                Return Nothing
            End If
        End If

    End Function

    Public Function ReturnPermanentRaster(ByRef pRaster As MapWinGIS.Grid, ByRef sOutputName As String) As MapWinGIS.Grid
        pRaster.Save(sOutputName)
        Return pRaster
    End Function

    Public Function ReturnRasterStretchColorRampCS(ByRef pRaster As MapWinGIS.Grid, ByRef strColor As String) As MapWinGIS.GridColorScheme
        Try
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
                Case Else
                    rFrom = CInt(Split(strColor, ",")(0))
                    gFrom = CInt(Split(strColor, ",")(1))
                    bFrom = CInt(Split(strColor, ",")(2))


                    rTo = CInt(Split(strColor, ",")(3))
                    gTo = CInt(Split(strColor, ",")(4))
                    bTo = CInt(Split(strColor, ",")(5))
            End Select



            Dim cs As New MapWinGIS.GridColorScheme
            Dim csbrk As New MapWinGIS.GridColorBreak
            csbrk.LowValue = pRaster.Minimum
            csbrk.HighValue = pRaster.Maximum
            csbrk.LowColor = Convert.ToUInt32(RGB(rFrom, gFrom, bFrom))
            csbrk.HighColor = Convert.ToUInt32(RGB(rTo, gTo, bTo))
            csbrk.Caption = csbrk.LowValue.ToString() + " - " + csbrk.HighValue.ToString()
            cs.InsertBreak(csbrk)

            Return cs
        Catch ex As Exception
            HandleError(True, "ReturnRasterStretchColorRampRender " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Function

    
    Public Function ClipBySelectedPoly(ByRef pAccumRunoffRaster As MapWinGIS.Grid, ByVal g_pSelectedPolyClip As MapWinGIS.Shape, ByVal outputFileName As String) As MapWinGIS.Grid
        Dim strtmp1 As String = IO.Path.GetTempFileName
        MapWinGeoProc.DataManagement.DeleteGrid(strtmp1 + ".bgd")
        pAccumRunoffRaster.Save(+".bgd")

        MapWinGeoProc.SpatialOperations.ClipGridWithPolygon(strtmp1, g_pSelectedPolyClip, outputFileName)

        Dim out As New MapWinGIS.Grid
        out.Open(outputFileName)
        Return out
    End Function



    Public Function BrowseForFileName(ByRef strType As String, ByRef frm As System.Windows.Forms.Form, ByRef strTitle As String) As String

        Dim pfilter As String

        Select Case strType
            Case "Feature"
                Dim shp As New MapWinGIS.Shapefile
                pfilter = shp.CdlgFilter
            Case "Raster"
                Dim g As New MapWinGIS.Grid
                pfilter = g.CdlgFilter
            Case Else
                pfilter = ""
        End Select

        Dim dlg As New Windows.Forms.OpenFileDialog
        dlg.Filter = pfilter
        dlg.Title = "Open " + strType
        If dlg.ShowDialog() = DialogResult.OK Then
            Return dlg.FileName
        Else
            Return ""
        End If


    End Function


    Public Sub CleanupRasterFolder(ByRef strWorkspacePath As String)
        Try
            'Used to cleanup the User's workspace and avoid the dreaded -2147467259 error

            'Dim pWorkspaceFactory As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory
            'Dim pWorkspace As ESRI.ArcGIS.Geodatabase.IWorkspace
            'Dim pDataset As ESRI.ArcGIS.Geodatabase.IDataset
            'Dim pEnumRasterDataset As ESRI.ArcGIS.Geodatabase.IEnumDataset
            'Dim pEnv As ESRI.ArcGIS.GeoAnalyst.IRasterAnalysisEnvironment

            'pWorkspaceFactory = New ESRI.ArcGIS.DataSourcesRaster.RasterWorkspaceFactory
            'pWorkspace = pWorkspaceFactory.OpenFromFile(strWorkspacePath, 0)

            'pEnumRasterDataset = pWorkspace.Datasets(ESRI.ArcGIS.Geodatabase.esriDatasetType.esriDTRasterDataset)
            'pEnv = New ESRI.ArcGIS.GeoAnalyst.RasterAnalysis

            'pDataset = pEnumRasterDataset.Next

            'Do While Not pDataset Is Nothing
            '    If InStr(1, pDataset.Name, pEnv.DefaultOutputRasterPrefix, CompareMethod.Text) > 0 Then
            '        If pDataset.CanDelete Then
            '            pDataset.Delete()
            '        End If
            '    End If
            '    pDataset = pEnumRasterDataset.Next
            'Loop

        Catch ex As Exception
            HandleError(True, "CleanupRasterFolder " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Sub

    Public Sub CleanGlobals()
        Try
            ''Sub to rid the world of stragling GRIDS, i.e. the ones established for global usse

            'If Not modMainRun.g_pFeatWorkspace Is Nothing Then
            '    modMainRun.g_pFeatWorkspace = Nothing
            'End If

            'If Not modMainRun.g_pRasWorkspace Is Nothing Then
            '    modMainRun.g_pRasWorkspace = Nothing
            'End If

            ''Had an 'elegant' solution using an Iarray to hold global rasters, but that didn't seem to do the
            ''job, so we have to manually set each and everyone to nothing

            'If Not g_pSCS100Raster Is Nothing Then
            '    g_pSCS100Raster = Nothing
            'End If

            'If Not g_pAbstractRaster Is Nothing Then
            '    g_pAbstractRaster = Nothing
            'End If

            'If Not g_pRunoffRaster Is Nothing Then
            '    g_pRunoffRaster = Nothing
            'End If

            'If Not g_pRunoffInchRaster Is Nothing Then
            '    g_pRunoffInchRaster = Nothing
            'End If

            'If Not g_pCellAreaSqMiRaster Is Nothing Then
            '    g_pCellAreaSqMiRaster = Nothing
            'End If

            'If Not g_pRunoffCFRaster Is Nothing Then
            '    g_pRunoffCFRaster = Nothing
            'End If

            'If Not g_pRunoffAFRaster Is Nothing Then
            '    g_pRunoffAFRaster = Nothing
            'End If

            'If Not g_pMetRunoffRaster Is Nothing Then
            '    g_pMetRunoffRaster = Nothing
            'End If

            'If Not g_pRunoffRaster Is Nothing Then
            '    g_pRunoffRaster = Nothing
            'End If

            'If Not g_pDEMRaster Is Nothing Then
            '    g_pDEMRaster = Nothing
            'End If

            'If Not g_pFlowAccRaster Is Nothing Then
            '    g_pFlowAccRaster = Nothing
            'End If

            'If Not g_pFlowDirRaster Is Nothing Then
            '    g_pFlowDirRaster = Nothing
            'End If

            'If Not g_pLSRaster Is Nothing Then
            '    g_pLSRaster = Nothing
            'End If

            'If Not g_pWaterShedFeatClass Is Nothing Then
            '    g_pWaterShedFeatClass = Nothing
            'End If

            'If Not g_KFactorRaster Is Nothing Then
            '    g_KFactorRaster = Nothing
            'End If

            'If Not g_pPrecipRaster Is Nothing Then
            '    g_pPrecipRaster = Nothing
            'End If

            'If Not g_pSoilsRaster Is Nothing Then
            '    g_pSoilsRaster = Nothing
            'End If

            'If Not g_LandCoverRaster Is Nothing Then
            '    g_LandCoverRaster = Nothing
            'End If

            'If Not g_pSelectedPolyClip Is Nothing Then
            '    g_pSelectedPolyClip = Nothing
            'End If
        Catch ex As Exception
            HandleError(True, "CleanGlobals " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Sub


    Public Function UniqueName(ByRef strTableName As String, ByRef strName As String) As Boolean
        Try

            Dim strCmdText As String

            strCmdText = "SELECT * FROM " & strTableName & " WHERE NAME LIKE '" & strName & "'"
            Dim cmdName As New OleDbCommand(strCmdText, g_DBConn)
            Dim datName As OleDbDataReader = cmdName.ExecuteReader()
            If datName.HasRows Then
                UniqueName = False
            Else
                UniqueName = True
            End If
            datName.Close()
        Catch ex As Exception
            HandleError(True, "UniqueName " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Function

    'Tests name inputs to insure unique values for databases
    Public Function CreateUniqueName(ByRef strTableName As String, ByRef strName As String) As String
        CreateUniqueName = ""
        Try
            Dim strCmdText As String
            Dim sCurrNum As String
            Dim strCurrNameRecord As String
            strCmdText = "SELECT * FROM " & strTableName '& " WHERE NAME LIKE '" & strName & "'"
            Dim cmd As New OleDbCommand(strCmdText, g_DBConn)
            Dim data As OleDbDataReader = cmd.ExecuteReader
            sCurrNum = "0"

            While data.Read()
                strCurrNameRecord = CStr(data("Name"))
                If InStr(1, strCurrNameRecord, strName, 1) > 0 Then
                    If IsNumeric(Right(strCurrNameRecord, 2)) Then
                        If (CShort(Right(strCurrNameRecord, 2)) > CShort(sCurrNum)) Then
                            sCurrNum = Right(strCurrNameRecord, 2)
                        Else
                            Exit While
                        End If
                    Else
                        If IsNumeric(Right(strCurrNameRecord, 1)) Then
                            If (CShort(Right(strCurrNameRecord, 1)) > CShort(sCurrNum)) Then
                                sCurrNum = Right(strCurrNameRecord, 1)
                            End If
                        End If
                    End If
                End If
            End While

            If sCurrNum = "0" Then
                CreateUniqueName = strName & "1"
            Else
                CreateUniqueName = strName & CStr(CShort(sCurrNum) + 1)
            End If

            data.Close()

        Catch ex As Exception
            HandleError(True, "CreateUniqueName " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Function


    Public Function GetUniqueName(ByRef Name As String, ByRef folderPath As String, ByVal Extension As String) As String
        GetUniqueName = ""
        Dim i As Integer = 0
        Dim nameAttempt As String

        Do
            i = i + 1
            nameAttempt = folderPath + IO.Path.DirectorySeparatorChar + Name + i.ToString + Extension
        Loop While IO.File.Exists(nameAttempt) And i < 1000

        If i < 1000 Then
            GetUniqueName = nameAttempt
        End If
    End Function

    Public Function ExportSelectedFeatures() As String
        ' Modified from http://www.mapwindow.org/wiki/index.php/MapWinGIS:SampleCode-VB_Net:ExportSelectedShapes
        Dim Result As Boolean
        Dim SelectedShape As MapWindow.Interfaces.SelectedShape
        Dim cdlSave As New SaveFileDialog
        Dim sFileName, sLayerType As String
        Dim iFileCnt As Integer = 1
        Dim myShapeFile, newShapefile As MapWinGIS.Shapefile
        Dim myShape As MapWinGIS.Shape
        Dim ShapefileType As MapWinGIS.ShpfileType
        Dim iShapeHandle, iFieldCnt As Integer

        ExportSelectedFeatures = Nothing
        If g_strSelectedExportPath <> "" Then
            Return g_strSelectedExportPath
        Else
            'First check to see if any features have been selected
            If g_MapWin.View.SelectedShapes.NumSelected <= 0 Then Exit Function

            Try
                myShapeFile = g_MapWin.Layers(g_MapWin.Layers.CurrentLayer).GetObject()

                'Determine if shape is polygon, line, or point
                sLayerType = Strings.LCase(g_MapWin.Layers(g_MapWin.Layers.CurrentLayer).LayerType.ToString)
                If InStr(sLayerType, "line", CompareMethod.Text) > 0 Then
                    ShapefileType = MapWinGIS.ShpfileType.SHP_POLYLINE
                ElseIf InStr(sLayerType, "polygon", CompareMethod.Text) > 0 Then
                    ShapefileType = MapWinGIS.ShpfileType.SHP_POLYGON
                ElseIf InStr(sLayerType, "point", CompareMethod.Text) > 0 Then
                    ShapefileType = MapWinGIS.ShpfileType.SHP_POINT
                End If

                sFileName = modUtil.GetUniqueName("selpoly", g_strWorkspace, ".shp")

                'Create the new shapefile
                newShapefile = New MapWinGIS.Shapefile
                newShapefile.CreateNew(sFileName, ShapefileType)

                'The new shapefile has no fields at this point
                For iFieldCnt = 0 To myShapeFile.NumFields - 1
                    newShapefile.EditInsertField(myShapeFile.Field(iFieldCnt), iFieldCnt)
                Next iFieldCnt

                'Start an edit session in the shapefile
                newShapefile.StartEditingShapes(True, Nothing)

                'Iterate through each of the selected feature
                For i As Integer = 0 To g_MapWin.View.SelectedShapes.NumSelected - 1
                    'Set to the selected shape
                    SelectedShape = g_MapWin.View.SelectedShapes(i)
                    myShape = myShapeFile.Shape(SelectedShape.ShapeIndex)

                    'insert the selected shape
                    iShapeHandle = newShapefile.NumShapes
                    Result = newShapefile.EditInsertShape(myShape, iShapeHandle)

                    'Populate the aspatial data
                    For iFieldCnt = 0 To myShapeFile.NumFields - 1
                        newShapefile.EditCellValue(iFieldCnt, iShapeHandle, myShapeFile.CellValue(iFieldCnt, SelectedShape.ShapeIndex))
                    Next iFieldCnt
                Next i

                newShapefile.StopEditingShapes()
                newShapefile.Close()
                Return sFileName
            Catch ex As Exception
                MsgBox("Error in exporting selected features.", MsgBoxStyle.Exclamation, "Exporting Selected Error")
            End Try
        End If


    End Function


    Public Delegate Function RasterMathCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
    Public Delegate Function RasterMathCellCalcNulls(ByVal Input1 As Single, ByVal Input1Null As Single, ByVal Input2 As Single, ByVal Input2Null As Single, ByVal Input3 As Single, ByVal Input3Null As Single, ByVal Input4 As Single, ByVal Input4Null As Single, ByVal Input5 As Single, ByVal Input5Null As Single) As Single

    Public Sub RasterMath(ByRef InputGrid1 As MapWinGIS.Grid, ByRef InputGrid2 As MapWinGIS.Grid, ByRef Inputgrid3 As MapWinGIS.Grid, ByRef Inputgrid4 As MapWinGIS.Grid, ByRef Inputgrid5 As MapWinGIS.Grid, ByRef Outputgrid As MapWinGIS.Grid, ByRef CellCalc As RasterMathCellCalc, Optional ByVal checkNullFirst As Boolean = True, Optional ByRef CellCalcNull As RasterMathCellCalcNulls = Nothing)
        Dim head1, head2, head3, head4, head5, headnew As MapWinGIS.GridHeader
        Dim ncol As Integer
        Dim nrow As Integer
        Dim nodata1, nodata2, nodata3, nodata4, nodata5 As Single
        Dim rowvals1(), rowvals2(), rowvals3(), rowvals4(), rowvals5(), rowvalsout() As Single

        head1 = InputGrid1.Header
        headnew = New MapWinGIS.GridHeader
        headnew.CopyFrom(head1)
        Outputgrid = New MapWinGIS.Grid()
        Outputgrid.CreateNew("", headnew, MapWinGIS.GridDataType.FloatDataType, headnew.NodataValue)
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
            nodata4 = nodata1
        End If
        ReDim rowvals1(ncol)
        ReDim rowvals2(ncol)
        ReDim rowvals3(ncol)
        ReDim rowvals4(ncol)
        ReDim rowvals5(ncol)
        ReDim rowvalsout(ncol)

        For row As Integer = 0 To nrow
            InputGrid1.GetRow(row, rowvals1(0))
            If Not InputGrid2 Is Nothing Then
                InputGrid2.GetRow(row, rowvals2(0))
            Else
                InputGrid1.GetRow(row, rowvals2(0))
            End If
            If Not Inputgrid3 Is Nothing Then
                Inputgrid3.GetRow(row, rowvals3(0))
            Else
                InputGrid1.GetRow(row, rowvals3(0))
            End If
            If Not Inputgrid4 Is Nothing Then
                Inputgrid4.GetRow(row, rowvals4(0))
            Else
                InputGrid1.GetRow(row, rowvals4(0))
            End If
            If Not Inputgrid5 Is Nothing Then
                Inputgrid5.GetRow(row, rowvals5(0))
            Else
                InputGrid1.GetRow(row, rowvals5(0))
            End If

            For col As Integer = 0 To ncol
                If checkNullFirst Then
                    If rowvals1(col) <> nodata1 And rowvals2(col) <> nodata2 And rowvals3(col) <> nodata3 And rowvals4(col) <> nodata4 Then
                        rowvalsout(col) = CellCalc.Invoke(rowvals1(col), rowvals2(col), rowvals3(col), rowvals4(col), rowvals5(col))
                    Else
                        rowvalsout(col) = nodata1
                    End If
                Else
                    rowvalsout(col) = CellCalcNull.Invoke(rowvals1(col), nodata1, rowvals2(col), nodata2, rowvals3(col), nodata3, rowvals4(col), nodata4, rowvals5(col), nodata5)
                End If
            Next

            Outputgrid.PutRow(row, rowvalsout(0))
        Next
    End Sub
End Module
