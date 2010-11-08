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

Imports System.Data.OleDb
Imports System.Windows.Forms

Module modUtil
    Public g_nspectPath As String
    Public g_nspectDocPath As String
    Public g_strWorkspace As String

    Public g_CurrentProjectPath As String

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

    Public g_frmProjectSetup As Windows.Forms.Form


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
                If IO.File.Exists(strName) Then
                    If strLyrName <> "" Then
                        g_MapWin.Layers.Add(strName, strLyrName)
                    Else
                        g_MapWin.Layers.Add(strName)
                    End If
                Else
                    Return False
                End If
            Else
                If IO.File.Exists(strName + ".shp") Then
                    If strLyrName <> "" Then
                        g_MapWin.Layers.Add(strName + ".shp", strLyrName)
                    Else
                        g_MapWin.Layers.Add(strName + ".shp")
                    End If
                Else
                    Return False
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

    Public Function AddInputFromGxBrowserText(ByRef txtInput As System.Windows.Forms.TextBox, ByRef strTitle As String, ByRef frm As System.Windows.Forms.Form, ByRef intType As Short) As MapWinGIS.Grid
        AddInputFromGxBrowserText = Nothing
        Try
            Dim dlgOpen As New Windows.Forms.OpenFileDialog
            Dim g As New MapWinGIS.Grid
            dlgOpen.Filter = g.CdlgFilter
            dlgOpen.Title = strTitle
            If dlgOpen.ShowDialog = DialogResult.OK Then
                txtInput.Text = dlgOpen.FileName
                g.Open(dlgOpen.FileName)
                Return g
            Else
                Return Nothing
            End If
        Catch ex As Exception
            MsgBox("The file you have choosen is not a valid GRID dataset.  Please select another.", MsgBoxStyle.Critical, "Invalid Data Type")
        End Try

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
        pRaster.Header.Projection = g_MapWin.Project.ProjectProjection
        Return pRaster
    End Function

    Public Function ReturnRasterStretchColorRampCS(ByRef pRaster As MapWinGIS.Grid, ByRef strColor As String) As MapWinGIS.GridColorScheme
        ReturnRasterStretchColorRampCS = Nothing
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



            Dim cs As New MapWinGIS.GridColorScheme
            Dim csbrk As New MapWinGIS.GridColorBreak
            csbrk.LowValue = pRaster.Minimum
            csbrk.HighValue = pRaster.Maximum
            csbrk.ColoringType = MapWinGIS.ColoringType.Gradient
            csbrk.LowColor = Convert.ToUInt32(RGB(rFrom, gFrom, bFrom))
            csbrk.HighColor = Convert.ToUInt32(RGB(rTo, gTo, bTo))
            csbrk.Caption = csbrk.LowValue.ToString() + " - " + csbrk.HighValue.ToString()
            cs.InsertBreak(csbrk)
            Return cs
        Catch ex As Exception
            HandleError(True, "ReturnRasterStretchColorRampRender " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Function

    Public Function ReturnUniqueRasterRenderer(ByRef pRaster As MapWinGIS.Grid, ByRef strStandardName As String) As Object
        Try
            'Create two colors, red, green
            'pColorRed.RGB = System.Convert.ToUInt32(RGB(214, 71, 0))
            'pColorGreen.RGB = System.Convert.ToUInt32(RGB(56, 168, 0))

            Dim cs As New MapWinGIS.GridColorScheme
            Dim csb1 As New MapWinGIS.GridColorBreak
            Dim csb2 As New MapWinGIS.GridColorBreak

            csb1.Caption = "Exceeds Standard"
            csb1.ColoringType = MapWinGIS.ColoringType.Gradient
            csb1.HighValue = 1
            csb1.LowValue = 1
            csb1.HighColor = System.Convert.ToUInt32(RGB(214, 71, 0))
            csb1.LowColor = System.Convert.ToUInt32(RGB(214, 71, 0))

            csb2.Caption = "Below Standard"
            csb2.ColoringType = MapWinGIS.ColoringType.Gradient
            csb2.HighValue = 2
            csb2.LowValue = 2
            csb2.HighColor = System.Convert.ToUInt32(RGB(56, 168, 0))
            csb2.LowColor = System.Convert.ToUInt32(RGB(56, 168, 0))

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
            HandleError(True, "ReturnHSVColorString " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Function

    Public Function CheckSpatialReference(ByRef pRasGeoDataset As MapWinGIS.Grid) As String
        CheckSpatialReference = ""
        Try
            Dim strprj As String = pRasGeoDataset.Header.Projection
            If strprj <> "" Then
                Return strprj
            Else
                If IO.Path.GetFileName(pRasGeoDataset.Filename) = "sta.adf" Then
                    If IO.File.Exists(IO.Path.GetDirectoryName(pRasGeoDataset.Filename) + IO.Path.DirectorySeparatorChar + "prj.adf") Then
                        Dim infile As New IO.StreamReader(IO.Path.GetDirectoryName(pRasGeoDataset.Filename) + IO.Path.DirectorySeparatorChar + "prj.adf")
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
        Catch ex As Exception
            HandleError(True, "CheckSpatialReference " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
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
            'Sub to rid the world of stragling GRIDS, i.e. the ones established for global usse


            'Had an 'elegant' solution using an Iarray to hold global rasters, but that didn't seem to do the
            'job, so we have to manually set each and everyone to nothing

            If Not g_pSCS100Raster Is Nothing Then
                g_pSCS100Raster.Close()
                g_pSCS100Raster = Nothing
            End If

            If Not g_pMetRunoffRaster Is Nothing Then
                g_pMetRunoffRaster.Close()
                g_pMetRunoffRaster = Nothing
            End If

            If Not g_pRunoffRaster Is Nothing Then
                g_pRunoffRaster.Close()
                g_pRunoffRaster = Nothing
            End If

            If Not g_pDEMRaster Is Nothing Then
                g_pDEMRaster.Close()
                g_pDEMRaster = Nothing
            End If

            If Not g_pFlowAccRaster Is Nothing Then
                g_pFlowAccRaster.Close()
                g_pFlowAccRaster = Nothing
            End If

            If Not g_pFlowDirRaster Is Nothing Then
                g_pFlowDirRaster.Close()
                g_pFlowDirRaster = Nothing
            End If

            If Not g_pLSRaster Is Nothing Then
                g_pLSRaster.Close()
                g_pLSRaster = Nothing
            End If

            If Not g_pWaterShedFeatClass Is Nothing Then
                g_pWaterShedFeatClass.Close()
                g_pWaterShedFeatClass = Nothing
            End If

            If Not g_KFactorRaster Is Nothing Then
                g_KFactorRaster.Close()
                g_KFactorRaster = Nothing
            End If

            If Not g_pPrecipRaster Is Nothing Then
                g_pPrecipRaster.Close()
                g_pPrecipRaster = Nothing
            End If

            If Not g_LandCoverRaster Is Nothing Then
                g_LandCoverRaster.Close()
                g_LandCoverRaster = Nothing
            End If

            If Not g_pSelectedPolyClip Is Nothing Then
                g_pSelectedPolyClip = Nothing
            End If
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




    Public Delegate Function RasterMathCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
    Public Delegate Function RasterMathCellCalcNulls(ByVal Input1 As Single, ByVal Input1Null As Single, ByVal Input2 As Single, ByVal Input2Null As Single, ByVal Input3 As Single, ByVal Input3Null As Single, ByVal Input4 As Single, ByVal Input4Null As Single, ByVal Input5 As Single, ByVal Input5Null As Single, ByVal OutNull As Single) As Single
    Public Sub RasterMath(ByRef InputGrid1 As MapWinGIS.Grid, ByRef InputGrid2 As MapWinGIS.Grid, ByRef Inputgrid3 As MapWinGIS.Grid, ByRef Inputgrid4 As MapWinGIS.Grid, ByRef Inputgrid5 As MapWinGIS.Grid, ByRef Outputgrid As MapWinGIS.Grid, ByRef CellCalc As RasterMathCellCalc, Optional ByVal checkNullFirst As Boolean = True, Optional ByRef CellCalcNull As RasterMathCellCalcNulls = Nothing)
        Dim head1, head2, head3, head4, head5, headnew As MapWinGIS.GridHeader
        Dim ncol As Integer
        Dim nrow As Integer
        Dim nodata1, nodata2, nodata3, nodata4, nodata5, nodataout As Single
        Dim rowvals1(), rowvals2(), rowvals3(), rowvals4(), rowvals5(), rowvalsout() As Single

        nodataout = -9999.0

        head1 = InputGrid1.Header
        headnew = New MapWinGIS.GridHeader
        headnew.CopyFrom(head1)
        headnew.NodataValue = nodataout
        Outputgrid = New MapWinGIS.Grid()
        Outputgrid.CreateNew("", headnew, MapWinGIS.GridDataType.FloatDataType, headnew.NodataValue)
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
                    If rowvals1(col) = nodata1 OrElse rowvals2(col) = nodata2 OrElse rowvals3(col) = nodata3 OrElse rowvals4(col) = nodata4 OrElse rowvals5(col) = nodata5 Then
                        rowvalsout(col) = nodataout
                    Else
                        rowvalsout(col) = CellCalc.Invoke(rowvals1(col), rowvals2(col), rowvals3(col), rowvals4(col), rowvals5(col), nodataout)
                    End If
                Else
                    rowvalsout(col) = CellCalcNull.Invoke(rowvals1(col), nodata1, rowvals2(col), nodata2, rowvals3(col), nodata3, rowvals4(col), nodata4, rowvals5(col), nodata5, nodataout)
                End If
            Next

            Outputgrid.PutRow(row, rowvalsout(0))
        Next
    End Sub

    Public Delegate Function RasterMathCellCalcWindow(ByRef InputBox1(,) As Single, ByRef InputBox2(,) As Single, ByRef InputBox3(,) As Single, ByRef InputBox4(,) As Single, ByRef InputBox5(,) As Single, ByVal OutNull As Single) As Single
    Public Delegate Function RasterMathCellCalcWindowNulls(ByRef InputBox1(,) As Single, ByVal Input1Null As Single, ByRef InputBox2(,) As Single, ByVal Input2Null As Single, ByRef InputBox3(,) As Single, ByVal Input3Null As Single, ByRef InputBox4(,) As Single, ByVal Input4Null As Single, ByRef InputBox5(,) As Single, ByVal Input5Null As Single, ByVal OutNull As Single) As Single
    Public Sub RasterMathWindow(ByRef InputGrid1 As MapWinGIS.Grid, ByRef InputGrid2 As MapWinGIS.Grid, ByRef Inputgrid3 As MapWinGIS.Grid, ByRef Inputgrid4 As MapWinGIS.Grid, ByRef Inputgrid5 As MapWinGIS.Grid, ByRef Outputgrid As MapWinGIS.Grid, ByRef CellCalcWindow As RasterMathCellCalcWindow, Optional ByVal checkNullFirst As Boolean = True, Optional ByRef CellCalcWindowNull As RasterMathCellCalcWindowNulls = Nothing)
        Dim head1, head2, head3, head4, head5, headnew As MapWinGIS.GridHeader
        Dim ncol As Integer
        Dim nrow As Integer
        Dim nodata1, nodata2, nodata3, nodata4, nodata5, nodataout As Single
        Dim rowvals1(3)(), rowvals2(3)(), rowvals3(3)(), rowvals4(3)(), rowvals5(3)(), rowvalsout() As Single
        Dim InputBox1(3, 3), InputBox2(3, 3), InputBox3(3, 3), InputBox4(3, 3), InputBox5(3, 3) As Single

        nodataout = -9999.0

        head1 = InputGrid1.Header
        headnew = New MapWinGIS.GridHeader
        headnew.CopyFrom(head1)
        headnew.NodataValue = nodataout
        Outputgrid = New MapWinGIS.Grid()
        Outputgrid.CreateNew("", headnew, MapWinGIS.GridDataType.FloatDataType, headnew.NodataValue)
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
        ReDim rowvals1(0)(ncol)
        ReDim rowvals2(0)(ncol)
        ReDim rowvals3(0)(ncol)
        ReDim rowvals4(0)(ncol)
        ReDim rowvals5(0)(ncol)
        ReDim rowvals1(1)(ncol)
        ReDim rowvals2(1)(ncol)
        ReDim rowvals3(1)(ncol)
        ReDim rowvals4(1)(ncol)
        ReDim rowvals5(1)(ncol)
        ReDim rowvals1(2)(ncol)
        ReDim rowvals2(2)(ncol)
        ReDim rowvals3(2)(ncol)
        ReDim rowvals4(2)(ncol)
        ReDim rowvals5(2)(ncol)
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
