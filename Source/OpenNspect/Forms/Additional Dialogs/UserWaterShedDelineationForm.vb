'********************************************************************************************************
'File Name: frmUserWshed.vb
'Description: Form for handling user defined watershed delins
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
Imports System.Windows.Forms
Imports System.IO
Imports MapWinGIS

Friend Class UserWaterShedDelineationForm
    Private _frmWS As WatershedDelineationsForm
    Private _frmPrj As MainForm
    Private _strDEM2BFileName As String
    Private _strNibbleName As String
    Private _dem_null As Double

#Region "Events"

    Private Sub cmdBrowseDEMFile_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) _
        Handles cmdBrowseDEMFile.Click
        Try
            'ReturnGRIDPath txtDEMFile, "Select DEM GRID"

            Dim pDEMRasterDataset As Grid

            pDEMRasterDataset = AddInputFromGxBrowserText(txtDEMFile, "Choose DEM GRID")
            If Not pDEMRasterDataset Is Nothing Then
                'Get the spatial reference
                Dim strProj As String = CheckSpatialReference(pDEMRasterDataset)
                If strProj = "" Then
                    MsgBox( _
                            "The GRID you have choosen has no spatial reference information.  Please define a projection before continuing.", _
                            MsgBoxStyle.Exclamation, "No Project Information Detected")
                    Exit Sub
                Else
                    If strProj.ToLower.Contains("units=m") Then
                        cboDEMUnits.SelectedIndex = 0
                    Else
                        cboDEMUnits.SelectedIndex = 1
                    End If

                    cboDEMUnits.Refresh()

                End If

            End If
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub cmdBrowseFlowAcc_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) _
        Handles cmdBrowseFlowAcc.Click
        Try
            ReturnGRIDPath(txtFlowAcc, "Select Flow Accumulation GRID")

        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub cmdBrowseFlowDir_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) _
        Handles cmdBrowseFlowDir.Click
        Try
            ReturnGRIDPath(txtFlowDir, "Select Flow Direction GRID")

        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub cmdBrowseLS_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) _
        Handles cmdBrowseLS.Click
        Try
            ReturnGRIDPath(txtLS, "Select Length-Slope GRID")

        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub cmdBrowseWS_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) _
        Handles cmdBrowseWS.Click
        Try
            txtWaterSheds.Text = BrowseForFileName("Feature")
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Protected Overrides Sub OK_Button_Click(sender As Object, e As EventArgs)
        Dim strCmdInsert As String

        ShowProgress("Validating input...", "Adding New Delineation...", 0, 3, 1, Me)
        If Not ValidateDataFormInput() Then
            CloseDialog()
            Exit Sub
        End If

        Try
            'ARA 10/29/2010 Using base dem and flow dir instead of expanded grids
            'modProgDialog.ProgDialog("Creating 2 Cell Buffer and Nibble GRIDs...", "Adding New Delineation...", 0, 3, 2, 0)
            'Return2BDEM(txtDEMFile.Text, txtFlowDir.Text)
            _strDEM2BFileName = txtDEMFile.Text
            _strNibbleName = txtFlowDir.Text

            ShowProgress("Updating Database...", "Adding New Delineation...", 0, 3, 2, Me)

            strCmdInsert = String.Format("INSERT INTO WSDelineation (Name, DEMFileName, DEMGridUnits, FlowDirFileName, FlowAccumFileName,FilledDEMFileName, HydroCorrected, StreamFileName, SubWSSize, WSFileName, LSFileName, NibbleFileName, DEM2bFileName)  VALUES ('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '0', '', '0', '{6}', '{7}', '{8}', '{9}')", _
                               CStr(txtWSDelinName.Text), _
                               CStr(txtDEMFile.Text), _
                               cboDEMUnits.SelectedIndex, _
                               txtFlowDir.Text, _
                               txtFlowAcc.Text, _
                               txtDEMFile.Text, _
                               txtWaterSheds.Text, _
                               txtLS.Text, _
                               _strNibbleName, _
                               _strDEM2BFileName)

            'Execute the statement.
            Using cmdIns As New DataHelper(strCmdInsert)
                cmdIns.ExecuteNonQuery()
            End Using
            System.Windows.Forms.Cursor.Current = Cursors.Default

            CloseDialog()

            'Confirm
            MsgBox(txtWSDelinName.Text & " successfully added.", MsgBoxStyle.OkOnly, "Record Added")

            If g_boolNewWShed Then
                'frmPrj.Show
                _frmPrj.cboWSDelin.Items.Clear()
                InitComboBox((_frmPrj.cboWSDelin), "WSDelineation")
                _frmPrj.cboWSDelin.SelectedIndex = GetCboIndex((txtWSDelinName.Text), (_frmPrj.cboWSDelin))
                MyBase.OK_Button_Click(sender, e)
            Else
                MyBase.OK_Button_Click(sender, e)
                _frmWS.Close()
            End If

        Catch ex As Exception
            MsgBox("An error occurred while processing your Watershed Delineation.", MsgBoxStyle.Critical, "Error")
            CloseDialog()
        End Try

    End Sub

#End Region

#Region "Helper"

    Public Sub Init(ByRef frmWS As WatershedDelineationsForm, ByRef frmPrj As MainForm)
        Try
            _frmWS = frmWS
            _frmPrj = frmPrj
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Function ValidateDataFormInput() As Boolean
        Try
            'check name
            If Len(Trim(txtWSDelinName.Text)) = 0 Then
                MsgBox("Please enter a name for your watershed delineation.", MsgBoxStyle.Information, "Name Missing")
                txtWSDelinName.Focus()
                ValidateDataFormInput = False
                Exit Function
            End If

            If Not Directory.Exists(g_nspectPath & "\wsdelin\" & txtWSDelinName.Text) Then
                Directory.CreateDirectory(g_nspectPath & "\wsdelin\" & txtWSDelinName.Text)
            Else
                MsgBox("Name in use.  Please select another.", MsgBoxStyle.Critical, "Choose New Name")
                txtWSDelinName.Focus()
                ValidateDataFormInput = False
                Exit Function
            End If

            'check dem
            If Len(Trim(txtDEMFile.Text)) = 0 Then
                MsgBox("Please select a DEM for your watershed delineation.", MsgBoxStyle.Information, "DEM Missing")
                txtDEMFile.Focus()
                ValidateDataFormInput = False
                Exit Function
            Else
                If Not (RasterExists((txtDEMFile.Text))) Then
                    MsgBox("The DEM selected does not appear to be valid.", MsgBoxStyle.Information, "Invalid Dataset")
                    txtDEMFile.Focus()
                    ValidateDataFormInput = False
                    Exit Function
                End If
            End If

            'check flowacc
            If Len(Trim(txtFlowAcc.Text)) = 0 Then
                MsgBox("Please select a Flow Accumulation Grid for your watershed delineation.", _
                        MsgBoxStyle.Information, "Flow Accumulation Grid Missing")
                txtFlowAcc.Focus()
                ValidateDataFormInput = False
                Exit Function
            Else
                If Not (RasterExists((txtFlowAcc.Text))) Then
                    MsgBox("The Flow Accumulation file selected does not appear to be valid.", MsgBoxStyle.Information, _
                            "Invalid Dataset")
                    txtFlowAcc.Focus()
                    ValidateDataFormInput = False
                    Exit Function
                End If
            End If

            'Check flowdir
            If Len(Trim(txtFlowDir.Text)) = 0 Then
                MsgBox("Please select a Flow Direction Grid for your watershed delineation.", MsgBoxStyle.Information, _
                        "Flow Direction Grid Missing")
                txtFlowDir.Focus()
                ValidateDataFormInput = False
                Exit Function
            Else
                If Not (RasterExists(txtFlowDir.Text)) Then
                    MsgBox("The Flow Direction file selected does not appear to be valid.", MsgBoxStyle.Information, _
                            "Invalid Dataset")
                    txtFlowDir.Focus()
                    ValidateDataFormInput = False
                    Exit Function
                Else
                    Dim tmpFlow As New Grid
                    tmpFlow.Open(txtFlowDir.Text)
                    'Really crude check for if an ESRI flow direction grid or TAUDEM. Fails if ESRI grid only has 1, 2, 4, and 8 flows, so leave option for using it. Most with only 1-8 will be non-esri flows though.
                    If tmpFlow.Maximum > 8 Then
                        tmpFlow.Close()
                    Else
                        Dim _
                            tmpres As DialogResult = _
                                MsgBox( _
                                        "The Flow Direction file selected does not seem to be an ESRI format flow direction grid. If this file is in TAUDEM format, click Yes to convert it to ESRI format. If you believe this file is ESRI format already, click No to override this error. Otherwise, click cancel to abort the process.", _
                                        MsgBoxStyle.YesNoCancel, "Flow Direction File Error")
                        If tmpres = System.Windows.Forms.DialogResult.Yes Then
                            Dim flowpath As String
                            If Path.GetFileName(txtFlowDir.Text) = "sta.adf" Then
                                flowpath = Path.GetDirectoryName(txtFlowDir.Text) + "_esri" + g_OutputGridExt
                            Else
                                flowpath = Path.GetDirectoryName(txtFlowDir.Text) + Path.DirectorySeparatorChar + _
                                           Path.GetFileNameWithoutExtension(txtFlowDir.Text) + "_esri" + _
                                           g_OutputGridExt
                            End If
                            Dim pESRID8Flow As New Grid
                            Dim tmphead As New GridHeader
                            tmphead.CopyFrom(tmpFlow.Header)
                            pESRID8Flow.CreateNew(flowpath, tmphead, GridDataType.FloatDataType, -1)
                            Dim tauD8ToESRIcalc As New RasterMathCellCalcNulls(AddressOf tauD8ToESRICellCalc)
                            RasterMath(tmpFlow, Nothing, Nothing, Nothing, Nothing, pESRID8Flow, Nothing, False, _
                                        tauD8ToESRIcalc)
                            pESRID8Flow.Header.NodataValue = -1
                            pESRID8Flow.Save(flowpath)
                            pESRID8Flow.Close()
                            txtFlowDir.Text = flowpath
                            tmpFlow.Close()
                            If Not (RasterExists(txtFlowDir.Text)) Then
                                MsgBox( _
                                        "The created Flow Direction file does not appear to be valid. Aborting the process.", _
                                        MsgBoxStyle.Information, "Invalid Dataset")
                                txtFlowDir.Focus()
                                ValidateDataFormInput = False
                                Exit Function
                            End If
                        ElseIf tmpres = System.Windows.Forms.DialogResult.No Then
                            tmpFlow.Close()
                        Else
                            tmpFlow.Close()
                            txtFlowDir.Focus()
                            ValidateDataFormInput = False
                            Exit Function
                        End If
                    End If
                End If
            End If

            'Check LS
            If Len(Trim(txtLS.Text)) = 0 Then
                MsgBox("Please select a Length-slope Grid for your watershed delineation.", MsgBoxStyle.Information, _
                        "Length Slope Grid Missing")
                txtLS.Focus()
                ValidateDataFormInput = False
                Exit Function
            Else
                If Not (RasterExists((txtLS.Text))) Then
                    MsgBox("The Length-slope file selected does not appear to be valid.", MsgBoxStyle.Information, _
                            "Invalid Dataset")
                    txtLS.Focus()
                    ValidateDataFormInput = False
                    Exit Function
                End If
            End If

            'Check watersheds
            If Len(Trim(txtWaterSheds.Text)) = 0 Then
                MsgBox("Please select a watershed shapefile for your watershed delineation.", MsgBoxStyle.Information, _
                        "Watershed Shapefile Missing")
                txtWaterSheds.Focus()
                ValidateDataFormInput = False
                Exit Function
            Else
                If Not (FeatureExists((txtWaterSheds.Text))) Then
                    MsgBox("The watersheds file selected does not appear to be valid.", MsgBoxStyle.Information, _
                            "Invalid Dataset")
                    txtWaterSheds.Focus()
                    ValidateDataFormInput = False
                    Exit Function
                End If
            End If

            'if we got through all that, return true.

            ValidateDataFormInput = True
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Function

    Private Sub ReturnGRIDPath(ByRef txtBox As TextBox, ByRef strTitle As String)
        Try
            Dim pDEMRasterDataset As Grid

            pDEMRasterDataset = AddInputFromGxBrowserText(txtBox, strTitle)
            If Not pDEMRasterDataset Is Nothing Then
                'Get the spatial reference
                If CheckSpatialReference(pDEMRasterDataset) Is Nothing Then
                    MsgBox( _
                            "The GRID you have choosen has no spatial reference information.  Please define a projection before continuing.", _
                            MsgBoxStyle.Exclamation, "No Project Information Detected")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub Return2BDEM(ByRef strDEMFileName As String, ByRef strFlowDirFileName As String)
        Try
            Dim pDEMOneCell As Grid = Nothing
            Dim pDEMTwoCell As Grid = Nothing
            Dim pDEMRaster As Grid = Nothing
            Dim pFlowDir As Grid = Nothing
            Dim pFlowDirBV As Grid = Nothing
            Dim pNibble As Grid = Nothing
            Dim pMask As Grid = Nothing
            Dim intCellSize As Short

            'pWorkspace = pRasterWorkspaceFactory.OpenFromFile(modUtil.g_nspectPath & "\wsdelin\" & txtWSDelinName.Text, 0)

            pDEMRaster = ReturnRaster(strDEMFileName)
            pFlowDir = ReturnRaster(strFlowDirFileName)
            intCellSize = pDEMRaster.Header.dX
            _dem_null = pDEMRaster.Header.NodataValue

            'STEP 1: ----------------------------------------------------------------------
            'Buffer the DEM by one cell
            Dim demonecalc As New RasterMathCellCalcWindowNulls(AddressOf focalminGrowCellCalc)
            RasterMathWindow(pDEMRaster, Nothing, Nothing, Nothing, Nothing, pDEMOneCell, Nothing, False, demonecalc)
            'strExpression = "Con(isnull([aml_fdem]), focalmin([aml_fdem]), [aml_fdem])"

            'END STEP 1: ------------------------------------------------------------------

            'STEP 2: ----------------------------------------------------------------------
            'Buffer the DEM buffer by one more cell
            Dim demtwocalc As New RasterMathCellCalcWindowNulls(AddressOf focalminGrowCellCalc)
            RasterMathWindow(pDEMOneCell, Nothing, Nothing, Nothing, Nothing, pDEMTwoCell, Nothing, False, demtwocalc)
            'strExpression = "Con(isnull([dem_b]), focalmin([dem_b]), [dem_b])"

            Dim fdronecalc As New RasterMathCellCalcWindowNulls(AddressOf focalminGrowCellCalc)
            RasterMathWindow(pFlowDir, Nothing, Nothing, Nothing, Nothing, pFlowDirBV, Nothing, False, demonecalc)

            Dim fdrtwocalc As New RasterMathCellCalcWindowNulls(AddressOf focalminGrowCellCalc)
            RasterMathWindow(pFlowDirBV, Nothing, Nothing, Nothing, Nothing, pNibble, Nothing, False, demtwocalc)

            ''STEP 3: ----------------------------------------------------------------------
            'Dim maskcalc As New RasterMathCellCalcNulls(AddressOf maskCellCalc)
            'RasterMath(pDEMTwoCell, Nothing, Nothing, Nothing, Nothing, pMask, Nothing, False, maskcalc)
            ''strExpression = "con([mask] >= 0, 1, 0)"

            ''STEP 4: ----------------------------------------------------------------------
            'Dim flowdvcalc As New RasterMathCellCalcNulls(AddressOf flowdvCellCalc)
            'RasterMath(pFlowDir, Nothing, Nothing, Nothing, Nothing, pFlowDirBV, Nothing, False, flowdvcalc)
            ''strExpression = "con(isnull([fdr_b]),0,[fdr_b])"

            ''Nibble
            'Dim nibcalc As New RasterMathCellCalcWindowNulls(AddressOf nibCellCalc)
            'RasterMathWindow(pFlowDirBV, pMask, Nothing, Nothing, Nothing, pNibble, Nothing, False, nibcalc)
            ''strExpression = "nibble([fdr_bv],[waia_reg], dataonly)"

            'Get nibble's path for use in the database
            _strNibbleName = GetUniqueName("nibble", g_strWorkspace, g_OutputGridExt)
            ReturnPermanentRaster(pNibble, _strNibbleName)
            pNibble.Close()
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

#End Region

#Region "Raster Math"

    Private Function focalminGrowCellCalc(ByRef InputBox1(,) As Single, ByVal Input1Null As Single, _
                                           ByRef InputBox2(,) As Single, ByVal Input2Null As Single, _
                                           ByRef InputBox3(,) As Single, ByVal Input3Null As Single, _
                                           ByRef InputBox4(,) As Single, ByVal Input4Null As Single, _
                                           ByRef InputBox5(,) As Single, ByVal Input5Null As Single, _
                                           ByVal OutNull As Single) As Single
        Try
            'strExpression = "Con(isnull([aml_fdem]), focalmin([aml_fdem]), [aml_fdem])"
            'focal min is the minimum non-nodata value of each cell in the neighborhood
            Dim minval As Single = Single.MaxValue
            If InputBox1(1, 1) = Input1Null Then
                For i As Integer = 0 To 2
                    For j As Integer = 0 To 2
                        If InputBox1(i, j) <> Input1Null Then
                            If InputBox1(i, j) < minval Then
                                minval = InputBox1(i, j)
                            End If
                        End If
                    Next
                Next

                If minval = Single.MaxValue Then
                    Return Input1Null
                Else
                    Return minval
                End If
            Else
                Return InputBox1(1, 1)
            End If
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Function

    Private Function maskCellCalc(ByVal Input1 As Single, ByVal Input1Null As Single, ByVal Input2 As Single, _
                                   ByVal Input2Null As Single, ByVal Input3 As Single, ByVal Input3Null As Single, _
                                   ByVal Input4 As Single, ByVal Input4Null As Single, ByVal Input5 As Single, _
                                   ByVal Input5Null As Single, ByVal OutNull As Single) As Single
        Try
            'strExpression = "con([mask] >= 0, 1, 0)"
            If Input1 <> Input1Null Then
                If Input1 >= 0 Then
                    Return 1
                Else
                    Return 0
                End If
            Else
                Return OutNull
            End If
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Function

    Private Function flowdvCellCalc(ByVal Input1 As Single, ByVal Input1Null As Single, ByVal Input2 As Single, _
                                     ByVal Input2Null As Single, ByVal Input3 As Single, ByVal Input3Null As Single, _
                                     ByVal Input4 As Single, ByVal Input4Null As Single, ByVal Input5 As Single, _
                                     ByVal Input5Null As Single, ByVal OutNull As Single) As Single
        Try
            If Input1 = Input1Null Then
                Return 0
            Else
                Return Input1
            End If
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Function

#End Region
End Class