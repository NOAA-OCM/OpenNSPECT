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

Friend Class NewFromExistingWaterShedDelineationForm
    Private _frmWS As WatershedDelineationsForm
    Private _frmPrj As MainForm
    Private _strDEM2BFileName As String
    Private _strNibbleName As String

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
                    Return
                Else
                    If strProj.ToLower.Contains("units=m") Then
                        cboDEMUnits.SelectedIndex = 0
                    Else
                        cboDEMUnits.SelectedIndex = 1
                    End If
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

        ShowProgress("Validating input...", "Adding New Delineation...", 3, 1, Me)
        If Not ValidateDataFormInput() Then
            CloseProgressDialog()
            Return
        End If

        Try
            'ARA 10/29/2010 Using base dem and flow dir instead of expanded grids
            _strDEM2BFileName = txtDEMFile.Text
            _strNibbleName = txtFlowDir.Text

            ShowProgress("Updating Database...", "Adding New Delineation...", 3, 2, Me)

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

            CloseProgressDialog()

            'Confirm
            MsgBox(txtWSDelinName.Text & " successfully added.", MsgBoxStyle.OkOnly, "Record Added")

            If g_boolNewWShed Then
                'frmPrj.Show
                _frmPrj.cboWaterShedDelineations.Items.Clear()
                InitComboBox((_frmPrj.cboWaterShedDelineations), "WSDelineation")
                _frmPrj.cboWaterShedDelineations.SelectedIndex = GetIndexOfEntry((txtWSDelinName.Text), (_frmPrj.cboWaterShedDelineations))
                MyBase.OK_Button_Click(sender, e)
            Else
                MyBase.OK_Button_Click(sender, e)
                _frmWS.Close()
            End If

        Catch ex As Exception
            MsgBox("An error occurred while processing your Watershed Delineation.", MsgBoxStyle.Critical, "Error")
            CloseProgressDialog()
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

            Dim dir As String = String.Format("{0}\wsdelin\{1}", g_nspectPath, txtWSDelinName.Text)
            If Not Directory.Exists(dir) Then
                Directory.CreateDirectory(dir)
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
                                flowpath = Path.GetDirectoryName(txtFlowDir.Text) + "_esri" + OutputGridExt
                            Else
                                flowpath = Path.GetDirectoryName(txtFlowDir.Text) + Path.DirectorySeparatorChar + _
                                           Path.GetFileNameWithoutExtension(txtFlowDir.Text) + "_esri" + _
                                           OutputGridExt
                            End If
                            Dim pESRID8Flow As New Grid
                            Dim tmphead As New GridHeader
                            tmphead.CopyFrom(tmpFlow.Header)
                            pESRID8Flow.CreateNew(flowpath, tmphead, GridDataType.FloatDataType, -1)
                            RasterMath(tmpFlow, Nothing, Nothing, Nothing, Nothing, pESRID8Flow, Nothing, False, _
                                        GetConverterToEsriFromTauDem())
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
                    Return
                End If
            End If
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub


#End Region

End Class