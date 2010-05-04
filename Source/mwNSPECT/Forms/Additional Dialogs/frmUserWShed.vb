Imports System.Data.OleDb
Friend Class frmUserWShed
	Inherits System.Windows.Forms.Form

    Private _frmWS As frmWatershedDelin
    Private _frmPrj As frmProjectSetup
    Private _strDEM2BFileName As String
    Private _strNibbleName As String
    Private _dem_null As Double

#Region "Events"
    Private Sub cmdBrowseDEMFile_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdBrowseDEMFile.Click

        'ReturnGRIDPath txtDEMFile, "Select DEM GRID"

        Dim pDEMRasterDataset As MapWinGIS.Grid

        pDEMRasterDataset = AddInputFromGxBrowserText(txtDEMFile, "Choose DEM GRID", Me, 0)

        'Get the spatial reference
        Dim strProj As String = CheckSpatialReference(pDEMRasterDataset)
        If strProj = "" Then
            MsgBox("The GRID you have choosen has no spatial reference information.  Please define a projection before continuing.", MsgBoxStyle.Exclamation, "No Project Information Detected")
            Exit Sub
        Else
            If strProj.ToLower.Contains("units=m") Then
                cboDEMUnits.SelectedIndex = 0
            Else
                cboDEMUnits.SelectedIndex = 1
            End If

            cboDEMUnits.Refresh()

        End If

    End Sub

    Private Sub cmdBrowseFlowAcc_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdBrowseFlowAcc.Click

        ReturnGRIDPath(txtFlowAcc, "Select Flow Accumulation GRID")

    End Sub

    Private Sub cmdBrowseFlowDir_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdBrowseFlowDir.Click

        ReturnGRIDPath(txtFlowDir, "Select Flow Direction GRID")

    End Sub

    Private Sub cmdBrowseLS_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdBrowseLS.Click

        ReturnGRIDPath(txtLS, "Select Length-Slope GRID")

    End Sub

    Private Sub cmdBrowseWS_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdBrowseWS.Click
        txtWaterSheds.Text = BrowseForFileName("Feature", Me, "Select Watersheds Shapefile")
    End Sub

    Private Sub cmdQuit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdQuit.Click
        Me.Close()
    End Sub

    Private Sub cmdCreate_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCreate.Click

        Dim strCmdInsert As String
        
        If Not ValidateDataFormInput() Then
            Exit Sub
        End If

        modProgDialog.ProgDialog("Validating input...", "Adding New Delineation...", 0, 3, 1, 0)

        Try
            modProgDialog.ProgDialog("Creating 2 Cell Buffer and Nibble GRIDs...", "Adding New Delineation...", 0, 3, 2, 0)

            Return2BDEM((txtDEMFile.Text), (txtFlowDir.Text))

            modProgDialog.ProgDialog("Updating Database...", "Adding New Delineation...", 0, 3, 2, 0)

            strCmdInsert = "INSERT INTO WSDelineation " & "(Name, DEMFileName, DEMGridUnits, FlowDirFileName, FlowAccumFileName," & "FilledDEMFileName, HydroCorrected, StreamFileName, SubWSSize, WSFileName, LSFileName, NibbleFileName, DEM2bFileName) " & " VALUES (" & "'" & CStr(txtWSDelinName.Text) & "', " & "'" & CStr(txtDEMFile.Text) & "', " & "'" & cboDEMUnits.SelectedIndex & "', " & "'" & txtFlowDir.Text & "', " & "'" & txtFlowAcc.Text & "', " & "'" & txtDEMFile.Text & "', " & "'" & "0" & "', " & "'" & "" & "', " & "'" & "0" & "', " & "'" & txtWaterSheds.Text & "', " & "'" & txtLS.Text & "', " & "'" & _strNibbleName & "', " & "'" & _strDEM2BFileName & "')"

            'Execute the statement.
            Dim cmdIns As New OleDbCommand(strCmdInsert, g_DBConn)
            cmdIns.ExecuteNonQuery()
            System.Windows.Forms.Cursor.Current = Windows.Forms.Cursors.Default

            modProgDialog.KillDialog()

            'Confirm
            MsgBox(txtWSDelinName.Text & " successfully added.", MsgBoxStyle.OkOnly, "Record Added")

            If g_boolNewWShed Then
                'frmPrj.Show
                _frmPrj.cboWSDelin.Items.Clear()
                modUtil.InitComboBox((_frmPrj.cboWSDelin), "WSDelineation")
                _frmPrj.cboWSDelin.SelectedIndex = modUtil.GetCboIndex((txtWSDelinName.Text), (_frmPrj.cboWSDelin))
                Me.Close()
            Else
                Me.Close()
                _frmWS.Close()
            End If


        Catch ex As Exception
            MsgBox("An error occurred while processing your Watershed Delineation.", MsgBoxStyle.Critical, "Error")
            modProgDialog.KillDialog()
        End Try
    End Sub


#End Region

#Region "Helper"
    Public Sub Init(ByRef frmWS As frmWatershedDelin, ByRef frmPrj As frmProjectSetup)
        _frmWS = frmWS
        _frmPrj = frmPrj
    End Sub

    Private Function ValidateDataFormInput() As Boolean
        'check name
        If Len(Trim(txtWSDelinName.Text)) = 0 Then
            MsgBox("Please enter a name for your watershed delineation.", MsgBoxStyle.Information, "Name Missing")
            txtWSDelinName.Focus()
            ValidateDataFormInput = False
            Exit Function
        End If


        If Not IO.Directory.Exists(modUtil.g_nspectPath & "\wsdelin\" & txtWSDelinName.Text) Then
            IO.Directory.CreateDirectory(modUtil.g_nspectPath & "\wsdelin\" & txtWSDelinName.Text)
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
            If Not (modUtil.RasterExists((txtDEMFile.Text))) Then
                MsgBox("The DEM selected does not appear to be valid.", MsgBoxStyle.Information, "Invalid Dataset")
                txtDEMFile.Focus()
                ValidateDataFormInput = False
                Exit Function
            End If
        End If

        'check flowacc
        If Len(Trim(txtFlowAcc.Text)) = 0 Then
            MsgBox("Please select a Flow Accumulation Grid for your watershed delineation.", MsgBoxStyle.Information, "Flow Accumulation Grid Missing")
            txtFlowAcc.Focus()
            ValidateDataFormInput = False
            Exit Function
        Else
            If Not (modUtil.RasterExists((txtFlowAcc.Text))) Then
                MsgBox("The Flow Accumulation file selected does not appear to be valid.", MsgBoxStyle.Information, "Invalid Dataset")
                txtFlowAcc.Focus()
                ValidateDataFormInput = False
                Exit Function
            End If
        End If

        'Check flowdir
        If Len(Trim(txtFlowDir.Text)) = 0 Then
            MsgBox("Please select a Flow Direction Grid for your watershed delineation.", MsgBoxStyle.Information, "Flow Direction Grid Missing")
            txtFlowDir.Focus()
            ValidateDataFormInput = False
            Exit Function
        Else
            If Not (modUtil.RasterExists((txtFlowDir.Text))) Then
                MsgBox("The Flow Direction file selected does not appear to be valid.", MsgBoxStyle.Information, "Invalid Dataset")
                txtFlowDir.Focus()
                ValidateDataFormInput = False
                Exit Function
            End If
        End If

        'Check LS
        If Len(Trim(txtLS.Text)) = 0 Then
            MsgBox("Please select a Length-slope Grid for your watershed delineation.", MsgBoxStyle.Information, "Length Slope Grid Missing")
            txtLS.Focus()
            ValidateDataFormInput = False
            Exit Function
        Else
            If Not (modUtil.RasterExists((txtLS.Text))) Then
                MsgBox("The Length-slope file selected does not appear to be valid.", MsgBoxStyle.Information, "Invalid Dataset")
                txtLS.Focus()
                ValidateDataFormInput = False
                Exit Function
            End If
        End If

        'Check watersheds
        If Len(Trim(txtWaterSheds.Text)) = 0 Then
            MsgBox("Please select a watershed shapefile for your watershed delineation.", MsgBoxStyle.Information, "Watershed Shapefile Missing")
            txtWaterSheds.Focus()
            ValidateDataFormInput = False
            Exit Function
        Else
            If Not (modUtil.FeatureExists((txtWaterSheds.Text))) Then
                MsgBox("The watersheds file selected does not appear to be valid.", MsgBoxStyle.Information, "Invalid Dataset")
                txtWaterSheds.Focus()
                ValidateDataFormInput = False
                Exit Function
            End If
        End If

        'if we got through all that, return true.

        ValidateDataFormInput = True
    End Function

    Private Sub ReturnGRIDPath(ByRef txtBox As System.Windows.Forms.TextBox, ByRef strTitle As String)

        Dim pDEMRasterDataset As MapWinGIS.Grid

        pDEMRasterDataset = AddInputFromGxBrowserText(txtBox, strTitle, Me, 0)

        'Get the spatial reference
        If CheckSpatialReference(pDEMRasterDataset) Is Nothing Then
            MsgBox("The GRID you have choosen has no spatial reference information.  Please define a projection before continuing.", MsgBoxStyle.Exclamation, "No Project Information Detected")
            Exit Sub
        End If

    End Sub

    Private Sub Return2BDEM(ByRef strDEMFileName As String, ByRef strFlowDirFileName As String)
        'TODO: Work out how to do focalmin and nibble raster algs
        'Dim strExpression As Object

        'Dim pDEMOneCell As MapWinGIS.Grid = Nothing
        'Dim pDEMTwoCell As MapWinGIS.Grid = Nothing
        'Dim pDEMRaster As MapWinGIS.Grid = Nothing
        'Dim pFlowDir As MapWinGIS.Grid = Nothing
        'Dim pFlowDirBV As MapWinGIS.Grid = Nothing
        'Dim pNibble As MapWinGIS.Grid = Nothing
        'Dim pMask As MapWinGIS.Grid = Nothing
        'Dim intCellSize As Short

        ''pWorkspace = pRasterWorkspaceFactory.OpenFromFile(modUtil.g_nspectPath & "\wsdelin\" & txtWSDelinName.Text, 0)

        'pDEMRaster = modUtil.ReturnRaster(strDEMFileName)
        'pFlowDir = modUtil.ReturnRaster(strFlowDirFileName)
        'intCellSize = pDEMRaster.Header.dX
        '_dem_null = pDEMRaster.Header.NodataValue

        ''STEP 1: ----------------------------------------------------------------------
        ''Buffer the DEM by one cell
        'Dim demonecalc As New RasterMathCellCalc(AddressOf demoneCellCalc)
        'RasterMath(pDEMRaster, Nothing, Nothing, Nothing, Nothing, pDEMOneCell, demonecalc)
        ''strExpression = "Con(isnull([aml_fdem]), focalmin([aml_fdem]), [aml_fdem])"

        ''END STEP 1: ------------------------------------------------------------------

        ''STEP 2: ----------------------------------------------------------------------
        ''Buffer the DEM buffer by one more cell
        'Dim demtwocalc As New RasterMathCellCalc(AddressOf demtwoCellCalc)
        'RasterMath(pDEMOneCell, Nothing, Nothing, Nothing, Nothing, pDEMTwoCell, demtwocalc)
        ''strExpression = "Con(isnull([dem_b]), focalmin([dem_b]), [dem_b])"

        ''STEP 3: ----------------------------------------------------------------------
        'Dim maskcalc As New RasterMathCellCalc(AddressOf maskCellCalc)
        'RasterMath(pDEMTwoCell, Nothing, Nothing, Nothing, Nothing, pMask, maskcalc)
        ''strExpression = "con([mask] >= 0, 1, 0)"


        ''STEP 4: ----------------------------------------------------------------------
        'Dim flowdvcalc As New RasterMathCellCalc(AddressOf flowdvCellCalc)
        'RasterMath(pFlowDir, Nothing, Nothing, Nothing, Nothing, pFlowDirBV, flowdvcalc)
        ''strExpression = "con(isnull([fdr_b]),0,[fdr_b])"

        ''Nibble
        'Dim nibcalc As New RasterMathCellCalc(AddressOf nibCellCalc)
        'RasterMath(pFlowDirBV, pMask, Nothing, Nothing, Nothing, pNibble, nibcalc)
        ''strExpression = "nibble([fdr_bv],[waia_reg], dataonly)"

        ''Get nibble's path for use in the database
        '_strNibbleName = modUtil.MakePerminentGrid(pNibble, (pWorkspace.PathName), "nibble")
    End Sub

#End Region

#Region "Raster Math"
    Private Function demoneCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single) As Single
        'strExpression = "Con(isnull([aml_fdem]), focalmin([aml_fdem]), [aml_fdem])"
        If Input1 = _dem_null Then
            'TODO: Redo to allow for neighbor searches
            'focal min is the minimum value of each cell in the neighborhood
            Return 0
        Else
            Return Input1
        End If
    End Function



#End Region


End Class