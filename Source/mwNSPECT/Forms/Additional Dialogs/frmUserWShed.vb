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
        
        modProgDialog.ProgDialog("Validating input...", "Adding New Delineation...", 0, 3, 1, Me)
        If Not ValidateDataFormInput() Then
            Exit Sub
        End If


        Try    
            'ARA 10/29/2010 Using base dem and flow dir instead of expanded grids
            'modProgDialog.ProgDialog("Creating 2 Cell Buffer and Nibble GRIDs...", "Adding New Delineation...", 0, 3, 2, 0)
            'Return2BDEM(txtDEMFile.Text, txtFlowDir.Text)
            _strDEM2BFileName = txtDEMFile.Text
            _strNibbleName = txtFlowDir.Text

            modProgDialog.ProgDialog("Updating Database...", "Adding New Delineation...", 0, 3, 2, Me)

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
        Dim pDEMOneCell As MapWinGIS.Grid = Nothing
        Dim pDEMTwoCell As MapWinGIS.Grid = Nothing
        Dim pDEMRaster As MapWinGIS.Grid = Nothing
        Dim pFlowDir As MapWinGIS.Grid = Nothing
        Dim pFlowDirBV As MapWinGIS.Grid = Nothing
        Dim pNibble As MapWinGIS.Grid = Nothing
        Dim pMask As MapWinGIS.Grid = Nothing
        Dim intCellSize As Short

        'pWorkspace = pRasterWorkspaceFactory.OpenFromFile(modUtil.g_nspectPath & "\wsdelin\" & txtWSDelinName.Text, 0)

        pDEMRaster = modUtil.ReturnRaster(strDEMFileName)
        pFlowDir = modUtil.ReturnRaster(strFlowDirFileName)
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
        _strNibbleName = modUtil.GetUniqueName("nibble", g_strWorkspace, ".bgd")
        modUtil.ReturnPermanentRaster(pNibble, _strNibbleName)
        pNibble.Close()
    End Sub

#End Region

#Region "Raster Math"
    Private Function focalminGrowCellCalc(ByRef InputBox1(,) As Single, ByVal Input1Null As Single, ByRef InputBox2(,) As Single, ByVal Input2Null As Single, ByRef InputBox3(,) As Single, ByVal Input3Null As Single, ByRef InputBox4(,) As Single, ByVal Input4Null As Single, ByRef InputBox5(,) As Single, ByVal Input5Null As Single, ByVal OutNull As Single) As Single
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
            Return 0
        Else
            Return InputBox1(1, 1)
        End If
    End Function

    Private Function maskCellCalc(ByVal Input1 As Single, ByVal Input1Null As Single, ByVal Input2 As Single, ByVal Input2Null As Single, ByVal Input3 As Single, ByVal Input3Null As Single, ByVal Input4 As Single, ByVal Input4Null As Single, ByVal Input5 As Single, ByVal Input5Null As Single, ByVal OutNull As Single) As Single
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
    End Function

    Private Function flowdvCellCalc(ByVal Input1 As Single, ByVal Input1Null As Single, ByVal Input2 As Single, ByVal Input2Null As Single, ByVal Input3 As Single, ByVal Input3Null As Single, ByVal Input4 As Single, ByVal Input4Null As Single, ByVal Input5 As Single, ByVal Input5Null As Single, ByVal OutNull As Single) As Single
        If Input1 = Input1Null Then
            Return 0
        Else
            Return Input1
        End If
    End Function

    Private Function nibCellCalc(ByRef InputBox1(,) As Single, ByVal Input1Null As Single, ByRef InputBox2(,) As Single, ByVal Input2Null As Single, ByRef InputBox3(,) As Single, ByVal Input3Null As Single, ByRef InputBox4(,) As Single, ByVal Input4Null As Single, ByRef InputBox5(,) As Single, ByVal Input5Null As Single, ByVal OutNull As Single) As Single
        'TODO: forming a nibble flow dir
    End Function
#End Region


End Class