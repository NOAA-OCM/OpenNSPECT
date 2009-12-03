Option Strict Off
Option Explicit On
Friend Class frmUserWShed
	Inherits System.Windows.Forms.Form
	Private m_App As ESRI.ArcGIS.Framework.IApplication
	Private m_strDEM2BFileName As String
	Private m_strNibbleName As String
	
	Private Sub cmdBrowseDEMFile_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdBrowseDEMFile.Click
		
		'ReturnGRIDPath txtDEMFile, "Select DEM GRID"
		
		Dim pDEMRasterDataset As ESRI.ArcGIS.Geodatabase.IRasterDataset
		Dim pDistUnit As ESRI.ArcGIS.Geometry.ILinearUnit
		Dim intUnit As Short
		Dim pProjCoord As ESRI.ArcGIS.Geometry.IProjectedCoordinateSystem
		Dim strInputDEM As String
		
		On Error GoTo ErrHandler
		
		pDEMRasterDataset = AddInputFromGxBrowserText(txtDEMFile, "Choose DEM GRID", Me, 0)
		
		'Get the spatial reference
		If CheckSpatialReference(pDEMRasterDataset) Is Nothing Then
			
			MsgBox("The GRID you have choosen has no spatial reference information.  Please define a projection before continuing.", MsgBoxStyle.Exclamation, "No Project Information Detected")
			Exit Sub
			
		Else
			
			pProjCoord = CheckSpatialReference(pDEMRasterDataset)
			pDistUnit = pProjCoord.CoordinateUnit
			intUnit = pDistUnit.MetersPerUnit
			
			If intUnit = 1 Then
				cboDEMUnits.SelectedIndex = 0
			Else
				cboDEMUnits.SelectedIndex = 1
			End If
			
			cboDEMUnits.Refresh()
			
		End If
		
		'UPGRADE_NOTE: Object pDEMRasterDataset may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pDEMRasterDataset = Nothing
		'UPGRADE_NOTE: Object pDistUnit may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pDistUnit = Nothing
		'UPGRADE_NOTE: Object pProjCoord may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pProjCoord = Nothing
		Exit Sub
		
ErrHandler: 
		Exit Sub
		
		
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
	
	
	Private Function ValidateDataFormInput() As Boolean
		
		Dim fso As New Scripting.FileSystemObject
		Dim Folder As Scripting.Folder
		
		'check name
		If Len(Trim(txtWSDelinName.Text)) = 0 Then
			MsgBox("Please enter a name for your watershed delineation.", MsgBoxStyle.Information, "Name Missing")
			txtWSDelinName.Focus()
			ValidateDataFormInput = False
			Exit Function
		End If
		
		
		If Not fso.FolderExists(modUtil.g_nspectPath & "\wsdelin\" & txtWSDelinName.Text) Then
			Folder = fso.CreateFolder(modUtil.g_nspectPath & "\wsdelin\" & txtWSDelinName.Text)
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
	Private Sub cmdCreate_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCreate.Click
		
		Dim strCmdInsert As String
		Dim strDEM2BFileName As String
		Dim strNibbleFileName As String
		
		If Not ValidateDataFormInput Then
			Exit Sub
		End If
		
		modProgDialog.ProgDialog("Validating input...", "Adding New Delineation...", 0, 3, 1, (m_App.hwnd))
		
		On Error GoTo ErrHandler
		
		modProgDialog.ProgDialog("Creating 2 Cell Buffer and Nibble GRIDs...", "Adding New Delineation...", 0, 3, 2, (m_App.hwnd))
		
		Return2BDEM((txtDEMFile.Text), (txtFlowDir.Text))
		
		modProgDialog.ProgDialog("Updating Database...", "Adding New Delineation...", 0, 3, 2, (m_App.hwnd))
		
		strCmdInsert = "INSERT INTO WSDelineation " & "(Name, DEMFileName, DEMGridUnits, FlowDirFileName, FlowAccumFileName," & "FilledDEMFileName, HydroCorrected, StreamFileName, SubWSSize, WSFileName, LSFileName, NibbleFileName, DEM2bFileName) " & " VALUES (" & "'" & CStr(txtWSDelinName.Text) & "', " & "'" & CStr(txtDEMFile.Text) & "', " & "'" & cboDEMUnits.SelectedIndex & "', " & "'" & txtFlowDir.Text & "', " & "'" & txtFlowAcc.Text & "', " & "'" & txtDEMFile.Text & "', " & "'" & "0" & "', " & "'" & "" & "', " & "'" & "0" & "', " & "'" & txtWaterSheds.Text & "', " & "'" & txtLS.Text & "', " & "'" & m_strNibbleName & "', " & "'" & m_strDEM2BFileName & "')"
		
		'Execute the statement.
		modUtil.g_ADOConn.Execute(strCmdInsert, ADODB.CommandTypeEnum.adCmdText)
		'UPGRADE_ISSUE: Unable to determine which constant to upgrade vbNormal to. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B3B44E51-B5F1-4FD7-AA29-CAD31B71F487"'
		'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = vbNormal
		
		modProgDialog.KillDialog()
		
		'Confirm
		MsgBox(txtWSDelinName.Text & " successfully added.", MsgBoxStyle.OKOnly, "Record Added")
		
		If g_boolNewWShed Then
			'frmPrj.Show
			frmPrj.Frame.Visible = True
			frmPrj.cboWSDelin.Items.Clear()
			modUtil.InitComboBox((frmPrj.cboWSDelin), "WSDelineation")
			frmPrj.cboWSDelin.SelectedIndex = modUtil.GetCboIndex((txtWSDelinName.Text), (frmPrj.cboWSDelin))
			Me.Close()
			frmNewWSDelin.Close()
		Else
			Me.Close()
			frmNewWSDelin.Close()
			frmWSDelin.Close()
		End If
		
		
		Exit Sub
		
ErrHandler: 
		MsgBox("An error occurred while processing your Watershed Delineation.", MsgBoxStyle.Critical, "Error")
		modProgDialog.KillDialog()
		
	End Sub
	
	Private Sub ReturnGRIDPath(ByRef txtBox As System.Windows.Forms.TextBox, ByRef strTitle As String)
		
		Dim pDEMRasterDataset As ESRI.ArcGIS.Geodatabase.IRasterDataset
		Dim pProjCoord As ESRI.ArcGIS.Geometry.IProjectedCoordinateSystem
		
		On Error GoTo ErrHandler
		
		pDEMRasterDataset = AddInputFromGxBrowserText(txtBox, strTitle, Me, 0)
		
		
		'Get the spatial reference
		If CheckSpatialReference(pDEMRasterDataset) Is Nothing Then
			
			MsgBox("The GRID you have choosen has no spatial reference information.  Please define a projection before continuing.", MsgBoxStyle.Exclamation, "No Project Information Detected")
			Exit Sub
			
		End If
		
		'Get the name
		'UPGRADE_NOTE: Object pDEMRasterDataset may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pDEMRasterDataset = Nothing
		
ErrHandler: 
		Exit Sub
	End Sub
	
	Private Sub Return2BDEM(ByRef strDEMFileName As String, ByRef strFlowDirFileName As String)
		Dim strExpression As Object
		
		Dim pMapAlgebraOp As ESRI.ArcGIS.SpatialAnalyst.IMapAlgebraOp
		Dim pEnv As ESRI.ArcGIS.GeoAnalyst.IRasterAnalysisEnvironment
		Dim pDEMOneCell As ESRI.ArcGIS.Geodatabase.IRaster
		Dim pDEMTwoCell As ESRI.ArcGIS.Geodatabase.IRaster
		Dim pDEMRaster As ESRI.ArcGIS.Geodatabase.IRaster
		Dim pDEMRasterProps As ESRI.ArcGIS.DataSourcesRaster.IRasterProps
		Dim pFlowDir As ESRI.ArcGIS.Geodatabase.IRaster
		Dim pFlowDirBV As ESRI.ArcGIS.Geodatabase.IRaster
		Dim pNibble As ESRI.ArcGIS.Geodatabase.IRaster
		Dim pMask As ESRI.ArcGIS.Geodatabase.IRaster
		Dim pRasterWorkspaceFactory As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory
		Dim pWorkspace As ESRI.ArcGIS.Geodatabase.IWorkspace
		Dim intCellSize As Short
		Dim pEnvelope As ESRI.ArcGIS.Geometry.IEnvelope
		
		pRasterWorkspaceFactory = New ESRI.ArcGIS.DataSourcesRaster.RasterWorkspaceFactory
		pWorkspace = pRasterWorkspaceFactory.OpenFromFile(modUtil.g_nspectPath & "\wsdelin\" & txtWSDelinName.Text, 0)
		pDEMRaster = modUtil.ReturnRaster(strDEMFileName)
		pFlowDir = modUtil.ReturnRaster(strFlowDirFileName)
		pDEMRasterProps = pDEMRaster
		
		intCellSize = pDEMRasterProps.MeanCellSize.X
		
		pEnvelope = pDEMRasterProps.Extent
		pEnvelope.Expand(intCellSize * 2, intCellSize * 2, False)
		
		pMapAlgebraOp = New ESRI.ArcGIS.SpatialAnalyst.RasterMapAlgebraOp
		pEnv = pMapAlgebraOp
		
		With pEnv
			.SetCellSize(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvValue, pDEMRasterProps.MeanCellSize.X)
			'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.OutWorkspace = pWorkspace
			'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutSpatialReference. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.OutSpatialReference = pDEMRasterProps.SpatialReference '.SpatialReference
			.SetExtent(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvValue, pEnvelope) 'pDEMRasterProps.Extent
		End With
		
		'STEP 1: ----------------------------------------------------------------------
		'Buffer the DEM by one cell
		'UPGRADE_WARNING: Couldn't resolve default property of object pDEMRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pMapAlgebraOp.BindRaster(pDEMRaster, "aml_fdem")
		'UPGRADE_WARNING: Couldn't resolve default property of object strExpression. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		strExpression = "Con(isnull([aml_fdem]), focalmin([aml_fdem]), [aml_fdem])"
		
		'UPGRADE_WARNING: Couldn't resolve default property of object strExpression. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pDEMOneCell = pMapAlgebraOp.Execute(strExpression)
		pMapAlgebraOp.UnbindRaster("aml_fdem")
		'END STEP 1: ------------------------------------------------------------------
		
		'STEP 2: ----------------------------------------------------------------------
		'Buffer the DEM buffer by one more cell
		'UPGRADE_WARNING: Couldn't resolve default property of object pDEMOneCell. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pMapAlgebraOp.BindRaster(pDEMOneCell, "dem_b")
		'UPGRADE_WARNING: Couldn't resolve default property of object strExpression. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		strExpression = "Con(isnull([dem_b]), focalmin([dem_b]), [dem_b])"
		
		'UPGRADE_WARNING: Couldn't resolve default property of object strExpression. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pDEMTwoCell = pMapAlgebraOp.Execute(strExpression)
		m_strDEM2BFileName = modUtil.MakePerminentGrid(pDEMTwoCell, (pWorkspace.PathName), "dem2b")
		pMapAlgebraOp.UnbindRaster("dem_b")
		
		'STEP 3: ----------------------------------------------------------------------
		'UPGRADE_WARNING: Couldn't resolve default property of object pDEMOneCell. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pMapAlgebraOp.BindRaster(pDEMOneCell, "mask")
		'UPGRADE_WARNING: Couldn't resolve default property of object strExpression. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		strExpression = "con([mask] >= 0, 1, 0)"
		'UPGRADE_WARNING: Couldn't resolve default property of object strExpression. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pMask = pMapAlgebraOp.Execute(strExpression)
		
		With pEnv
			.Mask = pMask
			'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.OutWorkspace = pWorkspace
		End With
		
		pEnv = pMapAlgebraOp
		
		'STEP 4: ----------------------------------------------------------------------
		With pMapAlgebraOp
			'UPGRADE_WARNING: Couldn't resolve default property of object pFlowDir. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.BindRaster(pFlowDir, "fdr_b")
		End With
		
		'UPGRADE_WARNING: Couldn't resolve default property of object strExpression. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		strExpression = "con(isnull([fdr_b]),0,[fdr_b])"
		
		'UPGRADE_WARNING: Couldn't resolve default property of object strExpression. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pFlowDirBV = pMapAlgebraOp.Execute(strExpression)
		
		pMapAlgebraOp.UnbindRaster("fdr_b")
		
		'Nibble
		'UPGRADE_WARNING: Couldn't resolve default property of object pFlowDirBV. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pMapAlgebraOp.BindRaster(pFlowDirBV, "fdr_bv")
		'UPGRADE_WARNING: Couldn't resolve default property of object pMask. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pMapAlgebraOp.BindRaster(pMask, "waia_reg")
		'UPGRADE_WARNING: Couldn't resolve default property of object strExpression. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		strExpression = "nibble([fdr_bv],[waia_reg], dataonly)"
		
		'UPGRADE_WARNING: Couldn't resolve default property of object strExpression. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pNibble = pMapAlgebraOp.Execute(strExpression)
		
		'Get nibble's path for use in the database
		m_strNibbleName = modUtil.MakePerminentGrid(pNibble, (pWorkspace.PathName), "nibble")
		
		With pMapAlgebraOp
			.UnbindRaster("fdr_bv")
			.UnbindRaster("waia_reg")
		End With
		
		'Cleanup
		'UPGRADE_NOTE: Object pMapAlgebraOp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pMapAlgebraOp = Nothing
		'UPGRADE_NOTE: Object pEnv may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pEnv = Nothing
		'UPGRADE_NOTE: Object pDEMOneCell may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pDEMOneCell = Nothing
		'UPGRADE_NOTE: Object pDEMTwoCell may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pDEMTwoCell = Nothing
		'UPGRADE_NOTE: Object pDEMRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pDEMRaster = Nothing
		'UPGRADE_NOTE: Object pDEMRasterProps may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pDEMRasterProps = Nothing
		'UPGRADE_NOTE: Object pRasterWorkspaceFactory may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRasterWorkspaceFactory = Nothing
		'UPGRADE_NOTE: Object pWorkspace may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pWorkspace = Nothing
		
	End Sub
	
	Public Sub init(ByVal pApp As ESRI.ArcGIS.Framework.IApplication)
		
		
		m_App = pApp
		
		
	End Sub
	
	
	Private Sub cmdQuit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdQuit.Click
		Me.Close()
	End Sub
End Class