Option Strict Off
Option Explicit On
Friend Class frmSoilsSetup
	Inherits System.Windows.Forms.Form
	' *************************************************************************************
	' *  Perot Systems Government Services
	' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
	' *  frmSoilsSetup
	' *************************************************************************************
	' *  Description: Form for allowing the user to create soils and k-factor soils grids
	' *
	' *
	' *  Called By:  frmSoils
	' *************************************************************************************
	
	Private m_App As ESRI.ArcGIS.Framework.IApplication 'Application handle
	Private m_pRasterProps As ESRI.ArcGIS.DataSourcesRaster.IRasterProps 'Raster props garnered from DEM
	
	Public Sub init(ByVal pApp As ESRI.ArcGIS.Framework.IApplication)
		
		m_App = pApp
		
	End Sub
	
	Private Sub cmdBrowseFile_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdBrowseFile.Click
		
		On Error GoTo ErrHandler
		
		'browse...get output filename
		dlgOpenOpen.FileName = CStr(Nothing)
		'UPGRADE_WARNING: CommonDialog variable was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="671167DC-EA81-475D-B690-7A40C7BF4A23"'
		With dlgOpen
			'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			.Filter = MSG6
			.Title = "Open Soils Dataset"
			.FilterIndex = 1
			'UPGRADE_WARNING: FileOpenConstants constant FileOpenConstants.cdlOFNHideReadOnly was upgraded to OpenFileDialog.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
			'UPGRADE_WARNING: MSComDlg.CommonDialog property dlgOpen.Flags was upgraded to dlgOpenOpen.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
			'UPGRADE_WARNING: FileOpenConstants constant FileOpenConstants.cdlOFNHideReadOnly was upgraded to OpenFileDialog.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
			.ShowReadOnly = False
			'UPGRADE_WARNING: MSComDlg.CommonDialog property dlgOpen.Flags was upgraded to dlgOpenOpen.CheckFileExists which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
			.CheckFileExists = True
			.CheckPathExists = True
			'UPGRADE_WARNING: The CommonDialog CancelError property is not supported in Visual Basic .NET. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"'
			.CancelError = True
			.ShowDialog()
		End With
		
		If Len(dlgOpenOpen.FileName) > 0 Then
			txtSoilsDS.Text = Trim(dlgOpenOpen.FileName)
			PopulateCbo()
		End If
		
ErrHandler: 
		Exit Sub
		
	End Sub
	
	Private Sub PopulateCbo()
		
		'Populate cboSoilFields & cboSoilFieldsK with the fields in the selected Soils layer
		Dim i As Short
		Dim pFields As ESRI.ArcGIS.Geodatabase.IFields
		Dim pFeatureClass As ESRI.ArcGIS.Geodatabase.IFeatureClass
		
		cboSoilFields.Items.Clear()
		cboSoilFieldsK.Items.Clear()
		
		pFeatureClass = modUtil.ReturnFeatureClass((txtSoilsDS.Text))
		'UPGRADE_WARNING: Couldn't resolve default property of object pFeatureClass.Fields. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pFields = pFeatureClass.Fields
		
		'Pop both cbos with field names
		For i = 1 To pFields.FieldCount - 1
			cboSoilFields.Items.Add(pFields.Field(i).Name)
			cboSoilFieldsK.Items.Add(pFields.Field(i).Name)
		Next i
		
		'Cleanup
		'UPGRADE_NOTE: Object pFields may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFields = Nothing
		'UPGRADE_NOTE: Object pFeatureClass may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFeatureClass = Nothing
		
	End Sub
	
	Private Sub cmdQuit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdQuit.Click
		
		Dim intvbYesNo As Short
		
		intvbYesNo = MsgBox("Do you want to save changes you made to soils setup?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Exclamation, "N-SPECT")
		
		If intvbYesNo = MsgBoxResult.Yes Then
			Call SaveSoils()
		ElseIf intvbYesNo = MsgBoxResult.No Then 
			Me.Close()
		ElseIf intvbYesNo = MsgBoxResult.Cancel Then 
			Exit Sub
		End If
		
		
	End Sub
	
	Private Sub cmdSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSave.Click
		
		Call SaveSoils()
		
	End Sub
	
	Private Sub SaveSoils()
		
		'Check data, if OK create soils grids
		If ValidateData Then
			If CreateSoilsGrid((txtSoilsDS.Text), (cboSoilFields.Text), cboSoilFieldsK.Text) Then
				If frmSoils.Visible Then
					frmSoils.cboSoils.Items.Clear()
					modUtil.InitComboBox((frmSoils.cboSoils), "Soils")
					frmSoils.cboSoils.SelectedIndex = modUtil.GetCboIndex((txtSoilsName.Text), (frmSoils.cboSoils))
					Me.Close()
				End If
			Else
				Exit Sub
			End If
		Else
			Exit Sub
		End If
		
		
	End Sub
	
	
	Private Function ValidateData() As Boolean
		
		If Len(txtSoilsName.Text) > 0 Then
			If modUtil.UniqueName("Soils", (txtSoilsName.Text)) Then
				ValidateData = True
			Else
				MsgBox("The name you have chosen is already in use.  Please select another.", MsgBoxStyle.Critical, "Select Unique Name")
				ValidateData = False
				txtSoilsName.Focus()
				Exit Function
			End If
		Else
			MsgBox("Please enter a name.", MsgBoxStyle.Critical, "Soils Name Missing")
			ValidateData = False
			txtSoilsName.Focus()
			Exit Function
			
		End If
		
		If Len(txtSoilsDS.Text) = 0 Then
			MsgBox("Please select a soils dataset.", MsgBoxStyle.Critical, "Soils Dataset Missing")
			txtSoilsDS.Focus()
			ValidateData = False
			Exit Function
		Else
			ValidateData = True
		End If
		
		If Len(cboSoilFields.Text) = 0 Then
			MsgBox("Please select a soils attribute.", MsgBoxStyle.Critical, "Choose Soils Attribute")
			cboSoilFields.Focus()
			ValidateData = False
			Exit Function
		Else
			ValidateData = True
		End If
		
		If Len(cboSoilFieldsK.Text) = 0 Then
			MsgBox("Please select a k-factor soils attribute.", MsgBoxStyle.Critical, "Choose K-Factor Attribute")
			cboSoilFieldsK.Focus()
			ValidateData = False
			Exit Function
		Else
			ValidateData = True
		End If
		
		If Len(txtMUSLEVal.Text) > 0 Then
			If IsNumeric(CDbl(txtMUSLEVal.Text)) Then
				ValidateData = True
			Else
				MsgBox("Please enter a numeric value for the MUSLE equation.", MsgBoxStyle.Critical, "Numeric Values Only")
				ValidateData = False
			End If
		Else
			MsgBox("Please enter a value for the MUSLE equation.", MsgBoxStyle.Critical, "Missing Value")
			txtMUSLEVal.Focus()
			ValidateData = False
		End If
		
		If Len(txtMUSLEExp.Text) > 0 Then
			If IsNumeric(CDbl(txtMUSLEExp.Text)) Then
				ValidateData = True
			Else
				MsgBox("Please enter a numeric value for the MUSLE equation.", MsgBoxStyle.Critical, "Numeric Values Only")
				ValidateData = False
			End If
		Else
			MsgBox("Please enter a value for the MUSLE equation.", MsgBoxStyle.Critical, "Missing Value")
			txtMUSLEExp.Focus()
			ValidateData = False
		End If
		
		
	End Function
	
	Private Function CreateSoilsGrid(ByRef strSoilsFileName As String, ByRef strHydFieldName As String, Optional ByRef strKFactor As String = "") As Boolean
		'Incoming:
		'strSoilsFileName: string of soils file name path
		'strHydFieldName: string of hydrologic group attribute
		'strKFactor: string of K factor attribute
		
		On Error GoTo ErrHandler
		
		Dim pSoilsFeatClass As ESRI.ArcGIS.Geodatabase.IFeatureClass 'Soils Featureclass
		Dim pSoilsFeatCursor As ESRI.ArcGIS.Geodatabase.IFeatureCursor 'Cursor to loop through soils
		Dim pSoilsFeature As ESRI.ArcGIS.Geodatabase.IFeature 'Feature for use
		Dim lngHydFieldIndex As Integer 'HydGroup Field Index
		Dim lngNewHydFieldIndex As Integer 'New Group Field index
		Dim pFieldEdit As ESRI.ArcGIS.Geodatabase.IFieldEdit 'FieldEdit, case we have to add one
		Dim pField As ESRI.ArcGIS.Geodatabase.IField 'Field
		Dim strHydValue As String 'Hyd Value to check
		Dim lngValue As Integer 'Count
		Dim strSoilsRas As String 'Soils Raster Name
		Dim strSoilsKRas As String 'Soils K raster Name
		Dim strCmd As String 'String to insert new stuff in dbase
		Dim strOutSoils As String 'OutSoils name
		Dim strOutKSoils As String
		
		Dim pConversionOp As ESRI.ArcGIS.GeoAnalyst.IConversionOp 'Feat to Grid convert
		Dim pConversionOpK As ESRI.ArcGIS.GeoAnalyst.IConversionOp 'K Feat to Grid convert
		Dim pSoilsRaster As ESRI.ArcGIS.Geodatabase.IRasterDataset 'Soils raster ds
		Dim pSoilsKRaster As ESRI.ArcGIS.Geodatabase.IRasterDataset 'Soils k factor ds
		Dim pSoilsDS As ESRI.ArcGIS.Geodatabase.IGeoDataset
		Dim pSoilsKDS As ESRI.ArcGIS.Geodatabase.IGeoDataset
		Dim pSoilsFeatureDescriptor As ESRI.ArcGIS.GeoAnalyst.IFeatureClassDescriptor 'Descript for soils
		Dim pSoilsKFeatureDescriptor As ESRI.ArcGIS.GeoAnalyst.IFeatureClassDescriptor 'Descript for soils K
		
		Dim pEnv As ESRI.ArcGIS.GeoAnalyst.IRasterAnalysisEnvironment
		Dim pWSFact As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory
		Dim pWS As ESRI.ArcGIS.DataSourcesRaster.IRasterWorkspace
		
		'Get the soils featurclass
		pSoilsFeatClass = modUtil.ReturnFeatureClass(strSoilsFileName)
		
		'Check for fields
		'UPGRADE_WARNING: Couldn't resolve default property of object pSoilsFeatClass.FindField. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		lngHydFieldIndex = pSoilsFeatClass.FindField(strHydFieldName)
		'UPGRADE_WARNING: Couldn't resolve default property of object pSoilsFeatClass.FindField. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		lngNewHydFieldIndex = pSoilsFeatClass.FindField("GROUP")
		
		'If the GROUP field is missing, we have to add it
		If lngNewHydFieldIndex = -1 Then
			pFieldEdit = New ESRI.ArcGIS.Geodatabase.Field
			With pFieldEdit
				.Name = "GROUP"
				.Type = ESRI.ArcGIS.Geodatabase.esriFieldType.esriFieldTypeInteger
				.length = 2
			End With
			pField = pFieldEdit
			
			'UPGRADE_WARNING: Couldn't resolve default property of object pSoilsFeatClass.AddField. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pSoilsFeatClass.AddField(pField)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object pSoilsFeatClass.FindField. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			lngNewHydFieldIndex = pSoilsFeatClass.FindField("GROUP")
		End If
		
		lngValue = 1
		
		'Get all features in a cursor
		pSoilsFeatCursor = pSoilsFeatClass.Update(Nothing, False)
		
		'Get all the features into the cursor
		pSoilsFeature = pSoilsFeatCursor.NextFeature
		
		'Now calc the Values
		Do While Not pSoilsFeature Is Nothing
			modProgDialog.ProgDialog("Calculating soils values...", "Processing Soils", 0, pSoilsFeatClass.FeatureCount(Nothing), lngValue, (m_App.hwnd))
			'Find the current value
			If modProgDialog.g_boolCancel Then
				'UPGRADE_WARNING: Couldn't resolve default property of object pSoilsFeature.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				strHydValue = pSoilsFeature.Value(lngHydFieldIndex)
				'Based on current value, change GROUP to appropriate setting
				Select Case strHydValue
					Case "A"
						'UPGRADE_WARNING: Couldn't resolve default property of object pSoilsFeature.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						pSoilsFeature.Value(lngNewHydFieldIndex) = 1
					Case "B"
						'UPGRADE_WARNING: Couldn't resolve default property of object pSoilsFeature.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						pSoilsFeature.Value(lngNewHydFieldIndex) = 2
					Case "C"
						'UPGRADE_WARNING: Couldn't resolve default property of object pSoilsFeature.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						pSoilsFeature.Value(lngNewHydFieldIndex) = 3
					Case "D"
						'UPGRADE_WARNING: Couldn't resolve default property of object pSoilsFeature.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						pSoilsFeature.Value(lngNewHydFieldIndex) = 4
					Case "A/B"
						'UPGRADE_WARNING: Couldn't resolve default property of object pSoilsFeature.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						pSoilsFeature.Value(lngNewHydFieldIndex) = 2
					Case "B/C"
						'UPGRADE_WARNING: Couldn't resolve default property of object pSoilsFeature.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						pSoilsFeature.Value(lngNewHydFieldIndex) = 3
					Case "C/D"
						'UPGRADE_WARNING: Couldn't resolve default property of object pSoilsFeature.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						pSoilsFeature.Value(lngNewHydFieldIndex) = 4
					Case "B/D"
						'UPGRADE_WARNING: Couldn't resolve default property of object pSoilsFeature.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						pSoilsFeature.Value(lngNewHydFieldIndex) = 4
					Case "A/D"
						'UPGRADE_WARNING: Couldn't resolve default property of object pSoilsFeature.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						pSoilsFeature.Value(lngNewHydFieldIndex) = 4
					Case ""
						MsgBox("Your soils dataset contains missing values for Hydrologic Soils Attribute.  Please correct.", MsgBoxStyle.Critical, "Missing Values Detected")
						CreateSoilsGrid = False
						modProgDialog.KillDialog()
						Exit Function
				End Select
				'Update row and move to next
				pSoilsFeatCursor.UpdateFeature(pSoilsFeature)
				pSoilsFeature = pSoilsFeatCursor.NextFeature
				lngValue = lngValue + 1
			Else
				'If they cancel, kill the dialog
				modProgDialog.KillDialog()
				Exit Function
			End If
		Loop 
		
		'Close dialog
		modProgDialog.KillDialog()
		
		'STEP 2:
		'Now do the conversion: Convert soils layer to GRID using new
		'Group field as the value
		'First set the descriptor to use the 'Group' field
		pSoilsFeatureDescriptor = New ESRI.ArcGIS.GeoAnalyst.FeatureClassDescriptor
		pSoilsFeatureDescriptor.Create(pSoilsFeatClass, Nothing, "GROUP")
		
		pSoilsDS = pSoilsFeatureDescriptor
		
		'Conversion Operation
		pConversionOp = New ESRI.ArcGIS.GeoAnalyst.RasterConversionOp
		
		'Set the spat Environment
		pEnv = pConversionOp
		
		'Get the goodies from the rasterprops
		With pEnv
			.SetCellSize(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvValue, m_pRasterProps.MeanCellSize.X)
			.SetExtent(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvValue, m_pRasterProps.Extent, True)
			'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutSpatialReference. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.OutSpatialReference = m_pRasterProps.SpatialReference
		End With
		
		'Set the workspace
		pWSFact = New ESRI.ArcGIS.DataSourcesRaster.RasterWorkspaceFactory
		pWS = pWSFact.OpenFromFile(modUtil.SplitWorkspaceName(strSoilsFileName), m_App.hwnd)
		
		modProgDialog.ProgDialog("Converting Soils Dataset...", "Processing Soils", 0, 2, 1, (m_App.hwnd))
		
		If modProgDialog.g_boolCancel Then
			strOutSoils = modUtil.GetUniqueName("soils", modUtil.SplitWorkspaceName(strSoilsFileName))
			pSoilsRaster = New ESRI.ArcGIS.DataSourcesRaster.RasterDataset
			'UPGRADE_WARNING: Couldn't resolve default property of object pWS. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pSoilsRaster = pConversionOp.ToRasterDataset(pSoilsDS, "GRID", pWS, strOutSoils)
			
			strSoilsRas = pSoilsRaster.CompleteName
			
			'STEP 3:
			'Now do the conversion: Convert soils layer to GRID using
			'k factor field as the value
			'If they are doing a K factor then repeat the process, this time using 'k' field
			If Len(strKFactor) > 0 Then
				
				pSoilsKFeatureDescriptor = New ESRI.ArcGIS.GeoAnalyst.FeatureClassDescriptor
				pSoilsKFeatureDescriptor.Create(pSoilsFeatClass, Nothing, strKFactor)
				
				pSoilsKDS = pSoilsKFeatureDescriptor
				pConversionOpK = New ESRI.ArcGIS.GeoAnalyst.RasterConversionOp
				pEnv = pConversionOpK
				
				With pEnv
					.SetCellSize(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvValue, m_pRasterProps.MeanCellSize.X)
					.SetExtent(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvValue, m_pRasterProps.Extent, True)
					'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutSpatialReference. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.OutSpatialReference = m_pRasterProps.SpatialReference
				End With
				
				modProgDialog.ProgDialog("Converting Soils K Dataset...", "Processing Soils", 0, 2, 2, (m_App.hwnd))
				
				strOutKSoils = modUtil.GetUniqueName("soilsk", modUtil.SplitWorkspaceName(strSoilsFileName))
				pSoilsKRaster = New ESRI.ArcGIS.DataSourcesRaster.RasterDataset
				'UPGRADE_WARNING: Couldn't resolve default property of object pWS. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pSoilsKRaster = pConversionOpK.ToRasterDataset(pSoilsKDS, "GRID", pWS, strOutKSoils)
				
				strSoilsKRas = pSoilsKRaster.CompleteName
				
			Else
				modProgDialog.KillDialog()
			End If
			
			'STEP 4:
			'Now enter all into database
			strCmd = "INSERT INTO SOILS (NAME,SOILSFILENAME,SOILSKFILENAME,MUSLEVal,MUSLEExp) VALUES ('" & Replace(txtSoilsName.Text, "'", "''") & "', '" & Replace(strSoilsRas, "'", "''") & "', '" & Replace(strSoilsKRas, "'", "''") & "', " & CDbl(txtMUSLEVal.Text) & ", " & CDbl(txtMUSLEExp.Text) & ")"
			
			g_ADOConn.Execute(strCmd, ADODB.CommandTypeEnum.adCmdText)
			
			modProgDialog.KillDialog()
			
		Else
			modProgDialog.KillDialog()
		End If
		
		CreateSoilsGrid = True
		
		'Cleanup
		'UPGRADE_NOTE: Object pSoilsFeatClass may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pSoilsFeatClass = Nothing
		'UPGRADE_NOTE: Object pSoilsFeatCursor may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pSoilsFeatCursor = Nothing
		'UPGRADE_NOTE: Object pSoilsFeature may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pSoilsFeature = Nothing
		'UPGRADE_NOTE: Object pFieldEdit may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFieldEdit = Nothing
		'UPGRADE_NOTE: Object pField may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pField = Nothing
		'UPGRADE_NOTE: Object pConversionOp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pConversionOp = Nothing
		'UPGRADE_NOTE: Object pConversionOpK may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pConversionOpK = Nothing
		'UPGRADE_NOTE: Object pSoilsRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pSoilsRaster = Nothing
		'UPGRADE_NOTE: Object pSoilsKRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pSoilsKRaster = Nothing
		'UPGRADE_NOTE: Object pSoilsDS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pSoilsDS = Nothing
		'UPGRADE_NOTE: Object pSoilsKDS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pSoilsKDS = Nothing
		'UPGRADE_NOTE: Object pSoilsFeatureDescriptor may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pSoilsFeatureDescriptor = Nothing
		'UPGRADE_NOTE: Object pSoilsKFeatureDescriptor may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pSoilsKFeatureDescriptor = Nothing
		'UPGRADE_NOTE: Object pEnv may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pEnv = Nothing
		
		Exit Function
		
ErrHandler: 
		MsgBox(Err.Number & ": " & Err.Description)
		CreateSoilsGrid = False
	End Function
	
	Private Sub cmdDEMBrowse_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDEMBrowse.Click
		'Browse for DEM
		Dim pDEMRasterDataset As ESRI.ArcGIS.Geodatabase.IRasterDataset
		pDEMRasterDataset = AddInputFromGxBrowserText(txtDEMFile, "Choose DEM Dataset", Me, 0)
		
		If Not pDEMRasterDataset Is Nothing Then
			m_pRasterProps = pDEMRasterDataset.CreateDefaultRaster
		Else
			MsgBox("The Raster Dataset you have chosen is invalid.", MsgBoxStyle.Critical, "DEM Error")
			Exit Sub
		End If
		
		'UPGRADE_NOTE: Object pDEMRasterDataset may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pDEMRasterDataset = Nothing
		
	End Sub
End Class