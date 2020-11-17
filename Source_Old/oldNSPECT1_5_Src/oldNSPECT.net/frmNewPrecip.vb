Option Strict Off
Option Explicit On
Friend Class frmNewPrecip
	Inherits System.Windows.Forms.Form
	' *************************************************************************************
	' *  Perot Systems Government Services
	' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
	' *  frmNewPrecip
	' *************************************************************************************
	' *  Description: Form for entering a new precipitation scenarios
	' *  within NSPECT
	' *
	' *  Called By:  frmPrecip
	' *************************************************************************************
	
	
	Public m_pInputPrecipDS As ESRI.ArcGIS.Geodatabase.IRasterDataset
	
	
	
	'UPGRADE_WARNING: Event cboTimePeriod.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboTimePeriod_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboTimePeriod.SelectedIndexChanged
		
		If cboTimePeriod.SelectedIndex = 0 Then
			lblRainingDays.Visible = True
			txtRainingDays.Visible = True
		Else
			lblRainingDays.Visible = False
			txtRainingDays.Visible = False
		End If
		
	End Sub
	
	Private Sub cmdBrowseFile_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdBrowseFile.Click
		
		Dim pPrecipRasterDataset As ESRI.ArcGIS.Geodatabase.IRasterDataset
		Dim pDistUnit As ESRI.ArcGIS.Geometry.ILinearUnit
		Dim intUnit As Short
		Dim pProjCoord As ESRI.ArcGIS.Geometry.IProjectedCoordinateSystem
		
		On Error GoTo ErrHandler
		
		m_pInputPrecipDS = AddInputFromGxBrowserText(txtPrecipFile, "Choose Precipitation GRID", frmPrecip, 0)
		
		If Not m_pInputPrecipDS Is Nothing Then
			
			pPrecipRasterDataset = m_pInputPrecipDS
			
			If CheckSpatialReference(pPrecipRasterDataset) Is Nothing Then
				
				MsgBox("The GRID you have choosen has no spatial reference information.  Please define a projection before continuing.", MsgBoxStyle.Exclamation, "No Project Information Detected")
				Exit Sub
				
			Else
				
				pProjCoord = CheckSpatialReference(pPrecipRasterDataset)
				pDistUnit = pProjCoord.CoordinateUnit
				intUnit = pDistUnit.MetersPerUnit
				
				If intUnit = 1 Then
					cboGridUnits.SelectedIndex = 0
				Else
					cboGridUnits.SelectedIndex = 1
				End If
				
				cboGridUnits.Refresh()
			End If
		End If
		
		'UPGRADE_NOTE: Object pPrecipRasterDataset may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pPrecipRasterDataset = Nothing
		'UPGRADE_NOTE: Object pDistUnit may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pDistUnit = Nothing
		'UPGRADE_NOTE: Object pProjCoord may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pProjCoord = Nothing
		
		Exit Sub
ErrHandler: 
		Exit Sub
		
		
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		
		Dim intvbYesNo As Short
		
		intvbYesNo = MsgBox("Are you sure you want to exit?  All changes not saved will be lost.", MsgBoxStyle.YesNo, "Exit?")
		
		If intvbYesNo = MsgBoxResult.Yes Then
			If IsLoaded("frmPrj") Then
				frmPrj.cboPrecipScen.SelectedIndex = 0
			End If
		Else
			Exit Sub
		End If
		
		Me.Close()
		
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		
		Dim intType As Short
		Dim intRainingDays As Short
		
		
		Dim strCmdInsert As String
		If CheckParams Then
			'Process the time period
			intType = cboTimePeriod.SelectedIndex
			If intType = 0 Then
				intRainingDays = CShort(txtRainingDays.Text)
			Else
				intRainingDays = 0
			End If
			
			
			'Compose the INSERT statement.
			strCmdInsert = "INSERT INTO PrecipScenario " & "(Name, Description, PrecipFileName, PrecipGridUnits, PrecipUnits, Type, PrecipType, RainingDays) VALUES (" & "'" & Replace(CStr(txtPrecipName.Text), "'", "''") & "', " & "'" & Replace(CStr(txtDesc.Text), "'", "''") & "', " & "'" & Replace(txtPrecipFile.Text, "'", "''") & "', " & "" & cboGridUnits.SelectedIndex & ", " & "" & cboPrecipUnits.SelectedIndex & ", " & "" & intType & ", " & "" & cboPrecipType.SelectedIndex & ", " & "" & intRainingDays & ")"
			
			
			Debug.Print(strCmdInsert)
			
			If modUtil.UniqueName("PrecipScenario", (txtPrecipName.Text)) Then
				'Execute the statement.
				modUtil.g_ADOConn.Execute(strCmdInsert, ADODB.CommandTypeEnum.adCmdText)
				'Confirm
				MsgBox(txtPrecipName.Text & " successfully added.", MsgBoxStyle.OKOnly, "Record Added")
				
				If IsLoaded("frmPrj") Then
					frmPrj.cboPrecipScen.Items.Clear()
					modUtil.InitComboBox((frmPrj.cboPrecipScen), "PrecipScenario")
					frmPrj.cboPrecipScen.Items.Insert(frmPrj.cboPrecipScen.Items.Count, "New precipitation scenario...")
					frmPrj.cboPrecipScen.SelectedIndex = modUtil.GetCboIndex((txtPrecipName.Text), (frmPrj.cboPrecipScen))
					Me.Close()
				End If
				
				If IsLoaded("frmPrecip") Then
					
					frmPrecip.cboScenName.Items.Clear()
					modUtil.InitComboBox((frmPrecip.cboScenName), "PrecipScenario")
					frmPrecip.cboScenName.SelectedIndex = modUtil.GetCboIndex((txtPrecipName.Text), (frmPrecip.cboScenName))
					Me.Close()
					
				End If
				
			Else
				MsgBox("Name already in use.  Please choose a different one.", MsgBoxStyle.Critical, "Name In Use")
				txtPrecipName.Focus()
				Exit Sub
			End If
			
			
		End If
		
		
	End Sub
	
	
	
	
	Private Sub txtDesc_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtDesc.Validating
		Dim Cancel As Boolean = eventArgs.Cancel
		
		txtDesc.Text = Replace(txtDesc.Text, "'", "")
		
		eventArgs.Cancel = Cancel
	End Sub
	
	'UPGRADE_WARNING: Event txtPrecipName.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtPrecipName_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPrecipName.TextChanged
		
		txtPrecipName.Text = Replace(txtPrecipName.Text, "'", "")
		
	End Sub
	
	Private Function CheckParams() As Boolean
		
		'Check the inputs of the form, before saving
		If Len(txtDesc.Text) = 0 Then
			MsgBox("Please enter a description for this scenario", MsgBoxStyle.Critical, "Description Missing")
			txtDesc.Focus()
			CheckParams = False
			Exit Function
		End If
		
		If txtPrecipFile.Text = " " Or txtPrecipFile.Text = "" Then
			MsgBox("Please select a valid precipitation GRID before saving.", MsgBoxStyle.Critical, "GRID Missing")
			txtPrecipFile.Focus()
			CheckParams = False
			Exit Function
		End If
		
		If cboGridUnits.Text = "" Then
			MsgBox("Please select GRID units.", MsgBoxStyle.Critical, "Units Missing")
			cboGridUnits.Focus()
			CheckParams = False
			Exit Function
		End If
		
		If cboPrecipUnits.Text = "" Then
			MsgBox("Please select precipitation units.", MsgBoxStyle.Critical, "Units Missing")
			cboPrecipUnits.Focus()
			CheckParams = False
			Exit Function
		End If
		
		If Len(cboPrecipType.Text) = 0 Then
			MsgBox("Please select a Precipitation Type.", MsgBoxStyle.Critical, "Precipitation Type Missing")
			cboPrecipType.Focus()
			CheckParams = False
			Exit Function
		End If
		
		If Len(cboTimePeriod.Text) = 0 Then
			MsgBox("Please select a Time Period.", MsgBoxStyle.Critical, "Precipitation Time Period Missing")
			cboTimePeriod.Focus()
			CheckParams = False
			Exit Function
		End If
		
		If cboTimePeriod.SelectedIndex = 0 Then
			If Not IsNumeric(txtRainingDays.Text) Or Len(txtRainingDays.Text) = 0 Then
				MsgBox("Please enter a numeric value for Raining Days.", MsgBoxStyle.Critical, "Raining Days Value Incorrect")
				txtRainingDays.Focus()
				CheckParams = False
				Exit Function
			End If
		End If
		
		'if it got through all that, then set it to true
		CheckParams = True
		
		
	End Function
End Class