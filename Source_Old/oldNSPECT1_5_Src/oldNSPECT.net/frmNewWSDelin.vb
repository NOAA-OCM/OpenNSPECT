Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmNewWSDelin
	Inherits System.Windows.Forms.Form
	' *************************************************************************************************
	' *  Perot Systems Government Services
	' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
	' *  frmNewWSDelin
	' *************************************************************************************************
	' *  Description: Form for browsing and maintaining watershed delineation
	' *  Scenarios within NSPECT
	' *  Called By:  frmWSDelin menu item New...
	' *************************************************************************************************
	
	Private boolChange(3) As Boolean 'Array set to track changes in controls: On Change, cmdCreate is enabled
	Public m_pInputDEMDS As ESRI.ArcGIS.Geodatabase.IRasterDataset 'DEM Dataset
	Public m_booProject As Boolean 'Boolean to determine if this was called from frmPrj
	
	Private m_strDemArray() As String
	Private m_strDemName As String 'DEM Name
	Private m_strFullDemName As String 'DEM GRID name
	Private m_strAccumFileName As String 'Accumulation GRID name
	Private m_strDirFileName As String 'Direction GRID name
	Private m_strFilledDEMFileName As String 'Filled DEM name
	Private m_strWShedFileName As String 'WShed Name
	Private m_strStreamLayer As String 'Stream Layer
	Private m_strLSFileName As String 'Lenght Slope GRID File name
	Private m_strNibbleName As String 'Nibble Grid File Name
	Private m_strDEM2BName As String '2 Cell DEM buffer grid path
	Private m_strWorkspace As String 'Workspace
	
	'Objects from the Watershed Delineation needed in the LS calc
	Private m_pFilledDEMGDS As ESRI.ArcGIS.Geodatabase.IGeoDataset
	Private m_pFlowDirRS As ESRI.ArcGIS.Geodatabase.IRasterDataset
	Private m_intGridUnits As Short 'Grid Units: 0 = meters, 1 = feet
	Private m_intCellSize As Short 'Cell Size of DEM Grid, used in Length Slope Calculation
	
	Private m_App As ESRI.ArcGIS.Framework.IApplication 'Application
	Private m_pMap As ESRI.ArcGIS.Carto.IMap 'Map
	
	Private Const dblSmall As Double = 0.03 '0.001    '
	Private Const dblMedium As Double = 0.06 '0.01    '-Subwatershed sizes (3%, 6%, 10%)
	Private Const dblLarge As Double = 0.1 '
	
	Private m_intSize As Short 'Index for Size Combo
	
	Const c_sModuleFileName As String = "frmNewWSDelin.frm"
	Private m_ParentHWND As Integer ' Set this to get correct parenting of Error handler forms
	
	'Hook ArcMap
	Public Sub init(ByVal pApp As ESRI.ArcGIS.Framework.IApplication)
		On Error GoTo ErrorHandler
		
		m_App = pApp
		Dim pDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
		pDoc = pApp.Document
		
		m_pMap = pDoc.FocusMap
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "init " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	
	'UPGRADE_WARNING: Event cboDEMUnits.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboDEMUnits_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboDEMUnits.SelectedIndexChanged
		On Error GoTo ErrorHandler
		
		boolChange(2) = True
		CheckEnabled()
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "cboDEMUnits_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	'UPGRADE_WARNING: Event cboStreamLayer.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboStreamLayer_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboStreamLayer.SelectedIndexChanged
		
		On Error GoTo ErrorHandler
		
		If cboStreamLayer.Text <> "" Then
			cmdOptions.Enabled = True
		Else
			cmdOptions.Enabled = False
		End If
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "cboStreamLayer_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	'UPGRADE_WARNING: Event cboSubWSSize.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboSubWSSize_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboSubWSSize.SelectedIndexChanged
		
		On Error GoTo ErrorHandler
		
		boolChange(3) = True
		CheckEnabled()
		m_intSize = cboSubWSSize.SelectedIndex
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "cboSubWSSize_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	
	
	'UPGRADE_WARNING: Event chkHydroCorr.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkHydroCorr_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkHydroCorr.CheckStateChanged
		On Error GoTo ErrorHandler
		
		Select Case chkHydroCorr.CheckState
			Case 1
				chkStreamAgree.Enabled = True
				cboStreamLayer.Enabled = True
				cmdOptions.Enabled = True
				g_boolHydCorr = True
				
			Case 0
				
				chkStreamAgree.Enabled = False
				cboStreamLayer.Enabled = False
				cmdOptions.Enabled = False
				g_boolHydCorr = False
				
		End Select
		
		CheckEnabled()
		
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "chkHydroCorr_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	'UPGRADE_WARNING: Event chkStreamAgree.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkStreamAgree_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkStreamAgree.CheckStateChanged
		On Error GoTo ErrorHandler
		
		Select Case chkStreamAgree.CheckState
			Case 1
				cboStreamLayer.Enabled = True
				lblStream(0).Enabled = True
				g_boolAgree = True
			Case 0
				cboStreamLayer.Enabled = False
				cmdOptions.Enabled = False
		End Select
		
		CheckEnabled()
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "chkStreamAgree_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	Private Sub cmdBrowseDEMFile_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdBrowseDEMFile.Click 'Load button for DEM File
		
		Dim pDEMRasterDataset As ESRI.ArcGIS.Geodatabase.IRasterDataset
		Dim pDistUnit As ESRI.ArcGIS.Geometry.ILinearUnit
		Dim intUnit As Short
		Dim pProjCoord As ESRI.ArcGIS.Geometry.IProjectedCoordinateSystem
		
		On Error GoTo ErrHandler
		
		m_pInputDEMDS = AddInputFromGxBrowserText(txtDEMFile, "Choose DEM GRID", Me, 0)
		pDEMRasterDataset = m_pInputDEMDS
		
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
		
		'Get the name
		m_strDemArray = Split(pDEMRasterDataset.CompleteName, "\")
		m_strDemName = m_strDemArray(UBound(m_strDemArray))
		
		'UPGRADE_NOTE: Object pDEMRasterDataset may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pDEMRasterDataset = Nothing
		'UPGRADE_NOTE: Object pDistUnit may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pDistUnit = Nothing
		'UPGRADE_NOTE: Object pProjCoord may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pProjCoord = Nothing
		
ErrHandler: 
		Exit Sub
		
		
	End Sub
	
	Private Sub cmdCreate_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCreate.Click
		
		Select Case Err.Number
			Case Is < 0
				Error(5)
			Case 1
				GoTo ErrHandler
		End Select
		
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		Dim intPos As Short
		Dim pRasterDataset As ESRI.ArcGIS.Geodatabase.IRasterDataset
		Dim fso As Scripting.FileSystemObject
		Dim Folder As Scripting.Folder
		
		Dim pWorkspace As ESRI.ArcGIS.Geodatabase.IWorkspace
		Dim pRasterWorkspaceFactory As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory
		Dim pPropertySet As ESRI.ArcGIS.esriSystem.IPropertySet
		
		pRasterWorkspaceFactory = New ESRI.ArcGIS.DataSourcesRaster.RasterWorkspaceFactory
		pPropertySet = New ESRI.ArcGIS.esriSystem.PropertySet
		fso = CreateObject("Scripting.FileSystemObject")
		
		If fso.FolderExists(txtDEMFile.Text) Or fso.FileExists(txtDEMFile.Text) Then
			
			If m_strDemName = "" Then
				'Get the names
				m_strDemArray = Split(txtDEMFile.Text, "\") 'Array
				m_strDemName = m_strDemArray(UBound(m_strDemArray)) 'Name of Raster
			End If
			
		Else
			MsgBox("The File you have choosen does not exist", MsgBoxStyle.Critical, "File Not Found")
			txtDEMFile.Focus()
			'UPGRADE_ISSUE: Unable to determine which constant to upgrade vbNormal to. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B3B44E51-B5F1-4FD7-AA29-CAD31B71F487"'
			'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			System.Windows.Forms.Cursor.Current = vbNormal
			Exit Sub
		End If
		
		intPos = InStrRev(txtDEMFile.Text, "\", -1)
		
		m_strWorkspace = VB.Left(txtDEMFile.Text, intPos)
		
		pRasterDataset = modUtil.OpenRasterDataset(m_strWorkspace, m_strDemName)
		
		Dim strCmdInsert As String
		If pRasterDataset Is Nothing Then
			MsgBox("Error:  Could not open DEM Raster.", MsgBoxStyle.Critical, "Could Not Open Dataset")
			Exit Sub
			
		Else
			
			If Not fso.FolderExists(modUtil.g_nspectPath & "\wsdelin\" & txtWSDelinName.Text) Then
				Folder = fso.CreateFolder(modUtil.g_nspectPath & "\wsdelin\" & txtWSDelinName.Text)
			Else
				MsgBox("Name in use.  Please select another.", MsgBoxStyle.Critical, "Choose New Name")
				txtWSDelinName.Focus()
				Exit Sub
			End If
			
			pWorkspace = pRasterWorkspaceFactory.OpenFromFile(modUtil.g_nspectPath & "\wsdelin\" & txtWSDelinName.Text, 0)
			
			'Give the call; if successful insert new record
			'UPGRADE_WARNING: Couldn't resolve default property of object pWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If DelineateWatershed(pRasterDataset, pWorkspace) Then
				
				'SQL Insert
				
				'DataBase Update
				'Compose the INSERT statement.
				strCmdInsert = "INSERT INTO WSDelineation " & "(Name, DEMFileName, DEMGridUnits, FlowDirFileName, FlowAccumFileName," & "FilledDEMFileName, HydroCorrected, StreamFileName, SubWSSize, WSFileName, LSFileName, NibbleFileName, DEM2bFileName) " & " VALUES (" & "'" & CStr(txtWSDelinName.Text) & "', " & "'" & CStr(txtDEMFile.Text) & "', " & "'" & cboDEMUnits.SelectedIndex & "', " & "'" & m_strDirFileName & "', " & "'" & m_strAccumFileName & "', " & "'" & m_strFilledDEMFileName & "', " & "'" & chkHydroCorr.CheckState & "', " & "'" & m_strStreamLayer & "', " & "'" & cboSubWSSize.SelectedIndex & "', " & "'" & m_strWShedFileName & "', " & "'" & m_strLSFileName & "', " & "'" & m_strNibbleName & "', " & "'" & m_strDEM2BName & "')"
				
				'Execute the statement.
				modUtil.g_ADOConn.Execute(strCmdInsert, ADODB.CommandTypeEnum.adCmdText)
				'UPGRADE_ISSUE: Unable to determine which constant to upgrade vbNormal to. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B3B44E51-B5F1-4FD7-AA29-CAD31B71F487"'
				'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
				'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
				System.Windows.Forms.Cursor.Current = vbNormal
				
				'Confirm
				MsgBox(txtWSDelinName.Text & " successfully added.", MsgBoxStyle.OKOnly, "Record Added")
				
				If g_boolNewWShed Then
					'frmPrj.Show
					frmPrj.Frame.Visible = True
					frmPrj.cboWSDelin.Items.Clear()
					modUtil.InitComboBox((frmPrj.cboWSDelin), "WSDelineation")
					frmPrj.cboWSDelin.SelectedIndex = modUtil.GetCboIndex((txtWSDelinName.Text), (frmPrj.cboWSDelin))
					Me.Close()
				Else
					Me.Close()
					frmWSDelin.Close()
				End If
			Else
				'UPGRADE_ISSUE: Unable to determine which constant to upgrade vbNormal to. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B3B44E51-B5F1-4FD7-AA29-CAD31B71F487"'
				'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
				'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
				System.Windows.Forms.Cursor.Current = vbNormal
				Exit Sub
			End If
		End If
		
		'Reset project boolean
		m_booProject = False
		
		'Cleanup
		'UPGRADE_NOTE: Object pRasterDataset may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRasterDataset = Nothing
		'UPGRADE_NOTE: Object fso may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		fso = Nothing
		'UPGRADE_NOTE: Object Folder may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Folder = Nothing
		'UPGRADE_NOTE: Object pWorkspace may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pWorkspace = Nothing
		'UPGRADE_NOTE: Object pRasterWorkspaceFactory may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRasterWorkspaceFactory = Nothing
		'UPGRADE_NOTE: Object pPropertySet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pPropertySet = Nothing
		
		Exit Sub
		
ErrHandler: 
		
		'UPGRADE_ISSUE: Unable to determine which constant to upgrade vbNormal to. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B3B44E51-B5F1-4FD7-AA29-CAD31B71F487"'
		'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = vbNormal
		MsgBox(Err.Number & Err.Description & " :New Watershed Delineation")
		
	End Sub
	
	Private Sub cmdQuit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdQuit.Click
		On Error GoTo ErrorHandler
		
		Dim intvbYesNo As Short
		
		intvbYesNo = MsgBox("Are you sure you want to exit?  Your changes will not be saved.", MsgBoxStyle.YesNo, "Exit")
		
		If intvbYesNo = MsgBoxResult.Yes Then
			If m_booProject Then
				frmPrj.cboWSDelin.SelectedIndex = 0
				Me.Close()
			Else
				Me.Close()
			End If
		Else
			Exit Sub
		End If
		
		m_booProject = False
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "cmdQuit_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	Private Sub frmNewWSDelin_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		On Error GoTo ErrorHandler
		
		'Init bool variables
		g_boolAgree = False
		g_boolHydCorr = False
		g_boolParams = False
		
		Dim i As Short
		
		For i = 0 To UBound(boolChange)
			boolChange(i) = False
		Next i
		
		AddFeatureLayerToComboBox(cboStreamLayer, m_pMap, "line")
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	Private Sub frmNewWSDelin_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		On Error GoTo ErrorHandler
		
		m_booProject = False
		'UPGRADE_NOTE: Object m_pMap may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_pMap = Nothing
		'UPGRADE_NOTE: Object m_App may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_App = Nothing
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "Form_Unload " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	'UPGRADE_WARNING: Event txtDEMFile.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtDEMFile_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDEMFile.TextChanged
		On Error GoTo ErrorHandler
		
		
		boolChange(1) = True
		CheckEnabled()
		
		
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "txtDEMFile_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	'UPGRADE_WARNING: Event txtWSDelinName.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtWSDelinName_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWSDelinName.TextChanged
		On Error GoTo ErrorHandler
		
		boolChange(0) = True
		CheckEnabled()
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "txtWSDelinName_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	
	Public Sub CheckEnabled()
		On Error GoTo ErrorHandler
		
		If chkHydroCorr.CheckState And chkStreamAgree.CheckState Then
			If boolChange(0) And boolChange(1) And boolChange(2) And boolChange(3) And g_boolParams Then
				cmdCreate.Enabled = True
			Else
				cmdCreate.Enabled = False
			End If
		ElseIf chkHydroCorr.CheckState And Not chkStreamAgree.CheckState Then 
			If boolChange(0) And boolChange(1) And boolChange(2) And boolChange(3) Then
				cmdCreate.Enabled = True
			Else
				cmdCreate.Enabled = False
			End If
		ElseIf Not chkHydroCorr.CheckState Then 
			If boolChange(0) And boolChange(1) And boolChange(2) And boolChange(3) Then
				cmdCreate.Enabled = True
			Else
				cmdCreate.Enabled = False
			End If
		End If
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "CheckEnabled " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	
	'***************************************************************************************************
	Private Function DelineateWatershed(ByRef pSurfaceDatasetIn As ESRI.ArcGIS.Geodatabase.IRasterDataset, ByRef pWorkspace As ESRI.ArcGIS.DataSourcesRaster.IRasterWorkspace) As Boolean
		
		On Error GoTo ErrorHandler
		
		'Map Stuff
		Dim pMap As ESRI.ArcGIS.Carto.IMap
		Dim pMxDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
		
		'Hydro Operations
		Dim pFillHydrologyOp As ESRI.ArcGIS.SpatialAnalyst.IHydrologyOp ' Fill
		Dim pFlowDirHydrologyOp As ESRI.ArcGIS.SpatialAnalyst.IHydrologyOp ' Flow Direction
		Dim pAccumHydrologyOp As ESRI.ArcGIS.SpatialAnalyst.IHydrologyOp ' Flow Accumulation
		Dim pWaterShedOp As ESRI.ArcGIS.SpatialAnalyst.IHydrologyOp ' Watershed
		Dim pStreamLinkOp As ESRI.ArcGIS.SpatialAnalyst.IHydrologyOp ' Streamlink
		Dim pBasinOp As ESRI.ArcGIS.SpatialAnalyst.IHydrologyOp 'Creates polys out of watershed grid
		
		'Declare the geodataset objects
		Dim pFlowDirRaster As ESRI.ArcGIS.Geodatabase.IRaster 'Flow Direction
		Dim pAccumRaster As ESRI.ArcGIS.Geodatabase.IRaster 'Flow Accumulation
		Dim pFillRaster As ESRI.ArcGIS.Geodatabase.IRaster 'Fill GDS
		Dim pWShedRaster As ESRI.ArcGIS.Geodatabase.IRaster 'WaterShed GDS
		Dim pBasinRaster As ESRI.ArcGIS.Geodatabase.IRaster 'Basin GDS
		Dim pSurface As ESRI.ArcGIS.Geodatabase.IGeoDataset 'Incoming surface
		
		Dim pEnv As ESRI.ArcGIS.GeoAnalyst.IRasterAnalysisEnvironment 'Analysis Environment
		Dim pDEMRasterp As ESRI.ArcGIS.Geodatabase.IRaster
		Dim pDEMRasterProps As ESRI.ArcGIS.DataSourcesRaster.IRasterProps
		
		Dim pAccumStats As ESRI.ArcGIS.DataSourcesRaster.IRasterStatistics 'Flow Accum Stats-used to get max
		Dim dblMax As Double
		Dim pExtr As ESRI.ArcGIS.SpatialAnalyst.IExtractionOp
		Dim pRasDes As ESRI.ArcGIS.GeoAnalyst.IRasterDescriptor
		Dim pQueryFilter As ESRI.ArcGIS.Geodatabase.IQueryFilter
		Dim pStream1 As ESRI.ArcGIS.Geodatabase.IGeoDataset
		Dim pStream2 As ESRI.ArcGIS.Geodatabase.IGeoDataset
		
		'Added 12/18
		Dim pEnvelope As ESRI.ArcGIS.Geometry.IEnvelope
		'Set pEnvelope = New Envelope
		
		'Featureclass objects
		Dim pBasinFeatClass As ESRI.ArcGIS.Geodatabase.IFeatureClass 'Basin Featureclass
		
		'Raster to Shape
		Dim pBasinRastConvert As ESRI.ArcGIS.GeoAnalyst.IRasterConvertHelper
		Dim pShedRastConvert As ESRI.ArcGIS.GeoAnalyst.IRasterConvertHelper
		
		Dim strProgTitle As String
		strProgTitle = "Watershed Delineation Processing..."
		
		'Set the map stuff up
		pMxDoc = m_App.Document
		pMap = pMxDoc.FocusMap
		
		' First, check for Spatial Analyst License
		If Not modUtil.CheckSpatialAnalystLicense Then
			MsgBox("There is no Spatial Analyst license available.", MsgBoxStyle.Critical, "Spatial Analyst License Necessary")
			DelineateWatershed = False
			Exit Function
		End If
		
		'Have to hide the existing forms, to be able to show the progress
		If g_boolNewWShed Then
			Me.Hide()
			frmPrj.Hide()
		ElseIf Not g_boolNewWShed Then 
			Me.Hide()
			frmWSDelin.Hide()
		End If
		
		'Initialize the Hydro Ops
		pFlowDirHydrologyOp = New ESRI.ArcGIS.SpatialAnalyst.RasterHydrologyOp
		pAccumHydrologyOp = New ESRI.ArcGIS.SpatialAnalyst.RasterHydrologyOp
		pFillHydrologyOp = New ESRI.ArcGIS.SpatialAnalyst.RasterHydrologyOp
		pWaterShedOp = New ESRI.ArcGIS.SpatialAnalyst.RasterHydrologyOp
		pBasinOp = New ESRI.ArcGIS.SpatialAnalyst.RasterHydrologyOp
		pExtr = New ESRI.ArcGIS.SpatialAnalyst.RasterExtractionOp
		pStreamLinkOp = New ESRI.ArcGIS.SpatialAnalyst.RasterHydrologyOp
		
		'Get the rasterprops through pDEMRasterp, a default raster of pSurfaceDatasetIn
		pDEMRasterp = pSurfaceDatasetIn.CreateDefaultRaster
		pDEMRasterProps = pDEMRasterp
		pSurface = pSurfaceDatasetIn
		
		'Get cell size from DEM; needed later in the Length Slope Calculation
		m_intCellSize = pDEMRasterProps.MeanCellSize.X
		
		'Expand the envelope 2 cell sizes to account for later analysis
		'pEnvelope.PutCoords pDEMRasterProps.Extent.XMin - (2 * m_intCellSize), _
		'pDEMRasterProps.Extent.YMin - (2 * m_intCellSize), _
		'pDEMRasterProps.Extent.XMax + (2 * m_intCellSize), _
		'pDEMRasterProps.Extent.YMax + (2 * m_intCellSize)
		
		pEnvelope = pDEMRasterProps.Extent
		pEnvelope.Expand(m_intCellSize * 2, m_intCellSize * 2, False)
		
		
		'Set the Environment
		pEnv = pFillHydrologyOp
		With pEnv
			.SetCellSize(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvValue, pDEMRasterProps.MeanCellSize.X)
			'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.OutWorkspace = pWorkspace
			'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutSpatialReference. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.OutSpatialReference = pSurface.SpatialReference
			.SetExtent(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvValue, pEnvelope) 'pDEMRasterProps.Extent
			.SetAsNewDefaultEnvironment()
		End With
		
		'STEP 1:  Fill the Surface
		'if hydrocorrect, then skip the Fill, just use the incoming DEM
		If chkHydroCorr.CheckState = 1 Then
			pFillRaster = pSurfaceDatasetIn.CreateDefaultRaster
		Else
			
			'Call to ProgDialog to use throughout process: keep user informed.
			modProgDialog.ProgDialog("Filling DEM...", strProgTitle, 0, 10, 1, (m_App.hwnd))
			
			If modProgDialog.g_boolCancel Then
				pFillRaster = pFillHydrologyOp.Fill(pSurface)
			Else
				GoTo ProgCancel
			End If
			
		End If 'End if Fill
		
		'STEP 2: Flow Direction
		pEnv = pFlowDirHydrologyOp
		With pEnv
			.SetCellSize(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvValue, pDEMRasterProps.MeanCellSize.X)
			'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.OutWorkspace = pWorkspace
			'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutSpatialReference. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.OutSpatialReference = pSurface.SpatialReference
			.SetExtent(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvValue, pEnvelope) 'pDEMRasterProps.Extent
		End With
		
		modProgDialog.ProgDialog("Computing Flow Direction...", strProgTitle, 0, 10, 2, (m_App.hwnd))
		
		If modProgDialog.g_boolCancel Then
			'UPGRADE_WARNING: Couldn't resolve default property of object pFillRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pFlowDirRaster = pFlowDirHydrologyOp.FlowDirection(pFillRaster, False, False)
		Else
			GoTo ProgCancel
		End If
		
		'STEP 3: Flow Accumulation
		pEnv = pAccumHydrologyOp
		With pEnv
			.SetCellSize(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvValue, pDEMRasterProps.MeanCellSize.X)
			'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.OutWorkspace = pWorkspace
			'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutSpatialReference. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.OutSpatialReference = pSurface.SpatialReference
			.SetExtent(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvValue, pEnvelope) 'pDEMRasterProps.Extent
		End With
		
		modProgDialog.ProgDialog("Computing Flow Accumulation...", strProgTitle, 0, 10, 3, (m_App.hwnd))
		
		If modProgDialog.g_boolCancel Then
			'UPGRADE_WARNING: Couldn't resolve default property of object pFlowDirRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pAccumRaster = pAccumHydrologyOp.FlowAccumulation(pFlowDirRaster)
		Else
			GoTo ProgCancel
		End If
		
		'STEP 4: Stream Layer
		'How it's done:
		'1. Get the Stat.Max of the Accum GRID
		'2. Multiply the Max * percentage as deemed by user: small = 0.01% med = 1% large = 10%
		'3. Query Accumulation GRID for all values large than result from #2
		'4. Use #3 in Watershed method with the Flow Direction Grid
		
		Dim pAccumBands As ESRI.ArcGIS.DataSourcesRaster.IRasterBandCollection
		Dim pAccumBand As ESRI.ArcGIS.DataSourcesRaster.IRasterBand
		
		pAccumBands = pAccumRaster
		pAccumBand = pAccumBands.Item(0)
		
		pAccumStats = pAccumBand.Statistics
		dblMax = pAccumStats.Maximum
		
		'Initialize ExtractionOp
		pRasDes = New ESRI.ArcGIS.GeoAnalyst.RasterDescriptor
		pQueryFilter = New ESRI.ArcGIS.Geodatabase.QueryFilter
		
		Dim dblSubShedSize As Double
		
		Select Case m_intSize
			Case 0 'small
				dblSubShedSize = dblMax * dblSmall
			Case 1 'medium
				dblSubShedSize = dblMax * dblMedium
			Case 2 'large
				dblSubShedSize = dblMax * dblLarge
		End Select
		
		pQueryFilter.WhereClause = " value > " & dblSubShedSize
		pRasDes.Create(pAccumRaster, pQueryFilter, "value")
		
		pEnv = pExtr
		With pEnv
			.SetCellSize(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvValue, pDEMRasterProps.MeanCellSize.X)
			'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.OutWorkspace = pWorkspace
			'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutSpatialReference. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.OutSpatialReference = pSurface.SpatialReference
			.SetExtent(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvValue, pEnvelope) 'pDEMRasterProps.Extent
		End With
		
		pStream1 = pExtr.Attribute(pRasDes)
		
		'Step 5: Using Hydrology Op to create stream network
		pEnv = pStreamLinkOp
		With pEnv
			.SetCellSize(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvValue, pDEMRasterProps.MeanCellSize.X)
			'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.OutWorkspace = pWorkspace
			'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutSpatialReference. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.OutSpatialReference = pSurface.SpatialReference
			.SetExtent(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvValue, pEnvelope) ' pDEMRasterProps.Extent
		End With
		
		modProgDialog.ProgDialog("Creating Stream Network...", strProgTitle, 0, 10, 4, (m_App.hwnd))
		
		If modProgDialog.g_boolCancel Then
			'UPGRADE_WARNING: Couldn't resolve default property of object pFlowDirRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pStream2 = pStreamLinkOp.StreamLink(pStream1, pFlowDirRaster)
		Else
			GoTo ProgCancel
		End If
		
		'Step 6: Do WaterShed Op
		pEnv = pWaterShedOp
		With pEnv
			.SetCellSize(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvValue, pDEMRasterProps.MeanCellSize.X)
			'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.OutWorkspace = pWorkspace
			'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutSpatialReference. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.OutSpatialReference = pSurface.SpatialReference
			.SetExtent(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvValue, pEnvelope) 'pDEMRasterProps.Extent
		End With
		
		modProgDialog.ProgDialog("Creating Watershed GRID...", strProgTitle, 0, 10, 5, (m_App.hwnd))
		
		If modProgDialog.g_boolCancel Then
			'UPGRADE_WARNING: Couldn't resolve default property of object pFlowDirRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pWShedRaster = pWaterShedOp.Watershed(pFlowDirRaster, pStream2)
		Else
			GoTo ProgCancel
		End If
		
		pBasinRastConvert = New ESRI.ArcGIS.GeoAnalyst.RasterConvertHelper
		
		'UPGRADE_WARNING: Couldn't resolve default property of object pWShedRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pBasinFeatClass = pBasinRastConvert.ToShapefile(pWShedRaster, ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolygon, pEnv)
		SetFeatureClassName(pBasinFeatClass, "basinpolytemp")
		
		'STEP 7: Basin
		'Ed Removed Basin operation from the code 11/19/2007
		'Now using polygons output from the watershed.
		'    Set pEnv = pBasinOp
		'    With pEnv
		'       .SetCellSize esriRasterEnvValue, pDEMRasterProps.MeanCellSize.X
		'       Set .OutWorkspace = pWorkspace
		'       Set .OutSpatialReference = pSurface.SpatialReference
		'       .SetExtent esriRasterEnvValue, pDEMRasterProps.Extent
		'    End With
		'
		'    modProgDialog.ProgDialog "Creating Watershed Shapefile...", strProgTitle, _
		''       0, 10, 6, m_App.hWnd
		'
		'    If modProgDialog.g_boolCancel Then
		'        Set pBasinRaster = pBasinOp.Basin(pFlowDirRaster)
		'
		'        With pEnv
		'            .SetCellSize esriRasterEnvValue, pDEMRasterProps.MeanCellSize.X
		'            Set .OutWorkspace = pWorkspace
		'            Set .OutSpatialReference = pSurface.SpatialReference
		'            .SetExtent esriRasterEnvValue, pDEMRasterProps.Extent
		'        End With
		'
		'        Set pBasinRastConvert = New RasterConvertHelper
		'
		'        Set pBasinFeatClass = pBasinRastConvert.ToShapefile(pBasinRaster, esriGeometryPolygon, pEnv)
		'        SetFeatureClassName pBasinFeatClass, "basinpolytemp"
		'    Else
		'       GoTo ProgCancel
		'    End If
		
		'STEP 8: Get rid of small polys in the basin shapefile along the coast
		Dim pFinalBasinClass As ESRI.ArcGIS.Geodatabase.IFeatureClass
		
		With pEnv
			.SetCellSize(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvValue, pDEMRasterProps.MeanCellSize.X)
			'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.OutWorkspace = pWorkspace
			'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutSpatialReference. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.OutSpatialReference = pSurface.SpatialReference
			.SetExtent(ESRI.ArcGIS.GeoAnalyst.esriRasterEnvSettingEnum.esriRasterEnvValue, pEnvelope) 'pDEMRasterProps.Extent
		End With
		
		pFinalBasinClass = RemoveSmallPolys(pBasinFeatClass, pFillRaster, pEnv)
		'END STEP 7
		
		'Save the flow direction as a new GRID
		modProgDialog.ProgDialog("Saving Fill GRID...", strProgTitle, 0, 10, 7, (m_App.hwnd))
		
		If modProgDialog.g_boolCancel Then
			
			'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Not Len(modUtil.MakePerminentGrid(pFillRaster, pEnv.OutWorkspace.PathName, "demfill")) > 0 Then
				MsgBox("Could Not Save DEM Fill GRID", MsgBoxStyle.Critical, "Error Saving File")
				DelineateWatershed = False
				Exit Function
			End If
			
			modProgDialog.ProgDialog("Saving Flow Direction GRID...", strProgTitle, 0, 10, 8, (m_App.hwnd))
			
			'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Not Len(modUtil.MakePerminentGrid(pFlowDirRaster, pEnv.OutWorkspace.PathName, "flowdir")) > 0 Then
				MsgBox("Could Not Save Flow Direction GRID", MsgBoxStyle.Critical, "Error Saving File")
				DelineateWatershed = False
				Exit Function
			End If
			
			modProgDialog.ProgDialog("Saving Flow Accumulation...", strProgTitle, 0, 10, 9, (m_App.hwnd))
			
			'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Not Len(modUtil.MakePerminentGrid(pAccumRaster, pEnv.OutWorkspace.PathName, "flowacc")) > 0 Then
				MsgBox("Could Not Save Flow Accumulation GRID", MsgBoxStyle.Critical, "Error Saving File")
				DelineateWatershed = False
				Exit Function
			End If
			
			modProgDialog.ProgDialog("Saving Watersheds", strProgTitle, 0, 10, 10, (m_App.hwnd))
			'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Not Len(modUtil.MakePerminentGrid(pWShedRaster, pEnv.OutWorkspace.PathName, "wshed")) > 0 Then
				MsgBox("Could Not Save Watershed GRID", MsgBoxStyle.Critical, "Error Saving File")
				DelineateWatershed = False
				Exit Function
			End If
			
		Else
			GoTo ProgCancel
		End If
		
		modProgDialog.KillDialog()
		
		'With all of that done, now go get the name of the LS Grid while actually computing said LS Grid
		'UPGRADE_WARNING: Couldn't resolve default property of object pWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_strLSFileName = CalcLengthSlope(pFillRaster, pFlowDirRaster, pAccumRaster, pEnv, "0", pWorkspace)
		
		'Now get file paths to throw back in Database
		'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_strAccumFileName = pEnv.OutWorkspace.PathName & "flowacc"
		'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_strDirFileName = pEnv.OutWorkspace.PathName & "flowdir"
		
		If chkHydroCorr.CheckState = 1 Then
			m_strFilledDEMFileName = txtDEMFile.Text
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_strFilledDEMFileName = pEnv.OutWorkspace.PathName & "demfill"
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_strWShedFileName = pEnv.OutWorkspace.PathName & "basinpoly"
		
		DelineateWatershed = True
		
		'Clean up
		'UPGRADE_NOTE: Object pMap may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pMap = Nothing
		'UPGRADE_NOTE: Object pMxDoc may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pMxDoc = Nothing
		'UPGRADE_NOTE: Object pSurface may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pSurface = Nothing
		'UPGRADE_NOTE: Object pFlowDirRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFlowDirRaster = Nothing
		'UPGRADE_NOTE: Object pAccumRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pAccumRaster = Nothing
		'UPGRADE_NOTE: Object pFillRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFillRaster = Nothing
		'UPGRADE_NOTE: Object pWShedRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pWShedRaster = Nothing
		'UPGRADE_NOTE: Object pBasinRaster may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pBasinRaster = Nothing
		'UPGRADE_NOTE: Object pEnv may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pEnv = Nothing
		'UPGRADE_NOTE: Object pAccumBands may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pAccumBands = Nothing
		'UPGRADE_NOTE: Object pAccumBand may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pAccumBand = Nothing
		'UPGRADE_NOTE: Object pBasinRastConvert may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pBasinRastConvert = Nothing
		'Hydro Operations
		'UPGRADE_NOTE: Object pFillHydrologyOp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFillHydrologyOp = Nothing
		'UPGRADE_NOTE: Object pFlowDirHydrologyOp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFlowDirHydrologyOp = Nothing
		'UPGRADE_NOTE: Object pAccumHydrologyOp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pAccumHydrologyOp = Nothing
		'UPGRADE_NOTE: Object pWaterShedOp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pWaterShedOp = Nothing
		'UPGRADE_NOTE: Object pStreamLinkOp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pStreamLinkOp = Nothing
		'UPGRADE_NOTE: Object pBasinOp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pBasinOp = Nothing
		
		Exit Function
ProgCancel: 
		modProgDialog.KillDialog()
		MsgBox("The watershed delineation process has been stopped by the user.  Changes have been discarded.", MsgBoxStyle.Critical, "Process Stopped")
		DelineateWatershed = False
		
ErrorHandler: 
		If Err.Number = -2147217297 Then 'User cancelled operation
			modProgDialog.g_boolCancel = False
			DelineateWatershed = False
		Else
			HandleError(False, "DelineateWatershed " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
			DelineateWatershed = False
			modProgDialog.KillDialog()
		End If
		
	End Function
	
	Private Function RemoveSmallPolys(ByRef pFeatureClass As ESRI.ArcGIS.Geodatabase.IFeatureClass, ByRef pDEMRaster As ESRI.ArcGIS.Geodatabase.IRaster, ByRef pEnv As ESRI.ArcGIS.GeoAnalyst.IRasterAnalysisEnvironment) As ESRI.ArcGIS.Geodatabase.IFeatureClass
		On Error GoTo ErrorHandler
		
		
		Dim pAlgebraOp As ESRI.ArcGIS.SpatialAnalyst.IMapAlgebraOp
		Dim pDEMOutline As ESRI.ArcGIS.Geodatabase.IRaster
		Dim pDEMGeo As ESRI.ArcGIS.Geodatabase.IGeoDataset
		Dim pDEMFeatClass As ESRI.ArcGIS.Geodatabase.IFeatureClass
		Dim pRasterConvert As ESRI.ArcGIS.GeoAnalyst.IConversionOp
		
		'Basin featureclass editing stuff
		Dim pWorkspace As ESRI.ArcGIS.Geodatabase.IWorkspace
		Dim pFeatureCursor As ESRI.ArcGIS.Geodatabase.IFeatureCursor
		Dim pFeature As ESRI.ArcGIS.Geodatabase.IFeature
		Dim pArea As ESRI.ArcGIS.Geometry.IArea
		Dim pExtension As ESRI.ArcGIS.esriSystem.IExtension
		Dim pExtMgr As ESRI.ArcGIS.esriSystem.IExtensionManager
		Dim pEditor As ESRI.ArcGIS.Editor.IEditor
		Dim pUID As New ESRI.ArcGIS.esriSystem.UID
		
		'Union goodies
		Dim pDEMTable As ESRI.ArcGIS.Geodatabase.ITable
		Dim pBasinTable As ESRI.ArcGIS.Geodatabase.ITable
		Dim pFeatClassName As ESRI.ArcGIS.Geodatabase.IFeatureClassName
		Dim pNewWSName As ESRI.ArcGIS.Geodatabase.IWorkspaceName
		Dim pDatasetName As ESRI.ArcGIS.Geodatabase.IDatasetName
		Dim pOutputFeatClass As ESRI.ArcGIS.Geodatabase.IFeatureClass
		Dim pBGP As ESRI.ArcGIS.Carto.IBasicGeoprocessor
		Dim pEnvNew As ESRI.ArcGIS.GeoAnalyst.IRasterAnalysisEnvironment
		Dim pTempDEM As ESRI.ArcGIS.Geodatabase.IDataset
		Dim pTempBasin As ESRI.ArcGIS.Geodatabase.IDataset
		
		Dim strExpression As String
		Dim dblArea As Double
		Dim strWorkspace As String
		pEnvNew = pEnv
		
		'UPGRADE_WARNING: Couldn't resolve default property of object pEnvNew.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		strWorkspace = pEnvNew.OutWorkspace.PathName
		
		pAlgebraOp = New ESRI.ArcGIS.SpatialAnalyst.RasterMapAlgebraOp
		pEnvNew = pAlgebraOp
		
		'#1 First step, get a 1, 0 representation of the DEM, so we can get the outline
		'UPGRADE_WARNING: Couldn't resolve default property of object pDEMRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pAlgebraOp.BindRaster(pDEMRaster, "DEMo")
		strExpression = "Con([DEMo] >= 0, 1, 0)"
		
		pDEMOutline = pAlgebraOp.Execute(strExpression)
		
		pAlgebraOp.UnbindRaster("DEMo")
		
		'#2 init the rasterconverter and create the poly
		pRasterConvert = New ESRI.ArcGIS.GeoAnalyst.RasterConversionOp
		'Convert the DEM outline to a polygon and save as 'demoutline'
		pWorkspace = SetFeatureShapeWorkspace(strWorkspace)
		'UPGRADE_WARNING: Couldn't resolve default property of object pDEMOutline. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pDEMFeatClass = pRasterConvert.ToFeatureData(pDEMOutline, ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolygon, pWorkspace, "demout")
		
		modUtil.AddFeatureLayer(m_App, pDEMFeatClass, "demout")
		modUtil.AddFeatureLayer(m_App, pFeatureClass, "basinpolytemp")
		
		'#3 determine size of 'small' watersheds, this is the area
		'Number of cells in the DEM area that are not null * a number dave came up with * CellSize Squared
		'dblArea = ((modUtil.ReturnCellCount(pDEMRaster)) * 0.004) * m_intCellSize * m_intCellSize
		
		'#4 Now with the Area of small sheds determined we can remove polygons that are too small.  To do this
		'   simply loop through the features and test the area.
		'Set pFeatureCursor = pFeatureClass.Update(Nothing, False)
		'Set pFeature = pFeatureCursor.NextFeature
		
		'Set up the Editor
		'pUID.Value = "{F8842F20-BB23-11D0-802B-0000F8037368}"
		
		'Set pExtMgr = m_App
		'Set pExtension = pExtMgr.FindExtension(pUID)
		'Set pEditor = pExtension
		
		'pEditor.StartEditing pWorkspace
		
		'Do While Not pFeature Is Nothing
		'    Set pArea = pFeature.Shape
		
		'    If pArea.Area <= dblArea Then
		'        pFeatureCursor.DeleteFeature
		'    End If
		
		'    Set pFeature = pFeatureCursor.NextFeature
		'Loop
		
		'Stop and Save
		'pFeatureCursor.Flush
		'pEditor.StopEditing True
		
		'Set pFeatureCursor = Nothing
		
		'Have to add the damn thing in to do the union in the next step
		'modUtil.AddFeatureLayer m_App, pFeatureClass, "Basin Polygons"
		
		'5#  Now, time to union the outline of the of the DEM with the newly paired down basin poly
		pDEMTable = m_pMap.Layer(1)
		pBasinTable = m_pMap.Layer(0)
		
		pFeatClassName = New ESRI.ArcGIS.Geodatabase.FeatureClassName
		
		' Set output location and feature class name
		pNewWSName = New ESRI.ArcGIS.Geodatabase.WorkspaceName
		pNewWSName.WorkspaceFactoryProgID = "esriDataSourcesFile.ShapeFileWorkspaceFactory.1"
		pNewWSName.PathName = pWorkspace.PathName
		
		pDatasetName = pFeatClassName
		pDatasetName.Name = "basinpoly"
		
		pDatasetName.WorkspaceName = pNewWSName
		
		' Set the tolerance.  Passing 0.0 causes the default tolerance to be used.
		' The default tolerance is 1/10,000 of the extent of the data frame's spatial domain
		Dim tol As Double
		tol = 0
		
		' Perform the union
		pBGP = New ESRI.ArcGIS.Carto.BasicGeoprocessor
		pOutputFeatClass = pBGP.Union(pDEMTable, False, pBasinTable, False, tol, pFeatClassName)
		
		RemoveSmallPolys = pOutputFeatClass
		
		'Cleanup
		'Remove the layers first
		m_pMap.DeleteLayer(m_pMap.Layer(0))
		m_pMap.DeleteLayer(m_pMap.Layer(0))
		
		pTempDEM = pDEMFeatClass
		pTempBasin = pFeatureClass
		
		If pTempDEM.CanDelete Then
			pTempDEM.Delete()
		End If
		
		If pTempBasin.CanDelete Then
			pTempBasin.Delete()
		End If
		
		'Cleanup
		'UPGRADE_NOTE: Object pAlgebraOp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pAlgebraOp = Nothing
		'UPGRADE_NOTE: Object pDEMOutline may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pDEMOutline = Nothing
		'UPGRADE_NOTE: Object pDEMFeatClass may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pDEMFeatClass = Nothing
		'UPGRADE_NOTE: Object pRasterConvert may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRasterConvert = Nothing
		'UPGRADE_NOTE: Object pWorkspace may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pWorkspace = Nothing
		'UPGRADE_NOTE: Object pFeature may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFeature = Nothing
		'UPGRADE_NOTE: Object pArea may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pArea = Nothing
		'UPGRADE_NOTE: Object pEditor may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pEditor = Nothing
		'UPGRADE_NOTE: Object pDEMTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pDEMTable = Nothing
		'UPGRADE_NOTE: Object pBasinTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pBasinTable = Nothing
		'UPGRADE_NOTE: Object pFeatClassName may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFeatClassName = Nothing
		'UPGRADE_NOTE: Object pNewWSName may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pNewWSName = Nothing
		'UPGRADE_NOTE: Object pDatasetName may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pDatasetName = Nothing
		'UPGRADE_NOTE: Object pBGP may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pBGP = Nothing
		'UPGRADE_NOTE: Object pExtMgr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pExtMgr = Nothing
		'UPGRADE_NOTE: Object pTempDEM may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pTempDEM = Nothing
		'UPGRADE_NOTE: Object pTempBasin may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pTempBasin = Nothing
		'UPGRADE_NOTE: Object pEnvNew may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pEnvNew = Nothing
		'UPGRADE_NOTE: Object pEditor may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pEditor = Nothing
		'UPGRADE_NOTE: Object pExtension may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pExtension = Nothing
		
		Exit Function
ErrorHandler: 
		HandleError(False, "RemoveSmallPolys " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
		
	End Function
	
	
	
	Private Sub txtWSDelinName_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtWSDelinName.Validating
		Dim Cancel As Boolean = eventArgs.Cancel
		On Error GoTo ErrorHandler
		
		Dim fso As Scripting.FileSystemObject
		fso = CreateObject("Scripting.FileSystemObject")
		
		If fso.FileExists(modUtil.g_nspectPath & "\" & txtWSDelinName.Text) Then
			MsgBox("A scenario named " & txtWSDelinName.Text & " already exists.  Please select another name.", MsgBoxStyle.Critical, "Duplicate Name")
			txtWSDelinName.Focus()
		End If
		
		GoTo EventExitSub
ErrorHandler: 
		HandleError(True, "txtWSDelinName_Validate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
EventExitSub: 
		eventArgs.Cancel = Cancel
	End Sub
	
	Private Function ReturnRasterDataset(ByRef pInComingGeoDataSet As ESRI.ArcGIS.Geodatabase.IGeoDataset, ByRef strName As String, ByRef pWorkspace As ESRI.ArcGIS.Geodatabase.IWorkspace) As ESRI.ArcGIS.Geodatabase.IGeoDataset
		On Error GoTo ErrorHandler
		
		Dim pBands As ESRI.ArcGIS.DataSourcesRaster.IRasterBandCollection
		Dim pBand As ESRI.ArcGIS.DataSourcesRaster.IRasterBand
		Dim pRasterDataset As ESRI.ArcGIS.Geodatabase.IRasterDataset
		Dim pNewGeo As ESRI.ArcGIS.Geodatabase.IGeoDataset
		
		pBands = pInComingGeoDataSet
		pBand = pBands.Item(0)
		pRasterDataset = pBand.RasterDataset
		
		Dim pTempDS As ESRI.ArcGIS.DataSourcesRaster.ITemporaryDataset
		pTempDS = pRasterDataset
		
		MsgBox("Rename Workspace: " & pWorkspace.PathName)
		MsgBox("rename type: " & pWorkspace.Type)
		
		ReturnRasterDataset = pTempDS.MakePermanentAs(strName, pWorkspace, "GRID")
		
		'UPGRADE_NOTE: Object pRasterDataset may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRasterDataset = Nothing
		'UPGRADE_NOTE: Object pBands may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pBands = Nothing
		'UPGRADE_NOTE: Object pBand may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pBand = Nothing
		'UPGRADE_NOTE: Object pInComingGeoDataSet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pInComingGeoDataSet = Nothing
		
		
		
		
		Exit Function
ErrorHandler: 
		HandleError(False, "ReturnRasterDataset " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Function
	
	
	Private Function CalcLengthSlope(ByRef pDEMRaster As ESRI.ArcGIS.Geodatabase.IRaster, ByRef pFlowDirRaster As ESRI.ArcGIS.Geodatabase.IRaster, ByRef pAccumRaster As ESRI.ArcGIS.Geodatabase.IRaster, ByRef pEnv As ESRI.ArcGIS.GeoAnalyst.IRasterAnalysisEnvironment, ByRef strUnits As String, ByRef pWS As ESRI.ArcGIS.Geodatabase.IWorkspace) As String
		
		'From the Delineate watershed, we are garnering the DEM, Flow Direction, flowaccum, environment and units
		'to be used in this function
		'Returns: Name of the LS File for use in the database table
		On Error GoTo ErrHandler
		
		'The rasters                                'Associated Steps
		Dim pDEMOneCell As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 1: 1 Cell Buffer DEM
		Dim pDEMTwoCell As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 2: 2 Cell Buffer DEM
		Dim pFlowDir As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 3: Flow Direction
		Dim pFlowDirBV As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 4: Pre Nibble Null
		Dim pMask As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 4: Mask
		Dim pNibble As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 4: Nibble
		Dim pDownSlope As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 5: Down Slope
		Dim pDownAngle As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 5a: Tweak down slope
		Dim pRelativeSlope As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 6: Relative Slope
		Dim pRelSlopeThreshold As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 7: Relative slope threshold
		Dim pSlopeBreak As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 7a: Slope Break
		Dim pFlowDirBreak As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 8: Flow Direction Break
		Dim pWeight As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 9: Weight GRID
		Dim pFlowLength As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 10: Flow Length
		Dim pFlowLengthFt As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 11: Flow Length to Feet
		Dim pSlopeExp As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 12: Slope Exponent
		Dim pRusleLFactor As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 13: Rusle L Factor
		Dim pRusleSFactor As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 14: Rusle S Factor
		Dim pLSFactor As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 15: LS Factor
		Dim pFinalLS As ESRI.ArcGIS.Geodatabase.IRaster 'STEP 15a: clippage
		Dim strProgTitle As String
		
		'Analysis Environment
		Dim pDEMGeoDS As ESRI.ArcGIS.Geodatabase.IGeoDataset 'GeoDataset to get spat. ref
		Dim pSpatRef As ESRI.ArcGIS.Geometry.ISpatialReference 'Spatial Reference
		
		'Create Map Algebra Operator
		Dim pMapAlgebraOp As ESRI.ArcGIS.SpatialAnalyst.IMapAlgebraOp 'Workhorse
		'String to hold calculations
		Dim strExpression As String
		
		pDEMGeoDS = pDEMRaster
		pSpatRef = pDEMGeoDS.SpatialReference
		
		If Not modUtil.CheckSpatialAnalystLicense Then
			MsgBox("No Spatial Analyst License Available.", MsgBoxStyle.Critical, "Pay Your Licensing Fee")
			CalcLengthSlope = "ERROR"
			Exit Function
		End If
		
		If pDEMRaster Is Nothing Or pFlowDirRaster Is Nothing Then
			MsgBox("CaclLengthSlope Error.")
			CalcLengthSlope = "ERROR"
		End If
		
		'Initialize the Map AlgebraOp, same thing as the Map Calculator
		'All of the following steps use the same methodology.  They take a Raster, bind it
		'to a symbol, in this case a string that represents them in some map calculation.
		'You then simply use the MapAlgebraOp to execute the expression.
		
		'Get a hold of the Spatial Reference of the DEM, and init MapAlgebra
		pMapAlgebraOp = New ESRI.ArcGIS.SpatialAnalyst.RasterMapAlgebraOp
		
		'Set the environment
		pEnv = pMapAlgebraOp
		
		strProgTitle = "Processing the LS GRID..."
		
		'STEP 1: ----------------------------------------------------------------------
		'Buffer the DEM by one cell
		modProgDialog.ProgDialog("Creating one cell buffer...", strProgTitle, 0, 15, 1, (m_App.hwnd))
		
		'UPGRADE_NOTE: Object pDEMOneCell may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pDEMOneCell = Nothing
		
		If modProgDialog.g_boolCancel Then
			
			'UPGRADE_WARNING: Couldn't resolve default property of object pDEMRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pMapAlgebraOp.BindRaster(pDEMRaster, "aml_fdem")
			strExpression = "Con(isnull([aml_fdem]), focalmin([aml_fdem]), [aml_fdem])"
			
			pDEMOneCell = pMapAlgebraOp.Execute(strExpression)
			
			pMapAlgebraOp.UnbindRaster("aml_fdem")
			
		Else
			GoTo ProgCancel
		End If
		
		'END STEP 1: ------------------------------------------------------------------
		
		'STEP 2: ----------------------------------------------------------------------
		'Buffer the DEM buffer by one more cell, that's 2
		modProgDialog.ProgDialog("Creating two cell buffer...", strProgTitle, 0, 15, 2, (m_App.hwnd))
		
		If modProgDialog.g_boolCancel Then
			
			'UPGRADE_WARNING: Couldn't resolve default property of object pDEMOneCell. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pMapAlgebraOp.BindRaster(pDEMOneCell, "dem_b")
			strExpression = "Con(isnull([dem_b]), focalmin([dem_b]), [dem_b])"
			
			pDEMTwoCell = pMapAlgebraOp.Execute(strExpression)
			
			m_strDEM2BName = modUtil.MakePerminentGrid(pDEMTwoCell, (pWS.PathName), "dem2b")
			
			pMapAlgebraOp.UnbindRaster("dem_b")
			
		Else
			GoTo ProgCancel
		End If
		'END STEP 2: ------------------------------------------------------------------
		
		'STEP 3: ----------------------------------------------------------------------
		'Flow Direction
		
		pFlowDir = pFlowDirRaster
		
		'END STEP 3: ------------------------------------------------------------------
		
		
		'STEP 3a: ---------------------------------------------------------------------
		modProgDialog.ProgDialog("Creating mask...", strProgTitle, 0, 15, 3, (m_App.hwnd))
		
		If modProgDialog.g_boolCancel Then
			
			With pMapAlgebraOp
				'UPGRADE_WARNING: Couldn't resolve default property of object pDEMOneCell. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(pDEMOneCell, "mask")
			End With
			
			'strExpression = "con(isnull([mask]),0,1)"
			strExpression = "con([mask] >= 0, 1, 0)"
			
			
			pMask = pMapAlgebraOp.Execute(strExpression)
			
			With pEnv
				.Mask = pMask
				'UPGRADE_WARNING: Couldn't resolve default property of object pEnv.OutWorkspace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.OutWorkspace = pWS
			End With
			
			pEnv = pMapAlgebraOp
		Else
			GoTo ProgCancel
		End If
		
		'STEP 4: ----------------------------------------------------------------------
		'Buffering Flow Direction and do the nibble to fill it in
		'Needed in case there is outflow from the DEM grid.
		'The following algorithms need to access the downslope DEM grid cell.
		'We find this by nibbling the original flow direction grid instead of recalculating
		'flow direction from the nibbled DEM because we want the elevation that is assumed
		'to be downstream of the edge cell.  Using flow direction on the buffered DEM
		'may not give that same result.
		modProgDialog.ProgDialog("Buffering slope direction...", strProgTitle, 0, 15, 4, (m_App.hwnd))
		
		If modProgDialog.g_boolCancel Then
			
			With pMapAlgebraOp
				'UPGRADE_WARNING: Couldn't resolve default property of object pFlowDir. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(pFlowDir, "fdr_b")
			End With
			strExpression = "con(isnull([fdr_b]),0,[fdr_b])"
			
			pFlowDirBV = pMapAlgebraOp.Execute(strExpression)
			
			pMapAlgebraOp.UnbindRaster("fdr_b")
			
			'Nibble
			'UPGRADE_WARNING: Couldn't resolve default property of object pFlowDirBV. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pMapAlgebraOp.BindRaster(pFlowDirBV, "fdr_bv")
			'UPGRADE_WARNING: Couldn't resolve default property of object pMask. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pMapAlgebraOp.BindRaster(pMask, "waia_reg")
			strExpression = "nibble([fdr_bv],[waia_reg], dataonly)"
			
			pNibble = pMapAlgebraOp.Execute(strExpression)
			
			'Get nibble's path for use in the database
			m_strNibbleName = modUtil.MakePerminentGrid(pNibble, (pWS.PathName), "nibble")
			
			With pMapAlgebraOp
				.UnbindRaster("fdr_bv")
				.UnbindRaster("waia_reg")
			End With
		Else
			GoTo ProgCancel
		End If
		'END STEP 4: ------------------------------------------------------------------
		
		'STEP 5: ----------------------------------------------------------------------
		'Calculate Slope
		'Actually this is calculating the SLOPE, in degrees, not the slope change.
		'Note that ESRI's SLOPE command should not be used here.  That command fits
		'a plane to the 3x3 grid surrounding the central point and assigns the central
		'point the slope of that plane.  The algorithm used here calculates only the slope
		'between the central point and it's immediate downstream neighbor.
		'That is what is needed by RUSLE.
		
		modProgDialog.ProgDialog("Calculating Slope change...", strProgTitle, 0, 15, 5, (m_App.hwnd))
		
		If modProgDialog.g_boolCancel Then
			
			With pMapAlgebraOp
				'UPGRADE_WARNING: Couldn't resolve default property of object pNibble. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(pNibble, "fdrnib")
				'UPGRADE_WARNING: Couldn't resolve default property of object pDEMTwoCell. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(pDEMTwoCell, "dem_2b")
			End With
			strExpression = "Con(([fdrnib] ge 0.5 and [fdrnib] lt 1.5), deg * atan(([dem_2b] - [dem_2b](1,0)) / (" & m_intCellSize & "))," & "Con(([fdrnib] ge 1.5 and [fdrnib] lt 3.0), deg * atan(([dem_2b] - [dem_2b](1,1)) / (" & m_intCellSize & " * 1.4142))," & "Con(([fdrnib] ge 3.0 and [fdrnib] lt 6.0), deg * atan(([dem_2b] - [dem_2b](0,1)) / (" & m_intCellSize & "))," & "Con(([fdrnib] ge 6.0 and [fdrnib] lt 12.0), deg * atan(([dem_2b] - [dem_2b](-1,1)) / (" & m_intCellSize & " * 1.4142))," & "Con(([fdrnib] ge 12.0 and [fdrnib] lt 24.0), deg * atan(([dem_2b] - [dem_2b](-1,0)) / (" & m_intCellSize & "))," & "Con(([fdrnib] ge 24.0 and [fdrnib] lt 48.0), deg * atan(([dem_2b] - [dem_2b](-1,-1)) / (" & m_intCellSize & " * 1.4142))," & "Con(([fdrnib] ge 48.0 and [fdrnib] lt 96.0), deg * atan(([dem_2b] - [dem_2b](0,-1)) / (" & m_intCellSize & "))," & "Con(([fdrnib] ge 96.0 and [fdrnib] lt 192.0), deg * atan(([dem_2b] - [dem_2b](1,-1)) / (" & m_intCellSize & " * 1.4142))," & "Con(([fdrnib] ge 192.0 and [fdrnib] le 255.0), deg * atan(([dem_2b] - [dem_2b](1,0)) / (" & m_intCellSize & "))," & "0.1 )))))))))"
			
			pDownSlope = pMapAlgebraOp.Execute(strExpression)
			'Cleanup
			With pMapAlgebraOp
				.UnbindRaster("fdrnib")
				.UnbindRaster("dem_2b")
			End With
		Else
			GoTo ProgCancel
		End If
		'END STEP 5 ---------------------------------------------------------------------
		
		'STEP 5a: -----------------------------------------------------------------------
		'Tweak slope where it equals 0 to 0.1
		'UPGRADE_WARNING: Couldn't resolve default property of object pDownSlope. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pMapAlgebraOp.BindRaster(pDownSlope, "dwnsltmp")
		strExpression = "Con([dwnsltmp] le 0, 0.1,[dwnsltmp])"
		
		pDownAngle = pMapAlgebraOp.Execute(strExpression)
		
		pMapAlgebraOp.UnbindRaster("dwnsltmp")
		
		'END STEP 5a: -------------------------------------------------------------------
		
		'STEP 6: ------------------------------------------------------------------------
		'Relative Slope Change
		modProgDialog.ProgDialog("Calculating Relative Slope Change...", strProgTitle, 0, 15, 6, (m_App.hwnd))
		
		If modProgDialog.g_boolCancel Then
			With pMapAlgebraOp
				'UPGRADE_WARNING: Couldn't resolve default property of object pDownAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(pDownAngle, "dwnslangle")
				'UPGRADE_WARNING: Couldn't resolve default property of object pNibble. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(pNibble, "fdrnib")
			End With
			strExpression = "Con([fdrnib] == 1, ([dwnslangle] - [dwnslangle](1,0)) / [dwnslangle]," & "Con([fdrnib] == 2, ([dwnslangle] - [dwnslangle](1,1)) / [dwnslangle]," & "Con([fdrnib] == 4, ([dwnslangle] - [dwnslangle](0,1)) / [dwnslangle]," & "Con([fdrnib] == 8, ([dwnslangle] - [dwnslangle](-1,1)) / [dwnslangle]," & "Con([fdrnib] == 16, ([dwnslangle] - [dwnslangle](-1,0)) / [dwnslangle]," & "Con([fdrnib] == 32, ([dwnslangle] - [dwnslangle](-1,-1)) / [dwnslangle]," & "Con([fdrnib] == 64, ([dwnslangle] - [dwnslangle](0,-1)) / [dwnslangle]," & "Con([fdrnib] == 128, ([dwnslangle] - [dwnslangle](1,-1)) / [dwnslangle]," & "Con([fdrnib] == 255, ([dwnslangle] - [dwnslangle](1,0)) / [dwnslangle]," & "0.1 )))))))))"
			
			pRelativeSlope = pMapAlgebraOp.Execute(strExpression)
			
			With pMapAlgebraOp
				.UnbindRaster("dwnslangle")
				.UnbindRaster("fdrnib")
			End With
		Else
			GoTo ProgCancel
		End If
		'END STEP 6 -----------------------------------------------------------------------
		
		'STEP 7: --------------------------------------------------------------------------
		'Identify breakpoints: relative difference where slope angle exceeds threshold values
		modProgDialog.ProgDialog("Identifying breakpoints...", strProgTitle, 0, 15, 7, (m_App.hwnd))
		
		If modProgDialog.g_boolCancel Then
			
			'UPGRADE_WARNING: Couldn't resolve default property of object pDownAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pMapAlgebraOp.BindRaster(pDownAngle, "dwnslangle")
			strExpression = "Con(([dwnslangle] gt 2.86240), 0.5, Con(([dwnslangle] le 2.86240), 0.7, 0.0 ))"
			
			pRelSlopeThreshold = pMapAlgebraOp.Execute(strExpression)
			
			pMapAlgebraOp.UnbindRaster("dwnslangle")
			'END STEP 7 -----------------------------------------------------------------------
			
			'STEP 7a: -------------------------------------------------------------------------
			'Slope break
			With pMapAlgebraOp
				'UPGRADE_WARNING: Couldn't resolve default property of object pRelSlopeThreshold. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(pRelSlopeThreshold, "threshold")
				'UPGRADE_WARNING: Couldn't resolve default property of object pRelativeSlope. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(pRelativeSlope, "delslprel")
			End With
			strExpression = "Con(([delslprel] gt [threshold]), 1, 0 )"
			
			pSlopeBreak = pMapAlgebraOp.Execute(strExpression)
			
			With pMapAlgebraOp
				.UnbindRaster("threshold")
				.UnbindRaster("delslprel")
			End With
		Else
			GoTo ProgCancel
		End If
		'END STEP 7a -------------------------------------------------------------------------
		
		'STEP 8 -------------------------------------------------------------------------------
		'Create Modified Flow Direction GRID
		modProgDialog.ProgDialog("Creating modified flow direction GRID...", strProgTitle, 0, 15, 8, (m_App.hwnd))
		
		If modProgDialog.g_boolCancel Then
			
			With pMapAlgebraOp
				'UPGRADE_WARNING: Couldn't resolve default property of object pSlopeBreak. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(pSlopeBreak, "slopebreak")
				'UPGRADE_WARNING: Couldn't resolve default property of object pFlowDir. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(pFlowDir, "fdr")
			End With
			strExpression = "Con([slopebreak] eq 0, [fdr], 0)"
			
			pFlowDirBreak = pMapAlgebraOp.Execute(strExpression)
			
			With pMapAlgebraOp
				.UnbindRaster("slopebreak")
				.UnbindRaster("fdr")
			End With
		Else
			GoTo ProgCancel
		End If
		'END STEP 8: ----------------------------------------------------------------------
		
		'STEP 9: --------------------------------------------------------------------------
		'Create weight grid
		modProgDialog.ProgDialog("Creating weight GRID...", strProgTitle, 0, 15, 9, (m_App.hwnd))
		
		If modProgDialog.g_boolCancel Then
			
			
			'Dave's comments:
			'This is an error.  In the original AML code, it is needed to correctly account for
			'diagonal flow in the flow length calculation of step 10.  However, ArcMap already
			'makes this correction.  There is another weighting function needed, however, to be
			'consistent with the procedure used in the original AML code. That is what should replace this con.
			
			'Removed 12/19/07
			'pMapAlgebraOp.BindRaster pFlowDir, "fdr"
			'strExpression = "Con([fdr] eq 2, 1.41421," & _
			''    "Con([fdr] eq 8, 1.41421," & _
			''    "Con([fdr] eq 32, 1.41421," & _
			''    "Con([fdr] eq 128, 1.41421," & _
			''    "1.0))))"
			'End Remove
			
			'Added 12/19/07
			'pAccumRaster is passed over from delin watershed
			'UPGRADE_WARNING: Couldn't resolve default property of object pAccumRaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pMapAlgebraOp.BindRaster(pAccumRaster, "flowacc")
			strExpression = "Con([flowacc] eq 0, 0.5,1.0)"
			
			pWeight = pMapAlgebraOp.Execute(strExpression)
			
			pMapAlgebraOp.UnbindRaster("flowacc")
			'End Added
		Else
			GoTo ProgCancel
		End If
		'END STEP 9 -------------------------------------------------------------------------
		
		'STEP 10: ---------------------------------------------------------------------------
		'Flow Length GRID
		modProgDialog.ProgDialog("Creating flow length GRID...", strProgTitle, 0, 15, 10, (m_App.hwnd))
		
		If modProgDialog.g_boolCancel Then
			
			With pMapAlgebraOp
				'UPGRADE_WARNING: Couldn't resolve default property of object pFlowDirBreak. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(pFlowDirBreak, "fdrbrk")
				'UPGRADE_WARNING: Couldn't resolve default property of object pWeight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(pWeight, "weight")
			End With
			strExpression = "FlowLength([fdrbrk], [weight], UPSTREAM)"
			
			pFlowLength = pMapAlgebraOp.Execute(strExpression)
			
			With pMapAlgebraOp
				.UnbindRaster("fdrbrk")
				.UnbindRaster("weight")
			End With
		Else
			GoTo ProgCancel
		End If
		'END STEP 10: -----------------------------------------------------------------------
		
		'STEP 11: ---------------------------------------------------------------------------
		'Convert Meters To Feet
		'TODO: Check measure units, won't have to do if already in Feet
		modProgDialog.ProgDialog("Checking measurement units...", strProgTitle, 0, 15, 11, (m_App.hwnd))
		
		If modProgDialog.g_boolCancel Then
			'UPGRADE_WARNING: Couldn't resolve default property of object pFlowLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pMapAlgebraOp.BindRaster(pFlowLength, "flowlen")
			strExpression = "[flowlen] / 0.3048"
			
			pFlowLengthFt = pMapAlgebraOp.Execute(strExpression)
			
			pMapAlgebraOp.UnbindRaster("flowlen")
		Else
			GoTo ProgCancel
		End If
		'END STEP 11: -----------------------------------------------------------------------
		
		'STEP 12: ---------------------------------------------------------------------------
		'Calculate the slope length exponent value 'M'
		modProgDialog.ProgDialog("Calculating slope length exponent value 'M'...", strProgTitle, 0, 15, 12, (m_App.hwnd))
		
		If modProgDialog.g_boolCancel Then
			
			'UPGRADE_WARNING: Couldn't resolve default property of object pDownAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pMapAlgebraOp.BindRaster(pDownAngle, "dwnslangle")
			strExpression = "Con(([dwnslangle] le 0.1), 0.01," & "Con(([dwnslangle] gt 0.1 and [dwnslangle] lt 0.2), 0.02," & "Con(([dwnslangle] ge 0.2 and [dwnslangle] lt 0.4), 0.04," & "Con(([dwnslangle] ge 0.4 and [dwnslangle] lt 0.85), 0.08," & "Con(([dwnslangle] ge 0.85 and [dwnslangle] lt 1.4), 0.14," & "Con(([dwnslangle] ge 1.4 and [dwnslangle] lt 2.0), 0.18," & "Con(([dwnslangle] ge 2.0 and [dwnslangle] lt 2.6), 0.22," & "Con(([dwnslangle] ge 2.6 and [dwnslangle] lt 3.1), 0.25," & "Con(([dwnslangle] ge 3.1 and [dwnslangle] lt 3.7), 0.28," & "Con(([dwnslangle] ge 3.7 and [dwnslangle] lt 5.2), 0.32," & "Con(([dwnslangle] ge 5.2 and [dwnslangle] lt 6.3), 0.35," & "Con(([dwnslangle] ge 6.3 and [dwnslangle] lt 7.4), 0.37," & "Con(([dwnslangle] ge 7.4 and [dwnslangle] lt 8.6), 0.40," & "Con(([dwnslangle] ge 8.6 and [dwnslangle] lt 10.3), 0.41," & "Con(([dwnslangle] ge 10.3 and [dwnslangle] lt 12.9), 0.44," & "Con(([dwnslangle] ge 12.9 and [dwnslangle] lt 15.7), 0.47," & "Con(([dwnslangle] ge 15.7 and [dwnslangle] lt 20.0), 0.49," & "Con(([dwnslangle] ge 20.0 and [dwnslangle] lt 25.8), 0.52," & "Con(([dwnslangle] ge 25.8 and [dwnslangle] lt 31.5), 0.54," & "Con(([dwnslangle] ge 31.5 and [dwnslangle] lt 37.2), 0.55," & "0.56))))))))))))))))))))"
			
			pSlopeExp = pMapAlgebraOp.Execute(strExpression)
			
			pMapAlgebraOp.UnbindRaster("dwnslangle")
		Else
			GoTo ProgCancel
		End If
		'END STEP 12: ------------------------------------------------------------------------
		
		'STEP 13: ----------------------------------------------------------------------------
		'Calculate the L-Factor
		modProgDialog.ProgDialog("Calculating the L factor...", strProgTitle, 0, 15, 13, (m_App.hwnd))
		
		If modProgDialog.g_boolCancel Then
			
			With pMapAlgebraOp
				'UPGRADE_WARNING: Couldn't resolve default property of object pFlowLengthFt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(pFlowLengthFt, "flowlenft")
				'UPGRADE_WARNING: Couldn't resolve default property of object pSlopeExp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(pSlopeExp, "new_slpexp")
			End With
			'This non-dimensionalizes flowlenft and, after raising the result to the power in new_slpexp, we have the L factor.
			strExpression = "Pow(([flowlenft] / 72.6), [new_slpexp])"
			
			pRusleLFactor = pMapAlgebraOp.Execute(strExpression)
			'AddRasterLayer Application, pRusleLFactor, "RussleL"
			
			With pMapAlgebraOp
				.UnbindRaster("flowlenft")
				.UnbindRaster("new_slpexp")
			End With
		Else
			GoTo ProgCancel
		End If
		'END STEP 13: ------------------------------------------------------------------------
		
		'STEP 14: ----------------------------------------------------------------------------
		'Calculate the S-Factor
		modProgDialog.ProgDialog("Creating flow length GRID...", strProgTitle, 0, 15, 14, (m_App.hwnd))
		
		If modProgDialog.g_boolCancel Then
			
			'Here is the calculation for S, which is not, actually the slope,
			'but IS a function of the slope between a cell and its immediate downstream neighbor.
			'UPGRADE_WARNING: Couldn't resolve default property of object pDownAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pMapAlgebraOp.BindRaster(pDownAngle, "dwnslangle")
			strExpression = "Con([dwnslangle] ge 5.1428, 16.8 * (sin(([dwnslangle] - 0.5) div deg))," & "10.8 * (sin(([dwnslangle] + 0.03) div deg)))"
			
			pRusleSFactor = pMapAlgebraOp.Execute(strExpression)
			
			pMapAlgebraOp.UnbindRaster("dwnslangle")
		Else
			GoTo ProgCancel
		End If
		'END STEP 14: -------------------------------------------------------------------------
		
		'STEP 15: ----------------------------------------------------------------------------
		'Calculate the LS Factor
		modProgDialog.ProgDialog("Calculating the LS Factor...", strProgTitle, 0, 15, 15, (m_App.hwnd))
		
		If modProgDialog.g_boolCancel Then
			
			'quick math to clip this bugger
			With pMapAlgebraOp
				'UPGRADE_WARNING: Couldn't resolve default property of object pRusleSFactor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(pRusleSFactor, "Sfactor")
				'UPGRADE_WARNING: Couldn't resolve default property of object pRusleLFactor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(pRusleLFactor, "Lfactor")
			End With
			
			strExpression = "[Sfactor] * [Lfactor]"
			
			pLSFactor = pMapAlgebraOp.Execute(strExpression)
			
			With pMapAlgebraOp
				.UnbindRaster("Sfactor")
				.UnbindRaster("Lfactor")
			End With
			'END STEP 15: -------------------------------------------------------------------------
			
			'STEP 15a: ----------------------------------------------------------------------------
			With pMapAlgebraOp
				'UPGRADE_WARNING: Couldn't resolve default property of object pLSFactor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(pLSFactor, "LSFactor")
				'UPGRADE_WARNING: Couldn't resolve default property of object pMask. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.BindRaster(pMask, "Mask")
			End With
			
			strExpression = "[LSFactor] * [Mask]"
			
			pFinalLS = pMapAlgebraOp.Execute(strExpression)
			
			With pMapAlgebraOp
				.UnbindRaster("LSFactor")
				.UnbindRaster("Mask")
			End With
		Else
			GoTo ProgCancel
		End If
		
		CalcLengthSlope = modUtil.MakePerminentGrid(pLSFactor, (pWS.PathName), "LSGrid")
		
		modProgDialog.KillDialog()
		
		'Cleanup
		'UPGRADE_NOTE: Object pDEMOneCell may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pDEMOneCell = Nothing 'STEP 1: 1 Cell Buffer DEM
		'UPGRADE_NOTE: Object pDEMTwoCell may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pDEMTwoCell = Nothing 'STEP 2: 2 Cell Buffer DEM
		'UPGRADE_NOTE: Object pFlowDir may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFlowDir = Nothing 'STEP 3: Flow Direction
		'UPGRADE_NOTE: Object pFlowDirBV may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFlowDirBV = Nothing 'STEP 4: Pre Nibble Null
		'UPGRADE_NOTE: Object pMask may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pMask = Nothing 'STEP 4: Mask
		'UPGRADE_NOTE: Object pNibble may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pNibble = Nothing 'STEP 4: Nibble
		'UPGRADE_NOTE: Object pDownSlope may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pDownSlope = Nothing 'STEP 5: Down Slope
		'UPGRADE_NOTE: Object pDownAngle may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pDownAngle = Nothing 'STEP 5a: Tweak down slope
		'UPGRADE_NOTE: Object pRelativeSlope may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRelativeSlope = Nothing 'STEP 6: Relative Slope
		'UPGRADE_NOTE: Object pRelSlopeThreshold may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRelSlopeThreshold = Nothing 'STEP 7: Relative slope threshold
		'UPGRADE_NOTE: Object pSlopeBreak may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pSlopeBreak = Nothing 'STEP 7a: Slope Break
		'UPGRADE_NOTE: Object pFlowDirBreak may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFlowDirBreak = Nothing 'STEP 8: Flow Direction Break
		'UPGRADE_NOTE: Object pWeight may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pWeight = Nothing 'STEP 9: Weight GRID
		'UPGRADE_NOTE: Object pFlowLength may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFlowLength = Nothing 'STEP 10: Flow Length
		'UPGRADE_NOTE: Object pFlowLengthFt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFlowLengthFt = Nothing 'STEP 11: Flow Length to Feet
		'UPGRADE_NOTE: Object pSlopeExp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pSlopeExp = Nothing 'STEP 12: Slope Exponent
		'UPGRADE_NOTE: Object pRusleLFactor may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRusleLFactor = Nothing 'STEP 13: Rusle L Factor
		'UPGRADE_NOTE: Object pRusleSFactor may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRusleSFactor = Nothing 'STEP 14: Rusle S Factor
		'UPGRADE_NOTE: Object pFinalLS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFinalLS = Nothing 'STEP 15: Final LS Factor
		'UPGRADE_NOTE: Object pLSFactor may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pLSFactor = Nothing 'STEP 15a: Masked LS Factor
		
		Exit Function
ProgCancel: 
		modProgDialog.KillDialog()
		MsgBox("The LS GRID calculation has been stopped by the user.  Changes have been discarded.", MsgBoxStyle.Critical, "Process Stopped")
		CalcLengthSlope = "ERROR"
		
		Exit Function
ErrHandler: 
		MsgBox(Err.Number & ": " & Err.Description)
		
	End Function
End Class