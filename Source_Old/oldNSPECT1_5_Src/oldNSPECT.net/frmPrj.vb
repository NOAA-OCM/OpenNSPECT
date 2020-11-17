Option Strict Off
Option Explicit On
Friend Class frmPrj
	Inherits System.Windows.Forms.Form
	' *************************************************************************************
	' *  Perot Systems Government Services
	' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
	' *  frmPrj
	' *************************************************************************************
	' *  Description:  The big one.  The main form called from Analysis... menu choice.
	' *
	' *
	' *  Called By:  frmPrj
	' *************************************************************************************
	
	
	Private m_strFileName As String 'Name of Open doc
	Private m_XMLPrjParams As clsXMLPrjFile 'xml doc that holds inputs
	Private m_bolFirstLoad As Boolean 'Is initial Load event
	Private m_booNew As Boolean 'New
	Private m_booExists As Boolean 'Has file been saved
	Private m_strOpenFileName As String 'String to hold open file name, if they change name, prompt to 'save as'
	Private m_strWorkspace As String 'String holding workspace, set it
	Private m_booAnnualPrecip As Boolean 'Is the precip scenario annual, if so = TRUE
	
	Private m_intCount As Short
	Private m_intMgmtCount As Short 'Count for management scenarios
	Private m_intLUCount As Short 'Count for Land Use grid
	Private m_intPollCount As Short 'Count for Pollutants in Pollutants tab
	Private m_intPollRow As Short 'Row for Pollutant grid
	Private m_intPollCol As Short 'Col for Pollutant grid
	Private m_intLCRow As Short 'Row for LCChange Grid
	Private m_intLCCol As Short 'Col for LCChange Grid
	Private m_intLURow As Short 'Row for mgmt scenarios
	Private m_intLUCol As Short 'col for mgmt scenarios
	
	Private m_strType As String 'Flag for deletion/creation of checkboxes
	Private m_strPrecipFile As String
	Private m_strWShed As String 'String
	
	'Font DPI API
	Private Declare Function GetDC Lib "user32" (ByVal hwnd As Integer) As Integer
	Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Integer, ByVal hDC As Integer) As Integer
	Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Integer, ByVal nIndex As Integer) As Integer
	
	' Win 32 Constant Declarations
	Private Const LOGPIXELSX As Short = 88 'Logical pixels/inch in X
	
	'ArcMap stuff
	Private m_pMap As ESRI.ArcGIS.Carto.IMap 'Ref to ArcMap.FocusMap
	Private m_pMxDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
	Private m_ParentHWND As Integer 'Set this to get correct parenting of Error handler forms
	
	'Public stuff
	Public m_App As ESRI.ArcGIS.Framework.IApplication 'Ref to ArcMap
	Public WithEvents m_pActiveViewEvents As ESRI.ArcGIS.Carto.Map 'Active View Events for tracking added data layers
	
	' The code file for "cBrowseForFolder" is copied from vbAccelerator.
	' Link: http://vbaccelerator.com/zip.asp?id=5160
	'Public WithEvents dlgBrowser As cBrowseForFolder 'Browse for folder add in: sets output directory
	
	' Constant used by the Error handler function - DO NOT REMOVE
	Const c_sModuleFileName As String = "frmPrj.frm"
	
	' The following is a workaround for modeless VB forms.
	' See: http://resources.esri.com/help/9.3/arcgisdesktop/com/COM/VB6/ModelessVBDialogs.htm
	Private m_Frame As ESRI.ArcGIS.Framework.IModelessFrame
	Public Function Frame() As ESRI.ArcGIS.Framework.IModelessFrame
		If m_Frame Is Nothing Then
			m_Frame = New ESRI.ArcGIS.Framework.ModelessFrame
			m_Frame.Create(Me)
			m_Frame.Caption = Me.Text
		End If
		Frame = m_Frame
	End Function
	
	Public Sub init(ByVal pApp As ESRI.ArcGIS.Framework.IApplication)
		On Error GoTo ErrorHandler
		
		'Called before form loads, means to initialize pMap and the ActiveView events
		
		m_App = pApp
		m_pMxDoc = m_App.Document
		m_pMap = m_pMxDoc.FocusMap
		m_pActiveViewEvents = m_pMap 'Set up to catch add layer event
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "init " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	'UPGRADE_WARNING: Event cboAreaLayer.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboAreaLayer_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboAreaLayer.SelectedIndexChanged
		'The floating combobox used in the Management scenarios, is populated with all polygon layers in the current map
		On Error GoTo ErrorHandler
		
		grdLCChanges.set_TextMatrix(m_intLCRow, m_intLCCol, cboAreaLayer.Text)
		cboAreaLayer.Visible = False
		
		grdLCChanges.set_TextMatrix(0, 0, "")
		
		Exit Sub
		
ErrorHandler: 
		HandleError(True, "cboAreaLayer_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	'UPGRADE_WARNING: Event cboClass.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboClass_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboClass.SelectedIndexChanged
		'The floating comboxbox used in the Management Scenarios grid, is populated with all polygon layers
		On Error GoTo ErrorHandler
		
		grdLCChanges.set_TextMatrix(m_intLCRow, m_intLCCol, cboClass.Text)
		cboClass.Visible = False
		
		Exit Sub
		
ErrorHandler: 
		HandleError(True, "cboClass_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	'UPGRADE_WARNING: Event cboCoeff.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboCoeff_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboCoeff.SelectedIndexChanged
		'The floating combobox used in the Pollutant tab for Coefficient
		On Error GoTo ErrorHandler
		
		grdCoeffs.set_TextMatrix(m_intPollRow, m_intPollCol, cboCoeff.Text)
		cboCoeff.Visible = False
		
		If cboCoeff.SelectedIndex = 4 Then
			
			g_intCoeffRow = m_intPollRow 'Global set up to hold what row we're on
			g_strCoeffCalc = grdCoeffs.get_TextMatrix(g_intCoeffRow, 6)
			
			frmPrjCalc.init(m_App)
			VB6.ShowForm(frmPrjCalc, VB6.FormShowConstants.Modal, Me)
			
		End If
		
		Exit Sub
		
ErrorHandler: 
		HandleError(True, "cboCoeff_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	'UPGRADE_WARNING: Event cboCoeffSet.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboCoeffSet_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboCoeffSet.SelectedIndexChanged
		'The floating combobox used in the Pollutant tab for Coefficient Set
		On Error GoTo ErrorHandler
		
		'Set the text of column 3 to the selected cbo Text
		grdCoeffs.set_TextMatrix(m_intPollRow, m_intPollCol, cboCoeffSet.Text)
		
		'Added 8/17/04; set Type to 'Type 1' as default
		grdCoeffs.set_TextMatrix(m_intPollRow, m_intPollCol + 1, "Type 1")
		
		cboCoeffSet.Visible = False
		
		Exit Sub
		
ErrorHandler: 
		HandleError(True, "cboCoeffSet_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	'UPGRADE_WARNING: Event cboLCLayer.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboLCLayer_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboLCLayer.SelectedIndexChanged
		'Landclass layer combobox
		On Error GoTo ErrorHandler
		
		cboLCUnits.SelectedIndex = modUtil.GetRasterDistanceUnits((cboLCLayer.Text), m_App)
		
		Exit Sub
		
ErrorHandler: 
		HandleError(True, "cboLCLayer_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	'UPGRADE_WARNING: Event cboLCType.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboLCType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboLCType.SelectedIndexChanged
		'Landclass type
		
		On Error GoTo ErrorHandler
		
		FillCboLCCLass()
		
		Exit Sub
		
ErrorHandler: 
		HandleError(True, "cboLCType_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	'UPGRADE_WARNING: Event cboPrecipScen.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboPrecipScen_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboPrecipScen.SelectedIndexChanged
		'Combobox for Precip Scenarios
		On Error GoTo ErrorHandler
		
		'Have to change Erosion tab based on Annual/Event driven rain event
		Dim rsEvent As New ADODB.Recordset
		Dim strEvent As String
		
		'If define, then open new window for new definition, else select from database
		If cboPrecipScen.Text = "New precipitation scenario..." Then
			'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
			Load(frmNewPrecip)
			VB6.ShowForm(frmNewPrecip, VB6.FormShowConstants.Modal, Me)
		Else
			strEvent = "SELECT * FROM PRECIPSCENARIO WHERE NAME LIKE '" & cboPrecipScen.Text & "'"
			rsEvent.Open(strEvent, g_ADOConn, ADODB.CursorTypeEnum.adOpenStatic)
			
			Select Case rsEvent.Fields("Type").Value
				Case 0 'Annual
					frmSDR.Visible = True
					frameRainFall.Visible = True
					chkCalcErosion.Text = "Calculate Erosion for Annual Type Precipitation Scenario"
					m_booAnnualPrecip = True 'Set flag
				Case 1 'Event
					frmSDR.Visible = False
					frameRainFall.Visible = False
					chkCalcErosion.Text = "Calculate Erosion for Event Type Precipitation Scenario"
					m_booAnnualPrecip = False 'Set flag
			End Select
			
			m_strPrecipFile = rsEvent.Fields("PrecipFileName").Value
			
			rsEvent.Close()
			'UPGRADE_NOTE: Object rsEvent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsEvent = Nothing
			
		End If
		
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "cboPrecipScen_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	'UPGRADE_WARNING: Event cboSoilsLayer.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboSoilsLayer_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboSoilsLayer.SelectedIndexChanged
		'Soils layer combobox
		On Error GoTo ErrorHandler
		
		
		Dim rsSoils As New ADODB.Recordset
		
		Dim strSelect As String
		strSelect = "SELECT * FROM Soils WHERE NAME LIKE '" & cboSoilsLayer.Text & "'"
		
		rsSoils.Open(strSelect, g_ADOConn, ADODB.CursorTypeEnum.adOpenStatic)
		
		lblKFactor.Text = rsSoils.Fields("SoilsKFileName").Value
		lblSoilsHyd.Text = rsSoils.Fields("SoilsFileName").Value
		
		'clean
		rsSoils.Close()
		'UPGRADE_NOTE: Object rsSoils may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsSoils = Nothing
		
		
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "cboSoilsLayer_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	'UPGRADE_WARNING: Event cboWQStd.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboWQStd_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboWQStd.SelectedIndexChanged
		On Error GoTo ErrorHandler
		
		Dim i As Short
		Dim j As Short
		
		If cboWQStd.Text = "New water quality standard..." Then
			'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
			Load(frmAddWQStd)
			VB6.ShowForm(frmAddWQStd, VB6.FormShowConstants.Modal, Me)
		Else
			Timer1.Interval = 10
			Timer1.Enabled = True
			PopPollutants()
			
		End If
		
		For i = 1 To grdCoeffs.Rows - 1
			grdCoeffs.set_TextMatrix(i, 3, "")
			grdCoeffs.set_TextMatrix(i, 4, "")
			
		Next i
		
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "cboWQStd_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	'UPGRADE_WARNING: Event cboWSDelin.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboWSDelin_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboWSDelin.SelectedIndexChanged
		On Error GoTo ErrorHandler
		
		
		If cboWSDelin.Text = "New watershed delineation..." Then
			
			g_boolNewWShed = True
			frmNewWSDelin.m_booProject = True
			frmNewWSDelin.init(m_App)
			VB6.ShowForm(frmNewWSDelin, VB6.FormShowConstants.Modal, Me)
			
		End If
		
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "cboWSDelin_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	'UPGRADE_WARNING: Event chkCalcErosion.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkCalcErosion_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkCalcErosion.CheckStateChanged
		On Error GoTo ErrorHandler
		
		frameRainFall.Enabled = chkCalcErosion.CheckState
		cboErodFactor.Enabled = chkCalcErosion.CheckState
		lblErodFactor.Enabled = chkCalcErosion.CheckState
		optUseGRID.Enabled = chkCalcErosion.CheckState
		optUseValue.Enabled = True 'chkCalcErosion.Value
		lblKFactor.Visible = chkCalcErosion.CheckState
		Label7.Visible = chkCalcErosion.CheckState
		
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "chkCalcErosion_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	'UPGRADE_WARNING: Event chkIgnore.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkIgnore_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkIgnore.CheckStateChanged
		Dim Index As Short = chkIgnore.GetIndex(eventSender)
		On Error GoTo ErrorHandler
		
		'Ignore column
		grdCoeffs.set_TextMatrix(Index, 1, chkIgnore(Index).CheckState)
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "chkIgnore_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	'UPGRADE_WARNING: Event chkIgnoreLU.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkIgnoreLU_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkIgnoreLU.CheckStateChanged
		Dim Index As Short = chkIgnoreLU.GetIndex(eventSender)
		On Error GoTo ErrorHandler
		
		'Ignore column value
		grdLU.set_TextMatrix(Index, 1, chkIgnoreLU(Index).CheckState)
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "chkIgnoreLU_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	'UPGRADE_WARNING: Event chkIgnoreMgmt.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkIgnoreMgmt_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkIgnoreMgmt.CheckStateChanged
		Dim Index As Short = chkIgnoreMgmt.GetIndex(eventSender)
		On Error GoTo ErrorHandler
		
		'Ignore column value
		grdLCChanges.set_TextMatrix(Index, 1, chkIgnoreMgmt(Index).CheckState)
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "chkIgnoreMgmt_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	'UPGRADE_WARNING: Event chkSDR.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkSDR_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkSDR.CheckStateChanged
		If chkSDR.CheckState = 1 Then
			txtSDRGRID.Enabled = True
			cmdOpenSDR.Enabled = True
		Else
			txtSDRGRID.Enabled = False
			cmdOpenSDR.Enabled = False
		End If
	End Sub
	
	'UPGRADE_WARNING: Event chkSelectedPolys.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkSelectedPolys_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkSelectedPolys.CheckStateChanged
		On Error GoTo ErrorHandler
		
		If chkSelectedPolys.CheckState = 1 Then
			cboSelectPoly.Enabled = True
			lblLayer.Enabled = True
		Else
			cboSelectPoly.Enabled = False
			lblLayer.Enabled = False
		End If
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "chkSelectedPolys_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	Private Sub cmdOpenSDR_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOpenSDR.Click
		
		Dim pRasterDataset As ESRI.ArcGIS.Geodatabase.IRasterDataset
		
		On Error GoTo ErrHandler
		pRasterDataset = AddInputFromGxBrowserText(txtSDRGRID, "Choose SDR GRID", Me, 0)
		
		Exit Sub
ErrHandler: 
		MsgBox("There was an error opening the selected file.", MsgBoxStyle.Critical, "Error Opening File")
	End Sub
	
	Private Sub cmdRun_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdRun.Click
		On Error GoTo ErrorHandler
		
		Dim rsWaterShed As New ADODB.Recordset 'RS to get WaterShed information
		Dim strWaterShed As String 'Connection string
		Dim rsPrecip As New ADODB.Recordset 'RS to get Precip info
		Dim strPrecip As String 'Connection String
		Dim pTempRaster As ESRI.ArcGIS.Geodatabase.IRaster 'Temp raster used to delete all globals at completion
		Dim pSelectedPolyLayer As ESRI.ArcGIS.Carto.ILayer 'Selected polygon layer.
		Dim lngWShedLayerIndex As Integer 'Watershed layer index
		
		Dim booLUItems As Boolean 'Are there Landuse Scenarios???
		Dim dictPollutants As New Scripting.Dictionary 'Dict to hold all pollutants
		Dim i As Integer
		Dim strProjectInfo As String 'String that will hold contents of prj file for inclusion in metatdata
		
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		
		'STEP 1: Save file, populate xml params: -------------------------------------------------------------------------
		If Not SaveXMLFile Then
			'UPGRADE_ISSUE: Unable to determine which constant to upgrade vbNormal to. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B3B44E51-B5F1-4FD7-AA29-CAD31B71F487"'
			'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			System.Windows.Forms.Cursor.Current = vbNormal
			Exit Sub
		End If
		
		'Init your global dictionary to hold the metadata records as well as the global xml prj file
		g_dicMetadata = New Scripting.Dictionary
		g_clsXMLPrjFile = m_XMLPrjParams
		'END STEP 1: -----------------------------------------------------------------------------------------------------
		
		
		'STEP 2: Identify if local effects are being used : --------------------------------------------------------------
		'Local Effects Global
		If m_XMLPrjParams.intLocalEffects = 1 Then
			g_booLocalEffects = True
		Else
			g_booLocalEffects = False
		End If
		'END STEP 2: -----------------------------------------------------------------------------------------------------
		
		'STEP 3: Find out if user is making use of only the selected Sheds -----------------------------------------------
		'Selected Sheds only
		If m_XMLPrjParams.intSelectedPolys = 1 Then
			g_booSelectedPolys = True
			lngWShedLayerIndex = modUtil.GetLayerIndex((cboSelectPoly.Text), m_App)
			pSelectedPolyLayer = m_pMap.Layer(lngWShedLayerIndex)
		Else
			g_booSelectedPolys = False
		End If
		'END STEP 3: ---------------------------------------------------------------------------------------------------------
		
		'Check for Spatial Analyst Extension, doing it once here to eliminate multiple checks down the road
		If Not CheckSpatialAnalystLicense Then
			MsgBox("Spatial Analyst Extension is required for N-SPECT and a license is not available.", MsgBoxStyle.Critical, "Spatial Analyst Required")
			Exit Sub
		End If
		
		'STEP 4: Get the Management Scenarios: ------------------------------------------------------------------------------------
		'If they're using, we send them over to modMgmtScen to implement
		If m_XMLPrjParams.clsMgmtScenHolder.Count > 0 Then
			modMgmtScen.MgmtScenSetup((m_XMLPrjParams.clsMgmtScenHolder), (m_XMLPrjParams.strLCGridType), (m_XMLPrjParams.strLCGridFileName), (m_XMLPrjParams.strProjectWorkspace))
		End If
		'END STEP 4: ---------------------------------------------------------------------------------------------------------
		
		'STEP 5: Pollutant Dictionary creation, needed for Landuse -----------------------------------------------------------
		'Go through and find the pollutants, if they're used and what the CoeffSet is
		'We're creating a dictionary that will hold Pollutant, Coefficient Set for use in the Landuse Scenarios
		For i = 1 To m_XMLPrjParams.clsPollItems.Count
			If m_XMLPrjParams.clsPollItems.Item(i).intApply = 1 Then
				dictPollutants.Add(m_XMLPrjParams.clsPollItems.Item(i).strPollName, m_XMLPrjParams.clsPollItems.Item(i).strCoeffSet)
			End If
		Next i
		'END STEP 5: ---------------------------------------------------------------------------------------------------------
		
		'STEP 6: Landuses sent off to modLanduse for processing -----------------------------------------------------
		For i = 1 To m_XMLPrjParams.clsLUItems.Count
			If m_XMLPrjParams.clsLUItems.Item(i).intApply = 1 Then
				booLUItems = True
				modLanduse.Begin((m_XMLPrjParams.strLCGridType), (m_XMLPrjParams.clsLUItems), dictPollutants, (m_XMLPrjParams.strLCGridFileName), m_pMap, (m_XMLPrjParams.strProjectWorkspace))
				Exit For
			Else
				booLUItems = False
			End If
		Next i
		'END STEP 6: ---------------------------------------------------------------------------------------------------------
		
		'STEP 7: ---------------------------------------------------------------------------------------------------------
		'Obtain Watershed values
		
		strWaterShed = "Select * from WSDelineation Where Name like '" & m_XMLPrjParams.strWaterShedDelin & "'"
		rsWaterShed.Open(strWaterShed, g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
		
		'END STEP 7: -----------------------------------------------------------------------------------------------------
		
		'STEP 8: ---------------------------------------------------------------------------------------------------------
		'Set the Analysis Environment and globals for output workspace
		
		modMainRun.SetGlobalEnvironment(rsWaterShed, (m_XMLPrjParams.strProjectWorkspace), m_pMap, pSelectedPolyLayer)
		
		'END STEP 8: -----------------------------------------------------------------------------------------------------
		
		'STEP 8a: --------------------------------------------------------------------------------------------------------
		'Added 1/08/2007 to account for non-adjacent polygons
		Dim pPolygon As ESRI.ArcGIS.Geometry.IPolygon4
		If m_XMLPrjParams.intSelectedPolys = 1 Then
			pPolygon = g_pSelectedPolyClip
			If modMainRun.CheckMultiPartPolygon(pPolygon) Then
				MsgBox("Warning: Your selected polygons are not adjacent.  Please select only polygons that are adjacent.", MsgBoxStyle.Critical, "Non-adjacent Polygons Detected")
				'UPGRADE_ISSUE: Unable to determine which constant to upgrade vbNormal to. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B3B44E51-B5F1-4FD7-AA29-CAD31B71F487"'
				'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
				'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
				System.Windows.Forms.Cursor.Current = vbNormal
				Exit Sub
			End If
		End If
		
		'STEP 9: ---------------------------------------------------------------------------------------------------------
		'Create the runoff GRID
		'Get the precip scenario stuff
		strPrecip = "Select * from PrecipScenario where name like '" & m_XMLPrjParams.strPrecipScenario & "'"
		rsPrecip.Open(strPrecip, g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
		
		'If there has been a land use added, then a new LCType has been created, hence we get it from g_strLCTypename
		Dim strLCType As String
		If booLUItems Then
			strLCType = modLanduse.g_strLCTypeName
		Else
			strLCType = m_XMLPrjParams.strLCGridType
		End If
		
		'Added 6/04 to account for different PrecipTypes
		modMainRun.g_intPrecipType = rsPrecip.Fields("PrecipType").Value
		
		If Not modRunoff.CreateRunoffGrid((m_XMLPrjParams.strLCGridFileName), strLCType, rsPrecip, (m_XMLPrjParams.strSoilsHydFileName)) Then
			Exit Sub
		End If
		'END STEP 9: -----------------------------------------------------------------------------------------------------
		
		'STEP 10: ---------------------------------------------------------------------------------------------------------
		'Process pollutants
		For i = 1 To m_XMLPrjParams.clsPollItems.Count
			If m_XMLPrjParams.clsPollItems.Item(i).intApply = 1 Then
				'If user is NOT ignoring the pollutant then send the whole item over along with LCType
				If Not modPollutantCalcs.PollutantConcentrationSetup(m_XMLPrjParams.clsPollItems.Item(i), (m_XMLPrjParams.strLCGridType), (m_XMLPrjParams.strWaterQuality)) Then
					Exit Sub
				End If
			End If
		Next i
		'END STEP 10: -----------------------------------------------------------------------------------------------------
		
		'Step 11: Erosion -------------------------------------------------------------------------------------------------
		'Check that they have chosen Erosion
		If m_XMLPrjParams.intCalcErosion = 1 Then
			If m_booAnnualPrecip Then 'If Annual (0) then TRUE, ergo RUSLE
				If m_XMLPrjParams.intRainGridBool Then
					If Not modRusle.RUSLESetup(rsWaterShed.Fields("NibbleFileName").Value, rsWaterShed.Fields("dem2bfilename").Value, (m_XMLPrjParams.strRainGridFileName), (m_XMLPrjParams.strSoilsKFileName), (m_XMLPrjParams.strSDRGridFileName), (m_XMLPrjParams.strLCGridType)) Then
						Exit Sub
					End If
				ElseIf m_XMLPrjParams.intRainConstBool Then 
					If Not modRusle.RUSLESetup(rsWaterShed.Fields("NibbleFileName").Value, rsWaterShed.Fields("dem2bfilename").Value, (m_XMLPrjParams.strRainGridFileName), (m_XMLPrjParams.strSoilsKFileName), (m_XMLPrjParams.strSDRGridFileName), (m_XMLPrjParams.strLCGridType), m_XMLPrjParams.dblRainConstValue) Then
						Exit Sub
					End If
				End If
				
			Else 'If event (1) then False, ergo MUSLE
				If Not modMUSLE.MUSLESetup((m_XMLPrjParams.strSoilsDefName), (m_XMLPrjParams.strSoilsKFileName), (m_XMLPrjParams.strLCGridType)) Then
					Exit Sub
				End If
			End If
		End If
		'STEP 11: ----------------------------------------------------------------------------------------------------------
		
		'STEP 12 : Cleanup any temp critters -------------------------------------------------------------------------------
		'g_DictTempNames holds the names of all temporary landuses and/or coefficient sets created during the Landuse scenario
		'portion of our program, for example CCAP1, or NitSet1.  We now must eliminate them from the database if they exist.
		If g_DictTempNames.Count > 0 Then
			If booLUItems Then
				modLanduse.Cleanup(g_DictTempNames, (m_XMLPrjParams.clsPollItems), (m_XMLPrjParams.strLCGridType))
			End If
		End If
		'END STEP 12: -------------------------------------------------------------------------------------------------------
		
		'STEP 13: -----------------------------------------------------------------------------------------------------------
		'g_pGroupLayer has been created earlier and has been taken on GRIDs since.  Now lets add it
		'Add the group layer.
		With g_pGroupLayer
			.Expanded = True 'Are going to 'expand' it
			'UPGRADE_WARNING: Couldn't resolve default property of object g_pGroupLayer.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.Name = m_XMLPrjParams.strProjectName 'The name equals whatever the user entered
		End With
		
		'UPGRADE_WARNING: Couldn't resolve default property of object g_pGroupLayer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_pMap.AddLayer(g_pGroupLayer)
		'END STEP 13: -------------------------------------------------------------------------------------------------------
		
		'STEP 14 save out group layer ---------------------------------------------------------------------------------------
		'UPGRADE_WARNING: Couldn't resolve default property of object g_pGroupLayer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		modUtil.ExportLayerToPath(g_pGroupLayer, m_XMLPrjParams.strProjectWorkspace & "\" & m_XMLPrjParams.strProjectName & ".lyr")
		'END STEP 14: -------------------------------------------------------------------------------------------------------
		
		'STEP 15: create string describing project parameters ---------------------------------------------------------------
		strProjectInfo = modUtil.ParseProjectforMetadata(m_XMLPrjParams, m_strFileName)
		'END STEP 15: -------------------------------------------------------------------------------------------------------
		
		'STEP 16: Apply the metadata to each of the rasters in the group layer ----------------------------------------------
		m_App.StatusBar.Message(0) = "Creating metadata for the N-SPECT group layer..."
		modUtil.CreateMetadata(g_pGroupLayer, strProjectInfo)
		'END STEP 16: -------------------------------------------------------------------------------------------------------
		
		'Cleanup ------------------------------------------------------------------------------------------------------------
		m_App.StatusBar.Message(0) = "Deleting temporary files..."
		rsWaterShed.Close()
		rsPrecip.Close()
		'UPGRADE_NOTE: Object rsWaterShed may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsWaterShed = Nothing
		'UPGRADE_NOTE: Object rsPrecip may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsPrecip = Nothing
		
		'Go into workspace and rid it of all rasters
		modUtil.CleanGlobals()
		modUtil.CleanupRasterFolder((m_XMLPrjParams.strProjectWorkspace))
		
		'UPGRADE_NOTE: Object g_pGroupLayer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		g_pGroupLayer = Nothing
		'UPGRADE_NOTE: Object g_DictTempNames may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		g_DictTempNames = Nothing
		'UPGRADE_NOTE: Object g_dicMetadata may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		g_dicMetadata = Nothing
		
		'UPGRADE_ISSUE: Unable to determine which constant to upgrade vbNormal to. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B3B44E51-B5F1-4FD7-AA29-CAD31B71F487"'
		'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = vbNormal
		
		m_App.StatusBar.Message(0) = "N-SPECT processing complete!"
		
		Me.Close()
		
		Exit Sub
		
UserCancel: 
		modProgDialog.KillDialog()
		'UPGRADE_ISSUE: Unable to determine which constant to upgrade vbNormal to. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B3B44E51-B5F1-4FD7-AA29-CAD31B71F487"'
		'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = vbNormal
		MsgBox("Processing has been stopped.", MsgBoxStyle.Information, "Analysis Stopped")
		
ErrorHandler: 
		HandleError(True, "cmdRun_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
		'UPGRADE_ISSUE: Unable to determine which constant to upgrade vbNormal to. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B3B44E51-B5F1-4FD7-AA29-CAD31B71F487"'
		'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = vbNormal
	End Sub
	
	Private Sub cmdSave_Click()
		On Error GoTo ErrorHandler
		
		dlgXMLOpen.FileName = CStr(Nothing)
		dlgXMLSave.FileName = CStr(Nothing)
		'UPGRADE_WARNING: CommonDialog variable was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="671167DC-EA81-475D-B690-7A40C7BF4A23"'
		With dlgXML
			
			'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			.Filter = MSG8
			.Title = MSG2
			.FilterIndex = 1
			'UPGRADE_WARNING: MSComDlg.CommonDialog property dlgXML.Flags was upgraded to dlgXMLSave.OverwritePrompt which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
			.OverwritePrompt = True
			.CheckPathExists = True
			.CheckPathExists = True
			.ShowDialog()
			
		End With
		
		If Len(dlgXMLOpen.FileName) > 0 Then
			m_strFileName = Trim(dlgXMLOpen.FileName)
			m_XMLPrjParams.SaveFile(m_strFileName)
		End If
		
		Exit Sub
ErrorHandler: 
		HandleError(False, "cmdSave_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	Private Sub LoadXMLFile()
		'On Error GoTo ErrorHandler
		
		'browse...get output filename
		Dim fso As Scripting.FileSystemObject
		Dim strFolder As String
		
		fso = New Scripting.FileSystemObject
		strFolder = modUtil.g_nspectDocPath & "\projects"
		If Not fso.FolderExists(strFolder) Then
			MkDir(strFolder)
		End If
		
		dlgXMLOpen.FileName = CStr(Nothing)
		dlgXMLSave.FileName = CStr(Nothing)
		'UPGRADE_WARNING: CommonDialog variable was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="671167DC-EA81-475D-B690-7A40C7BF4A23"'
		With dlgXML
			'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			.Filter = MSG8
			.InitialDirectory = strFolder
			.Title = "Open N-SPECT Project File"
			.FilterIndex = 1
			'UPGRADE_WARNING: FileOpenConstants constant FileOpenConstants.cdlOFNHideReadOnly was upgraded to OpenFileDialog.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
			'UPGRADE_WARNING: MSComDlg.CommonDialog property dlgXML.Flags was upgraded to dlgXMLOpen.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
			'UPGRADE_WARNING: FileOpenConstants constant FileOpenConstants.cdlOFNHideReadOnly was upgraded to OpenFileDialog.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
			.ShowReadOnly = False
			'UPGRADE_WARNING: MSComDlg.CommonDialog property dlgXML.Flags was upgraded to dlgXMLOpen.CheckFileExists which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
			.CheckFileExists = True
			.CheckPathExists = True
			.CheckPathExists = True
			.ShowDialog()
		End With
		
		If Len(dlgXMLOpen.FileName) > 0 Then
			m_strFileName = Trim(dlgXMLOpen.FileName)
			m_XMLPrjParams.XML = m_strFileName
			FillForm()
		Else
			Exit Sub
		End If
		
		'Pop this string with the incoming name, if they change, we'll prompt to 'save as'.
		m_strOpenFileName = txtProjectName.Text
		
		'UPGRADE_NOTE: Object fso may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		fso = Nothing
		
		'Exit Sub
		'ErrorHandler:
		'  HandleError False, "LoadXMLFile " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
	End Sub
	
	Private Function SaveXMLFile() As Boolean
		
		On Error GoTo ErrHandler
		
		Dim strFolder As String
		Dim intvbYesNo As Short
		Dim fso As Scripting.FileSystemObject
		
		fso = New Scripting.FileSystemObject
		
		strFolder = modUtil.g_nspectDocPath & "\projects"
		If Not fso.FolderExists(strFolder) Then
			MkDir(strFolder)
		End If
		
		If Not ValidateData Then 'check form inputs
			SaveXMLFile = False
			Exit Function
		End If
		
		'If it does not already exist, open Save As... dialog
		If Not m_booExists Then
			'UPGRADE_WARNING: CommonDialog variable was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="671167DC-EA81-475D-B690-7A40C7BF4A23"'
			With dlgXML
				'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				.Filter = MSG8
				.Title = "Save Project File As..."
				.InitialDirectory = strFolder
				.FilterIndex = 1
				'UPGRADE_WARNING: MSComDlg.CommonDialog property dlgXML.Flags was upgraded to dlgXMLSave.OverwritePrompt which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
				.OverwritePrompt = True
				'UPGRADE_WARNING: MSComDlg.CommonDialog property dlgXML.Flags was upgraded to dlgXMLOpen.CheckFileExists which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
				.CheckFileExists = True
				.CheckPathExists = True
				.CheckPathExists = True
				'UPGRADE_WARNING: The CommonDialog CancelError property is not supported in Visual Basic .NET. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"'
				.CancelError = True
				.FileName = txtProjectName.Text
				.ShowDialog()
			End With
			'check to make sure filename length is greater than zeros
			If Len(dlgXMLOpen.FileName) > 0 Then
				m_strFileName = Trim(dlgXMLOpen.FileName)
				m_booExists = True
				m_XMLPrjParams.SaveFile(m_strFileName)
				SaveXMLFile = True
			Else
				SaveXMLFile = False
				Exit Function
			End If
			
		Else
			'Now check to see if the name changed
			If m_strOpenFileName <> txtProjectName.Text Then
				intvbYesNo = MsgBox("You have changed the name of this project.  Would you like to save your settings as a new file?" & vbNewLine & vbTab & "Yes" & vbTab & " -    Save as new N-SPECT project file" & vbNewLine & vbTab & "No" & vbTab & " -    Save changes to current N-SPECT project file" & vbNewLine & vbTab & "Cancel" & vbTab & " -    Return to the project window", MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, "N-SPECT")
				
				If intvbYesNo = MsgBoxResult.Yes Then
					'UPGRADE_WARNING: CommonDialog variable was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="671167DC-EA81-475D-B690-7A40C7BF4A23"'
					With dlgXML
						'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
						.Filter = MSG8
						.Title = "Save Project File As..."
						.InitialDirectory = strFolder
						.FilterIndex = 1
						'UPGRADE_WARNING: MSComDlg.CommonDialog property dlgXML.Flags was upgraded to dlgXMLSave.OverwritePrompt which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
						.OverwritePrompt = True
						'UPGRADE_WARNING: MSComDlg.CommonDialog property dlgXML.Flags was upgraded to dlgXMLOpen.CheckFileExists which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
						.CheckFileExists = True
						.CheckPathExists = True
						.CheckPathExists = True
						'UPGRADE_WARNING: The CommonDialog CancelError property is not supported in Visual Basic .NET. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"'
						.CancelError = True
						.FileName = txtProjectName.Text
						.ShowDialog()
					End With
					'check to make sure filename length is greater than zeros
					If Len(dlgXMLOpen.FileName) > 0 Then
						m_strFileName = Trim(dlgXMLOpen.FileName)
						m_booExists = True
						m_XMLPrjParams.SaveFile(m_strFileName)
						SaveXMLFile = True
					Else
						SaveXMLFile = False
						Exit Function
					End If
				ElseIf intvbYesNo = MsgBoxResult.No Then 
					m_XMLPrjParams.SaveFile(m_strFileName)
					m_booExists = True
					SaveXMLFile = True
				Else
					SaveXMLFile = False
					Exit Function
				End If
			Else
				m_XMLPrjParams.SaveFile(m_strFileName)
				m_booExists = True
				SaveXMLFile = True
				
			End If
			
		End If
		
		'UPGRADE_NOTE: Object fso may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		fso = Nothing
		
		Exit Function
		
ErrHandler: 
		
		If Err.Number = 32755 Then
			SaveXMLFile = False
			Exit Function
		Else
			MsgBox(Err.Number & " " & Err.Description)
			SaveXMLFile = False
		End If
		
	End Function
	
	Private Sub FillForm()
		'On Error GoTo ErrorHandler
		On Error Resume Next
		Dim rsCurrWShed As New ADODB.Recordset
		Dim strCurrWShed As String
		Dim pCurrWShedPolyLayer As ESRI.ArcGIS.Carto.ILayer
		Dim lngCurrWshedPolyIndex As Integer
		Dim intYesNo As Short
		Dim pDBasinFeatureLayer As ESRI.ArcGIS.Carto.IFeatureLayer
		Dim pDBasinDataset As ESRI.ArcGIS.Geodatabase.IDataset
		Dim strBasinPoly As String
		Dim i As Short
		Dim z As Short
		Dim booNameMatch As Object
		Dim fso As New Scripting.FileSystemObject
		
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		
		txtProjectName.Text = m_XMLPrjParams.strProjectName
		txtOutputWS.Text = m_XMLPrjParams.strProjectWorkspace
		
		'Step 1:  LandCoverGrid
		'Check to see if the LC cover is in the map, if so, set the combobox
		If modUtil.LayerInMap((m_XMLPrjParams.strLCGridName), m_pMap) Then
			cboLCLayer.SelectedIndex = modUtil.GetCboIndex((m_XMLPrjParams.strLCGridName), cboLCLayer)
			cboLCLayer.Refresh()
		Else
			If fso.FileExists(m_XMLPrjParams.strLCGridFileName & ".lyr") Then
				
				modUtil.AddLayerFileToMap(m_XMLPrjParams.strLCGridFileName & ".lyr", m_pMap)
				
			ElseIf modUtil.AddRasterLayerToMapFromFileName((m_XMLPrjParams.strLCGridFileName), m_pMap) Then 
				
				With cboLCLayer
					'.AddItem m_XMLPrjParams.strLCGridName
					.Refresh()
					'.ListIndex = modUtil.GetCboIndex(m_XMLPrjParams.strLCGridName, cboLCLayer)
				End With
				
			Else
				intYesNo = MsgBox("Could not find the Land Cover dataset: " & m_XMLPrjParams.strLCGridFileName & ".  Would you like " & "to browse for it?", MsgBoxStyle.YesNo, "Cannot Locate Dataset")
				If intYesNo = MsgBoxResult.Yes Then
					m_XMLPrjParams.strLCGridFileName = modUtil.AddInputFromGxBrowser(cboLCLayer, Me, "Raster")
					If m_XMLPrjParams.strLCGridFileName <> "" Then
						If modUtil.AddRasterLayerToMapFromFileName((m_XMLPrjParams.strLCGridFileName), m_pMap) Then
							cboLCLayer.SelectedIndex = 0
						End If
					Else
						Exit Sub
					End If
				Else
					Exit Sub
				End If
			End If
			
		End If
		
		cboLCUnits.SelectedIndex = modUtil.GetCboIndex((m_XMLPrjParams.strLCGridUnits), cboLCUnits)
		cboLCUnits.Refresh()
		
		cboLCType.SelectedIndex = modUtil.GetCboIndex((m_XMLPrjParams.strLCGridType), cboLCType)
		cboLCType.Refresh()
		
		'Step 2: Soils - same process, if in doc and map, OK, else send em looking
		If modUtil.RasterExists((m_XMLPrjParams.strSoilsHydFileName)) Then
			cboSoilsLayer.SelectedIndex = modUtil.GetCboIndex((m_XMLPrjParams.strSoilsDefName), cboSoilsLayer)
			cboSoilsLayer.Refresh()
		Else
			MsgBox("Could not find soils dataset.  Please correct the soils definition in the Advanced Settings.", MsgBoxStyle.Critical, "Dataset Missing")
			Exit Sub
		End If
		
		'Step5: Precip Scenario
		cboPrecipScen.SelectedIndex = modUtil.GetCboIndex((m_XMLPrjParams.strPrecipScenario), cboPrecipScen)
		
		'Step6: Watershed Delineation
		cboWSDelin.SelectedIndex = modUtil.GetCboIndex((m_XMLPrjParams.strWaterShedDelin), cboWSDelin)
		
		'Add the basinpoly to the map
		strCurrWShed = "Select * from WSDelineation where Name Like '" & m_XMLPrjParams.strWaterShedDelin & "'"
		rsCurrWShed.Open(strCurrWShed, g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
		
		
		'**********************************************************************************************
		'**********************************************************************************************
		'Check to see if the Drainage Basins poly is already in the map
		'    For z = 0 To m_pMap.LayerCount - 1
		'
		'        If TypeOf m_pMap.Layer(z) Is IFeatureLayer Then
		'            Set pDBasinFeatureLayer = m_pMap.Layer(z)
		'            Set pDBasinDataset = pDBasinFeatureLayer
		'            strBasinPoly = Trim(pDBasinDataset.Workspace.PathName & "\" & pDBasinFeatureLayer.FeatureClass.AliasName)
		'
		'            If strBasinPoly <> rsCurrWShed!wsfilename Then
		'                booNameMatch = False
		'            Else
		'                booNameMatch = True
		'                Exit For
		'            End If
		'        End If
		'    Next z
		
		
		If Not modUtil.LayerInMapByFileName(rsCurrWShed.Fields("wsfilename").Value, m_pMap) Then
			If modUtil.AddFeatureLayerToMapFromFileName(rsCurrWShed.Fields("wsfilename").Value, m_pMap, rsCurrWShed.Fields("Name").Value & " " & "Drainage Basins") Then
				lngCurrWshedPolyIndex = modUtil.GetLayerIndex(rsCurrWShed.Fields("Name").Value & " " & "Drainage Basins", m_App)
				m_pMxDoc.ContentsView(0).Refresh(m_pMap.Layer(lngCurrWshedPolyIndex))
			Else
				MsgBox("Could not find watershed layer: " & rsCurrWShed.Fields("wsfilename").Value & " .  Please add the watershed layer to the map.", MsgBoxStyle.Critical, "File Not Found")
			End If
		End If
		
		
		'    If Not booNameMatch Then
		'        If modUtil.AddFeatureLayerToMapFromFileName(rsCurrWShed!wsfilename, m_pMap, rsCurrWShed!Name & " " & "Drainage Basins") Then
		'            lngCurrWshedPolyIndex = modUtil.GetLayerIndex(rsCurrWShed!Name & " " & "Drainage Basins", m_App)
		'            m_pMxDoc.ContentsView(0).Refresh m_pMap.Layer(lngCurrWshedPolyIndex)
		'        End If
		'    End If
		
		
		'**********************************************************************************************
		'**********************************************************************************************
		'    If modUtil.LayerInMap(rsCurrWShed!Name & " " & "Drainage Basins", m_pMap) Then
		'        Set pDBasinFeatureLayer = m_pMap.Layer(modUtil.GetLayerIndex(rsCurrWShed!Name & " " & "Drainage Basins", m_App))
		'        Set pDBasinDataset = pDBasinFeatureLayer
		'        strBasinPoly = Trim(pDBasinDataset.Workspace.PathName & "\" & pDBasinFeatureLayer.FeatureClass.AliasName)
		'        'if it is not the same one then add it
		'        If strBasinPoly <> rsCurrWShed!wsfilename Then
		'            If modUtil.AddFeatureLayerToMapFromFileName(rsCurrWShed!wsfilename, m_pMap, rsCurrWShed!Name & " " & "Drainage Basins") Then
		'                lngCurrWshedPolyIndex = modUtil.GetLayerIndex(rsCurrWShed!Name & " " & "Drainage Basins", m_App)
		'                m_pMxDoc.ContentsView(0).Refresh m_pMap.Layer(lngCurrWshedPolyIndex)
		'            End If
		'        End If
		'    Else
		'        'if not in the map do straight add in
		'        If modUtil.AddFeatureLayerToMapFromFileName(rsCurrWShed!wsfilename, m_pMap, rsCurrWShed!Name & " " & "Drainage Basins") Then
		'            lngCurrWshedPolyIndex = modUtil.GetLayerIndex(rsCurrWShed!Name & " " & "Drainage Basins", m_App)
		'            m_pMxDoc.ContentsView(0).Refresh m_pMap.Layer(lngCurrWshedPolyIndex)
		'        End If
		'    End If
		
		'**********************************************************************************************
		'**********************************************************************************************
		
		'Step7: Water Quality
		cboWQStd.SelectedIndex = modUtil.GetCboIndex((m_XMLPrjParams.strWaterQuality), cboWQStd)
		
		'Step8: LocalEffects/Selected Watersheds
		chkLocalEffects.CheckState = m_XMLPrjParams.intLocalEffects
		chkSelectedPolys.CheckState = m_XMLPrjParams.intSelectedPolys
		
		If chkSelectedPolys.CheckState = 1 Then
			'1st see if it's in the map
			If modUtil.LayerInMapByFileName((m_XMLPrjParams.strSelectedPolyFileName), m_pMap) Then
				cboSelectPoly.SelectedIndex = modUtil.GetCboIndex((m_XMLPrjParams.strSelectedPolyLyrName), cboSelectPoly)
				cboSelectPoly.Refresh()
			Else
				'Not there then add it
				If modUtil.AddFeatureLayerToMapFromFileName((m_XMLPrjParams.strSelectedPolyFileName), m_pMap, m_XMLPrjParams.strSelectedPolyLyrName) Then
					cboSelectPoly.Refresh()
					cboSelectPoly.SelectedIndex = modUtil.GetCboIndex((m_XMLPrjParams.strSelectedPolyLyrName), cboSelectPoly)
				Else
					'Can't find it, then send em searching
					intYesNo = MsgBox("Could not find the Selected Polygons file used to limit extent: " & m_XMLPrjParams.strSelectedPolyFileName & ".  Would you like to browse for it? ", MsgBoxStyle.YesNo, "Cannot Locate Dataset")
					If intYesNo = MsgBoxResult.Yes Then
						'if they want to look for it then give em the browser
						m_XMLPrjParams.strSelectedPolyFileName = modUtil.AddInputFromGxBrowser(cboSelectPoly, Me, "Feature")
						'if they actually find something, throw it in the map
						If Len(m_XMLPrjParams.strSelectedPolyFileName) > 0 Then
							If modUtil.AddFeatureLayerToMapFromFileName((m_XMLPrjParams.strSelectedPolyFileName), m_pMap) Then
								cboSelectPoly.SelectedIndex = modUtil.GetCboIndex(modUtil.SplitFileName((m_XMLPrjParams.strSelectedPolyFileName)), cboSelectPoly)
							End If
						End If
					Else
						m_XMLPrjParams.intSelectedPolys = 0
						chkSelectedPolys.CheckState = System.Windows.Forms.CheckState.Unchecked
					End If
				End If
			End If
		End If
		
		'Step: Erosion Tab - Calc Erosion, Erosion Attribute
		If m_XMLPrjParams.intCalcErosion = 1 Then
			chkCalcErosion.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkCalcErosion.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		
		'Step: Erosion Tab - Precip
		'Either they use the GRID
		optUseGRID.Checked = m_XMLPrjParams.intRainGridBool
		
		If optUseGRID.Checked Then
			If modUtil.LayerInMap((m_XMLPrjParams.strRainGridName), m_pMap) Then
				cboRainGrid.SelectedIndex = modUtil.GetCboIndex((m_XMLPrjParams.strRainGridName), cboRainGrid)
				cboRainGrid.Refresh()
			Else
				If modUtil.AddRasterLayerToMapFromFileName((m_XMLPrjParams.strRainGridFileName), m_pMap) Then
					With cboRainGrid
						.Items.Add(m_XMLPrjParams.strRainGridName)
						.Refresh()
						.SelectedIndex = modUtil.GetCboIndex((m_XMLPrjParams.strRainGridName), cboRainGrid)
					End With
				Else
					intYesNo = MsgBox("Could not find Rainfall GRID: " & m_XMLPrjParams.strRainGridName & ".  Would you like " & "to browse for it?", MsgBoxStyle.YesNo, "Cannot Locate Dataset")
					If intYesNo = MsgBoxResult.Yes Then
						m_XMLPrjParams.strRainGridFileName = modUtil.AddInputFromGxBrowser(cboRainGrid, Me, "Raster")
						If modUtil.AddRasterLayerToMapFromFileName((m_XMLPrjParams.strRainGridFileName), m_pMap) Then
							cboRainGrid.SelectedIndex = 0
						End If
					Else
						Exit Sub
					End If
				End If
			End If
		End If
		
		'Or they use a constant value
		optUseValue.Checked = m_XMLPrjParams.intRainConstBool
		
		If optUseValue.Checked Then
			txtRainValue.Text = CStr(m_XMLPrjParams.dblRainConstValue)
		End If
		
		'SDR GRID
		
		'If Not m_XMLPrjParams.intUseOwnSDR Is Nothing Then
		On Error GoTo VersionProblem
		If m_XMLPrjParams.intUseOwnSDR = 1 Then
			chkSDR.CheckState = System.Windows.Forms.CheckState.Checked
			txtSDRGRID.Text = m_XMLPrjParams.strLCGridFileName
		Else
			chkSDR.CheckState = System.Windows.Forms.CheckState.Unchecked
			txtSDRGRID.Text = m_XMLPrjParams.strSDRGridFileName
		End If
		'End If
		
		'Step Pollutants
		m_intPollCount = m_XMLPrjParams.clsPollItems.Count
		
		If m_intPollCount > 0 Then
			grdCoeffs.Rows = m_intPollCount + 1
			For i = 1 To m_intPollCount
				With grdCoeffs
					.row = m_XMLPrjParams.clsPollItems.Item(i).intID
					.set_TextMatrix(.row, 1, m_XMLPrjParams.clsPollItems.Item(i).intApply)
					.set_TextMatrix(.row, 2, m_XMLPrjParams.clsPollItems.Item(i).strPollName)
					.set_TextMatrix(.row, 3, m_XMLPrjParams.clsPollItems.Item(i).strCoeffSet)
					.set_TextMatrix(.row, 4, m_XMLPrjParams.clsPollItems.Item(i).strCoeff)
					.set_TextMatrix(.row, 5, CStr(m_XMLPrjParams.clsPollItems.Item(i).intThreshold))
					
					If Len(m_XMLPrjParams.clsPollItems.Item(i).strTypeDefXMLFile) > 0 Then
						.set_TextMatrix(.row, 6, CStr(m_XMLPrjParams.clsPollItems.Item(i).strTypeDefXMLFile))
					End If
					
				End With
			Next i
			
			ClearCheckBoxes(True)
			CreateCheckBoxes(True, True)
		End If
		
		'Step - Land Uses
		m_intLUCount = m_XMLPrjParams.clsLUItems.Count
		
		If m_intLUCount > 0 Then
			grdLU.Rows = m_intLUCount + 1
			For i = 1 To m_intLUCount
				With grdLU
					.row = m_XMLPrjParams.clsLUItems.Item(i).intID
					.set_TextMatrix(.row, 1, m_XMLPrjParams.clsLUItems.Item(i).intApply)
					.set_TextMatrix(.row, 2, m_XMLPrjParams.clsLUItems.Item(i).strLUScenName)
					.set_TextMatrix(.row, 3, m_XMLPrjParams.clsLUItems.Item(i).strLUScenXMLFile)
				End With
			Next i
			
			ClearLUCheckBoxes(True)
			CreateLUCheckBoxes(True)
			
		End If
		
		'Step Management Scenarios
		m_intMgmtCount = m_XMLPrjParams.clsMgmtScenHolder.Count
		
		If m_intMgmtCount > 0 Then
			grdLCChanges.Rows = m_intMgmtCount + 1
			For i = 1 To m_intMgmtCount
				With grdLCChanges
					.row = m_XMLPrjParams.clsMgmtScenHolder.Item(i).intID
					If modUtil.LayerInMap((m_XMLPrjParams.clsMgmtScenHolder.Item(i).strAreaName), m_pMap) Then
						.set_TextMatrix(i, 1, m_XMLPrjParams.clsMgmtScenHolder.Item(i).intApply)
						.set_TextMatrix(i, 2, m_XMLPrjParams.clsMgmtScenHolder.Item(i).strAreaName)
						.set_TextMatrix(i, 3, m_XMLPrjParams.clsMgmtScenHolder.Item(i).strChangeToClass)
					Else
						If modUtil.AddFeatureLayerToMapFromFileName((m_XMLPrjParams.clsMgmtScenHolder.Item(i).strAreaFileName), m_pMap) Then
							.set_TextMatrix(i, 1, m_XMLPrjParams.clsMgmtScenHolder.Item(i).intApply)
							.set_TextMatrix(i, 2, m_XMLPrjParams.clsMgmtScenHolder.Item(i).strAreaName)
							.set_TextMatrix(i, 3, m_XMLPrjParams.clsMgmtScenHolder.Item(i).strChangeToClass)
						Else
							intYesNo = MsgBox("Could not find Management Sceario Area Layer: " & m_XMLPrjParams.clsMgmtScenHolder.Item(i).strAreaFileName & ".  Would you like " & "to browse for it?", MsgBoxStyle.YesNo, "Cannot Locate Dataset:" & m_XMLPrjParams.clsMgmtScenHolder.Item(i).strAreaFileName)
							If intYesNo = MsgBoxResult.Yes Then
								m_XMLPrjParams.clsMgmtScenHolder.Item(i).strAreaFileName = modUtil.AddInputFromGxBrowser(cboAreaLayer, Me, "Feature")
								If modUtil.AddFeatureLayerToMapFromFileName((m_XMLPrjParams.clsMgmtScenHolder.Item(i).strAreaFileName), m_pMap) Then
									cboAreaLayer.SelectedIndex = 0
									.set_TextMatrix(i, 1, m_XMLPrjParams.clsMgmtScenHolder.Item(i).intApply)
									.set_TextMatrix(i, 2, m_XMLPrjParams.clsMgmtScenHolder.Item(i).strAreaName)
									.set_TextMatrix(i, 3, m_XMLPrjParams.clsMgmtScenHolder.Item(i).strChangeToClass)
								End If
							Else
								Exit Sub
							End If
						End If
					End If
				End With
			Next i
			
			ClearMgmtCheckBoxes(True, m_intMgmtCount)
			CreateMgmtCheckBoxes(True)
			
		End If
		
		'Reset to first tab
		SSTab1.SelectedIndex = 0
		m_booExists = True
		'UPGRADE_ISSUE: Unable to determine which constant to upgrade vbNormal to. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B3B44E51-B5F1-4FD7-AA29-CAD31B71F487"'
		'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = vbNormal
		
		'Cleanup
		rsCurrWShed.Close()
		'UPGRADE_NOTE: Object rsCurrWShed may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsCurrWShed = Nothing
		'UPGRADE_NOTE: Object pDBasinFeatureLayer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pDBasinFeatureLayer = Nothing
		'UPGRADE_NOTE: Object pDBasinDataset may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pDBasinDataset = Nothing
		
		Exit Sub
ErrorHandler: 
		HandleError(False, "FillForm " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
VersionProblem: 
		MsgBox("Version Problem")
	End Sub
	
	Private Sub cmdOpenWS_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOpenWS.Click
		On Error GoTo ErrorHandler
		
		Dim initFolder As String
		
		Dim shlShell As Shell32.Shell
		Dim shlFolder As Shell32.Folder3
		Dim fso As Scripting.FileSystemObject
		
		fso = New Scripting.FileSystemObject
		initFolder = modUtil.g_nspectDocPath & "\workspace"
		If Not fso.FolderExists(initFolder) Then
			MkDir(initFolder)
		End If
		
		shlShell = New Shell32.Shell
		shlFolder = shlShell.BrowseForFolder(Me.Handle.ToInt32, "Choose a directory for analysis output: ", &H1) ', initFolder)
		If Not shlFolder Is Nothing Then
			Me.txtOutputWS.Text = shlFolder.Self.path
			'MsgBox "Output folder: " & Me.txtOutputWS.Text
			m_strWorkspace = txtOutputWS.Text
		End If
		
		'Open workspace button
		'uses the addin folder browser; Reference: vbAccelerator folder browse library
		'dlgBrowser is initiated on Form Load
		'With dlgBrowser
		'    .hwndOwner = Me.hwnd
		'    .InitialDir = initFolder
		'    .FileSystemOnly = True
		'    .StatusText = True
		'    .EditBox = True
		'    .UseNewUI = True
		'    .Title = "Choose a Directory for Analysis Output:"
		'End With
		'txtOutputWS.Text = dlgBrowser.BrowseForFolder
		'm_strWorkspace = txtOutputWS.Text
		
		Exit Sub
		
ErrorHandler: 
		HandleError(True, "cmdOpenWS_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	
	Private Sub frmPrj_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		On Error GoTo ErrorHandler
		
		'Enables Shift + F1 to bring up NSPECT help.  Regular F1 brings in the darn ArcMap help
		
		Dim shiftdown As Short
		
		If (Shift And VB6.ShiftConstants.ShiftMask) > 0 Then
			
			If KeyCode = System.Windows.Forms.Keys.F1 Then
				HtmlHelp(0, modUtil.g_nspectPath & "\Help\nspect.chm", HH_DISPLAY_TOPIC, "project_setup.htm")
			End If
			
			If KeyCode = Shift + System.Windows.Forms.Keys.F7 Then
				frmAbout.ShowDialog()
			End If
			
		End If
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "Form_KeyDown " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	Private Sub frmPrj_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		On Error GoTo ErrorHandler
		
		
		'Check the DPI setting: This was put in place because the alignment of the checkboxes will be
		'messed up if DPI settings are anything other than Normal
		Dim lngMapDC As Integer
		Dim lngDPI As Integer
		lngMapDC = GetDC(Me.Handle.ToInt32)
		lngDPI = GetDeviceCaps(lngMapDC, LOGPIXELSX)
		ReleaseDC(Me.Handle.ToInt32, lngMapDC)
		
		If lngDPI <> 96 Then
			MsgBox("Warning: N-SPECT requires your font size to be 96 DPI." & vbNewLine & "Some controls may appear out of alignment on this form.", MsgBoxStyle.Critical, "Warning!")
		End If
		
		m_bolFirstLoad = True 'It's the first load
		m_booExists = False
		
		'Initialize the browse for folder vbAccelerator Reference
		'Set dlgBrowser = New cBrowseForFolder
		
		'define flexgrid for coefficients tab
		With grdCoeffs
			.col = .FixedCols
			.row = .FixedRows
			.Width = VB6.TwipsToPixelsX(7500)
			.set_ColWidth(0,  , 400)
			.set_ColWidth(1,  , 600)
			.set_ColWidth(2,  , 2200)
			.set_ColWidth(3,  , 2400)
			.set_ColWidth(4,  , 1800)
			.set_ColWidth(5,  , 0)
			.set_ColWidth(6,  , 0)
			.row = 0
			.col = 1
			.Text = "Apply"
			.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
			.col = 2
			.Text = "Pollutant Name"
			.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
			.col = 3
			.Text = "Coefficient Set"
			.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
			.col = 4
			.Text = "Which Coefficient"
			.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
			
		End With
		
		'define flexgrid for land cover change scenarios on Management Scenarios tab
		With grdLCChanges
			.col = .FixedCols
			.row = .FixedRows
			.Width = VB6.TwipsToPixelsX(6700)
			.set_ColWidth(0,  , 400)
			.set_ColWidth(1,  , 600)
			.set_ColWidth(2,  , 2800)
			.set_ColWidth(3,  , 2800)
			.row = 0
			.col = 1
			.Text = "Apply"
			.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
			.col = 2
			.Text = "Change area layer"
			.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
			.col = 3
			.Text = "Change to class"
			.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
		End With
		
		'define flexgrid for Point Sources tab
		With grdLU
			.col = .FixedCols
			.row = .FixedRows
			.Width = VB6.TwipsToPixelsX(5200 + (.GridLineWidth * (.get_Cols() + 1)) + 100)
			.set_ColWidth(0,  , 400)
			.set_ColWidth(1,  , 600)
			.set_ColWidth(2,  , 4200)
			.set_ColWidth(3,  , 0)
			.row = 0
			.col = 1
			.Text = "Apply"
			.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
			.col = 2
			.Text = "Land Use Scenario"
			.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
		End With
		
		'Fill the Form
		
		'ComboBox::LandCover Type
		modUtil.InitComboBox(cboLCType, "LCType")
		
		'ComboBox::Precipitation Scenarios
		modUtil.InitComboBox(cboPrecipScen, "PrecipScenario")
		cboPrecipScen.Items.Insert(cboPrecipScen.Items.Count, "New precipitation scenario...")
		
		'ComboBox::WaterShed Delineations
		modUtil.InitComboBox(cboWSDelin, "WSDelineation")
		cboWSDelin.Items.Insert(cboWSDelin.Items.Count, "New watershed delineation...")
		
		'ComboBox::WaterQuality Criteria
		modUtil.InitComboBox(cboWQStd, "WQCriteria")
		cboWQStd.Items.Insert(cboWQStd.Items.Count, "New water quality standard...")
		
		'Fill Land Cover cbo
		modUtil.AddRasterLayerToComboBox(cboLCLayer, m_pMap)
		
		'Fill Rain GRID cbo
		modUtil.AddRasterLayerToComboBox(cboRainGrid, m_pMap)
		
		'Soils, now a 'scenario', not just a datalayer
		modUtil.InitComboBox(cboSoilsLayer, "Soils")
		
		'Fill area
		modUtil.AddFeatureLayerToComboBox(cboAreaLayer, m_pMap, "poly")
		
		'Fill LandClass
		FillCboLCCLass()
		
		m_intMgmtCount = grdLCChanges.Rows - 1 'Number of mgmt scens
		m_intLUCount = grdLU.Rows - 1 'Number of landuses
		
		'Initialize parameter file
		m_XMLPrjParams = New clsXMLPrjFile
		
		Me.Text = "Untitled"
		
		'Find out what the deal is
		cboSelectPoly.Items.Clear()
		modUtil.AddFeatureLayerToComboBox(cboSelectPoly, m_pMap, "poly")
		
		chkSelectedPolys.Enabled = EnableChkWaterShed
		
		'Test workspace persistence
		If Len(m_strWorkspace) > 0 Then
			txtOutputWS.Text = m_strWorkspace
		End If
		
		
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	Private Sub cmdOutputBrowse_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOutputBrowse.Click
		On Error GoTo ErrHandler
		
		'browse...get output filename
		dlgXMLOpen.FileName = CStr(Nothing)
		dlgXMLSave.FileName = CStr(Nothing)
		'UPGRADE_WARNING: CommonDialog variable was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="671167DC-EA81-475D-B690-7A40C7BF4A23"'
		With dlgXML
			'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			.Filter = MSG6
			.Title = MSG7
			.FilterIndex = 1
			'UPGRADE_WARNING: FileOpenConstants constant FileOpenConstants.cdlOFNHideReadOnly was upgraded to OpenFileDialog.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
			'UPGRADE_WARNING: MSComDlg.CommonDialog property dlgXML.Flags was upgraded to dlgXMLOpen.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
			'UPGRADE_WARNING: FileOpenConstants constant FileOpenConstants.cdlOFNHideReadOnly was upgraded to OpenFileDialog.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
			.ShowReadOnly = False
			'UPGRADE_WARNING: MSComDlg.CommonDialog property dlgXML.Flags was upgraded to dlgXMLOpen.CheckFileExists which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
			.CheckFileExists = True
			.CheckPathExists = True
			.CheckPathExists = True
			'UPGRADE_WARNING: The CommonDialog CancelError property is not supported in Visual Basic .NET. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"'
			.CancelError = True
			.ShowDialog()
		End With
		
		If Len(dlgXMLOpen.FileName) > 0 Then
			txtOutputFile.Text = Trim(dlgXMLOpen.FileName)
			txtThemeName.Text = modUtil.SplitFileName((txtOutputFile.Text))
		End If
		
ErrHandler: 
		Exit Sub
		
	End Sub
	
	Private Sub cmdQuit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdQuit.Click
		On Error GoTo ErrorHandler
		
		Dim intvbYesNo As Short
		
		intvbYesNo = MsgBox("Do you want to save changes you made to " & Me.Text & "?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Exclamation, "N-SPECT")
		
		If intvbYesNo = MsgBoxResult.Yes Then
			If SaveXMLFile Then
				Me.Close()
			End If
		ElseIf intvbYesNo = MsgBoxResult.No Then 
			Me.Close()
		ElseIf intvbYesNo = MsgBoxResult.Cancel Then 
			Exit Sub
		End If
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "cmdQuit_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	Private Sub frmPrj_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		On Error GoTo ErrorHandler
		
		'Cleanup
		'UPGRADE_NOTE: Object m_pMap may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_pMap = Nothing
		'UPGRADE_NOTE: Object m_pActiveViewEvents may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_pActiveViewEvents = Nothing
		'UPGRADE_NOTE: Object m_pMxDoc may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_pMxDoc = Nothing
		'UPGRADE_NOTE: Object m_App may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_App = Nothing
		m_strOpenFileName = ""
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "Form_Unload " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	Private Sub grdCoeffs_MouseUpEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_MouseUpEvent) Handles grdCoeffs.MouseUpEvent
		
		m_intPollRow = grdCoeffs.row
		m_intPollCol = grdCoeffs.col
		
		
		If grdCoeffs.get_TextMatrix(m_intPollRow, 1) = "1" Then
			'We want to limit editing to only the coeff/coeff type columns
			If m_intPollCol = 4 And m_intPollRow >= 1 Then
				
				cboCoeffSet.Visible = False
				
				With cboCoeff
					.Visible = True
					
					.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(SSTab1.Left) + VB6.PixelsToTwipsX(grdCoeffs.Left) + grdCoeffs.CellLeft), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(SSTab1.Top) + VB6.PixelsToTwipsY(grdCoeffs.Top) + grdCoeffs.CellTop), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
					
					.Width = VB6.TwipsToPixelsX(grdCoeffs.CellWidth)
				End With
				
			ElseIf m_intPollCol = 3 And m_intPollRow >= 1 Then 
				
				cboCoeff.Visible = False
				
				With cboCoeffSet
					.Visible = True
					.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(SSTab1.Left) + VB6.PixelsToTwipsX(grdCoeffs.Left) + grdCoeffs.CellLeft), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(SSTab1.Top) + VB6.PixelsToTwipsY(grdCoeffs.Top) + grdCoeffs.CellTop), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
					.Width = VB6.TwipsToPixelsX(grdCoeffs.CellWidth)
					FillCboCoeffSet()
				End With
			Else
				cboCoeff.Visible = False
				cboCoeffSet.Visible = False
			End If
		Else
			cboCoeff.Visible = False
			cboCoeffSet.Visible = False
		End If
		
		
	End Sub
	
	
	Private Sub grdLCChanges_MouseDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_MouseDownEvent) Handles grdLCChanges.MouseDownEvent
		On Error GoTo ErrorHandler
		
		
		If eventArgs.Button = 2 Then
			'UPGRADE_ISSUE: Form method frmPrj.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			PopupMenu(mnuManagement)
		End If
		
		
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "grdLCChanges_MouseDown " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	
	Private Sub grdLCChanges_MouseUpEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_MouseUpEvent) Handles grdLCChanges.MouseUpEvent
		
		m_intLCRow = grdLCChanges.row
		m_intLCCol = grdLCChanges.col
		
		'We
		If m_intLCCol = 3 And m_intLCRow >= 1 Then
			
			cboAreaLayer.Visible = False
			
			With cboClass
				
				.Visible = True
				.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(SSTab1.Left) + VB6.PixelsToTwipsX(grdLCChanges.Left) + grdLCChanges.CellLeft), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(SSTab1.Top) + VB6.PixelsToTwipsY(grdLCChanges.Top) + grdLCChanges.CellTop), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
				
				.Width = VB6.TwipsToPixelsX(grdLCChanges.CellWidth)
			End With
			
		ElseIf m_intLCCol = 2 And m_intLCRow >= 1 Then 
			
			cboClass.Visible = False
			
			With cboAreaLayer
				.Visible = True
				.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(SSTab1.Left) + VB6.PixelsToTwipsX(grdLCChanges.Left) + grdLCChanges.CellLeft), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(SSTab1.Top) + VB6.PixelsToTwipsY(grdLCChanges.Top) + grdLCChanges.CellTop), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
				.Width = VB6.TwipsToPixelsX(grdLCChanges.CellWidth)
				
			End With
			
		Else
			
			cboClass.Visible = False
			cboAreaLayer.Visible = False
			
		End If
		
	End Sub
	
	Private Sub m_pActiveViewEvents_ItemAdded(ByVal Item As Object) Handles m_pActiveViewEvents.ItemAdded
		On Error GoTo ErrorHandler
		
		'Necessary to track items added/removed in this case to the Map, so cbos and such can update them darn selves
		
		Dim strLCLayer As String
		Dim strAreaLayer As String
		Dim strSelectPolyLayer As String
		Dim strRainLayer As String
		
		'Find out the current LcLayer selection
		strLCLayer = cboLCLayer.Text
		
		'Fill Land Cover cbo
		modUtil.AddRasterLayerToComboBox(cboLCLayer, m_pMap)
		
		'Return the cboLCLayer to original selection, if there was one
		If Len(strLCLayer) <> 0 Then
			cboLCLayer.SelectedIndex = modUtil.GetCboIndex(strLCLayer, cboLCLayer)
		End If
		
		'Fill Rain GRID cbo
		If cboRainGrid.Visible = True Then
			strRainLayer = cboRainGrid.Text
			modUtil.AddRasterLayerToComboBox(cboRainGrid, m_pMap)
			'Again, check for prior selection, if there was one, return to it
			If Len(strRainLayer) > 0 Then
				cboRainGrid.SelectedIndex = modUtil.GetCboIndex(strRainLayer, cboRainGrid)
			End If
			
		End If
		
		'Fill area
		strAreaLayer = cboAreaLayer.Text
		modUtil.AddFeatureLayerToComboBox(cboAreaLayer, m_pMap, "poly")
		
		If Len(strAreaLayer) <> 0 Then
			cboAreaLayer.SelectedIndex = modUtil.GetCboIndex(strAreaLayer, cboAreaLayer)
		End If
		
		'Fill SelectPolys
		strSelectPolyLayer = cboSelectPoly.Text
		modUtil.AddFeatureLayerToComboBox(cboSelectPoly, m_pMap, "poly")
		cboSelectPoly.Enabled = False
		
		If Len(strSelectPolyLayer) <> 0 Then
			cboSelectPoly.SelectedIndex = modUtil.GetCboIndex(strSelectPolyLayer, cboSelectPoly)
		End If
		
		chkSelectedPolys.Enabled = EnableChkWaterShed
		
		Exit Sub
ErrorHandler: 
		HandleError(False, "m_pActiveViewEvents_ItemAdded " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	
	Private Sub m_pActiveViewEvents_ItemDeleted(ByVal Item As Object) Handles m_pActiveViewEvents.ItemDeleted
		On Error GoTo ErrorHandler
		
		Dim pLayer As ESRI.ArcGIS.Carto.ILayer
		Dim strLCLayer As String
		Dim strAreaLayer As String
		Dim strSelectPolyLayer As String
		Dim strRainLayer As String
		
		pLayer = Item
		
		'Find out the current LcLayer selection
		strLCLayer = cboLCLayer.Text
		
		'Fill Land Cover cbo
		modUtil.AddRasterLayerToComboBox(cboLCLayer, m_pMap)
		
		'Return the cboLCLayer to original selection, if there was one, have to make sure
		'however that the item removed isn't the selected item
		If Len(strLCLayer) <> 0 Then
			If Not pLayer.Name = strLCLayer Then
				cboLCLayer.SelectedIndex = modUtil.GetCboIndex(strLCLayer, cboLCLayer)
			Else
				cboLCLayer.SelectedIndex = -1
			End If
		End If
		
		'Fill Rain GRID cbo
		If cboRainGrid.Visible = True Then
			strRainLayer = cboRainGrid.Text
			modUtil.AddRasterLayerToComboBox(cboRainGrid, m_pMap)
			'Again, check for prior selection, if there was one, return to it
			If Len(strRainLayer) > 0 Then
				If Not pLayer.Name = strRainLayer Then
					cboRainGrid.SelectedIndex = modUtil.GetCboIndex(strRainLayer, cboRainGrid)
				Else
					cboRainGrid.SelectedIndex = -1
				End If
			End If
			
		End If
		
		'Fill area
		strAreaLayer = cboAreaLayer.Text
		modUtil.AddFeatureLayerToComboBox(cboAreaLayer, m_pMap, "poly")
		
		If Len(strAreaLayer) <> 0 Then
			If Not pLayer.Name = strAreaLayer Then
				cboAreaLayer.SelectedIndex = modUtil.GetCboIndex(strAreaLayer, cboAreaLayer)
			Else
				cboAreaLayer.SelectedIndex = -1
			End If
		End If
		
		'Fill SelectPolys
		strSelectPolyLayer = cboSelectPoly.Text
		modUtil.AddFeatureLayerToComboBox(cboSelectPoly, m_pMap, "poly")
		
		If Len(strSelectPolyLayer) <> 0 Then
			If Not pLayer.Name = strSelectPolyLayer Then
				cboSelectPoly.SelectedIndex = modUtil.GetCboIndex(strSelectPolyLayer, cboSelectPoly)
			Else
				cboSelectPoly.SelectedIndex = -1
			End If
		End If
		
		chkSelectedPolys.Enabled = EnableChkWaterShed
		cboSelectPoly.Enabled = chkSelectedPolys.Enabled
		
		'Clean
		'UPGRADE_NOTE: Object pLayer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pLayer = Nothing
		
		
		
		Exit Sub
ErrorHandler: 
		HandleError(False, "m_pActiveViewEvents_ItemDeleted " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	Private Sub m_pActiveViewEvents_SelectionChanged() Handles m_pActiveViewEvents.SelectionChanged
		On Error GoTo ErrorHandler
		
		chkSelectedPolys.Enabled = EnableChkWaterShed
		
		Exit Sub
ErrorHandler: 
		HandleError(False, "m_pActiveViewEvents_SelectionChanged " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	Public Sub mnuExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuExit.Click
		On Error GoTo ErrorHandler
		
		Dim intvbYesNo As Short
		
		intvbYesNo = MsgBox("Do you want to save changes you made to " & Me.Text & "?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Exclamation, "N-SPECT")
		
		If intvbYesNo = MsgBoxResult.Yes Then
			If SaveXMLFile Then
				Me.Close()
			End If
		ElseIf intvbYesNo = MsgBoxResult.No Then 
			Me.Close()
		ElseIf intvbYesNo = MsgBoxResult.Cancel Then 
			Exit Sub
		End If
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "mnuExit_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	Public Sub mnuGeneralHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuGeneralHelp.Click
		On Error GoTo ErrorHandler
		
		'API call to help
		HtmlHelp(0, modUtil.g_nspectPath & "\Help\nspect.chm", HH_DISPLAY_TOPIC, "project_setup.htm")
		
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "mnuGeneralHelp_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	Public Sub mnuLUDelete_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLUDelete.Click
		On Error GoTo ErrorHandler
		
		'delete current row
		Dim row, R, C, col As Short
		
		With grdLU
			If .Rows > .FixedRows Then 'make sure we don't del header Rows
				For col = 1 To .get_Cols() - 1
					If ((Trim(.get_TextMatrix(.row, col)) > "" And col = 2) Or (.get_TextMatrix(.row, col) <> "0" And col = 1) Or (.get_TextMatrix(.row, col) <> "0" And col >= 3)) Then 'data?
						C = 1
						Exit For
					End If
				Next col
				If C Then
					R = MsgBox("There is data in Row" & Str(.row) & "! Delete anyway?", MsgBoxStyle.YesNo, "Delete Row!")
				End If
				If C = 0 Or R = MsgBoxResult.Yes Then 'no exist. data or YES
					If .row = .Rows - 1 And .Rows = 2 Then 'last row?
						'If .row = 1 And .Rows = 1 Then  'you want to leave 1 row, but empty it.
						.set_TextMatrix(.row, 1, 0)
						.set_TextMatrix(.row, 2, "")
						.set_TextMatrix(.row, 3, "")
					Else
						For row = .row To .Rows - 2 'move data up 1 row
							For col = 1 To .get_Cols() - 1
								.set_TextMatrix(row, col, .get_TextMatrix(row + 1, col))
							Next col
						Next row
						.Rows = .Rows - 1 'del last row
					End If
					
				End If
				
			End If
			
		End With
		
		m_intLUCount = grdLU.Rows
		
		m_intCount = grdLU.Rows
		ClearLUCheckBoxes(True, m_intCount + 1)
		CreateLUCheckBoxes(True)
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "mnuLUDelete_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	'Add row...
	Public Sub mnuManAppen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuManAppen.Click
		On Error GoTo ErrorHandler
		
		'add row to end of  grid
		With grdLCChanges
			.Rows = .Rows + 1
			.row = .Rows - 1
			.set_TextMatrix(.row, 1, "0")
		End With
		
		m_intMgmtCount = grdLCChanges.Rows
		
		CreateMgmtCheckBoxes(False, grdLCChanges.row)
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "mnuManAppen_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	Public Sub mnuManDelete_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuManDelete.Click
		On Error GoTo ErrorHandler
		
		'delete current row
		
		Dim row, R, C, col As Short
		
		With grdLCChanges
			
			If .Rows > .FixedRows Then 'make sure we don't del header Rows
				For col = 1 To .get_Cols() - 1
					If ((Trim(.get_TextMatrix(.row, col)) > "" And col = 2) Or (.get_TextMatrix(.row, col) <> "0" And col = 1) Or (.get_TextMatrix(.row, col) <> "0" And col >= 3)) Then 'data?
						C = 1
						Exit For
					End If
				Next col
				If C Then
					R = MsgBox("There is data in Row" & Str(.row) & " ! Delete anyway?", MsgBoxStyle.YesNo, "Delete Row!")
				End If
				If C = 0 Or R = MsgBoxResult.Yes Then 'no exist. data or YES
					If .row = .Rows - 1 Then 'last row?
						.row = .row - 1 'move active cell
						
					Else
						For row = .row To .Rows - 2 'move data up 1 row
							For col = 1 To .get_Cols() - 1
								.set_TextMatrix(row, col, .get_TextMatrix(row + 1, col))
							Next col
						Next row
					End If
					.Rows = .Rows - 1 'del last row
				Else
					Exit Sub
				End If
			End If
		End With
		
		m_intMgmtCount = grdLCChanges.Rows
		
		ClearMgmtCheckBoxes(True, m_intMgmtCount)
		CreateMgmtCheckBoxes(True)
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "mnuManDelete_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	Public Sub mnuManInsert_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuManInsert.Click
		On Error GoTo ErrorHandler
		
		'Insert row above current row in grdLChanges- Thanks, Andrew
		
		Dim row, R, col As Short
		With grdLCChanges
			If .row < .FixedRows Then 'make sure we don't insert above header Rows
				mnuManAppen_Click(mnuManAppen, New System.EventArgs())
			Else
				R = .row
				.Rows = .Rows + 1 'add a row
				
				For row = .Rows - 1 To R + 1 Step -1 'move data dn 1 row
					For col = 1 To .get_Cols() - 1
						.set_TextMatrix(row, col, .get_TextMatrix(row - 1, col))
					Next col
				Next row
				For col = 1 To .get_Cols() - 1 ' clear all cells in this row
					If (col = 1) Then
						.set_TextMatrix(R, col, "0")
					Else
						.set_TextMatrix(R, col, "")
					End If
				Next col
				
				
			End If
		End With
		
		
		m_intMgmtCount = grdLCChanges.Rows
		
		ClearMgmtCheckBoxes(True, m_intMgmtCount - 2)
		CreateMgmtCheckBoxes(True)
		
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "mnuManInsert_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	Public Sub MnuLUEdit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuLUEdit.Click
		On Error GoTo ErrorHandler
		
		mnuOptions.Enabled = False
		
		g_intManScenRow = CStr(m_intLURow)
		g_strLUScenFileName = grdLU.get_TextMatrix(m_intLURow, 3)
		
		With frmLUScen
			.init(m_App, (cboWQStd.Text))
			.Text = "Edit Land Use Scenario"
			VB6.ShowForm(frmLUScen, VB6.FormShowConstants.Modal, Me)
		End With
		
		Exit Sub
ErrorHandler: 
		HandleError(False, "MnuLUEdit_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	Public Sub MnuLUAdd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MnuLUAdd.Click
		On Error GoTo ErrorHandler
		
		Dim intRow As Short
		
		g_intManScenRow = CStr(m_intLURow)
		g_strLUScenFileName = grdLU.get_TextMatrix(m_intLURow, 3)
		
		'if there is data in the row....
		If g_strLUScenFileName <> "" Then
			
			grdLU.Rows = grdLU.Rows + 1 'Add a row
			g_strLUScenFileName = "" 'reset data
			g_intManScenRow = CStr(grdLU.Rows - 1) '
			intRow = CShort(g_intManScenRow)
			CreateLUCheckBoxes(False, intRow)
		Else
			g_intManScenRow = CStr(m_intLURow)
			g_strLUScenFileName = ""
		End If
		
		With frmLUScen
			.init(m_App, (cboWQStd.Text))
			.Text = "Add Land Use Scenario"
			VB6.ShowForm(frmLUScen, VB6.FormShowConstants.Modal, Me)
		End With
		
		
		Exit Sub
ErrorHandler: 
		HandleError(False, "MnuLUAdd_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	Private Sub grdLU_MouseDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_MouseDownEvent) Handles grdLU.MouseDownEvent
		On Error GoTo ErrorHandler
		
		Dim strName As String
		
		m_intLURow = grdLU.MouseRow
		m_intLUCol = grdLU.MouseCol
		
		'Set the popup to proper functionality, if current row, all for add, delete, edit
		'If empty row, disable edit, delete
		If eventArgs.Button = 2 And m_intLURow > 0 Then
			strName = grdLU.get_TextMatrix(m_intLURow, 2)
			If Len(strName) = 0 Then
				mnuLUEdit.Enabled = False
				mnuLUDelete.Enabled = False
				'UPGRADE_ISSUE: Form method frmPrj.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				PopupMenu(mnuOptions)
			Else
				mnuLUEdit.Enabled = True
				mnuLUDelete.Enabled = True
				'UPGRADE_ISSUE: Form method frmPrj.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				PopupMenu(mnuOptions)
			End If
		End If
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "grdLU_MouseDown " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	
	Public Sub mnuNew_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuNew.Click
		On Error GoTo ErrorHandler
		
		Dim intvbYesNo As Short
		
		intvbYesNo = MsgBox("Do you want to save changes you made to " & Me.Text & "?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Exclamation, "N-SPECT")
		
		If intvbYesNo = MsgBoxResult.Yes Then
			If SaveXMLFile Then
				ClearForm()
				frmPrj_Load(Me, New System.EventArgs())
			Else
				Exit Sub
			End If
		ElseIf intvbYesNo = MsgBoxResult.No Then 
			ClearForm()
			frmPrj_Load(Me, New System.EventArgs())
		ElseIf intvbYesNo = MsgBoxResult.Cancel Then 
			Exit Sub
		End If
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "mnuNew_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	Private Sub ClearForm()
		On Error GoTo ErrorHandler
		
		'Gotta clean up before new, clean form
		
		ClearCheckBoxes(True)
		ClearMgmtCheckBoxes(True, m_intMgmtCount)
		ClearLUCheckBoxes(True)
		
		'LandClass stuff
		cboLCLayer.Items.Clear()
		cboLCType.Items.Clear()
		
		'DBase scens
		cboPrecipScen.Items.Clear()
		cboWSDelin.Items.Clear()
		cboWQStd.Items.Clear()
		cboSoilsLayer.Items.Clear()
		
		'Text
		txtProjectName.Text = ""
		txtOutputWS.Text = ""
		
		'Checkboxes
		chkSelectedPolys.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkLocalEffects.CheckState = System.Windows.Forms.CheckState.Unchecked
		
		'Erosion
		cboRainGrid.Items.Clear()
		optUseGRID.Checked = True
		chkCalcErosion.CheckState = System.Windows.Forms.CheckState.Unchecked
		txtRainValue.Text = ""
		
		'txtOutputFile.Text = ""
		txtThemeName.Text = ""
		
		'clear the GRIDS
		grdLU.Clear()
		grdLU.Rows = 2
		grdLCChanges.Clear()
		grdLCChanges.Rows = 2
		grdCoeffs.Clear()
		
		
		
		Exit Sub
ErrorHandler: 
		HandleError(False, "ClearForm " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	
	Public Sub mnuOpen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuOpen.Click
		'On Error GoTo ErrorHandler
		
		LoadXMLFile()
		
		'Exit Sub
		'ErrorHandler:
		'HandleError True, "mnuOpen_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
	End Sub
	
	Public Sub mnuSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSave.Click
		On Error GoTo ErrorHandler
		
		SaveXMLFile()
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "mnuSave_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	Public Sub mnuSaveAs_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSaveAs.Click
		On Error GoTo ErrorHandler
		
		m_booExists = False
		SaveXMLFile()
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "mnuSaveAs_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	'UPGRADE_WARNING: Event optUseGRID.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optUseGRID_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optUseGRID.CheckedChanged
		If eventSender.Checked Then
			On Error GoTo ErrorHandler
			
			cboRainGrid.Enabled = optUseGRID.Checked
			txtRainValue.Enabled = optUseValue.Checked
			
			Exit Sub
ErrorHandler: 
			HandleError(True, "optUseGRID_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optUseValue.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optUseValue_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optUseValue.CheckedChanged
		If eventSender.Checked Then
			On Error GoTo ErrorHandler
			
			txtRainValue.Enabled = optUseValue.Checked
			cboRainGrid.Enabled = optUseGRID.Checked
			
			Exit Sub
ErrorHandler: 
			HandleError(True, "optUseValue_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
		End If
	End Sub
	
	'*********************************************************************************************************
	'FORM SPECIFIC FUNCTIONS/SUBS
	'*********************************************************************************************************
	Private Sub PopPollutants()
		On Error GoTo ErrorHandler
		
		
		Dim strSQLWQStd As String
		Dim rsWQStdCboClick As New ADODB.Recordset
		
		Dim strSQLWQStdPoll As String
		Dim rsWQStdPoll As New ADODB.Recordset
		
		Dim i As Short
		
		'Selection based on combo box
		strSQLWQStd = "SELECT * FROM WQCRITERIA WHERE NAME LIKE '" & cboWQStd.Text & "'"
		rsWQStdCboClick.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		rsWQStdCboClick.Open(strSQLWQStd, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
		
		If rsWQStdCboClick.RecordCount > 0 Then
			
			strSQLWQStdPoll = "SELECT POLLUTANT.NAME, POLL_WQCRITERIA.THRESHOLD " & "FROM POLL_WQCRITERIA INNER JOIN POLLUTANT " & "ON POLL_WQCRITERIA.POLLID = POLLUTANT.POLLID Where POLL_WQCRITERIA.WQCRITID = " & rsWQStdCboClick.Fields("WQCRITID").Value
			
			rsWQStdPoll.CursorLocation = ADODB.CursorLocationEnum.adUseClient
			rsWQStdPoll.Open(strSQLWQStdPoll, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
			
			grdCoeffs.Rows = rsWQStdPoll.RecordCount + 1
			
			For i = 1 To rsWQStdPoll.RecordCount
				
				grdCoeffs.set_TextMatrix(i, 2, rsWQStdPoll.Fields("Name").Value)
				grdCoeffs.set_TextMatrix(i, 5, rsWQStdPoll.Fields("Threshold").Value)
				rsWQStdPoll.MoveNext()
				
			Next i
			
			m_intCount = rsWQStdPoll.RecordCount
			
			'Clean it
			rsWQStdPoll.Close()
			'UPGRADE_NOTE: Object rsWQStdPoll may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsWQStdPoll = Nothing
			
			rsWQStdCboClick.Close()
			'UPGRADE_NOTE: Object rsWQStdCboClick may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsWQStdCboClick = Nothing
			
		Else
			
			MsgBox("Warning: There are no water quality standards remaining.  Please add a new one.", MsgBoxStyle.Critical, "Recordset Empty")
			'UPGRADE_NOTE: Object rsWQStdCboClick may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsWQStdCboClick = Nothing
		End If
		
		
		
		
		Exit Sub
ErrorHandler: 
		HandleError(False, "PopPollutants " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	Private Sub FillCboCoeffSet()
		On Error GoTo ErrorHandler
		
		Dim intCol As Short
		Dim strPollName As String
		Dim strSelectCoeff As String
		Dim rsCoeffSet As New ADODB.Recordset
		Dim i As Short
		
		cboCoeffSet.Items.Clear()
		
		intCol = m_intPollCol - 1
		strPollName = grdCoeffs.get_TextMatrix(m_intPollRow, intCol)
		
		strSelectCoeff = "SELECT POLLUTANT.POLLID, POLLUTANT.NAME, COEFFICIENTSET.NAME AS NAME2 FROM POLLUTANT INNER JOIN COEFFICIENTSET " & "ON POLLUTANT.POLLID = COEFFICIENTSET.POLLID Where POLLUTANT.NAME LIKE '" & strPollName & "'"
		
		rsCoeffSet.Open(strSelectCoeff, g_ADOConn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		
		For i = 0 To rsCoeffSet.RecordCount - 1
			cboCoeffSet.Items.Insert(i, rsCoeffSet.Fields("Name2").Value)
			rsCoeffSet.MoveNext()
		Next i
		
		'cleanup
		rsCoeffSet.Close()
		'UPGRADE_NOTE: Object rsCoeffSet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsCoeffSet = Nothing
		
		Exit Sub
ErrorHandler: 
		HandleError(False, "FillCboCoeffSet " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	Private Sub FillCboLCCLass()
		On Error GoTo ErrorHandler
		
		
		Dim strLCChanges As String
		Dim rsLCChanges As New ADODB.Recordset
		Dim i As Short
		
		strLCChanges = "SELECT LCCLASS.Name as Name2, LCTYPE.LCTYPEID FROM LCTYPE INNER JOIN LCCLASS " & "ON LCTYPE.LCTYPEID = LCCLASS.LCTYPEID WHERE LCTYPE.NAME LIKE '" & cboLCType.Text & "'"
		
		rsLCChanges.Open(strLCChanges, g_ADOConn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		
		rsLCChanges.MoveFirst()
		
		cboClass.Items.Clear()
		
		For i = 0 To rsLCChanges.RecordCount - 1
			cboClass.Items.Insert(i, rsLCChanges.Fields("Name2").Value)
			rsLCChanges.MoveNext()
		Next i
		
		rsLCChanges.Close()
		'UPGRADE_NOTE: Object rsLCChanges may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsLCChanges = Nothing
		
		Exit Sub
ErrorHandler: 
		HandleError(False, "FillCboLCCLass " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	Private Sub CreateCheckBoxes(ByRef booAll As Boolean, ByRef booUser As Boolean, Optional ByRef intRecNo As Short = 0)
		On Error GoTo ErrorHandler
		
		'booAll:  making all new ones
		'booUser:  flag set to determine if the boxes are being made during loading of a project file
		'option intRecNo: creation of 1 box
		
		Dim i As Short
		Dim j As Short
		Dim k As Short
		Dim strChkName As String
		j = 1
		i = 1
		
		SSTab1.SelectedIndex = 0
		
		If booAll Then
			For i = 1 To grdCoeffs.Rows - 1
				
				grdCoeffs.row = i
				grdCoeffs.col = 1
				
				'Set the alignment to center
				grdCoeffs.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
				
				k = i
				
				'UPGRADE_ISSUE: Controls method Controls.Add was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				Call Controls.Add("VB.CheckBox", "chk" & CStr(k), Me)
				
				chkIgnore.Load(k)
				
				chkIgnore(k).Parent = Me
				With chkIgnore(k)
					.Visible = True
					.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(SSTab1.Top) + VB6.PixelsToTwipsY(grdCoeffs.Top) + grdCoeffs.CellTop)
					.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(grdCoeffs.Left)) + (grdCoeffs.CellLeft) + (VB6.PixelsToTwipsX(SSTab1.Left)) + (grdCoeffs.CellWidth * 0.4))
					.Height = VB6.TwipsToPixelsY(195)
					.Width = VB6.TwipsToPixelsX(195)
					
					If booUser Then 'if during load event of project file...
						'If Threshold > 0 or The User has choosen to ignore then...
						If (CShort(grdCoeffs.get_TextMatrix(i, 5)) <> 0) And CDbl(grdCoeffs.get_TextMatrix(i, 1)) = 1 Then
							.CheckState = System.Windows.Forms.CheckState.Checked
							'.Enabled = False
						Else
							.CheckState = System.Windows.Forms.CheckState.Unchecked
							'.Enabled = False
						End If
					Else
						If (CShort(CDbl(grdCoeffs.get_TextMatrix(i, 5)) > 0)) Then
							.CheckState = System.Windows.Forms.CheckState.Unchecked
							.Enabled = True
						Else
							.CheckState = System.Windows.Forms.CheckState.Unchecked
							.Enabled = False
						End If
					End If
				End With
				
				grdCoeffs.set_TextMatrix(i, 1, CStr(chkIgnore(i).CheckState))
				Call Controls.RemoveAt("chk" & CStr(k))
			Next i
		End If
		
		
		
		Exit Sub
ErrorHandler: 
		HandleError(False, "CreateCheckBoxes " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	Private Sub ClearCheckBoxes(ByRef booAll As Boolean, Optional ByRef intRecNo As Short = 0)
		On Error GoTo ErrorHandler
		
		Dim k As Short
		
		SSTab1.SelectedIndex = 0
		
		If booAll Then
			
			'UPGRADE_WARNING: Couldn't resolve default property of object chkIgnore().UBound. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For k = 1 To chkIgnore().UBound
				chkIgnore.Unload(k)
			Next k
		Else
			chkIgnore.Unload(intRecNo)
		End If
		
		
		
		Exit Sub
ErrorHandler: 
		HandleError(False, "ClearCheckBoxes " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	Private Sub CreateMgmtCheckBoxes(ByRef booAll As Boolean, Optional ByRef intRecNo As Short = 0)
		On Error GoTo ErrorHandler
		
		
		Dim i As Short
		Dim j As Short
		Dim k As Short
		Dim strChkName As String
		j = 1
		i = 1
		
		SSTab1.SelectedIndex = 3
		
		If booAll Then
			For i = 1 To grdLCChanges.Rows - 1
				
				grdLCChanges.row = i
				grdLCChanges.col = 1
				
				'Set the alignment to center
				grdLCChanges.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
				
				k = i
				
				'UPGRADE_ISSUE: Controls method Controls.Add was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				Call Controls.Add("VB.CheckBox", "chkmgmt" & CStr(k), Me)
				
				chkIgnoreMgmt.Load(k)
				chkIgnoreMgmt(k).Parent = Me
				
				With chkIgnoreMgmt(k)
					.Visible = True
					.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(SSTab1.Top) + VB6.PixelsToTwipsY(grdLCChanges.Top) + grdLCChanges.CellTop)
					.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(grdLCChanges.Left) + (grdLCChanges.CellLeft) + (VB6.PixelsToTwipsX(SSTab1.Left)) + (grdLCChanges.CellWidth * 0.4))
					.Height = VB6.TwipsToPixelsY(195)
					.Width = VB6.TwipsToPixelsX(195)
					
					If grdLCChanges.get_TextMatrix(k, 1) <> "" Then
						.CheckState = CShort(grdLCChanges.get_TextMatrix(k, 1))
					Else
						.CheckState = System.Windows.Forms.CheckState.Unchecked
					End If
					grdLCChanges.set_TextMatrix(k, 1, chkIgnoreMgmt(k).CheckState)
				End With
				Call Controls.RemoveAt("chkmgmt" & CStr(k))
			Next i
		Else
			With grdLCChanges
				.row = intRecNo
				.col = 1
				.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
			End With
			
			k = intRecNo
			
			'UPGRADE_ISSUE: Controls method Controls.Add was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Call Controls.Add("VB.CheckBox", "chkmgmt" & CStr(k), Me)
			
			chkIgnoreMgmt.Load(k)
			chkIgnoreMgmt(k).Parent = Me
			
			With chkIgnoreMgmt(k)
				.Visible = True
				.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(SSTab1.Top) + VB6.PixelsToTwipsY(grdLCChanges.Top) + grdLCChanges.CellTop)
				.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(SSTab1.Left) + VB6.PixelsToTwipsX(grdLCChanges.Left) + grdLCChanges.CellLeft + (grdLCChanges.CellWidth * 0.4))
				.Height = VB6.TwipsToPixelsY(195)
				.Width = VB6.TwipsToPixelsX(195)
				.CheckState = CShort(grdLCChanges.get_TextMatrix(k, 1))
			End With
			Call Controls.RemoveAt("chkmgmt" & CStr(k))
		End If
		
		grdLCChanges.set_TextMatrix(0, 0, "")
		
		
		
		Exit Sub
ErrorHandler: 
		HandleError(False, "CreateMgmtCheckBoxes " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	Private Sub ClearMgmtCheckBoxes(ByRef booAll As Boolean, ByRef intCount As Short, Optional ByRef intRecNo As Short = 0)
		On Error GoTo ErrorHandler
		
		
		Dim k As Short
		
		SSTab1.SelectedIndex = 3
		
		If booAll Then
			
			'UPGRADE_WARNING: Couldn't resolve default property of object chkIgnoreMgmt().UBound. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For k = 1 To chkIgnoreMgmt().UBound
				chkIgnoreMgmt.Unload(k)
			Next k
		Else
			chkIgnoreMgmt.Unload(intRecNo)
		End If
		
		
		
		Exit Sub
ErrorHandler: 
		HandleError(False, "ClearMgmtCheckBoxes " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	Private Sub CreateLUCheckBoxes(ByRef booAll As Boolean, Optional ByRef intRecNo As Short = 0)
		On Error GoTo ErrorHandler
		
		
		Dim i As Short
		Dim j As Short
		Dim k As Short
		Dim strChkName As String
		j = 1
		i = 1
		
		SSTab1.SelectedIndex = 2
		
		If booAll Then
			For i = 1 To grdLU.Rows - 1
				
				grdLU.row = i
				grdLU.col = 1
				
				'Set the alignment to center
				grdLU.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
				
				k = i
				
				'UPGRADE_ISSUE: Controls method Controls.Add was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				Call Controls.Add("VB.CheckBox", "chklu" & CStr(k), Me)
				
				chkIgnoreLU.Load(k)
				
				chkIgnoreLU(k).Parent = Me
				With chkIgnoreLU(k)
					.Visible = True
					.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(SSTab1.Top) + VB6.PixelsToTwipsY(grdLU.Top) + grdLU.CellTop)
					.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(grdLU.Left) + (grdLU.CellLeft) + (VB6.PixelsToTwipsX(SSTab1.Left)) + (grdLU.CellWidth * 0.4))
					.Height = VB6.TwipsToPixelsY(195)
					.Width = VB6.TwipsToPixelsX(195)
					If grdLU.get_TextMatrix(i, 1) <> "" Then
						.CheckState = CShort(grdLU.get_TextMatrix(i, 1))
					Else
						.CheckState = System.Windows.Forms.CheckState.Unchecked
					End If
				End With
				grdLU.set_TextMatrix(k, 1, chkIgnoreLU(k).CheckState)
				Call Controls.RemoveAt("chklu" & CStr(k))
			Next i
		Else
			With grdLU
				.row = intRecNo
				.col = 1
				.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
			End With
			
			k = intRecNo
			
			'UPGRADE_ISSUE: Controls method Controls.Add was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Call Controls.Add("VB.CheckBox", "chk" & CStr(k), Me)
			chkIgnoreLU.Load(k)
			chkIgnoreLU(k).Parent = Me
			With chkIgnoreLU(k)
				.Visible = True
				.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(SSTab1.Top) + VB6.PixelsToTwipsY(grdLU.Top) + grdLU.CellTop)
				.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(grdLU.Left) + (grdLU.CellLeft) + (VB6.PixelsToTwipsX(SSTab1.Left)) + (grdLU.CellWidth * 0.4))
				.Height = VB6.TwipsToPixelsY(195)
				.Width = VB6.TwipsToPixelsX(195)
				.CheckState = System.Windows.Forms.CheckState.Unchecked
			End With
			grdLU.set_TextMatrix(k, 1, chkIgnoreLU(k).CheckState)
			Call Controls.RemoveAt("chk" & CStr(k))
		End If
		
		
		
		
		Exit Sub
ErrorHandler: 
		HandleError(False, "CreateLUCheckBoxes " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	Private Sub ClearLUCheckBoxes(ByRef booAll As Boolean, Optional ByRef intRecNo As Short = 0)
		On Error GoTo ErrorHandler
		
		
		Dim k As Short
		
		SSTab1.SelectedIndex = 2
		
		If booAll Then
			'UPGRADE_WARNING: Couldn't resolve default property of object chkIgnoreLU().UBound. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For k = 1 To chkIgnoreLU().UBound
				chkIgnoreLU.Unload(k)
			Next k
		Else
			chkIgnoreLU.Unload(intRecNo)
		End If
		
		
		
		Exit Sub
ErrorHandler: 
		HandleError(False, "ClearLUCheckBoxes " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	Private Sub SSTab1_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles SSTab1.SelectedIndexChanged
		Static PreviousTab As Short = SSTab1.SelectedIndex()
		On Error GoTo ErrorHandler
		
		
		Dim i As Short
		
		Select Case SSTab1.SelectedIndex
			
			Case 0
				For i = 0 To chkIgnore.UBound
					chkIgnore(i).Visible = True
				Next i
				
				For i = 0 To chkIgnoreMgmt.UBound
					chkIgnoreMgmt(i).Visible = False
				Next i
				
				For i = 0 To chkIgnoreLU.UBound
					chkIgnoreLU(i).Visible = False
				Next i
				
				cboAreaLayer.Visible = False
				cboClass.Visible = False
				
			Case 1
				For i = 0 To chkIgnore.UBound
					chkIgnore(i).Visible = False
				Next i
				
				For i = 0 To chkIgnoreMgmt.UBound
					chkIgnoreMgmt(i).Visible = False
				Next i
				
				For i = 0 To chkIgnoreLU.UBound
					chkIgnoreLU(i).Visible = False
				Next i
				
				cboCoeff.Visible = False
				cboCoeffSet.Visible = False
				cboAreaLayer.Visible = False
				cboClass.Visible = False
				
			Case 2
				For i = 0 To chkIgnore.UBound
					chkIgnore(i).Visible = False
				Next i
				
				For i = 0 To chkIgnoreMgmt.UBound
					chkIgnoreMgmt(i).Visible = False
				Next i
				
				For i = 0 To chkIgnoreLU.UBound
					chkIgnoreLU(i).Visible = True
				Next i
				
				cboCoeff.Visible = False
				cboCoeffSet.Visible = False
				cboAreaLayer.Visible = False
				cboClass.Visible = False
				
			Case 3
				For i = 0 To chkIgnore.UBound
					chkIgnore(i).Visible = False
				Next i
				
				For i = 0 To chkIgnoreMgmt.UBound
					chkIgnoreMgmt(i).Visible = True
				Next i
				
				For i = 0 To chkIgnoreLU.UBound
					chkIgnoreLU(i).Visible = False
				Next i
				
				cboCoeff.Visible = False
				cboCoeffSet.Visible = False
				cboAreaLayer.Visible = False
				cboClass.Visible = False
				
		End Select
		
		
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "SSTab1_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
		PreviousTab = SSTab1.SelectedIndex()
	End Sub
	
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
		On Error GoTo ErrorHandler
		
		'Use of the timer is necessary to create the checkboxes for all Grids that require them
		
		Timer1.Enabled = False
		
		If m_bolFirstLoad Then
			
			CreateCheckBoxes(True, False)
			CreateMgmtCheckBoxes(True)
			CreateLUCheckBoxes(True)
			
			m_bolFirstLoad = False
			
		ElseIf m_booNew Then 
			
			ClearCheckBoxes(True)
			CreateCheckBoxes(True, False)
			
			ClearMgmtCheckBoxes(True, m_intMgmtCount)
			CreateMgmtCheckBoxes(True)
			ClearLUCheckBoxes(True, m_intLUCount)
			CreateLUCheckBoxes(True)
			
			
		End If
		
		SSTab1.SelectedIndex = 0
		m_booNew = False
		
		
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "Timer1_Timer " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	Private Function ValidateData() As Boolean
		On Error GoTo ErrorHandler
		
		'Time to rifle through the form ensuring kosher data across the board.
		
		Dim i As Short
		Dim j As Short
		Dim strUpdatePrecip As String
		Dim strMgmnt As String
		Dim pLayer As ESRI.ArcGIS.Carto.ILayer
		Dim fso As New Scripting.FileSystemObject
		
		Dim clsParamsPrj As clsXMLPrjFile 'Just a holder for the xml
		clsParamsPrj = New clsXMLPrjFile
		
		'First check Selected Watersheds
		If chkSelectedPolys.Enabled = True And chkSelectedPolys.CheckState = 1 Then
			If Len(cboSelectPoly.Text) > 0 Then
				pLayer = m_pMap.Layer(modUtil.GetLayerIndex((cboSelectPoly.Text), m_App))
				If modUtil.GetSelectedFeatureCount(pLayer, m_pMap) > 0 Then
					ValidateData = True
				End If
			Else
				MsgBox("You have chosen 'Selected watersheds only'.  Please select watersheds.", MsgBoxStyle.Critical, "No Selected Features Found")
				ValidateData = False
				Exit Function
			End If
		End If
		
		'Project Name
		If Len(txtProjectName.Text) > 0 Then
			clsParamsPrj.strProjectName = Trim(txtProjectName.Text)
		Else
			MsgBox("Please enter a name for this project.", MsgBoxStyle.Information, "Enter Name")
			txtProjectName.Focus()
			ValidateData = False
			Exit Function
		End If
		
		'Working Directory
		If (Len(txtOutputWS.Text) > 0) And fso.FolderExists(txtOutputWS.Text) Then
			clsParamsPrj.strProjectWorkspace = Trim(txtOutputWS.Text)
		Else
			MsgBox("Please choose a valid output working directory.", MsgBoxStyle.Information, "Choose Workspace")
			txtOutputWS.Focus()
			ValidateData = False
			Exit Function
		End If
		
		'LandCover
		If cboLCLayer.Text = "" Then
			MsgBox("Please select a Land Cover layer before continuing.", MsgBoxStyle.Information, "Select Land Cover Layer")
			cboLCLayer.Focus()
			ValidateData = False
			Exit Function
		Else
			If modUtil.LayerInMap((cboLCLayer.Text), m_pMap) Then
				clsParamsPrj.strLCGridName = cboLCLayer.Text
				clsParamsPrj.strLCGridFileName = modUtil.GetRasterFileName((cboLCLayer.Text), m_App)
				clsParamsPrj.strLCGridUnits = CStr(cboLCUnits.SelectedIndex)
			Else
				MsgBox("The Land Cover layer you have choosen is not in the current map frame.", MsgBoxStyle.Information, "Layer Not Found")
				ValidateData = False
				Exit Function
			End If
		End If
		
		'LC Type
		If cboLCType.Text = "" Then
			MsgBox("Please select a Land Class Type before continuing.", MsgBoxStyle.Information, "Select Land Class Type")
			cboLCType.Focus()
			ValidateData = False
			Exit Function
		Else
			clsParamsPrj.strLCGridType = cboLCType.Text
		End If
		
		'REMOVED 11/30/2007 based on the whole issue of landcover records for clipped images.
		'Now Check LandCover, its table and whether or not the # of records matches in the databaset
		'    If Not (CheckLandCoverFields(clsParamsPrj.strLCGridType, modUtil.ReturnRaster(clsParamsPrj.strLCGridFileName))) Then
		'        MsgBox "The number of land cover classes in your " & clsParamsPrj.strLCGridName & _
		''               " GRID do not match the number entered " & vbNewLine & _
		''                " in the " & clsParamsPrj.strLCGridType & " land cover type.  " & _
		''                "Please refer to 'Land Cover Types' in the Advanced Settings before proceeding.", _
		''                vbInformation, "Records Not Compatible"
		'        ValidateData = False
		'        Exit Function
		'    End If
		
		'Soils - use definition to find datasets, if there use, if not tell the user
		If cboSoilsLayer.Text = "" Then
			MsgBox("Please select a Soils definition before continuing.", MsgBoxStyle.Information, "Select Soils Layer")
			cboSoilsLayer.Focus()
			ValidateData = False
			Exit Function
		Else
			If modUtil.RasterExists((lblSoilsHyd.Text)) Then
				clsParamsPrj.strSoilsDefName = cboSoilsLayer.Text
				clsParamsPrj.strSoilsHydFileName = lblSoilsHyd.Text
			Else
				MsgBox("The hydrologic soils layer " & lblSoilsHyd.Text & " you have selected is missing.  Please check you soils definition.", MsgBoxStyle.Information, "Soils Layer Not Found")
				ValidateData = False
				Exit Function
			End If
		End If
		
		'PrecipScenario
		'If the layer is in the map, get out, all is well- m_strPrecipFile is established on the
		'PrecipCbo Click event
		If modUtil.LayerInMap(modUtil.SplitFileName(m_strPrecipFile), m_pMap) Then
			clsParamsPrj.strPrecipScenario = cboPrecipScen.Text
		Else
			'Check if you can add it, if so, all is well
			If modUtil.RasterExists(m_strPrecipFile) Then
				clsParamsPrj.strPrecipScenario = cboPrecipScen.Text
			Else
				'Can't find it...well, then send user to Browse
				MsgBox("Unable to find precip dataset: " & m_strPrecipFile & ".  Please Correct", MsgBoxStyle.Information, "Cannot Find Dataset")
				m_strPrecipFile = modUtil.BrowseForFileName("Raster", Me, "Browse for Precipitation Dataset...")
				'If new one found, then we must update DataBase
				If Len(m_strPrecipFile) > 0 Then
					strUpdatePrecip = "UPDATE PrecipScenario SET precipScenario.PrecipFileName = '" & m_strPrecipFile & "'" & "WHERE NAME = '" & cboPrecipScen.Text & "'"
					g_ADOConn.Execute(strUpdatePrecip)
					'Now we can set the xmlParams
					clsParamsPrj.strPrecipScenario = cboPrecipScen.Text
					'modUtil.AddRasterLayerToMapFromFileName m_strPrecipFile, m_pMap
				Else
					MsgBox("Invalid File.", MsgBoxStyle.Information, "Invalid File")
					cboPrecipScen.Focus()
					ValidateData = False
				End If
			End If
		End If
		
		'Go out to a separate function for this one...WaterShed
		If ValidateWaterShed Then
			clsParamsPrj.strWaterShedDelin = cboWSDelin.Text
		Else
			MsgBox("There is a problems with the selected Watershed Delineation.", MsgBoxStyle.Information, "Watershed Delineation")
			ValidateData = False
			Exit Function
		End If
		
		'Water Quality
		If Len(cboWQStd.Text) > 0 Then
			clsParamsPrj.strWaterQuality = cboWQStd.Text
		Else
			MsgBox("Please select a water quality standard.", MsgBoxStyle.Information, "Water Quality Standard Missing")
			ValidateData = False
			Exit Function
		End If
		
		'Checkboxes, straight up values
		clsParamsPrj.intLocalEffects = chkLocalEffects.CheckState
		
		'Theoreretically, user could open file that had selected sheds.
		If chkSelectedPolys.Enabled = True Then
			clsParamsPrj.intSelectedPolys = chkSelectedPolys.CheckState
			clsParamsPrj.strSelectedPolyFileName = modUtil.GetFeatureFileName((cboSelectPoly.Text), m_App)
			clsParamsPrj.strSelectedPolyLyrName = cboSelectPoly.Text
		Else
			clsParamsPrj.intSelectedPolys = 0
		End If
		
		'Erosion Tab
		'Calc Erosion checkbox
		clsParamsPrj.intCalcErosion = chkCalcErosion.CheckState
		
		If chkCalcErosion.CheckState Then
			If modUtil.RasterExists((lblKFactor.Text)) Then
				clsParamsPrj.strSoilsKFileName = lblKFactor.Text
			Else
				MsgBox("The K Factor soils dataset " & lblSoilsHyd.Text & " you have selected is missing.  Please check your soils definition.", MsgBoxStyle.Information, "Soils K Factor Not Found")
				ValidateData = False
				Exit Function
			End If
			
			'Check the Rainfall Factor grid objects.
			If frameRainFall.Visible = True Then
				
				If optUseGRID.Checked Then
					
					If Len(cboRainGrid.Text) > 0 And (InStr(1, cboRainGrid.Text, cboLCLayer.Text, 1) = 0) Then
						clsParamsPrj.intRainGridBool = 1
						clsParamsPrj.intRainConstBool = 0
						clsParamsPrj.strRainGridName = cboRainGrid.Text
						clsParamsPrj.strRainGridFileName = modUtil.GetRasterFileName((cboRainGrid.Text), m_App)
					Else
						MsgBox("Please choose a rainfall Grid.", MsgBoxStyle.Information, "Select Rainfall GRID")
						SSTab1.SelectedIndex = 1
						ValidateData = False
						Exit Function
						
					End If
					
				ElseIf optUseValue.Checked Then 
					
					If Not IsNumeric(txtRainValue.Text) Then
						MsgBox("Numbers Only for Rain Values.", MsgBoxStyle.Information, "Numbers Only Please")
						txtRainValue.Focus()
					Else
						If CDbl(txtRainValue.Text) < 0 Then
							MsgBox("Positive values only please for rainfall values.", MsgBoxStyle.Information, "Postive Values Only")
							txtRainValue.Focus()
						Else
							clsParamsPrj.intRainConstBool = 1
							clsParamsPrj.dblRainConstValue = CDbl(txtRainValue.Text)
							clsParamsPrj.strRainGridFileName = ""
						End If
					End If
					
				Else
					MsgBox("You must choose a rainfall factor.", MsgBoxStyle.Information, "Rainfall Factor Missing")
					ValidateData = False
					Exit Function
				End If
			End If
			
			'Soil Delivery Ratio
			'Added 12/03/07 to account for soil delivery ratio GRID, user can now provide.
			If chkSDR.CheckState = 1 Then
				If Len(txtSDRGRID.Text) > 0 Then
					If modUtil.RasterExists((txtSDRGRID.Text)) Then
						clsParamsPrj.intUseOwnSDR = 1
						clsParamsPrj.strSDRGridFileName = txtSDRGRID.Text
					Else
						MsgBox("SDR GRID " & txtSDRGRID.Text & " not found.", MsgBoxStyle.Information, "SDR GRID Not Found")
						ValidateData = False
						Exit Function
					End If
				Else
					MsgBox("Please select an SDR GRID.", MsgBoxStyle.Information, "SDR GRID Not Selected")
					ValidateData = False
					Exit Function
				End If
			Else
				clsParamsPrj.intUseOwnSDR = 0
				clsParamsPrj.strSDRGridFileName = txtSDRGRID.Text
			End If
			
		End If
		
		'Managment Scenarios
		If ValidateMgmtScenario Then
			For i = 1 To grdLCChanges.Rows - 1
				If Len(grdLCChanges.get_TextMatrix(i, 2)) > 0 Then 'if they have entered one, then go ahead and add
					clsParamsPrj.clsMgmtScenItem = New clsXMLMgmtScenItem
					clsParamsPrj.clsMgmtScenItem.intID = i
					clsParamsPrj.clsMgmtScenItem.intApply = CShort(grdLCChanges.get_TextMatrix(i, 1))
					clsParamsPrj.clsMgmtScenItem.strAreaName = grdLCChanges.get_TextMatrix(i, 2)
					clsParamsPrj.clsMgmtScenItem.strAreaFileName = modUtil.GetFeatureFileName(grdLCChanges.get_TextMatrix(i, 2), m_App)
					clsParamsPrj.clsMgmtScenItem.strChangeToClass = grdLCChanges.get_TextMatrix(i, 3)
					clsParamsPrj.clsMgmtScenHolder.Add(clsParamsPrj.clsMgmtScenItem)
				End If
			Next i
		Else
			ValidateData = False
			grdLCChanges.Focus()
			Exit Function
		End If
		
		'Pollutants
		If ValidatePollutants Then
			For i = 1 To grdCoeffs.Rows - 1
				'Adding a New Pollutantant Item to the Project file
				clsParamsPrj.clsPollItem = New clsXMLPollutantItem
				clsParamsPrj.clsPollItem.intID = i
				clsParamsPrj.clsPollItem.intApply = CShort(grdCoeffs.get_TextMatrix(i, 1))
				clsParamsPrj.clsPollItem.strPollName = grdCoeffs.get_TextMatrix(i, 2)
				clsParamsPrj.clsPollItem.strCoeffSet = grdCoeffs.get_TextMatrix(i, 3)
				clsParamsPrj.clsPollItem.strCoeff = grdCoeffs.get_TextMatrix(i, 4)
				clsParamsPrj.clsPollItem.intThreshold = CShort(grdCoeffs.get_TextMatrix(i, 5))
				If grdCoeffs.get_TextMatrix(i, 6) <> "" Then
					clsParamsPrj.clsPollItem.strTypeDefXMLFile = grdCoeffs.get_TextMatrix(i, 6)
				End If
				clsParamsPrj.clsPollItems.Add(clsParamsPrj.clsPollItem)
			Next i
		Else
			ValidateData = False
			grdCoeffs.Focus()
			Exit Function
		End If
		
		'Land Uses
		For i = 1 To grdLU.Rows - 1
			If Len(grdLU.get_TextMatrix(i, 2)) > 0 Then
				clsParamsPrj.clsLUItem = New clsXMLLandUseItem
				clsParamsPrj.clsLUItem.intID = i
				clsParamsPrj.clsLUItem.intApply = CShort(grdLU.get_TextMatrix(i, 1))
				clsParamsPrj.clsLUItem.strLUScenName = grdLU.get_TextMatrix(i, 2)
				clsParamsPrj.clsLUItem.strLUScenXMLFile = grdLU.get_TextMatrix(i, 3)
				clsParamsPrj.clsLUItems.Add(clsParamsPrj.clsLUItem)
			End If
		Next i
		
		'If it gets to here, all is well
		ValidateData = True
		
		m_XMLPrjParams.XML = clsParamsPrj.XML
		
		'Cleanup
		'UPGRADE_NOTE: Object pLayer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pLayer = Nothing
		
		Exit Function
ErrorHandler: 
		HandleError(False, "ValidateData " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Function
	
	'UPGRADE_WARNING: Event txtProjectName.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtProjectName_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtProjectName.TextChanged
		On Error GoTo ErrorHandler
		
		'Make title of form = to what the user types in
		Me.Text = txtProjectName.Text
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "txtProjectName_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Sub
	
	Private Function ValidatePollutants() As Boolean
		On Error GoTo ErrorHandler
		
		'Function to validate pollutants
		Dim i As Short
		
		For i = 1 To grdCoeffs.Rows - 1
			'if the user isn't ignoring the pollutant, then check values
			If CDbl(grdCoeffs.get_TextMatrix(i, 1)) = 1 Then
				If Len(grdCoeffs.get_TextMatrix(i, 3)) = 0 Then
					MsgBox("Please select a coefficient set for pollutant: " & grdCoeffs.get_TextMatrix(i, 2), MsgBoxStyle.Critical, "Coefficient Set Missing")
					ValidatePollutants = False
					Exit Function
				Else
					If Len(grdCoeffs.get_TextMatrix(i, 4)) = 0 Then
						MsgBox("Please select a coefficient for pollutant: " & grdCoeffs.get_TextMatrix(i, 2), MsgBoxStyle.Critical, "Coefficient Missing")
						ValidatePollutants = False
						Exit Function
					Else
						ValidatePollutants = True
					End If
				End If
			Else
				ValidatePollutants = True
			End If
		Next i
		
		Exit Function
ErrorHandler: 
		HandleError(False, "ValidatePollutants " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Function
	
	Private Function ValidateWaterShed() As Boolean
		On Error GoTo ErrorHandler
		
		'Validate the Watershed
		Dim rsWShed As New ADODB.Recordset
		Dim strWShed As String
		Dim booUpdate As Boolean
		
		Dim strDEM As String
		Dim strFlowDirFileName As String
		Dim strFlowAccumFileName As String
		Dim strFilledDEMFileName As String
		
		booUpdate = False
		
		'Select record from current cbo Selection
		strWShed = "SELECT * FROM WSDELINEATION WHERE NAME LIKE '" & cboWSDelin.Text & "'"
		rsWShed.Open(strWShed, g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
		
		'Check to make sure all datasets exist, if not
		'DEM
		If Not modUtil.RasterExists(rsWShed.Fields("DEMFileName").Value) Then
			MsgBox("Unable to locate DEM dataset: " & rsWShed.Fields("DEMFileName").Value & ".", MsgBoxStyle.Critical, "Missing Dataset")
			strDEM = modUtil.BrowseForFileName("Raster", Me, "Browse for DEM...")
			If Len(strDEM) > 0 Then
				rsWShed.Fields("DEMFileName").Value = strDEM
				booUpdate = True
			Else
				ValidateWaterShed = False
				Exit Function
			End If
			'WaterShed Delineation
		ElseIf Not modUtil.FeatureExists(rsWShed.Fields("wsfilename").Value) Then 
			MsgBox("Unable to locate Watershed dataset: " & rsWShed.Fields("wsfilename").Value & ".", MsgBoxStyle.Critical, "Missing Dataset")
			strWShed = modUtil.BrowseForFileName("Feature", Me, "Browse for Watershed Dataset...")
			If Len(strWShed) > 0 Then
				rsWShed.Fields("wsfilename").Value = strWShed
				booUpdate = True
			Else
				ValidateWaterShed = False
				Exit Function
			End If
			'Flow Direction
		ElseIf Not modUtil.RasterExists(rsWShed.Fields("FlowDirFileName").Value) Then 
			MsgBox("Unable to locate Flow Direction GRID: " & rsWShed.Fields("FlowDirFileName").Value & ".", MsgBoxStyle.Critical, "Missing Dataset")
			strFlowDirFileName = modUtil.BrowseForFileName("Raster", Me, "Browse for Flow Direction GRID...")
			If Len(strFlowDirFileName) > 0 Then
				rsWShed.Fields("FlowDirFileName").Value = strFlowDirFileName
				booUpdate = True
			Else
				ValidateWaterShed = False
				Exit Function
			End If
			'Flow Accumulation
		ElseIf Not modUtil.RasterExists(rsWShed.Fields("FlowAccumFileName").Value) Then 
			MsgBox("Unable to locate Flow Accumulation GRID: " & rsWShed.Fields("FlowAccumFileName").Value & ".", MsgBoxStyle.Critical, "Missing Dataset")
			strFlowAccumFileName = modUtil.BrowseForFileName("Raster", Me, "Browse for Flow Accumulation GRID...")
			If Len(strFlowAccumFileName) > 0 Then
				rsWShed.Fields("FlowAccumFileName").Value = strFlowAccumFileName
				booUpdate = True
			Else
				ValidateWaterShed = False
				Exit Function
			End If
			'Check for non-hydro correct GRIDS
		ElseIf rsWShed.Fields("HydroCorrected").Value = 0 Then 
			If Not modUtil.RasterExists(rsWShed.Fields("FilledDEMFileName").Value) Then
				MsgBox("Unable to locate the Filled DEM: " & rsWShed.Fields("FilledDEMFileName").Value & ".", MsgBoxStyle.Critical, "Missing Dataset")
				strFilledDEMFileName = modUtil.BrowseForFileName("Raster", Me, "Browse for Filled DEM...")
				If Len(strFilledDEMFileName) > 0 Then
					rsWShed.Fields("FilledDEMFileName").Value = strFilledDEMFileName
					booUpdate = True
				Else
					ValidateWaterShed = False
					Exit Function
				End If
			End If
		End If
		
		If booUpdate Then
			rsWShed.Update()
		End If
		
		ValidateWaterShed = True
		
		rsWShed.Close()
		'UPGRADE_NOTE: Object rsWShed may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsWShed = Nothing
		
		
		
		Exit Function
ErrorHandler: 
		HandleError(False, "ValidateWaterShed " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Function
	
	Private Function ValidateMgmtScenario() As Boolean
		On Error GoTo ErrorHandler
		
		
		Dim i As Short
		Dim j As Short
		
		For i = 1 To grdLCChanges.Rows - 1
			If grdLCChanges.get_TextMatrix(i, 1) <> "0" Then
				For j = 1 To grdLCChanges.get_Cols() - 1
					
					Select Case j
						Case 2
							If Len(grdLCChanges.get_TextMatrix(i, j)) > 0 Then
								If Not (modUtil.LayerInMap(grdLCChanges.get_TextMatrix(i, j), m_pMap)) Then
									ValidateMgmtScenario = False
									Exit Function
								End If
							End If
						Case 3
							If Len(grdLCChanges.get_TextMatrix(i, j)) > 0 Then
								If grdLCChanges.get_TextMatrix(i, j) = "" Then
									ValidateMgmtScenario = False
									MsgBox("Please select a land class in cell " & i & " ," & j, MsgBoxStyle.Critical, "Missing Value")
									grdLCChanges.row = i
									grdLCChanges.col = j
									Exit Function
								End If
							End If
					End Select
				Next j
			End If
		Next i
		
		ValidateMgmtScenario = True
		
		
		
		Exit Function
ErrorHandler: 
		HandleError(False, "ValidateMgmtScenario " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
	End Function
	
	Private Function EnableChkWaterShed() As Boolean
		
		Dim rsWShed As New ADODB.Recordset
		Dim strWShed As String
		
		On Error GoTo ErrHandler
		
		strWShed = "SELECT WSFILENAME FROM WSDELINEATION WHERE NAME LIKE '" & cboWSDelin.Text & "'"
		rsWShed.Open(strWShed, g_ADOConn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
		
		m_strWShed = modUtil.SplitFileName(rsWShed.Fields("wsfilename").Value)
		
		If m_pMap.SelectionCount > 0 Then
			EnableChkWaterShed = True
		Else
			EnableChkWaterShed = False
		End If
		
		rsWShed.Close()
		'UPGRADE_NOTE: Object rsWShed may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsWShed = Nothing
		
		Exit Function
		
ErrHandler: 
		EnableChkWaterShed = False
		
	End Function

    Private Sub cboLCLayer_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboWSDelin.KeyDown, cboWQStd.KeyDown, cboPrecipScen.KeyDown, cboLCUnits.KeyDown, cboLCType.KeyDown, cboLCLayer.KeyDown
        e.SuppressKeyPress = True
    End Sub
End Class