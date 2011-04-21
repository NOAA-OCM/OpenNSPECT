Option Strict Off
Option Explicit On
Friend Class frmPrecipitation
    Inherits System.Windows.Forms.Form
    '	' *************************************************************************************
    '	' *  Perot Systems Government Services
    '	' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
    '	' *  frmPrecip
    '	' *************************************************************************************
    '	' *  Description: Form for browsing and maintaining precipitation scenarios
    '	' *  within NSPECT
    '	' *
    '	' *  Called By:  clsPrecip
    '	' *************************************************************************************



    '	Private rsPrecipCboClick As ADODB.Recordset
    '	Private rsPrecipDelete As ADODB.Recordset
    '	Private m_App As ESRI.ArcGIS.Framework.IApplication
    '	Private boolChange As Boolean
    '	Private boolLoad As Boolean

    '	Public m_pInputPrecipDS As ESRI.ArcGIS.Geodatabase.IRasterDataset


    '	Public Sub init(ByVal pApp As ESRI.ArcGIS.Framework.IApplication)

    '		m_App = pApp

    '	End Sub


    '	'UPGRADE_WARNING: Event cboGridUnits.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    '	Private Sub cboGridUnits_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboGridUnits.SelectedIndexChanged

    '		EnableSave()

    '	End Sub

    '	'UPGRADE_WARNING: Event cboPrecipType.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    '	Private Sub cboPrecipType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboPrecipType.SelectedIndexChanged
    '		If Not boolLoad Then
    '			EnableSave()
    '		End If

    '	End Sub

    '	'UPGRADE_WARNING: Event cboPrecipUnits.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    '	Private Sub cboPrecipUnits_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboPrecipUnits.SelectedIndexChanged

    '		EnableSave()

    '	End Sub

    '	'UPGRADE_WARNING: Event cboScenName.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    '	Private Sub cboScenName_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboScenName.SelectedIndexChanged

    '		Dim strSQLPrecip As String
    '		rsPrecipCboClick = New ADODB.Recordset

    '		strSQLPrecip = "SELECT * FROM PRECIPSCENARIO WHERE NAME LIKE '" & cboScenName.Text & "'"
    '		rsPrecipCboClick.CursorLocation = ADODB.CursorLocationEnum.adUseClient
    '		rsPrecipCboClick.Open(strSQLPrecip, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

    '		'Populate the controls...
    '		txtDesc.Text = rsPrecipCboClick.Fields("Description").Value
    '		txtPrecipFile.Text = rsPrecipCboClick.Fields("PrecipFileName").Value

    '		cboGridUnits.SelectedIndex = CShort(rsPrecipCboClick.Fields("PrecipGridUnits").Value)
    '		cboPrecipUnits.SelectedIndex = CShort(rsPrecipCboClick.Fields("PrecipUnits").Value)
    '		cboTimePeriod.SelectedIndex = rsPrecipCboClick.Fields("Type").Value
    '		cboPrecipType.SelectedIndex = rsPrecipCboClick.Fields("PrecipType").Value

    '		If rsPrecipCboClick.Fields("Type").Value = 0 Then
    '			txtRainingDays.Text = rsPrecipCboClick.Fields("RainingDays").Value
    '		End If

    '		cmdSave.Enabled = False


    '	End Sub

    '	'UPGRADE_WARNING: Event cboTimePeriod.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    '	Private Sub cboTimePeriod_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboTimePeriod.SelectedIndexChanged

    '		If cboTimePeriod.SelectedIndex = 0 Then
    '			lblRainingDays.Visible = True
    '			txtRainingDays.Visible = True
    '		Else
    '			lblRainingDays.Visible = False
    '			txtRainingDays.Visible = False
    '		End If

    '	End Sub

    '	Private Sub cmdBrowseFile_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdBrowseFile.Click

    '		Dim pPrecipRasterDataset As ESRI.ArcGIS.Geodatabase.IRasterDataset
    '		Dim pDistUnit As ESRI.ArcGIS.Geometry.ILinearUnit
    '		Dim intUnit As Short
    '		Dim pProjCoord As ESRI.ArcGIS.Geometry.IProjectedCoordinateSystem

    '		On Error GoTo ErrHandler

    '		m_pInputPrecipDS = AddInputFromGxBrowserText(txtPrecipFile, "Choose Precipitation GRID", Me, 0)

    '		If m_pInputPrecipDS Is Nothing Then
    '			Exit Sub
    '		Else

    '			pPrecipRasterDataset = m_pInputPrecipDS

    '			If CheckSpatialReference(pPrecipRasterDataset) Is Nothing Then

    '				MsgBox("The GRID you have choosen has no spatial reference information.  Please define a projection before continuing.", MsgBoxStyle.Exclamation, "No Project Information Detected")
    '				txtPrecipFile.Text = ""
    '				Exit Sub

    '			Else

    '				pProjCoord = CheckSpatialReference(pPrecipRasterDataset)
    '				pDistUnit = pProjCoord.CoordinateUnit
    '				intUnit = pDistUnit.MetersPerUnit

    '				If intUnit = 1 Then
    '					cboPrecipUnits.SelectedIndex = 0
    '				Else
    '					cboPrecipUnits.SelectedIndex = 1
    '				End If

    '				cboPrecipUnits.Refresh()
    '			End If

    '		End If


    '		'UPGRADE_NOTE: Object pPrecipRasterDataset may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '		pPrecipRasterDataset = Nothing
    '		'UPGRADE_NOTE: Object pDistUnit may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '		pDistUnit = Nothing
    '		'UPGRADE_NOTE: Object pProjCoord may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '		pProjCoord = Nothing

    '		Exit Sub
    'ErrHandler: 
    '		Exit Sub

    '	End Sub

    '	Private Sub cmdSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSave.Click

    '		If CheckParams = True Then

    '			rsPrecipCboClick.Fields("Name").Value = cboScenName.Text
    '			rsPrecipCboClick.Fields("Description").Value = txtDesc.Text
    '			rsPrecipCboClick.Fields("PrecipFileName").Value = txtPrecipFile.Text
    '			rsPrecipCboClick.Fields("PrecipGridUnits").Value = cboGridUnits.SelectedIndex
    '			rsPrecipCboClick.Fields("PrecipUnits").Value = cboPrecipUnits.SelectedIndex
    '			rsPrecipCboClick.Fields("PrecipType").Value = cboPrecipType.SelectedIndex
    '			rsPrecipCboClick.Fields("Type").Value = cboTimePeriod.SelectedIndex

    '			If cboTimePeriod.SelectedIndex = 0 Then
    '				rsPrecipCboClick.Fields("RainingDays").Value = CShort(txtRainingDays.Text)
    '			Else
    '				rsPrecipCboClick.Fields("RainingDays").Value = 0
    '			End If

    '			rsPrecipCboClick.Update()
    '			boolChange = False

    '			MsgBox(cboScenName.Text & " saved successfully.", MsgBoxStyle.OKOnly, "Record Saved")
    '			Me.Close()

    '			'UPGRADE_NOTE: Object rsPrecipCboClick may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '			rsPrecipCboClick = Nothing

    '		Else

    '			Exit Sub

    '		End If

    '	End Sub



    '	Private Sub frmPrecip_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

    '		'load data
    '		boolLoad = True
    '		modUtil.InitComboBox(cboScenName, "PRECIPSCENARIO")
    '		cmdSave.Enabled = False
    '		boolChange = False
    '		boolLoad = False

    '	End Sub

    '	Private Sub cmdQuit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdQuit.Click

    '		Dim intSave As Object

    '		If boolChange Then
    '			'UPGRADE_WARNING: Couldn't resolve default property of object intSave. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '			intSave = MsgBox("You have made changes to this record, are you sure you want to quit?", MsgBoxStyle.YesNo, "Quit?")

    '			If intSave = MsgBoxResult.Yes Then
    '				Me.Close()
    '			ElseIf intSave = MsgBoxResult.No Then 

    '				Exit Sub
    '			End If
    '		Else
    '			Me.Close()
    '		End If



    '	End Sub

    '	Private Sub frmPrecip_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

    '		rsPrecipCboClick.Close()
    '		'UPGRADE_NOTE: Object rsPrecipCboClick may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '		rsPrecipCboClick = Nothing

    '	End Sub

    '	Public Sub mnuNewPrecip_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuNewPrecip.Click

    '		VB6.ShowForm(frmNewPrecip, VB6.FormShowConstants.Modal, Me)

    '	End Sub

    '	Public Sub mnuDelPrecip_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDelPrecip.Click

    '		Dim intAns As Object
    '		Dim strSQLPrecipDel As String
    '		Dim cntrl As System.Windows.Forms.Control

    '		strSQLPrecipDel = "SELECT * FROM PRECIPSCENARIO WHERE NAME LIKE '" & cboScenName.Text & "'"

    '		If Not (cboScenName.Text = "") Then
    '			'UPGRADE_WARNING: Couldn't resolve default property of object intAns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '			intAns = MsgBox("Are you sure you want to delete the precipitation scenario '" & VB6.GetItemString(cboScenName, cboScenName.SelectedIndex) & "'?", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Confirm Delete")
    '			'code to handle response
    '			If intAns = MsgBoxResult.Yes Then

    '				'Set up a delete rs and get rid of it
    '				rsPrecipDelete = New ADODB.Recordset
    '				rsPrecipDelete.CursorLocation = ADODB.CursorLocationEnum.adUseClient
    '				rsPrecipDelete.Open(strSQLPrecipDel, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockOptimistic)

    '				rsPrecipDelete.Delete(ADODB.AffectEnum.adAffectCurrent)
    '				rsPrecipDelete.Update()
    '				MsgBox(VB6.GetItemString(cboScenName, cboScenName.SelectedIndex) & " deleted.", MsgBoxStyle.OKOnly, "Record Deleted")

    '				'Clear everything, clean up form
    '				cboScenName.Items.Clear()
    '				'cboGridUnits.Clear
    '				'cboPrecipUnits.Clear

    '				txtDesc.Text = ""
    '				'txtDuration.Text = ""
    '				txtPrecipFile.Text = ""

    '				modUtil.InitComboBox(cboScenName, "PRECIPSCENARIO")

    '				Me.Refresh()

    '			ElseIf intAns = MsgBoxResult.No Then 
    '				Exit Sub
    '			End If
    '		Else
    '			MsgBox("Please select a Precipitation Scenario", MsgBoxStyle.Critical, "No Scenario Selected")
    '		End If


    '	End Sub

    '	Public Sub mnuPrecipHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPrecipHelp.Click

    '		HtmlHelp(0, modUtil.g_nspectPath & "\Help\nspect.chm", HH_DISPLAY_TOPIC, "precip.htm")

    '	End Sub

    '	Private Sub optAnnual_Click()
    '		EnableSave()
    '	End Sub

    '	'UPGRADE_WARNING: Event txtDesc.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    '	Private Sub txtDesc_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDesc.TextChanged

    '		EnableSave()
    '		txtDesc.Text = Replace(txtDesc.Text, "'", "")

    '	End Sub

    '	Private Sub EnableSave()

    '		cmdSave.Enabled = True
    '		boolChange = True

    '	End Sub

    '	Private Sub txtDuration_Change()

    '		EnableSave()

    '	End Sub

    '	'UPGRADE_WARNING: Event txtPrecipFile.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    '	Private Sub txtPrecipFile_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPrecipFile.TextChanged

    '		EnableSave()

    '	End Sub

    '	Private Function CheckParams() As Boolean

    '		'Check the inputs of the form, before saving
    '		If Len(txtDesc.Text) = 0 Then
    '			MsgBox("Please enter a description for this scenario", MsgBoxStyle.Critical, "Description Missing")
    '			txtDesc.Focus()
    '			CheckParams = False
    '			Exit Function
    '		End If

    '		If txtPrecipFile.Text = " " Or txtPrecipFile.Text = "" Then
    '			MsgBox("Please select a valid precipitation GRID before saving.", MsgBoxStyle.Critical, "GRID Missing")
    '			txtPrecipFile.Focus()
    '			CheckParams = False
    '			Exit Function
    '		End If

    '		If cboGridUnits.Text = "" Then
    '			MsgBox("Please select GRID units.", MsgBoxStyle.Critical, "Units Missing")
    '			cboGridUnits.Focus()
    '			CheckParams = False
    '			Exit Function
    '		End If

    '		If cboPrecipUnits.Text = "" Then
    '			MsgBox("Please select precipitation units.", MsgBoxStyle.Critical, "Units Missing")
    '			cboPrecipUnits.Focus()
    '			CheckParams = False
    '			Exit Function
    '		End If

    '		If Len(cboPrecipType.Text) = 0 Then
    '			MsgBox("Please select a Precipitation Type.", MsgBoxStyle.Critical, "Precipitation Type Missing")
    '			cboPrecipType.Focus()
    '			CheckParams = False
    '			Exit Function
    '		End If

    '		If Len(cboTimePeriod.Text) = 0 Then
    '			MsgBox("Please select a Time Period.", MsgBoxStyle.Critical, "Precipitation Time Period Missing")
    '			cboTimePeriod.Focus()
    '			CheckParams = False
    '			Exit Function
    '		End If

    '		If cboTimePeriod.SelectedIndex = 0 Then
    '			If Not IsNumeric(txtRainingDays.Text) Or Len(txtRainingDays.Text) = 0 Then
    '				MsgBox("Please enter a numeric value for Raining Days.", MsgBoxStyle.Critical, "Raining Days Value Incorrect")
    '				txtRainingDays.Focus()
    '				CheckParams = False
    '				Exit Function
    '			End If
    '		End If

    '		'if it got through all that, then set it to true
    '		CheckParams = True


    '	End Function

    '	'UPGRADE_WARNING: Event txtRainingDays.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    '	Private Sub txtRainingDays_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRainingDays.TextChanged

    '		EnableSave()

    '	End Sub
End Class