Option Strict Off
Option Explicit On
Friend Class frmWatershedDelin
    Inherits System.Windows.Forms.Form
    '' *************************************************************************************
    '' *  Perot Systems Government Services
    '' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
    '' *  frmWSDelin
    '' *************************************************************************************
    '' *  Description: Form for browsing and maintaining watershed delineation
    '' *  Scenarios within NSPECT
    '' *
    '' *  Called By:  clsWSDelin
    '' *************************************************************************************


    'Dim rsWSDelinLoad As ADODB.Recordset 'Recordset on load event
    'Dim rsWSDelinCboClick As ADODB.Recordset 'Recordset to pop controls
    'Dim rsWSDelinDelete As ADODB.Recordset 'Recordset for deletions

    'Private m_App As ESRI.ArcGIS.Framework.IApplication
    'Public m_pInputPrecipDS As ESRI.ArcGIS.Geodatabase.IRasterDataset

    '' Variables used by the Error handler function

    'Public Sub init(ByVal pApp As ESRI.ArcGIS.Framework.IApplication)

    '	m_App = pApp

    'End Sub


    ''UPGRADE_WARNING: Event cboWSDelin.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    'Private Sub cboWSDelin_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboWSDelin.SelectedIndexChanged

    '	'String and recordset
    '	Dim strSQLPrecip As String
    '	rsWSDelinCboClick = New ADODB.Recordset

    '	'Selection based on combo box
    '	strSQLPrecip = "SELECT * FROM WSDELINEATION WHERE NAME LIKE '" & cboWSDelin.Text & "'"
    '	rsWSDelinCboClick.CursorLocation = ADODB.CursorLocationEnum.adUseClient
    '	rsWSDelinCboClick.Open(strSQLPrecip, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

    '	'Check for records
    '	If rsWSDelinCboClick.RecordCount > 0 Then
    '		'Populate the controls...
    '		txtDEMFile.Text = rsWSDelinCboClick.Fields("DEMFileName").Value
    '		cboDEMUnits.SelectedIndex = rsWSDelinCboClick.Fields("DEMGridUnits").Value
    '		txtStream.Text = rsWSDelinCboClick.Fields("StreamFileName").Value & ""
    '		chkHydroCorr.CheckState = rsWSDelinCboClick.Fields("HydroCorrected").Value
    '		cboWSSize.SelectedIndex = rsWSDelinCboClick.Fields("SubWSSize").Value
    '		txtWSFile.Text = rsWSDelinCboClick.Fields("wsfilename").Value & ""
    '		txtFlowAccumGrid.Text = rsWSDelinCboClick.Fields("FlowAccumFileName").Value & ""
    '		txtLSGrid.Text = rsWSDelinCboClick.Fields("LSFileName").Value & ""
    '		'Clean it
    '		'UPGRADE_NOTE: Object rsWSDelinCboClick may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '		rsWSDelinCboClick = Nothing

    '	Else
    '		MsgBox("Warning: There are no watershed delineation scenarios remaining.  Please add a new one.", MsgBoxStyle.Critical, "Recordset Empty")
    '		'UPGRADE_NOTE: Object rsWSDelinCboClick may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '		rsWSDelinCboClick = Nothing
    '		Exit Sub
    '	End If
    'End Sub

    'Private Sub cmdQuit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdQuit.Click

    '	Me.Close()

    'End Sub

    'Private Sub frmWSDelin_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

    '	modUtil.InitComboBox(cboWSDelin, "WSDELINEATION")

    'End Sub

    'Public Sub mnuNewExist_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuNewExist.Click
    '	frmUserWShed.init(m_App)
    '	VB6.ShowForm(frmUserWShed, VB6.FormShowConstants.Modal, Me)
    'End Sub

    'Public Sub mnuNewWSDelin_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuNewWSDelin.Click

    '	frmNewWSDelin.init(m_App)
    '	VB6.ShowForm(frmNewWSDelin, VB6.FormShowConstants.Modal, Me)

    'End Sub

    'Public Sub mnuDelWSDelin_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDelWSDelin.Click 'Delete menu item

    '	Dim intAns As Object
    '	Dim strSQLWSDel As String
    '	Dim fldWSDelin As Scripting.FileSystemObject
    '	Dim strFolder As String

    '	strSQLWSDel = "SELECT * FROM WSDELINEATION WHERE NAME LIKE '" & cboWSDelin.Text & "'"

    '	If Not (cboWSDelin.Text = "") Then
    '		'UPGRADE_WARNING: Couldn't resolve default property of object intAns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '		intAns = MsgBox("Are you sure you want to delete the watershed delineation scenario '" & VB6.GetItemString(cboWSDelin, cboWSDelin.SelectedIndex) & "'?", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Confirm Delete")
    '		'code to handle response
    '		If intAns = MsgBoxResult.Yes Then

    '			'Set up a delete rs and get rid of it
    '			rsWSDelinDelete = New ADODB.Recordset
    '			rsWSDelinDelete.CursorLocation = ADODB.CursorLocationEnum.adUseClient
    '			rsWSDelinDelete.Open(strSQLWSDel, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockOptimistic)

    '			rsWSDelinDelete.Delete(ADODB.AffectEnum.adAffectCurrent)
    '			rsWSDelinDelete.Update()

    '			fldWSDelin = CreateObject("Scripting.FileSystemObject")
    '			strFolder = modUtil.g_nspectPath & "\wsdelin\" & cboWSDelin.Text
    '			If fldWSDelin.FolderExists(strFolder) Then
    '				fldWSDelin.DeleteFolder(strFolder, True)
    '			End If

    '			'Confirm
    '			MsgBox(VB6.GetItemString(cboWSDelin, cboWSDelin.SelectedIndex) & " deleted.", MsgBoxStyle.OKOnly, "Record Deleted")

    '			'Clear everything, clean up form
    '			cboWSDelin.Items.Clear()
    '			chkHydroCorr.CheckState = System.Windows.Forms.CheckState.Unchecked
    '			txtDEMFile.Text = ""
    '			txtStream.Text = ""
    '			txtWSFile.Text = ""
    '			txtFlowAccumGrid.Text = ""

    '			modUtil.InitComboBox(cboWSDelin, "WSDELINEATION")

    '			Me.Refresh()

    '		ElseIf intAns = MsgBoxResult.No Then 
    '			Exit Sub
    '		End If
    '	Else
    '		MsgBox("Please select a watershed delineation", MsgBoxStyle.Critical, "No Scenario Selected")
    '	End If

    '	'cleanup
    '	'UPGRADE_NOTE: Object fldWSDelin may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '	fldWSDelin = Nothing
    '	rsWSDelinDelete.Close()
    '	'UPGRADE_NOTE: Object rsWSDelinDelete may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '	rsWSDelinDelete = Nothing


    'End Sub


    'Public Sub mnuWSDelin_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuWSDelin.Click

    '	HtmlHelp(0, modUtil.g_nspectPath & "\Help\nspect.chm", HH_DISPLAY_TOPIC, "wsdelin.htm")

    'End Sub
End Class