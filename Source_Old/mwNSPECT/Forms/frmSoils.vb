Option Strict Off
Option Explicit On
Friend Class frmSoils
	Inherits System.Windows.Forms.Form
	
    'Dim rsSoilsLoad As ADODB.Recordset
    'Dim rsSoilsCboClick As ADODB.Recordset
    'Dim rsSoilsDelete As ADODB.Recordset
    'Public m_App As ESRI.ArcGIS.Framework.IApplication

    'Public Sub init(ByVal pApp As ESRI.ArcGIS.Framework.IApplication)

    '	m_App = pApp

    'End Sub


    ''UPGRADE_WARNING: Event cboSoils.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    'Private Sub cboSoils_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboSoils.SelectedIndexChanged

    '	Dim strSQLSoils As String
    '	rsSoilsCboClick = New ADODB.Recordset

    '	strSQLSoils = "SELECT * FROM SOILS WHERE NAME LIKE '" & cboSoils.Text & "'"
    '	rsSoilsCboClick.CursorLocation = ADODB.CursorLocationEnum.adUseClient
    '	rsSoilsCboClick.Open(strSQLSoils, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

    '	'Populate the controls...
    '	txtSoilsGrid.Text = rsSoilsCboClick.Fields("SoilsFileName").Value
    '	txtSoilsKGrid.Text = rsSoilsCboClick.Fields("SoilsKFileName").Value

    'End Sub

    'Private Sub cmdQuit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdQuit.Click

    '	Me.Close()

    'End Sub

    'Private Sub cmdSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSave.Click
    '	Me.Close()
    'End Sub

    'Private Sub frmSoils_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
    '	'load data
    '	modUtil.InitComboBox(cboSoils, "SOILS")

    'End Sub

    'Public Sub mnuDelete_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDelete.Click

    '	Dim intAns As Object
    '	Dim strSQLSoilsDel As String
    '	Dim cntrl As System.Windows.Forms.Control

    '	strSQLSoilsDel = "SELECT * FROM SOILS WHERE NAME LIKE '" & cboSoils.Text & "'"

    '	If Not (cboSoils.Text = "") Then
    '		'UPGRADE_WARNING: Couldn't resolve default property of object intAns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '		intAns = MsgBox("Are you sure you want to delete the soils setup '" & VB6.GetItemString(cboSoils, cboSoils.SelectedIndex) & "'?", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Confirm Delete")
    '		'code to handle response
    '		If intAns = MsgBoxResult.Yes Then

    '			'Set up a delete rs and get rid of it
    '			rsSoilsDelete = New ADODB.Recordset
    '			rsSoilsDelete.CursorLocation = ADODB.CursorLocationEnum.adUseClient
    '			rsSoilsDelete.Open(strSQLSoilsDel, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockOptimistic)

    '			rsSoilsDelete.Delete(ADODB.AffectEnum.adAffectCurrent)
    '			rsSoilsDelete.Update()
    '			MsgBox(VB6.GetItemString(cboSoils, cboSoils.SelectedIndex) & " deleted.", MsgBoxStyle.OKOnly, "Record Deleted")

    '			'Clear everything, clean up form
    '			cboSoils.Items.Clear()

    '			txtSoilsGrid.Text = ""
    '			txtSoilsKGrid.Text = ""

    '			modUtil.InitComboBox(cboSoils, "SOILS")

    '			Me.Refresh()

    '		ElseIf intAns = MsgBoxResult.No Then 
    '			Exit Sub
    '		End If
    '	Else
    '		MsgBox("Please select a Soils Setup", MsgBoxStyle.Critical, "No Soils Setup Selected")
    '	End If


    'End Sub

    'Public Sub mnuNew_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuNew.Click

    '	frmSoilsSetup.init(m_App)
    '	VB6.ShowForm(frmSoilsSetup, VB6.FormShowConstants.Modal, Me)

    'End Sub




    'Public Sub mnuSoilsHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSoilsHelp.Click

    '	HtmlHelp(0, modUtil.g_nspectPath & "\Help\nspect.chm", HH_DISPLAY_TOPIC, "soils.htm")

    'End Sub
End Class