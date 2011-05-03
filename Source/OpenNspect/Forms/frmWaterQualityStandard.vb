Option Strict Off
Option Explicit On
Friend Class frmWaterQualityStandard
    Inherits System.Windows.Forms.Form
    '	' *************************************************************************************
    '	' *  Perot Systems Government Services
    '	' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
    '	' *  frmWQStd
    '	' *************************************************************************************
    '	' *  Description: Form for browsing and maintaining water quality standards
    '	' *  within NSPECT
    '	' *
    '	' *  Called By:  clsWQStd
    '	' *************************************************************************************
    '	Private rsWQStdLoad As ADODB.Recordset 'Load recordset
    '	Private rsWQStdCboClick As ADODB.Recordset 'cbo Click event
    '	Private rsWQStdPoll As ADODB.Recordset 'Pollutants of current standard
    '	Private rsWQStdDelete As ADODB.Recordset 'RS to be deleted

    '	Private m_App As ESRI.ArcGIS.Framework.IApplication
    '	Private m_strFileName As String
    '	Private strUndoText As String

    '	Private m_strUndoText As String

    '	Private intRow As Short 'Current Row
    '	Private intCol As Short 'Current Col.

    '	Private m_intWQRow As Short
    '	Private m_intWQCol As Short

    '	Private m_bolChange As Boolean 'Have records changed?
    '	Private intNumChanges As Short 'Counter for changes

    '	' Variables used by the Error handler function - DO NOT REMOVE
    '	Const c_sModuleFileName As String = "frmWQStd.frm"

    '	Public Sub init(ByVal pApp As ESRI.ArcGIS.Framework.IApplication)
    '		On Error GoTo ErrorHandler

    '		m_App = pApp

    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(True, "init " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub


    '	'UPGRADE_WARNING: Event cboWQStdName.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    '	Private Sub cboWQStdName_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboWQStdName.SelectedIndexChanged
    '		On Error GoTo ErrorHandler


    '		Dim intYesNo As Short

    '		If m_bolChange Then
    '			intYesNo = MsgBox("You have made changes to the data.  Would you like to save before coninuing?", MsgBoxStyle.YesNo, "Save Changes?")
    '			If intYesNo = MsgBoxResult.Yes Then
    '				UpdateData()
    '				m_bolChange = False
    '			ElseIf intYesNo = MsgBoxResult.No Then 
    '				m_bolChange = False
    '			End If

    '		End If

    '		Dim strSQLWQStd As String
    '		rsWQStdCboClick = New ADODB.Recordset

    '		Dim strSQLWQStdPoll As String
    '		rsWQStdPoll = New ADODB.Recordset


    '		'Selection based on combo box
    '		strSQLWQStd = "SELECT * FROM WQCRITERIA WHERE NAME LIKE '" & cboWQStdName.Text & "'"
    '		rsWQStdCboClick.CursorLocation = ADODB.CursorLocationEnum.adUseClient
    '		rsWQStdCboClick.Open(strSQLWQStd, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

    '		If rsWQStdCboClick.RecordCount > 0 Then

    '			txtWQStdDesc.Text = rsWQStdCboClick.Fields("Description").Value

    '			strSQLWQStdPoll = "SELECT POLLUTANT.NAME, POLL_WQCRITERIA.THRESHOLD, POLL_WQCRITERIA.POLL_WQCRITID " & "FROM POLL_WQCRITERIA INNER JOIN POLLUTANT " & "ON POLL_WQCRITERIA.POLLID = POLLUTANT.POLLID Where POLL_WQCRITERIA.WQCRITID = " & rsWQStdCboClick.Fields("WQCRITID").Value

    '			rsWQStdPoll.CursorLocation = ADODB.CursorLocationEnum.adUseClient
    '			rsWQStdPoll.Open(strSQLWQStdPoll, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

    '			grdWQStd.Recordset = rsWQStdPoll

    '			modUtil.InitWQStdGrid(grdWQStd)

    '			'Clean it
    '			'UPGRADE_NOTE: Object rsWQStdPoll may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '			rsWQStdPoll = Nothing
    '			'UPGRADE_NOTE: Object rsWQStdCboClick may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '			rsWQStdCboClick = Nothing

    '		Else

    '			MsgBox("Warning: There are no water quality standards remaining.  Please add a new one.", MsgBoxStyle.Critical, "Recordset Empty")
    '			'UPGRADE_NOTE: Object rsWQStdCboClick may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '			rsWQStdCboClick = Nothing

    '		End If



    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(True, "cboWQStdName_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub

    '	Private Sub cmdSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSave.Click

    '		On Error GoTo ErrHandler
    '		If ValidateData Then
    '			UpdateData()
    '			MsgBox("Data saved successfully.", MsgBoxStyle.Information, "Data Saved")
    '		End If

    '		Exit Sub
    'ErrHandler: 
    '		MsgBox("Error updating Water Quality Standards: " & Err.Number & vbNewLine & Err.Description, MsgBoxStyle.Critical, "Error")
    '	End Sub

    '	Private Sub frmWQStd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Click
    '		On Error GoTo ErrorHandler


    '		txtActiveCell.Visible = False



    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(True, "Form_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub

    '	Private Sub frmWQStd_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
    '		On Error GoTo ErrorHandler


    '		'Initialize the grid and then populate

    '		modUtil.InitComboBox(cboWQStdName, "WQCRITERIA")
    '		modUtil.InitWQStdGrid(grdWQStd)
    '		'Set some values
    '		m_bolChange = False
    '		intNumChanges = 0


    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub

    '	Private Sub cmdQuit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdQuit.Click
    '		On Error GoTo ErrorHandler

    '		Me.Close()


    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(True, "cmdQuit_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub

    '	Private Sub grdWQStd_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles grdWQStd.ClickEvent
    '		On Error GoTo ErrorHandler


    '		txtActiveCell.Visible = False

    '		intRow = grdWQStd.row
    '		intCol = grdWQStd.col

    '		If intCol = 2 Then

    '			With txtActiveCell

    '				.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(grdWQStd.Left) + grdWQStd.CellLeft), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(grdWQStd.Top) + grdWQStd.CellTop), VB6.TwipsToPixelsX(grdWQStd.CellWidth - 30), VB6.TwipsToPixelsY(grdWQStd.CellHeight - 75))

    '				If intRow <> 0 Then
    '					.Visible = True
    '					strUndoText = grdWQStd.get_TextMatrix(intRow, intCol)
    '					.Text = grdWQStd.get_TextMatrix(intRow, intCol)
    '					.Focus()
    '				End If

    '			End With

    '		End If



    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(True, "grdWQStd_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub



    '	'Options menu option processing
    '	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '	Public Sub mnuNewWQStd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuNewWQStd.Click
    '		On Error GoTo ErrorHandler

    '		VB6.ShowForm(frmAddWQStd, VB6.FormShowConstants.Modal, Me)


    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(True, "mnuNewWQStd_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub

    '	Public Sub mnuCopyWQStd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuCopyWQStd.Click
    '		On Error GoTo ErrorHandler


    '		VB6.ShowForm(frmCopyWQStd, VB6.FormShowConstants.Modal, Me)



    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(True, "mnuCopyWQStd_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub

    '	Public Sub mnuDelWQStd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDelWQStd.Click

    '		On Error GoTo ErrHandler
    '		Dim intAns As Short
    '		intAns = MsgBox("Are you sure you want to delete the Water Quality Standard '" & VB6.GetItemString(cboWQStdName, cboWQStdName.SelectedIndex) & "'?", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Confirm Delete")
    '		'code to handle response

    '		'WQ Recordset
    '		Dim strWQStdDelete As String
    '		Dim strWQPollDelete As String

    '		strWQStdDelete = "SELECT * FROM WQCriteria WHERE NAME LIKE '" & cboWQStdName.Text & "'"

    '		rsWQStdDelete = New ADODB.Recordset

    '		rsWQStdDelete.CursorLocation = ADODB.CursorLocationEnum.adUseClient
    '		rsWQStdDelete.Open(strWQStdDelete, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockOptimistic)

    '		strWQPollDelete = "Delete * FROM POLL_WQCRITERIA WHERE WQCRITID =" & rsWQStdDelete.Fields("WQCRITID").Value

    '		If Not (cboWQStdName.Text = "") Then
    '			'code to handle response
    '			If intAns = MsgBoxResult.Yes Then

    '				'Delete the WaterQuality Standard
    '				rsWQStdDelete.Delete(ADODB.AffectEnum.adAffectCurrent)
    '				rsWQStdDelete.Update()

    '				'modUtil.g_ADOConn.Execute strWQPollDelete

    '				MsgBox(VB6.GetItemString(cboWQStdName, cboWQStdName.SelectedIndex) & " deleted.", MsgBoxStyle.OKOnly, "Record Deleted")

    '				cboWQStdName.Items.Clear()
    '				modUtil.InitComboBox(cboWQStdName, "WQCRITERIA")
    '				Me.Refresh()

    '			ElseIf intAns = MsgBoxResult.No Then 
    '				Exit Sub
    '			End If
    '		Else
    '			MsgBox("Please select a water quality standard", MsgBoxStyle.Critical, "No Standard Selected")
    '		End If

    '		'Cleanup
    '		rsWQStdDelete.Close()
    '		'UPGRADE_NOTE: Object rsWQStdDelete may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '		rsWQStdDelete = Nothing

    '		Exit Sub
    'ErrHandler: 
    '		MsgBox("An Error occurred during deletion." & "  " & Err.Number & ": " & Err.Description, MsgBoxStyle.Critical, "Error")

    '	End Sub

    '	'Import Menu
    '	Public Sub mnuImpWQStd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuImpWQStd.Click
    '		On Error GoTo ErrorHandler

    '		VB6.ShowForm(frmImportWQStd, VB6.FormShowConstants.Modal, Me)

    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(True, "mnuImpWQStd_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub

    '	'Export Menu
    '	Public Sub mnuExpWQStd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuExpWQStd.Click
    '		On Error GoTo ErrorHandler

    '		Dim intAns As Short

    '		'browse...get output filename
    '		dlgCMD1Open.FileName = CStr(Nothing)
    '		dlgCMD1Save.FileName = CStr(Nothing)
    '		'UPGRADE_WARNING: CommonDialog variable was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="671167DC-EA81-475D-B690-7A40C7BF4A23"'
    '		With dlgCMD1
    '			'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
    '			.Filter = Replace(MSG1, "<name>", "Water Quality Standard")
    '			.Title = Replace(MSG3, "<name>", "Water Quality Standard")
    '			.FilterIndex = 1
    '			.DefaultExt = ".txt"
    '			'UPGRADE_WARNING: FileOpenConstants constant FileOpenConstants.cdlOFNHideReadOnly was upgraded to OpenFileDialog.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
    '			'UPGRADE_WARNING: MSComDlg.CommonDialog property dlgCMD1.Flags was upgraded to dlgCMD1Open.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
    '			'UPGRADE_WARNING: FileOpenConstants constant FileOpenConstants.cdlOFNHideReadOnly was upgraded to OpenFileDialog.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
    '			.ShowReadOnly = False
    '			'UPGRADE_WARNING: MSComDlg.CommonDialog property dlgCMD1.Flags was upgraded to dlgCMD1Save.OverwritePrompt which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
    '			.OverwritePrompt = True
    '			.ShowDialog()
    '		End With
    '		If Len(dlgCMD1Open.FileName) > 0 Then
    '			'Export Water Quality Standard to file - dlgCMD1.FileName
    '			ExportStandard((dlgCMD1Open.FileName))
    '		End If



    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(True, "mnuExpWQStd_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub

    '	Public Sub mnuWQHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuWQHelp.Click

    '		HtmlHelp(0, modUtil.g_nspectPath & "\Help\nspect.chm", HH_DISPLAY_TOPIC, "wq_stnds.htm")

    '	End Sub

    '	'UPGRADE_WARNING: Event txtActiveCell.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    '	Private Sub txtActiveCell_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtActiveCell.TextChanged
    '		On Error GoTo ErrorHandler

    '		'See the grd text to the text box
    '		grdWQStd.Text = txtActiveCell.Text
    '		m_bolChange = True



    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(True, "txtActiveCell_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub



    '	'Exports your current standard and pollutants to text or csv.
    '	Private Sub ExportStandard(ByRef strFileName As String)
    '		On Error GoTo ErrorHandler


    '		Dim fso As New Scripting.FileSystemObject
    '		Dim fl As Scripting.TextStream

    '		Dim rsNew As ADODB.Recordset
    '		Dim theLine As Object

    '		fl = fso.CreateTextFile(strFileName, True)

    '		'Write the name and descript.
    '		With fl
    '			.WriteLine(cboWQStdName.Text & "," & txtWQStdDesc.Text)
    '		End With

    '		Dim i As Short

    '		'Write name of pollutant and threshold
    '		For i = 1 To grdWQStd.Rows - 1
    '			fl.WriteLine(grdWQStd.get_TextMatrix(i, 1) & "," & grdWQStd.get_TextMatrix(i, 2))
    '		Next i

    '		fl.Close()

    '		'UPGRADE_NOTE: Object fso may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '		fso = Nothing
    '		'UPGRADE_NOTE: Object fl may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '		fl = Nothing



    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(False, "ExportStandard " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub

    '	Private Sub txtActiveCell_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtActiveCell.Enter
    '		On Error GoTo ErrorHandler


    '		strUndoText = txtActiveCell.Text

    '		txtActiveCell.SelectionStart = 0
    '		txtActiveCell.SelectionLength = Len(txtActiveCell.Text)



    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(True, "txtActiveCell_GotFocus " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub
    '	Private Sub txtActiveCell_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtActiveCell.KeyDown
    '		Dim KeyCode As Short = eventArgs.KeyCode
    '		Dim Shift As Short = eventArgs.KeyData \ &H10000
    '		On Error GoTo ErrorHandler

    '		'Handles some key inputs
    '		With grdWQStd

    '			Select Case KeyCode
    '				Case System.Windows.Forms.Keys.Escape 'if the user pressed escape, then get out without changing
    '					'.Text = UndoText
    '					txtActiveCell.Visible = False
    '					.Focus()
    '				Case 13 'if the user presses enter, get out of the textbox
    '					txtActiveCell.Visible = False
    '					.Focus()
    '				Case System.Windows.Forms.Keys.Up 'if the user presses the up arrow, get out of the textbox and move up a cell on the grid
    '					.Text = txtActiveCell.Text
    '					txtActiveCell.Visible = False
    '					.Focus()
    '					If .row > 0 Then
    '						.row = .row - 1
    '					Else
    '						.row = .row 'if the row is already on zero, don't move cells
    '					End If
    '					KeyMoveUpdate()
    '				Case System.Windows.Forms.Keys.Down 'if the user presses the down arrow, get out of the textbox and move down a cell on the grid
    '					.Text = txtActiveCell.Text
    '					txtActiveCell.Visible = False
    '					.Focus()
    '					If .row < grdWQStd.Rows - 1 Then
    '						.row = .row + 1
    '					Else
    '						.row = .row 'again, if the row is on the last row, don't move cells
    '					End If
    '					KeyMoveUpdate()
    '				Case System.Windows.Forms.Keys.Right 'if the user pressed the right arrow key, get out of the textbox and move right one cell in the grid
    '					txtActiveCell.Visible = False
    '					.Focus()
    '					.col = .col + 1
    '			End Select
    '		End With


    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(True, "txtActiveCell_KeyDown " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub

    '	Private Sub txtActiveCell_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtActiveCell.Leave
    '		On Error GoTo ErrorHandler

    '		With grdWQStd
    '			If intCol = .col And intRow = .row Then
    '				.Text = txtActiveCell.Text
    '				txtActiveCell.Visible = False
    '				.Focus()
    '			End If
    '		End With

    '		If CheckText(strUndoText, (txtActiveCell.Text)) > 0 Then
    '			cmdSave.Enabled = True
    '		End If

    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(True, "txtActiveCell_LostFocus " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub


    '	'Counter to determine if changes have been made
    '	Private Function CheckText(ByRef strOriginal As String, ByRef strNew As String) As Short
    '		On Error GoTo ErrorHandler


    '		If strOriginal <> strNew Then
    '			intNumChanges = intNumChanges + 1
    '			CheckText = intNumChanges
    '		End If

    '		Exit Function
    'ErrorHandler: 
    '		HandleError(False, "CheckText " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Function

    '	Private Sub txtWQStdDesc_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWQStdDesc.Click
    '		On Error GoTo ErrorHandler

    '		m_bolChange = True
    '		cmdSave.Enabled = m_bolChange


    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(True, "txtWQStdDesc_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub


    '	Private Sub KeyMoveUpdate()

    '		On Error GoTo ErrorHandler

    '		'This guy basically replicates the functionality of the grdWQstd_Click event and is used
    '		'in a couple of instances for moving around the grid.

    '		m_intWQRow = grdWQStd.row
    '		m_intWQCol = grdWQStd.col

    '		txtActiveCell.Visible = False

    '		If (m_intWQCol = 2) And (m_intWQRow > 0) Then

    '			With txtActiveCell

    '				.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(grdWQStd.Left) + grdWQStd.CellLeft), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(grdWQStd.Top) + grdWQStd.CellTop), VB6.TwipsToPixelsX(grdWQStd.CellWidth - 30), VB6.TwipsToPixelsY(grdWQStd.CellHeight - 150))

    '				If m_intWQRow <> 0 Then

    '					.Visible = True
    '					m_strUndoText = grdWQStd.get_TextMatrix(m_intWQRow, m_intWQCol)
    '					.Text = m_strUndoText

    '					If IsNumeric(m_strUndoText) Then
    '						txtActiveCell.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
    '					Else
    '						txtActiveCell.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
    '					End If

    '				End If
    '			End With
    '			txtActiveCell.Focus()

    '		ElseIf m_intWQCol = 0 Then 

    '			txtActiveCell.Visible = False

    '		End If

    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(False, "KeyMoveUpdate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub


    '	Private Sub UpdateData()
    '		On Error GoTo ErrorHandler


    '		Dim strSQLWQStd As String
    '		Dim strWQSelect As String
    '		Dim rsWQstd As New ADODB.Recordset
    '		rsWQStdCboClick = New ADODB.Recordset
    '		Dim rsPollUpdate As New ADODB.Recordset
    '		Dim i As Short

    '		Dim booYesNo As Short

    '		'Selection based on combo box, update Description
    '		strSQLWQStd = "SELECT * FROM WQCRITERIA WHERE NAME LIKE '" & cboWQStdName.Text & "'"

    '		With rsWQStdCboClick
    '			.CursorLocation = ADODB.CursorLocationEnum.adUseClient
    '			.Open(strSQLWQStd, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
    '		End With

    '		rsWQStdCboClick.Fields("Description").Value = txtWQStdDesc.Text
    '		rsWQStdCboClick.Update()

    '		'Now update Threshold values
    '		For i = 1 To grdWQStd.Rows - 1
    '			strWQSelect = "SELECT * from POLL_WQCRITERIA WHERE POLL_WQCRITID = " & grdWQStd.get_TextMatrix(i, 3)

    '			rsWQstd.Open(strWQSelect, g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

    '			rsWQstd.Fields("Threshold").Value = grdWQStd.get_TextMatrix(i, 2)
    '			rsWQstd.Update()
    '			rsWQstd.Close()
    '		Next i

    '		m_bolChange = False

    '		'Cleanup
    '		'UPGRADE_NOTE: Object rsWQstd may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '		rsWQstd = Nothing
    '		rsWQStdCboClick.Close()
    '		'UPGRADE_NOTE: Object rsWQStdCboClick may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '		rsWQStdCboClick = Nothing




    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(False, "UpdateData " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub

    '	Private Function ValidateData() As Boolean
    '		On Error GoTo ErrorHandler


    '		Dim i As Short

    '		For i = 1 To grdWQStd.Rows - 1

    '			If IsNumeric(grdWQStd.get_TextMatrix(i, 2)) Then
    '				If CShort(grdWQStd.get_TextMatrix(i, 2)) >= 0 Then
    '					ValidateData = True
    '				Else
    '					MsgBox("Warning: Values must be greater than or equal to 0.", MsgBoxStyle.Critical, "Invalid Value")
    '					grdWQStd.row = i
    '					grdWQStd.col = 2
    '					KeyMoveUpdate()
    '					ValidateData = False
    '				End If
    '			ElseIf grdWQStd.get_TextMatrix(i, 2) <> "" Then 
    '				MsgBox("Numeric values only please.", MsgBoxStyle.Critical, "Numeric Values Only")
    '				grdWQStd.row = i
    '				grdWQStd.col = 2
    '				KeyMoveUpdate()
    '				ValidateData = False
    '			End If
    '		Next 





    '		Exit Function
    'ErrorHandler: 
    '		HandleError(False, "ValidateData " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Function
End Class