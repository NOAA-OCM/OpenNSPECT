Option Strict Off
Option Explicit On
Friend Class frmLandCoverTypes
    Inherits System.Windows.Forms.Form
    '	' *************************************************************************************
    '	' *  Perot Systems Government Services
    '	' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
    '	' *  frmLCTypes
    '	' *************************************************************************************
    '	' *  Description: Form for browsing LC Types
    '	' *  within NSPECT
    '	' *
    '	' *  Called By:  NSPECT clsLCType
    '	' *************************************************************************************
    '	' *  Subs:
    '	' *     ExportLandCover(FileName as String) - creates text file of current classes
    '	' *         Called By: mnuExpLCType
    '	' *
    '	' *  Misc:  Uses an invisible menu called mnuPopUp for right click events on the FlexGrid
    '	' *************************************************************************************

    '	Private m_App As ESRI.ArcGIS.Framework.IApplication
    '	Private rsLCTypeCboClick As ADODB.Recordset 'Recordset
    '	Private rsCCAPDefault As ADODB.Recordset 'CCAP default recordset
    '	Private m_intRow As Short 'Current Row
    '	Private m_intCol As Short 'Current Col.
    '	Private m_intLCTypeID As Integer 'LCTypeID#
    '	Private m_intCount As Short 'Number of rows in old GRID
    '	Private m_bolGridChanged As Boolean 'Flag for whether or not grid values have changed
    '	Private m_bolSaved As Boolean 'Flag for saved/not saved changes
    '	Private m_bolFirstLoad As Boolean 'Is initial Load event
    '	Private m_bolBegin As Boolean
    '	Private m_strUndoText As String 'initial cell value used to track changes - defaults back on Esc
    '	Private m_strUndoDescrip As String 'same but for the Description
    '	Private m_intMouseButton As Short 'Integer for mouse button click - added to avoid right click change cell value problem

    '	Private clsLCClassData As New clsLCClassData 'Class that handles the data
    '	Dim WithEvents m_adoConn As ADODB.Connection 'BeginTrans, CommitTrans Event handler

    '	'UPGRADE_WARNING: Event chkWWL.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    '	Private Sub chkWWL_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkWWL.CheckStateChanged
    '		Dim Index As Short = chkWWL.GetIndex(eventSender)

    '		grdLCClasses.set_TextMatrix(Index, 8, chkWWL(Index).CheckState)

    '		If Not m_bolFirstLoad Then
    '			'm_bolGridChanged = True
    '			'MsgBox m_bolGridChanged
    '			CmdSaveEnabled()
    '		End If

    '	End Sub

    '	Private Sub chkWWL_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles chkWWL.MouseDown
    '		Dim Button As Short = eventArgs.Button \ &H100000
    '		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
    '		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
    '		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
    '		Dim Index As Short = chkWWL.GetIndex(eventSender)
    '		'MsgBox "MouseDown"
    '		m_bolGridChanged = True
    '		CmdSaveEnabled()
    '	End Sub

    '	Private Sub cmdRestore_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdRestore.Click
    '		'Restore Defaults Button - just read in NSPECT.LCCLASSDEFAULTS

    '		On Error GoTo Errhand

    '		Dim strCCAP As String

    '		'Check to make sure that's what they want
    '		intYesNo = MsgBox(strDefault, MsgBoxStyle.YesNo, strDefaultTitle)

    '		Dim i As Short
    '		If intYesNo = MsgBoxResult.Yes Then


    '			rsCCAPDefault = New ADODB.Recordset

    '			'Selection based on combo box
    '			strCCAP = "SELECT * From LCCLASSDEFAULTS"

    '			rsCCAPDefault.Open(strCCAP, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
    '			rsCCAPDefault.MoveFirst()

    '			grdLCClasses.Rows = rsCCAPDefault.RecordCount + 1
    '			'UPGRADE_NOTE: Refresh was upgraded to CtlRefresh. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    '			grdLCClasses.CtlRefresh()
    '			m_intCount = grdLCClasses.Rows

    '			'Clear out the old dataset - again, column 10 contains the LCClassID
    '			For i = 1 To rsCCAPDefault.RecordCount 'grdLCClasses.Rows - 1

    '				grdLCClasses.set_TextMatrix(i, 1, rsCCAPDefault.Fields("Value").Value)
    '				grdLCClasses.set_TextMatrix(i, 2, rsCCAPDefault.Fields("Name").Value)
    '				grdLCClasses.set_TextMatrix(i, 3, rsCCAPDefault.Fields("CN-A").Value)
    '				grdLCClasses.set_TextMatrix(i, 4, rsCCAPDefault.Fields("CN-B").Value)
    '				grdLCClasses.set_TextMatrix(i, 5, rsCCAPDefault.Fields("CN-C").Value)
    '				grdLCClasses.set_TextMatrix(i, 6, rsCCAPDefault.Fields("CN-D").Value)
    '				grdLCClasses.set_TextMatrix(i, 7, rsCCAPDefault.Fields("CoverFactor").Value)
    '				grdLCClasses.set_TextMatrix(i, 8, rsCCAPDefault.Fields("W_WL").Value)

    '				rsCCAPDefault.MoveNext()

    '			Next i

    '			ClearCheckBoxes(True, m_intCount)
    '			CreateCheckBoxes(True)

    '			rsCCAPDefault.Close()
    '			'UPGRADE_NOTE: Object rsCCAPDefault may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '			rsCCAPDefault = Nothing

    '			m_bolGridChanged = True
    '			CmdSaveEnabled()

    '		Else

    '			Exit Sub

    '		End If

    '		Exit Sub

    'Errhand: 
    '		MsgBox("There was an error loading the default CCAP data.", MsgBoxStyle.Critical, "Error Loading Data")


    '	End Sub

    '	Private Sub cmdSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSave.Click

    '		On Error GoTo ErrHandler
    '		txtActiveCell.Visible = False

    '		If ValidateGridValues Then
    '			UpdateValues()
    '			g_ADOConn.CommitTrans()

    '			MsgBox("Data saved successfully.", MsgBoxStyle.OKOnly, "Data Saved Successfully")

    '			'Reset the flags
    '			m_bolGridChanged = False
    '			m_bolSaved = True

    '			'UPGRADE_NOTE: Object m_adoConn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '			m_adoConn = Nothing

    '			Me.Close()
    '		End If
    '		Exit Sub
    'ErrHandler: 

    '		If Err.Number = -2147221504 Then
    '			MsgBox("The data values entered exceed the allowable precision of the database." & vbNewLine & "Data must not contain more than 4 values to the right of the decimal place." & vbNewLine & "Please correct your inputs before saving.", MsgBoxStyle.Information, "Precision Error")
    '			Exit Sub
    '		End If

    '		MsgBox("There was an error saving changes.", MsgBoxStyle.Critical, "Error Saving Changes")



    '		Exit Sub

    '	End Sub



    '	Private Sub frmLCTypes_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Click
    '		'If user clicks the form, get rid of the txtActiveCell

    '		If txtActiveCell.Visible = True Then
    '			grdLCClasses.Text = txtActiveCell.Text
    '			txtActiveCell.Visible = False
    '		End If


    '	End Sub

    '	Private Sub frmLCTypes_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

    '		m_adoConn = g_ADOConn

    '		'Set the flags
    '		m_bolSaved = False 'We haven't saved
    '		m_bolGridChanged = False 'Nothing's changed
    '		m_bolFirstLoad = True 'It's the first load

    '		'Initialize the Grid and populate the combobox
    '		modUtil.InitComboBox(cboLCType, "LCTYPE")

    '	End Sub


    '	'UPGRADE_WARNING: Event cboLCType.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    '	Private Sub cboLCType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboLCType.SelectedIndexChanged

    '		'On the cbo Click to change to a new LandClassType, check if there's been changes, prompt to save
    '		Dim strSQLLCType As String
    '		Dim strSQLLCClass As String
    '		If m_bolGridChanged And m_bolBegin Then

    '			intYesNo = MsgBox(strYesNo, MsgBoxStyle.YesNo, strYesNoTitle)
    '			'Since we're changing records for LCTypes...good time to save changes, ergo CommitTrans
    '			If intYesNo = MsgBoxResult.Yes Then
    '				UpdateValues()
    '				g_ADOConn.CommitTrans()
    '				m_bolGridChanged = False
    '				CmdSaveEnabled()
    '			ElseIf intYesNo = MsgBoxResult.No Then 
    '				g_ADOConn.RollbackTrans()
    '				m_bolGridChanged = False
    '				CmdSaveEnabled()
    '			End If

    '		Else

    '			m_intCount = grdLCClasses.Rows
    '			CheckCCAPDefault((cboLCType.Text))
    '			txtActiveCell.Visible = False

    '			'original
    '			rsLCTypeCboClick = New ADODB.Recordset


    '			'Selection based on combo box
    '			strSQLLCType = "SELECT * FROM LCTYPE WHERE NAME LIKE '" & cboLCType.Text & "'"

    '			With rsLCTypeCboClick
    '				.CursorLocation = ADODB.CursorLocationEnum.adUseClient
    '				.Open(strSQLLCType, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
    '			End With

    '			If rsLCTypeCboClick.RecordCount > 0 Then

    '				txtLCTypeDesc.Text = rsLCTypeCboClick.Fields("Description").Value

    '				strSQLLCClass = "SELECT LCCLASS.Value, LCCLASS.Name, LCCLASS.[CN-A], LCCLASS.[CN-B]," & " LCCLASS.[CN-C], LCCLASS.[CN-D], LCCLASS.CoverFactor, LCCLASS.W_WL, LCCLASS.LCTYPEID, LCCLASS.LCCLASSID FROM LCCLASS WHERE" & " LCTYPEID = " & rsLCTypeCboClick.Fields("LCTypeID").Value & " ORDER BY LCCLass.Value"

    '				With clsLCClassData
    '					.SQL = strSQLLCClass
    '				End With

    '				With grdLCClasses
    '					.Recordset = clsLCClassData.Recordset
    '					modUtil.InitLCClassesGrid(grdLCClasses)
    '					'UPGRADE_NOTE: Refresh was upgraded to CtlRefresh. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    '					.CtlRefresh()
    '				End With

    '				clsLCClassData.Recordset.MoveFirst()
    '				'Get the LCTypeID for use in updates, deletes, etc
    '				m_intLCTypeID = clsLCClassData.Recordset.Fields("LCTypeID").Value

    '				If Not m_bolBegin Then
    '					g_ADOConn.BeginTrans()
    '				End If
    '			Else

    '				MsgBox("Warning: There are no records remaining.  Please add a new one.", MsgBoxStyle.Critical, "Recordset Empty")

    '			End If

    '		End If

    '		If Not m_bolBegin Then
    '			g_ADOConn.BeginTrans()
    '		End If

    '		Timer1.Interval = 10
    '		Timer1.Enabled = True

    '	End Sub


    '	Public Sub mnuLCHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLCHelp.Click

    '		HtmlHelp(0, modUtil.g_nspectPath & "\Help\nspect.chm", HH_DISPLAY_TOPIC, "land_cover.htm")

    '	End Sub

    '	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick

    '		Timer1.Enabled = False
    '		If m_bolFirstLoad Then
    '			CreateCheckBoxes(True)
    '			m_bolFirstLoad = False
    '		Else
    '			ClearCheckBoxes(True, m_intCount)
    '			CreateCheckBoxes(True)
    '		End If

    '	End Sub

    '	Private Sub cmdQuit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdQuit.Click

    '		If Not m_bolSaved And m_bolGridChanged Then

    '			intYesNo = MsgBox("Do you want to save changes made to " & cboLCType.Text & "?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Exclamation, "N-SPECT")

    '			If intYesNo = MsgBoxResult.Yes Then

    '				If ValidateGridValues Then
    '					UpdateValues()
    '					g_ADOConn.CommitTrans()
    '					MsgBox("Data saved successfully.", MsgBoxStyle.OKOnly, "Save Successful")
    '					m_bolGridChanged = False
    '					m_bolSaved = True
    '					Me.Close()
    '				End If

    '			ElseIf intYesNo = MsgBoxResult.No Then 

    '				g_ADOConn.RollbackTrans()
    '				Me.Close()

    '			Else
    '				Exit Sub

    '			End If
    '		Else
    '			g_ADOConn.CommitTrans()
    '			Me.Close()
    '		End If

    '	End Sub

    '	Private Sub CmdSaveEnabled()

    '		cmdSave.Enabled = m_bolGridChanged

    '	End Sub

    '	Private Sub frmLCTypes_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
    '		'UPGRADE_NOTE: Object m_adoConn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '		m_adoConn = Nothing
    '	End Sub
    '	Private Sub m_adoConn_BeginTransComplete(ByVal TransactionLevel As Integer, ByVal pError As ADODB.Error, ByRef adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection) Handles m_adoConn.BeginTransComplete
    '		m_bolBegin = True
    '	End Sub
    '	Private Sub m_adoConn_CommitTransComplete(ByVal pError As ADODB.Error, ByRef adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection) Handles m_adoConn.CommitTransComplete
    '		m_bolBegin = False
    '	End Sub

    '	Private Sub m_adoConn_RollbackTransComplete(ByVal pError As ADODB.Error, ByRef adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection) Handles m_adoConn.RollbackTransComplete
    '		m_bolBegin = False
    '	End Sub

    '	'''''''''''''''''''''''''''''''''''''''
    '	'OPTIONS Menu
    '	'''''''''''''''''''''''''''''''''''''''

    '	'''''''''''''''''''''''''''
    '	'NEW Menu
    '	'''''''''''''''''''''''''''
    '	Public Sub mnuNewLCType_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuNewLCType.Click
    '		VB6.ShowForm(frmNewLCType, VB6.FormShowConstants.Modal, Me)
    '	End Sub

    '	'''''''''''''''''''''''''''
    '	'IMPORT Menu
    '	'''''''''''''''''''''''''''
    '	Public Sub mnuImpLCType_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuImpLCType.Click
    '		VB6.ShowForm(frmImportLCType, VB6.FormShowConstants.Modal, Me)
    '	End Sub

    '	'''''''''''''''''''''''''''
    '	'DELETE Menu
    '	'''''''''''''''''''''''''''
    '	Public Sub mnuDelLCType_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDelLCType.Click

    '		Dim intAns As Short
    '		intAns = MsgBox("Are you sure you want to delete the land cover type '" & VB6.GetItemString(cboLCType, cboLCType.SelectedIndex) & "' and all associated Coefficient Sets?", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Confirm Delete")

    '		Dim strLCTypeDelete As String
    '		Dim strLCClassDelete As String
    '		Dim rsLCTypeDelete As ADODB.Recordset
    '		Dim rsLCClassDelete As ADODB.Recordset
    '		If intAns = MsgBoxResult.Yes Then



    '			strLCTypeDelete = "SELECT * FROM LCTYPE WHERE NAME LIKE '" & cboLCType.Text & "'"

    '			rsLCTypeDelete = New ADODB.Recordset

    '			rsLCTypeDelete.CursorLocation = ADODB.CursorLocationEnum.adUseClient
    '			rsLCTypeDelete.Open(strLCTypeDelete, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockOptimistic)

    '			strLCClassDelete = "Delete * FROM LCCLASS WHERE LCTYPEID =" & rsLCTypeDelete.Fields("LCTypeID").Value

    '			If Not (cboLCType.Text = "") Then

    '				'code to handle response

    '				modUtil.g_ADOConn.Execute(strLCClassDelete)

    '				'Set up a delete rs and get rid of it
    '				rsLCTypeDelete.Delete(ADODB.AffectEnum.adAffectCurrent)
    '				rsLCTypeDelete.Update()

    '				MsgBox(VB6.GetItemString(cboLCType, cboLCType.SelectedIndex) & " deleted.", MsgBoxStyle.OKOnly, "Record Deleted")

    '				cboLCType.Items.Clear()
    '				m_bolGridChanged = False
    '				modUtil.InitComboBox(cboLCType, "LCType")
    '				Me.Refresh()

    '			Else
    '				MsgBox("Please select a Land class", MsgBoxStyle.Critical, "No Land Class Selected")
    '			End If



    '		ElseIf intAns = MsgBoxResult.No Then 
    '			m_bolGridChanged = False
    '			Exit Sub
    '		End If


    '	End Sub

    '	''''''''''''''''''''''''''
    '	'EXPORT Menu
    '	''''''''''''''''''''''''''
    '	Public Sub mnuExpLCType_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuExpLCType.Click
    '		Dim intAns As Short

    '		'browse...get output filename
    '		dlgCMD1Open.FileName = CStr(Nothing)
    '		dlgCMD1Save.FileName = CStr(Nothing)
    '		'UPGRADE_WARNING: CommonDialog variable was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="671167DC-EA81-475D-B690-7A40C7BF4A23"'
    '		With dlgCMD1
    '			'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
    '			.Filter = Replace(MSG1, "<name>", "Land Cover Type")
    '			.Title = Replace(MSG3, "<name>", "Land Cover Type")
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
    '			'Export land cover type to file - dlgCMD1.FileName
    '			ExportLandCover((dlgCMD1Open.FileName))
    '		End If


    '	End Sub

    '	'''***************************************
    '	'''mnuPopup Functionality - Andrew
    '	'''***************************************
    '	'Add row...
    '	Public Sub mnuAppend_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuAppend.Click

    '		clsLCClassData.LCTypeID = getLCTypeID
    '		clsLCClassData.AddNew()

    '		'add row to end of  grid
    '		With grdLCClasses
    '			.Rows = .Rows + 1
    '			.row = .Rows - 1
    '			.set_TextMatrix(.row, 0, "")
    '			.set_TextMatrix(.row, 1, "0")
    '			.set_TextMatrix(.row, 2, "Landclass" & .row)
    '			.set_TextMatrix(.row, 3, "0")
    '			.set_TextMatrix(.row, 4, "0")
    '			.set_TextMatrix(.row, 5, "0")
    '			.set_TextMatrix(.row, 6, "0")
    '			.set_TextMatrix(.row, 7, "0")
    '			.set_TextMatrix(.row, 8, "0")
    '			.set_TextMatrix(.row, 9, clsLCClassData.LCTypeID)
    '			.set_TextMatrix(.row, 10, g_intLCClassid)

    '		End With
    '		'IsCellVisible

    '		m_intCount = grdLCClasses.Rows

    '		CreateCheckBoxes(False, grdLCClasses.row)

    '		m_bolGridChanged = True
    '		CmdSaveEnabled()

    '	End Sub

    '	Public Sub mnuDeleteRow_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDeleteRow.Click
    '		'delete current row

    '		Dim row, R, C, col As Short
    '		Dim lngLCClassID As Integer

    '		With grdLCClasses

    '			lngLCClassID = CInt(.get_TextMatrix(m_intRow, 10))

    '			If .Rows > .FixedRows Then 'make sure we don't del header Rows
    '				For col = 1 To .get_Cols() - 1
    '					If ((Trim(.get_TextMatrix(.row, col)) > "" And col = 2) Or (.get_TextMatrix(.row, col) <> "0" And col = 1) Or (.get_TextMatrix(.row, col) <> "0" And col >= 3)) Then 'data?
    '						C = 1
    '						Exit For
    '					End If
    '				Next col
    '				If C Then
    '					R = MsgBox("There is data in Row" & Str(.row) & " ! Delete anyway?", MsgBoxStyle.YesNo, "Delete Row!")
    '				End If
    '				If C = 0 Or R = MsgBoxResult.Yes Then 'no exist. data or YES
    '					If .row = .Rows - 1 Then 'last row?
    '						.row = .row - 1 'move active cell
    '					Else
    '						For row = .row To .Rows - 2 'move data up 1 row
    '							For col = 1 To .get_Cols() - 1
    '								.set_TextMatrix(row, col, .get_TextMatrix(row + 1, col))
    '							Next col
    '						Next row
    '					End If
    '					.Rows = .Rows - 1 'del last row
    '					clsLCClassData.Load(lngLCClassID)
    '					clsLCClassData.Delete()
    '				End If
    '			End If
    '		End With

    '		m_intCount = grdLCClasses.Rows
    '		ClearCheckBoxes(True, m_intCount + 1)
    '		CreateCheckBoxes(True)

    '		m_bolGridChanged = True 'reset
    '		CmdSaveEnabled()

    '	End Sub

    '	Public Sub mnuInsertRow_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuInsertRow.Click
    '		'Insert row above current row in grdLCClasses- Thanks, Andrew

    '		'Get a hold of LCTYPEID, must have it to insert new records
    '		clsLCClassData.LCTypeID = m_intLCTypeID
    '		clsLCClassData.AddNew()

    '		Dim row, R, col As Short
    '		With grdLCClasses
    '			If .row < .FixedRows Then 'make sure we don't insert above header Rows
    '				mnuAppend_Click(mnuAppend, New System.EventArgs())
    '			Else
    '				R = .row
    '				.Rows = .Rows + 1 'add a row

    '				For row = .Rows - 1 To R + 1 Step -1 'move data dn 1 row
    '					For col = 1 To .get_Cols() - 1
    '						.set_TextMatrix(row, col, .get_TextMatrix(row - 1, col))
    '					Next col
    '				Next row
    '				For col = 1 To .get_Cols() - 1 ' clear all cells in this row
    '					If (col = 2) Then
    '						.set_TextMatrix(R, col, "")
    '					Else
    '						.set_TextMatrix(R, col, "0")
    '					End If
    '				Next col
    '				.set_TextMatrix(R, 9, clsLCClassData.LCTypeID)
    '				.set_TextMatrix(R, 10, g_intLCClassid)


    '			End If
    '		End With

    '		txtActiveCell.Visible = False
    '		m_intCount = grdLCClasses.Rows

    '		ClearCheckBoxes(True, m_intCount - 1)
    '		CreateCheckBoxes(True)

    '		m_bolGridChanged = True 'reset
    '		CmdSaveEnabled()

    '	End Sub

    '	Private Sub grdLCClasses_MouseDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_MouseDownEvent) Handles grdLCClasses.MouseDownEvent

    '		m_intRow = grdLCClasses.MouseRow
    '		m_intCol = grdLCClasses.MouseCol

    '		m_intMouseButton = eventArgs.Button

    '		If eventArgs.Button = 2 Then

    '			If m_intCol = 0 And cboLCType.Text <> "CCAP" Then
    '				With grdLCClasses
    '					.row = m_intRow
    '					.col = m_intCol
    '				End With
    '				modUtil.HighlightGridRow(grdLCClasses, m_intRow)
    '				'UPGRADE_ISSUE: Form method frmLCTypes.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '				PopupMenu(mnuPopUp)
    '			End If

    '		End If

    '	End Sub

    '	Private Sub grdLcClasses_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles grdLcClasses.ClickEvent

    '		m_intRow = grdLCClasses.MouseRow
    '		m_intCol = grdLCClasses.MouseCol

    '		txtActiveCell.Visible = False

    '		If m_intMouseButton = 1 Then

    '			If (m_intCol >= 1 And m_intCol < 8) And (m_intRow >= 1) Then

    '				With txtActiveCell

    '					.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(grdLCClasses.Left) + grdLCClasses.CellLeft), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(grdLCClasses.Top) + grdLCClasses.CellTop), VB6.TwipsToPixelsX(grdLCClasses.CellWidth - 30), VB6.TwipsToPixelsY(grdLCClasses.CellHeight - 75))

    '					If m_intRow <> 0 Then
    '						.Visible = True
    '						m_strUndoText = grdLCClasses.get_TextMatrix(m_intRow, m_intCol)
    '						.Text = grdLCClasses.get_TextMatrix(m_intRow, m_intCol)
    '						.Focus()
    '						.SelectionLength = Len(.Text)

    '						If m_intCol >= 3 And m_intCol <= 7 Then
    '							txtActiveCell.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
    '						Else
    '							If IsNumeric(m_strUndoText) Then
    '								txtActiveCell.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
    '							Else
    '								txtActiveCell.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
    '							End If
    '						End If
    '					End If
    '				End With

    '			ElseIf m_intCol = 0 Then 

    '				txtActiveCell.Visible = False

    '			End If
    '		Else
    '			Exit Sub
    '		End If
    '	End Sub

    '	'UPGRADE_WARNING: Event txtActiveCell.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    '	Private Sub txtActiveCell_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtActiveCell.TextChanged

    '		grdLCClasses.Text = txtActiveCell.Text

    '	End Sub


    '	Private Sub txtActiveCell_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtActiveCell.Enter

    '		m_strUndoText = txtActiveCell.Text

    '		txtActiveCell.SelectionStart = 0
    '		txtActiveCell.SelectionLength = Len(txtActiveCell.Text)

    '	End Sub

    '	Private Sub txtActiveCell_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtActiveCell.Validating
    '		Dim Cancel As Boolean = eventArgs.Cancel

    '		If Not m_bolGridChanged Then

    '			If txtActiveCell.Text <> m_strUndoText Then
    '				m_bolGridChanged = True
    '				CmdSaveEnabled()
    '			End If

    '		End If

    '		eventArgs.Cancel = Cancel
    '	End Sub

    '	Private Sub txtActiveCell_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtActiveCell.KeyDown
    '		Dim KeyCode As Short = eventArgs.KeyCode
    '		Dim Shift As Short = eventArgs.KeyData \ &H10000

    '		'Handles some key inputs from the txtActiveCell, basically provides cell to cell movement around the
    '		'GRID
    '		With grdLCClasses
    '			Select Case KeyCode
    '				Case System.Windows.Forms.Keys.Escape 'if the user pressed escape, then get out without changing
    '					.Text = m_strUndoText
    '					txtActiveCell.Visible = False
    '					.Focus()
    '				Case 13 'if the user presses enter, get out of the textbox
    '					.Text = txtActiveCell.Text
    '					txtActiveCell.Visible = False
    '				Case System.Windows.Forms.Keys.Up 'if the user presses the up arrow, get out of the textbox and move up a cell on the grid
    '					.Text = txtActiveCell.Text
    '					If .row > 0 Then
    '						.row = .row - 1
    '						KeyMoveUpdate()
    '					Else
    '						.row = .row 'if the row is already on zero, don't move cells
    '					End If
    '				Case System.Windows.Forms.Keys.Down 'if the user presses the down arrow, get out of the textbox and move down a cell on the grid
    '					.Text = txtActiveCell.Text
    '					txtActiveCell.Visible = False
    '					If .row < grdLCClasses.Rows - 1 Then
    '						.row = .row + 1
    '						KeyMoveUpdate()
    '					Else
    '						.row = .row 'again, if the row is on the last row, don't move cells
    '						.Text = txtActiveCell.Text
    '						txtActiveCell.Visible = False
    '					End If
    '				Case System.Windows.Forms.Keys.Left 'if the user pressed the right arrow key, get out of the textbox and move right one cell in the grid
    '					txtActiveCell.Visible = False
    '					If .col > 1 Then
    '						.col = .col - 1
    '						KeyMoveUpdate()
    '					Else

    '						If .row > 1 Then
    '							.col = 7
    '							.row = .row - 1
    '							KeyMoveUpdate()
    '						End If
    '					End If
    '				Case System.Windows.Forms.Keys.Right 'if the user pressed the right arrow key, get out of the textbox and move right one cell in the grid
    '					'.Text = txtActiveCell.Text
    '					txtActiveCell.Visible = False
    '					If .col < 7 Then
    '						.col = .col + 1
    '						KeyMoveUpdate()
    '					Else
    '						If .row < grdLCClasses.Rows - 1 Then
    '							.col = 1
    '							.row = .row + 1
    '							KeyMoveUpdate()
    '						End If
    '					End If
    '				Case System.Windows.Forms.Keys.Tab
    '					MsgBox("WWWWWWWWWWWOOOOOO")
    '			End Select
    '		End With
    '	End Sub

    '	Private Sub KeyMoveUpdate()

    '		'This guy basically replicates the functionality of the grdLCClasses_Click event and is used
    '		'in a couple of instances for moving around the grid.

    '		m_intRow = grdLCClasses.row
    '		m_intCol = grdLCClasses.col

    '		'chkWWL.Visible = False
    '		txtActiveCell.Visible = False

    '		If (m_intCol >= 1 And m_intCol < 8) And (m_intRow >= 1) Then

    '			With txtActiveCell

    '				.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(grdLCClasses.Left) + grdLCClasses.CellLeft), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(grdLCClasses.Top) + grdLCClasses.CellTop), VB6.TwipsToPixelsX(grdLCClasses.CellWidth - 30), VB6.TwipsToPixelsY(grdLCClasses.CellHeight - 75))

    '				If m_intRow <> 0 Then

    '					.Visible = True
    '					m_strUndoText = grdLCClasses.get_TextMatrix(m_intRow, m_intCol)
    '					.Text = m_strUndoText

    '					If m_intCol >= 3 And m_intCol <= 7 Then
    '						txtActiveCell.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
    '					Else
    '						If IsNumeric(m_strUndoText) Then
    '							txtActiveCell.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
    '						Else
    '							txtActiveCell.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
    '						End If
    '					End If

    '				End If
    '			End With
    '			txtActiveCell.Focus()

    '		ElseIf m_intCol = 0 Then 

    '			txtActiveCell.Visible = False

    '		End If

    '	End Sub

    '	'''***************************************
    '	'''Subs/Functions
    '	'''***************************************
    '	Private Sub ExportLandCover(ByRef strFileName As String)
    '		'Exports your current LCType/LCClasses to text or csv.

    '		Dim fso As New Scripting.FileSystemObject
    '		Dim fl As Scripting.TextStream
    '		Dim rsNew As ADODB.Recordset
    '		Dim theLine As Object

    '		fl = fso.CreateTextFile(strFileName, True)

    '		'Write the name and descript.
    '		With fl
    '			.WriteLine(cboLCType.Text & "," & txtLCTypeDesc.Text)
    '		End With

    '		Dim i As Short

    '		'Write name of pollutant and threshold
    '		For i = 1 To grdLCClasses.Rows - 1
    '			fl.WriteLine(grdLCClasses.get_TextMatrix(i, 1) & "," & grdLCClasses.get_TextMatrix(i, 2) & "," & grdLCClasses.get_TextMatrix(i, 3) & "," & grdLCClasses.get_TextMatrix(i, 4) & "," & grdLCClasses.get_TextMatrix(i, 5) & "," & grdLCClasses.get_TextMatrix(i, 6) & "," & grdLCClasses.get_TextMatrix(i, 7) & "," & grdLCClasses.get_TextMatrix(i, 8))

    '		Next i

    '		fl.Close()

    '		'UPGRADE_NOTE: Object fso may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '		fso = Nothing
    '		'UPGRADE_NOTE: Object fl may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '		fl = Nothing

    '	End Sub

    '	Private Function ValidateGridValues() As Boolean

    '		'Need to validate each grid value before saving.  Essentially we take it a row at a time,
    '		'then rifle through each column of each row.  Case Select tests each each x,y value depending
    '		'on column... eg Column 1 must be unique, 3-6 must be 1-100 range, 7 must be <= 1

    '		'Returns: True or False

    '		Dim varActive As Object 'txtActiveCell value
    '		Dim varColumn2Value As Object 'Value of Column 2 ([VALUE]) - have to check for unique
    '		Dim i As Short
    '		Dim j As Short
    '		Dim k As Short

    '		For i = 1 To grdLCClasses.Rows - 1

    '			For j = 1 To 7

    '				'UPGRADE_WARNING: Couldn't resolve default property of object varActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '				varActive = grdLCClasses.get_TextMatrix(i, j)

    '				Select Case j

    '					Case 1
    '						If Not IsNumeric(varActive) Then
    '							ErrorGenerator(Err1, i, j)
    '						Else
    '							For k = 1 To grdLCClasses.Rows - 1

    '								'UPGRADE_WARNING: Couldn't resolve default property of object varColumn2Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '								varColumn2Value = grdLCClasses.get_TextMatrix(k, 1)
    '								If k <> i Then 'Don't want to compare value to itself
    '									'UPGRADE_WARNING: Couldn't resolve default property of object varColumn2Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '									If varColumn2Value = grdLCClasses.get_TextMatrix(i, 1) Then
    '										ErrorGenerator(Err2, i, j)
    '										grdLCClasses.col = j
    '										grdLCClasses.row = i
    '										ValidateGridValues = False
    '										KeyMoveUpdate()
    '										Exit Function
    '									End If
    '								End If
    '							Next k
    '						End If


    '					Case 2
    '						If IsNumeric(varActive) Then
    '							ErrorGenerator(Err1, i, j)
    '							grdLCClasses.col = j
    '							grdLCClasses.row = i
    '							ValidateGridValues = False
    '							KeyMoveUpdate()
    '							Exit Function
    '						End If

    '					Case 3
    '						'UPGRADE_WARNING: Couldn't resolve default property of object varActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '						If Not IsNumeric(varActive) Or ((varActive < 0) Or (varActive > 1)) Or (Len(varActive) > 6) Then
    '							ErrorGenerator(Err1, i, j)
    '							grdLCClasses.col = j
    '							grdLCClasses.row = i
    '							ValidateGridValues = False
    '							KeyMoveUpdate()
    '							Exit Function
    '						End If

    '					Case 4
    '						'UPGRADE_WARNING: Couldn't resolve default property of object varActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '						If Not IsNumeric(varActive) Or ((varActive < 0) Or (varActive > 1)) Or (Len(varActive) > 6) Then
    '							ErrorGenerator(Err1, i, j)
    '							grdLCClasses.col = j
    '							grdLCClasses.row = i
    '							ValidateGridValues = False
    '							KeyMoveUpdate()
    '							Exit Function
    '						End If

    '					Case 5
    '						'UPGRADE_WARNING: Couldn't resolve default property of object varActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '						If Not IsNumeric(varActive) Or ((varActive < 0) Or (varActive > 1)) Or (Len(varActive) > 6) Then
    '							ErrorGenerator(Err1, i, j)
    '							grdLCClasses.col = j
    '							grdLCClasses.row = i
    '							ValidateGridValues = False
    '							KeyMoveUpdate()
    '							Exit Function
    '						End If

    '					Case 6
    '						'UPGRADE_WARNING: Couldn't resolve default property of object varActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '						If Not IsNumeric(varActive) Or ((varActive < 0) Or (varActive > 1)) Or (Len(varActive) > 6) Then
    '							ErrorGenerator(Err1, i, j)
    '							grdLCClasses.col = j
    '							grdLCClasses.row = i
    '							ValidateGridValues = False
    '							KeyMoveUpdate()
    '							Exit Function
    '						End If

    '					Case 7
    '						'UPGRADE_WARNING: Couldn't resolve default property of object varActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '						If Not IsNumeric(varActive) Or ((varActive < 0) Or (varActive > 1)) Or (Len(varActive) > 5) Then
    '							ErrorGenerator(Err3, i, j)
    '							grdLCClasses.col = j
    '							grdLCClasses.row = i
    '							ValidateGridValues = False
    '							KeyMoveUpdate()
    '							Exit Function
    '						End If
    '				End Select
    '			Next j
    '		Next i

    '		ValidateGridValues = True

    '	End Function

    '	Private Function getLCTypeID() As Integer

    '		getLCTypeID = m_intLCTypeID 'grdLCClasses.TextMatrix(1, 9)

    '	End Function


    '	Private Sub CheckCCAPDefault(ByRef strName As String)

    '		If strName = "CCAP" Then
    '			cmdRestore.Enabled = True
    '			mnuDelLCType.Enabled = False
    '			mnuPopUp.Enabled = False
    '			ToolTip1.SetToolTip(grdLCClasses, "")
    '		Else
    '			cmdRestore.Enabled = False
    '			mnuDelLCType.Enabled = True
    '			mnuPopUp.Enabled = True
    '			ToolTip1.SetToolTip(grdLCClasses, "Right click to add, delete, or insert a row")
    '		End If

    '	End Sub

    '	Private Sub UpdateValues()

    '		Dim i As Short

    '		For i = 1 To grdLCClasses.Rows - 1

    '			With clsLCClassData

    '				.Load(grdLCClasses.get_TextMatrix(i, 10))
    '				.Value = CInt(grdLCClasses.get_TextMatrix(i, 1))
    '				.Name = grdLCClasses.get_TextMatrix(i, 2)
    '				.CNA = CSng(grdLCClasses.get_TextMatrix(i, 3))
    '				.CNB = CSng(grdLCClasses.get_TextMatrix(i, 4))
    '				.CNC = CSng(grdLCClasses.get_TextMatrix(i, 5))
    '				.CND = CSng(grdLCClasses.get_TextMatrix(i, 6))
    '				.CoverFactor = CSng(grdLCClasses.get_TextMatrix(i, 7))
    '				.W_WL = CInt(grdLCClasses.get_TextMatrix(i, 8))
    '				.LCTypeID = CInt(grdLCClasses.get_TextMatrix(i, 9))
    '				.LCClassID = CInt(grdLCClasses.get_TextMatrix(i, 10))

    '				.SaveChanges()

    '			End With
    '		Next i

    '	End Sub

    '	Private Sub AddDefaultValues()

    '		Dim i As Short

    '		For i = 1 To grdLCClasses.Rows - 1

    '			With clsLCClassData

    '				.Value = CInt(grdLCClasses.get_TextMatrix(i, 1))
    '				.Name = grdLCClasses.get_TextMatrix(i, 2)
    '				.CNA = CSng(grdLCClasses.get_TextMatrix(i, 3))
    '				.CNB = CSng(grdLCClasses.get_TextMatrix(i, 4))
    '				.CNC = CSng(grdLCClasses.get_TextMatrix(i, 5))
    '				.CND = CSng(grdLCClasses.get_TextMatrix(i, 6))
    '				.CoverFactor = CSng(grdLCClasses.get_TextMatrix(i, 7))
    '				.W_WL = CInt(grdLCClasses.get_TextMatrix(i, 8))
    '				.LCTypeID = CInt(grdLCClasses.get_TextMatrix(i, 9))

    '				.AddNew()

    '			End With
    '		Next i

    '		ClearCheckBoxes(True, m_intCount)
    '		CreateCheckBoxes(True)


    '	End Sub

    '	Private Sub CreateCheckBoxes(ByRef booAll As Boolean, Optional ByRef intRecNo As Short = 0)

    '		Dim i As Short
    '		Dim j As Short
    '		Dim k As Short
    '		Dim strChkName As String
    '		j = 1
    '		i = 1

    '		If booAll Then
    '			For i = 1 To grdLCClasses.Rows - 1

    '				grdLCClasses.row = i
    '				grdLCClasses.col = 8

    '				'Set the alignment to cover up current numbers
    '				grdLCClasses.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter

    '				k = i

    '				'UPGRADE_ISSUE: Controls method Controls.Add was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '				Call Controls.Add("VB.CheckBox", "chk" & CStr(k), Me)
    '				chkWWL.Load(k)
    '				chkWWL(k).Parent = Me
    '				With chkWWL(k)
    '					.Visible = True
    '					.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(grdLCClasses.Top) + grdLCClasses.CellTop)
    '					.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(grdLCClasses.Left)) + (grdLCClasses.CellLeft) + (grdLCClasses.CellWidth * 0.4))
    '					.Height = VB6.TwipsToPixelsY(195)
    '					.Width = VB6.TwipsToPixelsX(195)
    '					.CheckState = CShort(grdLCClasses.get_TextMatrix(i, 8))
    '				End With
    '				Call Controls.RemoveAt("chk" & CStr(k))
    '			Next i
    '		Else

    '			With grdLCClasses
    '				.row = intRecNo
    '				.col = 8
    '				.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
    '			End With

    '			k = intRecNo

    '			'UPGRADE_ISSUE: Controls method Controls.Add was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '			Call Controls.Add("VB.CheckBox", "chk" & CStr(k), Me)
    '			chkWWL.Load(k)
    '			chkWWL(k).Parent = Me
    '			With chkWWL(k)
    '				.Visible = True
    '				.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(grdLCClasses.Top) + grdLCClasses.CellTop)
    '				.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(grdLCClasses.Left) + grdLCClasses.CellLeft + (grdLCClasses.CellWidth * 0.4))
    '				.Height = VB6.TwipsToPixelsY(195)
    '				.Width = VB6.TwipsToPixelsX(195)
    '				.CheckState = CShort(grdLCClasses.get_TextMatrix(k, 8))
    '			End With
    '			Call Controls.RemoveAt("chk" & CStr(k))
    '		End If

    '	End Sub

    '	Private Sub ClearCheckBoxes(ByRef booAll As Boolean, ByRef intCount As Short, Optional ByRef intRecNo As Short = 0)

    '		Dim k As Short

    '		If booAll Then

    '			'UPGRADE_WARNING: Couldn't resolve default property of object chkWWL().UBound. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '			For k = 1 To chkWWL().UBound 'intCount - 1
    '				chkWWL.Unload(k)
    '			Next k
    '		Else
    '			chkWWL.Unload(intRecNo)
    '		End If
    '	End Sub


    '	Public Sub init(ByVal pApp As ESRI.ArcGIS.Framework.IApplication)

    '		m_App = pApp

    '	End Sub
End Class