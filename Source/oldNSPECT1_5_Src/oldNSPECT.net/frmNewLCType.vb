Option Strict Off
Option Explicit On
Friend Class frmNewLCType
	Inherits System.Windows.Forms.Form
	' *************************************************************************************
	' *  Perot Systems Government Services
	' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
	' *  frmNewLCType
	' *************************************************************************************
	' *  Description: Form for entering a new Land Class Type
	' *  within NSPECT
	' *
	' *  Called By:  frmLCType menu New...
	' *************************************************************************************
	
	
	Private m_intRow As Short 'Current Row
	Private m_intCol As Short 'Current Col.
	Private m_intLCTypeID As Integer 'LCTypeID#
	
	Private m_bolGridChanged As Boolean 'Flag for whether or not grid values have changed
	Private m_bolSaved As Boolean 'Flag for saved/not saved changes
	Private m_bolFirstLoad As Boolean 'Is initial Load event
	Private m_bolBegin As Boolean
	Private m_intCount As Short
	
	Private m_strUndoText As String 'initial cell value - defaults back on Esc
	
	Private m_intMouseButton As Short 'Integer for mouse button click - added to avoid right click change cell value problem
	' Variables used by the Error handler function - DO NOT REMOVE
	Const c_sModuleFileName As String = "frmNewLCType.frm"
	
	
	'UPGRADE_WARNING: Event chkWWL.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkWWL_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkWWL.CheckStateChanged
		Dim Index As Short = chkWWL.GetIndex(eventSender)
		On Error GoTo ErrorHandler
		
		
		grdLCClasses.set_TextMatrix(Index, 8, chkWWL(Index).CheckState)
		
		
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "chkWWL_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		On Error GoTo ErrorHandler
		
		Dim strName As String
		Dim strDescript As String
		Dim strCmd As String 'INSERT function
		Dim arrParams(7) As Object 'Array that holds each row's contents
		Dim i As Short
		Dim j As Short
		
		
		If ValidateGridValues Then
			'Get rid of possible apostrophes in name
			strName = Trim(txtLCType.Text)
			strDescript = Trim(txtLCTypeDesc.Text)
			
			If modUtil.UniqueName("LCTYPE", strName) And (Trim(strName) <> "") Then
				strCmd = "INSERT INTO LCTYPE (NAME,DESCRIPTION) VALUES ('" & Replace(strName, "'", "''") & "', '" & Replace(strDescript, "'", "''") & "')"
				g_ADOConn.Execute(strCmd, ADODB.CommandTypeEnum.adCmdText)
			Else
				MsgBox("The name you have chosen is already in use.  Please select another.", MsgBoxStyle.Critical, "Select Unique Name")
				Exit Sub
			End If 'End unique name check
			
			'Now add GRID values
			
			For i = 1 To grdLCClasses.Rows - 1
				For j = 1 To 8 'Hard coded, bad, but ya know
					'UPGRADE_WARNING: Couldn't resolve default property of object arrParams(j - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					arrParams(j - 1) = grdLCClasses.get_TextMatrix(i, j)
				Next j
				
				AddLCClass(strName, arrParams)
				
			Next i
			
			MsgBox("Data saved successfully.", MsgBoxStyle.OKOnly, "Data Saved")
			
			Me.Close()
			frmLCTypes.Close()
		Else
			Exit Sub
		End If
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "cmdOK_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	Private Sub AddLCClass(ByRef strName As String, ByRef strParams() As Object)
		'Called from cmdOK_Click, this uses a passed array to insert new landclasses
		
		Dim strLCTypeAdd As String
		Dim strCmdInsert As String
		
		Dim rsLCType As ADODB.Recordset
		rsLCType = New ADODB.Recordset
		
		On Error GoTo ErrHandler
		'Get the WQCriteria values using the name
		strLCTypeAdd = "SELECT * FROM LCTYPE WHERE NAME = " & "'" & strName & "'"
		rsLCType.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		rsLCType.Open(strLCTypeAdd, g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
		
		'On Error GoTo dbaseerr
		strCmdInsert = "INSERT INTO LCCLASS([Value],[Name],[LCTYPEID],[CN-A],[CN-B],[CN-C],[CN-D],[CoverFactor],[W_WL]) VALUES(" & CStr(strParams(0)) & ",'" & CStr(strParams(1)) & "'," & CStr(rsLCType.Fields("LCTypeID").Value) & "," & CStr(strParams(2)) & "," & CStr(strParams(3)) & "," & CStr(strParams(4)) & "," & CStr(strParams(5)) & "," & CStr(strParams(6)) & "," & CStr(strParams(7)) & ")"
		
		g_ADOConn.Execute(strCmdInsert, ADODB.CommandTypeEnum.adCmdText)
		
		'UPGRADE_NOTE: Object rsLCType may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsLCType = Nothing
		
		Exit Sub
		
ErrHandler: 
		MsgBox("There is an error inserting records into LCClass.", MsgBoxStyle.Critical, Err.Number & ": " & Err.Description)
		Exit Sub
		
	End Sub
	
	
	Private Sub frmNewLCType_GotFocus(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.GotFocus
		On Error GoTo ErrorHandler
		
		txtLCType.Focus()
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "Form_GotFocus " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	Private Sub frmNewLCType_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		On Error GoTo ErrorHandler
		
		
		InitLCClassesGrid(grdLCClasses)
		
		With grdLCClasses
			.row = 1
			.set_TextMatrix(.row, 1, "0")
			.set_TextMatrix(.row, 2, "Landclass" & .row)
			.set_TextMatrix(.row, 3, "0")
			.set_TextMatrix(.row, 4, "0")
			.set_TextMatrix(.row, 5, "0")
			.set_TextMatrix(.row, 6, "0")
			.set_TextMatrix(.row, 7, "0")
			.set_TextMatrix(.row, 8, "0")
		End With
		
		CreateCheckBoxes(True)
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		On Error GoTo ErrorHandler
		
		Me.Close()
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "cmdCancel_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	Private Sub frmNewLCType_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Click
		On Error GoTo ErrorHandler
		
		txtActiveCell.Visible = False
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "Form_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	Private Sub cmdQuit_Click()
		On Error GoTo ErrorHandler
		
		Me.Close()
		
		Exit Sub
ErrorHandler: 
		HandleError(False, "cmdQuit_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	Private Sub grdLCClasses_MouseDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_MouseDownEvent) Handles grdLCClasses.MouseDownEvent
		On Error GoTo ErrorHandler
		
		m_intRow = grdLCClasses.MouseRow
		m_intCol = grdLCClasses.MouseCol
		
		m_intMouseButton = eventArgs.Button
		
		If eventArgs.Button = 2 Then
			
			If m_intCol = 0 Then
				With grdLCClasses
					.row = m_intRow
					.col = m_intCol
					modUtil.HighlightGridRow(grdLCClasses, m_intRow)
				End With
				'UPGRADE_ISSUE: Form method frmNewLCType.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				PopupMenu(mnuPopUp)
			End If
			
		End If
		
		Exit Sub
ErrorHandler: 
		HandleError(False, "grdLCClasses_MouseDown " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	Private Sub grdLcClasses_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles grdLcClasses.ClickEvent
		On Error GoTo ErrorHandler
		
		m_intRow = grdLCClasses.MouseRow
		m_intCol = grdLCClasses.MouseCol
		
		txtActiveCell.Visible = False
		
		If m_intMouseButton = 1 Then 'Right clicking snagged cell values, so we can only act on left click event
			
			If (m_intCol >= 1 And m_intCol <= 8) And (m_intRow >= 1) Then
				
				With txtActiveCell
					
					.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(grdLCClasses.Left) + grdLCClasses.CellLeft), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(grdLCClasses.Top) + grdLCClasses.CellTop), VB6.TwipsToPixelsX(grdLCClasses.CellWidth - 30), VB6.TwipsToPixelsY(grdLCClasses.CellHeight - 75))
					
					If m_intRow <> 0 Then
						.Visible = True
						m_strUndoText = grdLCClasses.get_TextMatrix(m_intRow, m_intCol)
						.Text = grdLCClasses.get_TextMatrix(m_intRow, m_intCol)
						.Focus()
						.SelectionLength = Len(.Text)
						
						If IsNumeric(m_strUndoText) Then
							txtActiveCell.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
						Else
							txtActiveCell.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
						End If
					End If
				End With
				
			ElseIf m_intCol = 0 Then 
				
				txtActiveCell.Visible = False
				
			End If
		End If
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "grdLcClasses_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	
	'UPGRADE_WARNING: Event txtActiveCell.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtActiveCell_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtActiveCell.TextChanged
		On Error GoTo ErrorHandler
		
		grdLCClasses.Text = txtActiveCell.Text
		m_bolGridChanged = True
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "txtActiveCell_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	Private Sub txtActiveCell_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtActiveCell.Enter
		On Error GoTo ErrorHandler
		
		txtActiveCell.SelectionStart = 0
		txtActiveCell.SelectionLength = Len(txtActiveCell.Text)
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "txtActiveCell_GotFocus " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	
	Private Sub txtActiveCell_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtActiveCell.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		On Error GoTo ErrorHandler
		
		'Handles some key inputs from the txtActiveCell, basically provides cell to cell movement around the
		'GRID
		
		With grdLCClasses
			Select Case KeyCode
				Case System.Windows.Forms.Keys.Escape 'if the user pressed escape, then get out without changing
					.Text = m_strUndoText
					txtActiveCell.Visible = False
					.Focus()
				Case 13 'if the user presses enter, get out of the textbox
					.Text = txtActiveCell.Text
					txtActiveCell.Visible = False
				Case System.Windows.Forms.Keys.Up 'if the user presses the up arrow, get out of the textbox and move up a cell on the grid
					.Text = txtActiveCell.Text
					If .row > 0 Then
						.row = .row - 1
						KeyMoveUpdate()
					Else
						.row = .row 'if the row is already on zero, don't move cells
					End If
				Case System.Windows.Forms.Keys.Down 'if the user presses the down arrow, get out of the textbox and move down a cell on the grid
					.Text = txtActiveCell.Text
					txtActiveCell.Visible = False
					If .row < grdLCClasses.Rows - 1 Then
						.row = .row + 1
						KeyMoveUpdate()
					Else
						.row = .row 'again, if the row is on the last row, don't move cells
						.Text = txtActiveCell.Text
						txtActiveCell.Visible = False
					End If
				Case System.Windows.Forms.Keys.Left 'if the user pressed the right arrow key, get out of the textbox and move right one cell in the grid
					txtActiveCell.Visible = False
					If .col > 1 Then
						.col = .col - 1
						KeyMoveUpdate()
					Else
						.col = .col
						.Text = txtActiveCell.Text
						txtActiveCell.Visible = False
					End If
				Case System.Windows.Forms.Keys.Right 'if the user pressed the right arrow key, get out of the textbox and move right one cell in the grid
					'.Text = txtActiveCell.Text
					txtActiveCell.Visible = False
					If .col < 8 Then
						.col = .col + 1
						KeyMoveUpdate()
					Else
						.col = .col
						.Text = txtActiveCell.Text
						txtActiveCell.Visible = False
					End If
				Case System.Windows.Forms.Keys.Tab
					MsgBox("WWWWWWWWWWWOOOOOO")
			End Select
		End With
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "txtActiveCell_KeyDown " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	Private Sub KeyMoveUpdate()
		On Error GoTo ErrorHandler
		
		'This guy basically replicates the functionality of the grdLCClasses_Click event and is used
		'in a couple of instances for moving around the grid.
		
		m_intRow = grdLCClasses.row
		m_intCol = grdLCClasses.col
		
		'cboWWL.Visible = False
		txtActiveCell.Visible = False
		
		If (m_intCol >= 1 And m_intCol < 8) And (m_intRow >= 1) Then
			
			With txtActiveCell
				
				.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(grdLCClasses.Left) + grdLCClasses.CellLeft), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(grdLCClasses.Top) + grdLCClasses.CellTop), VB6.TwipsToPixelsX(grdLCClasses.CellWidth - 30), VB6.TwipsToPixelsY(grdLCClasses.CellHeight - 75))
				
				If m_intRow <> 0 Then
					
					.Visible = True
					m_strUndoText = grdLCClasses.get_TextMatrix(m_intRow, m_intCol)
					.Text = m_strUndoText
					
					If IsNumeric(m_strUndoText) Then
						txtActiveCell.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
					Else
						txtActiveCell.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
					End If
					
				End If
			End With
			txtActiveCell.Focus()
			
		ElseIf m_intCol = 8 And m_intRow <> 0 Then 
			
		ElseIf m_intCol = 0 Then 
			
			txtActiveCell.Visible = False
			
		End If
		
		Exit Sub
ErrorHandler: 
		HandleError(False, "KeyMoveUpdate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	Public Sub mnuAppendRow_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuAppendRow.Click
		On Error GoTo ErrorHandler
		
		'add row to end of grid
		With grdLCClasses
			.Rows = .Rows + 1
			.row = .Rows - 1
			.set_TextMatrix(.row, 1, "0")
			.set_TextMatrix(.row, 2, "Landclass" & .row)
			.set_TextMatrix(.row, 3, "0")
			.set_TextMatrix(.row, 4, "0")
			.set_TextMatrix(.row, 5, "0")
			.set_TextMatrix(.row, 6, "0")
			.set_TextMatrix(.row, 7, "0")
			.set_TextMatrix(.row, 8, "0")
		End With
		
		CreateCheckBoxes(False, grdLCClasses.row)
		
		m_bolGridChanged = True 'reset
		cmdOK.Enabled = m_bolGridChanged
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "mnuAppendRow_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	Public Sub mnuDeleteRow_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDeleteRow.Click
		On Error GoTo ErrorHandler
		
		'delete current row
		Dim row, R, C, col As Short
		
		With grdLCClasses
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
				End If
			End If
		End With
		
		m_intCount = grdLCClasses.Rows
		ClearCheckBoxes(True, m_intCount + 1)
		CreateCheckBoxes(True)
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "mnuDeleteRow_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	Public Sub mnuInsertRow_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuInsertRow.Click
		On Error GoTo ErrorHandler
		
		'insert row above cuurent row
		Dim row, R, col As Short
		With grdLCClasses
			If .row < .FixedRows Then 'make sure we don't insert above header Rows
				mnuAppendRow_Click(mnuAppendRow, New System.EventArgs())
			Else
				R = .row
				.Rows = .Rows + 1 'add a row
				'.TextMatrix(.Rows - 1, 0) = .Rows - .FixedRows     'new row Title
				
				For row = .Rows - 1 To R + 1 Step -1 'move data dn 1 row
					For col = 1 To .get_Cols() - 1
						.set_TextMatrix(row, col, .get_TextMatrix(row - 1, col))
					Next col
				Next row
				For col = 1 To .get_Cols() - 1 ' clear all cells in this row
					If (col = 2) Then
						.set_TextMatrix(R, col, "")
					Else
						.set_TextMatrix(R, col, "0")
					End If
				Next col
			End If
		End With
		
		txtActiveCell.Visible = False
		
		m_intCount = grdLCClasses.Rows
		
		ClearCheckBoxes(True, m_intCount - 1)
		CreateCheckBoxes(True)
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "mnuInsertRow_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	Private Function ValidateGridValues() As Boolean
		On Error GoTo ErrorHandler
		
		'Need to validate each grid value before saving.  Essentially we take it a row at a time,
		'then rifle through each column of each row.  Case Select tests each each x,y value depending
		'on column... eg Column 1 must be unique, 3-6 must be 1-100 range, 7 must be <= 1
		
		'Returns: True or False
		
		Dim varActive As Object 'txtActiveCell value
		Dim varColumn2Value As Object 'Value of Column 2 ([VALUE]) - have to check for unique
		Dim i As Short
		Dim j As Short
		Dim k As Short
		
		For i = 1 To grdLCClasses.Rows - 1
			
			For j = 1 To 7
				
				'UPGRADE_WARNING: Couldn't resolve default property of object varActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				varActive = grdLCClasses.get_TextMatrix(i, j)
				
				Select Case j
					
					Case 1
						If Not IsNumeric(varActive) Then
							ErrorGenerator(Err1, i, j)
						Else
							For k = 1 To grdLCClasses.Rows - 1
								
								'UPGRADE_WARNING: Couldn't resolve default property of object varColumn2Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								varColumn2Value = grdLCClasses.get_TextMatrix(k, 1)
								If k <> i Then 'Don't want to compare value to itself
									'UPGRADE_WARNING: Couldn't resolve default property of object varColumn2Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									If varColumn2Value = grdLCClasses.get_TextMatrix(i, 1) Then
										ErrorGenerator(Err2, i, j)
										grdLCClasses.col = j
										grdLCClasses.row = i
										ValidateGridValues = False
										KeyMoveUpdate()
										Exit Function
									End If
								End If
							Next k
						End If
						
						
					Case 2
						If IsNumeric(varActive) Then
							ErrorGenerator(Err1, i, j)
							grdLCClasses.col = j
							grdLCClasses.row = i
							ValidateGridValues = False
							KeyMoveUpdate()
							Exit Function
						End If
						
					Case 3
						'UPGRADE_WARNING: Couldn't resolve default property of object varActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Not IsNumeric(varActive) Or ((varActive < 0) Or (varActive > 100)) Then
							ErrorGenerator(Err1, i, j)
							grdLCClasses.col = j
							grdLCClasses.row = i
							ValidateGridValues = False
							KeyMoveUpdate()
							Exit Function
						End If
						
					Case 4
						'UPGRADE_WARNING: Couldn't resolve default property of object varActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Not IsNumeric(varActive) Or ((varActive < 0) Or (varActive > 100)) Then
							ErrorGenerator(Err1, i, j)
							grdLCClasses.col = j
							grdLCClasses.row = i
							ValidateGridValues = False
							KeyMoveUpdate()
							Exit Function
						End If
						
					Case 5
						'UPGRADE_WARNING: Couldn't resolve default property of object varActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Not IsNumeric(varActive) Or ((varActive < 0) Or (varActive > 100)) Then
							ErrorGenerator(Err1, i, j)
							grdLCClasses.col = j
							grdLCClasses.row = i
							ValidateGridValues = False
							KeyMoveUpdate()
							Exit Function
						End If
						
					Case 6
						'UPGRADE_WARNING: Couldn't resolve default property of object varActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Not IsNumeric(varActive) Or ((varActive < 0) Or (varActive > 100)) Then
							ErrorGenerator(Err1, i, j)
							grdLCClasses.col = j
							grdLCClasses.row = i
							ValidateGridValues = False
							KeyMoveUpdate()
							Exit Function
						End If
						
					Case 7
						'UPGRADE_WARNING: Couldn't resolve default property of object varActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Not IsNumeric(varActive) Or ((varActive < 0) Or (varActive > 1)) Then
							ErrorGenerator(Err3, i, j)
							grdLCClasses.col = j
							grdLCClasses.row = i
							ValidateGridValues = False
							KeyMoveUpdate()
							Exit Function
						End If
				End Select
			Next j
		Next i
		
		ValidateGridValues = True
		
		Exit Function
ErrorHandler: 
		HandleError(False, "ValidateGridValues " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Function
	
	Private Sub CreateCheckBoxes(ByRef booAll As Boolean, Optional ByRef intRecNo As Short = 0)
		On Error GoTo ErrorHandler
		
		Dim i As Short
		Dim j As Short
		Dim k As Short
		Dim strChkName As String
		j = 1
		i = 1
		
		If booAll Then
			For i = 1 To grdLCClasses.Rows - 1
				
				grdLCClasses.row = i
				grdLCClasses.col = 8
				
				'Set the alignment to cover up current numbers
				grdLCClasses.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
				
				k = i
				
				'UPGRADE_ISSUE: Controls method Controls.Add was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				Call Controls.Add("VB.CheckBox", "chk" & CStr(k), Me)
				chkWWL.Load(k)
				chkWWL(k).Parent = Me
				With chkWWL(k)
					.Visible = True
					.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(grdLCClasses.Top) + grdLCClasses.CellTop)
					.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(grdLCClasses.Left) + grdLCClasses.CellLeft + (grdLCClasses.CellWidth * 0.4))
					.Height = VB6.TwipsToPixelsY(195)
					.Width = VB6.TwipsToPixelsX(195)
					.CheckState = CShort(grdLCClasses.get_TextMatrix(i, 8))
				End With
				Call Controls.RemoveAt("chk" & CStr(k))
			Next i
		Else
			
			With grdLCClasses
				.row = intRecNo
				.col = 8
				.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
			End With
			
			k = intRecNo
			
			'UPGRADE_ISSUE: Controls method Controls.Add was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Call Controls.Add("VB.CheckBox", "chk" & CStr(k), Me)
			chkWWL.Load(k)
			chkWWL(k).Parent = Me
			With chkWWL(k)
				.Visible = True
				.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(grdLCClasses.Top) + grdLCClasses.CellTop)
				.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(grdLCClasses.Left) + grdLCClasses.CellLeft + (grdLCClasses.CellWidth * 0.4))
				.Height = VB6.TwipsToPixelsY(195)
				.Width = VB6.TwipsToPixelsX(195)
				.CheckState = CShort(grdLCClasses.get_TextMatrix(k, 8))
			End With
			Call Controls.RemoveAt("chk" & CStr(k))
		End If
		
		Exit Sub
ErrorHandler: 
		HandleError(False, "CreateCheckBoxes " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	Private Sub ClearCheckBoxes(ByRef booAll As Boolean, ByRef intCount As Short, Optional ByRef intRecNo As Short = 0)
		On Error GoTo ErrorHandler
		
		Dim k As Short
		
		If booAll Then
			
			For k = 1 To intCount - 1
				chkWWL.Unload(k)
			Next k
		Else
			chkWWL.Unload(intRecNo)
		End If
		
		Exit Sub
ErrorHandler: 
		HandleError(False, "ClearCheckBoxes " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
End Class