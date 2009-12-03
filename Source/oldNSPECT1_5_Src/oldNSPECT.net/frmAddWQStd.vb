Option Strict Off
Option Explicit On
Friend Class frmAddWQStd
	Inherits System.Windows.Forms.Form
	' *************************************************************************************
	' *  frmAddWQStd
	' *  Perot Systems Government Services
	' *  Contact: Ed Dempsey  -  ed.dempsey@noaa.gov
	' *************************************************************************************
	' *  Description:  Form that allows for the addition of new water quality standard
	' *
	' *
	' *  Called By:  frmWQStd
	' *************************************************************************************
	
	
	Private intRow As Short 'Integer tracking current row
	Private intCol As Short 'Integer tracking current col
	Private strUndoText As String 'Text of txtActiveCell
	Private mChange As Boolean
	Private intNumChanges As Short
	
	' Variables used by the Error handler function - DO NOT REMOVE
	Const c_sModuleFileName As String = "frmAddWQStd.frm"
	
	
	Private Sub cmdSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSave.Click
		On Error GoTo ErrorHandler
		
		Dim strName As String
		Dim strDescript As String
		Dim strCmd As String
		
		'Get rid of possible apostrophes
		strName = Replace(Trim(txtWQStdName.Text), "'", "''")
		strDescript = Trim(txtWQStdDesc.Text)
		
		If Len(strName) = 0 Then
			MsgBox("Please enter a name for the water quality standard.", MsgBoxStyle.Critical, "Empty Name Field")
			txtWQStdName.Focus()
			Exit Sub
		Else
			'Name Check
			If modUtil.UniqueName("WQCRITERIA", (txtWQStdName.Text)) Then
				'Value check
				If CheckThreshValues Then
					strCmd = "INSERT INTO WQCRITERIA (NAME,DESCRIPTION) VALUES ('" & Replace(txtWQStdName.Text, "'", "''") & "', '" & Replace(strDescript, "'", "''") & "')"
					g_ADOConn.Execute(strCmd, ADODB.CommandTypeEnum.adCmdText)
				Else
					MsgBox("Threshold values must be numeric.", MsgBoxStyle.Critical, "Check Threshold Value")
					Exit Sub
				End If
			Else
				MsgBox("The name you have chosen is already in use.  Please select another.", MsgBoxStyle.Critical, "Select Unique Name")
				Exit Sub
			End If
		End If
		
		'If it gets here, time to add the pollutants
		Dim i As Short
		i = 0
		
		For i = 1 To grdWQStd.Rows - 1
			'Allow for numeric or nulls
			PollutantAdd((txtWQStdName.Text), grdWQStd.get_TextMatrix(i, 1), grdWQStd.get_TextMatrix(i, 2))
		Next i
		
		MsgBox(txtWQStdName.Text & " successfully added.", MsgBoxStyle.OKOnly, "Record Added")
		
		'Clean up stuff
		If modUtil.IsLoaded("frmWQStd") Then
			frmWQStd.cboWQStdName.Items.Clear()
			modUtil.InitComboBox((frmWQStd.cboWQStdName), "WQCRITERIA")
			frmWQStd.cboWQStdName.SelectedIndex = modUtil.GetCboIndex((txtWQStdName.Text), (frmWQStd.cboWQStdName))
		ElseIf modUtil.IsLoaded("frmPrj") Then 
			frmPrj.cboWQStd.Items.Clear()
			modUtil.InitComboBox((frmPrj.cboWQStd), "WQCRITERIA")
			frmPrj.cboWQStd.Items.Insert(frmPrj.cboWQStd.Items.Count, "Define a new water quality standard...")
			frmPrj.cboWQStd.SelectedIndex = modUtil.GetCboIndex((txtWQStdName.Text), (frmPrj.cboWQStd))
		End If
		
		Me.Close()
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "cmdSave_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	Private Function CheckThreshValues() As Boolean
		On Error GoTo ErrorHandler
		
		Dim i As Short
		i = 0
		
		For i = 1 To grdWQStd.Rows - 1
			If IsNumeric(grdWQStd.get_TextMatrix(i, 2)) Or grdWQStd.get_TextMatrix(i, 2) = "" Then
				CheckThreshValues = True
			Else
				CheckThreshValues = False
				Exit Function
			End If
		Next i
		
		Exit Function
ErrorHandler: 
		HandleError(False, "CheckThreshValues " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Function
	
	Private Sub frmAddWQStd_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		On Error GoTo ErrorHandler
		
		'Populate cbo with pollutant names
		Dim rsPollutant As ADODB.Recordset
		Dim strPollutant As String
		rsPollutant = New ADODB.Recordset
		
		mChange = False
		intNumChanges = 0
		
		modUtil.InitWQStdGrid(grdWQStd)
		
		strPollutant = "SELECT NAME FROM POLLUTANT ORDER BY NAME ASC"
		
		rsPollutant.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		rsPollutant.Open(strPollutant, g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
		
		Dim i As Short
		i = 1
		
		rsPollutant.MoveFirst()
		grdWQStd.Rows = rsPollutant.RecordCount + 1
		
		For i = 1 To rsPollutant.RecordCount
			grdWQStd.set_TextMatrix(i, 1, rsPollutant.Fields("Name").Value)
			rsPollutant.MoveNext()
		Next i
		
		'Cleanup
		rsPollutant.Close()
		'UPGRADE_NOTE: Object rsPollutant may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsPollutant = Nothing
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		On Error GoTo ErrorHandler
		
		Dim intvbYesNo As Short
		
		intvbYesNo = MsgBox("Are you sure you want to exit?  All changes not saved will be lost.", MsgBoxStyle.YesNo, "Exit?")
		
		If intvbYesNo = MsgBoxResult.Yes Then
			If IsLoaded("frmPrj") Then
				frmPrj.cboWQStd.SelectedIndex = 0
			End If
			
			Me.Close()
		Else
			Exit Sub
		End If
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "cmdCancel_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	Private Sub grdWQStd_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles grdWQStd.ClickEvent
		On Error GoTo ErrorHandler
		
		txtActiveCell.Visible = False
		
		intRow = grdWQStd.row
		intCol = grdWQStd.col
		
		If intCol = 2 Then
			
			With txtActiveCell
				
				.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(grdWQStd.Left) + grdWQStd.CellLeft), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(grdWQStd.Top) + grdWQStd.CellTop), VB6.TwipsToPixelsX(grdWQStd.CellWidth - 30), VB6.TwipsToPixelsY(grdWQStd.CellHeight - 75))
				
				If intRow <> 0 Then
					.Visible = True
					strUndoText = grdWQStd.get_TextMatrix(intRow, intCol)
					.Text = grdWQStd.get_TextMatrix(intRow, intCol)
					.Focus()
				End If
				
			End With
			
		End If
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "grdWQStd_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	'UPGRADE_WARNING: Event txtActiveCell.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtActiveCell_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtActiveCell.TextChanged
		On Error GoTo ErrorHandler
		
		'See the grd text to the text box
		grdWQStd.Text = txtActiveCell.Text
		
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "txtActiveCell_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	Private Sub txtActiveCell_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtActiveCell.Enter
		'Select all when clicked
		On Error GoTo ErrorHandler
		
		strUndoText = txtActiveCell.Text
		
		txtActiveCell.SelectionStart = 0
		txtActiveCell.SelectionLength = Len(txtActiveCell.Text)
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "txtActiveCell_GotFocus " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	'Key event handles for txtActiveCell
	Private Sub txtActiveCell_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtActiveCell.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		On Error GoTo ErrorHandler
		
		With grdWQStd
			
			Select Case KeyCode
				Case System.Windows.Forms.Keys.Escape 'if the user pressed escape, then get out without changing
					.Text = strUndoText
					txtActiveCell.Visible = False
					.Focus()
				Case 13 'if the user presses enter, get out of the textbox
					txtActiveCell.Visible = False
					.Focus()
				Case System.Windows.Forms.Keys.Up 'if the user presses the up arrow, get out of the textbox and move up a cell on the grid
					.Text = txtActiveCell.Text
					txtActiveCell.Visible = False
					.Focus()
					If .row > 0 Then
						.row = .row - 1
					Else
						.row = .row 'if the row is already on zero, don't move cells
					End If
					KeyMoveUpdate()
				Case System.Windows.Forms.Keys.Down 'if the user presses the down arrow, get out of the textbox and move down a cell on the grid
					.Text = txtActiveCell.Text
					txtActiveCell.Visible = False
					.Focus()
					If .row < grdWQStd.Rows - 1 Then
						.row = .row + 1
					Else
						.row = .row 'again, if the row is on the last row, don't move cells
					End If
					KeyMoveUpdate()
				Case System.Windows.Forms.Keys.Right 'if the user pressed the right arrow key, get out of the textbox and move right one cell in the grid
					txtActiveCell.Visible = False
					.Focus()
					.col = .col + 1
			End Select
			
		End With
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "txtActiveCell_KeyDown " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	Private Sub KeyMoveUpdate()
		On Error GoTo ErrorHandler
		
		'This guy basically replicates the functionality of the grdWQstd_Click event and is used
		'in a couple of instances for moving around the grid.
		
		intRow = grdWQStd.row
		intCol = grdWQStd.col
		
		txtActiveCell.Visible = False
		
		If (intCol = 2) And (intRow >= 1) Then
			
			With txtActiveCell
				
				.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(grdWQStd.Left) + grdWQStd.CellLeft), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(grdWQStd.Top) + grdWQStd.CellTop), VB6.TwipsToPixelsX(grdWQStd.CellWidth - 30), VB6.TwipsToPixelsY(grdWQStd.CellHeight - 150))
				
				If intRow <> 0 Then
					
					.Visible = True
					strUndoText = grdWQStd.get_TextMatrix(intRow, intCol)
					.Text = strUndoText
					
				End If
			End With
			txtActiveCell.Focus()
			
		ElseIf intCol = 0 Then 
			
			txtActiveCell.Visible = False
			
		End If
		
		
		
		Exit Sub
ErrorHandler: 
		HandleError(False, "KeyMoveUpdate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	
	Private Sub txtActiveCell_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtActiveCell.Leave
		On Error GoTo ErrorHandler
		
		
		With grdWQStd
			If intCol = .col And intRow = .row Then
				.Text = txtActiveCell.Text
				txtActiveCell.Visible = False
				.Focus()
			End If
		End With
		
		If CheckText(strUndoText, (grdWQStd.Text)) > 0 Then
			cmdSave.Enabled = True
		End If
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "txtActiveCell_LostFocus " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	
	'Counter to determine if changes have been made
	Private Function CheckText(ByRef strOriginal As String, ByRef strNew As String) As Short
		On Error GoTo ErrorHandler
		
		If strOriginal <> strNew Then
			intNumChanges = intNumChanges + 1
			CheckText = intNumChanges
		End If
		
		Exit Function
ErrorHandler: 
		HandleError(False, "CheckText " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Function
	
	
	Private Sub PollutantAdd(ByRef strName As String, ByRef strPoll As String, ByRef intThresh As String)
		On Error GoTo ErrorHandler
		
		Dim strPollAdd As String
		Dim strPollDetails As String
		Dim strCmdInsert As String
		
		Dim rsPollAdd As ADODB.Recordset
		Dim rsPollDetails As ADODB.Recordset
		
		rsPollAdd = New ADODB.Recordset
		rsPollDetails = New ADODB.Recordset
		
		'Get the WQCriteria values using the name
		strPollAdd = "SELECT * FROM WQCriteria WHERE NAME = " & "'" & strName & "'"
		rsPollAdd.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		rsPollAdd.Open(strPollAdd, g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
		
		'Get the pollutant particulars
		strPollDetails = "SELECT * FROM POLLUTANT WHERE NAME =" & "'" & strPoll & "'"
		rsPollDetails.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		rsPollDetails.Open(strPollDetails, g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
		
		If Trim(intThresh) = "" Then
			strCmdInsert = "INSERT INTO POLL_WQCRITERIA (PollID,WQCritID) VALUES ('" & rsPollDetails.Fields("POLLID").Value & "', '" & rsPollAdd.Fields("WQCRITID").Value & "')"
		Else
			strCmdInsert = "INSERT INTO POLL_WQCRITERIA (PollID,WQCritID,Threshold) VALUES ('" & rsPollDetails.Fields("POLLID").Value & "', '" & rsPollAdd.Fields("WQCRITID").Value & "'," & intThresh & ")"
		End If
		
		g_ADOConn.Execute(strCmdInsert, ADODB.CommandTypeEnum.adCmdText)
		
		'UPGRADE_NOTE: Object rsPollAdd may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsPollAdd = Nothing
		'UPGRADE_NOTE: Object rsPollDetails may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsPollDetails = Nothing
		
		Exit Sub
ErrorHandler: 
		HandleError(False, "PollutantAdd " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	'UPGRADE_WARNING: Event txtWQStdName.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtWQStdName_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWQStdName.TextChanged
		On Error GoTo ErrorHandler
		
		txtWQStdName.Text = Replace(txtWQStdName.Text, "'", "")
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "txtWQStdName_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
End Class