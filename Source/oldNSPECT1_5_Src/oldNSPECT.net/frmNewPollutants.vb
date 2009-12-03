Option Strict Off
Option Explicit On
Friend Class frmNewPollutants
	Inherits System.Windows.Forms.Form
	' *************************************************************************************
	' *  Perot Systems Government Services
	' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
	' *  frmPollutants
	' *************************************************************************************
	' *  Description: Form for browsing pollutants
	' *  within NSPECT
	' *
	' *  Called By:  NSPECT clsPollutants
	' *************************************************************************************
	' *  Subs:
	' *     ExportLandCover(FileName as String) - creates text file of current classes
	' *         Called By: mnuExpLCType
	' *
	' *  Misc:  Uses an invisible menu called mnuPopUp for right click events on the FlexGrid
	' *************************************************************************************
	
	
	Private rsLCTypeCboClick As New ADODB.Recordset 'RS on cboPollName click event
	Private rsCoeff As ADODB.Recordset
	Private rsLCType As ADODB.Recordset
	Private rsFullCoeff As ADODB.Recordset 'RS on cboCoeffSet click event
	Private rsCoeffs As ADODB.Recordset 'RS that fills FlexGrid
	Private rsWQStds As ADODB.Recordset 'Rs that fills FelexGrid on Standards Tab
	
	Private boolLoaded As Boolean
	Private boolChanged As Boolean 'Bool for enabling save button
	Private boolDescChanged As Boolean 'Bool for seeing if Description Changed
	Private boolSaved As Boolean 'Bool for whether or not things have saved
	
	Private m_intCurFrame As Short 'Current Frame visible in SSTab
	Private m_intPollRow As Short 'Row Number for grdPolldef
	Private m_intPollCol As Short 'Column Number for grdPollDef
	Private intRowWQ As Short 'Row Number for grdPollWQStd
	Private intColWQ As Short 'Column Number for grdPollDef
	Private m_intPollID As Short 'There's a need to have the PollID so we'll store it here
	Private m_intLCTypeID As Short 'Land Class (CCAP) ID - needed to add new coefficient sets
	Private m_intCoeffID As Short 'Key for CoefficientSetID - needed to add new coefficients 'See above
	
	Private m_strUndoText As String 'Text for txtActiveCell     |
	Private m_strUndoTextWQ As String 'Text for txtActiveCellWQ   |-all three used to detect change
	Private m_strUndoDesc As String 'Text for Description       |
	Private m_strLCType As String 'Need for name, we'll store here
	
	'UPGRADE_WARNING: Event cboLCType.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboLCType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboLCType.SelectedIndexChanged
		
		Dim strLCClasses As String
		Dim i As Short
		
		strLCClasses = "SELECT LCTYPE.LCTYPEID, LCCLASS.VALUE, LCCLASS.NAME, LCCLASS.LCCLASSID FROM LCTYPE INNER JOIN LCCLASS ON " & "LCTYPE.LCTYPEID = LCCLASS.LCTYPEID WHERE LCTYPE.NAME LIKE '" & cboLCType.Text & "'" & " ORDER BY LCCLASS.VALUE"
		
		rsLCTypeCboClick.Open(strLCClasses, g_ADOConn, ADODB.CursorTypeEnum.adOpenKeyset)
		
		grdPollDef.Clear()
		InitPollDefGrid(grdPollDef)
		grdPollDef.Rows = rsLCTypeCboClick.RecordCount + 1
		
		rsLCTypeCboClick.MoveFirst()
		
		'Actually add the records to the new set
		For i = 1 To rsLCTypeCboClick.RecordCount
			
			With grdPollDef
				.set_TextMatrix(i, 1, rsLCTypeCboClick.Fields("Value").Value)
				.set_TextMatrix(i, 2, rsLCTypeCboClick.Fields("Name").Value)
				.set_TextMatrix(i, 3, 0)
				.set_TextMatrix(i, 4, 0)
				.set_TextMatrix(i, 5, 0)
				.set_TextMatrix(i, 6, 0)
				.set_TextMatrix(i, 7, 0)
				.set_TextMatrix(i, 8, rsLCTypeCboClick.Fields("LCClassID").Value)
				
				rsLCTypeCboClick.MoveNext()
			End With
			
		Next i
		
		rsLCTypeCboClick.Close()
		'UPGRADE_NOTE: Object rsLCTypeCboClick may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsLCTypeCboClick = Nothing
		
	End Sub
	
	Private Sub cmdSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSave.Click
		
		If CheckForm Then
			
			If UpdateValues Then
				
				MsgBox(txtPollutant.Text & " successfully added.  Please enter value for associated water quality standards.", MsgBoxStyle.Information, "Pollutant Successfully Added")
				Me.Close()
				frmPollutants.SSTab1.SelectedIndex = 1
			End If
			
		End If
		
	End Sub
	
	Private Sub frmNewPollutants_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		Dim rsWQstd As New ADODB.Recordset
		Dim strSelectWQStd As String
		
		'Initialize the two flex Grids
		InitPollDefGrid(grdPollDef)
		InitPollWQStdGrid(grdWQStd)
		
		With Me
			.SSTab1.SelectedIndex = 0
			.SSTab1.TabPages.Item(1).Enabled = False
			.SSTab1.TabPages.Item(1).Visible = False
		End With
		
		modUtil.InitComboBox(cboLCType, "LCType")
		
		boolLoaded = True
		
		
	End Sub
	
	Private Sub cmdQuit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdQuit.Click
		
		Me.Close()
		
	End Sub
	
	Private Sub grdPollDef_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles grdPollDef.ClickEvent
		
		'Code for Poll Def Grid - moves txtActiveCell to appropriate cell
		
		m_intPollRow = grdPollDef.MouseRow
		m_intPollCol = grdPollDef.MouseCol
		
		txtActiveCell.Visible = False
		
		'Column greater than or equal to 1 and same for row
		If m_intPollCol >= 3 Then
			
			With txtActiveCell
				'Put in position
				.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(SSTab1.Left) + VB6.PixelsToTwipsX(grdPollDef.Left) + grdPollDef.CellLeft), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(SSTab1.Top) + VB6.PixelsToTwipsY(grdPollDef.Top) + grdPollDef.CellTop), VB6.TwipsToPixelsX(grdPollDef.CellWidth - 30), VB6.TwipsToPixelsY(grdPollDef.CellHeight - 150))
				
				If m_intPollRow <> 0 Then
					.Visible = True
					m_strUndoText = grdPollDef.get_TextMatrix(m_intPollRow, m_intPollCol)
					.Text = grdPollDef.get_TextMatrix(m_intPollRow, m_intPollCol)
					.Focus()
					
					If IsNumeric(m_strUndoText) Then
						txtActiveCell.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
					Else
						txtActiveCell.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
					End If
					
				End If
				
			End With
			
		End If
		
		
	End Sub
	
	Private Sub grdPollDef_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyDownEvent) Handles grdPollDef.KeyDownEvent
		
		With grdPollDef
			Select Case eventArgs.KeyCode
				Case System.Windows.Forms.Keys.Escape 'if the user pressed escape, then get out without changing
					.Text = m_strUndoText
					txtActiveCell.Visible = False
					.Focus()
					
				Case 13 'if the user presses enter, get out of the textbox
					.Text = txtActiveCell.Text
					txtActiveCell.Visible = False
					
				Case System.Windows.Forms.Keys.Up 'if the user presses the up arrow, get out of the textbox and move up a cell on the grid
					.Text = txtActiveCell.Text
					If .row > 1 Then
						.row = .row - 1
						KeyMoveUpdate()
					Else
						If .col < 6 Then
							.col = .col + 1
							.row = .Rows - 1
						Else
							.row = .row
						End If
						KeyMoveUpdate()
					End If
					
				Case System.Windows.Forms.Keys.Down 'if the user presses the down arrow, get out of the textbox and move down a cell on the grid
					.Text = txtActiveCell.Text
					txtActiveCell.Visible = False
					If .row < grdPollDef.Rows - 1 Then
						.row = .row + 1
						KeyMoveUpdate()
					Else
						If .col > 1 Then
							.row = 1 'again, if the row is on the last row, don't move cells
							.col = .col - 1
						Else
							.Text = txtActiveCell.Text
							txtActiveCell.Visible = False
						End If
						KeyMoveUpdate()
					End If
				Case System.Windows.Forms.Keys.Left 'if the user pressed the right arrow key, get out of the textbox and move right one cell in the grid
					txtActiveCell.Visible = False
					If .col > 3 Then
						.col = .col - 1
						KeyMoveUpdate()
						
					Else
						
						If .row > 1 And .col = 3 Then
							.col = 6
							.row = .row - 1
							KeyMoveUpdate()
						End If
					End If
				Case System.Windows.Forms.Keys.Right 'if the user pressed the right arrow key, get out of the textbox and move right one cell in the grid
					.Text = txtActiveCell.Text
					txtActiveCell.Visible = False
					If .col < 6 Then
						.col = .col + 1
						KeyMoveUpdate()
					Else
						If .row < grdPollDef.Rows - 1 Then
							.col = 3
							.row = .row + 1
							KeyMoveUpdate()
						End If
					End If
			End Select
		End With
		
		
	End Sub
	
	Private Sub grdWQStd_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles grdWQStd.ClickEvent
		
		Dim strUndoText As String
		
		txtActiveCellWQStd.Visible = False
		intRowWQ = grdWQStd.row
		intColWQ = grdWQStd.col
		
		'We want to limit editing to only the threshold column
		If intColWQ = 3 And intRowWQ >= 1 Then
			
			With txtActiveCellWQStd
				
				.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(SSTab1.Left) + VB6.PixelsToTwipsX(grdWQStd.Left) + grdWQStd.CellLeft), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(SSTab1.Top) + VB6.PixelsToTwipsY(grdWQStd.Top) + grdWQStd.CellTop), VB6.TwipsToPixelsX(grdWQStd.CellWidth - 30), VB6.TwipsToPixelsY(grdWQStd.CellHeight - 150))
				
				If intRowWQ <> 0 Then
					.Visible = True
					strUndoText = grdWQStd.get_TextMatrix(intRowWQ, intColWQ)
					.Text = grdWQStd.get_TextMatrix(intRowWQ, intColWQ)
					.Focus()
					
					If IsNumeric(strUndoText) Then
						txtActiveCellWQStd.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
					Else
						txtActiveCellWQStd.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
					End If
					
				End If
				
			End With
			
		End If
		
	End Sub
	
	
	'UPGRADE_WARNING: Event txtActiveCell.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtActiveCell_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtActiveCell.TextChanged
		
		grdPollDef.Text = txtActiveCell.Text
		boolChanged = True
		CmdSaveEnabled()
		
	End Sub
	
	Private Sub txtActiveCell_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtActiveCell.Enter
		
		txtActiveCell.SelectionStart = 0
		txtActiveCell.SelectionLength = Len(txtActiveCell.Text)
		
	End Sub
	
	Private Sub txtActiveCell_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtActiveCell.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		With grdPollDef
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
					If .row > 1 Then
						.row = .row - 1
						KeyMoveUpdate()
					Else
						If .col < 6 Then
							.col = .col + 1
							.row = .Rows - 1
						Else
							.row = .row
						End If
						KeyMoveUpdate()
					End If
					
				Case System.Windows.Forms.Keys.Down 'if the user presses the down arrow, get out of the textbox and move down a cell on the grid
					.Text = txtActiveCell.Text
					txtActiveCell.Visible = False
					If .row < grdPollDef.Rows - 1 Then
						.row = .row + 1
						KeyMoveUpdate()
					Else
						If .col > 1 Then
							.row = 1 'again, if the row is on the last row, don't move cells
							.col = .col - 1
						Else
							.Text = txtActiveCell.Text
							txtActiveCell.Visible = False
						End If
						KeyMoveUpdate()
					End If
				Case System.Windows.Forms.Keys.Left 'if the user pressed the right arrow key, get out of the textbox and move right one cell in the grid
					txtActiveCell.Visible = False
					If .col > 3 Then
						.col = .col - 1
						KeyMoveUpdate()
						
					Else
						
						If .row > 1 And .col = 3 Then
							.col = 6
							.row = .row - 1
							KeyMoveUpdate()
						End If
					End If
				Case System.Windows.Forms.Keys.Right 'if the user pressed the right arrow key, get out of the textbox and move right one cell in the grid
					.Text = txtActiveCell.Text
					txtActiveCell.Visible = False
					If .col < 6 Then
						.col = .col + 1
						KeyMoveUpdate()
					Else
						If .row < grdPollDef.Rows - 1 Then
							.col = 3
							.row = .row + 1
							KeyMoveUpdate()
						End If
					End If
			End Select
		End With
		
	End Sub
	
	Private Sub KeyMoveUpdate()
		'This guy basically replicates the functionality of the grdpolldef_Click event and is used
		'in a couple of instances for moving around the grid.
		
		m_intPollRow = grdPollDef.row
		m_intPollCol = grdPollDef.col
		
		txtActiveCell.Visible = False
		
		If (m_intPollCol >= 3 And m_intPollCol < 7) And (m_intPollRow >= 1) Then
			
			With txtActiveCell
				
				.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(SSTab1.Left) + VB6.PixelsToTwipsX(grdPollDef.Left) + grdPollDef.CellLeft), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(SSTab1.Top) + VB6.PixelsToTwipsY(grdPollDef.Top) + grdPollDef.CellTop), VB6.TwipsToPixelsX(grdPollDef.CellWidth - 30), VB6.TwipsToPixelsY(grdPollDef.CellHeight - 150))
				
				If m_intPollRow <> 0 Then
					
					.Visible = True
					m_strUndoText = grdPollDef.get_TextMatrix(m_intPollRow, m_intPollCol)
					.Text = m_strUndoText
					
					If IsNumeric(m_strUndoText) Then
						txtActiveCell.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
					Else
						txtActiveCell.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
					End If
					
				End If
			End With
			txtActiveCell.Focus()
			
		ElseIf m_intPollCol = 0 Then 
			
			txtActiveCell.Visible = False
			
		End If
		
	End Sub
	
	Private Sub txtActiveCell_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtActiveCell.Validating
		Dim Cancel As Boolean = eventArgs.Cancel
		
		If Not boolChanged Then
			
			If txtActiveCell.Text <> m_strUndoText Then
				boolChanged = True
				CmdSaveEnabled()
			End If
			
		End If
		eventArgs.Cancel = Cancel
	End Sub
	
	'UPGRADE_WARNING: Event txtActiveCellWQStd.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtActiveCellWQStd_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtActiveCellWQStd.TextChanged
		
		grdWQStd.Text = txtActiveCellWQStd.Text
		
	End Sub
	
	'Just making the entire cell text selected
	Private Sub txtActiveCellWQStd_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtActiveCellWQStd.Enter
		
		txtActiveCell.SelectionStart = 0
		txtActiveCellWQStd.SelectionLength = Len(txtActiveCellWQStd.Text)
		m_strUndoText = txtActiveCell.Text
		
	End Sub
	
	Public Sub mnuCoeffCopySet_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuCoeffCopySet.Click
		
		g_boolCopyCoeff = False
		'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
		Load(frmCopyCoeffSet)
		VB6.ShowForm(frmCopyCoeffSet, VB6.FormShowConstants.Modal, Me)
		
	End Sub
	
	Public Sub mnuCoeffNewSet_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuCoeffNewSet.Click
		
		g_boolAddCoeff = False
		'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
		Load(frmAddCoeffSet)
		VB6.ShowForm(frmAddCoeffSet, VB6.FormShowConstants.Modal, Me)
		
	End Sub
	
	Private Sub SSTab1_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles SSTab1.SelectedIndexChanged
		Static PreviousTab As Short = SSTab1.SelectedIndex()
		
		Select Case SSTab1.SelectedIndex
			Case 0
				txtActiveCellWQStd.Visible = False
			Case 1
				txtActiveCell.Visible = False
		End Select
		
		PreviousTab = SSTab1.SelectedIndex()
	End Sub
	
	'**********************************************************************************************************
	'Subs/Functions
	'**********************************************************************************************************
	Public Sub CopyCoefficient(ByRef strNewCoeffName As String, ByRef strCoeffSet As String)
		
		'General gist:  First we add new record to the Coefficient Set table using strNewCoeffName as
		'the name, PollID, LCTYPEID.  Once that's done, we'll add the coefficients
		'from the set being copied
		Dim strCopySet As String 'The Recordset of existing coefficients being copied
		Dim strLandClass As String 'Select for Landclass
		Dim i As Short
		
		Dim rsCopySet As New ADODB.Recordset 'First RS
		Dim rsLandClass As New ADODB.Recordset '1 record recordset to get a hold of LandClass
		
		strCopySet = "SELECT * FROM COEFFICIENTSET INNER JOIN COEFFICIENT ON COEFFICIENTSET.COEFFSETID = " & "COEFFICIENT.COEFFSETID WHERE COEFFICIENTSET.NAME LIKE '" & strCoeffSet & "'"
		
		rsCopySet.Open(strCopySet, g_ADOConn, ADODB.CursorTypeEnum.adOpenKeyset)
		
		
		Debug.Print(rsCopySet.RecordCount)
		
		'Step 1: Enter name
		txtCoeffSet.Text = strNewCoeffName
		
		'Clear things and set the rows to recordcount + 1, remember 1st row fixed
		grdPollDef.Clear()
		grdPollDef.Rows = rsCopySet.RecordCount + 1
		
		rsCopySet.MoveFirst()
		
		'Actually add the records to the new set
		For i = 1 To rsCopySet.RecordCount
			strLandClass = "SELECT * FROM LCCLASS WHERE LCCLASSID = " & rsCopySet.Fields("LCClassID").Value
			'Let's try one more ADO method, why not, righ?
			rsLandClass.Open(strLandClass, g_ADOConn, ADODB.CursorTypeEnum.adOpenKeyset)
			
			'Add the necessary components
			With grdPollDef
				.set_TextMatrix(i, 1, rsLandClass.Fields("Value").Value)
				.set_TextMatrix(i, 2, rsLandClass.Fields("Name").Value)
				.set_TextMatrix(i, 3, rsCopySet.Fields("Coeff1").Value)
				.set_TextMatrix(i, 4, rsCopySet.Fields("Coeff2").Value)
				.set_TextMatrix(i, 5, rsCopySet.Fields("Coeff3").Value)
				.set_TextMatrix(i, 6, rsCopySet.Fields("Coeff4").Value)
				
			End With
			rsLandClass.Close()
			rsCopySet.MoveNext()
		Next i
		
		
		'Set up everything to look good
		frmCopyCoeffSet.Close()
		
		
		'Cleanup
		rsCopySet.Close()
		
		'UPGRADE_NOTE: Object rsCopySet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsCopySet = Nothing
		'UPGRADE_NOTE: Object rsLandClass may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsLandClass = Nothing
		
	End Sub
	
	
	Private Sub CmdSaveEnabled()
		If boolChanged Or boolDescChanged Then
			cmdSave.Enabled = True
		Else
			cmdSave.Enabled = False
		End If
	End Sub
	
	'Grouped all of the checks here.
	Private Function CheckForm() As Boolean
		
		If Trim(txtPollutant.Text) = "" Then
			MsgBox("Please enter a name for the new pollutant", MsgBoxStyle.Critical, "Enter Name")
			CheckForm = False
			txtPollutant.Focus()
			txtPollutant.SelectionLength = Len(txtPollutant.Text)
			Exit Function
		ElseIf modUtil.UniqueName("Pollutant", txtPollutant.Text) Then 
			CheckForm = True
		Else
			MsgBox(Err4, MsgBoxStyle.Critical, "Name In Use")
			CheckForm = False
			txtPollutant.Focus()
			txtPollutant.SelectionLength = Len(txtPollutant.Text)
			Exit Function
		End If
		
		If Len(Trim(txtCoeffSet.Text)) = 0 Then
			MsgBox("Please enter a name for the new pollutant", MsgBoxStyle.Critical, "Enter Name")
			CheckForm = False
			txtPollutant.Focus()
			txtPollutant.SelectionLength = Len(txtPollutant.Text)
			Exit Function
		ElseIf modUtil.UniqueName("Coefficientset", (txtCoeffSet.Text)) Then 
			CheckForm = True
		Else
			MsgBox(Err4, MsgBoxStyle.Critical, "Name In Use")
			CheckForm = False
			txtCoeffSet.Focus()
			txtCoeffSet.SelectionLength = Len(txtPollutant.Text)
			Exit Function
		End If
		
		'Now if all is there and good, go on and check the grid values
		If ValidateGridValues Then
			CheckForm = True
		Else
			CheckForm = False
		End If
		
	End Function
	
	
	Private Function ValidateGridValues() As Boolean
		
		'Need to validate each grid value before saving.  Essentially we take it a row at a time,
		'then rifle through each column of each row.  Case Select tests each each x,y value depending
		'on column... 3-6 must be 1-100 range
		
		'Returns: True or False
		
		Dim varActive As Object 'txtActiveCell value
		Dim i As Short
		Dim j As Short
		
		For i = 1 To grdPollDef.Rows - 1
			
			For j = 3 To 6
				
				'UPGRADE_WARNING: Couldn't resolve default property of object varActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				varActive = grdPollDef.get_TextMatrix(i, j)
				
				'UPGRADE_WARNING: Couldn't resolve default property of object varActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If InStr(1, CStr(varActive), ".", CompareMethod.Text) > 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object varActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If (Len(Split(CStr(varActive), ".")(1)) > 4) Then
						ErrorGenerator(Err6, i, j)
						grdPollDef.col = j
						grdPollDef.row = i
						ValidateGridValues = False
						KeyMoveUpdate()
						Exit Function
					End If
				End If
				
				'UPGRADE_WARNING: Couldn't resolve default property of object varActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Not IsNumeric(varActive) Or (varActive < 0) Or (varActive > 1000) Then
					ErrorGenerator(Err6, i, j)
					grdPollDef.col = j
					grdPollDef.row = i
					ValidateGridValues = False
					KeyMoveUpdate()
					Exit Function
				End If
				
				
				
			Next j
			
		Next i
		
		ValidateGridValues = True
		
	End Function
	
	Private Function UpdateValues() As Boolean
		
		Dim strInsertPollutant As String 'Insert String for new poll
		Dim strSelectPollutant As String 'Select string for new poll
		Dim strSelectLCType As String 'Select string for LCType
		Dim strInsertCoeffSet As String 'Insert string for new coeff set
		Dim strSelectCoeffSet As String 'Select string for new coeff set
		Dim strInsertCoeffs As String 'Insert string for the coefficients
		Dim strSelectWQCrit As String 'Select string for Water Quality
		Dim strInsertWQCrit As String 'Insert string for Water Quality
		Dim strNewColor As String 'New color
		
		Dim rsNewPollutant As New ADODB.Recordset
		Dim rsNewCoeffSet As New ADODB.Recordset
		Dim rsLCType As New ADODB.Recordset
		Dim rsWQCrit As New ADODB.Recordset
		
		Dim i As Short
		Dim j As Short
		Dim k As Short
		
		On Error GoTo ErrHandler
		'Step 1a: Get a new color for this pollutant
		strNewColor = modUtil.ReturnHSVColorString
		
		'Step 1: Insert the New Pollutant
		strInsertPollutant = "INSERT INTO POLLUTANT(NAME, POLLTYPE, COLOR) VALUES ('" & Replace(Trim(txtPollutant.Text), "'", "''") & "', 0, " & "'" & strNewColor & "'" & ")"
		
		g_ADOConn.Execute(strInsertPollutant, ADODB.AffectEnum.adAffectCurrent)
		
		'Step 2: Select the newly inserted pollutant info
		strSelectPollutant = "SELECT * FROM POLLUTANT WHERE NAME LIKE '" & Trim(txtPollutant.Text) & "'"
		rsNewPollutant.Open(strSelectPollutant, g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic)
		
		'Step 2a: Select the WQ Standards
		strSelectWQCrit = "SELECT * FROM WQCriteria"
		rsWQCrit.Open(strSelectWQCrit, g_ADOConn, ADODB.CursorTypeEnum.adOpenKeyset)
		
		rsWQCrit.MoveFirst()
		
		For k = 1 To rsWQCrit.RecordCount
			
			strInsertWQCrit = "INSERT INTO POLL_WQCRITERIA (POLLID, WQCRITID, THRESHOLD) VALUES (" & rsNewPollutant.Fields("POLLID").Value & "," & rsWQCrit.Fields("WQCRITID").Value & "," & "0 )"
			g_ADOConn.Execute(strInsertWQCrit, ADODB.AffectEnum.adAffectCurrent)
			
			rsWQCrit.MoveNext()
		Next k
		
		'Step 3: Get the LCtype information
		strSelectLCType = "SELECT * FROM LCTYPE WHERE NAME LIKE '" & cboLCType.Text & "'"
		rsLCType.Open(strSelectLCType, g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic)
		
		'Step 4: Insert the New coefficient set
		strInsertCoeffSet = "INSERT INTO COEFFICIENTSET (NAME, DESCRIPTION, LCTYPEID, POLLID) VALUES (" & Replace(Trim(txtCoeffSet.Text), "'", "''") & "', '" & Replace(Trim(txtCoeffSetDesc.Text), "'", "''") & "'," & rsLCType.Fields("LCTypeID").Value & "," & rsNewPollutant.Fields("POLLID").Value & ")"
		g_ADOConn.Execute(strInsertCoeffSet, ADODB.AffectEnum.adAffectCurrent)
		
		'Step 5: Select the newly inserted coefficient set
		strSelectCoeffSet = "SELECT * FROM COEFFICIENTSET WHERE NAME LIKE '" & txtCoeffSet.Text & "'"
		rsNewCoeffSet.Open(strSelectCoeffSet, g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic)
		
		'Step 6: Insert the new coeffs for that set
		For i = 1 To grdPollDef.Rows - 1
			
			strInsertCoeffs = "INSERT INTO COEFFICIENT (COEFF1, COEFF2, COEFF3, COEFF4, COEFFSETID, LCCLASSID) VALUES (" & grdPollDef.get_TextMatrix(i, 3) & ", " & grdPollDef.get_TextMatrix(i, 4) & ", " & grdPollDef.get_TextMatrix(i, 5) & ", " & grdPollDef.get_TextMatrix(i, 6) & ", " & rsNewCoeffSet.Fields("CoeffSetID").Value & ", " & grdPollDef.get_TextMatrix(i, 8) & ")"
			
			g_ADOConn.Execute(strInsertCoeffs, ADODB.AffectEnum.adAffectCurrent)
			
		Next i
		
		'Cleanup
		rsNewPollutant.Close()
		rsLCType.Close()
		rsNewCoeffSet.Close()
		
		'UPGRADE_NOTE: Object rsNewPollutant may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsNewPollutant = Nothing
		'UPGRADE_NOTE: Object rsLCType may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsLCType = Nothing
		'UPGRADE_NOTE: Object rsNewCoeffSet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsNewCoeffSet = Nothing
		
		frmPollutants.cboPollName.Items.Clear()
		modUtil.InitComboBox((frmPollutants.cboPollName), "Pollutant")
		frmPollutants.cboPollName.SelectedIndex = modUtil.GetCboIndex((txtPollutant.Text), (frmPollutants.cboPollName))
		
		UpdateValues = True
		
		Exit Function
ErrHandler: 
		
		MsgBox("An error occurred while creating new pollutant." & vbNewLine & Err.Number & ": " & Err.Description, MsgBoxStyle.Critical, "Error")
		
	End Function
	
	Public Sub AddCoefficient(ByRef strCoeffName As String, ByRef strLCType As String)
		
		'General gist:  First we add new record to the Coefficient Set table using strCoeffName as
		'the name, m_intPollID as the PollID, and m_intLCTYPEID as the LCTypeID.  The last two are
		'garnered above during a cbo click event.  Once that's done, we'll add a series of blank
		'coefficients for the landclass type the user chooses...ie CCAP, NotCCAP, whatever
		
		Dim strNewLcType As String 'CmdString for inserting new coefficientset
		Dim strDefault As String '
		Dim strNewCoeff As String
		Dim strNewCoeffID As String 'Holder for the CoefficientSetID
		Dim strInsertNewCoeff As String 'Putting the newly created coefficients in their table
		Dim intCoeffSetID As Short
		Dim i As Short
		
		Dim rsCopySet As ADODB.Recordset 'First RS
		Dim rsCoeffSetID As New ADODB.Recordset '1 record recordset to get a hold of CoeffsetID
		Dim rsNewCoeff As New ADODB.Recordset
		
		strNewLcType = "INSERT INTO COEFFICIENTSET(NAME, POLLID, LCTYPEID) VALUES ('" & Replace(strCoeffName, "'", "''") & "'," & m_intPollID & "," & m_intLCTypeID & ")"
		
		'First need to add the coefficient set to that table
		g_ADOConn.Execute(strNewLcType, ADODB.AffectEnum.adAffectCurrent)
		
		'Get the Coefficient Set ID of the newly created coefficient set to populate Column # 8 in the GRid,
		'which by the way, is hidden from view.  InitPollDef sets the widths of col 7, 8 to 0
		strNewCoeffID = "SELECT COEFFSETID FROM COEFFICIENTSET " & "WHERE COEFFICIENTSET.NAME LIKE '" & strCoeffName & "'"
		
		rsCoeffSetID.Open(strNewCoeffID, g_ADOConn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
		intCoeffSetID = rsCoeffSetID.Fields("CoeffSetID").Value
		
		strDefault = "SELECT LCTYPE.LCTYPEID, LCCLASS.LCCLASSID, LCCLASS.NAME As valName, " & "LCCLASS.VAlue as valValue FROM LCTYPE " & "INNER JOIN LCCLASS ON LCCLASS.LCTYPEID = LCTYPE.LCTYPEID " & "WHERE LCTYPE.Name Like " & "'" & strLCType & "'"
		
		rsCopySet = New ADODB.Recordset
		rsCopySet.Open(strDefault, g_ADOConn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
		
		'Clear things and set the rows to recordcount + 1, remember 1st row fixed
		grdPollDef.Clear()
		grdPollDef.Rows = rsCopySet.RecordCount + 1
		
		
		rsCopySet.MoveFirst()
		modUtil.InitPollDefGrid(grdPollDef) 'Call this again to set it up as we cleared it
		
		Dim cmdInsert As New ADODB.Command
		
		'Now loopy loo to populate values.
		Dim strNewCoeff1 As String
		strNewCoeff1 = "SELECT * FROM COEFFICIENT"
		
		rsNewCoeff.Open(strNewCoeff1, g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
		
		For i = 1 To rsCopySet.RecordCount
			
			'Let's try one more ADO method, why not, righ?
			rsNewCoeff.AddNew()
			
			'Add the necessary components
			rsNewCoeff.Fields("Coeff1").Value = 0
			rsNewCoeff.Fields("Coeff2").Value = 0
			rsNewCoeff.Fields("Coeff3").Value = 0
			rsNewCoeff.Fields("Coeff4").Value = 0
			rsNewCoeff.Fields("CoeffSetID").Value = rsCoeffSetID.Fields("CoeffSetID").Value
			rsNewCoeff.Fields("LCClassID").Value = rsCopySet.Fields("LCClassID").Value
			
			With grdPollDef
				
				.set_TextMatrix(i, 1, rsCopySet.Fields("valValue").Value)
				.set_TextMatrix(i, 2, rsCopySet.Fields("valName").Value)
				.set_TextMatrix(i, 3, "0")
				.set_TextMatrix(i, 4, "0")
				.set_TextMatrix(i, 5, "0")
				.set_TextMatrix(i, 6, "0")
				.set_TextMatrix(i, 7, rsCoeffSetID.Fields("CoeffSetID").Value)
				.set_TextMatrix(i, 8, rsNewCoeff.Fields("coeffID").Value)
				'.TextMatrix(i, 9) = rsNewCoeff!CoeffID
				
			End With
			rsCopySet.MoveNext()
			
		Next i
		
		'Call the function to set everything to newly added Coefficient.
		'cboCoeffSet.ListIndex = GetCboIndex(strCoeffName, cboCoeffSet)
		
		'Cleanup
		rsCopySet.Close()
		rsCoeffSetID.Close()
		'rsNewCoeff.Close
		
		'UPGRADE_NOTE: Object rsCopySet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsCopySet = Nothing
		'UPGRADE_NOTE: Object rsCoeffSetID may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsCoeffSetID = Nothing
		'UPGRADE_NOTE: Object rsNewCoeff may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsNewCoeff = Nothing
		
		Me.Close()
		
	End Sub
End Class