Option Strict Off
Option Explicit On
Friend Class frmLUScen
	Inherits System.Windows.Forms.Form
	' *************************************************************************************
	' *  Perot Systems Government Services
	' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
	' *  frmLUScen
	' *************************************************************************************
	' *  Description: Form that takes in the user's stuff for a new landuse scenario.  Stores
	' *  the contents in an xml string in a hidden cell in 3rd column in grdLU on frmPrj.
	' *  This is so folks can make use of the Edit Scenario... menu choice as well.
	' *
	' *  Called By: frmPrj::Add/Edit Landuse scenario
	' *************************************************************************************
	' *  Subs:
	' *     Function CreateXMLFile:: creates XML string holding the forms info
	' *     Function ValidateData:: You guessed it.  Validates data.
	' *     Sub PopulateForm:: Throws values of the hidden cell in grdLU into form
	' *
	' *  Misc: employs clsPopUp, an API use of a popup menu.  Workaround for VB does not
	' *  support multiple popups on the same form.
	' *************************************************************************************
	
	
	Private m_App As ESRI.ArcGIS.Framework.IApplication
	Private m_pMap As ESRI.ArcGIS.Carto.IMap
	Private m_intRow As Short
	Private m_intCol As Short
	Private m_strUndoText As String
	Private clsManScen As clsXMLLUScen
	Private m_strWQStd As String
	
	' Variables used by the Error handler function - DO NOT REMOVE
	Const c_sModuleFileName As String = "frmLUScen.frm"
	
	'UPGRADE_WARNING: Event cboPollName.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboPollName_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboPollName.SelectedIndexChanged
		
		If m_intRow > 0 Then
			grdPoll.set_TextMatrix(m_intRow, m_intCol, cboPollName.Text)
			cboPollName.Visible = False
		End If
		
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		
		Me.Close()
		
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		
		If ValidateData Then 'ToDO: Validate data
			clsManScen.SaveFile(CreateXMLFile)
			Me.Close()
		End If
		
	End Sub
	
	Private Function CreateXMLFile() As String
		
		Dim i As Short
		Dim clsMan As New clsXMLLUScen
		
		With clsMan
			
			.strLUScenName = Trim(txtLUName.Text)
			.strLUScenLyrName = Trim(cboLULayer.Text)
			.strLUScenFileName = modUtil.GetFeatureFileName(.strLUScenLyrName, m_App)
			.intLUScenSelectedPoly = chkSelectedPolys.CheckState
			.intSCSCurveA = CDbl(txtLUCN(0).Text)
			.intSCSCurveB = CDbl(txtLUCN(1).Text)
			.intSCSCurveC = CDbl(txtLUCN(2).Text)
			.intSCSCurveD = CDbl(txtLUCN(3).Text)
			.lngCoverFactor = CDbl(txtLUCN(4).Text)
			.intWaterWetlands = chkWatWetlands.CheckState
			
			For i = 1 To grdPoll.Rows - 1
				clsMan.clsPollutant = New clsXMLLUScenPollItem
				.clsPollutant.intID = i
				.clsPollutant.strPollName = grdPoll.get_TextMatrix(i, 1)
				.clsPollutant.intType1 = CDbl(grdPoll.get_TextMatrix(i, 2))
				.clsPollutant.intType2 = CDbl(grdPoll.get_TextMatrix(i, 3))
				.clsPollutant.intType3 = CDbl(grdPoll.get_TextMatrix(i, 4))
				.clsPollutant.intType4 = CDbl(grdPoll.get_TextMatrix(i, 5))
				.clsPollItems.Add(.clsPollutant)
			Next i
			
		End With
		
		frmPrj.grdLU.set_TextMatrix(CInt(g_intManScenRow), 2, clsMan.strLUScenName)
		CreateXMLFile = clsMan.XML
		
	End Function
	
	
	Private Sub frmLUScen_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		modUtil.AddFeatureLayerToComboBox(cboLULayer, m_pMap, "poly")
		FillGrid()
		
		'define flexgrid for Land Use tab
		With grdPoll
			.col = .FixedCols
			.row = .FixedRows
			.Width = VB6.TwipsToPixelsX(5600 + (.GridLineWidth * (.get_Cols() + 1)) + 75)
			.set_ColWidth(0,  , 400)
			.set_ColWidth(1,  , 2000)
			.set_ColWidth(2,  , 800)
			.set_ColWidth(3,  , 800)
			.set_ColWidth(4,  , 800)
			.set_ColWidth(5,  , 800)
			.row = 0
			.col = 0
			.Text = ""
			.col = 1
			.Text = "Pollutant"
			.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
			.col = 2
			.Text = "Type 1"
			.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
			.col = 3
			.Text = "Type 2"
			.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
			.col = 4
			.Text = "Type 3"
			.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
			.col = 5
			.Text = "Type 4"
			.CellAlignment = MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter
			.Enabled = True
		End With
		
		clsManScen = New clsXMLLUScen
		
		If Len(g_strLUScenFileName) > 0 Then
			clsManScen.XML = g_strLUScenFileName
			grdPoll.Rows = clsManScen.clsPollItems.Count + 1
			PopulateForm()
		Else
			txtLUCN(0).Text = "0"
			txtLUCN(1).Text = "0"
			txtLUCN(2).Text = "0"
			txtLUCN(3).Text = "0"
		End If
		
	End Sub
	
	Public Sub init(ByVal pApp As ESRI.ArcGIS.Framework.IApplication, ByRef strWQStd As String)
		
		Dim pMxDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
		m_App = pApp
		
		pMxDoc = m_App.Document
		m_pMap = pMxDoc.FocusMap
		
		m_strWQStd = strWQStd
		
		
		
	End Sub
	
	Private Sub frmLUScen_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		
		'UPGRADE_NOTE: Object clsManScen may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		clsManScen = Nothing
		m_intRow = 0
		m_intCol = 0
		
	End Sub
	
	Private Sub grdPoll_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles grdPoll.ClickEvent
		
		m_intRow = grdPoll.MouseRow
		m_intCol = grdPoll.MouseCol
		
		If m_intCol > 1 And m_intRow >= 1 Then
			With txtTypes
				.Visible = True
				.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(grdPoll.Left) + grdPoll.CellLeft), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(grdPoll.Top) + grdPoll.CellTop), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
				.Width = VB6.TwipsToPixelsX(grdPoll.CellWidth - 30)
				.Height = VB6.TwipsToPixelsY(grdPoll.CellHeight - 30)
				.Text = grdPoll.get_TextMatrix(m_intRow, m_intCol)
			End With
			
		Else
			If m_intCol < 1 Then
				txtTypes.Visible = False
				cboPollName.Visible = False
				
			End If
		End If
		
	End Sub
	
	Private Sub grdPoll_MouseDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_MouseDownEvent) Handles grdPoll.MouseDownEvent
		On Error GoTo ErrorHandler
		
		m_intRow = grdPoll.MouseRow
		m_intCol = grdPoll.MouseCol
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "grdPoll_MouseDown " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	Private Sub AlterGrid(ByRef lngAlterType As Integer)
		Dim row, R, C, col As Short
		
		Dim lngLCClassID As Integer
		Select Case lngAlterType
			Case 1 'Append
				
				With grdPoll
					.Rows = .Rows + 1
					.row = .Rows - 1
				End With
				
			Case 2 'Insert
				
				With grdPoll
					If .row < .FixedRows Then 'make sure we don't insert above header Rows
						.Rows = .Rows + 1
						.row = .Rows - 1
					Else
						R = .row
						.Rows = .Rows + 1 'add a row
						
						For row = .Rows - 1 To R + 1 Step -1 'move data dn 1 row
							For col = 1 To .get_Cols() - 1
								.set_TextMatrix(row, col, .get_TextMatrix(row - 1, col))
							Next col
						Next row
						For col = 1 To .get_Cols() - 1 ' clear all cells in this row
							.set_TextMatrix(R, col, "")
						Next col
					End If
				End With
				
			Case 3 'Delete
				
				
				With grdPoll
					
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
		End Select
		
	End Sub
	
	'UPGRADE_WARNING: Event txtTypes.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtTypes_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTypes.TextChanged
		
		grdPoll.set_TextMatrix(m_intRow, m_intCol, txtTypes.Text)
		
	End Sub
	
	Private Sub txtTypes_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTypes.Enter
		
		txtTypes.SelectionStart = 0
		txtTypes.SelectionLength = Len(txtTypes.Text)
		m_strUndoText = txtTypes.Text
		
	End Sub
	
	Private Sub txtTypes_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtTypes.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		With grdPoll
			Select Case KeyCode
				Case System.Windows.Forms.Keys.Escape 'if the user pressed escape, then get out without changing
					.Text = m_strUndoText
					txtTypes.Visible = False
					.Focus()
					
				Case 13 'if the user presses enter, get out of the textbox
					.Text = txtTypes.Text
					txtTypes.Visible = False
					
				Case System.Windows.Forms.Keys.Up 'if the user presses the up arrow, get out of the textbox and move up a cell on the grid
					.Text = txtTypes.Text
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
					.Text = txtTypes.Text
					txtTypes.Visible = False
					If .row < grdPoll.Rows - 1 Then
						.row = .row + 1
						KeyMoveUpdate()
					Else
						If .col > 1 Then
							.row = 1 'again, if the row is on the last row, don't move cells
							.col = .col - 1
						Else
							.Text = txtTypes.Text
							txtTypes.Visible = False
						End If
						KeyMoveUpdate()
					End If
				Case System.Windows.Forms.Keys.Left 'if the user pressed the right arrow key, get out of the textbox and move right one cell in the grid
					txtTypes.Visible = False
					If .col >= 2 Then
						.col = .col - 1
						KeyMoveUpdate()
						
					Else
						
						If .row > 1 And .col = 2 Then
							.col = 5
							.row = .row - 1
							KeyMoveUpdate()
						End If
					End If
					
				Case System.Windows.Forms.Keys.Right 'if the user pressed the right arrow key, get out of the textbox and move right one cell in the grid
					.Text = txtTypes.Text
					txtTypes.Visible = False
					If .col < 5 Then
						.col = .col + 1
						KeyMoveUpdate()
					Else
						If .row < grdPoll.Rows - 1 Then
							.col = 2
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
		
		m_intRow = grdPoll.row
		m_intCol = grdPoll.col
		
		txtTypes.Visible = False
		
		If (m_intCol >= 2 And m_intCol <= 5) And (m_intRow >= 1) Then
			
			With txtTypes
				
				.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(grdPoll.Left) + grdPoll.CellLeft), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(grdPoll.Top) + grdPoll.CellTop), VB6.TwipsToPixelsX(grdPoll.CellWidth - 30), VB6.TwipsToPixelsY(grdPoll.CellHeight - 150))
				
				If m_intRow <> 0 Then
					
					.Visible = True
					m_strUndoText = grdPoll.get_TextMatrix(m_intRow, m_intCol)
					.Text = m_strUndoText
					.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
				End If
			End With
			
			txtTypes.Focus()
			
		ElseIf m_intCol = 0 Then 
			
			txtTypes.Visible = False
			
		End If
		
	End Sub
	
	Private Function ValidateData() As Boolean
		
		Dim i As Short
		
		'Project Name
		If Len(txtLUName.Text) = 0 Or Len(txtLUName.Text) > 30 Then
			MsgBox("Please enter a name for the scenario.  Names must be 30 characters or less.", MsgBoxStyle.Critical, "Enter Name")
			txtLUName.Focus()
			ValidateData = False
			Exit Function
		Else
			ValidateData = True
		End If
		
		'LandCoverLayer
		If cboLULayer.Text = "" Then
			MsgBox("Please select a layer before continuing.", MsgBoxStyle.Critical, "Select Layer")
			cboLULayer.Focus()
			ValidateData = False
			Exit Function
		Else
			If Not modUtil.LayerInMap(cboLULayer.Text, m_pMap) Then
				MsgBox("The layer you have choosen is not in the current map frame.", MsgBoxStyle.Critical, "Layer Not Found")
				ValidateData = False
				Exit Function
			End If
		End If
		
		'Check selected polygons
		If chkSelectedPolys.CheckState = 1 Then
			If m_pMap.SelectionCount = 0 Then
				MsgBox("You have chosen to use selected polygons from " & cboLULayer.Text & ", but the current map contains no selected features." & vbNewLine & "Please select features or N-SPECT will use the entire extent of " & cboLULayer.Text & " to apply this landuse scenario.", MsgBoxStyle.Information, "No Selected Features Found")
				ValidateData = False
			End If
		End If
		
		'SCS Curve Numbers
		'UPGRADE_WARNING: Couldn't resolve default property of object txtLUCN().UBound. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For i = 0 To txtLUCN().UBound
			If IsNumeric(Trim(txtLUCN(i).Text)) Then
				If CShort(txtLUCN(i).Text) > 0 Or CShort(txtLUCN(i).Text) <= 1 Then
					ValidateData = True
				End If
			Else
				MsgBox("SCS Values are to be numeric only in the range of 0 - 1.", MsgBoxStyle.Critical, "Check SCS Values")
				ValidateData = False
				txtLUCN(i).Focus()
				Exit Function
			End If
		Next i
		
	End Function
	
	Private Sub PopulateForm()
		
		Dim strScenName As String
		Dim strLyrName As String
		Dim i As Short
		
		Dim clsPollItem As clsXMLLUScenPollItem
		
		strScenName = clsManScen.strLUScenName
		strLyrName = clsManScen.strLUScenLyrName
		
		txtLUName.Text = strScenName
		
		If modUtil.LayerInMap(strLyrName, m_pMap) Then
			cboLULayer.SelectedIndex = modUtil.GetCboIndex(strLyrName, cboLULayer)
		End If
		
		chkSelectedPolys.CheckState = clsManScen.intLUScenSelectedPoly
		
		txtLUCN(0).Text = CStr(clsManScen.intSCSCurveA)
		txtLUCN(1).Text = CStr(clsManScen.intSCSCurveB)
		txtLUCN(2).Text = CStr(clsManScen.intSCSCurveC)
		txtLUCN(3).Text = CStr(clsManScen.intSCSCurveD)
		txtLUCN(4).Text = CStr(clsManScen.lngCoverFactor)
		chkWatWetlands.CheckState = clsManScen.intWaterWetlands
		
		grdPoll.Rows = clsManScen.clsPollItems.Count + 1
		
		For i = 1 To clsManScen.clsPollItems.Count
			With grdPoll
				.row = clsManScen.clsPollItems.Item(i).intID
				.set_TextMatrix(.row, 1, clsManScen.clsPollItems.Item(i).strPollName)
				.set_TextMatrix(.row, 2, CStr(clsManScen.clsPollItems.Item(i).intType1))
				.set_TextMatrix(.row, 3, CStr(clsManScen.clsPollItems.Item(i).intType2))
				.set_TextMatrix(.row, 4, CStr(clsManScen.clsPollItems.Item(i).intType3))
				.set_TextMatrix(.row, 5, CStr(clsManScen.clsPollItems.Item(i).intType4))
			End With
		Next i
		
	End Sub
	
	'Loads the variety of comboboxes throught the project using combobox and name of table
	Private Sub FillGrid()
		
		Dim strSQLWQStd As String
		Dim rsWQStdCboClick As New ADODB.Recordset
		
		Dim strSQLWQStdPoll As String
		Dim rsWQStdPoll As New ADODB.Recordset
		
		Dim i As Short
		
		'Selection based on combo box
		strSQLWQStd = "SELECT * FROM WQCRITERIA WHERE NAME LIKE '" & m_strWQStd & "'"
		rsWQStdCboClick.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		rsWQStdCboClick.Open(strSQLWQStd, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
		
		If rsWQStdCboClick.RecordCount > 0 Then
			
			strSQLWQStdPoll = "SELECT POLLUTANT.NAME, POLL_WQCRITERIA.THRESHOLD " & "FROM POLL_WQCRITERIA INNER JOIN POLLUTANT " & "ON POLL_WQCRITERIA.POLLID = POLLUTANT.POLLID Where POLL_WQCRITERIA.WQCRITID = " & rsWQStdCboClick.Fields("WQCRITID").Value
			
			rsWQStdPoll.CursorLocation = ADODB.CursorLocationEnum.adUseClient
			rsWQStdPoll.Open(strSQLWQStdPoll, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
			
			grdPoll.Rows = rsWQStdPoll.RecordCount + 1
			
			For i = 1 To rsWQStdPoll.RecordCount
				
				grdPoll.set_TextMatrix(i, 1, rsWQStdPoll.Fields("Name").Value)
				grdPoll.set_TextMatrix(i, 2, 0)
				grdPoll.set_TextMatrix(i, 3, 0)
				grdPoll.set_TextMatrix(i, 4, 0)
				grdPoll.set_TextMatrix(i, 5, 0)
				rsWQStdPoll.MoveNext()
				
			Next i
			
			'Clean it
			'UPGRADE_NOTE: Object rsWQStdPoll may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsWQStdPoll = Nothing
			'UPGRADE_NOTE: Object rsWQStdCboClick may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsWQStdCboClick = Nothing
			
		Else
			
			MsgBox("Warning: There are no water quality standards remaining.  Please add a new one.", MsgBoxStyle.Critical, "Recordset Empty")
			'UPGRADE_NOTE: Object rsWQStdCboClick may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsWQStdCboClick = Nothing
		End If
		
	End Sub
End Class