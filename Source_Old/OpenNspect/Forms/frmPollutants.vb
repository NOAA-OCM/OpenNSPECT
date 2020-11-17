Option Strict Off
Option Explicit On
Friend Class frmPollutants
	Inherits System.Windows.Forms.Form
    '	' *************************************************************************************
    '	' *  Perot Systems Government Services
    '	' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
    '	' *  frmPollutants
    '	' *************************************************************************************
    '	' *  Description: Form for browsing pollutants
    '	' *  within NSPECT
    '	' *
    '	' *  Called By:  NSPECT clsPollutants
    '	' *************************************************************************************
    '	' *  Subs/Functions:
    '	' *     ExportCoeffSet(strFileName As String): Exports data to textfile
    '	' *     KeyMoveUpdate:  Handles <- -> ^ key input from user while in data grid
    '	' *     ValidateGridValues:  Function returns Boolean value whether or not all data is
    '	' *         ok before updating..specifically if coeffs are numeric and in range 0 - 100
    '	' *
    '	' *  Misc:
    '	' *************************************************************************************


    '	Private m_App As ESRI.ArcGIS.Framework.IApplication

    '	Private rsPollCboClick As ADODB.Recordset 'RS on cboPollName click event
    '	Private rsCoeff As ADODB.Recordset
    '	Private rsLCType As ADODB.Recordset
    '	Private rsFullCoeff As ADODB.Recordset 'RS on cboCoeffSet click event
    '	Private rsCoeffs As ADODB.Recordset 'RS that fills FlexGrid
    '	Private rsWQStds As ADODB.Recordset 'Rs that fills FelexGrid on Standards Tab

    '	Private boolLoaded As Boolean
    '	Private boolChanged As Boolean 'Bool for enabling save button
    '	Private boolDescChanged As Boolean 'Boolship for seeing if Description Changed
    '	Private boolSaved As Boolean 'Boolship for whether or not things have saved

    '	Private m_intCurFrame As Short 'Current Frame visible in SSTab
    '	Private m_intPollRow As Short 'Row Number for grdPolldef
    '	Private m_intPollCol As Short 'Column Number for grdPollDef
    '	Private intRowWQ As Short 'Row Number for grdPollWQStd
    '	Private intColWQ As Short 'Column Number for grdPollDef
    '	Private m_intPollID As Short 'There's a need to have the PollID so we'll store it here
    '	Private m_intLCTypeID As Short 'Land Class (CCAP) ID - needed to add new coefficient sets
    '	Private m_intCoeffID As Short 'Key for CoefficientSetID - needed to add new coefficients 'See above

    '	Private m_strUndoText As String 'Text for txtActiveCell     |
    '	Private m_strUndoTextWQ As String 'Text for txtActiveCellWQ   |-all three used to detect change
    '	Private m_strUndoDesc As String 'Text for Description       |
    '	Private m_strLCType As String 'Need for name, we'll store here


    '	' Variables used by the Error handler function - DO NOT REMOVE
    '	Const c_sModuleFileName As String = "frmPollutants.frm"


    '	Private Sub cmdSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSave.Click
    '		On Error GoTo ErrorHandler


    '		If ValidateGridValues Then

    '			UpdateValues()
    '			boolSaved = True
    '			MsgBox(cboPollName.Text & " saved successfully.", MsgBoxStyle.Information, "N-SPECT")
    '			Me.Close()

    '		End If



    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(True, "cmdSave_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub

    '	Private Sub frmPollutants_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
    '		On Error GoTo ErrorHandler


    '		boolChanged = False

    '		'Initialize the two flex Grids
    '		InitPollDefGrid(grdPollDef)
    '		InitPollWQStdGrid(grdWQStd)

    '		'Toss in the names of all pollutants and call the cbo click event
    '		InitComboBox(cboPollName, "Pollutant")

    '		SSTab1.SelectedIndex = 0
    '		boolLoaded = True
    '		boolChanged = False
    '		boolSaved = False


    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub

    '	'UPGRADE_WARNING: Event cboPollName.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    '	Private Sub cboPollName_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboPollName.SelectedIndexChanged

    '		Dim strSQLPollutant As String
    '		Dim strSQLLCType As String
    '		Dim strSQLCoeff As String
    '		Dim strSQLWQStd As String

    '		On Error GoTo ErrorHandler


    '		'Check to see if things have changed
    '		If boolChanged Then

    '			intYesNo = MsgBox(strYesNo, MsgBoxStyle.YesNo, strYesNoTitle)

    '			If intYesNo = MsgBoxResult.Yes Then

    '				UpdateValues()

    '				boolChanged = False

    '				rsPollCboClick = New ADODB.Recordset

    '				'Selection based on combo box
    '				strSQLPollutant = "SELECT * FROM POLLUTANT WHERE NAME = '" & cboPollName.Text & "'"
    '				rsPollCboClick.Open(strSQLPollutant, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
    '				m_intPollID = rsPollCboClick.Fields("POLLID").Value

    '				rsCoeff = New ADODB.Recordset

    '				strSQLCoeff = "SELECT * FROM CoefficientSet WHERE POLLID = " & rsPollCboClick.Fields("POLLID").Value & ""
    '				rsCoeff.Open(strSQLCoeff, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)

    '				rsLCType = New ADODB.Recordset

    '				strSQLLCType = "SELECT NAME, LCTYPEID FROM LCTYPE WHERE LCTYPEID = " & rsCoeff.Fields("LCTypeID").Value & ""
    '				rsLCType.Open(strSQLLCType, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
    '				m_strLCType = rsLCType.Fields("Name").Value
    '				m_intLCTypeID = rsLCType.Fields("LCTypeID").Value

    '				'Fill everything based on that
    '				cboCoeffSet.Items.Clear()
    '				Do Until rsCoeff.EOF
    '					cboCoeffSet.Items.Add(rsCoeff.Fields("Name").Value)
    '					rsCoeff.MoveNext()
    '				Loop 

    '				cboCoeffSet.SelectedIndex = 0

    '				txtLCType.Text = m_strLCType

    '				'Fill the Water Quality Standards Tab
    '				rsWQStds = New ADODB.Recordset

    '				strSQLWQStd = "SELECT WQCRITERIA.Name, WQCRITERIA.Description," & "POLL_WQCRITERIA.Threshold, POLL_WQCRITERIA.POLL_WQCRITID FROM WQCRITERIA " & "LEFT OUTER JOIN POLL_WQCRITERIA ON WQCRITERIA.WQCRITID = POLL_WQCRITERIA.WQCRITID " & "Where POLL_WQCRITERIA.POLLID = " & rsPollCboClick.Fields("POLLID").Value

    '				rsWQStds.Open(strSQLWQStd, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
    '				grdWQStd.Recordset = rsWQStds

    '				'Cleanup
    '				'UPGRADE_NOTE: Object rsPollCboClick may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '				rsPollCboClick = Nothing
    '				'Set rsCoeff = Nothing
    '				'UPGRADE_NOTE: Object rsLCType may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '				rsLCType = Nothing
    '				'UPGRADE_NOTE: Object rsWQStds may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '				rsWQStds = Nothing

    '			ElseIf intYesNo = MsgBoxResult.No Then 

    '				boolChanged = False

    '				rsPollCboClick = New ADODB.Recordset

    '				'Selection based on combo box
    '				strSQLPollutant = "SELECT * FROM POLLUTANT WHERE NAME = '" & cboPollName.Text & "'"
    '				rsPollCboClick.Open(strSQLPollutant, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
    '				m_intPollID = rsPollCboClick.Fields("POLLID").Value

    '				rsCoeff = New ADODB.Recordset

    '				strSQLCoeff = "SELECT * FROM CoefficientSet WHERE POLLID = " & rsPollCboClick.Fields("POLLID").Value & ""
    '				rsCoeff.Open(strSQLCoeff, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)

    '				rsLCType = New ADODB.Recordset

    '				strSQLLCType = "SELECT NAME, LCTYPEID FROM LCTYPE WHERE LCTYPEID = " & rsCoeff.Fields("LCTypeID").Value & ""
    '				rsLCType.Open(strSQLLCType, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
    '				m_strLCType = rsLCType.Fields("Name").Value
    '				m_intLCTypeID = rsLCType.Fields("LCTypeID").Value

    '				'Fill everything based on that
    '				cboCoeffSet.Items.Clear()
    '				Do Until rsCoeff.EOF
    '					cboCoeffSet.Items.Add(rsCoeff.Fields("Name").Value)
    '					rsCoeff.MoveNext()
    '				Loop 

    '				cboCoeffSet.SelectedIndex = 0

    '				txtLCType.Text = m_strLCType

    '				'Fill the Water Quality Standards Tab
    '				rsWQStds = New ADODB.Recordset

    '				strSQLWQStd = "SELECT WQCRITERIA.Name, WQCRITERIA.Description," & "POLL_WQCRITERIA.Threshold, POLL_WQCRITERIA.POLL_WQCRITID FROM WQCRITERIA " & "LEFT OUTER JOIN POLL_WQCRITERIA ON WQCRITERIA.WQCRITID = POLL_WQCRITERIA.WQCRITID " & "Where POLL_WQCRITERIA.POLLID = " & rsPollCboClick.Fields("POLLID").Value

    '				rsWQStds.Open(strSQLWQStd, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
    '				grdWQStd.Recordset = rsWQStds

    '				'Cleanup
    '				'UPGRADE_NOTE: Object rsPollCboClick may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '				rsPollCboClick = Nothing
    '				'Set rsCoeff = Nothing
    '				'UPGRADE_NOTE: Object rsLCType may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '				rsLCType = Nothing
    '				'UPGRADE_NOTE: Object rsWQStds may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '				rsWQStds = Nothing

    '			End If

    '		Else

    '			boolChanged = False

    '			rsPollCboClick = New ADODB.Recordset

    '			'Selection based on combo box
    '			strSQLPollutant = "SELECT * FROM POLLUTANT WHERE NAME = '" & cboPollName.Text & "'"
    '			rsPollCboClick.Open(strSQLPollutant, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
    '			m_intPollID = rsPollCboClick.Fields("POLLID").Value

    '			rsCoeff = New ADODB.Recordset

    '			strSQLCoeff = "SELECT * FROM CoefficientSet WHERE POLLID = " & rsPollCboClick.Fields("POLLID").Value & ""
    '			rsCoeff.Open(strSQLCoeff, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)

    '			rsLCType = New ADODB.Recordset

    '			strSQLLCType = "SELECT NAME, LCTYPEID FROM LCTYPE WHERE LCTYPEID = " & rsCoeff.Fields("LCTypeID").Value & ""
    '			rsLCType.Open(strSQLLCType, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
    '			m_strLCType = rsLCType.Fields("Name").Value
    '			m_intLCTypeID = rsLCType.Fields("LCTypeID").Value

    '			'Fill everything based on that
    '			cboCoeffSet.Items.Clear()
    '			Do Until rsCoeff.EOF
    '				cboCoeffSet.Items.Add(rsCoeff.Fields("Name").Value)
    '				rsCoeff.MoveNext()
    '			Loop 

    '			cboCoeffSet.SelectedIndex = 0

    '			txtLCType.Text = m_strLCType

    '			'Fill the Water Quality Standards Tab
    '			rsWQStds = New ADODB.Recordset

    '			strSQLWQStd = "SELECT WQCRITERIA.Name, WQCRITERIA.Description," & "POLL_WQCRITERIA.Threshold, POLL_WQCRITERIA.POLL_WQCRITID FROM WQCRITERIA " & "LEFT OUTER JOIN POLL_WQCRITERIA ON WQCRITERIA.WQCRITID = POLL_WQCRITERIA.WQCRITID " & "Where POLL_WQCRITERIA.POLLID = " & rsPollCboClick.Fields("POLLID").Value

    '			rsWQStds.Open(strSQLWQStd, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
    '			grdWQStd.Recordset = rsWQStds

    '			'Cleanup
    '			'UPGRADE_NOTE: Object rsPollCboClick may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '			rsPollCboClick = Nothing
    '			'Set rsCoeff = Nothing
    '			'UPGRADE_NOTE: Object rsLCType may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '			rsLCType = Nothing
    '			'UPGRADE_NOTE: Object rsWQStds may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '			rsWQStds = Nothing

    '		End If



    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(True, "cboPollName_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub

    '	'UPGRADE_WARNING: Event cboCoeffSet.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    '	Private Sub cboCoeffSet_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboCoeffSet.SelectedIndexChanged
    '		On Error GoTo ErrorHandler


    '		txtActiveCell.Visible = False

    '		Dim strSQLFullCoeff As String
    '		Dim strSQLCoeffs As String
    '		If boolChanged Then

    '			intYesNo = MsgBox(strYesNo, MsgBoxStyle.YesNo, strYesNoTitle)

    '			If intYesNo = MsgBoxResult.Yes Then

    '				If ValidateGridValues Then
    '					UpdateValues()
    '					boolChanged = False
    '				End If
    '			Else
    '				GoTo notSave
    '			End If

    '		Else

    'notSave: 
    '			boolChanged = False
    '			grdPollDef.Clear()

    '			rsFullCoeff = New ADODB.Recordset

    '			strSQLFullCoeff = "SELECT COEFFICIENTSET.NAME, COEFFICIENTSET.DESCRIPTION, " & "COEFFICIENTSET.COEFFSETID, LCTYPE.NAME as NAME2 " & "FROM COEFFICIENTSET INNER JOIN LCTYPE " & "ON COEFFICIENTSET.LCTYPEID = LCTYPE.LCTYPEID " & "WHERE COEFFICIENTSET.NAME LIKE '" & cboCoeffSet.Text & "'"

    '			rsFullCoeff.Open(strSQLFullCoeff, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

    '			Debug.Print(rsFullCoeff.RecordCount)

    '			With txtCoeffSetDesc
    '				.Text = rsFullCoeff.Fields("Description").Value & ""
    '				.Refresh()
    '			End With

    '			txtLCType.Text = rsFullCoeff.Fields("Name2").Value & ""



    '			strSQLCoeffs = "SELECT LCCLASS.Value, LCCLASS.Name, COEFFICIENT.Coeff1 As Type1, COEFFICIENT.Coeff2 as Type2, " & "COEFFICIENT.Coeff3 as Type3, COEFFICIENT.Coeff4 as Type4, COEFFICIENT.CoeffID, COEFFICIENT.LCCLASSID " & "FROM LCCLASS LEFT OUTER JOIN COEFFICIENT " & "ON LCCLASS.LCCLASSID = COEFFICIENT.LCCLASSID " & "WHERE COEFFICIENT.COEFFSETID = " & rsFullCoeff.Fields("CoeffSetID").Value & " ORDER BY LCCLASS.VALUE"

    '			rsCoeffs = New ADODB.Recordset
    '			rsCoeffs.Open(strSQLCoeffs, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

    '			grdPollDef.Recordset = rsCoeffs

    '			'Cleanup
    '			rsCoeffs.Close()
    '			rsFullCoeff.Close()

    '			'UPGRADE_NOTE: Object rsFullCoeff may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '			rsFullCoeff = Nothing
    '			'UPGRADE_NOTE: Object rsCoeffs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '			rsCoeffs = Nothing

    '		End If



    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(True, "cboCoeffSet_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub


    '	Private Sub cmdQuit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdQuit.Click
    '		On Error GoTo ErrorHandler


    '		If (boolChanged Or boolDescChanged) And Not boolSaved Then

    '			intYesNo = MsgBox(strYesNo, MsgBoxStyle.YesNo, strYesNoTitle)

    '			If intYesNo = MsgBoxResult.Yes Then

    '				If ValidateGridValues Then
    '					UpdateValues()
    '					MsgBox("Data saved successfully.", MsgBoxStyle.OKOnly, "Save Successful")
    '					boolChanged = False
    '					boolDescChanged = False
    '					boolSaved = True
    '					Me.Close()
    '				End If

    '			Else

    '				Me.Close()

    '			End If
    '		Else

    '			Me.Close()
    '		End If



    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(True, "cmdQuit_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub


    '	Private Sub grdPollDef_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles grdPollDef.ClickEvent
    '		On Error GoTo ErrorHandler

    '		'Code for Poll Def Grid - moves txtActiveCell to appropriate cell

    '		m_intPollRow = grdPollDef.MouseRow
    '		m_intPollCol = grdPollDef.MouseCol

    '		txtActiveCell.Visible = False

    '		'Column greater than or equal to 1 and same for row
    '		If m_intPollCol >= 3 Then

    '			With txtActiveCell
    '				'Put in position
    '				.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(SSTab1.Left) + VB6.PixelsToTwipsX(grdPollDef.Left) + grdPollDef.CellLeft), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(SSTab1.Top) + VB6.PixelsToTwipsY(grdPollDef.Top) + grdPollDef.CellTop), VB6.TwipsToPixelsX(grdPollDef.CellWidth - 30), VB6.TwipsToPixelsY(grdPollDef.CellHeight - 150))

    '				If m_intPollRow <> 0 Then
    '					.Visible = True
    '					m_strUndoText = grdPollDef.get_TextMatrix(m_intPollRow, m_intPollCol)
    '					.Text = grdPollDef.get_TextMatrix(m_intPollRow, m_intPollCol)
    '					.Focus()

    '					'                If IsNumeric(m_strUndoText) Then
    '					'                    txtActiveCell.Alignment = 1
    '					'                Else
    '					'                    txtActiveCell.Alignment = 0
    '					'                End If

    '				End If

    '			End With

    '		End If



    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(True, "grdPollDef_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub


    '	Private Sub grdWQStd_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles grdWQStd.ClickEvent
    '		On Error GoTo ErrorHandler


    '		intRowWQ = grdWQStd.MouseRow
    '		intColWQ = grdWQStd.MouseCol

    '		txtActiveCellWQStd.Visible = False
    '		'We want to limit editing to only the threshold columns
    '		If intColWQ = 3 And intRowWQ >= 1 Then

    '			With txtActiveCellWQStd

    '				.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(SSTab1.Left) + VB6.PixelsToTwipsX(grdWQStd.Left) + grdWQStd.CellLeft), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(SSTab1.Top) + VB6.PixelsToTwipsY(grdWQStd.Top) + grdWQStd.CellTop), VB6.TwipsToPixelsX(grdWQStd.CellWidth - 30), VB6.TwipsToPixelsY(grdWQStd.CellHeight - 150))

    '				If intRowWQ <> 0 Then
    '					.Visible = True
    '					m_strUndoTextWQ = grdWQStd.get_TextMatrix(intRowWQ, intColWQ)
    '					.Text = grdWQStd.get_TextMatrix(intRowWQ, intColWQ)
    '					.Focus()

    '					If IsNumeric(m_strUndoTextWQ) Then
    '						txtActiveCellWQStd.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
    '					Else
    '						txtActiveCellWQStd.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
    '					End If

    '				End If

    '			End With

    '		End If



    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(True, "grdWQStd_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub



    '	Public Sub mnuCoeffHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuCoeffHelp.Click
    '		HtmlHelp(0, modUtil.g_nspectPath & "\Help\nspect.chm", HH_DISPLAY_TOPIC, "pol_coeftab.htm")
    '	End Sub

    '	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '	'Coefficients menu option processing
    '	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '	Public Sub mnuCoeffNewSet_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuCoeffNewSet.Click
    '		On Error GoTo ErrorHandler


    '		g_boolAddCoeff = True
    '		VB6.ShowForm(frmAddCoeffSet, VB6.FormShowConstants.Modal, Me)



    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(True, "mnuCoeffNewSet_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub

    '	Public Sub mnuCoeffDeleteSet_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuCoeffDeleteSet.Click

    '		'Using straight command text to rid ourselves of the dreaded coefficient sets

    '		On Error GoTo ErrHandler
    '		Dim intAns As Short
    '		intAns = MsgBox("Are you sure you want to delete the coefficient set '" & VB6.GetItemString(cboCoeffSet, cboCoeffSet.SelectedIndex) & "' associated with pollutant '" & VB6.GetItemString(cboPollName, cboPollName.SelectedIndex) & "'?", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Confirm Delete")

    '		'code to handle response
    '		Dim strDeleteCoeffSet As String
    '		If intAns = MsgBoxResult.Yes Then

    '			strDeleteCoeffSet = "DELETE * from COEFFICIENTSET WHERE NAME LIKE '" & cboCoeffSet.Text & "'"

    '			modUtil.g_ADOConn.Execute(strDeleteCoeffSet)

    '			MsgBox(cboCoeffSet.Text & " deleted.", MsgBoxStyle.OKOnly, "Record Deleted")

    '			cboPollName.Items.Clear()
    '			cboCoeffSet.Items.Clear()

    '			modUtil.InitComboBox(cboPollName, "Pollutant")

    '			Me.Refresh()

    '		Else

    '			Exit Sub

    '		End If

    '		Exit Sub

    'ErrHandler: 
    '		MsgBox("Error deleting coefficient set.", MsgBoxStyle.Critical, "Error")
    '		MsgBox(Err.Number & ": " & Err.Description)

    '	End Sub

    '	Public Sub mnuCoeffCopySet_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuCoeffCopySet.Click
    '		On Error GoTo ErrorHandler

    '		g_boolCopyCoeff = True
    '		frmCopyCoeffSet.init(rsCoeff)
    '		VB6.ShowForm(frmCopyCoeffSet, VB6.FormShowConstants.Modal, Me)

    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(True, "mnuCoeffCopySet_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub

    '	Public Sub mnuCoeffImportSet_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuCoeffImportSet.Click
    '		On Error GoTo ErrorHandler


    '		'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
    '		Load(frmImportCoeffSet)
    '		VB6.ShowForm(frmImportCoeffSet, VB6.FormShowConstants.Modal, Me)



    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(True, "mnuCoeffImportSet_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub



    '	'Export Menu
    '	Public Sub mnuCoeffExportSet_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuCoeffExportSet.Click
    '		On Error GoTo ErrorHandler

    '		Dim intAns As Short

    '		'browse...get output filename
    '		dlgCMD1Open.FileName = CStr(Nothing)
    '		dlgCMD1Save.FileName = CStr(Nothing)
    '		'UPGRADE_WARNING: CommonDialog variable was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="671167DC-EA81-475D-B690-7A40C7BF4A23"'
    '		With dlgCMD1
    '			'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
    '			.Filter = Replace(MSG1, "<name>", "Coefficient Set")
    '			.Title = Replace(MSG3, "<name>", "Coefficient Set")
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
    '			ExportCoeffSet((dlgCMD1Open.FileName))
    '		End If



    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(True, "mnuCoeffExportSet_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub

    '	'Exports your current standard and pollutants to text or csv.
    '	Private Sub ExportCoeffSet(ByRef strFileName As String)
    '		On Error GoTo ErrorHandler


    '		Dim fso As New Scripting.FileSystemObject
    '		Dim fl As Scripting.TextStream

    '		Dim rsNew As ADODB.Recordset
    '		Dim theLine As Object

    '		fl = fso.CreateTextFile(strFileName, True)

    '		Dim i As Short

    '		'Write name of pollutant and threshold
    '		For i = 1 To grdPollDef.Rows - 1
    '			fl.WriteLine(grdPollDef.get_TextMatrix(i, 1) & "," & grdPollDef.get_TextMatrix(i, 3) & "," & grdPollDef.get_TextMatrix(i, 4) & "," & grdPollDef.get_TextMatrix(i, 5) & "," & grdPollDef.get_TextMatrix(i, 6))

    '		Next i

    '		fl.Close()

    '		'UPGRADE_NOTE: Object fso may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '		fso = Nothing
    '		'UPGRADE_NOTE: Object fl may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '		fl = Nothing



    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(False, "ExportCoeffSet " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub

    '	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '	'Pollutants menu option processing
    '	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '	Public Sub mnuAddPoll_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuAddPoll.Click
    '		On Error GoTo ErrorHandler



    '		VB6.ShowForm(frmNewPollutants, VB6.FormShowConstants.Modal, Me)



    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(True, "mnuAddPoll_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub

    '	Public Sub mnuDeletePoll_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDeletePoll.Click
    '		On Error GoTo ErrorHandler


    '		Dim intAns As Short
    '		intAns = MsgBox("Are you sure you want to delete the pollutant '" & VB6.GetItemString(cboPollName, cboPollName.SelectedIndex) & "'?", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Confirm Delete")
    '		'code to handle response

    '		If intAns = MsgBoxResult.Yes Then
    '			DeletePollutant(VB6.GetItemString(cboPollName, cboPollName.SelectedIndex))
    '		Else
    '			Exit Sub
    '		End If



    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(True, "mnuDeletePoll_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub



    '	Public Sub mnuPollHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPollHelp.Click

    '		HtmlHelp(0, modUtil.g_nspectPath & "\Help\nspect.chm", HH_DISPLAY_TOPIC, "pollutants.htm")

    '	End Sub

    '	Private Sub SSTab1_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles SSTab1.SelectedIndexChanged
    '		Static PreviousTab As Short = SSTab1.SelectedIndex()
    '		On Error GoTo ErrorHandler

    '		'Clear the text box if user jumps from one tab to the next
    '		Select Case SSTab1.SelectedIndex
    '			Case 0
    '				txtActiveCellWQStd.Visible = False
    '			Case 1
    '				txtActiveCell.Visible = False
    '		End Select



    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(True, "SSTab1_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '		PreviousTab = SSTab1.SelectedIndex()
    '	End Sub

    '	'UPGRADE_WARNING: Event txtActiveCell.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    '	Private Sub txtActiveCell_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtActiveCell.TextChanged
    '		On Error GoTo ErrorHandler



    '		grdPollDef.Text = txtActiveCell.Text



    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(True, "txtActiveCell_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub


    '	Private Sub txtActiveCell_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtActiveCell.KeyDown
    '		Dim KeyCode As Short = eventArgs.KeyCode
    '		Dim Shift As Short = eventArgs.KeyData \ &H10000
    '		On Error GoTo ErrorHandler


    '		With grdPollDef
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
    '					If .row > 1 Then
    '						.row = .row - 1
    '						KeyMoveUpdate()
    '					Else
    '						If .col < 6 Then
    '							.col = .col + 1
    '							.row = .Rows - 1
    '						Else
    '							.row = .row
    '						End If
    '						KeyMoveUpdate()
    '					End If

    '				Case System.Windows.Forms.Keys.Down 'if the user presses the down arrow, get out of the textbox and move down a cell on the grid
    '					.Text = txtActiveCell.Text
    '					txtActiveCell.Visible = False
    '					If .row < grdPollDef.Rows - 1 Then
    '						.row = .row + 1
    '						KeyMoveUpdate()
    '					Else
    '						If .col > 1 Then
    '							.row = 1 'again, if the row is on the last row, don't move cells
    '							.col = .col - 1
    '						Else
    '							.Text = txtActiveCell.Text
    '							txtActiveCell.Visible = False
    '						End If
    '						KeyMoveUpdate()
    '					End If
    '				Case System.Windows.Forms.Keys.Left 'if the user pressed the right arrow key, get out of the textbox and move right one cell in the grid
    '					txtActiveCell.Visible = False
    '					If .col > 3 Then
    '						.col = .col - 1
    '						KeyMoveUpdate()

    '					Else

    '						If .row > 1 And .col = 3 Then
    '							.col = 6
    '							.row = .row - 1
    '							KeyMoveUpdate()
    '						End If
    '					End If
    '				Case System.Windows.Forms.Keys.Right 'if the user pressed the right arrow key, get out of the textbox and move right one cell in the grid
    '					.Text = txtActiveCell.Text
    '					txtActiveCell.Visible = False
    '					If .col < 6 Then
    '						.col = .col + 1
    '						KeyMoveUpdate()
    '					Else
    '						If .row < grdPollDef.Rows - 1 Then
    '							.col = 3
    '							.row = .row + 1
    '							KeyMoveUpdate()
    '						End If
    '					End If
    '			End Select
    '		End With



    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(True, "txtActiveCell_KeyDown " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub

    '	Private Sub KeyMoveUpdate()
    '		On Error GoTo ErrorHandler

    '		'This guy basically replicates the functionality of the grdpolldef_Click event and is used
    '		'in a couple of instances for moving around the grid.

    '		m_intPollRow = grdPollDef.row
    '		m_intPollCol = grdPollDef.col

    '		txtActiveCell.Visible = False

    '		If (m_intPollCol >= 3 And m_intPollCol < 7) And (m_intPollRow >= 1) Then

    '			With txtActiveCell

    '				.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(SSTab1.Left) + VB6.PixelsToTwipsX(grdPollDef.Left) + grdPollDef.CellLeft), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(SSTab1.Top) + VB6.PixelsToTwipsY(grdPollDef.Top) + grdPollDef.CellTop), VB6.TwipsToPixelsX(grdPollDef.CellWidth - 30), VB6.TwipsToPixelsY(grdPollDef.CellHeight - 150))

    '				If m_intPollRow <> 0 Then

    '					.Visible = True
    '					m_strUndoText = grdPollDef.get_TextMatrix(m_intPollRow, m_intPollCol)
    '					.Text = m_strUndoText
    '				End If
    '			End With

    '			txtActiveCell.Focus()

    '		ElseIf m_intPollCol = 0 Then 

    '			txtActiveCell.Visible = False

    '		End If

    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(False, "KeyMoveUpdate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub

    '	Private Sub txtActiveCell_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtActiveCell.Validating
    '		Dim Cancel As Boolean = eventArgs.Cancel
    '		On Error GoTo ErrorHandler


    '		If Not boolChanged Then

    '			If txtActiveCell.Text <> m_strUndoText Then
    '				boolChanged = True
    '				CmdSaveEnabled()
    '			End If

    '		End If


    '		GoTo EventExitSub
    'ErrorHandler: 
    '		HandleError(True, "txtActiveCell_Validate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    'EventExitSub: 
    '		eventArgs.Cancel = Cancel
    '	End Sub

    '	Private Sub txtActiveCell_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtActiveCell.Enter
    '		On Error GoTo ErrorHandler


    '		txtActiveCell.SelectionStart = 0
    '		txtActiveCell.SelectionLength = Len(txtActiveCell.Text)



    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(True, "txtActiveCell_GotFocus " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub


    '	'UPGRADE_WARNING: Event txtActiveCellWQStd.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    '	Private Sub txtActiveCellWQStd_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtActiveCellWQStd.TextChanged
    '		On Error GoTo ErrorHandler


    '		grdWQStd.Text = txtActiveCellWQStd.Text



    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(True, "txtActiveCellWQStd_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub

    '	'Just making the entire cell text selected
    '	Private Sub txtActiveCellWQStd_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtActiveCellWQStd.Enter
    '		On Error GoTo ErrorHandler

    '		txtActiveCellWQStd.SelectionStart = 0
    '		txtActiveCellWQStd.SelectionLength = Len(txtActiveCellWQStd.Text)
    '		m_strUndoTextWQ = txtActiveCellWQStd.Text


    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(True, "txtActiveCellWQStd_GotFocus " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub
    '	Private Sub txtActiveCellWQStd_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtActiveCellWQStd.Validating
    '		Dim Cancel As Boolean = eventArgs.Cancel

    '		On Error GoTo ErrorHandler


    '		If Not boolChanged Then

    '			If txtActiveCellWQStd.Text <> m_strUndoTextWQ Then
    '				boolChanged = True
    '				CmdSaveEnabled()
    '			End If

    '		End If


    '		GoTo EventExitSub
    'ErrorHandler: 
    '		HandleError(True, "txtActiveCell_Validate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)

    'EventExitSub: 
    '		eventArgs.Cancel = Cancel
    '	End Sub

    '	Private Function ValidateGridValues() As Boolean
    '		On Error GoTo ErrorHandler

    '		'Need to validate each grid value before saving.  Essentially we take it a row at a time,
    '		'then rifle through each column of each row.  Case Select tests each each x,y value depending
    '		'on column... 3-6 must be 1-100 range

    '		'Returns: True or False

    '		Dim varActive As Object 'txtActiveCell value
    '		Dim i As Short
    '		Dim j As Short
    '		Dim iQstd As Short
    '		Dim jQstd As Short

    '		For i = 1 To grdPollDef.Rows - 1

    '			For j = 3 To 6

    '				'UPGRADE_WARNING: Couldn't resolve default property of object varActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '				varActive = grdPollDef.get_TextMatrix(i, j)

    '				'UPGRADE_WARNING: Couldn't resolve default property of object varActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '				If InStr(1, CStr(varActive), ".", CompareMethod.Text) > 0 Then
    '					'UPGRADE_WARNING: Couldn't resolve default property of object varActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '					If (Len(Split(CStr(varActive), ".")(1)) > 4) Then
    '						ErrorGenerator(Err6, i, j)
    '						grdPollDef.col = j
    '						grdPollDef.row = i
    '						ValidateGridValues = False
    '						KeyMoveUpdate()
    '						Exit Function
    '					End If
    '				End If

    '				'UPGRADE_WARNING: Couldn't resolve default property of object varActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '				If Not IsNumeric(varActive) Or (varActive < 0) Or (varActive > 1000) Then
    '					ErrorGenerator(Err6, i, j)
    '					grdPollDef.col = j
    '					grdPollDef.row = i
    '					ValidateGridValues = False
    '					KeyMoveUpdate()
    '					Exit Function
    '				End If



    '			Next j

    '		Next i

    '		For iQstd = 1 To grdWQStd.Rows - 1
    '			'UPGRADE_WARNING: Couldn't resolve default property of object varActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '			varActive = grdWQStd.get_TextMatrix(iQstd, 3)

    '			'UPGRADE_WARNING: Couldn't resolve default property of object varActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '			If Not IsNumeric(varActive) Or (varActive < 0) Then
    '				ErrorGenerator(Err5, iQstd, 3)
    '				grdWQStd.col = 3
    '				grdWQStd.row = iQstd
    '				ValidateGridValues = False
    '				Exit Function
    '			End If
    '		Next iQstd

    '		ValidateGridValues = True

    '		Exit Function
    'ErrorHandler: 
    '		HandleError(False, "ValidateGridValues " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Function


    '	Private Sub txtCoeffSetDesc_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCoeffSetDesc.Enter
    '		On Error GoTo ErrorHandler

    '		m_strUndoDesc = txtCoeffSetDesc.Text

    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(True, "txtCoeffSetDesc_GotFocus " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub

    '	Private Sub txtCoeffSetDesc_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtCoeffSetDesc.Validating
    '		Dim Cancel As Boolean = eventArgs.Cancel
    '		On Error GoTo ErrorHandler


    '		'If Not boolDescChanged Then
    '		If m_strUndoDesc <> txtCoeffSetDesc.Text Then
    '			boolDescChanged = True
    '			CmdSaveEnabled()
    '		End If
    '		'End If

    '		GoTo EventExitSub
    'ErrorHandler: 
    '		HandleError(True, "txtCoeffSetDesc_Validate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    'EventExitSub: 
    '		eventArgs.Cancel = Cancel
    '	End Sub


    '	Private Sub DeletePollutant(ByRef strName As String)

    '		'Can ya guess what this one does?

    '		On Error GoTo ErrorHandler

    '		Dim strPollDelete As String
    '		Dim strPoll2Delete As String

    '		Dim rsPollDelete As ADODB.Recordset
    '		'Dim rsLCClassDelete As ADODB.Recordset

    '		strPollDelete = "Delete * FROM Pollutant WHERE NAME LIKE '" & strName & "'"


    '		modUtil.g_ADOConn.Execute(strPollDelete)

    '		MsgBox(strName & " deleted.", MsgBoxStyle.OKOnly, "Record Deleted")

    '		Me.cboPollName.Items.Clear()
    '		modUtil.InitComboBox((Me.cboPollName), "Pollutant")
    '		Me.Refresh()

    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(False, "DeletePollutant " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub

    '	Public Sub AddCoefficient(ByRef strCoeffName As String, ByRef strLCType As String)
    '		On Error GoTo ErrorHandler

    '		'General gist:  First we add new record to the Coefficient Set table using strCoeffName as
    '		'the name, m_intPollID as the PollID, and m_intLCTYPEID as the LCTypeID.  The last two are
    '		'garnered above during a cbo click event.  Once that's done, we'll add a series of blank
    '		'coefficients for the landclass type the user chooses...ie CCAP, NotCCAP, whatever

    '		Dim strNewLcType As String 'CmdString for inserting new coefficientset
    '		Dim strGetLcType As String
    '		Dim strDefault As String '
    '		Dim strNewCoeff As String
    '		Dim strNewCoeffID As String 'Holder for the CoefficientSetID
    '		Dim strInsertNewCoeff As String 'Putting the newly created coefficients in their table
    '		Dim intCoeffSetID As Short
    '		Dim i As Short

    '		Dim rsLCType As New ADODB.Recordset
    '		Dim rsCopySet As ADODB.Recordset 'First RS
    '		Dim rsCoeffSetID As New ADODB.Recordset '1 record recordset to get a hold of CoeffsetID
    '		Dim rsNewCoeff As New ADODB.Recordset

    '		strGetLcType = "SELECT * FROM LCTYPE WHERE NAME LIKE '" & strLCType & "'"
    '		rsLCType.Open(strGetLcType, g_ADOConn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

    '		strNewLcType = "INSERT INTO COEFFICIENTSET(NAME, POLLID, LCTYPEID) VALUES ('" & Replace(strCoeffName, "'", "''") & "'," & Replace(CStr(m_intPollID), "'", "''") & "," & Replace(rsLCType.Fields("LCTypeID").Value, "'", "''") & ")"

    '		'First need to add the coefficient set to that table
    '		g_ADOConn.Execute(strNewLcType, ADODB.AffectEnum.adAffectCurrent)

    '		'Get the Coefficient Set ID of the newly created coefficient set to populate Column # 8 in the GRid,
    '		'which by the way, is hidden from view.  InitPollDef sets the widths of col 7, 8 to 0
    '		strNewCoeffID = "SELECT COEFFSETID FROM COEFFICIENTSET " & "WHERE COEFFICIENTSET.NAME LIKE '" & strCoeffName & "'"

    '		rsCoeffSetID.Open(strNewCoeffID, g_ADOConn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
    '		intCoeffSetID = rsCoeffSetID.Fields("CoeffSetID").Value


    '		strDefault = "SELECT LCTYPE.LCTYPEID, LCCLASS.LCCLASSID, LCCLASS.NAME As valName, " & "LCCLASS.VAlue as valValue FROM LCTYPE " & "INNER JOIN LCCLASS ON LCCLASS.LCTYPEID = LCTYPE.LCTYPEID " & "WHERE LCTYPE.Name Like " & "'" & strLCType & "' ORDER BY LCCLASS.Value"

    '		Debug.Print(strDefault)
    '		rsCopySet = New ADODB.Recordset
    '		rsCopySet.Open(strDefault, g_ADOConn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)

    '		'Clear things and set the rows to recordcount + 1, remember 1st row fixed
    '		grdPollDef.Clear()
    '		grdPollDef.Rows = rsCopySet.RecordCount + 1


    '		modUtil.InitPollDefGrid(grdPollDef) 'Call this again to set it up as we cleared it

    '		Dim cmdInsert As New ADODB.Command

    '		'Now loopy loo to populate values.
    '		Dim strNewCoeff1 As String
    '		strNewCoeff1 = "SELECT * FROM COEFFICIENT"

    '		rsNewCoeff.Open(strNewCoeff1, g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

    '		i = 0

    '		rsCopySet.MoveFirst()
    '		'For i = 1 To rsCopySet.RecordCount
    '		Do While Not rsCopySet.EOF
    '			i = i + 1
    '			Debug.Print("Record: " & i & ": " & rsCopySet.Fields("valName").Value)
    '			'Let's try one more ADO method, why not, righ?
    '			rsNewCoeff.AddNew()


    '			'Add the necessary components
    '			rsNewCoeff.Fields("Coeff1").Value = 0
    '			rsNewCoeff.Fields("Coeff2").Value = 0
    '			rsNewCoeff.Fields("Coeff3").Value = 0
    '			rsNewCoeff.Fields("Coeff4").Value = 0
    '			rsNewCoeff.Fields("CoeffSetID").Value = rsCoeffSetID.Fields("CoeffSetID").Value
    '			rsNewCoeff.Fields("LCClassID").Value = rsCopySet.Fields("LCClassID").Value
    '			rsNewCoeff.Update()

    '			rsCopySet.MoveNext()
    '		Loop 

    '		'Cleanup
    '		rsCopySet.Close()
    '		rsCoeffSetID.Close()
    '		rsNewCoeff.Close()

    '		cboPollName.SelectedIndex = cboPollName.SelectedIndex

    '		cboCoeffSet.Items.Add(strCoeffName)

    '		'Call the function to set everything to newly added Coefficient.
    '		cboCoeffSet.SelectedIndex = GetCboIndex(strCoeffName, cboCoeffSet)

    '		txtLCType.Text = rsLCType.Fields("Name").Value

    '		'UPGRADE_NOTE: Object rsCopySet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '		rsCopySet = Nothing
    '		'UPGRADE_NOTE: Object rsCoeffSetID may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '		rsCoeffSetID = Nothing
    '		'UPGRADE_NOTE: Object rsNewCoeff may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '		rsNewCoeff = Nothing

    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(True, "AddCoefficient " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub

    '	Private Sub UpdateValues()
    '		On Error GoTo ErrorHandler


    '		Dim i As Short
    '		Dim rsPollUpdate As New ADODB.Recordset
    '		Dim strPollUpdate As String
    '		Dim strWQSelect As String

    '		Dim rsDescrip As New ADODB.Recordset
    '		Dim rsWQstd As New ADODB.Recordset

    '		If ValidateGridValues Then
    '			'Update

    '			For i = 1 To grdPollDef.Rows - 1

    '				strPollUpdate = "SELECT * From Coefficient Where CoeffID = " & grdPollDef.get_TextMatrix(i, 7)
    '				rsPollUpdate.Open(strPollUpdate, g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

    '				rsPollUpdate.Fields("Coeff1").Value = grdPollDef.get_TextMatrix(i, 3)
    '				rsPollUpdate.Fields("Coeff2").Value = grdPollDef.get_TextMatrix(i, 4)
    '				rsPollUpdate.Fields("Coeff3").Value = grdPollDef.get_TextMatrix(i, 5)
    '				rsPollUpdate.Fields("Coeff4").Value = grdPollDef.get_TextMatrix(i, 6)

    '				rsPollUpdate.Update()
    '				rsPollUpdate.Close()

    '			Next i

    '		End If

    '		Dim strUpdateDescription As Object
    '		If boolDescChanged Then

    '			'UPGRADE_WARNING: Couldn't resolve default property of object strUpdateDescription. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '			strUpdateDescription = "SELECT Description from CoefficientSet Where Name like '" & cboCoeffSet.Text & "'"

    '			rsDescrip.Open(strUpdateDescription, g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

    '			If Len(txtCoeffSetDesc.Text) = 0 Then
    '				rsDescrip.Fields("Description").Value = ""
    '			Else
    '				rsDescrip.Fields("Description").Value = txtCoeffSetDesc.Text
    '			End If

    '			rsDescrip.Update()
    '			rsDescrip.Close()

    '		End If

    '		For i = 1 To grdWQStd.Rows - 1
    '			strWQSelect = "SELECT * from POLL_WQCRITERIA WHERE POLL_WQCRITID = " & grdWQStd.get_TextMatrix(i, 4)

    '			rsWQstd.Open(strWQSelect, g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

    '			rsWQstd.Fields("Threshold").Value = grdWQStd.get_TextMatrix(i, 3)
    '			rsWQstd.Update()
    '			rsWQstd.Close()

    '		Next i

    '		'UPGRADE_NOTE: Object rsDescrip may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '		rsDescrip = Nothing
    '		'UPGRADE_NOTE: Object rsPollUpdate may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '		rsPollUpdate = Nothing
    '		'UPGRADE_NOTE: Object rsWQstd may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '		rsWQstd = Nothing


    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(False, "UpdateValues " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub

    '	Public Sub UpdateCoeffSet(ByRef adoRSCoeff As ADODB.Recordset, ByRef strCoeffName As String, ByRef strFileName As String)
    '		On Error GoTo ErrorHandler


    '		'General gist:  First we add new record to the Coefficient Set table using strCoeffName as
    '		'the name, m_intPollID as the PollID, and m_intLCTYPEID as the LCTypeID.  The last two are
    '		'garnered above during a cbo click event.  Once that's done, we'll add a series of
    '		'coefficients for the landclass based on the incoming textfile...strFileName


    '		Dim rsCopySet As ADODB.Recordset 'First RS
    '		Dim rsCoeffSetID As New ADODB.Recordset '1 record recordset to get a hold of CoeffsetID
    '		Dim rsNewCoeff As New ADODB.Recordset 'New Coefficient set

    '		Dim strNewLcType As String 'CmdString for inserting new coefficientset
    '		Dim strDefault As String '
    '		Dim strNewCoeff As String
    '		Dim strNewCoeffID As String 'Holder for the CoefficientSetID
    '		Dim strInsertNewCoeff As String 'Putting the newly created coefficients in their table
    '		Dim intCoeffSetID As Short
    '		Dim i As Short

    '		'Textfile related material
    '		Dim fso As New Scripting.FileSystemObject
    '		Dim fl As Scripting.TextStream
    '		Dim strLine As String
    '		Dim strValue As Short
    '		Dim intLine As Short

    '		adoRSCoeff.ReQuery()

    '		strNewLcType = "INSERT INTO COEFFICIENTSET(NAME, POLLID, LCTYPEID) VALUES ('" & Replace(strCoeffName, "'", "''") & "'," & m_intPollID & "," & adoRSCoeff.Fields("LCTypeID").Value & ")"

    '		'First need to add the coefficient set to that table
    '		g_ADOConn.Execute(strNewLcType, ADODB.AffectEnum.adAffectCurrent)

    '		'Get the Coefficient Set ID of the newly created coefficient set to populate Column # 8 in the GRid,
    '		'which by the way, is hidden from view.  InitPollDef sets the widths of col 7, 8 to 0
    '		strNewCoeffID = "SELECT COEFFSETID FROM COEFFICIENTSET " & "WHERE COEFFICIENTSET.NAME LIKE '" & strCoeffName & "'"

    '		rsCoeffSetID.Open(strNewCoeffID, g_ADOConn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
    '		intCoeffSetID = rsCoeffSetID.Fields("CoeffSetID").Value

    '		'Now turn attention to the TextFile...to get the users coefficient values
    '		fl = fso.OpenTextFile(strFileName, Scripting.IOMode.ForReading, True, Scripting.Tristate.TristateFalse)
    '		intLine = 0

    '		'Now loopy loo to populate values.
    '		Dim strNewCoeff1 As String
    '		strNewCoeff1 = "SELECT * FROM COEFFICIENT"
    '		rsNewCoeff.Open(strNewCoeff1, g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

    '		'adoRSCoeff.ReQuery
    '		Debug.Print(adoRSCoeff.RecordCount)

    '		adoRSCoeff.ReQuery()
    '		adoRSCoeff.MoveFirst()

    '		i = 0

    '		grdPollDef.Clear()
    '		grdPollDef.Rows = adoRSCoeff.RecordCount + 1

    '		Do While Not fl.AtEndOfStream

    '			strLine = fl.ReadLine
    '			'Value exits??
    '			strValue = CShort(Split(strLine, ",")(0))
    '			adoRSCoeff.MoveFirst()


    '			For i = 0 To adoRSCoeff.RecordCount - 1
    '				Debug.Print("")

    '				If adoRSCoeff.Fields("Value").Value = strValue Then

    '					'Let's try one more ADO method
    '					rsNewCoeff.AddNew()

    '					'Add the necessary components
    '					rsNewCoeff.Fields("Coeff1").Value = Split(strLine, ",")(1)
    '					rsNewCoeff.Fields("Coeff2").Value = Split(strLine, ",")(2)
    '					rsNewCoeff.Fields("Coeff3").Value = Split(strLine, ",")(3)
    '					rsNewCoeff.Fields("Coeff4").Value = Split(strLine, ",")(4)
    '					rsNewCoeff.Fields("CoeffSetID").Value = rsCoeffSetID.Fields("CoeffSetID").Value
    '					rsNewCoeff.Fields("LCClassID").Value = adoRSCoeff.Fields("LCClassID").Value
    '					rsNewCoeff.Update()

    '					With grdPollDef
    '						.set_TextMatrix(i, 1, strValue)
    '						.set_TextMatrix(i, 2, adoRSCoeff.Fields("Name").Value)
    '						.set_TextMatrix(i, 3, Split(strLine, ",")(2))
    '						.set_TextMatrix(i, 4, Split(strLine, ",")(2))
    '						.set_TextMatrix(i, 5, Split(strLine, ",")(2))
    '						.set_TextMatrix(i, 6, Split(strLine, ",")(2))
    '						.set_TextMatrix(i, 7, rsCoeffSetID.Fields("CoeffSetID").Value)
    '						.set_TextMatrix(i, 8, rsNewCoeff.Fields("coeffID").Value)
    '					End With
    '				End If
    '				adoRSCoeff.MoveNext()
    '			Next i

    '		Loop 

    '		frmImportCoeffSet.Close()
    '		cboCoeffSet.Items.Clear()
    '		InitComboBox(cboCoeffSet, "CoefficientSet")
    '		cboCoeffSet.SelectedIndex = modUtil.GetCboIndex(strCoeffName, cboCoeffSet)




    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(True, "UpdateCoeffSet " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub

    '	Public Sub CopyCoefficient(ByRef strNewCoeffName As String, ByRef strCoeffSet As String)
    '		On Error GoTo ErrorHandler


    '		'General gist:  First we add new record to the Coefficient Set table using strNewCoeffName as
    '		'the name, PollID, LCTYPEID.  Once that's done, we'll add the coefficients
    '		'from the set being copied
    '		Dim strCopySet As String 'The Recordset of existing coefficients being copied
    '		Dim strNewLcType As String 'CmdString for inserting new coefficientset               '
    '		Dim strNewCoeff As String
    '		Dim strNewCoeffID As String 'Holder for the CoefficientSetID
    '		Dim intCoeffSetID As Short
    '		Dim i As Short

    '		Dim rsCopySet As New ADODB.Recordset 'First RS
    '		Dim rsCoeffSetID As New ADODB.Recordset '1 record recordset to get a hold of CoeffsetID
    '		Dim rsNewCoeff As New ADODB.Recordset

    '		strCopySet = "SELECT * FROM COEFFICIENTSET INNER JOIN COEFFICIENT ON COEFFICIENTSET.COEFFSETID = " & "COEFFICIENT.COEFFSETID WHERE COEFFICIENTSET.NAME LIKE '" & strCoeffSet & "'"

    '		rsCopySet.Open(strCopySet, g_ADOConn, ADODB.CursorTypeEnum.adOpenKeyset)


    '		'INSERT: new Coefficient set taking the PollID and LCType ID from rsCopySet
    '		strNewLcType = "INSERT INTO COEFFICIENTSET(NAME, POLLID, LCTYPEID) VALUES ('" & Replace(strNewCoeffName, "'", "''") & "'," & rsCopySet.Fields("POLLID").Value & "," & rsCopySet.Fields("LCTypeID").Value & ")"

    '		'First need to add the coefficient set to that table
    '		g_ADOConn.Execute(strNewLcType, ADODB.AffectEnum.adAffectCurrent)

    '		'Get the Coefficient Set ID of the newly created coefficient set to populate Column # 8 in the GRid,
    '		'which by the way, is hidden from view.  InitPollDef sets the widths of col 7, 8 to 0
    '		strNewCoeffID = "SELECT COEFFSETID FROM COEFFICIENTSET " & "WHERE COEFFICIENTSET.NAME LIKE '" & strNewCoeffName & "'"

    '		rsCoeffSetID.Open(strNewCoeffID, g_ADOConn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockPessimistic)
    '		intCoeffSetID = rsCoeffSetID.Fields("CoeffSetID").Value

    '		'Now loopy loo to populate values.
    '		Dim strNewCoeff1 As String
    '		strNewCoeff1 = "SELECT * FROM COEFFICIENT"
    '		rsNewCoeff.Open(strNewCoeff1, g_ADOConn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)

    '		'Clear things and set the rows to recordcount + 1, remember 1st row fixed
    '		grdPollDef.Clear()
    '		grdPollDef.Rows = rsCopySet.RecordCount + 1

    '		rsCopySet.MoveFirst()

    '		'Actually add the records to the new set
    '		For i = 1 To rsCopySet.RecordCount

    '			'Let's try one more ADO method, why not, righ?
    '			rsNewCoeff.AddNew()

    '			'Add the necessary components
    '			rsNewCoeff.Fields("Coeff1").Value = rsCopySet.Fields("Coeff1").Value
    '			rsNewCoeff.Fields("Coeff2").Value = rsCopySet.Fields("Coeff2").Value
    '			rsNewCoeff.Fields("Coeff3").Value = rsCopySet.Fields("Coeff3").Value
    '			rsNewCoeff.Fields("Coeff4").Value = rsCopySet.Fields("Coeff4").Value
    '			rsNewCoeff.Fields("CoeffSetID").Value = intCoeffSetID
    '			rsNewCoeff.Fields("LCClassID").Value = rsCopySet.Fields("LCClassID").Value

    '			rsNewCoeff.Update()

    '			rsCopySet.MoveNext()

    '		Next i


    '		'Set up everything to look good
    '		cboPollName_SelectedIndexChanged(cboPollName, New System.EventArgs())
    '		cboCoeffSet.SelectedIndex = GetCboIndex(strNewCoeffName, cboCoeffSet)
    '		frmImportCoeffSet.Close()


    '		'Cleanup
    '		rsCopySet.Close()
    '		rsCoeffSetID.Close()
    '		rsNewCoeff.Close()

    '		'UPGRADE_NOTE: Object rsCopySet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '		rsCopySet = Nothing
    '		'UPGRADE_NOTE: Object rsCoeffSetID may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '		rsCoeffSetID = Nothing
    '		'UPGRADE_NOTE: Object rsNewCoeff may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '		rsNewCoeff = Nothing




    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(True, "CopyCoefficient " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub

    '	Private Sub CmdSaveEnabled()
    '		On Error GoTo ErrorHandler

    '		If boolChanged Or boolDescChanged Then
    '			cmdSave.Enabled = True
    '		Else
    '			cmdSave.Enabled = False
    '		End If


    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(False, "CmdSaveEnabled " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub

    '	'Form hooker
    '	Public Sub init(ByVal pApp As ESRI.ArcGIS.Framework.IApplication)
    '		On Error GoTo ErrorHandler

    '		m_App = pApp

    '		Exit Sub
    'ErrorHandler: 
    '		HandleError(True, "init " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    '	End Sub
End Class