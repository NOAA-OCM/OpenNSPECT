Option Strict Off
Option Explicit On
Friend Class frmImportWQStd
	Inherits System.Windows.Forms.Form
	' *************************************************************************************
	' *  Perot Systems Government Services
	' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
	' *  frmImportWQStd
	' *************************************************************************************
	' *  Description:  Allows for the copying of the contents of an existing wat qual standard
	' *
	' *
	' *  Called By:  frmWQStd
	' *************************************************************************************
	
	Private m_strFileName As String
	' Variables used by the Error handler function - DO NOT REMOVE
	Const c_sModuleFileName As String = "frmImportWQStd.frm"
	
	
	Private Sub cmdBrowse_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdBrowse.Click
		On Error GoTo ErrorHandler
		
		
		'browse...get output filename
		dlgCMD1Open.FileName = CStr(Nothing)
		'UPGRADE_WARNING: CommonDialog variable was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="671167DC-EA81-475D-B690-7A40C7BF4A23"'
		With dlgCMD1
			'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			.Filter = MSG1
			.Title = MSG2
			.FilterIndex = 1
			'UPGRADE_WARNING: FileOpenConstants constant FileOpenConstants.cdlOFNHideReadOnly was upgraded to OpenFileDialog.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
			'UPGRADE_WARNING: MSComDlg.CommonDialog property dlgCMD1.Flags was upgraded to dlgCMD1Open.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
			'UPGRADE_WARNING: FileOpenConstants constant FileOpenConstants.cdlOFNHideReadOnly was upgraded to OpenFileDialog.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
			.ShowReadOnly = False
			'UPGRADE_WARNING: MSComDlg.CommonDialog property dlgCMD1.Flags was upgraded to dlgCMD1Open.CheckFileExists which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
			.CheckFileExists = True
			.CheckPathExists = True
			.ShowDialog()
		End With
		If Len(dlgCMD1Open.FileName) > 0 Then
			txtImpFile.Text = Trim(dlgCMD1Open.FileName)
			m_strFileName = txtImpFile.Text
			cmdOK.Enabled = True
		End If
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "cmdBrowse_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		On Error GoTo ErrorHandler
		
		Me.Close()
		
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "cmdCancel_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		On Error GoTo ErrorHandler
		
		
		Dim fso As New Scripting.FileSystemObject
		Dim fl As Scripting.TextStream
		Dim strLine As String
		Dim intLine As Short
		Dim strName As String
		Dim strDescript As String
		Dim strPoll As String
		Dim strThresh As String
		Dim strCmd As String
		Dim rsNew As ADODB.Recordset
		
		fl = fso.OpenTextFile(m_strFileName, Scripting.IOMode.ForReading, True, Scripting.Tristate.TristateFalse)
		
		intLine = 0
		
		Do While Not fl.AtEndOfStream
			strLine = fl.ReadLine
			intLine = intLine + 1
			'MsgBox theLine
			
			If intLine = 1 Then
				
				strName = Trim(txtStdName.Text)
				strDescript = Split(strLine, ",")(1)
				
				If strName = "" Then
					
					MsgBox("Name is blank.  Please enter a name.", MsgBoxStyle.Critical, "Empty Name Field")
					txtStdName.Focus()
					Exit Sub
					
				Else
					
					strCmd = "INSERT INTO WQCRITERIA (NAME,DESCRIPTION) VALUES ('" & Replace(txtStdName.Text, "'", "''") & "', '" & Replace(strDescript, "'", "''") & "')"
					'Name Check
					If modUtil.UniqueName("WQCRITERIA", (txtStdName.Text)) Then
						g_ADOConn.Execute(strCmd, ADODB.CommandTypeEnum.adCmdText)
					Else
						MsgBox("The name you have chosen is already in use.  Please select another.", MsgBoxStyle.Critical, "Select Unique Name")
						Exit Sub
					End If
					
				End If
				
			Else
				
				strPoll = Split(strLine, ",")(0)
				strThresh = Split(strLine, ",")(1)
				'Insert the pollutant/threshold value into POLL_WQCRITERIA
				PollutantAdd(strName, strPoll, strThresh)
				
			End If
			
		Loop 
		
		fl.Close()
		
		'Cleanup
		frmWQStd.cboWQStdName.Items.Clear()
		modUtil.InitComboBox((frmWQStd.cboWQStdName), "WQCRITERIA")
		frmWQStd.cboWQStdName.SelectedIndex = modUtil.GetCboIndex((txtStdName.Text), (frmWQStd.cboWQStdName))
		
		'UPGRADE_NOTE: Object fso may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		fso = Nothing
		'UPGRADE_NOTE: Object fl may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		fl = Nothing
		'UPGRADE_NOTE: Object rsNew may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsNew = Nothing
		
		Me.Close()
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "cmdOK_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	Private Sub PollutantAdd(ByRef strName As String, ByRef strPoll As String, ByRef strThresh As String)
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
		
		strCmdInsert = "INSERT INTO POLL_WQCRITERIA (PollID,WQCritID,Threshold) VALUES ('" & rsPollDetails.Fields("POLLID").Value & "', '" & rsPollAdd.Fields("WQCRITID").Value & "'," & strThresh & ")"
		
		g_ADOConn.Execute(strCmdInsert, ADODB.CommandTypeEnum.adCmdText)
		
		'Cleanup
		rsPollAdd.Close()
		rsPollDetails.Close()
		
		'UPGRADE_NOTE: Object rsPollAdd may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsPollAdd = Nothing
		'UPGRADE_NOTE: Object rsPollDetails may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsPollDetails = Nothing
		
		Exit Sub
ErrorHandler: 
		HandleError(False, "PollutantAdd " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
End Class