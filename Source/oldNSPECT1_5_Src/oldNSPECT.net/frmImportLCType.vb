Option Strict Off
Option Explicit On
Friend Class frmImportLCType
	Inherits System.Windows.Forms.Form
	' *************************************************************************************
	' *  Perot Systems Government Services
	' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
	' *  frmNewPrecip
	' *************************************************************************************
	' *  Description:  Allows for the importing of an existing land cover type classification
	' *
	' *
	' *  Called By:  frmImportLCType
	' *************************************************************************************
	
	
	Private m_strFileName As String
	Private booName As Boolean 'Check if user put a name in
	Private booFile As Boolean 'Check if FileName is correct
	
	
	Private Sub cmdBrowse_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdBrowse.Click
		'browse...get output filename
		dlgCMD1Open.FileName = CStr(Nothing)
		'UPGRADE_WARNING: CommonDialog variable was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="671167DC-EA81-475D-B690-7A40C7BF4A23"'
		With dlgCMD1
			'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			.Filter = Replace(MSG1, "<name>", "Land Cover Type")
			.Title = Replace(MSG2, "<name>", "Land Cover Type")
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
			m_strFileName = Trim(dlgCMD1Open.FileName)
			booFile = True
		End If
		
		If booFile And booName Then
			cmdOK.Enabled = True
		End If
		
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		
		Dim fso As New Scripting.FileSystemObject
		Dim fl As Scripting.TextStream
		Dim strLine As String
		Dim intLine As Short
		Dim strName As String
		Dim strDescript As String
		Dim strParams(7) As Object
		Dim strCmd As String
		Dim rsNew As ADODB.Recordset
		Dim i As Short
		
		If fso.FileExists(m_strFileName) Then
			
			fl = fso.OpenTextFile(m_strFileName, Scripting.IOMode.ForReading, True, Scripting.Tristate.TristateFalse)
			
			intLine = 0
			
			Do While Not fl.AtEndOfStream
				strLine = fl.ReadLine
				intLine = intLine + 1
				'MsgBox theLine
				
				'Get the first line, supposed to contain Name, Description
				If intLine = 1 Then
					
					strName = Trim(txtLCType.Text)
					strDescript = Split(strLine, ",")(1)
					
					'Check if name is present, if not bark
					If strName = "" Then
						
						MsgBox("Name is blank.  Please enter a name.", MsgBoxStyle.Critical, "Empty Name Field")
						txtLCType.Focus()
						Exit Sub
						
					Else
						
						'Name Check, if cool perform
						If modUtil.UniqueName("LCTYPE", (txtLCType.Text)) Then
							strCmd = "INSERT INTO LCTYPE (NAME,DESCRIPTION) VALUES ('" & Replace(txtLCType.Text, "'", "''") & "', '" & Replace(strDescript, "'", "''") & "')"
							g_ADOConn.Execute(strCmd, ADODB.CommandTypeEnum.adCmdText)
						Else
							MsgBox("The name you have chosen is already in use.  Please select another.", MsgBoxStyle.Critical, "Select Unique Name")
							Exit Sub
						End If 'End unique name check
						
					End If 'end empty name check
					
				Else ' > line 1
					
					If Len(Trim(strLine)) > 0 Then
						
						i = 0
						
						'Create an array of lines ie Value,Descript,1,2,3,4,CoverFactor,W/WL
						For i = 0 To UBound(strParams)
							'UPGRADE_WARNING: Couldn't resolve default property of object strParams(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strParams(i) = Split(strLine, ",")(i)
						Next i
						
						'Check the values, if ok add them, if not rollback
						If modLCTypeValues.CheckGridValuesLCType(strParams) Then
							AddLCClass(strName, strParams)
						Else
							modLCTypeValues.RollBack((strName))
							GoTo Cleanup
						End If 'End check
					Else
						GoTo Cleanup
					End If
					
				End If
				
			Loop 
			
			fl.Close()
			'Redo the form and make newly added lctype first
			frmLCTypes.cboLCType.Items.Clear()
			modUtil.InitComboBox((frmLCTypes.cboLCType), "LCType")
			frmLCTypes.cboLCType.SelectedIndex = modUtil.GetCboIndex(strName, (frmLCTypes.cboLCType))
			Me.Close()
		Else
			MsgBox("The file you are pointing to does not exist. Please select another.", MsgBoxStyle.Critical, "File Not Found")
			'Cleanup
		End If
		
		Exit Sub
		
Cleanup: 
		frmLCTypes.cboLCType.Items.Clear()
		modUtil.InitComboBox((frmLCTypes.cboLCType), "LCTYPE")
		
		'UPGRADE_NOTE: Object fso may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		fso = Nothing
		'UPGRADE_NOTE: Object fl may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		fl = Nothing
		'UPGRADE_NOTE: Object rsNew may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsNew = Nothing
		
		Me.Close()
		
		
	End Sub
	
	Private Sub AddLCClass(ByRef strName As String, ByRef strParams() As Object)
		
		On Error GoTo ErrHandler
		
		Dim strLCTypeAdd As String
		Dim strCmdInsert As String
		
		Dim rsLCType As ADODB.Recordset
		rsLCType = New ADODB.Recordset
		
		
		'Get the WQCriteria values using the name
		strLCTypeAdd = "SELECT * FROM LCTYPE WHERE NAME = " & "'" & strName & "'"
		rsLCType.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		rsLCType.Open(strLCTypeAdd, g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
		
		'On Error GoTo dbaseerr
		
		strCmdInsert = "INSERT INTO LCCLASS([Value],[Name],[LCTYPEID],[CN-A],[CN-B],[CN-C],[CN-D],[CoverFactor],[W_WL]) VALUES(" & Replace(CStr(strParams(0)), "'", "''") & ",'" & Replace(CStr(strParams(1)), "'", "''") & "'," & Replace(CStr(rsLCType.Fields("LCTypeID").Value), "'", "''") & "," & Replace(CStr(strParams(2)), "'", "''") & "," & Replace(CStr(strParams(3)), "'", "''") & "," & Replace(CStr(strParams(4)), "'", "''") & "," & Replace(CStr(strParams(5)), "'", "''") & "," & Replace(CStr(strParams(6)), "'", "''") & "," & Replace(CStr(strParams(7)), "'", "''") & ")"
		
		g_ADOConn.Execute(strCmdInsert, ADODB.CommandTypeEnum.adCmdText)
		
		'UPGRADE_NOTE: Object rsLCType may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsLCType = Nothing
		
		Exit Sub
		
ErrHandler: 
		
		MsgBox("There was a problem updating the database.  Insure that your values meet the correct " & "value ranges for each field.", MsgBoxStyle.Critical, "Invalid Values Found")
		
		
		
	End Sub
	
	
	
	
	'UPGRADE_WARNING: Event txtLCType.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtLCType_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLCType.TextChanged
		
		If CBool(Trim(CStr(Len(txtLCType.Text)))) Then
			booName = True
		End If
		
		If booFile And booName Then
			cmdOK.Enabled = True
		End If
		
	End Sub
End Class