Option Strict Off
Option Explicit On
Friend Class frmCopyWQStd
	Inherits System.Windows.Forms.Form
	' *************************************************************************************
	' *  Perot Systems Government Services
	' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
	' *  frmCopyWQStd
	' *************************************************************************************
	' *  Description:  Allows for the copying of the contents of an existing wat qual standard
	' *
	' *
	' *  Called By:  frmWQStd
	' *************************************************************************************
	
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		
		Dim strStandard As String
		Dim rsStandard As ADODB.Recordset
		
		Dim strPollStandard As String
		Dim rsPollStandard As ADODB.Recordset
		
		Dim rsNewStandard As ADODB.Recordset
		Dim rsNewPollCriteria As ADODB.Recordset
		
		Dim strCmd As String
		Dim strCmd2 As String
		
		'Get the WQ stand info
		strStandard = "SELECT * FROM WQCriteria WHERE NAME LIKE '" & cboStdName.Text & "'"
		
		rsStandard = New ADODB.Recordset
		rsStandard.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		rsStandard.Open(strStandard, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
		
		'Get the related pollutant/thresholds
		strPollStandard = "SELECT * FROM POLL_WQCRITERIA WHERE WQCRITID =" & rsStandard.Fields("WQCRITID").Value
		
		rsPollStandard = New ADODB.Recordset
		rsPollStandard.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		rsPollStandard.Open(strPollStandard, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
		
		strCmd = "INSERT INTO WQCRITERIA (NAME,DESCRIPTION) VALUES ('" & Replace(Trim(txtStdName.Text), "'", "''") & "', '" & rsStandard.Fields("Description").Value & "')"
		
		If modUtil.UniqueName("WQCRITERIA", Trim(txtStdName.Text)) Then
			rsNewStandard = modUtil.g_ADOConn.Execute(strCmd)
		Else
			MsgBox(Err4, MsgBoxStyle.Critical, "Enter Unique Name")
			Exit Sub
		End If
		
		rsNewStandard.Open("Select * from WQCRITERIA WHERE NAME LIKE '" & Trim(txtStdName.Text) & "'", modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic)
		
		Dim i As Short
		i = 1
		
		rsPollStandard.MoveFirst()
		
		For i = 1 To rsPollStandard.RecordCount
			strCmd2 = "INSERT INTO POLL_WQCRITERIA (POLLID, WQCRITID, THRESHOLD) VALUES (" & rsPollStandard.Fields("POLLID").Value & ", " & rsNewStandard.Fields("WQCRITID").Value & "," & rsPollStandard.Fields("Threshold").Value & ")"
			rsNewPollCriteria = modUtil.g_ADOConn.Execute(strCmd2)
			rsPollStandard.MoveNext()
		Next i
		
		'Cleanup
		'UPGRADE_NOTE: Object rsStandard may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsStandard = Nothing
		'UPGRADE_NOTE: Object rsPollStandard may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsPollStandard = Nothing
		'UPGRADE_NOTE: Object rsNewStandard may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsNewStandard = Nothing
		'UPGRADE_NOTE: Object rsNewPollCriteria may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsNewPollCriteria = Nothing
		
		frmWQStd.cboWQStdName.Items.Clear()
		modUtil.InitComboBox((frmWQStd.cboWQStdName), "WQCRITERIA")
		
		Me.Close()
		
	End Sub
	
	Private Sub PollutantAdd(ByRef strName As String, ByRef strPoll As String, ByRef strThresh As String)
		
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
		
		'UPGRADE_NOTE: Object rsPollAdd may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsPollAdd = Nothing
		'UPGRADE_NOTE: Object rsPollDetails may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsPollDetails = Nothing
		
		
	End Sub
	
	Private Sub frmCopyWQStd_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		modUtil.InitComboBox(cboStdName, "WQCRITERIA")
		
	End Sub
End Class