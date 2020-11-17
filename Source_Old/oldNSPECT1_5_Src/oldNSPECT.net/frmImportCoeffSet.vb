Option Strict Off
Option Explicit On
Friend Class frmImportCoeffSet
	Inherits System.Windows.Forms.Form
	' *************************************************************************************
	' *  Perot Systems Government Services
	' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
	' *  frmImportCoeffSet
	' *************************************************************************************
	' *  Description:  Allows the user to import a text file containing a coeff set
	' *
	' *  Called By:  frmWQStd
	' *************************************************************************************
	
	' Variables used by the Error handler function - DO NOT REMOVE
	Const c_sModuleFileName As String = "frmImportCoeffSet.frm"
	
	
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
		
		
		If modUtil.UniqueName("CoefficientSet", (txtCoeffSetName.Text)) Then
			
			If modUtil.ValidateCoeffTextFile((txtImpFile.Text), (cboLCType.Text)) Then
				frmPollutants.UpdateCoeffSet(g_adoRSCoeff, (txtCoeffSetName.Text), (txtImpFile.Text))
			End If
			
		Else
			MsgBox("The name you have chosen is in use, please enter a different name.", MsgBoxStyle.Critical, "Name Detected")
			txtCoeffSetName.Focus()
			
		End If
		
		
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "cmdOK_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	Private Sub frmImportCoeffSet_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		On Error GoTo ErrorHandler
		
		InitComboBox(cboLCType, "LCTYPE")
		
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
End Class