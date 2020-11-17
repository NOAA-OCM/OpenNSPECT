Option Strict Off
Option Explicit On
Module modProgDialog
	' *************************************************************************************
	' *  Perot Systems Government Services
	' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
	' *  modProgDialog
	' *************************************************************************************
	' *  Description: Code for handling progress dialog throughout N-SPECT
	' *
	' *
	' *  Called By:  Various
	' *************************************************************************************
	Public g_pProgDialog As ESRI.ArcGIS.Framework.IProgressDialog2 'Prog Dialog
	Public g_pStepProgressor As ESRI.ArcGIS.esriSystem.IStepProgressor 'Step Progress
	Public g_pTrackCancel As ESRI.ArcGIS.esriSystem.ITrackCancel 'Cancel Button
	Public g_pProDlgFact As ESRI.ArcGIS.Framework.IProgressDialogFactory 'Factory
	Public g_boolCancel As Boolean
	
	Public Sub ProgDialog(ByRef strMessage As String, ByRef strTitle As String, ByRef lngMin As Integer, ByRef lngMax As Integer, ByRef lngValue As Integer, ByRef hwnd As Integer)
		'strMessage:  what's it doing
		'strTitle: Title of Dialog
		'lngMin: the Min value of the progress bar
		'lngMax: the max value of the progress bar
		'lngValue: current value of progress bar
		'hwnd: handle
		
		On Error GoTo ErrHandler
		
		'first time through, set things up
		If g_pProgDialog Is Nothing Then
			'create a CancelTracker
			g_pTrackCancel = New ESRI.ArcGIS.Display.CancelTracker
			
			With g_pTrackCancel
				.CancelOnClick = True
				.CancelOnKeyPress = True
			End With
			
			'create the ProgressDialog. This automatically displays the dialog
			g_pProDlgFact = New ESRI.ArcGIS.Framework.ProgressDialogFactory
			g_pProgDialog = g_pProDlgFact.Create(g_pTrackCancel, hwnd)
			
			With g_pProgDialog
				.Animation = ESRI.ArcGIS.Framework.esriProgressAnimationTypes.esriProgressGlobe
				.CancelEnabled = True
				.Title = strTitle
				.Description = "Processing...Please be patient."
			End With
			
			g_pStepProgressor = g_pProgDialog
			
		End If
		
		'these are the things that will change
		With g_pStepProgressor
			.MinRange = lngMin
			.MaxRange = lngMax
			.Position = lngValue
			'UPGRADE_WARNING: Couldn't resolve default property of object g_pStepProgressor.Message. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.Message = strMessage
			'UPGRADE_WARNING: Couldn't resolve default property of object g_pStepProgressor.Step. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.Step()
		End With
		
		'Global setup to Cancel the process
		g_boolCancel = g_pTrackCancel.Continue
		
		'g_pStepProgressor.Step
		
		Exit Sub
		
ErrHandler: 
		
		MsgBox("Error Occurring during ModProgDialog:" & Err.Number)
		
		'Kludge
		If Err.Number = 13 Then
			Resume Next
		End If
		
	End Sub
	
	Public Sub KillDialog()
		'Sub to kill all
		
		g_pProgDialog.HideDialog()
		
		'UPGRADE_NOTE: Object g_pTrackCancel may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		g_pTrackCancel = Nothing
		'UPGRADE_NOTE: Object g_pProDlgFact may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		g_pProDlgFact = Nothing
		'UPGRADE_NOTE: Object g_pStepProgressor may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		g_pStepProgressor = Nothing
		'UPGRADE_NOTE: Object g_pProgDialog may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		g_pProgDialog = Nothing
		
		
	End Sub
End Module