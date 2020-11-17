Attribute VB_Name = "modProgDialog"
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
Public g_pProgDialog As IProgressDialog2            'Prog Dialog
Public g_pStepProgressor As IStepProgressor         'Step Progress
Public g_pTrackCancel As ITrackCancel               'Cancel Button
Public g_pProDlgFact As IProgressDialogFactory      'Factory
Public g_boolCancel As Boolean

Public Sub ProgDialog(strMessage As String, strTitle As String, lngMin As Long, _
                      lngMax As Long, lngValue As Long, hwnd As Long)
'strMessage:  what's it doing
'strTitle: Title of Dialog
'lngMin: the Min value of the progress bar
'lngMax: the max value of the progress bar
'lngValue: current value of progress bar
'hwnd: handle

On Error GoTo ErrHandler:
    
'first time through, set things up
If g_pProgDialog Is Nothing Then
    'create a CancelTracker
    Set g_pTrackCancel = New CancelTracker
    
    With g_pTrackCancel
        .CancelOnClick = True
        .CancelOnKeyPress = True
    End With
    
    'create the ProgressDialog. This automatically displays the dialog
    Set g_pProDlgFact = New ProgressDialogFactory
    Set g_pProgDialog = g_pProDlgFact.Create(g_pTrackCancel, hwnd)
    
    With g_pProgDialog
        .Animation = esriProgressGlobe
        .CancelEnabled = True
        .Title = strTitle
        .Description = "Processing...Please be patient."
    End With
    
    Set g_pStepProgressor = g_pProgDialog
    
End If
    
'these are the things that will change
With g_pStepProgressor
    .MinRange = lngMin
    .MaxRange = lngMax
    .Position = lngValue
    .Message = strMessage
    .Step
End With

'Global setup to Cancel the process
g_boolCancel = g_pTrackCancel.Continue

'g_pStepProgressor.Step

Exit Sub

ErrHandler:
    
    MsgBox "Error Occurring during ModProgDialog:" & Err.Number
    
    'Kludge
    If Err.Number = 13 Then
        Resume Next
    End If

End Sub

Public Sub KillDialog()
'Sub to kill all
    
    g_pProgDialog.HideDialog
    
    Set g_pTrackCancel = Nothing
    Set g_pProDlgFact = Nothing
    Set g_pStepProgressor = Nothing
    Set g_pProgDialog = Nothing
    

End Sub


