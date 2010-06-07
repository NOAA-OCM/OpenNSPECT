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
    'Public g_pProgDialog As ESRI.ArcGIS.Framework.IProgressDialog2 'Prog Dialog
    'Public g_pStepProgressor As ESRI.ArcGIS.esriSystem.IStepProgressor 'Step Progress
    'Public g_pTrackCancel As ESRI.ArcGIS.esriSystem.ITrackCancel 'Cancel Button
    'Public g_pProDlgFact As ESRI.ArcGIS.Framework.IProgressDialogFactory 'Factory
    Public g_progdialog As frmProgressDialog
    Public g_boolCancel As Boolean

    Public Sub ProgDialog(ByRef strMessage As String, ByRef strTitle As String, ByRef lngMin As Integer, ByRef lngMax As Integer, ByRef lngValue As Integer, ByRef hwnd As Integer)
        'strMessage:  what's it doing
        'strTitle: Title of Dialog
        'lngMin: the Min value of the progress bar
        'lngMax: the max value of the progress bar
        'lngValue: current value of progress bar
        'hwnd: handle

        Try
            'first time through, set things up
            If g_progdialog Is Nothing Then
                g_boolCancel = True
                'create a CancelTracker
                g_progdialog = New frmProgressDialog
                g_frmProjectSetup.AddOwnedForm(g_progdialog)
                g_progdialog.Show()
            End If

            With g_progdialog
                .CancelEnabled = True
                .Title = strTitle
                .Description = strMessage
                .MinRange = lngMin
                .MaxRange = lngMax
                .Progress = lngValue
            End With

            Windows.Forms.Application.DoEvents()

        Catch ex As Exception
            MsgBox("Error Occurring during ModProgDialog:" & Err.Number)

        End Try

    End Sub

    Public Sub KillDialog()
        'Sub to kill all
        g_progdialog.Close()
        g_progdialog = Nothing
    End Sub
End Module