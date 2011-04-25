'********************************************************************************************************
'File Name: modProgDialog.vb
'Description: Functions handling the progress dialog of the model and plugin
'********************************************************************************************************
'The contents of this file are subject to the Mozilla Public License Version 1.1 (the "License"); 
'you may not use this file except in compliance with the License. You may obtain a copy of the License at 
'http://www.mozilla.org/MPL/ 
'Software distributed under the License is distributed on an "AS IS" basis, WITHOUT WARRANTY OF 
'ANY KIND, either express or implied. See the License for the specificlanguage governing rights and 
'limitations under the License. 
'
'Note: This code was converted from the vb6 NSPECT ArcGIS extension and so bears many of the old comments
'in the files where it was possible to leave them.
'
'Contributor(s): (Open source contributors should list themselves and their modifications here). 
'Oct 20, 2010:  Allen Anselmo allen.anselmo@gmail.com - 
'               Added licensing and comments to code

Module modProgDialog
    ' *************************************************************************************
    ' *  Perot Systems Government Services
    ' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
    ' *  modProgDialog
    ' *************************************************************************************
    ' *  Description: Code for handling progress dialog throughout OpenNSPECT
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


    Public Sub ProgDialog(ByRef strMessage As String, ByRef strTitle As String, ByRef lngMin As Integer, ByRef lngMax As Integer, ByRef lngValue As Integer, ByRef Owner As Windows.Forms.Form)
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
                g_progdialog.Show()
                If Not Owner Is Nothing Then
                    Owner.AddOwnedForm(g_progdialog)
                End If
            End If

            With g_progdialog
                .CancelEnabled = True
                .Title = strTitle
                .Description = strMessage
                .MinRange = lngMin
                .MaxRange = lngMax
                .Progress = lngValue
                If Not .TimerEnabled Then
                    .TimerEnabled = True
                End If
            End With

            Windows.Forms.Application.DoEvents()

        Catch ex As Exception
            MsgBox("Error Occurring during ModProgDialog:" & Err.Number)

        End Try

    End Sub


    Public Sub KillDialog()
        'Sub to kill all
        If Not g_progdialog Is Nothing Then
            g_progdialog.TimerEnabled = False
            g_progdialog.Close()
            g_progdialog = Nothing
        End If
    End Sub
End Module