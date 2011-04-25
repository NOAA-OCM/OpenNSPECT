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

''' <summary>
''' Code for handling progress dialog throughout OpenNSPECT
''' </summary>
''' <remarks></remarks>
Module modProgDialog
    Private ProgressForm As ProgressForm
    Public g_KeepRunning As Boolean

    ''' <summary>
    ''' Progs the dialog.
    ''' </summary>
    ''' <param name="message">what's it doing.</param>
    ''' <param name="title">Title of Dialog.</param>
    ''' <param name="min">the Min value of the progress bar.</param>
    ''' <param name="max">the max value of the progress bar.</param>
    ''' <param name="value">current value of progress bar.</param>
    ''' <param name="Owner">The owner.</param>
    Public Sub ShowProgress(ByRef message As String, ByRef title As String, ByRef min As Integer, ByRef max As Integer, ByRef value As Integer, ByRef Owner As Windows.Forms.Form)

        Try
            'first time through, set things up
            If ProgressForm Is Nothing Then
                g_KeepRunning = True
                ProgressForm = New ProgressForm
                ProgressForm.Show()
                If Not Owner Is Nothing Then
                    Owner.AddOwnedForm(ProgressForm)
                End If
            End If

            With ProgressForm
                .CancelEnabled = True
                .Title = title
                .Description = message
                .MinRange =
                .MaxRange = max
                .Progress = value
                If Not .TimerEnabled Then
                    .TimerEnabled = True
                End If
            End With

            Windows.Forms.Application.DoEvents()

        Catch ex As Exception
            HandleError("Error Occurring during ModProgDialog", ex)
        End Try

    End Sub


    ''' <summary>
    ''' Closes the dialog.
    ''' </summary>
    Public Sub CloseDialog()
        If Not ProgressForm Is Nothing Then
            ProgressForm.TimerEnabled = False
            ProgressForm.Close()
            ProgressForm = Nothing
        End If
    End Sub
End Module