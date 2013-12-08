'********************************************************************************************************
'File Name: frmCopyCoeffSet.vb
'Description: Form to copy coefficient set values
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
Imports System.Data.OleDb

Friend Class CopyCoefficientSetForm
    Private _frmPoll As PollutantsForm
    Private _frmNewPoll As NewPollutantForm

    Public Sub Init(ByRef cmdCoeffSet As OleDbCommand, ByRef frmPoll As PollutantsForm, ByRef frmNewPoll As NewPollutantForm)
        Try
            'The form is passed a recordest containing the names of all coefficient sets, allows for
            'easier populating

            If Not cmdCoeffSet Is Nothing Then
                Dim dataCoeff As OleDbDataReader = cmdCoeffSet.ExecuteReader()
                While dataCoeff.Read()
                    cboCoeffSet.Items.Add(dataCoeff("Name"))
                End While
                dataCoeff.Close()
            End If

            _frmPoll = frmPoll
            _frmNewPoll = frmNewPoll
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Protected Overrides Sub OK_Button_Click(sender As Object, e As EventArgs)
        Try
            If UniqueName("CoefficientSet", (txtCoeffSetName.Text)) And Trim(txtCoeffSetName.Text) <> "" Then
                If g_boolCopyCoeff Then
                    _frmPoll.CopyCoefficient(txtCoeffSetName.Text, cboCoeffSet.Text)
                Else
                    _frmNewPoll.CopyCoefficient(txtCoeffSetName.Text, cboCoeffSet.Text)
                End If
                MyBase.OK_Button_Click(sender, e)
            Else
                MsgBox("The name you have choosen for coefficient set is already in use.  Please pick another.", MsgBoxStyle.Critical, "Name In Use")
                With txtCoeffSetName
                    .SelectionStart = 0
                    .SelectionLength = Len(txtCoeffSetName.Text)
                    .Focus()
                End With
            End If
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub
End Class