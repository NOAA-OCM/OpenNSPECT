'********************************************************************************************************
'File Name: frmAddCoeffSet.vb
'Description: Form to ad coefficient sets
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

Friend Class NewCoefficientSetForm
    Private _frmPoll As PollutantsForm
    Private _frmNewPoll As NewPollutantForm

#Region "Events"

    Private Sub frmAddCoeffSet_Load (ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        Try
            InitComboBox (cboLCType, "LCTYPE")
        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub

    Private Sub txtCoeffSetName_TextChanged (ByVal sender As Object, ByVal e As EventArgs) _
        Handles txtCoeffSetName.TextChanged
        Try
            If Len (txtCoeffSetName.Text) > 0 Then
                OK_Button.Enabled = True
            Else
                OK_Button.Enabled = False
            End If
        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub

    Protected Overrides Sub OK_Button_Click (sender As Object, e As EventArgs)

        Try
            If UniqueName ("CoefficientSet", (txtCoeffSetName.Text)) Then
                'uses code in frmPollutants to do the work
                If g_boolAddCoeff Then
                    _frmPoll.AddCoefficient (txtCoeffSetName.Text, cboLCType.SelectedItem)
                Else
                    _frmNewPoll.AddCoefficient (txtCoeffSetName.Text, cboLCType.SelectedItem)
                End If
                Close()
            Else
                MsgBox (Err2, MsgBoxStyle.Critical, "Coefficient set name already in use.  Please enter new name")
                Return
            End If

            MyBase.OK_Button_Click (sender, e)
        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub

#End Region

#Region "Helper Functions"

    Public Sub Init (ByRef frmPoll As PollutantsForm, ByRef frmNewPoll As NewPollutantForm)
        Try
            _frmPoll = frmPoll
            _frmNewPoll = frmNewPoll
        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub

#End Region
End Class