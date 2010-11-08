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

Friend Class frmAddCoeffSet
    Inherits System.Windows.Forms.Form

    Private _frmPoll As frmPollutants
    Private _frmNewPoll As frmNewPollutants


#Region "Events"
    Private Sub frmAddCoeffSet_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        modUtil.InitComboBox(cboLCType, "LCTYPE")
    End Sub


    Private Sub txtCoeffSetName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCoeffSetName.TextChanged
        If Len(txtCoeffSetName.Text) > 0 Then
            cmdOK.Enabled = True
        Else
            cmdOK.Enabled = False
        End If
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click

        If modUtil.UniqueName("CoefficientSet", (txtCoeffSetName.Text)) Then
            'uses code in frmPollutants to do the work
            If g_boolAddCoeff Then
                _frmPoll.AddCoefficient(txtCoeffSetName.Text, cboLCType.SelectedItem)
            Else
                _frmNewPoll.AddCoefficient(txtCoeffSetName.Text, cboLCType.SelectedItem)
            End If
            Me.Close()
        Else
            MsgBox(Err2, MsgBoxStyle.Critical, "Coefficient set name already in use.  Please enter new name")
            Exit Sub
        End If

        Me.Close()
    End Sub
#End Region

#Region "Helper Functions"
    Public Sub Init(ByRef frmPoll As frmPollutants, ByRef frmNewPoll As frmNewPollutants)
        _frmPoll = frmPoll
        _frmNewPoll = frmNewPoll
    End Sub

#End Region


End Class