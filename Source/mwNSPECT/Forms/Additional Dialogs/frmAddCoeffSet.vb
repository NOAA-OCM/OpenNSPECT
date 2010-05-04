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