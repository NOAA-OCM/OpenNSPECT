Option Strict Off
Option Explicit On
Friend Class frmNewPollutants
	Inherits System.Windows.Forms.Form


    Private Sub frmNewPollutants_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub txtPollutant_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPollutant.TextChanged

    End Sub

    Private Sub txtCoeffSet_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCoeffSet.TextChanged

    End Sub

    Private Sub cboLCType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboLCType.SelectedIndexChanged

    End Sub

    Private Sub txtCoeffSetDesc_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCoeffSetDesc.TextChanged

    End Sub

    Private Sub cmdQuit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdQuit.Click
        Me.Close()
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click

    End Sub

    Private Sub mnuCoeffNewSet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCoeffNewSet.Click

    End Sub

    Private Sub mnuCoeffCopySet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCoeffCopySet.Click

    End Sub
End Class