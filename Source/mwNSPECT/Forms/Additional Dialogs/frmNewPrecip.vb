Option Strict Off
Option Explicit On
Friend Class frmNewPrecip
	Inherits System.Windows.Forms.Form

    Private _boolChange As Boolean

    Private Sub frmNewPrecip_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub txtPrecipName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPrecipName.TextChanged

    End Sub

    Private Sub txtDesc_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDesc.TextChanged

    End Sub

    Private Sub txtPrecipFile_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPrecipFile.TextChanged

    End Sub

    Private Sub cmdBrowseFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowseFile.Click

    End Sub

    Private Sub cboGridUnits_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboGridUnits.SelectedIndexChanged

    End Sub

    Private Sub cboPrecipUnits_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPrecipUnits.SelectedIndexChanged

    End Sub

    Private Sub cboTimePeriod_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTimePeriod.SelectedIndexChanged

    End Sub

    Private Sub txtRainingDays_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRainingDays.TextChanged

    End Sub

    Private Sub cboPrecipType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPrecipType.SelectedIndexChanged

    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click

        Dim intSave As Object

        If _boolChange Then
            intSave = MsgBox("You have made changes to this record, are you sure you want to quit?", MsgBoxStyle.YesNo, "Quit?")

            If intSave = MsgBoxResult.Yes Then
                Me.Close()
            ElseIf intSave = MsgBoxResult.No Then

                Exit Sub
            End If
        Else
            Me.Close()
        End If
    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click

    End Sub
End Class