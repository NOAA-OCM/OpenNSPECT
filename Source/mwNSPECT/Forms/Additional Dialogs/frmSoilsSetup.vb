Option Strict Off
Option Explicit On
Friend Class frmSoilsSetup
	Inherits System.Windows.Forms.Form


    Private Sub frmSoilsSetup_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub txtSoilsName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSoilsName.TextChanged

    End Sub

    Private Sub txtDEMFile_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDEMFile.TextChanged

    End Sub

    Private Sub cmdDEMBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDEMBrowse.Click

    End Sub

    Private Sub txtSoilsDS_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSoilsDS.TextChanged

    End Sub

    Private Sub cmdBrowseFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowseFile.Click

    End Sub

    Private Sub cboSoilFields_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSoilFields.SelectedIndexChanged

    End Sub

    Private Sub cboSoilFieldsK_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSoilFieldsK.SelectedIndexChanged

    End Sub

    Private Sub txtMUSLEVal_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMUSLEVal.TextChanged

    End Sub

    Private Sub txtMUSLEExp_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMUSLEExp.TextChanged

    End Sub

    Private Sub cmdQuit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdQuit.Click

        Dim intvbYesNo As Short

        intvbYesNo = MsgBox("Do you want to save changes you made to soils setup?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Exclamation, "N-SPECT")

        If intvbYesNo = MsgBoxResult.Yes Then
            SaveSoils()
        ElseIf intvbYesNo = MsgBoxResult.No Then
            Me.Close()
        ElseIf intvbYesNo = MsgBoxResult.Cancel Then
            Exit Sub
        End If
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click

    End Sub


    Private Sub SaveSoils()

        ''Check data, if OK create soils grids
        'If ValidateData Then
        '    If CreateSoilsGrid((txtSoilsDS.Text), (cboSoilFields.Text), cboSoilFieldsK.Text) Then
        '        If frmSoils.Visible Then
        '            frmSoils.cboSoils.Items.Clear()
        '            modUtil.InitComboBox((frmSoils.cboSoils), "Soils")
        '            frmSoils.cboSoils.SelectedIndex = modUtil.GetCboIndex((txtSoilsName.Text), (frmSoils.cboSoils))
        '            Me.Close()
        '        End If
        '    Else
        '        Exit Sub
        '    End If
        'Else
        '    Exit Sub
        'End If


    End Sub

End Class