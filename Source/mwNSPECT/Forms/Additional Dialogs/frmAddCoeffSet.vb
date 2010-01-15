Option Strict Off
Option Explicit On
Friend Class frmAddCoeffSet
	Inherits System.Windows.Forms.Form
    
#Region "Events"
    Private Sub frmAddCoeffSet_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        modUtil.InitComboBox(cboLCType, "LCTYPE")
    End Sub

    Private Sub txtCoeffSetName_TextChanged(ByVal sender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtCoeffSetName.TextChanged
        Dim Cancel As Boolean = eventArgs.Cancel

        If CDbl(Trim(CStr(Len(txtCoeffSetName.Text)))) <> 0 Then
            cmdOK.Enabled = True
        Else
            cmdOK.Enabled = False
        End If

        EventArgs.Cancel = Cancel
    End Sub

    Private Sub cboLCType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboLCType.SelectedIndexChanged

    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click

        'If modUtil.UniqueName("CoefficientSet", (txtCoeffSetName.Text)) Then
        '    'uses code in frmPollutants to do the work
        '    If g_boolAddCoeff Then
        '        frmPollutants.AddCoefficient((txtCoeffSetName.Text), VB6.GetItemString(cboLCType, cboLCType.SelectedIndex))
        '    Else
        '        frmNewPollutants.AddCoefficient((txtCoeffSetName.Text), VB6.GetItemString(cboLCType, cboLCType.SelectedIndex))
        '    End If
        'Else
        '    MsgBox(Err2, MsgBoxStyle.Critical, "Coefficient set name already in use.  Please enter new name")
        '    Exit Sub
        'End If

        'Me.Close()
    End Sub
#End Region

#Region "Helper Functions"

#End Region


End Class