Imports System.Data.OleDb
Friend Class frmCopyCoeffSet
    Inherits System.Windows.Forms.Form

    Private _frmPoll As frmPollutants
    Private _frmNewPoll As frmNewPollutants

    Public Sub Init(ByRef cmdCoeffSet As OleDbCommand, ByRef frmPoll As frmPollutants, ByRef frmNewPoll As frmNewPollutants)
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
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        If modUtil.UniqueName("CoefficientSet", (txtCoeffSetName.Text)) And Trim(txtCoeffSetName.Text) <> "" Then
            If g_boolCopyCoeff Then
                _frmPoll.CopyCoefficient(txtCoeffSetName.Text, cboCoeffSet.Text)
            Else
                _frmNewPoll.CopyCoefficient(txtCoeffSetName.Text, cboCoeffSet.Text)
            End If
            Me.Close()
        Else
            MsgBox("The name you have choosen for coefficient set is already in use.  Please pick another.", MsgBoxStyle.Critical, "Name In Use")
            With txtCoeffSetName
                .SelectionStart = 0
                .SelectionLength = Len(txtCoeffSetName.Text)
                .Focus()
            End With

        End If
    End Sub
End Class