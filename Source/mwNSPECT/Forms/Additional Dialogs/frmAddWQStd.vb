Imports System.Windows.Forms
Imports System.Data.OleDb
Friend Class frmAddWQStd
    Inherits System.Windows.Forms.Form

    Private _frmWQStd As frmWaterQualityStandard
    Private _frmPrj As frmProjectSetup
    Private _Change As Boolean

    Const c_sModuleFileName As String = "frmAddWQStd.frm"


#Region "Events"

    Private Sub frmAddWQStd_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try

            'Populate cbo with pollutant names
            _Change = False

            Dim strPollutant As String = "SELECT NAME FROM POLLUTANT ORDER BY NAME ASC"
            Dim pollCmd As New OleDbCommand(strPollutant, g_DBConn)
            Dim datPoll As OleDbDataReader = pollCmd.ExecuteReader

            Dim idx As Integer
            dgvWaterQuality.Rows.Clear()
            Do While datPoll.Read()
                idx = dgvWaterQuality.Rows.Add()
                dgvWaterQuality.Rows(idx).Cells("Pollutant").Value = datPoll.Item("Name")
            Loop
            datPoll.Close()
            cmdSave.Enabled = False
        Catch ex As Exception
            HandleError(True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try
    End Sub

    Private Sub txtWQStdName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtWQStdName.TextChanged
        txtWQStdName.Text = Replace(txtWQStdName.Text, "'", "")
    End Sub

    Private Sub txtWQStdDesc_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtWQStdDesc.TextChanged

    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Dim intvbYesNo As Short = MsgBox("Are you sure you want to exit?  All changes not saved will be lost.", MsgBoxStyle.YesNo, "Exit?")

        If intvbYesNo = MsgBoxResult.Yes Then
            If Not _frmPrj Is Nothing Then
                _frmPrj.cboWQStd.SelectedIndex = 0
            End If

            Me.Close()
        Else
            Exit Sub
        End If

    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Try
            Dim strName As String
            Dim strDescript As String
            Dim strCmd As String

            'Get rid of possible apostrophes
            strName = Replace(Trim(txtWQStdName.Text), "'", "''")
            strDescript = Trim(txtWQStdDesc.Text)

            If Len(strName) = 0 Then
                MsgBox("Please enter a name for the water quality standard.", MsgBoxStyle.Critical, "Empty Name Field")
                txtWQStdName.Focus()
                Exit Sub
            Else
                'Name Check
                If modUtil.UniqueName("WQCRITERIA", (txtWQStdName.Text)) Then
                    'Value check
                    If CheckThreshValues Then
                        strCmd = "INSERT INTO WQCRITERIA (NAME,DESCRIPTION) VALUES ('" & Replace(txtWQStdName.Text, "'", "''") & "', '" & Replace(strDescript, "'", "''") & "')"
                        Dim cmdInsert As New OleDbCommand(strCmd, g_DBConn)
                        cmdInsert.ExecuteNonQuery()
                    Else
                        MsgBox("Threshold values must be numeric.", MsgBoxStyle.Critical, "Check Threshold Value")
                        Exit Sub
                    End If
                Else
                    MsgBox("The name you have chosen is already in use.  Please select another.", MsgBoxStyle.Critical, "Select Unique Name")
                    Exit Sub
                End If
            End If

            'If it gets here, time to add the pollutants
            Dim i As Short
            i = 0

            For Each row As DataGridViewRow In dgvWaterQuality.Rows
                PollutantAdd(txtWQStdName.Text, row.Cells("Pollutant").Value, row.Cells("Threshold").Value)
            Next

            MsgBox(txtWQStdName.Text & " successfully added.", MsgBoxStyle.OkOnly, "Record Added")

            'Clean up stuff
            If Not _frmWQStd Is Nothing Then
                _frmWQStd.UpdateWQ(txtWQStdName.Text)
            ElseIf Not _frmPrj Is Nothing Then
                _frmPrj.UpdateWQ(txtWQStdName.Text)
            End If

            Me.Close()

        Catch ex As Exception
            HandleError(True, "cmdSave_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try
    End Sub

    Private Sub dgvWaterQuality_CellValueChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvWaterQuality.CellValueChanged
        cmdSave.Enabled = True
    End Sub
#End Region

#Region "Helper Functions"
    Public Sub Init(ByRef frmWQStd As frmWaterQualityStandard, ByRef frmPrj As frmProjectSetup)
        _frmWQStd = frmWQStd
        _frmPrj = frmPrj
    End Sub


    Private Function CheckThreshValues() As Boolean
        Try
            For Each row As DataGridViewRow In dgvWaterQuality.Rows
                If IsNumeric(row.Cells("Threshold").Value) Or row.Cells("Threshold").Value = "" Then
                    CheckThreshValues = True
                Else
                    CheckThreshValues = False
                    Exit Function
                End If
            Next
        Catch ex As Exception
            HandleError(False, "CheckThreshValues " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try
    End Function

    Private Sub PollutantAdd(ByRef strName As String, ByRef strPoll As String, ByRef intThresh As String)
        Try
            Dim strPollAdd As String
            Dim strPollDetails As String
            Dim strCmdInsert As String

            'Get the WQCriteria values using the name
            strPollAdd = "SELECT * FROM WQCriteria WHERE NAME = " & "'" & strName & "'"
            Dim cmdPollAdd As New OleDbCommand(strPollAdd, g_DBConn)
            Dim datPollAdd As OleDbDataReader = cmdPollAdd.ExecuteReader()
            datPollAdd.Read()

            'Get the pollutant particulars
            strPollDetails = "SELECT * FROM POLLUTANT WHERE NAME =" & "'" & strPoll & "'"
            Dim cmdPollDetails As New OleDbCommand(strPollDetails, g_DBConn)
            Dim datPollDetails As OleDbDataReader = cmdPollDetails.ExecuteReader()
            datPollDetails.Read()

            If Trim(intThresh) = "" Then
                strCmdInsert = "INSERT INTO POLL_WQCRITERIA (PollID,WQCritID) VALUES ('" & datPollDetails.Item("POLLID") & "', '" & datPollAdd.Item("WQCRITID") & "')"
            Else
                strCmdInsert = "INSERT INTO POLL_WQCRITERIA (PollID,WQCritID,Threshold) VALUES ('" & datPollDetails.Item("POLLID") & "', '" & datPollAdd.Item("WQCRITID") & "'," & intThresh & ")"
            End If
            Dim cmdInsert As New OleDbCommand(strCmdInsert, g_DBConn)
            cmdInsert.ExecuteNonQuery()

            datPollAdd.Close()
            datPollDetails.Close()

        Catch ex As Exception
            HandleError(False, "PollutantAdd " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try
    End Sub
#End Region


End Class