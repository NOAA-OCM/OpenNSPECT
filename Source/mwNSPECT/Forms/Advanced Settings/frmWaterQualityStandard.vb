'********************************************************************************************************
'File Name: frmWaterQualityStandard.vb
'Description: Form for displaying Water Quality standards data
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

Friend Class frmWaterQualityStandard
    Inherits System.Windows.Forms.Form

    Dim _bolChange As Boolean = False

    Const c_sModuleFileName As String = "frmWaterQualityStandard.vb"

#Region "Events"


    Private Sub frmWaterQualityStandard_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            _bolChange = False
            modUtil.InitComboBox(cboWQStdName, "WQCRITERIA")
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub cboWQStdName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboWQStdName.SelectedIndexChanged
        Try
            Dim intYesNo As Short

            If _bolChange Then
                intYesNo = MsgBox("You have made changes to the data.  Would you like to save before coninuing?", MsgBoxStyle.YesNo, "Save Changes?")
                If intYesNo = MsgBoxResult.Yes Then
                    UpdateData()
                    _bolChange = False
                ElseIf intYesNo = MsgBoxResult.No Then
                    _bolChange = False
                End If

            End If

            Dim strSQLWQStd As String
            Dim strSQLWQStdPoll As String

            'Selection based on combo box
            strSQLWQStd = "SELECT * FROM WQCRITERIA WHERE NAME LIKE '" & cboWQStdName.Text & "'"
            Dim WQCritCmd As New OleDbCommand(strSQLWQStd, modUtil.g_DBConn)
            Dim WQCrit As OleDbDataReader = WQCritCmd.ExecuteReader()

            If WQCrit.HasRows Then
                WQCrit.Read()
                txtWQStdDesc.Text = WQCrit.Item("Description")

                strSQLWQStdPoll = "SELECT POLLUTANT.NAME, POLL_WQCRITERIA.THRESHOLD, POLL_WQCRITERIA.POLL_WQCRITID " & "FROM POLL_WQCRITERIA INNER JOIN POLLUTANT " & "ON POLL_WQCRITERIA.POLLID = POLLUTANT.POLLID Where POLL_WQCRITERIA.WQCRITID = " & WQCrit.Item("WQCRITID")

                Dim WQCmd As New OleDbCommand(strSQLWQStdPoll, modUtil.g_DBConn)
                Dim WQ As New OleDbDataAdapter(WQCmd)
                Dim dt As New Data.DataTable()
                WQ.Fill(dt)

                dgvWaterQuality.DataSource = dt
            Else

                MsgBox("Warning: There are no water quality standards remaining.  Please add a new one.", MsgBoxStyle.Critical, "Recordset Empty")
            End If

            _bolChange = False
            OKEnable()

        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub txtWQStdDesc_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtWQStdDesc.TextChanged
        Try
            _bolChange = True
            OKEnable()
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub cmdQuit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdQuit.Click
        Try
            Close()
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)     'True, "cmdQuit_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try
    End Sub


    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Try
            If ValidateData() Then
                UpdateData()
                MsgBox("Data saved successfully.", MsgBoxStyle.Information, "Data Saved")
            End If

        Catch ex As Exception
            MsgBox("Error updating Water Quality Standards: " & Err.Number & vbNewLine & Err.Description, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub mnuNewWQStd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuNewWQStd.Click
        Try
            Dim addwq As New frmAddWQStd
            addwq.Init(Me, Nothing)
            addwq.ShowDialog()
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)     'True, "mnuNewWQStd_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try

    End Sub


    Private Sub mnuDelWQStd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDelWQStd.Click
        Try
            Dim intAns As Short
            intAns = MsgBox("Are you sure you want to delete the Water Quality Standard '" & cboWQStdName.SelectedItem & "'?", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Confirm Delete")
            'code to handle response

            'WQ Recordset
            Dim strWQStdDelete As String

            strWQStdDelete = "DELETE FROM WQCriteria WHERE NAME LIKE '" & cboWQStdName.Text & "'"
            Dim cmdWQ As New OleDbCommand(strWQStdDelete, g_DBConn)

            'strWQStdDelete = "SELECT * FROM WQCriteria WHERE NAME LIKE '" & cboWQStdName.Text & "'"
            'Dim cmdWQ As New OleDbCommand(strWQStdDelete, g_DBConn)
            'Dim wqSel As OleDbDataReader = cmdWQ.ExecuteReader()
            'wqSel.Read()
            'Doesn't seem to be used anymore
            'strWQPollDelete = "Delete * FROM POLL_WQCRITERIA WHERE WQCRITID =" & wqSel("WQCRITID").Value

            If Not (cboWQStdName.Text = "") Then
                'code to handle response
                If intAns = MsgBoxResult.Yes Then

                    'Delete the WaterQuality Standard
                    cmdWQ.ExecuteNonQuery()

                    MsgBox(cboWQStdName.SelectedItem & " deleted.", MsgBoxStyle.OkOnly, "Record Deleted")

                    cboWQStdName.Items.Clear()
                    modUtil.InitComboBox(cboWQStdName, "WQCRITERIA")
                    Me.Refresh()

                ElseIf intAns = MsgBoxResult.No Then
                    Exit Sub
                End If
            Else
                MsgBox("Please select a water quality standard", MsgBoxStyle.Critical, "No Standard Selected")
            End If

        Catch ex As Exception
            MsgBox("An Error occurred during deletion." & "  " & Err.Number & ": " & Err.Description, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub mnuCopyWQStd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCopyWQStd.Click
        Try
            Dim copywq As New frmCopyWQStd
            copywq.Init(Me)
            copywq.ShowDialog()
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)     'True, "mnuCopyWQStd_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try

    End Sub


    Private Sub mnuImpWQStd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuImpWQStd.Click
        Try
            Dim impwq As New frmImportWQStd()
            impwq.Init(Me)
            impwq.ShowDialog()
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)     'True, "mnuImpWQStd_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try

    End Sub


    Private Sub mnuExpWQStd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExpWQStd.Click
        Try
            Dim dlgSave As New Windows.Forms.SaveFileDialog
            dlgSave.Filter = Replace(MSG1, "<name>", "Water Quality Standard")
            dlgSave.Title = Replace(MSG3, "<name>", "Water Quality Standard")
            dlgSave.DefaultExt = ".txt"

            If dlgSave.ShowDialog = Windows.Forms.DialogResult.OK Then
                'Export Water Quality Standard to file - dlgCMD1.FileName
                ExportStandard(dlgSave.FileName)
            End If


        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)     'True, "mnuExpWQStd_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try

    End Sub


    Private Sub mnuWQHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuWQHelp.Click
        Try
            System.Windows.Forms.Help.ShowHelp(Me, modUtil.g_nspectPath & "\Help\nspect.chm", "wq_stnds.htm")
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub dgvWaterQuality_CellValueChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvWaterQuality.CellValueChanged
        Try
            _bolChange = True
            OKEnable()
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub
#End Region

#Region "Helper Functions"


    Private Sub UpdateData()
        Try
            Dim strSQLWQStd As String
            Dim strWQSelect As String
            Dim i As Short

            'Selection based on combo box, update Description
            strSQLWQStd = "SELECT * FROM WQCRITERIA WHERE NAME LIKE '" & cboWQStdName.Text & "'"
            Dim cmdWQ As New OleDbCommand(strSQLWQStd, g_DBConn)
            Dim adWQ As New OleDbDataAdapter(cmdWQ)
            Dim buWQ As New OleDbCommandBuilder(adWQ)
            buWQ.QuotePrefix = "["
            buWQ.QuoteSuffix = "]"
            Dim dt As New Data.DataTable
            adWQ.Fill(dt)
            dt.Rows(0)("Description") = txtWQStdDesc.Text
            adWQ.Update(dt)

            'Now update Threshold values
            For i = 0 To dgvWaterQuality.Rows.Count - 1
                strWQSelect = "SELECT * from POLL_WQCRITERIA WHERE POLL_WQCRITID = " & dgvWaterQuality.Rows(i).Cells(2).Value.ToString
                Dim cmdWQSel As New OleDbCommand(strWQSelect, g_DBConn)
                Dim adWQSel As New OleDbDataAdapter(cmdWQSel)
                Dim buWQSel As New OleDbCommandBuilder(adWQSel)
                buWQSel.QuotePrefix = "["
                buWQSel.QuoteSuffix = "]"
                Dim dtSel As New Data.DataTable
                adWQSel.Fill(dtSel)
                dtSel.Rows(0)("Threshold") = dgvWaterQuality.Rows(i).Cells(1).Value
                adWQSel.Update(dtSel)
            Next i

            _bolChange = False
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)     'False, "UpdateData " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try
    End Sub


    Private Function ValidateData() As Boolean
        Try
            Dim i As Short
            Dim val As String
            For i = 0 To dgvWaterQuality.Rows.Count - 1
                val = dgvWaterQuality.Rows(i).Cells(1).Value
                If IsNumeric(val) Then
                    If CShort(val) >= 0 Then
                        ValidateData = True
                    Else
                        MsgBox("Warning: Values must be greater than or equal to 0.", MsgBoxStyle.Critical, "Invalid Value")
                        Return False
                    End If
                ElseIf val <> "" Then
                    MsgBox("Numeric values only please.", MsgBoxStyle.Critical, "Numeric Values Only")
                    Return False
                End If
            Next
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)     'False, "ValidateData " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try
    End Function


    Public Sub UpdateWQ(ByVal strWQName As String)
        Try
            cboWQStdName.Items.Clear()
            modUtil.InitComboBox(cboWQStdName, "WQCRITERIA")
            cboWQStdName.SelectedIndex = modUtil.GetCboIndex(strWQName, cboWQStdName)
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub

    'Exports your current standard and pollutants to text or csv.


    Private Sub ExportStandard(ByRef strFileName As String)
        Try
            Dim out As New IO.StreamWriter(strFileName)

            'Write the name and descript.
            out.WriteLine(cboWQStdName.Text & "," & txtWQStdDesc.Text)

            Dim i As Short

            'Write name of pollutant and threshold
            For i = 0 To dgvWaterQuality.Rows.Count - 1
                out.WriteLine(dgvWaterQuality.Rows(i).Cells(0).Value & "," & dgvWaterQuality.Rows(i).Cells(1).Value)
            Next i

            out.Close()
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)     'False, "ExportStandard " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try
    End Sub


    Private Sub OKEnable()
        Try
            cmdSave.Enabled = _bolChange
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub

#End Region

End Class