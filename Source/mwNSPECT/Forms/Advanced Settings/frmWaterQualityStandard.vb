Imports System.Data.OleDb

Friend Class frmWaterQualityStandard
    Inherits System.Windows.Forms.Form

    Dim _bolChange As Boolean

    Const c_sModuleFileName As String = "frmWaterQualityStandard.vb"

#Region "Events"
    Private Sub frmWaterQualityStandard_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        modUtil.InitComboBox(cboWQStdName, "WQCRITERIA")
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



        Catch ex As Exception
            HandleError(True, "cboWQStdName_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try
    End Sub

    Private Sub txtWQStdDesc_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtWQStdDesc.TextChanged

    End Sub


    Private Sub cmdQuit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdQuit.Click
        Try
            Me.Close()
        Catch ex As Exception
            HandleError(True, "cmdQuit_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
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
            addwq.ShowDialog()
        Catch ex As Exception
            HandleError(True, "mnuNewWQStd_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try

    End Sub

    Private Sub mnuDelWQStd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDelWQStd.Click

        '        On Error GoTo ErrHandler
        '        Dim intAns As Short
        '        intAns = MsgBox("Are you sure you want to delete the Water Quality Standard '" & VB6.GetItemString(cboWQStdName, cboWQStdName.SelectedIndex) & "'?", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Confirm Delete")
        '        'code to handle response

        '        'WQ Recordset
        '        Dim strWQStdDelete As String
        '        Dim strWQPollDelete As String

        '        strWQStdDelete = "SELECT * FROM WQCriteria WHERE NAME LIKE '" & cboWQStdName.Text & "'"

        '        rsWQStdDelete = New ADODB.Recordset

        '        rsWQStdDelete.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        '        rsWQStdDelete.Open(strWQStdDelete, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockOptimistic)

        '        strWQPollDelete = "Delete * FROM POLL_WQCRITERIA WHERE WQCRITID =" & rsWQStdDelete.Fields("WQCRITID").Value

        '        If Not (cboWQStdName.Text = "") Then
        '            'code to handle response
        '            If intAns = MsgBoxResult.Yes Then

        '                'Delete the WaterQuality Standard
        '                rsWQStdDelete.Delete(ADODB.AffectEnum.adAffectCurrent)
        '                rsWQStdDelete.Update()

        '                'modUtil.g_ADOConn.Execute strWQPollDelete

        '                MsgBox(VB6.GetItemString(cboWQStdName, cboWQStdName.SelectedIndex) & " deleted.", MsgBoxStyle.OkOnly, "Record Deleted")

        '                cboWQStdName.Items.Clear()
        '                modUtil.InitComboBox(cboWQStdName, "WQCRITERIA")
        '                Me.Refresh()

        '            ElseIf intAns = MsgBoxResult.No Then
        '                Exit Sub
        '            End If
        '        Else
        '            MsgBox("Please select a water quality standard", MsgBoxStyle.Critical, "No Standard Selected")
        '        End If

        '        'Cleanup
        '        rsWQStdDelete.Close()
        '        'UPGRADE_NOTE: Object rsWQStdDelete may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        rsWQStdDelete = Nothing

        '        Exit Sub
        'ErrHandler:
        '        MsgBox("An Error occurred during deletion." & "  " & Err.Number & ": " & Err.Description, MsgBoxStyle.Critical, "Error")

    End Sub

    Private Sub mnuCopyWQStd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCopyWQStd.Click
        Try
            Dim copywq As New frmCopyWQStd
            copywq.ShowDialog()
        Catch ex As Exception
            HandleError(True, "mnuCopyWQStd_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try

    End Sub

    Private Sub mnuImpWQStd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuImpWQStd.Click
        Try
            Dim impwq As New frmImportWQStd()
            impwq.ShowDialog()
        Catch ex As Exception
            HandleError(True, "mnuImpWQStd_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try

    End Sub

    Private Sub mnuExpWQStd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExpWQStd.Click
        Try

            'Dim intAns As Short

            ''browse...get output filename
            'dlgCMD1Open.FileName = CStr(Nothing)
            'dlgCMD1Save.FileName = CStr(Nothing)
            'With dlgCMD1
            '    'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            '    .Filter = Replace(MSG1, "<name>", "Water Quality Standard")
            '    .Title = Replace(MSG3, "<name>", "Water Quality Standard")
            '    .FilterIndex = 1
            '    .DefaultExt = ".txt"
            '    'UPGRADE_WARNING: FileOpenConstants constant FileOpenConstants.cdlOFNHideReadOnly was upgraded to OpenFileDialog.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
            '    'UPGRADE_WARNING: MSComDlg.CommonDialog property dlgCMD1.Flags was upgraded to dlgCMD1Open.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
            '    'UPGRADE_WARNING: FileOpenConstants constant FileOpenConstants.cdlOFNHideReadOnly was upgraded to OpenFileDialog.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
            '    .ShowReadOnly = False
            '    'UPGRADE_WARNING: MSComDlg.CommonDialog property dlgCMD1.Flags was upgraded to dlgCMD1Save.OverwritePrompt which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
            '    .OverwritePrompt = True
            '    .ShowDialog()
            'End With
            'If Len(dlgCMD1Open.FileName) > 0 Then
            '    'Export Water Quality Standard to file - dlgCMD1.FileName
            '    ExportStandard((dlgCMD1Open.FileName))
            'End If


        Catch ex As Exception
            HandleError(True, "mnuExpWQStd_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try

    End Sub

    Private Sub mnuAddRow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddRow.Click

    End Sub

    Private Sub mnuInsertRow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuInsertRow.Click

    End Sub

    Private Sub mnuDeleteRow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDeleteRow.Click

    End Sub

    Private Sub mnuWQHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuWQHelp.Click
        System.Windows.Forms.Help.ShowHelp(Me, modUtil.g_nspectPath & "\Help\nspect.chm", "wq_stnds.htm")
    End Sub
#End Region

#Region "Helper Functions"

    Private Sub UpdateData()
        Try
            'Dim strSQLWQStd As String
            'Dim strWQSelect As String
            'Dim rsWQstd As New ADODB.Recordset
            'rsWQStdCboClick = New ADODB.Recordset
            'Dim rsPollUpdate As New ADODB.Recordset
            'Dim i As Short

            'Dim booYesNo As Short

            ''Selection based on combo box, update Description
            'strSQLWQStd = "SELECT * FROM WQCRITERIA WHERE NAME LIKE '" & cboWQStdName.Text & "'"

            'With rsWQStdCboClick
            '    .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            '    .Open(strSQLWQStd, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            'End With

            'rsWQStdCboClick.Fields("Description").Value = txtWQStdDesc.Text
            'rsWQStdCboClick.Update()

            ''Now update Threshold values
            'For i = 1 To grdWQStd.Rows - 1
            '    strWQSelect = "SELECT * from POLL_WQCRITERIA WHERE POLL_WQCRITID = " & grdWQStd.get_TextMatrix(i, 3)

            '    rsWQstd.Open(strWQSelect, g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            '    rsWQstd.Fields("Threshold").Value = grdWQStd.get_TextMatrix(i, 2)
            '    rsWQstd.Update()
            '    rsWQstd.Close()
            'Next i

            'm_bolChange = False

            ''Cleanup
            ''UPGRADE_NOTE: Object rsWQstd may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            'rsWQstd = Nothing
            'rsWQStdCboClick.Close()
            ''UPGRADE_NOTE: Object rsWQStdCboClick may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            'rsWQStdCboClick = Nothing

        Catch ex As Exception
            HandleError(False, "UpdateData " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try
    End Sub

    Private Function ValidateData() As Boolean
        Try
            'Dim i As Short

            'For i = 1 To grdWQStd.Rows - 1

            '    If IsNumeric(grdWQStd.get_TextMatrix(i, 2)) Then
            '        If CShort(grdWQStd.get_TextMatrix(i, 2)) >= 0 Then
            '            ValidateData = True
            '        Else
            '            MsgBox("Warning: Values must be greater than or equal to 0.", MsgBoxStyle.Critical, "Invalid Value")
            '            grdWQStd.row = i
            '            grdWQStd.col = 2
            '            KeyMoveUpdate()
            '            ValidateData = False
            '        End If
            '    ElseIf grdWQStd.get_TextMatrix(i, 2) <> "" Then
            '        MsgBox("Numeric values only please.", MsgBoxStyle.Critical, "Numeric Values Only")
            '        grdWQStd.row = i
            '        grdWQStd.col = 2
            '        KeyMoveUpdate()
            '        ValidateData = False
            '    End If
            'Next
        Catch ex As Exception
            HandleError(False, "ValidateData " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try
    End Function

    Public Sub UpdateWQ(ByVal strWQName As String)
        cboWQStdName.Items.Clear()
        modUtil.InitComboBox(cboWQStdName, "WQCRITERIA")
        cboWQStdName.SelectedIndex = modUtil.GetCboIndex(strWQName, cboWQStdName)
    End Sub

#End Region

End Class