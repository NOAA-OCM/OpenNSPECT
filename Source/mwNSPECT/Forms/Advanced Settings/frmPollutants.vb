Imports System.Data
Imports System.Data.OleDb
Friend Class frmPollutants
    Inherits System.Windows.Forms.Form

    Const c_sModuleFileName As String = "frmPollutants.vb"

    Private _boolLoaded As Boolean
    Private _boolChanged As Boolean 'Bool for enabling save button
    Private _boolDescChanged As Boolean 'Boolship for seeing if Description Changed
    Private _boolSaved As Boolean 'Boolship for whether or not things have saved


    Private _intCurFrame As Short 'Current Frame visible in SSTab
    Private _intPollRow As Short 'Row Number for grdPolldef
    Private _intPollCol As Short 'Column Number for grdPollDef
    Private _intRowWQ As Short 'Row Number for grdPollWQStd
    Private _intColWQ As Short 'Column Number for grdPollDef
    Private _intPollID As Short 'There's a need to have the PollID so we'll store it here
    Private _intLCTypeID As Short 'Land Class (CCAP) ID - needed to add new coefficient sets
    Private _intCoeffID As Short 'Key for CoefficientSetID - needed to add new coefficients 'See above

    Private _strUndoText As String 'Text for txtActiveCell     |
    Private _strUndoTextWQ As String 'Text for txtActiveCellWQ   |-all three used to detect change
    Private _strUndoDesc As String 'Text for Description       |
    Private _strLCType As String 'Need for name, we'll store here

#Region "Events"
    Private Sub frmPollutants_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            _boolChanged = False

            'Toss in the names of all pollutants and call the cbo click event
            InitComboBox(cboPollName, "Pollutant")

            SSTab1.SelectedIndex = 0
            _boolLoaded = True
            _boolChanged = False
            _boolSaved = False
        Catch ex As Exception
            HandleError(True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try
    End Sub

    Private Sub cboPollName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPollName.SelectedIndexChanged

        Dim strSQLPollutant As String
        Dim strSQLLCType As String
        Dim strSQLCoeff As String
        Dim strSQLWQStd As String

        Try
            'Check to see if things have changed
            If _boolChanged Then

                intYesNo = MsgBox(strYesNo, MsgBoxStyle.YesNo, strYesNoTitle)

                If intYesNo = MsgBoxResult.Yes Then

                    UpdateValues()

                    _boolChanged = False

                    'Selection based on combo box
                    strSQLPollutant = "SELECT * FROM POLLUTANT WHERE NAME = '" & cboPollName.Text & "'"
                    Dim pollCmd As New OleDbCommand(strSQLPollutant, modUtil.g_DBConn)
                    Dim poll As OleDbDataReader = pollCmd.ExecuteReader()
                    poll.Read()
                    _intPollID = poll.Item("POLLID")

                    strSQLCoeff = "SELECT * FROM CoefficientSet WHERE POLLID = " & poll.Item("POLLID") & ""
                    Dim coefCmd As New OleDbCommand(strSQLCoeff, modUtil.g_DBConn)
                    Dim coef As OleDbDataReader = coefCmd.ExecuteReader()
                    coef.Read()
                    
                    strSQLLCType = "SELECT NAME, LCTYPEID FROM LCTYPE WHERE LCTYPEID = " & coef.Item("LCTypeID") & ""
                    Dim LCCmd As New OleDbCommand(strSQLLCType, modUtil.g_DBConn)
                    Dim LC As OleDbDataReader = LCCmd.ExecuteReader()
                    LC.Read()
                    _strLCType = LC.Item("Name")
                    _intLCTypeID = LC.Item("LCTypeID")

                    'Fill everything based on that
                    cboCoeffSet.Items.Clear()
                    cboCoeffSet.Items.Add(coef.Item("Name"))
                    Do While coef.Read()
                        cboCoeffSet.Items.Add(coef.Item("Name"))
                    Loop
                    cboCoeffSet.SelectedIndex = 0


                    txtLCType.Text = _strLCType

                    'Fill the Water Quality Standards Tab
                    strSQLWQStd = "SELECT WQCRITERIA.Name, WQCRITERIA.Description," & "POLL_WQCRITERIA.Threshold, POLL_WQCRITERIA.POLL_WQCRITID FROM WQCRITERIA " & "LEFT OUTER JOIN POLL_WQCRITERIA ON WQCRITERIA.WQCRITID = POLL_WQCRITERIA.WQCRITID " & "Where POLL_WQCRITERIA.POLLID = " & poll.Item("POLLID")
                    Dim wqCmd As New OleDbCommand(strSQLWQStd, modUtil.g_DBConn)
                    Dim wq As New OleDbDataAdapter(wqCmd)
                    Dim dt As New Data.DataTable
                    wq.Fill(dt)
                    dgvWaterQuality.DataSource = dt
                ElseIf intYesNo = MsgBoxResult.No Then

                    _boolChanged = False

                    'Selection based on combo box
                    strSQLPollutant = "SELECT * FROM POLLUTANT WHERE NAME = '" & cboPollName.Text & "'"
                    Dim pollCmd As New OleDbCommand(strSQLPollutant, modUtil.g_DBConn)
                    Dim poll As OleDbDataReader = pollCmd.ExecuteReader()
                    poll.Read()
                    _intPollID = poll.Item("POLLID")

                    strSQLCoeff = "SELECT * FROM CoefficientSet WHERE POLLID = " & poll.Item("POLLID") & ""
                    Dim coefCmd As New OleDbCommand(strSQLCoeff, modUtil.g_DBConn)
                    Dim coef As OleDbDataReader = coefCmd.ExecuteReader()
                    coef.Read()

                    strSQLLCType = "SELECT NAME, LCTYPEID FROM LCTYPE WHERE LCTYPEID = " & coef.Item("LCTypeID") & ""
                    Dim LCCmd As New OleDbCommand(strSQLLCType, modUtil.g_DBConn)
                    Dim LC As OleDbDataReader = LCCmd.ExecuteReader()
                    LC.Read()
                    _strLCType = LC.Item("Name")
                    _intLCTypeID = LC.Item("LCTypeID")


                    'Fill everything based on that
                    cboCoeffSet.Items.Clear()
                    cboCoeffSet.Items.Add(coef.Item("Name"))
                    Do While coef.Read()
                        cboCoeffSet.Items.Add(coef.Item("Name"))
                    Loop
                    cboCoeffSet.SelectedIndex = 0

                    txtLCType.Text = _strLCType

                    'Fill the Water Quality Standards Tab
                    strSQLWQStd = "SELECT WQCRITERIA.Name, WQCRITERIA.Description," & "POLL_WQCRITERIA.Threshold, POLL_WQCRITERIA.POLL_WQCRITID FROM WQCRITERIA " & "LEFT OUTER JOIN POLL_WQCRITERIA ON WQCRITERIA.WQCRITID = POLL_WQCRITERIA.WQCRITID " & "Where POLL_WQCRITERIA.POLLID = " & poll.Item("POLLID")
                    Dim wqCmd As New OleDbCommand(strSQLWQStd, modUtil.g_DBConn)
                    Dim wq As New OleDbDataAdapter(wqCmd)
                    Dim dt As New Data.DataTable
                    wq.Fill(dt)
                    dgvWaterQuality.DataSource = dt
                End If
            Else

                _boolChanged = False

                'Selection based on combo box
                strSQLPollutant = "SELECT * FROM POLLUTANT WHERE NAME = '" & cboPollName.Text & "'"
                Dim pollCmd As New OleDbCommand(strSQLPollutant, modUtil.g_DBConn)
                Dim poll As OleDbDataReader = pollCmd.ExecuteReader()
                poll.Read()
                _intPollID = poll.item("POLLID")

                strSQLCoeff = "SELECT * FROM CoefficientSet WHERE POLLID = " & poll.Item("POLLID") & ""
                Dim coefCmd As New OleDbCommand(strSQLCoeff, modUtil.g_DBConn)
                Dim coef As OleDbDataReader = coefCmd.ExecuteReader()
                coef.Read()

                strSQLLCType = "SELECT NAME, LCTYPEID FROM LCTYPE WHERE LCTYPEID = " & coef.Item("LCTypeID") & ""
                Dim LCCmd As New OleDbCommand(strSQLLCType, modUtil.g_DBConn)
                Dim LC As OleDbDataReader = LCCmd.ExecuteReader()
                LC.Read()
                _strLCType = lc.item("Name")
                _intLCTypeID = lc.item("LCTypeID")

                'Fill everything based on that
                cboCoeffSet.Items.Clear()
                cboCoeffSet.Items.Add(coef.Item("Name"))
                Do While coef.Read()
                    cboCoeffSet.Items.Add(coef.Item("Name"))
                Loop
                cboCoeffSet.SelectedIndex = 0

                txtLCType.Text = _strLCType

                'Fill the Water Quality Standards Tab
                strSQLWQStd = "SELECT WQCRITERIA.Name, WQCRITERIA.Description," & "POLL_WQCRITERIA.Threshold, POLL_WQCRITERIA.POLL_WQCRITID FROM WQCRITERIA " & "LEFT OUTER JOIN POLL_WQCRITERIA ON WQCRITERIA.WQCRITID = POLL_WQCRITERIA.WQCRITID " & "Where POLL_WQCRITERIA.POLLID = " & poll.Item("POLLID")
                Dim wqCmd As New OleDbCommand(strSQLWQStd, modUtil.g_DBConn)
                Dim wq As New OleDbDataAdapter(wqCmd)
                Dim dt As New Data.DataTable
                wq.Fill(dt)
                dgvWaterQuality.DataSource = dt

            End If

        Catch ex As Exception
            HandleError(True, "cboPollName_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try
    End Sub

    Private Sub cboCoeffSet_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCoeffSet.SelectedIndexChanged
        Try

            If _boolChanged Then

                intYesNo = MsgBox(strYesNo, MsgBoxStyle.YesNo, strYesNoTitle)

                If intYesNo = MsgBoxResult.Yes Then

                    If ValidateGridValues() Then
                        UpdateValues()
                        _boolChanged = False
                    End If
                Else
                    NoSaveCoeffSetChange()
                End If

            Else
                NoSaveCoeffSetChange()
            End If

        Catch ex As Exception
            HandleError(True, "cboCoeffSet_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try
    End Sub

    Private Sub NoSaveCoeffSetChange()
        Dim strSQLFullCoeff As String
        Dim strSQLCoeffs As String
        _boolChanged = False

        strSQLFullCoeff = "SELECT COEFFICIENTSET.NAME, COEFFICIENTSET.DESCRIPTION, " & "COEFFICIENTSET.COEFFSETID, LCTYPE.NAME as NAME2 " & "FROM COEFFICIENTSET INNER JOIN LCTYPE " & "ON COEFFICIENTSET.LCTYPEID = LCTYPE.LCTYPEID " & "WHERE COEFFICIENTSET.NAME LIKE '" & cboCoeffSet.Text & "'"
        Dim coefCmd As New OleDbCommand(strSQLFullCoeff, modUtil.g_DBConn)
        Dim coef As OleDbDataReader = coefCmd.ExecuteReader()
        coef.Read()

        With txtCoeffSetDesc
            .Text = coef.Item("Description") & ""
            .Refresh()
        End With

        txtLCType.Text = coef.Item("Name2") & ""

        strSQLCoeffs = "SELECT LCCLASS.Value, LCCLASS.Name, COEFFICIENT.Coeff1 As Type1, COEFFICIENT.Coeff2 as Type2, " & "COEFFICIENT.Coeff3 as Type3, COEFFICIENT.Coeff4 as Type4, COEFFICIENT.CoeffID, COEFFICIENT.LCCLASSID " & "FROM LCCLASS LEFT OUTER JOIN COEFFICIENT " & "ON LCCLASS.LCCLASSID = COEFFICIENT.LCCLASSID " & "WHERE COEFFICIENT.COEFFSETID = " & coef.Item("CoeffSetID") & " ORDER BY LCCLASS.VALUE"
        Dim coefsCmd As New OleDbCommand(strSQLCoeffs, modUtil.g_DBConn)
        Dim coefs As New OleDbDataAdapter(coefsCmd)
        Dim dt As New Data.DataTable
        coefs.Fill(dt)
        dgvCoef.DataSource = dt

    End Sub


    Private Sub txtLCType_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLCType.TextChanged

    End Sub

    Private Sub txtCoeffSetDesc_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCoeffSetDesc.TextChanged

    End Sub


    Private Sub cmdQuit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdQuit.Click
        Try

            If (_boolChanged Or _boolDescChanged) And Not _boolSaved Then

                intYesNo = MsgBox(strYesNo, MsgBoxStyle.YesNo, strYesNoTitle)

                If intYesNo = MsgBoxResult.Yes Then

                    If ValidateGridValues() Then
                        UpdateValues()
                        MsgBox("Data saved successfully.", MsgBoxStyle.OkOnly, "Save Successful")
                        _boolChanged = False
                        _boolDescChanged = False
                        _boolSaved = True
                        Me.Close()
                    End If

                Else

                    Me.Close()

                End If
            Else

                Me.Close()
            End If
        Catch ex As Exception
            HandleError(True, "cmdQuit_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Try
            If ValidateGridValues() Then

                UpdateValues()
                _boolSaved = True
                MsgBox(cboPollName.Text & " saved successfully.", MsgBoxStyle.Information, "N-SPECT")
                Me.Close()

            End If
        Catch ex As Exception
            HandleError(True, "cmdSave_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try
    End Sub


    Private Sub mnuAddPoll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddPoll.Click
        Try
            Dim newPoll As New frmNewPollutants
            newPoll.ShowDialog()
        Catch ex As Exception
            HandleError(True, "mnuAddPoll_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try
    End Sub

    Private Sub mnuDeletePoll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDeletePoll.Click
        Try
            Dim intAns As Short
            intAns = MsgBox("Are you sure you want to delete the pollutant '" & cboPollName.Text & "'?", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Confirm Delete")
            'code to handle response

            If intAns = MsgBoxResult.Yes Then
                DeletePollutant(cboPollName.Text)
            Else
                Exit Sub
            End If
        Catch ex As Exception
            HandleError(True, "mnuDeletePoll_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try
    End Sub

    Private Sub mnuCoeffNewSet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCoeffNewSet.Click
        Try
            g_boolAddCoeff = True
            Dim addCoeff As New frmAddCoeffSet
            addCoeff.ShowDialog()
        Catch ex As Exception
            HandleError(True, "mnuCoeffNewSet_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try

    End Sub

    Private Sub mnuCoeffCopySet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCoeffCopySet.Click
        Try
            g_boolCopyCoeff = True
            Dim newCopyCoef As New frmCopyCoeffSet
            'TODO: Handle initializatin
            'newCopyCoef.init(rsCoeff)
            newCopyCoef.ShowDialog()
        Catch ex As Exception
            HandleError(True, "mnuCoeffCopySet_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try
    End Sub

    Private Sub mnuCoeffDeleteSet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCoeffDeleteSet.Click

        'Using straight command text to rid ourselves of the dreaded coefficient sets

        Try
            Dim intAns As Short
            intAns = MsgBox("Are you sure you want to delete the coefficient set '" & cboCoeffSet.Text & "' associated with pollutant '" & cboPollName.Text & "'?", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Confirm Delete")

            'code to handle response
            'Dim strDeleteCoeffSet As String
            If intAns = MsgBoxResult.Yes Then

                'strDeleteCoeffSet = "DELETE * from COEFFICIENTSET WHERE NAME LIKE '" & cboCoeffSet.Text & "'"

                'modUtil.g_ADOConn.Execute(strDeleteCoeffSet)

                'MsgBox(cboCoeffSet.Text & " deleted.", MsgBoxStyle.OkOnly, "Record Deleted")

                'cboPollName.Items.Clear()
                'cboCoeffSet.Items.Clear()

                'modUtil.InitComboBox(cboPollName, "Pollutant")

                'Me.Refresh()
            End If

        Catch ex As Exception
            MsgBox("Error deleting coefficient set.", MsgBoxStyle.Critical, "Error")
            MsgBox(Err.Number & ": " & Err.Description)
        End Try
    End Sub

    Private Sub mnuCoeffImportSet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCoeffImportSet.Click
        Try
            Dim newImportCoef As New frmImportCoeffSet
            newImportCoef.ShowDialog()
        Catch ex As Exception
            HandleError(True, "mnuCoeffImportSet_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try
    End Sub

    Private Sub mnuCoeffExportSet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCoeffExportSet.Click
        Try
            'Dim intAns As Short

            ''browse...get output filename
            'dlgCMD1Open.FileName = CStr(Nothing)
            'dlgCMD1Save.FileName = CStr(Nothing)
            ''UPGRADE_WARNING: CommonDialog variable was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="671167DC-EA81-475D-B690-7A40C7BF4A23"'
            'With dlgCMD1
            '    'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            '    .Filter = Replace(MSG1, "<name>", "Coefficient Set")
            '    .Title = Replace(MSG3, "<name>", "Coefficient Set")
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
            '    ExportCoeffSet((dlgCMD1Open.FileName))
            'End If
        Catch ex As Exception
            HandleError(True, "mnuCoeffExportSet_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try

    End Sub

    Private Sub mnuPollHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPollHelp.Click
        System.Windows.Forms.Help.ShowHelp(Me, modUtil.g_nspectPath & "\Help\nspect.chm", "pollutants.htm")
    End Sub

    Private Sub mnuCoeffHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCoeffHelp.Click
        System.Windows.Forms.Help.ShowHelp(Me, modUtil.g_nspectPath & "\Help\nspect.chm", "pol_coeftab.htm")
    End Sub

#End Region


#Region "Helper Functions"

    Private Function ValidateGridValues() As Boolean
        Try
            ''Need to validate each grid value before saving.  Essentially we take it a row at a time,
            ''then rifle through each column of each row.  Case Select tests each each x,y value depending
            ''on column... 3-6 must be 1-100 range

            ''Returns: True or False

            'Dim varActive As Object 'txtActiveCell value
            'Dim i As Short
            'Dim j As Short
            'Dim iQstd As Short
            'Dim jQstd As Short

            'For i = 1 To grdPollDef.Rows - 1

            '    For j = 3 To 6

            '        'UPGRADE_WARNING: Couldn't resolve default property of object varActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            '        varActive = grdPollDef.get_TextMatrix(i, j)

            '        'UPGRADE_WARNING: Couldn't resolve default property of object varActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            '        If InStr(1, CStr(varActive), ".", CompareMethod.Text) > 0 Then
            '            'UPGRADE_WARNING: Couldn't resolve default property of object varActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            '            If (Len(Split(CStr(varActive), ".")(1)) > 4) Then
            '                ErrorGenerator(Err6, i, j)
            '                grdPollDef.col = j
            '                grdPollDef.row = i
            '                ValidateGridValues = False
            '                KeyMoveUpdate()
            '                Exit Function
            '            End If
            '        End If

            '        'UPGRADE_WARNING: Couldn't resolve default property of object varActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            '        If Not IsNumeric(varActive) Or (varActive < 0) Or (varActive > 1000) Then
            '            ErrorGenerator(Err6, i, j)
            '            grdPollDef.col = j
            '            grdPollDef.row = i
            '            ValidateGridValues = False
            '            KeyMoveUpdate()
            '            Exit Function
            '        End If



            '    Next j

            'Next i

            'For iQstd = 1 To grdWQStd.Rows - 1
            '    'UPGRADE_WARNING: Couldn't resolve default property of object varActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            '    varActive = grdWQStd.get_TextMatrix(iQstd, 3)

            '    'UPGRADE_WARNING: Couldn't resolve default property of object varActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            '    If Not IsNumeric(varActive) Or (varActive < 0) Then
            '        ErrorGenerator(Err5, iQstd, 3)
            '        grdWQStd.col = 3
            '        grdWQStd.row = iQstd
            '        ValidateGridValues = False
            '        Exit Function
            '    End If
            'Next iQstd

            'ValidateGridValues = True

        Catch ex As Exception
            'HandleError(False, "ValidateGridValues " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try
    End Function

    Private Sub UpdateValues()
        Try
            'Dim i As Short
            'Dim rsPollUpdate As New ADODB.Recordset
            'Dim strPollUpdate As String
            'Dim strWQSelect As String

            'Dim rsDescrip As New ADODB.Recordset
            'Dim rsWQstd As New ADODB.Recordset

            'If ValidateGridValues() Then
            '    'Update

            '    For i = 1 To grdPollDef.Rows - 1

            '        strPollUpdate = "SELECT * From Coefficient Where CoeffID = " & grdPollDef.get_TextMatrix(i, 7)
            '        rsPollUpdate.Open(strPollUpdate, g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            '        rsPollUpdate.Fields("Coeff1").Value = grdPollDef.get_TextMatrix(i, 3)
            '        rsPollUpdate.Fields("Coeff2").Value = grdPollDef.get_TextMatrix(i, 4)
            '        rsPollUpdate.Fields("Coeff3").Value = grdPollDef.get_TextMatrix(i, 5)
            '        rsPollUpdate.Fields("Coeff4").Value = grdPollDef.get_TextMatrix(i, 6)

            '        rsPollUpdate.Update()
            '        rsPollUpdate.Close()

            '    Next i

            'End If

            'Dim strUpdateDescription As Object
            'If boolDescChanged Then

            '    'UPGRADE_WARNING: Couldn't resolve default property of object strUpdateDescription. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            '    strUpdateDescription = "SELECT Description from CoefficientSet Where Name like '" & cboCoeffSet.Text & "'"

            '    rsDescrip.Open(strUpdateDescription, g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            '    If Len(txtCoeffSetDesc.Text) = 0 Then
            '        rsDescrip.Fields("Description").Value = ""
            '    Else
            '        rsDescrip.Fields("Description").Value = txtCoeffSetDesc.Text
            '    End If

            '    rsDescrip.Update()
            '    rsDescrip.Close()

            'End If

            'For i = 1 To grdWQStd.Rows - 1
            '    strWQSelect = "SELECT * from POLL_WQCRITERIA WHERE POLL_WQCRITID = " & grdWQStd.get_TextMatrix(i, 4)

            '    rsWQstd.Open(strWQSelect, g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            '    rsWQstd.Fields("Threshold").Value = grdWQStd.get_TextMatrix(i, 3)
            '    rsWQstd.Update()
            '    rsWQstd.Close()

            'Next i

            ''UPGRADE_NOTE: Object rsDescrip may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            'rsDescrip = Nothing
            ''UPGRADE_NOTE: Object rsPollUpdate may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            'rsPollUpdate = Nothing
            ''UPGRADE_NOTE: Object rsWQstd may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            'rsWQstd = Nothing
        Catch ex As Exception
            'HandleError(False, "UpdateValues " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try
    End Sub

    Private Sub DeletePollutant(ByRef strName As String)

        '        'Can ya guess what this one does?

        '        On Error GoTo ErrorHandler

        '        Dim strPollDelete As String
        '        Dim strPoll2Delete As String

        '        Dim rsPollDelete As ADODB.Recordset
        '        'Dim rsLCClassDelete As ADODB.Recordset

        '        strPollDelete = "Delete * FROM Pollutant WHERE NAME LIKE '" & strName & "'"


        '        modUtil.g_ADOConn.Execute(strPollDelete)

        '        MsgBox(strName & " deleted.", MsgBoxStyle.OKOnly, "Record Deleted")

        '        Me.cboPollName.Items.Clear()
        '        modUtil.InitComboBox((Me.cboPollName), "Pollutant")
        '        Me.Refresh()

        '        Exit Sub
        'ErrorHandler:
        '        HandleError(False, "DeletePollutant " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
    End Sub
#End Region


End Class