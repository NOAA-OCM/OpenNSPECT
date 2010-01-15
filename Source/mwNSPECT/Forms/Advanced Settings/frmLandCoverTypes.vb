Imports System.Data
Imports System.Data.OleDb

Imports System.Windows.Forms

Friend Class frmLandCoverTypes
    Inherits System.Windows.Forms.Form

    Private _intRow As Short 'Current Row
    Private _intCol As Short 'Current Col.
    Private _intLCTypeID As Integer 'LCTypeID#
    Private _intCount As Short 'Number of rows in old GRID
    Private _bolGridChanged As Boolean 'Flag for whether or not grid values have changed
    Private _bolSaved As Boolean 'Flag for saved/not saved changes
    Private _bolFirstLoad As Boolean 'Is initial Load event
    Private _bolBegin As Boolean
    Private _strUndoText As String 'initial cell value used to track changes - defaults back on Esc
    Private _strUndoDescrip As String 'same but for the Description
    Private _intMouseButton As Short 'Integer for mouse button click - added to avoid right click change cell value problem


    Dim WithEvents _dbConn As OleDbConnection

#Region "Events"

    Private Sub frmLandCoverTypes_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _dbConn = g_DBConn

        'Set the flags
        _bolSaved = False 'We haven't saved
        _bolGridChanged = False 'Nothing's changed
        _bolFirstLoad = True 'It's the first load

        'Initialize the Grid and populate the combobox
        modUtil.InitComboBox(cmbxLCType, "LCTYPE")
    End Sub

    Private Sub cboLCType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbxLCType.SelectedIndexChanged

        'On the cbo Click to change to a new LandClassType, check if there's been changes, prompt to save
        Dim strSQLLCType As String
        Dim strSQLLCClass As String
        If _bolGridChanged And _bolBegin Then

            Dim intYesNo As MsgBoxResult = MsgBox(strYesNo, MsgBoxStyle.YesNo, strYesNoTitle)
            'Since we're changing records for LCTypes...good time to save changes, ergo CommitTrans
            If intYesNo = MsgBoxResult.Yes Then
                UpdateValues()
                _bolGridChanged = False
                CmdSaveEnabled()
            ElseIf intYesNo = MsgBoxResult.No Then
                _bolGridChanged = False
                CmdSaveEnabled()
            End If

        Else
            _intCount = dgvLCTypes.RowCount
            CheckCCAPDefault((cmbxLCType.Text))

            'Selection based on combo box
            strSQLLCType = "SELECT * FROM LCTYPE WHERE NAME LIKE '" & cmbxLCType.Text & "'"

            'original
            Dim LCTypeCmd As OleDbCommand = New OleDbCommand(strSQLLCType, _dbConn)
            Dim dataLCType As OleDbDataReader = LCTypeCmd.ExecuteReader()

            If dataLCType.HasRows Then
                dataLCType.Read()

                txtLCTypeDesc.Text = dataLCType.Item("Description")

                strSQLLCClass = "SELECT LCCLASS.Value, LCCLASS.Name, LCCLASS.[CN-A], LCCLASS.[CN-B]," & " LCCLASS.[CN-C], LCCLASS.[CN-D], LCCLASS.CoverFactor, LCCLASS.W_WL, LCCLASS.LCTYPEID, LCCLASS.LCCLASSID FROM LCCLASS WHERE" & " LCTYPEID = " & dataLCType.Item("LCTypeID") & " ORDER BY LCCLass.Value"
                Dim LCClassCmd As OleDbCommand = New OleDbCommand(strSQLLCClass, _dbConn)
                Dim dataLCClass As OleDbDataAdapter = New OleDbDataAdapter(LCClassCmd)
                Dim dt As New System.Data.DataTable
                dataLCClass.Fill(dt)
                dgvLCTypes.DataSource = dt

                If Not _bolBegin Then
                    _bolBegin = True
                End If
            Else
                MsgBox("Warning: There are no records remaining.  Please add a new one.", MsgBoxStyle.Critical, "Recordset Empty")
            End If
        End If

        If Not _bolBegin Then
            _bolBegin = True
        End If
    End Sub

    Private Sub cmbxLCType_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cmbxLCType.KeyDown
        e.SuppressKeyPress = True
    End Sub

    Private Sub txtLCTypeDesc_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLCTypeDesc.TextChanged
        CmdSaveEnabled()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click

        If Not _bolSaved And _bolGridChanged Then

            intYesNo = MsgBox("Do you want to save changes made to " & cmbxLCType.Text & "?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Exclamation, "N-SPECT")

            If intYesNo = MsgBoxResult.Yes Then

                If ValidateGridValues() Then
                    UpdateValues()
                    'TODO: save data
                    MsgBox("Data saved successfully.", MsgBoxStyle.OkOnly, "Save Successful")
                    _bolGridChanged = False
                    _bolSaved = True
                    _bolBegin = False
                    Me.Close()
                End If

            ElseIf intYesNo = MsgBoxResult.No Then

                'TODO: roll back data
                _bolBegin = False
                Me.Close()

            Else
                Exit Sub

            End If
        Else
            'TODO: save data
            _bolBegin = False
            Me.Close()
        End If

    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try
            If ValidateGridValues() Then
                UpdateValues()
                'TODO: Save data
                _bolBegin = False

                MsgBox("Data saved successfully.", MsgBoxStyle.OkOnly, "Data Saved Successfully")

                'Reset the flags
                _bolGridChanged = False
                _bolSaved = True

                Me.Close()
            End If
        Catch ex As Exception
            'TODO: Make sure this error functions
            If Err.Number = -2147221504 Then
                MsgBox("The data values entered exceed the allowable precision of the database." & vbNewLine & "Data must not contain more than 4 values to the right of the decimal place." & vbNewLine & "Please correct your inputs before saving.", MsgBoxStyle.Information, "Precision Error")
                Exit Sub
            End If

            MsgBox("There was an error saving changes.", MsgBoxStyle.Critical, "Error Saving Changes")
        End Try
    End Sub

    Private Sub btnRestoreDefaults_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRestoreDefaults.Click
        'Restore Defaults Button - just read in NSPECT.LCCLASSDEFAULTS

        Try
            'Dim strCCAP As String

            'Check to make sure that's what they want
            intYesNo = MsgBox(strDefault, MsgBoxStyle.YesNo, strDefaultTitle)

            'Dim i As Short
            If intYesNo = MsgBoxResult.Yes Then
                'rsCCAPDefault = New ADODB.Recordset

                ''Selection based on combo box
                'strCCAP = "SELECT * From LCCLASSDEFAULTS"

                'rsCCAPDefault.Open(strCCAP, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
                'rsCCAPDefault.MoveFirst()

                'grdLCClasses.Rows = rsCCAPDefault.RecordCount + 1
                ''UPGRADE_NOTE: Refresh was upgraded to CtlRefresh. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
                'grdLCClasses.CtlRefresh()
                'm_intCount = grdLCClasses.Rows

                ''Clear out the old dataset - again, column 10 contains the LCClassID
                'For i = 1 To rsCCAPDefault.RecordCount 'grdLCClasses.Rows - 1

                '    grdLCClasses.set_TextMatrix(i, 1, rsCCAPDefault.Fields("Value").Value)
                '    grdLCClasses.set_TextMatrix(i, 2, rsCCAPDefault.Fields("Name").Value)
                '    grdLCClasses.set_TextMatrix(i, 3, rsCCAPDefault.Fields("CN-A").Value)
                '    grdLCClasses.set_TextMatrix(i, 4, rsCCAPDefault.Fields("CN-B").Value)
                '    grdLCClasses.set_TextMatrix(i, 5, rsCCAPDefault.Fields("CN-C").Value)
                '    grdLCClasses.set_TextMatrix(i, 6, rsCCAPDefault.Fields("CN-D").Value)
                '    grdLCClasses.set_TextMatrix(i, 7, rsCCAPDefault.Fields("CoverFactor").Value)
                '    grdLCClasses.set_TextMatrix(i, 8, rsCCAPDefault.Fields("W_WL").Value)

                '    rsCCAPDefault.MoveNext()

                'Next
                
                _bolGridChanged = True
                CmdSaveEnabled()
            End If
        Catch ex As Exception
            MsgBox("There was an error loading the default CCAP data.", MsgBoxStyle.Critical, "Error Loading Data")
        End Try
    End Sub



    Private Sub mnuNewLCType_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuNewLCType.Click
        Dim newLC As New frmNewLCType
        newLC.ShowDialog()
    End Sub

    Private Sub mnuDelLCType_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDelLCType.Click

        Dim intAns As Short
        intAns = MsgBox("Are you sure you want to delete the land cover type '" & cmbxLCType.Text & "' and all associated Coefficient Sets?", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Confirm Delete")

        'Dim strLCTypeDelete As String
        'Dim strLCClassDelete As String
        'Dim rsLCTypeDelete As ADODB.Recordset
        'Dim rsLCClassDelete As ADODB.Recordset

        If intAns = MsgBoxResult.Yes Then
            'strLCTypeDelete = "SELECT * FROM LCTYPE WHERE NAME LIKE '" & cboLCType.Text & "'"

            'rsLCTypeDelete = New ADODB.Recordset

            'rsLCTypeDelete.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            'rsLCTypeDelete.Open(strLCTypeDelete, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockOptimistic)

            'strLCClassDelete = "Delete * FROM LCCLASS WHERE LCTYPEID =" & rsLCTypeDelete.Fields("LCTypeID").Value

            'If Not (cboLCType.Text = "") Then

            '    'code to handle response

            '    modUtil.g_ADOConn.Execute(strLCClassDelete)

            '    'Set up a delete rs and get rid of it
            '    rsLCTypeDelete.Delete(ADODB.AffectEnum.adAffectCurrent)
            '    rsLCTypeDelete.Update()

            '    MsgBox(VB6.GetItemString(cboLCType, cboLCType.SelectedIndex) & " deleted.", MsgBoxStyle.OkOnly, "Record Deleted")

            '    cboLCType.Items.Clear()
            '    m_bolGridChanged = False
            '    modUtil.InitComboBox(cboLCType, "LCType")
            '    Me.Refresh()

            'Else
            '    MsgBox("Please select a Land class", MsgBoxStyle.Critical, "No Land Class Selected")
            'End If
        ElseIf intAns = MsgBoxResult.No Then
            _bolGridChanged = False
            Exit Sub
        End If
    End Sub

    Private Sub mnuImpLCType_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuImpLCType.Click
        Dim impLC As New frmImportLCType
        impLC.ShowDialog()
    End Sub

    Private Sub mnuExpLCType_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExpLCType.Click
        'Dim intAns As Short

        ''browse...get output filename
        'dlgCMD1Open.FileName = CStr(Nothing)
        'dlgCMD1Save.FileName = CStr(Nothing)
        ''UPGRADE_WARNING: CommonDialog variable was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="671167DC-EA81-475D-B690-7A40C7BF4A23"'
        'With dlgCMD1Open
        '    'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        '    .Filter = Replace(MSG1, "<name>", "Land Cover Type")
        '    .Title = Replace(MSG3, "<name>", "Land Cover Type")
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
        '    'Export land cover type to file - dlgCMD1.FileName
        '    ExportLandCover((dlgCMD1Open.FileName))
        'End If
    End Sub

    Private Sub mnuAppend_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAppend.Click

        'clsLCClassData.LCTypeID = getLCTypeID
        'clsLCClassData.AddNew()

        ''add row to end of  grid
        'With grdLCClasses
        '    .Rows = .Rows + 1
        '    .row = .Rows - 1
        '    .set_TextMatrix(.row, 0, "")
        '    .set_TextMatrix(.row, 1, "0")
        '    .set_TextMatrix(.row, 2, "Landclass" & .row)
        '    .set_TextMatrix(.row, 3, "0")
        '    .set_TextMatrix(.row, 4, "0")
        '    .set_TextMatrix(.row, 5, "0")
        '    .set_TextMatrix(.row, 6, "0")
        '    .set_TextMatrix(.row, 7, "0")
        '    .set_TextMatrix(.row, 8, "0")
        '    .set_TextMatrix(.row, 9, clsLCClassData.LCTypeID)
        '    .set_TextMatrix(.row, 10, g_intLCClassid)

        'End With
        ''IsCellVisible

        'm_intCount = grdLCClasses.Rows

        'CreateCheckBoxes(False, grdLCClasses.row)

        'm_bolGridChanged = True
        'CmdSaveEnabled()
    End Sub

    Private Sub mnuInsertRow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuInsertRow.Click
        ''Insert row above current row in grdLCClasses- Thanks, Andrew

        ''Get a hold of LCTYPEID, must have it to insert new records
        'clsLCClassData.LCTypeID = m_intLCTypeID
        'clsLCClassData.AddNew()

        'Dim row, R, col As Short
        'With grdLCClasses
        '    If .row < .FixedRows Then 'make sure we don't insert above header Rows
        '        mnuAppend_Click(mnuAppend, New System.EventArgs())
        '    Else
        '        R = .row
        '        .Rows = .Rows + 1 'add a row

        '        For row = .Rows - 1 To R + 1 Step -1 'move data dn 1 row
        '            For col = 1 To .get_Cols() - 1
        '                .set_TextMatrix(row, col, .get_TextMatrix(row - 1, col))
        '            Next col
        '        Next row
        '        For col = 1 To .get_Cols() - 1 ' clear all cells in this row
        '            If (col = 2) Then
        '                .set_TextMatrix(R, col, "")
        '            Else
        '                .set_TextMatrix(R, col, "0")
        '            End If
        '        Next col
        '        .set_TextMatrix(R, 9, clsLCClassData.LCTypeID)
        '        .set_TextMatrix(R, 10, g_intLCClassid)


        '    End If
        'End With

        'txtActiveCell.Visible = False
        'm_intCount = grdLCClasses.Rows

        'ClearCheckBoxes(True, m_intCount - 1)
        'CreateCheckBoxes(True)

        'm_bolGridChanged = True 'reset
        'CmdSaveEnabled()
    End Sub

    Private Sub mnuDeleteRow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDeleteRow.Click
        ''delete current row

        'Dim row, R, C, col As Short
        'Dim lngLCClassID As Integer

        'With grdLCClasses

        '    lngLCClassID = CInt(.get_TextMatrix(m_intRow, 10))

        '    If .Rows > .FixedRows Then 'make sure we don't del header Rows
        '        For col = 1 To .get_Cols() - 1
        '            If ((Trim(.get_TextMatrix(.row, col)) > "" And col = 2) Or (.get_TextMatrix(.row, col) <> "0" And col = 1) Or (.get_TextMatrix(.row, col) <> "0" And col >= 3)) Then 'data?
        '                C = 1
        '                Exit For
        '            End If
        '        Next col
        '        If C Then
        '            R = MsgBox("There is data in Row" & Str(.row) & " ! Delete anyway?", MsgBoxStyle.YesNo, "Delete Row!")
        '        End If
        '        If C = 0 Or R = MsgBoxResult.Yes Then 'no exist. data or YES
        '            If .row = .Rows - 1 Then 'last row?
        '                .row = .row - 1 'move active cell
        '            Else
        '                For row = .row To .Rows - 2 'move data up 1 row
        '                    For col = 1 To .get_Cols() - 1
        '                        .set_TextMatrix(row, col, .get_TextMatrix(row + 1, col))
        '                    Next col
        '                Next row
        '            End If
        '            .Rows = .Rows - 1 'del last row
        '            clsLCClassData.Load(lngLCClassID)
        '            clsLCClassData.Delete()
        '        End If
        '    End If
        'End With

        'm_intCount = grdLCClasses.Rows
        'ClearCheckBoxes(True, m_intCount + 1)
        'CreateCheckBoxes(True)

        'm_bolGridChanged = True 'reset
        'CmdSaveEnabled()
    End Sub

    Private Sub mnuLCHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuLCHelp.Click
        Help.ShowHelp(Me, modUtil.g_nspectPath & "\Help\nspect.chm", "land_cover.htm")
    End Sub

#End Region

#Region "Helper Functions"
    Private Sub CmdSaveEnabled()

        btnSave.Enabled = _bolGridChanged

    End Sub

    Private Sub CheckCCAPDefault(ByRef strName As String)

        If strName = "CCAP" Then
            btnRestoreDefaults.Enabled = True
            mnuDelLCType.Enabled = False
            mnuPopUp.Enabled = False
            ToolTip1.SetToolTip(dgvLCTypes, "")
        Else
            btnRestoreDefaults.Enabled = False
            mnuDelLCType.Enabled = True
            mnuPopUp.Enabled = True
            ToolTip1.SetToolTip(dgvLCTypes, "Right click to add, delete, or insert a row")
        End If

    End Sub

    Private Sub UpdateValues()

        'Dim i As Short

        'For i = 1 To grdLCClasses.Rows - 1

        '    With clsLCClassData

        '        .Load(grdLCClasses.get_TextMatrix(i, 10))
        '        .Value = CInt(grdLCClasses.get_TextMatrix(i, 1))
        '        .Name = grdLCClasses.get_TextMatrix(i, 2)
        '        .CNA = CSng(grdLCClasses.get_TextMatrix(i, 3))
        '        .CNB = CSng(grdLCClasses.get_TextMatrix(i, 4))
        '        .CNC = CSng(grdLCClasses.get_TextMatrix(i, 5))
        '        .CND = CSng(grdLCClasses.get_TextMatrix(i, 6))
        '        .CoverFactor = CSng(grdLCClasses.get_TextMatrix(i, 7))
        '        .W_WL = CInt(grdLCClasses.get_TextMatrix(i, 8))
        '        .LCTypeID = CInt(grdLCClasses.get_TextMatrix(i, 9))
        '        .LCClassID = CInt(grdLCClasses.get_TextMatrix(i, 10))

        '        .SaveChanges()

        '    End With
        'Next i

    End Sub

    Private Function ValidateGridValues() As Boolean

        ''Need to validate each grid value before saving.  Essentially we take it a row at a time,
        ''then rifle through each column of each row.  Case Select tests each each x,y value depending
        ''on column... eg Column 1 must be unique, 3-6 must be 1-100 range, 7 must be <= 1

        ''Returns: True or False

        'Dim varActive As Object 'txtActiveCell value
        'Dim varColumn2Value As Object 'Value of Column 2 ([VALUE]) - have to check for unique
        'Dim i As Short
        'Dim j As Short
        'Dim k As Short

        'For i = 1 To grdLCClasses.Rows - 1

        '    For j = 1 To 7

        '        'UPGRADE_WARNING: Couldn't resolve default property of object varActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        varActive = grdLCClasses.get_TextMatrix(i, j)

        '        Select Case j

        '            Case 1
        '                If Not IsNumeric(varActive) Then
        '                    ErrorGenerator(Err1, i, j)
        '                Else
        '                    For k = 1 To grdLCClasses.Rows - 1

        '                        'UPGRADE_WARNING: Couldn't resolve default property of object varColumn2Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                        varColumn2Value = grdLCClasses.get_TextMatrix(k, 1)
        '                        If k <> i Then 'Don't want to compare value to itself
        '                            'UPGRADE_WARNING: Couldn't resolve default property of object varColumn2Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                            If varColumn2Value = grdLCClasses.get_TextMatrix(i, 1) Then
        '                                ErrorGenerator(Err2, i, j)
        '                                grdLCClasses.col = j
        '                                grdLCClasses.row = i
        '                                ValidateGridValues = False
        '                                KeyMoveUpdate()
        '                                Exit Function
        '                            End If
        '                        End If
        '                    Next k
        '                End If


        '            Case 2
        '                If IsNumeric(varActive) Then
        '                    ErrorGenerator(Err1, i, j)
        '                    grdLCClasses.col = j
        '                    grdLCClasses.row = i
        '                    ValidateGridValues = False
        '                    KeyMoveUpdate()
        '                    Exit Function
        '                End If

        '            Case 3
        '                'UPGRADE_WARNING: Couldn't resolve default property of object varActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                If Not IsNumeric(varActive) Or ((varActive < 0) Or (varActive > 1)) Or (Len(varActive) > 6) Then
        '                    ErrorGenerator(Err1, i, j)
        '                    grdLCClasses.col = j
        '                    grdLCClasses.row = i
        '                    ValidateGridValues = False
        '                    KeyMoveUpdate()
        '                    Exit Function
        '                End If

        '            Case 4
        '                'UPGRADE_WARNING: Couldn't resolve default property of object varActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                If Not IsNumeric(varActive) Or ((varActive < 0) Or (varActive > 1)) Or (Len(varActive) > 6) Then
        '                    ErrorGenerator(Err1, i, j)
        '                    grdLCClasses.col = j
        '                    grdLCClasses.row = i
        '                    ValidateGridValues = False
        '                    KeyMoveUpdate()
        '                    Exit Function
        '                End If

        '            Case 5
        '                'UPGRADE_WARNING: Couldn't resolve default property of object varActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                If Not IsNumeric(varActive) Or ((varActive < 0) Or (varActive > 1)) Or (Len(varActive) > 6) Then
        '                    ErrorGenerator(Err1, i, j)
        '                    grdLCClasses.col = j
        '                    grdLCClasses.row = i
        '                    ValidateGridValues = False
        '                    KeyMoveUpdate()
        '                    Exit Function
        '                End If

        '            Case 6
        '                'UPGRADE_WARNING: Couldn't resolve default property of object varActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                If Not IsNumeric(varActive) Or ((varActive < 0) Or (varActive > 1)) Or (Len(varActive) > 6) Then
        '                    ErrorGenerator(Err1, i, j)
        '                    grdLCClasses.col = j
        '                    grdLCClasses.row = i
        '                    ValidateGridValues = False
        '                    KeyMoveUpdate()
        '                    Exit Function
        '                End If

        '            Case 7
        '                'UPGRADE_WARNING: Couldn't resolve default property of object varActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                If Not IsNumeric(varActive) Or ((varActive < 0) Or (varActive > 1)) Or (Len(varActive) > 5) Then
        '                    ErrorGenerator(Err3, i, j)
        '                    grdLCClasses.col = j
        '                    grdLCClasses.row = i
        '                    ValidateGridValues = False
        '                    KeyMoveUpdate()
        '                    Exit Function
        '                End If
        '        End Select
        '    Next j
        'Next i

        ValidateGridValues = True

    End Function

    Private Sub AddDefaultValues()

        'Dim i As Short

        'For i = 1 To grdLCClasses.Rows - 1

        '    With clsLCClassData

        '        .Value = CInt(grdLCClasses.get_TextMatrix(i, 1))
        '        .Name = grdLCClasses.get_TextMatrix(i, 2)
        '        .CNA = CSng(grdLCClasses.get_TextMatrix(i, 3))
        '        .CNB = CSng(grdLCClasses.get_TextMatrix(i, 4))
        '        .CNC = CSng(grdLCClasses.get_TextMatrix(i, 5))
        '        .CND = CSng(grdLCClasses.get_TextMatrix(i, 6))
        '        .CoverFactor = CSng(grdLCClasses.get_TextMatrix(i, 7))
        '        .W_WL = CInt(grdLCClasses.get_TextMatrix(i, 8))
        '        .LCTypeID = CInt(grdLCClasses.get_TextMatrix(i, 9))

        '        .AddNew()

        '    End With
        'Next i

        'ClearCheckBoxes(True, m_intCount)
        'CreateCheckBoxes(True)


    End Sub

    Private Sub ExportLandCover(ByRef strFileName As String)
        ''Exports your current LCType/LCClasses to text or csv.

        'Dim fso As New Scripting.FileSystemObject
        'Dim fl As Scripting.TextStream
        'Dim rsNew As ADODB.Recordset
        'Dim theLine As Object

        'fl = fso.CreateTextFile(strFileName, True)

        ''Write the name and descript.
        'With fl
        '    .WriteLine(cboLCType.Text & "," & txtLCTypeDesc.Text)
        'End With

        'Dim i As Short

        ''Write name of pollutant and threshold
        'For i = 1 To grdLCClasses.Rows - 1
        '    fl.WriteLine(grdLCClasses.get_TextMatrix(i, 1) & "," & grdLCClasses.get_TextMatrix(i, 2) & "," & grdLCClasses.get_TextMatrix(i, 3) & "," & grdLCClasses.get_TextMatrix(i, 4) & "," & grdLCClasses.get_TextMatrix(i, 5) & "," & grdLCClasses.get_TextMatrix(i, 6) & "," & grdLCClasses.get_TextMatrix(i, 7) & "," & grdLCClasses.get_TextMatrix(i, 8))

        'Next i

        'fl.Close()

        ''UPGRADE_NOTE: Object fso may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        'fso = Nothing
        ''UPGRADE_NOTE: Object fl may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        'fl = Nothing

    End Sub

#End Region

    
End Class