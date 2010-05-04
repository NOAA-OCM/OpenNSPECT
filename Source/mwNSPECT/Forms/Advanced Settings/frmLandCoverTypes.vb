Imports System.Data
Imports System.Data.OleDb
Imports System.Windows.Forms

Friend Class frmLandCoverTypes
    Inherits System.Windows.Forms.Form

    Private _intRow As Short 'Current Row
    Private _intCol As Short 'Current Col.
    Private _intLCTypeID As Integer 'LCTypeID#
    Private _intLCClassID As Integer
    Private _intCount As Short 'Number of rows in old GRID
    Private _bolGridChanged As Boolean 'Flag for whether or not grid values have changed
    Private _bolSaved As Boolean 'Flag for saved/not saved changes
    Private _bolFirstLoad As Boolean 'Is initial Load event
    Private _bolBegin As Boolean
    Private _strUndoText As String 'initial cell value used to track changes - defaults back on Esc
    Private _strUndoDescrip As String 'same but for the Description
    Private _intMouseButton As Short 'Integer for mouse button click - added to avoid right click change cell value problem

    Private _LCAdapter As OleDbDataAdapter
    Private _cBuilder As OleDbCommandBuilder
    Private _dTable As DataTable
    Private _bSource As BindingSource

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
            If intYesNo = MsgBoxResult.Yes Then
                SaveToDB()
                _bolGridChanged = False
                CmdSaveEnabled()
                cboLCType_SelectedIndexChanged(eventSender, eventArgs)
            ElseIf intYesNo = MsgBoxResult.No Then
                _bolGridChanged = False
                CmdSaveEnabled()
                cboLCType_SelectedIndexChanged(eventSender, eventArgs)
            End If

        Else
            _intCount = dgvLCTypes.RowCount
            CheckCCAPDefault(cmbxLCType.Text)

            'Selection based on combo box
            strSQLLCType = "SELECT * FROM LCTYPE WHERE NAME LIKE '" & cmbxLCType.Text & "'"

            'original
            Dim LCTypeCmd As OleDbCommand = New OleDbCommand(strSQLLCType, _dbConn)
            Dim dataLCType As OleDbDataReader = LCTypeCmd.ExecuteReader()

            If dataLCType.HasRows Then
                dataLCType.Read()

                txtLCTypeDesc.Text = dataLCType.Item("Description")

                _intLCTypeID = dataLCType.Item("LCTypeID")

                strSQLLCClass = "SELECT LCCLASS.Value, LCCLASS.Name, LCCLASS.[CN-A], LCCLASS.[CN-B]," & " LCCLASS.[CN-C], LCCLASS.[CN-D], LCCLASS.CoverFactor, LCCLASS.W_WL, LCCLASS.LCTYPEID, LCCLASS.LCCLASSID FROM LCCLASS WHERE" & " LCTYPEID = " & dataLCType.Item("LCTypeID") & " ORDER BY LCCLass.Value"
                Dim LCClassCmd As OleDbCommand = New OleDbCommand(strSQLLCClass, _dbConn)
                _LCAdapter = New OleDbDataAdapter(LCClassCmd)
                _cBuilder = New OleDbCommandBuilder(_LCAdapter)
                _cBuilder.QuotePrefix = "["
                _cBuilder.QuoteSuffix = "]"
                _dTable = New DataTable

                _LCAdapter.Fill(_dTable)

                _dTable.Columns(0).DefaultValue = "0"
                _dTable.Columns(1).DefaultValue = "Landclass" + (_dTable.Rows.Count + 1).ToString
                _dTable.Columns(2).DefaultValue = "0"
                _dTable.Columns(3).DefaultValue = "0"
                _dTable.Columns(4).DefaultValue = "0"
                _dTable.Columns(5).DefaultValue = "0"
                _dTable.Columns(6).DefaultValue = "0"
                _dTable.Columns(7).DefaultValue = "0"
                _dTable.Columns(8).DefaultValue = _intLCTypeID
                '_dTable.Columns(9).DefaultValue = _intLCClassID


                _bSource = New BindingSource
                _bSource.DataSource = _dTable
                dgvLCTypes.DataSource = _bSource

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
                    SaveToDB()
                    MsgBox("Data saved successfully.", MsgBoxStyle.OkOnly, "Save Successful")
                    _bolGridChanged = False
                    _bolSaved = True
                    _bolBegin = False
                    Me.Close()
                End If

            ElseIf intYesNo = MsgBoxResult.No Then
                _bolBegin = False
                Me.Close()

            Else
                Exit Sub

            End If
        Else
            SaveToDB()
            MsgBox("Data saved successfully.", MsgBoxStyle.OkOnly, "Save Successful")
            _bolGridChanged = False
            _bolSaved = True
            _bolBegin = False
            Me.Close()
        End If

    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try
            If ValidateGridValues() Then
                SaveToDB()
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

            MsgBox("There was an error saving changes: " + ex.Message, MsgBoxStyle.Critical, "Error Saving Changes")
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
                
                'Selection based on combo box
                Dim strCCAP As String = "SELECT * From LCCLASSDEFAULTS"
                Dim cmdCCAP As New OleDbCommand(strCCAP, g_DBConn)
                Dim datCCAP As OleDbDataReader = cmdCCAP.ExecuteReader()

                Dim idx As Integer = 0
                Do While datCCAP.Read()
                    _dTable.Rows(idx).Item(0) = datCCAP.Item("Value")
                    _dTable.Rows(idx).Item(1) = datCCAP.Item("Name")
                    _dTable.Rows(idx).Item(2) = datCCAP.Item("CN-A")
                    _dTable.Rows(idx).Item(3) = datCCAP.Item("CN-B")
                    _dTable.Rows(idx).Item(4) = datCCAP.Item("CN-C")
                    _dTable.Rows(idx).Item(5) = datCCAP.Item("CN-D")
                    _dTable.Rows(idx).Item(6) = datCCAP.Item("CoverFactor")
                    _dTable.Rows(idx).Item(7) = datCCAP.Item("W_WL")

                    idx = idx + 1
                Loop
                datCCAP.Close()

                _bolGridChanged = True
                CmdSaveEnabled()
            End If
        Catch ex As Exception
            MsgBox("There was an error loading the default CCAP data.", MsgBoxStyle.Critical, "Error Loading Data")
        End Try
    End Sub



    Private Sub mnuNewLCType_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuNewLCType.Click
        Dim newLC As New frmNewLCType
        newLC.Init(Me)
        newLC.ShowDialog()
    End Sub

    Private Sub mnuDelLCType_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDelLCType.Click

        Dim intAns As Short
        intAns = MsgBox("Are you sure you want to delete the land cover type '" & cmbxLCType.Text & "' and all associated Coefficient Sets?", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Confirm Delete")

        Dim strLCTypeDelete As String
        Dim strLCClassDelete As String

        If intAns = MsgBoxResult.Yes Then
            strLCTypeDelete = "SELECT * FROM LCTYPE WHERE NAME LIKE '" & cmbxLCType.Text & "'"
            Dim cmdLCType As New OleDbCommand(strLCTypeDelete, g_DBConn)
            Dim datLC As OleDbDataReader = cmdLCType.ExecuteReader()
            datLC.Read()

            strLCClassDelete = "Delete * FROM LCCLASS WHERE LCTYPEID =" & datLC("LCTypeID")

            datLC.Close()
            If Not (cmbxLCType.Text = "") Then

                'code to handle response

                Dim cmdDel As New OleDbCommand(strLCClassDelete, g_DBConn)
                cmdDel.ExecuteNonQuery()

                strLCTypeDelete = "Delete * FROM LCTYPE WHERE NAME LIKE '" & cmbxLCType.Text & "'"
                cmdDel = New OleDbCommand(strLCTypeDelete, g_DBConn)
                cmdDel.ExecuteNonQuery()

                MsgBox(cmbxLCType.Text & " deleted.", MsgBoxStyle.OkOnly, "Record Deleted")

                cmbxLCType.Items.Clear()
                _bolGridChanged = False
                modUtil.InitComboBox(cmbxLCType, "LCType")
                Me.Refresh()

            Else
                MsgBox("Please select a Land class", MsgBoxStyle.Critical, "No Land Class Selected")
            End If
        ElseIf intAns = MsgBoxResult.No Then
            _bolGridChanged = False
            Exit Sub
        End If
    End Sub

    Private Sub mnuImpLCType_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuImpLCType.Click
        Dim impLC As New frmImportLCType
        impLC.Init(Me)
        impLC.ShowDialog()
    End Sub

    Private Sub mnuExpLCType_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExpLCType.Click
        Dim dlgsave As New SaveFileDialog
        dlgsave.Filter = Replace(MSG1, "<name>", "Land Cover Type")
        dlgsave.Title = Replace(MSG3, "<name>", "Land Cover Type")
        dlgsave.DefaultExt = ".txt"

        If dlgsave.ShowDialog = Windows.Forms.DialogResult.OK Then
            ExportLandCover(dlgsave.FileName)
        End If
    End Sub

    Private Sub mnuAppend_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAppend.Click
        AddRow()
    End Sub

    Private Sub mnuInsertRow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuInsertRow.Click
        InsertRow()
    End Sub

    Private Sub mnuDeleteRow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDeleteRow.Click
        DeleteRow()
    End Sub

    Private Sub mnuLCHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuLCHelp.Click
        Help.ShowHelp(Me, modUtil.g_nspectPath & "\Help\nspect.chm", "land_cover.htm")
    End Sub

    Private Sub dgvLCTypes_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgvLCTypes.MouseClick
        If e.Button = Windows.Forms.MouseButtons.Right Then
            cntxmnuGrid.Show(dgvLCTypes, New Drawing.Point(e.X, e.Y))
        End If

    End Sub

    Private Sub AddRowToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddRowToolStripMenuItem.Click
        AddRow()
    End Sub

    Private Sub InsertRowToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InsertRowToolStripMenuItem.Click
        InsertRow()
    End Sub

    Private Sub DeleteRowToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteRowToolStripMenuItem.Click
        DeleteRow()
    End Sub


    Private Sub dgvLCTypes_CellEndEdit(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvLCTypes.CellEndEdit
        _bolGridChanged = True
        CmdSaveEnabled()
    End Sub

    Private Sub dgvLCTypes_DataError(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvLCTypes.DataError
        MsgBox("Please enter a valid number in the cell on row " + (e.RowIndex + 1).ToString + " and column " + (e.ColumnIndex + 1).ToString + ".", MsgBoxStyle.Exclamation, "Data Error")
    End Sub
#End Region

#Region "Helper Functions"
    Private Sub SaveToDB()
        _bSource.EndEdit()
        _LCAdapter.Update(_dTable)
    End Sub

    Private Sub AddRow()
        _dTable.Columns(0).DefaultValue = _dTable.Rows.Count + 1
        _dTable.Columns(1).DefaultValue = "Landclass" + (_dTable.Rows.Count + 1).ToString

        Dim dr As DataRow = _dTable.NewRow()
        _dTable.Rows.Add(dr)
        _bolGridChanged = True
        CmdSaveEnabled()
    End Sub

    Private Sub InsertRow()
        If Not dgvLCTypes.CurrentRow Is Nothing Then
            _dTable.Columns(1).DefaultValue = "Landclass" + (_dTable.Rows.Count + 1).ToString
            Dim dr As DataRow = _dTable.NewRow()
            _dTable.Rows.InsertAt(dr, dgvLCTypes.CurrentRow.Index)
            _bolGridChanged = True
            CmdSaveEnabled()
        End If
    End Sub

    Private Sub DeleteRow()
        If Not dgvLCTypes.CurrentRow Is Nothing Then
            dgvLCTypes.Rows.Remove(dgvLCTypes.CurrentRow)
            _bolGridChanged = True
            CmdSaveEnabled()
        End If
    End Sub

    Private Sub CmdSaveEnabled()

        btnSave.Enabled = _bolGridChanged

    End Sub

    Private Sub CheckCCAPDefault(ByRef strName As String)

        If strName = "CCAP" Then
            btnRestoreDefaults.Enabled = True
            mnuDelLCType.Enabled = False
            mnuEdit.Enabled = False
            cntxmnuGrid.Enabled = False
            ToolTip1.SetToolTip(dgvLCTypes, "")
        Else
            btnRestoreDefaults.Enabled = False
            mnuDelLCType.Enabled = True
            mnuEdit.Enabled = True
            cntxmnuGrid.Enabled = True
            ToolTip1.SetToolTip(dgvLCTypes, "Right click to add, delete, or insert a row")
        End If

    End Sub


    Private Function ValidateGridValues() As Boolean

        ''Need to validate each grid value before saving.  Essentially we take it a row at a time,
        ''then rifle through each column of each row.  Case Select tests each each x,y value depending
        ''on column... eg Column 1 must be unique, 3-6 must be 1-100 range, 7 must be <= 1

        ''Returns: True or False

        Dim dr, dr2 As DataRow
        Dim val As Object

        For i As Integer = 0 To _dTable.Rows.Count - 1
            If _dTable.Rows(i).RowState <> DataRowState.Deleted Then
                dr = _dTable.Rows(i)
                For j As Integer = 0 To _dTable.Columns.Count - 1
                    val = dr.Item(j)
                    Select Case j
                        Case 0
                            If Not IsNumeric(val) Then
                                ErrorGenerator(Err1, i, j)
                                ValidateGridValues = False
                                Exit Function
                            Else
                                For k As Integer = 0 To _dTable.Rows.Count - 1
                                    If _dTable.Rows(k).RowState <> DataRowState.Deleted Then
                                        dr2 = _dTable.Rows(k)
                                        If k <> i Then 'Don't want to compare value to itself
                                            If val = dr2.Item(j) Then
                                                ErrorGenerator(Err2, i, j)
                                                ValidateGridValues = False
                                                Exit Function
                                            End If
                                        End If
                                    End If
                                Next k
                            End If

                        Case 1
                            If IsNumeric(val) Then
                                ErrorGenerator(Err1, i, j)
                                ValidateGridValues = False
                                Exit Function
                            End If

                        Case 2, 3, 4, 5
                            If Not IsNumeric(val) Or ((val < 0) Or (val > 1)) Or (Len(val.ToString) > 6) Then
                                ErrorGenerator(Err1, i, j)
                                ValidateGridValues = False
                                Exit Function
                            End If
                        Case 6
                            If Not IsNumeric(val) Or ((val < 0) Or (val > 1)) Or (Len(val.ToString) > 5) Then
                                ErrorGenerator(Err3, i, j)
                                ValidateGridValues = False
                                Exit Function
                            End If
                    End Select
                Next
            End If
        Next

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
        'Exports your current LCType/LCClasses to text or csv.

        Dim out As New IO.StreamWriter(strFileName)

        'Write the name and descript.
        out.WriteLine(cmbxLCType.Text + "," + txtLCTypeDesc.Text)

        'Write name of pollutant and threshold
        Dim dr As DataRow
        For i As Integer = 0 To _dTable.Rows.Count - 1
            If _dTable.Rows(i).RowState <> DataRowState.Deleted Then
                dr = _dTable.Rows(i)
                out.WriteLine(dr(0).ToString + "," + dr(1).ToString + "," + dr(2).ToString + "," + dr(3).ToString + "," + dr(4).ToString + "," + dr(5).ToString + "," + dr(6).ToString + "," + dr(7).ToString)
            End If
        Next

        out.Close()

    End Sub


    Public Sub UpdateCombo(ByVal strName As String)
        cmbxLCType.Items.Clear()
        _bolGridChanged = False
        modUtil.InitComboBox(cmbxLCType, "LCType")
        cmbxLCType.SelectedItem = strName
    End Sub

#End Region

    
End Class