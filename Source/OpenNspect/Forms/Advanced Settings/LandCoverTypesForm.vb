'********************************************************************************************************
'File Name: frmLandCoverTypes.vb
'Description: Form for displaying and editing Land Cover Types
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
Imports System.Data
Imports System.Windows.Forms
Imports System.Drawing
Imports System.IO

Friend Class LandCoverTypesForm
    Private _intLCTypeID As Integer
    'LCTypeID#

    Private _intCount As Short
    'Number of rows in old GRID
    Private _bolBegin As Boolean

    Private _LCAdapter As OleDbDataAdapter
    Private _cBuilder As OleDbCommandBuilder
    Private _dTable As DataTable
    Private _bSource As BindingSource

    Dim WithEvents _dbConn As OleDbConnection

#Region "Events"

    Private Sub frmLandCoverTypes_Load (ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        Try
            _dbConn = g_DBConn

            IsDirty = False
            'Nothing's changed

            'Initialize the Grid and populate the combobox
            InitComboBox (cmbxLCType, "LCTYPE")
        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub

    Private Sub cboLCType_SelectedIndexChanged (ByVal eventSender As Object, ByVal eventArgs As EventArgs) _
        Handles cmbxLCType.SelectedIndexChanged
        Try
            'On the cbo Click to change to a new LandClassType, check if there's been changes, prompt to save
            Dim strSQLLCType As String
            Dim strSQLLCClass As String
            If IsDirty And _bolBegin Then

                Dim intYesNo As MsgBoxResult = MsgBox (strYesNo, MsgBoxStyle.YesNo, strYesNoTitle)
                If intYesNo = MsgBoxResult.Yes Then
                    SaveToDB()
                    IsDirty = False
                    CmdSaveEnabled()
                    cboLCType_SelectedIndexChanged (eventSender, eventArgs)
                ElseIf intYesNo = MsgBoxResult.No Then
                    IsDirty = False
                    CmdSaveEnabled()
                    cboLCType_SelectedIndexChanged (eventSender, eventArgs)
                End If

            Else
                _intCount = dgvLCTypes.RowCount
                CheckCCAPDefault (cmbxLCType.Text)

                'Selection based on combo box
                strSQLLCType = "SELECT * FROM LCTYPE WHERE NAME LIKE '" & cmbxLCType.Text & "'"

                'original
                Dim LCTypeCmd As OleDbCommand = New OleDbCommand (strSQLLCType, _dbConn)
                Dim dataLCType As OleDbDataReader = LCTypeCmd.ExecuteReader()

                If dataLCType.HasRows Then
                    dataLCType.Read()

                    txtLCTypeDesc.Text = dataLCType.Item ("Description")

                    _intLCTypeID = dataLCType.Item ("LCTypeID")

                    strSQLLCClass = "SELECT LCCLASS.Value, LCCLASS.Name, LCCLASS.[CN-A], LCCLASS.[CN-B]," & _
                                    " LCCLASS.[CN-C], LCCLASS.[CN-D], LCCLASS.CoverFactor, LCCLASS.W_WL, LCCLASS.LCTYPEID, LCCLASS.LCCLASSID FROM LCCLASS WHERE" & _
                                    " LCTYPEID = " & dataLCType.Item ("LCTypeID") & " ORDER BY LCCLass.Value"
                    Dim LCClassCmd As OleDbCommand = New OleDbCommand (strSQLLCClass, _dbConn)
                    _LCAdapter = New OleDbDataAdapter (LCClassCmd)
                    _cBuilder = New OleDbCommandBuilder (_LCAdapter)
                    _cBuilder.QuotePrefix = "["
                    _cBuilder.QuoteSuffix = "]"
                    _dTable = New DataTable

                    _LCAdapter.Fill (_dTable)

                    _dTable.Columns (0).DefaultValue = "0"
                    _dTable.Columns (1).DefaultValue = "Landclass" + (_dTable.Rows.Count + 1).ToString
                    _dTable.Columns (2).DefaultValue = "0"
                    _dTable.Columns (3).DefaultValue = "0"
                    _dTable.Columns (4).DefaultValue = "0"
                    _dTable.Columns (5).DefaultValue = "0"
                    _dTable.Columns (6).DefaultValue = "0"
                    _dTable.Columns (7).DefaultValue = "0"
                    _dTable.Columns (8).DefaultValue = _intLCTypeID
                    '_dTable.Columns(9).DefaultValue = _intLCClassID

                    _bSource = New BindingSource
                    _bSource.DataSource = _dTable
                    dgvLCTypes.DataSource = _bSource

                    If Not _bolBegin Then
                        _bolBegin = True
                    End If
                Else
                    MsgBox ("Warning: There are no records remaining.  Please add a new one.", MsgBoxStyle.Critical, _
                            "Recordset Empty")
                End If
            End If

            If Not _bolBegin Then
                _bolBegin = True
            End If
        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub

    Private Sub cmbxLCType_KeyDown (ByVal sender As Object, ByVal e As KeyEventArgs) _
        Handles cmbxLCType.KeyDown
        Try
            e.SuppressKeyPress = True
        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub

    Private Sub txtLCTypeDesc_TextChanged (ByVal sender As Object, ByVal e As EventArgs) _
        Handles txtLCTypeDesc.TextChanged
        Try
            CmdSaveEnabled()
        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub

    Protected Overrides Sub OK_Button_Click (sender As Object, e As EventArgs)

        Try
            If ValidateGridValues() Then
                SaveToDB()
                _bolBegin = False

                MsgBox ("Data saved successfully.", MsgBoxStyle.OkOnly, "Data Saved Successfully")

                IsDirty = False

                MyBase.OK_Button_Click (sender, e)
            End If
        Catch ex As Exception
            'TODO: Make sure this error functions
            If Err.Number = - 2147221504 Then
                MsgBox ( _
                        "The data values entered exceed the allowable precision of the database." & vbNewLine & _
                        "Data must not contain more than 4 values to the right of the decimal place." & vbNewLine & _
                        "Please correct your inputs before saving.", MsgBoxStyle.Information, "Precision Error")
                Exit Sub
            End If

            MsgBox ("There was an error saving changes: " + ex.Message, MsgBoxStyle.Critical, "Error Saving Changes")
        End Try
    End Sub

    Private Sub btnRestoreDefaults_Click (ByVal sender As Object, ByVal e As EventArgs) _
        Handles btnRestoreDefaults.Click
        'Restore Defaults Button - just read in NSPECT.LCCLASSDEFAULTS

        Try
            'Dim strCCAP As String

            'Check to make sure that's what they want
            intYesNo = MsgBox (strDefault, MsgBoxStyle.YesNo, strDefaultTitle)

            'Dim i As Short
            If intYesNo = MsgBoxResult.Yes Then

                'Selection based on combo box
                Dim strCCAP As String = "SELECT * From LCCLASSDEFAULTS"
                Dim cmdCCAP As New DataHelper (strCCAP)
                Dim datCCAP As OleDbDataReader = cmdCCAP.ExecuteReader()

                Dim idx As Integer = 0
                Do While datCCAP.Read()
                    _dTable.Rows (idx).Item (0) = datCCAP.Item ("Value")
                    _dTable.Rows (idx).Item (1) = datCCAP.Item ("Name")
                    _dTable.Rows (idx).Item (2) = datCCAP.Item ("CN-A")
                    _dTable.Rows (idx).Item (3) = datCCAP.Item ("CN-B")
                    _dTable.Rows (idx).Item (4) = datCCAP.Item ("CN-C")
                    _dTable.Rows (idx).Item (5) = datCCAP.Item ("CN-D")
                    _dTable.Rows (idx).Item (6) = datCCAP.Item ("CoverFactor")
                    _dTable.Rows (idx).Item (7) = datCCAP.Item ("W_WL")

                    idx = idx + 1
                Loop
                datCCAP.Close()

                IsDirty = True
                CmdSaveEnabled()
            End If
        Catch ex As Exception
            MsgBox ("There was an error loading the default CCAP data.", MsgBoxStyle.Critical, "Error Loading Data")
        End Try
    End Sub

    Private Sub mnuNewLCType_Click (ByVal sender As Object, ByVal e As EventArgs) _
        Handles mnuNewLCType.Click
        Try
            Dim newLC As New NewLandCoverTypeForm
            newLC.Init (Me)
            newLC.ShowDialog()
        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub

    Private Sub mnuDelLCType_Click (ByVal sender As Object, ByVal e As EventArgs) _
        Handles mnuDelLCType.Click
        Try
            Dim intAns As Short
            intAns = _
                MsgBox ( _
                        "Are you sure you want to delete the land cover type '" & cmbxLCType.Text & _
                        "' and all associated Coefficient Sets?", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, _
                        "Confirm Delete")

            Dim strLCTypeDelete As String
            If intAns = MsgBoxResult.Yes Then
                strLCTypeDelete = "SELECT * FROM LCTYPE WHERE NAME LIKE '" & cmbxLCType.Text & "'"
                Dim cmdLCType As New DataHelper (strLCTypeDelete)
                Dim datLC As OleDbDataReader = cmdLCType.ExecuteReader()
                datLC.Read()

                datLC.Close()
                If Not (cmbxLCType.Text = "") Then

                    'code to handle response

                    Dim cmdDel As New DataHelper ("Delete * FROM LCCLASS WHERE LCTYPEID =" & datLC ("LCTypeID"))
                    cmdDel.ExecuteNonQuery()

                    strLCTypeDelete = "Delete * FROM LCTYPE WHERE NAME LIKE '" & cmbxLCType.Text & "'"
                    Dim cmdDel2 = New OleDbCommand (strLCTypeDelete, g_DBConn)
                    cmdDel2.ExecuteNonQuery()

                    MsgBox (cmbxLCType.Text & " deleted.", MsgBoxStyle.OkOnly, "Record Deleted")

                    cmbxLCType.Items.Clear()
                    IsDirty = False
                    InitComboBox (cmbxLCType, "LCType")
                    Me.Refresh()

                Else
                    MsgBox ("Please select a Land class", MsgBoxStyle.Critical, "No Land Class Selected")
                End If
            ElseIf intAns = MsgBoxResult.No Then
                IsDirty = False
                Exit Sub
            End If
        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub

    Private Sub mnuImpLCType_Click (ByVal sender As Object, ByVal e As EventArgs) _
        Handles mnuImpLCType.Click
        Try
            Dim impLC As New ImportLandCoverTypeForm
            impLC.Init (Me)
            impLC.ShowDialog()
        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub

    Private Sub mnuExpLCType_Click (ByVal sender As Object, ByVal e As EventArgs) _
        Handles mnuExpLCType.Click
        Try
            Dim dlgsave As New SaveFileDialog
            dlgsave.Filter = Replace (MSG1TextFile, "<name>", "Land Cover Type")
            dlgsave.Title = Replace (MSG3, "<name>", "Land Cover Type")
            dlgsave.DefaultExt = ".txt"

            If dlgsave.ShowDialog = System.Windows.Forms.DialogResult.OK Then
                ExportLandCover (dlgsave.FileName)
            End If
        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub

    Private Sub mnuAppend_Click (ByVal sender As Object, ByVal e As EventArgs) Handles mnuAppend.Click
        Try
            AddRow()
        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub

    Private Sub mnuInsertRow_Click (ByVal sender As Object, ByVal e As EventArgs) _
        Handles mnuInsertRow.Click
        Try
            InsertRow()
        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub

    Private Sub mnuDeleteRow_Click (ByVal sender As Object, ByVal e As EventArgs) _
        Handles mnuDeleteRow.Click
        Try
            DeleteRow()
        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub

    Private Sub mnuLCHelp_Click (ByVal sender As Object, ByVal e As EventArgs) Handles mnuLCHelp.Click
        Try
            Help.ShowHelp (Me, g_nspectPath & "\Help\nspect.chm", "land_cover.htm")
        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub

    Private Sub dgvLCTypes_MouseClick (ByVal sender As Object, ByVal e As MouseEventArgs) _
        Handles dgvLCTypes.MouseClick
        Try
            If e.Button = System.Windows.Forms.MouseButtons.Right Then
                cntxmnuGrid.Show (dgvLCTypes, New Point (e.X, e.Y))
            End If

        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub

    Private Sub AddRowToolStripMenuItem_Click (ByVal sender As Object, ByVal e As EventArgs) _
        Handles AddRowToolStripMenuItem.Click
        Try
            AddRow()
        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub

    Private Sub InsertRowToolStripMenuItem_Click (ByVal sender As Object, ByVal e As EventArgs) _
        Handles InsertRowToolStripMenuItem.Click
        Try
            InsertRow()
        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub

    Private Sub DeleteRowToolStripMenuItem_Click (ByVal sender As Object, ByVal e As EventArgs) _
        Handles DeleteRowToolStripMenuItem.Click
        Try
            DeleteRow()
        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub

    Private Sub dgvLCTypes_CellEndEdit (ByVal sender As Object, _
                                        ByVal e As DataGridViewCellEventArgs) _
        Handles dgvLCTypes.CellEndEdit
        Try
            IsDirty = True
            CmdSaveEnabled()
        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub

    Private Sub dgvLCTypes_DataError (ByVal sender As Object, _
                                      ByVal e As DataGridViewDataErrorEventArgs) _
        Handles dgvLCTypes.DataError
        Try
            MsgBox ( _
                    "Please enter a valid number in the cell on row " + (e.RowIndex + 1).ToString + " and column " + _
                    (e.ColumnIndex + 1).ToString + ".", MsgBoxStyle.Exclamation, "Data Error")
        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub

#End Region

#Region "Helper Functions"

    Private Sub SaveToDB()
        Try
            _bSource.EndEdit()
            _LCAdapter.Update (_dTable)
        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub

    Private Sub AddRow()
        Try
            _dTable.Columns (0).DefaultValue = _dTable.Rows.Count + 1
            _dTable.Columns (1).DefaultValue = "Landclass" + (_dTable.Rows.Count + 1).ToString

            Dim dr As DataRow = _dTable.NewRow()
            _dTable.Rows.Add (dr)
            IsDirty = True
            CmdSaveEnabled()
        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub

    Private Sub InsertRow()
        Try
            If Not dgvLCTypes.CurrentRow Is Nothing Then
                _dTable.Columns (1).DefaultValue = "Landclass" + (_dTable.Rows.Count + 1).ToString
                Dim dr As DataRow = _dTable.NewRow()
                _dTable.Rows.InsertAt (dr, dgvLCTypes.CurrentRow.Index)
                IsDirty = True
                CmdSaveEnabled()
            End If
        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub

    Private Sub DeleteRow()
        Try
            If Not dgvLCTypes.CurrentRow Is Nothing Then
                dgvLCTypes.Rows.Remove (dgvLCTypes.CurrentRow)
                IsDirty = True
                CmdSaveEnabled()
            End If
        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub

    Private Sub CmdSaveEnabled()
        OK_Button.Enabled = IsDirty
    End Sub

    Private Sub CheckCCAPDefault (ByRef strName As String)
        Try
            If strName = "CCAP" Then
                btnRestoreDefaults.Enabled = True
                mnuDelLCType.Enabled = False
                mnuEdit.Enabled = False
                cntxmnuGrid.Enabled = False
                ToolTip1.SetToolTip (dgvLCTypes, "")
            Else
                btnRestoreDefaults.Enabled = False
                mnuDelLCType.Enabled = True
                mnuEdit.Enabled = True
                cntxmnuGrid.Enabled = True
                ToolTip1.SetToolTip (dgvLCTypes, "Right click to add, delete, or insert a row")
            End If

        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub

    Private Function ValidateGridValues() As Boolean
        Try
            ''Need to validate each grid value before saving.  Essentially we take it a row at a time,
            ''then rifle through each column of each row.  Case Select tests each each x,y value depending
            ''on column... eg Column 1 must be unique, 3-6 must be 1-100 range, 7 must be <= 1

            ''Returns: True or False

            Dim dr, dr2 As DataRow
            Dim val As Object

            For i As Integer = 0 To _dTable.Rows.Count - 1
                If _dTable.Rows (i).RowState <> DataRowState.Deleted Then
                    dr = _dTable.Rows (i)
                    For j As Integer = 0 To _dTable.Columns.Count - 1
                        val = dr.Item (j)
                        Select Case j
                            Case 0
                                If Not IsNumeric (val) Then
                                    DisplayError (Err1, i, j)
                                    ValidateGridValues = False
                                    Exit Function
                                Else
                                    For k As Integer = 0 To _dTable.Rows.Count - 1
                                        If _dTable.Rows (k).RowState <> DataRowState.Deleted Then
                                            dr2 = _dTable.Rows (k)
                                            If k <> i Then 'Don't want to compare value to itself
                                                If val = dr2.Item (j) Then
                                                    DisplayError (Err2, i, j)
                                                    ValidateGridValues = False
                                                    Exit Function
                                                End If
                                            End If
                                        End If
                                    Next k
                                End If

                            Case 1
                                If IsNumeric (val) Then
                                    DisplayError (Err1, i, j)
                                    ValidateGridValues = False
                                    Exit Function
                                End If

                            Case 2, 3, 4, 5
                                If Not IsNumeric (val) Or ((val < 0) Or (val > 1)) Or (Len (val.ToString) > 6) Then
                                    DisplayError (Err1, i, j)
                                    ValidateGridValues = False
                                    Exit Function
                                End If
                            Case 6
                                If Not IsNumeric (val) Or ((val < 0) Or (val > 1)) Or (Len (val.ToString) > 5) Then
                                    DisplayError (Err3, i, j)
                                    ValidateGridValues = False
                                    Exit Function
                                End If
                        End Select
                    Next
                End If
            Next

            ValidateGridValues = True
        Catch ex As Exception
            HandleError (ex)
        End Try
    End Function

    Private Sub AddDefaultValues()
        Try
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

        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub

    Private Sub ExportLandCover (ByRef strFileName As String)
        Try
            'Exports your current LCType/LCClasses to text or csv.

            Dim out As New StreamWriter (strFileName)

            'Write the name and descript.
            out.WriteLine (cmbxLCType.Text + "," + txtLCTypeDesc.Text)

            'Write name of pollutant and threshold
            Dim dr As DataRow
            For i As Integer = 0 To _dTable.Rows.Count - 1
                If _dTable.Rows (i).RowState <> DataRowState.Deleted Then
                    dr = _dTable.Rows (i)
                    out.WriteLine ( _
                                   dr (0).ToString + "," + dr (1).ToString + "," + dr (2).ToString + "," + _
                                   dr (3).ToString + "," + dr (4).ToString + "," + dr (5).ToString + "," + _
                                   dr (6).ToString + "," + dr (7).ToString)
                End If
            Next

            out.Close()

        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub

    Public Sub UpdateCombo (ByVal strName As String)
        Try
            cmbxLCType.Items.Clear()
            IsDirty = False
            InitComboBox (cmbxLCType, "LCType")
            cmbxLCType.SelectedItem = strName
        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub

#End Region
End Class