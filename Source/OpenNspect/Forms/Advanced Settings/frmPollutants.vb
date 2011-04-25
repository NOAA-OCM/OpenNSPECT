'********************************************************************************************************
'File Name: frmPollutants.vb
'Description: Form for displaying and editting pollutants
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

Imports System.Data
Imports System.Data.OleDb
Friend Class frmPollutants
    Inherits System.Windows.Forms.Form

    Const c_sModuleFileName As String = "frmPollutants.vb"

    Private _boolLoaded As Boolean
    Private _boolChanged As Boolean 'Bool for enabling save button
    Private _boolDescChanged As Boolean 'Boolship for seeing if Description Changed
    Private _boolSaved As Boolean 'Boolship for whether or not things have saved

    Private _coefCmd As OleDbCommand

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

    Private _coefs As OleDbDataAdapter
    Private _dtCoeff As Data.DataTable

    Private _wq As OleDbDataAdapter
    Private _dtWQ As Data.DataTable


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
            HandleError(c_sModuleFileName, ex)
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
                    _coefCmd = New OleDbCommand(strSQLCoeff, modUtil.g_DBConn)
                    Dim coef As OleDbDataReader = _coefCmd.ExecuteReader()
                    coef.Read()

                    strSQLLCType = "SELECT NAME, LCTYPEID FROM LCTYPE WHERE LCTYPEID = " & coef.Item("LCTypeID") & ""
                    Dim LCCmd As New OleDbCommand(strSQLLCType, modUtil.g_DBConn)
                    Dim LC As OleDbDataReader = LCCmd.ExecuteReader()
                    LC.Read()
                    _strLCType = LC.Item("Name")
                    _intLCTypeID = LC.Item("LCTypeID")
                    LC.Close()

                    'Fill everything based on that
                    cboCoeffSet.Items.Clear()
                    cboCoeffSet.Items.Add(coef("Name"))
                    While coef.Read()
                        cboCoeffSet.Items.Add(coef("Name"))
                    End While
                    cboCoeffSet.SelectedIndex = 0
                    coef.Close()

                    txtLCType.Text = _strLCType

                    'Fill the Water Quality Standards Tab
                    strSQLWQStd = "SELECT WQCRITERIA.Name, WQCRITERIA.Description," & "POLL_WQCRITERIA.Threshold, POLL_WQCRITERIA.POLL_WQCRITID FROM WQCRITERIA " & "LEFT OUTER JOIN POLL_WQCRITERIA ON WQCRITERIA.WQCRITID = POLL_WQCRITERIA.WQCRITID " & "Where POLL_WQCRITERIA.POLLID = " & poll.Item("POLLID")
                    Dim wqCmd As New OleDbCommand(strSQLWQStd, modUtil.g_DBConn)
                    _wq = New OleDbDataAdapter(wqCmd)
                    _dtWQ = New Data.DataTable
                    _wq.Fill(_dtWQ)
                    dgvWaterQuality.DataSource = _dtWQ
                    poll.Close()

                ElseIf intYesNo = MsgBoxResult.No Then

                    _boolChanged = False

                    'Selection based on combo box
                    strSQLPollutant = "SELECT * FROM POLLUTANT WHERE NAME = '" & cboPollName.Text & "'"
                    Dim pollCmd As New OleDbCommand(strSQLPollutant, modUtil.g_DBConn)
                    Dim poll As OleDbDataReader = pollCmd.ExecuteReader()
                    poll.Read()
                    _intPollID = poll.Item("POLLID")

                    strSQLCoeff = "SELECT * FROM CoefficientSet WHERE POLLID = " & poll.Item("POLLID") & ""
                    _coefCmd = New OleDbCommand(strSQLCoeff, modUtil.g_DBConn)
                    Dim coef As OleDbDataReader = _coefCmd.ExecuteReader()
                    coef.Read()

                    strSQLLCType = "SELECT NAME, LCTYPEID FROM LCTYPE WHERE LCTYPEID = " & coef.Item("LCTypeID") & ""
                    Dim LCCmd As New OleDbCommand(strSQLLCType, modUtil.g_DBConn)
                    Dim LC As OleDbDataReader = LCCmd.ExecuteReader()
                    LC.Read()
                    _strLCType = LC.Item("Name")
                    _intLCTypeID = LC.Item("LCTypeID")
                    LC.Close()

                    'Fill everything based on that
                    cboCoeffSet.Items.Clear()
                    cboCoeffSet.Items.Add(coef.Item("Name"))
                    While coef.Read()
                        cboCoeffSet.Items.Add(coef.Item("Name"))
                    End While
                    cboCoeffSet.SelectedIndex = 0
                    coef.Close()

                    txtLCType.Text = _strLCType

                    'Fill the Water Quality Standards Tab
                    strSQLWQStd = "SELECT WQCRITERIA.Name, WQCRITERIA.Description," & "POLL_WQCRITERIA.Threshold, POLL_WQCRITERIA.POLL_WQCRITID FROM WQCRITERIA " & "LEFT OUTER JOIN POLL_WQCRITERIA ON WQCRITERIA.WQCRITID = POLL_WQCRITERIA.WQCRITID " & "Where POLL_WQCRITERIA.POLLID = " & poll.Item("POLLID")
                    Dim wqCmd As New OleDbCommand(strSQLWQStd, modUtil.g_DBConn)
                    _wq = New OleDbDataAdapter(wqCmd)
                    _dtWQ = New Data.DataTable
                    _wq.Fill(_dtWQ)
                    dgvWaterQuality.DataSource = _dtWQ
                    poll.Close()

                End If
            Else

                _boolChanged = False

                'Selection based on combo box
                strSQLPollutant = "SELECT * FROM POLLUTANT WHERE NAME = '" & cboPollName.Text & "'"
                Dim pollCmd As New OleDbCommand(strSQLPollutant, modUtil.g_DBConn)
                Dim poll As OleDbDataReader = pollCmd.ExecuteReader()
                poll.Read()
                _intPollID = poll.Item("POLLID")

                strSQLCoeff = "SELECT * FROM CoefficientSet WHERE POLLID = " & poll.Item("POLLID") & ""
                _coefCmd = New OleDbCommand(strSQLCoeff, modUtil.g_DBConn)
                Dim coef As OleDbDataReader = _coefCmd.ExecuteReader()
                coef.Read()

                strSQLLCType = "SELECT NAME, LCTYPEID FROM LCTYPE WHERE LCTYPEID = " & coef.Item("LCTypeID") & ""
                Dim LCCmd As New OleDbCommand(strSQLLCType, modUtil.g_DBConn)
                Dim LC As OleDbDataReader = LCCmd.ExecuteReader()
                LC.Read()
                _strLCType = LC.Item("Name")
                _intLCTypeID = LC.Item("LCTypeID")
                LC.Close()

                'Fill everything based on that
                cboCoeffSet.Items.Clear()
                cboCoeffSet.Items.Add(coef.Item("Name"))
                While coef.Read()
                    cboCoeffSet.Items.Add(coef.Item("Name"))
                End While
                cboCoeffSet.SelectedIndex = 0
                coef.Close()

                txtLCType.Text = _strLCType

                'Fill the Water Quality Standards Tab
                strSQLWQStd = "SELECT WQCRITERIA.Name, WQCRITERIA.Description," & "POLL_WQCRITERIA.Threshold, POLL_WQCRITERIA.POLL_WQCRITID FROM WQCRITERIA " & "LEFT OUTER JOIN POLL_WQCRITERIA ON WQCRITERIA.WQCRITID = POLL_WQCRITERIA.WQCRITID " & "Where POLL_WQCRITERIA.POLLID = " & poll.Item("POLLID")
                Dim wqCmd As New OleDbCommand(strSQLWQStd, modUtil.g_DBConn)
                _wq = New OleDbDataAdapter(wqCmd)
                _dtWQ = New Data.DataTable
                _wq.Fill(_dtWQ)
                dgvWaterQuality.DataSource = _dtWQ
                poll.Close()

            End If

            _boolChanged = False
            CmdSaveEnabled()
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
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
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub NoSaveCoeffSetChange()
        Try
            Dim strSQLFullCoeff As String
            Dim strSQLCoeffs As String

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
            _coefs = New OleDbDataAdapter(coefsCmd)
            _dtCoeff = New Data.DataTable
            _coefs.Fill(_dtCoeff)
            dgvCoef.DataSource = _dtCoeff

            _boolChanged = False
            _boolDescChanged = False
            CmdSaveEnabled()
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub txtCoeffSetDesc_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCoeffSetDesc.TextChanged
        Try
            _boolDescChanged = True
            CmdSaveEnabled()
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
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
                        Close()
                    End If

                Else

                    Close()

                End If
            Else

                Close()
            End If
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Try
            If ValidateGridValues() Then

                UpdateValues()
                _boolSaved = True
                MsgBox(cboPollName.Text & " saved successfully.", MsgBoxStyle.Information, "OpenNSPECT")
                Close()

            End If
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub mnuPollHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPollHelp.Click
        Try
            System.Windows.Forms.Help.ShowHelp(Me, modUtil.g_nspectPath & "\Help\nspect.chm", "pollutants.htm")
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub mnuCoeffHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCoeffHelp.Click
        Try
            System.Windows.Forms.Help.ShowHelp(Me, modUtil.g_nspectPath & "\Help\nspect.chm", "pol_coeftab.htm")
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub mnuAddPoll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddPoll.Click
        Try
            Dim newPoll As New frmNewPollutants
            newPoll.Init(Me)
            newPoll.ShowDialog()
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
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
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub mnuCoeffNewSet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCoeffNewSet.Click
        Try
            g_boolAddCoeff = True
            Dim addCoeff As New frmAddCoeffSet
            addCoeff.Init(Me, Nothing)
            addCoeff.ShowDialog()
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try

    End Sub


    Private Sub mnuCoeffCopySet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCoeffCopySet.Click
        Try
            g_boolCopyCoeff = True
            Dim newCopyCoef As New frmCopyCoeffSet
            newCopyCoef.Init(_coefCmd, Me, Nothing)
            newCopyCoef.ShowDialog()
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub mnuCoeffDeleteSet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCoeffDeleteSet.Click

        'Using straight command text to rid ourselves of the dreaded coefficient sets

        Try
            Dim intAns As Short
            intAns = MsgBox("Are you sure you want to delete the coefficient set '" & cboCoeffSet.Text & "' associated with pollutant '" & cboPollName.Text & "'?", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Confirm Delete")

            'code to handle response
            If intAns = MsgBoxResult.Yes Then

                Dim strDeleteCoeffSet As String = "DELETE * from COEFFICIENTSET WHERE NAME LIKE '" & cboCoeffSet.Text & "'"
                Dim cmdDelCoef As New DataHelper(strDeleteCoeffSet)
                cmdDelCoef.ExecuteNonQuery()

                MsgBox(cboCoeffSet.Text & " deleted.", MsgBoxStyle.OkOnly, "Record Deleted")

                cboPollName.Items.Clear()
                cboCoeffSet.Items.Clear()

                modUtil.InitComboBox(cboPollName, "Pollutant")

                Me.Refresh()
            End If

        Catch ex As Exception
            MsgBox("Error deleting coefficient set.", MsgBoxStyle.Critical, "Error")
            MsgBox(Err.Number & ": " & Err.Description)
        End Try
    End Sub


    Private Sub mnuCoeffImportSet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCoeffImportSet.Click
        Try
            Dim newImportCoef As New frmImportCoeffSet
            newImportCoef.Init(Me)
            newImportCoef.ShowDialog()
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub mnuCoeffExportSet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCoeffExportSet.Click
        Try
            Dim dlgSave As New Windows.Forms.SaveFileDialog

            'browse...get output filename
            dlgSave.Filter = Replace(MSG1TextFile, "<name>", "Coefficient Set")
            dlgSave.Title = Replace(MSG3, "<name>", "Coefficient Set")
            dlgSave.DefaultExt = ".txt"
            If dlgSave.ShowDialog = Windows.Forms.DialogResult.OK Then
                ExportCoeffSet(dlgSave.FileName)
            End If
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try

    End Sub


    Private Sub dgvCoef_CellValueChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvCoef.CellValueChanged, dgvWaterQuality.CellValueChanged
        Try
            _boolChanged = True
            CmdSaveEnabled()
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub
#End Region


#Region "Helper Functions"
    Private Sub CmdSaveEnabled()
        If _boolChanged Or _boolDescChanged Then
            cmdSave.Enabled = True
        Else
            cmdSave.Enabled = False
        End If

    End Sub


    Private Function ValidateGridValues() As Boolean
        Try
            'Need to validate each grid value before saving.  Essentially we take it a row at a time,
            'then rifle through each column of each row.  Case Select tests each each x,y value depending
            'on column... 3-6 must be 1-100 range

            'Returns: True or False

            Dim val As Double
            Dim i As Short
            Dim j As Short
            Dim iQstd As Short

            For i = 0 To dgvCoef.Rows.Count - 1

                For j = 2 To 5
                    val = dgvCoef.Rows(i).Cells(j).Value

                    If InStr(1, CStr(val), ".", CompareMethod.Text) > 0 Then
                        If (Len(Split(CStr(val), ".")(1)) > 4) Then
                            DisplayError(Err6, i, j)
                            Return False
                        End If
                    End If

                    If Not IsNumeric(val) Or (val < 0) Or (val > 1000) Then
                        DisplayError(Err6, i, j)
                        Return False
                    End If
                Next j
            Next i

            For iQstd = 0 To dgvWaterQuality.Rows.Count - 1
                val = dgvWaterQuality.Rows(iQstd).Cells(2).Value

                If Not IsNumeric(val) Or (val < 0) Then
                    DisplayError(Err5, iQstd, 3)
                    Return False
                End If
            Next iQstd

            ValidateGridValues = True

        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Function


    Private Sub UpdateValues()
        Try
            Dim strUpdateDescription As Object
            If _boolDescChanged Then
                strUpdateDescription = "SELECT Description from CoefficientSet Where Name like '" & cboCoeffSet.Text & "'"
                Dim cmdDesc As New DataHelper(strUpdateDescription)
                Dim adDesc = cmdDesc.GetAdapter()
                Dim buildDesc As New OleDbCommandBuilder(adDesc)
                buildDesc.QuotePrefix = "["
                buildDesc.QuoteSuffix = "]"
                Dim dt As New Data.DataTable
                adDesc.Fill(dt)

                If Len(txtCoeffSetDesc.Text) = 0 Then
                    dt.Rows(0)("Description") = ""
                Else
                    dt.Rows(0)("Description") = txtCoeffSetDesc.Text
                End If
                adDesc.Update(dt)
            End If

            Dim i As Short
            Dim strPollUpdate As String
            Dim strWQSelect As String

            If ValidateGridValues() Then
                'Update

                For i = 0 To dgvCoef.Rows.Count - 1

                    strPollUpdate = "SELECT * From Coefficient Where CoeffID = " & dgvCoef.Rows(i).Cells(6).Value.ToString
                    Dim cmdNewCoef As New DataHelper(strPollUpdate)
                    Dim adaptNewCoeff = cmdNewCoef.GetAdapter()
                    Dim cbuilder As New OleDbCommandBuilder(adaptNewCoeff)
                    cbuilder.QuotePrefix = "["
                    cbuilder.QuoteSuffix = "]"
                    Dim dt As New Data.DataTable
                    adaptNewCoeff.Fill(dt)

                    dt.Rows(0)("Coeff1") = dgvCoef.Rows(i).Cells(2).Value
                    dt.Rows(0)("Coeff2") = dgvCoef.Rows(i).Cells(3).Value
                    dt.Rows(0)("Coeff3") = dgvCoef.Rows(i).Cells(4).Value
                    dt.Rows(0)("Coeff4") = dgvCoef.Rows(i).Cells(5).Value

                    adaptNewCoeff.Update(dt)
                Next i
            End If


            For i = 0 To dgvWaterQuality.Rows.Count - 1
                strWQSelect = "SELECT * from POLL_WQCRITERIA WHERE POLL_WQCRITID = " & dgvWaterQuality.Rows(i).Cells(3).Value.ToString
                Dim cmdNewWQ As New DataHelper(strWQSelect)
                Dim adaptNewWQ = cmdNewWQ.GetAdapter()
                Dim wqbuilder As New OleDbCommandBuilder(adaptNewWQ)
                wqbuilder.QuotePrefix = "["
                wqbuilder.QuoteSuffix = "]"
                Dim dt As New Data.DataTable
                adaptNewWQ.Fill(dt)

                dt.Rows(0)("Threshold") = dgvWaterQuality.Rows(i).Cells(2).Value

                adaptNewWQ.Update(dt)
            Next i

        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub DeletePollutant(ByRef strName As String)
        Try
            Dim strPollDelete As String = "Delete * FROM Pollutant WHERE NAME LIKE '" & strName & "'"
            Dim cmdPollDel As New DataHelper(strPollDelete)
            cmdPollDel.ExecuteNonQuery()

            MsgBox(strName & " deleted.", MsgBoxStyle.OkOnly, "Record Deleted")

            Me.cboPollName.Items.Clear()
            modUtil.InitComboBox(Me.cboPollName, "Pollutant")
            Me.Refresh()

        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Public Sub AddCoefficient(ByRef strCoeffName As String, ByRef strLCType As String)
        Try
            'General gist:  First we add new record to the Coefficient Set table using strCoeffName as
            'the name, m_intPollID as the PollID, and m_intLCTYPEID as the LCTypeID.  The last two are
            'garnered above during a cbo click event.  Once that's done, we'll add a series of blank
            'coefficients for the landclass type the user chooses...ie CCAP, NotCCAP, whatever

            Dim strNewLcType As String 'CmdString for inserting new coefficientset
            Dim strGetLcType As String
            Dim strDefault As String '
            Dim strNewCoeffID As String 'Holder for the CoefficientSetID
            Dim intCoeffSetID As Short

            strGetLcType = "SELECT * FROM LCTYPE WHERE NAME LIKE '" & strLCType & "'"
            Dim cmdLCType As New DataHelper(strGetLcType)
            Dim datalctype As OleDbDataReader = cmdLCType.ExecuteReader()
            datalctype.Read()


            'First need to add the coefficient set to that table
            strNewLcType = "INSERT INTO COEFFICIENTSET(NAME, POLLID, LCTYPEID) VALUES ('" & Replace(strCoeffName, "'", "''") & "'," & Replace(CStr(_intPollID), "'", "''") & "," & Replace(datalctype("LCTypeID"), "'", "''") & ")"
            Dim cmdInsLC As New DataHelper(strNewLcType)
            cmdInsLC.ExecuteNonQuery()

            'Get the Coefficient Set ID of the newly created coefficient set to populate Column # 8 in the GRid,
            'which by the way, is hidden from view.  InitPollDef sets the widths of col 7, 8 to 0
            strNewCoeffID = "SELECT COEFFSETID FROM COEFFICIENTSET " & "WHERE COEFFICIENTSET.NAME LIKE '" & strCoeffName & "'"
            Dim cmdNewCoefID As New DataHelper(strNewCoeffID)
            Dim dataNewCoeffID As OleDbDataReader = cmdNewCoefID.ExecuteReader()
            dataNewCoeffID.Read()
            intCoeffSetID = dataNewCoeffID("CoeffSetID")
            dataNewCoeffID.Close()


            strDefault = "SELECT LCTYPE.LCTYPEID, LCCLASS.LCCLASSID, LCCLASS.NAME As valName, " & "LCCLASS.VAlue as valValue FROM LCTYPE " & "INNER JOIN LCCLASS ON LCCLASS.LCTYPEID = LCTYPE.LCTYPEID " & "WHERE LCTYPE.Name Like " & "'" & strLCType & "' ORDER BY LCCLASS.Value"
            Dim cmdCopySet As New DataHelper(strDefault)
            Dim dataCopySet As OleDbDataReader = cmdCopySet.ExecuteReader()


            'Now loopy loo to populate values.
            Dim strNewCoeff1 As String
            strNewCoeff1 = "SELECT * FROM COEFFICIENT"
            Dim cmdNewCoef As New DataHelper(strNewCoeff1)
            Dim adaptNewCoeff = cmdNewCoef.GetAdapter()
            Dim cbuilder As New OleDbCommandBuilder(adaptNewCoeff)
            cbuilder.QuotePrefix = "["
            cbuilder.QuoteSuffix = "]"
            Dim dt As New Data.DataTable
            adaptNewCoeff.Fill(dt)

            While dataCopySet.Read()
                Dim row As Data.DataRow = dt.NewRow()
                row("Coeff1") = 0
                row("Coeff2") = 0
                row("Coeff3") = 0
                row("Coeff4") = 0
                row("CoeffSetID") = intCoeffSetID
                row("LCClassID") = dataCopySet("LCClassID")
                dt.Rows.Add(row)
            End While
            adaptNewCoeff.Update(dt)
            dataCopySet.Close()

            cboPollName.SelectedIndex = cboPollName.SelectedIndex

            cboCoeffSet.Items.Add(strCoeffName)

            'Call the function to set everything to newly added Coefficient.
            cboCoeffSet.SelectedIndex = GetCboIndex(strCoeffName, cboCoeffSet)

            txtLCType.Text = datalctype("Name")
            datalctype.Close()

        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Public Sub CopyCoefficient(ByRef strNewCoeffName As String, ByRef strCoeffSet As String)
        Try
            'General gist:  First we add new record to the Coefficient Set table using strNewCoeffName as
            'the name, PollID, LCTYPEID.  Once that's done, we'll add the coefficients
            'from the set being copied
            Dim strCopySet As String 'The Recordset of existing coefficients being copied
            Dim strNewLcType As String 'CmdString for inserting new coefficientset               '
            Dim strNewCoeffID As String 'Holder for the CoefficientSetID
            Dim intCoeffSetID As Short

            strCopySet = "SELECT * FROM COEFFICIENTSET INNER JOIN COEFFICIENT ON COEFFICIENTSET.COEFFSETID = " & "COEFFICIENT.COEFFSETID WHERE COEFFICIENTSET.NAME LIKE '" & strCoeffSet & "'"
            Dim cmdCopySet As New DataHelper(strCopySet)
            Dim dataCopySet As OleDbDataReader = cmdCopySet.ExecuteReader()
            dataCopySet.Read()

            'INSERT: new Coefficient set taking the PollID and LCType ID from rsCopySet
            strNewLcType = "INSERT INTO COEFFICIENTSET(NAME, POLLID, LCTYPEID) VALUES ('" & Replace(strNewCoeffName, "'", "''") & "'," & dataCopySet("POLLID") & "," & dataCopySet("LCTypeID") & ")"

            'First need to add the coefficient set to that table
            Dim cmdInsLCType As New DataHelper(strNewLcType)
            cmdInsLCType.ExecuteNonQuery()

            'Get the Coefficient Set ID of the newly created coefficient set to populate Column # 8 in the GRid,
            'which by the way, is hidden from view.  InitPollDef sets the widths of col 7, 8 to 0
            strNewCoeffID = "SELECT COEFFSETID FROM COEFFICIENTSET " & "WHERE COEFFICIENTSET.NAME LIKE '" & strNewCoeffName & "'"
            Dim cmdNewCoefID As New DataHelper(strNewCoeffID)
            Dim dataNewCoeffID As OleDbDataReader = cmdNewCoefID.ExecuteReader()
            dataNewCoeffID.Read()
            intCoeffSetID = dataNewCoeffID("CoeffSetID")
            dataNewCoeffID.Close()

            'Now loopy loo to populate values.
            Dim strNewCoeff1 As String
            strNewCoeff1 = "SELECT * FROM COEFFICIENT"
            Dim cmdNewCoef As New DataHelper(strNewCoeff1)
            Dim adaptNewCoeff = cmdNewCoef.GetAdapter()
            Dim cbuilder As New OleDbCommandBuilder(adaptNewCoeff)
            cbuilder.QuotePrefix = "["
            cbuilder.QuoteSuffix = "]"
            Dim dt As New Data.DataTable
            adaptNewCoeff.Fill(dt)

            'Clear things and set the rows to recordcount + 1, remember 1st row fixed
            'dgvCoef.Rows.Clear()

            'Actually add the records to the new set
            Do
                Dim row As Data.DataRow = dt.NewRow()
                row("Coeff1") = dataCopySet("Coeff1")
                row("Coeff2") = dataCopySet("Coeff2")
                row("Coeff3") = dataCopySet("Coeff3")
                row("Coeff4") = dataCopySet("Coeff4")
                row("CoeffSetID") = intCoeffSetID
                row("LCClassID") = dataCopySet("LCClassID")
                dt.Rows.Add(row)
            Loop While dataCopySet.Read()
            adaptNewCoeff.Update(dt)
            dataCopySet.Close()

            'Set up everything to look good
            cboPollName_SelectedIndexChanged(cboPollName, New System.EventArgs())
            cboCoeffSet.SelectedIndex = GetCboIndex(strNewCoeffName, cboCoeffSet)
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub

    'Exports your current standard and pollutants to text or csv.


    Private Sub ExportCoeffSet(ByRef strFileName As String)
        Try
            Dim out As New IO.StreamWriter(strFileName)

            'Write name of pollutant and threshold
            For i As Integer = 0 To dgvCoef.Rows.Count - 1
                out.WriteLine(dgvCoef.Rows(i).Cells(0).Value.ToString & "," & dgvCoef.Rows(i).Cells(2).Value.ToString & "," & dgvCoef.Rows(i).Cells(3).Value.ToString & "," & dgvCoef.Rows(i).Cells(4).Value.ToString & "," & dgvCoef.Rows(i).Cells(5).Value.ToString)
            Next i
            out.Close()
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Public Sub UpdateCoeffSet(ByRef cmdCoeff As OleDbCommand, ByRef strCoeffName As String, ByRef strFileName As String)
        Try
            'General gist:  First we add new record to the Coefficient Set table using strCoeffName as
            'the name, m_intPollID as the PollID, and m_intLCTYPEID as the LCTypeID.  The last two are
            'garnered above during a cbo click event.  Once that's done, we'll add a series of
            'coefficients for the landclass based on the incoming textfile...strFileName

            Dim strNewLcType As String 'CmdString for inserting new coefficientset
            Dim strNewCoeffID As String 'Holder for the CoefficientSetID
            Dim intCoeffSetID As Short
            Dim i As Short

            'Textfile related material
            Dim strLine As String
            Dim strValue As Short
            Dim intLine As Short

            Dim dataCoeff As OleDbDataReader = cmdCoeff.ExecuteReader()
            dataCoeff.Read()
            strNewLcType = "INSERT INTO COEFFICIENTSET(NAME, POLLID, LCTYPEID) VALUES ('" & Replace(strCoeffName, "'", "''") & "'," & _intPollID & "," & dataCoeff("LCTypeID") & ")"
            dataCoeff.Close()

            'First need to add the coefficient set to that table
            Dim cmdInsCoef As New DataHelper(strNewLcType)
            cmdInsCoef.ExecuteNonQuery()

            'Get the Coefficient Set ID of the newly created coefficient set to populate Column # 8 in the GRid,
            'which by the way, is hidden from view.  InitPollDef sets the widths of col 7, 8 to 0
            strNewCoeffID = "SELECT COEFFSETID FROM COEFFICIENTSET " & "WHERE COEFFICIENTSET.NAME LIKE '" & strCoeffName & "'"
            Dim cmdNewCoefID As New DataHelper(strNewCoeffID)
            Dim dataNewCoeffID As OleDbDataReader = cmdNewCoefID.ExecuteReader()
            dataNewCoeffID.Read()
            intCoeffSetID = dataNewCoeffID("CoeffSetID")
            dataNewCoeffID.Close()

            'Now turn attention to the TextFile...to get the users coefficient values
            Dim read As New IO.StreamReader(strFileName)
            intLine = 0

            'Now loopy loo to populate values.
            Dim strNewCoeff1 As String
            strNewCoeff1 = "SELECT * FROM COEFFICIENT"
            Dim cmdNewCoef As New DataHelper(strNewCoeff1)
            Dim adaptNewCoeff = cmdNewCoef.GetAdapter()
            Dim cbuilder As New OleDbCommandBuilder(adaptNewCoeff)
            cbuilder.QuotePrefix = "["
            cbuilder.QuoteSuffix = "]"
            Dim dt As New Data.DataTable
            adaptNewCoeff.Fill(dt)

            i = 0

            _dtCoeff.Rows.Clear()

            Do While Not read.EndOfStream

                strLine = read.ReadLine
                'Value exits??
                strValue = CShort(Split(strLine, ",")(0))

                dataCoeff = cmdCoeff.ExecuteReader()
                While dataCoeff.Read()
                    If dataCoeff("Value") = strValue Then
                        Dim row As Data.DataRow = dt.NewRow()
                        row("Coeff1") = Split(strLine, ",")(1)
                        row("Coeff2") = Split(strLine, ",")(2)
                        row("Coeff3") = Split(strLine, ",")(3)
                        row("Coeff4") = Split(strLine, ",")(4)
                        row("CoeffSetID") = intCoeffSetID
                        row("LCClassID") = dataCoeff("LCClassID")
                        dt.Rows.Add(row)

                        Dim drow As Data.DataRow = _dtCoeff.NewRow()
                        drow(0) = strValue
                        drow(1) = dataCoeff("Name")
                        drow(2) = Split(strLine, ",")(1)
                        drow(3) = Split(strLine, ",")(2)
                        drow(4) = Split(strLine, ",")(3)
                        drow(5) = Split(strLine, ",")(4)
                        drow(6) = intCoeffSetID
                        drow(7) = row("coeffID")
                    End If
                End While
                adaptNewCoeff.Update(dt)
                dataCoeff.Close()
            Loop

            cboCoeffSet.Items.Clear()
            Dim coef As OleDbDataReader = _coefCmd.ExecuteReader()
            While coef.Read()
                cboCoeffSet.Items.Add(coef("Name"))
            End While
            cboCoeffSet.SelectedIndex = modUtil.GetCboIndex(strCoeffName, cboCoeffSet)
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub
#End Region


End Class