'********************************************************************************************************
'File Name: frmNewPollutants.vb
'Description: Form for entering new pollutants
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
Friend Class NewPollutantForm
    Inherits System.Windows.Forms.Form

    Private _boolLoaded As Boolean
    Private _boolChanged As Boolean

    Private _intPollID As Short 'There's a need to have the PollID so we'll store it here
    Private _intLCTypeID As Short 'Land Class (CCAP) ID - needed to add new coefficient sets

    Private _frmPoll As PollutantsForm
    Const c_sModuleFileName As String = "frmNewPollutants.vb"


#Region "Events"


    Private Sub frmNewPollutants_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            modUtil.InitComboBox(cboLCType, "LCType")

            _boolLoaded = True
            _boolChanged = False
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub cboLCType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboLCType.SelectedIndexChanged
        Try
            Dim strLCClasses As String

            strLCClasses = "SELECT LCTYPE.LCTYPEID, LCCLASS.VALUE, LCCLASS.NAME, LCCLASS.LCCLASSID FROM LCTYPE INNER JOIN LCCLASS ON " & "LCTYPE.LCTYPEID = LCCLASS.LCTYPEID WHERE LCTYPE.NAME LIKE '" & cboLCType.Text & "'" & " ORDER BY LCCLASS.VALUE"
            Using cmdLC As New DataHelper(strLCClasses)
                Using dataLC As OleDbDataReader = cmdLC.ExecuteReader()
                    dgvCoef.Rows.Clear()
                    'Actually add the records to the new set
                    Dim rowNum As Integer
                    While dataLC.Read()
                        rowNum = dgvCoef.Rows.Add()
                        dgvCoef.Rows(rowNum).Cells(0).Value = dataLC("Value")
                        dgvCoef.Rows(rowNum).Cells(1).Value = dataLC("Name")
                        dgvCoef.Rows(rowNum).Cells(2).Value = 0
                        dgvCoef.Rows(rowNum).Cells(3).Value = 0
                        dgvCoef.Rows(rowNum).Cells(4).Value = 0
                        dgvCoef.Rows(rowNum).Cells(5).Value = 0
                        dgvCoef.Rows(rowNum).Cells(6).Value = 0
                        dgvCoef.Rows(rowNum).Cells(7).Value = dataLC("LCClassID")
                    End While
                    dataLC.Close()
                End Using
            End Using
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub mnuCoeffNewSet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCoeffNewSet.Click
        Try
            g_boolAddCoeff = False
            Dim addCoeff As New NewCoefficientSetForm
            addCoeff.Init(Nothing, Me)
            addCoeff.ShowDialog()
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub mnuCoeffCopySet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCoeffCopySet.Click
        Try
            g_boolCopyCoeff = False
            Dim newCopyCoef As New CopyCoefficientSetForm
            newCopyCoef.Init(Nothing, Nothing, Me)
            newCopyCoef.ShowDialog()
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub dgvCoef_CellValueChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvCoef.CellValueChanged
        Try
            _boolChanged = True
            CmdSaveEnabled()
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub txtPollutant_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPollutant.TextChanged
        Try
            _boolChanged = True
            CmdSaveEnabled()
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub txtCoeffSet_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCoeffSet.TextChanged
        Try
            _boolChanged = True
            CmdSaveEnabled()
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub txtCoeffSetDesc_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCoeffSetDesc.TextChanged
        Try
            _boolChanged = True
            CmdSaveEnabled()
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub cmdQuit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdQuit.Click
        Try
            Close()
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Try
            If CheckForm() Then
                If UpdateValues() Then
                    MsgBox(txtPollutant.Text & " successfully added.  Please enter value for associated water quality standards.", MsgBoxStyle.Information, "Pollutant Successfully Added")
                    Close()
                    _frmPoll.SSTab1.SelectedIndex = 1
                End If
            End If
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub
#End Region

#Region "Helpers"


    Public Sub Init(ByRef frmPoll As PollutantsForm)
        Try
            _frmPoll = frmPoll
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub CmdSaveEnabled()
        Try
            If _boolChanged Then
                cmdSave.Enabled = True
            Else
                cmdSave.Enabled = False
            End If
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Function CheckForm() As Boolean
        Try
            If Trim(txtPollutant.Text) = "" Then
                MsgBox("Please enter a name for the new pollutant", MsgBoxStyle.Critical, "Enter Name")
                CheckForm = False
                txtPollutant.Focus()
                txtPollutant.SelectionLength = Len(txtPollutant.Text)
                Exit Function
            ElseIf modUtil.UniqueName("Pollutant", txtPollutant.Text) Then
                CheckForm = True
            Else
                MsgBox(Err4, MsgBoxStyle.Critical, "Name In Use")
                CheckForm = False
                txtPollutant.Focus()
                txtPollutant.SelectionLength = Len(txtPollutant.Text)
                Exit Function
            End If

            If Len(Trim(txtCoeffSet.Text)) = 0 Then
                MsgBox("Please enter a name for the new pollutant", MsgBoxStyle.Critical, "Enter Name")
                CheckForm = False
                txtPollutant.Focus()
                txtPollutant.SelectionLength = Len(txtPollutant.Text)
                Exit Function
            ElseIf modUtil.UniqueName("Coefficientset", (txtCoeffSet.Text)) Then
                CheckForm = True
            Else
                MsgBox(Err4, MsgBoxStyle.Critical, "Name In Use")
                CheckForm = False
                txtCoeffSet.Focus()
                txtCoeffSet.SelectionLength = Len(txtPollutant.Text)
                Exit Function
            End If

            'Now if all is there and good, go on and check the grid values
            If ValidateGridValues() Then
                CheckForm = True
            Else
                CheckForm = False
            End If

        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Function


    Private Function ValidateGridValues() As Boolean
        Try
            'Need to validate each grid value before saving.  Essentially we take it a row at a time,
            'then rifle through each column of each row.  Case Select tests each each x,y value depending
            'on column... 3-6 must be 1-100 range

            'Returns: True or False

            Dim val As Double

            For i As Integer = 0 To dgvCoef.Rows.Count - 1
                For j As Integer = 2 To 5

                    val = dgvCoef.Rows(i).Cells(j).Value

                    If InStr(1, val.ToString, ".", CompareMethod.Text) > 0 Then
                        If (Len(Split(val.ToString, ".")(1)) > 4) Then
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

            ValidateGridValues = True

        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Function


    Private Function UpdateValues() As Boolean

        Dim strInsertPollutant As String 'Insert String for new poll
        Dim strSelectPollutant As String 'Select string for new poll
        Dim strSelectLCType As String 'Select string for LCType
        Dim strInsertCoeffSet As String 'Insert string for new coeff set
        Dim strSelectCoeffSet As String 'Select string for new coeff set
        Dim strInsertCoeffs As String 'Insert string for the coefficients
        Dim strSelectWQCrit As String 'Select string for Water Quality
        Dim strInsertWQCrit As String 'Insert string for Water Quality
        Dim strNewColor As String 'New color


        Try
            'Step 1a: Get a new color for this pollutant
            strNewColor = modUtil.ReturnHSVColorString

            'Step 1: Insert the New Pollutant
            strInsertPollutant = "INSERT INTO POLLUTANT(NAME, POLLTYPE, COLOR) VALUES ('" & Replace(Trim(txtPollutant.Text), "'", "''") & "', 0, " & "'" & strNewColor & "'" & ")"
            Dim cmdinspol As New DataHelper(strInsertPollutant)
            cmdinspol.ExecuteNonQuery()

            'Step 2: Select the newly inserted pollutant info
            strSelectPollutant = "SELECT * FROM POLLUTANT WHERE NAME LIKE '" & Trim(txtPollutant.Text) & "'"
            Dim cmdSelPoll As New DataHelper(strSelectPollutant)
            Dim dataNewPoll As OleDbDataReader = cmdSelPoll.ExecuteReader()
            dataNewPoll.Read()

            'Step 2a: Select the WQ Standards
            strSelectWQCrit = "SELECT * FROM WQCriteria"
            Dim cmdWQCrit As New DataHelper(strSelectWQCrit)
            Dim dataWQCrit As OleDbDataReader = cmdWQCrit.ExecuteReader()

            While dataWQCrit.Read()
                strInsertWQCrit = "INSERT INTO POLL_WQCRITERIA (POLLID, WQCRITID, THRESHOLD) VALUES (" & dataNewPoll("POLLID") & "," & dataWQCrit("WQCRITID") & "," & "0 )"
                Dim cmdInsWQCrit As New DataHelper(strInsertWQCrit)
                cmdInsWQCrit.ExecuteNonQuery()
            End While
            dataWQCrit.Close()

            'Step 3: Get the LCtype information
            strSelectLCType = "SELECT * FROM LCTYPE WHERE NAME LIKE '" & cboLCType.Text & "'"
            Dim cmdNewType As New DataHelper(strSelectLCType)
            Dim dataNewType As OleDbDataReader = cmdNewType.ExecuteReader()
            dataNewType.Read()

            'Step 4: Insert the New coefficient set
            strInsertCoeffSet = "INSERT INTO COEFFICIENTSET (NAME, DESCRIPTION, LCTYPEID, POLLID) VALUES ('" & Replace(Trim(txtCoeffSet.Text), "'", "''") & "', '" & Replace(Trim(txtCoeffSetDesc.Text), "'", "''") & "'," & dataNewType("LCTypeID") & "," & dataNewPoll("POLLID") & ")"
            dataNewPoll.Close()
            dataNewType.Close()
            Dim cmdInsCoef As New DataHelper(strInsertCoeffSet)
            cmdInsCoef.ExecuteNonQuery()

            'Step 5: Select the newly inserted coefficient set
            strSelectCoeffSet = "SELECT * FROM COEFFICIENTSET WHERE NAME LIKE '" & txtCoeffSet.Text & "'"
            Dim cmdSelCoef As New DataHelper(strSelectCoeffSet)
            Dim dataCoeff As OleDbDataReader = cmdSelCoef.ExecuteReader()
            dataCoeff.Read()

            'Step 6: Insert the new coeffs for that set
            For i As Integer = 0 To dgvCoef.Rows.Count - 1
                strInsertCoeffs = "INSERT INTO COEFFICIENT (COEFF1, COEFF2, COEFF3, COEFF4, COEFFSETID, LCCLASSID) VALUES (" & dgvCoef.Rows(i).Cells(2).Value.ToString & ", " & dgvCoef.Rows(i).Cells(3).Value.ToString & ", " & dgvCoef.Rows(i).Cells(4).Value.ToString & ", " & dgvCoef.Rows(i).Cells(5).Value.ToString & ", " & dataCoeff("CoeffSetID") & ", " & dgvCoef.Rows(i).Cells(7).Value.ToString & ")"
                Dim cmdInsCoeffs As New DataHelper(strInsertCoeffs)
                cmdInsCoeffs.ExecuteNonQuery()
            Next i
            dataCoeff.Close()

            _frmPoll.cboPollName.Items.Clear()
            modUtil.InitComboBox(_frmPoll.cboPollName, "Pollutant")
            _frmPoll.cboPollName.SelectedIndex = modUtil.GetCboIndex((txtPollutant.Text), _frmPoll.cboPollName)

            UpdateValues = True
        Catch ex As Exception
            'ALFREDO
            MsgBox("An error occurred while creating new pollutant." & vbNewLine & Err.Number & ": " & Err.Description, MsgBoxStyle.Critical, "Error")
            HandleError(c_sModuleFileName, ex)
        End Try
    End Function


    Public Sub CopyCoefficient(ByRef strNewCoeffName As String, ByRef strCoeffSet As String)
        Try
            'General gist:  First we add new record to the Coefficient Set table using strNewCoeffName as
            'the name, PollID, LCTYPEID.  Once that's done, we'll add the coefficients
            'from the set being copied
            Dim strCopySet As String 'The Recordset of existing coefficients being copied
            Dim strLandClass As String 'Select for Landclass
            Dim i As Short = 0

            strCopySet = "SELECT * FROM COEFFICIENTSET INNER JOIN COEFFICIENT ON COEFFICIENTSET.COEFFSETID = " & "COEFFICIENT.COEFFSETID WHERE COEFFICIENTSET.NAME LIKE '" & strCoeffSet & "'"
            Dim cmdCopySet As New DataHelper(strCopySet)
            Dim dataCopySet As OleDbDataReader = cmdCopySet.ExecuteReader()

            'Step 1: Enter name
            txtCoeffSet.Text = strNewCoeffName

            'Clear things and set the rows to recordcount + 1, remember 1st row fixed
            dgvCoef.Rows.Clear()

            'Actually add the records to the new set
            While dataCopySet.Read()
                strLandClass = "SELECT * FROM LCCLASS WHERE LCCLASSID = " & dataCopySet("LCClassID")
                'Let's try one more ADO method, why not, righ?
                Dim cmdLC As New DataHelper(strLandClass)
                Dim dataLC As OleDbDataReader = cmdLC.ExecuteReader
                dataLC.Read()

                'Add the necessary components
                dgvCoef.Rows(i).Cells(0).Value = dataLC("Value")
                dgvCoef.Rows(i).Cells(1).Value = dataLC("Value")
                dgvCoef.Rows(i).Cells(2).Value = dataCopySet("Value")
                dgvCoef.Rows(i).Cells(3).Value = dataCopySet("Value")
                dgvCoef.Rows(i).Cells(4).Value = dataCopySet("Value")
                dgvCoef.Rows(i).Cells(5).Value = dataCopySet("Value")

                dataLC.Close()
                i = i + 1
            End While

        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Public Sub AddCoefficient(ByRef strCoeffName As String, ByRef strLCType As String)
        Try
            'TODO: verify this is even possible without the _intPollID and _intLCTypeID

            'General gist:  First we add new record to the Coefficient Set table using strCoeffName as
            'the name, m_intPollID as the PollID, and m_intLCTYPEID as the LCTypeID.  The last two are
            'garnered above during a cbo click event.  Once that's done, we'll add a series of blank
            'coefficients for the landclass type the user chooses...ie CCAP, NotCCAP, whatever

            Dim strNewLcType As String 'CmdString for inserting new coefficientset
            Dim strDefault As String '
            Dim strNewCoeffID As String 'Holder for the CoefficientSetID
            Dim intCoeffSetID As Short
            Dim i As Short = 0


            'First need to add the coefficient set to that table
            strNewLcType = "INSERT INTO COEFFICIENTSET(NAME, POLLID, LCTYPEID) VALUES ('" & Replace(strCoeffName, "'", "''") & "'," & _intPollID & "," & _intLCTypeID & ")"
            Dim cmdInsLC As New DataHelper(strNewLcType)
            cmdInsLC.ExecuteNonQuery()

            'Get the Coefficient Set ID of the newly created coefficient set to populate Column # 8 in the GRid,
            'which by the way, is hidden from view.  InitPollDef sets the widths of col 7, 8 to 0
            strNewCoeffID = "SELECT COEFFSETID FROM COEFFICIENTSET " & "WHERE COEFFICIENTSET.NAME LIKE '" & strCoeffName & "'"
            Dim cmdNewCoefID As New DataHelper(strNewCoeffID)
            Dim dataNewCoeffID As OleDbDataReader = cmdNewCoefID.ExecuteReader()
            dataNewCoeffID.Read()
            intCoeffSetID = dataNewCoeffID("CoeffSetID")

            strDefault = "SELECT LCTYPE.LCTYPEID, LCCLASS.LCCLASSID, LCCLASS.NAME As valName, " & "LCCLASS.VAlue as valValue FROM LCTYPE " & "INNER JOIN LCCLASS ON LCCLASS.LCTYPEID = LCTYPE.LCTYPEID " & "WHERE LCTYPE.Name Like " & "'" & strLCType & "'"
            Dim cmdCopySet As New DataHelper(strDefault)
            Dim dataCopySet As OleDbDataReader = cmdCopySet.ExecuteReader()

            'Clear things and set the rows to recordcount + 1, remember 1st row fixed
            dgvCoef.Rows.Clear()

            'Now loopy loo to populate values.
            Dim strNewCoeff1 As String
            strNewCoeff1 = "SELECT * FROM COEFFICIENT"
            Dim cmdNewCoef As New DataHelper(strNewCoeff1)
            Dim adaptNewCoeff = cmdNewCoef.GetAdapter()
            Dim dt As New Data.DataTable
            adaptNewCoeff.Fill(dt)

            While dataCopySet.Read()
                Dim row As Data.DataRow = dt.NewRow()
                row("Coeff1") = 0
                row("Coeff2") = 0
                row("Coeff3") = 0
                row("Coeff4") = 0
                row("CoeffSetID") = dataNewCoeffID("CoeffSetID")
                row("LCClassID") = dataCopySet("LCClassID")
                dt.Rows.Add(row)


                dgvCoef.Rows(i).Cells(0).Value = dataCopySet("valValue")
                dgvCoef.Rows(i).Cells(1).Value = dataCopySet("valName")
                dgvCoef.Rows(i).Cells(2).Value = "0"
                dgvCoef.Rows(i).Cells(3).Value = "0"
                dgvCoef.Rows(i).Cells(4).Value = "0"
                dgvCoef.Rows(i).Cells(5).Value = "0"
                dgvCoef.Rows(i).Cells(6).Value = dataNewCoeffID("CoeffSetID")
                dgvCoef.Rows(i).Cells(7).Value = row("coeffID")

                i = i + 1
            End While

            dataCopySet.Close()
            dataNewCoeffID.Close()


            Close()

        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub

#End Region


End Class