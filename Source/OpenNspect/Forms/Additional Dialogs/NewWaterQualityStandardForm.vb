'********************************************************************************************************
'File Name: frmAddWQStd.vb
'Description: Form to add new water quality standard
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

Imports System.Windows.Forms
Imports System.Data.OleDb
Friend Class NewWaterQualityStandardForm
    Inherits System.Windows.Forms.Form

    Const c_sModuleFileName As String = "frmAddWQStd.vb"

    Private _frmWQStd As WaterQualityStandardsForm
    Private _frmPrj As MainForm
    Private _Change As Boolean


#Region "Events"


    Private Sub frmAddWQStd_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'Populate cbo with pollutant names
            _Change = False

            Dim strPollutant As String = "SELECT NAME FROM POLLUTANT ORDER BY NAME ASC"
            Using pollCmd As New DataHelper(strPollutant)
                Dim datPoll As OleDbDataReader = pollCmd.ExecuteReader
                Dim idx As Integer

                dgvWaterQuality.Rows.Clear()
                Do While datPoll.Read()
                    idx = dgvWaterQuality.Rows.Add()
                    dgvWaterQuality.Rows(idx).Cells("Pollutant").Value = datPoll.Item("Name")
                Loop
                datPoll.Close()
            End Using
            cmdSave.Enabled = False
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub txtWQStdName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtWQStdName.TextChanged
        Try
            txtWQStdName.Text = Replace(txtWQStdName.Text, "'", "")
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Try
            Dim intvbYesNo As Short = MsgBox("Are you sure you want to exit?  All changes not saved will be lost.", MsgBoxStyle.YesNo, "Exit?")

            If intvbYesNo = MsgBoxResult.Yes Then
                If Not _frmPrj Is Nothing Then
                    _frmPrj.cboWQStd.SelectedIndex = 0
                End If

                Close()
            Else
                Exit Sub
            End If

        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
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
                    If CheckThreshValues() Then
                        strCmd = "INSERT INTO WQCRITERIA (NAME,DESCRIPTION) VALUES ('" & Replace(txtWQStdName.Text, "'", "''") & "', '" & Replace(strDescript, "'", "''") & "')"
                        Using cmdInsert As New DataHelper(strCmd)
                            cmdInsert.ExecuteNonQuery()
                        End Using
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

            Close()

        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub dgvWaterQuality_CellValueChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvWaterQuality.CellValueChanged
        Try
            cmdSave.Enabled = True

        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub
#End Region

#Region "Helper Functions"


    Public Sub Init(ByRef frmWQStd As WaterQualityStandardsForm, ByRef frmPrj As MainForm)
        Try
            _frmWQStd = frmWQStd
            _frmPrj = frmPrj
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
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
            HandleError(c_sModuleFileName, ex)
        End Try
    End Function


    Private Sub PollutantAdd(ByRef strName As String, ByRef strPoll As String, ByRef intThresh As String)
        Try
            Dim strPollAdd As String
            Dim strPollDetails As String
            Dim strCmdInsert As String

            'Get the WQCriteria values using the name
            strPollAdd = "SELECT * FROM WQCriteria WHERE NAME = " & "'" & strName & "'"
            Using cmdPollAdd As New DataHelper(strPollAdd)
                Dim datPollAdd As OleDbDataReader = cmdPollAdd.ExecuteReader()
                datPollAdd.Read()

                'Get the pollutant particulars
                strPollDetails = "SELECT * FROM POLLUTANT WHERE NAME =" & "'" & strPoll & "'"
                Using cmdPollDetails As New DataHelper(strPollDetails)
                    Dim datPollDetails As OleDbDataReader = cmdPollDetails.ExecuteReader()
                    datPollDetails.Read()
                    If Trim(intThresh) = "" Then
                        strCmdInsert = String.Format("INSERT INTO POLL_WQCRITERIA (PollID,WQCritID) VALUES ('{0}', '{1}')", datPollDetails.Item("POLLID"), datPollAdd.Item("WQCRITID"))
                    Else
                        strCmdInsert = String.Format("INSERT INTO POLL_WQCRITERIA (PollID,WQCritID,Threshold) VALUES ('{0}', '{1}', {2})", datPollDetails.Item("POLLID"), datPollAdd.Item("WQCRITID"), intThresh)
                    End If
                    Using cmdInsert As New DataHelper(strCmdInsert)
                        cmdInsert.ExecuteNonQuery()
                    End Using
                End Using
            End Using

        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub
#End Region


End Class