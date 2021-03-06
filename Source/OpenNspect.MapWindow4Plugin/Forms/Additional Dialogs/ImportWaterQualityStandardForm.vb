'********************************************************************************************************
'File Name: frmImportWQStd.vb
'Description: Form for importing Water Quality standards
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
Imports System.IO
Imports System.Data.OleDb
Imports Alias1 = System.Windows.Forms

Friend Class ImportWaterQualityStandardForm
    Private _frmWQ As WaterQualityStandardsForm
    Private _strFileName As String

    Private Sub cmdBrowse_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdBrowse.Click
        Try
            'browse...get output filename
            Using dlgOpen As New OpenFileDialog() With {.Filter = MSG1TextFile, .Title = MSG2}
                If dlgOpen.ShowDialog = Alias1.DialogResult.OK Then
                    txtImpFile.Text = Trim(dlgOpen.FileName)
                    _strFileName = txtImpFile.Text
                    OK_Button.Enabled = True
                End If
            End Using

        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Protected Overrides Sub OK_Button_Click(sender As Object, e As EventArgs)

        Try
            Dim strLine As String
            Dim intLine As Short
            Dim strName As String = ""
            Dim strDescript As String
            Dim strPoll As String
            Dim strThresh As String
            Dim strCmd As String

            Using read As New StreamReader(_strFileName)
                intLine = 0
                Do While Not read.EndOfStream
                    strLine = read.ReadLine
                    intLine = intLine + 1
                    If intLine = 1 Then
                        strName = Trim(txtStdName.Text)
                        strDescript = Split(strLine, ",")(1)
                        If strName = "" Then
                            MsgBox("Name is blank.  Please enter a name.", MsgBoxStyle.Critical, "Empty Name Field")
                            txtStdName.Focus()
                            Return
                        Else
                            strCmd = String.Format("INSERT INTO WQCRITERIA (NAME,DESCRIPTION) VALUES ('{0}', '{1}')", Replace(txtStdName.Text, "'", "''"), Replace(strDescript, "'", "''"))
                            'Name Check
                            If UniqueName("WQCRITERIA", (txtStdName.Text)) Then
                                Using cmdIns As New DataHelper(strCmd)
                                    cmdIns.ExecuteNonQuery()
                                End Using
                            Else
                                MsgBox("The name you have chosen is already in use.  Please select another.", MsgBoxStyle.Critical, "Select Unique Name")
                                Return
                            End If
                        End If
                    Else
                        strPoll = Split(strLine, ",")(0)
                        strThresh = Split(strLine, ",")(1)
                        'Insert the pollutant/threshold value into POLL_WQCRITERIA
                        PollutantAdd(strName, strPoll, strThresh)
                    End If
                Loop
            End Using

            'Cleanup
            _frmWQ.cboWQStdName.Items.Clear()
            InitComboBox(_frmWQ.cboWQStdName, "WQCRITERIA")
            _frmWQ.cboWQStdName.SelectedIndex = GetIndexOfEntry((txtStdName.Text), _frmWQ.cboWQStdName)
            MyBase.OK_Button_Click(sender, e)
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Public Sub Init(ByRef frmWQ As WaterQualityStandardsForm)
        Try
            _frmWQ = frmWQ
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub PollutantAdd(ByRef strName As String, ByRef strPoll As String, ByRef strThresh As String)
        Try

            Dim strPollAdd As String
            Dim strPollDetails As String
            Dim strCmdInsert As String

            'Get the WQCriteria values using the name
            strPollAdd = "SELECT * FROM WQCriteria WHERE NAME = " & "'" & strName & "'"
            Dim cmdPollAdd As New DataHelper(strPollAdd)
            Dim datapolladd As OleDbDataReader = cmdPollAdd.ExecuteReader
            datapolladd.Read()

            'Get the pollutant particulars
            strPollDetails = "SELECT * FROM POLLUTANT WHERE NAME =" & "'" & strPoll & "'"
            Dim cmdPollDet As New DataHelper(strPollDetails)
            Dim datapolldet As OleDbDataReader = cmdPollDet.ExecuteReader
            datapolldet.Read()

            strCmdInsert = "INSERT INTO POLL_WQCRITERIA (PollID,WQCritID,Threshold) VALUES ('" & datapolldet("POLLID") & "', '" & datapolladd("WQCRITID") & "'," & strThresh & ")"
            Dim cmdIns As New DataHelper(strCmdInsert)
            cmdIns.ExecuteNonQuery()

            datapolladd.Close()
            datapolldet.Close()
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub
End Class