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

Imports System.Data.OleDb
Friend Class ImportWaterQualityStandardForm
    Inherits System.Windows.Forms.Form

    Const c_sModuleFileName As String = "frmImportWQStd.vb"
    Private _frmWQ As WaterQualityStandardsForm
    Private _strFileName As String


    Private Sub cmdBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowse.Click
        Try
            'browse...get output filename
            Using dlgOpen As New Windows.Forms.OpenFileDialog() With {.Filter = MSG1TextFile, .Title = MSG2}
                If dlgOpen.ShowDialog = Windows.Forms.DialogResult.OK Then
                    txtImpFile.Text = Trim(dlgOpen.FileName)
                    _strFileName = txtImpFile.Text
                    cmdOK.Enabled = True
                End If
            End Using

        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Try
            Close()
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        Try


            Dim strLine As String
            Dim intLine As Short
            Dim strName As String = ""
            Dim strDescript As String
            Dim strPoll As String
            Dim strThresh As String
            Dim strCmd As String

            Dim read As New IO.StreamReader(_strFileName)

            intLine = 0

            Do While Not read.EndOfStream
                strLine = read.ReadLine
                intLine = intLine + 1
                'MsgBox theLine

                If intLine = 1 Then

                    strName = Trim(txtStdName.Text)
                    strDescript = Split(strLine, ",")(1)

                    If strName = "" Then

                        MsgBox("Name is blank.  Please enter a name.", MsgBoxStyle.Critical, "Empty Name Field")
                        txtStdName.Focus()
                        Exit Sub

                    Else

                        strCmd = "INSERT INTO WQCRITERIA (NAME,DESCRIPTION) VALUES ('" & Replace(txtStdName.Text, "'", "''") & "', '" & Replace(strDescript, "'", "''") & "')"
                        'Name Check
                        If modUtil.UniqueName("WQCRITERIA", (txtStdName.Text)) Then
                            Dim cmdIns As New DataHelper(strCmd)
                            cmdIns.ExecuteNonQuery()
                        Else
                            MsgBox("The name you have chosen is already in use.  Please select another.", MsgBoxStyle.Critical, "Select Unique Name")
                            Exit Sub
                        End If

                    End If

                Else

                    strPoll = Split(strLine, ",")(0)
                    strThresh = Split(strLine, ",")(1)
                    'Insert the pollutant/threshold value into POLL_WQCRITERIA
                    PollutantAdd(strName, strPoll, strThresh)

                End If

            Loop

            read.Close()

            'Cleanup
            _frmWQ.cboWQStdName.Items.Clear()
            modUtil.InitComboBox(_frmWQ.cboWQStdName, "WQCRITERIA")
            _frmWQ.cboWQStdName.SelectedIndex = modUtil.GetCboIndex((txtStdName.Text), _frmWQ.cboWQStdName)
            Close()
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Public Sub Init(ByRef frmWQ As WaterQualityStandardsForm)
        Try
            _frmWQ = frmWQ
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
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
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub
End Class