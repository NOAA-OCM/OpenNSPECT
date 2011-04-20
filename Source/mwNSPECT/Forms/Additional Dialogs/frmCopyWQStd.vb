'********************************************************************************************************
'File Name: frmCopyWQStd.vb
'Description: Form for copying one water quality standard to a new name
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
Friend Class frmCopyWQStd
    Inherits System.Windows.Forms.Form

    Private _frmWQStd As frmWaterQualityStandard
    Const c_sModuleFileName As String = "frmcopyWQStd.vb"


    Private Sub frmCopyWQStd_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Try
            modUtil.InitComboBox(cboStdName, "WQCRITERIA")

        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
        Try
            Me.Close()
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click

        Try
            Dim strStandard As String

            Dim strPollStandard As String

            Dim strCmd As String
            Dim strCmd2 As String

            'Get the WQ stand info
            strStandard = "SELECT * FROM WQCriteria WHERE NAME LIKE '" & cboStdName.Text & "'"
            Dim cmdstd As New OleDbCommand(strStandard, g_DBConn)
            Dim datastd As OleDbDataReader = cmdstd.ExecuteReader()
            datastd.Read()

            'Get the related pollutant/thresholds
            strPollStandard = "SELECT * FROM POLL_WQCRITERIA WHERE WQCRITID =" & datastd("WQCRITID").ToString
            Dim cmdPoll As New OleDbCommand(strPollStandard, g_DBConn)
            Dim datapoll As OleDbDataReader = cmdPoll.ExecuteReader()

            strCmd = "INSERT INTO WQCRITERIA (NAME,DESCRIPTION) VALUES ('" & Replace(Trim(txtStdName.Text), "'", "''") & "', '" & datastd("Description") & "')"

            If modUtil.UniqueName("WQCRITERIA", Trim(txtStdName.Text)) Then
                Dim cmdIns As New OleDbCommand(strCmd, g_DBConn)
                cmdIns.ExecuteNonQuery()
            Else
                MsgBox(Err4, MsgBoxStyle.Critical, "Enter Unique Name")
                Exit Sub
            End If
            Dim cmdNewStandard As New OleDbCommand("Select * from WQCRITERIA WHERE NAME LIKE '" & Trim(txtStdName.Text) & "'", g_DBConn)
            Dim datanewstd As OleDbDataReader = cmdNewStandard.ExecuteReader()
            datanewstd.Read()

            Dim i As Short
            i = 0

            While datapoll.Read()
                strCmd2 = "INSERT INTO POLL_WQCRITERIA (POLLID, WQCRITID, THRESHOLD) VALUES (" & datapoll("POLLID") & ", " & datanewstd("WQCRITID") & "," & datapoll("Threshold") & ")"
                Dim cmdIns2 As New OleDbCommand(strCmd2, g_DBConn)
                cmdIns2.ExecuteNonQuery()
            End While
            datastd.Close()
            datapoll.Close()
            datanewstd.Close()

            _frmWQStd.UpdateWQ(Trim(txtStdName.Text))
            Me.Close()

        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Public Sub Init(ByRef frmWQ As frmWaterQualityStandard)
        Try
            _frmWQStd = frmWQ
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


End Class