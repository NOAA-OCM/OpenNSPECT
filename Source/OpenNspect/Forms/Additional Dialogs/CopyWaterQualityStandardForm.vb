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
Friend Class CopyWaterQualityStandardForm
    Inherits System.Windows.Forms.Form

    Private _frmWQStd As WaterQualityStandardsForm
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
            Close()
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
            strStandard = String.Format("SELECT * FROM WQCriteria WHERE NAME LIKE '{0}'", cboStdName.Text)
            Using cmdstd As New DataHelper(strStandard)
                Using datastd As OleDbDataReader = cmdstd.ExecuteReader()
                    datastd.Read()
                    'Get the related pollutant/thresholds
                    strPollStandard = "SELECT * FROM POLL_WQCRITERIA WHERE WQCRITID =" & datastd("WQCRITID").ToString
                    Using datahelper As New DataHelper(strPollStandard)
                        Using datapoll = datahelper.ExecuteReader()
                            strCmd = String.Format("INSERT INTO WQCRITERIA (NAME,DESCRIPTION) VALUES ('{0}', '{1}')", Replace(Trim(txtStdName.Text), "'", "''"), datastd("Description"))
                            If modUtil.UniqueName("WQCRITERIA", Trim(txtStdName.Text)) Then
                                Using cmdIns As New DataHelper(strCmd)
                                    cmdIns.ExecuteNonQuery()
                                End Using
                            Else
                                MsgBox(Err4, MsgBoxStyle.Critical, "Enter Unique Name")
                                Exit Sub
                            End If
                            Using cmdNewStandard As New DataHelper(String.Format("Select * from WQCRITERIA WHERE NAME LIKE '{0}'", Trim(txtStdName.Text)))
                                Using datanewstd As OleDbDataReader = cmdNewStandard.ExecuteReader()
                                    datanewstd.Read()
                                    While datapoll.Read()
                                        strCmd2 = String.Format("INSERT INTO POLL_WQCRITERIA (POLLID, WQCRITID, THRESHOLD) VALUES ({0}, {1}, {2})", datapoll("POLLID"), datanewstd("WQCRITID"), datapoll("Threshold"))
                                        Using cmdIns2 As New DataHelper(strCmd2)
                                            cmdIns2.ExecuteNonQuery()
                                        End Using
                                    End While
                                    datastd.Close()
                                End Using
                            End Using
                        End Using
                    End Using
                End Using
            End Using

            _frmWQStd.UpdateWQ(Trim(txtStdName.Text))
            Close()

        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Public Sub Init(ByRef frmWQ As WaterQualityStandardsForm)
        Try
            _frmWQStd = frmWQ
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


End Class