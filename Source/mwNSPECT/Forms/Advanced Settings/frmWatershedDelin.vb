'********************************************************************************************************
'File Name: frmWatershedDelin.vb
'Description: Form for displaying and editing watershed delineation data
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
Friend Class frmWatershedDelin
    Inherits System.Windows.Forms.Form

    Private Const c_sModuleFileName As String = "frmWatershedDelin.vb"

#Region "Events"






    Private Sub frmWatershedDelin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            modUtil.InitComboBox(cboWSDelin, "WSDELINEATION")
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub







    Private Sub cboWSDelin_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboWSDelin.SelectedIndexChanged
        Try
            'String and recordset
            Dim strSQLDelin As String = "SELECT * FROM WSDELINEATION WHERE NAME LIKE '" & cboWSDelin.Text & "'"
            Dim delinCmd As New OleDbCommand(strSQLDelin, modUtil.g_DBConn)
            Dim delin As OleDbDataReader = delinCmd.ExecuteReader()

            'Check for records
            If delin.HasRows Then
                delin.Read()
                'Populate the controls...
                txtDEMFile.Text = delin.Item("DEMFileName")
                cboDEMUnits.SelectedIndex = delin.Item("DEMGridUnits")
                txtStream.Text = delin.Item("StreamFileName") & ""
                chkHydroCorr.CheckState = delin.Item("HydroCorrected")
                cboWSSize.SelectedIndex = delin.Item("SubWSSize")
                txtWSFile.Text = delin.Item("wsfilename") & ""
                txtFlowAccumGrid.Text = delin.Item("FlowAccumFileName") & ""
                txtLSGrid.Text = delin.Item("LSFileName") & ""
            Else
                MsgBox("Warning: There are no watershed delineation scenarios remaining.  Please add a new one.", MsgBoxStyle.Critical, "Recordset Empty")
            End If
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub







    Private Sub cmdQuit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdQuit.Click
        Try
            Me.Close()
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub







    Private Sub mnuNewWSDelin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuNewWSDelin.Click
        Try
            Dim newWS As New frmNewWSDelin
            newWS.Init(Me, Nothing)
            newWS.ShowDialog()
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub







    Private Sub mnuNewExist_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuNewExist.Click
        Try
            Dim newWS As New frmUserWShed
            newWS.Init(Me, Nothing)
            newWS.ShowDialog()
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub







    Private Sub mnuDelWSDelin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDelWSDelin.Click
        Try
            Dim intAns As Object
            Dim strSQLWSDel As String
            Dim strFolder As String

            strSQLWSDel = "DELETE FROM WSDELINEATION WHERE NAME LIKE '" & cboWSDelin.Text & "'"

            If Not (cboWSDelin.Text = "") Then
                intAns = MsgBox("Are you sure you want to delete the watershed delineation scenario '" & cboWSDelin.SelectedItem & "'?", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Confirm Delete")
                'code to handle response
                If intAns = MsgBoxResult.Yes Then

                    'Set up a delete rs and get rid of it
                    Dim cmdDel As New OleDbCommand(strSQLWSDel, g_DBConn)
                    cmdDel.ExecuteNonQuery()

                    strFolder = modUtil.g_nspectPath & "\wsdelin\" & cboWSDelin.Text
                    If IO.Directory.Exists(strFolder) Then
                        IO.Directory.Delete(strFolder, True)
                    End If

                    'Confirm
                    MsgBox(cboWSDelin.SelectedItem & " deleted.", MsgBoxStyle.OkOnly, "Record Deleted")

                    'Clear everything, clean up form
                    cboWSDelin.Items.Clear()
                    chkHydroCorr.CheckState = System.Windows.Forms.CheckState.Unchecked
                    txtDEMFile.Text = ""
                    txtStream.Text = ""
                    txtWSFile.Text = ""
                    txtFlowAccumGrid.Text = ""

                    modUtil.InitComboBox(cboWSDelin, "WSDELINEATION")

                    Me.Refresh()

                ElseIf intAns = MsgBoxResult.No Then
                    Exit Sub
                End If
            Else
                MsgBox("Please select a watershed delineation", MsgBoxStyle.Critical, "No Scenario Selected")
            End If
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub







    Private Sub mnuWSDelin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuWSDelin.Click
        Try
            System.Windows.Forms.Help.ShowHelp(Me, modUtil.g_nspectPath & "\Help\nspect.chm", "wsdelin.htm")
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub
#End Region

#Region "Helper Functions"

#End Region


End Class