'********************************************************************************************************
'File Name: frmSoils.vb
'Description: Form for displaying soils data
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
Friend Class frmSoils
    Inherits System.Windows.Forms.Form

#Region "Events"

    Private Sub frmSoils_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        modUtil.InitComboBox(cboSoils, "SOILS")
    End Sub

    Private Sub cboSoils_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSoils.SelectedIndexChanged

        Dim strSQLSoils As String = "SELECT * FROM SOILS WHERE NAME LIKE '" & cboSoils.Text & "'"
        Dim soilCmd As New OleDbCommand(strSQLSoils, modUtil.g_DBConn)
        Dim soil As OleDbDataReader = soilCmd.ExecuteReader()
        If soil.HasRows Then
            soil.Read()
            'Populate the controls...
            txtSoilsGrid.Text = soil.Item("SoilsFileName")
            txtSoilsKGrid.Text = soil.Item("SoilsKFileName")
        End If
        soil.Close()
    End Sub

    Private Sub txtSoilsGrid_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSoilsGrid.TextChanged

    End Sub

    Private Sub txtSoilsKGrid_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSoilsKGrid.TextChanged

    End Sub

    Private Sub cmdQuit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdQuit.Click
        Me.Close()
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Me.Close()
    End Sub

    Private Sub mnuNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuNew.Click
        Dim newsoil As New frmSoilsSetup
        newsoil.Init(Me)
        newsoil.ShowDialog()
    End Sub

    Private Sub mnuDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDelete.Click

        Dim intAns As Object
        Dim strSQLSoilsDel As String
        'Dim cntrl As System.Windows.Forms.Control

        strSQLSoilsDel = "DELETE FROM SOILS WHERE NAME LIKE '" & cboSoils.Text & "'"

        If Not (cboSoils.Text = "") Then
            intAns = MsgBox("Are you sure you want to delete the soils setup '" & cboSoils.SelectedItem & "'?", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Confirm Delete")
            'code to handle response
            If intAns = MsgBoxResult.Yes Then

                'Set up a delete rs and get rid of it
                Dim cmdDel As New OleDbCommand(strSQLSoilsDel, g_DBConn)
                cmdDel.ExecuteNonQuery()

                MsgBox(cboSoils.SelectedItem & " deleted.", MsgBoxStyle.OkOnly, "Record Deleted")

                'Clear everything, clean up form
                cboSoils.Items.Clear()

                txtSoilsGrid.Text = ""
                txtSoilsKGrid.Text = ""

                modUtil.InitComboBox(cboSoils, "SOILS")

                Me.Refresh()

            ElseIf intAns = MsgBoxResult.No Then
                Exit Sub
            End If
        Else
            MsgBox("Please select a Soils Setup", MsgBoxStyle.Critical, "No Soils Setup Selected")
        End If
    End Sub

    Private Sub mnuSoilsHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSoilsHelp.Click
        System.Windows.Forms.Help.ShowHelp(Me, modUtil.g_nspectPath & "\Help\nspect.chm", "soils.htm")
    End Sub

#End Region

#Region "Helper Functions"

#End Region


End Class