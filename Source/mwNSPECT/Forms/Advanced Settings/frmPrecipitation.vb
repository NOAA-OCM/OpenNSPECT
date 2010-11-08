'********************************************************************************************************
'File Name: frmPrecipitation.vb
'Description: Form for displaying Precipitation data
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

Friend Class frmPrecipitation
    Inherits System.Windows.Forms.Form

    Private _boolChange As Boolean
    Private _boolLoad As Boolean

    Private _pInputPrecipDS As MapWinGIS.Grid

#Region "Events"

    Private Sub frmPrecipitation_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _boolLoad = True
        modUtil.InitComboBox(cboScenName, "PRECIPSCENARIO")
        cmdSave.Enabled = False
        _boolChange = False
        _boolLoad = False
    End Sub

    Private Sub cboScenName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboScenName.SelectedIndexChanged

        Dim strSQLPrecip As String
        strSQLPrecip = "SELECT * FROM PRECIPSCENARIO WHERE NAME LIKE '" & cboScenName.Text & "'"

        Dim precipCmd As New OleDbCommand(strSQLPrecip, modUtil.g_DBConn)
        Dim precip As OleDbDataReader = precipCmd.ExecuteReader()
        precip.Read()
        'Populate the controls...
        txtDesc.Text = precip.Item("Description")
        txtPrecipFile.Text = precip.Item("PrecipFileName")

        cboGridUnits.SelectedIndex = CShort(precip.Item("PrecipGridUnits"))
        cboPrecipUnits.SelectedIndex = CShort(precip.Item("PrecipUnits"))
        cboTimePeriod.SelectedIndex = precip.Item("Type")
        cboPrecipType.SelectedIndex = precip.Item("PrecipType")

        If precip.Item("Type") = 0 Then
            txtRainingDays.Text = precip.Item("RainingDays")
        End If

        cmdSave.Enabled = False
    End Sub

    Private Sub txtDesc_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDesc.TextChanged

        EnableSave()
        txtDesc.Text = Replace(txtDesc.Text, "'", "")
    End Sub

    Private Sub txtPrecipFile_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPrecipFile.TextChanged

        EnableSave()
    End Sub

    Private Sub cmdBrowseFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowseFile.Click
        Dim pPrecipRasterDataset As MapWinGIS.Grid
        Dim strProj As String

        _pInputPrecipDS = modUtil.AddInputFromGxBrowserText(txtPrecipFile, "Choose Precipitation GRID", Me, 0)

        If _pInputPrecipDS Is Nothing Then
            Exit Sub
        Else

            pPrecipRasterDataset = _pInputPrecipDS
            strProj = CheckSpatialReference(pPrecipRasterDataset)
            If strProj = "" Then

                MsgBox("The GRID you have choosen has no spatial reference information.  Please define a projection before continuing.", MsgBoxStyle.Exclamation, "No Project Information Detected")
                txtPrecipFile.Text = ""
                Exit Sub

            Else
                If strProj.ToLower.Contains("units=m") Then
                    cboPrecipUnits.SelectedIndex = 0
                Else
                    cboPrecipUnits.SelectedIndex = 1
                End If

                cboPrecipUnits.Refresh()
            End If

        End If

    End Sub

    Private Sub cboGridUnits_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboGridUnits.SelectedIndexChanged
        EnableSave()
    End Sub

    Private Sub cboPrecipUnits_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPrecipUnits.SelectedIndexChanged
        EnableSave()
    End Sub

    Private Sub cboTimePeriod_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTimePeriod.SelectedIndexChanged

        If cboTimePeriod.SelectedIndex = 0 Then
            lblRainingDays.Visible = True
            txtRainingDays.Visible = True
        Else
            lblRainingDays.Visible = False
            txtRainingDays.Visible = False
        End If

    End Sub

    Private Sub txtRainingDays_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRainingDays.TextChanged

        EnableSave()

    End Sub

    Private Sub cboPrecipType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPrecipType.SelectedIndexChanged
        If Not _boolLoad Then
            EnableSave()
        End If
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click

        If CheckParams() = True Then
            Dim strSQLPrecip As String = "SELECT * FROM PRECIPSCENARIO WHERE NAME LIKE '" & cboScenName.Text & "'"

            Dim precipCmd As New OleDbCommand(strSQLPrecip, modUtil.g_DBConn)
            Dim precipAdapter As New OleDbDataAdapter(precipCmd)
            Dim cBuilder As New OleDbCommandBuilder(precipAdapter)
            cBuilder.QuotePrefix = "["
            cBuilder.QuoteSuffix = "]"
            Dim dt As New Data.DataTable
            precipAdapter.Fill(dt)

            dt.Rows(0)("Name") = cboScenName.Text
            dt.Rows(0)("Description") = txtDesc.Text
            dt.Rows(0)("PrecipFileName") = txtPrecipFile.Text
            dt.Rows(0)("PrecipGridUnits") = cboGridUnits.SelectedIndex
            dt.Rows(0)("PrecipUnits") = cboPrecipUnits.SelectedIndex
            dt.Rows(0)("PrecipType") = cboPrecipType.SelectedIndex
            dt.Rows(0)("Type") = cboTimePeriod.SelectedIndex

            If cboTimePeriod.SelectedIndex = 0 Then
                dt.Rows(0)("RainingDays") = CShort(txtRainingDays.Text)
            Else
                dt.Rows(0)("RainingDays") = 0
            End If
            precipAdapter.Update(dt)

            _boolChange = False

            MsgBox(cboScenName.Text & " saved successfully.", MsgBoxStyle.OkOnly, "Record Saved")
            Me.Close()
        Else
            Exit Sub

        End If
    End Sub

    Private Sub cmdQuit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdQuit.Click

        Dim intSave As Object

        If _boolChange Then
            intSave = MsgBox("You have made changes to this record, are you sure you want to quit?", MsgBoxStyle.YesNo, "Quit?")

            If intSave = MsgBoxResult.Yes Then
                Me.Close()
            ElseIf intSave = MsgBoxResult.No Then
                Exit Sub
            End If
        Else
            Me.Close()
        End If

    End Sub

    Private Sub mnuNewPrecip_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuNewPrecip.Click
        Dim newpre As New frmNewPrecip
        newpre.Init(Nothing, Me)
        newpre.ShowDialog()
    End Sub

    Private Sub mnuDelPrecip_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDelPrecip.Click

        Dim intAns As Object
        Dim strSQLPrecipDel As String
        'Dim cntrl As System.Windows.Forms.Control

        strSQLPrecipDel = "DELETE FROM PRECIPSCENARIO WHERE NAME LIKE '" & cboScenName.Text & "'"

        If Not (cboScenName.Text = "") Then
            intAns = MsgBox("Are you sure you want to delete the precipitation scenario '" & cboScenName.SelectedItem & "'?", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Confirm Delete")
            'code to handle response
            If intAns = MsgBoxResult.Yes Then
                'Set up a delete rs and get rid of it
                Dim cmdDel As New OleDbCommand(strSQLPrecipDel, g_DBConn)
                cmdDel.ExecuteNonQuery()

                MsgBox(cboScenName.SelectedItem & " deleted.", MsgBoxStyle.OkOnly, "Record Deleted")

                'Clear everything, clean up form
                cboScenName.Items.Clear()

                txtDesc.Text = ""
                'txtDuration.Text = ""
                txtPrecipFile.Text = ""

                modUtil.InitComboBox(cboScenName, "PRECIPSCENARIO")

                Me.Refresh()

            ElseIf intAns = MsgBoxResult.No Then
                Exit Sub
            End If
        Else
            MsgBox("Please select a Precipitation Scenario", MsgBoxStyle.Critical, "No Scenario Selected")
        End If
    End Sub

    Private Sub mnuPrecipHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPrecipHelp.Click
        System.Windows.Forms.Help.ShowHelp(Me, modUtil.g_nspectPath & "\Help\nspect.chm", "precip.htm")
    End Sub

    Private Sub cboScenName_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboTimePeriod.KeyDown, cboScenName.KeyDown, cboPrecipUnits.KeyDown, cboGridUnits.KeyDown
        e.SuppressKeyPress = True
    End Sub
#End Region

#Region "Helper Functions"
    Private Function CheckParams() As Boolean

        'Check the inputs of the form, before saving
        If Len(txtDesc.Text) = 0 Then
            MsgBox("Please enter a description for this scenario", MsgBoxStyle.Critical, "Description Missing")
            txtDesc.Focus()
            CheckParams = False
            Exit Function
        End If

        If txtPrecipFile.Text = " " Or txtPrecipFile.Text = "" Then
            MsgBox("Please select a valid precipitation GRID before saving.", MsgBoxStyle.Critical, "GRID Missing")
            txtPrecipFile.Focus()
            CheckParams = False
            Exit Function
        End If

        If cboGridUnits.Text = "" Then
            MsgBox("Please select GRID units.", MsgBoxStyle.Critical, "Units Missing")
            cboGridUnits.Focus()
            CheckParams = False
            Exit Function
        End If

        If cboPrecipUnits.Text = "" Then
            MsgBox("Please select precipitation units.", MsgBoxStyle.Critical, "Units Missing")
            cboPrecipUnits.Focus()
            CheckParams = False
            Exit Function
        End If

        If Len(cboPrecipType.Text) = 0 Then
            MsgBox("Please select a Precipitation Type.", MsgBoxStyle.Critical, "Precipitation Type Missing")
            cboPrecipType.Focus()
            CheckParams = False
            Exit Function
        End If

        If Len(cboTimePeriod.Text) = 0 Then
            MsgBox("Please select a Time Period.", MsgBoxStyle.Critical, "Precipitation Time Period Missing")
            cboTimePeriod.Focus()
            CheckParams = False
            Exit Function
        End If

        If cboTimePeriod.SelectedIndex = 0 Then
            If Not IsNumeric(txtRainingDays.Text) Or Len(txtRainingDays.Text) = 0 Then
                MsgBox("Please enter a numeric value for Raining Days.", MsgBoxStyle.Critical, "Raining Days Value Incorrect")
                txtRainingDays.Focus()
                CheckParams = False
                Exit Function
            End If
        End If

        'if it got through all that, then set it to true
        CheckParams = True


    End Function


    Private Sub EnableSave()

        cmdSave.Enabled = True
        _boolChange = True

    End Sub


    Public Sub UpdatePrecip(ByVal strPrecName As String)
        cboScenName.Items.Clear()
        modUtil.InitComboBox(cboScenName, "PrecipScenario")
        cboScenName.SelectedIndex = modUtil.GetCboIndex(strPrecName, cboScenName)
    End Sub

#End Region


End Class