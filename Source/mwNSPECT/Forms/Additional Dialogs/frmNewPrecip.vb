Imports System.Data.OleDb
Friend Class frmNewPrecip
    Inherits System.Windows.Forms.Form

    Private _pInputPrecipDS As String
    Private _boolChange As Boolean
    Private _frmPrj As frmProjectSetup
    Private _frmPrec As frmPrecipitation


#Region "Events"
    Private Sub frmNewPrecip_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub txtPrecipName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPrecipName.TextChanged
        txtPrecipName.Text = Replace(txtPrecipName.Text, "'", "")
    End Sub

    Private Sub txtDesc_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDesc.TextChanged

    End Sub

    Private Sub txtDesc_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtDesc.Validating
        Dim Cancel As Boolean = e.Cancel

        txtDesc.Text = Replace(txtDesc.Text, "'", "")

        e.Cancel = Cancel
    End Sub

    Private Sub txtPrecipFile_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPrecipFile.TextChanged

    End Sub

    Private Sub cmdBrowseFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowseFile.Click
        Dim dlgOpen As New Windows.Forms.OpenFileDialog

        Dim g As New MapWinGIS.Grid

        dlgOpen.Filter = g.CdlgFilter

        If dlgOpen.ShowDialog = Windows.Forms.DialogResult.OK Then
            If g.Open(dlgOpen.FileName) Then
                txtPrecipFile.Text = dlgOpen.FileName
                Dim proj As String = g.Header.Projection
                If proj = "" Then
                    MsgBox("The GRID you have choosen has no spatial reference information.  Please define a projection before continuing.", MsgBoxStyle.Exclamation, "No Project Information Detected")
                    Exit Sub
                Else
                    If proj.Contains("units=m") Then
                        cboGridUnits.SelectedIndex = 0
                    Else
                        cboGridUnits.SelectedIndex = 1
                    End If

                End If
            End If
        End If


    End Sub

    Private Sub cboGridUnits_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboGridUnits.SelectedIndexChanged

    End Sub

    Private Sub cboPrecipUnits_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPrecipUnits.SelectedIndexChanged

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

    End Sub

    Private Sub cboPrecipType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPrecipType.SelectedIndexChanged

    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click

        Dim intSave As Object

        If _boolChange Then
            intSave = MsgBox("You have made changes to this record, are you sure you want to quit?", MsgBoxStyle.YesNo, "Quit?")

            If intSave = MsgBoxResult.Yes Then
                If Not _frmPrj Is Nothing Then
                    _frmPrj.cboPrecipScen.SelectedIndex = 0
                End If
                Me.Close()
            ElseIf intSave = MsgBoxResult.No Then

                Exit Sub
            End If
        Else
            Me.Close()
        End If
    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        Dim intType As Short
        Dim intRainingDays As Short


        Dim strCmdInsert As String
        If CheckParams() Then
            'Process the time period
            intType = cboTimePeriod.SelectedIndex
            If intType = 0 Then
                intRainingDays = CShort(txtRainingDays.Text)
            Else
                intRainingDays = 0
            End If


            'Compose the INSERT statement.
            strCmdInsert = "INSERT INTO PrecipScenario " & "(Name, Description, PrecipFileName, PrecipGridUnits, PrecipUnits, Type, PrecipType, RainingDays) VALUES (" & "'" & Replace(CStr(txtPrecipName.Text), "'", "''") & "', " & "'" & Replace(CStr(txtDesc.Text), "'", "''") & "', " & "'" & Replace(txtPrecipFile.Text, "'", "''") & "', " & "" & cboGridUnits.SelectedIndex & ", " & "" & cboPrecipUnits.SelectedIndex & ", " & "" & intType & ", " & "" & cboPrecipType.SelectedIndex & ", " & "" & intRainingDays & ")"


            Debug.Print(strCmdInsert)

            If modUtil.UniqueName("PrecipScenario", txtPrecipName.Text) Then
                'Execute the statement.

                Dim cmdInsert As New OleDbCommand(strCmdInsert, g_DBConn)
                cmdInsert.ExecuteNonQuery()

                'Confirm
                MsgBox(txtPrecipName.Text & " successfully added.", MsgBoxStyle.OkOnly, "Record Added")



                If Not _frmPrj Is Nothing Then
                    _frmPrj.UpdatePrecip(txtPrecipName.Text)
                    Me.Close()
                End If

                If Not _frmPrec Is Nothing Then
                    _frmPrec.UpdatePrecip(txtPrecipName.Text)
                    Me.Close()
                End If

            Else
                MsgBox("Name already in use.  Please choose a different one.", MsgBoxStyle.Critical, "Name In Use")
                txtPrecipName.Focus()
                Exit Sub
            End If


        End If
    End Sub

#End Region

#Region "Helper Functions"
    Public Sub Init(ByRef frmPrj As frmProjectSetup, ByRef frmPrec As frmPrecipitation)
        _frmPrj = frmPrj
        _frmPrec = frmPrec
    End Sub


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
#End Region

End Class