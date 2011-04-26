'********************************************************************************************************
'File Name: frmNewPrecip.vb
'Description: Form for new precip data entry
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
Friend Class NewPrecipitationScenarioForm


    Private _frmPrj As MainForm
    Private _frmPrec As PrecipitationScenariosForm


#Region "Events"


    Private Sub txtPrecipName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPrecipName.TextChanged
        Try
            txtPrecipName.Text = Replace(txtPrecipName.Text, "'", "")
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub


    Private Sub txtDesc_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtDesc.Validating
        Try
            Dim Cancel As Boolean = e.Cancel

            txtDesc.Text = Replace(txtDesc.Text, "'", "")

            e.Cancel = Cancel
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub


    Private Sub cmdBrowseFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowseFile.Click
        Try
            Dim dlgOpen As New Windows.Forms.OpenFileDialog

            Dim g As New MapWinGIS.Grid

            dlgOpen.Filter = g.CdlgFilter

            If dlgOpen.ShowDialog = Windows.Forms.DialogResult.OK Then
                If g.Open(dlgOpen.FileName) Then
                    txtPrecipFile.Text = dlgOpen.FileName
                    Dim proj As String = g.Header.Projection
                    If IO.Path.GetFileName(dlgOpen.FileName) = "sta.adf" Then
                        If IO.File.Exists(IO.Path.GetDirectoryName(dlgOpen.FileName) + IO.Path.DirectorySeparatorChar + "prj.adf") Then
                            Dim infile As New IO.StreamReader(IO.Path.GetDirectoryName(dlgOpen.FileName) + IO.Path.DirectorySeparatorChar + "prj.adf")
                            If infile.ReadToEnd.Contains("METERS") Then
                                proj = "units=m"
                            End If
                        Else
                            MsgBox("The GRID you have choosen has no spatial reference information.  Please define a projection before continuing.", MsgBoxStyle.Exclamation, "No Project Information Detected")
                            Exit Sub
                        End If
                    End If
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


        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub


    Private Sub cboTimePeriod_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTimePeriod.SelectedIndexChanged
        Try
            If cboTimePeriod.SelectedIndex = 0 Then
                lblRainingDays.Visible = True
                txtRainingDays.Visible = True
            Else
                lblRainingDays.Visible = False
                txtRainingDays.Visible = False
            End If
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub
    Protected Overrides Sub Cancel_Button_Click(sender As Object, e As System.EventArgs)
        IsDirty = True
        MyBase.Cancel_Button_Click(sender, e)
        If Not _frmPrj Is Nothing Then
            _frmPrj.cboPrecipScen.SelectedIndex = 0
        End If
    End Sub
    Protected Overrides Sub OK_Button_Click(sender As Object, e As System.EventArgs)

        Try
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

                If modUtil.UniqueName("PrecipScenario", txtPrecipName.Text) Then
                    'Execute the statement.

                    Dim cmdInsert As New DataHelper(strCmdInsert)
                    cmdInsert.ExecuteNonQuery()

                    'Confirm
                    MsgBox(txtPrecipName.Text & " successfully added.", MsgBoxStyle.OkOnly, "Record Added")

                    If Not _frmPrec Is Nothing Then
                        _frmPrec.UpdatePrecip(txtPrecipName.Text)
                    End If

                    MyBase.OK_Button_Click(sender, e)

                Else
                    MsgBox("Name already in use.  Please choose a different one.", MsgBoxStyle.Critical, "Name In Use")
                    txtPrecipName.Focus()
                    Exit Sub
                End If

            End If
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

#End Region

#Region "Helper Functions"


    Public Sub Init(ByRef frmPrj As MainForm, ByRef frmPrec As PrecipitationScenariosForm)
        Try
            _frmPrj = frmPrj
            _frmPrec = frmPrec
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub


    Private Function CheckParams() As Boolean
        Try
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


        Catch ex As Exception
            HandleError(ex)
        End Try
    End Function
#End Region

End Class