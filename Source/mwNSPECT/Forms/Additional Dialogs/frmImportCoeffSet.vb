'********************************************************************************************************
'File Name: frmImportCoeffSet.vb
'Description: Form for importing coeffecient sets
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
Friend Class frmImportCoeffSet
    Inherits System.Windows.Forms.Form

    Const c_sModuleFileName As String = "frmImportCoeffSet.vb"
    Private _frmPoll As frmPollutants
    Private _cmdCoeff As OleDbCommand

    Private Sub frmImportCoeffSet_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        InitComboBox(cboLCType, "LCTYPE")
    End Sub

    Private Sub cmdBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowse.Click
        Try
            Dim dlgOpen As New Windows.Forms.OpenFileDialog
            'browse...get output filename
            dlgOpen.Filter = MSG1
            dlgOpen.Title = MSG2
            If dlgOpen.ShowDialog = Windows.Forms.DialogResult.OK Then
                txtImpFile.Text = Trim(dlgOpen.FileName)
            End If
        Catch ex As Exception
            HandleError(True, "cmdBrowse_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        Try
            If modUtil.UniqueName("CoefficientSet", (txtCoeffSetName.Text)) Then
                If ValidateCoeffTextFile(txtImpFile.Text, cboLCType.Text) Then
                    _frmPoll.UpdateCoeffSet(_cmdCoeff, txtCoeffSetName.Text, txtImpFile.Text)
                End If
            Else
                MsgBox("The name you have chosen is in use, please enter a different name.", MsgBoxStyle.Critical, "Name Detected")
                txtCoeffSetName.Focus()
            End If
            Me.Close()
        Catch ex As Exception
            HandleError(True, "cmdOK_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try
    End Sub


    Public Sub Init(ByRef frmPoll As frmPollutants)
        _frmPoll = frmPoll
    End Sub

    'Need to check the text file coming in from the import menu of the pollutant form.
    'Bringing the Text File itself, and the name of the LCType as picked by John User
    Private Function ValidateCoeffTextFile(ByRef strFileName As String, ByRef strLCTypeName As String) As Boolean
        Try

            Dim strLine As String
            Dim intLine As Short
            Dim strValue As String
            Dim strParams(7) As Object
            Dim strLCTypeNum As String
            Dim j As Short
            Dim dataLCType As OleDbDataReader

            ValidateCoeffTextFile = False

            'Gameplan is to find number of records(landclasses) in the chosen LCType. Then
            'compare that to the number of lines in the text file, and the [Value] field to
            'make sure both jive.  If not, bark at them...ruff, ruff

            strLCTypeNum = "SELECT LCTYPE.LCTYPEID, LCCLASS.NAME, LCCLASS.VALUE, LCCLASS.LCCLASSID FROM " & "LCTYPE INNER JOIN LCCLASS ON LCTYPE.LCTYPEID = LCCLASS.LCTYPEID " & "WHERE LCTYPE.NAME LIKE '" & strLCTypeName & "'"
            Dim cmdLCType As New OleDbCommand(strLCTypeNum, g_DBConn)
            If IO.File.Exists(strFileName) Then
                Dim read As New IO.StreamReader(strFileName)

                intLine = 0

                'CHECK 1: loop through the text file, compare value to [Value], make sure exists
                Do While Not read.EndOfStream

                    strLine = read.ReadLine
                    'Value exits??
                    strValue = Split(strLine, ",")(0)

                    j = 0

                    dataLCType = cmdLCType.ExecuteReader()
                    While dataLCType.Read()
                        If dataLCType("Value") = strValue Then
                            j = j + 1
                        End If
                    End While
                    dataLCType.Close()

                    If j = 0 Then
                        MsgBox("There is a value in your text file that does not exist in the Land Class Type: '" & strLCTypeName & "' Please check your text file in line: " & intLine + 1, MsgBoxStyle.OkOnly, "Data Import Error")
                    ElseIf j > 1 Then
                        MsgBox("There are records in your text file that contain the same value.  Please check line " & intLine, MsgBoxStyle.Critical, "Multiple values found")
                    ElseIf j = 1 Then
                        ValidateCoeffTextFile = True
                    End If

                    intLine = intLine + 1
                    Debug.Print(intLine)

                Loop

                dataLCType = cmdLCType.ExecuteReader()
                Dim iRows As Integer = 0
                While dataLCType.Read()
                    iRows = iRows + 1
                End While
                dataLCType.Close()

                'Final check, make sure same number of records in text file vs the
                If iRows = intLine Then
                    ValidateCoeffTextFile = True
                Else
                    MsgBox("The number of records in your import file do not match the number of records in the " & "Landclass '" & strLCTypeName & "'.  Your file should contain " & iRows & " records.", MsgBoxStyle.Critical, "Error Importing File")
                End If
            Else
                MsgBox("The file you are pointing to does not exist. Please select another.", MsgBoxStyle.Critical, "File Not Found")
                'Cleanup
            End If

            If ValidateCoeffTextFile Then
                _cmdCoeff = cmdLCType
            End If

        Catch ex As Exception
            HandleError(True, "ValidateCoeffTextFile " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, 0)
        End Try
    End Function
End Class