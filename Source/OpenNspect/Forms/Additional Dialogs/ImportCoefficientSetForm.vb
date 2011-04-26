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
Imports System.Windows.Forms
Imports System.IO

Friend Class ImportCoefficientSetForm
    Private _frmPoll As PollutantsForm
    Private _cmdCoeff As OleDbCommand

    Private Sub frmImportCoeffSet_Load (ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        Try
            InitComboBox (cboLCType, "LCTYPE")
        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub

    Private Sub cmdBrowse_Click (ByVal sender As Object, ByVal e As EventArgs) Handles cmdBrowse.Click
        Try
            Using dlgOpen As New OpenFileDialog()
                'browse...get output filename
                dlgOpen.Filter = MSG1TextFile
                dlgOpen.Title = MSG2
                If dlgOpen.ShowDialog = System.Windows.Forms.DialogResult.OK Then
                    txtImpFile.Text = Trim (dlgOpen.FileName)
                End If
            End Using
        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub

    Protected Overrides Sub OK_Button_Click (sender As Object, e As EventArgs)
        Try
            If UniqueName ("CoefficientSet", (txtCoeffSetName.Text)) Then
                If ValidateCoeffTextFile (txtImpFile.Text, cboLCType.Text) Then
                    _frmPoll.UpdateCoeffSet (_cmdCoeff, txtCoeffSetName.Text, txtImpFile.Text)
                End If
            Else
                MsgBox ("The name you have chosen is in use, please enter a different name.", MsgBoxStyle.Critical, _
                        "Name Detected")
                txtCoeffSetName.Focus()
            End If
            MyBase.OK_Button_Click (sender, e)
        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub

    Public Sub Init (ByRef frmPoll As PollutantsForm)
        _frmPoll = frmPoll
    End Sub

    'Need to check the text file coming in from the import menu of the pollutant form.
    'Bringing the Text File itself, and the name of the LCType as picked by John User

    Private Function ValidateCoeffTextFile (ByRef strFileName As String, ByRef strLCTypeName As String) As Boolean
        Try

            Dim strLine As String
            Dim intLine As Short
            Dim strValue As String
            Dim strLCTypeNum As String
            Dim j As Short
            Dim dataLCType As OleDbDataReader

            ValidateCoeffTextFile = False

            'Gameplan is to find number of records(landclasses) in the chosen LCType. Then
            'compare that to the number of lines in the text file, and the [Value] field to
            'make sure both jive.  If not, bark at them...ruff, ruff

            strLCTypeNum = _
                String.Format ( _
                               "SELECT LCTYPE.LCTYPEID, LCCLASS.NAME, LCCLASS.VALUE, LCCLASS.LCCLASSID FROM LCTYPE INNER JOIN LCCLASS ON LCTYPE.LCTYPEID = LCCLASS.LCTYPEID WHERE LCTYPE.NAME LIKE '{0}'", _
                               strLCTypeName)
            Dim cmdLCType As New DataHelper (strLCTypeNum)
            If File.Exists (strFileName) Then
                Using read As New StreamReader (strFileName)

                    intLine = 0

                    'CHECK 1: loop through the text file, compare value to [Value], make sure exists
                    Do While Not read.EndOfStream

                        strLine = read.ReadLine
                        'Value exits??
                        strValue = Split (strLine, ",") (0)

                        j = 0

                        dataLCType = cmdLCType.ExecuteReader()
                        While dataLCType.Read()
                            If dataLCType ("Value") = strValue Then
                                j = j + 1
                            End If
                        End While
                        dataLCType.Close()

                        If j = 0 Then
                            MsgBox ( _
                                    "There is a value in your text file that does not exist in the Land Class Type: '" & _
                                    strLCTypeName & "' Please check your text file in line: " & intLine + 1, _
                                    MsgBoxStyle.OkOnly, "Data Import Error")
                        ElseIf j > 1 Then
                            MsgBox ( _
                                    "There are records in your text file that contain the same value.  Please check line " & _
                                    intLine, MsgBoxStyle.Critical, "Multiple values found")
                        ElseIf j = 1 Then
                            ValidateCoeffTextFile = True
                        End If
                        intLine = intLine + 1
                        Debug.WriteLine (intLine)

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
                        MsgBox ( _
                                String.Format ( _
                                               "The number of records in your import file do not match the number of records in the Landclass '{0}'.  Your file should contain {1} records.", _
                                               strLCTypeName, iRows), MsgBoxStyle.Critical, "Error Importing File")
                    End If
                End Using

            Else
                MsgBox ("The file you are pointing to does not exist. Please select another.", MsgBoxStyle.Critical, _
                        "File Not Found")
                'Cleanup
            End If

            If ValidateCoeffTextFile Then
                _cmdCoeff = cmdLCType.GetCommand()
            End If

        Catch ex As Exception
            HandleError (ex)
        End Try
    End Function
End Class