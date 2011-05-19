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
Imports System.Windows.Forms
Imports System.IO
Imports System.Data

Friend Class ImportCoefficientSetForm
    Private _frmPoll As PollutantsForm

    Private Sub frmImportCoeffSet_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        Try
            InitComboBox(cboLCType, "LCTYPE")
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub cmdBrowse_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdBrowse.Click
        Try
            Using dlgOpen As New OpenFileDialog()
                'browse...get output filename
                dlgOpen.Filter = MSG1TextFile
                dlgOpen.Title = MSG2
                If dlgOpen.ShowDialog = System.Windows.Forms.DialogResult.OK Then
                    txtImpFile.Text = Trim(dlgOpen.FileName)
                End If
            End Using
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Protected Overrides Sub OK_Button_Click(sender As Object, e As EventArgs)
        Try
            If UniqueName("CoefficientSet", (txtCoeffSetName.Text)) Then
                If ValidateCoeffTextFile(txtImpFile.Text, cboLCType.Text) Then
                    _frmPoll.UpdateCoeffSet(cboLCType.Text, txtCoeffSetName.Text, txtImpFile.Text)
                End If
            Else
                MsgBox("The name you have chosen is in use, please enter a different name.", MsgBoxStyle.Critical, "Name Detected")
                txtCoeffSetName.Focus()
            End If
            MyBase.OK_Button_Click(sender, e)
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Public Sub Init(ByRef frmPoll As PollutantsForm)
        _frmPoll = frmPoll
    End Sub

    'Need to check the text file coming in from the import menu of the pollutant form.
    'Bringing the Text File itself, and the name of the LCType as picked by John User

    Private Function ValidateCoeffTextFile(ByRef fileName As String, ByRef landClassTypeName As String) As Boolean
        Try
            If Not File.Exists(fileName) Then
                MsgBox("The file you are pointing to does not exist. Please select another.", MsgBoxStyle.Critical, "File Not Found")
                Return False
            End If
            Dim strLCTypeNum As String = String.Format("SELECT LCTYPE.LCTYPEID, LCCLASS.NAME, LCCLASS.VALUE, LCCLASS.LCCLASSID FROM LCTYPE INNER JOIN LCCLASS ON LCTYPE.LCTYPEID = LCCLASS.LCTYPEID WHERE LCTYPE.NAME LIKE '{0}'", landClassTypeName)
            'find number of records(landclasses) in the chosen LCType. Then
            'compare that to the number of lines in the text file, and the [Value] field to
            'make sure both jive.  If not, bark at them...ruff, ruff
            Using data As New DataHelper(strLCTypeNum)

                Using streamReader As New StreamReader(fileName)

                    Dim intLine As Short

                    'CHECK 1: loop through the text file, compare value to [Value], make sure exists exactly once
                    Using records As New DataTable()
                        data.GetAdapter.Fill(records)

                        Do While Not streamReader.EndOfStream

                            Dim fileValue As String = Split(streamReader.ReadLine, ",")(0)

                            Dim timesValueAppears As Short = 0

                            For Each row As DataRow In records.Rows
                                If row("Value") = fileValue Then
                                    timesValueAppears += 1
                                End If
                            Next

                            If timesValueAppears = 0 Then
                                MsgBox("There is a value in your text file that does not exist in the Land Class Type: '" & landClassTypeName & "' Please check your text file in line: " & intLine + 1, MsgBoxStyle.OkOnly, "Data Import Error")
                                Return False
                            ElseIf timesValueAppears > 1 Then
                                MsgBox("There are records in your text file that contain the same value.  Please check line " & intLine, MsgBoxStyle.Critical, "Multiple values found")
                                Return False
                            End If
                            intLine = intLine + 1

                        Loop

                    End Using

                    Using dataLCType2 = data.ExecuteReader()
                        Dim iRows As Integer = 0
                        While dataLCType2.Read()
                            iRows = iRows + 1
                        End While

                        'Final check, make sure same number of records in text file vs the
                        If iRows <> intLine Then
                            MsgBox(String.Format("The number of records in your import file do not match the number of records in the Landclass '{0}'.  Your file should contain {1} records.", landClassTypeName, iRows), MsgBoxStyle.Critical, "Error Importing File")
                            Return False
                        End If
                    End Using
                End Using
            End Using
            Return True
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Function
End Class