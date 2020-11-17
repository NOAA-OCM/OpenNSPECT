'********************************************************************************************************
'File Name: frmImportLCType.vb
'Description: Form for importing Land Cover types
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
Imports System.Data.OleDb

Friend Class ImportLandCoverTypeForm
    Private _booName As Boolean
    'Check if user put a name in
    Private _booFile As Boolean
    'Check if FileName is correct
    Private _strFileName As String
    Private _parent As LandCoverTypesForm

    Public Sub Init(ByRef parent As LandCoverTypesForm)
        Try
            _parent = parent
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub txtLCType_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles txtLCType.TextChanged
        Try
            If CBool(Trim(CStr(Len(txtLCType.Text)))) Then
                _booName = True
            End If

            If _booFile And _booName Then
                OK_Button.Enabled = True
            End If
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub cmdBrowse_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdBrowse.Click
        Try
            Using dlgopen As New OpenFileDialog() With {
                .Filter = Replace(MSG1TextFile, "<name>", "Land Cover Type"), .Title = Replace(MSG2, "<name>", "Land Cover Type")}
                If dlgopen.ShowDialog = System.Windows.Forms.DialogResult.OK Then
                    txtImpFile.Text = dlgopen.FileName
                    _strFileName = dlgopen.FileName
                    _booFile = True
                End If
            End Using

            If _booFile And _booName Then
                OK_Button.Enabled = True
            End If
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Protected Overrides Sub OK_Button_Click(sender As Object, e As EventArgs)
        Try
            If File.Exists(_strFileName) Then
                Using read As New StreamReader(_strFileName)
                    Dim strline As String = ""
                    Dim strname As String = ""
                    Dim strDesc As String = ""
                    Dim line As Integer = 0
                    While Not read.EndOfStream
                        strline = read.ReadLine
                        line = line + 1
                        If line = 1 Then
                            strname = txtLCType.Text.Trim()
                            strDesc = strline.Split(",")(1)
                            If strname = "" Then
                                MsgBox("Name is blank.  Please enter a name.", MsgBoxStyle.Critical, "Empty Name Field")
                                txtLCType.Focus()
                                Return
                            Else
                                'Name Check, if cool perform
                                If UniqueName("LCTYPE", txtLCType.Text) Then
                                    Dim strCmd As String = String.Format("INSERT INTO LCTYPE (NAME,DESCRIPTION) VALUES ('{0}', '{1}')", Replace(txtLCType.Text, "'", "''"), Replace(strDesc, "'", "''"))
                                    Using cmdIns As New DataHelper(strCmd)
                                        cmdIns.ExecuteNonQuery()
                                    End Using
                                Else
                                    MsgBox("The name you have chosen is already in use.  Please select another.", MsgBoxStyle.Critical, "Select Unique Name")
                                    Return
                                End If
                                'End unique name check
                            End If
                            'end empty name check
                        Else
                            ' > line 1
                            If strline.Trim.Length > 0 Then
                                'Create an array of lines ie Value,Descript,1,2,3,4,CoverFactor,W/WL
                                Dim strparams() As String = strline.Trim.Split(",")
                                'Check the values, if ok add them, if not rollback
                                If CheckGridValuesLCType(strparams) Then
                                    AddLCClass(strname, strparams)
                                Else
                                    RollBackImport(strname)
                                    _parent.cmbxLCType.Items.Clear()
                                    InitComboBox(_parent.cmbxLCType, "LCTYPE")
                                End If
                                'End check
                            End If
                        End If
                        ' line = 1 or not
                    End While
                    _parent.cmbxLCType.Items.Clear()
                    InitComboBox(_parent.cmbxLCType, "LCTYPE")
                    _parent.cmbxLCType.SelectedIndex = GetIndexOfEntry(strname, _parent.cmbxLCType)
                End Using
                MyBase.OK_Button_Click(sender, e)
            Else
                MsgBox("The file you are pointing to does not exist. Please select another.", MsgBoxStyle.Critical, "File Not Found")
            End If

        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub AddLCClass(ByRef strName As String, ByRef strParams() As String)
        Try
            Dim strLCTypeAdd As String
            Dim strCmdInsert As String

            'Get the WQCriteria values using the name
            strLCTypeAdd = String.Format("SELECT * FROM LCTYPE WHERE NAME = '{0}'", strName)
            Using cmdType As New DataHelper(strLCTypeAdd)
                Dim dataType As OleDbDataReader = cmdType.ExecuteReader()
                dataType.Read()
                strCmdInsert = String.Format("INSERT INTO LCCLASS([Value],[Name],[LCTYPEID],[CN-A],[CN-B],[CN-C],[CN-D],[CoverFactor],[W_WL]) VALUES({0},'{1}',{2},{3},{4},{5},{6},{7},{8})", Replace(CStr(strParams(0)), "'", "''"), Replace(CStr(strParams(1)), "'", "''"), Replace(CStr(dataType("LCTypeID")), "'", "''"), Replace(CStr(strParams(2)), "'", "''"), Replace(CStr(strParams(3)), "'", "''"), Replace(CStr(strParams(4)), "'", "''"), Replace(CStr(strParams(5)), "'", "''"), Replace(CStr(strParams(6)), "'", "''"), Replace(CStr(strParams(7)), "'", "''"))
                dataType.Close()
                Using cmdIns As New DataHelper(strCmdInsert)
                    cmdIns.ExecuteNonQuery()
                End Using
            End Using
        Catch ex As Exception
            MsgBox("There was a problem updating the database.  Insure that your values meet the correct " & "value ranges for each field.", MsgBoxStyle.Critical, "Invalid Values Found")
        End Try
    End Sub

    Private Function CheckGridValuesLCType(ByRef aryValue() As String) As Boolean
        Try
            For i As Integer = 0 To aryValue.Length - 1

                Select Case i
                    Case 6
                        If aryValue(i) < 0 Or aryValue(i) > 1 Then
                            MsgBox("Cover Factor must be between 0 and 1.  Please check values", MsgBoxStyle.Critical, "Check Cover Number")
                            CheckGridValuesLCType = False
                        Else
                            CheckGridValuesLCType = True
                        End If
                End Select
            Next i
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Function

    Private Sub RollBackImport(ByRef strName As String)
        Try
            Dim strSQLDel As String = "DELETE FROM LCTYPE where NAME LIKE '" & strName & "'"
            Using cmdDel As New DataHelper(strSQLDel)
                cmdDel.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub
End Class