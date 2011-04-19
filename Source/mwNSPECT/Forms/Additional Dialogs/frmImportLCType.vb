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

Imports System.Data.OleDb
Friend Class frmImportLCType
    Inherits System.Windows.Forms.Form

    Private _booName As Boolean 'Check if user put a name in
    Private _booFile As Boolean 'Check if FileName is correct
    Private _strFileName As String
    Private _parent As frmLandCoverTypes
    Const c_sModuleFileName As String = "frmImportLCType.vb"

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="parent"></param>
    ''' <remarks></remarks>
    Public Sub Init(ByRef parent As frmLandCoverTypes)
        Try
            _parent = parent
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub txtLCType_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLCType.TextChanged
        Try
            If CBool(Trim(CStr(Len(txtLCType.Text)))) Then
                _booName = True
            End If

            If _booFile And _booName Then
                cmdOK.Enabled = True
            End If
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub cmdBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowse.Click
        Try
            Dim dlgopen As New Windows.Forms.OpenFileDialog
            dlgopen.Filter = Replace(MSG1, "<name>", "Land Cover Type")
            dlgopen.Title = Replace(MSG2, "<name>", "Land Cover Type")

            If dlgopen.ShowDialog = Windows.Forms.DialogResult.OK Then
                txtImpFile.Text = dlgopen.FileName
                _strFileName = dlgopen.FileName
                _booFile = True
            End If

            If _booFile And _booName Then
                cmdOK.Enabled = True
            End If
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Try
            Me.Close()
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        Try
            If IO.File.Exists(_strFileName) Then
                Dim read As New IO.StreamReader(_strFileName)
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
                            Exit Sub
                        Else
                            'Name Check, if cool perform
                            If modUtil.UniqueName("LCTYPE", txtLCType.Text) Then
                                Dim strCmd As String = "INSERT INTO LCTYPE (NAME,DESCRIPTION) VALUES ('" & Replace(txtLCType.Text, "'", "''") & "', '" & Replace(strDesc, "'", "''") & "')"
                                Dim cmdIns As New OleDbCommand(strCmd, g_DBConn)
                                cmdIns.ExecuteNonQuery()
                            Else
                                MsgBox("The name you have chosen is already in use.  Please select another.", MsgBoxStyle.Critical, "Select Unique Name")
                                Exit Sub
                            End If 'End unique name check

                        End If 'end empty name check
                    Else ' > line 1
                        If strline.Trim.Length > 0 Then
                            'Create an array of lines ie Value,Descript,1,2,3,4,CoverFactor,W/WL
                            Dim strparams() As String = strline.Trim.Split(",")

                            'Check the values, if ok add them, if not rollback
                            If CheckGridValuesLCType(strparams) Then
                                AddLCClass(strname, strparams)
                            Else
                                RollBackImport(strname)
                                _parent.cmbxLCType.Items.Clear()
                                modUtil.InitComboBox(_parent.cmbxLCType, "LCTYPE")
                            End If 'End check
                        End If

                    End If ' line = 1 or not
                End While

                _parent.cmbxLCType.Items.Clear()
                modUtil.InitComboBox(_parent.cmbxLCType, "LCTYPE")
                _parent.cmbxLCType.SelectedIndex = modUtil.GetCboIndex(strname, _parent.cmbxLCType)
                read.Close()
                Me.Close()
            Else
                MsgBox("The file you are pointing to does not exist. Please select another.", MsgBoxStyle.Critical, "File Not Found")
            End If

        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="strName"></param>
    ''' <param name="strParams"></param>
    ''' <remarks></remarks>
    Private Sub AddLCClass(ByRef strName As String, ByRef strParams() As String)
        Try
            Dim strLCTypeAdd As String
            Dim strCmdInsert As String


            'Get the WQCriteria values using the name
            strLCTypeAdd = "SELECT * FROM LCTYPE WHERE NAME = " & "'" & strName & "'"
            Dim cmdType As New OleDbCommand(strLCTypeAdd, g_DBConn)
            Dim dataType As OleDbDataReader = cmdType.ExecuteReader()
            dataType.Read()
            strCmdInsert = "INSERT INTO LCCLASS([Value],[Name],[LCTYPEID],[CN-A],[CN-B],[CN-C],[CN-D],[CoverFactor],[W_WL]) VALUES(" & Replace(CStr(strParams(0)), "'", "''") & ",'" & Replace(CStr(strParams(1)), "'", "''") & "'," & Replace(CStr(dataType("LCTypeID")), "'", "''") & "," & Replace(CStr(strParams(2)), "'", "''") & "," & Replace(CStr(strParams(3)), "'", "''") & "," & Replace(CStr(strParams(4)), "'", "''") & "," & Replace(CStr(strParams(5)), "'", "''") & "," & Replace(CStr(strParams(6)), "'", "''") & "," & Replace(CStr(strParams(7)), "'", "''") & ")"
            dataType.Close()

            Dim cmdIns As New OleDbCommand(strCmdInsert, g_DBConn)
            cmdIns.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("There was a problem updating the database.  Insure that your values meet the correct " & "value ranges for each field.", MsgBoxStyle.Critical, "Invalid Values Found")
        End Try
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="aryValue"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckGridValuesLCType(ByRef aryValue() As String) As Boolean
        Try
            For i As Integer = 0 To aryValue.Length - 1
                'Select Case i
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
            HandleError(c_sModuleFileName, ex)
        End Try
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="strName"></param>
    ''' <remarks></remarks>
    Private Sub RollBackImport(ByRef strName As String)
        Try
            Dim strSQLDel As String = "DELETE FROM LCTYPE where NAME LIKE '" & strName & "'"
            Dim cmdDel As New OleDbCommand(strSQLDel, g_DBConn)
            cmdDel.ExecuteNonQuery()
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub

End Class