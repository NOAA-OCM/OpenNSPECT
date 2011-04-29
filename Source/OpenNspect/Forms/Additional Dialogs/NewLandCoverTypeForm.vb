'********************************************************************************************************
'File Name: frmNewLCTpe.vb 
'Description: Form for new land cover type
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
Imports System.Drawing
Imports System.Data.OleDb

Friend Class NewLandCoverTypeForm
    Private _frmLC As LandCoverTypesForm

#Region "Events"

    Private Sub frmNewLCType_Load (ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        Try
            dgvLCTypes.Rows.Add()
            With dgvLCTypes.Rows (0)
                .Cells (0).Value = "0"
                .Cells (1).Value = "Landclass1"
                .Cells (2).Value = "0"
                .Cells (3).Value = "0"
                .Cells (4).Value = "0"
                .Cells (5).Value = "0"
                .Cells (6).Value = "0"
                .Cells (7).Value = "0"
            End With
        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub


    Private Sub AddRowToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles AddRowToolStripMenuItem.Click
        Try
            AddRow()
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub InsertRowToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles InsertRowToolStripMenuItem.Click
        Try
            InsertRow()
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub DeleteRowToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles DeleteRowToolStripMenuItem.Click
        Try
            DeleteRow()
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub dgvLCTypes_DataError(ByVal sender As Object, ByVal e As DataGridViewDataErrorEventArgs) Handles dgvLCTypes.DataError
        Try
            MsgBox("Please enter a valid number.")
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Protected Overrides Sub OK_Button_Click(sender As Object, e As EventArgs)
        Try
            Dim strName As String
            Dim strDescript As String
            Dim strCmd As String
            'INSERT function

            If ValidateGridValues() Then
                'Get rid of possible apostrophes in name
                strName = Trim(txtLCType.Text)
                strDescript = Trim(txtLCTypeDesc.Text)

                If UniqueName("LCTYPE", strName) And (Trim(strName) <> "") Then
                    strCmd = String.Format("INSERT INTO LCTYPE (NAME,DESCRIPTION) VALUES ('{0}', '{1}')", Replace(strName, "'", "''"), Replace(strDescript, "'", "''"))
                    Dim cmdStr As New DataHelper(strCmd)
                    cmdStr.ExecuteNonQuery()
                Else
                    MsgBox("The name you have chosen is already in use.  Please select another.", MsgBoxStyle.Critical, "Select Unique Name")
                    Return
                End If
                'End unique name check

                'Now add GRID values
                For Each row As DataGridViewRow In dgvLCTypes.Rows
                    AddLCClass(strName, row)
                Next

                MsgBox("Data saved successfully.", MsgBoxStyle.OkOnly, "Data Saved")

                _frmLC.UpdateCombo(strName)
                MyBase.OK_Button_Click(sender, e)
            End If
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

#End Region

#Region "Helper Functions'"

    Public Sub Init(ByRef frmLC As LandCoverTypesForm)
        Try
            _frmLC = frmLC
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub AddRow()
        Try
            Dim idx As Integer = dgvLCTypes.Rows.Add()
            With dgvLCTypes.Rows(idx)
                .Cells(0).Value = "0"
                .Cells(1).Value = "Landclass" + (dgvLCTypes.Rows.Count).ToString
                .Cells(2).Value = "0"
                .Cells(3).Value = "0"
                .Cells(4).Value = "0"
                .Cells(5).Value = "0"
                .Cells(6).Value = "0"
                .Cells(7).Value = "0"
            End With
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub InsertRow()
        Try
            If Not dgvLCTypes.CurrentRow Is Nothing Then
                Dim idx As Integer = dgvLCTypes.CurrentRow.Index
                dgvLCTypes.Rows.Insert(idx, 1)
                With dgvLCTypes.Rows(idx)
                    .Cells(0).Value = "0"
                    .Cells(1).Value = "Landclass" + (dgvLCTypes.Rows.Count).ToString
                    .Cells(2).Value = "0"
                    .Cells(3).Value = "0"
                    .Cells(4).Value = "0"
                    .Cells(5).Value = "0"
                    .Cells(6).Value = "0"
                    .Cells(7).Value = "0"
                End With
            End If
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub DeleteRow()
        Try
            If Not dgvLCTypes.CurrentRow Is Nothing Then
                dgvLCTypes.Rows.Remove(dgvLCTypes.CurrentRow)
            End If
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Function ValidateGridValues() As Boolean
        Try
            'Need to validate each grid value before saving.  Essentially we take it a row at a time,
            'then rifle through each column of each row.  Case Select tests each each x,y value depending
            'on column... eg Column 1 must be unique, 3-6 must be 1-100 range, 7 must be <= 1

            'Returns: True or False

            Dim val As String
            'txtActiveCell value
            Dim val2 As String
            'Value of Column 2 ([VALUE]) - have to check for unique
            Dim i As Short
            Dim j As Short
            Dim k As Short

            For i = 0 To dgvLCTypes.Rows.Count - 1

                For j = 0 To 6

                    val = dgvLCTypes.Rows(i).Cells(j).Value

                    Select Case j

                        Case 0
                            If Not IsNumeric(val) Then
                                DisplayError(Err1, i, j)
                            Else
                                For k = 0 To dgvLCTypes.Rows.Count - 1

                                    val2 = dgvLCTypes.Rows(k).Cells(0).Value
                                    If k <> i Then 'Don't want to compare value to itself
                                        If val2 = dgvLCTypes.Rows(i).Cells(0).Value Then
                                            DisplayError(Err2, i, j)
                                            Return False
                                        End If
                                    End If
                                Next k
                            End If

                        Case 1
                            If IsNumeric(val) Then
                                DisplayError(Err1, i, j)
                                Return False
                            End If

                        Case 2
                            If Not IsNumeric(val) Or ((val < 0) Or (val > 100)) Then
                                DisplayError(Err1, i, j)
                                Return False
                            End If

                        Case 3
                            If Not IsNumeric(val) Or ((val < 0) Or (val > 100)) Then
                                DisplayError(Err1, i, j)
                                Return False
                            End If

                        Case 4
                            If Not IsNumeric(val) Or ((val < 0) Or (val > 100)) Then
                                DisplayError(Err1, i, j)
                                Return False
                            End If

                        Case 5
                            If Not IsNumeric(val) Or ((val < 0) Or (val > 100)) Then
                                DisplayError(Err1, i, j)
                                Return False
                            End If

                        Case 6
                            If Not IsNumeric(val) Or ((val < 0) Or (val > 1)) Then
                                DisplayError(Err3, i, j)
                                Return False
                            End If
                    End Select
                Next j
            Next i

            ValidateGridValues = True

        Catch ex As Exception
            HandleError(ex)
        End Try
    End Function

    Private Sub AddLCClass(ByRef strName As String, ByRef row As DataGridViewRow)
        Try
            'Called from cmdOK_Click, this uses a passed array to insert new landclasses

            Dim strLCTypeAdd As String
            Dim strCmdInsert As String

            'Get the WQCriteria values using the name
            strLCTypeAdd = "SELECT * FROM LCTYPE WHERE NAME = " & "'" & strName & "'"
            Dim lctypeaddCmd As New DataHelper(strLCTypeAdd)
            Dim datLCType As OleDbDataReader = lctypeaddCmd.ExecuteReader()
            datLCType.Read()

            strCmdInsert = "INSERT INTO LCCLASS([Value],[Name],[LCTYPEID],[CN-A],[CN-B],[CN-C],[CN-D],[CoverFactor],[W_WL]) VALUES(" & CStr(row.Cells(0).Value) & ",'" & CStr(row.Cells(1).Value) & "'," & CStr(datLCType("LCTypeID")) & "," & CStr(row.Cells(2).Value) & "," & CStr(row.Cells(3).Value) & "," & CStr(row.Cells(4).Value) & "," & CStr(row.Cells(5).Value) & "," & CStr(row.Cells(6).Value) & "," & CStr(row.Cells(7).Value) & ")"
            Dim insertCmd As New DataHelper(strCmdInsert)
            insertCmd.ExecuteNonQuery()

            datLCType.Close()

        Catch ex As Exception
            MsgBox("There is an error inserting records into LCClass.", MsgBoxStyle.Critical, Err.Number & ": " & Err.Description)
        End Try

    End Sub

#End Region
End Class