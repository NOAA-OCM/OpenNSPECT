Imports System.Data.OleDb
Imports System.Windows.Forms
Friend Class frmNewLCType
    Inherits System.Windows.Forms.Form

    Private _frmLC As frmLandCoverTypes

    Const c_sModuleFileName As String = "frmNewLCType.vb"

#Region "Events"
    Private Sub frmNewLCType_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        dgvLCTypes.Rows.Add()
        With dgvLCTypes.Rows(0)
            .Cells(0).Value = "0"
            .Cells(1).Value = "Landclass1"
            .Cells(2).Value = "0"
            .Cells(3).Value = "0"
            .Cells(4).Value = "0"
            .Cells(5).Value = "0"
            .Cells(6).Value = "0"
            .Cells(7).Value = "0"
        End With
    End Sub

    Private Sub dgvLCTypes_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgvLCTypes.MouseClick
        If e.Button = Windows.Forms.MouseButtons.Right Then
            cntxmnuGrid.Show(dgvLCTypes, New Drawing.Point(e.X, e.Y))
        End If

    End Sub

    Private Sub AddRowToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddRowToolStripMenuItem.Click
        AddRow()
    End Sub

    Private Sub InsertRowToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InsertRowToolStripMenuItem.Click
        InsertRow()
    End Sub

    Private Sub DeleteRowToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteRowToolStripMenuItem.Click
        DeleteRow()
    End Sub

    Private Sub dgvLCTypes_DataError(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvLCTypes.DataError
        MsgBox("Please enter a valid number.")
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Try
            Me.Close()
        Catch ex As Exception
            HandleError(True, "cmdCancel_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try
    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        Try

            Dim strName As String
            Dim strDescript As String
            Dim strCmd As String 'INSERT function
            Dim arrParams(7) As Object 'Array that holds each row's contents


            If ValidateGridValues() Then
                'Get rid of possible apostrophes in name
                strName = Trim(txtLCType.Text)
                strDescript = Trim(txtLCTypeDesc.Text)

                If modUtil.UniqueName("LCTYPE", strName) And (Trim(strName) <> "") Then
                    strCmd = "INSERT INTO LCTYPE (NAME,DESCRIPTION) VALUES ('" & Replace(strName, "'", "''") & "', '" & Replace(strDescript, "'", "''") & "')"
                    Dim cmdStr As New OleDbCommand(strCmd, g_DBConn)
                    cmdStr.ExecuteNonQuery()
                Else
                    MsgBox("The name you have chosen is already in use.  Please select another.", MsgBoxStyle.Critical, "Select Unique Name")
                    Exit Sub
                End If 'End unique name check

                'Now add GRID values

                For Each row As DataGridViewRow In dgvLCTypes.Rows
                    AddLCClass(strName, row)
                Next



                MsgBox("Data saved successfully.", MsgBoxStyle.OkOnly, "Data Saved")

                Me.Close()
                _frmLC.UpdateCombo(strName)
            Else
                Exit Sub
            End If


        Catch ex As Exception
            HandleError(True, "cmdOK_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try
    End Sub
#End Region

#Region "Helper Functions'"

    Public Sub Init(ByRef frmLC As frmLandCoverTypes)
        _frmLC = frmLC
    End Sub

    Private Sub AddRow()
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
    End Sub

    Private Sub InsertRow()
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
    End Sub

    Private Sub DeleteRow()
        If Not dgvLCTypes.CurrentRow Is Nothing Then
            dgvLCTypes.Rows.Remove(dgvLCTypes.CurrentRow)
        End If
    End Sub

    Private Function ValidateGridValues() As Boolean
        Try
            'Need to validate each grid value before saving.  Essentially we take it a row at a time,
            'then rifle through each column of each row.  Case Select tests each each x,y value depending
            'on column... eg Column 1 must be unique, 3-6 must be 1-100 range, 7 must be <= 1

            'Returns: True or False

            Dim val As String 'txtActiveCell value
            Dim val2 As String 'Value of Column 2 ([VALUE]) - have to check for unique
            Dim i As Short
            Dim j As Short
            Dim k As Short

            For i = 0 To dgvLCTypes.Rows.Count - 1

                For j = 0 To 6

                    val = dgvLCTypes.Rows(i).Cells(j).Value

                    Select Case j

                        Case 1
                            If Not IsNumeric(val) Then
                                ErrorGenerator(Err1, i, j)
                            Else
                                For k = 0 To dgvLCTypes.Rows.Count - 1

                                    val2 = dgvLCTypes.Rows(k).Cells(0).Value
                                    If k <> i Then 'Don't want to compare value to itself
                                        If val2 = dgvLCTypes.Rows(i).Cells(0).Value Then
                                            ErrorGenerator(Err2, i, j)
                                            Return False
                                        End If
                                    End If
                                Next k
                            End If


                        Case 2
                            If IsNumeric(val) Then
                                ErrorGenerator(Err1, i, j)
                                Return False
                            End If

                        Case 3
                            If Not IsNumeric(val) Or ((val < 0) Or (val > 100)) Then
                                ErrorGenerator(Err1, i, j)
                                Return False
                            End If

                        Case 4
                            If Not IsNumeric(val) Or ((val < 0) Or (val > 100)) Then
                                ErrorGenerator(Err1, i, j)
                                Return False
                            End If

                        Case 5
                            If Not IsNumeric(val) Or ((val < 0) Or (val > 100)) Then
                                ErrorGenerator(Err1, i, j)
                                Return False
                            End If

                        Case 6
                            If Not IsNumeric(val) Or ((val < 0) Or (val > 100)) Then
                                ErrorGenerator(Err1, i, j)
                                Return False
                            End If

                        Case 7
                            If Not IsNumeric(val) Or ((val < 0) Or (val > 1)) Then
                                ErrorGenerator(Err3, i, j)
                                Return False
                            End If
                    End Select
                Next j
            Next i

            ValidateGridValues = True

        Catch ex As Exception
            HandleError(False, "ValidateGridValues " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try
    End Function



    Private Sub AddLCClass(ByRef strName As String, ByRef row As DataGridViewRow)
        Try
            'Called from cmdOK_Click, this uses a passed array to insert new landclasses

            Dim strLCTypeAdd As String
            Dim strCmdInsert As String

            'Get the WQCriteria values using the name
            strLCTypeAdd = "SELECT * FROM LCTYPE WHERE NAME = " & "'" & strName & "'"
            Dim lctypeaddCmd As New OleDbCommand(strLCTypeAdd, g_DBConn)
            Dim datLCType As OleDbDataReader = lctypeaddCmd.ExecuteReader()
            datLCType.Read()

            strCmdInsert = "INSERT INTO LCCLASS([Value],[Name],[LCTYPEID],[CN-A],[CN-B],[CN-C],[CN-D],[CoverFactor],[W_WL]) VALUES(" & CStr(row.Cells(0).Value) & ",'" & CStr(row.Cells(1).Value) & "'," & CStr(datLCType("LCTypeID")) & "," & CStr(row.Cells(2).Value) & "," & CStr(row.Cells(3).Value) & "," & CStr(row.Cells(4).Value) & "," & CStr(row.Cells(5).Value) & "," & CStr(row.Cells(6).Value) & "," & CStr(row.Cells(7).Value) & ")"
            Dim insertCmd As New OleDbCommand(strCmdInsert, g_DBConn)
            insertCmd.ExecuteNonQuery()

            datLCType.Close()

        Catch ex As Exception
            MsgBox("There is an error inserting records into LCClass.", MsgBoxStyle.Critical, Err.Number & ": " & Err.Description)
        End Try

    End Sub

#End Region

End Class