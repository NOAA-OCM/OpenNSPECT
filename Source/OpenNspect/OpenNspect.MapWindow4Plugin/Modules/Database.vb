Imports System.Data.OleDb
Imports System.Windows.Forms

Module Database

    Public g_DBConn As OleDbConnection

    Public Sub InitializeDBConnection()
        Dim connectionString As String = String.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}\nspect.mdb", g_nspectPath)
        Try
            If g_DBConn Is Nothing Then
                g_DBConn = New OleDbConnection(connectionString)
                g_DBConn.Open()
            End If

        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Public Function UniqueName(ByRef strTableName As String, ByRef strName As String) As Boolean
        Try
            Dim strCmdText As String = String.Format("SELECT * FROM {0} WHERE NAME LIKE '{1}'", strTableName, strName)
            Using cmdName As New DataHelper(strCmdText)
                Dim datName As OleDbDataReader = cmdName.ExecuteReader()
                UniqueName = Not datName.HasRows
                datName.Close()
            End Using
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Function

    'Tests name inputs to insure unique values for databases

    Public Function CreateUniqueName(ByRef strTableName As String, ByRef strName As String) As String
        CreateUniqueName = ""
        Try
            Dim strCmdText As String
            Dim sCurrNum As String
            Dim strCurrNameRecord As String
            strCmdText = "SELECT * FROM " & strTableName
            '& " WHERE NAME LIKE '" & strName & "'"
            Using cmd As New DataHelper(strCmdText)
                Dim data As OleDbDataReader = cmd.ExecuteReader
                sCurrNum = "0"
                While data.Read()
                    strCurrNameRecord = CStr(data("Name"))
                    If InStr(1, strCurrNameRecord, strName, 1) > 0 Then
                        If IsNumeric(Right(strCurrNameRecord, 2)) Then
                            If (CShort(Right(strCurrNameRecord, 2)) > CShort(sCurrNum)) Then
                                sCurrNum = Right(strCurrNameRecord, 2)
                            Else
                                Exit While
                            End If
                        Else
                            If IsNumeric(Right(strCurrNameRecord, 1)) Then
                                If (CShort(Right(strCurrNameRecord, 1)) > CShort(sCurrNum)) Then
                                    sCurrNum = Right(strCurrNameRecord, 1)
                                End If
                            End If
                        End If
                    End If
                End While
                If sCurrNum = "0" Then
                    CreateUniqueName = strName & "1"
                Else
                    CreateUniqueName = strName & CStr(CShort(sCurrNum) + 1)
                End If
                data.Close()
            End Using

        Catch ex As Exception
            HandleError(ex)
        End Try
    End Function

    ''' <summary>
    ''' Loads the variety of comboboxes throught the project using combobox and name of table.
    ''' </summary>
    ''' <param name="cbo">The dropdown.</param>
    ''' <param name="name">Name.</param>
    Public Sub InitComboBox(ByRef cbo As ComboBox, ByRef name As String)
        Try
            Dim strSelectStatement As String = String.Format("SELECT NAME FROM {0} ORDER BY NAME ASC", name)

            Using rsNamesCmd As OleDbCommand = New OleDbCommand(strSelectStatement, g_DBConn)
                Using rsNames As OleDbDataReader = rsNamesCmd.ExecuteReader()
                    If Not rsNames.HasRows Then
                        MsgBox("Warning.  There are no records remaining.  Please add a new one.", MsgBoxStyle.Critical, "Recordset Empty")
                        Return
                    Else
                        With cbo
                            Do While rsNames.Read()
                                .Items.Add(rsNames.Item("Name"))
                            Loop
                        End With
                        ' Select the first item in the combo by default
                        cbo.SelectedIndex = 0
                    End If
                End Using
            End Using
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub
End Module
