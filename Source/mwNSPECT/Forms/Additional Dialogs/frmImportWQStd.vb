Imports System.Data.OleDb
Friend Class frmImportWQStd
    Inherits System.Windows.Forms.Form

    Const c_sModuleFileName As String = "frmImportWQStd.vb"
    Private _frmWQ As frmWaterQualityStandard
    Private _strFileName As String


    Private Sub cmdBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowse.Click
        Try
            'browse...get output filename
            Dim dlgOpen As New Windows.Forms.OpenFileDialog
            dlgOpen.Filter = MSG1
            dlgOpen.Title = MSG2

            If dlgOpen.ShowDialog = Windows.Forms.DialogResult.OK Then
                txtImpFile.Text = Trim(dlgOpen.FileName)
                _strFileName = txtImpFile.Text
                cmdOK.Enabled = True
            End If

        Catch ex As Exception
            HandleError(True, "cmdBrowse_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)

        End Try
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


            Dim strLine As String
            Dim intLine As Short
            Dim strName As String = ""
            Dim strDescript As String
            Dim strPoll As String
            Dim strThresh As String
            Dim strCmd As String

            Dim read As New IO.StreamReader(_strFileName)

            intLine = 0

            Do While Not read.EndOfStream
                strLine = read.ReadLine
                intLine = intLine + 1
                'MsgBox theLine

                If intLine = 1 Then

                    strName = Trim(txtStdName.Text)
                    strDescript = Split(strLine, ",")(1)

                    If strName = "" Then

                        MsgBox("Name is blank.  Please enter a name.", MsgBoxStyle.Critical, "Empty Name Field")
                        txtStdName.Focus()
                        Exit Sub

                    Else

                        strCmd = "INSERT INTO WQCRITERIA (NAME,DESCRIPTION) VALUES ('" & Replace(txtStdName.Text, "'", "''") & "', '" & Replace(strDescript, "'", "''") & "')"
                        'Name Check
                        If modUtil.UniqueName("WQCRITERIA", (txtStdName.Text)) Then
                            Dim cmdIns As New OleDbCommand(strCmd, g_DBConn)
                            cmdIns.ExecuteNonQuery()
                        Else
                            MsgBox("The name you have chosen is already in use.  Please select another.", MsgBoxStyle.Critical, "Select Unique Name")
                            Exit Sub
                        End If

                    End If

                Else

                    strPoll = Split(strLine, ",")(0)
                    strThresh = Split(strLine, ",")(1)
                    'Insert the pollutant/threshold value into POLL_WQCRITERIA
                    PollutantAdd(strName, strPoll, strThresh)

                End If

            Loop

            read.Close()

            'Cleanup
            _frmWQ.cboWQStdName.Items.Clear()
            modUtil.InitComboBox(_frmWQ.cboWQStdName, "WQCRITERIA")
            _frmWQ.cboWQStdName.SelectedIndex = modUtil.GetCboIndex((txtStdName.Text), _frmWQ.cboWQStdName)
            Me.Close()
        Catch ex As Exception
            HandleError(True, "cmdOK_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try
    End Sub

    Public Sub Init(ByRef frmWQ As frmWaterQualityStandard)
        _frmWQ = frmWQ
    End Sub

    Private Sub PollutantAdd(ByRef strName As String, ByRef strPoll As String, ByRef strThresh As String)
        Try

            Dim strPollAdd As String
            Dim strPollDetails As String
            Dim strCmdInsert As String


            'Get the WQCriteria values using the name
            strPollAdd = "SELECT * FROM WQCriteria WHERE NAME = " & "'" & strName & "'"
            Dim cmdPollAdd As New OleDbCommand(strPollAdd, g_DBConn)
            Dim datapolladd As OleDbDataReader = cmdPollAdd.ExecuteReader
            datapolladd.Read()

            'Get the pollutant particulars
            strPollDetails = "SELECT * FROM POLLUTANT WHERE NAME =" & "'" & strPoll & "'"
            Dim cmdPollDet As New OleDbCommand(strPollDetails, g_DBConn)
            Dim datapolldet As OleDbDataReader = cmdPollDet.ExecuteReader
            datapolldet.Read()

            strCmdInsert = "INSERT INTO POLL_WQCRITERIA (PollID,WQCritID,Threshold) VALUES ('" & datapolldet("POLLID") & "', '" & datapolladd("WQCRITID") & "'," & strThresh & ")"
            Dim cmdIns As New OleDbCommand(strCmdInsert, g_DBConn)
            cmdIns.ExecuteNonQuery()

            datapolladd.Close()
            datapolldet.Close()
        Catch ex As Exception
            HandleError(False, "PollutantAdd " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try
    End Sub
End Class