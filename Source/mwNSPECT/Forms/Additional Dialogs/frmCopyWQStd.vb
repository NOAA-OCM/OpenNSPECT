Imports System.Data.OleDb
Friend Class frmCopyWQStd
    Inherits System.Windows.Forms.Form

    Private _frmWQStd As frmWaterQualityStandard

    Private Sub frmCopyWQStd_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        modUtil.InitComboBox(cboStdName, "WQCRITERIA")

    End Sub

    Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click

        Dim strStandard As String

        Dim strPollStandard As String

        Dim strCmd As String
        Dim strCmd2 As String

        'Get the WQ stand info
        strStandard = "SELECT * FROM WQCriteria WHERE NAME LIKE '" & cboStdName.Text & "'"
        Dim cmdstd As New OleDbCommand(strStandard, g_DBConn)
        Dim datastd As OleDbDataReader = cmdstd.ExecuteReader()
        datastd.Read()

        'Get the related pollutant/thresholds
        strPollStandard = "SELECT * FROM POLL_WQCRITERIA WHERE WQCRITID =" & datastd("WQCRITID").ToString
        Dim cmdPoll As New OleDbCommand(strPollStandard, g_DBConn)
        Dim datapoll As OleDbDataReader = cmdPoll.ExecuteReader()

        strCmd = "INSERT INTO WQCRITERIA (NAME,DESCRIPTION) VALUES ('" & Replace(Trim(txtStdName.Text), "'", "''") & "', '" & datastd("Description") & "')"

        If modUtil.UniqueName("WQCRITERIA", Trim(txtStdName.Text)) Then
            Dim cmdIns As New OleDbCommand(strCmd, g_DBConn)
            cmdIns.ExecuteNonQuery()
        Else
            MsgBox(Err4, MsgBoxStyle.Critical, "Enter Unique Name")
            Exit Sub
        End If
        Dim cmdNewStandard As New OleDbCommand("Select * from WQCRITERIA WHERE NAME LIKE '" & Trim(txtStdName.Text) & "'", g_DBConn)
        Dim datanewstd As OleDbDataReader = cmdNewStandard.ExecuteReader()
        datanewstd.Read()
        
        Dim i As Short
        i = 0

        While datapoll.Read()
            strCmd2 = "INSERT INTO POLL_WQCRITERIA (POLLID, WQCRITID, THRESHOLD) VALUES (" & datapoll("POLLID") & ", " & datanewstd("WQCRITID") & "," & datapoll("Threshold") & ")"
            Dim cmdIns2 As New OleDbCommand(strCmd2, g_DBConn)
            cmdIns2.ExecuteNonQuery()
        End While
        datastd.Close()
        datapoll.Close()
        datanewstd.Close()

        _frmWQStd.UpdateWQ(Trim(txtStdName.Text))
        Me.Close()

    End Sub


    Public Sub Init(ByRef frmWQ As frmWaterQualityStandard)
        _frmWQStd = frmWQ
    End Sub

    

End Class