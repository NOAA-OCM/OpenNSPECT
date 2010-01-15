Option Strict Off
Option Explicit On
Friend Class frmImportWQStd
	Inherits System.Windows.Forms.Form

    Const c_sModuleFileName As String = "frmImportWQStd.vb"


    Private Sub frmImportWQStd_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub txtStdName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtStdName.TextChanged

    End Sub

    Private Sub txtImpFile_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtImpFile.TextChanged

    End Sub

    Private Sub cmdBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowse.Click

    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Try
            Me.Close()
        Catch ex As Exception
            HandleError(True, "cmdCancel_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try
    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click

    End Sub
End Class