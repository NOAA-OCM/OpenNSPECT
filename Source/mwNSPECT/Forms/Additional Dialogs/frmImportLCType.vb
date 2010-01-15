Option Strict Off
Option Explicit On
Friend Class frmImportLCType
	Inherits System.Windows.Forms.Form
    

    Private Sub frmImportLCType_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub txtLCType_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLCType.TextChanged

    End Sub

    Private Sub txtImpFile_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtImpFile.TextChanged

    End Sub

    Private Sub cmdBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowse.Click

    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click

    End Sub
End Class