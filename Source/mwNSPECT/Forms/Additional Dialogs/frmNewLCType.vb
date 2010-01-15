Option Strict Off
Option Explicit On
Friend Class frmNewLCType
	Inherits System.Windows.Forms.Form


    Const c_sModuleFileName As String = "frmNewLCType.vb"

    Private Sub frmNewLCType_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub txtLCType_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLCType.TextChanged

    End Sub

    Private Sub txtLCTypeDesc_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLCTypeDesc.TextChanged

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