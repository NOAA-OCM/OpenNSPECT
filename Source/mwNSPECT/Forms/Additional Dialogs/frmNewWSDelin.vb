Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmNewWSDelin
	Inherits System.Windows.Forms.Form

    Private _booProject As Boolean

    Private Const c_sModuleFileName As String = "frmNewWSDelin.vb"

    Private Sub frmNewWSDelin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub txtWSDelinName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtWSDelinName.TextChanged

    End Sub

    Private Sub txtDEMFile_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDEMFile.TextChanged

    End Sub

    Private Sub cmdBrowseDEMFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowseDEMFile.Click

    End Sub

    Private Sub chkHydroCorr_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkHydroCorr.CheckedChanged

    End Sub

    Private Sub cboDEMUnits_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDEMUnits.SelectedIndexChanged

    End Sub

    Private Sub cboSubWSSize_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSubWSSize.SelectedIndexChanged

    End Sub

    Private Sub cmdCreate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCreate.Click

    End Sub

    Private Sub cmdQuit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdQuit.Click
        Try
            Dim intvbYesNo As Short

            intvbYesNo = MsgBox("Are you sure you want to exit?  Your changes will not be saved.", MsgBoxStyle.YesNo, "Exit")

            If intvbYesNo = MsgBoxResult.Yes Then
                If _booProject Then
                    'TODO: global project form
                    'frmPrj.cboWSDelin.SelectedIndex = 0
                    Me.Close()
                Else
                    Me.Close()
                End If
            Else
                Exit Sub
            End If

            _booProject = False


        Catch ex As Exception
            HandleError(True, "cmdQuit_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
        End Try

    End Sub
End Class