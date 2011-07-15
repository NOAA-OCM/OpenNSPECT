

Public Class BaseDialogForm

    ''' <summary>
    ''' Gets or sets a value indicating whether this instance is dirty, i.e. whether it has changes.
    ''' </summary>
    ''' <value>
    '''   <c>true</c> if this instance is dirty; otherwise, <c>false</c>.
    ''' </value>
    Public Property IsDirty() As Boolean

    Protected Overridable Sub OK_Button_Click(ByVal sender As Object, ByVal e As EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Protected Overridable Sub Cancel_Button_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Cancel_Button.Click

        If IsDirty AndAlso Not ConfirmCancel() Then
            Return
        End If

        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Public Function ConfirmCancel() As Boolean
        Return MsgBoxResult.Yes = MsgBox("Are you sure you want to cancel, losing unsaved changes?", MsgBoxStyle.YesNo, "Discard Changes?")
    End Function
End Class
