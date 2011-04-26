Imports System.Windows.Forms

Public Class BaseDialogForm
    Private _isDirty As Boolean
    ''' <summary>
    ''' Gets or sets a value indicating whether this instance is dirty, i.e. whether it has changes.
    ''' </summary>
    ''' <value>
    '''   <c>true</c> if this instance is dirty; otherwise, <c>false</c>.
    ''' </value>
    Public Property IsDirty() As Boolean
        Get
            Return _isDirty
        End Get
        Set(ByVal value As Boolean)
            _isDirty = value
        End Set
    End Property
    Protected Overridable Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Protected Overridable Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click

        If _isDirty AndAlso Not ConfirmCancel() Then
            Exit Sub
        End If

        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Public Function ConfirmCancel() As Boolean
        Return MsgBoxResult.Yes = MsgBox("You have made changes, are you sure you want to quit?", MsgBoxStyle.YesNo, "Quit?")
    End Function

End Class
