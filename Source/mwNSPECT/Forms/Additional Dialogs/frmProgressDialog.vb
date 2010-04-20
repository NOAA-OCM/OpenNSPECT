Public Class frmProgressDialog
    Public Property Title() As String
        Get
            Return Me.Text
        End Get
        Set(ByVal value As String)
            Me.Text = value
        End Set
    End Property

    Public Property Description() As String
        Get
            Return lblDesc.Text
        End Get
        Set(ByVal value As String)
            lblDesc.Text = value
        End Set
    End Property

    Public Property MinRange() As Integer
        Get
            Return pbMain.Minimum
        End Get
        Set(ByVal value As Integer)
            pbMain.Minimum = value
        End Set
    End Property

    Public Property MaxRange() As Integer
        Get
            Return pbMain.Maximum
        End Get
        Set(ByVal value As Integer)
            pbMain.Maximum = value
        End Set
    End Property

    Public Property Progress() As Integer
        Get
            Return pbMain.Value
        End Get
        Set(ByVal value As Integer)
            pbMain.Value = value
        End Set
    End Property

    Public Property CancelEnabled() As Boolean
        Get
            Return btnCancel.Enabled
        End Get
        Set(ByVal value As Boolean)
            btnCancel.Enabled = value
        End Set
    End Property

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        g_boolCancel = False
    End Sub
End Class