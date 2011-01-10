Public Class frmSelectShape
    Private stopclose As Boolean

    Public Sub Initialize()
        Me.Show()

        g_MapWin.View.CursorMode = MapWinGIS.tkCursorMode.cmSelection

        If Not g_cb Is Nothing Then g_cb.Visible = False
        If Not g_comp Is Nothing Then g_comp.Visible = False
        If Not g_luscen Is Nothing Then g_luscen.Visible = False
    End Sub

    Private Sub frmSelectShapes_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        If Not g_cb Is Nothing Then g_cb.SetSelectedShape()
        If Not g_comp Is Nothing Then g_comp.SetSelectedShape()
        If Not g_luscen Is Nothing Then g_luscen.SetSelectedShape()
        Me.Hide()
        If Not g_cb Is Nothing Then g_cb.Visible = True
        If Not g_comp Is Nothing Then g_comp.Visible = True
        If Not g_luscen Is Nothing Then g_luscen.Visible = True
    End Sub

    Private Sub btnDone_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDone.Click
        Me.Close()
    End Sub

    Public Sub disableDone(ByVal disable As Boolean)
        stopclose = disable
        btnDone.Enabled = Not disable
    End Sub

    Private Sub frmSelectShape_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        If stopclose Then
            e.Cancel = True
        End If
    End Sub

End Class