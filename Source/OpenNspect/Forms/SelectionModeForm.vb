'********************************************************************************************************
'File Name: frmSelectShape.vb
'Description: Main  input and model activation form
'********************************************************************************************************
'The contents of this file are subject to the Mozilla Public License Version 1.1 (the "License"); 
'you may not use this file except in compliance with the License. You may obtain a copy of the License at 
'http://www.mozilla.org/MPL/ 
'Software distributed under the License is distributed on an "AS IS" basis, WITHOUT WARRANTY OF 
'ANY KIND, either express or implied. See the License for the specificlanguage governing rights and 
'limitations under the License. 
'
'Contributor(s): (Open source contributors should list themselves and their modifications here). 
'Oct 20, 2010:  Allen Anselmo allen.anselmo@gmail.com - 
'               Added licensing and comments to code
Public Class SelectionModeForm
    Private stopclose As Boolean
    Const c_sModuleFileName As String = "frmSelectShape.vb"


    Public Sub InitializeAndShow()
        Try
            Me.Show()

            g_MapWin.View.CursorMode = MapWinGIS.tkCursorMode.cmSelection

            If Not g_cbMainForm Is Nothing Then g_cbMainForm.Visible = False
            If Not g_comp Is Nothing Then g_comp.Visible = False
            If Not g_luscen Is Nothing Then g_luscen.Visible = False
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub frmSelectShapes_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        If Not g_cbMainForm Is Nothing Then g_cbMainForm.SetSelectedShape()
        If Not g_comp Is Nothing Then g_comp.SetSelectedShape()
        If Not g_luscen Is Nothing Then g_luscen.SetSelectedShape()
        Me.Hide()
        If Not g_cbMainForm Is Nothing Then g_cbMainForm.Visible = True
        If Not g_comp Is Nothing Then g_comp.Visible = True
        If Not g_luscen Is Nothing Then g_luscen.Visible = True
    End Sub


    Private Sub btnDone_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDone.Click
        Try
            Close()
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Public Sub disableDone(ByVal disable As Boolean)
        Try
            stopclose = disable
            btnDone.Enabled = Not disable
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub frmSelectShape_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        Try
            If stopclose Then
                e.Cancel = True
            End If
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub

End Class