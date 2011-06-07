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
Imports System.Windows.Forms
Imports MapWinGIS

Public Class SelectionModeForm
    Private stopclose As Boolean

    Private Shared Sub SetFormsVisibility(visible As Boolean)
        If Not g_cbMainForm Is Nothing Then g_cbMainForm.Visible = visible
        If Not g_comp Is Nothing Then g_comp.Visible = visible
        If Not g_luscen Is Nothing Then g_luscen.Visible = visible
    End Sub
    Public Sub InitializeAndShow()
        Try
            Me.Show()

            MapWindowPlugin.MapWindowInstance.View.CursorMode = tkCursorMode.cmSelection

            SetFormsVisibility(False)
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub frmSelectShapes_FormClosed(ByVal sender As Object, ByVal e As FormClosedEventArgs) Handles MyBase.FormClosed
        Try
            If Not g_cbMainForm Is Nothing Then g_cbMainForm.SetSelectedShape(MapWindowPlugin.MapWindowInstance.Layers.CurrentLayer)
            If Not g_comp Is Nothing Then g_comp.SetSelectedShape()
            If Not g_luscen Is Nothing Then g_luscen.SetSelectedShape()

            SetFormsVisibility(True)
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub btnDone_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnDone.Click
        Close()
    End Sub

    Public Sub disableDone(ByVal disable As Boolean)
        Try
            stopclose = disable
            btnDone.Enabled = Not disable
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub frmSelectShape_FormClosing(ByVal sender As Object, ByVal e As FormClosingEventArgs) Handles MyBase.FormClosing
        Try
            If stopclose Then
                e.Cancel = True
            End If
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub
End Class