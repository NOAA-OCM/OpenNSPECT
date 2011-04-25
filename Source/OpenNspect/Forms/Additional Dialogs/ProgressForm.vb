'********************************************************************************************************
'File Name: frmProgressDialog.vb
'Description: Form used for the progress dialog
'********************************************************************************************************
'The contents of this file are subject to the Mozilla Public License Version 1.1 (the "License"); 
'you may not use this file except in compliance with the License. You may obtain a copy of the License at 
'http://www.mozilla.org/MPL/ 
'Software distributed under the License is distributed on an "AS IS" basis, WITHOUT WARRANTY OF 
'ANY KIND, either express or implied. See the License for the specificlanguage governing rights and 
'limitations under the License. 
'
'Note: This code was converted from the vb6 NSPECT ArcGIS extension and so bears many of the old comments
'in the files where it was possible to leave them.
'
'Contributor(s): (Open source contributors should list themselves and their modifications here). 
'Oct 20, 2010:  Allen Anselmo allen.anselmo@gmail.com - 
'               Added licensing and comments to code

Public Class ProgressForm
    Const c_sModuleFileName As String = "frmProgressDialog.vb"


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


    Private _timerEnabled As Boolean


    Public Property TimerEnabled() As Boolean
        Get
            Return _timerEnabled
        End Get
        Set(ByVal value As Boolean)
            _timerEnabled = value
            If value Then
                tmrEventDriver.Start()
            Else
                tmrEventDriver.Stop()
            End If
        End Set
    End Property


    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            g_boolCancel = False
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub tmrEventDriver_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrEventDriver.Tick
        Try
            Windows.Forms.Application.DoEvents()
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub
End Class