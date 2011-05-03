'********************************************************************************************************
'File Name: frmErrorDialog.vb
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
Imports System.Reflection
Imports System.IO

Public Class ErrorForm
    Private ReadOnly OriginalException As Exception
    Friend WithEvents lblErr As Label
    Friend WithEvents txtComments As TextBox
    Friend WithEvents btnCopy As Button

    Public Sub New(ByVal ex As Exception)
        MyBase.New()
        'This call is required by the Windows Form Designer.
        InitializeComponent()

        OriginalException = ex
    End Sub

    Private Sub cmdCopy_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdCopy.Click
        Try
            Clipboard.SetText(txtError.Text)
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub cmdClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdClose.Click
        Close()
    End Sub

    Private Sub frmErrorDialog_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        Try
            txtError.Text = String.Format("OpenNSPECT ({0}){1}{1}{2}", File.GetLastWriteTime(Assembly.GetExecutingAssembly().Location).ToShortDateString(), vbCrLf, OriginalException)
            txtError.SelectionStart = txtError.Text.Length
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub
End Class