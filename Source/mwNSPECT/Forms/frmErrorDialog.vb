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

Public Class frmErrorDialog
    Private m_exception As Exception
    Friend WithEvents lblErr As System.Windows.Forms.Label
    Friend WithEvents txtComments As System.Windows.Forms.TextBox
    Friend WithEvents btnCopy As System.Windows.Forms.Button
    Const c_sModuleFileName As String = "frmErrorDialog.vb"

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="ex"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal ex As System.Exception)
        MyBase.New()
        Try

            'This call is required by the Windows Form Designer.
            InitializeComponent()

            m_exception = ex
        Catch ext As Exception
            HandleError(c_sModuleFileName, ext)
        End Try
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub cmdCopy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCopy.Click
        Try
            Windows.Forms.Clipboard.SetText(txtError.Text)
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub frmErrorDialog_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            txtError.Text = "Open N-SPECT (" + System.IO.File.GetLastWriteTime(System.Reflection.Assembly.GetExecutingAssembly().Location).ToShortDateString() + ")" + vbCrLf + vbCrLf + m_exception.ToString() ' + vbNewLine + vbNewLine + MapWinUtility.MiscUtils.GetDebugInfo()
            txtError.SelectionStart = txtError.Text.Length
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub txtError_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtError.KeyDown
        Try
            e.SuppressKeyPress = True
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub
End Class