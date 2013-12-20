'********************************************************************************************************
'File Name: DataPrepForm.vb
'Description: Form for Preprocessing data from various sources to a consistent projection 
'             and clipping to a target Area of Interest
'********************************************************************************************************
'The contents of this file are subject to the Mozilla Public License Version 1.1 (the "License"); 
'you may not use this file except in compliance with the License. You may obtain a copy of the License at 
'http://www.mozilla.org/MPL/ 
'Software distributed under the License is distributed on an "AS IS" basis, WITHOUT WARRANTY OF 
'ANY KIND, either express or implied. See the License for the specificlanguage governing rights and 
'limitations under the License. 
'
'Contributor(s): (Open source contributors should list themselves and their modifications here). 
'Dec. 18, 2013:  Dave Eslinger dave.eslinger@noaa.gov 
'      Initial commit of working code to repository

Imports System.IO
Imports MapWindow.Interfaces
Imports MapWinGIS
Imports MapWinGeoProc
Imports MapWindow
Imports MapWinUtility
Imports MapWindow.Controls.GisToolbox
Imports MapWindow.Controls.Data


Public Class PollCoeffSelectionShapefileForm
    Inherits BaseDialogForm

    Dim sf As New Shapefile
    Dim sfName As String
    Dim tmpField As String
    Dim indexField As String = ""
    Dim pollName As String

    Private Sub txtPollSFBrowse_Click(sender As Object, e As EventArgs) Handles txtPollSFBrowse.Click
        Dim sf As New Shapefile
        Dim sfName As String
        Dim tmpField As String
        Dim indexField As String = ""

        diaPollSFOpen.Reset()
        diaPollSFOpen.Title = "Open Index Shapefile for "
        diaPollSFOpen.Filter = "Shapefiles|*.shp"
        If diaPollSFOpen.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txtPollSF.Text = diaPollSFOpen.FileName
            sfName = txtPollSF.Text
            sf.Open(sfName)
            Try
                If sf.NumFields > 0 Then
                    Me.cmboPollAttrib.Items.Clear()
                    For field As Integer = 0 To sf.NumFields - 1
                        tmpField = sf.Field(field).Name
                        Me.cmboPollAttrib.Items.Add(tmpField)
                    Next

                    'If (Me.cmboPollAttrib.SelectedIndex > -1) Then
                    '    indexField = cmboPollAttrib.Text
                    '    MsgBox("The selected attribute is " & indexField & " with value of " & cmboPollAttrib.SelectedIndex.ToString _
                    '           & " and indexField is " & indexField)
                    'End If
                End If
            Catch ex As Exception

            End Try
        End If
    End Sub

    ' End Sub
    Protected Overrides Sub OK_Button_Click(ByVal sender As Object, ByVal e As EventArgs)
        sfName = Me.txtPollSF.Text
        MsgBox(sfName & " and " & indexField & " need to be stored for " & pollName)
    End Sub

    Public Sub New(Optional ByVal pName As String = "Specify shapefile for picking pollutant coefficients")
        'Call to MyBase.New must be the very first in a constructor.
        MyBase.New()
        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        Me.Text = "Specify shapefile for picking " & pName & " coefficients"
        pollName = pName
    End Sub

    Private Sub cmboPollAttrib_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmboPollAttrib.SelectedIndexChanged
        indexField = cmboPollAttrib.Text
        MsgBox("The selected attribute is " & indexField & " with value of " & cmboPollAttrib.SelectedIndex.ToString)
        'Verify range of Attributes

        'Then, if verified, store in XML file

    End Sub
End Class