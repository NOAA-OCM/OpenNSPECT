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

                    If (Me.cmboPollAttrib.SelectedIndex > 0) Then
                        indexField = cmboPollAttrib.Text
                    End If
                    MsgBox("The selected attribute is " & indexField)
               End If
            Catch ex As Exception

            End Try
        End If
    End Sub
End Class