'********************************************************************************************************
'File Name: frmSoilsSetup.vb
'Description: Form for hadling Soils setup
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

Imports System.Data.OleDb
Friend Class frmSoilsSetup
    Inherits System.Windows.Forms.Form

    Const c_sModuleFileName As String = "frmSoilsSetup.vb"

    Private _frmSoil As frmSoils
    Private _pRasterProps As MapWinGIS.Grid

#Region "Events"


    Private Sub cmdDEMBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDEMBrowse.Click
        Try
            'Browse for DEM
            Dim pDEMRasterDataset As MapWinGIS.Grid
            pDEMRasterDataset = AddInputFromGxBrowserText(txtDEMFile, "Choose DEM Dataset", Me, 0)

            If Not pDEMRasterDataset Is Nothing Then
                _pRasterProps = pDEMRasterDataset
            Else
                MsgBox("The Raster Dataset you have chosen is invalid.", MsgBoxStyle.Critical, "DEM Error")
                Exit Sub
            End If
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub cmdBrowseFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowseFile.Click
        Try
            'browse...get output filename
            Dim dlgOpen As New Windows.Forms.OpenFileDialog

            dlgOpen.Filter = MSG6
            dlgOpen.Title = "Open Soils Dataset"

            If dlgOpen.ShowDialog = Windows.Forms.DialogResult.OK Then
                txtSoilsDS.Text = Trim(dlgOpen.FileName)
                PopulateCbo()
            End If

        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub cmdQuit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdQuit.Click
        Try
            Dim intvbYesNo As Short

            intvbYesNo = MsgBox("Do you want to save changes you made to soils setup?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Exclamation, "N-SPECT")

            If intvbYesNo = MsgBoxResult.Yes Then
                SaveSoils()
            ElseIf intvbYesNo = MsgBoxResult.No Then
                Close()
            ElseIf intvbYesNo = MsgBoxResult.Cancel Then
                Exit Sub
            End If
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Try
            SaveSoils()
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub
#End Region

#Region "Helper"


    Public Sub Init(ByRef frmSoil As frmSoils)
        Try
            _frmSoil = frmSoil
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Sub SaveSoils()
        Try
            'Check data, if OK create soils grids
            If ValidateData() Then
                If CreateSoilsGrid(txtSoilsDS.Text, cboSoilFields.Text, cboSoilFieldsK.Text) Then
                    If _frmSoil.Visible Then
                        _frmSoil.cboSoils.Items.Clear()
                        modUtil.InitComboBox(_frmSoil.cboSoils, "Soils")
                        _frmSoil.cboSoils.SelectedIndex = modUtil.GetCboIndex(txtSoilsName.Text, _frmSoil.cboSoils)
                        Close()
                    End If
                Else
                    Exit Sub
                End If
            Else
                Exit Sub
            End If


        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub


    Private Function ValidateData() As Boolean
        Try
            If Len(txtSoilsName.Text) > 0 Then
                If modUtil.UniqueName("Soils", (txtSoilsName.Text)) Then
                    ValidateData = True
                Else
                    MsgBox("The name you have chosen is already in use.  Please select another.", MsgBoxStyle.Critical, "Select Unique Name")
                    ValidateData = False
                    txtSoilsName.Focus()
                    Exit Function
                End If
            Else
                MsgBox("Please enter a name.", MsgBoxStyle.Critical, "Soils Name Missing")
                ValidateData = False
                txtSoilsName.Focus()
                Exit Function

            End If

            If Len(txtSoilsDS.Text) = 0 Then
                MsgBox("Please select a soils dataset.", MsgBoxStyle.Critical, "Soils Dataset Missing")
                txtSoilsDS.Focus()
                ValidateData = False
                Exit Function
            Else
                ValidateData = True
            End If

            If Len(cboSoilFields.Text) = 0 Then
                MsgBox("Please select a soils attribute.", MsgBoxStyle.Critical, "Choose Soils Attribute")
                cboSoilFields.Focus()
                ValidateData = False
                Exit Function
            Else
                ValidateData = True
            End If

            If Len(cboSoilFieldsK.Text) = 0 Then
                MsgBox("Please select a k-factor soils attribute.", MsgBoxStyle.Critical, "Choose K-Factor Attribute")
                cboSoilFieldsK.Focus()
                ValidateData = False
                Exit Function
            Else
                ValidateData = True
            End If

            If Len(txtMUSLEVal.Text) > 0 Then
                If IsNumeric(CDbl(txtMUSLEVal.Text)) Then
                    ValidateData = True
                Else
                    MsgBox("Please enter a numeric value for the MUSLE equation.", MsgBoxStyle.Critical, "Numeric Values Only")
                    ValidateData = False
                End If
            Else
                MsgBox("Please enter a value for the MUSLE equation.", MsgBoxStyle.Critical, "Missing Value")
                txtMUSLEVal.Focus()
                ValidateData = False
            End If

            If Len(txtMUSLEExp.Text) > 0 Then
                If IsNumeric(CDbl(txtMUSLEExp.Text)) Then
                    ValidateData = True
                Else
                    MsgBox("Please enter a numeric value for the MUSLE equation.", MsgBoxStyle.Critical, "Numeric Values Only")
                    ValidateData = False
                End If
            Else
                MsgBox("Please enter a value for the MUSLE equation.", MsgBoxStyle.Critical, "Missing Value")
                txtMUSLEExp.Focus()
                ValidateData = False
            End If


        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Function


    Private Function CreateSoilsGrid(ByRef strSoilsFileName As String, ByRef strHydFieldName As String, Optional ByRef strKFactor As String = "") As Boolean
        'Incoming:
        'strSoilsFileName: string of soils file name path
        'strHydFieldName: string of hydrologic group attribute
        'strKFactor: string of K factor attribute

        Try
            Dim pSoilsFeatClass As MapWinGIS.Shapefile 'Soils Featureclass
            Dim lngHydFieldIndex As Integer 'HydGroup Field Index
            Dim lngNewHydFieldIndex As Integer 'New Group Field index
            Dim lngKFieldIndex As Integer 'K factor Field index
            Dim strHydValue As String 'Hyd Value to check
            Dim lngValue As Integer 'Count
            Dim strCmd As String 'String to insert new stuff in dbase
            Dim strOutSoils As String 'OutSoils name
            Dim strOutKSoils As String

            'Get the soils featurclass
            pSoilsFeatClass = modUtil.ReturnFeature(strSoilsFileName)

            'Check for fields
            lngKFieldIndex = -1
            For i As Integer = 0 To pSoilsFeatClass.NumFields - 1
                If pSoilsFeatClass.Field(i).Name = strKFactor Then
                    lngKFieldIndex = i
                    Exit For
                End If
            Next
            lngHydFieldIndex = -1
            For i As Integer = 0 To pSoilsFeatClass.NumFields - 1
                If pSoilsFeatClass.Field(i).Name = strHydFieldName Then
                    lngHydFieldIndex = i
                    Exit For
                End If
            Next
            lngNewHydFieldIndex = -1
            For i As Integer = 0 To pSoilsFeatClass.NumFields - 1
                If pSoilsFeatClass.Field(i).Name.ToUpper = "GROUP" Then
                    lngNewHydFieldIndex = i
                    Exit For
                End If
            Next


            'If the GROUP field is missing, we have to add it
            If lngNewHydFieldIndex = -1 Then
                Dim fieldedit As New MapWinGIS.Field
                fieldedit.Name = "GROUP"
                fieldedit.Type = MapWinGIS.FieldType.INTEGER_FIELD
                fieldedit.Width = 2
                lngNewHydFieldIndex = pSoilsFeatClass.NumFields
                pSoilsFeatClass.EditInsertField(fieldedit, lngNewHydFieldIndex)
            End If

            lngValue = 1

            pSoilsFeatClass.StartEditingTable()
            'Now calc the Values
            For i As Integer = 0 To pSoilsFeatClass.NumShapes - 1
                modProgDialog.ProgDialog("Calculating soils values...", "Processing Soils", 0, pSoilsFeatClass.NumShapes, lngValue, Me)
                'Find the current value
                If modProgDialog.g_boolCancel Then
                    strHydValue = pSoilsFeatClass.CellValue(lngHydFieldIndex, i)
                    'Based on current value, change GROUP to appropriate setting
                    Select Case strHydValue
                        Case "A"
                            pSoilsFeatClass.EditCellValue(lngNewHydFieldIndex, i, 1)
                        Case "B"
                            pSoilsFeatClass.EditCellValue(lngNewHydFieldIndex, i, 2)
                        Case "C"
                            pSoilsFeatClass.EditCellValue(lngNewHydFieldIndex, i, 3)
                        Case "D"
                            pSoilsFeatClass.EditCellValue(lngNewHydFieldIndex, i, 4)
                        Case "A/B"
                            pSoilsFeatClass.EditCellValue(lngNewHydFieldIndex, i, 2)
                        Case "B/C"
                            pSoilsFeatClass.EditCellValue(lngNewHydFieldIndex, i, 3)
                        Case "C/D"
                            pSoilsFeatClass.EditCellValue(lngNewHydFieldIndex, i, 4)
                        Case "B/D"
                            pSoilsFeatClass.EditCellValue(lngNewHydFieldIndex, i, 4)
                        Case "A/D"
                            pSoilsFeatClass.EditCellValue(lngNewHydFieldIndex, i, 1)
                        Case ""
                            MsgBox("Your soils dataset contains missing values for Hydrologic Soils Attribute.  Please correct.", MsgBoxStyle.Critical, "Missing Values Detected")
                            CreateSoilsGrid = False
                            modProgDialog.KillDialog()
                            Exit Function
                    End Select
                    lngValue = lngValue + 1
                Else
                    'If they cancel, kill the dialog
                    modProgDialog.KillDialog()
                    Exit Function
                End If
            Next
            pSoilsFeatClass.StopEditingTable()
            pSoilsFeatClass.Save()
            pSoilsFeatClass.Close()

            'Close dialog
            modProgDialog.KillDialog()

            'STEP 2:
            'Now do the conversion: Convert soils layer to GRID using new
            'Group field as the value

            If modProgDialog.g_boolCancel Then
                modProgDialog.ProgDialog("Converting Soils Dataset...", "Processing Soils", 0, 2, 2, Me)

                strOutSoils = modUtil.GetUniqueName("soils", IO.Path.GetDirectoryName(strSoilsFileName), g_OutputGridExt)

                'Hand convert the soils shapefile to grids by creating new grids based on header of dem
                Dim dem As New MapWinGIS.Grid
                dem.Open(txtDEMFile.Text)
                Dim head As New MapWinGIS.GridHeader
                Dim headK As New MapWinGIS.GridHeader
                head.CopyFrom(dem.Header)
                headK.CopyFrom(dem.Header)
                dem.Close()

                Dim soilsshp As New MapWinGIS.Shapefile
                soilsshp.Open(strSoilsFileName)
                soilsshp.BeginPointInShapefile()

                Dim outSoils As New MapWinGIS.Grid
                Dim outSoilsK As New MapWinGIS.Grid

                outSoils.CreateNew(strOutSoils, head, MapWinGIS.GridDataType.DoubleDataType, head.NodataValue)

                If Len(strKFactor) > 0 Then
                    strOutKSoils = modUtil.GetUniqueName("soilsk", IO.Path.GetDirectoryName(strSoilsFileName), g_OutputGridExt)
                    outSoilsK.CreateNew(strOutKSoils, headK, MapWinGIS.GridDataType.DoubleDataType, head.NodataValue)
                Else
                    strOutKSoils = ""
                End If

                'Then cycle over each cell, testing the cell to proj for intersection with the shapefile. very cheesy, but should work
                Dim x, y As Double
                Dim idx As Integer
                Dim nc As Integer = head.NumberCols - 1
                Dim nr As Integer = head.NumberRows - 1
                For row As Integer = 0 To nr
                    modProgDialog.ProgDialog("Converting Soils Dataset...", "Processing Soils", 1, nr, row, Me)
                    For col As Integer = 0 To nc
                        outSoils.CellToProj(col, row, x, y)
                        idx = soilsshp.PointInShapefile(x, y)
                        If idx <> -1 Then
                            outSoils.Value(col, row) = soilsshp.CellValue(lngNewHydFieldIndex, idx)
                            If strOutKSoils <> "" Then
                                outSoilsK.Value(col, row) = soilsshp.CellValue(lngKFieldIndex, idx)
                            End If
                        End If
                    Next
                Next

                'After, save the grids and close them.
                outSoils.Save()
                outSoils.Close()
                If strOutKSoils <> "" Then
                    outSoilsK.Save()
                    outSoilsK.Close()
                End If

                soilsshp.EndPointInShapefile()
                soilsshp.Close()
                'STEP 4:
                'Now enter all into database
                strCmd = "INSERT INTO SOILS (NAME,SOILSFILENAME,SOILSKFILENAME,MUSLEVal,MUSLEExp) VALUES ('" & Replace(txtSoilsName.Text, "'", "''") & "', '" & Replace(strOutSoils, "'", "''") & "', '" & Replace(strOutKSoils, "'", "''") & "', " & CDbl(txtMUSLEVal.Text) & ", " & CDbl(txtMUSLEExp.Text) & ")"
                Dim cmdIns As New OleDbCommand(strCmd, g_DBConn)
                cmdIns.ExecuteNonQuery()

                modProgDialog.KillDialog()

            Else
                modProgDialog.KillDialog()
            End If

            CreateSoilsGrid = True
        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
            'MsgBox(Err.Number & ": " & Err.Description)
            CreateSoilsGrid = False
        End Try
    End Function


    Private Sub PopulateCbo()
        Try
            'Populate cboSoilFields & cboSoilFieldsK with the fields in the selected Soils layer
            Dim i As Short
            Dim pFeatureClass As MapWinGIS.Shapefile

            cboSoilFields.Items.Clear()
            cboSoilFieldsK.Items.Clear()

            pFeatureClass = modUtil.ReturnFeature(txtSoilsDS.Text)

            'Pop both cbos with field names
            For i = 0 To pFeatureClass.NumFields - 1
                cboSoilFields.Items.Add(pFeatureClass.Field(i).Name)
                cboSoilFieldsK.Items.Add(pFeatureClass.Field(i).Name)
            Next i

        Catch ex As Exception
            HandleError(c_sModuleFileName, ex)
        End Try
    End Sub
#End Region

End Class