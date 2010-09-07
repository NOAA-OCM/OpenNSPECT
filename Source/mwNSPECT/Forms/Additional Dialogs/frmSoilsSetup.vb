Imports System.Data.OleDb
Friend Class frmSoilsSetup
    Inherits System.Windows.Forms.Form

    Private _frmSoil As frmSoils
    Private _pRasterProps As MapWinGIS.Grid
   
#Region "Events"
    Private Sub cmdDEMBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDEMBrowse.Click
        'Browse for DEM
        Dim pDEMRasterDataset As MapWinGIS.Grid
        pDEMRasterDataset = AddInputFromGxBrowserText(txtDEMFile, "Choose DEM Dataset", Me, 0)

        If Not pDEMRasterDataset Is Nothing Then
            _pRasterProps = pDEMRasterDataset
        Else
            MsgBox("The Raster Dataset you have chosen is invalid.", MsgBoxStyle.Critical, "DEM Error")
            Exit Sub
        End If
    End Sub

    Private Sub cmdBrowseFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowseFile.Click
        'browse...get output filename
        Dim dlgOpen As New Windows.Forms.OpenFileDialog

        dlgOpen.Filter = MSG6
        dlgOpen.Title = "Open Soils Dataset"

        If dlgOpen.ShowDialog = Windows.Forms.DialogResult.OK Then
            txtSoilsDS.Text = Trim(dlgOpen.FileName)
            PopulateCbo()
        End If

    End Sub

    Private Sub cmdQuit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdQuit.Click

        Dim intvbYesNo As Short

        intvbYesNo = MsgBox("Do you want to save changes you made to soils setup?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Exclamation, "N-SPECT")

        If intvbYesNo = MsgBoxResult.Yes Then
            SaveSoils()
        ElseIf intvbYesNo = MsgBoxResult.No Then
            Me.Close()
        ElseIf intvbYesNo = MsgBoxResult.Cancel Then
            Exit Sub
        End If
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        SaveSoils()
    End Sub
#End Region

#Region "Helper"
    Public Sub Init(ByRef frmSoil As frmSoils)
        _frmSoil = frmSoil
    End Sub

    Private Sub SaveSoils()

        'Check data, if OK create soils grids
        If ValidateData() Then
            If CreateSoilsGrid((txtSoilsDS.Text), (cboSoilFields.Text), cboSoilFieldsK.Text) Then
                If _frmSoil.Visible Then
                    _frmSoil.cboSoils.Items.Clear()
                    modUtil.InitComboBox(_frmSoil.cboSoils, "Soils")
                    _frmSoil.cboSoils.SelectedIndex = modUtil.GetCboIndex(txtSoilsName.Text, _frmSoil.cboSoils)
                    Me.Close()
                End If
            Else
                Exit Sub
            End If
        Else
            Exit Sub
        End If


    End Sub

    Private Function ValidateData() As Boolean

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
            Dim strHydValue As String 'Hyd Value to check
            Dim lngValue As Integer 'Count
            Dim strCmd As String 'String to insert new stuff in dbase
            Dim strOutSoils As String 'OutSoils name
            Dim strOutKSoils As String

            'Get the soils featurclass
            pSoilsFeatClass = modUtil.ReturnFeature(strSoilsFileName)

            'Check for fields
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
                modProgDialog.ProgDialog("Calculating soils values...", "Processing Soils", 0, pSoilsFeatClass.NumShapes, lngValue, 0)
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

            'TODO: Test if this is working
            If modProgDialog.g_boolCancel Then
                strOutSoils = modUtil.GetUniqueName("soils", IO.Path.GetDirectoryName(strSoilsFileName), ".bgd")

                Dim cellsize As Double = _pRasterProps.Header.dX
                MapWinGeoProc.Utils.ShapefileToGrid(strSoilsFileName, strOutSoils, MapWinGIS.GridFileType.UseExtension, MapWinGIS.GridDataType.ShortDataType, "GROUP", cellsize, Nothing)

                'STEP 3:
                'Now do the conversion: Convert soils layer to GRID using
                'k factor field as the value
                'If they are doing a K factor then repeat the process, this time using 'k' field
                If Len(strKFactor) > 0 Then
                    modProgDialog.ProgDialog("Converting Soils K Dataset...", "Processing Soils", 0, 2, 2, 0)

                    strOutKSoils = modUtil.GetUniqueName("soilsk", IO.Path.GetDirectoryName(strSoilsFileName), ".bgd")

                    MapWinGeoProc.Utils.ShapefileToGrid(strSoilsFileName, strOutKSoils, MapWinGIS.GridFileType.UseExtension, MapWinGIS.GridDataType.DoubleDataType, strKFactor, cellsize, Nothing)
                Else
                    strOutKSoils = ""
                    modProgDialog.KillDialog()
                End If

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
            MsgBox(Err.Number & ": " & Err.Description)
            CreateSoilsGrid = False
        End Try
    End Function


    Private Sub PopulateCbo()

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

    End Sub
#End Region

End Class