Imports System.Data.OleDb
Imports System.Drawing

Friend Class frmProjectSetup
    Inherits System.Windows.Forms.Form

#Region "Class Vars"
    Private _strFileName As String 'Name of Open doc
    Private _strWorkspace As String 'String holding workspace, set it
    Private _strWShed As String 'String

    'Private _XMLPrjParams As clsXMLPrjFile 'xml doc that holds inputs

    Private _bolFirstLoad As Boolean 'Is initial Load event
    Private _booNew As Boolean 'New
    Private _booExists As Boolean 'Has file been saved

    Private _intMgmtCount As Short 'Count for management scenarios
    Private _intLUCount As Short 'Count for Land Use grid

    'Font DPI API
    Private Declare Function GetDC Lib "user32" (ByVal hwnd As Integer) As Integer
    Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Integer, ByVal hDC As Integer) As Integer
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Integer, ByVal nIndex As Integer) As Integer

    ' Win 32 Constant Declarations
    Private Const LOGPIXELSX As Short = 88 'Logical pixels/inch in X


    Const c_sModuleFileName As String = "frmProjectSetup.vb"
    Private m_ParentHWND As Integer 'Set this to get correct parenting of Error handler forms
#End Region

#Region "Events"

    Private Sub frmProjectSetup_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'Check the DPI setting: This was put in place because the alignment of the checkboxes will be
            'messed up if DPI settings are anything other than Normal
            Dim lngMapDC As Integer
            Dim lngDPI As Integer
            lngMapDC = GetDC(Me.Handle.ToInt32)
            lngDPI = GetDeviceCaps(lngMapDC, LOGPIXELSX)
            ReleaseDC(Me.Handle.ToInt32, lngMapDC)

            If lngDPI <> 96 Then
                MsgBox("Warning: N-SPECT requires your font size to be 96 DPI." & vbNewLine & "Some controls may appear out of alignment on this form.", MsgBoxStyle.Critical, "Warning!")
            End If

            _bolFirstLoad = True 'It's the first load
            _booExists = False

            'ComboBox::LandCover Type
            modUtil.InitComboBox(cboLCType, "LCType")

            'ComboBox::Precipitation Scenarios
            modUtil.InitComboBox(cboPrecipScen, "PrecipScenario")
            cboPrecipScen.Items.Insert(cboPrecipScen.Items.Count, "New precipitation scenario...")

            'ComboBox::WaterShed Delineations
            modUtil.InitComboBox(cboWSDelin, "WSDelineation")
            cboWSDelin.Items.Insert(cboWSDelin.Items.Count, "New watershed delineation...")

            'ComboBox::WaterQuality Criteria
            modUtil.InitComboBox(cboWQStd, "WQCriteria")
            cboWQStd.Items.Insert(cboWQStd.Items.Count, "New water quality standard...")

            'Fill Land Cover cbo
            'TODO:
            'modUtil.AddRasterLayerToComboBox(cboLCLayer, m_pMap)

            'Fill Rain GRID cbo
            'TODO:
            'modUtil.AddRasterLayerToComboBox(cboRainGrid, m_pMap)

            'Soils, now a 'scenario', not just a datalayer
            modUtil.InitComboBox(cboSoilsLayer, "Soils")

            'Fill area
            'TODO:
            'modUtil.AddFeatureLayerToComboBox(cboAreaLayer, m_pMap, "poly")

            'Fill LandClass
            FillCboLCCLass()

            '_intMgmtCount = grdLCChanges.Rows - 1 'Number of mgmt scens
            '_intLUCount = grdLU.Rows - 1 'Number of landuses

            'Initialize parameter file
            'TODO:
            'm_XMLPrjParams = New clsXMLPrjFile

            Me.Text = "Untitled"

            'Find out what the deal is
            cboSelectPoly.Items.Clear()
            'TODO:
            'modUtil.AddFeatureLayerToComboBox(cboSelectPoly, m_pMap, "poly")

            chkSelectedPolys.Enabled = EnableChkWaterShed

            'Test workspace persistence
            If Len(_strWorkspace) > 0 Then
                txtOutputWS.Text = _strWorkspace
            End If

        Catch ex As Exception
            HandleError(True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Sub


    Private Sub txtProjectName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtProjectName.TextChanged

    End Sub

    Private Sub txtOutputWS_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOutputWS.TextChanged

    End Sub

    Private Sub cmdOpenWS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOpenWS.Click

    End Sub

    Private Sub cboLCLayer_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboLCLayer.SelectedIndexChanged

    End Sub

    Private Sub cboLCUnits_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboLCUnits.SelectedIndexChanged

    End Sub

    Private Sub cboLCType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboLCType.SelectedIndexChanged

    End Sub

    Private Sub cboSoilsLayer_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSoilsLayer.SelectedIndexChanged

    End Sub

    Private Sub chkSelectedPolys_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSelectedPolys.CheckedChanged

    End Sub

    Private Sub cboSelectPoly_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSelectPoly.SelectedIndexChanged

    End Sub

    Private Sub chkLocalEffects_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLocalEffects.CheckedChanged

    End Sub

    Private Sub cboPrecipScen_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPrecipScen.SelectedIndexChanged

    End Sub

    Private Sub cboWSDelin_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboWSDelin.SelectedIndexChanged

    End Sub

    Private Sub cboWQStd_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboWQStd.SelectedIndexChanged

    End Sub


    Private Sub mnuNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuNew.Click
        Try
            Dim intvbYesNo As Short

            intvbYesNo = MsgBox("Do you want to save changes you made to " & Me.Text & "?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Exclamation, "N-SPECT")

            If intvbYesNo = MsgBoxResult.Yes Then
                If SaveXMLFile() Then
                    ClearForm()
                    frmProjectSetup_Load(Me, New System.EventArgs())
                Else
                    Exit Sub
                End If
            ElseIf intvbYesNo = MsgBoxResult.No Then
                ClearForm()
                frmProjectSetup_Load(Me, New System.EventArgs())
            ElseIf intvbYesNo = MsgBoxResult.Cancel Then
                Exit Sub
            End If

        Catch ex As Exception
            HandleError(True, "mnuNew_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Sub

    Private Sub mnuOpen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuOpen.Click
        LoadXMLFile()
    End Sub

    Private Sub mnuSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSave.Click
        Try
            SaveXMLFile()
        Catch ex As Exception
            HandleError(True, "mnuSave_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Sub

    Private Sub mnuSaveAs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSaveAs.Click
        Try
            _booExists = False
            SaveXMLFile()
        Catch ex As Exception
            HandleError(True, "mnuSaveAs_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Sub

    Private Sub mnuExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExit.Click
        Try
            Dim intvbYesNo As Short

            intvbYesNo = MsgBox("Do you want to save changes you made to " & Me.Text & "?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Exclamation, "N-SPECT")

            If intvbYesNo = MsgBoxResult.Yes Then
                If SaveXMLFile() Then
                    Me.Close()
                End If
            ElseIf intvbYesNo = MsgBoxResult.No Then
                Me.Close()
            ElseIf intvbYesNo = MsgBoxResult.Cancel Then
                Exit Sub
            End If

        Catch ex As Exception
            HandleError(True, "mnuExit_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Sub

    Private Sub mnuGeneralHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuGeneralHelp.Click
        System.Windows.Forms.Help.ShowHelp(Me, modUtil.g_nspectPath & "\Help\nspect.chm", "project_setup.htm")
    End Sub


    Private Sub cmdQuit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdQuit.Click
        Try
            Dim intvbYesNo As Short

            intvbYesNo = MsgBox("Do you want to save changes you made to " & Me.Text & "?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Exclamation, "N-SPECT")

            If intvbYesNo = MsgBoxResult.Yes Then
                If SaveXMLFile() Then
                    Me.Close()
                End If
            ElseIf intvbYesNo = MsgBoxResult.No Then
                Me.Close()
            ElseIf intvbYesNo = MsgBoxResult.Cancel Then
                Exit Sub
            End If


        Catch ex As Exception
            HandleError(True, "cmdQuit_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Sub

    Private Sub cmdRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRun.Click

    End Sub



    Private Sub cboLCLayer_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboWSDelin.KeyDown, cboWQStd.KeyDown, cboPrecipScen.KeyDown, cboLCUnits.KeyDown, cboLCType.KeyDown, cboLCLayer.KeyDown
        e.SuppressKeyPress = True
    End Sub
#End Region

#Region "Helper Functions"
    Private Sub ClearForm()
        '        On Error GoTo ErrorHandler

        '        'Gotta clean up before new, clean form

        '        ClearCheckBoxes(True)
        '        ClearMgmtCheckBoxes(True, m_intMgmtCount)
        '        ClearLUCheckBoxes(True)

        '        'LandClass stuff
        '        cboLCLayer.Items.Clear()
        '        cboLCType.Items.Clear()

        '        'DBase scens
        '        cboPrecipScen.Items.Clear()
        '        cboWSDelin.Items.Clear()
        '        cboWQStd.Items.Clear()
        '        cboSoilsLayer.Items.Clear()

        '        'Text
        '        txtProjectName.Text = ""
        '        txtOutputWS.Text = ""

        '        'Checkboxes
        '        chkSelectedPolys.CheckState = System.Windows.Forms.CheckState.Unchecked
        '        chkLocalEffects.CheckState = System.Windows.Forms.CheckState.Unchecked

        '        'Erosion
        '        cboRainGrid.Items.Clear()
        '        optUseGRID.Checked = True
        '        chkCalcErosion.CheckState = System.Windows.Forms.CheckState.Unchecked
        '        txtRainValue.Text = ""

        '        'txtOutputFile.Text = ""
        '        txtThemeName.Text = ""

        '        'clear the GRIDS
        '        grdLU.Clear()
        '        grdLU.Rows = 2
        '        grdLCChanges.Clear()
        '        grdLCChanges.Rows = 2
        '        grdCoeffs.Clear()



        '        Exit Sub
        'ErrorHandler:
        '        HandleError(False, "ClearForm " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
    End Sub

    Private Sub LoadXMLFile()
        ''On Error GoTo ErrorHandler

        ''browse...get output filename
        'Dim fso As Scripting.FileSystemObject
        'Dim strFolder As String

        'fso = New Scripting.FileSystemObject
        'strFolder = modUtil.g_nspectDocPath & "\projects"
        'If Not fso.FolderExists(strFolder) Then
        '    MkDir(strFolder)
        'End If

        'dlgXMLOpen.FileName = CStr(Nothing)
        'dlgXMLSave.FileName = CStr(Nothing)
        ''UPGRADE_WARNING: CommonDialog variable was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="671167DC-EA81-475D-B690-7A40C7BF4A23"'
        'With dlgXML
        '    'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        '    .Filter = MSG8
        '    .InitialDirectory = strFolder
        '    .Title = "Open N-SPECT Project File"
        '    .FilterIndex = 1
        '    'UPGRADE_WARNING: FileOpenConstants constant FileOpenConstants.cdlOFNHideReadOnly was upgraded to OpenFileDialog.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
        '    'UPGRADE_WARNING: MSComDlg.CommonDialog property dlgXML.Flags was upgraded to dlgXMLOpen.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
        '    'UPGRADE_WARNING: FileOpenConstants constant FileOpenConstants.cdlOFNHideReadOnly was upgraded to OpenFileDialog.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
        '    .ShowReadOnly = False
        '    'UPGRADE_WARNING: MSComDlg.CommonDialog property dlgXML.Flags was upgraded to dlgXMLOpen.CheckFileExists which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
        '    .CheckFileExists = True
        '    .CheckPathExists = True
        '    .CheckPathExists = True
        '    .ShowDialog()
        'End With

        'If Len(dlgXMLOpen.FileName) > 0 Then
        '    m_strFileName = Trim(dlgXMLOpen.FileName)
        '    m_XMLPrjParams.XML = m_strFileName
        '    FillForm()
        'Else
        '    Exit Sub
        'End If

        ''Pop this string with the incoming name, if they change, we'll prompt to 'save as'.
        'm_strOpenFileName = txtProjectName.Text

        ''UPGRADE_NOTE: Object fso may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        'fso = Nothing

        ''Exit Sub
        ''ErrorHandler:
        ''  HandleError False, "LoadXMLFile " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
    End Sub

    Private Function SaveXMLFile() As Boolean

        '        On Error GoTo ErrHandler

        '        Dim strFolder As String
        '        Dim intvbYesNo As Short
        '        Dim fso As Scripting.FileSystemObject

        '        fso = New Scripting.FileSystemObject

        '        strFolder = modUtil.g_nspectDocPath & "\projects"
        '        If Not fso.FolderExists(strFolder) Then
        '            MkDir(strFolder)
        '        End If

        '        If Not ValidateData Then 'check form inputs
        '            SaveXMLFile = False
        '            Exit Function
        '        End If

        '        'If it does not already exist, open Save As... dialog
        '        If Not m_booExists Then
        '            'UPGRADE_WARNING: CommonDialog variable was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="671167DC-EA81-475D-B690-7A40C7BF4A23"'
        '            With dlgXML
        '                'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        '                .Filter = MSG8
        '                .Title = "Save Project File As..."
        '                .InitialDirectory = strFolder
        '                .FilterIndex = 1
        '                'UPGRADE_WARNING: MSComDlg.CommonDialog property dlgXML.Flags was upgraded to dlgXMLSave.OverwritePrompt which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
        '                .OverwritePrompt = True
        '                'UPGRADE_WARNING: MSComDlg.CommonDialog property dlgXML.Flags was upgraded to dlgXMLOpen.CheckFileExists which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
        '                .CheckFileExists = True
        '                .CheckPathExists = True
        '                .CheckPathExists = True
        '                'UPGRADE_WARNING: The CommonDialog CancelError property is not supported in Visual Basic .NET. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"'
        '                .CancelError = True
        '                .FileName = txtProjectName.Text
        '                .ShowDialog()
        '            End With
        '            'check to make sure filename length is greater than zeros
        '            If Len(dlgXMLOpen.FileName) > 0 Then
        '                m_strFileName = Trim(dlgXMLOpen.FileName)
        '                m_booExists = True
        '                m_XMLPrjParams.SaveFile(m_strFileName)
        '                SaveXMLFile = True
        '            Else
        '                SaveXMLFile = False
        '                Exit Function
        '            End If

        '        Else
        '            'Now check to see if the name changed
        '            If m_strOpenFileName <> txtProjectName.Text Then
        '                intvbYesNo = MsgBox("You have changed the name of this project.  Would you like to save your settings as a new file?" & vbNewLine & vbTab & "Yes" & vbTab & " -    Save as new N-SPECT project file" & vbNewLine & vbTab & "No" & vbTab & " -    Save changes to current N-SPECT project file" & vbNewLine & vbTab & "Cancel" & vbTab & " -    Return to the project window", MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, "N-SPECT")

        '                If intvbYesNo = MsgBoxResult.Yes Then
        '                    'UPGRADE_WARNING: CommonDialog variable was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="671167DC-EA81-475D-B690-7A40C7BF4A23"'
        '                    With dlgXML
        '                        'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        '                        .Filter = MSG8
        '                        .Title = "Save Project File As..."
        '                        .InitialDirectory = strFolder
        '                        .FilterIndex = 1
        '                        'UPGRADE_WARNING: MSComDlg.CommonDialog property dlgXML.Flags was upgraded to dlgXMLSave.OverwritePrompt which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
        '                        .OverwritePrompt = True
        '                        'UPGRADE_WARNING: MSComDlg.CommonDialog property dlgXML.Flags was upgraded to dlgXMLOpen.CheckFileExists which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
        '                        .CheckFileExists = True
        '                        .CheckPathExists = True
        '                        .CheckPathExists = True
        '                        'UPGRADE_WARNING: The CommonDialog CancelError property is not supported in Visual Basic .NET. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"'
        '                        .CancelError = True
        '                        .FileName = txtProjectName.Text
        '                        .ShowDialog()
        '                    End With
        '                    'check to make sure filename length is greater than zeros
        '                    If Len(dlgXMLOpen.FileName) > 0 Then
        '                        m_strFileName = Trim(dlgXMLOpen.FileName)
        '                        m_booExists = True
        '                        m_XMLPrjParams.SaveFile(m_strFileName)
        '                        SaveXMLFile = True
        '                    Else
        '                        SaveXMLFile = False
        '                        Exit Function
        '                    End If
        '                ElseIf intvbYesNo = MsgBoxResult.No Then
        '                    m_XMLPrjParams.SaveFile(m_strFileName)
        '                    m_booExists = True
        '                    SaveXMLFile = True
        '                Else
        '                    SaveXMLFile = False
        '                    Exit Function
        '                End If
        '            Else
        '                m_XMLPrjParams.SaveFile(m_strFileName)
        '                m_booExists = True
        '                SaveXMLFile = True

        '            End If

        '        End If

        '        'UPGRADE_NOTE: Object fso may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '        fso = Nothing

        '        Exit Function

        'ErrHandler:

        '        If Err.Number = 32755 Then
        '            SaveXMLFile = False
        '            Exit Function
        '        Else
        '            MsgBox(Err.Number & " " & Err.Description)
        '            SaveXMLFile = False
        '        End If

    End Function


    Private Function EnableChkWaterShed() As Boolean

        Dim strWShed As String

        Try
            strWShed = "SELECT WSFILENAME FROM WSDELINEATION WHERE NAME LIKE '" & cboWSDelin.Text & "'"
            Dim wsCmd As New OleDbCommand(strWShed, modUtil.g_DBConn)
            Dim ws As OleDbDataReader = wsCmd.ExecuteReader()
            ws.Read()

            _strWShed = modUtil.SplitFileName(ws.Item("wsfilename"))

            'TODO: count selection
            'If _pMap.SelectionCount > 0 Then
            '    EnableChkWaterShed = True
            'Else
            '    EnableChkWaterShed = False
            'End If

        Catch ex As Exception
            EnableChkWaterShed = False
        End Try
    End Function

    Private Sub FillCboLCCLass()
        Try
            Dim strLCChanges As String
            Dim i As Short

            strLCChanges = "SELECT LCCLASS.Name as Name2, LCTYPE.LCTYPEID FROM LCTYPE INNER JOIN LCCLASS " & "ON LCTYPE.LCTYPEID = LCCLASS.LCTYPEID WHERE LCTYPE.NAME LIKE '" & cboLCType.Text & "'"
            Dim lcChangesCmd As New OleDbCommand(strLCChanges, modUtil.g_DBConn)
            Dim lcChanges As OleDbDataReader = lcChangesCmd.ExecuteReader()

            cboClass.Items.Clear()

            Do While lcChanges.Read()
                cboClass.Items.Insert(i, lcChanges.Item("Name2"))
            Loop

        Catch ex As Exception
            HandleError(False, "FillCboLCCLass " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Sub


#End Region


End Class