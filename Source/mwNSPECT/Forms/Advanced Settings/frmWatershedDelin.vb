Imports System.Data.OleDb
Friend Class frmWatershedDelin
    Inherits System.Windows.Forms.Form

#Region "Events"
    Private Sub frmWatershedDelin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        modUtil.InitComboBox(cboWSDelin, "WSDELINEATION")
    End Sub

    Private Sub cboWSDelin_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboWSDelin.SelectedIndexChanged

        'String and recordset
        Dim strSQLDelin As String = "SELECT * FROM WSDELINEATION WHERE NAME LIKE '" & cboWSDelin.Text & "'"
        Dim delinCmd As New OleDbCommand(strSQLDelin, modUtil.g_DBConn)
        Dim delin As OleDbDataReader = delinCmd.ExecuteReader()

        'Check for records
        If delin.HasRows Then
            delin.Read()
            'Populate the controls...
            txtDEMFile.Text = delin.Item("DEMFileName")
            cboDEMUnits.SelectedIndex = delin.Item("DEMGridUnits")
            txtStream.Text = delin.Item("StreamFileName") & ""
            chkHydroCorr.CheckState = delin.Item("HydroCorrected")
            cboWSSize.SelectedIndex = delin.Item("SubWSSize")
            txtWSFile.Text = delin.Item("wsfilename") & ""
            txtFlowAccumGrid.Text = delin.Item("FlowAccumFileName") & ""
            txtLSGrid.Text = delin.Item("LSFileName") & ""
        Else
            MsgBox("Warning: There are no watershed delineation scenarios remaining.  Please add a new one.", MsgBoxStyle.Critical, "Recordset Empty")
        End If
    End Sub

    Private Sub cmdQuit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdQuit.Click
        Me.Close()
    End Sub

    Private Sub mnuNewWSDelin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuNewWSDelin.Click
        Dim newWS As New frmNewWSDelin
        newWS.ShowDialog()
    End Sub

    Private Sub mnuNewExist_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuNewExist.Click
        Dim newWS As New frmUserWShed
        newWS.ShowDialog()
    End Sub

    Private Sub mnuDelWSDelin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDelWSDelin.Click

        'Dim intAns As Object
        'Dim strSQLWSDel As String
        'Dim fldWSDelin As Scripting.FileSystemObject
        'Dim strFolder As String

        'strSQLWSDel = "SELECT * FROM WSDELINEATION WHERE NAME LIKE '" & cboWSDelin.Text & "'"

        'If Not (cboWSDelin.Text = "") Then
        '    'UPGRADE_WARNING: Couldn't resolve default property of object intAns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '    intAns = MsgBox("Are you sure you want to delete the watershed delineation scenario '" & VB6.GetItemString(cboWSDelin, cboWSDelin.SelectedIndex) & "'?", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Confirm Delete")
        '    'code to handle response
        '    If intAns = MsgBoxResult.Yes Then

        '        'Set up a delete rs and get rid of it
        '        rsWSDelinDelete = New ADODB.Recordset
        '        rsWSDelinDelete.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        '        rsWSDelinDelete.Open(strSQLWSDel, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockOptimistic)

        '        rsWSDelinDelete.Delete(ADODB.AffectEnum.adAffectCurrent)
        '        rsWSDelinDelete.Update()

        '        fldWSDelin = CreateObject("Scripting.FileSystemObject")
        '        strFolder = modUtil.g_nspectPath & "\wsdelin\" & cboWSDelin.Text
        '        If fldWSDelin.FolderExists(strFolder) Then
        '            fldWSDelin.DeleteFolder(strFolder, True)
        '        End If

        '        'Confirm
        '        MsgBox(VB6.GetItemString(cboWSDelin, cboWSDelin.SelectedIndex) & " deleted.", MsgBoxStyle.OkOnly, "Record Deleted")

        '        'Clear everything, clean up form
        '        cboWSDelin.Items.Clear()
        '        chkHydroCorr.CheckState = System.Windows.Forms.CheckState.Unchecked
        '        txtDEMFile.Text = ""
        '        txtStream.Text = ""
        '        txtWSFile.Text = ""
        '        txtFlowAccumGrid.Text = ""

        '        modUtil.InitComboBox(cboWSDelin, "WSDELINEATION")

        '        Me.Refresh()

        '    ElseIf intAns = MsgBoxResult.No Then
        '        Exit Sub
        '    End If
        'Else
        '    MsgBox("Please select a watershed delineation", MsgBoxStyle.Critical, "No Scenario Selected")
        'End If

        ''cleanup
        ''UPGRADE_NOTE: Object fldWSDelin may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        'fldWSDelin = Nothing
        'rsWSDelinDelete.Close()
        ''UPGRADE_NOTE: Object rsWSDelinDelete may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        'rsWSDelinDelete = Nothing

    End Sub

    Private Sub mnuWSDelin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuWSDelin.Click
        System.Windows.Forms.Help.ShowHelp(Me, modUtil.g_nspectPath & "\Help\nspect.chm", "wsdelin.htm")
    End Sub
#End Region

#Region "Helper Functions"

#End Region


End Class