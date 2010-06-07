Imports System.Windows.Forms
Imports System.Data.OleDb
Friend Class frmLUScen
    Inherits System.Windows.Forms.Form
    Private _strUndoText As String
    Private _clsManScen As clsXMLLUScen
    Private _strWQStd As String
    Private _frmPrj As frmProjectSetup
    Private _stopClose As Boolean

#Region "Events"
    Private Sub frmLUScen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cboLULayer.Items.Clear()
        For i As Integer = 0 To g_MapWin.Layers.NumLayers - 1
            If g_MapWin.Layers(i).LayerType = MapWindow.Interfaces.eLayerType.PolygonShapefile Then
                cboLULayer.Items.Add(g_MapWin.Layers(i).Name)
            End If
        Next

        FillGrid()

        _clsManScen = New clsXMLLUScen

        If Len(g_strLUScenFileName) > 0 Then
            _clsManScen.XML = g_strLUScenFileName
            PopulateForm()
        Else
            _txtLUCN_0.Text = "0"
            _txtLUCN_1.Text = "0"
            _txtLUCN_2.Text = "0"
            _txtLUCN_3.Text = "0"
        End If
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub frmLUScen_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        If _stopClose Then
            e.Cancel = True
            _stopClose = False
        End If
    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click

        If ValidateData() Then
            CreateXMLFile()
            Me.Close()
        Else
            _stopClose = True
        End If
    End Sub

#End Region

#Region "Helper Subs"
    Public Sub init(ByVal strWQStd As String, ByRef frmProj As frmProjectSetup)
        _frmPrj = frmProj
        _strWQStd = strWQStd
    End Sub

    Private Function CreateXMLFile() As String

        Dim clsMan As New clsXMLLUScen

        With clsMan

            .strLUScenName = Trim(txtLUName.Text)
            .strLUScenLyrName = Trim(cboLULayer.Text)
            .strLUScenFileName = modUtil.GetLayerFilename(.strLUScenLyrName)
            .intLUScenSelectedPoly = chkSelectedPolys.CheckState
            .intSCSCurveA = CDbl(_txtLUCN_0.Text)
            .intSCSCurveB = CDbl(_txtLUCN_1.Text)
            .intSCSCurveC = CDbl(_txtLUCN_2.Text)
            .intSCSCurveD = CDbl(_txtLUCN_3.Text)
            .lngCoverFactor = CDbl(_txtLUCN_4.Text)
            .intWaterWetlands = chkWatWetlands.CheckState

            For Each row As DataGridViewRow In dgvCoef.Rows
                clsMan.clsPollutant = New clsXMLLUScenPollItem
                .clsPollutant.intID = row.Index + 1
                .clsPollutant.strPollName = row.Cells("Pollutant").Value
                .clsPollutant.intType1 = CDbl(row.Cells("Type1").Value)
                .clsPollutant.intType2 = CDbl(row.Cells("Type2").Value)
                .clsPollutant.intType3 = CDbl(row.Cells("Type3").Value)
                .clsPollutant.intType4 = CDbl(row.Cells("Type4").Value)
                .clsPollItems.Add(.clsPollutant)

            Next
        End With

        _frmPrj.SetLURow(CInt(g_intManScenRow), clsMan.strLUScenName, clsMan.XML)
        CreateXMLFile = clsMan.XML

    End Function

    Private Function ValidateData() As Boolean

        'Project Name
        If Len(txtLUName.Text) = 0 Or Len(txtLUName.Text) > 30 Then
            MsgBox("Please enter a name for the scenario.  Names must be 30 characters or less.", MsgBoxStyle.Critical, "Enter Name")
            txtLUName.Focus()
            ValidateData = False
            Exit Function
        Else
            ValidateData = True
        End If

        'LandCoverLayer
        If cboLULayer.Text = "" Then
            MsgBox("Please select a layer before continuing.", MsgBoxStyle.Critical, "Select Layer")
            cboLULayer.Focus()
            ValidateData = False
            Exit Function
        Else
            If Not modUtil.LayerInMap(cboLULayer.Text) Then
                MsgBox("The layer you have choosen is not in the current map frame.", MsgBoxStyle.Critical, "Layer Not Found")
                ValidateData = False
                Exit Function
            End If
        End If

        'Check selected polygons
        If chkSelectedPolys.CheckState = 1 Then
            If g_MapWin.View.SelectedShapes.NumSelected = 0 Then
                MsgBox("You have chosen to use selected polygons from " & cboLULayer.Text & ", but the current map contains no selected features." & vbNewLine & "Please select features or N-SPECT will use the entire extent of " & cboLULayer.Text & " to apply this landuse scenario.", MsgBoxStyle.Information, "No Selected Features Found")
                ValidateData = False
            End If
        End If

        'SCS Curve Numbers
        If IsNumeric(Trim(_txtLUCN_0.Text)) Then
            If CShort(_txtLUCN_0.Text) > 0 Or CShort(_txtLUCN_0.Text) <= 1 Then
                ValidateData = True
            End If
        Else
            MsgBox("SCS Values are to be numeric only in the range of 0 - 1.", MsgBoxStyle.Critical, "Check SCS Values")
            ValidateData = False
            _txtLUCN_0.Focus()
            Exit Function
        End If
        If IsNumeric(Trim(_txtLUCN_1.Text)) Then
            If CShort(_txtLUCN_1.Text) > 0 Or CShort(_txtLUCN_1.Text) <= 1 Then
                ValidateData = True
            End If
        Else
            MsgBox("SCS Values are to be numeric only in the range of 0 - 1.", MsgBoxStyle.Critical, "Check SCS Values")
            ValidateData = False
            _txtLUCN_1.Focus()
            Exit Function
        End If
        If IsNumeric(Trim(_txtLUCN_2.Text)) Then
            If CShort(_txtLUCN_2.Text) > 0 Or CShort(_txtLUCN_2.Text) <= 1 Then
                ValidateData = True
            End If
        Else
            MsgBox("SCS Values are to be numeric only in the range of 0 - 1.", MsgBoxStyle.Critical, "Check SCS Values")
            ValidateData = False
            _txtLUCN_2.Focus()
            Exit Function
        End If
        If IsNumeric(Trim(_txtLUCN_3.Text)) Then
            If CShort(_txtLUCN_3.Text) > 0 Or CShort(_txtLUCN_3.Text) <= 1 Then
                ValidateData = True
            End If
        Else
            MsgBox("SCS Values are to be numeric only in the range of 0 - 1.", MsgBoxStyle.Critical, "Check SCS Values")
            ValidateData = False
            _txtLUCN_3.Focus()
            Exit Function
        End If
        If IsNumeric(Trim(_txtLUCN_4.Text)) Then
            If CShort(_txtLUCN_4.Text) > 0 Or CShort(_txtLUCN_4.Text) <= 1 Then
                ValidateData = True
            End If
        Else
            MsgBox("SCS Values are to be numeric only in the range of 0 - 1.", MsgBoxStyle.Critical, "Check SCS Values")
            ValidateData = False
            _txtLUCN_4.Focus()
            Exit Function
        End If
    End Function

    Private Sub PopulateForm()

        Dim strScenName As String
        Dim strLyrName As String
        Dim i As Short

        strScenName = _clsManScen.strLUScenName
        strLyrName = _clsManScen.strLUScenLyrName

        txtLUName.Text = strScenName

        If modUtil.LayerInMap(strLyrName) Then
            cboLULayer.SelectedIndex = modUtil.GetCboIndex(strLyrName, cboLULayer)
        End If

        chkSelectedPolys.CheckState = _clsManScen.intLUScenSelectedPoly

        _txtLUCN_0.Text = CStr(_clsManScen.intSCSCurveA)
        _txtLUCN_1.Text = CStr(_clsManScen.intSCSCurveB)
        _txtLUCN_2.Text = CStr(_clsManScen.intSCSCurveC)
        _txtLUCN_3.Text = CStr(_clsManScen.intSCSCurveD)
        _txtLUCN_4.Text = CStr(_clsManScen.lngCoverFactor)
        chkWatWetlands.CheckState = _clsManScen.intWaterWetlands

        dgvCoef.Rows.Clear()
        Dim idx As Integer
        For i = 0 To _clsManScen.clsPollItems.Count - 1
            With dgvCoef
                idx = .Rows.Add()
                .Rows(idx).Cells("Pollutant").Value = _clsManScen.clsPollItems.Item(i).strpollname
                .Rows(idx).Cells("Type1").Value = _clsManScen.clsPollItems.Item(i).intType1
                .Rows(idx).Cells("Type2").Value = _clsManScen.clsPollItems.Item(i).intType2
                .Rows(idx).Cells("Type3").Value = _clsManScen.clsPollItems.Item(i).intType3
                .Rows(idx).Cells("Type4").Value = _clsManScen.clsPollItems.Item(i).intType4
            End With
        Next i

    End Sub

    Private Sub FillGrid()

        Dim strSQLWQStd As String
        Dim strSQLWQStdPoll As String


        'Selection based on combo box
        strSQLWQStd = "SELECT * FROM WQCRITERIA WHERE NAME LIKE '" & _strWQStd & "'"
        Dim cmdWQStdCboClick As New OleDbCommand(strSQLWQStd, g_DBConn)
        Dim dataWQStd As OleDbDataReader = cmdWQStdCboClick.ExecuteReader()
        dataWQStd.Read()

        If dataWQStd.HasRows Then

            strSQLWQStdPoll = "SELECT POLLUTANT.NAME, POLL_WQCRITERIA.THRESHOLD " & "FROM POLL_WQCRITERIA INNER JOIN POLLUTANT " & "ON POLL_WQCRITERIA.POLLID = POLLUTANT.POLLID Where POLL_WQCRITERIA.WQCRITID = " & dataWQStd.Item("WQCRITID")
            Dim cmdSQLWQStdPoll As New OleDbCommand(strSQLWQStdPoll, g_DBConn)
            Dim dataWQStdPoll As OleDbDataReader = cmdSQLWQStdPoll.ExecuteReader()

            dgvCoef.Rows.Clear()
            Dim idx As Integer
            While dataWQStdPoll.Read()
                idx = dgvCoef.Rows.Add()
                dgvCoef.Rows(idx).Cells("Pollutant").Value = dataWQStdPoll.Item("Name")
                dgvCoef.Rows(idx).Cells("Type1").Value = 0
                dgvCoef.Rows(idx).Cells("Type2").Value = 0
                dgvCoef.Rows(idx).Cells("Type3").Value = 0
                dgvCoef.Rows(idx).Cells("Type4").Value = 0
            End While

            dataWQStd.Close()
            dataWQStdPoll.Close()
        Else
            MsgBox("Warning: There are no water quality standards remaining.  Please add a new one.", MsgBoxStyle.Critical, "Recordset Empty")
        End If

    End Sub
#End Region

End Class