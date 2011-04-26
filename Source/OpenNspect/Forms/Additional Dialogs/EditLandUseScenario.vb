'********************************************************************************************************
'File Name: frmLUScen.vb
'Description: Form for entering Land Use Scenarios
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
Imports System.Collections.Generic
Imports System.Windows.Forms
Imports MapWindow.Interfaces
Imports System.Data.OleDb

Friend Class EditLandUseScenario
    Private _ManScen As XMLLUMangementScenario
    Private _strWQStd As String
    Private _frmPrj As MainForm
    Private _stopClose As Boolean
    Private _SelectLyrPath As String
    Private _SelectedShapes As List(Of Integer)

#Region "Events"

    Private Sub frmLUScen_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        Try
            g_luscen = Me
            cboLULayer.Items.Clear()
            For i As Integer = 0 To g_MapWin.Layers.NumLayers - 1
                If g_MapWin.Layers(i).LayerType = eLayerType.PolygonShapefile Then
                    cboLULayer.Items.Add(g_MapWin.Layers(i).Name)
                End If
            Next

            FillGrid()

            _ManScen = New XMLLUMangementScenario

            If Len(g_strLUScenFileName) > 0 Then
                _ManScen.XML = g_strLUScenFileName
                PopulateForm()
            Else
                _txtLUCN_0.Text = "0"
                _txtLUCN_1.Text = "0"
                _txtLUCN_2.Text = "0"
                _txtLUCN_3.Text = "0"
            End If
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub frmLUScen_FormClosing(ByVal sender As Object, _
                                       ByVal e As FormClosingEventArgs) Handles MyBase.FormClosing
        Try
            If _stopClose Then
                e.Cancel = True
                _stopClose = False
            End If
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Protected Overrides Sub OK_Button_Click(sender As Object, e As EventArgs)
        Try
            If ValidateData() Then
                CreateXmlFile()
                MyBase.OK_Button_Click(sender, e)
            Else
                _stopClose = True
            End If
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub btnSelect_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSelect.Click
        Try
            Dim selectfrm As New SelectionModeForm
            selectfrm.InitializeAndShow()
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub cboLULayer_MouseClick(ByVal sender As Object, ByVal e As MouseEventArgs) _
        Handles cboLULayer.MouseClick
        Try
            _SelectLyrPath = ""
            _SelectedShapes.Clear()
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

#End Region

#Region "Helper Subs"

    Public Sub init(ByVal strWQStd As String, ByRef frmProj As MainForm)
        Try
            _frmPrj = frmProj
            _strWQStd = strWQStd
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Function CreateXmlFile() As String
        Try
            Dim Man As New XMLLUMangementScenario

            With Man

                .strLUScenName = Trim(txtLUName.Text)
                .strLUScenLyrName = Trim(cboLULayer.Text)
                .strLUScenFileName = GetLayerFilename(.strLUScenLyrName)
                .intLUScenSelectedPoly = chkSelectedPolys.CheckState
                .intLUScenSelectedPolyList = _SelectedShapes
                .intSCSCurveA = CDbl(_txtLUCN_0.Text)
                .intSCSCurveB = CDbl(_txtLUCN_1.Text)
                .intSCSCurveC = CDbl(_txtLUCN_2.Text)
                .intSCSCurveD = CDbl(_txtLUCN_3.Text)
                .lngCoverFactor = CDbl(_txtLUCN_4.Text)
                .intWaterWetlands = chkWatWetlands.CheckState

                For Each row As DataGridViewRow In dgvCoef.Rows
                    Man.Pollutant = New XMLLUScenPollItem
                    .Pollutant.intID = row.Index + 1
                    .Pollutant.strPollName = row.Cells("Pollutant").Value
                    .Pollutant.intType1 = CDbl(row.Cells("Type1").Value)
                    .Pollutant.intType2 = CDbl(row.Cells("Type2").Value)
                    .Pollutant.intType3 = CDbl(row.Cells("Type3").Value)
                    .Pollutant.intType4 = CDbl(row.Cells("Type4").Value)
                    .PollItems.Add(.Pollutant)

                Next
            End With

            _frmPrj.SetLURow(CInt(g_intManScenRow), Man.strLUScenName, Man.XML)
            CreateXmlFile = Man.XML

        Catch ex As Exception
            HandleError(ex)
            CreateXmlFile = Nothing
        End Try
    End Function

    Private Function ValidateData() As Boolean
        Try
            'Project Name
            If Len(txtLUName.Text) = 0 Or Len(txtLUName.Text) > 30 Then
                MsgBox("Please enter a name for the scenario.  Names must be 30 characters or less.", _
                        MsgBoxStyle.Critical, "Enter Name")
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
                If Not LayerInMap(cboLULayer.Text) Then
                    MsgBox("The layer you have choosen is not in the current map frame.", MsgBoxStyle.Critical, _
                            "Layer Not Found")
                    ValidateData = False
                    Exit Function
                End If
            End If

            'Check selected polygons
            If chkSelectedPolys.CheckState = 1 Then
                If _SelectLyrPath = "" Or _SelectedShapes.Count = 0 Then
                    MsgBox( _
                            "You have chosen to use selected polygons from " & cboLULayer.Text & _
                            ", but there are no selected features." & vbNewLine & _
                            "Please select features or OpenNSPECT will use the entire extent of " & cboLULayer.Text & _
                            " to apply this landuse scenario.", MsgBoxStyle.Information, "No Selected Features Found")
                    ValidateData = False
                End If
            End If

            'SCS Curve Numbers
            If IsNumeric(Trim(_txtLUCN_0.Text)) Then
                If CShort(_txtLUCN_0.Text) > 0 Or CShort(_txtLUCN_0.Text) <= 1 Then
                    ValidateData = True
                End If
            Else
                MsgBox("SCS Values are to be numeric only in the range of 0 - 1.", MsgBoxStyle.Critical, _
                        "Check SCS Values")
                ValidateData = False
                _txtLUCN_0.Focus()
                Exit Function
            End If
            If IsNumeric(Trim(_txtLUCN_1.Text)) Then
                If CShort(_txtLUCN_1.Text) > 0 Or CShort(_txtLUCN_1.Text) <= 1 Then
                    ValidateData = True
                End If
            Else
                MsgBox("SCS Values are to be numeric only in the range of 0 - 1.", MsgBoxStyle.Critical, _
                        "Check SCS Values")
                ValidateData = False
                _txtLUCN_1.Focus()
                Exit Function
            End If
            If IsNumeric(Trim(_txtLUCN_2.Text)) Then
                If CShort(_txtLUCN_2.Text) > 0 Or CShort(_txtLUCN_2.Text) <= 1 Then
                    ValidateData = True
                End If
            Else
                MsgBox("SCS Values are to be numeric only in the range of 0 - 1.", MsgBoxStyle.Critical, _
                        "Check SCS Values")
                ValidateData = False
                _txtLUCN_2.Focus()
                Exit Function
            End If
            If IsNumeric(Trim(_txtLUCN_3.Text)) Then
                If CShort(_txtLUCN_3.Text) > 0 Or CShort(_txtLUCN_3.Text) <= 1 Then
                    ValidateData = True
                End If
            Else
                MsgBox("SCS Values are to be numeric only in the range of 0 - 1.", MsgBoxStyle.Critical, _
                        "Check SCS Values")
                ValidateData = False
                _txtLUCN_3.Focus()
                Exit Function
            End If
            If IsNumeric(Trim(_txtLUCN_4.Text)) Then
                If CShort(_txtLUCN_4.Text) > 0 Or CShort(_txtLUCN_4.Text) <= 1 Then
                    ValidateData = True
                End If
            Else
                MsgBox("SCS Values are to be numeric only in the range of 0 - 1.", MsgBoxStyle.Critical, _
                        "Check SCS Values")
                ValidateData = False
                _txtLUCN_4.Focus()
                Exit Function
            End If
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Function

    Private Sub PopulateForm()
        Try
            Dim strScenName As String
            Dim strLyrName As String
            Dim i As Short

            strScenName = _ManScen.strLUScenName
            strLyrName = _ManScen.strLUScenLyrName

            txtLUName.Text = strScenName

            If LayerInMap(strLyrName) Then
                cboLULayer.SelectedIndex = GetCboIndex(strLyrName, cboLULayer)
            End If

            chkSelectedPolys.CheckState = _ManScen.intLUScenSelectedPoly

            _SelectLyrPath = GetLayerFilename(_ManScen.strLUScenLyrName)
            _SelectedShapes = _ManScen.intLUScenSelectedPolyList
            lblSelected.Text = _SelectedShapes.Count.ToString + " selected"

            _txtLUCN_0.Text = CStr(_ManScen.intSCSCurveA)
            _txtLUCN_1.Text = CStr(_ManScen.intSCSCurveB)
            _txtLUCN_2.Text = CStr(_ManScen.intSCSCurveC)
            _txtLUCN_3.Text = CStr(_ManScen.intSCSCurveD)
            _txtLUCN_4.Text = CStr(_ManScen.lngCoverFactor)
            chkWatWetlands.CheckState = _ManScen.intWaterWetlands

            dgvCoef.Rows.Clear()
            Dim idx As Integer
            For i = 0 To _ManScen.PollItems.Count - 1
                With dgvCoef
                    idx = .Rows.Add()
                    .Rows(idx).Cells("Pollutant").Value = _ManScen.PollItems.Item(i).strpollname
                    .Rows(idx).Cells("Type1").Value = _ManScen.PollItems.Item(i).intType1
                    .Rows(idx).Cells("Type2").Value = _ManScen.PollItems.Item(i).intType2
                    .Rows(idx).Cells("Type3").Value = _ManScen.PollItems.Item(i).intType3
                    .Rows(idx).Cells("Type4").Value = _ManScen.PollItems.Item(i).intType4
                End With
            Next i

        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub FillGrid()
        Try
            Dim strSQLWQStd As String
            Dim strSQLWQStdPoll As String

            'Selection based on combo box
            strSQLWQStd = "SELECT * FROM WQCRITERIA WHERE NAME LIKE '" & _strWQStd & "'"
            Dim cmdWQStdCboClick As New DataHelper(strSQLWQStd)
            Dim dataWQStd As OleDbDataReader = cmdWQStdCboClick.ExecuteReader()
            dataWQStd.Read()

            If dataWQStd.HasRows Then

                strSQLWQStdPoll = "SELECT POLLUTANT.NAME, POLL_WQCRITERIA.THRESHOLD " & _
                                  "FROM POLL_WQCRITERIA INNER JOIN POLLUTANT " & _
                                  "ON POLL_WQCRITERIA.POLLID = POLLUTANT.POLLID Where POLL_WQCRITERIA.WQCRITID = " & _
                                  dataWQStd.Item("WQCRITID")
                Dim cmdSQLWQStdPoll As New DataHelper(strSQLWQStdPoll)
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
                MsgBox("Warning: There are no water quality standards remaining.  Please add a new one.", _
                        MsgBoxStyle.Critical, "Recordset Empty")
            End If

        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Public Sub SetSelectedShape()
        Try
            If g_MapWin.Layers.CurrentLayer <> -1 And g_MapWin.View.SelectedShapes.NumSelected > 0 Then
                chkSelectedPolys.Checked = True
                cboLULayer.Items.Clear()
                For i As Integer = 0 To g_MapWin.Layers.NumLayers - 1
                    If g_MapWin.Layers(i).LayerType = eLayerType.PolygonShapefile Then
                        cboLULayer.Items.Add(g_MapWin.Layers(i).Name)
                    End If
                Next
                cboLULayer.SelectedIndex = _
                    GetCboIndex(g_MapWin.Layers(g_MapWin.Layers.CurrentLayer).Name, cboLULayer)

                _SelectLyrPath = g_MapWin.Layers(g_MapWin.Layers.CurrentLayer).FileName
                _SelectedShapes = New List(Of Integer)
                For i As Integer = 0 To g_MapWin.View.SelectedShapes.NumSelected - 1
                    _SelectedShapes.Add(g_MapWin.View.SelectedShapes(i).ShapeIndex)
                Next
                lblSelected.Text = g_MapWin.View.SelectedShapes.NumSelected.ToString + " selected"
            End If
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

#End Region
End Class