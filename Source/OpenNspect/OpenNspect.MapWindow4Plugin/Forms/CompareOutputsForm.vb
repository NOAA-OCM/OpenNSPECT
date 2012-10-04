'********************************************************************************************************
'File Name: frmCompareOutputs.vb
'Description: Output Comparison Tool
'********************************************************************************************************
'The contents of this file are subject to the Mozilla Public License Version 1.1 (the "License"); 
'you may not use this file except in compliance with the License. You may obtain a copy of the License at 
'http://www.mozilla.org/MPL/ 
'Software distributed under the License is distributed on an "AS IS" basis, WITHOUT WARRANTY OF 
'ANY KIND, either express or implied. See the License for the specificlanguage governing rights and 
'limitations under the License. 
'
'Contributor(s): (Open source contributors should list themselves and their modifications here). 
'Dec 20, 2010:  Allen Anselmo allen.anselmo@gmail.com - 
'               Added licensing and comments to code
Imports System.Collections.Generic
Imports System.Windows.Forms
Imports System.IO
Imports LegendControl
Imports MapWinGIS
Imports OpenNspect.Xml

Public Class CompareOutputsForm
    Private _SelectLyrPath As String
    Private _SelectedShapes As List(Of Integer)

#Region "Events"

    Private Sub frmCompareOutputs_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        Try
            g_comp = Me
            RefreshLeft()
            RefreshRight()
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub chkbxLeftUseLegend_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles chkbxLeftUseLegend.CheckedChanged
        Try
            RefreshLeft()
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub chkbxRightUseLegend_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles chkbxRightUseLegend.CheckedChanged
        Try
            RefreshRight()
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub mnuitmAddToLegend_Click(ByVal sender As Object, ByVal e As EventArgs) Handles mnuitmAddToLegend.Click
        Try
            AddToLegendFromProj()
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub mnuitmExit_Click(ByVal sender As Object, ByVal e As EventArgs) Handles mnuitmExit.Click
        Try
            Close()
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub cmdQuit_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdQuit.Click
        Try
            Close()
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub cmdRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdRun.Click
        Try
            RunCompare()
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

#End Region

#Region "Helper Functions"

    Public Sub SetSelectedShape()
        Try
            If MapWindowPlugin.MapWindowInstance.Layers.CurrentLayer <> -1 And MapWindowPlugin.MapWindowInstance.View.SelectedShapes.NumSelected > 0 Then
                chkSelectedPolys.Checked = True
                _SelectLyrPath = MapWindowPlugin.MapWindowInstance.Layers(MapWindowPlugin.MapWindowInstance.Layers.CurrentLayer).FileName
                _SelectedShapes = New List(Of Integer)
                For i As Integer = 0 To MapWindowPlugin.MapWindowInstance.View.SelectedShapes.NumSelected - 1
                    _SelectedShapes.Add(MapWindowPlugin.MapWindowInstance.View.SelectedShapes(i).ShapeIndex)
                Next
                lblSelected.Text = MapWindowPlugin.MapWindowInstance.View.SelectedShapes.NumSelected.ToString + " selected"
            End If
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub RefreshLeft()
        Try
            If chkbxLeftUseLegend.Checked Then
                RefreshUsingLegend(lstbxLeft)
            Else
                RefreshUsingProjectDirectory(lstbxLeft)
            End If
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub RefreshRight()
        Try
            If chkbxRightUseLegend.Checked Then
                RefreshUsingLegend(lstbxRight)
            Else
                RefreshUsingProjectDirectory(lstbxRight)
            End If
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub RefreshUsingLegend(ByRef RefreshBox As ListBox)
        Try
            Dim glyr As Layer
            Dim lyr As MapWindow.Interfaces.Layer
            RefreshBox.Items.Clear()

            For i As Integer = 0 To MapWindowPlugin.MapWindowInstance.Layers.Groups.Count - 1
                glyr = MapWindowPlugin.MapWindowInstance.Layers.Groups(i).Item(0)
                lyr = MapWindowPlugin.MapWindowInstance.Layers(glyr.Handle)
                If GetTypeFromPath(lyr.FileName, lyr.Name) <> "" Then
                    RefreshBox.Items.Add(MapWindowPlugin.MapWindowInstance.Layers.Groups(i).Text + " (Position: " + i.ToString + ")")
                End If
            Next
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub RefreshUsingProjectDirectory(ByRef RefreshBox As ListBox)
        Try
            Dim prj As ProjectFile
            Dim filesExist As Boolean
            Dim strFolder As String = g_nspectDocPath & "\projects"
            Dim ProjectsList As String() = Directory.GetFiles(strFolder)
            RefreshBox.Items.Clear()

            For i As Integer = 0 To ProjectsList.Length - 1
                If Path.GetExtension(ProjectsList(i)) = ".xml" Then
                    Try
                        prj = New ProjectFile
                        prj.Xml = ProjectsList(i)
                        If prj.OutputItems.Count > 0 Then
                            filesExist = True
                            For j As Integer = 0 To prj.OutputItems.Count - 1
                                If Not File.Exists(prj.OutputItems.Item(j).strPath) Then
                                    filesExist = False
                                    Exit For
                                End If
                            Next
                            If filesExist Then
                                RefreshBox.Items.Add(Path.GetFileName(ProjectsList(i)))
                            End If
                        End If
                    Catch ex As Exception
                        ' xml failed to load so not a valid project file
                        HandleError(ex)
                    End Try
                End If
            Next
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Sub AddToLegendFromProj()
        Try
            Dim strFolder As String = g_nspectDocPath & "\projects"
            If Not Directory.Exists(strFolder) Then
                MkDir(strFolder)
            End If
            Dim dlgXmlOpen As New OpenFileDialog
            With dlgXmlOpen
                .Filter = MSG8XmlFile
                .InitialDirectory = strFolder
                .Title = "Open OpenNSPECT Project File"
                .FilterIndex = 1
                .ShowDialog()
            End With

            Dim prj As ProjectFile
            Dim filesExist As Boolean
            Dim tmprast As Grid
            Dim outitem As OutputItem
            If Len(dlgXmlOpen.FileName) > 0 Then
                Try
                    prj = New ProjectFile
                    prj.Xml = dlgXmlOpen.FileName
                    If prj.OutputItems.Count > 0 Then
                        filesExist = True
                        For j As Integer = 0 To prj.OutputItems.Count - 1
                            outitem = prj.OutputItems.Item(j)
                            If Not File.Exists(outitem.strPath) Then
                                filesExist = False
                                Exit For
                            End If
                        Next
                        If filesExist Then
                            Dim tmpgrp As Integer = MapWindowPlugin.MapWindowInstance.Layers.Groups.Add(prj.ProjectName)
                            For j As Integer = 0 To prj.OutputItems.Count - 1
                                outitem = prj.OutputItems.Item(j)
                                tmprast = New Grid
                                tmprast.Open(outitem.strPath)
                                AddOutputGridLayer(tmprast, outitem.strColor, outitem.booUseStretch, outitem.strName, "", tmpgrp, Nothing)
                            Next
                            RefreshLeft()
                            RefreshRight()
                        Else
                            MsgBox("The outputs associated with that project file cannot be found.", MsgBoxStyle.Exclamation, "Missing Output Files")
                        End If
                    End If
                Catch ex As Exception
                    MsgBox("That project file doesn't seem to be valid. Please make sure you select a valid NSPECT project file.", MsgBoxStyle.Exclamation, "Project File Not Valid")
                End Try
            Else
                Return
            End If
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Function GetListFromSelected(ByRef SelectCheckbox As CheckBox, ByRef SelectList As ListBox) As OutputItems
        Try
            Dim tmpOutItems As New OutputItems
            Dim outitem As OutputItem
            Dim grpnum As Integer
            Dim strgrp As String
            Dim glyr As Layer
            Dim tmplyr As MapWindow.Interfaces.Layer
            Dim prj As ProjectFile
            Dim strFolder As String = g_nspectDocPath & "\projects"

            If SelectCheckbox.Checked Then
                strgrp = SelectList.SelectedItem
                grpnum = strgrp.Substring(strgrp.LastIndexOf(" ")).Replace(")", "")
                For i As Integer = 0 To MapWindowPlugin.MapWindowInstance.Layers.Groups(grpnum).LayerCount - 1
                    glyr = MapWindowPlugin.MapWindowInstance.Layers.Groups(grpnum).Item(i)
                    tmplyr = MapWindowPlugin.MapWindowInstance.Layers(glyr.Handle)
                    If tmplyr IsNot Nothing Then
                        outitem = New OutputItem
                        outitem.strPath = tmplyr.FileName
                        outitem.strName = tmplyr.Name
                        outitem.strType = GetTypeFromPath(tmplyr.FileName, tmplyr.Name)
                        tmpOutItems.Add(outitem)
                    End If

                Next
            Else
                Try
                    prj = New ProjectFile
                    prj.Xml = strFolder + Path.DirectorySeparatorChar + SelectList.SelectedItem
                    tmpOutItems = prj.OutputItems
                Catch ex As Exception
                    MsgBox("The item selected in the left list does not seem to be a valid project file.", MsgBoxStyle.Exclamation, "Compare Error")
                End Try
            End If

            Return tmpOutItems
        Catch ex As Exception
            HandleError(ex)
            GetListFromSelected = Nothing
        End Try
    End Function

    Private Function GetTypeFromPath(ByVal path As String, ByVal name As String) As String
        Try
            Dim filename As String = IO.Path.GetFileName(path)
            If filename = "sta.adf" Or filename = "sta.bmp" Then
                Dim dir As String() = IO.Path.GetDirectoryName(path).Split(IO.Path.DirectorySeparatorChar)
                filename = dir(dir.Length - 1)
            End If

            'strfile.StartsWith("locaccum") Or strfile.StartsWith("runoff") Or strfile.StartsWith("locconc") Or strfile.StartsWith("accpoll") Or strfile.StartsWith("conc") Or strfile.StartsWith("wq") Or strfile.StartsWith("locrusle") Or strfile.StartsWith("RUSLE") Or strfile.StartsWith("locmusle") Or strfile.StartsWith("MUSLEmass")
            If filename.StartsWith("locaccum") Then
                Return "Runoff Local"
            ElseIf filename.StartsWith("runoff") Then
                Return "Runoff Accum"
            ElseIf filename.StartsWith("locconc") Then
                Return "Pollutant " + name.Split(" ")(0) + " Local"
            ElseIf filename.StartsWith("accpoll") Then
                Return "Pollutant " + name.Split(" ")(1) + " Accum"
            ElseIf filename.StartsWith("conc") Then
                Return "Pollutant " + name.Split(" ")(0) + " Conc"
            ElseIf filename.StartsWith("wq") Then
                Return "Pollutant " + name.Split(" ")(0) + " WQ"
            ElseIf filename.StartsWith("locrusle") Then
                Return "RUSLE Local"
            ElseIf filename.StartsWith("RUSLE") Then
                Return "RUSLE Accum"
            ElseIf filename.StartsWith("locmusle") Then
                Return "MUSLE Local"
            ElseIf filename.StartsWith("MUSLEmass") Then
                Return "MUSLE Accum"
            Else
                Return ""
            End If
        Catch ex As Exception
            HandleError(ex)
            GetTypeFromPath = ""
        End Try
    End Function

    Private Sub RunCompare()
        Try
            Dim leftOutItems, rightOutItems As OutputItems
            Dim gleft, gright, compout, comppercout As Grid
            Dim gout As New Grid
            Dim outstring As String
            Dim outgrpnum As Integer = -1

            If lstbxLeft.SelectedIndex <> -1 And lstbxRight.SelectedIndex <> -1 Then

                leftOutItems = GetListFromSelected(chkbxLeftUseLegend, lstbxLeft)
                rightOutItems = GetListFromSelected(chkbxRightUseLegend, lstbxRight)

                gleft = New Grid
                gright = New Grid
                gleft.Open(leftOutItems.Item(0).strPath)
                gright.Open(rightOutItems.Item(0).strPath)

                If gleft.Header.XllCenter = gright.Header.XllCenter Then
                    gleft.Close()
                    gright.Close()
                    For i As Integer = 0 To leftOutItems.Count - 1
                        For j As Integer = 0 To rightOutItems.Count - 1
                            If leftOutItems.Item(i).strType = rightOutItems.Item(j).strType Then
                                gleft = New Grid
                                gright = New Grid
                                gleft.Open(leftOutItems.Item(i).strPath)
                                gright.Open(rightOutItems.Item(j).strPath)

                                If outgrpnum = -1 Then
                                    outgrpnum = MapWindowPlugin.MapWindowInstance.Layers.Groups.Add("Compare Outputs")
                                End If
                                If g_Project.ProjectWorkspace = "" Then
                                    'If String.IsNullOrEmpty(g_Project.ProjectWorkspace) Then
                                    g_Project.ProjectWorkspace = g_nspectDocPath & "\workspace"
                                End If

                                Dim strSelectedExportPath As String = ExportSelectedFeatures(_SelectLyrPath, _SelectedShapes)
                                Dim pSelectedPolyClip As Shape = ReturnSelectGeometry(strSelectedExportPath)

                                'Straight left minus right comparison
                                Dim compcalc As New RasterMathCellCalc(AddressOf CompareCellCalc)
                                RasterMath(gleft, gright, Nothing, Nothing, Nothing, gout, compcalc)

                                outstring = GetUniqueFileName("comp_base", g_Project.ProjectWorkspace, OutputGridExt)
                                If chkSelectedPolys.Checked Then
                                    compout = ClipBySelectedPoly(gout, pSelectedPolyClip, outstring)
                                Else
                                    compout = CopyRaster(gout, outstring)
                                End If
                                AddOutputGridLayer(compout, "DiffAbsolute", True, leftOutItems.Item(i).strType + " Direct Comparison", "", outgrpnum, Nothing)
                                'AddOutputGridLayer(compout, "Red", True, leftOutItems.Item(i).strType + " Direct Comparison", "", outgrpnum, Nothing)

                                Dim percchangecalc As New RasterMathCellCalc(AddressOf PercChangeCellCalc)
                                RasterMath(gleft, gright, Nothing, Nothing, Nothing, gout, percchangecalc)

                                outstring = GetUniqueFileName("comp_perc", g_Project.ProjectWorkspace, OutputGridExt)
                                If chkSelectedPolys.Checked Then
                                    comppercout = ClipBySelectedPoly(gout, pSelectedPolyClip, outstring)
                                Else
                                    comppercout = CopyRaster(gout, outstring)
                                End If
                                AddOutputGridLayer(comppercout, "DiffPercent", True, leftOutItems.Item(i).strType + " Percentage Change", "", outgrpnum, Nothing)

                                Exit For
                            End If
                        Next
                    Next
                    Close()
                Else
                    MsgBox("The two datasets selected cannot be compared due to misaligned grids.", MsgBoxStyle.Exclamation, "Compare Error")
                End If
            Else
                MsgBox("Please select an output set from both lists to compare.", MsgBoxStyle.Exclamation, "Compare Error")
            End If
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

#End Region

#Region "Raster Math"

    Private Function CompareCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        Try
            Return Input1 - Input2
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Function

    Private Function PercChangeCellCalc(ByVal Input1 As Single, ByVal Input2 As Single, ByVal Input3 As Single, ByVal Input4 As Single, ByVal Input5 As Single, ByVal OutNull As Single) As Single
        Try
            If Input2 <> 0 Then
                Return (100 * (Input1 - Input2)) / Input2
            Else
                Return OutNull
            End If
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Function

#End Region
End Class