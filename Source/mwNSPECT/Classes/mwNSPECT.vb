'********************************************************************************************************
'File Name: mwNSPECT.vb
'Description: This class initializes and controls the plugin behavior on the MW menu
'********************************************************************************************************
'The contents of this file are subject to the Mozilla Public License Version 1.1 (the "License"); 
'you may not use this file except in compliance with the License. You may obtain a copy of the License at 
'http://www.mozilla.org/MPL/ 
'Software distributed under the License is distributed on an "AS IS" basis, WITHOUT WARRANTY OF 
'ANY KIND, either express or implied. See the License for the specificlanguage governing rights and 
'limitations under the License. 
'
'Contributor(s): (Open source contributors should list themselves and their modifications here). 
'Nov 23, 2009:  Allen Anselmo allen.anselmo@gmail.com - 
'               Added licensing and comments to original code

Imports System.IO
Imports System.Windows.Forms
Imports System.Collections.Generic
Public Class mwNSPECT
    Implements MapWindow.Interfaces.IPlugin

#Region "Private Variables"
    'Used for removing items on terminate
    Private _addedButtons As New System.Collections.Stack()
    Private _addedMenus As New System.Collections.Stack()

#End Region

#Region "Unused Plug-in Interface Elements"

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Handle"></param>
    ''' <param name="Location"></param>
    ''' <param name="Handled"></param>
    ''' <remarks></remarks>
    <CLSCompliant(False)> _
    Public Sub LegendDoubleClick(ByVal Handle As Integer, ByVal Location As MapWindow.Interfaces.ClickLocation, ByRef Handled As Boolean) Implements MapWindow.Interfaces.IPlugin.LegendDoubleClick

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Handle"></param>
    ''' <param name="Button"></param>
    ''' <param name="Location"></param>
    ''' <param name="Handled"></param>
    ''' <remarks></remarks>
    <CLSCompliant(False)> _
    Public Sub LegendMouseDown(ByVal Handle As Integer, ByVal Button As Integer, ByVal Location As MapWindow.Interfaces.ClickLocation, ByRef Handled As Boolean) Implements MapWindow.Interfaces.IPlugin.LegendMouseDown

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Handle"></param>
    ''' <param name="Button"></param>
    ''' <param name="Location"></param>
    ''' <param name="Handled"></param>
    ''' <remarks></remarks>
    <CLSCompliant(False)> _
    Public Sub LegendMouseUp(ByVal Handle As Integer, ByVal Button As Integer, ByVal Location As MapWindow.Interfaces.ClickLocation, ByRef Handled As Boolean) Implements MapWindow.Interfaces.IPlugin.LegendMouseUp

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Bounds"></param>
    ''' <param name="Handled"></param>
    ''' <remarks></remarks>
    Public Sub MapDragFinished(ByVal Bounds As System.Drawing.Rectangle, ByRef Handled As Boolean) Implements MapWindow.Interfaces.IPlugin.MapDragFinished

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Button"></param>
    ''' <param name="Shift"></param>
    ''' <param name="x"></param>
    ''' <param name="y"></param>
    ''' <param name="Handled"></param>
    ''' <remarks></remarks>
    Public Sub MapMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer, ByRef Handled As Boolean) Implements MapWindow.Interfaces.IPlugin.MapMouseUp

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="msg"></param>
    ''' <param name="Handled"></param>
    ''' <remarks></remarks>
    Public Sub Message(ByVal msg As String, ByRef Handled As Boolean) Implements MapWindow.Interfaces.IPlugin.Message

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="ProjectFile"></param>
    ''' <param name="SettingsString"></param>
    ''' <remarks></remarks>
    Public Sub ProjectSaving(ByVal ProjectFile As String, ByRef SettingsString As String) Implements MapWindow.Interfaces.IPlugin.ProjectSaving
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Handle"></param>
    ''' <param name="SelectInfo"></param>
    ''' <remarks></remarks>
    <CLSCompliant(False)> _
    Public Sub ShapesSelected(ByVal Handle As Integer, ByVal SelectInfo As MapWindow.Interfaces.SelectInfo) Implements MapWindow.Interfaces.IPlugin.ShapesSelected

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Handle"></param>
    ''' <remarks></remarks>
    Public Sub LayerSelected(ByVal Handle As Integer) Implements MapWindow.Interfaces.IPlugin.LayerSelected

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="ProjectFile"></param>
    ''' <param name="SettingsString"></param>
    ''' <remarks></remarks>
    Public Sub ProjectLoading(ByVal ProjectFile As String, ByVal SettingsString As String) Implements MapWindow.Interfaces.IPlugin.ProjectLoading

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Button"></param>
    ''' <param name="Shift"></param>
    ''' <param name="x"></param>
    ''' <param name="y"></param>
    ''' <param name="Handled"></param>
    ''' <remarks></remarks>
    Public Sub MapMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer, ByRef Handled As Boolean) Implements MapWindow.Interfaces.IPlugin.MapMouseDown

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="ScreenX"></param>
    ''' <param name="ScreenY"></param>
    ''' <param name="Handled"></param>
    ''' <remarks></remarks>
    Public Sub MapMouseMove(ByVal ScreenX As Integer, ByVal ScreenY As Integer, ByRef Handled As Boolean) Implements MapWindow.Interfaces.IPlugin.MapMouseMove

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Handle"></param>
    ''' <remarks></remarks>
    Public Sub LayerRemoved(ByVal Handle As Integer) Implements MapWindow.Interfaces.IPlugin.LayerRemoved

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Layers"></param>
    ''' <remarks></remarks>
    <CLSCompliant(False)> _
    Public Sub LayersAdded(ByVal Layers() As MapWindow.Interfaces.Layer) Implements MapWindow.Interfaces.IPlugin.LayersAdded

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub LayersCleared() Implements MapWindow.Interfaces.IPlugin.LayersCleared

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub MapExtentsChanged() Implements MapWindow.Interfaces.IPlugin.MapExtentsChanged


    End Sub
#End Region

#Region "Plug-in Information"
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Name() As String Implements MapWindow.Interfaces.IPlugin.Name
        Get
            Return "OpenNSPECT"
        End Get
    End Property

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Description() As String Implements MapWindow.Interfaces.IPlugin.Description
        Get
            Return "This plugin implements the Nonpoint-Source Pollution and Erosion Comparison Tool.  Primary developer: Allen Anselmo"
        End Get
    End Property

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Author() As String Implements MapWindow.Interfaces.IPlugin.Author
        Get
            Return "MapWindow Development Team"
        End Get
    End Property

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Version() As String Implements MapWindow.Interfaces.IPlugin.Version
        Get
            Return System.Diagnostics.FileVersionInfo.GetVersionInfo(Me.GetType.Assembly.Location).FileVersion
        End Get
    End Property

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property BuildDate() As String Implements MapWindow.Interfaces.IPlugin.BuildDate
        Get
            Return System.IO.File.GetLastWriteTime(Me.GetType.Assembly.Location)
        End Get
    End Property

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property SerialNumber() As String Implements MapWindow.Interfaces.IPlugin.SerialNumber
        Get
            Return "O+CW7YC/96VDT+A"
            ' This is matched to the author name in a serial number issuing program Dan ran.
        End Get
    End Property

#End Region

#Region "Start and Stop Methods"

    ''' <summary>
    ''' Event triggered on execution of the plugin
    ''' </summary>
    ''' <param name="MapWin"></param>
    ''' <param name="ParentHandle"></param>
    ''' <remarks></remarks>
    <CLSCompliant(False)> _
    Public Sub Initialize(ByVal MapWin As MapWindow.Interfaces.IMapWin, ByVal ParentHandle As Integer) Implements MapWindow.Interfaces.IPlugin.Initialize
        g_MapWin = MapWin  '  This sets global for use elsewhere in program
        g_handle = ParentHandle
        g_StatusBar = g_MapWin.StatusBar.AddPanel("", 2, 10, Windows.Forms.StatusBarPanelAutoSize.Spring)

        addMenus()
        ' addToolbars()


        Dim nspectPath As String

        ' Detects and sets the path to N-SPECT's application folder (installation directory)
        'nspectPath = My.Application.Info.DirectoryPath
        nspectPath = "C:\NSPECT\"

        If Right(nspectPath, 1) = "\" Then
            nspectPath = Left(nspectPath, Len(nspectPath) - 1)
        End If
        If Right(nspectPath, 4) = "\bin" Then
            nspectPath = Left(nspectPath, Len(nspectPath) - 4)
        End If

        modUtil.g_nspectPath = nspectPath

        modUtil.g_nspectDocPath = nspectPath

        'Initialize the database connection
        modUtil.DBConnection()

    End Sub

    ''' <summary>
    ''' Event triggered on termination of the plugin
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Terminate() Implements MapWindow.Interfaces.IPlugin.Terminate
        g_MapWin.StatusBar.RemovePanel(g_StatusBar)

        Dim t As MapWindow.Interfaces.Toolbar = g_MapWin.Toolbar
        While (_addedButtons.Count > 0)
            Try
                t.RemoveButton(_addedButtons.Pop().ToString())
            Catch ex As Exception

            End Try
        End While
        t.RemoveToolbar(g_ToolBarName)

        While (_addedMenus.Count > 0)
            Try
                g_MapWin.Menus.Remove(_addedMenus.Pop().ToString())
            Catch ex As Exception

            End Try
        End While
    End Sub
#End Region

#Region "Used Interface Methods"

    ''' <summary>
    ''' Triggered when a menu or toolbar item is clicked
    ''' </summary>
    ''' <param name="ItemName"></param>
    ''' <param name="Handled"></param>
    ''' <remarks></remarks>
    Public Sub ItemClicked(ByVal ItemName As String, ByRef Handled As Boolean) Implements MapWindow.Interfaces.IPlugin.ItemClicked
        Select Case ItemName
            Case g_mnuNSPECTAnalysis
                ShowAnalysisForm()
            Case g_mnuNSPECTCompare
                ShowCompareOUtputsForm()
            Case g_mnuNSPECTAdvLand
                ShowAdvLandForm()
            Case g_mnuNSPECTAdvPolutants
                ShowAdvPollutantsForm()
            Case g_mnuNSPECTAdvWQ
                ShowAdvWQForm()
            Case g_mnuNSPECTAdvPrecip
                ShowAdvPrecipForm()
            Case g_mnuNSPECTAdvWSDelin
                ShowAdvWSDelinForm()
            Case g_mnuNSPECTAdvSoils
                ShowAdvSoilsForm()
            Case g_mnuNSPECTHelp
                ShowHelpIntro()
        End Select
    End Sub

#End Region

#Region "Helper Functions"

#Region "   Menu/Toolbar Items"
    ''' <summary>
    ''' Sub used to add all the menus used by the mwNSPECT plugin
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub addMenus()
        Dim nil As Object
        nil = Nothing
        With g_MapWin.Menus
            .AddMenu(g_mnuNSPECTMain, nil, "NSPECT")
            _addedMenus.Push(g_mnuNSPECTMain)
            .AddMenu(g_mnuNSPECTAnalysis, g_mnuNSPECTMain, nil, "Run Analysis")
            _addedMenus.Push(g_mnuNSPECTAnalysis)
            .AddMenu("mnunspectsep0", g_mnuNSPECTMain, nil, "-")
            _addedMenus.Push("mnunspectsep0")
            .AddMenu(g_mnuNSPECTCompare, g_mnuNSPECTMain, nil, "Compare Outputs")
            _addedMenus.Push(g_mnuNSPECTCompare)
            .AddMenu("mnunspectsep1", g_mnuNSPECTMain, nil, "-")
            _addedMenus.Push("mnunspectsep1")
            .AddMenu(g_mnuNSPECTAdvSettings, g_mnuNSPECTMain, nil, "Advanced Settings")
            _addedMenus.Push(g_mnuNSPECTAdvSettings)
            .AddMenu(g_mnuNSPECTAdvLand, g_mnuNSPECTAdvSettings, nil, "Land Cover Types")
            _addedMenus.Push(g_mnuNSPECTAdvLand)
            .AddMenu(g_mnuNSPECTAdvPolutants, g_mnuNSPECTAdvSettings, nil, "Pollutants")
            _addedMenus.Push(g_mnuNSPECTAdvPolutants)
            .AddMenu(g_mnuNSPECTAdvWQ, g_mnuNSPECTAdvSettings, nil, "Water Quality Standards")
            _addedMenus.Push(g_mnuNSPECTAdvWQ)
            .AddMenu(g_mnuNSPECTAdvPrecip, g_mnuNSPECTAdvSettings, nil, "Precipitation Scenarios")
            _addedMenus.Push(g_mnuNSPECTAdvPrecip)
            .AddMenu(g_mnuNSPECTAdvWSDelin, g_mnuNSPECTAdvSettings, nil, "Watershed Delineations")
            _addedMenus.Push(g_mnuNSPECTAdvWSDelin)
            .AddMenu(g_mnuNSPECTAdvSoils, g_mnuNSPECTAdvSettings, nil, "Soils")
            _addedMenus.Push(g_mnuNSPECTAdvSoils)
            .AddMenu("mnunspectsep2", g_mnuNSPECTMain, nil, "-")
            _addedMenus.Push("mnunspectsep2")
            .AddMenu(g_mnuNSPECTHelp, g_mnuNSPECTMain, nil, "Help")
            _addedMenus.Push(g_mnuNSPECTHelp)

        End With
    End Sub

    ''' <summary>
    ''' Sub used to add all toolbars and buttons used by the mwNSPECT plugin
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub addToolbars()
        'g_MapWin.Toolbar.AddToolbar(g_ToolBarName)

        'Dim tlbrbtnfreehandCapture As MapWindow.Interfaces.ToolbarButton = g_MapWin.Toolbar.AddButton(g_btnFreehand, g_ToolBarName, "", "")
        'Dim freehand_ico As New Drawing.Icon(Me.GetType, "pencilFreehand.ico")
        'tlbrbtnfreehandCapture.Picture = freehand_ico
        'tlbrbtnfreehandCapture.Tooltip = "Manually Add a Feature"
        '_addedButtons.Push(g_btnFreehand)


        'Dim tlbrbtntmp As MapWindow.Interfaces.ToolbarButton = g_MapWin.Toolbar.AddButton("tmp", g_ToolBarName, "", "")
        '_addedButtons.Push("tmp")
    End Sub
#End Region

#Region "   Itemclicked Items"
    ''' <summary>
    ''' Shows the main Analysis form
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ShowAnalysisForm()
        g_frmProjectSetup = New frmProjectSetup
        g_frmProjectSetup.ShowDialog()
    End Sub

    ''' <summary>
    ''' Shows the outputs comparison form
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared Sub ShowCompareOutputsForm()
        Using tmp As New frmCompareOutputs()
            tmp.ShowDialog()
        End Using
    End Sub

    ''' <summary>
    ''' Shows the Land Cover form
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared Sub ShowAdvLandForm()
        Using tmp As New frmLandCoverTypes()
            tmp.ShowDialog()
        End Using
    End Sub

    ''' <summary>
    ''' Shows the Pollutant form
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared Sub ShowAdvPollutantsForm()
        Using tmp As New frmPollutants()
            tmp.ShowDialog()
        End Using
    End Sub

    ''' <summary>
    ''' Shows the Water Quality form
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared Sub ShowAdvWQForm()
        Using tmp As New frmWaterQualityStandard()
            tmp.ShowDialog()
        End Using
    End Sub

    ''' <summary>
    ''' Shows the Precipitation form
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared Sub ShowAdvPrecipForm()
        Using tmp As New frmPrecipitation()
            tmp.ShowDialog()
        End Using
    End Sub

    ''' <summary>
    ''' Shows the Watershed Delin form
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared Sub ShowAdvWSDelinForm()
        Using tmp As New frmWatershedDelin()
            tmp.ShowDialog()
        End Using
    End Sub

    ''' <summary>
    ''' Shows the Soils form
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared Sub ShowAdvSoilsForm()
        Using tmp As New frmSoils()
            tmp.ShowDialog()
        End Using
    End Sub

    ''' <summary>
    ''' Shows the Help form
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ShowHelpIntro()
        System.Windows.Forms.Help.ShowHelp(Nothing, modUtil.g_nspectPath & "\Help\nspect.chm", "Introduction.htm")
    End Sub

#End Region

#End Region

End Class

