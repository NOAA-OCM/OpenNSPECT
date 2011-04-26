'********************************************************************************************************
'File Name: MapWindowPlugin.vb
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
Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports MapWindow.Interfaces
Imports System.Reflection

Public Class MapWindowPlugin
    Implements IPlugin

    Private Const mnuNspectMain As String = "mnunspectMainMenu"
    Private Const mnuNspectAnalysis As String = "mnunspectAnalysis"
    Private Const mnuNspectCompare As String = "mnunspectCompare"
    Private Const mnuNspectAdvSettings As String = "mnunspectAdvancedSettings"
    Private Const mnuNspectAdvLand As String = "mnunspectLandCover"
    Private Const mnuNspectAdvPolutants As String = "mnunspectPollutants"
    Private Const mnuNspectAdvWQ As String = "mnunspectWaterQuality"
    Private Const mnuNspectAdvPrecip As String = "mnunspectPrecipitation"
    Private Const mnuNspectAdvWSDelin As String = "mnunspectWatershedDelineations"
    Private Const mnuNspectAdvSoils As String = "mnunspectSoils"
    Private Const mnuNspectHelp As String = "mnunspectHelp"
    Private Const mnuNspectAbout As String = "mnunspectAbout"

#Region "Private Variables"

    'Used for removing items on terminate
    Private _addedMenus As New Stack()

#End Region

#Region "Unused Plug-in Interface Elements"

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Handle"></param>
    ''' <param name="Location"></param>
    ''' <param name="Handled"></param>
    ''' <remarks></remarks>
    <CLSCompliant (False)> _
    Public Sub LegendDoubleClick (ByVal Handle As Integer, ByVal Location As ClickLocation, _
                                  ByRef Handled As Boolean) Implements IPlugin.LegendDoubleClick

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Handle"></param>
    ''' <param name="Button"></param>
    ''' <param name="Location"></param>
    ''' <param name="Handled"></param>
    ''' <remarks></remarks>
    <CLSCompliant (False)> _
    Public Sub LegendMouseDown (ByVal Handle As Integer, ByVal Button As Integer, _
                                ByVal Location As ClickLocation, ByRef Handled As Boolean) _
        Implements IPlugin.LegendMouseDown

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Handle"></param>
    ''' <param name="Button"></param>
    ''' <param name="Location"></param>
    ''' <param name="Handled"></param>
    ''' <remarks></remarks>
    <CLSCompliant (False)> _
    Public Sub LegendMouseUp (ByVal Handle As Integer, ByVal Button As Integer, _
                              ByVal Location As ClickLocation, ByRef Handled As Boolean) _
        Implements IPlugin.LegendMouseUp

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Bounds"></param>
    ''' <param name="Handled"></param>
    ''' <remarks></remarks>
    Public Sub MapDragFinished (ByVal Bounds As Rectangle, ByRef Handled As Boolean) _
        Implements IPlugin.MapDragFinished

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
    Public Sub MapMouseUp (ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer, _
                           ByRef Handled As Boolean) Implements IPlugin.MapMouseUp

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="msg"></param>
    ''' <param name="Handled"></param>
    ''' <remarks></remarks>
    Public Sub Message (ByVal msg As String, ByRef Handled As Boolean) Implements IPlugin.Message

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="ProjectFile"></param>
    ''' <param name="SettingsString"></param>
    ''' <remarks></remarks>
    Public Sub ProjectSaving (ByVal ProjectFile As String, ByRef SettingsString As String) _
        Implements IPlugin.ProjectSaving
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Handle"></param>
    ''' <param name="SelectInfo"></param>
    ''' <remarks></remarks>
    <CLSCompliant (False)> _
    Public Sub ShapesSelected (ByVal Handle As Integer, ByVal SelectInfo As SelectInfo) _
        Implements IPlugin.ShapesSelected

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Handle"></param>
    ''' <remarks></remarks>
    Public Sub LayerSelected (ByVal Handle As Integer) Implements IPlugin.LayerSelected

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="ProjectFile"></param>
    ''' <param name="SettingsString"></param>
    ''' <remarks></remarks>
    Public Sub ProjectLoading (ByVal ProjectFile As String, ByVal SettingsString As String) _
        Implements IPlugin.ProjectLoading

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
    Public Sub MapMouseDown (ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer, _
                             ByRef Handled As Boolean) Implements IPlugin.MapMouseDown

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="ScreenX"></param>
    ''' <param name="ScreenY"></param>
    ''' <param name="Handled"></param>
    ''' <remarks></remarks>
    Public Sub MapMouseMove (ByVal ScreenX As Integer, ByVal ScreenY As Integer, ByRef Handled As Boolean) _
        Implements IPlugin.MapMouseMove

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Handle"></param>
    ''' <remarks></remarks>
    Public Sub LayerRemoved (ByVal Handle As Integer) Implements IPlugin.LayerRemoved

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Layers"></param>
    ''' <remarks></remarks>
    <CLSCompliant (False)> _
    Public Sub LayersAdded (ByVal Layers() As Layer) _
        Implements IPlugin.LayersAdded

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub LayersCleared() Implements IPlugin.LayersCleared

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub MapExtentsChanged() Implements IPlugin.MapExtentsChanged

    End Sub

#End Region

#Region "Plug-in Information"

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Name() As String Implements IPlugin.Name
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
    Public ReadOnly Property Description() As String Implements IPlugin.Description
        Get
            Return _
                "This plugin implements the Nonpoint-Source Pollution and Erosion Comparison Tool.  Primary developer: Allen Anselmo"
        End Get
    End Property

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Author() As String Implements IPlugin.Author
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
    Public ReadOnly Property Version() As String Implements IPlugin.Version
        Get
            Return FileVersionInfo.GetVersionInfo (Me.GetType.Assembly.Location).FileVersion
        End Get
    End Property

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property BuildDate() As String Implements IPlugin.BuildDate
        Get
            Return File.GetLastWriteTime (Me.GetType.Assembly.Location)
        End Get
    End Property

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property SerialNumber() As String Implements IPlugin.SerialNumber
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
    <CLSCompliant (False)> _
    Public Sub Initialize (ByVal MapWin As IMapWin, ByVal ParentHandle As Integer) _
        Implements IPlugin.Initialize
        g_MapWin = MapWin
        '  This sets global for use elsewhere in program
        g_handle = ParentHandle
        g_StatusBar = g_MapWin.StatusBar.AddPanel ("", 2, 10, StatusBarPanelAutoSize.Spring)

        AddMenus()

        Dim nspectPath As String = "C:\NSPECT\"

        ' Detects and sets the path to OpenNSPECT's application folder (installation directory)
        'nspectPath = My.Application.Info.DirectoryPath

        If Right (nspectPath, 1) = "\" Then
            nspectPath = Left (nspectPath, Len (nspectPath) - 1)
        End If
        If Right (nspectPath, 4) = "\bin" Then
            nspectPath = Left (nspectPath, Len (nspectPath) - 4)
        End If

        g_nspectPath = nspectPath
        g_nspectDocPath = nspectPath

        'Initialize the database connection
        DBConnection()

    End Sub

    ''' <summary>
    ''' Event triggered on termination of the plugin
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Terminate() Implements IPlugin.Terminate
        g_MapWin.StatusBar.RemovePanel (g_StatusBar)

        While (_addedMenus.Count > 0)
            Try
                g_MapWin.Menus.Remove (_addedMenus.Pop().ToString())
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
    Public Sub ItemClicked (ByVal ItemName As String, ByRef Handled As Boolean) _
        Implements IPlugin.ItemClicked
        Select Case ItemName
            Case mnuNspectAnalysis
                ShowAnalysisForm()
            Case mnuNspectCompare
                ShowCompareOutputsForm()
            Case mnuNspectAdvLand
                ShowAdvLandForm()
            Case mnuNspectAdvPolutants
                ShowAdvPollutantsForm()
            Case mnuNspectAdvWQ
                ShowAdvWQForm()
            Case mnuNspectAdvPrecip
                ShowAdvPrecipForm()
            Case mnuNspectAdvWSDelin
                ShowAdvWSDelinForm()
            Case mnuNspectAdvSoils
                ShowAdvSoilsForm()
            Case mnuNspectHelp
                ShowHelpIntro()
            Case mnuNspectAbout
                ShowAboutForm()
        End Select
    End Sub

#End Region

#Region "   Menu/Toolbar Items"

    ''' <summary>
    ''' Add all the menus used by the plugin
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub AddMenus()
        Dim nil As Object = Nothing
        With g_MapWin.Menus
            .AddMenu (mnuNspectMain, nil, "OpenNSPECT")
            _addedMenus.Push (mnuNspectMain)
            .AddMenu (mnuNspectAnalysis, mnuNspectMain, nil, "Run Analysis...")
            _addedMenus.Push (mnuNspectAnalysis)
            .AddMenu ("mnunspectsep0", mnuNspectMain, nil, "-")
            _addedMenus.Push ("mnunspectsep0")
            .AddMenu (mnuNspectCompare, mnuNspectMain, nil, "Compare Outputs...")
            _addedMenus.Push (mnuNspectCompare)
            .AddMenu ("mnunspectsep1", mnuNspectMain, nil, "-")
            _addedMenus.Push ("mnunspectsep1")
            .AddMenu (mnuNspectAdvSettings, mnuNspectMain, nil, "Advanced Settings")
            _addedMenus.Push (mnuNspectAdvSettings)
            .AddMenu (mnuNspectAdvLand, mnuNspectAdvSettings, nil, "Land Cover Types...")
            _addedMenus.Push (mnuNspectAdvLand)
            .AddMenu (mnuNspectAdvPolutants, mnuNspectAdvSettings, nil, "Pollutants...")
            _addedMenus.Push (mnuNspectAdvPolutants)
            .AddMenu (mnuNspectAdvWQ, mnuNspectAdvSettings, nil, "Water Quality Standards...")
            _addedMenus.Push (mnuNspectAdvWQ)
            .AddMenu (mnuNspectAdvPrecip, mnuNspectAdvSettings, nil, "Precipitation Scenarios...")
            _addedMenus.Push (mnuNspectAdvPrecip)
            .AddMenu (mnuNspectAdvWSDelin, mnuNspectAdvSettings, nil, "Watershed Delineations...")
            _addedMenus.Push (mnuNspectAdvWSDelin)
            .AddMenu (mnuNspectAdvSoils, mnuNspectAdvSettings, nil, "Soils...")
            _addedMenus.Push (mnuNspectAdvSoils)
            .AddMenu ("mnunspectsep2", mnuNspectMain, nil, "-")
            _addedMenus.Push ("mnunspectsep2")
            .AddMenu (mnuNspectHelp, mnuNspectMain, nil, "Help...")
            _addedMenus.Push (mnuNspectHelp)
            .AddMenu (mnuNspectAbout, mnuNspectMain, nil, "About...")
            _addedMenus.Push (mnuNspectAbout)

        End With
    End Sub

#Region "   Itemclicked Items"

    ''' <summary>
    ''' Shows the main Analysis form
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ShowAnalysisForm()
        g_frmProjectSetup = New MainForm
        g_frmProjectSetup.ShowDialog()
    End Sub

    ''' <summary>
    ''' Shows the outputs comparison form
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared Sub ShowCompareOutputsForm()
        Using tmp As New CompareOutputsForm()
            tmp.ShowDialog()
        End Using
    End Sub

    ''' <summary>
    ''' Shows the Land Cover form
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared Sub ShowAdvLandForm()
        Using tmp As New LandCoverTypesForm()
            tmp.ShowDialog()
        End Using
    End Sub

    ''' <summary>
    ''' Shows the Pollutant form
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared Sub ShowAdvPollutantsForm()
        Using tmp As New PollutantsForm()
            tmp.ShowDialog()
        End Using
    End Sub

    ''' <summary>
    ''' Shows the Water Quality form
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared Sub ShowAdvWQForm()
        Using tmp As New WaterQualityStandardsForm()
            tmp.ShowDialog()
        End Using
    End Sub

    ''' <summary>
    ''' Shows the Precipitation form
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared Sub ShowAdvPrecipForm()
        Using tmp As New PrecipitationScenariosForm()
            tmp.ShowDialog()
        End Using
    End Sub

    ''' <summary>
    ''' Shows the Watershed Delin form
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared Sub ShowAdvWSDelinForm()
        Using tmp As New WatershedDelineationsForm()
            tmp.ShowDialog()
        End Using
    End Sub

    ''' <summary>
    ''' Shows the Soils form
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared Sub ShowAdvSoilsForm()
        Using tmp As New SoilsForm()
            tmp.ShowDialog()
        End Using
    End Sub

    ''' <summary>
    ''' Shows the Help form
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ShowHelpIntro()
        Help.ShowHelp (Nothing, g_nspectPath & "\Help\nspect.chm", "Introduction.htm")
    End Sub

#End Region

#End Region

    Private Shared Sub ShowAboutForm()
        Using form = New AboutForm()
            form.AppEntryAssembly = Assembly.GetExecutingAssembly
            form.ShowDialog()
        End Using
    End Sub
End Class

