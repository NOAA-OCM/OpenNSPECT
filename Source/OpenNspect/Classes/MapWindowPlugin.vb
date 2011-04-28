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
    Private Shared _mapwindowInstance As IMapWin

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
    Private _MenuItems As New Stack()

#End Region

    Public Shared ReadOnly Property MapWindowInstance() As IMapWin
        Get
            Return _mapwindowInstance
        End Get
    End Property
    Private ReadOnly Property ReferenceAssembly() As [Assembly]
        Get
            Return Me.GetType.Assembly
        End Get
    End Property

    Private _file As FileVersionInfo
    Private ReadOnly Property ReferenceFile() As FileVersionInfo
        Get
            If _file Is Nothing Then
                _file = FileVersionInfo.GetVersionInfo(ReferenceAssembly.Location)
            End If

            Return _file
        End Get
    End Property
#Region "Unused Plug-in Interface Elements"

    <CLSCompliant(False)> _
    Public Sub LegendDoubleClick(ByVal Handle As Integer, ByVal Location As ClickLocation, _
                                  ByRef Handled As Boolean) Implements IPlugin.LegendDoubleClick

    End Sub

    <CLSCompliant(False)> _
    Public Sub LegendMouseDown(ByVal Handle As Integer, ByVal Button As Integer, _
                                ByVal Location As ClickLocation, ByRef Handled As Boolean) _
        Implements IPlugin.LegendMouseDown

    End Sub

    <CLSCompliant(False)> _
    Public Sub LegendMouseUp(ByVal Handle As Integer, ByVal Button As Integer, _
                              ByVal Location As ClickLocation, ByRef Handled As Boolean) _
        Implements IPlugin.LegendMouseUp

    End Sub

    Public Sub MapDragFinished(ByVal Bounds As Rectangle, ByRef Handled As Boolean) _
        Implements IPlugin.MapDragFinished

    End Sub

    Public Sub MapMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer, _
                           ByRef Handled As Boolean) Implements IPlugin.MapMouseUp

    End Sub

    Public Sub Message(ByVal msg As String, ByRef Handled As Boolean) Implements IPlugin.Message

    End Sub

    Public Sub ProjectSaving(ByVal ProjectFile As String, ByRef SettingsString As String) _
        Implements IPlugin.ProjectSaving
    End Sub

    <CLSCompliant(False)> _
    Public Sub ShapesSelected(ByVal Handle As Integer, ByVal SelectInfo As SelectInfo) _
        Implements IPlugin.ShapesSelected

    End Sub

    Public Sub LayerSelected(ByVal Handle As Integer) Implements IPlugin.LayerSelected

    End Sub

    Public Sub ProjectLoading(ByVal ProjectFile As String, ByVal SettingsString As String) _
        Implements IPlugin.ProjectLoading

    End Sub

    Public Sub MapMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer, _
                             ByRef Handled As Boolean) Implements IPlugin.MapMouseDown

    End Sub

    Public Sub MapMouseMove(ByVal ScreenX As Integer, ByVal ScreenY As Integer, ByRef Handled As Boolean) _
        Implements IPlugin.MapMouseMove

    End Sub

    Public Sub LayerRemoved(ByVal Handle As Integer) Implements IPlugin.LayerRemoved

    End Sub

    <CLSCompliant(False)> _
    Public Sub LayersAdded(ByVal Layers() As Layer) _
        Implements IPlugin.LayersAdded

    End Sub

    Public Sub LayersCleared() Implements IPlugin.LayersCleared

    End Sub

    Public Sub MapExtentsChanged() Implements IPlugin.MapExtentsChanged

    End Sub

#End Region

#Region "Plug-in Information"

    Public ReadOnly Property Name() As String Implements IPlugin.Name
        Get
            Return ReferenceFile.ProductName
        End Get
    End Property

    Public ReadOnly Property Description() As String Implements IPlugin.Description
        Get
            Return ReferenceFile.Comments
        End Get
    End Property

    Public ReadOnly Property Author() As String Implements IPlugin.Author
        Get
            Return ReferenceFile.CompanyName
        End Get
    End Property

    Public ReadOnly Property Version() As String Implements IPlugin.Version
        Get
            Return ReferenceFile.FileVersion
        End Get
    End Property

    Public ReadOnly Property BuildDate() As String Implements IPlugin.BuildDate
        Get
            Return File.GetLastWriteTime(ReferenceAssembly.Location)
        End Get
    End Property

    Public ReadOnly Property SerialNumber() As String Implements IPlugin.SerialNumber
        Get
            Return String.Empty
        End Get
    End Property

#End Region

    Private Sub VerifyMapWindowVersion()
        Dim requiredVersion = New System.Version(4, 8, 2, 21243)
        Dim ver = Assembly.GetEntryAssembly().GetName().Version
        If ver < requiredVersion Then
            MessageBox.Show(String.Format("Your copy of MapWindow {0} is older than the required version {1} to run {2}. Visit http://www.mapwindow.org to update your copy.", ver, requiredVersion, Name))
        End If
    End Sub
#Region "Start and Stop Methods"

    <CLSCompliant(False)> _
    Public Sub Initialize(ByVal MapWin As IMapWin, ByVal ParentHandle As Integer) _
        Implements IPlugin.Initialize

        If MapWin Is Nothing Then
            Throw New ArgumentNullException("MapWin", "MapWin is nothing.")
        End If

        _mapwindowInstance = MapWin

        VerifyMapWindowVersion()

        AddMenus()

        Dim nspectPath As String = "C:\NSPECT\"
        ' Detects and sets the path to OpenNSPECT's application folder (installation directory)
        'nspectPath = My.Application.Info.DirectoryPath

        If Not Directory.Exists(nspectPath) Then
            MessageBox.Show("{0} is not a valid directory. You need to reinstall the plugin or relocate your data files to this location.", nspectPath)
        End If

        nspectPath = nspectPath.TrimEnd("\")

        g_nspectPath = nspectPath
        g_nspectDocPath = nspectPath

        'Initialize the database connection
        InitializeDBConnection()

    End Sub

    Public Sub Terminate() Implements IPlugin.Terminate
        For Each item As String In _MenuItems
            MapWindowInstance.Menus.Remove(item)
        Next
    End Sub

#End Region

#Region "Used Interface Methods"

    Public Sub ItemClicked(ByVal ItemName As String, ByRef Handled As Boolean) _
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

    Private Sub AddMenus()
        With MapWindowInstance.Menus
            .AddMenu(mnuNspectMain, Nothing, Nothing, "OpenNSPECT")
            _MenuItems.Push(mnuNspectMain)
            .AddMenu(mnuNspectAnalysis, mnuNspectMain, Nothing, "Run Analysis...")
            _MenuItems.Push(mnuNspectAnalysis)
            .AddMenu("mnunspectsep0", mnuNspectMain, Nothing, "-")
            _MenuItems.Push("mnunspectsep0")
            .AddMenu(mnuNspectCompare, mnuNspectMain, Nothing, "Compare Outputs...")
            _MenuItems.Push(mnuNspectCompare)
            .AddMenu("mnunspectsep1", mnuNspectMain, Nothing, "-")
            _MenuItems.Push("mnunspectsep1")

            .AddMenu(mnuNspectAdvSettings, mnuNspectMain, Nothing, "Advanced Settings")
            _MenuItems.Push(mnuNspectAdvSettings)
            .AddMenu(mnuNspectAdvLand, mnuNspectAdvSettings, Nothing, "Land Cover Types...")
            _MenuItems.Push(mnuNspectAdvLand)
            .AddMenu(mnuNspectAdvPolutants, mnuNspectAdvSettings, Nothing, "Pollutants...")
            _MenuItems.Push(mnuNspectAdvPolutants)
            .AddMenu(mnuNspectAdvWQ, mnuNspectAdvSettings, Nothing, "Water Quality Standards...")
            _MenuItems.Push(mnuNspectAdvWQ)
            .AddMenu(mnuNspectAdvPrecip, mnuNspectAdvSettings, Nothing, "Precipitation Scenarios...")
            _MenuItems.Push(mnuNspectAdvPrecip)
            .AddMenu(mnuNspectAdvWSDelin, mnuNspectAdvSettings, Nothing, "Watershed Delineations...")
            _MenuItems.Push(mnuNspectAdvWSDelin)

            .AddMenu(mnuNspectAdvSoils, mnuNspectAdvSettings, Nothing, "Soils...")
            _MenuItems.Push(mnuNspectAdvSoils)
            .AddMenu("mnunspectsep2", mnuNspectMain, Nothing, "-")
            _MenuItems.Push("mnunspectsep2")
            .AddMenu(mnuNspectHelp, mnuNspectMain, Nothing, "Help...")
            _MenuItems.Push(mnuNspectHelp)
            .AddMenu(mnuNspectAbout, mnuNspectMain, Nothing, "About...")
            _MenuItems.Push(mnuNspectAbout)
        End With
    End Sub

#Region "   Itemclicked Items"

    Private Sub ShowAnalysisForm()
        g_frmProjectSetup = New MainForm
        g_frmProjectSetup.ShowDialog()
    End Sub

    Private Shared Sub ShowCompareOutputsForm()
        Using tmp As New CompareOutputsForm()
            tmp.ShowDialog()
        End Using
    End Sub

    Private Shared Sub ShowAdvLandForm()
        Using tmp As New LandCoverTypesForm()
            tmp.ShowDialog()
        End Using
    End Sub

    Private Shared Sub ShowAdvPollutantsForm()
        Using tmp As New PollutantsForm()
            tmp.ShowDialog()
        End Using
    End Sub

    Private Shared Sub ShowAdvWQForm()
        Using tmp As New WaterQualityStandardsForm()
            tmp.ShowDialog()
        End Using
    End Sub

    Private Shared Sub ShowAdvPrecipForm()
        Using tmp As New PrecipitationScenariosForm()
            tmp.ShowDialog()
        End Using
    End Sub

    Private Shared Sub ShowAdvWSDelinForm()
        Using tmp As New WatershedDelineationsForm()
            tmp.ShowDialog()
        End Using
    End Sub

    Private Shared Sub ShowAdvSoilsForm()
        Using tmp As New SoilsForm()
            tmp.ShowDialog()
        End Using
    End Sub

    Private Sub ShowHelpIntro()
        Help.ShowHelp(Nothing, g_nspectPath & "\Help\nspect.chm", "Introduction.htm")
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
