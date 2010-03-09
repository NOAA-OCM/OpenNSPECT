'********************************************************************************************************
'File Name: Globals.vb
'Description: This screen is used to specify parameters for adding a new field.
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

Module Globals
#Region "Base Globals"
    Public g_MapWin As MapWindow.Interfaces.IMapWin
    Public g_handle As Integer
    Public g_StatusBar As System.Windows.Forms.StatusBarPanel

#End Region

#Region "Menu Globals"
    Public g_mnuNSPECTMain As String = "mnunspectMainMenu"
    Public g_mnuNSPECTAnalysis As String = "mnunspectAnalysis"
    Public g_mnuNSPECTAdvSettings As String = "mnunspectAdvancedSettings"
    Public g_mnuNSPECTAdvLand As String = "mnunspectLandCover"
    Public g_mnuNSPECTAdvPolutants As String = "mnunspectPollutants"
    Public g_mnuNSPECTAdvWQ As String = "mnunspectWaterQuality"
    Public g_mnuNSPECTAdvPrecip As String = "mnunspectPrecipitation"
    Public g_mnuNSPECTAdvWSDelin As String = "mnunspectWatershedDelineations"
    Public g_mnuNSPECTAdvSoils As String = "mnunspectSoils"
    Public g_mnuNSPECTHelp As String = "mnunspectHelp"

#End Region

#Region "Toolbar Globals"
    Public g_ToolBarName As String = "nspectToolbar"

#End Region


End Module
