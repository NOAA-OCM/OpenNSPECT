Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsHelp_NET.clsHelp")> Public Class clsHelp
	Implements ESRI.ArcGIS.SystemUI.ICommand
	' *************************************************************************************
	' *  Perot Systems Government Services
	' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
	' *  clsHelp
	' *************************************************************************************
	' *  Description:  Class for Help menu option off NSPECT main menu
	' *
	' *
	' *  Called By:  Help... menu
	' *************************************************************************************
	Private m_App As ESRI.ArcGIS.Framework.IApplication
	
	'I threw the Ado.Close event here.  Keeps the Access dbase from locking.
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		
		On Error Resume Next
		If Not (g_ADOConn Is Nothing) Then
			
			g_ADOConn.Close()
			
		End If
		
		'UPGRADE_NOTE: Object g_ADOConn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		g_ADOConn = Nothing
		'UPGRADE_NOTE: Object m_App may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_App = Nothing
		
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	Private ReadOnly Property ICommand_Bitmap() As ESRI.ArcGIS.esriSystem.OLE_HANDLE Implements ESRI.ArcGIS.SystemUI.ICommand.Bitmap
		Get
			
		End Get
	End Property
	
	Private ReadOnly Property ICommand_Caption() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Caption
		Get
			ICommand_Caption = "Help..."
		End Get
	End Property
	
	Private ReadOnly Property ICommand_Category() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Category
		Get
			ICommand_Category = "N-SPECT"
		End Get
	End Property
	
	Private ReadOnly Property ICommand_Checked() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Checked
		Get
			
		End Get
	End Property
	
	Private ReadOnly Property ICommand_Enabled() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Enabled
		Get
			ICommand_Enabled = True
		End Get
	End Property
	
	Private ReadOnly Property ICommand_HelpContextID() As Integer Implements ESRI.ArcGIS.SystemUI.ICommand.HelpContextID
		Get
			
		End Get
	End Property
	
	Private ReadOnly Property ICommand_HelpFile() As String Implements ESRI.ArcGIS.SystemUI.ICommand.HelpFile
		Get
			
		End Get
	End Property
	
	Private ReadOnly Property ICommand_Message() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Message
		Get
			ICommand_Message = "Help"
		End Get
	End Property
	
	Private ReadOnly Property ICommand_Name() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Name
		Get
			ICommand_Name = "N-SPECTToolHelp"
		End Get
	End Property
	
	Private ReadOnly Property ICommand_Tooltip() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Tooltip
		Get
			ICommand_Tooltip = "Choose for help on N-SPECT"
		End Get
	End Property
	
	Private Sub ICommand_OnClick() Implements ESRI.ArcGIS.SystemUI.ICommand.OnClick
		On Error GoTo ErrHandler
		'Load the help file
		HtmlHelp(0, modUtil.g_nspectPath & "\Help\nspect.chm", HH_DISPLAY_TOPIC, "Introduction.htm")
		Exit Sub
		
ErrHandler: 
		MsgBox("Could not find NSPECT help.  Please check Help directory for NSPECT.chm.", MsgBoxStyle.Critical, "Help Not Found")
	End Sub
	
	Private Sub ICommand_OnCreate(ByVal hook As Object) Implements ESRI.ArcGIS.SystemUI.ICommand.OnCreate
		m_App = hook
	End Sub
End Class