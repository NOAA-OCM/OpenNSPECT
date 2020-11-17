Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsLCType_NET.clsLCType")> Public Class clsLCType
	Implements ESRI.ArcGIS.SystemUI.ICommand
	' *************************************************************************************
	' *  Perot Systems Government Services
	' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
	' *  clsLCType
	' *************************************************************************************
	' *  Description:  Class for LandCover Types menu option off NSPECT advanced menu
	' *
	' *
	' *  Called By:  Land Cover Types... menu
	' *************************************************************************************
	
	' Define a class module
	Private m_App As ESRI.ArcGIS.Framework.IApplication
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
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
			ICommand_Caption = "Land Cover Types..."
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
			ICommand_Message = "Land Cover Types"
		End Get
	End Property
	
	Private ReadOnly Property ICommand_Name() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Name
		Get
			ICommand_Name = "HIToolLCType"
		End Get
	End Property
	
	Private ReadOnly Property ICommand_Tooltip() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Tooltip
		Get
			ICommand_Tooltip = "Choose Land Cover Type"
		End Get
	End Property
	
	Private Sub ICommand_OnClick() Implements ESRI.ArcGIS.SystemUI.ICommand.OnClick
		' initialize and show the form
		frmLCTypes.init(m_App)
		frmLCTypes.ShowDialog()
		
	End Sub
	
	Private Sub ICommand_OnCreate(ByVal hook As Object) Implements ESRI.ArcGIS.SystemUI.ICommand.OnCreate
		m_App = hook
	End Sub
End Class