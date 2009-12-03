Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsSoils_NET.clsSoils")> Public Class clsSoils
	Implements ESRI.ArcGIS.SystemUI.ICommand
	' *************************************************************************************
	' *  Perot Systems Government Services
	' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
	' *  clsSoils
	' *************************************************************************************
	' *  Description: Class that defines the main choice "Soils" in the advanced setting
	' *
	' *
	' *  Called By: Soils menu choice
	' *************************************************************************************
	' *  Subs:
	' *
	' *  Misc:
	' *************************************************************************************
	' Define a class module
	
	Private m_App As ESRI.ArcGIS.Framework.IApplication
	' Variables used by the Error handler function - DO NOT REMOVE
	Const c_sModuleFileName As String = "clsSoils.cls"
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		On Error GoTo ErrorHandler
		
		'UPGRADE_NOTE: Object m_App may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_App = Nothing
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "Class_Terminate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	Private ReadOnly Property ICommand_Bitmap() As ESRI.ArcGIS.esriSystem.OLE_HANDLE Implements ESRI.ArcGIS.SystemUI.ICommand.Bitmap
		Get
			On Error GoTo ErrorHandler
			
			Exit Property
ErrorHandler: 
			HandleError(True, "ICommand_Bitmap " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
		End Get
	End Property
	
	Private ReadOnly Property ICommand_Caption() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Caption
		Get
			On Error GoTo ErrorHandler
			
			ICommand_Caption = "Soils..."
			
			Exit Property
ErrorHandler: 
			HandleError(True, "ICommand_Caption " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
		End Get
	End Property
	
	Private ReadOnly Property ICommand_Category() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Category
		Get
			On Error GoTo ErrorHandler
			
			ICommand_Category = "N-SPECT"
			
			Exit Property
ErrorHandler: 
			HandleError(True, "ICommand_Category " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
		End Get
	End Property
	
	Private ReadOnly Property ICommand_Checked() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Checked
		Get
			On Error GoTo ErrorHandler
			
			Exit Property
ErrorHandler: 
			HandleError(True, "ICommand_Checked " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
		End Get
	End Property
	
	Private ReadOnly Property ICommand_Enabled() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Enabled
		Get
			On Error GoTo ErrorHandler
			
			ICommand_Enabled = True
			
			Exit Property
ErrorHandler: 
			HandleError(True, "ICommand_Enabled " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
		End Get
	End Property
	
	Private ReadOnly Property ICommand_HelpContextID() As Integer Implements ESRI.ArcGIS.SystemUI.ICommand.HelpContextID
		Get
			On Error GoTo ErrorHandler
			
			Exit Property
ErrorHandler: 
			HandleError(True, "ICommand_HelpContextID " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
		End Get
	End Property
	
	Private ReadOnly Property ICommand_HelpFile() As String Implements ESRI.ArcGIS.SystemUI.ICommand.HelpFile
		Get
			On Error GoTo ErrorHandler
			
			Exit Property
ErrorHandler: 
			HandleError(True, "ICommand_HelpFile " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
		End Get
	End Property
	
	Private ReadOnly Property ICommand_Message() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Message
		Get
			On Error GoTo ErrorHandler
			
			ICommand_Message = "Soils"
			
			Exit Property
ErrorHandler: 
			HandleError(True, "ICommand_Message " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
		End Get
	End Property
	
	Private ReadOnly Property ICommand_Name() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Name
		Get
			On Error GoTo ErrorHandler
			
			ICommand_Name = "HIToolSoils"
			
			Exit Property
ErrorHandler: 
			HandleError(True, "ICommand_Name " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
		End Get
	End Property
	
	Private ReadOnly Property ICommand_Tooltip() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Tooltip
		Get
			On Error GoTo ErrorHandler
			
			ICommand_Tooltip = "Create Soils GRIDs"
			
			Exit Property
ErrorHandler: 
			HandleError(True, "ICommand_Tooltip " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
		End Get
	End Property
	
	Private Sub ICommand_OnClick() Implements ESRI.ArcGIS.SystemUI.ICommand.OnClick
		On Error GoTo ErrorHandler
		
		'initialize and show the form
		frmSoils.init(m_App)
		frmSoils.ShowDialog()
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "ICommand_OnClick " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	Private Sub ICommand_OnCreate(ByVal hook As Object) Implements ESRI.ArcGIS.SystemUI.ICommand.OnCreate
		On Error GoTo ErrorHandler
		
		m_App = hook
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "ICommand_OnCreate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
End Class