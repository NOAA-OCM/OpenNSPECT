Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsToolSetup_NET.clsToolSetup")> Public Class clsToolSetup
	Implements ESRI.ArcGIS.SystemUI.IMenuDef
	Implements ESRI.ArcGIS.Framework.IShortcutMenu
	' *************************************************************************************
	' *  Perot Systems Government Services
	' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
	' *  clsToolSetup
	' *************************************************************************************
	' *  Description: Class that defines the tool setup sub-menu in the N-Spect menu
	' *
	' *
	' *  Called By:
	' *************************************************************************************
	' *  Subs:
	' *
	' *  Misc:
	' ************************************************************************************* 'It's a shortcut menu
	
	Private m_pApp As ESRI.ArcGIS.Framework.IApplication
	' Variables used by the Error handler function - DO NOT REMOVE
	Const c_sModuleFileName As String = "clsToolSetup.cls"
	
	Private ReadOnly Property IMenuDef_Caption() As String Implements ESRI.ArcGIS.SystemUI.IMenuDef.Caption
		Get
			On Error GoTo ErrorHandler
			
			IMenuDef_Caption = "Advanced Settings"
			
			Exit Property
ErrorHandler: 
			HandleError(True, "IMenuDef_Caption " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
		End Get
	End Property
	
	Private ReadOnly Property IMenuDef_ItemCount() As Integer Implements ESRI.ArcGIS.SystemUI.IMenuDef.ItemCount
		Get
			On Error GoTo ErrorHandler
			
			IMenuDef_ItemCount = 6
			
			Exit Property
ErrorHandler: 
			HandleError(True, "IMenuDef_ItemCount " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
		End Get
	End Property
	
	Private ReadOnly Property IMenuDef_Name() As String Implements ESRI.ArcGIS.SystemUI.IMenuDef.Name
		Get
			On Error GoTo ErrorHandler
			
			IMenuDef_Name = "NSPECTMenu"
			
			Exit Property
ErrorHandler: 
			HandleError(True, "IMenuDef_Name " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
		End Get
	End Property
	
	Private Sub IMenuDef_GetItemInfo(ByVal pos As Integer, ByVal itemDef As ESRI.ArcGIS.SystemUI.IItemDef) Implements ESRI.ArcGIS.SystemUI.IMenuDef.GetItemInfo
		On Error GoTo ErrorHandler
		
		Select Case pos
			Case 0 'Land Cover Types
				itemDef.id = "NSPECT.clsLCType"
				itemDef.Group = False
			Case 1 'Pollutants
				itemDef.id = "NSPECT.clsPollutants"
				itemDef.Group = False
			Case 2 'Water Quality Standards
				itemDef.id = "NSPECT.clsWQStd"
			Case 3 'Precipitation Scenarios
				itemDef.id = "NSPECT.clsPrecip"
				itemDef.Group = False
			Case 4 'Delineate Watershed
				itemDef.id = "NSPECT.clsWSDelin"
				itemDef.Group = False
			Case 5 'Soils
				itemDef.id = "NSPECT.clsSoils"
				itemDef.Group = False
		End Select
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "IMenuDef_GetItemInfo " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
End Class