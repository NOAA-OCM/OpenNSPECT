Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsMnuProject_NET.clsMnuProject")> Public Class clsMnuProject
	Implements ESRI.ArcGIS.SystemUI.IMenuDef
	Implements ESRI.ArcGIS.Framework.IRootLevelMenu
	' *************************************************************************************
	' *  Perot Systems Government Services
	' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
	' *  clsMnuProject
	' *************************************************************************************
	' *  Description: Class that defines the main menu in the N-Spect Toolbar
	' *
	' *
	' *  Called By:
	' *************************************************************************************
	' *  Subs:
	' *
	' *
	' *  Misc:
	' *************************************************************************************
	
	' This class module implements the following interfaces 'Menu Definition 'Root level so it can hang with the likes of File..., Edit...
	
	' Variables used by the Error handler function - DO NOT REMOVE
	Const c_sModuleFileName As String = "clsMnuProject.cls"
	
	Private ReadOnly Property IMenuDef_Caption() As String Implements ESRI.ArcGIS.SystemUI.IMenuDef.Caption
		Get
			On Error GoTo ErrorHandler
			
			' Menu caption
			IMenuDef_Caption = "N-SPECT"
			
			Exit Property
ErrorHandler: 
			HandleError(True, "IMenuDef_Caption " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
		End Get
	End Property
	Private ReadOnly Property IMenuDef_ItemCount() As Integer Implements ESRI.ArcGIS.SystemUI.IMenuDef.ItemCount
		Get
			On Error GoTo ErrorHandler
			
			' Menu count
			IMenuDef_ItemCount = 3
			
			Exit Property
ErrorHandler: 
			HandleError(True, "IMenuDef_ItemCount " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
		End Get
	End Property
	
	Private ReadOnly Property IMenuDef_Name() As String Implements ESRI.ArcGIS.SystemUI.IMenuDef.Name
		Get
			On Error GoTo ErrorHandler
			
			'Menu name
			IMenuDef_Name = "clsMnuProject"
			
			Exit Property
ErrorHandler: 
			HandleError(True, "IMenuDef_Name " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
		End Get
	End Property
	
	'The menu choices in the main menu of the Toolbar
	Private Sub IMenuDef_GetItemInfo(ByVal pos As Integer, ByVal itemDef As ESRI.ArcGIS.SystemUI.IItemDef) Implements ESRI.ArcGIS.SystemUI.IMenuDef.GetItemInfo
		On Error GoTo ErrorHandler
		
		Select Case Err.Number
			Case Is < 0
				Error(5)
			Case 1
				GoTo Errhand
		End Select
		' Menu items
		Select Case pos
			Case 0
				itemDef.id = "NSPECT.clsNewAnalysis"
				itemDef.Group = False
			Case 1
				'Tool Setup -> (it's a shortcut menu)
				itemDef.id = "NSPECT.clsToolSetup"
				itemDef.Group = True
				'New Analysis...
			Case 2
				itemDef.id = "NSPECT.clsHelp"
				itemDef.Group = True
		End Select
		Exit Sub
Errhand: 
		MsgBox("Error with the IMenuDef in clsMnuProject", MsgBoxStyle.Critical, "error with MenuDef")
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "IMenuDef_GetItemInfo " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
End Class