Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsMainToolbar_NET.clsMainToolbar")> Public Class clsMainToolbar
	Implements ESRI.ArcGIS.SystemUI.IToolBarDef
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
	
	
	' Variables used by the Error handler function - DO NOT REMOVE
	Const c_sModuleFileName As String = "clsMainToolbar.cls"
	
	Private Declare Function SetEnvironmentVariable Lib "kernel32"  Alias "SetEnvironmentVariableA"(ByVal lpName As String, ByVal lpValue As String) As Integer
	
	Private ReadOnly Property IToolBarDef_Caption() As String Implements ESRI.ArcGIS.SystemUI.IToolBarDef.Caption
		Get
			On Error GoTo ErrorHandler
			
			IToolBarDef_Caption = "N-SPECT " & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Revision
			
			Exit Property
ErrorHandler: 
			HandleError(True, "IToolBarDef_Caption " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
		End Get
	End Property
	
	Private ReadOnly Property IToolBarDef_ItemCount() As Integer Implements ESRI.ArcGIS.SystemUI.IToolBarDef.ItemCount
		Get
			On Error GoTo ErrorHandler
			
			IToolBarDef_ItemCount = 1
			
			Exit Property
ErrorHandler: 
			HandleError(True, "IToolBarDef_ItemCount " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
		End Get
	End Property
	
	Private ReadOnly Property IToolBarDef_Name() As String Implements ESRI.ArcGIS.SystemUI.IToolBarDef.Name
		Get
			On Error GoTo ErrorHandler
			
			IToolBarDef_Name = "N-SPECT"
			
			Exit Property
ErrorHandler: 
			HandleError(True, "IToolBarDef_Name " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
		End Get
	End Property
	
	Private Sub IToolBarDef_GetItemInfo(ByVal pos As Integer, ByVal itemDef As ESRI.ArcGIS.SystemUI.IItemDef) Implements ESRI.ArcGIS.SystemUI.IToolBarDef.GetItemInfo
		On Error GoTo ErrorHandler
		
		Select Case pos
			
			Case 0
				itemDef.id = "NSPECT.clsMnuProject"
				itemDef.Group = False
				
		End Select
		
		Exit Sub
ErrorHandler: 
		HandleError(True, "IToolBarDef_GetItemInfo " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		On Error GoTo ErrorHandler
		
		Dim fso As Scripting.FileSystemObject
		Dim nspectPath As String
		Dim docPath As String
		
		fso = New Scripting.FileSystemObject
		
		' Detects and sets the path to N-SPECT's application folder (installation directory)
		nspectPath = My.Application.Info.DirectoryPath
		If Right(nspectPath, 1) = "\" Then
			nspectPath = Left(nspectPath, Len(nspectPath) - 1)
		End If
		If Right(nspectPath, 4) = "\bin" Then
			nspectPath = Left(nspectPath, Len(nspectPath) - 4)
		End If
		'MsgBox "N-SPECT path: " & nspectPath
		modUtil.g_nspectPath = nspectPath
		'SetEnvironmentVariable "NSPECTDAT", nspectPath
		'MsgBox "NSPECTDAT: " & Environ("NSPECTDAT")
		
		' Detects and sets the path to N-SPECT's document folder (within user's "My Documents")
		'If Len(Environ("USERPROFILE")) = 0 Then
		'    Err.Raise vbObjectError + 100, "frmPrj.cmdOpenWS_Click", "Environment variable USERPROFILE is not available."
		'End If
		'docPath = Environ("USERPROFILE") & "\My Documents\" & App.CompanyName
		'If Not fso.FolderExists(docPath) Then
		'    MkDir docPath
		'End If
		'docPath = docPath & "\" & App.ProductName
		'If Not fso.FolderExists(docPath) Then
		'    MkDir docPath
		'End If
		'modUtil.g_nspectDocPath = docPath
		modUtil.g_nspectDocPath = nspectPath
		
		modUtil.ADOConnection()
		
		Exit Sub
		
ErrorHandler: 
		HandleError(True, "Class_Initialize " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 4)
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
End Class